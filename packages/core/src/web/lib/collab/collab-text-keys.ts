/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";
import {
  applyCollaborativeTextDelta,
  applyCollaborativeTextEdit,
  carryTextEditThroughDeltas,
  collaborativeTextAt,
  collaborativeTextChanges,
  collaborativeTextContainerAt,
  jsonValueAt,
  textContainersAlike,
  textContainerWithEditAt,
  textSplice,
  type TextSplice,
  withJsonValueAt,
  withTextContainerWords,
  writeCollaborativeTextContainer,
  type CollaborativeTextDeltaOp,
  type CollaborativeTextEdit,
  type JsonValue,
} from "./json-crdt";
import { updateDeletes } from "./remote-update-payload";

/** A noted keystroke: `base` with [start, end) replaced by `insert`, in the
 * text at `path`, with its container before and after. */
export interface CollaborativeTextKey extends CollaborativeTextEdit {
  path: readonly string[];
  container: { previous: Record<string, JsonValue>; next: Record<string, JsonValue> };
}

export type CollaborativeTextKeyWrite =
  | {
    outcome: "written" | "carried" | "carried-by-words";
    /** The container as the room now holds it. */
    value: Record<string, JsonValue>;
    /** The room's text after the key, and where the key left the caret in it. */
    roomText: string;
    caret: number;
    /** Carried: the editor's words after the key, its caret in them, and the
     * delta that turns them into the room's. */
    carried?: { viewText: string; viewCaret: number; viewToRoom: CollaborativeTextDeltaOp[] };
  }
  | { outcome: string };

/** The key went into the room (into the words typed in, or carried). */
export function textKeyWritten(outcome: string | void | undefined): boolean {
  return outcome === "written" || outcome === "carried" || outcome === "carried-by-words";
}

/** The key went into the room elsewhere than the editor's words show it:
 * the editor is handed the room's words. */
export function textKeyCarried(outcome: string | void | undefined): boolean {
  return outcome === "carried" || outcome === "carried-by-words";
}

/** Versions kept of a text being edited (a key may be typed into any of them). */
const HISTORY_LIMIT = 256;
/** How long the editor may show the words a key was carried from. */
const STALE_VIEW_MS = 10_000;

export class CollaborativeTextKeys {
  private readonly histories = new Map<string, {
    text: Y.Text;
    versions: Array<{ text: string; delta: CollaborativeTextDeltaOp[] | null }>;
    /** How many of the oldest versions were dropped (HISTORY_LIMIT). */
    dropped: number;
  }>();
  private readonly staleViews = new Map<string, { path: readonly string[]; views: Record<string, JsonValue>[]; at: number }>();
  /** The editor's words after each key carried (`write`), with the delta to
   * the room's words after it and that version's place in the history: a
   * next key typed into those words, before the editor drew the room's, is
   * carried from them. */
  private readonly viewVersions = new Map<string, Array<{ text: string; delta: CollaborativeTextDeltaOp[]; version: number }>>();

  constructor(private readonly root: Y.Map<unknown>) {
    root.observeDeep(this.record as never);
  }

  destroy(): void {
    this.root.unobserveDeep(this.record as never);
    this.histories.clear();
    this.staleViews.clear();
  }

  /** The text at `path` is being edited: its versions are kept from now on. */
  track(path: readonly string[]): void {
    const text = collaborativeTextAt(this.root, path);
    if (text) this.history(path.join("\u0000"), text);
  }

  private history(key: string, text: Y.Text) {
    let history = this.histories.get(key);
    if (!history || history.text !== text) {
      history = { text, versions: [{ text: text.toString(), delta: null }], dropped: 0 };
      this.histories.set(key, history);
      this.viewVersions.delete(key);
    }
    return history;
  }

  /** Every transaction on the room (anyone's): a new version of each text kept. */
  private readonly record = (events: Y.YEvent<never>[]) => {
    if (this.histories.size === 0) return;
    const changed = new Map<string, CollaborativeTextDeltaOp[]>();
    for (const change of collaborativeTextChanges(events)) changed.set(change.path.join("\u0000"), [...change.delta]);
    for (const [key, history] of this.histories) {
      const now = history.text.toString();
      const last = history.versions[history.versions.length - 1]!;
      if (last.text === now) continue;
      const delta = changed.get(key) ?? null;
      history.versions.push({ text: now, delta: delta && applyCollaborativeTextDelta(last.text, delta) === now ? delta : null });
      if (history.versions.length > HISTORY_LIMIT) {
        const drop = history.versions.length - HISTORY_LIMIT;
        history.versions.splice(0, drop);
        history.dropped += drop;
      }
    }
  };

  /** The room's deltas of a text since it last read `base` (a version of
   * the room's, or the editor's words after a carried key), or null. */
  private deltasSince(key: string, base: string): CollaborativeTextDeltaOp[][] | null {
    const history = this.histories.get(key);
    if (!history) return null;
    const { versions } = history;
    const since = (index: number, first: CollaborativeTextDeltaOp[][]) => {
      const deltas = [...first];
      for (let next = index + 1; next < versions.length; next += 1) {
        const delta = versions[next]!.delta;
        if (!delta) return null;
        deltas.push(delta);
      }
      return deltas;
    };
    for (let index = versions.length - 1; index >= 0; index -= 1) {
      if (versions[index]!.text === base) return since(index, []);
    }
    const views = this.viewVersions.get(key) ?? [];
    for (let index = views.length - 1; index >= 0; index -= 1) {
      const view = views[index]!;
      // Its room version must still be in the kept history (offsets shift as
      // the oldest are dropped: `dropped` counts them).
      const at = view.version - history.dropped;
      if (view.text === base && at >= 0 && at < versions.length) return since(at, [view.delta]);
    }
    return null;
  }

  /**
   * Write the key into the room now, in `doc`'s transaction of `origin`:
   * into the words it was typed in when the room holds them, else carried
   * to where they now stand. Any other outcome names why it was not
   * written (the caller publishes it as before).
   */
  write(doc: Y.Doc, origin: unknown, edit: CollaborativeTextKey): CollaborativeTextKeyWrite {
    const key = edit.path.join("\u0000");
    const containerPath = edit.path.slice(0, -1);
    const text = collaborativeTextAt(this.root, edit.path);
    if (!text) return { outcome: "no-text" };
    this.history(key, text);
    const { container } = edit;
    const roomText = text.toString();
    const splice = { start: edit.start, end: edit.end, insert: edit.insert };
    let written: Record<string, JsonValue> | null = null;
    if (roomText === edit.base) {
      doc.transact(() => {
        written = writeCollaborativeTextContainer(this.root, containerPath, container.previous, container.next,
          { base: edit.base, ...splice });
      }, origin);
      if (written) {
        // Written into the words the editor typed in: it is in step.
        this.staleViews.delete(key);
        return { outcome: "written", value: written, roomText: text.toString(), caret: edit.start + edit.insert.length };
      }
    }
    // The words it was typed into are none the room or the editor is known
    // to have held (their history began after them): the room's change of
    // them is read as the one splice between the two, and the key keeps its
    // place against the characters around it, as a publication's string
    // comparison placed it -- but written as a key, so the editor's words
    // that lack a collaborator's characters never read as deleting them.
    const known = this.deltasSince(key, edit.base);
    const since = known ?? [spliceDelta(textSplice(edit.base, roomText))];
    const carried = carryTextEditThroughDeltas(edit.base, splice, since);
    const room = collaborativeTextContainerAt(this.root, containerPath);
    const roomNext = carried && room
      ? textContainerWithEditAt(room, container.previous, container.next, { base: edit.base, ...splice }, carried)
      : null;
    if (!carried || !room || !roomNext) return { outcome: "room-differs-uncarried" };
    doc.transact(() => {
      written = writeCollaborativeTextContainer(this.root, containerPath, room, roomNext,
        { base: roomText, start: carried.start, end: carried.end, insert: edit.insert });
    }, origin);
    if (!written) return { outcome: "room-differs-unwritten" };
    // The editor shows the words it typed in until it is handed the room's:
    // its next key may be typed into them (`deltasSince`).
    const history = this.histories.get(key);
    if (history) {
      const viewText = applyCollaborativeTextEdit({ base: edit.base, ...splice });
      const list = this.viewVersions.get(key) ?? [];
      list.push({ text: viewText, delta: carried.viewToRoom, version: history.dropped + history.versions.length - 1 });
      if (list.length > 64) list.splice(0, list.length - 64);
      this.viewVersions.set(key, list);
    }
    const views = this.staleViews.get(key) ?? { path: containerPath, views: [], at: 0 };
    views.views.push(container.next);
    if (views.views.length > 32) views.views.splice(0, views.views.length - 32);
    views.at = Date.now();
    this.staleViews.set(key, views);
    return {
      outcome: known ? "carried" : "carried-by-words", value: written, roomText: text.toString(), caret: carried.start + edit.insert.length,
      carried: {
        viewText: applyCollaborativeTextEdit({ base: edit.base, ...splice }),
        viewCaret: edit.start + edit.insert.length,
        viewToRoom: carried.viewToRoom,
      },
    };
  }

  /**
   * `local` (an editor's document) with each text container that still
   * shows words a key was carried from (`write`) given the words `base`
   * (the room's) has there: no edit of this member's.
   */
  settle<T extends JsonValue>(local: T, base: JsonValue): T {
    if (this.staleViews.size === 0) return local;
    let out: JsonValue = local;
    const now = Date.now();
    for (const [key, entry] of this.staleViews) {
      if (now - entry.at > STALE_VIEW_MS) {
        this.staleViews.delete(key);
        continue;
      }
      const mine = jsonValueAt(out, entry.path);
      const room = jsonValueAt(base, entry.path);
      if (!isRecord(mine) || !isRecord(room) || textContainersAlike(mine, room) ||
          !entry.views.some((view) => textContainersAlike(view, mine))) continue;
      out = withJsonValueAt(out, entry.path, withTextContainerWords(mine, room));
    }
    return out as T;
  }
}

/** A splice of a text as a Y.Text delta of it. */
function spliceDelta({ start, end, insert }: TextSplice): CollaborativeTextDeltaOp[] {
  const delta: CollaborativeTextDeltaOp[] = [];
  if (start > 0) delta.push({ retain: start });
  if (end > start) delta.push({ delete: end - start });
  if (insert) delta.push({ insert });
  return delta;
}

function isRecord(value: JsonValue | undefined): value is Record<string, JsonValue> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Noted keys that waited, written by `write` (the room's write of a key when
 * typed) one by one in the order they were typed, each carried from the
 * words it was typed into. A text's keys after one not written keep waiting
 * with it. Returns the notes not written, in their order.
 */
export function writeWaitingTextEdits<T extends CollaborativeTextEdit & { path: readonly string[]; container?: CollaborativeTextKey["container"] }>(
  noted: T[],
  write: (edit: CollaborativeTextKey) => string,
): T[] {
  const stopped = new Set<string>();
  const left: T[] = [];
  for (const edit of noted) {
    const key = edit.path.join("\u0000");
    if (stopped.has(key) || !edit.container) {
      stopped.add(key);
      left.push(edit);
      continue;
    }
    const outcome = write({ ...edit, container: edit.container });
    if (textKeyWritten(outcome)) continue;
    stopped.add(key);
    left.push(edit);
  }
  return left;
}

/** How long, and how many, keys typed into words the room does not hold
 * wait for a peer's words (`OrphanTextEdits`). */
const ORPHAN_EDITS_MS = 120_000;
const ORPHAN_EDITS_LIMIT = 512;
















/** The items a publication wrote: this member's client, clocks [from, to). */
export interface OwnItems { client: number; from: number; to: number }

interface OrphanEntry<T> { edit: T; at: number; own: OwnItems | null }

export class OrphanTextEdits<T extends CollaborativeTextEdit & { path: readonly string[]; container?: CollaborativeTextKey["container"] }> {
  /** Each key, when it was kept, and the items its publication wrote. */
  private edits: Array<OrphanEntry<T>> = [];

  get size(): number {
    return this.edits.length;
  }

  /** After a publication of `noted` (which wrote `own`'s items): those whose
   * words are not in `root`. */
  keep(root: Y.Map<unknown>, noted: readonly T[], own: OwnItems | null): void {
    const now = Date.now();
    for (const edit of noted) {
      if (edit.container && !collaborativeTextAt(root, edit.path)) this.edits.push({ edit, at: now, own });
    }
    if (this.edits.length > ORPHAN_EDITS_LIMIT) this.edits.splice(0, this.edits.length - ORPHAN_EDITS_LIMIT);
  }

  /** After a change of this member's own: the keys whose words it brought go. */
  forgetPresent(root: Y.Map<unknown>): void {
    if (this.edits.length > 0) this.edits = this.edits.filter(({ edit }) => !collaborativeTextAt(root, edit.path));
  }

  /**
   * A peer's update (`payload`, when the room knows it) just brought words
   * some keys were typed into: those keys the peer did not hold, each text's
   * in order, taken off the list. A peer that read this member's words where
   * they were written (the drawing) before replacing them holds them, and
   * its update deletes the items that wrote them: those keys are not written
   * twice. Without the update, the peer's words tell it: from the first key
   * whose words they are not.
   */
  arrived(root: Y.Map<unknown>, payload?: Uint8Array): T[] {
    const now = Date.now();
    this.edits = this.edits.filter(({ at }) => now - at <= ORPHAN_EDITS_MS);
    const byText = new Map<string, Array<OrphanEntry<T>>>();
    const left: Array<OrphanEntry<T>> = [];
    for (const entry of this.edits) {
      if (!collaborativeTextAt(root, entry.edit.path)) {
        left.push(entry);
        continue;
      }
      const key = entry.edit.path.join("\u0000");
      byText.set(key, [...(byText.get(key) ?? []), entry]);
    }
    this.edits = left;
    const out: T[] = [];
    for (const entries of byText.values()) {
      if (payload && entries.every(({ own }) => own)) {
        out.push(...entries.filter(({ own }) => !updateDeletes(payload, own!.client, own!.from, own!.to)).map(({ edit }) => edit));
        continue;
      }
      const words = collaborativeTextAt(root, entries[0]!.edit.path)?.toString() ?? "";
      let from = 0;
      for (let at = entries.length - 1; at >= 0; at -= 1) {
        if (applyCollaborativeTextEdit(entries[at]!.edit) === words) {
          from = at + 1;
          break;
        }
      }
      out.push(...entries.slice(from).map(({ edit }) => edit));
    }
    return out;
  }

  clear(): void {
    this.edits = [];
  }
}
