/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  applyCollaborativeTextDelta,
  mapIndexThroughTextDelta,
  textSplice,
  textSpliceIsGuess,
  type CollaborativeTextDeltaOp,
  type JsonValue,
} from "@mrnote/core/web/lib/collab/json-crdt.ts";
import type { PageRemoteContentDetail } from "@mrnote/core/web/hooks/usePageContentSync.ts";
import { mapCaretThroughTextChange } from "@mrnote/core/web/lib/presentation-collab-text.ts";
import { presentationPreviewPlainText, presentationPreviewSelection } from "@mrnote/core/web/lib/presentation/text-selection.ts";

export interface TextRange {
  start: number;
  end: number;
}

/**
 * The words a browser's replacement (`insertReplacementText`: a spelling
 * suggestion, an autocorrection) put in place of others, given by the event
 * that made it; null for any other input.
 */
export function browserReplacement(event: Event | null | undefined): string | null {
  const input = event as (InputEvent & { dataTransfer?: DataTransfer | null }) | null | undefined;
  if (!input || input.inputType !== "insertReplacementText") return null;
  const words = input.data ?? input.dataTransfer?.getData("text/plain") ?? null;
  return words ? words.replace(/\r\n?/g, "\n") : null;
}

/**
 * Where a browser's replacement was made: the one range of `before` whose
 * replacement by `replacement` makes `after`. The browser replaces a word away
 * from the caret (the caret stays after what was typed since), so neither the
 * caret nor the selection says where; the replacing words do, when they fit
 * one place only ("helo" -> "hello": comparing the two texts alone could put
 * the new "l" after either "l", TRIAGE 49). Null when they fit no place or
 * more than one.
 */
export function replacementRange(before: string, after: string, replacement: string): TextRange | null {
  const removed = before.length - (after.length - replacement.length);
  if (!replacement || removed < 0) return null;
  let found: TextRange | null = null;
  for (let at = after.indexOf(replacement); at >= 0; at = after.indexOf(replacement, at + 1)) {
    if (at > before.length - removed) break;
    if (after.slice(0, at) !== before.slice(0, at) || after.slice(at + replacement.length) !== before.slice(at + removed)) continue;
    if (found) return null;
    found = { start: at, end: at + removed };
  }
  return found;
}

/** A slide object's words in the room (json-crdt collaborative text). */
export function slideBlockTextPath(slideId: string, blockId: string): string[] {
  return ["slides", slideId, "blocks", blockId, "text"];
}

/** A table cell's words in the room (json-crdt collaborative text). */
export function tableCellTextPath(slideId: string, blockId: string, row: number, column: number): string[] {
  return ["slides", slideId, "blocks", blockId, "rows", String(row), String(column), "text"];
}

/** A keystroke's text container (a slide object, a table cell) before and
 * after it, as the room writes it when it is typed (usePageContentSync
 * `noteTextEdit`); undefined when either is not a plain object. */
export function textContainerOf(
  previous: unknown,
  next: unknown,
): { previous: Record<string, JsonValue>; next: Record<string, JsonValue> } | undefined {
  const plain = (value: unknown): value is Record<string, JsonValue> =>
    value !== null && typeof value === "object" && !Array.isArray(value);
  return plain(previous) && plain(next) ? { previous, next } : undefined;
}

/** The keystroke that turned a cell's `before` into `after`, read from the
 * caret it left when it can be, else the one change between the two. */
export function cellKeystrokeSplice(
  before: string,
  after: string,
  caretAfter: number | undefined,
): { start: number; end: number; insert: string } | null {
  if (before === after) return null;
  // Without a caret, a key read off the two texts among the same letters is a
  // guess of where it went: not noted (the publication refuses a guess, TRIAGE 47).
  return caretSplice(before, after, caretAfter) ?? (textSpliceIsGuess(before, after) ? null : textSplice(before, after));
}












export function cellKeystrokeEdit(
  live: string,
  drawn: string | undefined,
  after: string,
  caretAfter: number | undefined,
): { base: string; start: number; end: number; insert: string } | null {
  const viaLive = cellKeystrokeSplice(live, after, caretAfter);
  if (drawn === undefined || drawn === live) return viaLive ? { base: live, ...viaLive } : null;
  const viaDrawn = cellKeystrokeSplice(drawn, after, caretAfter);
  if (viaDrawn && (!viaLive || spliceSize(viaDrawn) < spliceSize(viaLive))) return { base: drawn, ...viaDrawn };
  return viaLive ? { base: live, ...viaLive } : null;
}

/** The room's delta of one slide object's words in a delivery, if it has one. */
export function deliveredBlockTextDelta(
  detail: PageRemoteContentDetail | undefined,
  slideId: string,
  blockId: string,
): readonly CollaborativeTextDeltaOp[] | null {
  const change = detail?.textChanges?.find(({ path }) =>
    path.length === 5 && path[0] === "slides" && path[1] === slideId &&
    path[2] === "blocks" && path[3] === blockId && path[4] === "text");
  return change ? change.delta : null;
}

/**
 * The caret this member placed at `range` of `before`, where the room's
 * relative positions carried it in `after` (`PageRemoteContentDetail.carets`),
 * or null when the room holds no such caret for the text at `path` (placed
 * elsewhere, or in words the room had moved past).
 */
export function deliveredAnchoredCaret(
  detail: PageRemoteContentDetail | undefined,
  path: readonly string[],
  before: string,
  range: TextRange,
  after: string,
): TextRange | null {
  const key = path.join("\u0000");
  const caret = detail?.carets?.find((candidate) => candidate.path.join("\u0000") === key);
  if (!caret || caret.text !== after || caret.from.text !== before ||
      caret.from.start !== range.start || caret.from.end !== range.end) return null;
  return { start: caret.start, end: caret.end };
}

/**
 * `range` of `before` carried to `after`. With the change's delta (and the
 * delta is of `before`, producing `after`), exactly; otherwise by comparing
 * the two strings. A collapsed caret stays behind a peer's insertion at it;
 * a selection takes in neither edge's insertion.
 */
export function mapTextRangeThroughChange(
  range: TextRange,
  before: string,
  after: string,
  delta: readonly CollaborativeTextDeltaOp[] | null,
): TextRange {
  if (before === after && !delta) return range;
  if (delta && applyCollaborativeTextDelta(before, delta) === after) {
    if (range.start === range.end) {
      const at = mapIndexThroughTextDelta(range.start, delta, false);
      return { start: at, end: at };
    }
    const start = mapIndexThroughTextDelta(range.start, delta, true);
    return { start, end: Math.max(start, mapIndexThroughTextDelta(range.end, delta, false)) };
  }
  const start = mapCaretThroughTextChange(range.start, before, after);
  return { start, end: Math.max(start, mapCaretThroughTextChange(range.end, before, after)) };
}

/**
 * The peers' changes of slide objects' words that reached the live deck
 * since it was last drawn, per object: a keystroke read from the drawn words
 * is carried through them to where it belongs in the live words. `deltas`
 * is null once a change without a known delta arrived (no exact carry).
 */
export type UnrenderedTextChanges = Map<string, { base: string; deltas: CollaborativeTextDeltaOp[][] | null }>;

const changeKey = (slideId: string, blockId: string) => `${slideId}\u0000${blockId}`;

/** Record one delivery's change of an object's words (`before` -> `after`). */
export function recordUnrenderedTextChange(
  changes: UnrenderedTextChanges,
  slideId: string,
  blockId: string,
  before: string,
  after: string,
  delta: readonly CollaborativeTextDeltaOp[] | null,
): void {
  if (before === after) return;
  const key = changeKey(slideId, blockId);
  const entry = changes.get(key) ?? { base: before, deltas: [] };
  const exact = delta !== null && applyCollaborativeTextDelta(before, delta) === after;
  entry.deltas = entry.deltas && exact ? [...entry.deltas, [...delta]] : null;
  changes.set(key, entry);
}

/**
 * A keystroke typed into `base` (the drawn words) at `edit`, carried into
 * the live words: the live range, or null when it cannot be carried exactly
 * (the caller then merges as before).
 */
export function carryEditThroughUnrendered(
  changes: UnrenderedTextChanges,
  slideId: string,
  blockId: string,
  base: string,
  live: string,
  edit: TextRange,
): TextRange | null {
  const entry = changes.get(changeKey(slideId, blockId));
  if (!entry || entry.base !== base || !entry.deltas) return null;
  let text = base;
  let range = edit;
  for (const delta of entry.deltas) {
    const next = applyCollaborativeTextDelta(text, delta);
    if (next === null) return null;
    range = mapTextRangeThroughChange(range, text, next, delta);
    text = next;
  }
  return text === live ? range : null;
}

/** The caret in the focused drawn (rich) body of a slide object -- a text
 * box's or a shape's label -- if any. */
export function focusedRichPreviewCaret(root: HTMLElement | null): { blockId: string; text: string; range: TextRange } | null {
  const active = root?.ownerDocument.activeElement;
  if (!root || !(active instanceof HTMLElement) || !root.contains(active)) return null;
  const host = active.closest<HTMLElement>('[data-selection-owner="rich-preview"], [data-shape-rich-editor="true"]');
  if (!host || !root.contains(host)) return null;
  const blockId = host.closest("[data-block-id]")?.getAttribute("data-block-id");
  if (!blockId) return null;
  const text = presentationPreviewPlainText(host);
  const range = presentationPreviewSelection(host, text.length);
  return range ? { blockId, text, range: { start: range.start, end: range.end } } : null;
}

/** The caret in the words of a slide object that hold the keyboard: its
 * drawn (rich) body, or the field its words are typed through (a file's
 * label whose drawn words take no keys of their own), if any. */
export function focusedWordsCaret(root: HTMLElement | null): { blockId: string; text: string; range: TextRange } | null {
  const rich = focusedRichPreviewCaret(root);
  if (rich) return rich;
  const active = root?.ownerDocument.activeElement;
  if (!root || !(active instanceof HTMLTextAreaElement) || !root.contains(active)) return null;
  const blockId = active.closest("[data-block-id]")?.getAttribute("data-block-id");
  return blockId ? { blockId, text: active.value, range: { start: active.selectionStart, end: active.selectionEnd } } : null;
}

/** The focused field of a table cell (PresentationCellRichTextEditor: its
 * words and caret read as a textarea's are), if any. */
export function focusedTableCellField(root: HTMLElement | null): {
  field: HTMLElement & { carryCaret?: (words: string, start: number, end: number) => void };
  blockId: string; row: number; column: number; text: string; range: TextRange;
} | null {
  const active = root?.ownerDocument.activeElement as
    (HTMLElement & { value?: unknown; selectionStart?: unknown; selectionEnd?: unknown }) | null | undefined;
  if (!root || !active?.isContentEditable || !root.contains(active)) return null;
  const cell = active.closest<HTMLElement>('[data-testid="presentation-table-cell"]');
  const blockId = cell?.closest("[data-block-id]")?.getAttribute("data-block-id");
  const row = Number(cell?.getAttribute("data-table-row"));
  const column = Number(cell?.getAttribute("data-table-col"));
  if (!cell || !blockId || !Number.isInteger(row) || !Number.isInteger(column) || typeof active.value !== "string" ||
      typeof active.selectionStart !== "number" || typeof active.selectionEnd !== "number") return null;
  return { field: active, blockId, row, column, text: active.value,
    range: { start: active.selectionStart, end: active.selectionEnd } };
}

/**
 * The keystroke `before` -> `after` as a splice at `selection` (the
 * selection it replaced, read before the input), or null when `after` is
 * not `before` with that selection replaced.
 */
export function keystrokeSplice(
  before: string,
  after: string,
  selection: TextRange | undefined,
): { start: number; end: number; insert: string } | null {
  if (!selection) return null;
  const { start, end } = selection;
  if (start < 0 || start > end || end > before.length) return null;
  const tail = before.length - end;
  if (after.length < start + tail) return null;
  if (!after.startsWith(before.slice(0, start)) || !after.endsWith(before.slice(end))) return null;
  return { start, end, insert: after.slice(start, after.length - tail) };
}

/**
 * The keystroke `before` -> `after` as a splice, read from the caret it left
 * (`caretAfter`, collapsed): a letter typed or a character deleted at it.
 * Null when `after` is not such an edit of `before`.
 */
export function caretSplice(
  before: string,
  after: string,
  caretAfter: number | undefined,
): { start: number; end: number; insert: string } | null {
  if (caretAfter === undefined || caretAfter < 0 || caretAfter > after.length) return null;
  const grown = after.length - before.length;
  if (grown > 0) {
    const start = caretAfter - grown;
    if (start < 0 || after.slice(0, start) !== before.slice(0, start) || after.slice(caretAfter) !== before.slice(start)) return null;
    return { start, end: start, insert: after.slice(start, caretAfter) };
  }
  if (grown < 0) {
    const end = caretAfter - grown;
    if (end > before.length || after.slice(0, caretAfter) !== before.slice(0, caretAfter) || after.slice(caretAfter) !== before.slice(end)) return null;
    return { start: caretAfter, end, insert: "" };
  }
  return null;
}

/**
 * A keystroke read from the drawn words of one slide object (`rendered`),
 * as it belongs in its live words (`live`, which a collaborator's delivery
 * may have changed since the screen was drawn). `onLive`: write `text` over
 * the live object; else over the drawn one, as before (no exact carry was
 * possible and the caller merges). `splice` is the edit of the words it is
 * written over, when exactly known (the room then writes it there). `shift`
 * moves the caret the keystroke left.
 */
export function carryKeystrokeToLive(options: {
  changes: UnrenderedTextChanges;
  slideId: string;
  blockId: string;
  rendered: string;
  live: string;
  typed: string;
  selection?: TextRange;
  caretAfter?: number;
}): {
  onLive: boolean;
  text: string;
  selection?: TextRange;
  shift: number;
  splice: { start: number; end: number; insert: string } | null;
  /** How the keystroke was read: `exact` (at its selection or caret, in the
   * drawn or the live words), `compared` (no selection or caret told it: the
   * field's words compared with the live ones, which they differ from by
   * less than from the drawn ones) or `stale` (the live words are not the
   * drawn ones and neither read it: a comparison with the drawn words would
   * take the collaborator's newer letters for this key's, see below). */
  read: KeystrokeRead;
} {
  const { changes, slideId, blockId, rendered, live, typed, selection, caretAfter } = options;
  const read = (words: string) => keystrokeSplice(words, typed, selection) ?? caretSplice(words, typed, caretAfter);
  // Read as typed into the drawn words first (what the screen showed); only
  // words that cannot be those are the live ones (an earlier key's commit
  // not drawn yet).
  const drawn = read(rendered);
  if (live === rendered) {
    // The range the key was read at (its selection, or the caret it left when
    // the selection captured for it was another key's): a selection that did
    // not read it is not handed on (TRIAGE 49: the caller then read the key
    // off the two texts, a guess among repeated letters).
    return { onLive: false, text: typed, selection: drawn ? { start: drawn.start, end: drawn.end } : selection, shift: 0, splice: drawn,
      read: drawn ? "exact" : "compared" };
  }
  if (drawn) {
    const carried = carryEditThroughUnrendered(changes, slideId, blockId, rendered, live,
      { start: drawn.start, end: drawn.end });
    if (carried) {
      return {
        onLive: true,
        text: `${live.slice(0, carried.start)}${drawn.insert}${live.slice(carried.end)}`,
        selection: carried,
        shift: carried.start - drawn.start,
        splice: { ...carried, insert: drawn.insert },
        read: "exact",
      };
    }
    return { onLive: false, text: typed, selection, shift: 0, splice: null, read: "exact" };
  }
  const onLive = read(live);
  if (onLive) return { onLive: true, text: typed, selection: { start: onLive.start, end: onLive.end }, shift: 0, splice: onLive, read: "exact" };
  // Neither selection nor caret told the key (a field whose change event
  // carries neither). The field already shows a collaborator's delivery the
  
  // 25 ms after A's "l" at its end was drawn into the field): compared with
  // the drawn words, every character after the caret differs at the end, and
  // the key read as "replace all of them with this key and all of them
  // again" -- written so into the room it deleted the words and put a copy
  // back, A's concurrent letter twice. The field's words are the live ones
  // with this key in them when that comparison is the smaller change.
  const viaLive = textSplice(live, typed);
  const viaDrawn = textSplice(rendered, typed);
  if (spliceSize(viaLive) < spliceSize(viaDrawn)) {
    return { onLive: true, text: typed, selection: { start: viaLive.start, end: viaLive.end }, shift: 0, splice: viaLive, read: "compared" };
  }
  return { onLive: false, text: typed, selection, shift: 0, splice: null, read: "stale" };
}

/** How `carryKeystrokeToLive` read a keystroke. */
export type KeystrokeRead = "exact" | "compared" | "stale";

/** Characters a splice removes plus the ones it writes. */
function spliceSize(splice: { start: number; end: number; insert: string }): number {
  return splice.end - splice.start + splice.insert.length;
}

/**
 * A key typed into a drawn body right after a collaborator's delivery, moved
 * to the caret the delivery carried (`delivered`, in the words `before` the
 * key was typed into) when the browser put it elsewhere: the body is drawn
 * again with the carried caret, but a key taken before that render lands at
 * the caret's old offset. Only a plain insertion at a collapsed caret is
 * moved; null when the key is already there or is anything else.
 */
export function keyAtDeliveredCaret(
  before: string,
  typed: string,
  delivered: { words: string; start: number; end: number },
  caretAfter?: number,
): { text: string; start: number; shift: number } | null {
  if (delivered.words !== before || delivered.start !== delivered.end || typed.length <= before.length) return null;
  const splice = (caretAfter !== undefined ? caretSplice(before, typed, caretAfter) : null) ?? textSplice(before, typed);
  if (splice.start !== splice.end || splice.insert.length === 0 || splice.start === delivered.start) return null;
  if (delivered.start < 0 || delivered.start > before.length) return null;
  return {
    text: before.slice(0, delivered.start) + splice.insert + before.slice(delivered.start),
    start: delivered.start,
    shift: delivered.start - splice.start,
  };
}
