/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Generic JSON <-> Yjs mapping so EVERY page window co-edits, not just the
 * ones with a native CRDT binding.
 *
 * Rich text (docs), code and notebook cells have character-level bindings
 * already. The remaining editors — spreadsheet, presentation, PDF
 * annotations — own a plain JSON document, and a whole-document "last save
 * wins" would make two people's simultaneous work destroy each other. This
 * module mirrors that JSON into a Yjs tree so concurrent edits merge at the
 * granularity that matters to the user:
 *
 *   object  -> Y.Map, one entry per key      (different cells / settings merge)
 *   array of {id} -> Y.Map keyed by id + Y.Array of ids for order
 *                    (different slides / annotations merge; reorder is
 *                     last-writer-wins on the order array only)
 *   comment thread -> the same keyed representation, using its message id,
 *                     native Office id, or deterministic legacy position
 *   array (other) -> replaced wholesale       (short primitive lists)
 *   scalar  -> stored as-is                   (same key concurrently = LWW)
 *   collaborative text -> Y.Text              (characters merge; see
 *                     `isCollaborativeTextField`: a slide object's `text`)
 *
 * Only CHANGED paths are written, so an editor re-emitting an unchanged
 * document produces no Yjs update and therefore no network traffic.
 */

import * as Y from "yjs";

export type JsonValue =
  | string
  | number
  | boolean
  | null
  | JsonValue[]
  | { [key: string]: JsonValue };

/** Marks a Y.Map that represents an id-keyed array (see module docs). */
const LIST_MARKER = "__mrn_list__";
const LIST_ORDER = "__mrn_order__";
const LIST_ITEMS = "__mrn_items__";

function isPlainObject(value: unknown): value is Record<string, JsonValue> {
  return (
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    Object.getPrototypeOf(value) !== null
  );
}

/** Optional JS object properties are absent on the JSON wire. Never turn
 * an undefined field into an authored null, or treat it as a known peer key. */
function hasJsonProperty(value: Record<string, JsonValue>, key: string): boolean {
  return Object.hasOwn(value, key) && value[key] !== undefined;
}

const COMMENT_MESSAGE_KEYS = new Set([
  "id",
  "text",
  "author",
  "actorId",
  "ts",
  "nativeId",
  "personId",
  "dateTime",
]);

function commentMessageIdentity(
  item: Record<string, JsonValue>,
  index: number,
): string | null {
  if (
    typeof item.text !== "string" ||
    item.text.length === 0 ||
    Object.keys(item).some((key) => hasJsonProperty(item, key) && !COMMENT_MESSAGE_KEYS.has(key))
  ) {
    return null;
  }
  if (typeof item.id === "string" && item.id.length > 0) {
    return `comment:${item.id}`;
  }
  if (typeof item.nativeId === "string" && item.nativeId.length > 0) {
    return `comment:native:${item.nativeId}`;
  }
  if (
    typeof item.actorId === "string" &&
    item.actorId.length > 0 &&
    item.actorId.length <= 200 &&
    typeof item.ts === "number" &&
    Number.isFinite(item.ts)
  ) {
    return `comment:actor:${item.actorId.slice(0, 128)}:${item.ts}`;
  }
  return `comment:legacy:${index}`;
}

/** The stable per-item identities for a mergeable list, or null for LWW arrays. */
function keyedListIds(value: unknown, field?: string): string[] | null {
  // NOTE: an EMPTY array is deliberately decided by the caller using the
  // existing container shape, so a list does not flip representation between
  // empty and non-empty and lose the first concurrently-added item.
  if (!Array.isArray(value) || value.length === 0) return null;
  if (field === "thread" && value.every(isPlainObject)) {
    const ids = value.map((item, index) =>
      commentMessageIdentity(item as Record<string, JsonValue>, index),
    );
    if (
      ids.every((id): id is string => id !== null) &&
      new Set(ids).size === ids.length
    ) {
      return ids;
    }
  }
  if (
    value.every(
      (item) =>
        isPlainObject(item) &&
        typeof item.id === "string" &&
        item.id.length > 0,
    )
  ) {
    const ids = value.map((item) => (item as { id: string }).id);
    return new Set(ids).size === ids.length ? ids : null;
  }
  return null;
}

// ---------------------------------------------------------------------------
// Collaborative text
//
// Two people typing into the same slide text box used to be last-writer-wins
// on the whole string: each keystroke rewrote `text`, so one person's words
// vanished under the other's next key. Those strings are stored as Y.Text
// instead, and every write is a character splice, so concurrent typing keeps
// both people's characters.
//
// The policy is deliberately narrow and structural: the `text` key of an
// entry of an id-keyed list named `blocks` (a presentation slide's objects —
// text boxes and shapes). Other `text` keys — PDF annotations, comment
// thread messages — stay scalars: their editors replace them whole, and a
// different wire shape there would be a model change for those rooms.
//
// The server adapter (apps/api/app/services/page_causal_crdt.py) applies the
// same policy; the two must stay in step.
// ---------------------------------------------------------------------------

const COLLABORATIVE_TEXT_KEY = "text";








const TABLE_ROWS = "\u0000rows";
const TABLE_ROW = "\u0000row";
const TABLE_CELL = "\u0000cell";
const COLLABORATIVE_TEXT_LISTS: ReadonlySet<string> = new Set(["blocks", TABLE_CELL]);

/** The context of `key`'s value in a map that is an entry of `itemOf`: a
 *  slide object's `rows` (a table's; another object's rows hold scalars, on
 *  which the context changes nothing), or nothing. The server adapter
 *  (page_causal_crdt.py) reads the same structure. */
function childArrayContext(key: string, itemOf: string | undefined): string | undefined {
  return key === "rows" && itemOf === "blocks" ? TABLE_ROWS : undefined;
}

/** The context of an entry of an array whose context is `context`. */
function arrayEntryContext(context: string | undefined): string | undefined {
  return context === TABLE_ROWS ? TABLE_ROW : context === TABLE_ROW ? TABLE_CELL : undefined;
}

/** `key` belongs to an object that is an entry of the keyed list `itemOf`. */
function isCollaborativeTextField(key: string, itemOf: string | undefined): boolean {
  return key === COLLABORATIVE_TEXT_KEY && itemOf !== undefined &&
    COLLABORATIVE_TEXT_LISTS.has(itemOf);
}

const isHighSurrogate = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLowSurrogate = (code: number) => code >= 0xdc00 && code <= 0xdfff;

/** One contiguous edit of `base`: replace [start, end) with `insert`. */
export interface TextSplice {
  start: number;
  end: number;
  insert: string;
}

/**
 * The single splice turning `from` into `to` (common prefix and suffix
 * removed). Offsets are UTF-16 code units — the unit of both JS strings and
 * Y.Text — and never fall between the two halves of a surrogate pair, so a
 * concurrent edit can not be spliced into the middle of an emoji.
 */
export function textSplice(from: string, to: string): TextSplice {
  const limit = Math.min(from.length, to.length);
  let prefix = 0;
  while (prefix < limit && from.charCodeAt(prefix) === to.charCodeAt(prefix)) prefix += 1;
  if (prefix > 0 && isHighSurrogate(from.charCodeAt(prefix - 1)) &&
      (prefix < from.length || prefix < to.length)) {
    // The pair's high halves agree but its low halves (or presence) differ.
    prefix -= 1;
  }
  let suffix = 0;
  while (suffix < limit - prefix &&
    from.charCodeAt(from.length - 1 - suffix) === to.charCodeAt(to.length - 1 - suffix)) {
    suffix += 1;
  }
  if (suffix > 0 && isLowSurrogate(from.charCodeAt(from.length - suffix))) {
    // The suffix starts on a low half whose high half differs.
    suffix -= 1;
  }
  return {
    start: prefix,
    end: from.length - suffix,
    insert: to.slice(prefix, to.length - suffix),
  };
}

const isNoopSplice = (splice: TextSplice) => splice.start === splice.end && splice.insert === "";

function applySplices(base: string, splices: readonly TextSplice[]): string {
  let out = "";
  let at = 0;
  for (const splice of splices) {
    out += base.slice(at, splice.start) + splice.insert;
    at = splice.end;
  }
  return out + base.slice(at);
}

/**
 * Three-way merge of two concurrent edits of one string, each reduced to its
 * single splice against `base`. Nothing either side typed is lost:
 *
 * - disjoint edits are both applied;
 * - two insertions at one point are ordered deterministically (by text);
 * - an edit of one range that extends the other side's edit of the same range
 *   (an identical edit included) is applied once, as the longer edit: a
 *   rebase whose "theirs" already holds the start of this author's own typing
 *   (a keystroke committed but not yet rendered) must not type it twice;
 * - an insertion strictly inside the other side's deleted range survives:
 *   the deletion is split around it (the deleter's replacement text first);
 * - overlapping deletions delete their union, and both replacement texts are
 *   kept, in position order.
 */
export function mergeText3(base: string, mine: string, theirs: string): string {
  if (mine === theirs) return mine;
  if (theirs === base) return mine;
  if (mine === base) return theirs;
  const a = textSplice(base, mine);
  const b = textSplice(base, theirs);
  if (isNoopSplice(a)) return theirs;
  if (isNoopSplice(b)) return mine;
  if (a.start === b.start && a.end === b.end) {
    if (a.insert.startsWith(b.insert)) return mine;
    if (b.insert.startsWith(a.insert)) return theirs;
  }
  const [first, second] = [a, b].sort((left, right) =>
    left.start - right.start || left.end - right.end ||
    (left.insert < right.insert ? -1 : left.insert > right.insert ? 1 : 0));
  // Sorting puts an insertion at a range's start before that range, so these
  // touch without overlapping and are applied as two splices.
  if (first.end <= second.start) return applySplices(base, [first, second]);
  // Overlap. An insertion strictly inside the other's range keeps its text.
  // (`first` can not be the insertion here: it would sort before `second`.)
  if (second.start === second.end) {
    return applySplices(base, [{ start: first.start, end: first.end, insert: first.insert + second.insert }]);
  }
  const end = Math.max(first.end, second.end);
  const insert = first.insert === second.insert ? first.insert : first.insert + second.insert;
  return applySplices(base, [{ start: first.start, end, insert }]);
}

/**
 * Convert every collaborative text a legacy room still holds as a scalar into
 * a Y.Text of the same string. The JSON projection does not change.
 *
 * A writer runs this once when it joins a room, in a transaction its personal
 * history does not capture: the conversion is nobody's edit, so undoing the
 * first keystroke in a legacy box removes that keystroke instead of reverting
 * the box to its scalar (and with it every character a peer typed since).
 */
export function upgradeCollaborativeText(target: Y.Map<unknown>): boolean {
  let touched = false;
  const visit = (value: unknown, field?: string, itemOf?: string) => {
    if (isListMap(value)) {
      const items = value.get(LIST_ITEMS);
      if (!(items instanceof Y.Map)) return;
      for (const item of Array.from(items.values())) visit(item, undefined, field);
      return;
    }
    if (value instanceof Y.Map) {
      for (const [key, child] of Array.from(value.entries())) {
        if (typeof child === "string" && isCollaborativeTextField(key, itemOf)) {
          value.set(key, new Y.Text(child));
          touched = true;
        } else if (child instanceof Y.Array) {
          visit(child, key, childArrayContext(key, itemOf));
        } else {
          visit(child, key);
        }
      }
      return;
    }
    if (value instanceof Y.Array) {
      const entryOf = arrayEntryContext(itemOf);
      for (const child of value.toArray()) visit(child, undefined, entryOf);
    }
  };
  visit(target);
  return touched;
}

/** Apply the splice turning the Y.Text's current string into `next`. */
export function spliceYText(text: Y.Text, next: string): boolean {
  const current = text.toString();
  if (current === next) return false;
  const { start, end, insert } = textSplice(current, next);
  if (end > start) text.delete(start, end - start);
  if (insert) text.insert(start, insert);
  return true;
}

// ---------------------------------------------------------------------------
// Exact text positions
//
// A splice found by comparing two strings (`textSplice`) can only guess where
// an edit was made when the characters around it repeat ("aa" -> "aaa"), and
// a caret carried across a change by that guess lands on the wrong side of a
// peer's characters. Two people typing in one text box then scrambled each

// caret left the caret where it was, and the next keys went into the middle
// of what had just been typed. These helpers carry positions exactly: a
// peer's change arrives as the Y.Text's own delta, and this member's
// keystroke is written where it was typed (`CollaborativeTextEdit`).
// ---------------------------------------------------------------------------

/** One step of a Y.Text delta, reduced to characters (formatting-only
 *  changes are retains). */
export type CollaborativeTextDeltaOp = { retain: number } | { insert: string } | { delete: number };

/** A peer's change of one collaborative text, at its JSON path
 *  (e.g. ["slides", slideId, "blocks", blockId, "text"]). */
export interface CollaborativeTextChange {
  path: readonly string[];
  delta: readonly CollaborativeTextDeltaOp[];
}

/** This member's edit of one collaborative text: `base` with
 *  [start, end) replaced by `insert`, exactly where it was typed. */
export interface CollaborativeTextEdit {
  base: string;
  start: number;
  end: number;
  insert: string;
}

/** `edit` applied to its base. */
export function applyCollaborativeTextEdit(edit: CollaborativeTextEdit): string {
  return edit.base.slice(0, edit.start) + edit.insert + edit.base.slice(edit.end);
}

function validTextEdit(edit: CollaborativeTextEdit, previous: unknown, next: string): boolean {
  return edit.base === previous && ownTextEdit(edit, next);
}

/** `edit` is exactly the change from its own base to `next`. */
function ownTextEdit(edit: CollaborativeTextEdit, next: string): boolean {
  return edit.start >= 0 && edit.start <= edit.end && edit.end <= edit.base.length &&
    applyCollaborativeTextEdit(edit) === next;
}

/**
 * The text an author's edit was typed on when it is not the base the caller
 * measured the change from: a peer's characters reached the room (and moved
 * its base on) before the editor showed them, and a key typed meanwhile was
 * typed on the text without them. Null when the edit is of `previous` (or is
 * not exactly the change to `next`).
 */
function staleEditBase(edit: CollaborativeTextEdit | undefined, previous: unknown, next: string): string | null {
  return rebaseStaleTextEdits && edit && edit.base !== previous && ownTextEdit(edit, next) ? edit.base : null;
}

/**
 * `edit` (of its own base) moved onto `room`, which a peer changed from that
 * base: the peer's change is taken as one splice of the base, and the edit
 * keeps its place against the characters around it. An edit at the very
 * place a peer inserted goes before the peer's characters (the caret rule of
 * `mapIndexThroughTextDelta`). Null when the two changes overlap.
 */
function rebaseTextEdit(edit: CollaborativeTextEdit, room: string): TextSplice | null {
  const peer = textSplice(edit.base, room);
  if (isNoopSplice(peer)) return { start: edit.start, end: edit.end, insert: edit.insert };
  if (edit.end <= peer.start) return { start: edit.start, end: edit.end, insert: edit.insert };
  if (edit.start >= peer.end && !(edit.start === peer.start)) {
    const shift = peer.insert.length - (peer.end - peer.start);
    return { start: edit.start + shift, end: edit.end + shift, insert: edit.insert };
  }
  return null;
}

/**
 * The edits a publication writes exactly, by the room text they edit:
 * several notes of one text (keys typed before one publication) as one edit
 * of the first one's base, when each is of the text the one before left;
 * a text whose notes do not chain is left to the string comparison.
 */
export function collaborativeTextEditsFor(
  root: Y.Map<unknown>,
  noted: ReadonlyArray<CollaborativeTextEdit & { path: readonly string[] }>,
): Map<Y.Text, CollaborativeTextEdit> | undefined {
  const chains = new Map<Y.Text, { base: string; text: string } | null>();
  for (const { path, ...edit } of noted) {
    const text = collaborativeTextAt(root, path);
    if (!text) continue;
    if (!ownTextEdit(edit, applyCollaborativeTextEdit(edit))) continue;
    const chain = chains.get(text);
    if (chain === null) continue;
    if (chain === undefined) chains.set(text, { base: edit.base, text: applyCollaborativeTextEdit(edit) });
    else if (chain.text === edit.base) chain.text = applyCollaborativeTextEdit(edit);
    else chains.set(text, null);
  }
  let out: Map<Y.Text, CollaborativeTextEdit> | undefined;
  for (const [text, chain] of chains) {
    if (!chain) continue;
    const { start, end, insert } = textSplice(chain.base, chain.text);
    (out ??= new Map()).set(text, { base: chain.base, start, end, insert });
  }
  return out;
}

/** `base` with `delta` applied; null when the delta is not of `base`. */
export function applyCollaborativeTextDelta(
  base: string,
  delta: readonly CollaborativeTextDeltaOp[],
): string | null {
  let out = "";
  let at = 0;
  for (const op of delta) {
    if ("retain" in op) {
      if (at + op.retain > base.length) return null;
      out += base.slice(at, at + op.retain);
      at += op.retain;
    } else if ("insert" in op) {
      out += op.insert;
    } else {
      if (at + op.delete > base.length) return null;
      at += op.delete;
    }
  }
  return out + base.slice(at);
}

/**
 * Where `index` of the text before `delta` stands after it. A peer's
 * insertion exactly at `index` goes after it (`stickRight` false: a caret
 * stays behind this member's own last character, so each person's words
 * stay together) or before it (`stickRight`: the start of a selection does
 * not take in what a peer typed at its edge). Inside a deleted range the
 * index moves to where the range was.
 */
export function mapIndexThroughTextDelta(
  index: number,
  delta: readonly CollaborativeTextDeltaOp[],
  stickRight = false,
): number {
  let old = 0;
  let mapped = index;
  for (const op of delta) {
    if (old > index) break;
    if ("retain" in op) {
      old += op.retain;
    } else if ("insert" in op) {
      if (old < index || (old === index && stickRight)) mapped += op.insert.length;
    } else {
      const end = old + op.delete;
      mapped -= Math.max(0, Math.min(end, index) - old);
      old = end;
    }
  }
  return mapped;
}

/**
 * The collaborative text changes a room transaction made, read from its
 * events (they are only readable while the events are dispatched). A text
 * whose change is not a plain character delta is left out.
 */
export function collaborativeTextChanges(events: readonly Y.YEvent<never>[]): CollaborativeTextChange[] {
  const out: CollaborativeTextChange[] = [];
  for (const event of events) {
    const target: unknown = event.target;
    if (!(target instanceof Y.Text)) continue;
    // `event.path` runs from the observed root to the text itself.
    const path = (event.path as Array<string | number>).filter((key) => key !== LIST_ITEMS).map(String);
    if (path[path.length - 1] !== COLLABORATIVE_TEXT_KEY) continue;
    const delta: CollaborativeTextDeltaOp[] = [];
    let plain = true;
    for (const op of event.delta as Array<{ insert?: unknown; retain?: number; delete?: number }>) {
      if (typeof op.retain === "number") delta.push({ retain: op.retain });
      else if (typeof op.delete === "number") delta.push({ delete: op.delete });
      else if (typeof op.insert === "string") delta.push({ insert: op.insert });
      else { plain = false; break; }
    }
    if (plain) out.push({ path, delta });
  }
  return out;
}

/** The room's collaborative text at a JSON `path`, or null. */
export function collaborativeTextAt(root: Y.Map<unknown>, path: readonly string[]): Y.Text | null {
  const node = sharedNodeAt(root, path);
  return node instanceof Y.Text ? node : null;
}

/** The room's shared value at a JSON `path` (keyed lists by id, plain
 *  arrays -- a table's rows and cells -- by index), or undefined. */
function sharedNodeAt(root: Y.Map<unknown>, path: readonly string[]): unknown {
  let node: unknown = root;
  for (const key of path) {
    if (node instanceof Y.Array) {
      const index = /^(0|[1-9]\d*)$/.test(key) ? Number(key) : -1;
      if (index < 0 || index >= node.length) return undefined;
      node = node.get(index);
      continue;
    }
    if (!(node instanceof Y.Map)) return undefined;
    const map = isListMap(node) ? node.get(LIST_ITEMS) : node;
    if (!(map instanceof Y.Map)) return undefined;
    node = map.get(key);
  }
  return node;
}

/**
 * The text container at `path` -- a slide object (`[…, "blocks", id]`) or a
 * table cell (`[…, "blocks", id, "rows", row, column]`) -- and the context
 * its map is written in, or null.
 */
function textContainerAt(root: Y.Map<unknown>, path: readonly string[]): { map: Y.Map<unknown>; itemOf: string } | null {
  const n = path.length;
  const itemOf = n >= 2 && path[n - 2] === "blocks" ? "blocks"
    : n >= 5 && path[n - 3] === "rows" && path[n - 5] === "blocks" ? TABLE_CELL
    : null;
  if (!itemOf) return null;
  const node = sharedNodeAt(root, path);
  return node instanceof Y.Map && !isListMap(node) ? { map: node, itemOf } : null;
}

/** Whether two text containers hold the same words formatted alike: per
 * character (runs split differently but styled the same are alike), else
 * as the same values. */
function sameFormattedWords(left: Record<string, JsonValue>, right: Record<string, JsonValue>): boolean {
  if (!jsonEquals(left[COLLABORATIVE_TEXT_KEY] ?? null, right[COLLABORATIVE_TEXT_KEY] ?? null)) return false;
  const a = encodeFormatting(left);
  const b = encodeFormatting(right);
  if (a && b) {
    return a.kind === b.kind && a.units.length === b.units.length &&
      a.units.every((unit, index) => stableJson(unit as JsonValue) === stableJson(b.units[index] as JsonValue)) &&
      stableJson((a.tail ?? null) as JsonValue) === stableJson((b.tail ?? null) as JsonValue);
  }
  return !a && !b && jsonEquals(left[RUNS_KEY] ?? null, right[RUNS_KEY] ?? null) &&
    jsonEquals(left[PARAGRAPHS_KEY] ?? null, right[PARAGRAPHS_KEY] ?? null);
}















export function writeCollaborativeTextContainer(
  root: Y.Map<unknown>,
  path: readonly string[],
  previous: Record<string, JsonValue>,
  next: Record<string, JsonValue>,
  edit?: CollaborativeTextEdit,
): Record<string, JsonValue> | null {
  const container = textContainerAt(root, path);
  if (!container) return null;
  const { map, itemOf } = container;
  // The words and their formatting must be the ones typed in (the edit and
  // the formatting delta are measured from them); any other field is written
  // as a delta, as a publication writes it.
  if (!sameFormattedWords(yToJson(map) as Record<string, JsonValue>, previous)) return null;
  const text = map.get(COLLABORATIVE_TEXT_KEY);
  const exact = text instanceof Y.Text && edit && typeof next.text === "string" &&
    validTextEdit(edit, previous.text, next.text) ? new Map([[text, edit]]) : null;
  const outer = exactTextEdits;
  const outerRebase = rebaseStaleTextEdits;
  exactTextEdits = exact;
  rebaseStaleTextEdits = false;
  try {
    applyMapDelta(map, next, previous, itemOf);
  } finally {
    exactTextEdits = outer;
    rebaseStaleTextEdits = outerRebase;
  }
  return yToJson(map) as Record<string, JsonValue>;
}

/** The room's text container (a slide object or a table cell) at `path`, as
 * JSON, or null. */
export function collaborativeTextContainerAt(root: Y.Map<unknown>, path: readonly string[]): Record<string, JsonValue> | null {
  const container = textContainerAt(root, path);
  if (!container) return null;
  const value = yToJson(container.map);
  return isPlainObject(value) ? value : null;
}

/**
 * An edit typed into `base`, a text the room held, carried over the room's
 * changes of it since (`deltas`, oldest first: a collaborator's characters,
 * and this member's own keys written meanwhile): the range of the room's
 * text it now replaces, and the delta that turns the author's text after the
 * edit into the room's after it (for the editor to catch up, its caret moved
 * by it). The edit's insertion goes behind the character it was typed after,
 * before a collaborator's characters inserted at the same place (a caret
 * stays behind its owner's last letter, `mapIndexThroughTextDelta`); a
 * replaced range does not take in what a collaborator typed at its edges.
 * Null when a delta is not of the text it follows.
 */
export function carryTextEditThroughDeltas(
  base: string,
  edit: { start: number; end: number; insert: string },
  deltas: readonly (readonly CollaborativeTextDeltaOp[])[],
): { start: number; end: number; viewToRoom: CollaborativeTextDeltaOp[] } | null {
  if (edit.start < 0 || edit.start > edit.end || edit.end > base.length) return null;
  // Each character of the room's text by where it came from: the base's
  // index, or a collaborator's (negative, numbered).
  let tokens: number[] = Array.from({ length: base.length }, (_, index) => index);
  let chars: string[] = base.split("");
  let peer = 0;
  let start = edit.start;
  let end = edit.end;
  const collapsed = edit.start === edit.end;
  for (const delta of deltas) {
    const nextTokens: number[] = [];
    const nextChars: string[] = [];
    let at = 0;
    for (const op of delta) {
      if ("retain" in op) {
        if (at + op.retain > tokens.length) return null;
        for (let k = 0; k < op.retain; k += 1) { nextTokens.push(tokens[at]!); nextChars.push(chars[at]!); at += 1; }
      } else if ("insert" in op) {
        for (const char of op.insert.split("")) { peer += 1; nextTokens.push(-peer); nextChars.push(char); }
      } else {
        if (at + op.delete > tokens.length) return null;
        at += op.delete;
      }
    }
    for (; at < tokens.length; at += 1) { nextTokens.push(tokens[at]!); nextChars.push(chars[at]!); }
    start = mapIndexThroughTextDelta(start, delta, !collapsed);
    end = collapsed ? start : Math.max(start, mapIndexThroughTextDelta(end, delta, false));
    tokens = nextTokens;
    chars = nextChars;
  }
  // The author's text after the edit, and the room's, by the same tokens.
  const OWN = 1 << 30;
  const ownTokens = Array.from({ length: edit.insert.length }, (_, index) => OWN + index);
  const view = [
    ...Array.from({ length: edit.start }, (_, index) => index),
    ...ownTokens,
    ...Array.from({ length: base.length - edit.end }, (_, index) => edit.end + index),
  ];
  const room = [...tokens.slice(0, start), ...ownTokens, ...tokens.slice(end)];
  const roomChars = [...chars.slice(0, start), ...edit.insert.split(""), ...chars.slice(end)];
  const inView = new Set(view);
  const inRoom = new Set(room);
  const viewToRoom: CollaborativeTextDeltaOp[] = [];
  const push = (op: CollaborativeTextDeltaOp) => {
    const last = viewToRoom[viewToRoom.length - 1];
    if (last && "retain" in last && "retain" in op) last.retain += op.retain;
    else if (last && "delete" in last && "delete" in op) last.delete += op.delete;
    else if (last && "insert" in last && "insert" in op) last.insert += op.insert;
    else viewToRoom.push({ ...op });
  };
  let i = 0;
  let j = 0;
  while (i < view.length || j < room.length) {
    if (i < view.length && j < room.length && view[i] === room[j]) { push({ retain: 1 }); i += 1; j += 1; }
    else if (i < view.length && !inRoom.has(view[i]!)) { push({ delete: 1 }); i += 1; }
    else if (j < room.length && !inView.has(room[j]!)) { push({ insert: roomChars[j]! }); j += 1; }
    else return null;
  }
  const last = viewToRoom[viewToRoom.length - 1];
  if (last && "retain" in last) viewToRoom.pop();
  return { start, end, viewToRoom };
}

/** Whether two text containers (slide objects, table cells) hold the same
 * words formatted alike (`writeCollaborativeTextContainer`'s test). */
export function textContainersAlike(left: JsonValue | undefined, right: JsonValue | undefined): boolean {
  return isPlainObject(left) && isPlainObject(right) && sameFormattedWords(left, right);
}

/** The text container's own words and formatting fields. */
const TEXT_CONTAINER_WORD_KEYS = [COLLABORATIVE_TEXT_KEY, "runs", "paragraphs"] as const;

/**
 * `local` with its words and their formatting (`text`, `runs`,
 * `paragraphs`) taken from `room`, every other field as it is.
 */
export function withTextContainerWords(
  local: Record<string, JsonValue>,
  room: Record<string, JsonValue>,
): Record<string, JsonValue> {
  const out: Record<string, JsonValue> = { ...local };
  for (const key of TEXT_CONTAINER_WORD_KEYS) {
    if (hasJsonProperty(room, key) && room[key] !== undefined) out[key] = room[key]!;
    else delete out[key];
  }
  return out;
}

/** The value at a JSON `path` of a plain document (keyed list entries by id,
 * plain arrays by index), or undefined. */
export function jsonValueAt(doc: JsonValue | undefined, path: readonly string[]): JsonValue | undefined {
  let node: JsonValue | undefined = doc;
  for (const key of path) {
    if (Array.isArray(node)) {
      let index = node.findIndex((item) => isPlainObject(item) && item.id === key);
      if (index < 0 && /^(0|[1-9]\d*)$/.test(key) && !keyedListIds(node)) index = Number(key);
      node = index >= 0 ? node[index] : undefined;
    } else if (isPlainObject(node) && hasJsonProperty(node, key)) {
      node = node[key];
    } else {
      return undefined;
    }
  }
  return node;
}

/** A formatting's attributes per unit written back as the container's
 * `runs` or `paragraphs` (the JSON `decodeFormatting` reads from a text). */
function formattingJson(kind: FormatKind, text: string, units: readonly Attributes[], tail: Attributes | null): JsonValue {
  if (kind === "runs") {
    const runs: Array<{ text: string; style?: Attributes }> = [];
    for (let index = 0; index < text.length; index += 1) appendRun(runs, text[index]!, runStyleOf(units[index]));
    return runs as JsonValue;
  }
  type Paragraph = { text: string; runs: Array<{ text: string; style?: Attributes }>; attrs: Attributes | null };
  const paragraphs: Paragraph[] = [{ text: "", runs: [], attrs: null }];
  for (let index = 0; index < text.length; index += 1) {
    const current = paragraphs[paragraphs.length - 1]!;
    if (text[index] === "\n") {
      current.attrs = units[index] ?? {};
      paragraphs.push({ text: "", runs: [], attrs: null });
      continue;
    }
    current.text += text[index];
    appendRun(current.runs, text[index]!, runStyleOf(units[index]));
  }
  paragraphs[paragraphs.length - 1]!.attrs = tail ?? {};
  return paragraphs.map((paragraph) => {
    const out: Attributes = { text: paragraph.text };
    let hasRuns = false;
    for (const [key, value] of Object.entries(paragraph.attrs ?? {})) {
      if (!key.startsWith(PARAGRAPH_ATTRIBUTE)) continue;
      if (key === PARAGRAPH_RUNS_FLAG) hasRuns = value === true;
      else out[key.slice(PARAGRAPH_ATTRIBUTE.length)] = value;
    }
    if (hasRuns) out[RUNS_KEY] = paragraph.runs as JsonValue;
    return out;
  }) as JsonValue;
}

/**
 * The room's text container (`room`, as the room holds it) with this
 * member's edit of it written where the edit now stands in the room's words
 * (`at`: [start, end) of the room's text): the edit typed into the words
 * `previous` held (`edit`, of `previous.text`, giving `next`), while the
 * room's had moved on with a collaborator's characters. The typed characters
 * keep the formatting `next` gives them, the room's keep theirs; any other
 * field the edit changed is changed. Null when the formatting of the two is
 * not of one kind (the caller publishes the edit as before).
 */
export function textContainerWithEditAt(
  room: Record<string, JsonValue>,
  previous: Record<string, JsonValue>,
  next: Record<string, JsonValue>,
  edit: CollaborativeTextEdit,
  at: { start: number; end: number },
): Record<string, JsonValue> | null {
  const roomText = room[COLLABORATIVE_TEXT_KEY];
  if (typeof roomText !== "string" || typeof next.text !== "string" || typeof previous.text !== "string") return null;
  if (!validTextEdit(edit, previous.text, next.text)) return null;
  if (at.start < 0 || at.start > at.end || at.end > roomText.length) return null;
  const text = roomText.slice(0, at.start) + edit.insert + roomText.slice(at.end);
  const out: Record<string, JsonValue> = { ...room };
  // Other fields the author changed with the key (a box growing with its words).
  for (const key of new Set([...Object.keys(previous), ...Object.keys(next)])) {
    if ((TEXT_CONTAINER_WORD_KEYS as readonly string[]).includes(key)) continue;
    const was = hasJsonProperty(previous, key) ? previous[key] : undefined;
    const is = hasJsonProperty(next, key) ? next[key] : undefined;
    if (was !== undefined && is !== undefined && jsonEquals(was, is)) continue;
    if (is === undefined) delete out[key];
    else out[key] = is;
  }
  out[COLLABORATIVE_TEXT_KEY] = text;
  const roomFormat = encodeFormatting(room);
  const nextFormat = encodeFormatting(next);
  const hasFormatting = (value: Record<string, JsonValue>) =>
    (hasJsonProperty(value, RUNS_KEY) && value[RUNS_KEY] !== undefined) ||
    (hasJsonProperty(value, PARAGRAPHS_KEY) && value[PARAGRAPHS_KEY] !== undefined);
  if (!roomFormat && !nextFormat) {
    // Plain words on both sides: nothing but the text to write.
    return hasFormatting(room) || hasFormatting(next) ? null : out;
  }
  if (!roomFormat || !nextFormat || roomFormat.kind !== nextFormat.kind) return null;
  const typed = nextFormat.units.slice(edit.start, edit.start + edit.insert.length);
  if (typed.length !== edit.insert.length) return null;
  const units = [...roomFormat.units.slice(0, at.start), ...typed, ...roomFormat.units.slice(at.end)];
  const previousFormat = encodeFormatting(previous);
  const tailChanged = previousFormat?.kind === nextFormat.kind &&
    stableJson((previousFormat.tail ?? null) as JsonValue) !== stableJson((nextFormat.tail ?? null) as JsonValue);
  const tail = tailChanged ? nextFormat.tail : roomFormat.tail;
  const owned = ownedKey(roomFormat.kind);
  out[owned] = formattingJson(roomFormat.kind, text, units, tail);
  // Flat runs a paragraphs body also carries (as a value) spell the words too.
  if (roomFormat.kind === "paragraphs" && Array.isArray(room[RUNS_KEY])) {
    const moved = spliceRuns(room[RUNS_KEY] as JsonValue[], roomText, { start: at.start, end: at.end, insert: edit.insert });
    if (!moved) return null;
    out[RUNS_KEY] = moved;
  }
  return out;
}

/**
 * `doc` with the value at a JSON `path` (keyed list entries by id, plain
 * arrays by index, as the room's paths name them) replaced by `value`, the
 * objects along the path copied. `doc` itself when the path is not there.
 */
export function withJsonValueAt(doc: JsonValue, path: readonly string[], value: JsonValue): JsonValue {
  if (path.length === 0) return value;
  const [key, ...rest] = path;
  if (Array.isArray(doc)) {
    let index = doc.findIndex((item) => isPlainObject(item) && item.id === key);
    if (index < 0 && /^(0|[1-9]\d*)$/.test(key!) && !keyedListIds(doc)) index = Number(key);
    if (index < 0 || index >= doc.length) return doc;
    const child = withJsonValueAt(doc[index]!, rest, value);
    if (child === doc[index]) return doc;
    const out = doc.slice();
    out[index] = child;
    return out;
  }
  if (isPlainObject(doc) && hasJsonProperty(doc, key!)) {
    const child = withJsonValueAt(doc[key!]!, rest, value);
    return child === doc[key!] ? doc : { ...doc, [key!]: child };
  }
  return doc;
}

/**
 * A place in a room text that stays with its characters (a Yjs relative
 * position): `index` of the text as it stands now, bound to the character
 * before it (`assoc` < 0: a caret stays behind its owner's last letter when
 * a peer types at it) or after it.
 */
export function collaborativeTextAnchor(text: Y.Text, index: number, assoc: -1 | 0 = -1): Y.RelativePosition {
  return Y.createRelativePositionFromTypeIndex(text, Math.max(0, Math.min(index, text.length)), assoc);
}

/** Where `anchor` stands in `text` now, or null (another text's, or gone). */
export function resolveCollaborativeTextAnchor(text: Y.Text, anchor: Y.RelativePosition): number | null {
  if (!text.doc) return null;
  const at = Y.createAbsolutePositionFromRelativePosition(anchor, text.doc);
  return at && at.type === text ? at.index : null;
}

/** The edits `applyJsonDelta` writes exactly, by the room text they edit. */
let exactTextEdits: ReadonlyMap<Y.Text, CollaborativeTextEdit> | null = null;
/** Whether an exact edit typed on a text the base had moved past is rebased
 *  onto the room (`staleEditBase`); set for one `applyJsonDelta`. */
let rebaseStaleTextEdits = false;

function scalarEquals(a: unknown, b: unknown): boolean {
  return a === b || (Number.isNaN(a) && Number.isNaN(b));
}

/** Deep-equality for the JSON subset — used to skip unchanged writes. */
export function jsonEquals(a: JsonValue, b: JsonValue): boolean {
  if (scalarEquals(a, b)) return true;
  if (Array.isArray(a) && Array.isArray(b)) {
    return (
      a.length === b.length && a.every((item, i) => jsonEquals(item, b[i]))
    );
  }
  if (isPlainObject(a) && isPlainObject(b)) {
    const aKeys = Object.keys(a).filter((key) => hasJsonProperty(a, key));
    const bKeys = Object.keys(b).filter((key) => hasJsonProperty(b, key));
    if (aKeys.length !== bKeys.length) return false;
    return aKeys.every(
      (key) =>
        hasJsonProperty(b, key) && jsonEquals(a[key], (b as Record<string, JsonValue>)[key]),
    );
  }
  return false;
}

/**
 * `field` is the key `value` is stored under; `itemOf` names the keyed list
 * `value` is an entry of (it decides which child keys are collaborative text).
 */
function toYValue(value: JsonValue, field?: string, itemOf?: string): unknown {
  // An EMPTY array becomes a list-map too, not a plain Y.Array.
  //
  // The representation must not flip between empty and non-empty. While a
  // list was a plain array, two members adding their FIRST item concurrently
  // each replaced the whole array, and one of them lost their work — on a
  // fresh PDF, that is the first annotation two people draw. `isKeyedList`
  // has always documented this hazard and deferred to a `preferKeyed` the
  // caller was supposed to supply; no such thing was ever implemented, so
  
  //
  // Seeding it as a list-map costs nothing: `yToJson` reads an empty one back
  // as `[]`, and a key that turns out to hold a NON-keyed array is replaced
  // wholesale on its first real value, exactly as before.
  const listIds = keyedListIds(value, field);
  if (listIds || (Array.isArray(value) && value.length === 0)) {
    const list = new Y.Map<unknown>();
    list.set(LIST_MARKER, true);
    const order = new Y.Array<string>();
    order.push(listIds ?? []);
    list.set(LIST_ORDER, order);
    const items = new Y.Map<unknown>();
    for (const [index, item] of (
      value as Record<string, JsonValue>[]
    ).entries()) {
      const id = listIds?.[index];
      if (id) items.set(id, toYValue(item, undefined, field));
    }
    list.set(LIST_ITEMS, items);
    return list;
  }
  if (Array.isArray(value)) {
    const arr = new Y.Array<unknown>();
    const entryOf = arrayEntryContext(itemOf);
    arr.push(value.map((item) => toYValue(item, undefined, entryOf)));
    return arr;
  }
  if (isPlainObject(value)) {
    const map = new Y.Map<unknown>();
    const formatting = itemOf !== undefined && COLLABORATIVE_TEXT_LISTS.has(itemOf) ? encodeFormatting(value) : null;
    for (const [key, entry] of Object.entries(value)) {
      if (entry === undefined) continue;
      if (formatting && key === ownedKey(formatting.kind)) continue;
      if (formatting && key === COLLABORATIVE_TEXT_KEY) {
        // The words with their formatting (collaborative formatting).
        const text = new Y.Text();
        const words = entry as string;
        const keys = formatting.units.map((attrs) => stableJson(attrs));
        let at = 0;
        while (at < words.length) {
          let end = at + 1;
          while (end < words.length && keys[end] === keys[at]) end += 1;
          text.insert(at, words.slice(at, end), { ...(formatting.units[at] ?? {}) });
          at = end;
        }
        map.set(key, text);
        continue;
      }
      map.set(key, typeof entry === "string" && isCollaborativeTextField(key, itemOf)
        ? new Y.Text(entry)
        : toYValue(entry, key, childArrayContext(key, itemOf)));
    }
    if (formatting) {
      map.set(FORMAT_MARKER, formatting.kind);
      if (formatting.tail) map.set(FORMAT_TAIL, new Y.Map<unknown>(Object.entries(formatting.tail)));
    }
    return map;
  }
  return value;
}

function isListMap(value: unknown): value is Y.Map<unknown> {
  return value instanceof Y.Map && value.get(LIST_MARKER) === true;
}

/** Read a Yjs subtree back as plain JSON. */
export function yToJson(value: unknown): JsonValue {
  return readYJson(value, yToJson);
}

/** `runs` with `text` of theirs replaced as the splice says; a typed word takes
 * the style of the run it follows. Null when they do not partition that text. */
function spliceRuns(runs: JsonValue[], before: string, { start, end, insert }: TextSplice): JsonValue[] | null {
  if (!runs.every((run) => isPlainObject(run) && typeof run.text === "string")) return null;
  if ((runs as Array<{ text: string }>).map((run) => run.text).join("") !== before) return null;
  const out: JsonValue[] = [];
  let at = 0;
  let inserted = false;
  for (const run of runs as Array<Record<string, JsonValue> & { text: string }>) {
    const from = at;
    at += run.text.length;
    let words = run.text.slice(0, Math.max(0, start - from));
    if (!inserted && start <= at) {
      words += insert;
      inserted = true;
    }
    words += run.text.slice(Math.max(0, end - from));
    if (words.length > 0) out.push({ ...run, text: words });
  }
  if (!inserted && insert.length > 0) out.push({ text: insert });
  return out;
}

/** The paragraphs of a text they no longer spell, moved to spell it: only those
 * the change touches, keeping their settings and the words' styles. */
function alignParagraphs(paragraphs: JsonValue[], text: string): JsonValue[] | null {
  if (!paragraphs.every((paragraph) => isPlainObject(paragraph) && typeof paragraph.text === "string")) return null;
  const list = paragraphs as Array<Record<string, JsonValue> & { text: string }>;
  const joined = list.map((paragraph) => paragraph.text).join("\n");
  if (joined === text) return null;
  const change = textSplice(joined, text);
  let first = -1;
  let last = list.length - 1;
  let firstAt = 0;
  let lastAt = 0;
  let at = 0;
  for (let index = 0; index < list.length; index += 1) {
    const ends = at + list[index]!.text.length;
    if (first < 0 && change.start <= ends) {
      first = index;
      firstAt = at;
    }
    if (first >= 0 && change.end <= ends) {
      last = index;
      lastAt = at;
      break;
    }
    at = ends + 1;
  }
  const touched = list.slice(first, last + 1);
  const words = touched.map((paragraph) => paragraph.text).join("\n");
  const local = { start: change.start - firstAt, end: change.end - firstAt, insert: change.insert };
  const merged: Record<string, JsonValue> = {
    ...touched[0]!,
    text: list[first]!.text.slice(0, change.start - firstAt) + change.insert + list[last]!.text.slice(change.end - lastAt),
  };
  delete merged.runs;
  if (touched.every((paragraph) => Array.isArray(paragraph.runs))) {
    const runs = touched.flatMap((paragraph, index) => {
      const own = paragraph.runs as JsonValue[];
      if (index === touched.length - 1) return own;
      const style = isPlainObject(own[own.length - 1]) ? (own[own.length - 1] as Record<string, JsonValue>).style : undefined;
      return [...own, style === undefined ? { text: "\n" } : { text: "\n", style }];
    });
    const moved = spliceRuns(runs, words, local);
    if (moved) merged.runs = moved;
  }
  return [...list.slice(0, first), merged, ...list.slice(last + 1)];
}











function alignValueFormatting(block: Record<string, JsonValue>): void {
  const text = block.text;
  if (typeof text !== "string") return;
  const paragraphs = block[PARAGRAPHS_KEY];
  if (Array.isArray(paragraphs) && paragraphs.length > 0) {
    const aligned = alignParagraphs(paragraphs, text);
    if (aligned) block[PARAGRAPHS_KEY] = aligned;
  }
  const runs = block[RUNS_KEY];
  if (Array.isArray(runs) && runs.every((run) => isPlainObject(run) && typeof run.text === "string")) {
    const before = (runs as Array<{ text: string }>).map((run) => run.text).join("");
    if (before !== text) {
      const moved = spliceRuns(runs, before, textSplice(before, text));
      if (moved) block[RUNS_KEY] = moved;
    }
  }
}

function readYJson(value: unknown, child: (value: unknown) => JsonValue): JsonValue {
  if (isListMap(value)) {
    const order = value.get(LIST_ORDER);
    const items = value.get(LIST_ITEMS);
    if (!(order instanceof Y.Array) || !(items instanceof Y.Map)) return [];
    const out: JsonValue[] = [];
    const seen = new Set<string>();
    for (const id of order.toArray() as string[]) {
      if (seen.has(id)) continue;
      seen.add(id);
      if (items.has(id)) out.push(child(items.get(id)));
    }
    // Entries added concurrently by a peer whose order write lost the race
    // must still appear — appended rather than dropped.
    for (const id of items.keys()) {
      if (!seen.has(id)) out.push(child(items.get(id)));
    }
    return out;
  }
  if (value instanceof Y.Map) {
    const out: Record<string, JsonValue> = {};
    const kind = formatKindOf(value);
    for (const [key, entry] of value.entries()) {
      if (key === FORMAT_MARKER || key === FORMAT_TAIL) continue;
      out[key] = child(entry);
    }
    // A block whose formatting is its text's reads it back from there;
    // formatting a writer of the older model set as a value is read as written.
    const text = value.get(COLLABORATIVE_TEXT_KEY);
    if (kind && !value.has(ownedKey(kind)) && text instanceof Y.Text) {
      out[ownedKey(kind)] = decodeFormatting(kind, text, value.get(FORMAT_TAIL));
    }
    if (text instanceof Y.Text) alignValueFormatting(out);
    return out;
  }
  if (value instanceof Y.Array) {
    return value.toArray().map((entry) => child(entry));
  }
  if (value instanceof Y.Text) return value.toString();
  return (value ?? null) as JsonValue;
}

/**
 * Each shared type's JSON as last read, until a transaction changes it or
 * anything below it. Reading a whole page document is then the cost of the
 * path an edit changed, not of the document: a keystroke in one text box of
 * a 24-slide deck read every slide again, several times per key.
 */
const sharedJsonOf = new WeakMap<Y.AbstractType<unknown>, JsonValue>();
const sharedJsonDocs = new WeakSet<Y.Doc>();
/** Every object a shared read returned (see {@link isSharedJsonRead}). */
const sharedJsonReads = new WeakSet<object>();

/**
 * Whether `value` is an object {@link yToJsonShared} returned. A shared read
 * is read-only (the next read of an unchanged subtree is the same object), so
 * what is derived from it -- its canonical text, a copy taken to freeze it --
 * stands for as long as the object does, exactly as for a frozen object.
 */
export function isSharedJsonRead(value: unknown): boolean {
  return value !== null && typeof value === "object" && sharedJsonReads.has(value);
}

function followSharedJson(doc: Y.Doc): void {
  if (sharedJsonDocs.has(doc)) return;
  sharedJsonDocs.add(doc);
  // Before the transaction's observers run: they read the document.
  doc.on("beforeObserverCalls", (transaction: Y.Transaction) => {
    for (const type of transaction.changed.keys()) {
      for (
        let current: Y.AbstractType<unknown> | null = type as unknown as Y.AbstractType<unknown>;
        current;
        current = (current._item?.parent as Y.AbstractType<unknown> | undefined) ?? null
      ) {
        sharedJsonOf.delete(current);
      }
    }
  });
}

/**
 * {@link yToJson}, sharing every unchanged subtree with the previous read of
 * the same document: two reads with no change between them are the same
 * object, and after an edit only the changed path is new. The result is
 * READ-ONLY -- a caller that mutates it corrupts the next read.
 */
export function yToJsonShared(value: unknown): JsonValue {
  if (!(value instanceof Y.AbstractType) || !value.doc) return readYJson(value, yToJsonShared);
  followSharedJson(value.doc);
  const cached = sharedJsonOf.get(value);
  if (cached !== undefined) return cached;
  const json = readYJson(value, yToJsonShared);
  // Under test a caller that writes into a shared read fails loudly.
  if (json !== null && typeof json === "object") {
    if (process.env.NODE_ENV === "test") Object.freeze(json);
    sharedJsonReads.add(json);
  }
  sharedJsonOf.set(value, json);
  return json;
}

/** Session-local history over the actual shared tree. Remote transactions and
 * initial hydration must use a different origin from the caller's local edits.
 * Replaying the Yjs operation keeps peer changes instead of replacing their
 * current document with an older editor snapshot. */
export function createJsonCollaborationHistory(
  root: Y.Map<unknown>,
  options: {
    localOrigin: object | symbol;
    canWrite: () => boolean;
    /** Apply technical updates normally, without treating them as user input
     * or as peer edits that protect creations/order or split a capture. */
    maintenanceOrigins?: ReadonlySet<unknown>;
  },
) {
  // Y.Text counts as a container: a peer typing into a text this member
  // created (or upgraded from a legacy scalar) must keep that text alive when
  // the creation is undone, exactly like a peer-edited object.
  type Container = Y.Map<unknown> | Y.Array<unknown> | Y.Text;
  type StackItem = Y.UndoManager["undoStack"][number];
  const doc = root.doc;
  if (!doc) throw new Error("JSON history requires an attached collaboration document.");
  const maintenanceOrigins = new Set(options.maintenanceOrigins);
  if (maintenanceOrigins.has(options.localOrigin)) throw new Error("Maintenance origin cannot also be the local author origin.");
  const peerTouched = new WeakSet<Container>();
  const listeners = new Set<() => void>();
  const peerOrderRevision = new WeakMap<Y.Array<unknown>, number>();
  const orderRevisions = new WeakMap<StackItem, Map<Y.Array<unknown>, number>>();
  let replaying: { stack: StackItem[]; entries: StackItem[] } | null = null;
  const isContainer = (value: unknown): value is Container =>
    value instanceof Y.Map || value instanceof Y.Array || value instanceof Y.Text;
  const isListOrder = (value: unknown): value is Y.Array<unknown> => {
    if (!(value instanceof Y.Array)) return false;
    const list = value._item?.parent;
    return isListMap(list) && list.get(LIST_ORDER) === value;
  };
  const isPreservedCreation = (type: Container, entry: StackItem): boolean =>
    Boolean(type._item && peerTouched.has(type) && Y.isDeleted(entry.insertions, type._item.id));
  const preserveCreation = (type: Container): boolean => {
    // UndoManager may skip an obsolete stack item and pop the next one in the
    // same transaction. Its remaining stack length identifies the current item.
    const entry = replaying?.entries[replaying.stack.length];
    return Boolean(entry && isPreservedCreation(type, entry));
  };
  const manager = new Y.UndoManager(root, {
    trackedOrigins: new Set([options.localOrigin]),
    captureTransaction: (transaction) => !maintenanceOrigins.has(transaction.origin),
    captureTimeout: 500,
    deleteFilter(item) {
      // Undoing creation must not delete a container now used by a peer, nor
      // strip its original attributes one by one while retaining peer fields.
      if (item.content.getContent().some((value) => isContainer(value) && preserveCreation(value))) return false;
      let parent: unknown = item.parent;
      while (isContainer(parent)) {
        if (preserveCreation(parent)) return false;
        parent = parent._item?.parent;
      }
      // Preserve the position of that retained object as well as its contents.
      if (item.parent instanceof Y.Array) {
        const list = item.parent._item?.parent;
        if (isListMap(list) && list.get(LIST_ORDER) === item.parent) {
          const items = list.get(LIST_ITEMS);
          if (items instanceof Y.Map && item.content.getContent().some((id) => {
            const value = typeof id === "string" ? items.get(id) : null;
            return isContainer(value) && preserveCreation(value);
          })) return false;
        }
      }
      return true;
    },
  });
  const captureOrderRevisions = ({ stackItem, changedParentTypes }: {
    stackItem: StackItem;
    changedParentTypes: Y.Transaction["changedParentTypes"];
  }) => {
    const revisions = orderRevisions.get(stackItem) ?? new Map<Y.Array<unknown>, number>();
    for (const type of changedParentTypes.keys()) {
      if (isListOrder(type)) revisions.set(type, peerOrderRevision.get(type) ?? 0);
    }
    orderRevisions.set(stackItem, revisions);
  };
  manager.on("stack-item-added", captureOrderRevisions);
  manager.on("stack-item-updated", captureOrderRevisions);
  const observePeer = (transaction: Y.Transaction) => {
    if (transaction.origin === options.localOrigin || transaction.origin === manager ||
        maintenanceOrigins.has(transaction.origin)) return;
    // Do not merge local history across an intervening peer edit. Otherwise
    // one capture interval could treat a peer's new order as our own baseline.
    for (const type of transaction.changedParentTypes.keys()) {
      if (isContainer(type)) {
        peerTouched.add(type);
        if (type === root) manager.stopCapturing();
      }
      if (isListOrder(type)) peerOrderRevision.set(type, (peerOrderRevision.get(type) ?? 0) + 1);
    }
  };
  doc.on("afterTransaction", observePeer);
  let disposed = false;
  let historyMutationDepth = 0;
  const mutateHistory = <T>(operation: () => T): T => {
    historyMutationDepth++;
    try { return operation(); } finally { historyMutationDepth--; }
  };
  const replay = (direction: "undo" | "redo"): Record<string, JsonValue> | null => mutateHistory(() => {
    if (disposed || !options.canWrite()) return null;
    manager.stopCapturing();
    const stack = direction === "undo" ? manager.undoStack : manager.redoStack;
    // List order is one shared property: its Y.Array stores a full ordering,
    // so native array undo would resurrect old positions ahead of a peer's
    // replacement. Exclude only conflicting order structs from replay; object
    // maps still undo by identity. A restored object missing from the newer
    // order is appended by yToJson, preserving every peer's relative position.
    // The same applies when creation is retained for a peer: its order change
    // must not create a phantom undo step or remove that retained object.
    doc.transact((transaction) => {
      for (const entry of stack) {
        const protectedOrders = new Set<Y.Array<unknown>>();
        for (const [order, revision] of orderRevisions.get(entry) ?? []) {
          const list = order._item?.parent;
          const items = isListMap(list) ? list.get(LIST_ITEMS) : null;
          if ((peerOrderRevision.get(order) ?? 0) !== revision ||
            (items instanceof Y.Map && Array.from(items.values()).some((value) =>
              isContainer(value) && isPreservedCreation(value, entry)))) {
            protectedOrders.add(order);
          }
        }
        if (protectedOrders.size === 0) continue;
        for (const field of ["insertions", "deletions"] as const) {
          const filtered = Y.createDeleteSet();
          Y.iterateDeletedStructs(transaction, entry[field], (item) => {
            if (item instanceof Y.Item && protectedOrders.has(item.parent as Y.Array<unknown>)) return;
            const ranges = filtered.clients.get(item.id.client) ?? [];
            ranges.push({ clock: item.id.clock, len: item.length });
            filtered.clients.set(item.id.client, ranges);
          });
          entry[field] = filtered;
        }
      }
    }, manager);
    replaying = { stack, entries: stack.slice() };
    try {
      if (!manager[direction]()) {
        // Obsolete/conflicting entries can be consumed without a Yjs popped
        // event. Subscribers still need to disable the now-empty action.
        if (stack.length !== replaying.entries.length) listeners.forEach((listener) => listener());
        return null;
      }
      return yToJson(root) as Record<string, JsonValue>;
    } finally { replaying = null; }
  });
  const replayStep = (direction: "undo" | "redo", token: unknown): Record<string, JsonValue> | null => mutateHistory(() => {
    if (disposed || replaying || !options.canWrite()) return null;
    const stack = direction === "undo" ? manager.undoStack : manager.redoStack;
    if (!token || stack.at(-1) !== token) return null;
    // A workbook action may span multiple Y.Docs. Each shard must consume
    // only its part of that action: Yjs otherwise skips an obsolete entry
    // and can undo an older, unrelated action on this shard in the same call.
    const older = stack.splice(0, stack.length - 1);
    const opposite = direction === "undo" ? manager.redoStack : manager.undoStack;
    // Yjs also consults both stacks to distinguish earlier local deletions
    // from peer map changes. Keep that causal information visible while only
    // the selected source entry can be popped.
    opposite.unshift(...older);
    try {
      return replay(direction);
    } finally {
      opposite.splice(0, older.length);
      stack.unshift(...older);
      // Native stack notifications ran while older entries were fenced off.
      // Publish the restored availability even when this step was obsolete.
      listeners.forEach((listener) => listener());
    }
  });
  const discardStep = (direction: "undo" | "redo", token: unknown): boolean => {
    if (disposed || historyMutationDepth > 0) return false;
    const stack = direction === "undo" ? manager.undoStack : manager.redoStack;
    const index = stack.findIndex((entry) => entry === token);
    if (index < 0) return false;
    return mutateHistory(() => {
      manager.stopCapturing();
      const [entry] = stack.splice(index, 1);
      orderRevisions.delete(entry);
      // Bookkeeping only. UndoManager.clear() changes document retention and
      // would release items that another remaining entry may still need.
      // Keep those flags conservative until the document/history is retired.
      listeners.forEach((listener) => listener());
      return true;
    });
  };
  return {
    get canUndo() { return !disposed && options.canWrite() && manager.undoStack.length > 0; },
    get canRedo() { return !disposed && options.canWrite() && manager.redoStack.length > 0; },
    /** How many steps of this member's each stack holds (a history across
     * rooms counts them: lib/collab/multi-room-history). */
    get undoDepth() { return disposed ? 0 : manager.undoStack.length; },
    get redoDepth() { return disposed ? 0 : manager.redoStack.length; },
    undo: () => replay("undo"),
    redo: () => replay("redo"),
    get undoToken(): unknown { return disposed ? null : manager.undoStack.at(-1) ?? null; },
    get redoToken(): unknown { return disposed ? null : manager.redoStack.at(-1) ?? null; },
    undoStep: (token: unknown) => replayStep("undo", token),
    redoStep: (token: unknown) => replayStep("redo", token),
    /** Remove this exact stack-entry identity, including below the top. No
     * inverse is applied. Allowed while read-only; false for absent/retired
     * tokens or during history replay/notification. */
    discardStep,
    stopCapturing: () => { if (!disposed) manager.stopCapturing(); },
    /** The next local change joins the last undo step, as Yjs joins changes
     * made within its capture interval: a correction worked out from that
     * step at once (a table's rows fitted to the words just typed) is undone
     * and redone with it, not on its own. */
    continueLastStep: () => {
      if (!disposed && manager.undoStack.length > 0) manager.lastChange = Date.now();
    },
    subscribe(listener: () => void) {
      if (disposed) return () => {};
      listeners.add(listener);
      manager.on("stack-item-added", listener);
      manager.on("stack-item-popped", listener);
      manager.on("stack-cleared", listener);
      return () => {
        listeners.delete(listener);
        manager.off("stack-item-added", listener);
        manager.off("stack-item-popped", listener);
        manager.off("stack-cleared", listener);
      };
    },
    destroy() {
      if (disposed) return;
      disposed = true;
      listeners.clear();
      doc.off("afterTransaction", observePeer);
      manager.destroy();
    },
  };
}

/** Keep local history when a shard's network session is retired. The replica
 * receives real Yjs updates, preserving client/item identities and tombstones;
 * a JSON mirror would lose peer ABA changes and deleted values during reload. */
export function createJsonCollaborationHistoryReplica(options: {
  rootKey: string;
  scopeKey: string;
  localOrigin: object | symbol;
  replayOrigin: object | symbol;
  canWrite: () => boolean;
  maintenanceOrigins?: ReadonlySet<unknown>;
}) {
  const maintenanceOrigins = new Set(options.maintenanceOrigins);
  if (maintenanceOrigins.has(options.localOrigin) || maintenanceOrigins.has(options.replayOrigin)) {
    throw new Error("Maintenance origins cannot also be author or history replay origins.");
  }
  const replica = new Y.Doc();
  const root = replica.getMap<unknown>(options.rootKey);
  const remoteOrigin = Symbol("history-replica-remote");
  let disposed = false;
  let binding: { doc: Y.Doc; detach: () => void } | null = null;
  const writable = () => !disposed && binding !== null &&
    !binding.doc.isDestroyed && options.canWrite();
  const history = createJsonCollaborationHistory(root, {
    localOrigin: options.localOrigin,
    canWrite: writable,
    maintenanceOrigins,
  });
  const replay = (direction: "undo" | "redo", token?: unknown) => {
    if (!writable() || !binding) return null;
    const current = binding;
    // Catch up even if the network doc was replaced between user actions.
    Y.applyUpdate(replica, Y.encodeStateAsUpdate(current.doc), remoteOrigin);
    const result = token === undefined
      ? history[direction]()
      : direction === "undo" ? history.undoStep(token) : history.redoStep(token);
    if (!result) return null;
    // Include any missing dependencies retained for history, not an old JSON
    // snapshot. Live peer structs already present in the target are untouched.
    const update = Y.encodeStateAsUpdate(replica, Y.encodeStateVector(current.doc));
    Y.applyUpdate(current.doc, update, options.replayOrigin);
    return yToJson(current.doc.getMap(options.rootKey)) as Record<string, JsonValue>;
  };
  return {
    get canUndo() { return history.canUndo; },
    get canRedo() { return history.canRedo; },
    get undoToken() { return history.undoToken; },
    get redoToken() { return history.redoToken; },
    undo: () => replay("undo"),
    redo: () => replay("redo"),
    undoStep: (token: unknown) => replay("undo", token),
    redoStep: (token: unknown) => replay("redo", token),
    // No catch-up or live-document transaction: detached replicas can prune
    // a retired compound token while retaining their older usable history.
    discardStep: history.discardStep,
    stopCapturing: history.stopCapturing,
    subscribe: history.subscribe,
    bind(doc: Y.Doc, scopeKey: string): (() => void) | null {
      if (disposed || doc.isDestroyed || scopeKey !== options.scopeKey) return null;
      binding?.detach();
      Y.applyUpdate(replica, Y.encodeStateAsUpdate(doc), remoteOrigin);
      const current = { doc, detach: () => {} };
      const receive = (update: Uint8Array, origin: unknown) => {
        if (binding !== current || origin === options.replayOrigin) return;
        // Keep maintenance structs/tombstones in the replica, with their
        // classification intact. Ignoring the bytes would later make the
        // catch-up import mistake them for a new peer change.
        Y.applyUpdate(replica, update,
          origin === options.localOrigin || maintenanceOrigins.has(origin) ? origin : remoteOrigin);
      };
      current.detach = () => {
        doc.off("update", receive);
        if (binding === current) binding = null;
      };
      binding = current;
      doc.on("update", receive);
      return current.detach;
    },
    destroy() {
      if (disposed) return;
      disposed = true;
      binding?.detach();
      history.destroy();
      replica.destroy();
    },
  };
}

/**
 * Write `next` into `target`, touching only what actually changed. Must run
 * inside a transaction (the caller owns the origin tag).
 */
export function applyJsonToYMap(
  target: Y.Map<unknown>,
  next: Record<string, JsonValue>,
): void {
  for (const key of Array.from(target.keys())) {
    if (!hasJsonProperty(next, key)) target.delete(key);
  }
  for (const [key, value] of Object.entries(next)) {
    if (value !== undefined) applyJsonEntry(target, key, value);
  }
}

/**
 * `itemOf`: `value` is an entry of the keyed list of that name.
 * `parentItemOf`: `parent` itself is such an entry (decides whether `key` is
 * a collaborative text field).
 */
function applyJsonEntry(
  parent: Y.Map<unknown>,
  key: string,
  value: JsonValue,
  itemOf?: string,
  parentItemOf?: string,
): void {
  const current = parent.get(key);
  const nextIds = keyedListIds(value, key);

  if (nextIds) {
    if (!isListMap(current)) {
      parent.set(key, toYValue(value, key, itemOf));
      return;
    }
    const order = current.get(LIST_ORDER) as Y.Array<string>;
    const items = current.get(LIST_ITEMS) as Y.Map<unknown>;
    for (const id of Array.from(items.keys())) {
      if (!nextIds.includes(id)) items.delete(id);
    }
    for (const [index, item] of (value as JsonValue[]).entries()) {
      applyJsonEntry(items, nextIds[index] as string, item, key);
    }
    const currentIds = order.toArray();
    if (
      currentIds.length !== nextIds.length ||
      currentIds.some((id, index) => id !== nextIds[index])
    ) {
      order.delete(0, order.length);
      order.push(nextIds);
    }
    return;
  }

  if (isPlainObject(value)) {
    if (!(current instanceof Y.Map) || isListMap(current)) {
      parent.set(key, toYValue(value, key, itemOf));
      return;
    }
    const currentMap = current as Y.Map<unknown>;
    const collaborative = itemOf !== undefined && COLLABORATIVE_TEXT_LISTS.has(itemOf);
    const formatted = collaborative ? applyFormattedText(currentMap, value, null) : null;
    if (collaborative && formatted === null) dropFormatting(currentMap);
    for (const existing of Array.from(currentMap.keys())) {
      if ((existing === FORMAT_MARKER || existing === FORMAT_TAIL) && formatted !== null) continue;
      if (!hasJsonProperty(value, existing)) currentMap.delete(existing);
    }
    for (const [childKey, childValue] of Object.entries(value)) {
      if (formatted !== null && (childKey === formatted.owned || childKey === COLLABORATIVE_TEXT_KEY)) continue;
      if (childValue !== undefined) applyJsonEntry(currentMap, childKey, childValue, undefined, itemOf);
    }
    return;
  }

  if (Array.isArray(value)) {
    // Non-keyed arrays are replaced only when they actually differ, so an
    // untouched list never produces an update.
    if (current instanceof Y.Array && jsonEquals(yToJson(current), value)) {
      return;
    }
    parent.set(key, toYValue(value, key, childArrayContext(key, parentItemOf)));
    return;
  }

  if (typeof value === "string" && isCollaborativeTextField(key, parentItemOf)) {
    if (current instanceof Y.Text) {
      spliceYText(current, value);
      return;
    }
    // A legacy scalar (or absent) text becomes a Y.Text on its first change;
    // an unchanged one stays as it is so a re-emitted document writes nothing.
    if (current !== value) parent.set(key, new Y.Text(value));
    return;
  }

  if (current instanceof Y.Map || current instanceof Y.Array || current instanceof Y.Text) {
    parent.set(key, value);
    return;
  }
  if (!scalarEquals(current, value)) {
    parent.set(key, value);
  }
}

// ---------------------------------------------------------------------------
// Delta application
//
// `applyJsonToYMap` writes a document as a REPLACEMENT: keys absent from the
// incoming object are deleted. For a shared room that is wrong — the incoming
// object is ONE member's view, and a key missing from it is usually a peer's
// addition this member has not ingested yet, not a deletion. Publishing that

//
// The delta form takes BOTH the document being published and the document the
// author last knew about, so it can tell the two apart:
//   - changed / added: differs from `previous`      -> written
//   - deleted:         present in `previous`, gone  -> deleted
//   - unknown:         in neither                   -> LEFT ALONE (peer data)
// ---------------------------------------------------------------------------

/** Apply only what changed between `previous` and `next`. Returns true when
 *  anything was written. Must run inside a transaction. */
export function applyJsonDelta(
  target: Y.Map<unknown>,
  next: Record<string, JsonValue>,
  previous: Record<string, JsonValue>,
  options?: {
    /** This member's text edits, written where they were typed (see
     *  `CollaborativeTextEdit`) instead of where a string comparison puts
     *  them. An edit that is not exactly the change of its text is ignored. */
    textEdits?: ReadonlyMap<Y.Text, CollaborativeTextEdit>;
    /** An edit in `textEdits` whose base is not `previous`'s text -- the
     *  author typed before a peer's characters that reached the room were
     *  shown to them -- is written where it was typed, moved over the peer's
     *  characters (a later part's room, whose base is the room's own latest
     *  content). Several notes of one text must then come as one edit
     *  (`collaborativeTextEditsFor`). */
    rebaseStaleEdits?: boolean;
  },
): boolean {
  if (options?.textEdits?.size) {
    const outer = exactTextEdits;
    const outerRebase = rebaseStaleTextEdits;
    exactTextEdits = options.textEdits;
    rebaseStaleTextEdits = options.rebaseStaleEdits === true;
    try {
      return applyJsonDelta(target, next, previous);
    } finally {
      exactTextEdits = outer;
      rebaseStaleTextEdits = outerRebase;
    }
  }
  let touched = false;
  for (const key of Object.keys(previous)) {
    if (hasJsonProperty(previous, key) && !hasJsonProperty(next, key) && target.has(key)) {
      target.delete(key);
      touched = true;
    }
  }
  for (const [key, value] of Object.entries(next)) {
    if (value === undefined) continue;
    if (key in previous && jsonEquals(value, previous[key])) continue;
    if (applyJsonDeltaEntry(target, key, value, previous[key])) touched = true;
  }
  return touched;
}


// ---------------------------------------------------------------------------
// Collaborative character and paragraph formatting
//
// A slide object's formatting -- its `runs` (a partition of `text`) or its
// `paragraphs` (each paragraph's settings and runs) -- was one value: two
// people formatting different words of one text box at once kept one
// person's value, and the other's bold, colour or alignment was lost. When
// the formatting is canonical (see `encodeFormatting`) it is stored as the
// Y.Text's own formatting instead and the block map says which kind
// (`FORMAT_MARKER`): each run style key is an attribute of the characters it
// covers; each paragraph setting an attribute (`PARAGRAPH_ATTRIBUTE` + key) of
// the line break that ends the paragraph, or of the block's `FORMAT_TAIL` map
// for the last paragraph, which ends without one. Concurrent formatting then
// merges by character and by attribute. Other formatting stays a value, as
// before. The server adapter (page_causal_crdt.py) reads and writes the same
// model.
// ---------------------------------------------------------------------------

const RUNS_KEY = "runs";
const PARAGRAPHS_KEY = "paragraphs";
export const FORMAT_MARKER = "__mrn_format__";
export const FORMAT_TAIL = "__mrn_tail__";
/** Prefix of a paragraph setting's attribute; never a run style key. */
const PARAGRAPH_ATTRIBUTE = "¶";
/** The paragraph has runs (else its words are plain). */
const PARAGRAPH_RUNS_FLAG = `${PARAGRAPH_ATTRIBUTE}runs`;

type FormatKind = "runs" | "paragraphs";
type Attributes = Record<string, JsonValue>;
/**
 * A block's formatting as attributes: one per UTF-16 unit of `text`, and for
 * paragraphs the last paragraph's (`tail`).
 */
type Formatting = { kind: FormatKind; units: Attributes[]; tail: Attributes | null };

const ownedKey = (kind: FormatKind) => (kind === "runs" ? RUNS_KEY : PARAGRAPHS_KEY);

function formatKindOf(block: Y.Map<unknown>): FormatKind | null {
  const kind = block.get(FORMAT_MARKER);
  return kind === "runs" || kind === "paragraphs" ? kind : null;
}

function stableJson(value: JsonValue | undefined): string {
  if (isPlainObject(value)) {
    return `{${Object.keys(value).filter((key) => value[key] !== undefined).sort().map((key) => `${JSON.stringify(key)}:${stableJson(value[key])}`).join(",")}}`;
  }
  if (Array.isArray(value)) return `[${value.map((item) => stableJson(item)).join(",")}]`;
  return JSON.stringify(value ?? null);
}

/** Each run's attributes, or null when `runs` over `text` are not canonical. */
function canonicalRunAttributes(text: string, runs: unknown): Attributes[] | null {
  if (!Array.isArray(runs)) return null;
  let joined = "";
  let previous: string | null = null;
  const out: Attributes[] = [];
  for (const run of runs) {
    if (!isPlainObject(run)) return null;
    if (Object.keys(run).some((key) => key !== "text" && key !== "style")) return null;
    if (typeof run.text !== "string" || run.text.length === 0) return null;
    let attrs: Attributes = {};
    if (hasJsonProperty(run, "style")) {
      const style = run.style;
      if (!isPlainObject(style) || Object.keys(style).length === 0) return null;
      for (const [key, value] of Object.entries(style)) {
        if (value === null || value === undefined || key.startsWith(PARAGRAPH_ATTRIBUTE)) return null;
      }
      attrs = style;
    }
    const key = stableJson(attrs);
    if (key === previous) return null;
    previous = key;
    joined += run.text;
    for (let index = 0; index < run.text.length; index += 1) out.push(attrs);
  }
  return joined === text ? out : null;
}

/** A paragraph's settings as attributes, or null when they cannot be. */
function paragraphAttributes(paragraph: Record<string, JsonValue>): Attributes | null {
  const attrs: Attributes = {};
  for (const [key, value] of Object.entries(paragraph)) {
    if (key === "text" || key === RUNS_KEY || value === undefined) continue;
    if (value === null || key.startsWith(PARAGRAPH_ATTRIBUTE)) return null;
    attrs[`${PARAGRAPH_ATTRIBUTE}${key}`] = value;
  }
  if (hasJsonProperty(paragraph, RUNS_KEY)) attrs[PARAGRAPH_RUNS_FLAG] = true;
  return attrs;
}

/**
 * A block's formatting as attributes of its text, or null when it cannot be
 * (it stays a value). Paragraphs are canonical when each is an object whose
 * `text` holds no line break, they join with line breaks into `text`, no
 * setting is null, and their runs (where present) are canonical; runs when
 * each is `{text, style?}` with words, no style value null, no two
 * neighbours alike, and together exactly `text`.
 */
export function encodeFormatting(block: Record<string, JsonValue>): Formatting | null {
  const text = block.text;
  if (typeof text !== "string") return null;
  const paragraphs = block[PARAGRAPHS_KEY];
  if (Array.isArray(paragraphs) && paragraphs.length > 0) {
    const units: Attributes[] = [];
    let tail: Attributes | null = null;
    let joined = "";
    let canonical = true;
    paragraphs.forEach((paragraph, index) => {
      if (!canonical) return;
      if (!isPlainObject(paragraph) || typeof paragraph.text !== "string" || paragraph.text.includes("\n")) {
        canonical = false;
        return;
      }
      const attrs = paragraphAttributes(paragraph);
      const runs = hasJsonProperty(paragraph, RUNS_KEY)
        ? paragraph.text.length === 0 && Array.isArray(paragraph.runs) && paragraph.runs.length === 0
          ? []
          : canonicalRunAttributes(paragraph.text, paragraph.runs)
        : Array.from({ length: paragraph.text.length }, () => ({}));
      if (!attrs || !runs) {
        canonical = false;
        return;
      }
      units.push(...runs);
      joined += paragraph.text;
      if (index < paragraphs.length - 1) {
        units.push(attrs);
        joined += "\n";
      } else {
        tail = attrs;
      }
    });
    if (canonical && joined === text) return { kind: "paragraphs", units, tail };
  }
  if (hasJsonProperty(block, RUNS_KEY)) {
    const runs = block[RUNS_KEY];
    const units = text.length === 0 && Array.isArray(runs) && runs.length === 0 ? [] : canonicalRunAttributes(text, runs);
    if (units) return { kind: "runs", units, tail: null };
  }
  return null;
}

type DeltaOp = { insert?: unknown; attributes?: Attributes };

function runStyleOf(attributes: Attributes | undefined): Attributes {
  const style: Attributes = {};
  for (const [key, value] of Object.entries(attributes ?? {})) {
    if (!key.startsWith(PARAGRAPH_ATTRIBUTE)) style[key] = value;
  }
  return style;
}

function appendRun(runs: Array<{ text: string; style?: Attributes }>, text: string, style: Attributes): void {
  if (text.length === 0) return;
  const last = runs[runs.length - 1];
  if (last && jsonEquals((last.style ?? {}) as JsonValue, style as JsonValue)) {
    last.text += text;
    return;
  }
  runs.push(Object.keys(style).length > 0 ? { text, style: { ...style } } : { text });
}

/** The block's formatting read back from its text (and tail). */
function decodeFormatting(kind: FormatKind, text: Y.Text, tail: unknown): JsonValue {
  const ops = text.toDelta() as DeltaOp[];
  if (kind === "runs") {
    const runs: Array<{ text: string; style?: Attributes }> = [];
    for (const op of ops) if (typeof op.insert === "string") appendRun(runs, op.insert, runStyleOf(op.attributes));
    return runs as JsonValue;
  }
  type Paragraph = { text: string; runs: Array<{ text: string; style?: Attributes }>; attrs: Attributes | null };
  const paragraphs: Paragraph[] = [{ text: "", runs: [], attrs: null }];
  const settingsOf = (attributes: Attributes | undefined): Attributes | null => {
    const out: Attributes = {};
    for (const [key, value] of Object.entries(attributes ?? {})) if (key.startsWith(PARAGRAPH_ATTRIBUTE)) out[key] = value;
    return Object.keys(out).length > 0 ? out : null;
  };
  for (const op of ops) {
    if (typeof op.insert !== "string") continue;
    const pieces = op.insert.split("\n");
    pieces.forEach((piece, index) => {
      const current = paragraphs[paragraphs.length - 1]!;
      current.text += piece;
      appendRun(current.runs, piece, runStyleOf(op.attributes));
      if (index < pieces.length - 1) {
        current.attrs = settingsOf(op.attributes);
        paragraphs.push({ text: "", runs: [], attrs: null });
      }
    });
  }
  paragraphs[paragraphs.length - 1]!.attrs = tail instanceof Y.Map ? settingsOf(tail.toJSON() as Attributes) : null;
  // A paragraph whose break carries no settings (one a peer's concurrent
  // edit split off) takes its neighbour's, as a paragraph split by Enter
  // does in PowerPoint.
  paragraphs.forEach((paragraph, index) => {
    if (paragraph.attrs) return;
    for (let at = index - 1; at >= 0; at -= 1) {
      if (paragraphs[at]!.attrs) {
        paragraph.attrs = paragraphs[at]!.attrs;
        return;
      }
    }
    paragraph.attrs = paragraphs.slice(index + 1).find((other) => other.attrs)?.attrs ?? {};
  });
  return paragraphs.map((paragraph) => {
    const out: Attributes = { text: paragraph.text };
    let hasRuns = false;
    for (const [key, value] of Object.entries(paragraph.attrs ?? {})) {
      if (key === PARAGRAPH_RUNS_FLAG) hasRuns = value === true;
      else out[key.slice(PARAGRAPH_ATTRIBUTE.length)] = value;
    }
    if (hasRuns) out[RUNS_KEY] = paragraph.runs as JsonValue;
    return out;
  }) as JsonValue;
}

/** Per UTF-16 unit, the attributes the text holds there. */
function attributesByUnit(text: Y.Text): Attributes[] {
  const out: Attributes[] = [];
  for (const op of text.toDelta() as DeltaOp[]) {
    if (typeof op.insert !== "string") continue;
    for (let index = 0; index < op.insert.length; index += 1) out.push(op.attributes ?? {});
  }
  return out;
}

/** The patch that turns `current` attributes into `wanted` (null removes). */
function attributePatch(current: Attributes, wanted: Attributes): Record<string, JsonValue | null> | null {
  const patch: Record<string, JsonValue | null> = {};
  for (const [key, value] of Object.entries(wanted)) {
    if (!hasJsonProperty(current, key) || !jsonEquals(current[key] as JsonValue, value)) patch[key] = value;
  }
  for (const key of Object.keys(current)) if (!hasJsonProperty(wanted, key)) patch[key] = null;
  return Object.keys(patch).length > 0 ? patch : null;
}

/** The attributes the author changed from `had` to `wanted` (null removes). */
function changedAttributes(had: Attributes, wanted: Attributes): Record<string, JsonValue | null> | null {
  const changed: Record<string, JsonValue | null> = {};
  for (const key of new Set([...Object.keys(had), ...Object.keys(wanted)])) {
    const was = hasJsonProperty(had, key) ? had[key] : undefined;
    const is = hasJsonProperty(wanted, key) ? wanted[key] : undefined;
    if (was !== undefined && is !== undefined && jsonEquals(was, is)) continue;
    changed[key] = is === undefined ? null : is;
  }
  return Object.keys(changed).length > 0 ? changed : null;
}

function tailAttributes(block: Y.Map<unknown>): Attributes {
  const tail = block.get(FORMAT_TAIL);
  return tail instanceof Y.Map ? (tail.toJSON() as Attributes) : {};
}

/** Write each unit's patch (`at` in the room's text; the text's length is the tail). */
function writePatches(
  block: Y.Map<unknown>,
  text: Y.Text,
  patches: ReadonlyArray<{ at: number; patch: Record<string, JsonValue | null> }>,
): boolean {
  let touched = false;
  const length = text.length;
  let at = 0;
  while (at < patches.length) {
    const first = patches[at]!;
    if (first.at >= length) {
      let tail = block.get(FORMAT_TAIL);
      if (!(tail instanceof Y.Map)) {
        tail = new Y.Map<unknown>();
        block.set(FORMAT_TAIL, tail);
      }
      for (const [key, value] of Object.entries(first.patch)) {
        if (value === null) (tail as Y.Map<unknown>).delete(key);
        else (tail as Y.Map<unknown>).set(key, value);
      }
      touched = true;
      at += 1;
      continue;
    }
    const key = stableJson(first.patch as JsonValue);
    let end = at + 1;
    while (
      end < patches.length &&
      patches[end]!.at < length &&
      patches[end]!.at === patches[end - 1]!.at + 1 &&
      stableJson(patches[end]!.patch as JsonValue) === key
    ) {
      end += 1;
    }
    text.format(first.at, end - at, first.patch);
    touched = true;
    at = end;
  }
  return touched;
}

/** Make the text's formatting exactly `wanted` (the text already reads as the author's). */
function syncFormatting(block: Y.Map<unknown>, text: Y.Text, wanted: Formatting): boolean {
  const current = attributesByUnit(text);
  const patches: Array<{ at: number; patch: Record<string, JsonValue | null> }> = [];
  wanted.units.forEach((attrs, index) => {
    const patch = attributePatch(current[index] ?? {}, attrs);
    if (patch) patches.push({ at: index, patch });
  });
  let touched = writePatches(block, text, patches);
  if (wanted.tail) {
    const patch = attributePatch(tailAttributes(block), wanted.tail);
    if (patch && writePatches(block, text, [{ at: text.length, patch }])) touched = true;
  } else if (block.has(FORMAT_TAIL)) {
    block.delete(FORMAT_TAIL);
    touched = true;
  }
  return touched;
}

/**
 * The author changed the formatting from `before` (over `previousText`) to
 * `after` (over `nextText`); the room's text now holds the author's splice
 * and whatever peers typed. Only the attributes the author changed are
 * written, where those characters now stand: a peer's formatting of other
 * words, or of other attributes of the same words, stays. Without `before`
 * (no base of this kind) the author's formatting is written whole, still
 * only where the author's characters stand.
 */
function applyFormattingDelta(
  block: Y.Map<unknown>,
  text: Y.Text,
  previousText: string,
  before: Formatting | null,
  nextText: string,
  after: Formatting,
): boolean {
  const room = text.toString();
  const roomUnits = attributesByUnit(text);
  const roomTail = tailAttributes(block);
  const edit = exactTextEdits?.get(text);
  const author = edit && validTextEdit(edit, previousText, nextText)
    ? { start: edit.start, end: edit.end, insert: edit.insert }
    : textSplice(previousText, nextText);
  const toRoom = textSplice(nextText, room);
  const insertedEnd = author.start + author.insert.length;
  const withTail = after.kind === "paragraphs";
  const patches: Array<{ at: number; patch: Record<string, JsonValue | null> }> = [];
  const last = withTail ? nextText.length : nextText.length - 1;
  for (let index = 0; index <= last; index += 1) {
    const isTail = index === nextText.length;
    // Where this character of the author's text stands in the room now.
    if (!isTail && index >= toRoom.start && index < toRoom.end) continue; // a peer changed it
    const roomIndex = isTail ? room.length : index < toRoom.start ? index : index - (toRoom.end - toRoom.start) + toRoom.insert.length;
    const wanted = (isTail ? after.tail : after.units[index]) ?? {};
    let patch: Record<string, JsonValue | null> | null;
    if (!before || (!isTail && index >= author.start && index < insertedEnd)) {
      // Characters the author typed (or no base): exactly as the author has them.
      patch = attributePatch((isTail ? roomTail : roomUnits[roomIndex]) ?? {}, wanted);
    } else {
      const previousIndex = isTail
        ? previousText.length
        : index < author.start ? index : index - author.insert.length + (author.end - author.start);
      const had = (previousIndex === previousText.length ? before.tail : before.units[previousIndex]) ?? {};
      patch = changedAttributes(had, wanted);
    }
    if (patch) patches.push({ at: roomIndex, patch });
  }
  let touched = writePatches(block, text, patches);
  if (!withTail && block.has(FORMAT_TAIL)) {
    block.delete(FORMAT_TAIL);
    touched = true;
  }
  return touched;
}

/**
 * Write a collaborative block's `text` and its formatting together.
 * `previous` is the author's last view of the block (delta mode), or null
 * (replace mode). Returns null when the formatting is not canonical (the
 * caller writes it as a value), else the key it owns and whether anything
 * was written.
 */
function applyFormattedText(
  block: Y.Map<unknown>,
  next: Record<string, JsonValue>,
  previous: Record<string, JsonValue> | null,
): { owned: string; touched: boolean } | null {
  const wanted = encodeFormatting(next);
  if (!wanted) return null;
  const nextText = next.text as string;
  const owned = ownedKey(wanted.kind);
  let touched = false;
  let current = block.get(COLLABORATIVE_TEXT_KEY);
  const previousText = previous && typeof previous.text === "string" ? previous.text : null;
  if (!(current instanceof Y.Text)) {
    const created = new Y.Text();
    block.set(COLLABORATIVE_TEXT_KEY, created);
    current = created;
    touched = true;
  }
  const text = current as Y.Text;
  const kind = formatKindOf(block);
  // Formatting set as a value by a writer of the older model is read as
  // written; on this write it moves into the text.
  const valueFormatting = block.has(owned);
  if (previousText !== null) {
    if (applyTextDelta(block, COLLABORATIVE_TEXT_KEY, text, nextText, previousText)) touched = true;
    // An edit typed on a text the room had moved past (see applyTextDelta):
    // the formatting is measured from that text, and the author's earlier
    // formatting is not known of it -- the author's characters are written
    // as they have them, the peer's stay as the peer has them.
    const typedOn = staleEditBase(exactTextEdits?.get(text), previousText, nextText);
    const before = typedOn === null && kind === wanted.kind && !valueFormatting && previous ? encodeFormatting(previous) : null;
    if (applyFormattingDelta(block, text, typedOn ?? previousText, before?.kind === wanted.kind ? before : null, nextText, wanted)) touched = true;
  } else {
    if (spliceYText(text, nextText)) touched = true;
    if (syncFormatting(block, text, wanted)) touched = true;
  }
  if (valueFormatting) {
    block.delete(owned);
    touched = true;
  }
  if (kind !== wanted.kind) {
    block.set(FORMAT_MARKER, wanted.kind);
    touched = true;
  }
  return { owned, touched };
}

/** The formatting leaves the text (it is a value again, or gone). */
function dropFormatting(block: Y.Map<unknown>): boolean {
  if (!block.has(FORMAT_MARKER) && !block.has(FORMAT_TAIL)) return false;
  const text = block.get(COLLABORATIVE_TEXT_KEY);
  if (text instanceof Y.Text) {
    const patches: Array<{ at: number; patch: Record<string, JsonValue | null> }> = [];
    attributesByUnit(text).forEach((attrs, index) => {
      const patch = attributePatch(attrs, {});
      if (patch) patches.push({ at: index, patch });
    });
    writePatches(block, text, patches);
  }
  block.delete(FORMAT_MARKER);
  block.delete(FORMAT_TAIL);
  return true;
}

/** `itemOf` / `parentItemOf`: as for `applyJsonEntry`. */
function applyJsonDeltaEntry(
  parent: Y.Map<unknown>,
  key: string,
  next: JsonValue,
  previous: JsonValue | undefined,
  itemOf?: string,
  parentItemOf?: string,
): boolean {
  const current = parent.get(key);
  const nextIds =
    keyedListIds(next, key) ??
    (Array.isArray(next) && next.length === 0 && isListMap(current)
      ? []
      : null);

  // Keyed lists merge per id; only the ids the author actually touched are
  // written, so a peer's concurrently added entry survives.
  if (nextIds && isListMap(current)) {
    const order = current.get(LIST_ORDER) as Y.Array<string>;
    const items = current.get(LIST_ITEMS) as Y.Map<unknown>;
    const prevList = Array.isArray(previous) ? previous : [];
    const prevIds = keyedListIds(prevList, key) ?? [];
    const prevById = new Map(
      prevList.flatMap((item, index) => {
        const id = prevIds[index];
        return id ? [[id, item] as const] : [];
      }),
    );
    let touched = false;
    for (const [id] of prevById) {
      if (!nextIds.includes(id) && items.has(id)) {
        items.delete(id);
        touched = true;
      }
    }
    for (const [index, item] of (next as JsonValue[]).entries()) {
      const id = nextIds[index] as string;
      const before = prevById.get(id);
      if (before && jsonEquals(item, before)) continue;
      if (applyJsonDeltaEntry(items, id, item, before, key)) touched = true;
    }
    const currentOrder = order.toArray();
    const orderChangedLocally =
      prevIds.length !== nextIds.length ||
      prevIds.some((id, index) => id !== nextIds[index]);
    if (orderChangedLocally) {
      // Preserve ids the author never saw by appending them after the
      // author's ordering instead of dropping them.
      const unseen = currentOrder.filter(
        (id) => !nextIds.includes(id) && !prevIds.includes(id),
      );
      const merged = [...nextIds, ...unseen];
      if (
        currentOrder.length !== merged.length ||
        currentOrder.some((id, index) => id !== merged[index])
      ) {
        order.delete(0, order.length);
        order.push(merged);
        touched = true;
      }
    }
    return touched;
  }

  if (isPlainObject(next) && current instanceof Y.Map && !isListMap(current)) {
    return applyMapDelta(current as Y.Map<unknown>, next, isPlainObject(previous) ? previous : {}, itemOf);
  }

  if (typeof next === "string" && isCollaborativeTextField(key, parentItemOf)) {
    return applyTextDelta(parent, key, current, next, previous);
  }

  const arrayContext = childArrayContext(key, parentItemOf);
  if (current instanceof Y.Array && applyArrayInPlace(current, next, previous, key, arrayContext)) return true;

  // Leaf (or a container whose SHAPE changed): write it whole. This is a real
  // authored change, so overwriting the peer's value at the same path is the
  // intended last-writer-wins.
  const existing = yToJson(current);
  if (current !== undefined && jsonEquals(existing, next)) return false;
  parent.set(key, toYValue(next, key, itemOf ?? arrayContext));
  return true;
}

/**
 * The author changed the map `childMap` (an entry of the keyed list
 * `itemOf`, or a table cell) from `prevObject` to `next`: only what changed
 * is written, a slide object's or a cell's words with their formatting.
 */
function applyMapDelta(
  childMap: Y.Map<unknown>,
  next: Record<string, JsonValue>,
  prevObject: Record<string, JsonValue>,
  itemOf: string | undefined,
): boolean {
  let touched = false;
  // A slide object's words and their formatting are written together.
  const collaborative = itemOf !== undefined && COLLABORATIVE_TEXT_LISTS.has(itemOf);
  const formattedChanged =
    collaborative &&
    [COLLABORATIVE_TEXT_KEY, RUNS_KEY, PARAGRAPHS_KEY].some((field) => !jsonEquals(next[field] ?? null, prevObject[field] ?? null));
  const heldKind = formatKindOf(childMap);
  const formatted = formattedChanged ? applyFormattedText(childMap, next, prevObject) : null;
  if (formatted?.touched) touched = true;
  if (formattedChanged && formatted === null && dropFormatting(childMap)) touched = true;
  // Formatting the text held until now is a value again (unchanged by the
  // author, so the loop below would not write it).
  const released = formattedChanged && heldKind && ownedKey(heldKind) !== formatted?.owned ? ownedKey(heldKind) : null;
  if (released && hasJsonProperty(next, released) && next[released] !== undefined) {
    childMap.set(released, toYValue(next[released], released));
    touched = true;
  }
  for (const childKey of Object.keys(prevObject)) {
    if (hasJsonProperty(prevObject, childKey) && !hasJsonProperty(next, childKey) && childMap.has(childKey)) {
      childMap.delete(childKey);
      touched = true;
    }
  }
  for (const [childKey, childValue] of Object.entries(next)) {
    if (childValue === undefined) continue;
    if (
      childKey in prevObject &&
      jsonEquals(childValue, prevObject[childKey])
    ) {
      continue;
    }
    if (formatted !== null && (childKey === formatted.owned || childKey === COLLABORATIVE_TEXT_KEY)) continue;
    if (childKey === released) continue;
    if (
      applyJsonDeltaEntry(
        childMap,
        childKey,
        childValue,
        prevObject[childKey],
        undefined,
        itemOf,
      )
    ) {
      touched = true;
    }
  }
  return touched;
}















function applyArrayInPlace(
  current: Y.Array<unknown>,
  next: JsonValue,
  previous: JsonValue | undefined,
  key: string,
  context?: string,
): boolean {
  if (!Array.isArray(next) || !Array.isArray(previous) || next.length !== previous.length ||
      current.length !== previous.length || keyedListIds(next, key)) return false;
  const changed: number[] = [];
  for (let index = 0; index < next.length; index += 1) {
    if (!jsonEquals(next[index], previous[index])) changed.push(index);
  }
  if (changed.length === 0) return false;
  const entryOf = arrayEntryContext(context);
  // Every entry about to be written must still be what the author changed.
  // A table's rows and cells are merged instead (a peer typing in the same
  // cell, or another cell of the row): a row needs only to still be a row of
  // as many cells, a cell a map, and each is written by its own changes.
  for (const index of changed) {
    const entry = current.get(index);
    if (context === TABLE_ROWS) {
      if (!(entry instanceof Y.Array) || !Array.isArray(previous[index]) ||
          entry.length !== (previous[index] as JsonValue[]).length) return false;
      continue;
    }
    if (context === TABLE_ROW) {
      if (!(entry instanceof Y.Map) || isListMap(entry) || !isPlainObject(previous[index]) ||
          !isPlainObject(next[index])) return false;
      continue;
    }
    if (!jsonEquals(yToJson(entry), previous[index])) return false;
  }
  for (const index of changed) {
    const entry = current.get(index);
    const after = next[index], before = previous[index];
    if (context === TABLE_ROW && entry instanceof Y.Map) {
      applyMapDelta(entry, after as Record<string, JsonValue>, before as Record<string, JsonValue>, TABLE_CELL);
      continue;
    }
    // A plain map (not a keyed list, not a text with its formatting encoded,
    // whose keys are written only through their own paths).
    if (isPlainObject(after) && isPlainObject(before) && entry instanceof Y.Map &&
        entry.get(LIST_MARKER) !== true && !entry.has(FORMAT_MARKER)) {
      const map = entry as Y.Map<unknown>;
      for (const childKey of Object.keys(before)) {
        if (hasJsonProperty(before, childKey) && !hasJsonProperty(after, childKey) && map.has(childKey)) map.delete(childKey);
      }
      for (const [childKey, childValue] of Object.entries(after)) {
        if (childValue === undefined) continue;
        if (childKey in before && jsonEquals(childValue, before[childKey])) continue;
        applyJsonDeltaEntry(map, childKey, childValue, before[childKey]);
      }
      continue;
    }
    if (Array.isArray(after) && entry instanceof Y.Array && applyArrayInPlace(entry, after, before, key, entryOf)) continue;
    current.delete(index, 1);
    current.insert(index, [toYValue(after, undefined, entryOf)]);
  }
  return true;
}

/**
 * The author changed a collaborative text from `previous` to `next`. Only
 * their own splice is written; characters a peer typed meanwhile — already in
 * the room's text but not in `previous` — stay where the peer put them.
 */
function applyTextDelta(
  parent: Y.Map<unknown>,
  key: string,
  current: unknown,
  next: string,
  previous: JsonValue | undefined,
): boolean {
  if (current instanceof Y.Text) {
    const room = current.toString();
    const edit = exactTextEdits?.get(current);
    if (edit && room === previous && validTextEdit(edit, previous, next)) {
      if (edit.end > edit.start) current.delete(edit.start, edit.end - edit.start);
      if (edit.insert) current.insert(edit.start, edit.insert);
      return edit.end > edit.start || edit.insert.length > 0;
    }
    // The author typed on a text a peer's characters had not reached yet
    // (the room's base moved on to them first): only the author's own edit
    // is written, where they typed it. Measured from the room's base, the
    // author's text read as the peer's characters deleted -- the text box
    // was rewritten whole, and two windows doing it at once wrote it twice
    
    const typedOn = staleEditBase(edit, previous, next);
    if (edit && typedOn !== null) {
      const moved = rebaseTextEdit(edit, room);
      if (moved) {
        if (moved.end > moved.start) current.delete(moved.start, moved.end - moved.start);
        if (moved.insert) current.insert(moved.start, moved.insert);
        return moved.end > moved.start || moved.insert.length > 0;
      }
      return spliceYText(current, mergeText3(typedOn, next, room));
    }
    // Without a text base (the author added the field) there is no splice
    // to rebase: the author's value is written, as for any other leaf.
    const merged = typeof previous !== "string" || room === previous
      ? next
      : mergeText3(previous, next, room);
    return spliceYText(current, merged);
  }
  // A legacy room still holds the string as a scalar: upgrade it in place on
  // this first change. A peer's write that already arrived is merged in; one
  // still in flight is a scalar write racing this Y.Text, and the map key
  // keeps only one of them — a one-time last-writer-wins per legacy text box,
  // after which every edit is a character splice.
  const merged = typeof current === "string" && typeof previous === "string" && current !== previous
    ? mergeText3(previous, next, current)
    : next;
  parent.set(key, new Y.Text(merged));
  return true;
}
