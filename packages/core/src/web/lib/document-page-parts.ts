/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { assertNoStringifiedPageKey } from "./page-key-guard";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { apiGet, apiPut, ApiRequestError } from "./api";
import { reportClientError } from "./errors/report-beacon";
import { sameDocumentAuthorContent } from "../__cut__/m318e23f593";
import { grownByTyping, mergeCanonicalBlocks, sameAuthoredBlock, sameAuthoredBody, type MergeBlock } from "./collab/canonical-block-merge";
import { WordPartRoom, type LostLetters } from "./collab/document-word-part-room";
import { holdPageForPartRoom } from "./collab/part-room-hold";
import { holdDocumentPartOwners, layoutSnapshotRows } from "./docs-layout-snapshot-store";
import type { KeptPartLayout } from "../components/console/editor/native-document/paged/layoutSnapshot";
import type { Schema } from "@tiptap/pm/model";

export const DOCUMENT_PARTS_SCHEMA = "mrnote.document-parts.v1";
/** The imports a document held as parts comes from (the server's
 * `PARTS_SOURCE_KINDS`): a Word document, and a text or Markdown file. */
const DOCUMENT_PARTS_SOURCE_KINDS: ReadonlySet<string> = new Set(["docx", "markdown", "text"]);
const MAX_DOCUMENT_PARTS = 2_000;
/** Layout receipts describe the whole loaded document, never the body alone. */
const PRIVATE_LAYOUT_ATTRS = ["pageBlockSizes", "pageBlockLayoutKey", "pageRegionLayoutReceipt"] as const;


const PART_SAVE_DEBOUNCE_MS = 400;
const PART_SAVE_RETRY_MS = 3_000;
/** The longest wait between tries of a part's save that does not go through. */
const PART_SAVE_RETRY_MOST_MS = 30_000;
/** How many times a part save refused for holding a block another part holds
 * (`document_part_blocks_in_other_part`) is tried again before the view stops:
 * 3, 6, 12, 24, 30, 30 s -- a block moved across parts is saved out of the
 * part it left well within that. */
const PART_BLOCKS_ELSEWHERE_TRIES = 6;

/** The wait before try `tries` (1, 2, ...) of a part's save: doubling, at most PART_SAVE_RETRY_MOST_MS. */
export function partSaveRetryDelay(tries: number): number {
  return Math.min(PART_SAVE_RETRY_MOST_MS, PART_SAVE_RETRY_MS * 2 ** Math.max(0, tries - 1));
}
/** A room not read this long after it was asked for leaves its part's edits to
 * the part's own save (it is read later if it ever comes). */
const ROOM_FIRST_READ_WAIT_MS = 15_000;
/** The body's room key on the server (`BODY_PART_KEY`). */
const BODY_ROOM_KEY = "body";
/** How soon after a change of the view its rooms hear of it, at the least. */
const ROOM_PUBLISH_MS = 120;
/** ... and at the most, however long the last publish took. */
const ROOM_PUBLISH_MOST_MS = 2_000;
/** How far below the viewport the next part starts loading. */
const LOAD_AHEAD_MARGIN = `${2400}px 0px`;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * A deck held as parts (PPT progressive open, stage 3): the same parts as a
 * long document, a slide the unit. The body holds the deck-level fields and
 * the first slides; its `deckParts` marker names every slide in show order
 * and the part it is in (the server's `validate_deck_parts_entry`). Only the
 * deck side reads it (lib/deck-page-parts); it is here because which pages
 * are held as parts is one question.
 */
export const DECK_PARTS_KEY = "deckParts";
export const DECK_PARTS_KIND = "slides";
/** What the order states of a slide of a later part (`_DECK_ORDER_FIELDS`). */
export const DECK_ORDER_FIELDS = [
  "slideId", "title", "hidden", "section", "masterId", "layoutId", "hasNotes", "hasAnimations",
  "transition", "commentCount", "widthEmu", "heightEmu",
] as const;

export interface DeckOrderEntry {
  readonly id: string;
  readonly part: number;
  readonly slideId?: string;
  readonly title?: string;
  /** A later part's slide is hidden or shown here alone (the order is the
   * one source of it; the part's own slide says nothing). */
  readonly hidden?: boolean;
  /** The named section the slide is in. */
  readonly section?: string;
  readonly masterId?: string;
  readonly layoutId?: string;
  readonly hasNotes?: boolean;
  readonly hasAnimations?: boolean;
  readonly transition?: string;
  readonly commentCount?: number;
  readonly widthEmu?: number;
  readonly heightEmu?: number;
}

export interface DeckPartsMarker {
  readonly schema: typeof DOCUMENT_PARTS_SCHEMA;
  readonly kind: typeof DECK_PARTS_KIND;
  readonly count: number;
  readonly order: readonly DeckOrderEntry[];
  /** Parts the import read otherwise than it first wrote them, after their
   * author edited them: the author is asked, nothing is replaced. */
  readonly reimportNeeded?: readonly number[];
}

const DECK_MARKER_KEYS: ReadonlySet<string> = new Set(["schema", "kind", "count", "order", "reimportNeeded"]);
const DECK_ENTRY_KEYS: ReadonlySet<string> = new Set(["id", "part", ...DECK_ORDER_FIELDS]);

/** The body's well-formed `deckParts` marker, or null (the server's rules:
 * 2..2000 parts, every slide named once in a part the count holds, every
 * part naming a slide). */
export function deckPartsMarker(content: unknown): DeckPartsMarker | null {
  if (!isRecord(content)) return null;
  const marker = content[DECK_PARTS_KEY];
  if (!isRecord(marker) || marker.schema !== DOCUMENT_PARTS_SCHEMA || marker.kind !== DECK_PARTS_KIND) return null;
  if (Object.keys(marker).some((key) => !DECK_MARKER_KEYS.has(key))) return null;
  const count = marker.count;
  if (typeof count !== "number" || !Number.isInteger(count) || count < 2 || count > MAX_DOCUMENT_PARTS) return null;
  const order = marker.order;
  if (!Array.isArray(order) || order.length === 0) return null;
  const seen = new Set<string>();
  const used = new Set<number>();
  for (const entry of order) {
    if (!isRecord(entry) || typeof entry.id !== "string" || !entry.id || seen.has(entry.id)) return null;
    const part = entry.part;
    if (typeof part !== "number" || !Number.isInteger(part) || part < 0 || part >= count) return null;
    if (Object.keys(entry).some((key) => !DECK_ENTRY_KEYS.has(key))) return null;
    seen.add(entry.id);
    used.add(part);
  }
  // A later part may name no slide: its slides deleted, kept in the part for Undo (the
  // server's `validate_deck_parts_entry`). Read as no marker, the window saved a body
  // with none, and every save of that deck was refused for good.
  if (!used.has(0)) return null;
  const reimport = marker.reimportNeeded;
  if (reimport !== undefined && (
    !Array.isArray(reimport) || new Set(reimport).size !== reimport.length ||
    reimport.some((index) => typeof index !== "number" || !Number.isInteger(index) || index < 1 || index >= count)
  )) return null;
  return marker as unknown as DeckPartsMarker;
}

/** How many parts a page body names (1 for an ordinary page). */
export function documentPartsCount(content: unknown): number {
  if (!isRecord(content)) return 1;
  const deck = deckPartsMarker(content);
  if (deck) return deck.count;
  const marker = content.oversizedImport;
  if (
    !isRecord(marker) ||
    !DOCUMENT_PARTS_SOURCE_KINDS.has(String(marker.sourceKind))
  ) {
    return 1;
  }
  const parts = marker.parts;
  if (!isRecord(parts) || parts.schema !== DOCUMENT_PARTS_SCHEMA) return 1;
  const count = parts.count;
  return typeof count === "number" && Number.isInteger(count) && count >= 2 && count <= MAX_DOCUMENT_PARTS
    ? count
    : 1;
}






const KNOWN_PARTS_PAGES_KEY = "mrnote:docs:parts-pages:v1";
const KNOWN_PARTS_PAGES_KEPT = 200;
const knownPartsPages = ((): Set<string> => {
  try {
    const raw = typeof window === "undefined" ? null : window.sessionStorage.getItem(KNOWN_PARTS_PAGES_KEY);
    const list: unknown = raw ? JSON.parse(raw) : [];
    return new Set(Array.isArray(list) ? list.filter((id): id is string => typeof id === "string") : []);
  } catch {
    return new Set<string>();
  }
})();












export function documentBodyCutNeedsSaying(stored: unknown, saved: unknown): boolean {
  if (documentPartsCount(stored) < 2 || !isRecord(stored) || !isRecord(saved)) return false;
  const before = Array.isArray(stored.content) ? stored.content.length : 0;
  const after = Array.isArray(saved.content) ? saved.content.length : 0;
  return before - after > PART_JUMP_BLOCKS && after < before / 2;
}

export function knownDocumentPartsPage(pageId: string | null | undefined): boolean {
  return Boolean(pageId && knownPartsPages.has(pageId));
}

export function rememberDocumentPartsPage(pageId: string | null | undefined, parts: boolean): void {
  assertNoStringifiedPageKey(pageId, "document-page-parts rememberDocumentPartsPage");
  if (!pageId || knownPartsPages.has(pageId) === parts) return;
  if (parts) knownPartsPages.add(pageId);
  else knownPartsPages.delete(pageId);
  try {
    window.sessionStorage.setItem(KNOWN_PARTS_PAGES_KEY, JSON.stringify([...knownPartsPages].slice(-KNOWN_PARTS_PAGES_KEPT)));
  } catch {
    // Kept for this page's life only, as before.
  }
}

function withoutLayoutAttrs(attrs: unknown): string {
  if (!isRecord(attrs)) return "{}";
  const kept = { ...attrs };
  for (const key of PRIVATE_LAYOUT_ATTRS) delete kept[key];
  return JSON.stringify(kept);
}

function withoutPrivateLayoutAttrs(attrs: Record<string, unknown>): Record<string, unknown> {
  const kept = { ...attrs };
  for (const key of PRIVATE_LAYOUT_ATTRS) delete kept[key];
  return kept;
}

/** Whether the view's document attributes add nothing to the room's: only unset
 * values the room does not state. An editor fills its schema's defaults in, and
 * an opening that merely read the page wrote them as an edit (`encryptWithPassword:
 * false` on two pages nobody typed in, each moving the page's version on). */
function attrsSayNothingNew(view: Record<string, unknown>, room: Readonly<Record<string, unknown>>): boolean {
  return Object.entries(view).every(([key, value]) =>
    JSON.stringify(value) === JSON.stringify(room[key]) ||
    (!(key in room) && (value === null || value === undefined || value === false || value === "")));
}

function blockId(block: unknown): string | null {
  if (!isRecord(block) || !isRecord(block.attrs)) return null;
  const id = block.attrs.block_id;
  return typeof id === "string" && id ? id : null;
}

/** How far a page held as parts must be loaded for a whole-document draft
 * to be merged into its view (`documentDraftReach`). */
export interface DocumentDraftReach {
  /** Whether the draft says anything `base` does not. */
  readonly changed: boolean;
  /** Load through this block of `base` (the last one the draft changed or
   * deleted, or the one after a block it added); null: nothing to load, or
   * everything (`toEnd`). */
  readonly through: string | null;
  /** The draft added a block after every block of `base`: load every part. */
  readonly toEnd: boolean;
}

/**
 * A draft of the whole document kept for a page now held as parts -- its tab
 * typed while the page was split, or came back to a page split meanwhile
 * with edits it had not saved -- is merged into the parts view block by block
 * against `base`, the body it was edited from (`mergeCanonicalBlocks`). The
 * view holds only the parts loaded so far, so the merge waits until every
 * block it touches is loaded: this says through which. Blocks are matched by
 * identity and compared as authored (`sameAuthoredBlock`); a block without an
 * identity counts as changed.
 */
export function documentDraftReach(base: readonly unknown[], draft: readonly unknown[]): DocumentDraftReach {
  const baseIndex = new Map<string, number>();
  base.forEach((block, index) => {
    const id = blockId(block);
    if (id !== null && !baseIndex.has(id)) baseIndex.set(id, index);
  });
  // For each draft position, the base index of the first block at or after
  // it that base holds (where a block added there goes in the view).
  const nextHeld: number[] = new Array(draft.length + 1);
  nextHeld[draft.length] = -1;
  for (let index = draft.length - 1; index >= 0; index -= 1) {
    const id = blockId(draft[index]);
    const at = id !== null ? baseIndex.get(id) : undefined;
    nextHeld[index] = at !== undefined ? at : nextHeld[index + 1];
  }
  let furthest = -1;
  let toEnd = false;
  let changed = false;
  const kept = new Set<string>();
  draft.forEach((block, index) => {
    const id = blockId(block);
    const at = id !== null ? baseIndex.get(id) : undefined;
    if (id !== null) kept.add(id);
    if (at !== undefined) {
      if (!sameAuthoredBlock(base[at], block)) {
        changed = true;
        furthest = Math.max(furthest, at);
      }
      return;
    }
    changed = true;
    const after = nextHeld[index + 1];
    if (after < 0) toEnd = true;
    else furthest = Math.max(furthest, after);
  });
  base.forEach((block, index) => {
    const id = blockId(block);
    if (id === null || !kept.has(id)) {
      changed = true;
      furthest = Math.max(furthest, index);
    }
  });
  return { changed, toEnd, through: !toEnd && furthest >= 0 ? blockId(base[furthest]) : null };
}

/**
 * How many blocks of the document a draft changed that the view does not
 * hold: each is `base`'s (the body the draft was made on), changed in `draft`
 * as authored (`sameAuthoredBlock`, as documentDraftReach reads it), and not
 * in `view`. In a view still loading its parts such a block is in a part not
 * loaded yet (or another member deleted it): mergeCanonicalBlocks keeps a
 * change over a deletion and puts it after the last block loaded, and its part
 * brings it again when it loads (R482-D04 LIVE). Zero: the merge needs no
 * part it lacks.
 */
export function draftChangesOutsideView<T>(
  base: readonly T[],
  draft: readonly T[],
  view: readonly T[],
  idOf: (block: T) => unknown,
): number {
  const inView = new Set<unknown>();
  for (const block of view) inView.add(idOf(block));
  const was = new Map<unknown, T>();
  for (const block of base) {
    const id = idOf(block);
    if (typeof id === "string" && !was.has(id)) was.set(id, block);
  }
  let outside = 0;
  for (const block of draft) {
    const id = idOf(block);
    if (typeof id !== "string" || inView.has(id)) continue;
    const old = was.get(id);
    if (old !== undefined && !sameAuthoredBlock(old, block)) outside += 1;
  }
  return outside;
}


















export function draftBaseInEditorForm<T>(
  base: readonly T[],
  editorFormOf: ((blocks: readonly T[]) => readonly unknown[] | null) | undefined,
): readonly T[] {
  if (!editorFormOf || base.length === 0) return base;
  let read: readonly unknown[] | null = null;
  try {
    read = editorFormOf(base);
  } catch {
    read = null;
  }
  if (!read || read.length !== base.length) return base;
  for (let index = 0; index < base.length; index += 1) {
    if (blockId(read[index]) !== blockId(base[index])) return base;
  }
  return read as readonly T[];
}

/** How many of `draft`'s blocks `base` holds (by id) written otherwise as
 * authored (`sameAuthoredBlock`), with the first few named by type and Word
 * element: the handover trace's view of what a merge will read as edits. */
export function draftBlocksChangedFrom(base: readonly unknown[], draft: readonly unknown[]): {
  changed: number;
  sample: Array<{ type: unknown; anchor: string | null }>;
} {
  const was = new Map<string, unknown>();
  for (const block of base) {
    const id = blockId(block);
    if (id !== null && !was.has(id)) was.set(id, block);
  }
  let changed = 0;
  const sample: Array<{ type: unknown; anchor: string | null }> = [];
  for (const block of draft) {
    const id = blockId(block);
    const old = id !== null ? was.get(id) : undefined;
    if (old === undefined || sameAuthoredBlock(old, block)) continue;
    changed += 1;
    if (sample.length < 24) sample.push({ type: isRecord(block) ? block.type : null, anchor: sourceAnchorOf(block) });
  }
  return { changed, sample };
}

/** A part's blocks with each block's place by id (null when a block has none
 * or one id stands twice: such a part is not spliced into). */
function blockPlaces(blocks: readonly unknown[]): Map<string, number> | null {
  const places = new Map<string, number>();
  for (let index = 0; index < blocks.length; index += 1) {
    const id = blockId(blocks[index]);
    if (!id || places.has(id)) return null;
    places.set(id, index);
  }
  return places;
}

/** Harness switch (the trace on): how many publishes read the changed sheets only, and how many the whole view. */
function countPublish(kind: "scoped" | "whole"): void {
  if (typeof window === "undefined") return;
  const host = window as Window & { __docsHandoverTrace?: unknown[]; __partRoomPublishes?: Record<string, number> };
  if (!Array.isArray(host.__docsHandoverTrace)) return;
  const counts = host.__partRoomPublishes ?? (host.__partRoomPublishes = {});
  counts[kind] = (counts[kind] ?? 0) + 1;
}

/** How long after this member changed a block another's deletion of it whole leaves it. */
const KEPT_FROM_DELETION_MS = 10_000;

/** The room's change took out whole blocks this member changed a moment ago
 * (and that did not move to another part): the room's blocks with those put
 * back, each after the block before it (as the room had them before), or
 * null when nothing is to be kept. */
function keptFromDeletion(
  live: { mine: Map<string, number> },
  before: readonly unknown[],
  after: readonly unknown[],
  index: number,
  owners?: ReadonlyMap<string, number>,
): unknown[] | null {
  if (live.mine.size === 0) return null;
  const now = Date.now();
  const afterIds = new Set(after.map(blockId));
  const lost = new Set<string>();
  for (const block of before) {
    const id = blockId(block);
    if (!id || afterIds.has(id)) continue;
    const at = live.mine.get(id);
    if (at === undefined || now - at > KEPT_FROM_DELETION_MS) continue;
    if (owners && owners.has(id) && owners.get(id) !== index) continue;
    lost.add(id);
  }
  if (lost.size === 0) return null;
  const out = [...after];
  let previous: string | null = null;
  for (const block of before) {
    const id = blockId(block);
    if (id && lost.has(id)) {
      const at = previous === null ? -1 : out.findIndex((item) => blockId(item) === previous);
      out.splice(at + 1, 0, block);
    }
    if (id && (afterIds.has(id) || lost.has(id))) previous = id;
  }
  return out;
}

/** `text` put into a block's inline content at `offset` (each letter one, each
 * inline object one), taking the marks of the text it lands in. */
function withTextAt(block: Record<string, unknown>, offset: number | null, text: string): Record<string, unknown> {
  const content = Array.isArray(block.content) ? block.content.map((item) => (isRecord(item) ? { ...item } : item)) : [];
  let remain = offset ?? Number.POSITIVE_INFINITY;
  for (let at = 0; at < content.length; at += 1) {
    const item = content[at];
    if (isRecord(item) && item.type === "text" && typeof item.text === "string") {
      if (remain <= item.text.length) {
        item.text = item.text.slice(0, remain) + text + item.text.slice(remain);
        return { ...block, content };
      }
      remain -= item.text.length;
    } else {
      if (remain === 0) {
        content.splice(at, 0, { type: "text", text });
        return { ...block, content };
      }
      remain -= 1;
    }
  }
  const last = content[content.length - 1];
  if (isRecord(last) && last.type === "text" && typeof last.text === "string") last.text += text;
  else content.push({ type: "text", text });
  return { ...block, content };
}

/** The room's blocks with this member's letters a change took out without
 * naming them (WordPartRoom's `lost`) written back where they stood, or null
 * when there are none to write. */
function withLettersKept(
  blocks: readonly unknown[],
  lost: readonly LostLetters[] | undefined,
  /** The room's blocks: a block put back whole from before the change holds its letters already. */
  roomBlocks: readonly unknown[],
): unknown[] | null {
  if (!lost || lost.length === 0) return null;
  const inRoom = new Set(roomBlocks.map(blockId));
  let changed = false;
  const out = blocks.map((block) => {
    const id = blockId(block);
    if (!id || !inRoom.has(id) || !isRecord(block)) return block;
    const runs = lost.filter((entry) => entry.block === id);
    if (runs.length === 0) return block;
    changed = true;
    // From the end backwards, so each offset is read before anything is put in front of it.
    const ordered = [...runs.filter((entry) => entry.offset === null), ...runs.filter((entry) => entry.offset !== null)
      .sort((a, b) => (b.offset as number) - (a.offset as number))];
    return ordered.reduce<Record<string, unknown>>((current, entry) => withTextAt(current, entry.offset, entry.text), block);
  });
  return changed ? out : null;
}

/** The size of a text block's inline content as the editor counts it: each
 * letter one, each inline object one. */
function inlineSizeOf(block: unknown): number {
  const content = isRecord(block) && Array.isArray(block.content) ? block.content : [];
  let size = 0;
  for (const item of content) {
    if (isRecord(item) && item.type === "text" && typeof item.text === "string") size += item.text.length;
    else size += 1;
  }
  return size;
}

/** A block's JSON, kept for the very object (a part's unchanged blocks are the
 * same objects publish after publish). */
const blockJson = new WeakMap<object, string>();
function jsonOfBlock(block: unknown): string {
  if (!block || typeof block !== "object") return JSON.stringify(block);
  let json = blockJson.get(block);
  if (json === undefined) {
    json = JSON.stringify(block);
    blockJson.set(block, json);
  }
  return json;
}

/** Whether the handover trace is on (the harness set `window.__docsHandoverTrace = []`). */
function docsTraceOn(): boolean {
  return typeof window !== "undefined" && Array.isArray((window as Window & { __docsHandoverTrace?: unknown[] }).__docsHandoverTrace);
}

/** The Word element a top-level block was read from (`source.sourceAnchor`), or null. */
function sourceAnchorOf(block: unknown): string | null {
  const source = isRecord(block) && isRecord(block.attrs) && isRecord(block.attrs.source) ? block.attrs.source : null;
  return source && typeof source.sourceAnchor === "string" && source.sourceAnchor ? source.sourceAnchor : null;
}

/** Add (or, with `by` -1, take back) how often each Word element stands in `blocks`. */
export function countSourceAnchors(blocks: readonly unknown[], into: Map<string, number>, by = 1): void {
  for (const block of blocks) {
    const anchor = sourceAnchorOf(block);
    if (anchor) into.set(anchor, (into.get(anchor) ?? 0) + by);
  }
}

/** How many times `blocks` hold a Word element beyond what the stored
 * document held it (`stored`, countSourceAnchors; once for an element not in
 * it). Above zero, some blocks stand twice: a merge or a load put a block of
 * the document in again (R482-D04 LIVE). */
export function anchorsBeyondStored(blocks: readonly unknown[], stored: ReadonlyMap<string, number>): number {
  const seen = new Map<string, number>();
  let beyond = 0;
  for (const block of blocks) {
    const anchor = sourceAnchorOf(block);
    if (!anchor) continue;
    const count = (seen.get(anchor) ?? 0) + 1;
    seen.set(anchor, count);
    if (count > Math.max(1, stored.get(anchor) ?? 0)) beyond += 1;
  }
  return beyond;
}

/** The top-level source anchors `blocks` hold more than once (a Word element stored twice). */
function repeatedSourceAnchors(blocks: readonly unknown[]): number {
  return anchorsBeyondStored(blocks, new Map());
}

/** How many of `a`'s block ids `b` holds (a trace of how two readings of a part line up). */
function sharedIds(a: readonly unknown[], b: readonly unknown[]): number {
  const ids = new Set(b.map(blockId).filter((id): id is string => id !== null));
  return a.reduce<number>((count, item) => count + (ids.has(blockId(item) ?? "\u0000") ? 1 : 0), 0);
}

/** Harness switch (`window.__docsHandoverTrace = []`, NoteEditor's): each part edit and block. */
function traceDocsHandoverPart(entry: Record<string, unknown>): void {
  const sink = typeof window === "undefined" ? null : (window as Window & { __docsHandoverTrace?: unknown[] }).__docsHandoverTrace;
  if (Array.isArray(sink) && sink.length < 400) sink.push({ t: Math.round(performance.now()), ...entry });
}

/** Harness switch (`window.__docsTraceWords = [...]`, with the trace on): which of
 * these words `blocks` hold, for following a word through the part rooms'
 * publish, remote change and read. Undefined without it (no work). */
function tracedWords(blocks: readonly unknown[] | null | undefined): string[] | undefined {
  const words = typeof window === "undefined" ? null : (window as Window & { __docsTraceWords?: unknown }).__docsTraceWords;
  if (!Array.isArray(words) || words.length === 0 || !blocks) return undefined;
  const text = JSON.stringify(blocks);
  return words.filter((word): word is string => typeof word === "string" && text.includes(word));
}
const tracedPublishes = new Map<number, string>();

/**
 * Whether two readings of a part say the same (sameDocumentAuthorContent): what
 * opening it fills in by itself (a heading's anchor, ids on nested items, schema
 * defaults, a mark order) is not an edit, and a save of it wrote a part nobody
 * had touched as soon as anything else in the document was typed.
 */
function samePartContent(left: readonly unknown[], right: readonly unknown[]): boolean {
  try {
    return sameDocumentAuthorContent({ content: [...left] }, { content: [...right] });
  } catch {
    return false;
  }
}

/**
 * samePartContent(blocks, the saved part), block by block: a block whose JSON
 * is the saved block's own is the same, and only the others are read again
 * and compared (the comparison is block by block anyway). Each pause in
 * typing read the whole saved part back (JSON.parse) and keyed every block of
 * it to tell the one part being typed into had changed: 0.48 s of 10 s of
 * typing on R286-D01 (124 pages, CPU profile kh32). The saved part as a whole
 * where its blocks are not known, or are not the very blocks it was made of.
 * `window.__docsPartCompareByBlock = false`: the whole part every time (old).
 */
export function samePartAsSaved(
  blocks: readonly unknown[],
  each: readonly string[],
  saved: string,
  savedBlocks: readonly string[] | null,
): boolean {
  const byBlock = typeof window === "undefined" || (window as Window & { __docsPartCompareByBlock?: boolean }).__docsPartCompareByBlock !== false;
  if (byBlock && savedBlocks && savedBlocks.length === each.length && `[${savedBlocks.join(",")}]` === saved) {
    for (let index = 0; index < each.length; index += 1) {
      if (each[index] === savedBlocks[index]) continue;
      if (!samePartContent([blocks[index]], [JSON.parse(savedBlocks[index]!) as unknown])) return false;
    }
    return true;
  }
  return samePartContent(blocks, JSON.parse(saved) as unknown[]);
}

export type DocumentPartsSplit =
  | { readonly ok: true; readonly parts: readonly (readonly unknown[])[] }
  | { readonly ok: false; readonly reason: DocumentPartsBlock };

/**
 * The document's top-level blocks, back in the parts they belong to:
 * `parts[k]` for each part 0..loaded. `owner` names the part each block was
 * loaded or last saved in. A block not in it follows the block before it.
 */
export function splitDocumentParts(
  content: unknown,
  owner: ReadonlyMap<string, number>,
  loaded: number,
): DocumentPartsSplit {
  if (!isRecord(content) || !Array.isArray(content.content)) return { ok: false, reason: "unreadable" };
  const parts: unknown[][] = Array.from({ length: loaded + 1 }, () => []);
  const seen = new Set<string>();
  let current = 0;
  for (const block of content.content) {
    const id = blockId(block);
    if (id !== null && !seen.has(id)) {
      seen.add(id);
      const known = owner.get(id);
      // A block of a later loaded part starts that part. One of an earlier
      // part (moved across an edge) or of a part not loaded stays where it
      // now sits, as a block met twice does.
      if (known !== undefined && known > current && known <= loaded) current = known;
    }
    parts[current].push(block);
  }
  return { ok: true, parts: fillEmptyParts(parts) };
}

/** Every part holds a block (the server keeps no empty part): a part left
 * with none takes the first block of the next part that can spare one, else
 * the last of the part before it, else an empty paragraph. Document order is
 * kept; only where one part ends moves. */
function fillEmptyParts(parts: unknown[][]): unknown[][] {
  for (let index = 0; index < parts.length; index += 1) {
    if (parts[index].length > 0) continue;
    let later = index + 1;
    while (later < parts.length && parts[later].length === 0) later += 1;
    if (later < parts.length && parts[later].length > 1) {
      parts[index].push(parts[later].shift());
      continue;
    }
    const earlier = index - 1;
    if (earlier >= 0 && parts[earlier].length > 1) {
      parts[index].unshift(parts[earlier].pop());
      continue;
    }
    parts[index].push({ type: "paragraph" });
  }
  return parts;
}

/**
 * The editor root's `data-document-parts-loading` / `data-document-parts-left`
 * (read by the pages' ledger, which holds its structural rebalance while parts
 * come in): a later part being read, and how many are not in the view yet.
 * Null -- no attributes -- for a page not held as parts, and once every part
 * is in and none is being read.
 */
export function documentPartsArrivingAttrs(
  partsPage: boolean,
  state: Pick<DocumentPartsState, "count" | "loaded" | "loading" | "eager">,
): { loading: "true" | "false"; left: string; eager: "true" | "false" } | null {
  if (!partsPage) return null;
  const left = Math.max(0, state.count - 1 - state.loaded);
  if (left === 0 && !state.loading) return null;
  return { loading: state.loading ? "true" : "false", left: String(left), eager: state.eager ? "true" : "false" };
}

/** Whether a part coming in is read part by part (`serializeParts`):
 * `window.__documentPartsScopedLoad = true` (off until measured on R397-D28). */
function scopedPartsLoad(): boolean {
  return typeof window !== "undefined" && (window as Window & { __documentPartsScopedLoad?: boolean }).__documentPartsScopedLoad === true;
}

/** How far at once a part's block count may move with no hand behind it. */
const PART_JUMP_BLOCKS = 20;
/** How long a gesture that can take or bring many blocks at once answers for such a move. */
const PART_JUMP_GESTURE_MS = 60_000;

/**
 * Whether a part's blocks moved by more than an edit moves them unless someone
 * meant it: over 20 blocks, and under half or over one and a half times what it
 * held. The count is the check because a duplicated part comes back with new ids
 * (nothing repeats), and a part emptied by a mis-split still splits "ok".
 */











export function withPartsLoadedSince<T>(
  base: readonly T[],
  draft: readonly T[],
  view: readonly T[],
  idOf: (block: T) => unknown,
): T[] {
  const baseIds = new Set<string>();
  for (const block of base) {
    const id = idOf(block);
    if (typeof id === "string") baseIds.add(id);
  }
  const viewById = new Map<string, T>();
  for (const block of view) {
    const id = idOf(block);
    if (typeof id === "string" && !viewById.has(id)) viewById.set(id, block);
  }
  const start = "\u0000start";
  const after = new Map<string, T[]>();
  let anchor = start;
  let added = 0;
  for (const block of draft) {
    const id = idOf(block);
    if (typeof id !== "string") continue;
    if (baseIds.has(id)) { anchor = id; continue; }
    const held = viewById.get(id);
    if (held === undefined) continue;
    const list = after.get(anchor);
    if (list) list.push(held); else after.set(anchor, [held]);
    added += 1;
  }
  if (!added) return [...base];
  const out: T[] = [...(after.get(start) ?? [])];
  for (const block of base) {
    out.push(block);
    const id = idOf(block);
    const list = typeof id === "string" ? after.get(id) : undefined;
    if (list) { out.push(...list); after.delete(id as string); }
  }
  return out;
}











export function withStoredIdsForEmptyBlocks<T>(stored: readonly T[], view: readonly T[]): T[] {
  const idOf = (item: unknown) => {
    const id = isRecord(item) && isRecord(item.attrs) ? item.attrs.block_id : undefined;
    return typeof id === "string" ? id : null;
  };
  // No word in it (a text node holds at least one character).
  const empty = (item: unknown): boolean => !/"text":"[^"]/.test(JSON.stringify(item));
  const withId = (item: unknown, id: string | null) => isRecord(item) && isRecord(item.attrs)
    ? { ...item, attrs: { ...item.attrs, block_id: id } } : item;
  const storedIds = new Set(stored.map(idOf).filter((id): id is string => id !== null));
  const viewIds = new Set(view.map(idOf).filter((id): id is string => id !== null));
  let changed = false;
  const out = view.map((item, index) => {
    const own = idOf(item);
    const there = stored[index];
    const theirs = idOf(there);
    if (theirs === null || own === theirs || (own !== null && storedIds.has(own)) || viewIds.has(theirs)) return item;
    if (!empty(item) || !empty(there) || !sameAuthoredBlock(withId(item, null), withId(there, null))) return item;
    changed = true;
    return withId(item, theirs) as T;
  });
  return changed ? out : [...view];
}

export function partCountJumped(before: number, now: number): boolean {
  if (Math.abs(now - before) <= PART_JUMP_BLOCKS) return false;
  return now < before / 2 || now > before * 1.5;
}

const REFUSED_TEXT_PREFIX = "mrnote:document-parts-refused:v1:";
const REFUSED_TEXT_MAX_CHARS = 4_000;
const REFUSED_TEXT_WINDOW_MS = 10 * 60_000;

/**
 * What a reader typed lately, kept when the save gate stops a view. Text only,
 * and only to be offered back: the view it came from is not to be trusted, so
 * nothing here enters the document or a save by itself.
 */
export function readRefusedPartText(pageId: string): string | null {
  try {
    const raw = window.localStorage.getItem(REFUSED_TEXT_PREFIX + pageId);
    const read = raw ? (JSON.parse(raw) as { text?: unknown }) : null;
    return typeof read?.text === "string" && read.text ? read.text : null;
  } catch {
    return null;
  }
}
function writeRefusedPartText(pageId: string, text: string | null): void {
  try {
    if (!text) window.localStorage.removeItem(REFUSED_TEXT_PREFIX + pageId);
    else window.localStorage.setItem(REFUSED_TEXT_PREFIX + pageId, JSON.stringify({ state: "refused-by-gate-unverified", at: Date.now(), text }));
  } catch {
    // A full or closed store: the report still says what happened.
  }
}

/** The top-level ids a document holds more than once. */
function repeatedBlockIds(blocks: readonly unknown[]): number {
  const seen = new Set<string>();
  let repeated = 0;
  for (const block of blocks) {
    const id = blockId(block);
    if (!id) continue;
    if (seen.has(id)) repeated += 1;
    seen.add(id);
  }
  return repeated;
}











export function storedFormOfBlocks(blocks: readonly unknown[], schema: Schema): unknown[] {
  const droppable = new Map<string, ReadonlySet<string>>();
  const nullDefaults = (kind: "node" | "mark", type: string): ReadonlySet<string> | null => {
    const key = `${kind}:${type}`;
    const known = droppable.get(key);
    if (known) return known;
    const spec = kind === "node" ? schema.nodes[type] : schema.marks[type];
    if (!spec) return null;
    const names = new Set<string>();
    const attrs = (spec as unknown as { attrs?: Record<string, { hasDefault?: boolean; default?: unknown }> }).attrs ?? {};
    for (const [name, attr] of Object.entries(attrs)) {
      if (attr.hasDefault && attr.default === null) names.add(name);
    }
    droppable.set(key, names);
    return names;
  };
  const strip = (value: unknown, kind: "node" | "mark"): unknown => {
    if (!isRecord(value) || typeof value.type !== "string") return value;
    let out: Record<string, unknown> = value;
    const names = isRecord(value.attrs) ? nullDefaults(kind, value.type) : null;
    if (names && names.size && isRecord(value.attrs) && Object.entries(value.attrs).some(([name, item]) => item === null && names.has(name))) {
      const attrs = Object.fromEntries(Object.entries(value.attrs).filter(([name, item]) => !(item === null && names.has(name))));
      out = { ...value };
      if (Object.keys(attrs).length) out.attrs = attrs;
      else delete out.attrs;
    }
    if (kind === "node" && Array.isArray(value.content)) {
      const content = value.content.map((item) => strip(item, "node"));
      if (content.some((item, index) => item !== (value.content as unknown[])[index])) out = { ...out, content };
    }
    if (kind === "node" && Array.isArray(value.marks)) {
      const marks = value.marks.map((item) => strip(item, "mark"));
      if (marks.some((item, index) => item !== (value.marks as unknown[])[index])) out = { ...out, marks };
    }
    return out;
  };
  return blocks.map((block) => strip(block, "node"));
}

/** The page body to save: the document's attrs (without the loaded
 * document's layout receipts) over the body's blocks, with every top-level
 * key the stored body has beyond them (the parts marker above all). */
export function bodyContentForSave(
  merged: Record<string, unknown>,
  storedBody: Record<string, unknown>,
  bodyBlocks: readonly unknown[],
): Record<string, unknown> {
  const attrs = isRecord(merged.attrs) ? { ...merged.attrs } : null;
  const storedAttrs = isRecord(storedBody.attrs) ? storedBody.attrs : {};
  if (attrs) {
    // The loaded whole's layout receipts are not the body's: the body keeps
    // the ones it has (an unedited body is then the stored body exactly).
    for (const key of PRIVATE_LAYOUT_ATTRS) {
      if (key in storedAttrs) attrs[key] = storedAttrs[key];
      else delete attrs[key];
    }
  }
  const extra = Object.fromEntries(
    Object.entries(storedBody).filter(([key]) => key !== "type" && key !== "attrs" && key !== "content"),
  );
  return {
    ...merged,
    ...extra,
    ...(attrs ? { attrs } : {}),
    content: [...bodyBlocks],
  };
}

/** `GET /pages/{id}/document-parts/{index}`, for a document's part and a
 * deck's alike (a deck's names its `status`: a part the import has not read,
 * or could not, holds no content yet). */
export interface DocumentPartResponse {
  readonly pageId?: string;
  readonly partIndex: number;
  readonly partCount: number;
  readonly partId: string;
  readonly contentSha256?: string;
  readonly contentVersion?: number;
  readonly revision?: number;
  readonly status?: "ready" | "pending" | "failed";
  readonly reason?: string | null;
  /** The part is saved on its own (`PUT`) from an editor that may write. */
  readonly editable?: boolean;
  readonly content: {
    readonly type?: unknown;
    readonly content?: unknown;
    readonly slides?: unknown;
    readonly media?: unknown;
    readonly deckReceipts?: unknown;
  } | null;
}

/** One later part of a page held as parts. `init` for a reader that must
 * see the part as the server holds it now (a deck's part the import is
 * still writing: `cache: "no-store"`). */
export function readDocumentPart(pageId: string, index: number, init?: RequestInit): Promise<DocumentPartResponse> {
  const path = `/api/v1/pages/${encodeURIComponent(pageId)}/document-parts/${index}`;
  return init ? apiGet<DocumentPartResponse>(path, init) : apiGet<DocumentPartResponse>(path);
}

interface DocumentPartReceipt {
  readonly partIndex: number;
  readonly revision: number;
}

export type DocumentPartsBlock = "order" | "empty-part" | "duplicate-block" | "unreadable" | "count-jump";

export interface DocumentPartsState {
  /** How many parts the page holds (1: an ordinary page). */
  readonly count: number;
  /** How many later parts the editor holds now. */
  readonly loaded: number;
  /** A later part is being read. */
  readonly loading: boolean;
  /** Later parts with edits not yet saved, or being saved. */
  readonly saving: boolean;
  /** Why the document cannot be saved part by part (read-only meanwhile). */
  readonly blocked: DocumentPartsBlock | "conflict" | null;
  /** The last part save failed and is being retried. */
  readonly retrying: boolean;
  /** The import is still reading the next part (it opened from the first),
   * or could not. */
  readonly importing?: "pending" | "failed" | null;
  /** Typed in a view the save gate stopped, offered back and never applied. */
  readonly refusedText?: string | null;
  /** Every part is being read at once (loadsPartsAtOnce). */
  readonly eager?: boolean;
}

export interface UseDocumentPagePartsOptions {
  readonly pageId: string | null | undefined;
  /** The page body as saved (`autosave.page.content_json`). */
  readonly body: unknown;
  /** Changes with each hydration of the editor (a reload starts over). */
  readonly hydration: number;
  /** The editor holds the hydrated body. */
  readonly hydrated: boolean;
  /** The parts are saved from this view (the page is writable). */
  readonly writable: boolean;
  /** Append a part's blocks (as stored) at the end of the document; false
   * when the editor cannot take them now. */
  readonly appendBlocks: (blocks: readonly unknown[], partIndex?: number, keptLayout?: KeptPartLayout | null) => boolean;
  /** The pages kept for a part (paged/layoutSnapshot), asked before it is
   * appended, handed to `appendBlocks`. */
  readonly keptLayoutOf?: (blocks: readonly unknown[], partIndex: number) => Promise<KeptPartLayout | null>;
  /** The editor's pages are still being laid out (the first layout, the
   * background's, a pass waiting for typing to pause): a read-ahead part waits
   * for them rather than for the end of what is loaded to hold still twice. */
  readonly layoutBusy?: () => boolean;
  /** The editor's first screen is still to come (its first layout settling,
   * or the display font plan still to answer): the parts read at once
   * (loadsPartsAtOnce) are read and go in after it, unless the reader goes
   * further first. */
  readonly firstScreenPending?: () => boolean;
  /** The whole document as a save serializes it, or null. */
  readonly serialize: () => Record<string, unknown> | null;
  /** The view serialized for the part rooms alone: its blocks as `serialize`
   * gives them, without the page's layout records (measuring every sheet of a
   * long document for a receipt the rooms never carry). Falls back to
   * `serialize`. */
  readonly serializeForRooms?: () => Record<string, unknown> | null;
  




  readonly serializeParts?: (owner: ReadonlyMap<string, number>, loaded: number) => unknown[][] | null;
  /** Stored blocks as this editor serializes them (what `serialize` would
   * give for them), or null. */
  readonly editorFormOf?: (blocks: readonly unknown[]) => unknown[] | null;
  /** The block ids the view holds now (each block's, or its first piece's for
   * a block cut across sheets), read from the editor's document without
   * serializing it; null when there is no view. A part coming in is checked
   * against them (see loadNext). */
  readonly viewBlockIds?: () => ReadonlySet<string> | null;
  /** Make the editor's blocks `current` (a part's, by block id) read
   * `merged`; false when it cannot (nothing is changed then). */
  readonly mergePartIntoView?: (
    current: readonly unknown[],
    merged: readonly unknown[],
    /** Where the caret goes, when the room could follow it. */
    caret?: { readonly block: string; readonly offset: number },
  ) => boolean;
  /** Edit in live rooms: the view's means of being one room's client. Absent:
   * the parts are saved as above, and no room is asked for. */
  readonly partRooms?: DocumentPartRoomsView | null;
  /** Blocks as they are stored (`storedFormOfBlocks` with the editor's
   * schema): what a part's PUT and the body's save send. Absent: as the
   * editor serializes them. */
  readonly storedForm?: (blocks: readonly unknown[]) => unknown[];
  




  readonly hydrationToken?: () => unknown;
}

export interface DocumentPartRoomsView {
  /** The flat schema of a room's blocks; null until the editor exists. */
  readonly schema: () => Schema | null;
  /** The view's document (its identity says whether anything changed). */
  readonly viewDoc: () => unknown;
  /** The view is composing (IME): nothing is read from it or drawn into it. */
  readonly composing: () => boolean;
  /** Make the view's document attributes read these; only those it has. */
  readonly applyAttrs?: (attrs: Readonly<Record<string, unknown>>) => void;
  /** The body's room saved it into the page: the page's version now. */
  readonly bodySaved?: (contentVersion: number) => void;
  /** Everything this view holds is in its rooms, as far as they are in use. */
  readonly settled?: () => void;
  /** Another member said where they are (the body's room carries it). */
  readonly presence?: (bytes: Uint8Array) => void;
  /** This member's body room was (re)joined: they say where they are. */
  readonly joined?: () => void;
  /** The caret: the block it is in (`block_id`) and its offset in the block's
   * inline content; null when not in one block's text. */
  readonly caret?: () => { readonly block: string; readonly offset: number } | null;
  /** The blocks of the part of the view that changed between two of its
   * documents (`viewDoc`), whole blocks, serialized as `serializeForRooms`
   * gives them: `before` from `previous`, `after` from `next`, the blocks
   * around them the same in both. Null when that cannot be told cheaply (the
   * whole view is serialized then). */
  readonly changedBlocks?: (previous: unknown, next: unknown) => { before: unknown[]; after: unknown[] } | null;
  



  readonly partBlocks?: (index: number, owner: ReadonlyMap<string, number>, loaded: number) => unknown[] | null;
}

/** One live room this view holds. */
interface LivePartRoom {
  readonly index: number;
  readonly room: WordPartRoom;
  /** When the part's input started waiting for this room (it came into the view). */
  askedAt: number;
  /** The part as last published, JSON; null when the view was last changed
   * by the room (the next look compares with the room itself). */
  viewJson: string | null;
  viewAttrs: string | null;
  /** The part as it and the room last agreed in a save: what a fresh reading
   * of the room is merged against. */
  base: readonly unknown[];
  /** The part is in the view. A later part's room is joined as soon as the
   * part is read, while its blocks are still being laid out; until they are
   * in, its reading and the others' changes wait (nothing to merge them into). */
  inView: boolean;
  /** The room was read before the part was in the view: read once it is. */
  readWaiting: boolean;
  /** The part as the view last published it or was drawn from the room
   * (null: not known block by block), with each block's place by id: what a
   * key's publish splices the changed blocks into. */
  viewBlocks: readonly unknown[] | null;
  viewPlaces: Map<string, number> | null;
  /** When this member last changed each block of the part (by id). */
  mine: Map<string, number>;
}

/** What a closing window holds of a document held as parts (closingParts). */
export interface DocumentPartsClosing {
  readonly body: readonly unknown[];
  readonly attrs: unknown;
  /** The body's blocks as the editor read them when it opened the page. */
  readonly openedBody: readonly unknown[] | null;
  /** The body's room holds every word of it: nothing of it is to be sent. */
  readonly bodyInRoom?: boolean;
  /** The view could not be published to its rooms as the window went (an IME
   * composition): what it holds beyond what they acknowledged is in no room,
   * and a room's acknowledgement says nothing of it. Not forced out (a word
   * half composed), it is left to the window's AutoRecover copy, which is
   * kept. */
  readonly viewUnpublished?: true;
  readonly parts: ReadonlyArray<{
    readonly index: number;
    readonly current: readonly unknown[];
    readonly saved: readonly unknown[] | null;
    readonly sending: readonly unknown[] | null;
    /** The revision the save on its way was made from (null: none on its way). */
    readonly sendingRevision: number | null;
    /** Its reader cut or grew it at once in this reading, as a part's PUT says
     * (`large_removal`): the server refuses such a cut said by no one. */
    readonly largeRemoval?: boolean;
  }>;
}

export interface DocumentPageParts {
  readonly state: DocumentPartsState;
  /** See DocumentPartsClosing. */
  readonly closingParts: () => DocumentPartsClosing | null;
  /** The view's serialization has just been made: what it holds goes to the
   * rooms now. They need none of the page save's debounce, and a word that
   * waited for it was lost with a window gone in those seconds without a page
   * hide (and showed to the others seconds late). */
  readonly publishEdits: () => void;
  /** The element the next part loads when it comes near the viewport. */
  readonly sentinelRef: (element: HTMLElement | null) => void;
  /** A page draft as the page saves it: the body alone. Later parts that
   * changed are saved on their own. Unchanged for an ordinary page. Null: the
   * body's own room holds it, and the page's save sends nothing of it.
   * Undefined: not yet (the view is not this reading's body as opened): sent
   * now, nothing would be saved of it while the save said it was; asked
   * again a moment later. */
  readonly prepareContent: (content: Record<string, unknown>) => Record<string, unknown> | null | undefined;
  /** Whether a body cut prepared by this view is said meant (`large_removal`). */
  readonly bodyCutSaidMeant: () => boolean;
  /** Tell the page's other members where this one is. */
  readonly sendPresence: (bytes: Uint8Array) => void;
  /** Load the parts up to the one holding top-level block `id`, in order
   * (the place a reader was at before the page was split into parts); false
   * when no part holds it or loading stopped. */
  readonly loadThroughBlock: (id: string) => Promise<boolean>;
  /** Edits entering the view that no key here made (keys held through a reload). */
  readonly noteHand: () => void;
  /** The view's document changed: its rooms hear of it shortly. */
  readonly noteViewChanged: () => void;
  readonly dismissRefusedText: () => void;
  /** Whether the view holds the document as stored now, part by part: the
   * body as `storedBody` says it, and every later part it has read as the part
   * is stored (a part saved since this view read it -- a closing tab's edits
   * landing after it loaded -- must say in the view what it says stored). A
   * part a live room holds is the room's to bring. Null: not known (a part
   * could not be read, or the view is not this reading). */
  readonly viewHoldsStored: (storedBody: unknown) => Promise<boolean | null>;
  /** Whether what the view holds is the reading these parts were reset for:
   * false between a reload's content going into the view and the parts
   * starting over for it (the parts loaded, and their count, are then the
   * reading before's). */
  readonly viewIsThisReading: () => boolean;
  /** Blocks about to go into the view holding a Word element more often than
   * the stored document: the view is stopped and true is said (nothing of
   * them may go in); false when they hold each at most as often. */
  readonly refuseBlocksTwice: (blocks: readonly unknown[], what: string) => boolean;
}

interface PartRecord {
  partId: string;
  revision: number;
  /** The part's blocks as last saved (or as first serialized), JSON. */
  saved: string;
  /** The same blocks one by one (null: not known block by block, as after a
   * recovered edit not yet saved), and how many the stored part holds. */
  savedBlocks: readonly string[] | null;
  baseCount: number;
  /** A recovered edit's part as stored, in the editor's form: what a conflict over
   * it is merged against. (Against nothing, every block read as changed on both
   * sides, and a block that differed in what the two forms fill in came back twice.) */
  recoveredBase?: readonly unknown[] | null;
}

/**
 * A later part's edit not yet saved, kept in this tab's session storage as
 * the page body's unsaved draft is: a reload (or a crash of the page) before
 * its save reached the server brings it back. It is the stored part at
 * `rev` with `del` blocks from `start` replaced by `insert` (blocks as the
 * editor serializes them); `sending` says the save of exactly this content
 * was on its way.
 */
export interface DocumentPartRecovery {
  readonly rev: number;
  readonly count: number;
  readonly start: number;
  readonly del: number;
  readonly insert: readonly unknown[];
  readonly sending: boolean;
}

const PART_RECOVERY_KEY_PREFIX = "mrnote:document-part-recovery:v1:";

function partRecoveryKey(pageId: string): string {
  return `${PART_RECOVERY_KEY_PREFIX}${pageId}`;
}

export function readDocumentPartRecoveries(pageId: string): Record<string, DocumentPartRecovery> {
  try {
    const raw = window.sessionStorage.getItem(partRecoveryKey(pageId));
    const parsed: unknown = raw ? JSON.parse(raw) : null;
    return isRecord(parsed) ? (parsed as Record<string, DocumentPartRecovery>) : {};
  } catch {
    return {};
  }
}

export function writeDocumentPartRecovery(pageId: string, index: number, entry: DocumentPartRecovery | null): void {
  try {
    const all = readDocumentPartRecoveries(pageId);
    if (entry) all[String(index)] = entry;
    else delete all[String(index)];
    const key = partRecoveryKey(pageId);
    if (Object.keys(all).length === 0) window.sessionStorage.removeItem(key);
    else window.sessionStorage.setItem(key, JSON.stringify(all));
  } catch {
    // Storage full or unavailable: the edit is still saved as usual, it only
    // would not come back after a reload before that.
  }
}

/** The edit that turns the stored part (`saved`, block by block, or only
 * its block count when not known block by block) into `current`. */
export function documentPartRecovery(
  saved: readonly string[] | null,
  baseCount: number,
  current: readonly string[],
  rev: number,
  sending: boolean,
): DocumentPartRecovery {
  if (!saved) {
    return { rev, count: baseCount, start: 0, del: baseCount, insert: current.map((json) => JSON.parse(json)), sending };
  }
  let start = 0;
  while (start < saved.length && start < current.length && saved[start] === current[start]) start += 1;
  let end = 0;
  while (
    end < saved.length - start && end < current.length - start &&
    saved[saved.length - 1 - end] === current[current.length - 1 - end]
  ) end += 1;
  return {
    rev,
    count: saved.length,
    start,
    del: saved.length - start - end,
    insert: current.slice(start, current.length - end).map((json) => JSON.parse(json)),
    sending,
  };
}

/** The stored part's blocks with a kept edit applied, or null when the edit
 * does not belong to these blocks. */
export function applyDocumentPartRecovery(
  stored: readonly unknown[],
  entry: DocumentPartRecovery,
): unknown[] | null {
  if (
    !isRecord(entry) || !Array.isArray(entry.insert) || entry.count !== stored.length ||
    !Number.isInteger(entry.start) || !Number.isInteger(entry.del) ||
    entry.start < 0 || entry.del < 0 || entry.start + entry.del > stored.length
  ) return null;
  const blocks = [...stored.slice(0, entry.start), ...entry.insert, ...stored.slice(entry.start + entry.del)];
  return blocks.length > 0 ? blocks : null;
}

/**
 * A kept edit of a part saved since from elsewhere (its revision moved on, so
 * `applyDocumentPartRecovery` cannot place it), put into the part as it is
 * now with nothing of either side lost: each block the edit wrote goes in --
 * in its own block's place when the part holds it (taken whole when it only
 * typed into the stored one, left when the stored one already holds it,
 * kept beside it as a copy when both changed it), after the block before it
 * otherwise. What the edit deleted stays (the block may be another member's
 * since). It used to be dropped, and the page made read-only ("conflict"):
 * every word of it lost on a reload.
 */
export function mergeDocumentPartRecovery(
  stored: readonly unknown[],
  entry: DocumentPartRecovery,
): { blocks: unknown[]; copies: number } | null {
  if (!isRecord(entry) || !Array.isArray(entry.insert) || !Number.isInteger(entry.start)) return null;
  const blocks = [...stored];
  let copies = 0;
  let after = Math.min(Math.max(0, entry.start), blocks.length) - 1;
  for (const item of entry.insert) {
    const id = blockId(item);
    const at = id ? blocks.findIndex((block) => blockId(block) === id) : -1;
    if (at < 0) {
      blocks.splice(after + 1, 0, item);
      after += 1;
      continue;
    }
    const held = blocks[at];
    if (sameAuthoredBlock(held, item) || grownByTyping(item, held)) {
      after = at;
    } else if (grownByTyping(held, item)) {
      blocks[at] = item;
      after = at;
    } else {
      const copy = structuredClone(item) as Record<string, unknown>;
      if (isRecord(copy.attrs)) copy.attrs = { ...copy.attrs, block_id: globalThis.crypto?.randomUUID?.() ?? `kept-${Date.now().toString(36)}-${copies}` };
      blocks.splice(at + 1, 0, copy);
      copies += 1;
      after = at + 1;
    }
  }
  return blocks.length > 0 ? { blocks, copies } : null;
}

/**
 * A document held as parts in the editor: its later parts loaded as the
 * reader nears the end, and (writable) saved part by part.
 */
/** What one look at the sentinel reads. */
export interface PartLoadReading {
  /** The sentinel's top, viewport-relative (`getBoundingClientRect().top`). */
  readonly sentinelTop: number;
  readonly viewportHeight: number;
  /** The scroll position of the sentinel's scroll container. */
  readonly scrollTop: number;
}

/** What the loader keeps between looks. */
export interface PartLoadMemory {
  /** The sentinel's top at the previous look, or null before the first. */
  readonly settledTop: number | null;
  /** The scroll position the last look-ahead load was made at, or null. */
  readonly loadedAtScroll: number | null;
}

export const EMPTY_PART_LOAD_MEMORY: PartLoadMemory = { settledTop: null, loadedAtScroll: null };

/** The look-ahead below the viewport within which the next part loads. */
export const PART_LOAD_AHEAD_PX = 2400;

/**
 * Whether this look at the sentinel loads the next part.
 *
 * A part loads when the end of what is loaded comes near the viewport --
 * and only then. Two things kept parts loading on their own: a load that
 * finished while the sentinel still read "near" queued the next at once,
 * and the sentinel is read before the appended part is laid out, so it
 * reads near again (AI70H H18 B16: a 15,600-paragraph specification loaded
 * every part back to back for four minutes while nobody scrolled, and the
 * AI panel could not take a keystroke). So the sentinel must have stayed
 * where it was since the previous look (the appended part's layout has
 * settled), and a look-ahead load -- the sentinel below the viewport, not
 * yet in it -- is made once per scroll position: the reader moves, the
 * next part follows.
 */
/** How soon a part the import is still reading is asked for again. */
const IMPORTING_PART_RETRY_MS = 1200;

export function shouldLoadNextPart(
  memory: PartLoadMemory,
  reading: PartLoadReading,
): { load: boolean; memory: PartLoadMemory } {
  const settled = memory.settledTop === reading.sentinelTop;
  const next: PartLoadMemory = { settledTop: reading.sentinelTop, loadedAtScroll: memory.loadedAtScroll };
  if (reading.sentinelTop >= reading.viewportHeight + PART_LOAD_AHEAD_PX) return { load: false, memory: next };
  if (!settled) return { load: false, memory: next };
  const inView = reading.sentinelTop < reading.viewportHeight;
  if (!inView && memory.loadedAtScroll === reading.scrollTop) return { load: false, memory: next };
  return { load: true, memory: { settledTop: reading.sentinelTop, loadedAtScroll: reading.scrollTop } };
}

/** The reader must have left everything alone this long for a part to load unasked. */
export const PART_BACKGROUND_QUIET_MS = 2_000;
/** How often the end is looked at while the next part has a kept layout. */
const KEPT_PART_LOOK_MS = 300;
/** Looks in a row (one each 1.5 s) the end of what is loaded must have held still. */
export const PART_BACKGROUND_STILL_LOOKS = 2;
/** The parts a tab takes unasked; past them the reader's coming near still loads more. */
export const PART_BACKGROUND_CAP = 40;
/** A tablet's page is killed for memory long before a desktop's: a 390-page book held 700 MB. */
export const PART_BACKGROUND_CAP_TOUCH = 16;





















export const PART_EAGER_MAX = 12;

/** Whether stored blocks hold a footnote reference (`footnoteRef`): a document
 * that does is laid out whole on every pass (footnotesAllowProgressiveLayout),
 * so its parts coming in at once made each pass the whole document's
 * (R286-D01: one task of 9.4 s where the parts one by one held 2.8 s at most). */
export function storedBlocksHoldFootnoteRef(blocks: readonly unknown[]): boolean {
  const stack: unknown[] = [...blocks];
  while (stack.length > 0) {
    const item = stack.pop();
    if (Array.isArray(item)) {
      for (const child of item) stack.push(child);
      continue;
    }
    if (!item || typeof item !== "object") continue;
    const node = item as { type?: unknown; content?: unknown };
    if (node.type === "footnoteRef") return true;
    if (Array.isArray(node.content)) stack.push(node.content);
  }
  return false;
}

/** Whether a document of `count` parts reads all of them at once on opening.
 * `window.__docsPartsAtOnce = false` keeps the one-after-another reading (A/B). */
/** How often a part held for the first screen looks again, and the longest
 * it waits for it (a font plan that never answers, a layout that never
 * settles): after that the parts go on as before. */
const FIRST_SCREEN_LOOK_MS = 200;
const FIRST_SCREEN_WAIT_MAX_MS = 20_000;

/** Whether parts read ahead go in back to back (see loadNext's end). */
export function partsBackToBack(): boolean {
  return typeof window === "undefined"
    || (window as Window & { __docsPartsBackToBack?: boolean }).__docsPartsBackToBack !== false;
}

export function loadsPartsAtOnce(count: number, touch: boolean): boolean {
  if (typeof window !== "undefined" && (window as Window & { __docsPartsAtOnce?: boolean }).__docsPartsAtOnce === false) return false;
  return !touch && count > 1 && count <= PART_EAGER_MAX;
}











export function shouldLoadPartInBackground(reading: {
  readonly stillLooks: number;
  readonly quietMs: number;
  readonly visible: boolean;
  readonly nextIndex: number;
  readonly touch: boolean;
  /** The next part opens on a layout kept for it (layoutSnapshot): it stands
   * hollow at once, nothing of it laid out, so one look at a still end will do. */
  readonly kept?: boolean;
}): boolean {
  if (!reading.visible || reading.stillLooks < (reading.kept ? 1 : PART_BACKGROUND_STILL_LOOKS)) return false;
  if (reading.quietMs < PART_BACKGROUND_QUIET_MS) return false;
  return reading.nextIndex < (reading.touch ? PART_BACKGROUND_CAP_TOUCH : PART_BACKGROUND_CAP);
}

/** The scroll position of `element`'s scroll container (the window's when
 *  no ancestor scrolls). */
function scrollTopOf(element: HTMLElement): number {
  let node: HTMLElement | null = element.parentElement;
  while (node && node !== document.body) {
    const overflow = getComputedStyle(node).overflowY;
    if ((overflow === "auto" || overflow === "scroll") && node.scrollHeight > node.clientHeight) {
      return node.scrollTop;
    }
    node = node.parentElement;
  }
  return window.scrollY;
}

export function useDocumentPageParts(options: UseDocumentPagePartsOptions): DocumentPageParts {
  const { pageId, body, hydration, hydrated } = options;
  const count = documentPartsCount(body);
  const [state, setState] = useState<DocumentPartsState>({
    count, loaded: 0, loading: false, saving: false, blocked: null, retrying: false,
  });
  const optionsRef = useRef(options);
  optionsRef.current = options;
  const owner = useRef(new Map<string, number>());
  const records = useRef(new Map<number, PartRecord>());
  const loadedRef = useRef(0);
  /** Every part loads at once (loadsPartsAtOnce), decided per hydration. */
  const eagerRef = useRef(false);
  /** The later parts read ahead when they load at once, by index. */
  const prefetched = useRef(new Map<number, Promise<DocumentPartResponse | null>>());
  /** Parts read ahead whose read came back ready (loadNext goes on to them). */
  const readyAhead = useRef(new Set<number>());
  /** The hydration (generation) whose parts were all asked for at once. */
  const eagerStarted = useRef(-1);
  /** The hydration (generation) whose parts' reads went out. */
  const eagerReadsStarted = useRef(-1);
  /** Runs `then` once the first screen is in, the reader goes further, or
   * FIRST_SCREEN_WAIT_MAX_MS have gone by (afterFirstScreen below). */
  const afterFirstScreenRef = useRef<(started: number, then: () => void) => void>((_started, then) => then());
  /** A look at the sentinel is held for the first screen (lookAtSentinel). */
  const lookHeld = useRef(false);
  const loadingRef = useRef(false);
  const blockedRef = useRef<DocumentPartsState["blocked"]>(null);
  const pending = useRef(new Map<number, string>());
  /** The pending parts' blocks one by one (what a kept edit is made from). */
  const pendingBlocks = useRef(new Map<number, readonly string[]>());
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  /** Tries in a row of the parts' saves that did not go through (the retry's back-off). */
  const saveTries = useRef(0);
  /** Parts whose conflict could not be merged at once, said once each. */
  const conflictReported = useRef(new Set<number>());
  const savingRef = useRef(false);
  /** A later part's blocks on their way to the server, one by one. */
  const sendingBlocks = useRef(new Map<number, readonly string[]>());
  /** The body's blocks as the editor serialized them when it was opened. */
  const openedBody = useRef<unknown[] | null>(null);
  // Counts baselines taken (the body room joins once there is one).
  const [baselineTaken, setBaselineTaken] = useState(0);
  // The body's block count as this view last prepared it for a save.
  const preparedBodyCount = useRef<number | null>(null);
  const nearEnd = useRef(false);
  const generation = useRef(0);
  /** The body as this editor serialized it when it was opened (or as last
   * prepared for saving), and the body draft prepared then. */
  const bodyBaseline = useRef<{ blocks: string; attrs: string; prepared: Record<string, unknown> | null } | null>(null);
  const rooms = useRef(new Map<number, LivePartRoom>());
  /** `hydrationToken` as of the last reset: a view holding another is not this reading's. */
  const resetToken = useRef<unknown>(undefined);
  const viewIsThisReading = useCallback((): boolean => {
    const token = optionsRef.current.hydrationToken;
    if (!token || token() === resetToken.current) return true;
    traceDocsHandoverPart({ step: "view-not-this-reading" });
    return false;
  }, []);
  // Which part holds each block, for what keeps the document's layout part by
  // part (paged/layoutSnapshot).
  useEffect(() => (pageId ? holdDocumentPartOwners(pageId, () => owner.current) : undefined), [pageId]);
  /** The view's document as last published into the rooms or drawn from them. */
  const roomsInStepWith = useRef<unknown>(null);

  const publish = useCallback((patch: Partial<DocumentPartsState>) => {
    setState((current) => {
      const next = { ...current, ...patch };
      return Object.keys(next).every((key) => next[key as keyof DocumentPartsState] === current[key as keyof DocumentPartsState])
        ? current
        : next;
    });
  }, []);

  const block = useCallback((reason: NonNullable<DocumentPartsState["blocked"]>) => {
    if (blockedRef.current) return;
    traceDocsHandoverPart({ step: "parts-blocked", reason });
    blockedRef.current = reason;
    // What was waiting to be saved stays kept (the part's recovery): the
    // page's next opening merges it in (mergeDocumentPartRecovery). Not from a
    // view whose parts no longer read as they were stored: that is the damage.
    for (const [index, blocks] of pendingBlocks.current) {
      const record = records.current.get(index);
      if (record && pageId && reason !== "count-jump") {
        writeDocumentPartRecovery(pageId, index, documentPartRecovery(record.savedBlocks, record.baseCount, blocks, record.revision, false));
      }
    }
    pending.current.clear();
    pendingBlocks.current.clear();
    publish({ blocked: reason, saving: false });
  }, [pageId, publish]);

  const closeRooms = useCallback(() => {
    for (const live of rooms.current.values()) live.room.destroy();
    rooms.current.clear();
    roomsInStepWith.current = null;
  }, []);

  // A part and the body as last known to read right. Moves with every look, so
  // an edit a block at a time never adds up to a jump.
  const partCounts = useRef(new Map<number, number>());
  // How often each Word element (a top-level block's `source.sourceAnchor`)
  // stands in the body and the parts loaded, as the server stored them. No
  // gesture puts an element in a second time (a paste or a copy brings no
  // `source`; a conflict's copy is a new block): a view holding one more often
  // than this is a part's blocks entered twice (R482-D04 LIVE: a part's 17
  // tables stored twice), and it saves nothing (partsReadRight).
  const storedAnchors = useRef(new Map<string, number>());
  const lastBulkGestureAt = useRef<number | null>(null);
  // Whether anyone has done anything here. Until then what differs from the
  // stored part is the editor's own reading of it, never an edit: saved, it wrote
  
  // empty paragraph gone), and a part entered twice as 1271 blocks of 636.
  const handAt = useRef<number | null>(null);
  const noteHand = useCallback(() => { handAt.current = performance.now(); }, []);
  const typedLately = useRef<Array<{ at: number; text: string }>>([]);
  useEffect(() => {
    const typed = (text: string | null | undefined) => {
      if (!text) return;
      const since = performance.now() - REFUSED_TEXT_WINDOW_MS;
      typedLately.current = [...typedLately.current.filter((item) => item.at >= since), { at: performance.now(), text }].slice(-REFUSED_TEXT_MAX_CHARS);
    };
    const onInput = (event: Event) => {
      noteHand();
      const input = event as InputEvent;
      if (input.inputType === "insertText" || input.inputType === "insertFromComposition") typed(input.data);
    };
    const onComposed = (event: Event) => { noteHand(); typed((event as CompositionEvent).data); };
    const onPaste = (event: Event) => { noteHand(); typed((event as ClipboardEvent).clipboardData?.getData("text/plain")); };
    window.addEventListener("beforeinput", onInput, true);
    window.addEventListener("compositionend", onComposed, true);
    window.addEventListener("paste", onPaste, true);
    return () => {
      window.removeEventListener("beforeinput", onInput, true);
      window.removeEventListener("compositionend", onComposed, true);
      window.removeEventListener("paste", onPaste, true);
    };
  }, [noteHand]);
  useEffect(() => {
    const bulk = () => { lastBulkGestureAt.current = performance.now(); noteHand(); };
    const onKey = (event: KeyboardEvent) => {
      if (!["Shift", "Control", "Alt", "Meta"].includes(event.key)) noteHand();
      const selection = typeof document === "undefined" ? null : document.getSelection();
      const target = event.target;
      if (event.key === "Backspace" || event.key === "Delete" || event.ctrlKey || event.metaKey ||
          // Undo and redo are chords. A key outside the page is a menu's Enter.
          (selection && !selection.isCollapsed && event.key.length === 1) ||
          !(target instanceof Element) || !target.closest(".ProseMirror")) bulk();
    };
    // A click in the page is a caret; one outside it may be the menu's Delete or Insert.
    const onPointer = (event: Event) => {
      noteHand();
      const target = event.target;
      if (!(target instanceof Element) || !target.closest(".ProseMirror")) bulk();
    };
    window.addEventListener("keydown", onKey, true);
    window.addEventListener("pointerdown", onPointer, true);
    for (const name of ["paste", "cut", "drop"]) window.addEventListener(name, bulk, true);
    return () => {
      window.removeEventListener("keydown", onKey, true);
      window.removeEventListener("pointerdown", onPointer, true);
      for (const name of ["paste", "cut", "drop"]) window.removeEventListener(name, bulk, true);
    };
  }, [noteHand]);
  const bulkGestureLately = useCallback(() => lastBulkGestureAt.current !== null &&
    performance.now() - lastBulkGestureAt.current < PART_JUMP_GESTURE_MS, []);
  







  const refuseJump = useCallback((index: number, before: number, now: number, what: string) => {
    traceDocsHandoverPart({ step: "part-save-refused", what, index, before, now });
    reportClientError(
      Object.assign(new Error(`document_part_save_refused ${what} part ${index}: ${before} to ${now} blocks`), { name: "DocumentPartSaveRefused" }),
      { source: "window", windowType: "document" },
    );
    if (pageId) {
      const earlier = readRefusedPartText(pageId);
      const text = typedLately.current.map((item) => item.text).join("");
      if (text) writeRefusedPartText(pageId, `${earlier ?? ""}${text}`.slice(-REFUSED_TEXT_MAX_CHARS));
      typedLately.current = [];
    }
    block("count-jump");
  }, [block, pageId]);
  /**
   * Blocks about to go into the view (a whole view's draft merged into the
   * parts, NoteEditor adoptDraftIntoParts) holding a Word element more often
   * than the stored document does: the view stops as a damaged one does and
   * nothing of them goes in (true). Read here, before the editor has them:
   * there the block-id plugin takes `source` from every paragraph after the
   * first that claims one Word paragraph, so a paragraph put in twice no
   * longer says so, and the guards after it (partsReadRight, the save's) see
   * only its tables twice (R482-D04 LIVE: a part's tables and an empty
   * paragraph stored twice, the paragraph written as `<w:p><w:r/></w:p>`).
   */
  const refuseBlocksTwice = useCallback((blocks: readonly unknown[], what: string): boolean => {
    if (blockedRef.current) return true;
    const twice = anchorsBeyondStored(blocks, storedAnchors.current);
    if (twice === 0) return false;
    refuseJump(0, 0, twice, what);
    return true;
  }, [refuseJump]);
  /** Whether partsReadRight would stop these parts (it says nothing, keeps nothing). */
  const partsJumped = useCallback((parts: readonly (readonly unknown[])[]): boolean => {
    if (anchorsBeyondStored(parts.flat(), storedAnchors.current) > 0) return true;
    if (bulkGestureLately()) return false;
    if (repeatedBlockIds(parts.flat()) > PART_JUMP_BLOCKS) return true;
    return parts.some((blocks, index) => {
      const before = partCounts.current.get(index);
      return before !== undefined && partCountJumped(before, blocks.length);
    });
  }, [bulkGestureLately]);
  /** Whether a draft handed in holds other top-level blocks than the view does now. */
  const draftIsNotTheView = useCallback((content: Record<string, unknown>): boolean => {
    const now = optionsRef.current.serialize();
    const held = isRecord(now) && Array.isArray(now.content) ? now.content : null;
    const drafted = Array.isArray(content.content) ? content.content : null;
    if (!held || !drafted) return false;
    if (held.length !== drafted.length) return true;
    return held.some((block, index) => blockId(block) !== blockId(drafted[index]));
  }, []);
  const partsReadRight = useCallback((parts: readonly (readonly unknown[])[]): boolean => {
    if (blockedRef.current) return false;
    // Whatever the reader just did: no gesture holds a Word element twice.
    const twice = anchorsBeyondStored(parts.flat(), storedAnchors.current);
    if (twice > 0) {
      refuseJump(0, 0, twice, "repeated-anchors");
      return false;
    }
    if (!bulkGestureLately()) {
      const repeated = repeatedBlockIds(parts.flat());
      if (repeated > PART_JUMP_BLOCKS) {
        refuseJump(0, 0, repeated, "repeated-ids");
        return false;
      }
      for (let index = 0; index < parts.length; index += 1) {
        const before = partCounts.current.get(index);
        if (before !== undefined && partCountJumped(before, parts[index].length)) {
          refuseJump(index, before, parts[index].length, "count");
          return false;
        }
      }
    }
    parts.forEach((blocks, index) => { partCounts.current.set(index, blocks.length); });
    return true;
  }, [bulkGestureLately, refuseJump]);

  // A new hydration (first open, a reload) holds the body alone again.
  useEffect(() => {
    traceDocsHandoverPart({ step: "parts-reset", hydration, loaded: loadedRef.current, bodyBlocks: Array.isArray((optionsRef.current.body as { content?: unknown } | null)?.content) ? ((optionsRef.current.body as { content: unknown[] }).content.length) : null });
    generation.current += 1;
    resetToken.current = optionsRef.current.hydrationToken?.();
    closeRooms();
    bodyBaseline.current = null;
    openedBody.current = null;
    preparedBodyCount.current = null;
    sendingBlocks.current.clear();
    owner.current = new Map();
    records.current = new Map();
    partCounts.current = new Map();
    pending.current.clear();
    pendingBlocks.current.clear();
    loadedRef.current = 0;
    loadingRef.current = false;
    blockedRef.current = null;
    prefetched.current.clear();
    readyAhead.current.clear();
    const stored = optionsRef.current.body;
    storedAnchors.current = new Map();
    if (isRecord(stored) && Array.isArray(stored.content)) {
      countSourceAnchors(stored.content, storedAnchors.current);
      partCounts.current.set(0, stored.content.length);
      for (const block of stored.content) {
        const id = blockId(block);
        if (id) owner.current.set(id, 0);
      }
    }
    eagerRef.current = typeof window !== "undefined"
      && loadsPartsAtOnce(documentPartsCount(stored), window.matchMedia?.("(pointer: coarse)").matches ?? false);
    publish({ count: documentPartsCount(stored), loaded: 0, loading: false, saving: false, blocked: null, retrying: false, refusedText: pageId ? readRefusedPartText(pageId) : null, eager: eagerRef.current });
    // The body as saved is read once per hydration, as the editor holds it.
  }, [closeRooms, pageId, hydration, publish]);

  useEffect(() => closeRooms, [closeRooms]);

  // The body as this editor holds it right after it was opened: what an
  // unedited draft's body is compared with (the stored form is the server's,
  // not the editor's own normalization of it).
  // A view typed into before this is read (keys held through a reload, typed
  // the moment it took input) holds more than the body: taken as the body,
  // those keys read as no edit, the stored body was sent in their place and
  
  // stored). Such a view is no baseline: every draft of it is saved.
  // Asked again by each later look (prepareContent) while there is none: a
  // view that filled in after it opened (its body in, nothing typed into it
  // that a later draft could not still carry) takes its baseline then, and
  
  // body a moment after its hydration and was left with no baseline: its
  // words were shown and never saved).
  const takeBodyBaseline = useCallback((): boolean => {
    if (!optionsRef.current.hydrated || bodyBaseline.current || loadedRef.current !== 0) return Boolean(bodyBaseline.current);
    const serialized = optionsRef.current.serialize();
    if (!isRecord(serialized) || !Array.isArray(serialized.content)) return false;
    const stored = optionsRef.current.body;
    // As stored when it says the same (samePartContent): read by sameAuthoredBody
    // alone, the editor's own reading of every opened body (its schema's defaults
    // on each block) was "edited", every draft of it was saved, and opening a
    
    const asStored = !isRecord(stored) || !Array.isArray(stored.content) || sameAuthoredBody(serialized, stored) ||
      samePartContent(serialized.content as unknown[], stored.content);
    if (!asStored && documentBodyCutNeedsSaying(stored, serialized)) {
      // Most of the body missing from the view as it opened is no edit of it:
      
      // load: each held its loading placeholder, one empty paragraph, as the
      // page came in as parts; the body room's first reading merged that
      // against the body as stored -- every block deleted -- and the page saved
      // it, three long documents' bodies of 1,109 blocks written as one). No
      // baseline: nothing of it is saved and its room is not joined until a
      // reading of the page that holds the body.
      traceDocsHandoverPart({ step: "baseline-not-the-body", view: serialized.content.length,
        stored: isRecord(stored) && Array.isArray(stored.content) ? stored.content.length : null });
      return false;
    }
    if (!asStored) traceDocsHandoverPart({ step: "baseline-edited" });
    bodyBaseline.current = {
      blocks: asStored ? JSON.stringify(serialized.content) : "",
      attrs: withoutLayoutAttrs(serialized.attrs),
      prepared: null,
    };
    // The body room merges against the body as opened: the stored one, as
    // this editor writes it, when the view already holds more (or the
    // room's first reading took the keys back out as not the room's).
    openedBody.current = asStored || !isRecord(stored) || !Array.isArray(stored.content)
      ? serialized.content as unknown[]
      : optionsRef.current.editorFormOf?.(stored.content) ?? serialized.content as unknown[];
    setBaselineTaken((count) => count + 1);
    return true;
  }, []);
  useEffect(() => { takeBodyBaseline(); }, [hydrated, hydration, takeBodyBaseline]);

  // A part's edit not yet saved is kept (and, once saved, let go) as the
  // page body's unsaved draft is: see DocumentPartRecovery.
  const keepPartEdit = useCallback((index: number, sending: boolean) => {
    if (!pageId) return;
    const record = records.current.get(index);
    const current = pendingBlocks.current.get(index);
    if (!record || !current) return;
    writeDocumentPartRecovery(
      pageId, index, documentPartRecovery(record.savedBlocks, record.baseCount, current, record.revision, sending),
    );
  }, [pageId]);

  /** A part another save moved on: the server's blocks, this editor's and
   * the ones both last agreed on merged (`mergeCanonicalBlocks`, a block both
   * changed kept twice), shown in the editor and queued on the new revision. */
  const mergeConflictedPart = useCallback(async (index: number, record: PartRecord): Promise<"merged" | "later"> => {
    const current = optionsRef.current;
    if (!pageId || !current.editorFormOf || !current.mergePartIntoView) return "later";
    try {
      let fresh: DocumentPartResponse;
      try {
        fresh = await readDocumentPart(pageId, index, { cache: "no-store" });
      } catch {
        // The part could not be read now (the server busy, the network gone): later.
        return "later";
      }
      const stored = Array.isArray(fresh.content?.content) ? (fresh.content?.content as unknown[]) : null;
      if (!stored || fresh.partIndex !== index || typeof fresh.revision !== "number") return "later";
      const remote = current.editorFormOf(stored);
      const serialized = current.serialize();
      const split = serialized ? splitDocumentParts(serialized, owner.current, loadedRef.current) : null;
      if (!remote || !split?.ok) return "later";
      const local = split.parts[index] as MergeBlock[];
      const base = (record.savedBlocks
        ? record.savedBlocks.map((json) => JSON.parse(json))
        : record.saved ? JSON.parse(record.saved) : record.recoveredBase ?? []) as MergeBlock[];
      const merged = mergeCanonicalBlocks(base, local, remote as MergeBlock[], { typedIntoBoth: true, mergeText: true });
      if (docsTraceOn()) traceDocsHandoverPart({ step: "part-conflict-shape", index, base: base.length, local: local.length,
        remote: remote.length, merged: merged.blocks.length, localInRemote: sharedIds(local, remote), localInBase: sharedIds(local, base),
        conflicts: merged.conflicts, anchorsTwice: repeatedSourceAnchors(merged.blocks) });
      // A part holding a Word element twice is not saved (see roomRead).
      if (anchorsBeyondStored(merged.blocks, storedAnchors.current) > 0) {
        refuseJump(index, remote.length, merged.blocks.length, "conflict-merge-repeated-anchors");
        return "merged";
      }
      // A view the merge cannot be written into exactly (a block cut across pages
      // being laid out): this view's edit stays kept, the part is merged again later.
      if (merged.changedLocal && !current.mergePartIntoView(local, merged.blocks)) return "later";
      partCounts.current.delete(index);
      for (const item of merged.blocks) {
        const id = blockId(item);
        if (id) owner.current.set(id, index);
      }
      const remoteEach = remote.map((item) => JSON.stringify(item));
      record.revision = fresh.revision;
      record.saved = `[${remoteEach.join(",")}]`;
      record.savedBlocks = remoteEach;
      record.baseCount = remote.length;
      const each = merged.blocks.map((item) => JSON.stringify(item));
      const json = `[${each.join(",")}]`;
      if (json !== record.saved) {
        pending.current.set(index, json);
        pendingBlocks.current.set(index, each);
        keepPartEdit(index, false);
      } else {
        writeDocumentPartRecovery(pageId, index, null);
      }
      return "merged";
    } catch {
      return "later";
    }
  }, [keepPartEdit, pageId, refuseJump]);

  /** Whether a part's edits go to its room: it takes input, or may still (it
   * was asked for a moment ago; its first reading merges what was typed). */
  const roomHolds = useCallback((index: number): boolean => {
    const live = rooms.current.get(index);
    if (!live || !live.inView || !optionsRef.current.partRooms) return false;
    if (live.room.accepting) return true;
    return live.room.status === "connecting" && Date.now() - live.askedAt < ROOM_FIRST_READ_WAIT_MS;
  }, []);

  const flushSaves = useCallback(async () => {
    if (savingRef.current || !pageId) return;
    savingRef.current = true;
    const started = generation.current;
    try {
      while (pending.current.size > 0 && !blockedRef.current && generation.current === started) {
        // A part's room that may still come first takes its edit at its reading
        // (roomRead); one that does not come leaves it to this save.
        const index = [...pending.current.keys()].find((key) => !roomHolds(key));
        if (index === undefined) break;
        const json = pending.current.get(index) as string;
        const sentBlocks = pendingBlocks.current.get(index) ?? null;
        pending.current.delete(index);
        pendingBlocks.current.delete(index);
        const record = records.current.get(index);
        if (!record || record.saved === json) continue;
        const blocks = JSON.parse(json) as unknown[];
        const known = partCounts.current.get(index);
        if (known !== undefined && partCountJumped(known, blocks.length) && !bulkGestureLately()) {
          refuseJump(index, known, blocks.length, "save");
          return;
        }
        if (anchorsBeyondStored(blocks, storedAnchors.current) > 0) {
          refuseJump(index, record.baseCount, blocks.length, "save-repeated-anchors");
          return;
        }
        if (sentBlocks) {
          writeDocumentPartRecovery(pageId, index, documentPartRecovery(
            record.savedBlocks, record.baseCount, sentBlocks, record.revision, true,
          ));
        }
        if (sentBlocks) sendingBlocks.current.set(index, sentBlocks);
        try {
          // Said only by a reader who did something of the kind in this reading:
          // the server refuses such a cut otherwise (document_part_large_removal).
          const declared = lastBulkGestureAt.current !== null && partCountJumped(record.baseCount, blocks.length);
          const receipt = await apiPut<DocumentPartReceipt>(
            `/api/v1/pages/${encodeURIComponent(pageId)}/document-parts/${index}`,
            {
              content_json: { type: "doc", content: optionsRef.current.storedForm?.(blocks) ?? blocks },
              expected_revision: record.revision,
              ...(declared ? { large_removal: true } : {}),
            },
          );
          sendingBlocks.current.delete(index);
          if (generation.current !== started) return;
          record.revision = receipt.revision;
          record.saved = json;
          record.savedBlocks = sentBlocks;
          record.baseCount = sentBlocks ? sentBlocks.length : record.baseCount;
          // A newer edit of the part is kept against what is stored now.
          if (pending.current.has(index)) keepPartEdit(index, false);
          else writeDocumentPartRecovery(pageId, index, null);
          saveTries.current = 0;
          publish({ retrying: false });
        } catch (error) {
          sendingBlocks.current.delete(index);
          if (generation.current !== started) return;
          if (error instanceof ApiRequestError && error.code === "document_part_large_removal") {
            refuseJump(index, record.baseCount, blocks.length, "server");
            return;
          }
          if (error instanceof ApiRequestError && error.code === "document_part_blocks_twice") {
            refuseJump(index, record.baseCount, blocks.length, "server-repeated-anchors");
            return;
          }
          // The part brings in a block another stored part holds: a block
          // moved across parts waits for the part it left to be saved without
          // it (tried again below); one still held there after every try is
          // a block put in twice (R482-D04 LIVE), and the view stops.
          const heldElsewhere = error instanceof ApiRequestError && error.code === "document_part_blocks_in_other_part";
          if (heldElsewhere && saveTries.current >= PART_BLOCKS_ELSEWHERE_TRIES) {
            refuseJump(index, record.baseCount, blocks.length, "server-blocks-in-other-part");
            return;
          }
          if (!heldElsewhere && error instanceof ApiRequestError && error.status === 409) {
            // The part was saved elsewhere (another tab, another member):
            // both sides are kept, merged block by block, and saved again
            
            if (await mergeConflictedPart(index, record) === "merged") {
              if (generation.current !== started) return;
              saveTries.current = 0;
              continue;
            }
            if (generation.current !== started) return;
            // Not merged now (the part unreadable a moment, the view not taking
            // the merge yet): the edit stays kept and is merged and saved again
            // shortly. Stopping here stopped every later save of the page for
            // the reading: words typed after a busy server's answer were in the
            
            // read, and nothing was saved again).
            if (!conflictReported.current.has(index)) {
              conflictReported.current.add(index);
              reportClientError(Object.assign(new Error(`part ${index} not merged yet`), { name: "DocumentPartMergeLater" }));
            }
          }
          // Kept for the next attempt unless a newer edit replaced it.
          if (!pending.current.has(index)) {
            pending.current.set(index, json);
            if (sentBlocks) pendingBlocks.current.set(index, sentBlocks);
          }
          keepPartEdit(index, false);
          publish({ retrying: true });
          saveTries.current += 1;
          saveTimer.current = setTimeout(() => {
            saveTimer.current = null;
            void flushSaves();
          }, partSaveRetryDelay(saveTries.current));
          return;
        }
      }
    } finally {
      savingRef.current = false;
      if (generation.current === started) {
        publish({ saving: pending.current.size > 0 && !blockedRef.current });
        if (pending.current.size > 0 && !blockedRef.current && !saveTimer.current) {
          saveTimer.current = setTimeout(() => {
            saveTimer.current = null;
            void flushSaves();
          }, PART_SAVE_DEBOUNCE_MS);
        }
      }
    }
  }, [bulkGestureLately, keepPartEdit, mergeConflictedPart, pageId, publish, refuseJump, roomHolds]);

  const scheduleSaves = useCallback(() => {
    if (saveTimer.current) clearTimeout(saveTimer.current);
    publish({ saving: true });
    saveTimer.current = setTimeout(() => {
      saveTimer.current = null;
      void flushSaves();
    }, PART_SAVE_DEBOUNCE_MS);
  }, [flushSaves, publish]);

  useEffect(() => () => {
    if (saveTimer.current) clearTimeout(saveTimer.current);
  }, []);

  // Leaving with a part not yet saved asks first, as any unsaved page does.
  useEffect(() => {
    if (!state.saving) return;
    const warn = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [state.saving]);

  /** The rooms' part of what is still unsaved, said as the page's status says. */
  const noteRoomsState = useCallback(() => {
    let unsaved = false;
    for (const live of rooms.current.values()) if (live.room.unsaved) unsaved = true;
    publish({ saving: unsaved || pending.current.size > 0 });
    if (!unsaved && pending.current.size === 0 && roomHolds(0)) optionsRef.current.partRooms?.settled?.();
  }, [publish, roomHolds]);

  /**
   * Whatever the view holds that its rooms have not heard of goes to them now.
   * It is read from the view itself and never taken from a serialized draft: a
   * draft made before a member's words were drawn would take them out again.
   * False when the view cannot be read now.
   */
  const publishViewToRooms = useCallback((): boolean => {
    const current = optionsRef.current;
    const view = current.partRooms;
    if (!view || rooms.current.size === 0 || blockedRef.current || !current.writable) return true;
    // No room takes input (none read yet, or none to be had where part rooms
    // are off): nothing to read the view for.
    if (![...rooms.current.values()].some((live) => live.room.accepting && live.inView)) {
      noteRoomsState();
      return true;
    }
    if (view.composing() || !viewIsThisReading()) return false;
    const doc = view.viewDoc();
    if (doc && doc === roomsInStepWith.current) {
      noteRoomsState();
      return true;
    }
    /** One part as the view now holds it, to its room (and, the body, its attrs). */
    const publishPart = (index: number, blocks: readonly unknown[], attrs: Record<string, unknown> | undefined): boolean => {
      const live = rooms.current.get(index);
      if (!live?.room.accepting || !live.inView) return true;
      if (index > 0 && handAt.current === null && records.current.get(index)?.saved) return true;
      const each = blocks.map(jsonOfBlock);
      const json = `[${each.join(",")}]`;
      if (live.viewBlocks && live.viewPlaces && json !== live.viewJson) {
        // Which blocks this member changed: a change of theirs outweighs another's deletion (onRemote).
        const now = Date.now();
        for (let at = 0; at < blocks.length; at += 1) {
          const id = blockId(blocks[at]);
          if (!id) continue;
          const place = live.viewPlaces.get(id);
          // As authored: the part last drawn from the room is in the room's
          // form, and a block only written otherwise is no change of this member's.
          if (place === undefined || (jsonOfBlock(live.viewBlocks[place]) !== each[at] &&
            !sameAuthoredBlock(live.viewBlocks[place], blocks[at]))) live.mine.set(id, now);
        }
      }
      const attrsJson = attrs ? JSON.stringify(attrs) : null;
      const attrsMoved = attrsJson !== live.viewAttrs;
      const keep = () => {
        live.viewBlocks = blocks;
        live.viewPlaces = blockPlaces(blocks);
      };
      if (json === live.viewJson && !attrsMoved) { keep(); return true; }
      const newAttrs = attrs && attrsMoved && !attrsSayNothingNew(attrs, live.room.attrs()) ? attrs : undefined;
      if (newAttrs && docsTraceOn()) {
        const notes = (value: unknown) => (Array.isArray(value) ? value.length : null);
        traceDocsHandoverPart({ step: "pr-publish-attrs", index, notes: notes(newAttrs.footnotes), roomNotes: notes(live.room.attrs()?.footnotes) });
      }
      if (json === live.viewJson && !newAttrs) {
        live.viewAttrs = attrsJson;
        keep();
        return true;
      }
      try {
        live.room.publish(blocks, newAttrs);
      } catch (error) {
        console.warn("[part-room] a part the schema does not take was not published", error);
        block("unreadable");
        return false;
      }
      const published = tracedWords(blocks);
      if (published && tracedPublishes.get(index) !== published.join()) {
        tracedPublishes.set(index, published.join());
        traceDocsHandoverPart({ step: "pr-publish", index, words: published, room: tracedWords(live.room.blocks()) });
      }
      live.viewJson = json;
      live.viewAttrs = attrsJson;
      keep();
      if (index > 0 && pageId && live.room.unsaved) {
        const record = records.current.get(index);
        if (record) {
          writeDocumentPartRecovery(pageId, index, documentPartRecovery(record.savedBlocks, record.baseCount, each, record.revision, false));
        }
      }
      return true;
    };
    // A key (or a few) since the view and its rooms were last in step: only the
    // sheets it changed are serialized, and their blocks spliced into each
    // part as last published. Taken only when it is plainly the same part
    // with some blocks' contents changed -- every block there before is there
    // after, in the same order, each one known as that part's, and the blocks
    // around them read as published -- else the whole view is read as below.
    // Serializing the whole view for each publish (a long document's every
    // block, its page layout measured) kept two members typing in a book
    
    const scoped = (() => {
      const before = roomsInStepWith.current;
      if (!before || !doc || !view.changedBlocks) return null;
      // The document's attributes (the body's room carries them) are read whole.
      if ((before as { attrs?: unknown }).attrs !== (doc as { attrs?: unknown }).attrs) return null;
      const changed = view.changedBlocks(before, doc);
      if (!changed) return null;
      const was = changed.before.map(blockId), now = changed.after.map(blockId);
      if (was.length !== now.length || was.some((id, index) => !id || id !== now[index])) return null;
      const byPart = new Map<number, Array<{ id: string; block: unknown; old: unknown }>>();
      for (let index = 0; index < now.length; index += 1) {
        const id = now[index] as string;
        const part = owner.current.get(id);
        if (part === undefined || part > loadedRef.current) return null;
        const list = byPart.get(part) ?? [];
        list.push({ id, block: changed.after[index], old: changed.before[index] });
        byPart.set(part, list);
      }
      const parts = new Map<number, unknown[]>();
      for (const [part, list] of byPart) {
        const live = rooms.current.get(part);
        if (!live?.room.accepting || !live.inView) continue;
        if (!live.viewBlocks || !live.viewPlaces) return null;
        const next = [...live.viewBlocks];
        for (const { id, block: changedBlock, old } of list) {
          const place = live.viewPlaces.get(id);
          if (place === undefined) return null;
          // The scoped reading of the block as it was must be the block as published.
          if (!sameAuthoredBlock(old, live.viewBlocks[place])) return null;
          next[place] = changedBlock;
        }
        parts.set(part, next);
      }
      return parts;
    })();
    countPublish(scoped ? "scoped" : "whole");
    if (scoped) {
      for (const [index, blocks] of scoped) {
        if (!publishPart(index, blocks, undefined)) { noteRoomsState(); return true; }
      }
      if (!blockedRef.current) roomsInStepWith.current = doc;
      noteRoomsState();
      return true;
    }
    const serialized = (current.serializeForRooms ?? current.serialize)();
    if (!serialized) return false;
    const split = splitDocumentParts(serialized, owner.current, loadedRef.current);
    if (split.ok && !partsReadRight(split.parts)) return false;
    if (!split.ok) {
      block(split.reason);
      return true;
    }
    for (let index = 0; index < split.parts.length; index += 1) {
      const blocks = split.parts[index];
      for (const item of blocks) {
        const id = blockId(item);
        if (id) owner.current.set(id, index);
      }
      const attrs = index === 0 && isRecord(serialized.attrs) ? withoutPrivateLayoutAttrs(serialized.attrs) : undefined;
      if (!publishPart(index, blocks, attrs)) { noteRoomsState(); return true; }
    }
    if (!blockedRef.current) roomsInStepWith.current = doc;
    noteRoomsState();
    return true;
  }, [block, noteRoomsState, pageId, partsReadRight, viewIsThisReading]);

  /** A room read (first, or afresh after the server seeded it again): what it
   * says, what the view holds and the two's last agreement are merged, the
   * view shows the result and the room is told what it did not know. */
  const roomRead = useCallback((live: LivePartRoom) => {
    const current = optionsRef.current;
    const view = current.partRooms;
    if (!view || blockedRef.current || !current.editorFormOf || !current.mergePartIntoView) return;
    // The part as the view shows it: read by itself (a later part), else from
    // the whole view. Each later part's reading serialized the whole view
    // until then, a document of 43 parts 43 times, each longer than the last
    
    const own = live.index > 0 ? view.partBlocks?.(live.index, owner.current, loadedRef.current) ?? null : null;
    const serialized = own ? null : (current.serializeForRooms ?? current.serialize)();
    const split = serialized ? splitDocumentParts(serialized, owner.current, loadedRef.current) : null;
    const remote = current.editorFormOf(live.room.blocks());
    const shownBlocks = own ?? (split?.ok ? split.parts[live.index] : undefined);
    if (!shownBlocks || !remote) {
      block("unreadable");
      return;
    }
    // The view's own blocks (what it shows), and as read against the room.
    const shown = shownBlocks as MergeBlock[];
    // Nobody changed the part since it and the room last agreed: the room
    // holds the stored part (no member's change, saved or not, the view has
    // not got) and the view shows it as it came in. Nothing to merge, draw,
    // publish or save -- a document only opened read each part, often wrote
    // a block back that differed only in how it was written, published that
    // and so saved the part (the same comparison a save makes:
    // sameAuthoredBlock).
    const sameAs = (blocks: readonly unknown[]) => blocks.length === live.base.length &&
      blocks.every((item, at) => sameAuthoredBlock(item, live.base[at]));
    if (live.index > 0 && !pending.current.has(live.index) && !pendingBlocks.current.has(live.index) &&
      (sameAs(live.room.blocks()) || sameAs(remote)) && sameAs(shown)) {
      partCounts.current.delete(live.index);
      for (const item of shown) {
        const id = blockId(item);
        if (id) owner.current.set(id, live.index);
      }
      const json = JSON.stringify(shown);
      live.viewJson = json;
      live.viewBlocks = shown;
      live.viewPlaces = blockPlaces(shown);
      noteRoomsState();
      return;
    }
    const local = withStoredIdsForEmptyBlocks(live.base as MergeBlock[], shown);
    // The room and the view must be two readings of one part: a room holding
    // none of the blocks the view shows of it (the part as loaded does), with
    // blocks the view changed, is not merged -- merged, each changed block went
    // in at the part's start beside the room's own (a part's blocks twice).
    // One whose blocks the view left as loaded is the room's (another member
    // wrote the part anew): the merge takes it.
    if (local.length > 0 && remote.length > 0 && sharedIds(local, remote) === 0 &&
        sharedIds(local, live.base) > 0 && !sameAuthoredBody({ content: local }, { content: live.base })) {
      refuseJump(live.index, remote.length, local.length, "room-read-disjoint");
      return;
    }
    const merged = mergeCanonicalBlocks(live.base as MergeBlock[], local, remote as MergeBlock[], { typedIntoBoth: true, mergeText: true });
    if (docsTraceOn()) traceDocsHandoverPart({ step: "pr-read-shape", index: live.index, base: live.base.length, local: local.length,
      remote: remote.length, merged: merged.blocks.length, localInRemote: sharedIds(local, remote), localInBase: sharedIds(local, live.base),
      remoteInBase: sharedIds(remote, live.base), conflicts: merged.conflicts, anchorsTwice: repeatedSourceAnchors(merged.blocks) });
    // A merge that holds a Word element more often than the document does is
    // no merge of two readings of one part (R482-D04 LIVE: a part's tables
    // put in at its start a second time): nothing of it is shown, published
    // or saved, and the view stops as a damaged one does.
    const twice = anchorsBeyondStored(merged.blocks, storedAnchors.current);
    if (twice > 0) {
      refuseJump(live.index, remote.length, merged.blocks.length, "room-read-repeated-anchors");
      return;
    }
    const mergedWords = tracedWords(merged.blocks);
    if (mergedWords) traceDocsHandoverPart({ step: "pr-read", index: live.index, base: tracedWords(live.base),
      local: tracedWords(local), remote: tracedWords(remote), merged: mergedWords, conflicts: merged.conflicts });
    // A block the room only writes otherwise (`sameAuthoredBlock`: a key's
    // order, the file's `source`, an attribute at its default) says nothing new
    // to the view, which keeps its own. Written in, a paragraph cut across two
    // pages stood whole on the first and was cut again by the next pass, however
    // it fell then: each part of a book that opened rewrote a dozen blocks, and
    
    // of one book differed in, and every page a kept layout had right).
    const localBlocks = new Map(local.map((item) => [blockId(item), item]));
    const viewBlocks = merged.blocks.map((item) => {
      const own = localBlocks.get(blockId(item));
      return own && sameAuthoredBlock(own, item) ? own : item;
    });
    const viewJson = JSON.stringify(viewBlocks);
    const viewChanged = viewJson !== JSON.stringify(shown);
    if (viewChanged && !current.mergePartIntoView(shown, viewBlocks)) {
      block("conflict");
      return;
    }
    partCounts.current.delete(live.index);
    for (const item of merged.blocks) {
      const id = blockId(item);
      if (id) owner.current.set(id, live.index);
    }
    const remoteJson = JSON.stringify(remote);
    const mergedJson = JSON.stringify(merged.blocks);
    if (mergedJson !== remoteJson && current.writable) {
      try {
        live.room.publish(merged.blocks);
      } catch (error) {
        console.warn("[part-room] a part the schema does not take was not published", error);
        block("unreadable");
        return;
      }
    }
    live.viewJson = viewJson;
    live.viewBlocks = viewBlocks;
    live.viewPlaces = blockPlaces(viewBlocks);
    if (live.index === 0 && isRecord(serialized?.attrs)) {
      // The room's attributes stand: a body's are changed one at a time, so
      // one changed before this reading came is published with the next edit.
      if (docsTraceOn()) {
        const notes = (value: unknown) => (Array.isArray(value) ? value.length : null);
        traceDocsHandoverPart({ step: "pr-read-attrs", roomNotes: notes(live.room.attrs().footnotes), viewNotes: notes(serialized?.attrs && (serialized.attrs as Record<string, unknown>).footnotes) });
      }
      view.applyAttrs?.(live.room.attrs());
      live.viewAttrs = null;
    }
    // What the room holds beyond the stored part, left by a member who went
    // before their save: it is saved now, as anything of this member's is.
    if (remoteJson !== JSON.stringify(live.base) && current.writable) live.room.requestSave();
    // What waited for this reading is in the room now (its merge has `local`).
    pending.current.delete(live.index);
    pendingBlocks.current.delete(live.index);
    // This room's part is in step now, not every room's: what the view holds
    // that another room has not heard of (typed a moment before, its
    // serialization not run yet) stays to be published. Taken as in step, the
    
    // s7c9-1: three keys typed back into the body of a reopened view, a later
    // part's room read 760 ms after, the keys never published: 27 typed, 24
    // stored; s7c2-1 likewise).
    if (viewChanged) roomsInStepWith.current = null;
    noteRoomsState();
  }, [block, noteRoomsState, refuseJump]);

  const startRoom = useCallback((index: number, partKey: string, base: readonly unknown[], inView = true): LivePartRoom | null => {
    const current = optionsRef.current;
    const view = current.partRooms;
    const schema = view?.schema();
    if (!pageId || !view || !schema || rooms.current.has(index)) return null;
    const started = generation.current;
    // The body's room delivers what others save: the window keeps its content.
    const releaseHold = index === 0 ? holdPageForPartRoom(pageId) : null;
    const live: LivePartRoom = {
      index,
      askedAt: Date.now(),
      viewJson: null,
      viewAttrs: null,
      base,
      inView,
      readWaiting: false,
      viewBlocks: null,
      viewPlaces: null,
      mine: new Map(),
      room: new WordPartRoom({
        pageId,
        partKey,
        schema,
        // Another's change waits while the part is not in the view yet.
        beforeRemote: () => live.inView && publishViewToRooms(),
        onRead: () => {
          if (generation.current !== started) return;
          if (!live.inView) {
            live.readWaiting = true;
            return;
          }
          roomRead(live);
        },
        caret: () => optionsRef.current.partRooms?.caret?.() ?? null,
        onRemote: (before, after, lost, caret) => {
          if (docsTraceOn() && (before.length !== after.length || repeatedSourceAnchors(after) > 0)) {
            traceDocsHandoverPart({ step: "pr-remote-shape", index, before: before.length, after: after.length,
              shared: sharedIds(after, before), anchorsTwice: repeatedSourceAnchors(after) });
          }
          const words = tracedWords(after), wordsBefore = tracedWords(before);
          if (words && (words.join() !== wordsBefore?.join() || generation.current !== started || blockedRef.current)) {
            traceDocsHandoverPart({ step: "pr-remote", index, before: wordsBefore, after: words,
              stale: generation.current !== started, blocked: blockedRef.current });
          }
          if (generation.current !== started || blockedRef.current) return;
          // A paragraph this member typed into a moment ago that another took
          // out whole stays, with what was typed (a change outweighs a deletion,
          // as mergeCanonicalBlocks has it): the room deletes the whole element
          // and every letter typed into it with it.
          // And this member's letters a change took out with their text run, unseen, go back.
          const resurrected = keptFromDeletion(live, before, after, index, owner.current);
          const kept = withLettersKept(resurrected ?? after, lost, after) ?? resurrected;
          const target = kept ?? after;
          // The caret the room could not follow (its text run went) in a block
          // whose letters of this member's were written back: after them, where
          // this member's next key goes on from (it went to the paragraph's
          // start, or the next paragraph, and the word was cut in two).
          let caretTo = caret;
          if (!caretTo && lost && lost.length > 0) {
            const here = optionsRef.current.partRooms?.caret?.() ?? null;
            const block = here && lost.some((entry) => entry.block === here.block && entry.offset === null)
              ? target.find((item) => blockId(item) === here.block) : undefined;
            if (here && block) caretTo = { block: here.block, offset: inlineSizeOf(block) };
          }
          if (!optionsRef.current.mergePartIntoView?.(before, target, caretTo)) {
            const beforeIds = new Set(before.map(blockId)), targetIds = new Set(target.map(blockId));
            traceDocsHandoverPart({ step: "pr-merge-refused", index, kept: Boolean(kept), lost: lost ? lost.map((entry) => entry.block) : [],
              caret: caretTo ?? null, gone: [...beforeIds].filter((id) => !targetIds.has(id)).slice(0, 5),
              added: [...targetIds].filter((id) => !beforeIds.has(id)).slice(0, 5) });
            block("conflict");
            return;
          }
          if (kept) {
            try {
              live.room.publish(kept);
              traceDocsHandoverPart({ step: "pr-kept", index, blocks: kept.length - after.length });
            } catch (error) {
              console.warn("[part-room] a part the schema does not take was not published", error);
            }
          }
          // A member's words: this part's count is the room's now, no jump of ours.
          partCounts.current.delete(index);
          for (const item of target) {
            const id = blockId(item);
            if (id) owner.current.set(id, index);
          }
          live.viewJson = null;
          // The view's part is the room's now (as the room reads it: a later
          // key's blocks are spliced into it, compared as authored).
          live.viewBlocks = target;
          live.viewPlaces = blockPlaces(target);
          // The view is as the rooms say now: a later look finds nothing to
          // publish until it changes again.
          roomsInStepWith.current = optionsRef.current.partRooms?.viewDoc() ?? null;
        },
        onRemoteAttrs: (attrs) => {
          if (generation.current !== started || blockedRef.current) return;
          optionsRef.current.partRooms?.applyAttrs?.(attrs);
          live.viewAttrs = null;
          roomsInStepWith.current = optionsRef.current.partRooms?.viewDoc() ?? null;
        },
        onAcknowledged: () => {
          if (generation.current !== started) return;
          // The server holds them now, whatever becomes of this tab.
          if (index > 0) writeDocumentPartRecovery(pageId, index, null);
          noteRoomsState();
        },
        onSaved: (receipt) => {
          if (generation.current !== started) return;
          const answer = isRecord(receipt) ? receipt : {};
          const blocks = live.viewJson ? (JSON.parse(live.viewJson) as unknown[]) : null;
          if (blocks) live.base = blocks;
          if (index === 0 && typeof answer.contentVersion === "number") {
            optionsRef.current.partRooms?.bodySaved?.(answer.contentVersion);
          }
          const record = records.current.get(index);
          if (record) {
            if (typeof answer.revision === "number") record.revision = answer.revision;
            if (live.viewJson && blocks) {
              const each = blocks.map((item) => JSON.stringify(item));
              record.saved = live.viewJson;
              record.savedBlocks = each;
              record.baseCount = each.length;
            }
          }
          noteRoomsState();
        },
        onStatus: (status) => {
          if (status === "denied" || status === "closed") releaseHold?.();
          if (generation.current !== started) return;
          noteRoomsState();
          if (index === 0 && status === "synced") optionsRef.current.partRooms?.joined?.();
        },
        ...(index === 0
          ? { onPresence: (bytes: Uint8Array) => optionsRef.current.partRooms?.presence?.(bytes) }
          : {}),
        onRefused: (code) => {
          // Kept in the room, unsaved, and said by the page's status; the
          // room's next change or the next opening saves it again.
          console.warn("[part-room] the server refused this part's change", { index, code });
        },
      }),
    };
    rooms.current.set(index, live);
    live.room.start();
    return live;
  }, [block, noteRoomsState, pageId, publishViewToRooms, roomRead]);

  // The body's room, as soon as the editor holds the body as opened.
  useEffect(() => {
    if (!hydrated || !optionsRef.current.partRooms || !bodyBaseline.current || !openedBody.current) return;
    if (documentPartsCount(optionsRef.current.body) < 2 || blockedRef.current) return;
    startRoom(0, BODY_ROOM_KEY, openedBody.current);
  }, [hydrated, hydration, startRoom, baselineTaken]);

  /** A page draft as the page saves it: the body alone, or null when its room
   * holds the body too and the page's save has nothing of it to send. */
  const prepareContent = useCallback((content: Record<string, unknown>): Record<string, unknown> | null | undefined => {
    const stored = optionsRef.current.body;
    if (documentPartsCount(stored) < 2 || !isRecord(stored)) return content;
    const storedBlocks = Array.isArray(stored.content) ? stored.content : [];
    // The body as this page read it, sent again where nothing new of it is
    // to be saved. It is the body the page was opened with, not the body as
    // stored now: once the body's room has saved an edit (the page's version
    // moved on with it), sent under that version it wrote the opened body
    // back over the edit, the room was seeded again from it and the edit
    
    // paragraph and a heading stored, then the body back at 222 blocks and
    // all three gone). Nothing of the body is sent instead (null), which is
    // what each of these cases means. `window.__docsPartsStaleBodyResent =
    // true`: the body as opened is sent again (old).
    const unchangedBody = (): Record<string, unknown> | null =>
      typeof window !== "undefined" && (window as Window & { __docsPartsStaleBodyResent?: boolean }).__docsPartsStaleBodyResent === true
        ? { ...stored, content: storedBlocks }
        : null;
    if (blockedRef.current || !optionsRef.current.writable) {
      // Nothing is saved from here: the body stays what it is.
      return unchangedBody();
    }
    // Until the editor holds the body as opened (its reading of it is taken
    // right after hydration), what it holds is not the body: a draft made
    // then (an empty placeholder while the page was still loading) would
    // replace the stored body with nothing.
    // Not yet (undefined): the view is not the body as this reading opened
    // it, or not read as opened yet. Sent as the stored body, the save said
    
    // fin4-2, f7a-2: keys typed the moment a reload's view took input, shown
    // and never stored). The caller asks again a moment later.
    if (!optionsRef.current.hydrated || !takeBodyBaseline() || !viewIsThisReading()) {
      traceDocsHandoverPart({ step: "prepare-not-yet" });
      return undefined;
    }
    const split = splitDocumentParts(content, owner.current, loadedRef.current);
    if (split.ok && !blockedRef.current && partsJumped(split.parts) && draftIsNotTheView(content)) {
      // A draft serialized before this reading (the page read again under it,
      // its save made a moment later) is not the view: split against this
      
      // a body of 389 blocks drafted as 1827, the view itself 389). Not yet:
      // the page drafts the view again.
      traceDocsHandoverPart({ step: "prepare-stale-draft" });
      return undefined;
    }
    if (split.ok && !partsReadRight(split.parts)) return unchangedBody();
    publishViewToRooms();
    if (blockedRef.current) return unchangedBody();
    if (!split.ok) {
      block(split.reason);
      return unchangedBody();
    }
    // A loaded part none of whose blocks the view holds any more is not an
    // edit of it: the view was read again (a reload whose hydration has yet to
    // reach this hook), or its blocks are on their way in. Split so, it came
    // out empty and was saved as one borrowed block -- every part of a book,
    // the rest of each kept only in a recovery that then applied it
    
    // saved; the next look, with the view whole again, saves what changed.
    const held = new Set<number>();
    if (isRecord(content) && Array.isArray(content.content)) {
      for (const item of content.content) {
        const id = blockId(item);
        const known = id ? owner.current.get(id) : undefined;
        if (known !== undefined) held.add(known);
      }
    }
    {
      const sizes = split.parts.map((blocks) => blocks.length);
      const odd = sizes.some((size, index) => {
        const saved = records.current.get(index)?.savedBlocks?.length;
        return index > 0 && typeof saved === "number" && Math.abs(size - saved) > Math.max(3, saved * 0.05);
      });
      if (odd) {
        const firstOwners = split.parts.map((blocks) => {
          const id = blockId(blocks[0]);
          return id ? owner.current.get(id) ?? null : "none";
        });
        traceDocsHandoverPart({ step: "split-odd", loaded: loadedRef.current, view: Array.isArray(content.content) ? content.content.length : null,
          sizes, saved: split.parts.map((_, index) => records.current.get(index)?.savedBlocks?.length ?? null), firstOwners,
          ownerSize: owner.current.size, held: [...held].sort((a, b) => a - b) });
      }
    }
    let changed = false;
    split.parts.forEach((blocks, index) => {
      if (index > 0 && !held.has(index) && records.current.has(index)) {
        traceDocsHandoverPart({ step: "part-not-in-view", index, loaded: loadedRef.current });
        return;
      }
      for (const item of blocks) {
        const id = blockId(item);
        if (id) owner.current.set(id, index);
      }
      // Only a room taking input holds a part's edit. One still being connected
      // to holds nothing yet: the edit waits as any unsaved edit does, kept for a
      // reload, until that room reads it or turns out not to be had. A word
      // typed with the room out of reach was in no save and no kept edit.
      if (index === 0 || (optionsRef.current.partRooms && rooms.current.get(index)?.room.accepting)) return;
      const record = records.current.get(index);
      if (!record) return;
      const each = blocks.map((item) => JSON.stringify(item));
      const json = `[${each.join(",")}]`;
      if (handAt.current === null && record.saved !== "" && !pending.current.has(index)) {
        // Read as it reads now from here on; see handAt.
        record.saved = json;
        if (record.savedBlocks) record.savedBlocks = each;
        return;
      }
      const unchanged = json === record.saved ||
        (record.saved !== "" && samePartAsSaved(blocks, each, record.saved, record.savedBlocks));
      if (unchanged && json !== record.saved) {
        // Filled in by itself: compared with the part as it reads now from here on.
        record.saved = json;
        if (record.savedBlocks) record.savedBlocks = each;
      }
      if (unchanged && !pending.current.has(index)) return;
      if (unchanged) {
        pending.current.delete(index);
        pendingBlocks.current.delete(index);
        writeDocumentPartRecovery(pageId ?? "", index, null);
      } else if (pending.current.get(index) !== json) {
        traceDocsHandoverPart({ step: "part-edit", index, blocks: blocks.length, saved: record.savedBlocks?.length ?? null, loaded: loadedRef.current });
        pending.current.set(index, json);
        pendingBlocks.current.set(index, each);
        keepPartEdit(index, false);
        changed = true;
      }
    });
    if (changed) scheduleSaves();
    if (roomHolds(0)) return null;
    // A body nobody edited since it was opened or last prepared (the draft a
    // part's load or a later part's edit made) is that same body again:
    // nothing new of it is saved.
    const baseline = bodyBaseline.current;
    const blocksJson = JSON.stringify(split.parts[0]);
    const attrsJson = withoutLayoutAttrs(content.attrs);
    if (baseline && baseline.blocks === blocksJson && baseline.attrs === attrsJson) {
      return baseline.prepared ?? unchangedBody();
    }
    // Most of the body gone at once with nobody having asked is no edit of
    
    // under load: its loading placeholder, one empty paragraph, saved as the
    // body of three long documents; the part counts above had not been read
    // for it). Measured against the body as stored, or as last prepared here
    // (a reader's cut prepared once is not asked about again while its save
    // is on its way).
    const shrunkFrom = preparedBodyCount.current !== null && preparedBodyCount.current < storedBlocks.length
      ? { ...stored, content: storedBlocks.slice(0, preparedBodyCount.current) } : stored;
    if (!bulkGestureLately() && documentBodyCutNeedsSaying(shrunkFrom, { content: split.parts[0] })) {
      refuseJump(0, storedBlocks.length, split.parts[0].length, "body");
      return unchangedBody();
    }
    const prepared = bodyContentForSave(content, stored, optionsRef.current.storedForm?.(split.parts[0]) ?? split.parts[0]);
    bodyBaseline.current = { blocks: blocksJson, attrs: attrsJson, prepared };
    preparedBodyCount.current = split.parts[0].length;
    return prepared;
  }, [block, bulkGestureLately, draftIsNotTheView, keepPartEdit, pageId, partsJumped, partsReadRight, publishViewToRooms, refuseJump, roomHolds, scheduleSaves, takeBodyBaseline, viewIsThisReading]);
  /** Whether this reading's reader did something that cuts many blocks at
   * once (a part's PUT says `large_removal` on the same ground). */
  const bodyCutSaidMeant = useCallback(() => lastBulkGestureAt.current !== null, []);

  /** For a window that is closing (lib/page-closing-edits): the body's
   * blocks as the editor holds them now, and each later part that is not as
   * last saved -- its blocks now, as last saved, and as on their way. Null
   * when this view saves nothing of the parts. Synchronous. */
  const closingParts = useCallback((): DocumentPartsClosing | null => {
    const current = optionsRef.current;
    if (documentPartsCount(current.body) < 2 || blockedRef.current || !current.writable || !current.hydrated ||
        !bodyBaseline.current || !viewIsThisReading() || handAt.current === null) return null;
    const serialized = current.serialize();
    if (!serialized) return null;
    const split = splitDocumentParts(serialized, owner.current, loadedRef.current);
    if (!split.ok || !partsReadRight(split.parts)) return null;
    // Everything goes to the rooms as the window goes; a room that has
    // acknowledged all of it holds it, and the server saves it (a room left
    // with nobody in it). Only what is not known to have arrived is sent here.
    let published = false;
    try {
      published = publishViewToRooms();
    } catch {
      // The closing edits below still carry it.
    }
    const roomHoldsAll = (index: number) => {
      const live = rooms.current.get(index);
      return Boolean(live?.room.live && !live.room.unacknowledged);
    };
    const parts: Array<DocumentPartsClosing["parts"][number]> = [];
    split.parts.forEach((blocks, index) => {
      if (index === 0 || roomHoldsAll(index)) return;
      const record = records.current.get(index);
      if (!record) return;
      const each = blocks.map((item) => JSON.stringify(item));
      const sending = sendingBlocks.current.get(index) ?? null;
      const unchanged = `[${each.join(",")}]` === record.saved ||
        (record.saved !== "" && samePartAsSaved(blocks, each, record.saved, record.savedBlocks));
      if (unchanged && !sending) return;
      const parse = (list: readonly string[]) => list.map((json) => JSON.parse(json) as unknown);
      parts.push({
        index,
        current: blocks,
        saved: record.savedBlocks ? parse(record.savedBlocks) : record.saved ? (JSON.parse(record.saved) as unknown[]) : null,
        sending: sending ? parse(sending) : null,
        sendingRevision: sending ? record.revision : null,
        ...(lastBulkGestureAt.current !== null && partCountJumped(record.baseCount, blocks.length) ? { largeRemoval: true } : {}),
      });
    });
    return {
      body: split.parts[0], attrs: serialized.attrs, openedBody: openedBody.current, parts,
      bodyInRoom: roomHoldsAll(0),
      ...(published ? {} : { viewUnpublished: true as const }),
    };
  }, [partsReadRight, publishViewToRooms, viewIsThisReading]);

  const viewHoldsStored = useCallback(async (storedBody: unknown): Promise<boolean | null> => {
    const current = optionsRef.current;
    if (!current.pageId || !viewIsThisReading() || !isRecord(storedBody)) return null;
    const roomHolds = (index: number) => Boolean(rooms.current.get(index)?.room.live);
    let viewParts: readonly (readonly unknown[])[] | null = null;
    const readView = (): readonly (readonly unknown[])[] | null => {
      if (viewParts) return viewParts;
      const serialized = current.serialize();
      if (!serialized || !viewIsThisReading()) return null;
      const split = splitDocumentParts(serialized, owner.current, loadedRef.current);
      if (!split.ok) return null;
      viewParts = split.parts;
      return viewParts;
    };
    if (!roomHolds(0)) {
      const parts = readView();
      const storedBlocks = Array.isArray(storedBody.content) ? storedBody.content : null;
      if (!parts || !storedBlocks) return null;
      if (!samePartContent(parts[0] ?? [], storedBlocks)) return false;
    }
    for (const [index, record] of [...records.current.entries()].sort(([a], [b]) => a - b)) {
      if (index === 0 || roomHolds(index)) continue;
      let stored: DocumentPartResponse;
      try {
        stored = await readDocumentPart(current.pageId, index, { cache: "no-store" });
      } catch {
        return null;
      }
      // Read or saved by this view at the revision stored: it holds it.
      if (typeof stored.revision !== "number" || stored.revision <= record.revision) continue;
      const parts = readView();
      if (!parts) return null;
      const blocks = parts[index];
      const storedBlocks = stored.content?.content;
      if (!Array.isArray(blocks) || !Array.isArray(storedBlocks) || !samePartContent(blocks, storedBlocks)) return false;
    }
    return true;
  }, [viewIsThisReading]);

  const publishEdits = useCallback(() => {
    try {
      publishViewToRooms();
    } catch (error) {
      // The page save's own turn publishes again.
      console.warn("[part-room] the view could not be published", error);
    }
  }, [publishViewToRooms]);

  // The view changed (a key): its rooms hear of it within a moment, typing or
  // not. It waited for the page's own serialization, which waits for the typing
  // to pause (lib/typing-session, up to 12 s): the other member saw nothing of
  
  // sheet or two (publishViewToRooms' scoped reading); the next one waits at
  // least four times as long as the last one took, so a slow machine's typing
  // is never queued behind them.
  const publishTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const publishSpacing = useRef(ROOM_PUBLISH_MS);
  const noteViewChanged = useCallback(() => {
    if (publishTimer.current !== null || ![...rooms.current.values()].some((live) => live.room.accepting && live.inView)) return;
    publishTimer.current = setTimeout(() => {
      publishTimer.current = null;
      const started = performance.now();
      publishEdits();
      publishSpacing.current = Math.min(ROOM_PUBLISH_MOST_MS, Math.max(ROOM_PUBLISH_MS, 4 * (performance.now() - started)));
    }, publishSpacing.current);
  }, [publishEdits]);
  useEffect(() => () => {
    if (publishTimer.current !== null) clearTimeout(publishTimer.current);
    publishTimer.current = null;
  }, []);

  const importRetry = useRef<ReturnType<typeof setTimeout> | null>(null);
  const loadNextRef = useRef<(() => Promise<void>) | null>(null);
  useEffect(() => () => {
    if (importRetry.current) clearTimeout(importRetry.current);
    importRetry.current = null;
  }, []);
  const loadNext = useCallback(async () => {
    const current = optionsRef.current;
    // A view read again that this hook has yet to hear of (a render later) holds
    // the body alone while the hook still counts the parts it had: the next one
    // entered there, then every part again from the first once the reset came.
    if (!pageId || loadingRef.current || !current.hydrated || blockedRef.current || !viewIsThisReading()) return;
    const total = documentPartsCount(current.body);
    const index = loadedRef.current + 1;
    if (index >= total) return;
    loadingRef.current = true;
    publish({ loading: true });
    const started = generation.current;
    let early: LivePartRoom | null = null;
    const dropEarly = () => {
      if (!early || rooms.current.get(index) !== early) return;
      early.room.destroy();
      rooms.current.delete(index);
    };
    try {
      // Read ahead with the others when every part loads at once (loadsPartsAtOnce);
      // a read that failed there is made again here.
      const ahead = prefetched.current.get(index);
      prefetched.current.delete(index);
      readyAhead.current.delete(index);
      const part = (ahead ? await ahead : null) ?? await readDocumentPart(pageId, index, { cache: "no-store" });
      if (generation.current !== started || !viewIsThisReading()) return;
      if (part.status === "pending" || part.status === "failed") {
        // A long document opened from its first part: the import has yet to
        // read this one, and answers it soon; nothing to add meanwhile.
        publish({ importing: part.status });
        if (part.status === "pending") {
          if (importRetry.current) clearTimeout(importRetry.current);
          importRetry.current = setTimeout(() => {
            importRetry.current = null;
            if (generation.current === started) void loadNextRef.current?.();
          }, IMPORTING_PART_RETRY_MS);
        }
        return;
      }
      publish({ importing: null });
      const stored = Array.isArray(part.content?.content) ? (part.content?.content as unknown[]) : null;
      if (!stored || part.partIndex !== index || part.partCount !== total) return;
      const revision = typeof part.revision === "number" ? part.revision : 0;
      // An edit of this part kept from before a reload: back in, unless the
      // part was saved since from elsewhere (then it cannot be applied, and
      // the page says so). One whose own save had reached the server is done.
      let blocks = stored;
      let recovered = false;
      let lostToConflict = false;
      const kept = readDocumentPartRecoveries(pageId)[String(index)];
      if (kept && kept.del === 0 && Array.isArray(kept.insert) && kept.insert.length === 0) {
        // Says nothing: not an edit to bring back and save again.
        writeDocumentPartRecovery(pageId, index, null);
      } else if (kept) {
        const applied = kept.rev === revision ? applyDocumentPartRecovery(stored, kept) : null;
        if (applied && partCountJumped(stored.length, applied.length)) {
          // Came with the tab, with nobody here having done it: no edit to bring back.
          traceDocsHandoverPart({ step: "part-recovery-jump", index, stored: stored.length, applied: applied.length });
          writeDocumentPartRecovery(pageId, index, null);
        } else if (applied) {
          blocks = applied;
          recovered = true;
        } else if (kept.sending && kept.rev + 1 === revision) {
          // Its own save landed: the part is the edit.
          writeDocumentPartRecovery(pageId, index, null);
        } else {
          // Saved since from elsewhere: merged into the part as it is now
          // (mergeDocumentPartRecovery), never dropped.
          const merged = mergeDocumentPartRecovery(stored, kept);
          if (merged) {
            blocks = merged.blocks;
            recovered = true;
            traceDocsHandoverPart({ step: "part-recovery-merged", index, copies: merged.copies });
          } else {
            writeDocumentPartRecovery(pageId, index, null);
            lostToConflict = true;
          }
        }
      }
      // A part whose blocks the view holds already is not put in a second time:
      // each of them would stand twice, the copy the view holds and the part's
      // own (R482-D04 LIVE: a part's 17 tables and a paragraph put in before
      // the part loaded, by a whole view's draft merged into a view that had
      // not loaded it). In the editor the block-id plugin gives the second of
      // each a new id and takes `source` from a paragraph, so no later look
      // tells the paragraph from a new one. The view stops as a damaged one
      // does; reading the page again gives the document as stored.
      const held = optionsRef.current.viewBlockIds?.() ?? null;
      if (held && held.size > 0) {
        const twice = blocks.reduce<number>((count, item) => {
          const id = blockId(item);
          return count + (id !== null && held.has(id) ? 1 : 0);
        }, 0);
        if (twice > 0) {
          refuseJump(index, blocks.length, twice, "part-load-blocks-in-view");
          return;
        }
      }
      // Its room is joined now, while the part is laid out (seconds for a long
      // part, a minute on a busy machine): read and caught up by the time the
      
      // 60-130: a part's room read 46 s after its blocks came, its others'
      // words not shown meanwhile). It holds its reading and others' changes
      // until the part is in the view (LivePartRoom.inView).
      early = optionsRef.current.partRooms && !lostToConflict ? startRoom(index, part.partId, [], false) : null;
      // The layout kept for it (paged/layoutSnapshot), asked before anything of
      // it is owned: while that is asked the view holds none of its blocks, and
      
      // 528 blocks prepared as 1, refused by the save's count check). With no
      // layout to ask for, nothing waits.
      const askKept = recovered ? undefined : optionsRef.current.keptLayoutOf;
      const keptLayout = askKept ? await askKept(blocks, index) : null;
      if (askKept && generation.current !== started) { dropEarly(); return; }
      // Owned before they enter the editor: a save the append itself
      // triggers must read them as this part's, never as the part before.
      const added: string[] = [];
      for (const item of blocks) {
        const id = blockId(item);
        if (id && !owner.current.has(id)) {
          owner.current.set(id, index);
          added.push(id);
        }
      }
      loadedRef.current = index;
      countSourceAnchors(stored, storedAnchors.current);
      traceDocsHandoverPart({ step: "part-load", index, blocks: blocks.length, recovered, owned: blocks.length - added.length,
        ...(docsTraceOn() ? { anchorsTwice: repeatedSourceAnchors([...((optionsRef.current.serialize() as { content?: unknown[] } | null)?.content ?? []), ...blocks]) } : {}) });
      if (!optionsRef.current.appendBlocks(blocks, index, keptLayout)) {
        for (const id of added) owner.current.delete(id);
        countSourceAnchors(stored, storedAnchors.current, -1);
        loadedRef.current = index - 1;
        dropEarly();
        return;
      }
      // The part as this editor serializes it: what its later saves compare
      // with (the stored form may differ in defaults the editor fills). Read
      // part by part where that is on: the whole view serialized for each part
      // that came in, its pages measured, was 43 serializations of a longer and
      
      // way (serializeViewParts), so the count check below reads the same.
      const byParts = scopedPartsLoad() ? optionsRef.current.serializeParts?.(owner.current, index) ?? null : null;
      const serialized = byParts ? null : optionsRef.current.serialize();
      const split: DocumentPartsSplit | null = byParts ? { ok: true, parts: byParts }
        : serialized ? splitDocumentParts(serialized, owner.current, index) : null;
      if (!split || !split.ok) {
        block(split && !split.ok ? split.reason : "unreadable");
        return;
      }
      // The part must read back as it went in, and nothing else may have moved.
      partCounts.current.set(index, blocks.length);
      if (!partsReadRight(split.parts)) { dropEarly(); return; }
      const each = split.parts[index].map((item) => JSON.stringify(item));
      const json = `[${each.join(",")}]`;
      records.current.set(index, {
        partId: part.partId,
        revision,
        // A recovered edit is not what the server holds: it saves again.
        saved: recovered ? "" : json,
        savedBlocks: recovered || each.length !== stored.length ? null : each,
        baseCount: stored.length,
        recoveredBase: recovered ? optionsRef.current.editorFormOf?.(stored) ?? null : null,
      });
      publish({ loaded: index });
      if (optionsRef.current.partRooms && !lostToConflict) {
        // The part as stored is what the room is merged against, so an edit
        // recovered from before a reload is this member's over it.
        const base = recovered ? optionsRef.current.editorFormOf?.(stored) ?? stored : (JSON.parse(json) as unknown[]);
        const live = early && rooms.current.get(index) === early ? early : null;
        if (live) {
          live.base = base;
          live.inView = true;
          live.askedAt = Date.now();
          if (live.readWaiting) {
            live.readWaiting = false;
            roomRead(live);
          }
        } else {
          startRoom(index, part.partId, base);
        }
      }
      if (recovered) {
        pending.current.set(index, json);
        pendingBlocks.current.set(index, each);
        scheduleSaves();
      }
      if (lostToConflict) block("conflict");
    } catch {
      // Read again when the reader comes near the end once more.
      if (!rooms.current.get(index)?.inView) dropEarly();
    } finally {
      if (generation.current === started) {
        loadingRef.current = false;
        publish({ loading: false });
        // The next part is the next look's to decide (shouldLoadNextPart):
        // chaining it here loaded part after part while the sentinel still
        // read near, before the appended part was laid out.
        nearEnd.current = false;
        // A part read ahead (loadsPartsAtOnce) is not chained on at once
        // either: it goes in at the pace of any part loading by itself (two
        // still looks 1.5 s apart, the reader quiet for 2 s). Each part going in
        // is its blocks and its room's state (a 1.0-2.3 s task, mostly the
        // room's sync), then its pages laid out: back to back, or one per
        // still look 300 ms apart, a book's 7 later parts took the 20 s after
        
        
        // Since then the ledger holds the pages past the reader while parts are
        // still coming in (holdForArrivingParts) and lays them out together, so a
        // part read ahead goes in right after the one before while the reader is
        // quiet (no key or scroll for PART_BACKGROUND_QUIET_MS; one holds the next
        
        // 42.3/41.6 s against 49.4/44.1 s at the paced rate, editable 4.4-4.6 s
        // against 6.1-7.7 s, all parts in by 18-25 s instead of 37-45 s; the
        // longest task 4.0-4.9 s against 2.9-3.3 s.
        // `window.__docsPartsBackToBack = false`: the paced rate, as before.
        if (
          eagerRef.current
          && partsBackToBack()
          && prefetched.current.has(loadedRef.current + 1)
          && readyAhead.current.has(loadedRef.current + 1)
          && Date.now() - lastInputAt.current >= PART_BACKGROUND_QUIET_MS
        ) {
          const next = started;
          window.setTimeout(() => {
            if (generation.current === next && !loadingRef.current && !blockedRef.current) void loadNextRef.current?.();
          }, 0);
        }
      }
    }
  }, [block, pageId, partsReadRight, publish, refuseJump, roomRead, scheduleSaves, startRoom, viewIsThisReading]);
  loadNextRef.current = loadNext;

  // The next part loads when the end of what is loaded nears the viewport.
  const observer = useRef<IntersectionObserver | null>(null);
  const sentinel = useRef<HTMLElement | null>(null);
  const loadMemory = useRef<PartLoadMemory>(EMPTY_PART_LOAD_MEMORY);
  const stillLooks = useRef(0);
  const lastDocumentEnd = useRef<number | null>(null);
  const lastInputAt = useRef(0);
  const lookAtSentinelRef = useRef<(() => void) | null>(null);
  /** One look at the sentinel: the next part loads when shouldLoadNextPart says so. */
  const lookAtSentinel = useCallback(() => {
    const element = sentinel.current;
    if (!element?.isConnected || loadingRef.current) return;
    const rect = element.getBoundingClientRect();
    if (rect.width <= 0) return;
    const sentinelTop = Math.round(rect.top);
    const scrollTop = scrollTopOf(element);
    // Still in the document, wherever the reader has it scrolled to.
    const documentEnd = sentinelTop + scrollTop;
    stillLooks.current = lastDocumentEnd.current === documentEnd ? stillLooks.current + 1 : 0;
    lastDocumentEnd.current = documentEnd;
    const decision = shouldLoadNextPart(loadMemory.current, {
      sentinelTop,
      viewportHeight: window.innerHeight,
      scrollTop,
    });
    loadMemory.current = decision.memory;
    if (decision.load || shouldLoadPartInBackground({
      stillLooks: stillLooks.current,
      quietMs: Date.now() - lastInputAt.current,
      visible: document.visibilityState === "visible",
      nextIndex: loadedRef.current + 1,
      touch: window.matchMedia?.("(pointer: coarse)").matches ?? false,
      kept: Boolean(pageId) && layoutSnapshotRows(pageId!, loadedRef.current + 1).length > 0,
    })) {
      // Parts read all at once go in after the first screen (afterFirstScreen),
      // unless the reader goes further first.
      if (eagerRef.current && (optionsRef.current.firstScreenPending?.() ?? false)) {
        if (!lookHeld.current) {
          lookHeld.current = true;
          afterFirstScreenRef.current(generation.current, () => {
            lookHeld.current = false;
            lookAtSentinelRef.current?.();
          });
        }
        if (!(lastInputAt.current > 0)) return;
      }
      stillLooks.current = 0;
      nearEnd.current = true;
      void loadNext();
    }
  }, [loadNext, pageId]);
  lookAtSentinelRef.current = lookAtSentinel;
  // Anything the reader does holds the next unasked part back.
  useEffect(() => {
    const note = () => { lastInputAt.current = Date.now(); };
    const options = { capture: true, passive: true };
    const kinds = ["keydown", "pointerdown", "wheel", "scroll", "touchstart", "compositionstart"];
    for (const kind of kinds) document.addEventListener(kind, note, options);
    return () => {
      for (const kind of kinds) document.removeEventListener(kind, note, options);
    };
  }, []);
  const sentinelRef = useCallback((element: HTMLElement | null) => {
    observer.current?.disconnect();
    observer.current = null;
    nearEnd.current = false;
    loadMemory.current = EMPTY_PART_LOAD_MEMORY;
    sentinel.current = element;
    if (!element || typeof IntersectionObserver === "undefined") return;
    observer.current = new IntersectionObserver((entries) => {
      if (entries.some((entry) => entry.isIntersecting)) lookAtSentinel();
    }, { rootMargin: LOAD_AHEAD_MARGIN });
    observer.current.observe(element);
  }, [lookAtSentinel]);

  useEffect(() => {
    if (hydrated) lookAtSentinel();
  }, [hydrated, hydration, lookAtSentinel]);

  // Every part read ahead at once (loadsPartsAtOnce) the moment the first is
  // in the view; they go in from the first later part on (loadNext's chain).
  // What waits for the first screen: the parts read all at once and their
  // going in. The reader going further (a scroll, a key, a jump to a later
  // heading) does not wait.
  const afterFirstScreen = useCallback((started: number, then: () => void) => {
    const since = Date.now();
    const look = () => {
      if (generation.current !== started) return;
      const pending = optionsRef.current.firstScreenPending?.() ?? false;
      if (!pending || lastInputAt.current > since || Date.now() - since >= FIRST_SCREEN_WAIT_MAX_MS) {
        then();
        return;
      }
      setTimeout(look, FIRST_SCREEN_LOOK_MS);
    };
    look();
  }, []);
  afterFirstScreenRef.current = afterFirstScreen;
  // The parts read ahead, all at once (loadsPartsAtOnce).
  const readAllAhead = useCallback((started: number) => {
    const id = pageId;
    if (!id || eagerReadsStarted.current === started || generation.current !== started) return;
    eagerReadsStarted.current = started;
    const total = documentPartsCount(optionsRef.current.body);
    // A part already in, or being read by the reader going on (loadNext), is
    // not asked for twice.
    for (let index = 1; index < total; index += 1) {
      if (index <= loadedRef.current || (loadingRef.current && index === loadedRef.current + 1)) continue;
      const read = readDocumentPart(id, index, { cache: "no-store", priority: "low" } as RequestInit).catch(() => null);
      prefetched.current.set(index, read);
      // Ready to go in back to back (loadNext's end): a part the import has
      // yet to read waits for the look that finds it read.
      void read.then((response) => { if (response?.status === "ready") readyAhead.current.add(index); });
    }
  }, [pageId]);
  // Every part read ahead at once (loadsPartsAtOnce) once the first is in the
  // view and its first screen is in; each goes in at the pace of a part
  // loading by itself (shouldLoadPartInBackground). Read with the first screen, a book's 7 later parts
  // (each a request and a parse, then each part's pages) took the thread its
  // font plan was waiting for: answered at 17-18 s instead of 5.6 s, the pages
  // shown in the fallback faces and laid out again (a 865-page book of 8
  
  // further (a scroll, a jump to a later heading) reads them at once.
  useEffect(() => {
    if (!hydrated || !pageId || !eagerRef.current || loadedRef.current !== 0 || eagerStarted.current === generation.current) return;
    const started = generation.current;
    eagerStarted.current = started;
    afterFirstScreen(started, () => readAllAhead(started));
  }, [afterFirstScreen, hydrated, hydration, pageId, readAllAhead]);

  // The observer reports a change of intersection only: the end of what is
  // loaded already in view when the page opens (or brought there by a
  // layout pass, not a scroll) is found by asking where it is.
  useEffect(() => {
    if (!hydrated || state.loaded + 1 >= count || state.blocked) return;
    // A part with a kept layout is looked for sooner (shouldLoadPartInBackground).
    const kept = Boolean(pageId) && layoutSnapshotRows(pageId!, state.loaded + 1).length > 0;
    const timer = window.setInterval(lookAtSentinel, kept ? KEPT_PART_LOOK_MS : 1500);
    return () => window.clearInterval(timer);
  }, [count, hydrated, lookAtSentinel, pageId, state.blocked, state.loaded]);

  useEffect(() => () => observer.current?.disconnect(), []);

  const dismissRefusedText = useCallback(() => {
    if (pageId) writeRefusedPartText(pageId, null);
    publish({ refusedText: null });
  }, [pageId, publish]);

  const loadThroughBlock = useCallback(async (id: string): Promise<boolean> => {
    const started = generation.current;
    // A load already under way (the sentinel's) is waited for, at most a
    // minute; each of ours must add a part or the walk stops.
    let waits = 0;
    while (!owner.current.has(id)) {
      if (generation.current !== started || blockedRef.current) { traceDocsHandoverPart({ step: "through-stopped", id, generation: generation.current !== started, blocked: blockedRef.current }); return false; }
      if (loadingRef.current) {
        if ((waits += 1) > 600) return false;
        await new Promise((resolve) => window.setTimeout(resolve, 100));
        continue;
      }
      const before = loadedRef.current;
      if (before + 1 >= documentPartsCount(optionsRef.current.body)) { traceDocsHandoverPart({ step: "through-missing", id, loaded: before }); return false; }
      await loadNext();
      if (loadedRef.current === before) { traceDocsHandoverPart({ step: "through-stuck", id, loaded: before }); return false; }
    }
    if (generation.current !== started) traceDocsHandoverPart({ step: "through-generation", id });
    return generation.current === started;
  }, [loadNext]);

  const sendPresence = useCallback((bytes: Uint8Array) => {
    rooms.current.get(0)?.room.sendPresence(bytes);
  }, []);

  return useMemo(
    () => ({ state: { ...state, count }, sentinelRef, prepareContent, bodyCutSaidMeant, loadThroughBlock, closingParts, publishEdits, noteViewChanged, sendPresence, noteHand, dismissRefusedText, viewHoldsStored, viewIsThisReading, refuseBlocksTwice }),
    [bodyCutSaidMeant, closingParts, count, dismissRefusedText, loadThroughBlock, noteHand, noteViewChanged, prepareContent, publishEdits, refuseBlocksTwice, sendPresence, sentinelRef, state, viewHoldsStored, viewIsThisReading],
  );
}
