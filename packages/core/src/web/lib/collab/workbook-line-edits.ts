/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";
import type { JsonValue } from "./json-crdt";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "../spreadsheet/grid-sizing";
import {
  mapCellKeyThroughLineEdit,
  type SheetLineEdit,
} from "../spreadsheet/structural";

/** The journal's top-level type in each shard document. */
export const WORKBOOK_LINE_EDIT_JOURNAL = "workbook-line-edits";
/** Entries kept: a receiver only needs the edits that arrive while it is
 * connected, never a history. */
export const WORKBOOK_LINE_EDIT_JOURNAL_LIMIT = 32;
const MAX_LINE_EDIT_COUNT = 1_048_576;
const INTENT_TTL_MS = 5_000;
const OWN_WRITE_RECORD_LIMIT = 128;
const OWN_WRITE_RECORD_TTL_MS = 120_000;

export interface WorkbookLineEdit extends SheetLineEdit {
  id: string;
  sheetId: string;
}

export interface WorkbookLineEditEntry extends WorkbookLineEdit {
  /** The author's Yjs client id in this shard document. */
  author: number;
  /** Base64 Yjs state vector: what the author's editor had shown. */
  seen: string;
}

function isLineEdit(value: unknown): value is WorkbookLineEdit {
  if (!value || typeof value !== "object") return false;
  const entry = value as Record<string, unknown>;
  return (
    typeof entry.id === "string" && entry.id.length > 0 && entry.id.length <= 64 &&
    typeof entry.sheetId === "string" && entry.sheetId.length > 0 && entry.sheetId.length <= 200 &&
    (entry.axis === "row" || entry.axis === "col") &&
    (entry.op === "insert" || entry.op === "delete") &&
    Number.isSafeInteger(entry.at) && (entry.at as number) >= 0 && (entry.at as number) < MAX_LINE_EDIT_COUNT &&
    Number.isSafeInteger(entry.count) && (entry.count as number) >= 1 && (entry.count as number) <= MAX_LINE_EDIT_COUNT
  );
}

function isLineEditEntry(value: unknown): value is WorkbookLineEditEntry {
  if (!isLineEdit(value)) return false;
  const entry = value as unknown as Record<string, unknown>;
  return Number.isSafeInteger(entry.author) && typeof entry.seen === "string" && entry.seen.length <= 16_384;
}

function newLineEditId(): string {
  const random = globalThis.crypto?.randomUUID?.();
  return random ?? `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

// ── Author side: the editor announces an edit, the session journals it ──

/** The sheet's extent right after the edit: the projection that carries
 * the edit is the first one of that sheet with exactly this extent. */
export interface WorkbookLineEditExtent {
  rowCount: number;
  columnCount: number;
}

const pendingIntents = new Map<string, Array<{ edit: WorkbookLineEdit; extent: WorkbookLineEditExtent; at: number }>>();

/** The editor made `edit` on page `pageId`, leaving the sheet at `extent`;
 * the next local projection the session publishes for that page with the
 * sheet at that extent carries it into the journal. */
export function announceWorkbookLineEdit(
  pageId: string,
  edit: Omit<WorkbookLineEdit, "id">,
  extent: WorkbookLineEditExtent,
): WorkbookLineEdit | null {
  const candidate = { ...edit, id: newLineEditId() };
  if (!pageId || !isLineEdit(candidate)) return null;
  const list = pendingIntents.get(pageId) ?? [];
  list.push({ edit: candidate, extent: { rowCount: extent.rowCount, columnCount: extent.columnCount }, at: Date.now() });
  pendingIntents.set(pageId, list);
  return candidate;
}

/** The announced, not yet journaled edits of `pageId` that `extentOf` (a
 * sheet id -> its extent in the projection being published, or null when
 * that projection does not hold the sheet) confirms, in order. Every intent
 * of the page is consumed: one its projection never carried is dropped
 * rather than attached to a later, unrelated one. */
export function takeWorkbookLineEditIntents(
  pageId: string,
  extentOf: (sheetId: string) => WorkbookLineEditExtent | null,
): WorkbookLineEdit[] {
  const list = pendingIntents.get(pageId);
  if (!list) return [];
  pendingIntents.delete(pageId);
  const now = Date.now();
  return list.filter((intent) => {
    if (now - intent.at > INTENT_TTL_MS) return false;
    const extent = extentOf(intent.edit.sheetId);
    return extent !== null && extent.rowCount === intent.extent.rowCount && extent.columnCount === intent.extent.columnCount;
  }).map((intent) => intent.edit);
}

function encodeBase64(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function decodeBase64(text: string): Uint8Array | null {
  try {
    const binary = atob(text);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    return bytes;
  } catch {
    return null;
  }
}

export function encodeLineEditSeen(stateVector: Uint8Array): string {
  return encodeBase64(stateVector);
}

/** The author's state vector of an entry; an unreadable one reads as empty
 * (nothing seen), which makes a writer re-issue its recent writes -- a
 * compare-and-write that is a no-op where they already are right. */
export function decodeLineEditSeen(entry: Pick<WorkbookLineEditEntry, "seen">): Map<number, number> {
  const bytes = decodeBase64(entry.seen);
  if (!bytes) return new Map();
  try {
    return Y.decodeStateVector(bytes);
  } catch {
    return new Map();
  }
}

function journalOf(doc: Y.Doc): Y.Array<unknown> {
  return doc.getArray<unknown>(WORKBOOK_LINE_EDIT_JOURNAL);
}

/** Append `entries` (inside the caller's transaction) and keep the journal
 * bounded. */
export function appendWorkbookLineEdits(doc: Y.Doc, entries: readonly WorkbookLineEditEntry[]): void {
  if (entries.length === 0) return;
  const journal = journalOf(doc);
  journal.push(entries.map((entry) => ({ ...entry })));
  const excess = journal.length - WORKBOOK_LINE_EDIT_JOURNAL_LIMIT;
  if (excess > 0) journal.delete(0, excess);
}

/** The journal's valid entries, in order. */
export function readWorkbookLineEdits(doc: Y.Doc): WorkbookLineEditEntry[] {
  return journalOf(doc).toArray().filter(isLineEditEntry);
}

/** This client's own clock in `doc`: the clock its next item takes. */
export function ownWorkbookShardClock(doc: Y.Doc): number {
  return Y.decodeStateVector(Y.encodeStateVector(doc)).get(doc.clientID) ?? 0;
}

// ── Receiver side: the editor follows edits that arrived from others ──

const REMOTE_EDIT_TTL_MS = 5_000;
const PUBLISHED_EDIT_TTL_MS = 600_000;
const PUBLISHED_EDIT_LIMIT = 256;
const remoteEdits = new Map<string, Array<{ edit: WorkbookLineEdit; at: number }>>();
const remoteListeners = new Map<string, Set<() => void>>();





const publishedEdits = new Map<string, Map<string, number>>();

/** Records `id` as handed to the editor (or this window's own); false when
 * it already was. */
function notePublishedEdit(pageId: string, id: string): boolean {
  const now = Date.now();
  const published = publishedEdits.get(pageId) ?? new Map<string, number>();
  for (const [known, at] of published) if (now - at > PUBLISHED_EDIT_TTL_MS) published.delete(known);
  if (published.has(id)) return false;
  published.set(id, now);
  for (const known of published.keys()) {
    if (published.size <= PUBLISHED_EDIT_LIMIT) break;
    published.delete(known);
  }
  publishedEdits.set(pageId, published);
  return true;
}

/** This window journaled `edits` itself: they are never followed as a
 * collaborator's, also when they come back under the client id of a room
 * session it has reopened since. */
export function markOwnWorkbookLineEdits(pageId: string, edits: Iterable<Pick<WorkbookLineEdit, "id">>): void {
  for (const edit of edits) notePublishedEdit(pageId, edit.id);
}

/** A collaborator's edit arrived in page `pageId`'s rooms. */
export function publishRemoteWorkbookLineEdit(pageId: string, edit: WorkbookLineEdit): void {
  if (!notePublishedEdit(pageId, edit.id)) return;
  const list = remoteEdits.get(pageId) ?? [];
  list.push({
    edit: { id: edit.id, sheetId: edit.sheetId, axis: edit.axis, op: edit.op, at: edit.at, count: edit.count },
    at: Date.now(),
  });
  remoteEdits.set(pageId, list.slice(-WORKBOOK_LINE_EDIT_JOURNAL_LIMIT));
  for (const listener of remoteListeners.get(pageId) ?? []) listener();
}

/** The collaborators' edits of `pageId` not yet followed, in order. The
 * editor takes them with the delivery that carries their moved cells: when
 * `shown` says the workbook being loaded does not show them yet, they stay
 * queued for the load that does (and none are returned). One no delivery
 * brought within a few seconds is never applied to a later, unrelated one. */
export function takeRemoteWorkbookLineEdits(
  pageId: string,
  shown?: (edits: readonly WorkbookLineEdit[]) => boolean,
): WorkbookLineEdit[] {
  const now = Date.now();
  const list = (remoteEdits.get(pageId) ?? []).filter((item) => now - item.at <= REMOTE_EDIT_TTL_MS);
  const edits = list.map((item) => item.edit);
  if (edits.length > 0 && shown && !shown(edits)) {
    remoteEdits.set(pageId, list);
    return [];
  }
  remoteEdits.delete(pageId);
  return edits;
}

/** Whether a sheet that went from `before` to `after` shows `edits` (the
 * edits of that sheet): its extent moved the way they move it, on each axis
 * they change. A load that left the extent where it was does not hold them
 * yet, or holds them already — unless no load could move it (below). */
export function workbookLineEditsShownBy(
  edits: readonly SheetLineEdit[],
  before: WorkbookLineEditExtent,
  after: WorkbookLineEditExtent,
): boolean {
  for (const axis of ["row", "col"] as const) {
    const expected = edits.reduce((sum, edit) => edit.axis === axis ? sum + (edit.op === "insert" ? edit.count : -edit.count) : sum, 0);
    if (expected === 0) continue;
    const moved = axis === "row" ? after.rowCount - before.rowCount : after.columnCount - before.columnCount;
    if (Math.sign(moved) === Math.sign(expected)) continue;
    // Edits the extent cannot show: an insert into a sheet already at the end
    // of the address space (the lines move, the extent stays at the end), and
    // a delete wholly past the stored extent (the address space below moves,
    // the extent does not). Waiting for an extent change, the entry being
    // typed below them never followed.
    const extent = axis === "row" ? before.rowCount : before.columnCount;
    const ceiling = axis === "row" ? MAX_GRID_ROWS : MAX_GRID_COLUMNS;
    const unshowable = moved === 0 && edits
      .filter((edit) => edit.axis === axis)
      .every((edit) => (edit.op === "insert" ? extent >= ceiling : edit.at >= extent));
    if (!unshowable) return false;
  }
  return true;
}

export function subscribeRemoteWorkbookLineEdits(pageId: string, listener: () => void): () => void {
  const set = remoteListeners.get(pageId) ?? new Set();
  set.add(listener);
  remoteListeners.set(pageId, set);
  return () => {
    set.delete(listener);
    if (set.size === 0) remoteListeners.delete(pageId);
  };
}

// ── A writer's own recent cell writes, for re-issuing ──

type CellValue = Record<string, JsonValue>;

interface OwnWriteRecord {
  /** Own clock range [start, end) of the transaction that wrote it. */
  start: number;
  end: number;
  /** Pre-edit address -> the entry written (null: cleared). */
  cells: Map<string, CellValue | null>;
  /** The transaction carried a line edit of this writer's own. */
  structural: boolean;
  time: number;
}

/** A shard's own cell writes, recent enough to have been concurrent with a
 * collaborator's line edit, kept at current addresses. */
export class WorkbookOwnCellWrites {
  private records: OwnWriteRecord[] = [];

  record(start: number, end: number, cells: Map<string, CellValue | null>, structural = false): void {
    if (end <= start || (cells.size === 0 && !structural)) return;
    this.records.push({ start, end, cells, structural, time: Date.now() });
    this.prune();
  }

  /** Whether this writer made a line edit the other author had not seen. */
  hasUnseenStructural(seenClock: number): boolean {
    return this.records.some((record) => record.structural && record.end > seenClock);
  }

  /** Writes the author had not seen (own clock at or past `seenClock`),
   * merged in write order. */
  unseen(seenClock: number): Map<string, CellValue | null> {
    const merged = new Map<string, CellValue | null>();
    for (const record of this.records) {
      if (record.structural || record.end <= seenClock) continue;
      for (const [key, value] of record.cells) {
        merged.delete(key);
        merged.set(key, value);
      }
    }
    return merged;
  }

  /** The edit moved every recorded address. Writes the author had not seen
   * are replaced by their re-issue (`reissued`, own clock range
   * [start, end)); seen writes keep their records at the new addresses. */
  followLineEdit(edit: SheetLineEdit, seenClock: number, reissue: { start: number; end: number; cells: Map<string, CellValue | null> } | null): void {
    const kept: OwnWriteRecord[] = [];
    for (const record of this.records) {
      if (!record.structural && record.end > seenClock && reissue) continue;
      const cells = new Map<string, CellValue | null>();
      for (const [key, value] of record.cells) {
        const moved = mapCellKeyThroughLineEdit(key, edit);
        if (moved !== null) cells.set(moved, value);
      }
      kept.push({ ...record, cells });
    }
    this.records = kept;
    if (reissue && reissue.end > reissue.start && reissue.cells.size > 0) {
      this.records.push({ ...reissue, structural: false, time: Date.now() });
    }
    this.prune();
  }

  private prune(): void {
    const now = Date.now();
    this.records = this.records
      .filter((record) => now - record.time <= OWN_WRITE_RECORD_TTL_MS)
      .slice(-OWN_WRITE_RECORD_LIMIT);
  }
}

function recordOf(value: unknown): Record<string, JsonValue> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, JsonValue> : {};
}

/**
 * The cells a local merge wrote: entries the editor changed from `base` to
 * `next` (cheap: unchanged cells keep their objects) that the room really
 * took from `room` to `merged` (a collaborator's cell the editor had rendered
 * is in both and is not this window's write). Null: the cell was cleared.
 */
export function ownWorkbookCellWrites(
  base: unknown,
  next: unknown,
  room: unknown,
  merged: unknown,
  equals: (left: JsonValue, right: JsonValue) => boolean,
): Map<string, CellValue | null> {
  const from = recordOf(base), to = recordOf(next), before = recordOf(room), after = recordOf(merged);
  const written = new Map<string, CellValue | null>();
  const consider = (key: string) => {
    const was = Object.hasOwn(from, key) ? from[key] : undefined;
    const now = Object.hasOwn(to, key) ? to[key] : undefined;
    if (was === now || (was !== undefined && now !== undefined && equals(was, now))) return;
    const roomWas = Object.hasOwn(before, key) ? before[key] : undefined;
    const roomNow = Object.hasOwn(after, key) ? after[key] : undefined;
    if (roomWas === roomNow || (roomWas !== undefined && roomNow !== undefined && equals(roomWas, roomNow))) return;
    if (roomNow === undefined) written.set(key, null);
    else if (roomNow && typeof roomNow === "object" && !Array.isArray(roomNow)) written.set(key, roomNow as CellValue);
  };
  for (const key of Object.keys(to)) consider(key);
  for (const key of Object.keys(from)) if (!Object.hasOwn(to, key)) consider(key);
  return written;
}

/** Test seam: forget every queued intent and remote edit. */
export function resetWorkbookLineEditQueuesForTests(): void {
  pendingIntents.clear();
  remoteEdits.clear();
  publishedEdits.clear();
}
