/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { pageDocumentJsonEquals } from "@mrnote/core/web/lib/page-save-acknowledgement.ts";

/** The page's annotation rows before and after one row edit (every managed
 * object -- comment, shape, ink, stamp, link, text box -- is a row). */
export interface PdfRowStepRows {
  readonly before: readonly JsonValue[];
  readonly after: readonly JsonValue[];
}

function rowId(row: JsonValue): string | null {
  if (!row || typeof row !== "object" || Array.isArray(row)) return null;
  const id = (row as { id?: unknown }).id;
  return typeof id === "string" && id.trim() ? id : null;
}

/** A row as a step compares it: as JSON holds it (a key the window left
 * undefined is none, as the page saves it), and without what the reader
 * reads of the PDF's own object it stands for (`nativeSource`: its flags,
 * whether it has an appearance), read again with every saved revision, not
 * changed by an edit. */
function comparable(row: JsonValue | null): JsonValue | null {
  if (!row || typeof row !== "object" || Array.isArray(row)) return row;
  const rest: Record<string, JsonValue> = { ...(row as Record<string, JsonValue>) };
  delete rest.nativeSource;
  return JSON.parse(JSON.stringify(rest)) as JsonValue;
}

function sameRow(left: JsonValue | null, right: JsonValue | null): boolean {
  return pageDocumentJsonEquals(comparable(left), comparable(right));
}

function byId(rows: readonly JsonValue[]): Map<string, JsonValue> | null {
  const map = new Map<string, JsonValue>();
  for (const row of rows) {
    const id = rowId(row);
    if (!id || map.has(id)) return null;
    map.set(id, row);
  }
  return map;
}










export function pdfRowStepTarget(
  current: readonly JsonValue[],
  step: PdfRowStepRows,
  direction: "undo" | "redo",
): JsonValue[] | null {
  // A step that changed none of these rows (a form fill's annotations):
  // they stay as they are now.
  if (pageDocumentJsonEquals(step.before as JsonValue[], step.after as JsonValue[])) return [...current];
  const from = byId(direction === "undo" ? step.after : step.before);
  const toRows = direction === "undo" ? step.before : step.after;
  const to = byId(toRows);
  const now = byId(current);
  if (!from || !to || !now) return null;
  const changed = new Set<string>();
  for (const id of new Set([...from.keys(), ...to.keys()])) {
    if (!sameRow(from.get(id) ?? null, to.get(id) ?? null)) changed.add(id);
  }
  // (Rows that differ only in what the reader read of their PDF object:
  // the step changed none of them.)
  if (changed.size === 0) return [...current];
  for (const id of changed) {
    if (!sameRow(now.get(id) ?? null, from.get(id) ?? null)) return null;
  }
  return laid(current, toRows, to, now, changed, rowId);
}







export function pdfRestoreOriginsLeaveRows(origins: readonly (string | null | undefined)[]): boolean {
  return origins.length > 0 && origins.every((origin) => Boolean(origin) && !origin!.startsWith("rows:"));
}

/**
 * An Undo (or Redo) read before it is planned, as the chain of steps shown
 * ahead holds it (the rows of the newest one in the chain are the rows the
 * page restores until it is taken or refused):
 * - a step of rows alone (`rows`): the rows it leaves, shown at once;
 * - a step that moves none of the page's rows (`leavesRows`: a line's
 *   words): the rows as they are -- it holds its place among the steps read
 *   at their press;
 * - null for a step of rows that no longer stand as it left them, and for
 *   one that moves rows without being rows alone (its edit changed a
 *   collection beside them, or it is several records): it is shown when it
 *   is staged, from the page as it then stands, and holds no place here.
 */
export function pdfRowsAheadOfStep(
  current: readonly JsonValue[],
  step: Readonly<{ rows: PdfRowStepRows | null; leavesRows: boolean }>,
  direction: "undo" | "redo",
): Readonly<{ shown: JsonValue[]; ahead: boolean }> | null {
  if (step.rows) {
    const shown = pdfRowStepTarget(current, step.rows, direction);
    return shown ? { shown, ahead: true } : null;
  }
  // (Held as "rows unchanged", a step that takes rows away kept them: the
  // rows its staging put back were restored as the chain's, the undone mark
  // stayed the window's, and the page's next row edit wrote it again.)
  return step.leavesRows ? { shown: [...current], ahead: false } : null;
}

/** `current` with the entries `changed` names as `to` holds them: in place
 * when it holds them, dropped when `to` has none, and one only `to` has
 * after the entry it followed there. */
function laid(current: readonly JsonValue[], toRows: readonly JsonValue[], to: ReadonlyMap<string, JsonValue>,
  now: ReadonlyMap<string, JsonValue>, changed: ReadonlySet<string>,
  keyOf: (entry: JsonValue) => string | null): JsonValue[] {
  const result: JsonValue[] = [];
  for (const row of current) {
    const id = keyOf(row)!;
    if (!changed.has(id)) result.push(row);
    else if (to.has(id)) result.push(to.get(id)!);
  }
  toRows.forEach((row, index) => {
    const id = keyOf(row)!;
    if (!changed.has(id) || now.has(id)) return;
    let at = 0;
    for (let previous = index - 1; previous >= 0; previous -= 1) {
      const found = result.findIndex((candidate) => keyOf(candidate) === keyOf(toRows[previous]!));
      if (found >= 0) { at = found + 1; break; }
    }
    result.splice(at, 0, row);
  });
  return result;
}

/**
 * An edit of a collection whose entries `keyOf` tells apart, laid onto what
 * the page holds now when it was made on something else (a queue planned
 * again on a version another window saved): the entries it changed as it
 * left them, the ones it took away gone, every other entry as the page
 * holds it. Where the page changed one of the edit's own entries meanwhile,
 * the edit's stands: it is this window's unsaved work, and the rest of that
 * save is kept. Null when the entries cannot be told apart.
 */
export function pdfKeyedEntriesRebased(
  current: readonly JsonValue[],
  step: PdfRowStepRows,
  keyOf: (entry: JsonValue) => string | null,
  same: (left: JsonValue | null, right: JsonValue | null) => boolean = pageDocumentJsonEquals,
): JsonValue[] | null {
  const keyed = (rows: readonly JsonValue[]) => {
    const map = new Map<string, JsonValue>();
    for (const row of rows) {
      const id = keyOf(row);
      if (id === null || map.has(id)) return null;
      map.set(id, row);
    }
    return map;
  };
  const from = keyed(step.before), to = keyed(step.after), now = keyed(current);
  if (!from || !to || !now) return null;
  const changed = new Set<string>();
  for (const id of new Set([...from.keys(), ...to.keys()])) {
    if (!same(from.get(id) ?? null, to.get(id) ?? null)) changed.add(id);
  }
  return changed.size === 0 ? [...current] : laid(current, step.after, to, now, changed, keyOf);
}

/** A row edit laid onto the rows the page holds now (pdfKeyedEntriesRebased,
 * rows told apart by id and compared as a step compares them). */
export function pdfRowEditRebased(current: readonly JsonValue[], step: PdfRowStepRows): JsonValue[] | null {
  if (pageDocumentJsonEquals(step.before as JsonValue[], step.after as JsonValue[])) return [...current];
  return pdfKeyedEntriesRebased(current, step, rowId, sameRow);
}

/** What names an entry of a page collection a row edit changes beside its
 * rows (default: its id): a page's turn, size, crop and move its page as
 * uploaded, a label range the page it starts at. */
const PDF_COLLECTION_ENTRY_NAME: Readonly<Record<string, string>> = {
  pageMoves: "page", pageTransforms: "page", pageSizes: "page", cropBoxes: "page", pageLabelRanges: "start",
};

/**
 * An edit of one of the page's collections (`key`) laid onto what the page
 * holds there now, when another save changed the same collection meanwhile:
 * entry by entry (pdfKeyedEntriesRebased), the deleted pages as a set, the
 * document's decorations (headers, footers, watermark, page numbers) one
 * property at a time -- so a page one window deleted and a page another
 * turned, moved or deleted both stand. Undefined when the collection cannot
 * be told entry by entry (labels the page names none of on one side). The
 * values are compared as given: pass them as a step compares them.
 */
export function pdfCollectionEditRebased(key: string, current: JsonValue,
  step: Readonly<{ before: JsonValue; after: JsonValue }>): JsonValue | undefined {
  const { before, after } = step;
  if (pageDocumentJsonEquals(before, after)) return current;
  if (pageDocumentJsonEquals(current, before)) return after;
  const record = (value: JsonValue | null | undefined) =>
    value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, JsonValue> : null;
  if (key === "pdfDecorations") {
    const now = record(current), from = record(before), to = record(after);
    if (!now || !from || !to) return undefined;
    const next: Record<string, JsonValue> = { ...now };
    for (const name of new Set([...Object.keys(from), ...Object.keys(to)])) {
      if (pageDocumentJsonEquals(from[name] ?? null, to[name] ?? null)) continue;
      if (to[name] === undefined) delete next[name];
      else next[name] = to[name]!;
    }
    return next;
  }
  if (!Array.isArray(current) || !Array.isArray(before) || !Array.isArray(after)) return undefined;
  if (key === "deletedPages") {
    const was = new Set(before.map(Number)), is = new Set(after.map(Number)), next = new Set(current.map(Number));
    for (const page of is) if (!was.has(page)) next.add(page);
    for (const page of was) if (!is.has(page)) next.delete(page);
    return [...next].filter((page) => Number.isInteger(page)).sort((left, right) => left - right);
  }
  const name = PDF_COLLECTION_ENTRY_NAME[key] ?? "id";
  const keyOf = (entry: JsonValue) => {
    const held = record(entry)?.[name];
    return typeof held === "string" && held ? `s:${held}`
      : typeof held === "number" && Number.isFinite(held) ? `n:${held}` : null;
  };
  return pdfKeyedEntriesRebased(current, { before, after }, keyOf) ?? undefined;
}

/**
 * A row edit as the step of another edit takes it back when it joined that
 * step (a text box written after the delete of the line that became it):
 * only the rows it wrote itself. The first row edit of a page also names,
 * among the rows it leaves, what the reader has read of the PDF's own
 * objects since the page was saved (its comments). Taken back with the box
 * they were on the page again a moment later, and the step's Redo, which
 * then found them where it had none, refused the whole step ("rows moved").
 * A new row `owns` does not name is one the page held already as far as
 * the step goes: it stands in its `before` too (in the place the edit left
 * it, so the box comes back after the row it followed).
 */
export function pdfRowStepOfOwnRows<Step extends PdfRowStepRows>(step: Step, owns: (id: string) => boolean): Step {
  const before = new Set<string>();
  for (const row of step.before) {
    const id = rowId(row);
    if (id) before.add(id);
  }
  const read = step.after.filter((row) => {
    const id = rowId(row);
    return id !== null && !before.has(id) && !owns(id);
  });
  return read.length === 0 ? step : { ...step, before: [...step.before, ...read] };
}

/** A row edit with the page collections it changed beside its rows (its
 * companions: a text box's write changes the page's boxes with its rows). */
export interface PdfRowStepRowsAndCompanions extends PdfRowStepRows {
  readonly companions?: Readonly<{
    before: Readonly<Record<string, JsonValue>>;
    after: Readonly<Record<string, JsonValue>>;
  }>;
}

/** Two edits of one list of entries told apart by id, the second made right
 * after the first, as one: before the first, after the second. */
function joinedLists(runBefore: readonly JsonValue[], runAfter: readonly JsonValue[],
  nextBefore: readonly JsonValue[], nextAfter: readonly JsonValue[],
): { before: JsonValue[]; after: JsonValue[] } | null {
  const left = byId(runAfter);
  const first = byId(runBefore);
  const found = byId(nextBefore);
  if (!left || !first || !found || !byId(nextAfter)) return null;
  const changed = new Set<string>();
  for (const id of new Set([...left.keys(), ...first.keys()])) {
    if (!sameRow(left.get(id) ?? null, first.get(id) ?? null)) changed.add(id);
  }
  return {
    before: changed.size === 0 ? [...nextBefore] : laid(nextBefore, runBefore, first, found, changed, rowId),
    after: [...nextAfter],
  };
}



















export function pdfRowStepRunJoined<Step extends PdfRowStepRowsAndCompanions>(
  run: Step,
  next: PdfRowStepRowsAndCompanions,
  collections: readonly string[] = [],
): Step | null {
  const rows = joinedLists(run.before, run.after, next.before, next.after);
  if (!rows) return null;
  const side = (companions: PdfRowStepRowsAndCompanions["companions"], key: string) =>
    companions && companions.after[key] !== undefined
      ? { before: companions.before[key] ?? [], after: companions.after[key]! } : null;
  const keys = new Set([...Object.keys(run.companions?.after ?? {}), ...Object.keys(next.companions?.after ?? {})]);
  const before: Record<string, JsonValue> = {};
  const after: Record<string, JsonValue> = {};
  for (const key of keys) {
    if (!collections.includes(key)) return null;
    const ran = side(run.companions, key);
    const then = side(next.companions, key);
    if ([ran, then].some((each) => each && !(Array.isArray(each.before) && Array.isArray(each.after)))) return null;
    // (Changed by one of them only: as that one left it.)
    const joined = ran && then
      ? joinedLists(ran.before as JsonValue[], ran.after as JsonValue[], then.before as JsonValue[], then.after as JsonValue[])
      : ran ?? then;
    if (!joined) return null;
    before[key] = joined.before;
    after[key] = joined.after;
  }
  return { ...run, ...rows, ...(keys.size ? { companions: { before, after } } : {}) };
}
