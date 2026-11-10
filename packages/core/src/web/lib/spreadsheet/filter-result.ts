/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_GRID_ROWS } from "./grid-sizing";
import { parseA1Range } from "./print-settings";
import pageEnvelope from "./spreadsheet-page-envelope.shared.json";

export const FILTER_RESULT_KEY = "filterResult" as const;
/** The longest `criteria` a page holds (a hex digest). */
export const MAX_FILTER_RESULT_CRITERIA_CHARS = 64;
const EXPIRED_CRITERIA = "expired";
/** The most rows a result keeps: the longest array a page holds. A filter
 *  that hides more keeps no result and is read off the criteria, as before. */
export const MAX_FILTER_RESULT_ROWS = pageEnvelope.maxArrayLength;

export interface FilterResult {
  rows: number[];
  criteria?: string;
}

type Json = Record<string, unknown>;

export interface FilterResultSheet {
  filters?: unknown;
  conditionFilters?: unknown;
  colorFilters?: unknown;
  dataTables?: unknown;
  filterResult?: unknown;
}

function isRecord(value: unknown): value is Json {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function nonEmpty(value: unknown): boolean {
  return isRecord(value) && Object.keys(value).length > 0;
}

/** Keys sorted at every level, so equal criteria spell the same. */
function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (!isRecord(value)) return value;
  return Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonical(value[key])]));
}

function digest(text: string): string {
  
  let first = 0x811c9dc5;
  let second = 0x01000193 ^ text.length;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    first = Math.imul(first ^ code, 0x01000193);
    second = Math.imul(second ^ code, 0x5bd1e995);
  }
  return (first >>> 0).toString(16).padStart(8, "0") + (second >>> 0).toString(16).padStart(8, "0");
}

/** What the sheet's filters ask, or null when none filters. A table's
 *  criteria are keyed by its id and its columns' ids, which an insert does not
 *  change; a range is not a criterion (an insert inside it filters nothing
 *  again, as in Excel). */
export function filterCriteriaSignature(sheet: FilterResultSheet): string | null {
  const tables = Array.isArray(sheet.dataTables)
    ? (sheet.dataTables as unknown[])
        .filter((table): table is Json => isRecord(table) && nonEmpty(table.columnFilters))
        .map((table) => [String(table.id ?? ""), table.columnFilters])
    : [];
  const worksheet = (["filters", "conditionFilters", "colorFilters"] as const)
    .filter((key) => nonEmpty(sheet[key]))
    .map((key) => [key, sheet[key]]);
  if (!worksheet.length && !tables.length) return null;
  return digest(JSON.stringify(canonical({ worksheet, tables })));
}

type MovableFilterSheet = FilterResultSheet & { autoFilterRef?: unknown };
const COLUMN_KEYED = ["filters", "colorFilters", "conditionFilters"] as const;













export function filterCriteriaMovedFrom(earlier: MovableFilterSheet, sheet: MovableFilterSheet): boolean {
  const now = filterCriteriaSignature(sheet);
  if (now === null) return false;
  if (filterCriteriaSignature(earlier) === now) return true;
  const before = typeof earlier.autoFilterRef === "string" ? parseA1Range(earlier.autoFilterRef) : null;
  const after = typeof sheet.autoFilterRef === "string" ? parseA1Range(sheet.autoFilterRef) : null;
  if (!before || !after || before.r1 !== after.r1 || before.r2 !== after.r2) return false;
  const keys = new Set<number>();
  for (const field of COLUMN_KEYED) {
    const map = earlier[field];
    if (isRecord(map)) for (const key of Object.keys(map)) keys.add(Number(key));
  }
  if ([...keys].some((key) => !Number.isInteger(key))) return false;
  const width = (range: { c1: number; c2: number }) => range.c2 - range.c1;
  let delta: number;
  let places: number[];
  if (after.c1 !== before.c1) {
    delta = after.c1 - before.c1;
    if (width(after) !== width(before)) return false;
    // Lines before the range: every column of it moves.
    places = [delta > 0 ? before.c1 : before.c1 + delta];
  } else {
    delta = width(after) - width(before);
    if (delta === 0) return false;
    const inside = new Set<number>([before.c1 + 1, before.c2 + 1]);
    for (const key of keys) {
      inside.add(key);
      inside.add(key + 1);
    }
    places = [...inside].filter((at) => at > before.c1 && at <= before.c2 + 1);
  }
  const shift = (index: number, at: number): number | null => {
    if (index < at) return index;
    if (delta < 0 && index < at - delta) return null;
    return index + delta;
  };
  for (const at of places) {
    if (at < 0) continue;
    const moved: Json = { ...(earlier as Json) };
    let kept = true;
    for (const field of COLUMN_KEYED) {
      const map = earlier[field];
      if (!isRecord(map)) continue;
      const next: Json = {};
      for (const key of Object.keys(map)) {
        const column = shift(Number(key), at);
        if (column === null) kept = false;
        else next[String(column)] = map[key];
      }
      moved[field] = next;
    }
    if (kept && filterCriteriaSignature(moved as FilterResultSheet) === now) return true;
  }
  return false;
}

/** `after`'s filter result when a column insert or delete only moved the
 *  criteria `before` kept it under (`filterCriteriaMovedFrom`): the same rows,
 *  under the criteria as they now read. Undefined when there is nothing to
 *  carry over (no result, or one of other criteria: it is applied again). */
export function filterResultMovedWithColumns(
  before: MovableFilterSheet,
  after: MovableFilterSheet,
): FilterResult | undefined {
  const stored = readFilterResult(before);
  const was = filterCriteriaSignature(before);
  const now = filterCriteriaSignature(after);
  if (!stored || was === null || now === null || was === now || stored.criteria !== was) return undefined;
  if (!filterCriteriaMovedFrom(before, after)) return undefined;
  return { rows: stored.rows, criteria: now };
}

/** The stored result as the page holds it, or null when it holds none that
 *  reads (rows sorted, unique, in the grid). */
export function readFilterResult(sheet: FilterResultSheet): FilterResult | null {
  const raw = sheet.filterResult;
  if (!isRecord(raw) || !Array.isArray(raw.rows)) return null;
  const rows: number[] = [];
  for (const entry of raw.rows) {
    if (typeof entry !== "number" || !Number.isInteger(entry) || entry < 0 || entry >= MAX_GRID_ROWS) return null;
    if (rows.length && entry <= rows[rows.length - 1]!) return null;
    rows.push(entry);
  }
  const criteria = raw.criteria;
  if (criteria !== undefined && (typeof criteria !== "string" || criteria.length > MAX_FILTER_RESULT_CRITERIA_CHARS)) {
    return null;
  }
  return criteria === undefined ? { rows } : { rows, criteria };
}

/** The rows the sheet's filters hide now: the stored result while the sheet
 *  still asks what it asked when it was applied (or as imported), else null
 *  (the filter is to be applied). */
export function storedFilterHiddenRows(sheet: FilterResultSheet): ReadonlySet<number> | null {
  const stored = readFilterResult(sheet);
  if (!stored) return null;
  const criteria = filterCriteriaSignature(sheet);
  if (criteria === null) return null;
  if (stored.criteria !== undefined && stored.criteria !== criteria) return null;
  return new Set(stored.rows);
}

/** Whether the sheet's stored result is not the one to keep: there is none
 *  while a filter asks, its filters changed since it was applied, or no
 *  filter asks any more. */
export function filterResultStale(sheet: FilterResultSheet): boolean {
  const criteria = filterCriteriaSignature(sheet);
  if (criteria === null) return sheet.filterResult !== undefined;
  const stored = readFilterResult(sheet);
  return !stored || (stored.criteria !== undefined && stored.criteria !== criteria);
}

/** Whether the sheet's filters are to be applied now and their result kept:
 *  they ask something else than `previousCriteria` (what the sheet asked when
 *  the caller last looked; undefined when it has not), or the sheet keeps a
 *  result that is not theirs. A sheet that never kept a result stays as it
 *  was until its filters change: nothing is written on opening an old page,
 *  and an edited value is not what applies a filter. */
export function filtersToApply(sheet: FilterResultSheet, previousCriteria: string | null | undefined): boolean {
  if (previousCriteria !== undefined && previousCriteria !== filterCriteriaSignature(sheet)) return true;
  return sheet.filterResult !== undefined && filterResultStale(sheet);
}

/** The sheet with its result expired (Data > Reapply): whoever reads it next
 *  applies its filters and keeps what they hide. Never stored. */
export function withFilterResultExpired<S extends FilterResultSheet>(sheet: S): S {
  return { ...sheet, filterResult: { rows: [], criteria: EXPIRED_CRITERIA } } as S;
}

/** The sheet keeping `rows` as its filters' result (applied now); without
 *  any filter, the sheet without a result. The same sheet when nothing
 *  changes. */
export function withFilterResult<S extends FilterResultSheet>(sheet: S, rows: Iterable<number>): S {
  const criteria = filterCriteriaSignature(sheet);
  if (criteria === null) return withoutFilterResult(sheet);
  const sorted = [...new Set(rows)]
    .filter((row) => Number.isInteger(row) && row >= 0 && row < MAX_GRID_ROWS)
    .sort((left, right) => left - right);
  if (sorted.length > MAX_FILTER_RESULT_ROWS) return withoutFilterResult(sheet);
  const stored = readFilterResult(sheet);
  if (
    stored &&
    stored.criteria === criteria &&
    stored.rows.length === sorted.length &&
    stored.rows.every((row, index) => row === sorted[index])
  ) {
    return sheet;
  }
  return { ...sheet, filterResult: { rows: sorted, criteria } } as S;
}

/** The sheet with no stored result: the next reader applies its filters
 *  (Data > Reapply). */
export function withoutFilterResult<S extends FilterResultSheet>(sheet: S): S {
  if (sheet.filterResult === undefined) return sheet;
  const next = { ...sheet };
  delete next.filterResult;
  return next;
}

/** An imported result (rows, no criteria) takes the sheet's criteria, so a
 *  criterion changed later is told apart from the ones the file was saved
 *  with. */
export function withFilterResultCriteriaAdopted<S extends FilterResultSheet>(sheet: S): S {
  const stored = readFilterResult(sheet);
  if (!stored || stored.criteria !== undefined) return sheet;
  const criteria = filterCriteriaSignature(sheet);
  if (criteria === null) return withoutFilterResult(sheet);
  return { ...sheet, filterResult: { rows: stored.rows, criteria } } as S;
}

/** The stored result after a row insert or delete moved its rows (`remap`
 *  gives a row's new index, or null for a deleted row); an inserted row is
 *  shown. Undefined when the sheet holds none. */
export function remappedFilterResult(
  raw: unknown,
  remap: (row: number) => number | null,
): FilterResult | undefined {
  const stored = readFilterResult({ filterResult: raw });
  if (!stored) return undefined;
  const rows = stored.rows
    .map(remap)
    .filter((row): row is number => row !== null && row >= 0 && row < MAX_GRID_ROWS)
    .sort((left, right) => left - right);
  return stored.criteria === undefined ? { rows } : { rows, criteria: stored.criteria };
}
