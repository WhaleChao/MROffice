/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  COLUMN_STYLE_RUNS_KEY,
  expandLineStyleRuns,
  foldLineStyles,
  lineKeySpan,
  MAX_COLUMN_LINE_STYLES,
  withColumnStylesExpanded,
  withFoldedColumnStyles,
  type LineStyleMap,
} from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";

type ColumnMap = Record<string, unknown>;
type Sheetish = Record<string, unknown>;

/** Each map a column range fills, as the page names it, and its run field. */
const COLUMN_MAPS: ReadonlyArray<readonly [field: "columnWidths" | "bestFitColumns" | "nativeGrid.columns", runs: string]> = [
  ["columnWidths", "columnWidthRuns"],
  ["bestFitColumns", "bestFitColumnRuns"],
  ["nativeGrid.columns", "nativeGridColumnRuns"],
];
/** `manuallyHiddenCols` keeps its runs as `[first, last]` pairs. */
export const HIDDEN_COLUMN_RUNS_KEY = "hiddenColumnRuns";




const ROW_LISTS: ReadonlyArray<readonly [field: "manuallyHiddenRows" | "advancedFilterState.hiddenRows", runs: string]> = [
  ["manuallyHiddenRows", "hiddenRowRuns"],
  ["advancedFilterState.hiddenRows", "advancedFilterHiddenRowRuns"],
];
const MAX_ROWS = 1_048_576;
const RUN_FIELDS = [
  ...COLUMN_MAPS.map(([, runs]) => runs), HIDDEN_COLUMN_RUNS_KEY, COLUMN_STYLE_RUNS_KEY, ...ROW_LISTS.map(([, runs]) => runs),
];

function rowListOf(sheet: Sheetish, field: string): unknown {
  if (field === "advancedFilterState.hiddenRows") {
    const state = sheet.advancedFilterState;
    return record(state) ? state.hiddenRows : undefined;
  }
  return sheet[field];
}

function setRowList(sheet: Sheetish, field: string, rows: number[]): void {
  if (field === "advancedFilterState.hiddenRows") {
    if (record(sheet.advancedFilterState)) sheet.advancedFilterState = { ...sheet.advancedFilterState, hiddenRows: rows };
    return;
  }
  if (rows.length) sheet[field] = rows;
  else delete sheet[field];
}

function lineRuns(raw: unknown, limit: number): Array<[number, number]> {
  if (!Array.isArray(raw)) return [];
  return raw.filter((item): item is [number, number] =>
    Array.isArray(item) && item.length === 2 && Number.isInteger(item[0]) && Number.isInteger(item[1]) &&
    item[0] >= 0 && item[0] < item[1] && item[1] < limit);
}

function splitRuns(lines: unknown[], limit: number): { singles: number[]; spans: Array<[number, number]> } {
  const ordered = [...new Set(lines.filter((value): value is number => Number.isInteger(value) && (value as number) >= 0 && (value as number) < limit))]
    .sort((left, right) => left - right);
  const singles: number[] = [];
  const spans: Array<[number, number]> = [];
  for (let at = 0; at < ordered.length;) {
    let last = at;
    while (last + 1 < ordered.length && ordered[last + 1] === ordered[last] + 1) last += 1;
    if (last > at) spans.push([ordered[at], ordered[last]]);
    else singles.push(ordered[at]);
    at = last + 1;
  }
  return { singles, spans };
}

function record(value: unknown): value is ColumnMap {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function mapOf(sheet: Sheetish, field: string): ColumnMap | undefined {
  if (field === "nativeGrid.columns") {
    const grid = sheet.nativeGrid;
    return record(grid) && record(grid.columns) ? grid.columns : undefined;
  }
  const value = sheet[field];
  return record(value) ? value : undefined;
}

function hiddenRuns(raw: unknown): Array<[number, number]> {
  if (!Array.isArray(raw)) return [];
  return raw.filter((item): item is [number, number] =>
    Array.isArray(item) && item.length === 2 && Number.isInteger(item[0]) && Number.isInteger(item[1]) &&
    item[0] >= 0 && item[0] < item[1] && item[1] < MAX_COLUMN_LINE_STYLES);
}

/** Whether a stored sheet keeps any column map as runs. */
export function sheetHasColumnRuns(sheet: unknown): boolean {
  return record(sheet) && RUN_FIELDS.some((field) => sheet[field] !== undefined);
}

/** A stored sheet with every column map held per column and no run fields;
 * the same object when it has none. */
export function withColumnRunsExpanded<S extends object>(sheet: S): S {
  if (!sheetHasColumnRuns(sheet)) return sheet;
  const source = sheet as unknown as Sheetish;
  const out: Sheetish = { ...(withColumnStylesExpanded(sheet) as unknown as Sheetish) };
  for (const [field, runsKey] of COLUMN_MAPS) {
    const runs = out[runsKey];
    delete out[runsKey];
    if (!record(runs)) continue;
    const expanded = expandLineStyleRuns(runs as LineStyleMap);
    if (!expanded || expanded === runs) continue;
    const own = mapOf(source, field);
    const merged: ColumnMap = { ...expanded, ...(own ?? {}) };
    if (field === "nativeGrid.columns") {
      if (record(out.nativeGrid)) out.nativeGrid = { ...out.nativeGrid, columns: merged };
    } else {
      out[field] = merged;
    }
  }
  const spans = hiddenRuns(out[HIDDEN_COLUMN_RUNS_KEY]);
  delete out[HIDDEN_COLUMN_RUNS_KEY];
  if (spans.length) {
    const own = Array.isArray(source.manuallyHiddenCols) ? source.manuallyHiddenCols : [];
    const columns = new Set<number>(own.filter((value): value is number => Number.isInteger(value)));
    for (const [first, last] of spans) for (let column = first; column <= last; column += 1) columns.add(column);
    out.manuallyHiddenCols = [...columns].sort((left, right) => left - right);
  }
  for (const [field, runsKey] of ROW_LISTS) {
    const rowSpans = lineRuns(out[runsKey], MAX_ROWS);
    delete out[runsKey];
    if (!rowSpans.length) continue;
    const own = rowListOf(source, field);
    const rows = new Set<number>(Array.isArray(own) ? own.filter((value): value is number => Number.isInteger(value)) : []);
    for (const [first, last] of rowSpans) for (let row = first; row <= last; row += 1) rows.add(row);
    setRowList(out, field, [...rows].sort((left, right) => left - right));
  }
  return out as unknown as S;
}

/** A sheet as the page stores it: in each column map, every run of two or
 * more neighbouring columns with one value in the map's run field. */
export function withColumnRunsFolded<S extends object>(sheet: S): S {
  const expanded = withColumnRunsExpanded(sheet) as unknown as Sheetish;
  const out: Sheetish = { ...(withFoldedColumnStyles(expanded) as Sheetish) };
  for (const [field, runsKey] of COLUMN_MAPS) {
    const own = mapOf(out, field);
    if (!own || Object.keys(own).length < 2) continue;
    // Validated keys only: a run key in a map is a reader's, not a column.
    const lines: LineStyleMap = {};
    for (const key in own) if (lineKeySpan(key, MAX_COLUMN_LINE_STYLES, false)) lines[key] = own[key] as LineStyleMap[string];
    const folded = foldLineStyles(lines, MAX_COLUMN_LINE_STYLES) ?? {};
    const singles: ColumnMap = {};
    const runs: ColumnMap = {};
    for (const key in folded) (key.includes(":") ? runs : singles)[key] = folded[key];
    if (!Object.keys(runs).length) continue;
    if (field === "nativeGrid.columns") {
      // The grid keeps its (perhaps empty) columns map: an editor from
      // before runs refuses a grid without one.
      out.nativeGrid = { ...(out.nativeGrid as ColumnMap), columns: singles };
    } else if (Object.keys(singles).length) {
      out[field] = singles;
    } else {
      delete out[field];
    }
    out[runsKey] = runs;
  }
  const hidden = out.manuallyHiddenCols;
  if (Array.isArray(hidden) && hidden.length >= 2) {
    const columns = [...new Set(hidden.filter((value): value is number => Number.isInteger(value) && value >= 0 && value < MAX_COLUMN_LINE_STYLES))]
      .sort((left, right) => left - right);
    const singles: number[] = [];
    const spans: Array<[number, number]> = [];
    for (let at = 0; at < columns.length;) {
      let last = at;
      while (last + 1 < columns.length && columns[last + 1] === columns[last] + 1) last += 1;
      if (last > at) spans.push([columns[at], columns[last]]);
      else singles.push(columns[at]);
      at = last + 1;
    }
    if (spans.length) {
      out[HIDDEN_COLUMN_RUNS_KEY] = spans;
      if (singles.length) out.manuallyHiddenCols = singles;
      else delete out.manuallyHiddenCols;
    }
  }
  for (const [field, runsKey] of ROW_LISTS) {
    const own = rowListOf(out, field);
    if (!Array.isArray(own) || own.length < 2) continue;
    const { singles, spans } = splitRuns(own, MAX_ROWS);
    if (!spans.length) continue;
    out[runsKey] = spans;
    setRowList(out, field, singles);
  }
  return out as unknown as S;
}
