/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { cellMapChangesSince } from "./cell-map-lineage";








































export type LineStyle = Record<string, unknown>;
export type LineStyleMap = Record<string, LineStyle>;

/** Row formats a sheet keeps, one entry a row (Excel keeps a format on every
 * row it states). It was 10,000, the old grid's rows: formatting 20,000 rows
 * left the rows past the 10,000th unformatted, unsaid. As the row sizes
 * (MAX_GRID_ROW_SIZE_OVERRIDES), and the API's own caps with it. */
export const MAX_LINE_STYLES_PER_AXIS = 200_000;
/** Blank positions one sheet materializes for its line formats; past this
 * the rest stay derivable-only (an entry there still inherits on write). */
export const MAX_LINE_STYLE_MATERIALIZED_CELLS = 200_000;

/**
 * The blank cells {@link materializeLineStyles} wrote: each wears its line's
 * format and nothing else, so a reader that already reads the lines' formats
 * (the sheet's fonts, `useSpreadsheetFontFaces`) reads nothing new from it. A
 * cell that is written later is a new object, outside this set. Walking the
 * 250,000 a 1.6 MB workbook's column formats wrote took 0.3-0.4 s of its open.
 */
const materializedLineStyleBlanks = new WeakSet<object>();

export function isMaterializedLineStyleBlank(cell: unknown): boolean {
  return typeof cell === "object" && cell !== null && materializedLineStyleBlanks.has(cell);
}





export const MAX_LINE_STYLE_MATERIALIZED_CELLS_PER_WORKBOOK = 1_000_000;

const LINE_KEY = /^(0|[1-9]\d{0,6})$/;
const LINE_RUN_KEY = /^(0|[1-9]\d{0,6}):(0|[1-9]\d{0,6})$/;
/** Columns of a sheet (XFD): a column format may cover every one. */
export const MAX_COLUMN_LINE_STYLES = 16_384;

/** The inclusive lines a stored key names, or null (a run runs forwards). */
export function lineKeySpan(key: string, lineLimit: number, runs = true): [number, number] | null {
  if (LINE_KEY.test(key)) {
    const line = Number(key);
    return line < lineLimit ? [line, line] : null;
  }
  const run = runs ? LINE_RUN_KEY.exec(key) : null;
  if (!run) return null;
  const first = Number(run[1]);
  const last = Number(run[2]);
  return last > first && last < lineLimit ? [first, last] : null;
}

const expandedRuns = new WeakMap<object, LineStyleMap>();
/** One entry per line of a stored map whose column formats may be runs (a
 * run's lines share its style object); the same map when it has none. */
export function expandLineStyleRuns(map: LineStyleMap | undefined, lineLimit = MAX_COLUMN_LINE_STYLES): LineStyleMap | undefined {
  if (!map) return map;
  const known = expandedRuns.get(map);
  if (known) return known;
  let out: LineStyleMap | null = null;
  for (const key in map) {
    if (LINE_KEY.test(key)) continue;
    const span = lineKeySpan(key, lineLimit);
    out ??= {};
    if (!span) continue;
    for (let line = span[0]; line <= span[1]; line += 1) out[String(line)] ??= map[key];
  }
  if (!out) {
    expandedRuns.set(map, map);
    return map;
  }
  // A key of its own wins over a run that also covers its line.
  for (const key in map) if (LINE_KEY.test(key)) out[key] = map[key];
  expandedRuns.set(map, out);
  return out;
}






export function foldLineStyles(map: LineStyleMap | undefined, cap = MAX_LINE_STYLES_PER_AXIS): LineStyleMap | undefined {
  const lines = expandLineStyleRuns(map);
  if (!lines) return lines;
  const indexes = Object.keys(lines).filter((key) => LINE_KEY.test(key)).map(Number).sort((left, right) => left - right);
  const out: LineStyleMap = {};
  let count = 0;
  for (let at = 0; at < indexes.length && count < cap;) {
    const first = indexes[at];
    const style = lines[String(first)];
    let last = first;
    while (at + 1 < indexes.length && indexes[at + 1] === last + 1 && sameLineStyle(lines[String(indexes[at + 1])], style)) {
      at += 1;
      last += 1;
    }
    at += 1;
    out[first === last ? String(first) : `${first}:${last}`] = style;
    count += 1;
  }
  return count > 0 ? out : undefined;
}

/** The field a sheet stores its column format runs in. */
export const COLUMN_STYLE_RUNS_KEY = "columnStyleRuns";
type ColumnStyled = { columnStyles?: LineStyleMap; columnStyleRuns?: LineStyleMap };

/** A sheet's stored column formats as one map: its runs, then its own keys
 * (a run key among them too), which win. Undefined when it has none. */
export function storedColumnStyles(sheet: object): LineStyleMap | undefined {
  const { columnStyles, columnStyleRuns } = sheet as ColumnStyled;
  const runs = columnStyleRuns && typeof columnStyleRuns === "object" && !Array.isArray(columnStyleRuns) ? columnStyleRuns : undefined;
  if (!runs) return columnStyles;
  const own: LineStyleMap = {};
  for (const key in runs) if (LINE_RUN_KEY.test(key)) own[key] = runs[key];
  return { ...own, ...(columnStyles ?? {}) };
}

/** One past the last column carrying a format of its own (its key, or the
 * end of a run key), 0 when none does: the print's spill-tail rule takes a
 * styled column as used (lookwin p86b-lw; the server's styled_column_edge).
 * A `<col>` record that runs to the sheet's last column (XFD) is not counted:
 * lookwin p222-lw, A:C and D:XFD in one style, a note ending in D prints its
 * tail on every page (D:XFC: none); R457-X52's Intro, A:D, E and F:XFD in
 * Arial, prints F11:F13's. The import folds neighbouring records of one style
 * into one run, so of the run that ends at the sheet's last column only the
 * columns up to the last one with a width of its own count (p221-lw axfd's
 * custom-width A:C, R162-X08 Metadata's C:Z). */
export function styledColumnEdge(sheet: object): number {
  const lines = expandLineStyleRuns(storedColumnStyles(sheet)) ?? {};
  const styled = new Map<number, LineStyle>();
  for (const key in lines) {
    const style = lines[key];
    const col = Number(key);
    if (style && typeof style === "object" && Object.keys(style).length && Number.isInteger(col) && col >= 0 && col < MAX_COLUMN_LINE_STYLES) styled.set(col, style);
  }
  const { columnWidths, columnWidthRuns } = sheet as { columnWidths?: unknown; columnWidthRuns?: unknown };
  const widths = new Map<number, unknown>();
  for (const map of [columnWidthRuns, columnWidths]) {
    if (!map || typeof map !== "object" || Array.isArray(map)) continue;
    for (const key in map as Record<string, unknown>) {
      const span = lineKeySpan(key, MAX_COLUMN_LINE_STYLES);
      if (!span) continue;
      for (let col = span[0]; col <= span[1]; col += 1) widths.set(col, (map as Record<string, unknown>)[key]);
    }
  }
  // Where the run of one value that ends at the sheet's last column begins.
  const trailingRunStart = <V,>(map: Map<number, V>, same: (left: V, right: V) => boolean) => {
    const last = map.get(MAX_COLUMN_LINE_STYLES - 1);
    if (last === undefined) return MAX_COLUMN_LINE_STYLES;
    let start = MAX_COLUMN_LINE_STYLES - 1;
    while (start > 0) {
      const before = map.get(start - 1);
      if (before === undefined || !same(before, last)) break;
      start -= 1;
    }
    return start;
  };
  const widthEnd = trailingRunStart(widths, (left, right) => left === right);
  let widthEdge = 0;
  for (const col of widths.keys()) if (col < widthEnd) widthEdge = Math.max(widthEdge, col + 1);
  const runStart = trailingRunStart(styled, sameLineStyle);
  let edge = 0;
  for (const col of styled.keys()) if (col < runStart || col < widthEdge) edge = Math.max(edge, col + 1);
  return edge;
}

/** A sheet as the editor holds it: one column format entry per column, no runs. */
export function withColumnStylesExpanded<S extends object>(sheet: S): S {
  const stored = storedColumnStyles(sheet);
  const expanded = expandLineStyleRuns(stored);
  if (expanded === stored && (sheet as ColumnStyled).columnStyleRuns === undefined) return sheet;
  const out = { ...sheet } as S & ColumnStyled;
  delete out.columnStyleRuns;
  const own: LineStyleMap = {};
  for (const key in expanded) if (LINE_KEY.test(key)) own[key] = expanded[key];
  if (Object.keys(own).length) out.columnStyles = own;
  else delete out.columnStyles;
  return out;
}

/** A sheet as the page stores it: a run of neighbouring columns with one
 * format in `columnStyleRuns`, a lone column in `columnStyles`. */
export function withFoldedColumnStyles<S extends object>(sheet: S): S {
  const stored = storedColumnStyles(sheet);
  if (!stored) return sheet;
  const folded = foldLineStyles(stored, MAX_COLUMN_LINE_STYLES) ?? {};
  const own: LineStyleMap = {};
  const runs: LineStyleMap = {};
  for (const key in folded) (LINE_KEY.test(key) ? own : runs)[key] = folded[key];
  const out = { ...sheet } as S & ColumnStyled;
  if (Object.keys(own).length) out.columnStyles = own;
  else delete out.columnStyles;
  if (Object.keys(runs).length) out.columnStyleRuns = runs;
  else delete out.columnStyleRuns;
  return out;
}

export interface LineStyledCell {
  value?: unknown;
  kind?: unknown;
  style?: unknown;
  [key: string]: unknown;
}

export interface LineStyledSheet {
  cells: Record<string, LineStyledCell>;
  rowCount: number;
  columnCount: number;
  rowStyles?: LineStyleMap;
  columnStyles?: LineStyleMap;
}

/**
 * Validate a persisted map: string line keys inside the grid's axis limit,
 * each value a style the caller's normalizer accepts and that is not empty.
 * Absent (undefined) when nothing survives, so a sheet without line formats
 * serializes byte-for-byte as before.
 */
export function normalizeLineStyles(
  raw: unknown,
  lineLimit: number,
  normalizeStyle: (style: unknown) => LineStyle | null | undefined,
  


  keepEmpty = false,
): LineStyleMap | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const out: LineStyleMap = {};
  let count = 0;
  // Runs (`"8:16383"`, column formats only) are held one entry per line,
  // sharing the run's normalized style; a line's own key wins over a run.
  const runs = !keepEmpty;
  const spans = Object.keys(raw)
    .map((key) => [key, lineKeySpan(key, lineLimit, runs)] as const)
    .filter((entry): entry is readonly [string, [number, number]] => entry[1] !== null)
    .sort((left, right) => left[1][0] - right[1][0] || (left[1][1] - left[1][0]) - (right[1][1] - right[1][0]));
  // Every column may carry a format (a sheet's runs arrive here one key per
  // column, lib/spreadsheet/column-runs): cut at 10,000, R67-X05's formats
  
  const cap = runs ? MAX_COLUMN_LINE_STYLES : MAX_LINE_STYLES_PER_AXIS;
  for (const [key, [first, last]] of spans) {
    if (count >= cap) break;
    const rawStyle = (raw as Record<string, unknown>)[key];
    if (first !== last) {
      const style = normalizeStyle(rawStyle);
      if (!style || Object.keys(style).length === 0) continue;
      for (let line = first; line <= last; line += 1) {
        if ((raw as Record<string, unknown>)[String(line)] === undefined) out[String(line)] = style;
      }
      count += 1;
      continue;
    }
    const index = first;
    if (
      keepEmpty &&
      rawStyle !== null &&
      typeof rawStyle === "object" &&
      !Array.isArray(rawStyle) &&
      Object.keys(rawStyle).length === 0
    ) {
      out[String(index)] = {};
      count += 1;
      continue;
    }
    const style = normalizeStyle(rawStyle);
    if (!style || Object.keys(style).length === 0) continue;
    out[String(index)] = style;
    count += 1;
  }
  if (count === 0) return undefined;
  // Every key is a line's own: expandLineStyleRuns need not test them all
  // for a run again.
  expandedRuns.set(out, out);
  return out;
}

/** The line format a blank position reads: its row's, else its column's. */
export function lineStyleAt(
  rowStyles: LineStyleMap | undefined,
  columnStyles: LineStyleMap | undefined,
  row: number,
  col: number,
): LineStyle | undefined {
  return rowStyles?.[String(row)] ?? columnStyles?.[String(col)] ?? (columnStyles ? expandLineStyleRuns(columnStyles)?.[String(col)] : undefined);
}

function cloneStyle(style: LineStyle): LineStyle {
  return JSON.parse(JSON.stringify(style)) as LineStyle;
}

/**
 * Write the line formats onto the blank positions of the grid as ordinary
 * `{ value: "", kind: "text", style }` cells. With `from` (the grid's previous
 * extent) only the NEW band — rows at or past `from.rowCount`, columns at or
 * past `from.columnCount` — is materialized; the old band already was. A row
 * format takes the whole row, a column format the rows without one. Returns
 * the same sheet object when nothing is added.
 */
export function materializeLineStyles<S extends LineStyledSheet>(
  sheet: S,
  from?: { rowCount: number; columnCount: number },
  budget = MAX_LINE_STYLE_MATERIALIZED_CELLS,
  /** Only the rows and columns before these (a print's occupied range). */
  within?: { rows: number; cols: number },
): S {
  const rowStyles = sheet.rowStyles ?? {};
  const columnStyles = expandLineStyleRuns(storedColumnStyles(sheet)) ?? {};
  const rowKeys = Object.keys(rowStyles);
  const colKeys = Object.keys(columnStyles);
  if (rowKeys.length === 0 && colKeys.length === 0) return sheet;
  const fromRows = from?.rowCount ?? 0;
  const fromCols = from?.columnCount ?? 0;
  let cells: Record<string, LineStyledCell> | null = null;
  let made = 0;
  // One style object per line format, as cells of one format share theirs
  // (lib/spreadsheet/cell-style-table): a clone per blank cost 270 ms of an
  // open of 100,000 cells.
  const blanks = new Map<LineStyle, { style: LineStyle; locked: boolean }>();
  const put = (row: number, col: number, style: LineStyle): boolean => {
    const key = `${row}:${col}`;
    if (sheet.cells[key] !== undefined) return true;
    if (cells && cells[key] !== undefined) return true;
    if (made >= budget) return false;
    cells ??= { ...sheet.cells };
    // A line's default protection (an unlocked column of a template) is the
    // cell's own lock, not a style key (B501).
    let blank = blanks.get(style);
    if (!blank) {
      const { locked, ...visual } = style;
      blank = { style: cloneStyle(visual), locked: locked === false };
      blanks.set(style, blank);
    }
    const written: LineStyledCell = {
      value: "",
      kind: "text",
      style: blank.style,
      ...(blank.locked ? { locked: false } : {}),
    };
    materializedLineStyleBlanks.add(written);
    cells[key] = written;
    made += 1;
    return true;
  };
  rows: for (const rowKey of rowKeys) {
    const row = Number(rowKey);
    if (!Number.isInteger(row) || row < 0 || row >= sheet.rowCount) continue;
    // A row whose format is the default shows its blanks unformatted: no
    // cell to write, and the column formats stay off it (below).
    if (Object.keys(rowStyles[rowKey]).length === 0) continue;
    if (within && row >= within.rows) continue;
    for (let col = 0; col < Math.min(sheet.columnCount, within?.cols ?? Infinity); col += 1) {
      if (row < fromRows && col < fromCols) continue;
      if (!put(row, col, rowStyles[rowKey])) break rows;
    }
  }
  cols: for (const colKey of colKeys) {
    const col = Number(colKey);
    if (!Number.isInteger(col) || col < 0 || col >= sheet.columnCount) continue;
    if (within && col >= within.cols) continue;
    for (let row = 0; row < Math.min(sheet.rowCount, within?.rows ?? Infinity); row += 1) {
      if (rowStyles[String(row)] !== undefined) continue;
      if (row < fromRows && col < fromCols) continue;
      if (!put(row, col, columnStyles[colKey])) break cols;
    }
  }
  return cells ? { ...sheet, cells } : sheet;
}










export const LINE_BLANKS_WRITTEN_UP_TO = 50_000;

let lineBlankBudgetAsked: number | null | undefined;
function askedLineBlankBudget(): number | null {
  if (lineBlankBudgetAsked !== undefined) return lineBlankBudgetAsked;
  lineBlankBudgetAsked = null;
  if (typeof window !== "undefined") {
    let asked = /[?&]lineBlankBudget=(\d+)/.exec(window.location.search)?.[1] ?? null;
    if (asked === null) {
      try {
        asked = window.localStorage.getItem("mrai.lineBlankBudget");
      } catch {
        asked = null;
      }
    }
    if (asked !== null && /^\d+$/.test(asked)) lineBlankBudgetAsked = Number(asked);
  }
  return lineBlankBudgetAsked;
}

const lineBlankOfStyle = new WeakMap<LineStyle, LineStyledCell>();
function lineBlankOf(style: LineStyle): LineStyledCell {
  let blank = lineBlankOfStyle.get(style);
  if (!blank) {
    const { locked, ...visual } = style;
    blank = { value: "", kind: "text", style: visual, ...(locked === false ? { locked: false } : {}) };
    materializedLineStyleBlanks.add(blank);
    lineBlankOfStyle.set(style, blank);
  }
  return blank;
}

/** The blank a position of the grid wears when it has no cell: what
 *  {@link materializeLineStyles} would have written there. For reading only
 *  (the object is shared by every position of that line format). */
export function lineBlankAt(sheet: object, row: number, col: number): LineStyledCell | undefined {
  const { rowStyles, columnStyles, rowCount, columnCount } = sheet as Partial<LineStyledSheet>;
  if (!rowStyles && !columnStyles) return undefined;
  if (row >= (rowCount ?? 0) || col >= (columnCount ?? 0)) return undefined;
  const ownRow = rowStyles?.[String(row)];
  // A row whose format is the default keeps the column formats off it.
  if (ownRow !== undefined && Object.keys(ownRow).length === 0) return undefined;
  const line = ownRow ?? columnStyles?.[String(col)];
  return line ? lineBlankOf(line) : undefined;
}

/** Positions {@link writeLineBlanksIn} looks at, past which it stops: a whole
 *  sheet selected on a sheet formatted in every column is 80 million. */
const LINE_BLANK_WRITE_VISITS = 2_000_000;
const lineBlanksWrittenIn = new WeakMap<object, string>();

/**
 * Write the line-format blanks of a rectangle into the sheet's cell map, in
 * place: commands read a selection's cells for their format and their lock.
 * The map says nothing new by it (a save folds these blanks back into the
 * line's format, {@link compactLineStyleCells}), so it keeps its identity and
 * nothing is rendered, undone or saved because of it.
 */
export function writeLineBlanksIn<S extends LineStyledSheet>(
  sheet: S,
  rect: { r1: number; c1: number; r2: number; c2: number },
): void {
  const { rowStyles, columnStyles, cells } = sheet;
  if (!rowStyles && !columnStyles) return;
  const r1 = Math.max(0, rect.r1), c1 = Math.max(0, rect.c1);
  const r2 = Math.min(sheet.rowCount - 1, rect.r2), c2 = Math.min(sheet.columnCount - 1, rect.c2);
  if (r2 < r1 || c2 < c1) return;
  const asked = `${r1}:${c1}:${r2}:${c2}`;
  if (lineBlanksWrittenIn.get(cells) === asked) return;
  lineBlanksWrittenIn.set(cells, asked);
  const put = (row: number, col: number, style: LineStyle) => {
    const key = `${row}:${col}`;
    if (cells[key] !== undefined) return;
    const blank = { ...lineBlankOf(style) };
    materializedLineStyleBlanks.add(blank);
    cells[key] = blank;
  };
  const lineColumns: number[] = [];
  if (columnStyles) {
    if (c2 - c1 < 256) {
      for (let col = c1; col <= c2; col += 1) if (columnStyles[String(col)]) lineColumns.push(col);
    } else {
      for (const key of Object.keys(columnStyles)) {
        const col = Number(key);
        if (Number.isInteger(col) && col >= c1 && col <= c2) lineColumns.push(col);
      }
    }
  }
  let visits = 0;
  for (let row = r1; row <= r2 && visits < LINE_BLANK_WRITE_VISITS; row += 1) {
    const ownRow = rowStyles?.[String(row)];
    if (ownRow !== undefined) {
      if (Object.keys(ownRow).length === 0) continue;
      for (let col = c1; col <= c2; col += 1) put(row, col, ownRow);
      visits += c2 - c1 + 1;
    } else {
      for (const col of lineColumns) put(row, col, columnStyles![String(col)]);
      visits += lineColumns.length;
    }
  }
}

/**
 * A commit of a sheet whose line-format blanks are not written: a cell a command
 * wrote at a position that had none, saying nothing of its look, wears the
 * line's format, as it would have kept the blank's. A command that means to
 * write no format writes over {@link writeLineBlanksIn} first (paste, fill).
 * In place: the map is the command's own, made for this commit.
 */
export function newCellsTakeLineFormats<S extends LineStyledSheet>(previous: S, next: S): void {
  if (next.cells === previous.cells || !(next.rowStyles || next.columnStyles)) return;
  if (lineStyleDemand(next) <= (askedLineBlankBudget() ?? LINE_BLANKS_WRITTEN_UP_TO)) return;
  const before = previous.cells, cells = next.cells;
  const wear = (key: string) => {
    if (before[key] !== undefined) return;
    const cell = cells[key];
    if (!cell || cell.style !== undefined) return;
    const separator = key.indexOf(":");
    const blank = separator > 0 ? lineBlankAt(next, Number(key.slice(0, separator)), Number(key.slice(separator + 1))) : undefined;
    if (blank) cells[key] = { ...cell, style: blank.style };
  };
  // The keys the edit changed, when the commit knows them: this ran at every entry.
  const changed = cellMapChangesSince(cells, before);
  if (changed) for (const key of changed) wear(key);
  else for (const key in cells) wear(key);
}

/** Whether a line format of this sheet unlocks its cells: a blank's lock is
 *  read from the cell by what walks a protected sheet, so these stay written. */
function linesCarryALock(sheet: LineStyledSheet): boolean {
  const unlocks = (lines: LineStyleMap | undefined) => {
    for (const key in lines) if (lines[key].locked === false) return true;
    return false;
  };
  return unlocks(sheet.rowStyles) || unlocks(expandLineStyleRuns(storedColumnStyles(sheet)));
}

/** How many blank positions {@link materializeLineStyles} would write at
 * most, before the cells already there are counted out. */
function lineStyleDemand(sheet: LineStyledSheet): number {
  const rowStyles = sheet.rowStyles ?? {};
  let styledRows = 0;
  let formattedRows = 0;
  for (const rowKey of Object.keys(rowStyles)) {
    const row = Number(rowKey);
    if (!Number.isInteger(row) || row < 0 || row >= sheet.rowCount) continue;
    styledRows += 1;
    if (Object.keys(rowStyles[rowKey]).length > 0) formattedRows += 1;
  }
  // The columns inside the sheet's count: a format over every column is
  // 16,384 keys, of a sheet often twenty columns wide.
  const columnFormats = expandLineStyleRuns(storedColumnStyles(sheet)) ?? {};
  let columns = 0;
  for (let col = 0; col < Math.min(sheet.columnCount, MAX_COLUMN_LINE_STYLES); col += 1) {
    if (columnFormats[String(col)] !== undefined) columns += 1;
  }
  return formattedRows * sheet.columnCount + columns * Math.max(0, sheet.rowCount - styledRows);
}

/**
 * Every sheet of a workbook with its line formats materialized
 * ({@link materializeLineStyles}), within one budget for the workbook: a
 * sheet that needs less than an even share takes what it needs, and the rest
 * is shared evenly by the others (never past the sheet budget). Past its
 * share a sheet's blanks stay derivable-only, as past the sheet budget.
 */
export function materializeWorkbookLineStyles<S extends LineStyledSheet>(
  sheets: S[],
  workbookBudget = MAX_LINE_STYLE_MATERIALIZED_CELLS_PER_WORKBOOK,
  /** Opening on this sheet: a large workbook's other sheets wait to be shown. */
  openingOn?: string,
  /** Only these sheets write their blanks (the rest are kept as stored); each
   *  keeps the budget it has among all of them. */
  writes?: (sheet: S) => boolean,
): S[] {
  const limit = openingOn === undefined ? null : askedLineBlankBudget() ?? LINE_BLANKS_WRITTEN_UP_TO;
  const demands = sheets.map((sheet) => {
    if (!(sheet.rowStyles || sheet.columnStyles || storedColumnStyles(sheet))) return 0;
    const demand = lineStyleDemand(sheet);
    // Past the limit a window writes none: they are read as they are needed.
    if (limit !== null && demand > limit && !linesCarryALock(sheet)) return 0;
    return Math.min(MAX_LINE_STYLE_MATERIALIZED_CELLS, demand);
  });
  const budgets = new Array<number>(sheets.length).fill(0);
  let left = workbookBudget;
  let sharing = demands.filter((demand) => demand > 0).length;
  for (const index of demands.map((_, i) => i).sort((a, b) => demands[a] - demands[b])) {
    if (demands[index] === 0) continue;
    budgets[index] = Math.min(demands[index], Math.floor(left / sharing));
    left -= budgets[index];
    sharing -= 1;
  }
  const total = demands.reduce((sum, demand) => sum + demand, 0);
  const defers = (sheet: S) => openingOn !== undefined && total >= LINE_STYLE_DEFER_FROM_CELLS &&
    (sheet as { id?: unknown }).id !== openingOn && sheets.some((candidate) => (candidate as { id?: unknown }).id === openingOn);
  return sheets.map((sheet, index) => {
    if (demands[index] === 0) return sheet;
    if (writes && !writes(sheet)) return sheet;
    if (defers(sheet)) {
      deferredLineStyleBlanks.set(sheet.cells, budgets[index]);
      return sheet;
    }
    return materializeLineStyles(sheet, undefined, budgets[index]);
  });
}

const VALUED_CELLS_EXTENT = new WeakMap<object, { rows: number; cols: number }>();

/**
 * The range a sheet's real cells occupy, for its print: the import's
 * occupiedExtent (`occupied`) widened to the cells holding a value or a
 * formula now; null when the import read none. A tiled workbook's page body
 * is projected from its stub, which leaves every valued cell to the tiles
 * (they own them: a tiled save, sort and recalc write them there), so its
 * extent stopped at the stub's last valued row (R13-X27: 50 rows for 51,
 * R379-X18: 5 for 140); the cells the window holds carry the tiles' too. A
 * whole read's extent already holds every valued cell, and a value entered
 * past it is in Excel's used range as well. The API's writer reads the same
 * (`_occupied_extent_with_values`).
 */
export function occupiedExtentWithValues(
  cells: Readonly<Record<string, unknown>>,
  occupied: unknown,
): { rows: number; cols: number } | null {
  const extent = occupied && typeof occupied === "object" ? occupied as { rows?: unknown; cols?: unknown } : null;
  const importedRows = extent?.rows;
  const importedCols = extent?.cols;
  if (typeof importedRows !== "number" || typeof importedCols !== "number") return null;
  const valued = valuedCellsExtent(cells);
  return { rows: Math.max(importedRows, valued.rows), cols: Math.max(importedCols, valued.cols) };
}

/** The rows and columns down and across to the last cell holding a value or a formula (0 for none). */
export function valuedCellsExtent(cells: Readonly<Record<string, unknown>>): { rows: number; cols: number } {
  let valued = VALUED_CELLS_EXTENT.get(cells);
  if (!valued) {
    let rows = 0;
    let cols = 0;
    for (const key in cells) {
      const cell = cells[key] as { value?: unknown; formula?: unknown } | undefined;
      if (!cell || typeof cell !== "object") continue;
      if ((cell.value === undefined || cell.value === null || cell.value === "")
        && !(typeof cell.formula === "string" && cell.formula !== "")) continue;
      const separator = key.indexOf(":");
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (separator <= 0 || !Number.isInteger(row) || !Number.isInteger(col)) continue;
      if (row + 1 > rows) rows = row + 1;
      if (col + 1 > cols) cols = col + 1;
    }
    valued = { rows, cols };
    VALUED_CELLS_EXTENT.set(cells, valued);
  }
  return valued;
}

/**
 * A sheet about to print with its line formats' blanks written, as the API's
 * print paints them (`_tiled_line_painted_blanks`): a window opening a large
 * workbook writes none past {@link LINE_BLANKS_WRITTEN_UP_TO} and reads them
 * as needed, so a printed row the file writes nowhere stood at the default
 * row where Windows Excel stands it at its empty positions' faces (R13-X27:
 * rows 1-6 under Calibri 11 columns of an Arial 10 workbook print 15 pt, 12.75
 * here, every page's table 2-3 pt short), and a row format's fill was not
 * drawn across its blanks (R38-X05's navy title band). Past the API's budget
 * of 200,000 blanks only the range the file's cells occupy is painted
 * (`occupiedExtent`); never over a cell. The same sheet when nothing is added.
 */
export function printedLineBlanksSheet<S extends LineStyledSheet>(sheet: S): S {
  if (!(sheet.rowStyles || sheet.columnStyles || storedColumnStyles(sheet))) return sheet;
  const demand = lineStyleDemand(sheet);
  if (demand === 0) return sheet;
  if (demand <= MAX_LINE_STYLE_MATERIALIZED_CELLS) return materializeLineStyles(sheet, undefined, MAX_LINE_STYLE_MATERIALIZED_CELLS);
  // The occupied range with the cells holding a value now, as the print's
  // used range reads it (occupiedExtentWithValues): a tiled workbook's stub
  // keeps no valued cell, so its imported extent can be none at all (R492-X03
  // taxa: 0 x 0 for the whole read's 59 x 15, its P:XFD column format past
  // the budget) and no blank of its column formats printed.
  const occupied = occupiedExtentWithValues(sheet.cells, (sheet as { occupiedExtent?: unknown }).occupiedExtent);
  const rows = occupied?.rows;
  const cols = occupied?.cols;
  if (typeof rows !== "number" || typeof cols !== "number" || !Number.isInteger(rows) || !Number.isInteger(cols) || rows <= 0 || cols <= 0) {
    return sheet;
  }
  return materializeLineStyles(sheet, undefined, MAX_LINE_STYLE_MATERIALIZED_CELLS, { rows, cols });
}

/** From this many line-format blanks a workbook writes only those of the sheet
 *  it opens on. The 1.6 MB benchmark held 190,000 of them among its 256,000
 *  cells, 157,000 on sheets nobody was looking at, and every pass of the first
 *  render walked them all. Under 50,000 the 60-chart workbook of 40 sheets still
 *  wrote all of theirs at its opening (0.22 s); a sheet shown writes its own. */
const LINE_STYLE_DEFER_FROM_CELLS = 5_000;
const deferredLineStyleBlanks = new WeakMap<object, number>();

/** The budget a sheet with these cells waits to write its line-format blanks
 *  with, once it is shown (`materializeWorkbookLineStyles`); until then they
 *  are derivable only, as past a budget. */
export function deferredLineStyleBudget(cells: object): number | undefined {
  return deferredLineStyleBlanks.get(cells);
}

/** Whether two styles say the same thing, whatever the key order (a key set
 *  to undefined says nothing). Asked of every blank a save folds: spelling
 *  both sides out was 22 s of a 51 s run on four sheets of 10,000 cells
 *  (R32-X04), so nothing is built to compare them. */
export function sameLineStyle(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (Array.isArray(left)) {
    if (!Array.isArray(right) || left.length !== right.length) return false;
    return left.every((item, index) => sameLineStyle(item ?? null, right[index] ?? null));
  }
  if (!left || !right || typeof left !== "object" || typeof right !== "object" || Array.isArray(right)) return false;
  const first = left as Record<string, unknown>;
  const second = right as Record<string, unknown>;
  let held = 0;
  for (const key in first) {
    if (first[key] === undefined) continue;
    held += 1;
    if (second[key] === undefined || !sameLineStyle(first[key], second[key])) return false;
  }
  for (const key in second) {
    if (second[key] !== undefined) held -= 1;
  }
  return held === 0;
}

const DERIVABLE_CELL_KEYS = new Set(["value", "kind", "style"]);









function derivableBlankKey(cell: LineStyledCell, field: string): boolean {
  if (DERIVABLE_CELL_KEYS.has(field)) return true;
  const own = cell as { locked?: unknown };
  return field === "display" || field === "chartValueCache" || (field === "locked" && own.locked === false);
}

/**
 * Drop the blank cells a line format fully describes — no value, no formula,
 * no comment or link or lock, and a style equal to the line's — so the page
 * stores the format once. The inverse of {@link materializeLineStyles}; a
 * blank the user formatted differently from its line stays a cell.
 */
export function compactLineStyleCells<S extends LineStyledSheet>(sheet: S): S {
  if (!sheet.rowStyles && !sheet.columnStyles) return sheet;
  const source = sheet.cells;
  const keys = Object.keys(source);
  // Each line format split once into its lock and its look, with the last
  // cell style compared with it: the blanks of one line share one style object
  // (materializeLineStyles), so a sheet's 200,000 of them compare a handful of
  // times. Splitting and comparing per blank was most of a line edit's host
  
  // changed during the walk, so a style compared once compares the same.
  const lines = new Map<LineStyle, { locked: boolean; visual: LineStyle; style: unknown; same: boolean }>();
  let dropped: Uint8Array | null = null;
  let droppedCount = 0;
  for (let index = 0; index < keys.length; index += 1) {
    const key = keys[index]!;
    const cell = source[key];
    if (!cell || typeof cell !== "object") continue;
    if (cell.value !== "" && cell.value !== undefined && cell.value !== null) continue;
    if (Object.keys(cell).some((field) => !derivableBlankKey(cell, field))) continue;
    const separator = key.indexOf(":");
    if (separator <= 0) continue;
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    const line = lineStyleAt(sheet.rowStyles, sheet.columnStyles, row, col);
    if (!line) continue;
    let held = lines.get(line);
    if (!held) {
      // A line's default protection is opened as the cell's own lock, not a style
      // key (materializeLineStyles): read as a style, the unlocked rows of a
      // template never folded back (73,000 blanks on ONS CPI's Table 37, and
      
      const { locked: lineLocked, ...lineVisual } = line as LineStyle & { locked?: unknown };
      held = { locked: lineLocked === false, visual: lineVisual, style: undefined, same: false };
      lines.set(line, held);
    }
    if ((cell.locked === false) !== held.locked) continue;
    const style = cell.style ?? NO_STYLE;
    if (style !== held.style) {
      held.style = style;
      held.same = sameLineStyle(style, held.visual);
    }
    if (!held.same) continue;
    dropped ??= new Uint8Array(keys.length);
    dropped[index] = 1;
    droppedCount += 1;
  }
  if (!dropped) return sheet;
  // The cells kept, in their order: written afresh when most go (a copy of
  // 200,000 cells to delete 199,000 of them), else copied and the few deleted.
  let cells: Record<string, LineStyledCell>;
  // (A "__proto__" key, own in a parsed map, is not assigned: the copy keeps it.)
  if (droppedCount * 2 > keys.length && !Object.prototype.hasOwnProperty.call(source, "__proto__")) {
    cells = {};
    for (let index = 0; index < keys.length; index += 1) {
      if (dropped[index] !== 1) cells[keys[index]!] = source[keys[index]!]!;
    }
  } else {
    cells = { ...source };
    for (let index = 0; index < keys.length; index += 1) {
      if (dropped[index] === 1) delete cells[keys[index]!];
    }
  }
  return { ...sheet, cells };
}

/** A blank's style when it has none: compared, never written. */
const NO_STYLE: LineStyle = Object.freeze({}) as LineStyle;

/**
 * The line format after a patch, merged the way the caller merges a cell's
 * style (the same rules for cleared keys), for lines `from`..`to` of one axis.
 * An empty result removes the entry; an empty map is dropped.
 */
export function patchLineStyles<S extends LineStyledSheet>(
  sheet: S,
  axis: "row" | "col",
  from: number,
  to: number,
  merge: (existing: LineStyle | undefined) => LineStyle | undefined,
): S {
  const field = axis === "row" ? "rowStyles" : "columnStyles";
  const next: LineStyleMap = { ...((axis === "col" ? expandLineStyleRuns(sheet[field]) : sheet[field]) ?? {}) };
  // Every column may carry a format (the page stores them as runs).
  const cap = axis === "col" ? MAX_COLUMN_LINE_STYLES : MAX_LINE_STYLES_PER_AXIS;
  let size = Object.keys(next).length;
  let changed = false;
  for (let line = from; line <= to; line += 1) {
    const key = String(line);
    const merged = merge(next[key]);
    if (merged && Object.keys(merged).length > 0) {
      if (next[key] === undefined && size >= cap) continue;
      if (next[key] === undefined) size += 1;
      next[key] = merged;
      changed = true;
    } else if (next[key] !== undefined) {
      delete next[key];
      size -= 1;
      changed = true;
    }
  }
  if (!changed) return sheet;
  const out = { ...sheet } as S & Record<string, unknown>;
  if (Object.keys(next).length > 0) out[field] = next;
  else delete out[field];
  return out;
}
