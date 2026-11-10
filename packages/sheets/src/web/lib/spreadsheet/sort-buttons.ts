/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Sort (the Sort dialog's OK, top to bottom) and the quick sorts'
 * shared row rewrite as pure sheet transactions. Moved out of
 * SpreadsheetWindow so the ribbon and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps its dialog authority,
 * Sheet View sessions, macro recording, the background planner for large
 * ranges and the notices.
 */

import { isFormulaError, type SheetEvaluator, type SpillOverlay } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import type { CellMap, SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { recordCellMapRowPermutation } from "@mrnote/core/web/lib/spreadsheet/cell-map-lineage.ts";
import { remapMergesByRowPermutation, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import type { SortFormatPlacement, SortLevel, SortOn } from "./multi-sort";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { normalizeManualHidden } from "@mrnote/core/web/lib/spreadsheet/manual-hidden.ts";
import { hiddenRowsFromGroups, type OutlineGroup } from "@mrnote/core/web/lib/spreadsheet/outline-groups.ts";
import { relocateSortedCell } from "./sort";
import { chartsAfterRowSort, objectsAfterRowSort, type SortObjectsGrid } from "./sort-objects";
import type { DrawingObjectLike } from "./object-buttons";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** The sort a sheet records (Excel's worksheet sortState). */
export interface SpreadsheetSortState {
  ref: string;
  conditions: Array<{ ref: string; descending?: boolean }>;
}

/** The parts of a sheet a row sort reads and writes. */
export interface SortButtonSheet extends SortObjectsGrid {
  rowCount: number;
  columnCount: number;
  cells: CellMap;
  merges?: MergeSpan[];
  sortState?: SpreadsheetSortState;
  /** Drawings: one inside a single sorted row travels with it (Excel). */
  objects?: DrawingObjectLike[];
  /** The sheet's own charts, by their cell anchors (lib/spreadsheet/chart-anchor). */
  charts?: Array<{ id: string; drawingAnchor?: unknown }>;
  /** Tables: a sort that moves part of one is refused (rowSortCutsTable). */
  dataTables?: ReadonlyArray<{
    range: { r1: number; c1: number; r2: number; c2: number };
    headerRow?: boolean;
    totalsRow?: boolean;
  }>;
}

/**
 * Whether sorting rows `r1..r2` across columns `c1..c2` would move part of a
 * table. Excel sorts cells of a table only inside its body (measured in Excel
 * for Mac 16: a range sort reaching into a table's header row or past its
 * edges fails with nothing moved, one inside the body sorts); the table's own
 * sort is its header menu's. Moving a header cell renamed nothing while the
 * table kept the old name over a number (sheet model fuzz).
 */
export function rowSortCutsTable(
  sheet: Pick<SortButtonSheet, "dataTables">,
  span: { r1: number; c1: number; r2: number; c2: number },
): boolean {
  return (sheet.dataTables ?? []).some((table) => {
    const { r1, c1, r2, c2 } = table.range;
    const bottom = r2 + (table.totalsRow ? 1 : 0);
    const meets = span.r1 <= bottom && span.r2 >= r1 && span.c1 <= c2 && span.c2 >= c1;
    if (!meets) return false;
    const bodyTop = r1 + (table.headerRow === false ? 0 : 1);
    return !(span.r1 >= bodyTop && span.r2 <= r2 && span.c1 >= c1 && span.c2 <= c2);
  });
}

function columnLetters(index: number): string {
  let value = index + 1;
  let text = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    text = String.fromCharCode(65 + remainder) + text;
    value = Math.floor((value - 1) / 26);
  }
  return text;
}

function rangeA1(range: Rect): string {
  const start = `${columnLetters(range.c1)}${range.r1 + 1}`;
  const end = `${columnLetters(range.c2)}${range.r2 + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/**
 * The merges after a row sort of rows `firstDestinationRow..` that puts
 * `sourceOrder` there, or null when one would be torn. With `columns` (the
 * sorted range's columns; only they move) a merge wholly beside them stays
 * where it is, one inside them moves with its rows, and one reaching across
 * their edge refuses the sort when any of its rows moves, as Excel's "To do
 * this, all the merged cells need to be the same size" does. Without them
 * every merge on the rows is moved with the rows (the whole-row sort). A
 * merge beside a single sorted column used to travel with that column's
 * values and cover the cell next to it, which kept its value (sheet model
 * fuzz).
 */
export function remapRowSortMergesFailClosed(
  spans: readonly MergeSpan[] | undefined,
  firstDestinationRow: number,
  sourceOrder: readonly number[],
  columns?: { c1: number; c2: number },
): MergeSpan[] | null {
  if (!spans || spans.length === 0) return [];
  let maxOrderRow = firstDestinationRow + sourceOrder.length - 1;
  for (const span of spans) {
    maxOrderRow = Math.max(maxOrderRow, span.row + span.rowSpan - 1);
  }
  const order = Array.from({ length: maxOrderRow + 1 }, (_, row) => row);
  sourceOrder.forEach((sourceRow, offset) => {
    order[firstDestinationRow + offset] = sourceRow;
  });
  if (!columns) {
    const remapped = remapMergesByRowPermutation(spans, order);
    return remapped.length === spans.length ? remapped : null;
  }
  const lastSortedRow = firstDestinationRow + sourceOrder.length - 1;
  const out: MergeSpan[] = [];
  let insideSize: string | null = null;
  let insideArea = 0;
  for (const span of spans) {
    const lastCol = span.col + span.colSpan - 1;
    const lastRow = span.row + span.rowSpan - 1;
    if (lastCol < columns.c1 || span.col > columns.c2 || lastRow < firstDestinationRow || span.row > lastSortedRow) {
      out.push(span);
      continue;
    }
    // A merge reaching out of the sorted block, across a column edge or
    // across its first or last row (a header merged down into the data,
    // R76-X13's A2:A3 over a sort from row 4): any of its rows moving would
    
    // header's label among the data).
    if (span.col < columns.c1 || lastCol > columns.c2 || span.row < firstDestinationRow || lastRow > lastSortedRow) {
      for (let row = Math.max(span.row, firstDestinationRow); row <= Math.min(lastRow, lastSortedRow); row += 1) {
        if (order[row] !== row) return null;
      }
      out.push(span);
      continue;
    }
    // "To do this, all the merged cells need to be the same size."
    const size = `${span.rowSpan}x${span.colSpan}`;
    if (insideSize !== null && insideSize !== size) return null;
    insideSize = size;
    insideArea += span.rowSpan * span.colSpan;
    const [moved] = remapMergesByRowPermutation([span], order);
    if (!moved) return null;
    out.push(moved);
  }
  // A cell of its own beside them is a cell of another size: a block with a
  // merged title row and plain rows under it is refused, as Excel refuses it
  // (measured: evidence/xprobe-misc4, A3:D3 merged in A3:D7; xlsx-rounds
  // R98-X07 UI C had moved the title among the data).
  if (insideSize !== null && insideArea !== (lastSortedRow - firstDestinationRow + 1) * (columns.c2 - columns.c1 + 1)) {
    return null;
  }
  return out;
}

/**
 * A blank in the first row of a one-cell sort's current region, between its first
 * and last columns. Excel's header guess takes no header then, whatever else the
 * row holds, bold or text over numbers; a blank in the first or last column does
 * not count (measured on 56 sorts: evidence/xprobe-headblank and xprobe-headblank2,
 * 11 blank patterns over five columns under four key columns, and narrower regions).
 */
export function regionFirstRowHasInnerBlank(
  isBlank: (column: number) => boolean,
  region: { c1: number; c2: number },
): boolean {
  for (let column = region.c1 + 1; column < region.c2; column += 1) {
    if (isBlank(column)) return true;
  }
  return false;
}

/**
 * A merge in the first row of a one-cell sort's current region. Excel's header
 * guess refuses such a region whatever is under it, where a stated header sorts
 * the rows below (measured: evidence/xprobe-sortmerge, A1:B1 merged over data,
 * and R121-X30's "Cities Data"): the row is no header to step past, and the sort
 * then meets the merge and refuses.
 */
export function regionFirstRowHoldsMerge(
  merges: readonly MergeSpan[] | undefined,
  region: { r1: number; c1: number; c2: number },
): boolean {
  return (merges ?? []).some(
    (span) =>
      span.row <= region.r1 &&
      span.row + span.rowSpan > region.r1 &&
      span.col <= region.c2 &&
      span.col + span.colSpan > region.c1 &&
      span.rowSpan * span.colSpan > 1,
  );
}

export function valueSortState(
  rowRange: { r1: number; r2: number },
  columnCount: number,
  conditions: ReadonlyArray<{ column: number; direction: "asc" | "desc" }>,
  // The sorted block's first column: a region need not start at column A.
  firstColumn = 0,
): SpreadsheetSortState {
  return {
    ref: rangeA1({
      r1: rowRange.r1,
      c1: firstColumn,
      r2: rowRange.r2,
      c2: columnCount - 1,
    }),
    conditions: conditions.map((condition) => ({
      ref: rangeA1({
        r1: rowRange.r1,
        c1: condition.column,
        r2: rowRange.r2,
        c2: condition.column,
      }),
      ...(condition.direction === "desc" ? { descending: true } : {}),
    })),
  };
}

/**
 * Numeric authority for AutoFilter comparisons. Display text is intentionally
 * not parsed here: `$1,200`, localized separators, and formatted formula
 * results must rank by their actual cell value rather than by UI text.
 */
export function actualFilterNumberForCell(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
): number | null {
  const key = `${row}:${col}`;
  if (cell?.formula) {
    if (overlay.blocked.has(key)) return null;
    const computed = evaluator.getComputed(row, col);
    return typeof computed === "number" && Number.isFinite(computed)
      ? computed
      : null;
  }
  if (!cell?.value) {
    const spilled = overlay.spilled.get(key);
    return typeof spilled === "number" && Number.isFinite(spilled)
      ? spilled
      : null;
  }
  // Legacy pages may omit `kind` even though the stored value is the exact
  // numeric scalar. An explicit non-number kind remains outside the numeric
  // domain, while missing kind can be recovered without consulting display.
  if (cell.kind !== undefined && cell.kind !== "number") return null;
  const value = Number(cell.value);
  return Number.isFinite(value) ? value : null;
}

/**
 * Sort authority for values that are neither numbers nor text: a logical or an
 * error, read from the evaluator exactly like the numeric authority above and
 * never from display text. A typed TRUE is a logical (the evaluator coerces the
 * literal, as Excel does); a text '#text' or a typed "#N/A" stays text.
 */
export function actualSortTypeForCell(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
):
  | { valueType: "boolean"; booleanValue: boolean }
  | { valueType: "error" }
  | { valueType: "text" }
  | null {
  const key = `${row}:${col}`;
  if (cell?.formula && overlay.blocked.has(key)) return { valueType: "error" };
  const value = cell?.formula || cell?.value
    ? evaluator.getComputed(row, col)
    : overlay.spilled.get(key);
  if (typeof value === "boolean") return { valueType: "boolean", booleanValue: value };
  if (isFormulaError(value)) return { valueType: "error" };
  // The evaluator already answered "not a number" (the caller asked first):
  // what it computed is text, and the key says so, so '2 and a formula's
  // "100" sort among the words instead of being re-read as numbers.
  if (typeof value === "string" && value !== "") return { valueType: "text" };
  return null;
}

/** A stored sortState read back, or undefined when it does not describe a
 *  sort of whole value columns inside the sheet. */
export function normalizeSpreadsheetSortState(
  raw: unknown,
  rowCount: number,
  columnCount: number,
): SpreadsheetSortState | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const candidate = raw as {
    ref?: unknown;
    conditions?: unknown;
  };
  if (typeof candidate.ref !== "string" || !Array.isArray(candidate.conditions)) {
    return undefined;
  }
  const stateRange = parseA1Range(candidate.ref.replace(/\$/g, ""));
  if (
    !stateRange ||
    stateRange.r1 < 0 ||
    stateRange.c1 < 0 ||
    stateRange.r2 >= rowCount ||
    stateRange.c2 >= columnCount ||
    candidate.conditions.length < 1 ||
    candidate.conditions.length > 64
  ) {
    return undefined;
  }
  const conditions: SpreadsheetSortState["conditions"] = [];
  let conditionRows: { r1: number; r2: number } | null = null;
  for (const rawCondition of candidate.conditions) {
    if (!rawCondition || typeof rawCondition !== "object") return undefined;
    const condition = rawCondition as {
      ref?: unknown;
      descending?: unknown;
    };
    if (
      typeof condition.ref !== "string" ||
      (condition.descending !== undefined &&
        typeof condition.descending !== "boolean")
    ) {
      return undefined;
    }
    const range = parseA1Range(condition.ref.replace(/\$/g, ""));
    if (
      !range ||
      range.c1 !== range.c2 ||
      range.r1 < stateRange.r1 ||
      range.r2 > stateRange.r2 ||
      range.c1 < stateRange.c1 ||
      range.c2 > stateRange.c2
    ) {
      return undefined;
    }
    if (
      conditionRows !== null &&
      (conditionRows.r1 !== range.r1 || conditionRows.r2 !== range.r2)
    ) {
      return undefined;
    }
    conditionRows = { r1: range.r1, r2: range.r2 };
    conditions.push({
      ref: rangeA1(range),
      ...(condition.descending === true ? { descending: true } : {}),
    });
  }
  return {
    ref: rangeA1(stateRange),
    conditions,
  };
}

/** What Reapply re-sorts: the recorded sort's rows, key columns and order. */
export interface ReapplySortPlan {
  sortState: SpreadsheetSortState;
  /** The block the sortState records (its key columns sit inside it). */
  sortRange: Rect;
  dataR1: number;
  dataR2: number;
  levels: SortLevel[];
}

/**
 * Data > Reapply for the worksheet's own sort: the recorded sortState as the
 * rows to sort and the key levels (values, case-insensitive). Null when the
 * sheet records no usable sort.
 */
export function reapplySortPlan(sheet: Pick<SortButtonSheet, "rowCount" | "columnCount" | "sortState">): ReapplySortPlan | null {
  const sortState = normalizeSpreadsheetSortState(sheet.sortState, sheet.rowCount, sheet.columnCount);
  if (!sortState) return null;
  const sortRange = parseA1Range(sortState.ref.replace(/\$/g, ""));
  const conditionRanges = sortState.conditions.map((condition) => ({
    condition,
    range: parseA1Range(condition.ref.replace(/\$/g, "")),
  }));
  const firstCondition = conditionRanges[0]?.range;
  if (!sortRange || !firstCondition) return null;
  const dataR1 = firstCondition.r1;
  const dataR2 = firstCondition.r2;
  if (dataR2 <= dataR1) return null;
  const levels: SortLevel[] = [];
  for (const entry of conditionRanges) {
    const range = entry.range;
    if (
      !range ||
      range.r1 !== dataR1 ||
      range.r2 !== dataR2 ||
      range.c1 !== range.c2 ||
      range.c1 < sortRange.c1 ||
      range.c1 > sortRange.c2
    ) {
      return null;
    }
    levels.push({
      column: range.c1 - sortRange.c1,
      direction: entry.condition.descending === true ? "desc" : "asc",
      sortOn: "values",
      caseSensitive: false,
    });
  }
  return { sortState, sortRange, dataR1, dataR2, levels };
}

/** One level of the Sort dialog, as its rows hold it. */
export interface SortDialogLevel {
  /** The absolute sheet column whose values are the key. */
  column: number;
  direction: "asc" | "desc";
  sortOn?: SortOn;
  formatKey?: string;
  placement?: SortFormatPlacement;
  /** Index into the dialog's custom lists (Order > Custom List). */
  customListIndex?: number;
}

/** The Sort dialog's OK for a top-to-bottom sort. */
export interface RowSortRequest {
  /** The range the dialog sorts, header row included when it has one. */
  range: Rect;
  hasHeader: boolean;
  /** The range was the current region around one cell (the header guess and
   * a recorded macro read it); either way only the range's own cells move. */
  currentRegion: boolean;
  caseSensitive: boolean;
  /**
   * Sort Options ▸ Method (XLS-A065): "stroke" orders Han text by stroke
   * count; "letter" (or absent) keeps the locale's own order.
   */
  method?: "letter" | "stroke";
  locale: string;
  levels: readonly SortDialogLevel[];
  customLists?: readonly (readonly string[])[];
}

/** The rows a dialog sort reorders and the columns that move with them. */
export interface RowSortSpan {
  dataR1: number;
  r2: number;
  moveC1: number;
  moveC2: number;
}

/**
 * What moves with each sorted row: only the sorted range's own cells, as
 * Excel sorts a range. A title above it, a total below it or a block beside
 * it -- a selection's neighbours past a blank column included -- stays put;
 * data right beside a selection is offered by the Sort Warning before the
 * sort, never moved silently. Whole rows move only when a command asks for
 * them by name. A header row is pinned. Null when nothing is left to sort.
 */
export function rowSortSpan(
  sheet: Pick<SortButtonSheet, "columnCount">,
  request: Pick<RowSortRequest, "range" | "hasHeader">,
  wholeRows = false,
): RowSortSpan | null {
  const { r1, r2, c1, c2 } = request.range;
  const dataR1 = request.hasHeader ? r1 + 1 : r1;
  if (dataR1 > r2) return null;
  return wholeRows
    ? { dataR1, r2, moveC1: 0, moveC2: sheet.columnCount - 1 }
    : { dataR1, r2, moveC1: c1, moveC2: c2 };
}

/**
 * The dialog's levels as key offsets into the range's rows (offset 0 = the
 * range's first column), each with its direction, the shared case flag and,
 * for Order > Custom List, that list's ordered keys.
 */
export function rowSortLevels(request: RowSortRequest): SortLevel[] {
  const { c1 } = request.range;
  const customLists = request.customLists ?? [];
  return request.levels.map((lvl) => ({
    column: lvl.column - c1,
    direction: lvl.direction,
    sortOn: lvl.sortOn ?? "values",
    caseSensitive: request.caseSensitive,
    locale: request.locale,
    ...(request.method === "stroke" ? { method: "stroke" as const } : {}),
    ...((lvl.sortOn ?? "values") === "values" &&
    lvl.customListIndex !== undefined &&
    customLists[lvl.customListIndex]
      ? { orderKeys: customLists[lvl.customListIndex] }
      : {}),
    ...((lvl.sortOn ?? "values") !== "values"
      ? { formatKey: lvl.formatKey, placement: lvl.placement }
      : {}),
  }));
}

/**
 * The sortState a plain value sort records (case-insensitive, no custom
 * list, no format key), which Reapply and the export replay; any other sort
 * records none.
 */
export function rowSortStateFor(
  request: RowSortRequest,
  span: RowSortSpan,
): SpreadsheetSortState | undefined {
  return !request.caseSensitive &&
    request.method !== "stroke" &&
    request.levels.every(
      (level) =>
        (level.sortOn ?? "values") === "values" &&
        level.customListIndex === undefined,
    )
    ? valueSortState(
        { r1: span.dataR1, r2: span.r2 },
        span.moveC2 + 1,
        request.levels.map((level) => ({
          column: level.column,
          direction: level.direction,
        })),
        span.moveC1,
      )
    : undefined;
}

/**
 * Whether a cell is part of a dynamic array of more than one cell: its anchor
 * (the formula) or a cell its result spills into. The host answers from the
 * same calculation it shows.
 */
export type SortArrayCellAt = (row: number, col: number) => boolean;

/**
 * Whether moving rows `dataR1..` into `sourceOrder` across `c1..c2` would move
 * a cell of a dynamic array (a row that keeps its place moves nothing). Every
 * row writer of a sort asks it before writing (B522).
 */
export function rowMoveTouchesArray(
  arrayCellAt: SortArrayCellAt | undefined,
  dataR1: number,
  sourceOrder: readonly number[],
  c1: number,
  c2: number,
): boolean {
  if (!arrayCellAt) return false;
  for (let offset = 0; offset < sourceOrder.length; offset += 1) {
    const row = dataR1 + offset;
    if (sourceOrder[offset] === row) continue;
    for (let c = c1; c <= c2; c += 1) {
      if (arrayCellAt(row, c)) return true;
    }
  }
  return false;
}

/**
 * The dynamic arrays of a sheet as a sort sees them: every formula whose
 * result spills more than one cell claims its footprint (anchor included).
 * An anchor blocked by #SPILL! spills nothing and is an ordinary formula.
 * Footprints are read on the first question, once.
 */
export function sortArrayCellsOf(
  cells: Readonly<Record<string, { formula?: unknown; value?: unknown } | undefined>>,
  getArray: (row: number, col: number) => unknown,
  blocked?: ReadonlySet<string>,
): SortArrayCellAt {
  let footprints: Array<{ r1: number; c1: number; r2: number; c2: number }> | null = null;
  return (row, col) => {
    if (!footprints) {
      footprints = [];
      for (const [key, cell] of Object.entries(cells)) {
        const raw = cell?.formula ?? cell?.value;
        if (typeof raw !== "string" || !raw.startsWith("=") || blocked?.has(key)) continue;
        const [r, c] = key.split(":").map(Number);
        if (!Number.isInteger(r) || !Number.isInteger(c)) continue;
        let block: unknown;
        try {
          block = getArray(r, c);
        } catch {
          continue;
        }
        if (!Array.isArray(block) || block.length === 0 || !Array.isArray(block[0])) continue;
        const height = block.length;
        const width = (block[0] as unknown[]).length;
        if (height * width <= 1) continue;
        footprints.push({ r1: r, c1: c, r2: r + height - 1, c2: c + width - 1 });
      }
    }
    return footprints.some((f) => row >= f.r1 && row <= f.r2 && col >= f.c1 && col <= f.c2);
  };
}

/**
 * The recorded sort cut to the sheet's stored extent. A sort of a selection
 * that runs past the last stored row (or a whole column) moves nothing out
 * there, since those rows are blank and blanks sort last, but a sortState
 * naming them was dropped by the loader, so Reapply and the saved file lost
 * the sort on reopening (sheet model fuzz). Undefined when nothing of the
 * key columns is inside the extent.
 */
export function sortStateInsideExtent(
  sortState: SpreadsheetSortState | undefined,
  rowCount: number,
  columnCount: number,
): SpreadsheetSortState | undefined {
  if (!sortState) return sortState;
  const clamp = (ref: string) => {
    const range = parseA1Range(ref.replace(/\$/g, ""));
    if (!range) return null;
    const cut = {
      r1: range.r1,
      c1: range.c1,
      r2: Math.min(range.r2, rowCount - 1),
      c2: Math.min(range.c2, columnCount - 1),
    };
    return cut.r1 <= cut.r2 && cut.c1 <= cut.c2 ? cut : null;
  };
  const whole = clamp(sortState.ref);
  if (!whole) return undefined;
  const conditions: SpreadsheetSortState["conditions"] = [];
  for (const condition of sortState.conditions) {
    const range = clamp(condition.ref);
    if (!range) return undefined;
    conditions.push({ ...condition, ref: rangeA1(range) });
  }
  const ref = rangeA1(whole);
  return ref === sortState.ref && conditions.every((condition, index) => condition.ref === sortState.conditions[index].ref)
    ? sortState
    : { ref, conditions };
}

/**
 * The sheet with rows `dataR1..dataR1+sourceOrder.length-1` rewritten in
 * `sourceOrder` (each entry the source row that lands there) across columns
 * `c1..c2`, formulas rebased by their row move, merges remapped through the
 * same permutation, and `sortState` recorded (undefined removes it). Rows
 * already in order only record the sortState. "table" when part of a table
 * would move (rowSortCutsTable). "merge" when a merged cell
 * would be split; "array" when a row that moves holds part of a dynamic
 * array (`arrayCellAt`), which Excel refuses ("You can't change part of an
 * array"): only the anchor is stored, so moving it re-spilled the whole
 * array from its new row while the records beside the spilled values moved
 * without them (B522: 30C/20B/10A became blank A, blank B, 30C and a total
 * of 60 read 30). Nothing is written in either case; null when nothing
 * would change.
 */











export function rowsASortKeepsInPlace(
  sheet: { manuallyHiddenRows?: unknown; advancedFilterState?: unknown; rowCount?: unknown; outlineGroups?: unknown },
  filterHidden: ReadonlySet<number>,
): ReadonlySet<number> {
  const rowCount = typeof sheet.rowCount === "number" ? sheet.rowCount : 0;
  const advanced = sheet.advancedFilterState;
  const extra = [
    ...normalizeManualHidden(sheet.manuallyHiddenRows, rowCount),
    ...normalizeManualHidden(
      advanced && typeof advanced === "object" ? (advanced as { hiddenRows?: unknown }).hiddenRows : undefined,
      rowCount,
    ),
    ...(Array.isArray(sheet.outlineGroups) ? hiddenRowsFromGroups(sheet.outlineGroups as OutlineGroup[]) : []),
  ].filter((row) => !filterHidden.has(row));
  return extra.length ? new Set([...filterHidden, ...extra]) : filterHidden;
}








export function sourceOrderAmongVisibleRows(
  firstRow: number,
  lastRow: number,
  hidden: ReadonlySet<number>,
  sortVisible: (rows: number[]) => number[],
): number[] {
  const visible: number[] = [];
  for (let row = firstRow; row <= lastRow; row += 1) if (!hidden.has(row)) visible.push(row);
  const sorted = sortVisible(visible);
  let next = 0;
  const order: number[] = [];
  for (let row = firstRow; row <= lastRow; row += 1) order.push(hidden.has(row) ? row : sorted[next++]);
  return order;
}

const AREA_LINK_FIELDS = ["hyperlink", "hyperlinkTooltip", "hyperlinkDisplay", "hyperlinkInArea"] as const;




function areaLinksStayInPlace(before: CellMap, after: CellMap, r1: number, r2: number, c1: number, c2: number): void {
  const places: [string, Record<string, unknown>][] = [];
  for (const key of Object.keys(before)) {
    const cell = before[key] as (SpreadsheetCell & { hyperlinkInArea?: boolean }) | undefined;
    if (!cell?.hyperlinkInArea) continue;
    const at = key.indexOf(":");
    const row = Number(key.slice(0, at));
    const col = Number(key.slice(at + 1));
    if (row < r1 || row > r2 || col < c1 || col > c2) continue;
    const link: Record<string, unknown> = {};
    for (const field of AREA_LINK_FIELDS) {
      if (cell[field as keyof typeof cell] !== undefined) link[field] = cell[field as keyof typeof cell];
    }
    places.push([key, link]);
  }
  if (places.length === 0) return;
  for (const key of Object.keys(after)) {
    const cell = after[key] as (SpreadsheetCell & { hyperlinkInArea?: boolean }) | undefined;
    if (!cell?.hyperlinkInArea) continue;
    const at = key.indexOf(":");
    const row = Number(key.slice(0, at));
    const col = Number(key.slice(at + 1));
    if (row < r1 || row > r2 || col < c1 || col > c2) continue;
    const rest = { ...cell } as Record<string, unknown>;
    for (const field of AREA_LINK_FIELDS) delete rest[field];
    after[key] = rest as unknown as SpreadsheetCell;
  }
  // A link of one cell brought onto the place keeps its cell: the file then holds
  // both, as Excel writes them (B6 sorted to B5 inside B2:B8).
  for (const [key, link] of places) {
    const moved = after[key] as (SpreadsheetCell & { hyperlink?: unknown }) | undefined;
    if (moved?.hyperlink) continue;
    after[key] = { ...(moved ?? { value: "", kind: "text" }), ...link } as SpreadsheetCell;
  }
}

export function rowSortedSheet<S extends SortButtonSheet>(
  sheet: S,
  span: { dataR1: number; c1: number; c2: number },
  sourceOrder: readonly number[],
  sortState: SpreadsheetSortState | undefined,
  arrayCellAt?: SortArrayCellAt,
): S | "merge" | "array" | "table" | null {
  const { dataR1, c1, c2 } = span;
  sortState = sortStateInsideExtent(sortState, sheet.rowCount, sheet.columnCount);
  if (sourceOrder.every((source, offset) => source === dataR1 + offset)) {
    if (JSON.stringify(sheet.sortState) === JSON.stringify(sortState)) return null;
    const next = { ...sheet };
    if (sortState !== undefined) next.sortState = sortState;
    else delete next.sortState;
    return next;
  }
  if (rowMoveTouchesArray(arrayCellAt, dataR1, sourceOrder, c1, c2)) return "array";
  if (rowSortCutsTable(sheet, { r1: dataR1, c1, r2: dataR1 + sourceOrder.length - 1, c2 })) return "table";
  const remappedMerges = remapRowSortMergesFailClosed(sheet.merges, dataR1, sourceOrder, { c1, c2 });
  if (remappedMerges === null) return "merge";
  // Every cell outside the span as it was, in its order; then each source
  // row at its sorted position, so empty cells in a moved row stay empty.
  // Built afresh: a copy of the whole map with the span deleted from it and
  // placed again was 1.3 s of a 728,000-cell sort, half of it the copy and
  // the deletes (the keys come out in the same order either way).
  const nextCells: CellMap = {};
  const r2 = dataR1 + sourceOrder.length - 1;
  for (const key of Object.keys(sheet.cells)) {
    const at = key.indexOf(":");
    const row = Number(key.slice(0, at));
    const col = Number(key.slice(at + 1));
    if (at > 0 && row >= dataR1 && row <= r2 && col >= c1 && col <= c2 && key === `${row}:${col}`) continue;
    nextCells[key] = sheet.cells[key];
  }
  const columns: string[] = [];
  for (let c = c1; c <= c2; c += 1) columns.push(`:${c}`);
  const relocated: string[] = [];
  sourceOrder.forEach((srcRow, offset) => {
    const destRow = dataR1 + offset;
    const source = String(srcRow);
    const destination = String(destRow);
    for (const column of columns) {
      const cell = sheet.cells[source + column];
      if (cell !== undefined) {
        const moved = relocateSortedCell(cell, destRow - srcRow, 0);
        if (moved !== cell) relocated.push(destination + column);
        nextCells[destination + column] = moved;
      }
    }
  });
  areaLinksStayInPlace(sheet.cells, nextCells, dataR1, r2, c1, c2);
  // What the sort did: what is derived from the cells follows it, and a tiled
  // sheet's save sends the order instead of every cell it moved.
  recordCellMapRowPermutation(sheet.cells, nextCells, { dataR1, c1, c2, sourceOrder: [...sourceOrder], relocated });
  const next: S = { ...sheet, cells: nextCells };
  if (sortState !== undefined) next.sortState = sortState;
  else delete next.sortState;
  if (sheet.merges !== undefined) {
    if (remappedMerges.length > 0) next.merges = remappedMerges;
    else delete next.merges;
  }
  const destinations = new Map(sourceOrder.map((srcRow, offset) => [srcRow, dataR1 + offset]));
  const objects = objectsAfterRowSort(sheet.objects, sheet, { r1: dataR1, r2, c1, c2 }, (row) => destinations.get(row));
  const charts = chartsAfterRowSort(sheet.charts, sheet, { r1: dataR1, r2, c1, c2 }, (row) => destinations.get(row));
  if (charts !== sheet.charts) (next as SortButtonSheet).charts = charts;
  if (objects !== sheet.objects) (next as SortButtonSheet).objects = objects;
  return next;
}
