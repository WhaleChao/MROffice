/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert / delete rows and columns on a native spreadsheet sheet.
 *
 * Cells are re-keyed by their shifted coordinates, formulas in every cell are
 * rewritten so references keep pointing at the right cells (or surface #REF!
 * when their target line is deleted), and the sheet dimension is adjusted
 * within schema bounds. Pure and generic so the operations are unit-testable
 * and reusable across the spreadsheet model.
 */

import { nearestKnownCellMap } from "./cell-map-lineage";
import { recordFontDrawn } from "./cell-record-font";
import { filterResultMovedWithColumns, remappedFilterResult } from "./filter-result";
import { spreadsheetDefaultNames } from "./default-names";
import {
  adjustFormulaForStructuralChange,
  columnToLetters,
  shiftFormula,
  structuralChangeMovesThreeDReference,
  type StructuralAxis,
  type StructuralChangeScope,
} from "./formula-engine";
import {
  isCellRefString,
  isRangeRefString,
} from "./named-ranges";
import { defaultTableColumnName, invalidateDataTableTotalFormulaColumn } from "./data-table";
import { parseA1Range } from "./print-settings";
import {
  normalizeWhatIfTables,
  whatIfTableBlocks,
  whatIfTableFootprint,
  whatIfTableRef,
  type WhatIfTableDefinition,
} from "./what-if-table-definition";
import { protectRelativeNameRefs } from "./name-relative-refs";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS, MAX_GRID_ROW_SIZE_OVERRIDES } from "./grid-sizing";
import { expandLineStyleRuns, MAX_COLUMN_LINE_STYLES, MAX_LINE_STYLES_PER_AXIS, type LineStyleMap } from "./line-styles";
import { spreadsheetObjectOffsetsLayoutRelative } from "./layout-relative-anchor";
import {
  anchorBoxTurned,
  drawingBoxAxisAfterLineEdit,
  sheetDrawingBoxAxes,
  turnDrawingAnchorBox,
  type DrawingBoxSheet,
} from "./drawing-rotation-box";
import {
  splitAfterStructuralEdit,
  type SplitStructureSheet,
} from "./split-pane-structure";
import {
  BLANK_CELL_FORMATS_HINT,
  shiftBlankCellFormatsHint,
  type BlankCellFormatsHintCarrier,
} from "./blank-cell-formats";

// The worksheet address space (Excel's 1,048,576 x 16,384): an insert may grow
// a sheet up to here and never past it.
const MAX_ROWS = MAX_GRID_ROWS;
const MAX_COLS = MAX_GRID_COLUMNS;

/**
 * Safety budget for one insert: an N-line insert is N single-line remap passes
 * over the cells and every overlay, so the line count is bounded independently
 * of the address space. 10,000 is the most any insert could run before the
 * address space grew past 10,000 rows, so no sheet pays more than it used to.
 */
export const MAX_STRUCTURAL_INSERT_LINES = 10_000;

/** Why an insert of lines cannot run, for the caller to tell the user. */
export type StructuralInsertRefusal =
  /** Cells that hold data would be pushed past the last row/column. */
  | "sheet-end"
  /** More lines than {@link MAX_STRUCTURAL_INSERT_LINES} in one insert. */
  | "line-budget";






export function cellHoldsSheetEndContent(cell: StructuralCell): boolean {
  const held = cell as StructuralCell & {
    style?: Record<string, unknown>;
    comment?: unknown;
    locked?: boolean;
    formulaHidden?: boolean;
    runs?: unknown[];
  };
  if ((held.value ?? "") !== "" || held.formula) return true;
  if (held.style && Object.keys(held.style).length > 0) return true;
  if (held.comment || held.locked === false || held.formulaHidden === true) return true;
  return Array.isArray(held.runs) && held.runs.length > 0;
}

/** The last line index on `axis` that any stored cell occupies, or -1. */
function lastStoredLine(
  cells: Record<string, StructuralCell>,
  axis: StructuralAxis,
  held: (cell: StructuralCell) => boolean = () => true,
): number {
  let last = -1;
  for (const [key, cell] of Object.entries(cells)) {
    const separator = key.indexOf(":");
    if (separator <= 0) continue;
    const line = Number(
      axis === "row" ? key.slice(0, separator) : key.slice(separator + 1),
    );
    if (Number.isInteger(line) && line > last && cell && held(cell)) last = line;
  }
  return last;
}








function lastLineHeldAtSheetEnd(sheet: StructuralSheet, axis: StructuralAxis): number {
  let last = lastStoredLine(sheet.cells, axis, cellHoldsSheetEndContent);
  for (const merge of sheet.merges ?? []) {
    const far = axis === "row" ? merge.row + merge.rowSpan - 1 : merge.col + merge.colSpan - 1;
    if (Number.isInteger(far) && far > last) last = far;
  }
  last = Math.max(last, lastAnchoredLine(sheet.objects, axis));
  for (const table of sheet.dataTables ?? []) {
    const far = axis === "row" ? table.range?.r2 : table.range?.c2;
    if (typeof far === "number" && Number.isInteger(far) && far > last) last = far;
  }
  for (const chart of sheet.charts ?? []) {
    const line = chart.drawingAnchor
      ? lastAnchoredLine([{ anchor: chart.drawingAnchor }], axis)
      : axis === "row" ? chart.anchor?.row : chart.anchor?.col;
    if (typeof line === "number" && Number.isInteger(line) && line > last) last = line;
  }
  return last;
}

/**
 * Whether inserting `count` lines on `axis` is refused, and why; null when it
 * can run. Excel refuses an insert that would shift nonblank cells off the
 * worksheet ("To prevent possible loss of data, Excel cannot shift nonblank
 * cells off of the worksheet") and otherwise lets the blank tail fall off, so
 * a sheet already at the last row can still take an insert while its data
 * fits. The stored-cell scan runs only when the extent would pass the end.
 */
export function structuralInsertRefusal(
  sheet: StructuralSheet,
  axis: StructuralAxis,
  count: number,
): StructuralInsertRefusal | null {
  const n = Math.trunc(count);
  if (n < 1) return null;
  if (n > MAX_STRUCTURAL_INSERT_LINES) return "line-budget";
  const current = axis === "row" ? sheet.rowCount : sheet.columnCount;
  const ceiling = axis === "row" ? MAX_ROWS : MAX_COLS;
  if (current + n <= ceiling) return null;
  return lastLineHeldAtSheetEnd(sheet, axis) + n >= ceiling ? "sheet-end" : null;
}

/** One legacy (Ctrl+Shift+Enter) array formula's range, 0-based inclusive. */
export interface ArrayFormulaBlock {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /** The "row:col" key of the anchor cell that carries `arrayRef`. */
  anchor: string;
}

/**
 * The legacy array formula ranges a sheet holds: each anchor's `arrayRef`
 * (the file's `<f t="array" ref>`), the range Excel fills from that one
 * formula. A dynamic array (`spillRef`) is not one: Excel lets its anchor be
 * edited alone and its spill is never typed into.
 */
const arrayFormulaBlocksOfMap = new WeakMap<object, ArrayFormulaBlock[]>();

export function arrayFormulaBlocks(
  sheet: Pick<StructuralSheet, "cells">,
  /** The cells that can be an array formula's anchor (a formula's), when the
   * caller's cell index names them: a sorted 730,000-cell sheet's walk of
   * every cell was 0.2 s of each edit's render. */
  formulaKeys?: readonly string[],
): ArrayFormulaBlock[] {
  const blocks: ArrayFormulaBlock[] = [];
  // Read in place: a 50k-row sheet holds half a million cells, and this runs
  // on every render of the ribbon (an entries() copy of them cost 4 ms a
  // render, three renders a wheel tick while a large workbook scrolls).
  const cells = (sheet.cells ?? {}) as Record<string, unknown>;
  const visit = (key: string): void => {
    if (!Object.prototype.hasOwnProperty.call(cells, key)) return;
    const cell = cells[key];
    if (!cell) return;
    const ref = (cell as { arrayRef?: unknown }).arrayRef;
    if (typeof ref !== "string" || !ref) return;
    const area = parseA1Range(ref.replace(/\$/g, ""));
    if (!area) return;
    blocks.push({ ...area, anchor: key });
  };
  if (formulaKeys) {
    for (const key of formulaKeys) visit(key);
    return blocks;
  }
  // The whole walk is kept by map, and followed from the map before at the keys
  // that changed: Delete asked it at every key press (20 ms on 100,000 cells).
  const kept = arrayFormulaBlocksOfMap.get(cells);
  if (kept) return kept;
  const near = nearestKnownCellMap(cells as object, (map) => arrayFormulaBlocksOfMap.get(map));
  if (near) {
    const changed = new Set(near.keys);
    for (const block of near.value) if (!changed.has(block.anchor)) blocks.push(block);
    for (const key of near.keys) visit(key);
  } else {
    for (const key in cells) visit(key);
  }
  arrayFormulaBlocksOfMap.set(cells, blocks);
  return blocks;
}

/** The array formula range `row`/`col` lies in, or null. */
export function arrayFormulaBlockAt(
  sheet: Pick<StructuralSheet, "cells">,
  row: number,
  col: number,
): ArrayFormulaBlock | null {
  return (
    arrayFormulaBlocks(sheet).find(
      (block) => row >= block.r1 && row <= block.r2 && col >= block.c1 && col <= block.c2,
    ) ?? null
  );
}

/** Excel's "You can't change part of an array": the edit touches an array
 *  formula's range without covering the whole of it. */
export type StructuralArrayRefusal = "array";

/**
 * Whether inserting (`op` "insert") `count` lines at `at`, or deleting
 * (`op` "delete") the lines `[at, at + count - 1]`, would change PART of a
 * legacy array formula's range. Excel refuses both ("You can't change part
 * of an array"; Microsoft Support, "Guidelines and examples of array
 * formulas": "You can move or delete an entire array formula, but you can't
 * move or delete part of it"): a line inserted inside the range would split
 * it, and deleted lines that cut the range without removing it whole would
 * shrink it. Lines inserted above or below, and a delete that removes every
 * line of the range, are fine.
 */
export function structuralLineArrayRefusal(
  sheet: Pick<StructuralSheet, "cells">,
  axis: StructuralAxis,
  at: number,
  count: number,
  op: "insert" | "delete",
): StructuralArrayRefusal | null {
  const n = Math.trunc(count);
  if (n < 1 || !Number.isInteger(at) || at < 0) return null;
  for (const block of arrayFormulaBlocks(sheet)) {
    const start = axis === "row" ? block.r1 : block.c1;
    const end = axis === "row" ? block.r2 : block.c2;
    if (op === "insert") {
      if (at > start && at <= end) return "array";
      continue;
    }
    const last = at + n - 1;
    const intersects = at <= end && last >= start;
    const covers = at <= start && last >= end;
    if (intersects && !covers) return "array";
  }
  return null;
}

/**
 * Whether changing or clearing the cells of `rect` (0-based inclusive) would
 * change part of a legacy array formula: the rectangle meets a range without
 * covering it whole. A rectangle covering the whole range (and more) may be
 * cleared, as Excel clears an array when its entire range is selected.
 */
export function arrayPartEditRefused(
  sheet: Pick<StructuralSheet, "cells">,
  rect: { r1: number; c1: number; r2: number; c2: number },
): boolean {
  return arrayFormulaBlocks(sheet).some((block) => {
    const intersects =
      rect.r1 <= block.r2 && rect.r2 >= block.r1 && rect.c1 <= block.c2 && rect.c2 >= block.c1;
    const covers =
      rect.r1 <= block.r1 && rect.r2 >= block.r2 && rect.c1 <= block.c1 && rect.c2 >= block.c2;
    return intersects && !covers;
  });
}

/**
 * Whether typing into, clearing or editing the cell `row`/`col` would change
 * part of a legacy array formula. Excel lets the array be re-entered only
 * from its anchor with the whole range selected (then Ctrl+Shift+Enter
 * rewrites every cell of it); any other entry inside the range is refused
 * with "You can't change part of an array".
 */
export function arrayPartEntryBlockedAt(
  sheet: Pick<StructuralSheet, "cells">,
  selection: { r1: number; c1: number; r2: number; c2: number },
  row: number,
  col: number,
  /** The sheet's blocks where the caller holds them for this snapshot. */
  blocks?: readonly ArrayFormulaBlock[],
): boolean {
  const block = blocks
    ? (blocks.find((each) => row >= each.r1 && row <= each.r2 && col >= each.c1 && col <= each.c2) ?? null)
    : arrayFormulaBlockAt(sheet, row, col);
  if (!block) return false;
  const whole =
    selection.r1 === block.r1 &&
    selection.r2 === block.r2 &&
    selection.c1 === block.c1 &&
    selection.c2 === block.c2;
  return !(whole && row === block.r1 && col === block.c1);
}

/**
 * How many of `count` lines starting at `at` lie inside the persisted extent
 * (rowCount/columnCount) and so can be deleted. The grid shows blank lines
 * past the extent on demand; deleting those changes nothing, so a selection
 * that runs past the extent deletes only its in-extent part.
 */
export function deletableLineCount(
  sheet: StructuralSheet,
  axis: StructuralAxis,
  at: number,
  count: number,
): number {
  const n = Math.trunc(count);
  if (n < 1 || !Number.isInteger(at) || at < 0) return 0;
  // A line past the stored extent has no cells, but the sheet's address space
  // does not end at the extent: a formula naming A100 loses that row when row
  // 100 is deleted, wherever the content ends (live audit B152). The extent
  // itself only shrinks by the lines that were inside it (see deleteLine).
  return n;
}

export interface StructuralCell {
  value?: string;
  formula?: string;
  kind?: "text" | "number";
}

/** A "row:col"-keyed overlay (e.g. sparklines) that shifts with its anchor cell. */
export type StructuralCellKeyed = Record<string, unknown>;

/** A rectangular range bound to grid coordinates (e.g. a data-validation rule). */
export interface StructuralRanged {
  range: { r1: number; c1: number; r2: number; c2: number };
}

/**
 * A structured table (ListObject): a range PLUS an optional per-column totals
 * map keyed by the column's OFFSET from `range.c1` (0-based). The range shifts
 * like a validation rule, but the totals offsets must ALSO shift on a COLUMN op
 * — a totals entry on the deleted column drops and later offsets move — else the
 * totals row aggregates the wrong physical columns after an insert/delete.
 */
export interface StructuralDataTable extends StructuralRanged {
  name?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  totals?: Record<number, unknown>;
  totalLabels?: Record<number, unknown>;
  totalFormulas?: Record<number, unknown>;
  calculatedFormulas?: Record<number, unknown>;
  /** Calculated columns the file marks `array="1"`, by column offset: each
   *  row's cell is its own array formula, and the page holds no recipe. */
  arrayCalculatedColumns?: Record<number, unknown>;
  columns?: Array<{ id: number; name: string }>;
  columnFilters?: Record<string, unknown>;
  sortState?: {
    conditions: Array<{ columnId: number; descending?: boolean }>;
  };
}

/** A single-column-bound overlay (e.g. a slicer): its `column` shifts with cols. */
export interface StructuralColumnBound {
  column: number;
}

/**
 * An anchor+span merged block (Excel ListObject parity, see merge-cells.ts):
 * a rectangle stored as a top-left (row, col) plus a rowSpan/colSpan. NOT a
 * {r1,c1,r2,c2} rect — the real `merges` model carries this exact shape, so the
 * remap converts to inclusive edges, shifts them, then converts back.
 */
export interface StructuralMerge {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

/**
 * An axis-bound, inclusive LINE range (outline group, see outline-groups.ts):
 * a contiguous [start, end] run on one axis. Only the matching axis moves; a
 * row group is untouched by a column op and vice versa (the opposite-axis
 * analogue of a column-bound slicer, but a RANGE rather than a single line).
 */
export interface StructuralLineRange {
  axis: "row" | "col";
  start: number;
  end: number;
}

/**
 * A floating CHART (see chart.ts): a {r1,c1,r2,c2} SOURCE range PLUS an anchor
 * cell it floats near. BOTH move on a structural change — the range exactly like
 * a validation rule (reuse remapRanged) and the anchor cell exactly like a
 * sparkline's "row:col" key (reuse the coord shift). A chart whose ENTIRE source
 * range is deleted is kept with `sourceRefError` (Excel's #REF! series).
 */
export interface StructuralChart {
  range: { r1: number; c1: number; r2: number; c2: number };
  anchor: { row: number; col: number };
  sourceRefError?: true;
  /** The chart's cell anchor (lib/spreadsheet/chart-anchor): it moves as a
   * drawing object's does, and a Move-and-size chart whose every line is
   * deleted is deleted, as Excel deletes it. */
  drawingAnchor?: StructuralObjectAnchor;
}

/**
 * A cell anchor point of a floating drawing OBJECT (image / shape). The grid
 * coordinate is `row` for the row axis and EITHER `column` or `col` for the
 * column axis (the object model accepts both spellings). Kept a LOOSE structural
 * contract (all-optional, NO index signature) so the concrete object-model anchor
 * — which carries extra within-cell pixel / EMU offset fields — is assignable to
 * it; the remap reads only the line index and spreads the rest through untouched.
 */
export interface StructuralAnchorPoint {
  row?: number;
  column?: number;
  col?: number;
  /** Within-cell EMU offsets; a point on a deleted line loses its offset. */
  rowOffset?: number;
  columnOffset?: number;
  colOffset?: number;
}

/**
 * A floating drawing object (image / shape, see shapes.ts + the object layer):
 * cell-anchored from a top-left point and optionally to a bottom-right point.
 * BOTH points slide with an insert/delete on the active axis exactly like a
 * chart's anchor — UNLESS the anchor is `absolute` (pinned to the page in EMU,
 * Excel "Don't move or size with cells"), in which case it is left untouched.
 * Carried through verbatim previously = the blind spot: a shape floating over the
 * WRONG cells after a row/column insert or delete.
 */
export interface StructuralObjectAnchor {
  kind?: "oneCell" | "twoCell" | "absolute";
  from?: StructuralAnchorPoint;
  to?: StructuralAnchorPoint;
}

export interface StructuralObject {
  anchor?: StructuralObjectAnchor;
  /** Object kind; only a `chart` carries a grid-bound A1 source range. */
  type?: string;
  /** Chart objects only: the SINGLE A1 source range ("Sheet1!B2:B10"). */
  range?: string;
  /** Chart objects only: the legacy plural A1 source ranges. */
  ranges?: string[];
}

export interface StructuralSheet {
  /** Identity and tab name: what a PivotTable on ANOTHER worksheet names its
   * source sheet by (`sourceSheetId`, or an imported report's sheet name). */
  id?: string;
  name?: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, StructuralCell>;
  /** Per-cell overlays keyed "row:col" (sparklines): keys shift with the cell. */
  sparklines?: StructuralCellKeyed;
  /** Range-bound rules (data validation): r1/c1/r2/c2 shift; dropped if erased. */
  validations?: StructuralRanged[];
  /** Range-bound structured tables (ListObject): r1/c1/r2/c2 shift; dropped if erased. */
  dataTables?: StructuralDataTable[];
  /** Range-bound conditional formats: r1/c1/r2/c2 shift; dropped if erased. */
  conditionalFormats?: StructuralRanged[];
   
  ignoredErrorRanges?: StructuralRanged[];
   
  hyperlinkAreas?: StructuralRanged[];
  /** Column-bound overlays (slicers): `column` shifts; dropped if its col is deleted. */
  slicers?: StructuralColumnBound[];
  /** Watched cells (Watch Window): the point shifts on its line's axis; dropped when that line is deleted. */
  watches?: Array<{ row: number; col: number }>;
  /** Anchor+span merged blocks: edges shift; collapse-to-1x1 (or fully erased) drops. */
  merges?: StructuralMerge[];
  /** Axis-bound line ranges (outline groups): the matching axis shifts; collapse drops. */
  outlineGroups?: StructuralLineRange[];
  /** Floating charts: source range AND anchor cell shift; an erased range marks the source #REF!. */
  charts?: StructuralChart[];
  /** Floating drawing objects (images / shapes): each cell anchor's from/to point
   * shifts on the active axis (an absolute, page-pinned anchor is left untouched);
   * an anchor on the deleted line moves to the deletion point; a two-cell object
   * whose every line is deleted is dropped (Excel). */
  objects?: StructuralObject[];
  /** Per-line size overrides keyed by stringified ABSOLUTE line index (column for
   * columnWidths, row for rowHeights): only the matching axis moves; a key on the
   * deleted line drops, later keys shift. */
  columnWidths?: Record<string | number, number>;
  /** Best-fit column marks (a fitted, never user-set width), keyed like
   * columnWidths and moved with them. */
  bestFitColumns?: Record<string, true>;
  nativeGrid?: import("./native-typography").NativeGrid;
  rowHeights?: Record<string | number, number>;
  /** Per-column AutoFilter keyed by stringified column index: shifts on a column
   * op, dropped when its column is deleted (kept in sync with its slicer). */
  filters?: Record<string | number, unknown>;
  /** Filter-by-Color / Text-or-Number condition filters, BOTH keyed by column
   * index like `filters` (they drive the same hidden-rows set), so they shift on
   * a column op and drop with a deleted column — kept in sync with `filters`. */
  colorFilters?: Record<string | number, unknown>;
  conditionFilters?: Record<string | number, unknown>;
  /** The worksheet AutoFilter's A1 range ("A2:H22"): it widens, narrows and
   * shifts with the grid like Excel's _FilterDatabase; an erased range drops. */
  autoFilterRef?: string;
  /** The file's `xr:uid` of its `<autoFilter>`: it lives and goes with the ref. */
  autoFilterId?: string;
  namedSheetViews?: readonly unknown[];
  /** The sheet's Excel ``Database`` range (the package's _xlnm.Database
   * name, "A4:G237"): it moves with the grid as Excel moves every name; an
   * erased range drops (the package then keeps its name as it was). */
  databaseRef?: string;
  /** The worksheet's last value sort (Reapply, the file's sortState): its
   * range and key columns follow the grid like the AutoFilter range. */
  sortState?: { ref: string; conditions: Array<{ ref: string; descending?: boolean }> };
  /** Defined names (named ranges): an A1 `ref` string ("B2:B10") that shifts with
   * the grid (#REF! when its target is erased — the name is kept, not dropped). */
  namedRanges?: StructuralNamedRange[];
  /** A pivot recipe with an A1 `source` range string that shifts with the grid. */
  pivot?: StructuralPivot;
  /** The sheet's other PivotTables; each shifts as `pivot` does. */
  morePivots?: StructuralPivot[];
  /** Scenarios (Scenario Manager): each carries single-cell A1 `ref`s that shift
   * with the grid; a ref ON the deleted line is DROPPED from its scenario. */
  scenarios?: StructuralScenario[];
  /** What-If Data Tables: each moves with the grid as a whole, its input
   *  cells with their cells; a table the edit cuts ends (what-if-table-definition). */
  whatIfTables?: WhatIfTableDefinition[];
  /** Manually hidden row indices (Excel "Hide rows", distinct from outline
   * grouping): each absolute index shifts with a row op; an index ON the deleted
   * row drops, later indices move up. Untouched by a column op. */
  manuallyHiddenRows?: number[];
  /** Reading 998: those of manuallyHiddenRows the file writes 0 points high. */
  zeroHeightRows?: number[];
  /** Reading 944: rows past rowCount a sheet hidden by default still shows. */
  rowsShownPastEnd?: number[];
  /** Reading 945: hidden by the file's run, not by zeroHeight. */
  rowsHiddenByFileRun?: true;
  /** The rows the sheet's filters hid when last applied (filter-result). */
  filterResult?: unknown;
  /** Advanced Filter's in-place result: its list and criteria ranges follow
   * their cells like Excel's _FilterDatabase and Criteria names, its hidden
   * rows move with their rows; it ends when its header row or its whole body
   * is deleted. */
  advancedFilterState?: { ref: string; hiddenRows: number[]; criteria?: string };
  /** Print settings whose fields name grid lines, all moved in the same edit:
   * the ``printArea`` A1 range (dropped when its whole range is erased), the
   * repeated title rows/columns ("1:2" / "A:B", shrunk or dropped with their
   * lines) and the manual page breaks (0-based index of the line a page starts
   * at; a break goes with that line). Other fields pass through untouched. */
  printSettings?: {
    printArea?: string;
    repeatTitleRows?: string;
    repeatTitleCols?: string;
    manualRowBreaks?: number[];
    manualColBreaks?: number[];
  };
  /** Freeze-pane band size (freezeRows/freezeCols), a line BOUNDARY on its own
   * axis: inserting/deleting lines BEFORE it moves it so the same content stays
   * pinned. The split-divider position (splitRow/splitCol) is window state that
   * keeps its place through an edit (split-pane-structure.ts). */
  freezeRows?: number;
  freezeCols?: number;
  splitRow?: number;
  splitCol?: number;
  /** Manually hidden column indices (Excel "Hide columns"): each absolute index
   * shifts with a column op; an index ON the deleted column drops. */
  manuallyHiddenCols?: number[];
  /** Row / column formats (line-styles.ts), keyed by absolute line index like
   * rowHeights / columnWidths: a key on the deleted line drops, later keys
   * shift; an inserted line takes the format of the line before it. */
  rowStyles?: Record<string | number, unknown>;
  columnStyles?: Record<string | number, unknown>;
  /** The chain could not be kept (too long, a fork): the write-back moves
   * nothing past the page and says so ({@link withLineMoveRecorded}). */
  lineMovesUnchained?: boolean;
}

/** One whole-line edit of a sheet's chain: `delta` > 0 inserted at the
 * 0-based `at`, < 0 deleted from it; `prev` is the entry before ("" for the
 * first). Each entry is a sheet key of its own (`lineMove:<id>`): a shared
 * workbook merges sheet keys one by one, so two collaborators' entries both
 * stay, and two naming the same `prev` show the fork. */
export interface SheetLineMoveEntry {
  id: string;
  prev: string;
  axis: StructuralAxis;
  at: number;
  delta: number;
}

export const LINE_MOVE_KEY_PREFIX = "lineMove:";
/** Entries a chain keeps; past them the sheet is unchained. */
export const MAX_SHEET_LINE_MOVES = 256;

function lineMoveId(): string {
  const random = globalThis.crypto?.randomUUID?.();
  return (random ?? `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`).replace(/-/g, "").slice(0, 24);
}

/** A sheet's chain in order, or null when its entries fork (two name the
 * same one before them) or do not all follow from the first. */
export function sheetLineMoveChain(sheet: object): SheetLineMoveEntry[] | null {
  const byPrev = new Map<string, SheetLineMoveEntry>();
  let count = 0;
  for (const [key, value] of Object.entries(sheet)) {
    if (!key.startsWith(LINE_MOVE_KEY_PREFIX) || !value || typeof value !== "object") continue;
    const entry = value as Partial<SheetLineMoveEntry>;
    if (typeof entry.prev !== "string" || byPrev.has(entry.prev)) return null;
    byPrev.set(entry.prev, { ...(entry as SheetLineMoveEntry), id: key.slice(LINE_MOVE_KEY_PREFIX.length) });
    count += 1;
  }
  const chain: SheetLineMoveEntry[] = [];
  for (let entry = byPrev.get(""); entry; entry = byPrev.get(entry.id)) chain.push(entry);
  return chain.length === count ? chain : null;
}









export function withLineMoveRecorded<T extends StructuralSheet>(
  sheet: T,
  next: T,
  axis: StructuralAxis,
  at: number,
  delta: number,
): T {
  if (next === sheet || delta === 0 || next.lineMovesUnchained) return next;
  const chain = sheetLineMoveChain(sheet);
  if (chain === null || chain.length >= MAX_SHEET_LINE_MOVES) return { ...next, lineMovesUnchained: true };
  const id = lineMoveId();
  const prev = chain.length > 0 ? chain[chain.length - 1]!.id : "";
  return { ...next, [`${LINE_MOVE_KEY_PREFIX}${id}`]: { prev, axis, at, delta } };
}
























export interface ConditionalFormatFileCopy {
  after: string;
  src: string;
  srcAfter: string;
  rect: [number, number, number, number];
  at: [number, number];
  page: boolean;
  n: number;
  tiles?: [number, number];
  transpose?: true;
  fill?: [number, number, number, number];
  cut?: true;
  

  merge?: true;
  

  demote?: readonly object[];
}

/** How a Copy laid its block: copies of it, transposed, or filled over.
 * `rows`: the source row each copied row came from, for a Copy that left
 * rows out (the visible rows of a filtered table), whose runs are recorded
 * one by one. */
export interface ConditionalFormatFileCopyLayout {
  tiles?: [number, number];
  transpose?: true;
  fill?: [number, number, number, number];
  rows?: readonly number[];
  cut?: true;
  merge?: true;
  demote?: readonly object[];
}

export const CF_FILE_COPY_KEY_PREFIX = "cfCopy:";
/** Records a sheet keeps; past them a Copy pastes what the page lists only. */
export const MAX_CF_FILE_COPIES = 64;

/** `dest` with the record of a Copy from `source` (the same object when the
 * Copy stays on one sheet), or `dest` itself when a chain it would be placed
 * in has forked or the sheet holds {@link MAX_CF_FILE_COPIES} already. */
export function withConditionalFormatFileCopyRecorded<T extends object>(
  dest: T,
  source: object,
  copy: { rect: [number, number, number, number]; at: [number, number]; page: boolean } & ConditionalFormatFileCopyLayout,
): T {
  const destChain = sheetLineMoveChain(dest);
  const sourceChain = source === dest ? destChain : sheetLineMoveChain(source);
  const sourceId = (source as { id?: unknown }).id;
  if (destChain === null || sourceChain === null || (typeof sourceId !== "string" && typeof sourceId !== "number")) return dest;
  const held = Object.keys(dest).filter((key) => key.startsWith(CF_FILE_COPY_KEY_PREFIX)).length;
  if (held >= MAX_CF_FILE_COPIES) return dest;
  const rows = copy.rows;
  if (rows && rows.some((row, index) => row !== copy.rect[0] + index)) {
    // A Copy that left rows out pastes each run of the rows it kept closed
    // up under the one before, as the page lays its rules (keptRowRuns,
    // snapshotRulesForKeptRowsFormatPaste): one record a run, in order, or
    
    const runs: Array<{ from: number; to: number; offset: number }> = [];
    rows.forEach((row, index) => {
      const last = runs[runs.length - 1];
      if (last && row === last.to + 1) last.to = row;
      else runs.push({ from: row, to: row, offset: index });
    });
    if (
      copy.tiles || copy.transpose || copy.fill || copy.cut ||
      rows.length !== copy.rect[2] - copy.rect[0] + 1 ||
      held + runs.length > MAX_CF_FILE_COPIES
    ) {
      return dest;
    }
    let out = dest;
    for (const run of runs) {
      out = withConditionalFormatFileCopyRecorded(out, source === dest ? out : source, {
        rect: [run.from, copy.rect[1], run.to, copy.rect[3]],
        at: [copy.at[0] + run.offset, copy.at[1]],
        page: copy.page,
        ...(copy.merge ? { merge: true as const } : {}),
      });
    }
    return out;
  }
  const record: ConditionalFormatFileCopy = {
    after: destChain.length > 0 ? destChain[destChain.length - 1]!.id : "",
    src: String(sourceId),
    srcAfter: sourceChain.length > 0 ? sourceChain[sourceChain.length - 1]!.id : "",
    rect: copy.rect,
    at: copy.at,
    page: copy.page,
    n: held,
    ...(copy.tiles && (copy.tiles[0] !== 1 || copy.tiles[1] !== 1) ? { tiles: copy.tiles } : {}),
    ...(copy.transpose ? { transpose: true as const } : {}),
    ...(copy.fill ? { fill: copy.fill } : {}),
    ...(copy.cut ? { cut: true as const } : {}),
    ...(copy.merge ? { merge: true as const } : {}),
    ...(copy.cut && copy.demote?.length ? { demote: copy.demote } : {}),
  };
  return { ...dest, [`${CF_FILE_COPY_KEY_PREFIX}${lineMoveId()}`]: record };
}

/** A defined name bound to an A1 reference string (named-ranges.ts). */
export interface StructuralNamedRange {
  ref: string;
}

/** A pivot recipe bound to an A1 source-range string (pivot.ts). */
export interface StructuralPivot {
  source: string;
  /** The worksheet the source range is on when that is not the PivotTable's own
   * sheet: by id for an authored report, by name for an imported one. With
   * neither, the source is on the PivotTable's sheet. */
  sourceSheetId?: string;
  native?: { sourceSheetName?: string };
  /** "scenario": a Scenario PivotTable report, whose source is no range. */
  cacheSource?: string;
  /** Pivot field offsets are 0-based RELATIVE to the source range's first
   * column. On a COLUMN op they rebase exactly like a data table's offset-keyed
   * totals (the source string alone moving leaves them stale); a row/column
   * field offset whose column is deleted is dropped, and a value/filter/grouping
   * field whose source column is deleted is removed. A row op leaves them. */
  rows?: number[];
  columns?: number[];
  values?: Array<{ colIndex?: number }>;
  filters?: Array<{ colIndex?: number }>;
  grouping?: Array<{ colIndex?: number }>;
}

/** One {cell → value} override inside a scenario (scenarios.ts). */
export interface StructuralScenarioCell {
  ref: string;
  value: string;
}

/**
 * A named scenario (Scenario Manager): a set of single-cell A1 overrides. Each
 * `cells[].ref` shifts with the grid like a formula reference, BUT — unlike a
 * named range — a ref ON a deleted line is DROPPED from the scenario (the cell
 * it overrode no longer exists, so there is nothing to write there), mirroring
 * how a deleted cell is removed from `cells`.
 */
export interface StructuralScenario {
  cells: StructuralScenarioCell[];
}

/** The edited sheet's identity for reference rewriting: `First!A1` written on
 * First (or anywhere else) targets First and moves with its grid. A sheet with
 * no name keeps the legacy unqualified-only rewrite. */
function structuralScopeOf(sheet: Pick<StructuralSheet, "name">): StructuralChangeScope | undefined {
  return typeof sheet.name === "string" && sheet.name ? { sheetName: sheet.name } : undefined;
}

function rewriteFormula<C extends StructuralCell>(
  cell: C,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
  





  ownSheet = true,
): C {
  if (typeof cell.formula === "string" && cell.formula.startsWith("=")) {
    const next = adjustFormulaForStructuralChange(cell.formula, axis, at, delta, scope);
    const rewritten: C = { ...cell, formula: next, value: next };
    // A held result (keptResult: a VBA or add-in function's cached value and
    // the inputs it read) moves with its formula, as the file's cached <v>
    // does: left behind, it named the old formula and areas, and the page
    // and the export read the cell differently (xlsx-rounds R431-X08 E4/E6:
    // =countColoredCells($Q$6:$S$65,Q69) under a row inserted at 45).
    const kept = (cell as { keptResult?: unknown }).keptResult;
    if (kept !== undefined && next !== cell.formula) {
      const moved = keptResultMovedWithLines(kept, cell.formula, next, axis, at, delta, scope);
      if (moved === undefined) delete (rewritten as { keptResult?: unknown }).keptResult;
      else (rewritten as { keptResult?: unknown }).keptResult = moved;
    }
    if (!ownSheet) return rewritten;
    // An array formula's range ("M5:P10") moves with its anchor: the save
    // writes it only from the cell it starts at. A legacy CSE range
    // (arrayRef) moves whole -- a line edit cutting through one is refused
    // (structuralLineArrayRefusal), as Excel refuses it. A range the edit
    // erases leaves the cell a plain formula.
    const range = (cell as { arrayRef?: unknown }).arrayRef;
    if (typeof range === "string" && range) {
      const shifted = shiftA1RangeString(range, axis, at, delta);
      if (isErasedA1Range(shifted)) delete (rewritten as { arrayRef?: unknown }).arrayRef;
      else (rewritten as { arrayRef?: unknown }).arrayRef = shifted;
    }
    const spill = (cell as { spillRef?: unknown }).spillRef;
    if (typeof spill === "string" && spill) {
      const moved = spillRefMovedWithAnchor(spill, axis, at, delta);
      if (moved) (rewritten as { spillRef?: unknown }).spillRef = moved;
      else delete (rewritten as { spillRef?: unknown }).spillRef;
    }
    return rewritten;
  }
  return cell;
}

/** One line coordinate after a line edit: `at`.. shifted by `delta` (an
 *  insert), or null when a delete removes it. */
function lineAfterEdit(index: number, at: number, delta: number): number | null {
  if (delta > 0) return index >= at ? index + delta : index;
  if (index >= at && index < at - delta) return null;
  return index >= at - delta ? index + delta : index;
}

/**
 * A cell's `keptResult` after a line edit moved its formula from `formula` to
 * `next`: the formula text follows, and each input area on the edited sheet
 * moves (grows over lines inserted inside it, shrinks over lines deleted from
 * it) with the values it holds, as the reference it stands for does. A result
 * that named another formula is stale and stays as it was; undefined drops it
 * (an area the edit erases whole).
 */
function keptResultMovedWithLines(
  kept: unknown,
  formula: string,
  next: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): unknown {
  if (typeof kept !== "object" || kept === null || Array.isArray(kept)) return kept;
  const held = kept as { formula?: unknown; inputs?: unknown };
  if (held.formula !== formula) return kept;
  if (!Array.isArray(held.inputs)) return { ...held, formula: next };
  const inputs: unknown[] = [];
  for (const input of held.inputs) {
    const entry = input as { sheet?: unknown; area?: unknown; cells?: unknown } | null;
    const area = entry?.area;
    if (
      !entry || !scope || typeof entry.sheet !== "string" || entry.sheet.toUpperCase() !== scope.sheetName.toUpperCase() ||
      !Array.isArray(area) || area.length !== 4 || !area.every((n) => Number.isSafeInteger(n))
    ) {
      inputs.push(input);
      continue;
    }
    const [top, left, bottom, right] = area as [number, number, number, number];
    const [low, high] = axis === "row" ? [top, bottom] : [left, right];
    let first: number | null;
    let last: number | null;
    if (delta > 0) {
      first = low >= at ? low + delta : low;
      last = high >= at ? high + delta : high;
    } else {
      // the lines that remain of [low, high], renumbered
      const end = at - delta;
      const keptLow = low >= at && low < end ? end : low;
      const keptHigh = high >= at && high < end ? at - 1 : high;
      if (keptLow > keptHigh) return undefined;
      first = lineAfterEdit(keptLow, at, delta);
      last = lineAfterEdit(keptHigh, at, delta);
      if (first === null || last === null) return undefined;
    }
    const cells: Record<string, unknown> = {};
    for (const [key, value] of Object.entries((entry.cells ?? {}) as Record<string, unknown>)) {
      const match = /^(\d+):(\d+)$/.exec(key);
      if (!match) continue;
      const row = Number(match[1]);
      const col = Number(match[2]);
      const moved = lineAfterEdit(axis === "row" ? row : col, at, delta);
      if (moved === null) continue;
      cells[axis === "row" ? `${moved}:${col}` : `${row}:${moved}`] = value;
    }
    inputs.push({
      ...entry,
      area: axis === "row" ? [first, left, last, right] : [top, first, bottom, last],
      cells,
    });
  }
  return { ...held, formula: next, inputs };
}

function remapCells<C extends StructuralCell>(
  cells: Record<string, C>,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): Record<string, C> {
  const out: Record<string, C> = {};
  for (const key of Object.keys(cells)) {
    const [rawRow, rawCol] = key.split(":");
    const row = Number(rawRow);
    const col = Number(rawCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) {
      out[key] = cells[key];
      continue;
    }
    const coord = axis === "row" ? row : col;
    let nextRow = row;
    let nextCol = col;
    if (delta > 0) {
      if (coord >= at) {
        if (axis === "row") nextRow = row + 1;
        else nextCol = col + 1;
      }
    } else {
      if (coord === at) continue; // cell on the deleted line is removed
      if (coord > at) {
        if (axis === "row") nextRow = row - 1;
        else nextCol = col - 1;
      }
    }
    out[`${nextRow}:${nextCol}`] = rewriteFormula(cells[key], axis, at, delta, scope);
  }
  return out;
}

/** Shift a single coordinate by a structural change (or null if it is erased). */
function shiftCoord(coord: number, at: number, delta: number): number | null {
  if (delta > 0) return coord >= at ? coord + delta : coord;
  if (coord === at) return null; // on the deleted line
  return coord > at ? coord + delta : coord;
}

/** Shift a freeze/split BOUNDARY (the count of pinned leading lines / the split
 * position). A boundary at index ``value`` grows when lines are inserted strictly
 * before it and shrinks when a line before it is deleted; an op at/after the
 * boundary leaves it. 0 / absent means no freeze/split and is returned as-is. */
function shiftBoundary(
  value: number | undefined,
  at: number,
  delta: number,
): number | undefined {
  if (typeof value !== "number" || value <= 0) return value;
  if (delta > 0) return at < value ? value + delta : value;
  return at < value ? value - 1 : value; // delete one line before the boundary
}

/**
 * Shift an A1 range/ref STRING for a structural change, reusing the formula
 * engine's range adjuster so "B1:B5" -> "B1:B6" on a row insert at/above (etc.).
 * A leading sheet qualifier ("Sheet1!B1:B5") is peeled before the engine sees it
 * — the tokenizer rejects "!" and would return the string verbatim, leaving it
 * stale — then re-attached. When the target is erased the engine surfaces
 * "#REF!", which is PRESERVED here (a named range / pivot source keeps a broken
 * ref the way Excel does); callers that need to DROP on erase check for it.
 * With a `scope`, a string qualified with ANOTHER sheet's name is left as it
 * is (that sheet's cells did not move), and so is an unqualified one when the
 * scope is `qualifiedOnly` (it names its own, unedited sheet).
 */
function shiftA1RangeString(
  range: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): string {
  const bang = range.lastIndexOf("!");
  const prefix = bang >= 0 ? range.slice(0, bang + 1) : "";
  const body = bang >= 0 ? range.slice(bang + 1) : range;
  if (scope) {
    const quoted = prefix.slice(0, -1).replace(/^=/, "");
    const named = quoted.startsWith("'") && quoted.endsWith("'") && quoted.length >= 2
      ? quoted.slice(1, -1).replace(/''/g, "'")
      : quoted;
    if (prefix ? named.toUpperCase() !== scope.sheetName.toUpperCase() : scope.qualifiedOnly) return range;
  }
  return prefix + adjustFormulaForStructuralChange(body, axis, at, delta);
}

/**
 * True when a shifted A1 range string represents an ERASED target — the engine
 * surfaces it either as "#REF!" or, for a single-line range whose only line was
 * deleted, as a degenerate/inverted range whose end crosses before its start
 * (e.g. "B2:A10" / "B3:B2"). Callers either drop the overlay (sparkline) or
 * normalize it to "#REF!" (named range / pivot).
 */
function isErasedA1Range(rangeStr: string): boolean {
  if (rangeStr.includes("#REF!")) return true;
  const bang = rangeStr.lastIndexOf("!");
  const body = bang >= 0 ? rangeStr.slice(bang + 1) : rangeStr;
  const m = /^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/.exec(body);
  if (m) {
    const c1 = lettersToColumnIndex(m[1]);
    const r1 = Number(m[2]);
    const c2 = lettersToColumnIndex(m[3]);
    const r2 = Number(m[4]);
    if (c2 < c1 || r2 < r1) return true;
  }
  return false;
}

/**
 * Adjust the SOURCE-RANGE STRING carried inside a sparkline value for a
 * structural change. Returns the adjusted string, or null when the WHOLE source
 * range is erased (see isErasedA1Range) — a null result means the sparkline has
 * no data to draw and is dropped, mirroring how a validation/table whose entire
 * range collapses is dropped.
 */
function adjustSparklineRange(
  range: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): string | null {
  const next = shiftA1RangeString(range, axis, at, delta, scope);
  return isErasedA1Range(next) ? null : next;
}

/**
 * A dynamic array's stamped spill range (spillRef) after a whole-line edit:
 * its anchor moves, its size stays. The size of a spill is the size of the
 * formula's last result, never of the cells it once covered (measured in
 * Excel for Mac 16: =SEQUENCE(5) in A1 still spills A1:A5 after a row is
 * inserted or deleted at row 3; under Manual calculation =B1:B5*10 keeps its
 * five old results in A1:A5 although its reference became B1:B6 or B1:B4).
 * Shifted as a range, an insert inside the spill stretched it to A1:A6 and a
 * delete shrank it, and a stamp no calculation replaced (Manual calculation,
 * the server's stamping of a workbook) was saved so. The next calculation
 * stamps the size of its own result. A spill pushed past the sheet's last
 * line is the anchor alone, as a blocked spill is. Null when the edit erases
 * the anchor.
 */
function spillRefMovedWithAnchor(
  spillRef: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
): string | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(spillRef.trim());
  if (!match) {
    const shifted = shiftA1RangeString(spillRef, axis, at, delta);
    return isErasedA1Range(shifted) ? null : shifted;
  }
  const col = lettersToColumnIndex(match[1]);
  const row = Number(match[2]) - 1;
  const height = match[4] ? Number(match[4]) - row : 1;
  const width = match[3] ? lettersToColumnIndex(match[3]) - col + 1 : 1;
  const anchor = shiftA1RangeString(`${columnToLetters(col)}${row + 1}`, axis, at, delta);
  const moved = /^([A-Z]{1,3})(\d+)$/.exec(anchor);
  if (!moved) return null;
  const nextCol = lettersToColumnIndex(moved[1]);
  const nextRow = Number(moved[2]) - 1;
  if (height < 1 || width < 1 || (height === 1 && width === 1)) return anchor;
  const lastRow = nextRow + height - 1;
  const lastCol = nextCol + width - 1;
  if (lastRow >= MAX_ROWS || lastCol >= MAX_COLS) return anchor;
  return `${anchor}:${columnToLetters(lastCol)}${lastRow + 1}`;
}

/** Convert spreadsheet column letters (A, B, …, AA) to a 0-based index. */
function lettersToColumnIndex(letters: string): number {
  let n = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i += 1) {
    n = n * 26 + (upper.charCodeAt(i) - 64);
  }
  return n - 1;
}

/**
 * Remap a "row:col"-keyed overlay map (e.g. sparklines) the same way cells move:
 * an entry on the deleted line is dropped; later entries shift by `delta`. When
 * `adjustValueRange` is supplied (sparklines), the value's embedded source-range
 * STRING is also adjusted for the structural change — and an entry whose entire
 * source range is erased is dropped — so the chart never reads stale cells after
 * an insert/delete that intersects its source.
 */
function remapKeyed(
  keyed: StructuralCellKeyed | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
  adjustValueRange?: (value: unknown) => unknown | null,
): StructuralCellKeyed | undefined {
  if (!keyed) return keyed;
  const out: StructuralCellKeyed = {};
  for (const key of Object.keys(keyed)) {
    const [rawRow, rawCol] = key.split(":");
    const row = Number(rawRow);
    const col = Number(rawCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) {
      out[key] = keyed[key];
      continue;
    }
    let value: unknown = keyed[key];
    if (adjustValueRange) {
      value = adjustValueRange(value);
      if (value === null) continue; // entire source range erased -> drop
    }
    if (axis === "row") {
      const nextRow = shiftCoord(row, at, delta);
      if (nextRow === null) continue;
      out[`${nextRow}:${col}`] = value;
    } else {
      const nextCol = shiftCoord(col, at, delta);
      if (nextCol === null) continue;
      out[`${row}:${nextCol}`] = value;
    }
  }
  return out;
}

/**
 * Adjust a single sparkline value's source-range string for a structural change.
 * Returns the value with its `range` rewritten, or null when the whole source
 * range is erased (so the caller drops the sparkline). A value without a string
 * `range` is returned unchanged (the key still shifts).
 */
function adjustSparklineValue(
  value: unknown,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): unknown | null {
  if (!value || typeof value !== "object") return value;
  const v = value as Record<string, unknown>;
  if (typeof v.range !== "string") return value;
  const nextRange = adjustSparklineRange(v.range, axis, at, delta, scope);
  if (nextRange === null) return null;
  if (nextRange === v.range) return value;
  return { ...v, range: nextRange };
}

/** The formula texts a range-bound rule carries: a custom validation or
 * conditional-format `formula`, a list validation's `sourceRange`, a number
 * validation's bounds written as a reference or formula ("$T$50"). */
const RULE_FORMULA_FIELDS = ["formula", "sourceRange", "minFormula", "maxFormula"] as const;
/** A conditional format's thresholds that may be formulas: a colour scale's
 * and a data bar's `min` / `mid` / `max` and an icon set's `stops`. */
const RULE_THRESHOLD_FIELDS = ["min", "mid", "max"] as const;

/** Every formula text a rule carries, its thresholds' included. */
function ruleFormulaTexts(rule: object): unknown[] {
  const record = rule as Record<string, unknown>;
  const threshold = (value: unknown) =>
    value && typeof value === "object" ? (value as { formula?: unknown }).formula : undefined;
  return [
    ...RULE_FORMULA_FIELDS.map((field) => record[field]),
    ...RULE_THRESHOLD_FIELDS.map((field) => threshold(record[field])),
    ...(Array.isArray(record.stops) ? record.stops.map(threshold) : []),
  ];
}

/**
 * Rewrite a rule's formula texts for a structural edit. Excel keeps these
 * formulas relative to the rule's top-left cell, so when that cell's own line
 * is deleted the formula is first re-anchored on the next line, which becomes
 * the top-left one (`=A2>5` on A2:A10 stays `=A2>5` after row 2 is deleted),
 * and only then are its references moved with the grid. `anchor` is null for
 * a rule on another sheet, whose own cells do not move.
 */
function remapRuleFormulas<R extends object>(
  rule: R,
  anchor: number | null,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope: StructuralChangeScope | undefined,
): R {
  const remap = (text: string) => {
    const anchored = anchor !== null && delta < 0 && anchor === at
      ? shiftFormula(text, axis === "row" ? 1 : 0, axis === "col" ? 1 : 0)
      : text;
    return adjustFormulaForStructuralChange(anchored, axis, at, delta, scope);
  };
  // A threshold's formula moves as the rule's own formulas do (live audit
  // S-033: a number rule's bounds $T$50/$U$50 stayed on row 50 after a row
  // was inserted above, so the rule then read empty cells).
  const remapThreshold = <V>(value: V): V => {
    const formula = value && typeof value === "object" ? (value as { formula?: unknown }).formula : undefined;
    if (typeof formula !== "string" || !formula) return value;
    const moved = remap(formula);
    return moved === formula ? value : { ...value, formula: moved };
  };
  let next = rule;
  for (const field of RULE_FORMULA_FIELDS) {
    const text = (rule as Record<string, unknown>)[field];
    if (typeof text !== "string" || !text) continue;
    const moved = remap(text);
    if (moved !== text) next = { ...next, [field]: moved };
  }
  for (const field of RULE_THRESHOLD_FIELDS) {
    const value = (rule as Record<string, unknown>)[field];
    const moved = remapThreshold(value);
    if (moved !== value) next = { ...next, [field]: moved };
  }
  const stops = (rule as Record<string, unknown>).stops;
  if (Array.isArray(stops)) {
    const moved = stops.map(remapThreshold);
    if (moved.some((stop, index) => stop !== stops[index])) next = { ...next, stops: moved };
  }
  return next;
}

/**
 * Remap range-bound rules (validation): each edge shifts with the grid. A rule
 * whose ENTIRE range is on the single deleted line collapses and is dropped; a
 * range straddling the deleted line shrinks by one. Insert always widens a range
 * that starts at/after the insert point.
 */
function remapRanged<R extends StructuralRanged>(
  ranged: R[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): R[] | undefined {
  if (!ranged) return ranged;
  const out: R[] = [];
  for (const rule of ranged) {
    const { r1, c1 } = rule.range;
    // A rule may cover several areas (Excel's comma-separated sqref); each
    // shifts on its own, an erased one leaves the list, and the first area
    // stays the rule's `range`, as the API's reader requires.
    const rawAreas = (rule as { areas?: unknown }).areas;
    const areas = Array.isArray(rawAreas) && rawAreas.length > 0
      ? (rawAreas as StructuralRanged["range"][])
      : null;
    const shiftedAreas = areas
      ? areas.map((area) => shiftRangeBox(area, axis, at, delta)).filter((area): area is StructuralRanged["range"] => area !== null)
      : null;
    const range = shiftedAreas ? shiftedAreas[0] : shiftRangeBox(rule.range, axis, at, delta);
    if (!range) continue;
    const moved = scope ? remapRuleFormulas(rule, axis === "row" ? r1 : c1, axis, at, delta, scope) : rule;
    if (shiftedAreas) {
      out.push({ ...moved, range, areas: shiftedAreas.length > 1 ? shiftedAreas : undefined } as R);
      if (shiftedAreas.length <= 1) delete (out[out.length - 1] as { areas?: unknown }).areas;
      continue;
    }
    out.push({ ...moved, range });
  }
  return out;
}

/** One {r1,c1,r2,c2} box after a line insert/delete, or null when erased. */
function shiftRangeBox(
  box: StructuralRanged["range"],
  axis: StructuralAxis,
  at: number,
  delta: number,
): StructuralRanged["range"] | null {
  const { r1, c1, r2, c2 } = box;
  let nr1 = r1;
  let nr2 = r2;
  let nc1 = c1;
  let nc2 = c2;
  if (axis === "row") {
    if (delta > 0) {
      if (r1 >= at) nr1 = r1 + delta;
      if (r2 >= at) nr2 = r2 + delta;
    } else {
      // Delete: the start moves up if past the line; the end moves up if at or
      // past it. A range fully on the deleted single line (r1==r2==at) erases.
      if (r1 === at && r2 === at) return null;
      if (r1 > at) nr1 = r1 - 1;
      if (r2 >= at && r2 < MAX_ROWS - 1) nr2 = r2 - 1;
    }
  } else if (delta > 0) {
    if (c1 >= at) nc1 = c1 + delta;
    if (c2 >= at) nc2 = c2 + delta;
  } else {
    if (c1 === at && c2 === at) return null;
    if (c1 > at) nc1 = c1 - 1;
    if (c2 >= at && c2 < MAX_COLS - 1) nc2 = c2 - 1;
  }
  // A range that reaches the sheet's last line (a whole-column rule, a merge
  // across a whole row) keeps reaching it: Excel never grows it past the
  // grid, and a start past the grid is a range the grid cannot hold. A delete
  // before that line leaves it there too (xlsx-rounds evidence/xprobe-gridend:
  // a rule on H1:XFD2 reads G1:XFD2 with column C deleted).
  if (nr1 >= MAX_ROWS || nc1 >= MAX_COLS) return null;
  return { r1: nr1, c1: nc1, r2: Math.min(nr2, MAX_ROWS - 1), c2: Math.min(nc2, MAX_COLS - 1) };
}

/**
 * Remap structured tables (ListObject): the {r1,c1,r2,c2} range shifts exactly
 * like a validation rule (a table fully on the single deleted line drops), AND —
 * on a COLUMN op — the offset-keyed `totals` map is remapped so each per-column
 * aggregate stays on its physical column. Each old offset maps to its absolute
 * column (`c1 + offset`), shifts with the grid, then back to an offset against
 * the NEW `c1`; a totals entry on the deleted column is dropped, and any offset
 * that lands outside the new range is dropped. A ROW op leaves `totals` untouched
 * (offsets are column-relative). Without the totals remap the totals row
 * aggregates the WRONG columns after a column insert/delete.
 */
function remapDataTables<T extends StructuralDataTable>(
  tables: T[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): T[] | undefined {
  if (!tables) return tables;
  const out: T[] = [];
  for (const table of tables) {
    const { r1, c1, r2, c2 } = table.range;
    let nr1 = r1;
    let nr2 = r2;
    let nc1 = c1;
    let nc2 = c2;
    let deletesPhysicalTotalsRow = false;
    if (axis === "row") {
      if (delta > 0) {
        if (r1 >= at) nr1 = r1 + delta;
        if (r2 >= at) nr2 = r2 + delta;
        else if (table.totalsRow && at === r2 + 1) nr2 = r2 + delta;
      } else {
        if (r1 === at && r2 === at) continue;
        if (table.totalsRow && at === r2 + 1) {
          deletesPhysicalTotalsRow = true;
        }
        if (r1 > at) nr1 = r1 - 1;
        if (r2 >= at) nr2 = r2 - 1;
      }
    } else if (delta > 0) {
      if (c1 >= at) nc1 = c1 + delta;
      if (c2 >= at) nc2 = c2 + delta;
    } else {
      if (c1 === at && c2 === at) continue;
      if (c1 > at) nc1 = c1 - 1;
      if (c2 >= at) nc2 = c2 - 1;
    }
    const next = { ...table, range: { r1: nr1, c1: nc1, r2: nr2, c2: nc2 } };
    if (deletesPhysicalTotalsRow) {
      delete next.totalsRow;
      delete next.totals;
      delete next.totalLabels;
      delete next.totalFormulas;
    }
    // The totals map is keyed by column offset from c1, so only a COLUMN op
    // moves it; a ROW op keeps the original offsets (carried via the spread).
    // Every caller has already checked that the map exists.
    const remapOffsetMap = (
      source: Record<number, unknown>,
    ): Record<number, unknown> => {
      const remapped: Record<number, unknown> = {};
      const width = nc2 - nc1;
      for (const rawKey of Object.keys(source)) {
        const offset = Number(rawKey);
        if (!Number.isInteger(offset)) continue;
        const absCol = c1 + offset;
        let nextAbs: number;
        if (delta > 0) {
          nextAbs = absCol >= at ? absCol + delta : absCol;
        } else {
          if (absCol === at) continue; // this column was deleted
          nextAbs = absCol > at ? absCol - 1 : absCol;
        }
        const nextOffset = nextAbs - nc1;
        if (nextOffset >= 0 && nextOffset <= width) {
          remapped[nextOffset] = source[offset];
        }
      }
      return remapped;
    };
    if (axis === "col") {
      if (table.totals) next.totals = remapOffsetMap(table.totals);
      if (table.totalLabels) {
        next.totalLabels = remapOffsetMap(table.totalLabels);
      }
      if (table.totalFormulas) {
        let remapped = remapOffsetMap(table.totalFormulas);
        if (delta < 0 && at >= c1 && at <= c2) {
          const deletedColumnName = table.columns?.[at - c1]?.name;
          if (deletedColumnName) {
            remapped = Object.fromEntries(
              Object.entries(remapped ?? {}).map(([offset, formula]) => [
                offset,
                typeof formula === "string"
                  ? invalidateDataTableTotalFormulaColumn(
                      formula,
                      table.name ?? "",
                      deletedColumnName,
                    )
                  : formula,
              ]),
            );
          }
        }
        if (Object.keys(remapped).length > 0) next.totalFormulas = remapped;
        else delete next.totalFormulas;
      }
      if (table.calculatedFormulas) {
        let remapped = remapOffsetMap(table.calculatedFormulas);
        remapped = Object.fromEntries(
          Object.entries(remapped ?? {}).map(([offset, rawFormula]) => {
            if (typeof rawFormula !== "string") return [offset, rawFormula];
            let formula = adjustFormulaForStructuralChange(
              rawFormula,
              axis,
              at,
              delta,
              scope,
            );
            if (delta < 0 && at >= c1 && at <= c2) {
              const deletedColumnName = table.columns?.[at - c1]?.name;
              if (deletedColumnName) {
                formula = invalidateDataTableTotalFormulaColumn(
                  formula,
                  table.name ?? "",
                  deletedColumnName,
                );
              }
            }
            return [offset, formula];
          }),
        );
        if (Object.keys(remapped).length > 0) {
          next.calculatedFormulas = remapped;
        } else {
          delete next.calculatedFormulas;
        }
      }
      if (table.arrayCalculatedColumns) {
        const remapped = remapOffsetMap(table.arrayCalculatedColumns);
        if (Object.keys(remapped).length > 0) next.arrayCalculatedColumns = remapped;
        else delete next.arrayCalculatedColumns;
      }
      if (table.columns) {
        const removedColumnId =
          delta < 0
            ? table.columns.find((_, offset) => c1 + offset === at)?.id
            : undefined;
        const columns = table.columns.filter(
          (_, offset) => !(delta < 0 && c1 + offset === at),
        );
        if (delta > 0 && at > c1 && at <= c2) {
          const offset = at - c1;
          // Excel names the inserted column "Column" with the first number no
          // column of the table holds, from 1 -- not by its place: Windows Excel
          // 16 named a column inserted 29th in Table128 "Column1" (office-truth
          
          const name = defaultTableColumnName(
            columns.map((column) => column.name),
            (number) => `${spreadsheetDefaultNames().column}${number}`,
          );
          const id = Math.max(0, ...columns.map((column) => column.id)) + 1;
          columns.splice(offset, 0, { id, name });
        }
        next.columns = columns;
        if (removedColumnId !== undefined && table.columnFilters) {
          const filters = { ...table.columnFilters };
          delete filters[String(removedColumnId)];
          if (Object.keys(filters).length > 0) next.columnFilters = filters;
          else delete next.columnFilters;
        }
        if (removedColumnId !== undefined && table.sortState) {
          const conditions = table.sortState.conditions.filter(
            (condition) => condition.columnId !== removedColumnId,
          );
          if (conditions.length > 0) next.sortState = { conditions };
          else delete next.sortState;
        }
      }
    }
    if (axis === "row" && table.calculatedFormulas) {
      const firstDataRow = r1 + (table.headerRow ? 1 : 0);
      // The recipe reads like the column's first data cell, and follows the
      // edit as that cell does: a row put in the table moves a reference below
      
      // row changes which cell is first: the recipe is the next row's formula,
      // moved by the edit, read from the first row again, so what it says
      // relative to its own row stays.
      const followed = (formula: string): string => {
        if (nr1 !== r1 || at !== firstDataRow) {
          return adjustFormulaForStructuralChange(formula, axis, at, delta, scope);
        }
        return delta > 0
          ? shiftFormula(adjustFormulaForStructuralChange(formula, axis, at, delta, scope), -delta, 0)
          : adjustFormulaForStructuralChange(shiftFormula(formula, 1, 0), axis, at, delta, scope);
      };
      next.calculatedFormulas = Object.fromEntries(
        Object.entries(table.calculatedFormulas).map(([offset, rawFormula]) => [
          offset,
          typeof rawFormula === "string" ? followed(rawFormula) : rawFormula,
        ]),
      );
    }
    if (axis === "row" && next.totalFormulas) {
      // A total's cell stands in one row and moves whole: its text follows.
      next.totalFormulas = Object.fromEntries(
        Object.entries(next.totalFormulas).map(([offset, rawFormula]) => [
          offset,
          typeof rawFormula === "string"
            ? adjustFormulaForStructuralChange(rawFormula, axis, at, delta, scope)
            : rawFormula,
        ]),
      );
    }
    out.push(next);
  }
  return out;
}

/**
 * Remap an index-keyed map (column-width / row-height / AutoFilter) whose keys
 * are stringified ABSOLUTE line indices on `axis`: a key on the deleted line is
 * dropped, keys past `at` shift by `delta`, earlier keys are untouched. Only
 * moves when `active` (the map's axis matches the op axis); a non-integer key
 * rides through. Mirrors remapCells' index math so a per-line override / column
 * filter tracks its physical line on insert/delete (else the custom size or
 * filter silently attaches to the wrong line after the next edit).
 */
function remapIndexKeyedMap<V>(
  map: Record<string | number, V> | undefined,
  active: boolean,
  at: number,
  delta: number,
  // An entry pushed past the grid's last line (a style on the last column
  // when a column is inserted) leaves the sheet, as it does in Excel.
  ceiling: number = MAX_GRID_ROWS,
): Record<string | number, V> | undefined {
  if (!map || !active) return map;
  const out: Record<string, V> = {};
  for (const key of Object.keys(map)) {
    const idx = Number(key);
    if (!Number.isInteger(idx)) {
      out[key] = map[key];
      continue;
    }
    const next = shiftCoord(idx, at, delta);
    if (next === null || next >= ceiling) continue; // deleted line / past the grid -> drop
    out[String(next)] = map[key];
  }
  return out;
}

/**
 * Remap defined names (named ranges): each carries an A1 `ref` string that must
 * shift with the grid like a formula reference. Unlike a sparkline, an erased
 * target is KEPT as "#REF!" (the name still exists, it just errors when used —
 * Excel behaviour), not dropped.
 */
/** A reference text with every quoted sheet prefix ('My Sheet'!) written
 * bare, so two spellings of the same reference compare equal. */
function unquotedSheetPrefixes(text: string): string {
  return text.replace(/'((?:[^']|'')+)'!/g, (_match, name: string) => `${name.replace(/''/g, "'")}!`);
}

function remapNamedRanges<N extends StructuralNamedRange>(
  named: N[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): N[] | undefined {
  if (!named) return named;
  return named.map((entry) => {
    if (typeof entry.ref !== "string") return entry;
    // A name for another name of this workbook (`[0]!Application`, an old
    // file's macro-sheet name) points at no cell: Excel keeps it as it is.
    // The formula rewrite dropped the `[0]!` on the edited sheet's copy, and
    
    if (/^\[0\]![^!:$]+$/.test(entry.ref)) return entry;
    // A named FORMULA / LAMBDA (not a plain cell/range — e.g. "LAMBDA(x,x*2)",
    // "SUM(A1:A10)") shifts its INTERNAL cell refs like any formula, but must NOT
    // be collapsed wholesale to "#REF!" when one ref erases: the formula keeps its
    // other terms (the erased ref becomes inline #REF!, exactly like a cell
    // formula). The whole-entry erasure is range-only semantics. With a scope,
    // each of its references moves only when it targets the edited sheet.
    const plain = isCellRefString(entry.ref) || isRangeRefString(entry.ref);
    // A relative name's offsets (from the cell that uses it) are not cells
    // and do not move, as in Excel; its absolute axes do.
    const relative = protectRelativeNameRefs(entry.ref, axis);
    const source = relative?.text ?? entry.ref;
    const adjusted = scope && !plain
      ? adjustFormulaForStructuralChange(source, axis, at, delta, scope)
      : shiftA1RangeString(source, axis, at, delta, scope);
    const next = relative ? relative.restore(adjusted) : adjusted;
    // A name the edit does not move keeps its text: the formula rewrite
    // re-spells a sheet prefix it leaves pointing at the same cells
    // ('参数与图表'!$G$1 -> 参数与图表!$G$1). A workbook-scoped name is
    // mirrored on every sheet, and only the edited sheet's copy was
    // re-spelled, so the copies disagreed and the export declined the
    // edited sheet's every edit (defined_name_identity_ambiguous).
    if (next === entry.ref || unquotedSheetPrefixes(next) === unquotedSheetPrefixes(entry.ref)) return entry;
    if (!plain) {
      return { ...entry, ref: next };
    }
    return { ...entry, ref: isErasedA1Range(next) ? "#REF!" : next };
  });
}

/**
 * Remap scenarios (Scenario Manager): each scenario's per-cell A1 `ref` shifts
 * with the grid like a formula reference (reusing the SAME shiftA1RangeString +
 * isErasedA1Range helpers as remapNamedRanges), BUT the drop rule differs:
 *  - INSERT above/at the ref's line shifts it outward (e.g. row insert at 0
 *    moves "B2" -> "B3"); a ref before the insert point is untouched.
 *  - DELETE of the ref's OWN line ERASES it (the engine surfaces "#REF!") and
 *    the cell is DROPPED from the scenario — unlike a named range, a scenario
 *    override has no cell to write once its target is deleted, exactly as the
 *    deleted cell itself is removed from `cells`.
 *  - DELETE of a line BEFORE the ref shifts it inward ("B5" -> "B4").
 * A scenario whose cells all drop is KEPT as an empty scenario (it can still be
 * re-selected by name and simply writes nothing), matching normalizeScenario.
 */
function remapScenarios<S extends StructuralScenario>(
  scenarios: S[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): S[] | undefined {
  if (!scenarios) return scenarios;
  return scenarios.map((scenario) => {
    if (!Array.isArray(scenario.cells)) return scenario;
    const cells: StructuralScenarioCell[] = [];
    for (const cell of scenario.cells) {
      if (typeof cell.ref !== "string") {
        cells.push(cell);
        continue;
      }
      const next = shiftA1RangeString(cell.ref, axis, at, delta);
      if (isErasedA1Range(next)) continue; // ref on the deleted line -> drop
      cells.push({ ...cell, ref: next });
    }
    return { ...scenario, cells };
  });
}

/**
 * Remap What-If Data Tables (what-if-table-definition) over a line edit: the
 * whole table -- its results with the row and column of values beside them
 * -- shifts as a range does, and each input cell as a reference does. A table
 * the edit erased, cut or split, or whose input cell it deleted, is dropped
 * (fail closed: its result cells keep the values they show). The window
 * refuses such edits first ("Cannot change part of a data table.").
 */
function remapWhatIfTables(
  tables: readonly WhatIfTableDefinition[],
  axis: StructuralAxis,
  at: number,
  delta: number,
): WhatIfTableDefinition[] {
  const next: WhatIfTableDefinition[] = [];
  for (const block of whatIfTableBlocks(tables)) {
    const footprint = whatIfTableFootprint(block);
    const moved = shiftA1RangeString(whatIfTableRef(footprint), axis, at, delta);
    const area = isErasedA1Range(moved) ? null : parseA1Range(moved.replace(/\$/g, ""));
    if (
      !area ||
      area.r2 - area.r1 !== footprint.r2 - footprint.r1 ||
      area.c2 - area.c1 !== footprint.c2 - footprint.c1
    ) {
      continue;
    }
    const input = (cell: { row: number; col: number } | null): string | null | undefined => {
      if (!cell) return undefined;
      const shifted = shiftA1RangeString(`${columnToLetters(cell.col)}${cell.row + 1}`, axis, at, delta);
      return isErasedA1Range(shifted) ? null : shifted.replace(/\$/g, "");
    };
    const rowInput = input(block.rowInput);
    const colInput = input(block.colInput);
    if (rowInput === null || colInput === null) continue;
    next.push({
      ref: whatIfTableRef({ r1: area.r1 + 1, c1: area.c1 + 1, r2: area.r2, c2: area.c2 }),
      ...(rowInput ? { rowInput } : {}),
      ...(colInput ? { colInput } : {}),
    });
  }
  return normalizeWhatIfTables(next);
}

/**
 * Remap a pivot recipe: its A1 `source` range string shifts with the grid (an
 * erased source becomes "#REF!", which the pivot renders as empty/error rather
 * than silently reading the wrong cells).
 */
function remapPivot<P extends StructuralPivot>(
  pivot: P | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): P | undefined {
  if (!pivot || typeof pivot.source !== "string" || pivot.source.length === 0) {
    return pivot;
  }
  // A Scenario PivotTable report has no worksheet range behind it (its
  // `source` is only the cache's shape), so line edits never reach it.
  if (pivot.cacheSource === "scenario") return pivot;
  const shifted = shiftA1RangeString(pivot.source, axis, at, delta);
  const nextSource = isErasedA1Range(shifted) ? "#REF!" : shifted;
  // Field offsets are SOURCE-column-relative, so only a COLUMN op moves them; a
  // row op (or an erased source) leaves them as carried by the spread.
  // An imported (native) report reads its fields from the workbook's pivot
  // cache, as Excel does: a column inserted through, or deleted from, the
  // source range resizes the range and nothing else until the report is
  // refreshed, so its field offsets stay on the cache's columns (structural
  
  // refused). Refresh then finds each field again by name and drops one
  // whose column is gone (refreshPivotSnapshot).
  const oldBounds = parseA1Range(pivot.source);
  const newBounds = parseA1Range(nextSource);
  if (axis !== "col" || !oldBounds || !newBounds || pivot.native) {
    return { ...pivot, source: nextSource };
  }
  const oldC1 = oldBounds.c1;
  const newC1 = newBounds.c1;
  const width = newBounds.c2 - newC1;
  // An offset -> its new offset (relative to the NEW first column), or null when
  // its absolute column was the deleted one or it falls outside the new source.
  const remapOffset = (offset: number): number | null => {
    if (!Number.isInteger(offset) || offset < 0) return offset;
    const nextAbs = shiftCoord(oldC1 + offset, at, delta);
    if (nextAbs === null) return null;
    const nextOffset = nextAbs - newC1;
    return nextOffset >= 0 && nextOffset <= width ? nextOffset : null;
  };
  const remapOffsets = (offsets: number[] | undefined): number[] | undefined =>
    Array.isArray(offsets)
      ? offsets
          .map(remapOffset)
          .filter((value): value is number => value !== null)
      : offsets;
  const remapFields = <F extends { colIndex?: number }>(
    fields: F[] | undefined,
  ): F[] | undefined => {
    if (!Array.isArray(fields)) return fields;
    const out: F[] = [];
    for (const field of fields) {
      if (typeof field?.colIndex !== "number") {
        out.push(field);
        continue;
      }
      const nextOffset = remapOffset(field.colIndex);
      if (nextOffset === null) continue; // source column deleted -> drop the field
      out.push({ ...field, colIndex: nextOffset });
    }
    return out;
  };
  return {
    ...pivot,
    source: nextSource,
    ...(pivot.rows !== undefined ? { rows: remapOffsets(pivot.rows) } : {}),
    ...(pivot.columns !== undefined
      ? { columns: remapOffsets(pivot.columns) }
      : {}),
    ...(pivot.values !== undefined ? { values: remapFields(pivot.values) } : {}),
    ...(pivot.filters !== undefined
      ? { filters: remapFields(pivot.filters) }
      : {}),
    ...(pivot.grouping !== undefined
      ? { grouping: remapFields(pivot.grouping) }
      : {}),
  };
}

/** The report with its output anchor ("A9" / "A9:D12") shifted for a line
 *  edit on its own sheet; an anchor the edit erases stays where it was. */
function shiftPivotOutput<P extends StructuralPivot>(
  pivot: P,
  axis: StructuralAxis,
  at: number,
  delta: number,
): P {
  const output = (pivot as { output?: unknown }).output;
  if (typeof output !== "string" || !output) return pivot;
  const shifted = shiftA1RangeString(output, axis, at, delta);
  if (shifted === output || isErasedA1Range(shifted)) return pivot;
  return { ...pivot, output: shifted };
}

/**
 * Whether `pivot`, a PivotTable of `host`, reads its source from `edited`.
 * Mirrors the editor's source-sheet lookup: `sourceSheetId` first, then an
 * imported report's sheet name, else the PivotTable's own sheet.
 */
function pivotSourceIsOnSheet(
  pivot: StructuralPivot,
  host: Pick<StructuralSheet, "id" | "name">,
  edited: Pick<StructuralSheet, "id" | "name">,
): boolean {
  if (typeof pivot.sourceSheetId === "string" && pivot.sourceSheetId) {
    return pivot.sourceSheetId === edited.id;
  }
  const nativeName = pivot.native?.sourceSheetName;
  if (typeof nativeName === "string" && nativeName) {
    return nativeName === edited.name;
  }
  return host === edited || (host.id !== undefined && host.id === edited.id);
}

/**
 * Excel adjusts a PivotTable's source range like any other reference when rows
 * or columns are inserted or deleted ON THE SOURCE SHEET, wherever the report
 * itself lives: a line inserted inside the range grows it, a deleted one
 * shrinks it, and a line added just past its end leaves it alone. `host` is a
 * worksheet OTHER than the edited one; its PivotTables that read `edited` get
 * the same remap a same-sheet report gets, `lines` times (positive = inserted
 * at `at`, negative = deleted from `at`). Returns `host` itself when none of
 * its PivotTables reads the edited sheet.
 */
export function remapPivotSourcesOnEditedSheet<T extends StructuralSheet>(
  host: T,
  edited: Pick<StructuralSheet, "id" | "name">,
  axis: StructuralAxis,
  at: number,
  lines: number,
): T {
  const passes = Math.abs(Math.trunc(lines));
  if (passes === 0 || (!host.pivot && !host.morePivots?.length)) return host;
  const delta = lines > 0 ? 1 : -1;
  let touched = false;
  const remap = <P extends StructuralPivot>(pivot: P): P => {
    if (!pivotSourceIsOnSheet(pivot, host, edited)) return pivot;
    let next: P = pivot;
    for (let i = 0; i < passes; i += 1) {
      next = remapPivot(next, axis, at, delta) ?? next;
    }
    touched = true;
    return next;
  };
  const pivot = host.pivot ? remap(host.pivot) : host.pivot;
  const morePivots = host.morePivots?.map(remap);
  if (!touched) return host;
  return {
    ...host,
    ...(host.pivot ? { pivot } : {}),
    ...(host.morePivots ? { morePivots } : {}),
  };
}

/**
 * Excel rewrites every reference whose TARGET is the edited sheet, wherever it
 * is written: `=First!A1` on Second reads `=First!A2` once a row is inserted
 * above row 1 of First, and a deleted target is #REF!. `host` is a worksheet
 * OTHER than the edited one; its cell formulas, custom validation and
 * conditional-format formulas, list-validation sources, defined names,
 * sparkline and chart-object sources and table column formulas that name
 * `edited` move `lines` times (positive = inserted at `at`, negative = deleted
 * from `at`). Its own cells, ranges and anchors do not move, and a reference
 * to any other sheet is untouched. Returns `host` itself when nothing changes.
 */
export function remapReferencesToEditedSheet<T extends StructuralSheet>(
  host: T,
  edited: Pick<StructuralSheet, "id" | "name">,
  axis: StructuralAxis,
  at: number,
  lines: number,
): T {
  const count = Math.trunc(lines);
  if (count === 0 || !edited.name || host === edited || (host.id !== undefined && host.id === edited.id)) return host;
  const scope: StructuralChangeScope = { sheetName: edited.name, qualifiedOnly: true };
  // Only text that names the sheet can move; skip tokenizing everything else.
  // A name with an apostrophe is written with it doubled inside the quotes
  // (='it''s'!F8), so the doubled form names the sheet too; matching the
  // plain name alone skipped those references, which then read the shifted
  // cells instead of following them (sheet model fuzz).
  const needles = [...new Set([edited.name.toUpperCase(), edited.name.replace(/'/g, "''").toUpperCase()])];
  const mentions = (text: unknown): text is string => {
    if (typeof text !== "string") return false;
    const upper = text.toUpperCase();
    return needles.some((needle) => upper.includes(needle));
  };
  // An insert moves every reference by the whole count in one rewrite; a
  // delete is `count` single-line deletes at `at`, as deleteLine does.
  const passes = count > 0 ? [count] : Array.from({ length: -count }, () => -1);
  let next: T = host;
  for (const delta of passes) {
    const patch: Partial<StructuralSheet> = {};
    let cells: Record<string, StructuralCell> | undefined;
    for (const [key, cell] of Object.entries(next.cells)) {
      if (!mentions(cell.formula)) continue;
      const moved = rewriteFormula(cell, axis, at, delta, scope, false);
      if (moved.formula === cell.formula) continue;
      cells ??= { ...next.cells };
      cells[key] = moved;
    }
    if (cells) patch.cells = cells;
    const rules = <R extends object>(list: R[] | undefined): R[] | undefined => {
      if (!list?.some((rule) => ruleFormulaTexts(rule).some(mentions))) return undefined;
      return list.map((rule) => remapRuleFormulas(rule, null, axis, at, delta, scope));
    };
    const validations = rules(next.validations);
    if (validations) patch.validations = validations;
    const conditionalFormats = rules(next.conditionalFormats);
    if (conditionalFormats) patch.conditionalFormats = conditionalFormats;
    if (next.namedRanges?.some((entry) => mentions(entry.ref))) {
      patch.namedRanges = remapNamedRanges(next.namedRanges, axis, at, delta, scope);
    }
    if (next.sparklines && Object.values(next.sparklines).some((value) => mentions((value as { range?: unknown } | null)?.range))) {
      const sparklines: StructuralCellKeyed = {};
      for (const [key, value] of Object.entries(next.sparklines)) {
        const moved = adjustSparklineValue(value, axis, at, delta, scope);
        if (moved !== null) sparklines[key] = moved;
      }
      patch.sparklines = sparklines;
    }
    if (next.objects?.some((object) => mentions(object.range) || object.ranges?.some(mentions))) {
      patch.objects = remapObjects(next.objects, axis, at, delta, scope, false);
    }
    if (next.dataTables?.some((table) => Object.values(table.calculatedFormulas ?? {}).some(mentions) || Object.values(table.totalFormulas ?? {}).some(mentions))) {
      const moveMap = (map: Record<number, unknown> | undefined) => map && Object.fromEntries(Object.entries(map).map(([offset, formula]) =>
        [offset, mentions(formula) ? adjustFormulaForStructuralChange(formula, axis, at, delta, scope) : formula]));
      patch.dataTables = next.dataTables.map((table) => ({
        ...table,
        ...(table.calculatedFormulas ? { calculatedFormulas: moveMap(table.calculatedFormulas) } : {}),
        ...(table.totalFormulas ? { totalFormulas: moveMap(table.totalFormulas) } : {}),
      }));
    }
    if (Object.keys(patch).length > 0) next = { ...next, ...patch } as T;
  }
  return next;
}

/**
 * Whether inserting/deleting `lines` at `at` on `edited` would have to move a
 * 3-D reference (`First:Third!A5`) on any sheet of `workbook` whose span
 * includes `edited`: see structuralChangeMovesThreeDReference. A span whose
 * end sheets cannot be found is taken to include it.
 */
export function structuralEditMovesThreeDReference(
  workbook: ReadonlyArray<StructuralSheet>,
  edited: Pick<StructuralSheet, "id" | "name">,
  axis: StructuralAxis,
  at: number,
  lines: number,
): boolean {
  const count = Math.trunc(lines);
  if (count === 0) return false;
  const order = workbook.map((sheet) => (sheet.name ?? "").toUpperCase());
  const index = workbook.findIndex((sheet) => sheet === edited || (edited.id !== undefined && sheet.id === edited.id));
  const spans = (from: string, to: string) => {
    const a = order.indexOf(from.toUpperCase()), b = order.indexOf(to.toUpperCase());
    return a < 0 || b < 0 || index < 0 || (index >= Math.min(a, b) && index <= Math.max(a, b));
  };
  const delta = count > 0 ? count : -1;
  const texts = (sheet: StructuralSheet): unknown[] => [
    ...Object.values(sheet.cells).map((cell) => cell.formula),
    ...(sheet.validations ?? []).flatMap((rule) => ruleFormulaTexts(rule)),
    ...(sheet.conditionalFormats ?? []).flatMap((rule) => ruleFormulaTexts(rule)),
    ...(sheet.namedRanges ?? []).map((entry) => entry.ref),
  ];
  return workbook.some((sheet) => texts(sheet).some((text) =>
    typeof text === "string" && text.includes(":") && text.includes("!") &&
    structuralChangeMovesThreeDReference(text, axis, at, delta, spans)));
}

/**
 * Remap column-bound overlays (slicers): only a COLUMN insert/delete moves them;
 * a slicer on the deleted column is dropped, later columns shift by `delta`.
 */
function remapColumnBound<O extends StructuralColumnBound>(
  bound: O[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): O[] | undefined {
  if (!bound || axis !== "col") return bound;
  const out: O[] = [];
  for (const overlay of bound) {
    const nextCol = shiftCoord(overlay.column, at, delta);
    if (nextCol === null) continue; // the slicer's own column was deleted
    out.push({ ...overlay, column: nextCol });
  }
  return out;
}

/**
 * Remap watched cells (Watch Window): each is a single grid point whose
 * coordinate on the ACTIVE axis shifts like any cell, and which is DROPPED when
 * that exact line is deleted (the watch had no surviving target). A point on the
 * other axis is untouched. Mirrors the slicer column-drop semantics, one point.
 */
function remapWatches<W extends { row: number; col: number }>(
  watches: W[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): W[] | undefined {
  if (!watches) return watches;
  const out: W[] = [];
  for (const w of watches) {
    const next = shiftCoord(axis === "row" ? w.row : w.col, at, delta);
    if (next === null) continue; // the watched cell's own line was deleted
    out.push({
      ...w,
      row: axis === "row" ? next : w.row,
      col: axis === "col" ? next : w.col,
    });
  }
  return out;
}

/**
 * Remap anchor+span merged blocks (the `merges` model from merge-cells.ts).
 * Each block is a top-left (row, col) + rowSpan/colSpan; it is converted to
 * inclusive edges on the active axis, both edges are shifted like a range
 * (start moves only when past the line, end moves at-or-past it), then converted
 * back to anchor+span. Edge cases, matching Excel + the merge model:
 *  - INSERT widens a block straddling the insert point and shifts a block fully
 *    after it; a block fully before it is untouched.
 *  - DELETE of a line inside the block shrinks it by one on that axis.
 *  - A block that COLLAPSES to a single cell (1x1) after a delete is DROPPED — a
 *    1x1 merge covers nothing (mergeSpanFor / normalizeSpan reject 1x1, so the
 *    model treats it as "no merge"; we keep that invariant).
 *  - A block fully on the single deleted line erases entirely (span -> 0) and is
 *    dropped.
 * Only the ACTIVE axis moves; the other axis's anchor/span are untouched.
 */
function remapMerges<M extends StructuralMerge>(
  merges: M[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): M[] | undefined {
  if (!merges) return merges;
  const out: M[] = [];
  for (const merge of merges) {
    let start = axis === "row" ? merge.row : merge.col;
    let end = start + (axis === "row" ? merge.rowSpan : merge.colSpan) - 1;
    if (delta > 0) {
      if (start >= at) start += delta;
      if (end >= at) end += delta;
    } else {
      // Delete: the start moves up only if strictly past the line; the end moves
      // up at or past it. A block whose every line is the single deleted line
      // collapses (end < start) and is dropped below.
      if (start === at && end === at) continue;
      if (start > at) start -= 1;
      if (end >= at) end -= 1;
    }
    const limit = (axis === "row" ? MAX_ROWS : MAX_COLS) - 1;
    if (start > limit) continue; // pushed off the grid
    if (end > limit) end = limit; // a whole-row/column merge keeps the last line
    const span = end - start + 1;
    if (span < 1) continue; // fully erased
    const nextRowSpan = axis === "row" ? span : merge.rowSpan;
    const nextColSpan = axis === "col" ? span : merge.colSpan;
    // A merge that has shrunk to a single cell (1x1) is meaningless — drop it.
    if (nextRowSpan <= 1 && nextColSpan <= 1) continue;
    out.push({
      ...merge,
      row: axis === "row" ? start : merge.row,
      col: axis === "col" ? start : merge.col,
      rowSpan: nextRowSpan,
      colSpan: nextColSpan,
    });
  }
  return out;
}

/**
 * Remap an axis-bound inclusive LINE range (outline groups). Only a change on
 * the group's OWN axis moves it (a row group ignores a column op, like a slicer
 * ignores a row op — but here it is a {start,end} range, not a single line):
 *  - INSERT widens a group straddling the insert point and shifts one fully
 *    after it; a group fully before it is untouched.
 *  - DELETE of a line inside the group shrinks it by one; a group fully on the
 *    single deleted line collapses (end < start) and is dropped.
 *  - A group that shrinks to a single line (start === end) is DROPPED: a group
 *    needs ≥2 lines to have a body to hide (normalizeOutlineGroups rejects
 *    end <= start), so we keep that invariant.
 */
/** The indexes of the collapsed groups on `axis` that a line inserted at `at`
 *  lands inside. */
function groupsOpenedByInsertion(
  groups: readonly StructuralLineRange[],
  axis: StructuralAxis,
  at: number,
  delta: number,
): Set<number> {
  const opened = new Set<number>();
  if (delta <= 0) return opened;
  groups.forEach((group, index) => {
    // At the line right after the group too: the inserted line takes the outline level of the line before it, so the
    // group grows over it (``remapLineRange``) and opens as above (Windows Excel 16, xlsx-rounds
    // probe/night14/x20lvl: B hidden at level 1, C collapsed; a column at C reads C level 1 shown, D not collapsed).
    if (group.axis === axis && (group as { collapsed?: boolean }).collapsed === true && group.start < at && at <= group.end + 1) {
      opened.add(index);
    }
  });
  return opened;
}

function remapLineRange<G extends StructuralLineRange>(
  groups: G[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): G[] | undefined {
  if (!groups) return groups;
  const out: G[] = [];
  for (const group of groups) {
    // A group lives on one axis; a change on the OTHER axis never moves it.
    if (group.axis !== axis) {
      out.push(group);
      continue;
    }
    let { start, end } = group;
    if (delta > 0) {
      if (start >= at) start += delta;
      // A line inserted right after the group's last takes that line's outline level, as Excel gives an inserted
      // row the level of the row above and a column that of the column to its left (Windows Excel 16, xlsx-rounds
      // R432-X20 STRUCT, probe/night14/x20lvl: B at level 1, a column inserted at C reads level 1).
      if (end >= at - 1 && start < at) end += delta;
      else if (end >= at) end += delta;
    } else {
      if (start === at && end === at) continue; // fully on the deleted line
      if (start > at) start -= 1;
      if (end >= at) end -= 1;
    }
    const limit = (axis === "row" ? MAX_ROWS : MAX_COLS) - 1;
    if (start > limit) continue; // pushed off the grid
    end = Math.min(end, limit);
    if (end < start) continue; // a one-line group is legal (Excel groups a single row)
    out.push({ ...group, start, end });
  }
  return out;
}

/**
 * Remap floating charts. A chart owns TWO grid-bound coordinates that must both
 * track an insert/delete or it corrupts on the next edit:
 *  - the SOURCE RANGE {r1,c1,r2,c2}: shifted/shrunk exactly like a validation
 *    rule (reuse remapRanged). The chart is an object of its own, so deleting
 *    every line of its data never deletes it: as Excel turns the series into
 *    `#REF!` and keeps the chart's title, size, place and format, the chart is
 *    kept with `sourceRefError` and a one-cell placeholder range at the
 *    deletion point (it plots nothing until Select Data picks new data). A
 *    #REF! source never moves again.
 *  - the ANCHOR cell {row,col}: shifted exactly like a sparkline's "row:col"
 *    key (reuse shiftCoord on the active axis). A sheet chart is one-cell
 *    anchored (Excel "Move but don't size with cells") or pinned at x/y, and
 *    Excel keeps such an object when the lines under it are deleted, so an
 *    anchor on the deleted line clamps to the deletion point — the chart
 *    floats one cell up/left, onto the new occupant of that position.
 */
function remapCharts<C extends StructuralChart>(
  charts: C[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
): C[] | undefined {
  if (!charts) return charts;
  return charts.map((current) => {
    let chart = current;
    if (!chart.sourceRefError) {
      // The validation-style range remap shifts edges, shrinks a straddle, and
      // erases a range whose only line is the deleted one.
      const moved = remapRanged([chart], axis, at, delta)?.[0];
      chart = moved ?? {
        ...chart,
        sourceRefError: true,
        range:
          axis === "row"
            ? { r1: at, c1: chart.range.c1, r2: at, c2: chart.range.c1 }
            : { r1: chart.range.r1, c1: at, r2: chart.range.r1, c2: at },
      };
    }
    const coord = axis === "row" ? chart.anchor.row : chart.anchor.col;
    // shiftCoord returns null when the anchor is ON the deleted line; the chart
    // survives, so clamp the anchor to the deletion point.
    const shifted = shiftCoord(coord, at, delta);
    const nextCoord = shifted === null ? Math.max(0, at) : shifted;
    return {
      ...chart,
      anchor:
        axis === "row"
          ? { row: nextCoord, col: chart.anchor.col }
          : { row: chart.anchor.row, col: nextCoord },
      ...(chart.drawingAnchor
        ? { drawingAnchor: shiftDrawingAnchorLines(chart.drawingAnchor, axis, at, delta) }
        : {}),
    };
  }).filter((chart, index) =>
    // Every line under a Move-and-size-with-cells chart deleted: Excel
    
    !(delta < 0 && drawingAnchorClosedByDelete(charts[index].drawingAnchor, chart.drawingAnchor, axis)));
}













function shiftAnchorPoint<P extends StructuralAnchorPoint>(
  point: P | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
  



  far = false,
  /** Keep the offset of a point on a deleted line (a layout-relative SmartArt member). */
  keepOffset = false,
): P | undefined {
  if (!point) return point;
  if (far && delta > 0) {
    const place = anchorPointPlace(point, axis);
    if (place && place[0] === at && place[1] <= 0) return point;
  }
  if (axis === "row") {
    if (typeof point.row !== "number") return point;
    const shifted = shiftCoord(point.row, at, delta);
    return shifted === null
      ? { ...point, row: Math.max(0, at), ...(point.rowOffset !== undefined && !keepOffset ? { rowOffset: 0 } : {}) }
      : { ...point, row: shifted };
  }
  // Column op: the col index is stored as `column` OR `col` — shift whichever is
  // present (preferring `column`), leaving the spelling as found.
  const key = typeof point.column === "number" ? "column" : typeof point.col === "number" ? "col" : null;
  if (!key) return point;
  const shifted = shiftCoord(point[key] as number, at, delta);
  if (shifted !== null) return { ...point, [key]: shifted };
  if (keepOffset) return { ...point, [key]: Math.max(0, at) };
  return {
    ...point,
    [key]: Math.max(0, at),
    ...(point.columnOffset !== undefined ? { columnOffset: 0 } : {}),
    ...(point.colOffset !== undefined ? { colOffset: 0 } : {}),
  };
}

/** A marker's place on one axis, for comparing two markers: line, then offset. */
function anchorPointPlace(
  point: StructuralAnchorPoint | undefined,
  axis: StructuralAxis,
): [number, number] | null {
  if (!point) return null;
  const line = axis === "row" ? point.row : (point.column ?? point.col);
  if (typeof line !== "number") return null;
  const offset = Number(axis === "row" ? point.rowOffset : (point.columnOffset ?? point.colOffset));
  return [line, Number.isFinite(offset) ? offset : 0];
}

/** A two-cell anchor whose markers meet on `axis`: none of it is left there. */
function sameCell(a: StructuralAnchorPoint, b: StructuralAnchorPoint): boolean {
  return a.row === b.row && (a.column ?? a.col) === (b.column ?? b.col);
}

function twoCellAnchorCollapsed(
  anchor: StructuralObjectAnchor | undefined,
  axis: StructuralAxis,
): boolean {
  // A SmartArt member closes with its graphic's frame.
  const frame = (anchor as { smartArtFrame?: StructuralObjectAnchor } | undefined)?.smartArtFrame;
  if (frame && typeof frame === "object") return twoCellAnchorCollapsed(frame, axis);
  if (anchor?.kind !== "twoCell") return false;
  const from = anchorPointPlace(anchor.from, axis);
  const to = anchorPointPlace(anchor.to, axis);
  return Boolean(from && to && (to[0] < from[0] || (to[0] === from[0] && to[1] <= from[1])));
}

/**
 * Shift one A1 source-range STRING carried on a chart object ("Sheet1!B2:B10"),
 * keeping an erased target as "#REF!" — the SAME convention as a named range /
 * pivot source (the object is kept; a broken ref errors when read). Reuses
 * shiftA1RangeString (which peels/re-attaches the sheet qualifier and shifts
 * $-absolute refs).
 */
function shiftChartObjectA1(
  ref: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): string {
  const next = shiftA1RangeString(ref, axis, at, delta, scope);
  if (next === ref) return ref;
  return isErasedA1Range(next) ? "#REF!" : next;
}

/** The areas of a reference union a chart series names, `(Sheet1!$A$1,Sheet1!$A$5)`;
 *  null for any other reference. Commas inside a quoted sheet name are its own. */
function chartReferenceUnionAreas(ref: string): string[] | null {
  if (!ref.startsWith("(") || !ref.endsWith(")")) return null;
  const areas: string[] = [];
  let quoted = false;
  let start = 1;
  for (let index = 1; index < ref.length - 1; index += 1) {
    const char = ref[index];
    if (char === "'") quoted = !quoted;
    else if (char === "," && !quoted) {
      areas.push(ref.slice(start, index));
      start = index + 1;
    }
  }
  areas.push(ref.slice(start, ref.length - 1));
  return areas.length > 1 && areas.every((area) => area.length > 0) ? areas : null;
}





function shiftChartSeriesA1(
  ref: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): string {
  const areas = chartReferenceUnionAreas(ref);
  if (areas) {
    const moved = areas.map((area) => shiftChartSeriesA1(area, axis, at, delta, scope));
    return moved.every((area, index) => area === areas[index]) ? ref : `(${moved.join(",")})`;
  }
  const next = shiftChartObjectA1(ref, axis, at, delta, scope);
  if (next !== "#REF!" || ref.startsWith("(")) return next;
  return `${ref.slice(0, ref.lastIndexOf("!") + 1)}#REF!`;
}

/**
 * Remap floating drawing objects (images / shapes / charts). A CHART object also
 * carries its grid-bound SOURCE as A1 STRINGS in `range` (singular) and/or the
 * legacy `ranges` plural; like a named range / pivot source those shift with the
 * grid (erased -> "#REF!"), so an imported chart keeps reading the right cells
 * after a row/column insert/delete instead of going stale. Separately, every
 * object's cell anchor `from`/`to` point slides on the active axis (an `absolute`
 * page-pinned anchor is left untouched — Excel's "Don't move or size with cells").
 * An anchor on the deleted line moves to the deletion point; only a two-cell
 * (Move and size with cells) object left with nothing between its markers is
 * deleted, as Excel deletes it. A non-chart object (or a chart with no
 * range/ranges) and an unaffected object serialize byte-for-byte unchanged.
 */
function remapObjects<O extends StructuralObject>(
  objects: O[] | undefined,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
  /** False for objects on another sheet: only their sources can target the
   * edited sheet; their anchors sit on their own, unedited grid. */
  moveAnchors = true,
): O[] | undefined {
  if (!objects) return objects;
  // Excel deletes a Move-and-size-with-cells object whose every row (or
  // column) is deleted, and keeps every other kind (Excel for Mac 16,
  
  // it; a Move-but-don't-size shape moves to the deletion line, a free
  // floating one stays). A grouped object goes only with its whole group.
  const collapsed = new Set<O>();
  const mapped = objects.map((object) => {
    const next = remapObject(object);
    if (
      moveAnchors &&
      delta < 0 &&
      twoCellAnchorCollapsed(next.anchor, axis) &&
      !twoCellAnchorCollapsed(object.anchor, axis)
    ) {
      collapsed.add(next);
    }
    return next;
  });
  if (collapsed.size === 0) return mapped;
  const groupOf = (object: O) => {
    const id = (object as { groupId?: unknown }).groupId;
    return typeof id === "string" && id ? id : null;
  };
  const survivingGroups = new Set(
    mapped.filter((object) => !collapsed.has(object)).map(groupOf).filter((id): id is string => id !== null),
  );
  return mapped.filter((object) => {
    if (!collapsed.has(object)) return true;
    const group = groupOf(object);
    return group !== null && survivingGroups.has(group);
  });

  function remapObject(object: O): O {
    let next = object;
    // A chart object's A1 source range/ranges follow the grid (the anchor below
    // is handled independently). Non-chart objects have no range.
    if (object.type === "chart") {
      const patch: { range?: string; ranges?: string[] } = {};
      if (typeof object.range === "string") {
        const r = shiftChartObjectA1(object.range, axis, at, delta, scope);
        if (r !== object.range) patch.range = r;
      }
      if (Array.isArray(object.ranges)) {
        let changed = false;
        const nextRanges = object.ranges.map((rg) => {
          if (typeof rg !== "string") return rg;
          const r = shiftChartSeriesA1(rg, axis, at, delta, scope);
          if (r !== rg) changed = true;
          return r;
        });
        if (changed) patch.ranges = nextRanges;
      }
      // Each series' own references (title, categories, values, X values,
      // bubble sizes) move with the grid too; they are the complete list of
      // what the chart part reads, where `ranges` is their summary.
      const seriesRefs = (object as { seriesRefs?: unknown }).seriesRefs;
      let seriesChanged = false;
      const nextSeriesRefs = Array.isArray(seriesRefs)
        ? seriesRefs.map((entry) => {
            if (!entry || typeof entry !== "object") return entry;
            const moved: Record<string, unknown> = { ...(entry as Record<string, unknown>) };
            for (const key of ["title", "categories", "values", "xValues", "bubbleSizes"]) {
              const value = moved[key];
              if (typeof value !== "string" || !value) continue;
              const r = shiftChartSeriesA1(value, axis, at, delta, scope);
              if (r !== value) {
                moved[key] = r;
                seriesChanged = true;
              }
            }
            return moved;
          })
        : undefined;
      if (patch.range !== undefined || patch.ranges !== undefined || seriesChanged) {
        next = { ...object, ...patch, ...(seriesChanged ? { seriesRefs: nextSeriesRefs } : {}) };
      }
    }
    const anchor = next.anchor;
    // Absolute (page-pinned) anchors do not track cells; leave the anchor as-is.
    if (!moveAnchors || !anchor || anchor.kind === "absolute") return next;
    return {
      ...next,
      anchor: shiftDrawingAnchorLines(anchor, axis, at, delta, {
        keepOffsets: spreadsheetObjectOffsetsLayoutRelative(next),
      }),
    };
  }
}

/**
 * A cell anchor after `delta` whole lines are inserted (positive) or deleted
 * (negative) at line `at` on `axis`, as Excel moves a drawing's markers: the
 * same reading for Insert/Delete rows or columns and for Insert/Delete cells
 * that shift the object's column (or row). An absolute anchor is returned as
 * it is (Don't move or size with cells).
 */
export function shiftDrawingAnchorLines<A extends StructuralObjectAnchor>(
  anchor: A,
  axis: StructuralAxis,
  at: number,
  delta: number,
  /** `keepOffsets`: a layout-relative SmartArt member keeps its offsets. */
  options: { readonly keepOffsets?: boolean } = {},
): A {
  // A SmartArt member goes with its graphic's frame, the one anchored
  // object Excel has: the frame's markers move (and a two-cell frame
  // stretches), the member stays on the frame's corner cell at its place in
  // the layout (lib/spreadsheet/object-buttons readSmartArtFrame).
  const frame = (anchor as { smartArtFrame?: StructuralObjectAnchor & Record<string, unknown> }).smartArtFrame;
  if (frame && typeof frame === "object" && frame.from && anchor.from && sameCell(anchor.from, frame.from)) {
    if (frame.kind === "absolute" || delta === 0) return anchor;
    const moved = shiftDrawingAnchorLines(frame, axis, at, delta);
    const offsetKey = axis === "row" ? "rowOffset" : "columnOffset";
    const lineOf = (point: StructuralAnchorPoint) => (axis === "row" ? point.row : (point.column ?? point.col));
    const offsetOf = (point: StructuralAnchorPoint) =>
      Number((axis === "row" ? point.rowOffset : (point.columnOffset ?? point.colOffset)) ?? 0) || 0;
    const from = anchor.from;
    const line = lineOf(moved.from!);
    const shiftedOffset = Math.max(0, offsetOf(from) - offsetOf(frame.from) + offsetOf(moved.from!));
    const nextFrom: StructuralAnchorPoint =
      axis === "row"
        ? { ...from, row: line, [offsetKey]: shiftedOffset }
        : typeof from.column === "number"
          ? { ...from, column: line, columnOffset: shiftedOffset }
          : { ...from, col: line, colOffset: shiftedOffset };
    return { ...anchor, from: nextFrom, smartArtFrame: moved };
  }
  if (anchor.kind === "absolute" || delta === 0) return anchor;
  let next: A = anchor;
  const keep = options.keepOffsets === true;
  // A delete of N lines is N single-line deletes at the same index.
  for (let pass = 0; pass < (delta < 0 ? -delta : 1); pass += 1) {
    const step = delta < 0 ? -1 : delta;
    next = {
      ...next,
      ...(next.from ? { from: shiftAnchorPoint(next.from, axis, at, step, false, keep) } : {}),
      ...(next.to ? { to: shiftAnchorPoint(next.to, axis, at, step, true, keep) } : {}),
    };
  }
  return next;
}

/** How a line edit moved an anchor's markers on `axis`: not at all, every
 * one by the same `delta` lines (offsets kept), or some other way. */
function anchorLinesMoved(
  before: StructuralObjectAnchor,
  after: StructuralObjectAnchor,
  axis: StructuralAxis,
  delta: number,
): "none" | "all" | "mixed" {
  const moves = (["from", "to"] as const).flatMap((key) => {
    const was = anchorPointPlace(before[key], axis);
    const now = anchorPointPlace(after[key], axis);
    if (!was && !now) return [];
    if (!was || !now || was[1] !== now[1]) return [NaN];
    return [now[0] - was[0]];
  });
  if (moves.every((move) => move === 0)) return "none";
  return moves.every((move) => move === delta) ? "all" : "mixed";
}

/**
 * The sheet after a row or column insert or delete (`after`, its objects
 * moved by their frames), with each picture or shape turned a quarter moved
 * by its Excel box instead (drawing-rotation-box): Excel moves the box's
 * markers, and the frame is that box turned back on the grid after the edit.
 * A turned object whose box and frame the edit moves alike (both before it,
 * both past it) keeps the frame's own move.
 */
function withTurnedObjectsMovedByBox<T extends StructuralSheet>(
  before: T,
  after: T,
  axis: StructuralAxis,
  at: number,
  delta: number,
): T {
  const turnable = (object: StructuralObject) => {
    const record = object as StructuralObject & { id?: unknown; groupId?: unknown };
    const anchor = record.anchor as (StructuralObjectAnchor & { rotation?: unknown; smartArtFrame?: unknown }) | undefined;
    return (
      (record.type === "image" || record.type === "shape") &&
      typeof record.id === "string" &&
      !record.groupId &&
      Boolean(anchor) &&
      !anchor!.smartArtFrame &&
      (anchor!.kind === "oneCell" || anchor!.kind === "twoCell") &&
      anchorBoxTurned(anchor!.rotation)
    );
  };
  const turned = before.objects?.filter(turnable);
  if (!turned?.length || !after.objects || delta === 0) return after;
  const grid = sheetDrawingBoxAxes(before as unknown as DrawingBoxSheet);
  const edited = {
    columns: axis === "col" ? drawingBoxAxisAfterLineEdit(grid.columns, at, delta) : grid.columns,
    rows: axis === "row" ? drawingBoxAxisAfterLineEdit(grid.rows, at, delta) : grid.rows,
  };
  const frames = new Map(turned.map((object) => [(object as { id: string }).id, object.anchor!]));
  let changed = false;
  const objects = after.objects.map((object) => {
    const frame = frames.get((object as { id?: string }).id ?? "");
    if (!frame || !object.anchor) return object;
    const box = turnDrawingAnchorBox(frame, grid.columns, grid.rows);
    if (!box) return object;
    const movedBox = shiftDrawingAnchorLines(box, axis, at, delta);
    const boxMove = anchorLinesMoved(box, movedBox, axis, delta);
    if (boxMove !== "mixed" && boxMove === anchorLinesMoved(frame, object.anchor, axis, delta)) return object;
    const placed = turnDrawingAnchorBox(movedBox, edited.columns, edited.rows);
    if (!placed) return object;
    changed = true;
    return { ...object, anchor: { ...object.anchor, ...placed } };
  });
  return changed ? { ...after, objects } : after;
}

/** A two-cell (Move and size with cells) anchor a delete closed on `axis`:
 * Excel deletes its object. */
export function drawingAnchorClosedByDelete(
  before: StructuralObjectAnchor | undefined,
  after: StructuralObjectAnchor | undefined,
  axis: StructuralAxis,
): boolean {
  return twoCellAnchorCollapsed(after, axis) && !twoCellAnchorCollapsed(before, axis);
}

/**
 * Remap a list of ABSOLUTE line indices (manually hidden rows / columns) on
 * `axis`: only the matching axis moves; an index ON the deleted line is DROPPED,
 * later indices shift by `delta` (reusing shiftCoord, the same per-index math as
 * remapCells / remapIndexKeyedMap). A non-finite/non-integer entry is dropped.
 * Returns a NEW, de-duplicated array (or the input verbatim on the other axis).
 * Carried through stale previously would mean the wrong row/col silently
 * collapses after the next insert/delete.
 */
function remapHiddenIndices(
  indices: number[] | undefined,
  active: boolean,
  at: number,
  delta: number,
): number[] | undefined {
  if (!indices || !active) return indices;
  const seen = new Set<number>();
  const out: number[] = [];
  for (const raw of indices) {
    if (!Number.isInteger(raw)) continue;
    const next = shiftCoord(raw, at, delta);
    if (next === null) continue; // on the deleted line -> drop
    if (seen.has(next)) continue;
    seen.add(next);
    out.push(next);
  }
  return out;
}

/** Shift an inclusive [start, end] line run for one structural pass: an insert
 * at or before the start moves it, one inside widens it; a delete shrinks it,
 * or erases it (null) when its only line goes. */
function shiftLineRun(
  start: number,
  end: number,
  at: number,
  delta: number,
): [number, number] | null {
  if (delta > 0) return [start >= at ? start + delta : start, end >= at ? end + delta : end];
  if (start === at && end === at) return null;
  return [start > at ? start - 1 : start, end >= at ? end - 1 : end];
}

/** Print titles ("1:2" rows / "A:B" columns) follow their lines as Excel's
 * Print_Titles name does: moved by an insert before them, widened by one
 * inside, shrunk by a partial delete, and cleared once every title line is
 * deleted. Only the edited axis's titles are passed in. */
function shiftPrintTitles(
  value: string,
  axis: StructuralAxis,
  at: number,
  delta: number,
): string | undefined {
  const range = parseA1Range(value);
  if (!range) return value;
  const run = shiftLineRun(
    axis === "row" ? range.r1 : range.c1,
    axis === "row" ? range.r2 : range.c2,
    at,
    delta,
  );
  if (!run) return undefined;
  return axis === "row"
    ? `${run[0] + 1}:${run[1] + 1}`
    : `${columnToLetters(run[0])}:${columnToLetters(run[1])}`;
}

/** Manual page breaks (the 0-based line each new page starts at) belong to
 * that line, as Excel's Range.PageBreak does: an insert at or before it pushes
 * it on, deleting that line removes its break, later breaks move up. A break
 * before the first line means nothing and drops; coinciding breaks are one. */
function shiftManualBreaks(
  breaks: readonly number[],
  at: number,
  delta: number,
): number[] {
  const out = new Set<number>();
  for (const line of breaks) {
    if (!Number.isInteger(line)) continue;
    const next = shiftCoord(line, at, delta);
    if (next === null || next <= 0) continue;
    out.add(next);
  }
  return [...out].sort((a, b) => a - b);
}

/**
 * One structural pass over the print settings: the print area, the repeated
 * title rows/columns and the manual page breaks of the edited axis move
 * together, so the edited sheet still repeats and paginates the same content.
 * Returns the input itself when nothing moves.
 */
function remapPrintSettings<P extends NonNullable<StructuralSheet["printSettings"]>>(
  settings: P,
  axis: StructuralAxis,
  at: number,
  delta: number,
  scope?: StructuralChangeScope,
): P {
  const next: P = { ...settings };
  let changed = false;
  // The print area is an A1 range that tracks the grid like a named range; a
  // fully-erased area is dropped (Excel clears a print area whose every
  // column/row is deleted). Excel keeps several comma-separated areas in one
  // name; each block shifts on its own and an erased block leaves the list.
  if (typeof settings.printArea === "string" && settings.printArea) {
    const shifted = settings.printArea
      .split(",")
      .map((block) => block.trim())
      .filter((block) => block.length > 0)
      .map((block) => shiftA1RangeString(block, axis, at, delta, scope))
      .filter((block) => !isErasedA1Range(block));
    if (shifted.length === 0) delete next.printArea;
    else next.printArea = shifted.join(",");
    changed = true;
  }
  const titleKey = axis === "row" ? "repeatTitleRows" : "repeatTitleCols";
  const titles = settings[titleKey];
  if (typeof titles === "string" && titles) {
    const shifted = shiftPrintTitles(titles, axis, at, delta);
    if (shifted !== titles) {
      if (shifted === undefined) delete next[titleKey];
      else next[titleKey] = shifted;
      changed = true;
    }
  }
  const breakKey = axis === "row" ? "manualRowBreaks" : "manualColBreaks";
  const breaks = settings[breakKey];
  if (Array.isArray(breaks) && breaks.length > 0) {
    const shifted = shiftManualBreaks(breaks, at, delta);
    if (shifted.length !== breaks.length || shifted.some((line, index) => line !== breaks[index])) {
      if (shifted.length === 0) delete next[breakKey];
      else next[breakKey] = shifted;
      changed = true;
    }
  }
  return changed ? next : settings;
}

/** Apply the keyed/ranged/column-bound overlay remaps that ride alongside cells. */
function remapOverlays<T extends StructuralSheet>(
  sheet: T,
  axis: StructuralAxis,
  at: number,
  delta: number,
): Partial<T> {
  const patch: Partial<StructuralSheet> = {};
  // References written with a sheet name follow only when that name is this
  // sheet's; one naming another sheet stays where it points.
  const scope = structuralScopeOf(sheet);
  // A sparkline shifts by its anchor key AND has its embedded source-range
  // STRING adjusted (remapKeyed alone would leave "B1:B5" pointing at the old
  // cells); a sparkline whose entire source range is erased is dropped.
  if (sheet.sparklines)
    patch.sparklines = remapKeyed(sheet.sparklines, axis, at, delta, (value) =>
      adjustSparklineValue(value, axis, at, delta, scope),
    );
  if (sheet.validations) patch.validations = remapRanged(sheet.validations, axis, at, delta, scope);
  // A structured table is range-bound like a validation rule, BUT it also carries
  // a per-column `totals` map keyed by column offset — so it needs a dedicated
  // remap that shifts the totals offsets on a column op (else the totals row
  // aggregates the wrong physical columns after an insert/delete). The range
  // remap matches validations; a table fully on the deleted line drops.
  if (sheet.dataTables) patch.dataTables = remapDataTables(sheet.dataTables, axis, at, delta, scope);
  // A conditional-format rule is range-bound exactly like a validation rule —
  // same {r1,c1,r2,c2} shape — so it reuses the same range remap (shift on
  // insert, shrink on delete-inside, drop when its whole range is the deleted
  // line). Previously carried through verbatim at OLD coordinates = the blind
  // spot: a rule recoloring a deleted/shifted band after the next edit.
  if (sheet.conditionalFormats)
    patch.conditionalFormats = remapRanged(sheet.conditionalFormats, axis, at, delta, scope);
  // The sheet's Ignore Error ranges move with their cells, as Excel moves an
  
  if (sheet.ignoredErrorRanges)
    patch.ignoredErrorRanges = remapRanged(sheet.ignoredErrorRanges, axis, at, delta);
  // A hyperlink the sheet holds over a large area moves as Excel moves its
  // <hyperlink ref>: grown by an insert inside it, gone with its last line,
  
  if (sheet.hyperlinkAreas)
    patch.hyperlinkAreas = remapRanged(sheet.hyperlinkAreas, axis, at, delta);
  if (sheet.slicers) patch.slicers = remapColumnBound(sheet.slicers, axis, at, delta);
  // Watched cells (Watch Window): a single point that shifts on its line's axis,
  // and is DROPPED when that exact line is deleted — else the pinned watch would
  // silently point at whatever content slid into its old coordinate.
  if (sheet.watches) patch.watches = remapWatches(sheet.watches, axis, at, delta);
  // Anchor+span merged blocks: edges shift; a block that collapses to 1x1 or is
  // fully erased is dropped (a 1x1 merge is meaningless, matching the model).
  if (sheet.merges) patch.merges = remapMerges(sheet.merges, axis, at, delta);
  // Axis-bound outline groups: only the matching axis moves; a group whose
  // every line is deleted is dropped.
  let opened = new Set<number>();
  if (sheet.outlineGroups) {
    patch.outlineGroups = remapLineRange(sheet.outlineGroups, axis, at, delta);
    // A line inserted inside a collapsed group shows, and the group is no longer
    // collapsed: its other lines stay hidden, each on its own (Excel 16 for Mac,
    // xlsx-rounds evidence/xprobe-outlineinsert: F:G collapsed, a column at G
    // shows between hidden F and H, and I loses its collapsed flag; R120-X04 E8).
    if (delta > 0) opened = groupsOpenedByInsertion(sheet.outlineGroups, axis, at, delta);
    if (opened.size > 0) {
      patch.outlineGroups = patch.outlineGroups?.map((group, index) => (opened.has(index) ? { ...group, collapsed: false } : group));
    }
  }
  // Floating charts: BOTH the source range AND the anchor cell shift; a chart
  // whose entire source range is deleted stays, its source #REF!. Carried through at OLD
  // coordinates previously = the exact blind spot — a chart reading stale cells
  // (or floating over the wrong region) after the next insert/delete.
  if (sheet.charts) patch.charts = remapCharts(sheet.charts, axis, at, delta);
  // Floating drawing objects (images / shapes): the cell anchor from/to point
  // slides like a chart anchor so the shape stays over the same cells (an
  // absolute, page-pinned anchor is left untouched). Carried through verbatim
  // previously = the blind spot — a shape floating over the WRONG region after an
  // insert/delete (and inconsistent with charts, which already track).
  if (sheet.objects) patch.objects = remapObjects(sheet.objects, axis, at, delta, scope);
  // Index-keyed-by-line maps: a per-line size override / column AutoFilter is
  // keyed by absolute line index, so its key must shift with its line (only the
  // matching axis). Carried through verbatim previously = the same blind spot —
  // the size/filter silently attaching to the wrong line after the next edit
  // (and the filter de-syncing from its slicer, which IS remapped).
  if (sheet.columnWidths)
    patch.columnWidths = remapIndexKeyedMap(sheet.columnWidths, axis === "col", at, delta, MAX_GRID_COLUMNS);
  if (sheet.bestFitColumns && axis === "col") {
    const marks = remapIndexKeyedMap(sheet.bestFitColumns, true, at, delta, MAX_GRID_COLUMNS);
    patch.bestFitColumns =
      marks && Object.keys(marks).length > 0
        ? (marks as Record<string, true>)
        : undefined;
  }
  if (sheet.nativeGrid && axis === "col") {
    patch.nativeGrid = { ...sheet.nativeGrid, columns: remapIndexKeyedMap(sheet.nativeGrid.columns, true, at, delta, MAX_GRID_COLUMNS) ?? {} };
  }
  if (sheet.nativeGrid?.rowMetrics && axis === "row") {
    patch.nativeGrid = {...sheet.nativeGrid,rowMetrics:{...sheet.nativeGrid.rowMetrics,rows:remapIndexKeyedMap(sheet.nativeGrid.rowMetrics.rows,true,at,delta)??{}}};
  }
  if (sheet.rowHeights)
    patch.rowHeights = remapIndexKeyedMap(sheet.rowHeights, axis === "row", at, delta);
  // The rows a file marks thickTop / thickBot (officeRowThickMarkExtras) move
  // with their rows too; an inserted row carries none.
  const thickMarks = (sheet as { rowThickEdges?: unknown }).rowThickEdges;
  if (axis === "row" && thickMarks && typeof thickMarks === "object" && !Array.isArray(thickMarks))
    (patch as Record<string, unknown>).rowThickEdges = remapIndexKeyedMap(thickMarks as Record<string, number>, true, at, delta) ?? {};
  if (sheet.filters)
    patch.filters = remapIndexKeyedMap(sheet.filters, axis === "col", at, delta, MAX_GRID_COLUMNS);
  // Filter-by-Color and Text/Number condition filters are ALSO keyed by column
  // index and feed the same hidden-rows computation; rebasing only `filters` left
  // these stranded at the old column, so after a column op they hid rows for the
  // wrong column (desyncing from the value filter that did move).
  if (sheet.colorFilters)
    patch.colorFilters = remapIndexKeyedMap(
      sheet.colorFilters,
      axis === "col",
      at,
      delta,
      MAX_GRID_COLUMNS,
    );
  if (sheet.conditionFilters)
    patch.conditionFilters = remapIndexKeyedMap(
      sheet.conditionFilters,
      axis === "col",
      at,
      delta,
      MAX_GRID_COLUMNS,
    );
  // A1-string-bound recipes: a named range's `ref` and a pivot's `source` are A1
  // strings that must track the grid like a formula reference (an erased target
  // becomes #REF!, kept not dropped). Stale previously = a name/pivot reading the
  // wrong cells after insert/delete.
  if (sheet.namedRanges)
    patch.namedRanges = remapNamedRanges(sheet.namedRanges, axis, at, delta, scope);
  // Only a PivotTable whose source is on THIS sheet: a report that reads another
  // worksheet keeps its range when its own sheet's grid moves (the edited-source
  // side is remapPivotSourcesOnEditedSheet).
  const ownSource = (pivot: StructuralPivot) =>
    pivotSourceIsOnSheet(pivot, sheet, sheet);
  // Every report on this sheet moves its OUTPUT anchor with the grid (the
  // rendered block sits in these cells, which just moved), and only a
  // report whose source is here moves its source too.
  const moved = (pivot: StructuralPivot): StructuralPivot => {
    const withSource = ownSource(pivot) ? (remapPivot(pivot, axis, at, delta) ?? pivot) : pivot;
    return shiftPivotOutput(withSource, axis, at, delta);
  };
  if (sheet.pivot) patch.pivot = moved(sheet.pivot);
  if (sheet.morePivots) patch.morePivots = sheet.morePivots.map(moved);
  // Print area, print titles and manual page breaks name grid lines; they move
  // together in this pass so the printout keeps the same content per page.
  if (sheet.printSettings) {
    const printSettings = remapPrintSettings(sheet.printSettings, axis, at, delta, scope);
    // Settings left empty (the only print area erased) are no settings,
    // as Clear Print Area leaves them; `{}` read back as none.
    if (printSettings !== sheet.printSettings) {
      patch.printSettings = Object.keys(printSettings).length > 0 ? printSettings : undefined;
    }
  }
  // The worksheet AutoFilter range (the sheet's _FilterDatabase) follows its
  // cells as Excel moves it: a line inserted inside it widens it, one deleted
  // inside it narrows it, lines before it shift it, and an erased range ends
  // the filter. Carried through verbatim previously = the row a user inserted
  // inside the list left its last row outside the exported filter.
  if (typeof sheet.autoFilterRef === "string" && sheet.autoFilterRef) {
    const shifted = shiftA1RangeString(sheet.autoFilterRef, axis, at, delta, scope);
    if (shifted !== sheet.autoFilterRef) {
      patch.autoFilterRef = isErasedA1Range(shifted) ? undefined : shifted;
      if (patch.autoFilterRef === undefined && !sheet.namedSheetViews?.length) patch.autoFilterId = undefined;
    }
    // The filter's header row deleted: the filter goes, criteria and all
    
    // Sheet Views hold their criteria against it and keep it.
    if (
      axis === "row" &&
      delta < 0 &&
      !sheet.namedSheetViews?.length &&
      parseA1Range(sheet.autoFilterRef)?.r1 === at
    ) {
      patch.autoFilterRef = undefined;
      // The file's uid went with it: alone it is a sheet the schema refuses,
      // and the delete was refused with it (R123-X07, B4:AZ5 with xr:uid).
      patch.autoFilterId = undefined;
      patch.filters = undefined;
      patch.colorFilters = undefined;
      patch.conditionFilters = undefined;
      patch.filterResult = undefined;
    }
  }
  if (typeof sheet.databaseRef === "string" && sheet.databaseRef) {
    const shifted = shiftA1RangeString(sheet.databaseRef, axis, at, delta, scope);
    if (shifted !== sheet.databaseRef) {
      patch.databaseRef = isErasedA1Range(shifted) ? undefined : shifted;
    }
  }
  // The recorded sort follows its rows and key columns the same way (Excel
  // keeps a sortState's ref in step with the grid). It was carried verbatim,
  // so a row deleted above it left it naming rows past the sheet, and the
  // sort was dropped on reopening (sheet model fuzz); a deleted key column
  // or an erased range ends it.
  if (sheet.sortState && typeof sheet.sortState.ref === "string") {
    const state = sheet.sortState;
    // A range narrowed to one cell is written as that cell ("A13"), the
    // form the sortState is read back in.
    const single = (range: string) => range.replace(/^(\$?[A-Za-z]+\$?\d+):\1$/, "$1");
    const ref = single(shiftA1RangeString(state.ref, axis, at, delta, scope));
    const conditions = state.conditions.map((condition) => ({
      ...condition,
      ref: single(shiftA1RangeString(condition.ref, axis, at, delta, scope)),
    }));
    if (ref !== state.ref || conditions.some((condition, index) => condition.ref !== state.conditions[index].ref)) {
      patch.sortState = isErasedA1Range(ref) || conditions.some((condition) => isErasedA1Range(condition.ref))
        ? undefined
        : { ref, conditions };
    }
  }
  // Freeze-pane band boundaries move with their own axis so the same rows/cols
  // stay pinned after an insert/delete (carried through verbatim previously =
  // the WRONG lines frozen once a line is added/removed before the boundary).
  // A split is window state, not a band: insertLine/deleteLine settle it once
  // for the whole edit (splitAfterStructuralEdit).
  if (axis === "row") {
    if (sheet.freezeRows !== undefined)
      patch.freezeRows = shiftBoundary(sheet.freezeRows, at, delta);
  } else {
    if (sheet.freezeCols !== undefined)
      patch.freezeCols = shiftBoundary(sheet.freezeCols, at, delta);
  }
  // Scenarios (Scenario Manager): each per-cell A1 `ref` shifts with the grid;
  // a ref ON the deleted line is DROPPED from its scenario (the override has no
  // cell to write once its target is gone). Carried through verbatim previously
  // = the same blind spot — a scenario writing the WRONG cells after an
  // insert/delete that intersects its refs.
  if (sheet.scenarios)
    patch.scenarios = remapScenarios(sheet.scenarios, axis, at, delta);
  // What-If Data Tables move whole with their cells; one the edit cut, or
  // whose input cell it deleted, ends and its cells keep their values.
  if (sheet.whatIfTables) {
    const tables = remapWhatIfTables(sheet.whatIfTables, axis, at, delta);
    patch.whatIfTables = tables.length > 0 ? tables : undefined;
  }
  // Manually hidden rows/cols: absolute line indices that shift like a cell on
  // the matching axis (an index ON the deleted line drops). Stale previously =
  // the wrong row/col staying collapsed after an insert/delete.
  if (sheet.manuallyHiddenRows)
    patch.manuallyHiddenRows = remapHiddenIndices(
      sheet.manuallyHiddenRows,
      axis === "row",
      at,
      delta,
    );
  if (sheet.manuallyHiddenCols)
    patch.manuallyHiddenCols = remapHiddenIndices(
      sheet.manuallyHiddenCols,
      axis === "col",
      at,
      delta,
    );
  // The rows 0 points high (reading 998) move with their rows likewise.
  if (sheet.zeroHeightRows && axis === "row")
    patch.zeroHeightRows = remapHiddenIndices(sheet.zeroHeightRows, true, at, delta);
  // A row still shown past the hidden default (reading 944, row 1,048,576)
  // moves with the rows; one an insert pushes off the grid's end goes, as
  // Excel's last row does.
  // A delete also brings fresh rows in at the grid's end, which such a sheet
  // shows: its file states no hidden default (Windows Excel 16, xlsx-rounds
  
  // 1,048,574 and rows 1,048,575-1,048,576 came in shown).
  // A sheet hidden by its file's run to the grid's end (reading 945, no
  // zeroHeight) shows them likewise (probe: rows 2-3 deleted from a run to
  // 1,048,576, rows 1,048,575-1,048,576 came in shown). An insert keeps the
  // run hidden to the end there too.
  if (
    axis === "row" &&
    ((sheet.rowsShownPastEnd && sheet.rowsShownPastEnd.length > 0) ||
      (sheet.rowsHiddenByFileRun === true && delta < 0))
  ) {
    const moved = (remapHiddenIndices(sheet.rowsShownPastEnd ?? [], true, at, delta) ?? []).filter((row) => row < MAX_ROWS);
    for (let row = Math.max(0, MAX_ROWS + delta); delta < 0 && row < MAX_ROWS; row += 1) moved.push(row);
    const shown = [...new Set(moved)].sort((left, right) => left - right);
    patch.rowsShownPastEnd = shown.length > 0 ? shown : undefined;
  }
  if (opened.size > 0) {
    const stayHidden = new Set<number>(axis === "row" ? patch.manuallyHiddenRows ?? [] : patch.manuallyHiddenCols ?? []);
    // A line a group inside it still hides (collapsed, and not opened by this
    // insert) stays that group's: the file states it hidden under a collapsed
    // summary, which reads back as the outline's, not hidden by hand
    // (xlsx-rounds R462-X38 E4: rows 207:208 and 211:212 under their own
    // collapsed groups inside the opened 205:223).
    const stillCollapsed = (patch.outlineGroups ?? []).filter(
      (group, index) => group.axis === axis && (group as { collapsed?: boolean }).collapsed === true && !opened.has(index),
    );
    for (const index of opened) {
      const group = sheet.outlineGroups![index]!;
      for (let line = group.start; line <= group.end; line += 1) {
        const moved = line >= at ? line + delta : line;
        if (!stillCollapsed.some((inner) => inner.start <= moved && moved <= inner.end)) stayHidden.add(moved);
      }
    }
    const sorted = [...stayHidden].sort((a, b) => a - b);
    if (axis === "row") patch.manuallyHiddenRows = sorted;
    else patch.manuallyHiddenCols = sorted;
  }
  // The rows the filters hid when last applied move with their rows; an
  // inserted row is shown, a deleted one goes (Excel filters nothing again).
  if (sheet.filterResult !== undefined && axis === "row" && !("filterResult" in patch))
    patch.filterResult = remappedFilterResult(sheet.filterResult, (row) => shiftCoord(row, at, delta));
  // A column edit that only moves the criteria's columns is the same filter:
  
  // again. One that takes a criterion's column leaves it to be applied again.
  if (sheet.filterResult !== undefined && axis === "col" && !("filterResult" in patch)) {
    const carried = filterResultMovedWithColumns(sheet as never, { ...sheet, ...patch } as never);
    if (carried) patch.filterResult = carried;
  }
  if (sheet.advancedFilterState) {
    const state = sheet.advancedFilterState;
    const list = parseA1Range(state.ref);
    const ref = shiftA1RangeString(state.ref, axis, at, delta, scope);
    const shifted = isErasedA1Range(ref) ? null : parseA1Range(ref);
    const headerDeleted =
      axis === "row" && delta < 0 && list !== null && at === list.r1;
    if (!shifted || headerDeleted || shifted.r2 <= shifted.r1) {
      patch.advancedFilterState = undefined;
    } else {
      const hiddenRows = (
        remapHiddenIndices(state.hiddenRows, axis === "row", at, delta) ?? []
      )
        .filter((row) => row > shifted.r1 && row <= shifted.r2)
        .sort((left, right) => left - right);
      const criteria =
        state.criteria === undefined
          ? undefined
          : shiftA1RangeString(state.criteria, axis, at, delta, scope);
      patch.advancedFilterState = {
        ref,
        hiddenRows,
        ...(criteria !== undefined && !isErasedA1Range(criteria)
          ? { criteria }
          : {}),
      };
    }
  }
  // Row / column formats are keyed by line index like the sizes, and move
  // with them (B029): a row format follows its row through an insert or a
  // delete above it and drops with a deleted row.
  if (sheet.rowStyles) {
    const rows = remapIndexKeyedMap(sheet.rowStyles, axis === "row", at, delta);
    patch.rowStyles = rows && Object.keys(rows).length > 0 ? rows : undefined;
  }
  if (sheet.columnStyles) {
    // A column format the page stores as a run ("8:16383") moves column by column.
    const cols = remapIndexKeyedMap(
      axis === "col" ? expandLineStyleRuns(sheet.columnStyles as LineStyleMap) : sheet.columnStyles,
      axis === "col", at, delta, MAX_GRID_COLUMNS,
    );
    patch.columnStyles = cols && Object.keys(cols).length > 0 ? cols : undefined;
  }
  // The blank-cell ranges a sheet was loaded with (a symbol-keyed save hint,
  // see blank-cell-formats.ts) follow the cells they describe, so ranges the
  // edit did not touch persist unchanged instead of being rebuilt.
  const blankCellFormats = (sheet as BlankCellFormatsHintCarrier)[
    BLANK_CELL_FORMATS_HINT
  ];
  if (blankCellFormats) {
    (patch as BlankCellFormatsHintCarrier)[BLANK_CELL_FORMATS_HINT] =
      shiftBlankCellFormatsHint(blankCellFormats, axis, at, delta > 0 ? 1 : -1);
  }
  return patch as Partial<T>;
}

function structuralTableFormulasAreOwned(sheet: StructuralSheet): boolean {
  for (const table of sheet.dataTables ?? []) {
    if (table.totalsRow && table.totalFormulas) {
      const totalsRow = table.range.r2 + 1;
      for (const [rawOffset, rawFormula] of Object.entries(table.totalFormulas)) {
        const offset = Number(rawOffset);
        if (!Number.isInteger(offset) || typeof rawFormula !== "string") return false;
        const cell = sheet.cells[`${totalsRow}:${table.range.c1 + offset}`];
        const physical = cell?.formula ?? cell?.value;
        if (physical !== rawFormula) return false;
      }
    }
    if (table.calculatedFormulas) {
      for (const [rawOffset, rawFormula] of Object.entries(
        table.calculatedFormulas,
      )) {
        const offset = Number(rawOffset);
        if (!Number.isInteger(offset) || typeof rawFormula !== "string") return false;
      }
    }
  }
  return true;
}

/**
 * The calculated-column cells that do not hold their column formula: Excel's
 * calculated-column exceptions (a typed constant, another formula, a blank).
 * They move with the rows and columns around them and are never rewritten to
 * the column formula, so an insert or delete keeps them as they are
 * (Microsoft Support, "Use calculated columns in an Excel table"). Keyed like
 * `cells`, so `remapCells` moves them exactly as it moves the cells.
 */
function structuralCalculatedExceptions(
  sheet: StructuralSheet,
): Record<string, StructuralCell> {
  const exceptions: Record<string, StructuralCell> = {};
  for (const table of sheet.dataTables ?? []) {
    if (!table.calculatedFormulas) continue;
    const firstDataRow = table.range.r1 + (table.headerRow ? 1 : 0);
    for (const [rawOffset, rawFormula] of Object.entries(
      table.calculatedFormulas,
    )) {
      const offset = Number(rawOffset);
      if (!Number.isInteger(offset) || typeof rawFormula !== "string") continue;
      for (let row = firstDataRow; row <= table.range.r2; row += 1) {
        const key = `${row}:${table.range.c1 + offset}`;
        const cell = sheet.cells[key];
        if (
          (cell?.formula ?? cell?.value) !==
          shiftFormula(rawFormula, row - firstDataRow, 0)
        ) {
          exceptions[key] = { value: "" };
        }
      }
    }
  }
  return exceptions;
}

function structuralTableColumnFormulaRemapSupported(
  sheet: StructuralSheet,
  at: number,
): boolean {
  return (sheet.dataTables ?? []).every((table) => {
    if (
      (!table.totalFormulas && !table.calculatedFormulas) ||
      at < table.range.c1 ||
      at > table.range.c2
    ) {
      return true;
    }
    return (
      Array.isArray(table.columns) &&
      table.columns.length === table.range.c2 - table.range.c1 + 1
    );
  });
}

function syncStructuralTableFormulaCells<C extends StructuralCell>(
  cells: Record<string, C>,
  tables: StructuralDataTable[] | undefined,
  exceptions: Record<string, StructuralCell> = {},
): Record<string, C> {
  let next = cells;
  for (const table of tables ?? []) {
    if (table.totalsRow && table.totalFormulas) {
      const totalsRow = table.range.r2 + 1;
      for (const [rawOffset, rawFormula] of Object.entries(table.totalFormulas)) {
        const offset = Number(rawOffset);
        if (!Number.isInteger(offset) || typeof rawFormula !== "string") continue;
        const key = `${totalsRow}:${table.range.c1 + offset}`;
        const current = next[key];
        if (!current || (current.formula === rawFormula && current.value === rawFormula)) {
          continue;
        }
        if (next === cells) next = { ...cells };
        next[key] = { ...current, formula: rawFormula, value: rawFormula };
      }
    }
    if (table.calculatedFormulas) {
      const firstDataRow = table.range.r1 + (table.headerRow ? 1 : 0);
      for (const [rawOffset, rawFormula] of Object.entries(
        table.calculatedFormulas,
      )) {
        const offset = Number(rawOffset);
        if (!Number.isInteger(offset) || typeof rawFormula !== "string") continue;
        for (let row = firstDataRow; row <= table.range.r2; row += 1) {
          const formula = shiftFormula(rawFormula, row - firstDataRow, 0);
          const key = `${row}:${table.range.c1 + offset}`;
          if (Object.prototype.hasOwnProperty.call(exceptions, key)) continue;
          const current = next[key];
          if (current?.formula === formula && current.value === formula) {
            continue;
          }
          // A row that was there holds its own formula, moved by the edit as
          // every cell's is; the recipe fills what the edit brought.
          if (current?.formula) continue;
          if (next === cells) next = { ...cells };
          // A new cell is written whole, as a typed formula is (its kind
          // is "text"); the loader would otherwise add the kind on reopening.
          next[key] = current
            ? { ...current, formula, value: formula }
            : ({ formula, value: formula, kind: "text" as const } as C);
        }
      }
    }
  }
  return next;
}

/**
 * A row inserted into a table takes its calculated columns' formulas (Microsoft
 * Support, "Use calculated columns in an Excel table"; xlsx-rounds xprobe23). An
 * array calculated column has no recipe on the page, so the new cell takes the
 * neighbouring row's formula, as an array of its own.
 */
function fillArrayCalculatedCells<C extends StructuralCell>(
  cells: Record<string, C>,
  tables: StructuralDataTable[] | undefined,
  at: number,
): Record<string, C> {
  let next = cells;
  for (const table of tables ?? []) {
    if (!table.arrayCalculatedColumns) continue;
    const firstDataRow = table.range.r1 + (table.headerRow ? 1 : 0);
    if (at < firstDataRow || at > table.range.r2) continue;
    const neighbour = at > firstDataRow ? at - 1 : at + 1;
    for (const rawOffset of Object.keys(table.arrayCalculatedColumns)) {
      const col = table.range.c1 + Number(rawOffset);
      if (!Number.isInteger(col)) continue;
      const source = next[`${neighbour}:${col}`] as
        | (StructuralCell & { spillRef?: unknown; arrayRef?: unknown })
        | undefined;
      const mark = typeof source?.spillRef === "string" ? "spillRef" : "arrayRef";
      const ref = source?.[mark];
      const single = typeof ref === "string" ? /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref) : null;
      if (!source?.formula || !single || Number(single[2]) !== neighbour + 1) continue;
      const formula = shiftFormula(source.formula, at - neighbour, 0);
      const key = `${at}:${col}`;
      if (next === cells) next = { ...cells };
      next[key] = {
        ...next[key],
        formula,
        value: formula,
        kind: "text",
        [mark]: `${single[1]}${at + 1}`,
      } as C;
    }
  }
  return next;
}





function withCellsPastTheEndDropped<T extends StructuralSheet>(sheet: T, axis: StructuralAxis, ceiling: number): T {
  let cells: Record<string, StructuralCell> | null = null;
  for (const key of Object.keys(sheet.cells)) {
    const separator = key.indexOf(":");
    if (separator <= 0) continue;
    const line = Number(axis === "row" ? key.slice(0, separator) : key.slice(separator + 1));
    if (!Number.isInteger(line) || line < ceiling) continue;
    cells ??= { ...sheet.cells };
    delete cells[key];
  }
  return cells ? { ...sheet, cells } : sheet;
}

/**
 * Insert `count` blank rows/columns at index `at`, shifting later lines outward
 * (Excel/WPS "Insert N rows" from an N-line selection). `count` defaults to 1
 * (the single-line op); a count < 1 is a no-op. The per-remap INSERT branch
 * encodes single-line semantics (a coord at/after `at` shifts by exactly one),
 * so — symmetric with deleteLine — an N-line insert is `count` successive
 * single-line inserts at the SAME index `at`: each pass pushes the blank already
 * at `at` (and everything after it) outward by one, leaving N blank lines in
 * `[at, at+N-1]` and the original `at`-content at `at+N`. This reuses the fully
 * tested 1-line remap for cells, formulas, merges, ranges, outline groups,
 * charts and every overlay without duplicating its edge cases. An insert that
 * {@link structuralInsertRefusal} refuses (data would leave the worksheet, or
 * the line budget) returns the sheet unchanged; callers ask that function
 * first to tell the user why. Otherwise all `count` lines are inserted and the
 * extent grows, stopping at the last row/column (only the blank tail drops).
 */
export function insertLine<T extends StructuralSheet>(
  sheet: T,
  axis: StructuralAxis,
  at: number,
  count = 1,
): T {
  const n = Math.trunc(count);
  if (n < 1) return sheet;
  if (!structuralTableFormulasAreOwned(sheet)) return sheet;
  if (structuralInsertRefusal(sheet, axis, n) !== null) return sheet;
  if (structuralLineArrayRefusal(sheet, axis, at, n, "insert") !== null) return sheet;
  const ceiling = axis === "row" ? MAX_ROWS : MAX_COLS;
  const passes = n;
  let next: T = sheet;
  for (let i = 0; i < passes; i += 1) {
    let cells = remapCells(next.cells, axis, at, 1, structuralScopeOf(next));
    const overlays = remapOverlays(next, axis, at, 1);
    cells = syncStructuralTableFormulaCells(
      cells,
      overlays.dataTables as StructuralDataTable[] | undefined,
      remapCells(structuralCalculatedExceptions(next), axis, at, 1, structuralScopeOf(next)),
    );
    if (axis === "row") {
      cells = fillArrayCalculatedCells(
        cells,
        overlays.dataTables as StructuralDataTable[] | undefined,
        at,
      );
    }
    if (axis === "col" && next.dataTables) {
      const nextTables = overlays.dataTables as
        | StructuralDataTable[]
        | undefined;
      next.dataTables.forEach((table, tableIndex) => {
        if (
          table.headerRow === false ||
          at <= table.range.c1 ||
          at > table.range.c2
        ) {
          return;
        }
        const nextTable = nextTables?.[tableIndex];
        const insertedOffset = nextTable ? at - nextTable.range.c1 : -1;
        const columnName = nextTable?.columns?.[insertedOffset]?.name;
        if (!nextTable || !columnName) return;
        cells = {
          ...cells,
          [`${nextTable.range.r1}:${at}`]: { value: columnName, kind: "text" as const },
        };
      });
    }
    next =
      axis === "row"
        ? {
            ...next,
            ...overlays,
            rowCount: Math.min(next.rowCount + 1, ceiling),
            cells,
          }
        : {
            ...next,
            ...overlays,
            columnCount: Math.min(next.columnCount + 1, ceiling),
            cells,
          };
  }
  next = withTurnedObjectsMovedByBox(sheet, next, axis, at, n);
  next = withCellsPastTheEndDropped(next, axis, ceiling);
  next = withAreaLinksAcrossInsertedLines(next, axis, at, n);
  return withLineMoveRecorded(sheet, splitAfterStructuralEdit(
    sheet as unknown as SplitStructureSheet,
    next as unknown as SplitStructureSheet,
    axis,
  ) as unknown as T, axis, at, n);
}








function withAreaLinksAcrossInsertedLines<T extends StructuralSheet>(
  sheet: T,
  axis: StructuralAxis,
  at: number,
  count: number,
): T {
  if (at < 1) return sheet;
  type Layer = { hyperlink?: unknown; hyperlinkTooltip?: unknown };
  type Linked = StructuralCell & Layer & { hyperlinkInArea?: boolean; hyperlinkArea?: Layer };
  const cells = sheet.cells as Record<string, Linked>;
  // The area's link a cell holds: its own (hyperlinkInArea), or the one
  // beneath a link of the cell's own (hyperlinkArea). The area's element
  // grows over the inserted lines either way -- Excel moves its ref -- so a
  // cell above or below linked on its own does not stop it (xlsx-rounds
  // R450-X65 E4: B6 linked on its own inside B3:B11, a row inserted at 7).
  const areaLink = (cell: Linked | undefined): Layer | null =>
    cell?.hyperlinkInArea === true
      ? cell
      : cell?.hyperlinkArea && typeof cell.hyperlinkArea.hyperlink === "string"
        ? cell.hyperlinkArea
        : null;
  let grown: Record<string, Linked> | null = null;
  for (const key of Object.keys(cells)) {
    const before = areaLink(cells[key]);
    if (!before) continue;
    const parsed = parseCellKey(key);
    if (!parsed || (axis === "row" ? parsed.row : parsed.col) !== at - 1) continue;
    const place = (line: number) => (axis === "row" ? `${line}:${parsed.col}` : `${parsed.row}:${line}`);
    const after = areaLink(cells[place(at + count)]);
    if (!after || after.hyperlink !== before.hyperlink) continue;
    for (let line = at; line < at + count; line += 1) {
      const target = place(line);
      (grown ??= { ...cells })[target] = {
        ...(cells[target] ?? { value: "", kind: "text" }),
        hyperlink: before.hyperlink,
        ...(before.hyperlinkTooltip === undefined ? {} : { hyperlinkTooltip: before.hyperlinkTooltip }),
        hyperlinkInArea: true,
      };
    }
  }
  return grown ? { ...sheet, cells: grown } : sheet;
}

/**
 * Excel's default Insert Option, "Format Same As Above" (rows) / "Format Same
 * As Left" (columns): the `count` lines just inserted at `at` take the cell
 * FORMATS of the line before them, never its values, so a row inserted inside
 * a bordered, number-formatted table looks like the table instead of a gap.
 * An insert at the first line has nothing before it and stays unformatted. A
 * cell the insert itself filled (a table's new column header) keeps its value.
 * `span` limits the copy to part of the line (inclusive columns for a row,
 * rows for a column), as when a table grows by one line of its own.
 */
function withoutMergeFarEdge(
  style: unknown,
  sheet: StructuralSheet,
  row: number,
  col: number,
  axis: StructuralAxis,
): unknown {
  if (!style || typeof style !== "object") return style;
  const merge = (sheet.merges ?? []).find((m) => m.row === row && m.col === col);
  if (!merge) return style;
  // A column insert copies along the row: the anchor's bottom belongs to the
  // area's last row. A row insert copies along the column: its right belongs
  // to the area's last column.
  const far = axis === "col" ? (merge.rowSpan > 1 ? "bottom" : null) : merge.colSpan > 1 ? "right" : null;
  if (!far) return style;
  const next = { ...(style as Record<string, unknown>) };
  const flag = far === "bottom" ? "borderBottom" : "borderRight";
  if (next.border === true) {
    delete next.border;
    for (const side of ["Top", "Right", "Bottom", "Left"]) {
      if (`border${side}` !== flag) next[`border${side}`] = true;
    }
  }
  delete next[flag];
  const edges = next.borderEdges;
  if (edges && typeof edges === "object" && far in (edges as Record<string, unknown>)) {
    const rest = { ...(edges as Record<string, unknown>) };
    delete rest[far];
    if (Object.keys(rest).length > 0) next.borderEdges = rest;
    else delete next.borderEdges;
  }
  // A far line the anchor's own record does state, hidden by the area's outline
  
  const own = (merge as { coveredStyles?: Record<string, { borderEdges?: Record<string, unknown> }> }).coveredStyles?.["0:0"]
    ?.borderEdges?.[far];
  if (own) {
    next.borderEdges = { ...((next.borderEdges as Record<string, unknown> | undefined) ?? {}), [far]: own };
    next[flag] = true;
  }
  return next;
}

/** An index-keyed map (line formats, row sizes) brought back within a cap
 *  by dropping its highest lines. */
function withinCap<V>(map: Record<string | number, V>, cap: number): Record<string | number, V> {
  const keys = Object.keys(map);
  if (keys.length <= cap) return map;
  const numeric = keys.filter((key) => Number.isInteger(Number(key))).sort((a, b) => Number(b) - Number(a));
  const out = { ...map };
  for (const key of numeric.slice(0, keys.length - cap)) delete out[key];
  return out;
}

export function inheritInsertedLineFormats<T extends StructuralSheet>(
  sheet: T,
  axis: StructuralAxis,
  at: number,
  count = 1,
  span?: { from: number; to: number },
): T {
  const n = Math.trunc(count);
  if (n < 1 || at < 1) return sheet;
  const source = at - 1;
  let cells: Record<string, StructuralCell> | null = null;
  for (const key of Object.keys(sheet.cells)) {
    const separator = key.indexOf(":");
    if (separator <= 0) continue;
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if ((axis === "row" ? row : col) !== source) continue;
    const cross = axis === "row" ? col : row;
    if (span && (cross < span.from || cross > span.to)) continue;
    const sourceCell = sheet.cells[key] as { style?: unknown; locked?: unknown; formulaHidden?: unknown };
    // A merge anchor's page style folds the area's outline in (its bottom
    // and right lines come from the area's far cells); Excel's Insert copies
    // the anchor's OWN record, which does not draw the far edge.
    const style = recordFontDrawn(withoutMergeFarEdge(sourceCell.style, sheet, row, col, axis));
    const hasStyle =
      style !== undefined && style !== null && typeof style === "object" && Object.keys(style).length > 0;
    // Protection (Format Cells > Protection) is part of the format Excel's
    // Insert copies from the neighbour: an inserted row in an unlocked input
    // block stays unlocked.
    const protection = {
      ...(sourceCell.locked === false ? { locked: false } : {}),
      ...(sourceCell.formulaHidden === true ? { formulaHidden: true } : {}),
    };
    if (!hasStyle && Object.keys(protection).length === 0) continue;
    for (let offset = 0; offset < n; offset += 1) {
      const target =
        axis === "row" ? `${at + offset}:${col}` : `${row}:${at + offset}`;
      const existing = (cells ?? sheet.cells)[target] as
        | (StructuralCell & { style?: unknown; locked?: unknown; formulaHidden?: unknown })
        | undefined;
      if (existing?.style !== undefined || existing?.locked !== undefined || existing?.formulaHidden !== undefined) {
        continue;
      }
      cells ??= { ...sheet.cells };
      cells[target] = {
        value: "",
        kind: "text" as const,
        ...existing,
        ...(hasStyle ? { style } : {}),
        ...protection,
      } as StructuralCell;
    }
  }
  // The line's own format (rowStyles / columnStyles) is inherited the same
  // way: an inserted row formatted like the row above it keeps that format
  // on the columns added later, as the row above does (B029).
  const field = axis === "row" ? "rowStyles" : "columnStyles";
  const lineStyles = axis === "col" ? expandLineStyleRuns(sheet[field] as LineStyleMap | undefined) : sheet[field];
  const sourceStyle = lineStyles?.[String(source)];
  let nextLineStyles: Record<string | number, unknown> | null = null;
  if (sourceStyle !== undefined && sourceStyle !== null && typeof sourceStyle === "object") {
    for (let offset = 0; offset < n; offset += 1) {
      const key = String(at + offset);
      if ((nextLineStyles ?? lineStyles)?.[key] !== undefined) continue;
      nextLineStyles ??= { ...lineStyles };
      nextLineStyles[key] = JSON.parse(JSON.stringify(sourceStyle));
    }
  }
  // The line's size too: Range.Insert's default CopyOrigin
  // (xlFormatFromLeftOrAbove) gives an inserted row the height of the row
  // above and an inserted column the width of the one to its left. A table's
  // own row (`span`) shifts cells, not a whole line, so no size moves.
  const sizes = span ? null : inheritedLineSizes(sheet, axis, source, at, n);
  // A page holds at most so many line formats / row sizes (the page schema's
  // caps; an import keeps the first ones). The inserted line's own entry
  // must not push the map past the cap -- the save was refused whole
  // (xlsx-rounds R16-X25: a column format on all 16,141 columns, a column
  // inserted): the entry the shift pushed furthest goes, as one past the
  // import's horizon already did; the file keeps its own <col> spans.
  if (nextLineStyles) nextLineStyles = withinCap(nextLineStyles, axis === "col" ? MAX_COLUMN_LINE_STYLES : MAX_LINE_STYLES_PER_AXIS);
  if (sizes && (sizes as { rowHeights?: Record<string, number> }).rowHeights) {
    const capped = sizes as { rowHeights?: Record<string, number> };
    capped.rowHeights = withinCap(capped.rowHeights ?? {}, MAX_GRID_ROW_SIZE_OVERRIDES) as Record<string, number>;
  }
  if (!cells && !nextLineStyles && !sizes) return sheet;
  return {
    ...sheet,
    ...(cells ? { cells } : {}),
    ...(nextLineStyles ? { [field]: nextLineStyles } : {}),
    ...(sizes ?? {}),
  };
}

/** A size map with `source`'s entry copied onto the `count` lines from `at`
 * that state none, or null when nothing changes. */
function withSourceSize<V>(
  map: Record<string | number, V> | undefined,
  source: number,
  at: number,
  count: number,
): Record<string, V> | null {
  const size = map?.[String(source)];
  if (!map || size === undefined || size === null) return null;
  let next: Record<string, V> | null = null;
  for (let offset = 0; offset < count; offset += 1) {
    const key = String(at + offset);
    if ((next ?? map)[key] !== undefined) continue;
    next ??= { ...(map as Record<string, V>) };
    next[key] = typeof size === "object" ? (JSON.parse(JSON.stringify(size)) as V) : size;
  }
  return next;
}

function inheritedLineSizes<T extends StructuralSheet>(
  sheet: T,
  axis: StructuralAxis,
  source: number,
  at: number,
  count: number,
): Partial<T> | null {
  const out: Record<string, unknown> = {};
  if (axis === "row") {
    const heights = withSourceSize(sheet.rowHeights, source, at, count);
    if (heights) out.rowHeights = heights;
    const metrics = sheet.nativeGrid?.rowMetrics;
    const rows = metrics ? withSourceSize(metrics.rows, source, at, count) : null;
    if (rows && sheet.nativeGrid && metrics) {
      out.nativeGrid = { ...sheet.nativeGrid, rowMetrics: { ...metrics, rows } };
    }
  } else {
    const widths = withSourceSize(sheet.columnWidths, source, at, count);
    if (widths) out.columnWidths = widths;
    const columns = sheet.nativeGrid ? withSourceSize(sheet.nativeGrid.columns, source, at, count) : null;
    if (columns && sheet.nativeGrid) out.nativeGrid = { ...sheet.nativeGrid, columns };
  }
  return Object.keys(out).length ? (out as Partial<T>) : null;
}

/**
 * Delete `count` rows/columns starting at index `at`, shifting later lines
 * inward (Excel/WPS "Delete N rows" from an N-line selection). `count` defaults
 * to 1; a count < 1 is a no-op. Unlike insert, the delete branch of every remap
 * encodes single-line semantics (a coord ON `at` drops; a coord past it moves up
 * by one), so an N-line delete is `count` successive single-line deletes at the
 * SAME index `at` — each pass removes the line that has shifted into `at` — which
 * reuses the fully-tested 1-line remap for cells, formulas, merges, ranges,
 * outline groups, charts and every other overlay without duplicating its edge
 * cases. Only the lines inside the persisted extent are deleted
 * ({@link deletableLineCount}): the blank lines the grid shows past the extent
 * are not stored, so deleting them is a no-op and a selection running past the
 * extent deletes its in-extent part. The extent never shrinks below one line
 * or below the last stored cell.
 */
export function deleteLine<T extends StructuralSheet>(
  sheet: T,
  axis: StructuralAxis,
  at: number,
  count = 1,
): T {
  const n = deletableLineCount(sheet, axis, at, count);
  if (n < 1) return sheet;
  if (structuralLineArrayRefusal(sheet, axis, at, n, "delete") !== null) return sheet;
  if (
    !structuralTableFormulasAreOwned(sheet) ||
    (axis === "col" && !structuralTableColumnFormulaRemapSupported(sheet, at))
  ) {
    return sheet;
  }
  let next: T = sheet;
  for (let i = 0; i < n; i += 1) {
    let cells = remapCells(next.cells, axis, at, -1, structuralScopeOf(next));
    const overlays = remapOverlays(next, axis, at, -1);
    cells = syncStructuralTableFormulaCells(
      cells,
      overlays.dataTables as StructuralDataTable[] | undefined,
      remapCells(structuralCalculatedExceptions(next), axis, at, -1, structuralScopeOf(next)),
    );
    next =
      axis === "row"
        ? {
            ...next,
            ...overlays,
            rowCount: at < next.rowCount ? Math.max(next.rowCount - 1, 1) : next.rowCount,
            cells,
          }
        : {
            ...next,
            ...overlays,
            columnCount:
              at < next.columnCount ? Math.max(next.columnCount - 1, 1) : next.columnCount,
            cells,
          };
  }
  next = withTurnedObjectsMovedByBox(sheet, next, axis, at, -n);
  // A stored cell past the extent (older content) keeps the extent covering
  // it, and so does a floating object: an anchor on the deleted last line
  // clamps to that line (remapObjects), which the shrunken extent no longer
  // holds — and the page validator refuses an anchor outside the grid.
  const extent = axis === "row" ? next.rowCount : next.columnCount;
  const lastCell = Math.max(
    lastStoredLine(next.cells, axis),
    lastAnchoredLine(next.objects, axis),
    lastAnchoredLine(next.charts?.map((chart) => ({ anchor: chart.drawingAnchor })), axis),
  );
  if (lastCell >= extent) {
    next =
      axis === "row"
        ? { ...next, rowCount: lastCell + 1 }
        : { ...next, columnCount: lastCell + 1 };
  }
  // A split window keeps its bars and panes, and the extent holds them.
  return withLineMoveRecorded(sheet, splitAfterStructuralEdit(
    sheet as unknown as SplitStructureSheet,
    next as unknown as SplitStructureSheet,
    axis,
  ) as unknown as T, axis, at, -n);
}

/** The last grid line a cell-anchored object's from/to point sits on, -1 when none. */
function lastAnchoredLine(objects: StructuralObject[] | undefined, axis: StructuralAxis): number {
  let last = -1;
  for (const object of objects ?? []) {
    const anchor = object.anchor;
    // A SmartArt member's graphic frame holds cells of its own.
    const frame = (anchor as { smartArtFrame?: StructuralObjectAnchor } | undefined)?.smartArtFrame;
    if (frame && typeof frame === "object") last = Math.max(last, lastAnchoredLine([{ anchor: frame }], axis));
    if (!anchor || anchor.kind === "absolute") continue;
    for (const point of [anchor.from, anchor.to]) {
      if (!point) continue;
      const line = axis === "row" ? point.row : (point.column ?? point.col);
      if (typeof line === "number" && Number.isInteger(line) && line > last) last = line;
    }
  }
  return last;
}

// ---------------------------------------------------------------------------
// Co-authoring: one whole-line edit as a coordinate transform
//
// A collaborator's Insert/Delete rows or columns reaches the other windows as
// moved cells. These pure maps let a receiver move what is its own -- the
// active cell, the selection, the open entry, and cell writes it made that
// the author had not seen yet -- the way the edit moved everything else, so a
// value typed beside a label stays beside it (Excel co-authoring).
// ---------------------------------------------------------------------------

/** One whole-line edit as {@link insertLine} / {@link deleteLine} apply it:
 * `count` lines at index `at` on `axis`. A delete's count is the number of
 * lines actually deleted ({@link deletableLineCount}). */
export interface SheetLineEdit {
  axis: StructuralAxis;
  op: "insert" | "delete";
  at: number;
  count: number;
}

/** Where line `line` is after `edit`; null when the edit deleted it. */
export function mapLineThroughLineEdit(line: number, edit: SheetLineEdit): number | null {
  const n = Math.max(0, Math.trunc(edit.count));
  if (n === 0 || line < edit.at) return line;
  if (edit.op === "insert") return line + n;
  return line < edit.at + n ? null : line - n;
}

/** The line, before `edit`, whose content is at `line` after it; null for a
 * line the edit inserted (it holds nothing that existed before). */
export function unmapLineThroughLineEdit(line: number, edit: SheetLineEdit): number | null {
  const n = Math.max(0, Math.trunc(edit.count));
  if (n === 0 || line < edit.at) return line;
  if (edit.op === "delete") return line + n;
  return line < edit.at + n ? null : line - n;
}

/** A cell's address after `edit`; null when its row/column was deleted. */
export function mapCellThroughLineEdit(
  cell: { row: number; col: number },
  edit: SheetLineEdit,
): { row: number; col: number } | null {
  const mapped = mapLineThroughLineEdit(edit.axis === "row" ? cell.row : cell.col, edit);
  if (mapped === null) return null;
  return edit.axis === "row" ? { row: mapped, col: cell.col } : { row: cell.row, col: mapped };
}

function parseCellKey(key: string): { row: number; col: number } | null {
  const separator = key.indexOf(":");
  if (separator <= 0) return null;
  const row = Number(key.slice(0, separator));
  const col = Number(key.slice(separator + 1));
  return Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0 ? { row, col } : null;
}

/** A "row:col" cell key after `edit`; null when deleted. Other keys pass. */
export function mapCellKeyThroughLineEdit(key: string, edit: SheetLineEdit): string | null {
  const cell = parseCellKey(key);
  if (!cell) return key;
  const mapped = mapCellThroughLineEdit(cell, edit);
  return mapped ? `${mapped.row}:${mapped.col}` : null;
}

/** The "row:col" key, before `edit`, whose content `key` holds after it. */
export function unmapCellKeyThroughLineEdit(key: string, edit: SheetLineEdit): string | null {
  const cell = parseCellKey(key);
  if (!cell) return key;
  const line = unmapLineThroughLineEdit(edit.axis === "row" ? cell.row : cell.col, edit);
  if (line === null) return null;
  return edit.axis === "row" ? `${line}:${cell.col}` : `${cell.row}:${line}`;
}

/** A cell entry as it reads after `edit` when it existed before it: a formula
 * keeps naming the cells it named (references across the edited lines shift,
 * references into deleted lines become #REF!), exactly as {@link insertLine}
 * and {@link deleteLine} rewrite every stored cell. */
export function remapCellEntryThroughLineEdit<C extends StructuralCell>(
  cell: C,
  edit: SheetLineEdit,
  sheetName?: string,
): C {
  const n = Math.max(0, Math.trunc(edit.count));
  const scope = sheetName ? { sheetName } : undefined;
  let next = cell;
  for (let pass = 0; pass < n; pass += 1) {
    next = rewriteFormula(next, edit.axis, edit.at, edit.op === "insert" ? 1 : -1, scope);
  }
  return next;
}

/**
 * The cells a writer re-issues after it receives a collaborator's line edit
 * that was made without some of its own cell writes: `writes` maps a pre-edit
 * "row:col" address to the entry the writer stored there (null: it cleared
 * the cell), in write order.
 *
 * The author moved every cell it knew. A write it did not know stayed at its
 * old address, where the key-level merge either kept it or let the author's
 * moved value win. The writer re-issues each such entry at its transformed
 * address (a formula rewritten as the edit rewrote every other cell) and, where
 * its own entry is what the old address still holds, puts back what the edit
 * moved there: the pre-edit content of {@link unmapCellKeyThroughLineEdit},
 * rewritten likewise, or nothing for an inserted line. An entry in a deleted
 * line goes with its line. An old address someone else has written since is
 * left alone. Only cells that differ from `after` are returned, so applying
 * the result again changes nothing.
 *
 * `before` / `after` are the writer's cells just before and just after the
 * edit arrived.
 */
export function concurrentCellWriteCorrection<C extends StructuralCell>(args: {
  before: Readonly<Record<string, C>>;
  after: Readonly<Record<string, C>>;
  writes: ReadonlyMap<string, C | null>;
  edit: SheetLineEdit;
  sheetName?: string;
  equals?: (left: C | null, right: C | null) => boolean;
}): Record<string, C | null> {
  const { before, after, writes, edit, sheetName } = args;
  const equals = args.equals ?? ((left: C | null, right: C | null) => JSON.stringify(left) === JSON.stringify(right));
  const read = (cells: Readonly<Record<string, C>>, key: string): C | null =>
    Object.hasOwn(cells, key) ? cells[key] : null;
  const rewrite = (value: C | null) =>
    value === null ? null : remapCellEntryThroughLineEdit(value, edit, sheetName);
  const want = new Map<string, C | null>();
  const targets = new Set<string>();
  for (const key of writes.keys()) {
    const target = mapCellKeyThroughLineEdit(key, edit);
    if (target !== null) targets.add(target);
  }
  for (const [key, value] of writes) {
    const target = mapCellKeyThroughLineEdit(key, edit);
    const stillMine = equals(read(after, key), value);
    if (target !== null) {
      if (target !== key) want.set(target, rewrite(value));
      else if (stillMine) want.set(key, rewrite(value));
    }
    if (target === key || targets.has(key) || !stillMine || value === null) continue;
    const source = unmapCellKeyThroughLineEdit(key, edit);
    want.set(key, source === null ? null : rewrite(read(before, source)));
  }
  const patch: Record<string, C | null> = {};
  for (const [key, value] of want) {
    if (!equals(read(after, key), value)) patch[key] = value;
  }
  return patch;
}
