/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Cells (Insert / Delete rows, columns and cells, Format > Row Height,
 * Column Width, Default Width, Hide & Unhide) and the row/column header menus
 * as pure sheet transactions. Moved out of SpreadsheetWindow so the ribbon and
 * the spreadsheet editor commands run the same code; behaviour is unchanged.
 * Protection notices, the selection, the validation alert, the macro recorder
 * and the measuring AutoFit stay in the window.
 */

import type { ExcelFontRecordClass } from "@mrnote/core/web/lib/office-fonts/excel-windows-faces.ts";
import { officePrimaryFontFamily } from "@mrnote/core/web/components/office-ribbon/officeTypographyCatalog.ts";
import { officePrintScale, type OfficePrintScale } from "./office-print-metrics";
import {
  cellShiftMetadataBlocker,
  canAdjustCellShiftFormulas,
  deleteCellsShift,
  insertCellsShift,
  insertCellsShiftDropsCells,
  remapCellShiftCharts,
  remapCellShiftConditionalFormats,
  remapCellShiftDataTables,
  remapCellShiftDrawings,
  remapCellShiftMerges,
  remapCellShiftNamedRanges,
  remapCellShiftPivotsOnSheet,
  remapCellShiftPrintSettings,
  remapCellShiftScenarios,
  remapCellShiftSlicers,
  remapCellShiftSparklines,
  remapCellShiftValidations,
  type CellShiftMetadataSheet,
  type ShiftDirection,
  type ShiftFormulaCell,
} from "./cell-shift";
import { withBestFitColumns, type CellButtonsSheet } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  adjustFormulaForCellShift,
  canAdjustFormulaForCellShift,
  type FormulaCellShiftOperation,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import {
  clampColumnWidth,
  clampRowHeight,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
  officeDigitWidth,
  officeIndentStepPx,
  officeRowHeight,
  officeRowMetricsKnown,
  setLineSizes,
  type SizeMap,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import {
  hideLines,
  normalizeManualHidden,
  serializeManualHidden,
  setLinesHidden,
  unhideLines,
} from "@mrnote/core/web/lib/spreadsheet/manual-hidden.ts";
import {
  editNativeColumns,
  editNativeRowSizes,
  nativeColumnCssClamp,
  nativeDefaultWidth,
  nativeWidthToPixels,
  pixelsToNativeWidth,
  syncNativeRowMetrics,
  type NativeGrid,
} from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";
import { rewriteSheetFormulaReferences } from "./sheet-formula-references";
import { insertedRowsKeepTheirShowing } from "./sheet-extent";
import {
  deletableLineCount,
  deleteLine,
  inheritInsertedLineFormats,
  insertLine,
  remapPivotSourcesOnEditedSheet,
  remapReferencesToEditedSheet,
  structuralInsertRefusal,
  structuralLineArrayRefusal,
  type StructuralInsertRefusal,
  type StructuralSheet,
} from "@mrnote/core/web/lib/spreadsheet/structural.ts";
import { normalizePivotRange } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { whatIfTableCellShiftRefused, whatIfTableLineRefused } from "@mrnote/core/web/lib/spreadsheet/what-if-table-definition.ts";

type Axis = "row" | "col";
type CellRange = { r1: number; c1: number; r2: number; c2: number };

/** What a structural edit means for the OTHER worksheets of the workbook. */
export type FollowOtherSheets<S> = (candidate: S) => S;

/**
 * A sheet edit as the editor commits it. Legacy Watch Window coordinates are
 * preserve-only workbook payload: structural helpers still remap that old
 * field, but session watches live outside the workbook, so the sheet's own
 * payload is restored and a row/column edit cannot create a hidden content
 * mutation. An imported grid's row metrics follow a changed row height.
 */
export function committedSheetEdit<S extends {
  watches?: unknown;
  nativeGrid?: NativeGrid;
  rowHeights?: SizeMap;
  defaultRowHeight?: number;
}>(sheet: S, nextSheet: S): S {
  let committedSheet = nextSheet;
  if (Object.prototype.hasOwnProperty.call(sheet, "watches")) {
    committedSheet = { ...nextSheet, watches: sheet.watches };
  } else if (Object.prototype.hasOwnProperty.call(nextSheet, "watches")) {
    committedSheet = { ...nextSheet };
    delete committedSheet.watches;
  }
  if (committedSheet.nativeGrid?.rowMetrics &&
      (committedSheet.rowHeights !== sheet.rowHeights || committedSheet.defaultRowHeight !== sheet.defaultRowHeight || committedSheet.nativeGrid !== sheet.nativeGrid)) {
    committedSheet = {...committedSheet,nativeGrid:syncNativeRowMetrics(committedSheet.nativeGrid,committedSheet.rowHeights,committedSheet.defaultRowHeight)};
  }
  return committedSheet;
}

/**
 * Excel adjusts a PivotTable's source like any reference when rows or columns
 * are inserted or deleted on the SOURCE sheet, wherever the report lives. The
 * edited sheet's own PivotTables move inside insertLine/deleteLine; this is
 * the same remap for the reports on the other worksheets, written in the
 * structural edit's own commit. `lines` > 0 were inserted at `at`, < 0
 * deleted from it. Every other reference those sheets make to the edited one
 * (`=First!A1`, names, validation and chart sources) follows in the same
 * commit, so one Undo restores the edit and every rewrite.
 */
export function pivotSourcesFollowLines<S extends StructuralSheet>(
  edited: S,
  axis: Axis,
  at: number,
  lines: number,
): FollowOtherSheets<S> | undefined {
  if (lines === 0) return undefined;
  return (candidate) =>
    remapReferencesToEditedSheet(
      remapPivotSourcesOnEditedSheet(candidate, edited, axis, at, lines),
      edited, axis, at, lines,
    );
}

/** Why Insert Rows / Columns did not run. */
export type LineInsertRefusal = StructuralInsertRefusal | "three-d" | "array" | "data-table" | "pivot";

/**
 * Insert `count` whole rows (columns) at `at` (Home > Cells > Insert Sheet
 * Rows / Columns, the header menu's Insert): the lines take the format of
 * the line before them, as Excel's Insert Options default does. A 3-D
 * reference (`First:Third!A5`) spanning the edited sheet cannot follow a
 * line edit on one sheet of its span, so the edit is refused.
 */
export function insertedLinesSheet<S extends StructuralSheet>(
  sheet: S,
  workbook: ReadonlyArray<StructuralSheet>,
  axis: Axis,
  at: number,
  count: number,
): { sheet: S; followOthers: FollowOtherSheets<S> | undefined } | LineInsertRefusal {
  const refusal = structuralInsertRefusal(sheet, axis, count);
  if (refusal) return refusal;
  if (linesCutPivotReport(sheet, axis, at, count, "insert")) return "pivot";
  // Excel: "You can't change part of an array" -- a line through a legacy
  // array formula's range would split it.
  if (structuralLineArrayRefusal(sheet, axis, at, count, "insert")) return "array";
  // "Cannot change part of a data table." -- a line through a What-If Data
  // Table would split it.
  if (whatIfTableLineRefused(sheet.whatIfTables, axis, at, count, "insert")) return "data-table";
  // A 3-D reference over this sheet (='1:31'!B21) does not stop the insert:
  // Excel inserts and leaves the reference as written, while the cells move
  // (Mac Excel 16, excel-formats/xprobe-rounds1113 H; xlsx-rounds R393-X20 ..
  // R397-X16). The rewrite never moves a 3-D reference. A group edit, which
  // moves every sheet of the span, still asks (sheet-group).
  const inserted = insertLine(sheet, axis, at, count);
  // The lines moved are the lines inserted. The extent's growth is not: at
  // the 1,048,576-row (16,384-column) cap it is capped, while the cells past
  // `at` still move by `count`, and other sheets' references must follow
  // them (O005).
  const lines = inserted === sheet ? 0 : count;
  const formatted = inserted === sheet ? inserted : inheritInsertedLineFormats(inserted, axis, at, count);
  return {
    // On a sheet whose rows are hidden by default, the rows moved past the
    // old extent keep showing; the insert says which rows are hidden itself
    
    sheet: formatted === sheet || axis !== "row" ? formatted : insertedRowsKeepTheirShowing(sheet, formatted, at, count),
    followOthers: pivotSourcesFollowLines(sheet, axis, at, lines),
  };
}

/**
 * Whether deleting rows `at..at+count-1` would take a table's header row and
 * leave some of its rows behind. Excel refuses that delete (measured in
 * Excel for Mac 16: Delete on the sheet row holding a table's header leaves
 * the sheet as it was; rows covering the whole table delete the table): the
 * header IS the columns' names, and the first record left behind would
 * become a header whose cells name nothing.
 */
export function rowDeleteCutsTableHeader(
  sheet: Pick<StructuralSheet, "dataTables">,
  at: number,
  count: number,
): boolean {
  const last = at + count - 1;
  return (sheet.dataTables ?? []).some((table) => {
    const { r1, r2 } = table.range;
    if (table.headerRow === false || r1 < at || r1 > last) return false;
    const bottom = r2 + (table.totalsRow ? 1 : 0);
    return bottom > last;
  });
}

/**
 * The lines of each report read from a file, on `axis`: its output
 * (`native.outputRef`) and, across rows, its filter rows above that output
 * with one blank row between (`native.savedLayout`).
 */
function pivotReportLines(sheet: StructuralSheet, axis: Axis): Array<{ filters: { first: number; last: number } | null; first: number; last: number }> {
  const reports = [sheet.pivot, ...(sheet.morePivots ?? [])];
  const out: Array<{ filters: { first: number; last: number } | null; first: number; last: number }> = [];
  for (const report of reports) {
    const native = (report as { native?: { outputRef?: unknown; savedLayout?: { pageFieldCount?: unknown; rowPageCount?: unknown } } } | undefined)?.native;
    const read = typeof native?.outputRef === "string" ? normalizePivotRange(native.outputRef) : null;
    if (!read) continue;
    // The size is the file's; where it stands now is the anchor, which follows
    // line edits above and beside the report (the read ref stays as it was).
    const stands = (report as { output?: unknown } | undefined)?.output;
    const anchor = typeof stands === "string" ? normalizePivotRange(stands) : null;
    const dr = anchor ? anchor.r1 - read.r1 : 0;
    const dc = anchor ? anchor.c1 - read.c1 : 0;
    const output = { r1: read.r1 + dr, r2: read.r2 + dr, c1: read.c1 + dc, c2: read.c2 + dc };
    const rows = axis === "row" && native?.savedLayout?.pageFieldCount ? Number(native.savedLayout.rowPageCount) : 0;
    out.push({
      filters: rows > 0 ? { first: output.r1 - 1 - rows, last: output.r1 - 2 } : null,
      first: axis === "row" ? output.r1 : output.c1,
      last: axis === "row" ? output.r2 : output.c2,
    });
  }
  return out;
}












function linesCutPivotReport(sheet: StructuralSheet, axis: Axis, at: number, count: number, op: "insert" | "delete"): boolean {
  return pivotReportLines(sheet, axis).some(({ filters, first, last }) => {
    // Across rows the output's first row is through it, under filters at least
    // (the case measured); its first column moves the whole report right.
    if (op === "insert") return (axis === "row" && filters ? first <= at : first < at) && at <= last;
    const end = at + count - 1;
    const takesOutput = at <= first && end >= last;
    if (filters && at <= filters.last && end >= filters.first && !(takesOutput && at <= filters.first)) return true;
    return at <= last && end >= first && !takesOutput;
  });
}

/**
 * Delete `requested` whole rows (columns) from `at`. Only lines inside the
 * persisted extent exist to delete: the blank lines the grid shows past
 * rowCount are on-demand, so a span running past the extent deletes its
 * stored part. Null when nothing is deleted; "table-header" when the rows
 * take a table's header row but not the whole table (Excel refuses it).
 */
export function deletedLinesSheet<S extends StructuralSheet>(
  sheet: S,
  workbook: ReadonlyArray<StructuralSheet>,
  axis: Axis,
  at: number,
  requested: number,
): { sheet: S; count: number; followOthers: FollowOtherSheets<S> | undefined } | "three-d" | "array" | "pivot" | "data-table" | "table-header" | null {
  const count = deletableLineCount(sheet, axis, at, requested);
  if (count === 0) return null;
  if (axis === "row" && rowDeleteCutsTableHeader(sheet, at, count)) return "table-header";
  if (linesCutPivotReport(sheet, axis, at, count, "delete")) return "pivot";
  // Deleted lines that cut a legacy array formula's range without removing
  // it whole would shrink the array (Excel refuses it).
  if (structuralLineArrayRefusal(sheet, axis, at, count, "delete")) return "array";
  // Deleted lines that cut a What-If Data Table without removing it whole.
  if (whatIfTableLineRefused(sheet.whatIfTables, axis, at, count, "delete")) return "data-table";
  // As for an insert: Excel deletes and leaves a 3-D reference as written
  // (xprobe-rounds1113 H).
  const next = deleteLine(sheet, axis, at, count);
  return {
    sheet: next,
    count,
    followOthers: pivotSourcesFollowLines(sheet, axis, at, next === sheet ? 0 : -count),
  };
}

/**
 * The same for Insert / Delete Cells with a shift: a report on another sheet
 * whose source block shifts whole on the edited sheet follows it. A partial
 * crossing was refused before the edit (remapCellShiftPivotsOnSheet), so the
 * remap here cannot fail.
 */
function pivotSourcesFollowCellShift<S extends CellShiftMetadataSheet>(
  edited: Pick<CellShiftMetadataSheet, "id" | "name">,
  range: CellRange,
  direction: ShiftDirection,
): FollowOtherSheets<S> {
  return (candidate) =>
    remapCellShiftPivotsOnSheet(candidate, edited, range, direction).sheet;
}

/** The cell-shift operation as a PEER worksheet sees it: only references
 * qualified with the edited sheet's name move. */
function peerCellShiftOperation(
  edited: { id: string; name: string },
  range: CellRange,
  direction: ShiftDirection,
): FormulaCellShiftOperation {
  return { range, direction, sheetName: edited.name || edited.id, qualifiedOnly: true };
}

/** True when every formula of a peer sheet can follow the shift (a range the
 * band crosses only in part stays as written; an unreadable one refuses). */
function peerCellShiftFormulasSupported(
  candidate: { cells: Record<string, ShiftFormulaCell> },
  operation: FormulaCellShiftOperation,
): boolean {
  for (const cell of Object.values(candidate.cells)) {
    const source =
      cell.formula ??
      (typeof cell.value === "string" && cell.value.startsWith("=") ? cell.value : null);
    if (source !== null && !canAdjustFormulaForCellShift(source, operation)) return false;
  }
  return true;
}

/**
 * What an Insert / Delete Cells shift means for the OTHER worksheets: their
 * PivotTable sources follow (above) and so does every formula, name,
 * validation or chart of theirs that names the shifted cells on the edited
 * sheet (=数据!G21 becomes =数据!G22 when G20 shifts down; live audit B148),
 * the way this sheet's own references already do.
 */
export function peerSheetsFollowCellShift<S extends CellShiftMetadataSheet & { id: string; name: string }>(
  edited: S,
  range: CellRange,
  direction: ShiftDirection,
): FollowOtherSheets<S> {
  const followPivots = pivotSourcesFollowCellShift<S>(edited, range, direction);
  const operation = peerCellShiftOperation(edited, range, direction);
  return (candidate) => {
    const withPivots = followPivots(candidate);
    if (candidate.id === edited.id) return withPivots;
    return rewriteSheetFormulaReferences(withPivots, (formula) => {
      const result = adjustFormulaForCellShift(formula, operation);
      return result.supported ? result.formula : formula;
    });
  };
}

export interface CellShiftButtonSheet extends CellShiftMetadataSheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, ShiftFormulaCell>;
  whatIfTables?: unknown;
}

/**
 * Insert Cells (shift down / right) or Delete Cells (shift up / left) on
 * `range` (Excel "Insert Cells…" / "Delete Cells…"): only the cells on the
 * range's rows (columns) move, and every formula, validation, conditional
 * format, merge, table, PivotTable, slicer, name, print area, chart,
 * drawing, scenario and sparkline that reads them follows. Null when a part
 * of the workbook cannot safely follow the shift (the window says so). An
 * insert that pushes cells past the sheet's extent grows it, up to the
 * schema bounds.
 */
export function cellShiftedSheet<S extends CellShiftButtonSheet>(
  sheet: S,
  workbook: readonly S[],
  range: CellRange,
  direction: ShiftDirection,
  bounds: { maxRows: number; maxCols: number },
): { sheet: S; followOthers: FollowOtherSheets<S> } | null {
  if (
    !canAdjustCellShiftFormulas(sheet.cells, range, direction, sheet.name) ||
    cellShiftMetadataBlocker(sheet, range, direction, workbook) !== null ||
    // A What-If Data Table never moves in part (nor here, whole).
    whatIfTableCellShiftRefused(sheet.whatIfTables, range, direction)
  ) {
    return null;
  }
  const validationPlan = remapCellShiftValidations(sheet.validations, range, direction, sheet.name);
  const conditionalPlan = remapCellShiftConditionalFormats(sheet.conditionalFormats, range, direction, sheet.name);
  const mergePlan = remapCellShiftMerges(sheet.merges, range, direction);
  const tablePlan = remapCellShiftDataTables(sheet.dataTables, range, direction);
  // This sheet's PivotTables: a source here follows a whole-block shift
  // (a report whose data is on another sheet keeps its source), and an
  // output anchor may not be moved. The reports on OTHER sheets that read
  // this one get the same source remap, in the same commit
  // (`pivotSourcesFollowCellShift`), and refuse the shift the same way.
  const pivotPlan = remapCellShiftPivotsOnSheet(sheet, sheet, range, direction);
  const otherSheetPivotsSupported = workbook.every(
    (candidate) =>
      candidate.id === sheet.id ||
      remapCellShiftPivotsOnSheet(candidate, sheet, range, direction).supported,
  );
  const peerOperation = peerCellShiftOperation(sheet, range, direction);
  const peerFormulasSupported = workbook.every(
    (candidate) =>
      candidate.id === sheet.id || peerCellShiftFormulasSupported(candidate, peerOperation),
  );
  const slicerPlan = remapCellShiftSlicers(sheet.slicers, sheet.dataTables, tablePlan.items);
  const namePlan = remapCellShiftNamedRanges(sheet.namedRanges, range, direction, sheet.id, sheet.name);
  const printPlan = remapCellShiftPrintSettings(sheet.printSettings, range, direction, sheet.name);
  const chartPlan = remapCellShiftCharts(sheet.charts, range, direction);
  const drawingPlan = remapCellShiftDrawings(sheet.objects, range, direction, sheet.name);
  const scenarioPlan = remapCellShiftScenarios(sheet.scenarios, range, direction, sheet.name);
  const sparklinePlan = remapCellShiftSparklines(sheet.sparklines, range, direction, sheet.name);
  if (
    !validationPlan.supported ||
    !conditionalPlan.supported ||
    !mergePlan.supported ||
    !tablePlan.supported ||
    !pivotPlan.supported ||
    !otherSheetPivotsSupported ||
    !peerFormulasSupported ||
    !slicerPlan.supported ||
    !namePlan.supported ||
    !printPlan.supported ||
    !chartPlan.supported ||
    !drawingPlan.supported ||
    !scenarioPlan.supported ||
    !sparklinePlan.supported
  ) {
    return null;
  }
  const inserting = direction === "down" || direction === "right";
  // A stored cell pushed off the sheet is refused, never dropped (B331).
  if (
    (direction === "down" || direction === "right") &&
    insertCellsShiftDropsCells(sheet.cells, range, direction, bounds.maxRows, bounds.maxCols)
  ) {
    return null;
  }
  const cells =
    direction === "down" || direction === "right"
      ? insertCellsShift(sheet.cells, range, direction, bounds.maxRows, bounds.maxCols, sheet.name)
      : deleteCellsShift(sheet.cells, range, direction, sheet.name);
  // Grow the sheet so a shifted cell that lands past the current extent stays
  // visible (clamped to the schema bound by the helper's drop rule).
  const height = range.r2 - range.r1 + 1;
  const width = range.c2 - range.c1 + 1;
  const next = {
    ...sheet,
    cells,
    ...(inserting
      ? {
          rowCount:
            direction === "down" ? Math.min(bounds.maxRows, sheet.rowCount + height) : sheet.rowCount,
          columnCount:
            direction === "right" ? Math.min(bounds.maxCols, sheet.columnCount + width) : sheet.columnCount,
        }
      : {}),
    ...(validationPlan.items !== undefined ? { validations: validationPlan.items } : {}),
    ...(conditionalPlan.items !== undefined ? { conditionalFormats: conditionalPlan.items } : {}),
    ...(mergePlan.items !== undefined ? { merges: mergePlan.items } : {}),
    ...(tablePlan.items !== undefined ? { dataTables: tablePlan.items } : {}),
    ...(pivotPlan.sheet.pivot ? { pivot: pivotPlan.sheet.pivot } : {}),
    ...(pivotPlan.sheet.morePivots ? { morePivots: pivotPlan.sheet.morePivots } : {}),
    ...(slicerPlan.items !== undefined ? { slicers: slicerPlan.items } : {}),
    ...(namePlan.items !== undefined ? { namedRanges: namePlan.items } : {}),
    ...(printPlan.item !== sheet.printSettings ? { printSettings: printPlan.item } : {}),
    ...(chartPlan.items !== undefined ? { charts: chartPlan.items } : {}),
    ...(drawingPlan.items !== undefined ? { objects: drawingPlan.items } : {}),
    ...(scenarioPlan.items !== undefined ? { scenarios: scenarioPlan.items } : {}),
    ...(sparklinePlan.item !== undefined ? { sparklines: sparklinePlan.item } : {}),
  } as S;
  return { sheet: next, followOthers: peerSheetsFollowCellShift(sheet, range, direction) };
}

/**
 * Format Rows / Format Columns (resizing, auto-fitting or hiding a line) on a
 * protected sheet is the sheet's Format rows / Format columns permission
 * (Excel's Protect Sheet list). A span outside the sheet is refused too.
 */
export function lineFormatProtected<S extends { rowCount: number; columnCount: number }>(
  sheet: S,
  axis: Axis,
  start: number,
  end: number,
  protectionAllows: (sheet: S, permission: "formatRows" | "formatColumns") => boolean,
): boolean {
  // "Outside the sheet" is outside the worksheet's address space. A blank
  // column past the last stored one is a column like any other, and saying
  // the sheet is protected because nothing is stored there yet is wrong
  // (live audit B381).
  const lineCount = axis === "row" ? MAX_GRID_ROWS : MAX_GRID_COLUMNS;
  if (
    !Number.isSafeInteger(start) ||
    !Number.isSafeInteger(end) ||
    start < 0 ||
    end < start ||
    end >= lineCount
  ) {
    return true;
  }
  return !protectionAllows(sheet, axis === "row" ? "formatRows" : "formatColumns");
}

export interface HiddenLineSheet {
  rowCount: number;
  columnCount: number;
  manuallyHiddenRows?: number[];
  manuallyHiddenCols?: number[];
  /** Reading 998: those of manuallyHiddenRows the file writes 0 points high. */
  zeroHeightRows?: number[];
}

/** A row shown again is no longer one hidden by a height of 0 points
 *  (reading 998): `zeroHeightRows` keeps only rows still hidden. */
function keepZeroHeightRowsHidden<S extends HiddenLineSheet>(next: S): S {
  if (!Array.isArray(next.zeroHeightRows)) return next;
  const hidden = new Set(next.manuallyHiddenRows ?? []);
  const kept = next.zeroHeightRows.filter((row) => hidden.has(row));
  if (kept.length === next.zeroHeightRows.length) return next;
  if (kept.length > 0) next.zeroHeightRows = kept;
  else delete next.zeroHeightRows;
  return next;
}

/** Set equality of two hidden-line lists. */
export function numberListsEqual(
  left: readonly number[] | ReadonlySet<number>,
  right: readonly number[] | ReadonlySet<number> | undefined,
): boolean {
  const leftSet: ReadonlySet<number> = Array.isArray(left)
    ? new Set(left)
    : (left as ReadonlySet<number>);
  const rightSet: ReadonlySet<number> = Array.isArray(right)
    ? new Set(right)
    : ((right ?? new Set<number>()) as ReadonlySet<number>);
  return (
    leftSet.size === rightSet.size &&
    Array.from(leftSet).every((value) => rightSet.has(value))
  );
}

/** The lines the sheet hides by hand, to the grid's last line: a hidden tail
 *  past the stored extent (`<col min="6" max="16384" hidden="1">` on a sheet
 *  whose cells end at E) is the sheet's too. Read to the extent only, a Hide
 *  or Unhide of another line stored what was inside it and the tail came out
 *  shown (xlsx-rounds R218-X02 E22). */
function manualHidden(sheet: HiddenLineSheet, axis: Axis): Set<number> {
  return axis === "row"
    ? normalizeManualHidden(sheet.manuallyHiddenRows, MAX_GRID_ROWS)
    : normalizeManualHidden(sheet.manuallyHiddenCols, MAX_GRID_COLUMNS);
}

function withManualHidden<S extends HiddenLineSheet>(
  sheet: S,
  axis: Axis,
  current: ReadonlySet<number>,
  hidden: ReadonlySet<number>,
): S | null {
  const next = serializeManualHidden(hidden);
  if (numberListsEqual(current, next)) return null;
  const key = axis === "row" ? "manuallyHiddenRows" : "manuallyHiddenCols";
  const sheetNext: S = { ...sheet };
  if (next) sheetNext[key] = next;
  else delete sheetNext[key];
  if (axis === "row") keepZeroHeightRowsHidden(sheetNext);
  // A line hidden past the stored extent stores that line, as a width given
  // to one does: the grid and the loader read the hidden lines inside the
  // extent only, so Hide on a blank column past it did nothing visible and
  // the next Hide or Unhide dropped it (sheet model fuzz).
  // Only a line this edit hides grows it; one the sheet already hid past it does not.
  let last: number | undefined;
  for (const line of hidden) if (!current.has(line) && (last === undefined || line > last)) last = line;
  if (last !== undefined) {
    if (axis === "row" && last >= sheetNext.rowCount) sheetNext.rowCount = Math.min(MAX_GRID_ROWS, last + 1);
    if (axis === "col" && last >= sheetNext.columnCount) sheetNext.columnCount = Math.min(MAX_GRID_COLUMNS, last + 1);
  }
  return sheetNext;
}

/**
 * Manual Hide / Unhide (Excel header right-click, Home > Format > Hide &
 * Unhide): Hide collapses the span [start, end]; Unhide reveals every hidden
 * line inside it (Excel: select across or bracket a hidden run, then
 * Unhide). Distinct from outline collapse. Null when nothing changes; the
 * list is dropped when empty so the no-hidden save stays byte-identical.
 */
export function manualHiddenSpanSheet<S extends HiddenLineSheet>(
  sheet: S,
  axis: Axis,
  start: number,
  end: number,
  hide: boolean,
): S | null {
  const current = manualHidden(sheet, axis);
  return withManualHidden(sheet, axis, current, hide ? hideLines(current, start, end) : unhideLines(current, start, end));
}

/**
 * Excel hides a column given 0 in Column Width (a row given 0 in Row Height)
 * rather than sizing it: the size it had comes back when it is shown again.
 * Exactly these lines, however scattered. Null when nothing changes.
 */
export function linesHiddenSheet<S extends HiddenLineSheet>(
  sheet: S,
  axis: Axis,
  lines: readonly number[],
  hide: boolean,
): S | null {
  const current = manualHidden(sheet, axis);
  return withManualHidden(sheet, axis, current, setLinesHidden(current, lines, hide));
}

/** An imported workbook's grid as sized for display and print (below). */
export interface ResolvedNativeGrid {
  grid: NativeGrid;
  dpi: number;
  digitWidth: number;
  fontRowHeight: number;
  fontReplaced: boolean;
  /** The automatic default row is the grid font's own (fontRowHeight), not
   * the height the file states: a theme replaced the font, or the file states
   * a default that is not custom and the tables know the face. Windows Excel
   * draws such a sheet's rows at its own default row for the font (lookwin
   * probe p26: 14.4, 14.5 or 20 pt stated without customHeight all draw 15 pt
   * in Calibri 11); a file saved at 120 dpi states 14.4. */
  automaticDefaultRow: boolean;
  /** The face the grid's rows are measured in (fontRowHeight's). */
  rowFont: { family: string; sizePt: number; record?: ExcelFontRecordClass };
  widths: Record<string, number>;
  defaultWidth: number | undefined;
  /** Windows Excel's printer scale of this grid (officePrintScale): rows by
   * the grid font, columns by the face they are counted in. */
  printScale?: OfficePrintScale;
  /** One indent level on the screen, px: three of the grid font's spaces
   * (officeIndentStepPx, lookwin p87-lw). */
  indentStepPx?: number;
}

/**
 * An imported workbook's grid, sized as Office sizes it and as the printed
 * page prints it: the 96dpi grid the model stores; columns count the digits
 * of the grid font and an automatic row takes that font's line box, both from
 * the tables measured off the font files Office ships. A theme that replaces
 * the grid font re-measures the untouched source columns and the automatic
 * default row in the new face, as Excel does.
 */
export function resolveSpreadsheetNativeGrid(
  grid: NativeGrid | undefined,
  widths: Record<string, number> | undefined,
  defaultWidth: number | undefined,
  themeId?: string,
  bodyFamily?: string,
  headingFamily?: string,
  // The grid font as the UI's language draws it (`nativeGridFontForLocale`):
  // Chinese Excel draws a theme-bound Normal font in the theme's Hans face
  // and counts its columns in that face's digit, 等线 11's 8 px where the
  // file's Calibri has 7 (X-329). Rows keep the workbook's own heights.
  localeFamily?: string,
  // The UI's face also measures the rows (a Western UI's Latin face).
  localeRows = false,
): ResolvedNativeGrid | undefined {
  const changedTheme = grid && themeId && (grid.sourceThemeFontId ?? "office") !== themeId;
  const measurementFont = grid?.metricsFont ?? grid?.font;
  const replacement = measurementFont?.themeRole === "major" ? headingFamily : bodyFamily;
  const displayFamily = measurementFont && (changedTheme && measurementFont.themeRole && replacement
    ? officePrimaryFontFamily(replacement) : measurementFont.family);
  if (!grid || !displayFamily) return undefined;
  const dpi = 96;
  const columnFamily = !changedTheme && measurementFont?.themeRole && localeFamily && localeFamily !== displayFamily
    ? officePrimaryFontFamily(localeFamily) : displayFamily;
  const rowFamily = localeRows ? columnFamily : displayFamily;
  // The record's class (a face Windows lacks, drawn by it) belongs to the
  // file's own face, not to a theme or UI face put in its place.
  const recordOf = (family: string) => family === measurementFont?.family && measurementFont.record ? { record: measurementFont.record } : {};
  const font = { family: rowFamily, sizePt: measurementFont?.sizePt, ...recordOf(rowFamily) };
  const digitWidth = officeDigitWidth({ family: columnFamily, sizePt: measurementFont?.sizePt, ...recordOf(columnFamily) });
  const fontRowHeight = officeRowHeight(font);
  const resolved: Record<string, number> = { ...widths };
  for (const [index, column] of Object.entries(grid.columns)) {
    if (widths?.[index] === column.pixels) resolved[index] = nativeWidthToPixels(column.width, digitWidth, dpi);
  }
  // The workbook's own row heights describe its own font; only a theme
  // change puts a different face on the sheet, whose rows must be re-derived.
  const fontReplaced = Boolean(changedTheme && measurementFont?.themeRole && replacement);
  const automaticDefaultRow = fontReplaced || officeRowMetricsKnown(font.family, font.sizePt, font.record);
  const rowPrint = officePrintScale(font, digitWidth);
  const printScale = columnFamily === rowFamily
    ? rowPrint
    : { x: officePrintScale({ family: columnFamily, sizePt: measurementFont?.sizePt, ...recordOf(columnFamily) }, digitWidth).x, y: rowPrint.y };
  return { grid, dpi, digitWidth, fontRowHeight, fontReplaced, automaticDefaultRow, indentStepPx: officeIndentStepPx(font),
    rowFont: { family: rowFamily, sizePt: measurementFont?.sizePt ?? 11, ...recordOf(rowFamily) }, widths: resolved, printScale,
    defaultWidth: defaultWidth === grid.defaultPixels
      ? nativeDefaultWidth(grid, digitWidth, dpi, measurementFont
        ? officeDigitWidth({ family: measurementFont.family, sizePt: measurementFont.sizePt, ...recordOf(measurementFont.family) })
        : undefined)
      : defaultWidth };
}

export interface LineSizeSheet extends CellButtonsSheet {
  nativeGrid?: NativeGrid;
  rowHeights?: SizeMap;
  defaultRowHeight?: number;
  manuallyHiddenRows?: number[];
  manuallyHiddenCols?: number[];
  zeroHeightRows?: number[];
}

/**
 * Sanitize a persisted column-width / row-height map (Excel parity): keep only
 * integer in-bounds line indices, clamp each value to the draggable range, and
 * drop entries that resolve back to the default (redundant). Returns undefined
 * when nothing survives so the field can be omitted entirely.
 */
export function normalizeSizeMap(
  raw: unknown,
  lineCount: number,
  clamp: (value: unknown, fallback: number) => number,
  defaultSize: number,
  preserveExplicitDefault = false,
): SizeMap | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const fallback = clamp(defaultSize, 1);
  const out: SizeMap = {};
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    const index = Number(key);
    if (!Number.isInteger(index) || index < 0 || index >= lineCount) continue;
    const size = clamp(value, fallback);
    if (size !== fallback || preserveExplicitDefault) out[index] = size;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

export function sizeMapsEqual(
  left: SizeMap | undefined,
  right: SizeMap | undefined,
): boolean {
  const leftEntries = Object.entries(left ?? {});
  const rightEntries = Object.entries(right ?? {});
  if (leftEntries.length !== rightEntries.length) return false;
  return leftEntries.every(([key, value]) => right?.[key] === value);
}

/** Show the given lines in the same step (a size typed for a hidden line):
 * whether any is hidden, and the write that shows them. */
function revealing<S extends LineSizeSheet>(
  sheet: S,
  axis: Axis,
  lines: readonly number[],
  reveal: boolean | undefined,
): { changes: boolean; withShown: (next: S) => S } {
  const unchanged = { changes: false, withShown: (next: S) => next };
  if (!reveal) return unchanged;
  const current = manualHidden(sheet, axis);
  const shown = setLinesHidden(current, lines, false);
  if (shown.size === current.size) return unchanged;
  const key = axis === "row" ? "manuallyHiddenRows" : "manuallyHiddenCols";
  const list = serializeManualHidden(shown);
  return {
    changes: true,
    withShown: (next) => {
      if (list) next[key] = list;
      else delete next[key];
      return axis === "row" ? keepZeroHeightRowsHidden(next) : next;
    },
  };
}

/**
 * Several column widths, in pixels, as ONE model change (Column Width…, a
 * multi-column resize or AutoFit). `reveal`: Excel shows a hidden column
 * that is given a width in the same step. `bestFit`: the widths were fitted
 * to the content, so the columns stay best-fit; a width the user set ends
 * it. An imported grid stores the width in its own native units. Null when
 * nothing changes or an index is outside the sheet.
 */
export function columnWidthsSheet<S extends LineSizeSheet>(
  sheet: S,
  sizes: ReadonlyArray<{ index: number; size: number }>,
  options: {
    nativeGrid: ResolvedNativeGrid | undefined;
    defaultWidth: number;
    reveal?: boolean;
    bestFit?: boolean;
  },
): S | null {
  const valid = sizes.filter(
    ({ index }) => Number.isSafeInteger(index) && index >= 0 && index < MAX_GRID_COLUMNS,
  );
  if (valid.length === 0 || valid.length !== sizes.length) return null;
  const lines = valid.map(({ index }) => index);
  // A width given to a column past the stored extent stores that column.
  const columnCount = Math.max(sheet.columnCount, ...lines.map((index) => index + 1));
  if (columnCount !== sheet.columnCount) sheet = { ...sheet, columnCount };
  const { changes: changesHidden, withShown } = revealing(sheet, "col", lines, options.reveal);
  const { nativeGrid, defaultWidth } = options;
  if (sheet.nativeGrid && !nativeGrid) return null;
  if (nativeGrid) {
    const edited = valid.map(({ index, size }) => ({ index, size: clampColumnWidth(size, defaultWidth) }));
    const nextGrid = editNativeColumns(nativeGrid.grid, edited, nativeGrid.digitWidth, nativeGrid.dpi);
    if (!changesHidden && edited.every(({ index }) => Math.abs(nativeColumnCssClamp(nativeGrid.widths[index], nativeGrid.defaultWidth ?? defaultWidth) - nativeWidthToPixels(nextGrid.columns[index].width, nativeGrid.digitWidth, nativeGrid.dpi)) < 0.01)) return null;
    return withShown({ ...sheet, columnWidths: { ...sheet.columnWidths, ...Object.fromEntries(edited.map(({ index }) => [index, nextGrid.columns[index].pixels])) }, nativeGrid: nextGrid });
  }
  const nextWidths = setLineSizes(sheet.columnWidths, valid, defaultWidth, clampColumnWidth);
  const marked = withBestFitColumns(sheet, lines, options.bestFit === true);
  if (!changesHidden && marked === sheet && sizeMapsEqual(sheet.columnWidths, nextWidths)) return null;
  const next: S = { ...marked };
  if (nextWidths) next.columnWidths = nextWidths;
  else delete next.columnWidths;
  return withShown(next);
}

/**
 * Several row heights, in pixels, as ONE model change (Row Height…, a
 * multi-row resize or AutoFit); `reveal` as for columns. A row whose content
 * would grow it on its own (`autoRow`) keeps an explicit default height so
 * the height set stays. An imported grid records the height in points too.
 * Null when nothing changes or an index is outside the sheet.
 */
export function rowHeightsSheet<S extends LineSizeSheet>(
  sheet: S,
  sizes: ReadonlyArray<{ index: number; size: number }>,
  options: {
    defaultHeight: number;
    autoRow: (row: number) => boolean;
    reveal?: boolean;
  },
): S | null {
  const valid = sizes.filter(
    ({ index }) => Number.isSafeInteger(index) && index >= 0 && index < MAX_GRID_ROWS,
  );
  if (valid.length === 0 || valid.length !== sizes.length) return null;
  // A height given to a row past the stored extent stores that row.
  const rowCount = Math.max(sheet.rowCount, ...valid.map(({ index }) => index + 1));
  if (rowCount !== sheet.rowCount) sheet = { ...sheet, rowCount };
  const { changes: changesHidden, withShown } = revealing(sheet, "row", valid.map(({ index }) => index), options.reveal);
  if (sheet.nativeGrid?.rowMetrics) {
    const nextGrid = editNativeRowSizes(sheet.nativeGrid, valid);
    const heights = { ...sheet.rowHeights, ...Object.fromEntries(valid.map(({ index }) => [index, nextGrid.rowMetrics!.rows[index].pixels])) };
    return withShown({ ...sheet, nativeGrid: nextGrid, rowHeights: heights });
  }
  const { defaultHeight } = options;
  let nextHeights = setLineSizes(sheet.rowHeights, valid, defaultHeight, clampRowHeight);
  for (const { index, size } of valid) {
    if (options.autoRow(index) && clampRowHeight(size, defaultHeight) === defaultHeight) {
      nextHeights = { ...nextHeights, [index]: defaultHeight };
    }
  }
  if (!changesHidden && sizeMapsEqual(sheet.rowHeights, nextHeights)) return null;
  const next: S = { ...sheet };
  if (nextHeights) next.rowHeights = nextHeights;
  else delete next.rowHeights;
  return withShown(next);
}

/**
 * Home > Format > Default Width (Excel's Standard Width), in pixels: the
 * width of every column without its own. An imported grid stores it in its
 * native units. Null when nothing changes.
 */
export function defaultColumnWidthSheet<S extends LineSizeSheet>(
  sheet: S,
  width: number,
  options: { nativeGrid: ResolvedNativeGrid | undefined; defaultWidth: number; standardWidth: number },
): S | null {
  const nextDefault = clampColumnWidth(width, options.standardWidth);
  const { nativeGrid } = options;
  if (sheet.nativeGrid && !nativeGrid) return null;
  if (nativeGrid) {
    const nativeWidth = pixelsToNativeWidth(nextDefault, nativeGrid.digitWidth, nativeGrid.dpi);
    const pixels = nativeWidthToPixels(nativeWidth, nativeGrid.digitWidth, nativeGrid.dpi);
    return { ...sheet, defaultColumnWidth: pixels,
      nativeGrid: { ...nativeGrid.grid, defaultPixels: pixels, defaultWidth: nativeWidth,
        defaultProfile: { dpi: nativeGrid.dpi as 72 | 96, digitWidth: nativeGrid.digitWidth } } };
  }
  if (nextDefault === options.defaultWidth) return null;
  const next: S = { ...sheet };
  if (nextDefault === options.standardWidth) {
    delete next.defaultColumnWidth;
  } else {
    next.defaultColumnWidth = nextDefault;
  }
  const nextWidths = normalizeSizeMap(sheet.columnWidths, sheet.columnCount, clampColumnWidth, nextDefault);
  if (nextWidths) next.columnWidths = nextWidths;
  else delete next.columnWidths;
  return next;
}
