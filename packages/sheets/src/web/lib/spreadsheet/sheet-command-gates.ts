/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  spreadsheetVisibleControlDecision,
  type SpreadsheetVisibleCommandContext,
  type SpreadsheetVisibleCommandReason,
  type SpreadsheetVisibleControlId,
} from "./visible-command-authority";
import {
  canDeleteLines,
  rangeHasLockedCell as storedRangeHasLockedCell,
  sheetProtectionAllows as storedSheetProtectionAllows,
  sheetProtectionPermissions,
  UNPROTECTED_SHEET_PERMISSIONS,
  type SheetProtectionPermission,
  type SheetProtectionPermissions,
} from "@mrnote/core/web/lib/spreadsheet/protection.ts";
import {
  expandRangeToMergeClosure,
  mergeAcrossSpans,
  mergeContentCarry,
  mergeSpanFor,
  normalizeMerges,
  rangeFullyMerged,
  rangeHasMerge,
  type MergeSpan,
} from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { mergeRectTouchesTable } from "./merge-buttons";
import {
  cellShiftedSheet,
  committedSheetEdit,
  deletedLinesSheet,
  insertedLinesSheet,
  type CellShiftButtonSheet,
  type FollowOtherSheets,
} from "./row-column-buttons";
import {
  canAdjustCellShiftFormulas,
  cellShiftMetadataBlocker,
  cellShiftRangeSensitiveFormulas,
  insertCellsShiftDropsCells,
  type ShiftDirection,
} from "./cell-shift";
import {
  arrayPartEditRefused,
  deletableLineCount,
  type StructuralSheet,
} from "@mrnote/core/web/lib/spreadsheet/structural.ts";
import {
  clearedSheetForRange,
  type ClearButtonSheet,
} from "./clear-buttons";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

export interface SheetCommandRect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Why a sheet command did not run: the authority's reason, or the notice
 *  key of the transaction that refused. */
export type SheetCommandRefusalKey =
  | SpreadsheetVisibleCommandReason
  | "spreadsheet.protectionBlocked"
  | "spreadsheet.structuralThreeDReferenceBlocked"
  | "spreadsheet.arrayPartBlocked"
  | "spreadsheet.whatIfTablePartBlocked"
  | "spreadsheet.pivotPartBlocked"
  | "spreadsheet.cellShiftRefusedTable"
  | "spreadsheet.cellShiftRefusedMerge"
  | "spreadsheet.cellShiftRefusedPivot"
  | "spreadsheet.cellShiftRefusedGeneric"
  | "spreadsheet.pivotPartEditBlocked"
  | "spreadsheet.uiCopy.insertWouldShiftCellsOffSheet"
  | "spreadsheet.uiCopy.insertLineBudgetExceeded"
  | "spreadsheet.uiCopy.cantChangePartOfArray"
  | "spreadsheet.uiCopy.canTSortBecauseMergedCellsWouldBeSplit"
  | "spreadsheet.mergeDiscardsValuesConfirm"
  | "spreadsheet.chartsLoadingStructureBlocked";

export type SheetCommandRefused = { readonly kind: "refused"; readonly reasonKey: SheetCommandRefusalKey };

const refused = (reasonKey: SheetCommandRefusalKey): SheetCommandRefused => ({ kind: "refused", reasonKey });

/**
 * Whether a chart imported with its model deferred is still waiting for it
 * (CHART-DEFER-DESIGN.md): its references are not on the page yet, so an
 * edit that moves lines or cells would leave them pointing at the old
 * places. The page fills them in within seconds of opening.
 */
export function chartsStillLoading(
  workbook: ReadonlyArray<{ objects?: ReadonlyArray<unknown> }>,
  editedSheetName?: string,
): boolean {
  // A placeholder the server ran a command over names the sheets its
  // references read (`refSheets`, read off its chart part): one that reads
  // none of the edited sheet's cells has nothing to move (R116-X02: a row
  // insert filled 183 charts in to move the references of a few).
  const edited = editedSheetName?.toLowerCase();
  // A sheet the import left a placeholder (lib/spreadsheet/deferred-sheets)
  // holds none of what an edit across sheets reads or moves until it is in.
  if (workbook.some((sheet) => (sheet as { sheetPending?: unknown }).sheetPending === true)) return true;
  return workbook.some((sheet) =>
    (sheet.objects ?? []).some((object) => {
      const placeholder = object as { chartPending?: unknown; refSheets?: unknown } | null;
      if (placeholder?.chartPending !== true) return false;
      if (edited === undefined || !Array.isArray(placeholder.refSheets)) return true;
      return placeholder.refSheets.some((name) => typeof name !== "string" || name.toLowerCase() === edited);
    }));
}

// ── Protection, as one entry point reads it ─────────────────────────────

type ProtectedSheetLike = Parameters<typeof storedRangeHasLockedCell>[0] & {
  protectionOptions?: Parameters<typeof sheetProtectionPermissions>[0];
};

/**
 * How an entry point reads protection. The window reads its session (a
 * paused protection, an edit range unlocked this session); a headless run
 * reads the stored sheet (`STORED_SHEET_PROTECTION`).
 */
export interface SheetProtectionReading<S> {
  effective(sheet: S): boolean;
  allows(sheet: S, permission: SheetProtectionPermission): boolean;
  rangeHasLockedCell(sheet: S, range: SheetCommandRect): boolean;
}

export const STORED_SHEET_PROTECTION: SheetProtectionReading<object> = Object.freeze({
  effective: (sheet: object) => (sheet as ProtectedSheetLike).protected === true,
  allows: (sheet: object, permission: SheetProtectionPermission) =>
    storedSheetProtectionAllows(sheet as ProtectedSheetLike, permission),
  rangeHasLockedCell: (sheet: object, range: SheetCommandRect) =>
    storedRangeHasLockedCell(sheet as ProtectedSheetLike, range),
});

interface ExtentSheet {
  rowCount: number;
  columnCount: number;
  protectionOptions?: Parameters<typeof sheetProtectionPermissions>[0];
}

/** Excel's Protect Sheet permissions on the sheet: all open when it is not
 *  (effectively) protected. */
export function sheetCommandPermissions<S extends ExtentSheet>(
  sheet: S,
  protection: SheetProtectionReading<S>,
): Readonly<SheetProtectionPermissions> {
  return protection.effective(sheet)
    ? sheetProtectionPermissions(sheet.protectionOptions)
    : UNPROTECTED_SHEET_PERMISSIONS;
}

/**
 * The whole-line structure edits the sheet's protection opens for a
 * selection: Insert rows / columns by permission alone, Delete rows / columns
 * only for lines whose every cell is unlocked (Worksheet.Protect
 * AllowDeletingRows / AllowDeletingColumns).
 */
export function sheetStructuralLineEdits<S extends ExtentSheet>(
  sheet: S,
  rect: SheetCommandRect,
  protection: SheetProtectionReading<S>,
): { insertRows: boolean; insertColumns: boolean; deleteRows: boolean; deleteColumns: boolean } {
  const effective = protection.effective(sheet);
  const allows = sheetCommandPermissions(sheet, protection);
  const lastRow = Math.max(0, sheet.rowCount - 1);
  const lastColumn = Math.max(0, sheet.columnCount - 1);
  return {
    insertRows: allows.insertRows,
    insertColumns: allows.insertColumns,
    // The lines' own Locked format decides, not an edit range unlocked
    // this session (canDeleteLines).
    deleteRows:
      allows.deleteRows &&
      (!effective ||
        canDeleteLines(sheet as never, "row", rect.r1, Math.min(rect.r2, lastRow), sheet.columnCount)),
    deleteColumns:
      allows.deleteColumns &&
      (!effective ||
        canDeleteLines(sheet as never, "col", rect.c1, Math.min(rect.c2, lastColumn), sheet.rowCount)),
  };
}

/**
 * Sorting under protection: the sheet's Allow Sort permission and every cell
 * of the selection unlocked (`lockedSelection` false). Always open on an
 * unprotected sheet.
 */
export function sheetSortAuthorized<S extends ExtentSheet>(
  sheet: S,
  protection: SheetProtectionReading<S>,
  /** The selection's Locked format: "mixed" when it holds both. */
  lockedSelection: boolean | "mixed" | null,
): boolean {
  return (
    !protection.effective(sheet) ||
    (sheet.protectionOptions?.allowSort === true && lockedSelection === false)
  );
}

/**
 * What the merge commands read about a selection: the range they act on (a
 * selection overlapping a merge takes its whole area, as Excel selects a
 * merged cell's area), whether it spans cells and columns, whether it touches
 * a table (Excel never merges in one).
 */
export function sheetMergeSelection(
  sheet: object,
  spans: readonly MergeSpan[],
  rect: SheetCommandRect,
): {
  target: SheetCommandRect;
  mergeSelectionMultiCell: boolean;
  mergeSelectionMultiColumn: boolean;
  mergeSelectionInTable: boolean;
  selectionHasMerge: boolean;
} {
  const target = expandRangeToMergeClosure(rect, spans as MergeSpan[]);
  return {
    target,
    mergeSelectionMultiCell: target.r1 !== target.r2 || target.c1 !== target.c2,
    mergeSelectionMultiColumn: target.c1 !== target.c2,
    mergeSelectionInTable: mergeRectTouchesTable(sheet, target),
    selectionHasMerge: rangeHasMerge(spans as MergeSpan[], rect),
  };
}

interface ContextSheet extends ExtentSheet {
  name?: string;
  cells: Record<string, unknown>;
  merges?: unknown;
  chartSheet?: unknown;
}

/**
 * The authority's context for `rect` on `sheet` with nothing else selected
 * and no Sheet View in front: what a headless run (the AI's commands, the
 * Copilot) has. The fields a sheet command's predicate reads are derived as
 * the window derives them; the window's own view state (a selected object, a
 * sorted Sheet View, an import in progress) is absent here by construction.
 */
export function sheetCommandAuthorityContext<S extends ContextSheet>(
  sheet: S,
  rect: SheetCommandRect,
  protection: SheetProtectionReading<S>,
  /** `true` when the page saves back into an imported xlsx (patchback). */
  imported = false,
): SpreadsheetVisibleCommandContext {
  const effective = protection.effective(sheet);
  const locked = protection.rangeHasLockedCell(sheet, rect);
  const shiftFormulas = () =>
    cellShiftRangeSensitiveFormulas(sheet.cells as Record<string, never>);
  let candidates: ReturnType<typeof shiftFormulas> | null = null;
  const shiftSupported = (axis: "down" | "right") => {
    candidates ??= shiftFormulas();
    return canAdjustCellShiftFormulas(candidates, rect, axis, sheet.name);
  };
  const structural = !sheet.chartSheet;
  const merge = sheetMergeSelection(sheet, normalizeMerges(sheet.merges), rect);
  return {
    activeComment: false,
    activeFilter: false,
    activeTable: false,
    canOpenWorkbookWindow: false,
    commentedCellCount: 0,
    fullscreenSupported: false,
    hasMultiColumnSelection: rect.c2 > rect.c1,
    hasMultiRowSelection: rect.r2 > rect.r1,
    hasMultiSelection: false,
    hasSelection: true,
    importInProgress: false,
    nativeWriteRoute: imported ? "patchback" : "from-scratch",
    selectionLockStateAvailable: true,
    selectionLockState: locked,
    selectionEditable: !effective || !locked,
    sheetProtected: effective,
    protectionAllows: sheetCommandPermissions(sheet, protection),
    formatCellsAllowed: sheetCommandPermissions(sheet, protection).formatCells,
    structuralLineEdits: sheetStructuralLineEdits(sheet, rect, protection),
    sortAuthorized: sheetSortAuthorized(sheet, protection, locked),
    mergeSelectionMultiCell: merge.mergeSelectionMultiCell,
    mergeSelectionMultiColumn: merge.mergeSelectionMultiColumn,
    mergeSelectionInTable: merge.mergeSelectionInTable,
    selectionHasMerge: merge.selectionHasMerge,
    homeObjectSelected: false,
    // A chart sheet has no cells to restructure.
    homeStructureControlsAvailable: structural,
    get cellShiftDownAvailable() { return structural && shiftSupported("down"); },
    get cellShiftRightAvailable() { return structural && shiftSupported("right"); },
    get cellShiftUpAvailable() { return structural && shiftSupported("down"); },
    get cellShiftLeftAvailable() { return structural && shiftSupported("right"); },
    tableCapacityAvailable: true,
    versionHistoryAvailable: false,
  } as SpreadsheetVisibleCommandContext;
}

/** The authority's decision for one control: null when it may run. */
export function sheetCommandRefusal(
  controlId: SpreadsheetVisibleControlId,
  context: SpreadsheetVisibleCommandContext,
): SheetCommandRefused | null {
  const decision = spreadsheetVisibleControlDecision(controlId, context);
  return decision.enabled
    ? null
    : refused(decision.reasonKey ?? "spreadsheet.capability.commandUnavailable");
}

// ── Insert / Delete whole rows and columns ──────────────────────────────

/** The control each whole-line edit answers to in the window. */
export const LINE_EDIT_CONTROL: Readonly<
  Record<"insert" | "delete", Readonly<Record<"row" | "col", SpreadsheetVisibleControlId>>>
> = Object.freeze({
  insert: Object.freeze({ row: "XLSCTL.HOME.CELLS.INSERT.ROW", col: "XLSCTL.HOME.CELLS.INSERT.COLUMN" }),
  delete: Object.freeze({ row: "XLSCTL.HOME.CELLS.DELETE.ROW", col: "XLSCTL.HOME.CELLS.DELETE.COLUMN" }),
});

/**
 * The workbook's sheets after a structural edit of the sheet at `index`: the
 * edit committed as the sheet commit commits it, every other sheet following
 * it (its references to the edited sheet, its PivotTables' sources) in the
 * same step.
 */
export function sheetsAfterEdit<S>(
  sheets: readonly S[],
  index: number,
  next: S,
  followOthers: FollowOtherSheets<S> | undefined,
): S[] {
  const before = sheets[index];
  const committed = committedSheetEdit(before as never, next as never) as S;
  return sheets.map((candidate, position) =>
    position === index ? committed : followOthers ? followOthers(candidate) : candidate,
  );
}

export type LineEditOutcome<S> =
  | {
      readonly kind: "done";
      readonly sheet: S;
      /** Lines inserted or deleted (a delete past the stored extent deletes
       *  its stored part). */
      readonly count: number;
      readonly followOthers: FollowOtherSheets<S> | undefined;
    }
  | SheetCommandRefused
  /** Nothing to delete: the lines lie past the stored extent. */
  | { readonly kind: "none" };

/**
 * Insert or delete `requested` whole rows (columns) at `at`: the transaction
 * of Home > Insert / Delete, the header and cell menus, the keys on whole
 * lines, a macro's replay and the AI's row and column commands. The sheet's
 * Insert / Delete rows and columns permissions decide on a protected sheet,
 * and a delete needs every cell of its lines unlocked.
 */
export function lineEditOutcome<S extends StructuralSheet>(
  sheet: S,
  workbook: ReadonlyArray<StructuralSheet>,
  op: "insert" | "delete",
  axis: "row" | "col",
  at: number,
  requested: number,
  protection: SheetProtectionReading<S>,
): LineEditOutcome<S> {
  const permission: SheetProtectionPermission =
    op === "insert"
      ? axis === "row" ? "insertRows" : "insertColumns"
      : axis === "row" ? "deleteRows" : "deleteColumns";
  if (protection.effective(sheet) && !protection.allows(sheet, permission)) {
    return refused("spreadsheet.protectionBlocked");
  }
  if (chartsStillLoading(workbook, sheet.name)) return refused("spreadsheet.chartsLoadingStructureBlocked");
  if (op === "insert") {
    const inserted = insertedLinesSheet(sheet, workbook, axis, at, requested);
    if (typeof inserted === "string") {
      return refused(
        inserted === "three-d"
          ? "spreadsheet.structuralThreeDReferenceBlocked"
          : inserted === "array"
            ? "spreadsheet.arrayPartBlocked"
            : inserted === "pivot"
              ? "spreadsheet.pivotPartBlocked"
              : inserted === "data-table"
              ? "spreadsheet.whatIfTablePartBlocked"
              : inserted === "sheet-end"
                ? "spreadsheet.uiCopy.insertWouldShiftCellsOffSheet"
                : "spreadsheet.uiCopy.insertLineBudgetExceeded",
      );
    }
    return { kind: "done", sheet: inserted.sheet, count: requested, followOthers: inserted.followOthers };
  }
  // Only lines inside the persisted extent exist to delete: the blank lines
  // the grid shows past it are on-demand (deletedLinesSheet).
  const deleted = deletedLinesSheet(sheet, workbook, axis, at, requested);
  if (!deleted) return { kind: "none" };
  const deleting =
    typeof deleted === "string" ? deletableLineCount(sheet, axis, at, requested) : deleted.count;
  // Excel's Delete rows / columns: only lines whose every cell is unlocked.
  if (
    protection.effective(sheet) &&
    !canDeleteLines(
      sheet as never,
      axis,
      at,
      at + deleting - 1,
      axis === "row" ? sheet.columnCount : sheet.rowCount,
    )
  ) {
    return refused("spreadsheet.protectionBlocked");
  }
  if (typeof deleted === "string") {
    return refused(
      deleted === "three-d"
        ? "spreadsheet.structuralThreeDReferenceBlocked"
        : deleted === "array"
          ? "spreadsheet.arrayPartBlocked"
          : deleted === "pivot"
          ? "spreadsheet.pivotPartBlocked"
          : deleted === "data-table"
            ? "spreadsheet.whatIfTablePartBlocked"
            : "spreadsheet.cellShiftRefusedTable",
    );
  }
  return { kind: "done", sheet: deleted.sheet, count: deleted.count, followOthers: deleted.followOthers };
}

// ── Insert / Delete cells (the shift dialog) ────────────────────────────

/** The control each Insert / Delete Cells choice answers to. */
export const CELL_SHIFT_CONTROL: Readonly<Record<ShiftDirection, SpreadsheetVisibleControlId>> = Object.freeze({
  down: "XLSCTL.HOME.CELLS.INSERT.SHIFT_DOWN",
  right: "XLSCTL.HOME.CELLS.INSERT.SHIFT_RIGHT",
  up: "XLSCTL.HOME.CELLS.DELETE.SHIFT_UP",
  left: "XLSCTL.HOME.CELLS.DELETE.SHIFT_LEFT",
});

export type CellShiftOutcome<S> =
  | { readonly kind: "done"; readonly sheet: S; readonly followOthers: FollowOtherSheets<S> | undefined }
  | SheetCommandRefused;

/**
 * Insert or Delete Cells at `range`, shifting `direction`: the dialog's OK.
 * Excel will not push a nonblank cell off the sheet (B331), and answers a
 * shift that would cut a table, a merged cell or a PivotTable with its own
 * message (XLS-A009).
 */
export function cellShiftOutcome<S extends CellShiftButtonSheet>(
  sheet: S,
  workbook: ReadonlyArray<CellShiftButtonSheet>,
  range: SheetCommandRect,
  direction: ShiftDirection,
  limits: { maxRows: number; maxCols: number } = { maxRows: MAX_GRID_ROWS, maxCols: MAX_GRID_COLUMNS },
): CellShiftOutcome<S> {
  if (chartsStillLoading(workbook as ReadonlyArray<{ objects?: ReadonlyArray<unknown> }>, sheet.name)) return refused("spreadsheet.chartsLoadingStructureBlocked");
  if (
    (direction === "down" || direction === "right") &&
    insertCellsShiftDropsCells(sheet.cells, range, direction, limits.maxRows, limits.maxCols)
  ) {
    return refused("spreadsheet.uiCopy.insertWouldShiftCellsOffSheet");
  }
  const shifted = cellShiftedSheet(sheet, workbook, range, direction, limits);
  if (!shifted) {
    const blocker = cellShiftMetadataBlocker(
      sheet as never,
      range,
      direction,
      workbook as never,
    );
    return refused(
      blocker === "merge"
        ? "spreadsheet.cellShiftRefusedMerge"
        : blocker === "table"
          ? "spreadsheet.cellShiftRefusedTable"
          : blocker === "pivot"
            ? "spreadsheet.cellShiftRefusedPivot"
            : "spreadsheet.cellShiftRefusedGeneric",
    );
  }
  return {
    kind: "done",
    sheet: shifted.sheet as S,
    followOthers: shifted.followOthers as unknown as FollowOtherSheets<S> | undefined,
  };
}

// ── Clear Contents ──────────────────────────────────────────────────────

/** The PivotTable reports on the sheet, as the entry point can see them. */
export interface SheetPivotReports {
  /** The report the whole of which `range` is, if any (Delete removes it). */
  entire(range: SheetCommandRect): boolean;
  /** Whether `range` touches any report. */
  intersects(range: SheetCommandRect): boolean;
}

export const NO_PIVOT_REPORTS: SheetPivotReports = Object.freeze({
  entire: () => false,
  intersects: () => false,
});

export type ClearContentsOutcome<S> =
  | { readonly kind: "done"; readonly sheet: S }
  /** The whole of a PivotTable selected: the command removes the PivotTable
   *  (Clear > Remove PivotTable), which the entry point does its own way. */
  | { readonly kind: "remove-pivot" }
  | SheetCommandRefused
  | { readonly kind: "unchanged" };

/**
 * Clear Contents over `range`, whichever way it is asked for (the Delete key,
 * Home > Clear > Clear Contents, the cell and header menus, the touch bar, the
 * AI's clear): a whole PivotTable is removed, a part of one or of a legacy
 * array is refused, then the Clear transaction (`clearedSheetForRange`) runs.
 */
export function clearContentsOutcome<S extends ClearButtonSheet & StructuralSheet>(
  sheet: S,
  range: SheetCommandRect,
  protection: SheetProtectionReading<S>,
  pivots: SheetPivotReports,
  grid: { rowCount: number; columnCount: number },
  /** A row a filter hides. Excel clears only the cells a filter leaves
   *  visible (rows hidden by hand clear with the rest), so the hidden
   *  records keep their values (audit P-FILTER-CLEAR-HIDDEN-VALUES). */
  isRowFiltered?: (row: number) => boolean,
): ClearContentsOutcome<S> {
  if (pivots.entire(range)) return { kind: "remove-pivot" };
  if (pivots.intersects(range)) return refused("spreadsheet.pivotPartEditBlocked");
  // Over part of a legacy array formula's range nothing changes (Excel:
  // "You can't change part of an array"); the whole range clears it.
  if (arrayPartEditRefused(sheet, range)) return refused("spreadsheet.arrayPartBlocked");
  let current = sheet;
  let changed = false;
  const runs = visibleRowRuns(range, isRowFiltered);
  // Over many runs, only those holding a stored cell can change anything.
  const storedRows = runs.length > 1 ? rowsWithCellsIn(sheet.cells, range) : null;
  let stored = 0;
  for (const run of runs) {
    if (storedRows) {
      while (stored < storedRows.length && storedRows[stored] < run.r1) stored += 1;
      if (stored >= storedRows.length || storedRows[stored] > run.r2) continue;
    }
    const part = { ...range, r1: run.r1, r2: run.r2 };
    const base = current;
    const outcome = clearedSheetForRange(base, part, "contents", {
      protectionEffective: protection.effective(base),
      insertHyperlinksAllowed: protection.allows(base, "insertHyperlinks"),
      rangeHasLockedCell: (candidate) => protection.rangeHasLockedCell(base, candidate),
      gridRowCount: grid.rowCount,
      gridColumnCount: grid.columnCount,
    });
    if (outcome === "protected") return refused("spreadsheet.protectionBlocked");
    if (outcome === "what-if-table") return refused("spreadsheet.whatIfTablePartBlocked");
    if (outcome) {
      current = outcome;
      changed = true;
    }
  }
  return changed ? { kind: "done", sheet: current } : { kind: "unchanged" };
}

/** The rows of `range` that hold a stored cell, ascending. */
function rowsWithCellsIn(cells: Readonly<Record<string, unknown>>, range: SheetCommandRect): number[] {
  const rows = new Set<number>();
  for (const key of Object.keys(cells)) {
    const at = key.indexOf(":");
    const row = Number(key.slice(0, at));
    const col = Number(key.slice(at + 1));
    if (row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2) rows.add(row);
  }
  return [...rows].sort((a, b) => a - b);
}

/** The runs of rows of `range` a filter leaves visible, top to bottom; the
 *  whole range when nothing in it is filtered. */
function visibleRowRuns(
  range: SheetCommandRect,
  isRowFiltered: ((row: number) => boolean) | undefined,
): Array<{ r1: number; r2: number }> {
  if (!isRowFiltered) return [{ r1: range.r1, r2: range.r2 }];
  const runs: Array<{ r1: number; r2: number }> = [];
  let start: number | null = null;
  for (let row = range.r1; row <= range.r2; row += 1) {
    if (isRowFiltered(row)) {
      if (start !== null) runs.push({ r1: start, r2: row - 1 });
      start = null;
    } else if (start === null) {
      start = row;
    }
  }
  if (start !== null) runs.push({ r1: start, r2: range.r2 });
  return runs;
}

// ── Sort ────────────────────────────────────────────────────────────────

/**
 * Why `rowSortedSheet` refused a sort, in the window's words: a merge it would
 * tear, part of a legacy array, part of a table. Every refusal the shared row
 * sort returns is one of these three; a caller that let one through wrote a
 * refusal string as a sheet (the gate went red over the missing "table" once).
 */
export function sortRefusalKey(refusal: "merge" | "array" | "table"): SheetCommandRefusalKey {
  return refusal === "array"
    ? "spreadsheet.uiCopy.cantChangePartOfArray"
    : refusal === "table"
      ? "spreadsheet.cellShiftRefusedTable"
      : "spreadsheet.uiCopy.canTSortBecauseMergedCellsWouldBeSplit";
}

/** The control the AI's sort answers to in the window (Data > Sort). */
export const SORT_CONTROL: SpreadsheetVisibleControlId = "XLSCTL.DATA.SORT.ASC";

// ── Merge ───────────────────────────────────────────────────────────────

/** Whether merging `spans` keeps only the upper-left value of one of them
 *  (more than one of its cells holds a value). */
export function mergeSpansDiscardValues(
  cells: Readonly<Record<string, { value?: unknown } | undefined>>,
  spans: readonly MergeSpan[],
): boolean {
  return spans.some(
    (span) => mergeContentCarry(span, (row, col) => cells[`${row}:${col}`]?.value as never).dataCells > 1,
  );
}

/**
 * Whether a merge button on `rect` would keep only the upper-left value of
 * cells that hold more than one: the window asks first ("Merging cells only
 * keeps the upper-left value"), and an entry point that cannot ask refuses
 * with that question's key.
 */
export function mergeDiscardsValues(
  sheet: { cells: Readonly<Record<string, { value?: unknown } | undefined>>; merges?: unknown },
  rect: SheetCommandRect,
  button: "merge_center" | "merge_across" | "merge_cells" | "unmerge",
): boolean {
  if (button === "unmerge") return false;
  const spans = normalizeMerges(sheet.merges);
  const target = expandRangeToMergeClosure(rect, spans);
  // Merge & Center on a merged selection unmerges it.
  if (button === "merge_center" && rangeFullyMerged(spans, target)) return false;
  const planned =
    button === "merge_across"
      ? mergeAcrossSpans(target)
      : [mergeSpanFor(target)].filter((span): span is MergeSpan => span !== null);
  return mergeSpansDiscardValues(sheet.cells, planned);
}
