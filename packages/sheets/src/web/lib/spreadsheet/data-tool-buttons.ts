/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Data Tools and Outline: Remove Duplicates' OK, Subtotal's OK, and
 * the outline writes (Group, Ungroup, Auto Outline, show/hide detail) as pure
 * sheet transactions. Moved out of SpreadsheetWindow so the ribbon and the
 * spreadsheet editor commands run the same code; behaviour is unchanged. The
 * window keeps its dialog authority, toasts and status lines.
 */

import { writeCellValueInto, type CellMap, type SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  normalizeDataTables,
  tableTotalCellOwnsValue,
  tableTotalCellValue,
  type DataTable,
} from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { rangeHasMerge } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import {
  autoOutlineGroups,
  createGroup,
  type OutlineGroup,
  type OutlineSummarySettings,
} from "@mrnote/core/web/lib/spreadsheet/outline-groups.ts";
import { removeDuplicateRows } from "./remove-duplicates";
import { relocateSortedCell } from "./sort";
import type { SortFilterRow } from "./sort-filter";
import {
  deleteLine,
  insertLine,
  structuralInsertRefusal,
  type StructuralLineRange,
  type StructuralSheet,
} from "@mrnote/core/web/lib/spreadsheet/structural.ts";
import { planSubtotals, type SubtotalFunction } from "./subtotal";
import {
  tableHeadersForCells,
  tableRangeResizeBlocked,
  type LockedCellTest,
  type TableButtonSheet,
} from "./table-buttons";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** What a cell shows (formulas resolved, number formats applied). */
export type CellDisplay = (cell: SpreadsheetCell | undefined, row: number, col: number) => string;

/** Why Remove Duplicates' OK refused. */
export type RemoveDuplicatesRefusal =
  | "noData"
  | "tableMismatch"
  | "tableBlocked"
  | "customTotals"
  | "protected"
  | "noColumns"
  | "merge"
  | "spill"
  | "outlined";

export interface RemovedDuplicates<S> {
  /** The sheet after the command; null when no duplicate was found. */
  sheet: S | null;
  removed: number;
  unique: number;
}

/**
 * Remove Duplicates' OK on `range`: rows whose values in `keyColumns`
 * (absolute columns, compared as displayed, case-insensitively) repeat an
 * earlier row are removed and the kept rows close up inside the range; a
 * header row stays. Over a Table the range must be the table's own, and its
 * Totals Row follows the kept rows up.
 */
export function removeDuplicatesSheet<S extends TableButtonSheet>(
  sheet: S,
  request: { range: Rect; hasHeader: boolean; keyColumns: ReadonlySet<number> },
  context: {
    display: CellDisplay;
    /** A dynamic array spills into the cell. */
    spilled: (row: number, col: number) => boolean;
    rangeHasLockedCell: LockedCellTest<S>;
  },
): RemovedDuplicates<S> | RemoveDuplicatesRefusal {
  const { r1, r2, c1, c2 } = request.range;
  // A header row (when toggled) is kept verbatim and excluded from the dedup:
  // the data starts one row below, so the header is never moved or compared.
  const dataR1 = request.hasHeader ? r1 + 1 : r1;
  if (dataR1 > r2) return "noData";
  const dataTables = normalizeDataTables(sheet.dataTables, {
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  });
  const overlapsTableRange = (table: DataTable) =>
    table.range.r1 <= r2 && table.range.r2 >= r1 && table.range.c1 <= c2 && table.range.c2 >= c1;
  const exactlyMatchesTableRange = (table: DataTable) =>
    table.range.r1 === r1 && table.range.r2 === r2 && table.range.c1 === c1 && table.range.c2 === c2;
  const overlappingTables = dataTables.filter(overlapsTableRange);
  const matchingTable = overlappingTables.find(exactlyMatchesTableRange);
  if (overlappingTables.length > 0 && (matchingTable === undefined || overlappingTables.length > 1)) {
    return "tableMismatch";
  }
  if (
    matchingTable &&
    (!request.hasHeader || tableRangeResizeBlocked(sheet, matchingTable, { totalsRowMovable: true }))
  ) {
    return "tableBlocked";
  }
  // A table's Totals Row stays under the rows that remain (Excel keeps it
  // through Remove Duplicates): every totals cell must still be the
  // table's own recipe so it can be rewritten one row up per removed row.
  const matchingTableHeaders = matchingTable ? tableHeadersForCells(matchingTable, sheet.cells) : null;
  if (matchingTable?.totalsRow) {
    const totalsRow = matchingTable.range.r2 + 1;
    if (
      !matchingTableHeaders ||
      matchingTableHeaders.some(
        (header, offset) =>
          !tableTotalCellOwnsValue(
            matchingTable,
            offset,
            header,
            String(sheet.cells[`${totalsRow}:${c1 + offset}`]?.value ?? ""),
          ),
      )
    ) {
      return "customTotals";
    }
    if (context.rangeHasLockedCell(sheet, { r1: totalsRow, r2: totalsRow, c1, c2 })) return "protected";
  }
  // The key columns: the checked subset, expressed as offsets into the selection
  // so they line up with the pure-row arrays (index 0 = first selected column).
  const keyOffsets: number[] = [];
  for (let c = c1; c <= c2; c += 1) {
    if (request.keyColumns.has(c)) keyOffsets.push(c - c1);
  }
  if (keyOffsets.length === 0) return "noColumns";
  // Excel: "You cannot remove duplicate values from outline data that is
  // outlined or that has subtotals. To remove duplicates, you must remove
  // both the outline and the subtotals." Any row group over the data rows,
  // or column group over the range's columns, refuses the run.
  const outlineGroups = (sheet as { outlineGroups?: unknown }).outlineGroups;
  if (
    Array.isArray(outlineGroups) &&
    (outlineGroups as Partial<OutlineGroup>[]).some((group) =>
      typeof group?.start === "number" &&
      typeof group.end === "number" &&
      (group.axis === "row"
        ? group.start <= r2 && group.end >= dataR1
        : group.axis === "col" && group.start <= c2 && group.end >= c1),
    )
  ) {
    return "outlined";
  }
  // A partial-range collapse cannot preserve a merged rectangle's geometry.
  if (sheet.merges && rangeHasMerge(sheet.merges, { r1: dataR1, r2, c1, c2 })) return "merge";
  // A dynamic array's spilled cells are not the command's to clear or
  // reorder, and their values are what they show, not blanks (live audit
  // B158): a selection over a spill is refused before anything is written.
  for (let r = dataR1; r <= r2; r += 1) {
    for (let c = c1; c <= c2; c += 1) {
      if (context.spilled(r, c)) return "spill";
    }
  }
  // Build the pure rows from the RESOLVED display values (so equality matches
  // what the user sees / what AutoFilter compares), tagging each with its source
  // row index so the kept rows can be rewritten with full cell fidelity.
  const indexRows: SortFilterRow[] = [];
  const sourceRowOf: number[] = [];
  for (let r = dataR1; r <= r2; r += 1) {
    const cellsRow: Array<{ value: string } | undefined> = [];
    for (let c = c1; c <= c2; c += 1) {
      const cell = sheet.cells[`${r}:${c}`];
      cellsRow.push(cell ? { value: context.display(cell, r, c) } : undefined);
    }
    indexRows.push(cellsRow);
    sourceRowOf.push(r);
  }
  const { keptRows, removedCount } = removeDuplicateRows(
    indexRows.map((row, i) => [...row, { value: String(sourceRowOf[i]) }]),
    keyOffsets,
  );
  if (removedCount === 0) return { sheet: null, removed: 0, unique: keptRows.length };
  // The source row index lives at the last appended cell of each kept row.
  const keptSourceRows = keptRows.map((row) => Number(row[row.length - 1]?.value));
  // Sheet protection: authenticate exactly the selected rectangle that will be
  // rewritten, rather than unrelated cells elsewhere on the same worksheet row.
  if (context.rangeHasLockedCell(sheet, { r1: dataR1, r2, c1, c2 })) return "protected";
  // Rectangle-only rewrite within [dataR1, r2] x [c1, c2]: clear the selected
  // cells first, then place each kept source row at the top in first-seen order.
  // Full cell objects travel, so styles/comments/number formats survive; formulas
  // rebase by their row delta. Everything outside the selected columns is exact.
  const nextCells: CellMap = { ...sheet.cells };
  for (let r = dataR1; r <= r2; r += 1) {
    for (let c = c1; c <= c2; c += 1) {
      delete nextCells[`${r}:${c}`];
    }
  }
  keptSourceRows.forEach((srcRow, offset) => {
    const destRow = dataR1 + offset;
    for (let c = c1; c <= c2; c += 1) {
      const cell = sheet.cells[`${srcRow}:${c}`];
      if (cell !== undefined) {
        nextCells[`${destRow}:${c}`] = relocateSortedCell(cell, destRow - srcRow, 0);
      }
    }
  });
  const nextTableRange = matchingTable
    ? { ...matchingTable.range, r2: matchingTable.range.r2 - removedCount }
    : null;
  const nextTable = matchingTable && nextTableRange ? { ...matchingTable, range: nextTableRange } : null;
  if (nextTable?.totalsRow && matchingTable && matchingTableHeaders) {
    // Move the physical Totals Row up under the kept rows: the cell (its
    // number format included) travels, and its total is rewritten from the
    // table's recipe in structured form, so an imported A1 span does not
    // keep pointing at rows the table no longer has.
    const fromRow = matchingTable.range.r2 + 1;
    const toRow = nextTable.range.r2 + 1;
    for (let offset = 0; offset < matchingTableHeaders.length; offset += 1) {
      const col = c1 + offset;
      const cell = sheet.cells[`${fromRow}:${col}`];
      delete nextCells[`${fromRow}:${col}`];
      if (cell) nextCells[`${toRow}:${col}`] = { ...cell };
      writeCellValueInto(nextCells, toRow, col, tableTotalCellValue(nextTable, offset, matchingTableHeaders[offset]));
    }
  }
  const nextDataTables = nextTable
    ? dataTables.map((table) => (table.id === nextTable.id ? nextTable : table))
    : dataTables;
  return {
    sheet: { ...sheet, cells: nextCells, ...(nextTableRange ? { dataTables: nextDataTables } : {}) },
    removed: removedCount,
    unique: keptRows.length,
  };
}

/** Why Subtotal's OK refused. */
export type SubtotalRefusal = "needColumns" | "noData" | "protected" | "tooManyRows";

/** The sheet parts Subtotal reads and writes. */
export type SubtotalSheet = TableButtonSheet & StructuralSheet & { outlineGroups?: OutlineGroup[]; outlineSummaryBelow?: boolean };

/**
 * Subtotal's OK over `range`: a SUBTOTAL row after each run of equal values
 * in `groupColumn` (absolute), totalling `valueColumns` with `fn`, a grand
 * total at the bottom, and outline groups so each block collapses under its
 * total. The summary rows are inserted as Insert Rows does, so every
 * formula, merge, table, chart and group on the sheet follows them.
 */
export function subtotaledSheet<S extends SubtotalSheet>(
  original: S,
  request: {
    range: Rect;
    hasHeader: boolean;
    groupColumn: number | null;
    valueColumns: readonly number[];
    fn: SubtotalFunction;
    /** "{group} Total": the group's value replaces {group}. */
    labelTemplate: string;
    grandTotalLabel: string;
    /** "Replace current subtotals" (default true; XLS-A064). */
    replace?: boolean;
    /** "Page break between groups" (default false; XLS-A064). */
    pageBreaks?: boolean;
    /** "Summary below data" (default true; XLS-A064). */
    summaryBelow?: boolean;
  },
  context: { display: CellDisplay; rangeHasLockedCell: LockedCellTest<S>; maxRows: number },
): { sheet: S; inserted: number } | SubtotalRefusal {
  if (request.groupColumn === null || request.valueColumns.length === 0) return "needColumns";
  const summaryBelow = request.summaryBelow !== false;
  // Protection: a subtotal rewrites the region + the rows below it shift down,
  // so refuse on a protected sheet with any locked cell in or below the region.
  {
    const firstData = request.hasHeader ? request.range.r1 + 1 : request.range.r1;
    const { c1, c2 } = request.range;
    if (context.rangeHasLockedCell(original, { r1: firstData, r2: original.rowCount - 1, c1, c2 })) {
      return "protected";
    }
  }
  // "Replace current subtotals" (Excel's default): the list's earlier
  // subtotals go first, exactly as Remove All takes them (XLS-A064).
  let sheet: S = original;
  let rangeBottom = request.range.r2;
  if (request.replace !== false) {
    const removed = removeSubtotalsSheet(original, request.range, context);
    if (typeof removed === "object") {
      sheet = removed.sheet;
      rangeBottom -= removed.removed;
    } else if (removed !== "none") {
      return removed === "protected" ? "protected" : "noData";
    }
  }
  // The header row (when the list has one) stays where it is, above the
  // detail rows, as Excel leaves it.
  const r1 = request.hasHeader ? request.range.r1 + 1 : request.range.r1;
  const r2 = Math.min(rangeBottom, sheet.rowCount - 1);
  const { c1, c2 } = request.range;
  if (r2 < r1) return "noData";
  // Kept (not replaced) subtotals stay put and the new ones nest inside them.
  const existingSummaryRows = new Set<number>();
  if (request.replace === false) {
    for (let r = r1; r <= r2; r += 1) {
      if (rowHoldsSubtotal(sheet, r, c1, c2)) existingSummaryRows.add(r - r1);
    }
  }
  // Read the region's DISPLAY values (formulas resolved) as the detail matrix.
  const dataRows: string[][] = [];
  for (let r = r1; r <= r2; r += 1) {
    const row: string[] = [];
    for (let c = c1; c <= c2; c += 1) {
      row.push(context.display(sheet.cells[`${r}:${c}`], r, c));
    }
    dataRows.push(row);
  }
  // The helper works in column offsets RELATIVE to the region's left edge.
  const plan = planSubtotals(dataRows, {
    groupColumn: request.groupColumn - c1,
    valueColumns: request.valueColumns.map((c) => c - c1),
    function: request.fn,
    firstDataRow: r1,
    firstDataColumn: c1,
    labelTemplate: request.labelTemplate,
    grandTotalLabel: request.grandTotalLabel,
    grandTotal: true,
    summaryBelow,
    existingSummaryRows,
    pageBreaks: request.pageBreaks === true,
  });
  if (plan.insertedCount === 0) return "noData";
  // The rebuilt region is TALLER than the original (inserted summary rows). It
  // grows the sheet; rows that were below the region shift DOWN by the number
  // of inserted rows so they are not overwritten.
  const grew = plan.rows.length - dataRows.length;
  const newRowCount = sheet.rowCount + grew;
  if (newRowCount > context.maxRows) return "tooManyRows";
  // 1) Insert the summary rows as Insert Rows does, one block at a time from
  //    the bottom up, so every formula, merge, table, chart and outline
  //    group on the sheet is rewritten for the shift: a detail =B92*2 that
  //    moves to row 93 reads B93, and an E90=C92 outside the region follows
  //    its cell to C93 (live audit B104). Each block's `at` is the ORIGINAL
  //    row the block goes in front of, valid because later blocks are
  //    inserted first.
  const blocks: Array<{ at: number; count: number }> = [];
  let detailSeen = 0;
  for (let i = 0; i < plan.rows.length; i += 1) {
    if (plan.rowSources[i] !== null) {
      detailSeen += 1;
      continue;
    }
    const at = r1 + detailSeen;
    const last = blocks[blocks.length - 1];
    if (last && last.at === at) last.count += 1;
    else blocks.push({ at, count: 1 });
  }
  let structural: S = sheet;
  for (const block of [...blocks].reverse()) {
    const refusal = structuralInsertRefusal(structural, "row", block.count);
    const inserted = refusal ? structural : insertLine(structural, "row", block.at, block.count);
    if (inserted === structural) return "tooManyRows";
    structural = inserted;
  }
  const nextCells: CellMap = { ...structural.cells };
  // 2) Write ONLY the inserted SUMMARY rows (rowSources[i] === null) from the
  //    plan — the SUBTOTAL formulas + the group/grand-total labels. The DETAIL
  //    rows were already relocated above with their original content.
  for (let i = 0; i < plan.rows.length; i += 1) {
    if (plan.rowSources[i] !== null) continue;
    const row = plan.rows[i];
    for (let j = 0; j < row.length; j += 1) {
      writeCellValueInto(nextCells, r1 + i, c1 + j, row[j]);
    }
  }
  // 3) Outline groups: the plan's groups are relative to the first data row, so
  //    shift each by r1 to get absolute sheet rows.
  //    The sheet's other groups (shifted by the inserts) stay; the new ones
  //    nest inside any kept subtotal's groups.
  const groups = createGroup(
    structural.outlineGroups ?? [],
    plan.groups.map((g) => ({ start: g.start + r1, end: g.end + r1 })),
    "row",
  );
  const sheetNext: S = {
    ...structural,
    cells: nextCells,
    rowCount: Math.max(structural.rowCount, newRowCount),
  };
  if (plan.groups.length > 0) {
    sheetNext.outlineGroups = groups;
    // "Summary below data" also sets the sheet's outline setting, so each
    // toggle sits on its total row: below the detail (Excel's default, never
    // stored) or above it when the box is cleared.
    if (summaryBelow) delete sheetNext.outlineSummaryBelow;
    else sheetNext.outlineSummaryBelow = false;
  }
  // 4) "Page break between groups": a manual break before every group after
  //    the first (XLS-A064).
  if (plan.pageBreakOffsets.length > 0) {
    const breaks = new Set(structural.printSettings?.manualRowBreaks ?? []);
    for (const offset of plan.pageBreakOffsets) breaks.add(r1 + offset);
    sheetNext.printSettings = {
      ...(structural.printSettings ?? {}),
      manualRowBreaks: [...breaks].sort((a, b) => a - b),
    };
  }
  return { sheet: sheetNext, inserted: plan.insertedCount };
}

/** A formula cell that is an outline subtotal (Subtotal writes these). */
const SUBTOTAL_FORMULA_RE = /^=\s*SUBTOTAL\s*\(/i;

/** Whether row `row` holds a SUBTOTAL formula between columns c1..c2. */
function rowHoldsSubtotal(sheet: SubtotalSheet, row: number, c1: number, c2: number): boolean {
  for (let c = c1; c <= c2; c += 1) {
    const cell = sheet.cells[`${row}:${c}`] as { value?: unknown; formula?: unknown } | undefined;
    if (!cell) continue;
    const formula = typeof cell.formula === "string" ? cell.formula : cell.value;
    if (typeof formula === "string" && SUBTOTAL_FORMULA_RE.test(formula)) return true;
  }
  return false;
}

/**
 * Subtotal's "Remove All" over the list `range` (XLS-A064): every row of the
 * list holding a SUBTOTAL formula is deleted as Delete Rows does (formulas,
 * merges, tables and charts follow), and with them the list's row outline and
 * the manual page breaks inside it: Excel's Remove All takes the outline and
 * the "Page break between groups" breaks it built along with the summary
 * rows. "none" when the list has nothing to remove.
 */
export function removeSubtotalsSheet<S extends SubtotalSheet>(
  sheet: S,
  range: Rect,
  context: { rangeHasLockedCell: LockedCellTest<S> },
): { sheet: S; removed: number } | "none" | "protected" | "blocked" {
  const r1 = Math.max(0, range.r1);
  const r2 = Math.min(range.r2, sheet.rowCount - 1);
  if (r2 < r1) return "none";
  const rows: number[] = [];
  for (let r = r1; r <= r2; r += 1) {
    if (rowHoldsSubtotal(sheet, r, range.c1, range.c2)) rows.push(r);
  }
  const listEnd = r2 - rows.length;
  const intersectsList = (group: StructuralLineRange) =>
    group.axis === "row" && group.start <= listEnd && group.end >= r1;
  const breakInList = (line: number) => line > r1 && line <= listEnd;
  if (
    rows.length === 0 &&
    !(sheet.outlineGroups ?? []).some(intersectsList) &&
    !(sheet.printSettings?.manualRowBreaks ?? []).some(breakInList)
  ) {
    return "none";
  }
  if (rows.length > 0 && context.rangeHasLockedCell(sheet, { r1, r2: sheet.rowCount - 1, c1: range.c1, c2: range.c2 })) {
    return "protected";
  }
  let next: S = sheet;
  // Bottom-up, so each deletion leaves the rows above it where they were.
  for (const row of [...rows].reverse()) {
    const deleted = deleteLine(next, "row", row, 1);
    if (deleted === next) return "blocked";
    next = deleted;
  }
  const groups: OutlineGroup[] = next.outlineGroups ?? [];
  const outline = groups.filter((group) => !intersectsList(group));
  const result: S = { ...next };
  if (outline.length > 0) result.outlineGroups = outline;
  else delete result.outlineGroups;
  const breaks = next.printSettings?.manualRowBreaks;
  if (breaks && breaks.some(breakInList)) {
    const kept = breaks.filter((line) => !breakInList(line));
    const printSettings = { ...next.printSettings };
    if (kept.length > 0) printSettings.manualRowBreaks = kept;
    else delete printSettings.manualRowBreaks;
    result.printSettings = printSettings;
  }
  return { sheet: result, removed: rows.length };
}

/** The sheet parts the outline writes touch. */
export interface OutlineSheet extends OutlineSummarySettings {
  rowCount: number;
  columnCount: number;
  cells: Readonly<Record<string, { value?: string }>>;
  outlineGroups?: OutlineGroup[];
  outlineGroupsNativeBlocker?: string;
}

/**
 * The sheet with its outline replaced by `next` (empty drops it). "native"
 * when an imported sheet's outline cannot be written back exactly (the
 * window asks this only on a page saved through the native writer).
 */
export function outlineGroupsSheet<S extends OutlineSheet>(
  sheet: S,
  next: OutlineGroup[],
  nativeBlocked: boolean,
): S | "native" {
  if (nativeBlocked && sheet.outlineGroupsNativeBlocker) return "native";
  const sheetNext: S = { ...sheet };
  if (next.length > 0) sheetNext.outlineGroups = next;
  else delete sheetNext.outlineGroups;
  return sheetNext;
}

/**
 * Auto Outline: groups from the SUM/SUBTOTAL summary cells that sit directly
 * below a vertical block or right of a horizontal block (where the Outline
 * Settings put summaries), folded into the existing groups by the same
 * nesting rules. Null when nothing new is found.
 */
export function autoOutlinedGroups(
  sheet: OutlineSheet,
  outlineGroups: readonly OutlineGroup[],
): OutlineGroup[] | null {
  const auto = autoOutlineGroups({
    rowCount: Number(sheet.rowCount) || 0,
    columnCount: Number(sheet.columnCount) || 0,
    formulaAt: (row, col) => sheet.cells[`${row}:${col}`]?.value,
    // Excel's Auto Outline reads summaries where the Outline Settings put
    // them: above / left of the detail once those boxes are cleared.
    settings: sheet,
  });
  if (auto.length === 0) return null;
  let next = [...outlineGroups];
  for (const group of auto) {
    next = createGroup(next, [{ start: group.start, end: group.end }], group.axis);
  }
  if (next.length === outlineGroups.length) return null; // nothing new added
  return next;
}
