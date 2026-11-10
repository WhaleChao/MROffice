/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Advanced as a pure sheet transaction: the list and criteria ranges
 * read as the grid shows them (numbers stay numbers, formula criteria
 * evaluated per record), then either the matching records copied to another
 * place on the sheet (formats and formulas travel, rebased) or Advanced
 * Filter's own hidden rows stored in place. Moved out of SpreadsheetWindow so
 * the dialog and the spreadsheet editor commands run the same code; behaviour
 * is unchanged. The window keeps the mode authority (protection, the xlsx
 * write route), the dialog text and the commit.
 */

import type { CellMap, SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  applyAdvancedFilter,
  MAX_CRITERIA_ROWS,
  type FilterCellValue,
} from "./advanced-filter";
import { rangesOverlap } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { spreadsheetCellIsFormula } from "@mrnote/core/web/lib/spreadsheet/display-cache.ts";
import { spreadsheetCellIsText } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import { columnToLetters, type CellValue } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { rangeHasMerge, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { relocateSortedCell } from "./sort";

type Rect = { r1: number; c1: number; r2: number; c2: number };

export interface AdvancedFilterSheet {
  /** The sheet's tab name: a range qualified by it ("Sheet1!$A$1:$C$6") is
   * this sheet's range. */
  name?: string;
  cells: CellMap;
  rowCount: number;
  columnCount: number;
  merges?: MergeSpan[];
  /** The in-place result, with the criteria range it ran with (Excel's
   * sheet-scoped Criteria name). */
  advancedFilterState?: { ref: string; hiddenRows: number[]; criteria?: string };
  autoFilterRef?: string;
  autoFilterId?: string;
  namedSheetViewsOwnsAutoFilter?: true;
  namedSheetViews?: ReadonlyArray<{ filterId?: string }>;
}

export interface AdvancedFilterPlanRequest {
  mode: "in-place" | "copy";
  list: string;
  criteria: string;
  copyTo: string;
  uniqueOnly: boolean;
}

export type AdvancedFilterPlanError =
  | { kind: "ranges" }
  | { kind: "headers" }
  | { kind: "invalid-criteria"; criteria: string[] }
  | { kind: "missing-headers"; headers: string[] }
  /** More condition rows than the filter compiles: refused whole (O013). */
  | { kind: "criteria-rows"; max: number }
  | { kind: "copy-to" }
  | { kind: "copy-to-columns" }
  | { kind: "copy-header-blank" }
  | { kind: "copy-header"; header: string }
  | { kind: "copy-destination" }
  | { kind: "copy-protected" }
  | { kind: "protected" }
  /** Named Sheet Views filter through the AutoFilter the filter would replace. */
  | { kind: "named-sheet-views" };

/**
 * A criteria cell's formula (Excel's computed criteria), or null. Only a cell
 * that IS a formula by the engine's own rule: text that begins with "=" (an
 * apostrophe entry, a string from the file, the Text format) is a value to
 * match, never a formula to run on every record (O020).
 */
function criteriaCellFormula(cell: SpreadsheetCell | undefined): string | null {
  if (!spreadsheetCellIsFormula(cell)) return null;
  return typeof cell?.formula === "string" && cell.formula.trim().startsWith("=")
    ? cell.formula.trim()
    : typeof cell?.value === "string" && cell.value.trim().startsWith("=")
      ? cell.value.trim()
      : null;
}

/**
 * The criteria formulas a plan evaluates, each over the list's records in its
 * criteria column and anchored at the first record, as `read.evaluateFormula`
 * is asked for them: for a caller that has them evaluated ahead (C2, the
 * calculation session). Empty when the ranges are unusable.
 */
export function advancedFilterCriteriaFormulas(
  sheet: Pick<AdvancedFilterSheet, "name" | "cells" | "rowCount" | "columnCount">,
  request: Pick<AdvancedFilterPlanRequest, "list" | "criteria">,
): Array<{ formula: string; anchorRow: number; anchorCol: number; range: Rect }> {
  const list = parseAdvancedFilterRange(request.list, sheet.name);
  const criteria = parseAdvancedFilterRange(request.criteria, sheet.name);
  if (!list || !criteria || list.r2 <= list.r1 || criteria.r2 <= criteria.r1 || list.r2 >= sheet.rowCount || criteria.r2 >= sheet.rowCount ||
      criteria.c2 >= sheet.columnCount || list.r1 < 0 || criteria.r1 < 0 || criteria.c1 < 0) return [];
  const out = new Map<string, { formula: string; anchorRow: number; anchorCol: number; range: Rect }>();
  for (let row = criteria.r1 + 1; row <= criteria.r2; row += 1) {
    for (let col = criteria.c1; col <= criteria.c2; col += 1) {
      const formula = criteriaCellFormula(sheet.cells[`${row}:${col}`]);
      if (formula) out.set(`${col}\u0000${formula}`, { formula, anchorRow: list.r1 + 1, anchorCol: col, range: { r1: list.r1 + 1, c1: col, r2: list.r2, c2: col } });
    }
  }
  return [...out.values()];
}

/**
 * A range as Excel's Advanced Filter boxes take it: "A1:C6", absolute
 * "$A$1:$C$6" (how Excel fills the List range), or qualified by this sheet's
 * name ("Sheet1!$A$1:$C$6", "'Q1 data'!A1:C6", how a pick on the sheet
 * writes it). A range qualified by another sheet's name is not this sheet's.
 */
export function parseAdvancedFilterRange(text: string, sheetName?: string): Rect | null {
  let ref = text.trim();
  const bang = ref.lastIndexOf("!");
  if (bang >= 0) {
    let qualifier = ref.slice(0, bang).trim();
    if (qualifier.length >= 2 && qualifier.startsWith("'") && qualifier.endsWith("'")) {
      qualifier = qualifier.slice(1, -1).replace(/''/g, "'");
    }
    if (!sheetName || qualifier.toLocaleLowerCase() !== sheetName.toLocaleLowerCase()) return null;
    ref = ref.slice(bang + 1).trim();
  }
  return parseA1Range(ref.replace(/\$/g, ""));
}

function rangeA1(range: Rect): string {
  const start = `${columnToLetters(range.c1)}${range.r1 + 1}`;
  const end = `${columnToLetters(range.c2)}${range.r2 + 1}`;
  return start === end ? start : `${start}:${end}`;
}

/**
 * Advanced Filter. `read` is how the grid shows a cell, a formula's result,
 * a criteria formula's verdict for one record, whether a cell may be written
 * and whether it lies in a dynamic array's spill. A copy that would write
 * nothing new (a pre-seeded header and no match) leaves the sheet as it is.
 */
export function advancedFilterPlan<S extends AdvancedFilterSheet>(
  sheet: S,
  request: AdvancedFilterPlanRequest,
  read: {
    display: (cell: SpreadsheetCell | undefined, row: number, col: number) => string;
    computed: (row: number, col: number) => CellValue;
    evaluateFormula: (formula: string, row: number, col: number, anchorRow: number, anchorCol: number) => CellValue;
    isEditable: (row: number, col: number) => boolean;
    spilled: (row: number, col: number) => boolean;
    sheetProtected: boolean;
  },
): { ok: false; error: AdvancedFilterPlanError } | { ok: true; sheet: S; matched: number; duplicatesRemoved: number } {
  const list = parseAdvancedFilterRange(request.list, sheet.name);
  const criteria = parseAdvancedFilterRange(request.criteria, sheet.name);
  const inBounds = (range: Rect) =>
    range.r1 >= 0 &&
    range.r2 >= range.r1 &&
    range.r2 < sheet.rowCount &&
    range.c1 >= 0 &&
    range.c2 >= range.c1 &&
    range.c2 < sheet.columnCount;
  if (
    !list ||
    !criteria ||
    !inBounds(list) ||
    !inBounds(criteria) ||
    list.r2 <= list.r1 ||
    criteria.r2 <= criteria.r1 ||
    rangesOverlap(list, criteria)
  ) {
    return { ok: false, error: { kind: "ranges" } };
  }
  // Every condition row is ORed in; one the filter would not compile is
  // refused before anything is copied or hidden, never dropped (O013).
  if (criteria.r2 - criteria.r1 > MAX_CRITERIA_ROWS) {
    return { ok: false, error: { kind: "criteria-rows", max: MAX_CRITERIA_ROWS } };
  }
  const readRow = (row: number, c1: number, c2: number) => {
    const values: string[] = [];
    for (let col = c1; col <= c2; col += 1) values.push(read.display(sheet.cells[`${row}:${col}`], row, col));
    return values;
  };
  // A record's numbers stay numbers, however they are formatted: a
  // 1,234.56 amount is greater than 1000 (live audit B103). Text and
  // errors are read as they are shown.
  const readTypedRow = (row: number, c1: number, c2: number): FilterCellValue[] => {
    const values: FilterCellValue[] = [];
    for (let col = c1; col <= c2; col += 1) {
      const cell = sheet.cells[`${row}:${col}`];
      if (cell && spreadsheetCellIsFormula(cell)) {
        const computed = read.computed(row, col);
        if (typeof computed === "number" || typeof computed === "boolean") {
          values.push(computed);
          continue;
        }
      } else if (cell && cell.kind === "number" && !spreadsheetCellIsText(cell)) {
        const number = Number(cell.value);
        if (cell.value !== "" && Number.isFinite(number)) {
          values.push(number);
          continue;
        }
      }
      values.push(read.display(cell, row, col));
    }
    return values;
  };
  const headers = readRow(list.r1, list.c1, list.c2);
  if (headers.some((header) => header.trim() === "")) return { ok: false, error: { kind: "headers" } };
  const rows: FilterCellValue[][] = [];
  for (let row = list.r1 + 1; row <= list.r2; row += 1) rows.push(readTypedRow(row, list.c1, list.c2));
  const criteriaRows: FilterCellValue[][] = [];
  const criteriaFormulas: Array<Array<string | null>> = [];
  for (let row = criteria.r1 + 1; row <= criteria.r2; row += 1) {
    criteriaRows.push(readTypedRow(row, criteria.c1, criteria.c2));
    const formulas: Array<string | null> = [];
    for (let col = criteria.c1; col <= criteria.c2; col += 1) {
      formulas.push(criteriaCellFormula(sheet.cells[`${row}:${col}`]));
    }
    criteriaFormulas.push(formulas);
  }
  const result = applyAdvancedFilter({
    headers,
    rows,
    criteria: {
      headers: readRow(criteria.r1, criteria.c1, criteria.c2),
      rows: criteriaRows,
      formulas: criteriaFormulas,
    },
    uniqueOnly: request.uniqueOnly,
    evaluateFormulaCriterion: (formula, position) => {
      const value = read.evaluateFormula(
        formula,
        list.r1 + 1 + position.dataRow,
        criteria.c1 + position.criteriaColumn,
        list.r1 + 1,
        criteria.c1 + position.criteriaColumn,
      );
      return typeof value === "boolean" ? value : null;
    },
  });
  if (result.invalidCriteria.length > 0) {
    return { ok: false, error: { kind: "invalid-criteria", criteria: result.invalidCriteria } };
  }
  if (result.unmatchedHeaders.length > 0) {
    return { ok: false, error: { kind: "missing-headers", headers: result.unmatchedHeaders } };
  }
  const outcome = { matched: result.matched.length, duplicatesRemoved: result.duplicatesRemoved };
  if (request.mode === "copy") {
    const copyTo = parseAdvancedFilterRange(request.copyTo, sheet.name);
    if (!copyTo || copyTo.r1 !== copyTo.r2 || !inBounds(copyTo)) return { ok: false, error: { kind: "copy-to" } };
    // Excel lets Copy to point either at one blank anchor (copy every list
    // column) or at a pre-seeded row of list headers (copy/reorder only
    // those columns). The seeded header cells themselves stay as they are:
    // users commonly format them before running Advanced Filter.
    const requestedLabels = readRow(copyTo.r1, copyTo.c1, copyTo.c2).map((header) => header.trim());
    const hasRequestedHeader = requestedLabels.some(Boolean);
    let preserveDestinationHeaders = false;
    let sourceColumnOffsets: number[];
    if (!hasRequestedHeader) {
      if (copyTo.c1 !== copyTo.c2) return { ok: false, error: { kind: "copy-to-columns" } };
      sourceColumnOffsets = headers.map((_header, index) => index);
    } else {
      if (requestedLabels.some((label) => label === "")) return { ok: false, error: { kind: "copy-header-blank" } };
      const usedSourceColumns = new Set<number>();
      sourceColumnOffsets = [];
      for (const label of requestedLabels) {
        const wanted = label.toLowerCase();
        const matchingColumns = headers.flatMap((header, index) =>
          header.trim().toLowerCase() === wanted ? [index] : [],
        );
        const matchingColumn = matchingColumns[0];
        if (matchingColumns.length !== 1 || matchingColumn === undefined || usedSourceColumns.has(matchingColumn)) {
          return { ok: false, error: { kind: "copy-header", header: label } };
        }
        usedSourceColumns.add(matchingColumn);
        sourceColumnOffsets.push(matchingColumn);
      }
      preserveDestinationHeaders = true;
    }
    const width = sourceColumnOffsets.length;
    const output = { r1: copyTo.r1, c1: copyTo.c1, r2: copyTo.r1 + result.matched.length, c2: copyTo.c1 + width - 1 };
    const firstOutputRow = preserveDestinationHeaders ? output.r1 + 1 : output.r1;
    if (
      !inBounds(output) ||
      rangesOverlap(output, list) ||
      rangesOverlap(output, criteria) ||
      (sheet.merges && (rangeHasMerge(sheet.merges, list) || rangeHasMerge(sheet.merges, output)))
    ) {
      return { ok: false, error: { kind: "copy-destination" } };
    }
    for (let row = firstOutputRow; row <= output.r2; row += 1) {
      for (let col = output.c1; col <= output.c2; col += 1) {
        if (!read.isEditable(row, col) || read.spilled(row, col)) return { ok: false, error: { kind: "copy-protected" } };
      }
    }
    // The header (for a blank anchor) and the matched rows in one sheet
    // transaction; full cell objects keep formatting, formulas and links,
    // and formulas rebase as Sort and Remove Duplicates rebase them.
    const sourceRows = result.matched.map((rowIndex) => list.r1 + 1 + rowIndex);
    if (!preserveDestinationHeaders) sourceRows.unshift(list.r1);
    if (preserveDestinationHeaders && sourceRows.length === 0) return { ok: true, sheet, ...outcome };
    const nextCells: CellMap = { ...sheet.cells };
    for (let row = firstOutputRow; row <= output.r2; row += 1) {
      for (let col = output.c1; col <= output.c2; col += 1) delete nextCells[`${row}:${col}`];
    }
    sourceRows.forEach((sourceRow, rowOffset) => {
      const targetRow = firstOutputRow + rowOffset;
      sourceColumnOffsets.forEach((sourceOffset, targetOffset) => {
        const sourceCol = list.c1 + sourceOffset;
        const targetCol = output.c1 + targetOffset;
        const cell = sheet.cells[`${sourceRow}:${sourceCol}`];
        if (cell !== undefined) {
          nextCells[`${targetRow}:${targetCol}`] = relocateSortedCell(cell, targetRow - sourceRow, targetCol - sourceCol);
        }
      });
    });
    return { ok: true, sheet: { ...sheet, cells: nextCells }, ...outcome };
  }
  if (read.sheetProtected) return { ok: false, error: { kind: "protected" } };
  // Excel's in-place Advanced Filter replaces a worksheet AutoFilter that
  // filters nothing (its arrows go; a sheet has one filter). The caller has
  // already refused an AutoFilter that filters. Named Sheet Views filter
  // through that AutoFilter, so it stays and the filter is refused.
  const autoFilterId = sheet.autoFilterId?.toUpperCase();
  if (
    autoFilterId &&
    (sheet.namedSheetViews ?? []).some((view) => view.filterId?.toUpperCase() === autoFilterId)
  ) {
    return { ok: false, error: { kind: "named-sheet-views" } };
  }
  // One transaction replaces only Advanced Filter's prior row contribution.
  // Manual hides, outline collapse and table filters stay hidden on their own.
  const matched = new Set(result.matched);
  const hiddenRows = rows.flatMap((_row, index) => (matched.has(index) ? [] : [list.r1 + 1 + index]));
  const filtered: S = {
    ...sheet,
    advancedFilterState: { ref: rangeA1(list), hiddenRows, criteria: rangeA1(criteria) },
  };
  delete filtered.autoFilterRef;
  delete filtered.autoFilterId;
  delete filtered.namedSheetViewsOwnsAutoFilter;
  return { ok: true, sheet: filtered, ...outcome };
}
