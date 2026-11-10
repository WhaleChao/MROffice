/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Filter and a column's AutoFilter menu (values, Text / Number
 * Filters, Clear Filter) as pure sheet transactions. Moved out of
 * SpreadsheetWindow so the ribbon and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps Sheet View sessions,
 * the table branch's notices and the menu state.
 *
 * A worksheet holds one AutoFilter: `autoFilterRef` (its range, header row
 * first) and per-column criteria keyed by the absolute column index:
 * `filters` (the values left checked, as displayed), `colorFilters` and
 * `conditionFilters` (one Text/Number Filter per column).
 */

import { surroundingDataRegion, type DataRegionRange } from "@mrnote/core/web/lib/spreadsheet/data-region.ts";
import {
  normalizeConditionFilter,
  type ConditionFilter,
} from "@mrnote/core/web/lib/spreadsheet/filter-conditions.ts";
import {
  autoFilterRangeForToggle,
  worksheetAutoFilterRange,
} from "./worksheet-autofilter";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** The parts of a worksheet the filter buttons read and write. */
export interface FilterButtonSheet {
  rowCount: number;
  columnCount: number;
  cells: Readonly<Record<string, { value?: string; formula?: string }>>;
  autoFilterRef?: string;
  autoFilterId?: string;
  namedSheetViewsOwnsAutoFilter?: true;
  namedSheetViews?: ReadonlyArray<{ filterId?: string }>;
  filters?: Record<string, string[]>;
  colorFilters?: Record<string, string[]>;
  conditionFilters?: Record<string, ConditionFilter>;
  /** An in-place Advanced Filter's result (its own hidden rows). */
  advancedFilterState?: unknown;
  /** Rows hidden by hand (0-based; lib/spreadsheet/manual-hidden). */
  manuallyHiddenRows?: readonly number[];
}













export function withAutoFilterRowsShown<S extends FilterButtonSheet>(
  sheet: S,
  range: DataRegionRange | null,
): S {
  if (!range) return sheet;
  const inData = (row: number) => row > range.r1 && row <= range.r2;
  // A collapsed outline group hides its rows by the same flag: the rows of it
  // the filter covers follow the filter (Excel 16 for Mac, xlsx-rounds
  // evidence/xprobe-filtermanual: rows 6:7 in a collapsed group, row 7 kept by
  // the filter shows; R256-X05 E11, 14 rows of a collapsed level-2 group the
  // filter kept stayed hidden). The group is expanded, and its rows the range
  // does not cover (its header among them) stay hidden, by hand.
  const groups = (sheet as { outlineGroups?: ReadonlyArray<{ axis?: string; start: number; end: number; collapsed?: boolean }> }).outlineGroups;
  const opened = Array.isArray(groups)
    ? groups.filter((g) => g.axis === "row" && g.collapsed === true && g.start <= range.r2 && g.end > range.r1)
    : [];
  const hidden = Array.isArray(sheet.manuallyHiddenRows) ? sheet.manuallyHiddenRows : [];
  const kept = hidden.filter((row) => !inData(row));
  if (opened.length === 0 && kept.length === hidden.length) return sheet;
  const next: S = { ...sheet };
  if (opened.length > 0) {
    const outside = new Set(kept);
    for (const g of opened) for (let row = g.start; row <= g.end; row += 1) if (!inData(row)) outside.add(row);
    kept.splice(0, kept.length, ...[...outside].sort((a, b) => a - b));
    (next as { outlineGroups?: unknown }).outlineGroups = groups!.map((g) => (opened.includes(g) ? { ...g, collapsed: false } : g));
  }
  if (kept.length > 0) next.manuallyHiddenRows = kept;
  else delete next.manuallyHiddenRows;
  return next;
}

/** The AutoFilter range a filter write leaves on the sheet. */
function filteredRange(sheet: FilterButtonSheet, autoFilter: DataRegionRange | null): DataRegionRange | null {
  return sheet.autoFilterRef ? sheetAutoFilterRange(sheet) : autoFilter;
}

/** A table's occupied range, as the AutoFilter toggle tests overlap. */
interface TableExtent {
  range: Rect;
  totalsRow?: boolean;
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

/** The sheet's AutoFilter range (stored, or inferred for a legacy sheet that
 *  holds criteria without one), or null when it has none. */
export function sheetAutoFilterRange(sheet: FilterButtonSheet): DataRegionRange | null {
  return worksheetAutoFilterRange({
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    cells: sheet.autoFilterRef ? {} : sheet.cells,
    autoFilterRef: sheet.autoFilterRef,
    filters: sheet.filters,
    colorFilters: sheet.colorFilters,
    conditionFilters: sheet.conditionFilters,
  });
}

/** Whether the sheet's AutoFilter holds a criterion (Excel's FilterMode). */
function filtering(sheet: FilterButtonSheet): boolean {
  return [sheet.filters, sheet.colorFilters, sheet.conditionFilters].some(
    (map) => map !== undefined && map !== null && Object.keys(map).length > 0,
  );
}

/**
 * Data > Filter outside a Table: turns the AutoFilter off (removing the arrows
 * and every criterion, so all rows show again) when the sheet has one, or on
 * over the selection / the current region around the active cell.
 * "namedViews" when Named Sheet Views hold criteria against it;
 * "overlapsTable" when the new range would overlap a Table.
 */
export function autoFilterToggledSheet<S extends FilterButtonSheet>(
  sheet: S,
  autoFilter: DataRegionRange | null,
  active: { row: number; col: number },
  selection: Rect,
  merges: ReadonlyArray<{ row: number; col: number; rowSpan: number; colSpan: number }>,
  tables: readonly TableExtent[],
): S | "namedViews" | "overlapsTable" {
  if (autoFilter) {
    const autoFilterId = sheet.autoFilterId?.toUpperCase();
    if (
      autoFilterId &&
      (sheet.namedSheetViews ?? []).some(
        (view) => view.filterId?.toUpperCase() === autoFilterId,
      )
    ) {
      // Named Sheet Views hold their criteria against this AutoFilter's
      // identity and range; removing it would orphan every view's filter.
      return "namedViews";
    }
    const next: S = { ...sheet };
    delete next.autoFilterRef;
    delete next.autoFilterId;
    delete next.namedSheetViewsOwnsAutoFilter;
    delete next.filters;
    delete next.colorFilters;
    delete next.conditionFilters;
    // Filter off shows every row the filter hid and every row hidden by hand
    
    return filtering(sheet) ? withAutoFilterRowsShown(next, autoFilter) : next;
  }
  const range = autoFilterRangeForToggle(sheet, active, selection, merges);
  if (
    tables.some(
      ({ range: table, totalsRow }) =>
        range.r1 <= (totalsRow ? table.r2 + 1 : table.r2) &&
        range.r2 >= table.r1 &&
        range.c1 <= table.c2 &&
        range.c2 >= table.c1,
    )
  ) {
    return "overlapsTable";
  }
  // Filter on shows the rows hidden by hand under its range as well (Excel 16
  
  return withAutoFilterRowsShown({ ...sheet, autoFilterRef: rangeA1(range) }, range);
}

/**
 * The range a column's filter acts on: the AutoFilter's when the sheet has
 * one (null when the column is outside it), otherwise the data region of the
 * column around the active cell or its first filled cell.
 */
export function ordinaryFilterRange(
  sheet: FilterButtonSheet,
  autoFilter: DataRegionRange | null,
  col: number,
  active: { row: number; col: number },
): DataRegionRange | null {
  if (sheet.autoFilterRef || autoFilter) {
    return autoFilter && col >= autoFilter.c1 && col <= autoFilter.c2
      ? autoFilter
      : null;
  }
  const regionAt = (row: number) =>
    surroundingDataRegion(sheet.cells, row, col, sheet.rowCount, sheet.columnCount);
  let region = active.col === col ? regionAt(active.row) : null;
  for (let row = 0; !region && row < sheet.rowCount; row += 1) {
    const cell = sheet.cells[`${row}:${col}`];
    if (cell && (cell.formula || (cell.value ?? "") !== "")) {
      region = regionAt(row);
      break;
    }
  }
  return region && region.r2 > region.r1 ? region : null;
}

/**
 * The first filter set on a sheet that has no AutoFilter range establishes
 * one (ordinaryFilterRange), as Excel's Filter button does, so the filter owns
 * its table and nothing else; an XLSX export then carries the right
 * <autoFilter ref>.
 */
export function establishedAutoFilterRef(
  sheet: FilterButtonSheet,
  autoFilter: DataRegionRange | null,
  col: number,
  active: { row: number; col: number },
): { autoFilterRef?: string } {
  if (sheet.autoFilterRef) return {};
  const range = ordinaryFilterRange(sheet, autoFilter, col, active);
  return range ? { autoFilterRef: rangeA1(range) } : {};
}

/**
 * A sheet saved before its AutoFilter range was recorded (criteria, no ref)
 * records that range with its first filter write of any kind. Clearing a
 * criterion then keeps the arrows, as Excel's Clear keeps AutoFilter on.
 */
export function recordedLegacyAutoFilter(
  sheet: FilterButtonSheet,
  autoFilter: DataRegionRange | null,
): { autoFilterRef?: string } {
  return !sheet.autoFilterRef && autoFilter ? { autoFilterRef: rangeA1(autoFilter) } : {};
}

/** A column filter write: the criteria maps (for a Sheet View) and the sheet. */
export interface ColumnFilterPlan<S> {
  filters: Record<string, string[]> | undefined;
  conditionFilters: Record<string, ConditionFilter> | undefined;
  sheet: S;
}

/**
 * A column's value filter: the values left checked, as displayed ("" is
 * (Blanks)); null clears it. A column carries one kind of criterion, so its
 * Text/Number Filter goes.
 */
export function columnValueFilterPlan<S extends FilterButtonSheet>(
  sheet: S,
  autoFilter: DataRegionRange | null,
  col: number,
  allowed: string[] | null,
  active: { row: number; col: number },
): ColumnFilterPlan<S> {
  const key = String(col);
  const nextFilters: Record<string, string[]> = { ...(sheet.filters || {}) };
  if (allowed === null) delete nextFilters[key];
  else nextFilters[key] = allowed;
  const nextConditions = { ...(sheet.conditionFilters || {}) };
  delete nextConditions[key];
  const filters = Object.keys(nextFilters).length > 0 ? nextFilters : undefined;
  const conditionFilters = Object.keys(nextConditions).length > 0 ? nextConditions : undefined;
  return {
    filters,
    conditionFilters,
    sheet: appliedFilterSheet(sheet, autoFilter, {
      ...sheet,
      ...(allowed !== null
        ? establishedAutoFilterRef(sheet, autoFilter, col, active)
        : recordedLegacyAutoFilter(sheet, autoFilter)),
      filters,
      conditionFilters,
    }),
  };
}




export function appliedFilterSheet<S extends FilterButtonSheet>(
  before: S,
  autoFilter: DataRegionRange | null,
  next: S,
): S {
  if (!filtering(before) && !filtering(next)) return next;
  return withAutoFilterRowsShown(next, filteredRange(next, autoFilter));
}

/**
 * A column's Text / Number Filter (replaces the value checklist for that
 * column; colour filters stay); null clears it. Null when the condition is
 * not one the filter can hold.
 */
export function columnConditionFilterPlan<S extends FilterButtonSheet>(
  sheet: S,
  autoFilter: DataRegionRange | null,
  col: number,
  cond: ConditionFilter | null,
  active: { row: number; col: number },
): ColumnFilterPlan<S> | null {
  const key = String(col);
  const nextCond: Record<string, ConditionFilter> = { ...(sheet.conditionFilters || {}) };
  if (cond === null) {
    delete nextCond[key];
  } else {
    const normalized = normalizeConditionFilter(cond);
    if (!normalized) return null;
    nextCond[key] = normalized;
  }
  const nextFilters = { ...(sheet.filters || {}) };
  delete nextFilters[key];
  const filters = Object.keys(nextFilters).length > 0 ? nextFilters : undefined;
  const conditionFilters = Object.keys(nextCond).length > 0 ? nextCond : undefined;
  return {
    filters,
    conditionFilters,
    sheet: appliedFilterSheet(sheet, autoFilter, {
      ...sheet,
      ...(cond !== null
        ? establishedAutoFilterRef(sheet, autoFilter, col, active)
        : recordedLegacyAutoFilter(sheet, autoFilter)),
      filters,
      conditionFilters,
    }),
  };
}

/**
 * Data > Clear: every criterion of the sheet's AutoFilter goes (values,
 * colours, conditions); the AutoFilter itself stays on. Null when no column
 * is filtered.
 */
export function clearedFiltersSheet<S extends FilterButtonSheet>(
  sheet: S,
  autoFilter: DataRegionRange | null,
): S | null {
  const hasValue = sheet.filters && Object.keys(sheet.filters).length > 0;
  const hasColor = sheet.colorFilters && Object.keys(sheet.colorFilters).length > 0;
  const hasCond = sheet.conditionFilters && Object.keys(sheet.conditionFilters).length > 0;
  // Data > Clear also ends an in-place Advanced Filter, as Excel's Clear
  // (Show All) does: the rows it hid show again; rows hidden by hand stay.
  // An AutoFilter's criteria cleared show every row under it, those hidden
  
  const hasAdvanced = sheet.advancedFilterState !== undefined;
  if (!hasValue && !hasColor && !hasCond && !hasAdvanced) return null;
  const next: S =
    hasValue || hasColor || hasCond
      ? {
          ...sheet,
          ...recordedLegacyAutoFilter(sheet, autoFilter),
          filters: undefined,
          colorFilters: undefined,
          conditionFilters: undefined,
        }
      : { ...sheet };
  if (hasAdvanced) delete next.advancedFilterState;
  return hasValue || hasColor || hasCond ? withAutoFilterRowsShown(next, filteredRange(next, autoFilter)) : next;
}
