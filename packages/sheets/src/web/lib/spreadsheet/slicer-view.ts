/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { type DataTable, type DataTableColumn, type DataTableColumnFilter } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { distinctColumnValues, filterToSlicerSelection, slicerSelectionState, type Slicer } from "./slicer";

/**
 * What a slicer SHOWS — its caption, its items and which are lit — resolved
 * from the workbook alone, so the grid's panel and the printout (the PDF
 * view's slicer pictures) read one answer. Excel prints a slicer as it shows
 * (Format Slicer > Properties > Print object, on by default).
 *
 * `unbound` says the slicer's Table column or PivotTable field is gone: it
 * shows no items. Why a bound slicer cannot be USED (a native filter, a
 * read-only PivotTable) is the grid's to say; it changes nothing shown.
 */
export interface SpreadsheetSlicerView {
  readonly title: string;
  readonly values: readonly string[];
  readonly selected: ReadonlySet<string>;
  readonly filtered: boolean;
  readonly unbound?: "table" | "pivot";
}

export interface SpreadsheetTableSlicerBinding {
  readonly table: DataTable;
  readonly column: DataTableColumn;
  readonly worksheetColumn: number;
  readonly rowStart: number;
  readonly rowEnd: number;
}

/** A Table slicer's column in its Table (the grid's resolveTableSlicerBinding). */
export function tableSlicerBinding(
  slicer: Slicer,
  dataTables: readonly DataTable[],
): SpreadsheetTableSlicerBinding | null {
  const source = slicer.source;
  if (source?.kind !== "table") return null;
  const table = dataTables.find((candidate) => candidate.id === source.tableId);
  if (!table?.columns) return null;
  const offset = table.columns.findIndex((column) => column.id === source.tableColumnId);
  if (offset < 0) return null;
  const worksheetColumn = table.range.c1 + offset;
  if (worksheetColumn !== slicer.column) return null;
  return {
    table,
    column: table.columns[offset],
    worksheetColumn,
    rowStart: table.range.r1 + (table.headerRow ? 1 : 0),
    rowEnd: table.range.r2,
  };
}

/**
 * A column's distinct shown values between two rows, the slicer's items
 * (the grid's slicerColumnValues): `shown(row, col)` is the cell's text as
 * the grid draws it.
 */
export function slicerColumnValues(
  shown: (row: number, col: number) => string,
  rowCount: number,
  col: number,
  rowStart = 1,
  rowEnd = rowCount - 1,
): string[] {
  const first = Math.max(0, Math.min(rowCount, rowStart));
  const last = Math.max(first - 1, Math.min(rowCount - 1, rowEnd));
  if (last < first) return [];
  const resolved: Record<string, { value?: string } | undefined> = {};
  for (let r = first; r <= last; r += 1) resolved[`${r - first + 1}:${col}`] = { value: shown(r, col) };
  return distinctColumnValues(resolved, col, last - first + 2);
}

export interface SpreadsheetSlicerViewContext {
  readonly dataTables: readonly DataTable[];
  /** The sheet's worksheet AutoFilter, by column. */
  readonly filters: Readonly<Record<string, readonly string[] | undefined>> | undefined;
  readonly rowCount: number;
  /** A cell's text as the grid draws it. */
  readonly shown: (row: number, col: number) => string;
  /** A PivotTable slicer's field: its items and its filter (pivotSlicerBinding). */
  readonly pivot: (slicer: Slicer) => { readonly values: readonly string[]; readonly filterValues?: readonly string[] | null } | null;
  /** A column's letters, the caption a slicer with nothing better shows. */
  readonly columnLabel: (col: number) => string;
}

export function spreadsheetSlicerView(slicer: Slicer, context: SpreadsheetSlicerViewContext): SpreadsheetSlicerView {
  if (slicer.source?.kind === "table") {
    const binding = tableSlicerBinding(slicer, context.dataTables);
    if (!binding) {
      return { title: slicer.title?.trim() || context.columnLabel(slicer.column), values: [], selected: new Set(), filtered: false, unbound: "table" };
    }
    const filter: DataTableColumnFilter | undefined = binding.table.columnFilters?.[String(binding.column.id)];
    const values = slicerColumnValues(context.shown, context.rowCount, binding.worksheetColumn, binding.rowStart, binding.rowEnd);
    return {
      title: slicer.title?.trim() || binding.column.name || context.columnLabel(binding.worksheetColumn),
      values,
      selected: new Set(filterToSlicerSelection(filter?.values, values)),
      filtered: slicerSelectionState(filter?.values, values).filtered,
    };
  }
  if (slicer.source?.kind === "pivot") {
    const binding = context.pivot(slicer);
    if (!binding) {
      return { title: slicer.title?.trim() || slicer.source.fieldId, values: [], selected: new Set(), filtered: false, unbound: "pivot" };
    }
    return {
      title: slicer.title?.trim() || slicer.source.fieldId,
      values: binding.values,
      selected: new Set(filterToSlicerSelection(binding.filterValues ?? undefined, binding.values)),
      filtered: slicerSelectionState(binding.filterValues ?? undefined, binding.values).filtered,
    };
  }
  const headerText = context.shown(0, slicer.column).trim();
  const values = slicerColumnValues(context.shown, context.rowCount, slicer.column);
  const filter = context.filters?.[String(slicer.column)];
  return {
    title: slicer.title?.trim() || headerText || context.columnLabel(slicer.column),
    values,
    selected: new Set(filterToSlicerSelection(filter ? [...filter] : undefined, values)),
    filtered: slicerSelectionState(filter ? [...filter] : undefined, values).filtered,
  };
}
