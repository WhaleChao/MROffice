/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { surroundingDataRegion, type DataRegionRange } from "@mrnote/core/web/lib/spreadsheet/data-region.ts";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";

type AutoFilterCells = Readonly<
  Record<string, { value?: unknown; formula?: string }>
>;
type AutoFilterMerges = ReadonlyArray<{
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}>;

/** The parts of a worksheet the AutoFilter range is read from. */
export interface WorksheetAutoFilterSource {
  rowCount: number;
  columnCount: number;
  cells: AutoFilterCells;
  autoFilterRef?: string;
  filters?: Readonly<Record<string, unknown>>;
  colorFilters?: Readonly<Record<string, unknown>>;
  conditionFilters?: Readonly<Record<string, unknown>>;
}

/** Columns that carry an ordinary worksheet criterion (value, colour or condition). */
export function worksheetFilterCriterionColumns(
  source: WorksheetAutoFilterSource,
): number[] {
  const columns = new Set<number>();
  for (const map of [
    source.filters,
    source.colorFilters,
    source.conditionFilters,
  ]) {
    for (const key of Object.keys(map ?? {})) {
      const column = Number(key);
      if (
        Number.isInteger(column) &&
        String(column) === key &&
        column >= 0 &&
        column < source.columnCount
      ) {
        columns.add(column);
      }
    }
  }
  return [...columns].sort((left, right) => left - right);
}

/**
 * The worksheet AutoFilter in force, header row first, or null when the sheet
 * has none (Excel: "Filters are either on for an entire range, or off").
 *
 * A stored `autoFilterRef` is the range. A sheet saved before the editor
 * recorded the range can still carry criteria; those filter rows 2.. of the
 * whole sheet under a header on row 1 (filterVisibilityGroups). Its range is
 * that same row span, over the header row's current region widened to every
 * criterion column, so the arrows sit over the columns that filter and
 * recording this range later hides exactly the rows that are hidden now.
 */
export function worksheetAutoFilterRange(
  source: WorksheetAutoFilterSource,
): DataRegionRange | null {
  if (source.autoFilterRef) {
    const stored = parseA1Range(source.autoFilterRef.replace(/\$/g, ""));
    return stored &&
      stored.r2 < source.rowCount &&
      stored.c2 < source.columnCount
      ? stored
      : null;
  }
  const columns = worksheetFilterCriterionColumns(source);
  if (columns.length === 0 || source.rowCount < 1) return null;
  let c1 = columns[0];
  let c2 = columns[columns.length - 1];
  for (const column of columns) {
    const header = surroundingDataRegion(
      source.cells,
      0,
      column,
      source.rowCount,
      source.columnCount,
    );
    if (!header) continue;
    c1 = Math.min(c1, header.c1);
    c2 = Math.max(c2, header.c2);
  }
  return { r1: 0, c1, r2: source.rowCount - 1, c2 };
}

function occupied(cells: AutoFilterCells, row: number, col: number) {
  const cell = cells[`${row}:${col}`];
  return (
    !!cell &&
    (Boolean(cell.formula) ||
      (cell.value !== undefined && cell.value !== null && cell.value !== ""))
  );
}

/**
 * The range Excel's Filter command turns AutoFilter on for: the selection
 * when it is more than one cell, otherwise the current region around the
 * active cell (a blank cell next to data joins that data's region; a blank
 * cell with no data around it filters itself). A selection of one row, or of
 * whole columns, takes the rows of data below its first row, so selecting the
 * header row filters the list under it.
 */
export function autoFilterRangeForToggle(
  source: WorksheetAutoFilterSource,
  active: { row: number; col: number },
  selection: DataRegionRange,
  merges: AutoFilterMerges = [],
): DataRegionRange {
  const { cells, rowCount, columnCount } = source;
  const regionAround = (row: number, col: number) =>
    surroundingDataRegion(
      cells,
      row,
      col,
      rowCount,
      columnCount,
      undefined,
      merges,
    );
  const multiCell =
    selection.r1 !== selection.r2 || selection.c1 !== selection.c2;
  if (multiCell) {
    const wholeColumns = selection.r1 === 0 && selection.r2 >= rowCount - 1;
    if (!wholeColumns && selection.r1 !== selection.r2) return selection;
    let bottom = selection.r1;
    if (wholeColumns) {
      for (const key of Object.keys(cells)) {
        const separator = key.indexOf(":");
        const row = Number(key.slice(0, separator));
        const col = Number(key.slice(separator + 1));
        if (
          col >= selection.c1 &&
          col <= selection.c2 &&
          row > bottom &&
          row < rowCount &&
          occupied(cells, row, col)
        ) {
          bottom = row;
        }
      }
    } else {
      for (let col = selection.c1; col <= selection.c2; col += 1) {
        const region = occupied(cells, selection.r1, col)
          ? regionAround(selection.r1, col)
          : null;
        if (!region) continue;
        bottom = Math.max(bottom, region.r2);
        // One region answers for every header cell it already spans.
        col = Math.max(col, region.c2);
      }
    }
    return { ...selection, r2: bottom };
  }
  const own = regionAround(active.row, active.col);
  if (own) return own;
  for (let row = active.row - 1; row <= active.row + 1; row += 1) {
    for (let col = active.col - 1; col <= active.col + 1; col += 1) {
      if (row < 0 || col < 0 || row >= rowCount || col >= columnCount) continue;
      const region = occupied(cells, row, col) ? regionAround(row, col) : null;
      if (!region) continue;
      return {
        r1: Math.min(region.r1, active.row),
        c1: Math.min(region.c1, active.col),
        r2: Math.max(region.r2, active.row),
        c2: Math.max(region.c2, active.col),
      };
    }
  }
  return {
    r1: active.row,
    c1: active.col,
    r2: active.row,
    c2: active.col,
  };
}
