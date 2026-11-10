/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Sort a spreadsheet's used rows by a key column.
 *
 * Whole rows are permuted, so values keep their row associations. Formula cells
 * relocate with their row: relative references shift by the source-to-target
 * displacement, absolute references remain fixed, and stale display caches are
 * dropped so recalculation cannot reuse an old result. Values compare with the
 * shared sort-filter comparator: numbers before text, blanks always last.
 *
 * Merged-cell spans travel with the rows they cover: each span's anchor row is
 * remapped through the SAME row permutation, and any span whose covered rows no
 * longer stay contiguous (the merge would be torn apart) is dropped — so the
 * returned sheet's `merges` always stay consistent with its sorted `cells`.
 */

import {
  remapMergesByRowPermutation,
  type MergeSpan,
} from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { shiftFormula } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import type { DrawingObjectLike } from "./object-buttons";
import { chartsAfterRowSort, objectsAfterRowSort, type SortObjectsGrid } from "./sort-objects";
import {
  compareCells,
  normalizeFillColor,
  sortKeyOf as sortKey,
} from "./sort-filter";

export type SortDirection = "asc" | "desc";

export interface SortableCell {
  value?: string;
  formula?: string;
  /** Optional computed result cache; removed when a formula changes position. */
  display?: string;
  chartValueCache?: unknown;
  /**
   * Optional cell style. Only `style.backgroundColor` is read, and ONLY by the
   * "sort by color" path ({@link sortByColumnColor}); the value-based
   * {@link sortByColumn} ignores it entirely, so its behaviour is unchanged.
   */
  style?: { backgroundColor?: string; [key: string]: unknown };
}

export interface SortableSheet extends SortObjectsGrid {
  cells: Record<string, SortableCell>;
  /** Optional merged-cell spans, remapped (or dropped) through the sort. */
  merges?: MergeSpan[];
  /** Drawings: one inside a single sorted row travels with it (Excel). */
  objects?: DrawingObjectLike[];
  /** The sheet's own charts, by their cell anchors (lib/spreadsheet/chart-anchor). */
  charts?: Array<{ id: string; drawingAnchor?: unknown }>;
}

export type SortKeyResolver = (
  cell: SortableCell | undefined,
  row: number,
  col: number,
) => string;

export interface SortResult<T> {
  sheet: T;
  /** Reserved for structural refusals; formula-bearing rows are supported. */
  skipped: boolean;
}

/**
 * Relocate a cell through a sort permutation using Excel/WPS fill semantics.
 *
 * Plain cells travel by identity. A moved formula cell is cloned, its relative
 * references are shifted by the row/column displacement, its `value` and
 * `formula` stay aligned, and its stale display cache is removed.
 */
export function relocateSortedCell<T extends SortableCell>(
  cell: T,
  rowDelta: number,
  colDelta: number,
): T {
  const formula = cell.formula;
  if (
    !formula ||
    !formula.startsWith("=") ||
    (rowDelta === 0 && colDelta === 0)
  ) {
    return cell;
  }
  const shifted = shiftFormula(formula, rowDelta, colDelta);
  const relocated: SortableCell = {
    ...cell,
    formula: shifted,
    value: shifted,
  };
  delete relocated.display;
  delete relocated.chartValueCache;
  return relocated as T;
}

/**
 * Heuristic header detection (Excel-like): treat row 0 as a header — and keep
 * it pinned — when the key column's first cell is text and at least one row
 * below it in that column is numeric.
 */
function hasHeaderRow(
  rows: Array<Array<SortableCell | undefined>>,
  keyCol: number,
  keyResolver?: SortKeyResolver,
): boolean {
  if (rows.length < 2) return false;
  const resolvedKey = (row: number): SortableCell | undefined =>
    keyResolver
      ? { value: keyResolver(rows[row][keyCol], row, keyCol) }
      : rows[row][keyCol];
  if (sortKey(resolvedKey(0)).kind !== "string") return false;
  for (let r = 1; r < rows.length; r += 1) {
    if (sortKey(resolvedKey(r)).kind === "number") return true;
  }
  return false;
}

/**
 * Scan the used region: returns the bounds and the row-major matrix, or `null`
 * when there is nothing meaningful to sort. Shared by the value sort and color
 * sort so both bound the grid identically.
 */
function scanUsedRegion(sheet: SortableSheet): {
  maxRow: number;
  maxCol: number;
  rows: Array<Array<SortableCell | undefined>>;
} | null {
  let maxRow = -1;
  let maxCol = -1;
  for (const key of Object.keys(sheet.cells)) {
    const [rawRow, rawCol] = key.split(":");
    const row = Number(rawRow);
    const col = Number(rawCol);
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  if (maxRow < 1) return null; // nothing meaningful to sort
  const rows: Array<Array<SortableCell | undefined>> = [];
  for (let r = 0; r <= maxRow; r += 1) {
    const row: Array<SortableCell | undefined> = [];
    for (let c = 0; c <= maxCol; c += 1) {
      row.push(sheet.cells[`${r}:${c}`]);
    }
    rows.push(row);
  }
  return { maxRow, maxCol, rows };
}

/**
 * Rebuild a sheet from a row PERMUTATION (`order[destRow] = srcRow`), carrying
 * every cell with its row and remapping merged spans through the SAME
 * permutation (extended with identity for rows below the sorted region). Shared
 * by the value sort and the color sort.
 */
function rebuildFromOrder<T extends SortableSheet>(
  sheet: T,
  rows: Array<Array<SortableCell | undefined>>,
  maxRow: number,
  maxCol: number,
  order: number[],
): SortResult<T> {
  const cells: Record<string, SortableCell> = {};
  order.forEach((srcRow, destRow) => {
    for (let c = 0; c <= maxCol; c += 1) {
      const cell = rows[srcRow][c];
      if (cell !== undefined) {
        cells[`${destRow}:${c}`] = relocateSortedCell(
          cell,
          destRow - srcRow,
          0,
        );
      }
    }
  });
  const destinations = new Map(order.map((srcRow, destRow) => [srcRow, destRow]));
  const objects = objectsAfterRowSort(sheet.objects, sheet, { r1: 0, r2: maxRow, c1: 0, c2: maxCol }, (row) => destinations.get(row));
  const sortedCharts = chartsAfterRowSort(sheet.charts, sheet, { r1: 0, r2: maxRow, c1: 0, c2: maxCol }, (row) => destinations.get(row));
  if (sortedCharts !== sheet.charts) sheet = { ...sheet, charts: sortedCharts };
  if (objects !== sheet.objects) sheet = { ...sheet, objects };
  if (sheet.merges && sheet.merges.length > 0) {
    let maxMergeRow = maxRow;
    for (const span of sheet.merges) {
      maxMergeRow = Math.max(maxMergeRow, span.row + span.rowSpan - 1);
    }
    const fullOrder = [...order];
    for (let r = maxRow + 1; r <= maxMergeRow; r += 1) fullOrder.push(r);
    const merges = remapMergesByRowPermutation(sheet.merges, fullOrder);
    return { sheet: { ...sheet, cells, merges }, skipped: false };
  }
  return { sheet: { ...sheet, cells }, skipped: false };
}

/** Sort the used rows by `keyCol`, relocating formulas with their source rows. */
export function sortByColumn<T extends SortableSheet>(
  sheet: T,
  keyCol: number,
  direction: SortDirection,
  keyResolver?: SortKeyResolver,
): SortResult<T> {
  const scan = scanUsedRegion(sheet);
  if (!scan) return { sheet, skipped: false };
  const { maxRow, maxCol, rows } = scan;

  // Keep a detected header row pinned at the top; sort only the data rows.
  const startRow = hasHeaderRow(rows, keyCol, keyResolver) ? 1 : 0;
  const dataIndices: number[] = [];
  for (let r = startRow; r <= maxRow; r += 1) dataIndices.push(r);
  dataIndices.sort((a, b) => {
    const aKey = keyResolver
      ? { value: keyResolver(rows[a][keyCol], a, keyCol) }
      : rows[a][keyCol];
    const bKey = keyResolver
      ? { value: keyResolver(rows[b][keyCol], b, keyCol) }
      : rows[b][keyCol];
    // The one spreadsheet comparator: Excel's type order, blanks last.
    return compareCells(aKey, bKey, direction, true);
  });
  const order = [
    ...Array.from({ length: startRow }, (_, r) => r),
    ...dataIndices,
  ];

  return rebuildFromOrder(sheet, rows, maxRow, maxCol, order);
}

/**
 * Sort the used rows of a sheet so the rows whose `keyCol` cell has the CHOSEN
 * fill color float to the top (Excel/WPS "Sort by Color"). The order is stable:
 * the chosen-color rows keep their relative order and so do all the rest, so a
 * uniform-color column is a no-op. A detected header row stays pinned; formula
 * cells relocate and rebase like the value sort. The chosen color is NORMALIZED
 * (see normalizeFillColor), so passing "#FFF" matches a cell filled "#ffffff";
 * pass the empty string to bring NO-FILL rows to the top. Reads only
 * `cell.style.backgroundColor` when determining row order.
 */
export function sortByColumnColor<T extends SortableSheet>(
  sheet: T,
  keyCol: number,
  chosenColor: string,
): SortResult<T> {
  const scan = scanUsedRegion(sheet);
  if (!scan) return { sheet, skipped: false };
  const { maxRow, maxCol, rows } = scan;

  const target = normalizeFillColor(chosenColor);
  const startRow = hasHeaderRow(rows, keyCol) ? 1 : 0;
  // Stable partition: chosen-color rows (rank 0) before the rest (rank 1), each
  // bucket keeping its original order. A plain index tie-break makes it explicit.
  const dataIndices: Array<{ row: number; rank: number }> = [];
  for (let r = startRow; r <= maxRow; r += 1) {
    const rank =
      normalizeFillColor(rows[r][keyCol]?.style?.backgroundColor) === target
        ? 0
        : 1;
    dataIndices.push({ row: r, rank });
  }
  dataIndices.sort((a, b) =>
    a.rank !== b.rank ? a.rank - b.rank : a.row - b.row,
  );
  const order = [
    ...Array.from({ length: startRow }, (_, r) => r),
    ...dataIndices.map((entry) => entry.row),
  ];

  return rebuildFromOrder(sheet, rows, maxRow, maxCol, order);
}
