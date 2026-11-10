/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native presentation table selection (PowerPoint / WPS "Select" group).
 *
 * The editor stores physical row-array coordinates for the focused textarea,
 * while users select logical table rows and columns. These helpers bridge the
 * two models and keep merged cells whole: a row/column selection expands across
 * the complete rowspan/colspan of the focused cell, and formatting includes any
 * cell whose logical rectangle intersects the selected rectangle.
 */

import {
  buildOccupancyGrid,
  logicalPositionOfCell,
  tableLogicalWidth,
  type MergeableCell,
} from "./presentation-table-merge";

export type TableSelectionScope = "cells" | "row" | "column" | "table";

export interface TableCellSelectionPoint {
  /** Source row-array index. */
  row: number;
  /** Source cell index inside the row array. */
  col: number;
}

export interface TableSelectionRectangle {
  row1: number;
  row2: number;
  col1: number;
  col2: number;
}

export type TableStructureAxis = "row" | "column";

export interface TableAxisSelectionRange {
  start: number;
  end: number;
  count: number;
}

function positiveSpan(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return 1;
  return Math.max(1, Math.floor(value));
}

function cellRectangle(
  rows: MergeableCell[][],
  point: TableCellSelectionPoint,
): TableSelectionRectangle | null {
  const cell = rows[point.row]?.[point.col];
  const position = logicalPositionOfCell(rows, point.row, point.col);
  if (!cell || !position) return null;
  return {
    row1: position.row,
    row2: position.row + positiveSpan(cell.rowspan) - 1,
    col1: position.col,
    col2: position.col + positiveSpan(cell.colspan) - 1,
  };
}

/**
 * Every cell of the table: what a table selected by its border stands for
 * when Home formats it (PowerPoint formats all of its cells), or `null` for
 * a table without cells.
 */
export function tableWholeSelection(rows: MergeableCell[][]): TableSelectionRectangle | null {
  const grid = buildOccupancyGrid(rows);
  const colCount = tableLogicalWidth(rows);
  if (!grid || grid.length === 0 || colCount === 0) return null;
  return { row1: 0, row2: grid.length - 1, col1: 0, col2: colCount - 1 };
}

/**
 * Resolve the logical rectangle addressed by a focused cell, optional
 * Shift-selection anchor, and WPS/PowerPoint selection command.
 */
export function tableSelectionRectangle(
  rows: MergeableCell[][],
  focus: TableCellSelectionPoint,
  anchor: TableCellSelectionPoint | null,
  scope: TableSelectionScope,
): TableSelectionRectangle | null {
  const grid = buildOccupancyGrid(rows);
  const focusRectangle = cellRectangle(rows, focus);
  if (!grid || !focusRectangle) return null;
  const rowCount = grid.length;
  const colCount = tableLogicalWidth(rows);
  if (rowCount === 0 || colCount === 0) return null;
  const anchorRectangle = anchor ? cellRectangle(rows, anchor) : null;

  if (scope === "table") {
    return { row1: 0, row2: rowCount - 1, col1: 0, col2: colCount - 1 };
  }
  if (scope === "row") {
    return {
      row1: anchorRectangle
        ? Math.min(focusRectangle.row1, anchorRectangle.row1)
        : focusRectangle.row1,
      row2: Math.min(
        rowCount - 1,
        anchorRectangle
          ? Math.max(focusRectangle.row2, anchorRectangle.row2)
          : focusRectangle.row2,
      ),
      col1: 0,
      col2: colCount - 1,
    };
  }
  if (scope === "column") {
    return {
      row1: 0,
      row2: rowCount - 1,
      col1: anchorRectangle
        ? Math.min(focusRectangle.col1, anchorRectangle.col1)
        : focusRectangle.col1,
      col2: Math.min(
        colCount - 1,
        anchorRectangle
          ? Math.max(focusRectangle.col2, anchorRectangle.col2)
          : focusRectangle.col2,
      ),
    };
  }

  if (!anchorRectangle) return focusRectangle;
  return {
    row1: Math.min(focusRectangle.row1, anchorRectangle.row1),
    row2: Math.max(focusRectangle.row2, anchorRectangle.row2),
    col1: Math.min(focusRectangle.col1, anchorRectangle.col1),
    col2: Math.max(focusRectangle.col2, anchorRectangle.col2),
  };
}

/**
 * Resolve the contiguous row/column span affected by an Office structural
 * command. A Shift-selected cell rectangle addresses both axes; an explicit
 * Select Row / Select Column scope addresses only its matching axis. The
 * opposite scope and whole-table scope fall back to the focused slot so
 * selecting a column can never accidentally delete every table row.
 */
export function tableAxisSelectionRange(
  selection: TableSelectionRectangle | null | undefined,
  scope: TableSelectionScope,
  axis: TableStructureAxis,
  fallbackIndex: number,
  axisCount: number,
): TableAxisSelectionRange | null {
  if (!Number.isInteger(axisCount) || axisCount <= 0) return null;
  const fallback = Math.min(
    axisCount - 1,
    Math.max(0, Math.floor(fallbackIndex)),
  );
  const matchingScope =
    scope === "cells" ||
    (axis === "row" && scope === "row") ||
    (axis === "column" && scope === "column");
  if (!selection || !matchingScope) {
    return { start: fallback, end: fallback, count: 1 };
  }
  const rawStart = axis === "row" ? selection.row1 : selection.col1;
  const rawEnd = axis === "row" ? selection.row2 : selection.col2;
  if (!Number.isFinite(rawStart) || !Number.isFinite(rawEnd)) {
    return { start: fallback, end: fallback, count: 1 };
  }
  const start = Math.min(
    axisCount - 1,
    Math.max(0, Math.floor(Math.min(rawStart, rawEnd))),
  );
  const end = Math.min(
    axisCount - 1,
    Math.max(start, Math.floor(Math.max(rawStart, rawEnd))),
  );
  return { start, end, count: end - start + 1 };
}

/** Whether a native cell (including its complete span) touches the selection. */
export function tableCellIntersectsSelection(
  cell: MergeableCell,
  position: { row: number; col: number } | null,
  selection: TableSelectionRectangle,
): boolean {
  if (!position) return false;
  const cellRow2 = position.row + positiveSpan(cell.rowspan) - 1;
  const cellCol2 = position.col + positiveSpan(cell.colspan) - 1;
  return (
    position.row <= selection.row2 &&
    cellRow2 >= selection.row1 &&
    position.col <= selection.col2 &&
    cellCol2 >= selection.col1
  );
}
