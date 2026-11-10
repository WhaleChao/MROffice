/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Keyboard navigation for native presentation table cells (PowerPoint / WPS).
 *
 * The persisted table uses HTML-style merged rows, so keyboard movement must
 * address the logical occupancy grid instead of assuming `row[cellIndex]` maps
 * directly to a visual column. These helpers are pure and DOM-free; the editor
 * owns focus and row insertion.
 */

import {
  buildOccupancyGrid,
  logicalCellPositions,
  logicalPositionOfCell,
  type MergeableCell,
} from "@mrnote/core/web/lib/presentation-table-merge.ts";

export interface TableCellTarget {
  row: number;
  cellIndex: number;
}

export type TableTabTarget =
  | ({ kind: "cell" } & TableCellTarget)
  | { kind: "append-row"; row: number };

export type TableArrowKey = "ArrowLeft" | "ArrowRight" | "ArrowUp" | "ArrowDown";

function orderedCellTargets(rows: MergeableCell[][]): TableCellTarget[] {
  const positions = logicalCellPositions(rows);
  const targets: Array<TableCellTarget & { logicalRow: number; logicalCol: number }> =
    [];
  for (let row = 0; row < rows.length; row += 1) {
    for (let cellIndex = 0; cellIndex < rows[row].length; cellIndex += 1) {
      const position = positions[row]?.[cellIndex];
      if (!position) continue;
      targets.push({
        row,
        cellIndex,
        logicalRow: position.row,
        logicalCol: position.col,
      });
    }
  }
  targets.sort(
    (left, right) =>
      left.logicalRow - right.logicalRow ||
      left.logicalCol - right.logicalCol,
  );
  return targets.map(({ row, cellIndex }) => ({ row, cellIndex }));
}

/**
 * Resolve Tab / Shift+Tab in logical row-major order. Forward Tab from the last
 * native cell requests one appended row, matching PowerPoint/WPS table editing.
 * Backward Tab from the first cell returns `null` so normal browser focus can
 * leave the table.
 */
export function tableTabTarget(
  rows: MergeableCell[][],
  current: TableCellTarget,
  backward = false,
): TableTabTarget | null {
  const targets = orderedCellTargets(rows);
  const index = targets.findIndex(
    (target) =>
      target.row === current.row && target.cellIndex === current.cellIndex,
  );
  if (index < 0) return null;
  const next = index + (backward ? -1 : 1);
  if (next >= 0 && next < targets.length) {
    return { kind: "cell", ...targets[next] };
  }
  if (!backward && next >= targets.length) {
    return { kind: "append-row", row: rows.length };
  }
  return null;
}

/**
 * Move one logical grid position for a cell-selection shortcut. If the adjacent
 * coordinate is covered by a merged cell, focus returns to that cell's physical
 * anchor; ragged gaps are skipped until a real occupant or table edge is found.
 */
export function tableArrowTarget(
  rows: MergeableCell[][],
  current: TableCellTarget,
  key: TableArrowKey,
): TableCellTarget | null {
  const origin = logicalPositionOfCell(rows, current.row, current.cellIndex);
  const grid = buildOccupancyGrid(rows);
  if (!origin || !grid) return null;
  const rowStep = key === "ArrowUp" ? -1 : key === "ArrowDown" ? 1 : 0;
  const colStep = key === "ArrowLeft" ? -1 : key === "ArrowRight" ? 1 : 0;
  let row = origin.row + rowStep;
  let col = origin.col + colStep;
  while (
    row >= 0 &&
    row < grid.length &&
    col >= 0 &&
    col < (grid[row]?.length ?? 0)
  ) {
    const occupant = grid[row]?.[col];
    if (
      occupant &&
      (occupant.anchorRow !== current.row ||
        occupant.cellIndex !== current.cellIndex)
    ) {
      return {
        row: occupant.anchorRow,
        cellIndex: occupant.cellIndex,
      };
    }
    row += rowStep;
    col += colStep;
  }
  return null;
}
