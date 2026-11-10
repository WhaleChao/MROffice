/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// WAI-ARIA grid coordinates for the spreadsheet. ----------------------------
// Pure helpers so the (skeptic-probed) off-by-one against the header band is
// testable in isolation, independent of the React render or the virtualization
// window.
//
// The grid is a `role="grid"` with TWO header bands:
//   - a column-letter header row (A, B, C…) across the top, and
//   - a row-number header column (1, 2, 3…) down the left.
//
// Index convention (1-based, per WAI-ARIA, "Numbering starts with 1, not 0"):
//   aria-rowcount / aria-colcount report the FULL logical sizes INCLUDING the
//   header band — i.e. data rows + the one column-letter header row, and data
//   columns + the one row-number header column. A virtualized grid renders only
//   a window of rows, so the COUNT must come from the sheet's logical size, not
//   the rendered DOM (a screen reader must hear "row 3 of 100001", never
//   "row 3 of 12").
//
//   A data cell's aria-rowindex / aria-colindex are aligned to the user-visible
//   A1 reference so a screen reader announces the SAME coordinate the user sees:
//   the cell at spreadsheet B3 reports aria-rowindex=3 (row 3) and aria-colindex=2
//   (column B is the 2nd column, A=1). The logical (0-based) grid index is
//   therefore offset by +1. Because the data indices are A1-aligned, the
//   column-letter header row carries no competing aria-rowindex (aria-rowindex is
//   optional); it stays discoverable via role="row" + role="columnheader" cells.

/** aria-rowcount for the grid: data rows + the column-letter header row. */
export function gridAriaRowCount(rowCount: number): number {
  return Math.max(0, Math.floor(rowCount)) + 1;
}

/** aria-colcount for the grid: data columns + the row-number header column. */
export function gridAriaColCount(columnCount: number): number {
  return Math.max(0, Math.floor(columnCount)) + 1;
}

/**
 * aria-rowindex for a data cell / data row, given its 0-based LOGICAL row.
 * A1-aligned: logical row 0 (spreadsheet "row 1") → 1, logical row 2 ("row 3")
 * → 3. The LOGICAL row (not the rendered/virtualized DOM position) is the input,
 * so a cell scrolled deep into a virtualized sheet still reports its true number.
 */
export function cellAriaRowIndex(logicalRow: number): number {
  return Math.floor(logicalRow) + 1;
}

/**
 * aria-colindex for a data cell / column header, given its 0-based LOGICAL
 * column. A1-aligned: column A (logical 0) → 1, column B (logical 1) → 2.
 */
export function cellAriaColIndex(logicalCol: number): number {
  return Math.floor(logicalCol) + 1;
}

/**
 * The live-region announcement for the active cell on a selection change:
 * the A1 coordinate, plus the displayed value when the cell is non-empty
 * (e.g. "B3, 42"). An empty cell announces just its coordinate ("B3").
 */
export function announceCell(label: string, value: string): string {
  const trimmed = (value ?? "").trim();
  return trimmed ? `${label}, ${trimmed}` : label;
}
