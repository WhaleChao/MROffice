/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helper for "Locate / Fill Blanks" (WPS Smart Toolbox 定位空值 / 填充空白单元格,
 * Excel "Go To Special > Blanks" + fill-down).
 *
 * Two operations over a rectangular range:
 *   - LOCATE  → the coordinates of every blank cell in the range (no writes), so
 *     the caller can select them for the user to act on next.
 *   - FILL    → the cell writes that fill every blank with either a CONSTANT value
 *     or the value of the nearest non-blank cell ABOVE it in the same column
 *     (fill-down, the common 填充空白=上方值 case).
 *
 * A cell is BLANK when it is absent OR its raw `value` is empty / whitespace-only —
 * the SAME emptiness test the grid's `writeCellValueInto` uses, so "fill the blanks"
 * targets exactly the cells that render as empty and that a fresh write would treat
 * as empty. (A formula cell that evaluates to "" is NOT blank: it has a non-empty
 * raw `value` starting with "=", so it is left alone — matching Excel, whose Go To
 * Special > Blanks excludes formula cells.)
 *
 * TOTAL and side-effect-free: an empty range, a single cell, an all-blank range, an
 * all-filled range (→ no writes), and an empty constant all return cleanly. The
 * caller applies the writes through the normal grid edit path, which is what gives
 * undo / autosave / sheet-protection for free.
 */

/** A cell read by the blank scan: only the raw `value` matters. */
export interface FillBlanksCell {
  value?: string;
}

/** The sheet cells the scan reads: a `row:col` -> cell map (Excel-style sparse). */
export type FillBlanksCells = Record<string, FillBlanksCell | undefined>;

/** The inclusive rectangle the scan operates on. */
export interface FillBlanksRange {
  r1: number;
  r2: number;
  c1: number;
  c2: number;
}

/** A blank cell coordinate the LOCATE pass found. */
export interface BlankCellCoord {
  row: number;
  col: number;
}

/** A single cell write the FILL pass produces. */
export interface FillBlanksWrite {
  row: number;
  col: number;
  value: string;
}

/** How FILL chooses each blank's new value. */
export type FillBlanksMode =
  | { kind: "constant"; value: string }
  // Fill-down: each blank takes the value of the nearest non-blank cell ABOVE it
  // (within the range, same column). Consecutive blanks cascade — they all inherit
  // the last non-blank value seen above them, exactly like Excel's fill-down. A
  // blank with no non-blank cell above it inside the range stays empty (no write).
  | { kind: "above" };

/** True when a cell is empty (absent or an empty `value`). A cell holding
 * spaces is text, not a blank: Fill Blanks and Go To Special > Blanks share
 * this definition, and neither overwrites it. */
export function isBlankCell(cell: FillBlanksCell | undefined): boolean {
  return (cell?.value ?? "") === "";
}

/** Normalize a range so r1<=r2 and c1<=c2 regardless of how it was passed. */
function normalizeRange(range: FillBlanksRange): FillBlanksRange {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}

/**
 * Return the coordinates of every BLANK cell inside `range`, in row-major order
 * (top-to-bottom, then left-to-right within a row). Reads `cells` only; never
 * mutates. An all-filled range yields `[]`.
 */
export function locateBlankCells(
  cells: FillBlanksCells,
  range: FillBlanksRange,
): BlankCellCoord[] {
  const { r1, r2, c1, c2 } = normalizeRange(range);
  const out: BlankCellCoord[] = [];
  for (let row = r1; row <= r2; row += 1) {
    for (let col = c1; col <= c2; col += 1) {
      if (isBlankCell(cells[`${row}:${col}`])) out.push({ row, col });
    }
  }
  return out;
}

/**
 * Compute the cell writes that fill the blanks in `range`.
 *
 * - `{ kind: "constant", value }` — every blank in the range is written with the
 *   given constant (even an empty constant is honoured: it produces writes that
 *   clear-to-empty, which the caller's write path treats as a no-op clear; callers
 *   that want to skip an empty constant should guard before calling).
 * - `{ kind: "above" }` — fill-down: each blank takes the nearest non-blank value
 *   ABOVE it in the same column. The running value is seeded from the cell DIRECTLY
 *   above the range's top row (`r1 - 1`) when that cell is non-blank, so a fill that
 *   starts mid-column still inherits the column's value above the selection (Excel
 *   behaviour). A blank with no non-blank value available above it (top of the
 *   column / range with nothing above) produces NO write and stays empty.
 *
 * Only BLANK cells ever produce a write — non-blank cells are never touched, so the
 * write set is minimal and a range with no blanks returns `[]`. Writes are emitted
 * in column-major-by-row order; their relative order does not matter to the caller.
 */
export function planFillBlanks(
  cells: FillBlanksCells,
  range: FillBlanksRange,
  mode: FillBlanksMode,
): FillBlanksWrite[] {
  const { r1, r2, c1, c2 } = normalizeRange(range);
  const writes: FillBlanksWrite[] = [];

  if (mode.kind === "constant") {
    for (let row = r1; row <= r2; row += 1) {
      for (let col = c1; col <= c2; col += 1) {
        if (isBlankCell(cells[`${row}:${col}`])) {
          writes.push({ row, col, value: mode.value });
        }
      }
    }
    return writes;
  }

  // Fill-down: walk each column top-to-bottom, carrying the last non-blank value.
  for (let col = c1; col <= c2; col += 1) {
    // Seed from the cell directly above the range's first row, if it is non-blank,
    // so a selection that begins below filled data still inherits that value.
    const aboveCell = r1 > 0 ? cells[`${r1 - 1}:${col}`] : undefined;
    let carried: string | null =
      r1 > 0 && !isBlankCell(aboveCell) ? (aboveCell?.value ?? "") : null;
    for (let row = r1; row <= r2; row += 1) {
      const cell = cells[`${row}:${col}`];
      if (isBlankCell(cell)) {
        // Only fill when we actually have a value above; otherwise leave it empty.
        if (carried !== null) writes.push({ row, col, value: carried });
      } else {
        carried = cell?.value ?? "";
      }
    }
  }
  return writes;
}
