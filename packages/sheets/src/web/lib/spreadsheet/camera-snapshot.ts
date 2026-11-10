/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Camera tool (Excel/WPS "Camera" / "Picture of range"): the pure data side of a
 * floating object that shows a LIVE picture of a source cell range — a read-only
 * mini-grid of the range's current display values that updates as those cells
 * change. This module turns a range + a display-value getter into a bounded 2D
 * matrix the renderer paints; it is dependency-free and DOM-free so the bounding
 * + extraction is exhaustively unit-testable. The host wires it over the existing
 * formula evaluator's getDisplay(row, col).
 */

export interface CameraRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface CameraSnapshot {
  /** Row-major display strings for the (bounded) range. */
  rows: string[][];
  /** True when the source range exceeded the cell caps and was clipped. */
  truncated: boolean;
  /** The full source dimensions, before clipping. */
  sourceRowCount: number;
  sourceColCount: number;
}

/** Default render caps — a camera is a glanceable thumbnail, not a second grid. */
export const CAMERA_MAX_ROWS = 50;
export const CAMERA_MAX_COLS = 20;

/** Normalize a range so r1≤r2, c1≤c2, and all bounds are non-negative integers. */
export function normalizeCameraRange(range: CameraRange): CameraRange {
  const r1 = Math.max(0, Math.trunc(Math.min(range.r1, range.r2)));
  const r2 = Math.max(0, Math.trunc(Math.max(range.r1, range.r2)));
  const c1 = Math.max(0, Math.trunc(Math.min(range.c1, range.c2)));
  const c2 = Math.max(0, Math.trunc(Math.max(range.c1, range.c2)));
  return { r1, c1, r2, c2 };
}

/**
 * Extract a bounded display-value matrix for `range`, reading each cell through
 * `getDisplay`. Rows/cols beyond {@link CAMERA_MAX_ROWS}/{@link CAMERA_MAX_COLS}
 * (overridable) are clipped and flagged so the camera stays a thumbnail and a
 * pathological range can't paint a million cells. A null/undefined getDisplay
 * result renders as an empty string.
 */
export function rangeToDisplayMatrix(
  range: CameraRange,
  getDisplay: (row: number, col: number) => string | null | undefined,
  opts: { maxRows?: number; maxCols?: number } = {},
): CameraSnapshot {
  const norm = normalizeCameraRange(range);
  const maxRows = Math.max(1, Math.trunc(opts.maxRows ?? CAMERA_MAX_ROWS));
  const maxCols = Math.max(1, Math.trunc(opts.maxCols ?? CAMERA_MAX_COLS));
  const sourceRowCount = norm.r2 - norm.r1 + 1;
  const sourceColCount = norm.c2 - norm.c1 + 1;
  const drawRows = Math.min(sourceRowCount, maxRows);
  const drawCols = Math.min(sourceColCount, maxCols);
  const rows: string[][] = [];
  for (let r = 0; r < drawRows; r += 1) {
    const row: string[] = [];
    for (let c = 0; c < drawCols; c += 1) {
      const value = getDisplay(norm.r1 + r, norm.c1 + c);
      row.push(value == null ? "" : String(value));
    }
    rows.push(row);
  }
  return {
    rows,
    truncated: drawRows < sourceRowCount || drawCols < sourceColCount,
    sourceRowCount,
    sourceColCount,
  };
}
