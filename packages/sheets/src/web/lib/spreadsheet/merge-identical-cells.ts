/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Merge identical consecutive cells (WPS Smart Toolbox "合并相同单元格" parity).
 *
 * Given a selection range and a column to operate on, scans vertically and
 * groups consecutive cells that share exactly the same (non-empty) text value
 * into "runs". Each run of 2+ equal cells is collapsed into a single merged
 * cell using the existing MergeSpan mechanism (the anchor keeps the value;
 * covered cells are cleared). Single-cell runs and blank cells are left alone.
 *
 * The function is PURE: it returns new cells + merges arrays without mutating
 * the input. Callers are expected to pipe the result through the same
 * `addMerge` path used by the toolbar's Merge Cells button.
 *
 * Only vertical merging (within a column, across rows) is supported — matching
 * the WPS / Excel "Merge Same Cells" convention which always merges within a
 * column. Horizontal runs are not merged.
 */

import { addMerge, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";

export interface CellLike {
  value?: string;
  formula?: string;
  [key: string]: unknown;
}

export interface MergeIdenticalResult {
  /** Updated cell map (covered cells cleared, anchor cells unchanged). */
  cells: Record<string, CellLike>;
  /** Updated merges array (new spans appended via addMerge). */
  merges: MergeSpan[];
  /** Number of distinct merge spans created. */
  spansCreated: number;
}

/**
 * Describe the runs of identical consecutive values in a column over a row
 * range. Returns an array of `{ row, count }` descriptors: `row` is the
 * first row of the run and `count` is the run length. Only runs of 2+ equal
 * non-empty values are returned; singletons and blank cells are omitted.
 *
 * @param cells  - flat "row:col" cell map
 * @param col    - 0-based column index to scan
 * @param r1     - first row of the scan range (inclusive)
 * @param r2     - last row of the scan range (inclusive)
 */
export function findEqualRuns(
  cells: Record<string, CellLike>,
  col: number,
  r1: number,
  r2: number,
): Array<{ row: number; count: number }> {
  const runs: Array<{ row: number; count: number }> = [];
  if (r1 > r2 || col < 0) return runs;

  let runStart = r1;
  let runValue: string | null = null;
  let runLen = 0;

  const commitRun = () => {
    if (runValue !== null && runLen >= 2) {
      runs.push({ row: runStart, count: runLen });
    }
  };

  for (let r = r1; r <= r2; r += 1) {
    const cell = cells[`${r}:${col}`];
    // Only non-formula, non-empty values are eligible for merging (formulas
    // must not be silently cleared, and blank cells don't carry a value to
    // display in the merged anchor).
    const hasFormula =
      typeof cell?.formula === "string" && cell.formula.startsWith("=");
    const val =
      !hasFormula && cell?.value !== undefined && cell.value !== ""
        ? cell.value
        : null;

    if (val !== null && val === runValue) {
      runLen += 1;
    } else {
      // Close previous run.
      commitRun();
      // Start new run (or blank/formula pause).
      runStart = r;
      runValue = val;
      runLen = val !== null ? 1 : 0;
    }
  }
  // Close the final run.
  commitRun();

  return runs;
}

/**
 * Apply identical-cell merging over every column in the selection range.
 * For each column, vertical runs of 2+ consecutive equal (non-empty, non-
 * formula) cells are collapsed into merged spans. Covered cells are cleared.
 *
 * The result is suitable for passing directly to `replaceActiveSheet`.
 *
 * @param cells    - current flat "row:col" cell map
 * @param merges   - current merge spans (will be extended)
 * @param r1 / r2 - row bounds of the selection (inclusive, any order OK)
 * @param c1 / c2 - column bounds of the selection (inclusive, any order OK)
 */
export function mergeIdenticalCells(
  cells: Record<string, CellLike>,
  merges: readonly MergeSpan[],
  r1: number,
  r2: number,
  c1: number,
  c2: number,
): MergeIdenticalResult {
  const rowMin = Math.min(r1, r2);
  const rowMax = Math.max(r1, r2);
  const colMin = Math.min(c1, c2);
  const colMax = Math.max(c1, c2);

  let currentMerges: MergeSpan[] = [...merges];
  const nextCells: Record<string, CellLike> = { ...cells };
  let spansCreated = 0;

  for (let col = colMin; col <= colMax; col += 1) {
    const runs = findEqualRuns(nextCells, col, rowMin, rowMax);
    for (const run of runs) {
      const span: MergeSpan = {
        row: run.row,
        col,
        rowSpan: run.count,
        colSpan: 1,
      };
      // Use addMerge which handles overlap removal (same semantics as the
      // existing toolbar Merge Cells command).
      currentMerges = addMerge(currentMerges, span);
      // Clear the covered (non-anchor) cells so the model stays consistent:
      // covered cells must hold no value (formula engine reads them as blank).
      for (let r = run.row + 1; r < run.row + run.count; r += 1) {
        delete nextCells[`${r}:${col}`];
      }
      spansCreated += 1;
    }
  }

  return { cells: nextCells, merges: currentMerges, spansCreated };
}
