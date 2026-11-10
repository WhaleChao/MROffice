/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface DataRegionRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export const MAX_INFERRED_DATA_REGION_CELLS = 500_000;
/** The region a command the person asked for takes (a sort, a filter, a table):
 * past 50,000 cells, a list of 5,000 rows and 10 columns, this used to answer
 * nothing, and Sort and Filter said the active cell held no data. */
export const MAX_COMMAND_DATA_REGION_CELLS = 5_000_000;
/** What one selection hands an AI analysis. */
export const MAX_ANALYZED_REGION_CELLS = 50_000;

type RegionCells = Readonly<Record<string, { value?: unknown; formula?: string }>>;
type RegionMerges = ReadonlyArray<{ row: number; col: number; rowSpan: number; colSpan: number }>;

function regionCellFilled(cells: RegionCells, r: number, c: number): boolean {
  return regionCellHoldsData(cells[`${r}:${c}`]);
}

/** Whether a cell counts as data to a region (a formula, or a value): all a
 * region scan reads of it, so cells changed otherwise (a formula's new
 * result, a format) leave every region as it was. */
export function regionCellHoldsData(cell: { value?: unknown; formula?: string } | undefined): boolean {
  return (
    !!cell &&
    (Boolean(cell.formula) ||
      (cell.value !== undefined && cell.value !== null && cell.value !== ""))
  );
}

/** A cell holds data, or is covered by a merge whose top-left cell does. */
function regionCellOccupied(cells: RegionCells, merges: RegionMerges, r: number, c: number): boolean {
  if (regionCellFilled(cells, r, c)) return true;
  for (const merge of merges) {
    if (
      r >= merge.row &&
      r < merge.row + merge.rowSpan &&
      c >= merge.col &&
      c < merge.col + merge.colSpan
    ) {
      return regionCellFilled(cells, merge.row, merge.col);
    }
  }
  return false;
}

/** A bounded CurrentRegion scan, not the worksheet's whole used range.
 * Blank rows/columns bound the region; holes within its rectangle do not.
 * Each newly enclosed coordinate is visited once, including empty holes, so
 * disconnected values inside the growing rectangle can extend its boundary.
 * No array proportional to worksheet dimensions or distant stored cells.
 */
export function surroundingDataRegion(
  cells: Readonly<Record<string, { value?: unknown; formula?: string }>>,
  row: number,
  col: number,
  rowCount: number,
  columnCount: number,
  maxCells = MAX_COMMAND_DATA_REGION_CELLS,
  // Excel's CurrentRegion takes a merged cell whole: the blank cells a merge
  // covers belong to the data, so a region never stops halfway through one
  // (and a sort over it can still see, and refuse, a merge it would tear).
  merges: ReadonlyArray<{
    row: number;
    col: number;
    rowSpan: number;
    colSpan: number;
  }> = [],
): DataRegionRange | null {
  if (
    ![row, col, rowCount, columnCount, maxCells].every(Number.isSafeInteger) ||
    row < 0 ||
    col < 0 ||
    row >= rowCount ||
    col >= columnCount ||
    maxCells < 1
  )
    return null;
  const occupied = (r: number, c: number) => regionCellOccupied(cells, merges, r, c);
  if (!occupied(row, col)) return null;
  let range: DataRegionRange = { r1: row, c1: col, r2: row, c2: col };
  const queue: Array<readonly [number, number]> = [[row, col]];
  for (let cursor = 0; cursor < queue.length; cursor += 1) {
    const [r, c] = queue[cursor];
    for (
      let nr = Math.max(0, r - 1);
      nr <= Math.min(rowCount - 1, r + 1);
      nr += 1
    ) {
      for (
        let nc = Math.max(0, c - 1);
        nc <= Math.min(columnCount - 1, c + 1);
        nc += 1
      ) {
        if (
          (nr >= range.r1 &&
            nr <= range.r2 &&
            nc >= range.c1 &&
            nc <= range.c2) ||
          !occupied(nr, nc)
        )
          continue;
        const next = {
          r1: Math.min(range.r1, nr),
          c1: Math.min(range.c1, nc),
          r2: Math.max(range.r2, nr),
          c2: Math.max(range.c2, nc),
        };
        if ((next.r2 - next.r1 + 1) * (next.c2 - next.c1 + 1) > maxCells)
          return null;
        // Enqueue only the disjoint strips added to the previous rectangle.
        for (let rr = next.r1; rr < range.r1; rr += 1)
          for (let cc = next.c1; cc <= next.c2; cc += 1) queue.push([rr, cc]);
        for (let rr = range.r2 + 1; rr <= next.r2; rr += 1)
          for (let cc = next.c1; cc <= next.c2; cc += 1) queue.push([rr, cc]);
        for (let rr = range.r1; rr <= range.r2; rr += 1) {
          for (let cc = next.c1; cc < range.c1; cc += 1) queue.push([rr, cc]);
          for (let cc = range.c2 + 1; cc <= next.c2; cc += 1)
            queue.push([rr, cc]);
        }
        range = next;
      }
    }
  }
  return range;
}

/**
 * Excel's Sort Warning (and Remove Duplicates'): data in a column right beside
 * a multi-cell selection, within its rows, is data the user has not selected.
 * The region "Expand the selection" would take -- the selection joined with
 * the CurrentRegion of that data -- or null when nothing borders it, and the
 * selection is sorted as selected without asking (Microsoft Support,
 * "Expanding a selection while sorting data or removing duplicates").
 */
export function sortSelectionExpansion(
  cells: RegionCells,
  rect: DataRegionRange,
  rowCount: number,
  columnCount: number,
  merges: RegionMerges = [],
): DataRegionRange | null {
  if (
    ![rect.r1, rect.r2, rect.c1, rect.c2, rowCount, columnCount].every(Number.isSafeInteger) ||
    rect.r1 < 0 ||
    rect.c1 < 0 ||
    rect.r2 < rect.r1 ||
    rect.c2 < rect.c1
  )
    return null;
  const r2 = Math.min(rect.r2, rowCount - 1);
  const sides = [rect.c1 - 1, rect.c2 + 1].filter((c) => c >= 0 && c < columnCount);
  let seed: readonly [number, number] | null = null;
  if ((r2 - rect.r1 + 1) * sides.length <= Object.keys(cells).length + merges.length) {
    for (const c of sides) {
      for (let r = rect.r1; r <= r2 && !seed; r += 1) {
        if (regionCellOccupied(cells, merges, r, c)) seed = [r, c];
      }
    }
  } else {
    // A tall selection (whole columns) beside a sparse sheet: read the stored
    // cells and merges instead of every row.
    const candidates: Array<readonly [number, number]> = [];
    for (const key of Object.keys(cells)) {
      const separator = key.indexOf(":");
      candidates.push([Number(key.slice(0, separator)), Number(key.slice(separator + 1))]);
    }
    for (const merge of merges) {
      for (let c = merge.col; c < merge.col + merge.colSpan; c += 1) candidates.push([Math.max(merge.row, rect.r1), c]);
    }
    seed = candidates.find(([r, c]) =>
      r >= rect.r1 && r <= r2 && sides.includes(c) && regionCellOccupied(cells, merges, r, c)) ?? null;
  }
  if (!seed) return null;
  const region = surroundingDataRegion(cells, seed[0], seed[1], rowCount, columnCount, MAX_COMMAND_DATA_REGION_CELLS, merges);
  if (!region) return null;
  return {
    r1: Math.min(region.r1, rect.r1),
    c1: Math.min(region.c1, rect.c1),
    r2: Math.max(region.r2, r2),
    c2: Math.max(region.c2, rect.c2),
  };
}
