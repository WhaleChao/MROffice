/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface OfficeMergedRows {
  /** The merge's first row. */
  readonly start: number;
  /** Its last row (inclusive). */
  readonly end: number;
  /** The height its content needs across all its rows, margins included. */
  readonly need: number;
}

/**
 * `heights` (each row's own height) with every merge's shortfall added to its
 * last row. Rows out of range are ignored; the input is not changed.
 */
export function officeRowsUnderMerges(
  heights: readonly number[],
  merges: readonly OfficeMergedRows[],
): number[] {
  const out = heights.map((height) => (Number.isFinite(height) && height > 0 ? height : 0));
  const settled = merges
    .filter((merge) => merge.start >= 0 && merge.end < out.length && merge.end >= merge.start && Number.isFinite(merge.need))
    .slice()
    .sort((a, b) => a.end - b.end);
  for (const { start, end, need } of settled) {
    let total = 0;
    for (let row = start; row <= end; row += 1) total += out[row]!;
    if (need > total) out[end]! += need - total;
  }
  return out;
}
