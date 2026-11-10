/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How much a clipboard paste may bring into the grid at once. Excel pastes
 * whatever fits the sheet; this bound keeps one paste's parse and write
 * within a budget the editor can commit as a single undo step. A paste that
 * exceeds it is cut to the bound and the caller says so: a silent cut lost
 * the tail of a 1,001-row paste with a "saved" status.
 */
export const PASTE_MAX_ROWS = 100_000;
/** Excel's column limit (XFD). */
export const PASTE_MAX_COLS = 16_384;

export interface LimitedPastedGrid<T> {
  rows: T[][];
  /** True when rows or columns beyond the bound were dropped. */
  truncated: boolean;
}

export function limitPastedGrid<T>(
  rows: readonly (readonly T[])[],
  maxRows = PASTE_MAX_ROWS,
  maxCols = PASTE_MAX_COLS,
): LimitedPastedGrid<T> {
  let truncated = rows.length > maxRows;
  const limited = rows.slice(0, maxRows).map((row) => {
    if (row.length > maxCols) truncated = true;
    return row.slice(0, maxCols);
  });
  return { rows: limited, truncated };
}
