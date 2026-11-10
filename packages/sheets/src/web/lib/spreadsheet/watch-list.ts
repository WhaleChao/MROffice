/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Watch Window (Excel Formulas ▸ Watch Window parity): a small, deduplicated
 * list of cells the user pins so their live value stays visible in a side panel
 * while editing far-away cells in a large sheet.
 *
 * Pure + dependency-free: the model stores only coordinates; the value + formula
 * are re-derived at render from the formula evaluator (never persisted). The
 * runtime re-clamps a loaded list via {@link normalizeWatches} so an out-of-range
 * or malformed entry from a hand-edited file is dropped, and the count is capped.
 */

export interface WatchedCell {
  row: number;
  col: number;
}

/** The hard ceiling on watched cells per sheet (mirrors the schema `.max(64)`). */
export const MAX_WATCHES = 64;

function sameCell(a: WatchedCell, b: WatchedCell): boolean {
  return a.row === b.row && a.col === b.col;
}

/**
 * Validate + dedupe a (possibly untrusted) watch list against the grid bounds.
 * Non-integer / negative / out-of-range / duplicate entries are dropped, order
 * is preserved, and the result is capped at {@link MAX_WATCHES}.
 */
export function normalizeWatches(
  raw: unknown,
  bounds: { rowCount: number; columnCount: number },
): WatchedCell[] {
  if (!Array.isArray(raw)) return [];
  const rowCount = Math.max(0, Math.trunc(bounds.rowCount));
  const columnCount = Math.max(0, Math.trunc(bounds.columnCount));
  const out: WatchedCell[] = [];
  for (const entry of raw) {
    if (out.length >= MAX_WATCHES) break;
    const row = (entry as WatchedCell)?.row;
    const col = (entry as WatchedCell)?.col;
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(col) ||
      row < 0 ||
      col < 0 ||
      row >= rowCount ||
      col >= columnCount
    ) {
      continue;
    }
    const cell = { row, col };
    if (out.some((w) => sameCell(w, cell))) continue;
    out.push(cell);
  }
  return out;
}

/** Whether `cell` is already watched. */
export function hasWatch(list: readonly WatchedCell[], cell: WatchedCell): boolean {
  return list.some((w) => sameCell(w, cell));
}

/**
 * Add `cell` to the watch list (idempotent; no-op if already present or the list
 * is full). Returns a new array — the input is never mutated.
 */
export function addWatch(
  list: readonly WatchedCell[],
  cell: WatchedCell,
): WatchedCell[] {
  if (
    !Number.isInteger(cell.row) ||
    !Number.isInteger(cell.col) ||
    cell.row < 0 ||
    cell.col < 0
  ) {
    return [...list];
  }
  if (hasWatch(list, cell) || list.length >= MAX_WATCHES) return [...list];
  return [...list, { row: cell.row, col: cell.col }];
}

/** Remove `cell` from the watch list (no-op if absent). Returns a new array. */
export function removeWatch(
  list: readonly WatchedCell[],
  cell: WatchedCell,
): WatchedCell[] {
  return list.filter((w) => !sameCell(w, cell));
}
