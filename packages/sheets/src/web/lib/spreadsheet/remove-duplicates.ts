/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helper for "Remove Duplicates" (Excel / WPS Data tab parity).
 *
 * Scans a matrix of rows and drops later rows whose chosen key columns match an
 * earlier row, keeping the FIRST-SEEN row in its original order (Excel's
 * behaviour). Equality is Excel's: two cells are duplicates when their
 * DISPLAYED values are the same text, compared without regard to case and
 * WITHOUT any other cleaning. Leading and trailing spaces are part of the
 * value ("foo", " foo" and "foo " are three records; trimming them is the
 * user's TRIM, not this command's), so this key is deliberately independent
 * of the sort / AutoFilter normalization in `sortKeyOf`, which trims. An
 * empty cell and an empty display string are the same blank; a cell holding
 * only spaces is not blank.
 *
 * TOTAL and side-effect-free: an empty `rows`, a single row, blanks, and an empty
 * `keyColumns` (no columns chosen → every row is its own empty key → every row is
 * a "duplicate" of the first, matching Excel, which requires at least one column)
 * all return cleanly without throwing. The caller decides how to write the kept
 * rows back (whole-row rewrite through the normal mutation path, so undo + autosave
 * + protection all fire).
 */

import type { SortFilterRow } from "./sort-filter";

/** The outcome of a de-duplication pass over a row matrix. */
export interface RemoveDuplicatesResult<T extends SortFilterRow> {
  /** The kept rows, first-seen order preserved (a subset of the input rows). */
  keptRows: T[];
  /** How many rows were dropped as duplicates (rows.length - keptRows.length). */
  removedCount: number;
}

/**
 * Build the equality key for one row from the chosen `keyColumns`. Each part is
 * the cell's display string exactly as shown (untrimmed; an empty cell is ""),
 * case-folded, joined with a NUL separator so values can't collide across the
 * column boundary (e.g. ["a","bc"] vs ["ab","c"], or ["a ","b"] vs ["a"," b"],
 * stay distinct).
 */
function rowKey(row: SortFilterRow, keyColumns: readonly number[]): string {
  const parts: string[] = [];
  for (const col of keyColumns) {
    parts.push((row[col]?.value ?? "").toLowerCase());
  }
  // A serialized array keeps every field boundary whatever the fields hold:
  // a NUL inside a value joined two different records into one key (B492).
  return JSON.stringify(parts);
}

/**
 * Remove duplicate rows by the chosen key columns, keeping the first occurrence
 * of each distinct key in input order. The input is never mutated.
 *
 * @param rows the data rows (exclude a header row BEFORE calling if relevant; the
 *   caller owns header handling)
 * @param keyColumns the column indices that define a row's identity (order does
 *   not matter; duplicates among them are harmless)
 * @returns the kept rows (first-seen order) and the count removed
 */
export function removeDuplicateRows<T extends SortFilterRow>(
  rows: readonly T[],
  keyColumns: readonly number[],
): RemoveDuplicatesResult<T> {
  const seen = new Set<string>();
  const keptRows: T[] = [];
  for (const row of rows) {
    const key = rowKey(row, keyColumns);
    if (seen.has(key)) continue;
    seen.add(key);
    keptRows.push(row);
  }
  return { keptRows, removedCount: rows.length - keptRows.length };
}
