/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helper for the multi-level "Sort" dialog (Excel / WPS "Sort by col A,
 * then by col B, then by col C" parity).
 *
 * Excel's Sort dialog lets the user stack N ordered sort levels: the rows are
 * ordered by the first level, ties are broken by the second, then the third, and
 * so on — a *lexicographic* compare over the chosen key columns, each with its
 * own ascending/descending direction. This module implements that as a STABLE
 * multi-key comparator over a plain row-major matrix so it can be unit-tested in
 * isolation and reused by the React grid (which bridges its `row:col` record into
 * this shape, exactly like the single-key sort / AutoFilter / Remove-Duplicates
 * paths do).
 *
 * It deliberately REUSES the single-key value comparator from `sort-filter.ts`
 * (`compareCells`): values order by type (numbers, text, FALSE, TRUE, errors),
 * blanks always sink to the bottom regardless of direction, and a per-level
 * direction flips the sign — so a one-level sort here is identical to a
 * `sortRows` call with the same text options. A level carries those options:
 * Excel's default sort is case-insensitive ("apple" == "Apple" for ordering),
 * and the dialog's "Case sensitive" toggle puts the lowercase form first
 * (measured in Excel: a | A | ab | Ab | b | B).
 *
 * Header handling (range extraction / reinsertion) lives in {@link splitHeader}
 * and {@link reattachHeader}: the caller pins the header row when "my data has
 * headers" is on, sorts only the data rows, then puts the header back on top.
 * Keeping that here makes the header math testable instead of buried in the grid.
 *
 * Everything is TOTAL and side-effect-free: empty rows, a single row, an empty
 * level list (no-op order), blanks and duplicate key columns across levels all
 * return cleanly without throwing. The input is never mutated.
 */

import {
  compareCells,
  sortKeyOf,
  type SortDirection,
  type SortFilterCell,
  type SortFilterRow,
} from "./sort-filter";
import type { ChineseCollation } from "@mrnote/core/web/lib/spreadsheet/locale-sort.ts";

export type { SortDirection, SortFilterCell, SortFilterRow };

/** The value or displayed format a sort level compares. */
export type SortOn = "values" | "cellColor" | "fontColor" | "icon";

/** Where the selected display format is grouped within its level. */
export type SortFormatPlacement = "first" | "last";

/** Display metadata resolved by the editor adapter for format-aware sorting. */
export interface SortDisplayKeyCell extends SortFilterCell {
  readonly fillColor?: string;
  readonly fontColor?: string;
  readonly iconKey?: string;
}

/**
 * One sort level: the key column (index into each row), its direction, and
 * whether text comparison at this level is case sensitive. `typeAware` mirrors
 * the single-key sort's flag (numbers compare numerically and before text); it
 * defaults to true and the dialog always sorts on values, so callers normally
 * omit it.
 */
export interface SortLevel {
  /** Key column index into each row (0-based, into the extracted matrix). */
  column: number;
  /** "asc" (A→Z) or "desc" (Z→A) for THIS level. */
  direction: SortDirection;
  /** Values by default; the other modes group one selected displayed format. */
  sortOn?: SortOn;
  /** Exact normalized fill/font/icon key selected for a format sort. */
  formatKey?: string;
  /** Selected formats sort first by default, or last when explicitly requested. */
  placement?: SortFormatPlacement;
  /**
   * When true, text that differs only in case orders lowercase first ("apple"
   * < "Apple"), Excel's case-sensitive sort. Default false = Excel's
   * case-insensitive sort, where they tie and keep their input order.
   */
  caseSensitive?: boolean;
  /**
   * Locale whose collation orders text (the UI locale, so a Chinese UI sorts
   * Chinese by pinyin as zh Excel does). Absent: the engine's default locale.
   */
  locale?: string;
  /** Explicit Chinese method (pinyin / stroke); absent = the locale's own. */
  method?: ChineseCollation;
  /**
   * When true (default), numbers compare numerically and sort before text;
   * blanks sink last. When false, everything compares as text. Matches
   * `sort-filter`'s `compareCells` `typeAware`.
   */
  typeAware?: boolean;
  /**
   * CUSTOM-LIST order (Excel "Sort ▸ Order ▸ Custom List"): an ordered list of
   * display strings. When present, cells are ordered by their POSITION in this
   * list rather than alphabetically/numerically — e.g. ["Low","Medium","High"]
   * sorts those three in that sequence regardless of spelling. Matching is
   * case-INSENSITIVE on the trimmed display text (Excel parity: "low" matches
   * "Low"). A value NOT in the list sorts AFTER every listed value, in the
   * normal ascending value order, and blanks stay last. "desc" is the list
   * read backwards (Excel's reversed-list Order entry): only the listed block
   * reverses; the unlisted tail stays below it, ascending. Absent/empty =
   * normal ordering.
   */
  orderKeys?: readonly string[];
}

/**
 * Build the case-insensitive position lookup for a custom-list order: each
 * trimmed-lowercased key → its first index in the list. Earlier entries win on
 * duplicates so the list's own order is authoritative. Returns null when the
 * list is empty (no custom order in effect).
 */
function customOrderIndex(
  orderKeys: readonly string[] | undefined,
): Map<string, number> | null {
  if (!orderKeys || orderKeys.length === 0) return null;
  const index = new Map<string, number>();
  for (let i = 0; i < orderKeys.length; i += 1) {
    const key = orderKeys[i].trim().toLowerCase();
    if (!index.has(key)) index.set(key, i);
  }
  return index.size > 0 ? index : null;
}

/**
 * Compare two cells at a single sort level, exactly as the single-key sort
 * does ({@link compareCells}): Excel's type order (numbers, text, FALSE, TRUE,
 * errors), blanks always last, direction flipping everything else. The level
 * supplies the text order:
 *
 *   - Case INSENSITIVE (default, Excel's default sort): "Apple" and "apple"
 *     are equal for ordering (a true tie → the caller's stable index tie-break
 *     preserves input order, measured in Excel as Apple | apple).
 *   - Case SENSITIVE: base letters still decide first (so "apple"/"Apple"
 *     group together), then the LOWERCASE form sorts first, measured in Excel
 *     as a | A | ab | Ab | b | B. Descending reverses it.
 *   - `locale` / `method`: the collation, so zh text orders by pinyin.
 */
export function compareCellsAtLevel(
  a: SortFilterCell | undefined,
  b: SortFilterCell | undefined,
  level: SortLevel,
): number {
  const sortOn = level.sortOn ?? "values";
  if (sortOn !== "values") {
    const target = level.formatKey?.trim().toLowerCase() ?? "";
    // Invalid format levels fail closed as stable no-ops. The UI rejects them
    // before execution; keeping this pure helper total prevents implicit value
    // sorting if a malformed recipe reaches it.
    if (target === "") return 0;
    const keyOf = (cell: SortFilterCell | undefined): string => {
      const displayCell = cell as SortDisplayKeyCell | undefined;
      if (sortOn === "cellColor") return displayCell?.fillColor ?? "";
      if (sortOn === "fontColor") return displayCell?.fontColor ?? "";
      return displayCell?.iconKey ?? "";
    };
    const aMatches = keyOf(a).trim().toLowerCase() === target;
    const bMatches = keyOf(b).trim().toLowerCase() === target;
    if (aMatches === bMatches) return 0;
    const selectedFirst = level.placement !== "last";
    return aMatches === selectedFirst ? -1 : 1;
  }
  const typeAware = level.typeAware ?? true;
  // Custom-list order (Excel "Order ▸ Custom List"): compare by list POSITION
  // first. Listed values order by their index, reversed for "desc" (the list
  // read backwards); every unlisted value, blanks included, sorts after the
  // listed block in the normal ASCENDING order, blanks last, in both
  // directions: a custom sort places the matching items first and the
  // remaining items in ascending order underneath.
  const customIndex = customOrderIndex(level.orderKeys);
  if (customIndex !== null) {
    const ra =
      sortKeyOf(a).kind === "empty"
        ? undefined
        : customIndex.get((a?.value ?? "").trim().toLowerCase());
    const rb =
      sortKeyOf(b).kind === "empty"
        ? undefined
        : customIndex.get((b?.value ?? "").trim().toLowerCase());
    if (ra !== undefined && rb !== undefined) {
      if (ra === rb) return 0;
      const cmp = ra < rb ? -1 : 1;
      return level.direction === "desc" ? -cmp : cmp;
    }
    if (ra !== undefined) return -1;
    if (rb !== undefined) return 1;
    return compareCells(a, b, "asc", typeAware, level);
  }
  return compareCells(a, b, level.direction, typeAware, level);
}

/**
 * Lexicographic multi-key compare across `levels`: compare on level 0, and only
 * fall through to level 1 (then 2, …) on a tie. Returns 0 when every level ties
 * (the caller's stable tie-break on original index then preserves input order).
 * An empty `levels` array means "no ordering" → always 0 (a stable no-op sort).
 */
export function compareRowsByLevels(
  a: SortFilterRow,
  b: SortFilterRow,
  levels: readonly SortLevel[],
): number {
  for (const level of levels) {
    const cmp = compareCellsAtLevel(a[level.column], b[level.column], level);
    if (cmp !== 0) return cmp;
  }
  return 0;
}

/**
 * Stable multi-key sort of a row matrix by the ordered `levels`, returning a NEW
 * array (input never mutated). Whole rows move together, so every column travels
 * with its keys. The sort is STABLE: rows that tie on every level keep their
 * original relative order — implemented with a decorate-sort-undecorate that
 * tie-breaks on the original index, so stability does not depend on the host
 * engine's `Array.prototype.sort` guarantees.
 *
 * @param rows the rows to sort (row-major; cells may be undefined)
 * @param levels the ordered sort levels (level 0 is primary). Empty = no-op.
 */
export function multiSortRows<T extends SortFilterRow>(
  rows: readonly T[],
  levels: readonly SortLevel[],
): T[] {
  const decorated = rows.map((row, index) => ({ row, index }));
  decorated.sort((a, b) => {
    const cmp = compareRowsByLevels(a.row, b.row, levels);
    return cmp !== 0 ? cmp : a.index - b.index;
  });
  return decorated.map((entry) => entry.row);
}

/**
 * Compute the source-row PERMUTATION a multi-key sort produces, given the rows'
 * original source-row indices. This is the grid-facing entry point: the caller
 * builds `rows` (the resolved key values for the data span) paired with the
 * matching `sourceRows` (their actual sheet row indices), and gets back the
 * source rows in their sorted order, ready to re-lay-out the span.
 *
 * `rows[i]` and `sourceRows[i]` MUST line up (same length); a mismatch is
 * tolerated by pairing up to the shorter length. The sort is stable, so a sheet
 * that is already in order (or has only ties) yields `sourceRows` unchanged.
 *
 * @param rows the data rows' resolved key cells (row-major)
 * @param sourceRows the sheet row index of each `rows[i]`
 * @param levels the ordered sort levels
 * @returns the source row indices in sorted order
 */
export function sortedSourceOrder(
  rows: readonly SortFilterRow[],
  sourceRows: readonly number[],
  levels: readonly SortLevel[],
): number[] {
  const n = Math.min(rows.length, sourceRows.length);
  const decorated: Array<{ row: SortFilterRow; source: number; index: number }> =
    [];
  for (let i = 0; i < n; i += 1) {
    decorated.push({ row: rows[i], source: sourceRows[i], index: i });
  }
  decorated.sort((a, b) => {
    const cmp = compareRowsByLevels(a.row, b.row, levels);
    return cmp !== 0 ? cmp : a.index - b.index;
  });
  return decorated.map((entry) => entry.source);
}

/**
 * Compute the source-COLUMN permutation a left-to-right (sort by row) multi-key
 * sort produces — the Excel "Sort ▸ Options ▸ Sort left to right" orientation.
 * Where {@link sortedSourceOrder} permutes ROWS keyed by columns, this permutes
 * COLUMNS keyed by ROWS: each "unit" is a spreadsheet column, and a sort level's
 * `column` field is reinterpreted as the KEY ROW offset to compare on.
 *
 * `columns[i]` is the resolved cell values of source column `sourceCols[i]`, laid
 * out as a row-major single "row" whose entries are indexed by ROW offset (so
 * `columns[i][rowOffset]` is the value at that row in that column). The levels'
 * `column` is the row offset to order by; everything else (direction, case,
 * type-aware, stability) behaves exactly like the row sort. Returns the source
 * column indices in their sorted left-to-right order.
 *
 * @param columns each source column's key cells, indexed by row offset
 * @param sourceCols the sheet column index of each `columns[i]`
 * @param levels the ordered sort levels (their `column` = key ROW offset)
 */
export function sortedColumnOrder(
  columns: readonly SortFilterRow[],
  sourceCols: readonly number[],
  levels: readonly SortLevel[],
): number[] {
  // Identical mechanics to sortedSourceOrder; the only difference is the caller's
  // interpretation of a "row" as a column and a level's `column` as a key row.
  return sortedSourceOrder(columns, sourceCols, levels);
}

/** A range split into an optional pinned header row and the data rows below it. */
export interface HeaderSplit<T> {
  /** The header row, or null when `hasHeader` is false / the range is empty. */
  header: T | null;
  /** The data rows (everything below the header). */
  data: T[];
}

/**
 * Split a range's rows into a pinned header (row 0 when `hasHeader`) and the data
 * rows. With `hasHeader` false — or fewer than one row — there is no header and
 * every row is data. The input array is not mutated (data is a shallow copy).
 */
export function splitHeader<T>(
  rows: readonly T[],
  hasHeader: boolean,
): HeaderSplit<T> {
  if (!hasHeader || rows.length === 0) {
    return { header: null, data: rows.slice() };
  }
  return { header: rows[0], data: rows.slice(1) };
}

/**
 * Re-attach a pinned header (if any) on top of the sorted data rows — the inverse
 * of {@link splitHeader}. A null header just returns the data rows. This keeps the
 * header verbatim at the top of the range, exactly where Excel leaves it.
 */
export function reattachHeader<T>(header: T | null, data: readonly T[]): T[] {
  return header === null ? data.slice() : [header, ...data];
}
