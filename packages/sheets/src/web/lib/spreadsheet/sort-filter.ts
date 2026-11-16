/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for spreadsheet range sort + AutoFilter (Excel / WPS parity).
 *
 * These functions operate on a plain row-major matrix of cells so they can be
 * unit-tested in isolation and reused by the React grid (which stores cells in a
 * `row:col` record and bridges into this shape). Two distinct operations live
 * here:
 *
 *   - SORT: permute whole rows by a key column in Excel's type order: numbers,
 *     text, logical values (FALSE before TRUE), errors, and blanks, which sink
 *     to the bottom regardless of direction. The sort is *stable* — equal keys
 *     keep their original relative
 *     order — so a multi-pass sort behaves predictably and every other column in
 *     a row travels with its key (formulas included; the caller decides how to
 *     relocate any formula references, see SpreadsheetWindow's sort wiring).
 *
 *   - FILTER (AutoFilter): given a column and a set of allowed display values,
 *     compute which row indices stay visible. Filtering is a *view* operation —
 *     it never mutates or deletes data, it only reports which rows to hide.
 *
 * Filtering treats display text as its source of truth. Sorting accepts a
 * separate authenticated numeric channel so formatted values such as `$1,200`
 * and `28%` still compare by their raw/computed scalar. Callers that do not
 * provide that channel retain the legacy numeric-text fallback. Logical values
 * and errors come the same way, through `valueType`: display text alone never
 * makes a cell a logical or an error, so a text '#text' stays text.
 */

import {
  compareSpreadsheetSortText,
  type SpreadsheetTextOrder,
} from "@mrnote/core/web/lib/spreadsheet/locale-sort.ts";

export type { SpreadsheetTextOrder };

export type SortDirection = "asc" | "desc";

/** A single cell as far as sort/filter logic cares. */
export interface SortFilterCell {
  /** Display text used by filtering, labels, and non-numeric sorting. */
  value?: string;
  /** Authenticated raw/computed numeric scalar; never parsed from formatted UI. */
  numericValue?: number;
  /**
   * Authenticated computed type of a value that is not a number: a logical
   * (TRUE/FALSE), an error (#N/A, #DIV/0!, ...), or text. Supplied by the
   * evaluator, like `numericValue`, never guessed from display text. A cell
   * typed "text" is text even when it reads as a number ('2, a Text (@) cell,
   * a formula returning "100"): it sorts among the words, as Excel sorts it,
   * and never through the display-text number parse below, which exists only
   * for callers that supply no type at all.
   */
  valueType?: "boolean" | "error" | "text";
  /** The logical value when `valueType` is "boolean". */
  booleanValue?: boolean;
}

/** A row is a (possibly sparse) array of cells; `undefined` means an empty cell. */
export type SortFilterRow = ReadonlyArray<SortFilterCell | undefined>;

export type SortKeyKind = "number" | "string" | "boolean" | "error" | "empty";

interface SortKey {
  kind: SortKeyKind;
  /** Numeric value; 0/1 for FALSE/TRUE. */
  n: number;
  s: string;
  authoritativeNumeric: boolean;
}

/**
 * Excel's ascending order of value types (Microsoft Support, "Sort data in a
 * range or table"): numbers, text, logical values, errors. Blanks come after
 * all of them in either direction and are handled apart.
 */
const SORT_TYPE_RANK: Record<Exclude<SortKeyKind, "empty">, number> = {
  number: 0,
  string: 1,
  boolean: 2,
  error: 3,
};

/** How a comparison orders text; see {@link SpreadsheetTextOrder}. */
export type CompareCellsOptions = SpreadsheetTextOrder;










export function sortKeyOf(cell: SortFilterCell | undefined): SortKey {
  const text = cell?.value ?? "";
  const raw = text.trim();
  if (cell?.valueType === "error") {
    return { kind: "error", n: 0, s: raw, authoritativeNumeric: false };
  }
  if (cell?.valueType === "boolean") {
    return {
      kind: "boolean",
      n: cell.booleanValue === true ? 1 : 0,
      s: raw,
      authoritativeNumeric: false,
    };
  }
  if (cell?.valueType === "text") {
    return text === ""
      ? { kind: "empty", n: 0, s: "", authoritativeNumeric: false }
      : { kind: "string", n: 0, s: text, authoritativeNumeric: false };
  }
  if (
    typeof cell?.numericValue === "number" &&
    Number.isFinite(cell.numericValue)
  ) {
    return {
      kind: "number",
      n: cell.numericValue,
      s: raw || String(cell.numericValue),
      authoritativeNumeric: true,
    };
  }
  if (text === "") {
    return { kind: "empty", n: 0, s: "", authoritativeNumeric: false };
  }
  if (raw === "") {
    return { kind: "string", n: 0, s: text, authoritativeNumeric: false };
  }
  const n = Number(raw);
  if (Number.isFinite(n) && raw !== "") {
    return { kind: "number", n, s: raw, authoritativeNumeric: false };
  }
  return {
    kind: "string",
    n: 0,
    s: text,
    authoritativeNumeric: false,
  };
}

/**
 * Compare two cells for sorting. When `typeAware` is true (Excel default),
 * values order by type first, ascending: numbers (numerically), text, FALSE,
 * TRUE, errors (all equal, so a stable sort keeps their order); otherwise
 * everything compares as text. Descending reverses all of that, but blanks
 * always sort last, independent of `direction` (Excel parity). Text compares
 * case-insensitively unless `options.caseSensitive`, in the collation of
 * `options.locale` / `options.method` (see compareSpreadsheetText).
 * Returns a direction-applied comparator value (negative => a before b).
 */
export function compareCells(
  a: SortFilterCell | undefined,
  b: SortFilterCell | undefined,
  direction: SortDirection,
  typeAware: boolean,
  options?: CompareCellsOptions,
): number {
  const ka = sortKeyOf(a);
  const kb = sortKeyOf(b);
  // Blanks last, no matter the sort direction.
  if (ka.kind === "empty" && kb.kind === "empty") return 0;
  if (ka.kind === "empty") return 1;
  if (kb.kind === "empty") return -1;

  let cmp: number;
  if (!typeAware) {
    cmp = compareSpreadsheetSortText(ka.s, kb.s, options);
  } else if (ka.kind !== kb.kind) {
    cmp = SORT_TYPE_RANK[ka.kind] - SORT_TYPE_RANK[kb.kind];
  } else if (ka.kind === "boolean") {
    cmp = ka.n - kb.n; // FALSE before TRUE
  } else if (ka.kind === "error") {
    cmp = 0; // errors tie; the stable sort keeps them in input order
  } else if (ka.kind === "number") {
    cmp = ka.n - kb.n;
    // Deterministic tiebreak: two values that are numerically equal but written
    // differently ("1" vs "1.0", "01" vs "1") fall back to a string compare so
    // the order is stable across engines instead of leaning on a no-op. Explicit
    // numeric authority deliberately stays tied: formatting must not reorder
    // equal raw/computed values.
    if (
      cmp === 0 &&
      !ka.authoritativeNumeric &&
      !kb.authoritativeNumeric
    ) {
      cmp = ka.s.localeCompare(kb.s);
    }
  } else {
    cmp = compareSpreadsheetSortText(ka.s, kb.s, options);
  }
  return direction === "desc" ? -cmp : cmp;
}

/**
 * Sort a matrix of rows by `colIndex`, returning a NEW array of rows (the input
 * is not mutated). Whole rows move together, so every column travels with its
 * key. The sort is stable: rows whose key compares equal keep their original
 * relative order. Blanks sink to the bottom regardless of direction.
 *
 * @param rows the rows to sort (row-major; cells may be undefined)
 * @param colIndex the key column
 * @param dir "asc" | "desc"
 * @param typeAware when true (default), values order by type (numbers, text,
 *   logicals, errors); when false, all values compare as text
 * @param options text order: case sensitivity, locale and Chinese method
 */
export function sortRows<T extends SortFilterRow>(
  rows: readonly T[],
  colIndex: number,
  dir: SortDirection,
  typeAware = true,
  options?: CompareCellsOptions,
): T[] {
  // Decorate-sort-undecorate to keep the sort stable across engines (JS Array
  // sort stability is only guaranteed for modern engines, but tie-breaking on
  // the original index makes the intent explicit and robust).
  const decorated = rows.map((row, index) => ({ row, index }));
  decorated.sort((a, b) => {
    const cmp = compareCells(
      a.row[colIndex],
      b.row[colIndex],
      dir,
      typeAware,
      options,
    );
    return cmp !== 0 ? cmp : a.index - b.index;
  });
  return decorated.map((entry) => entry.row);
}











export function filterValueKey(value: string): string {
  return filterValueText(value).toLowerCase();
}



export function filterValueText(value: string): string {
  return value.replace(/^ +| +$/g, "");
}

/** The ticked values of one filter as the keys cells are matched by. */
export function filterValueKeys(values: Iterable<string>): Set<string> {
  const keys = new Set<string>();
  for (const value of values) keys.add(filterValueKey(String(value)));
  return keys;
}

/**
 * Compute the ordered set of distinct display values in a column, for the
 * AutoFilter dropdown. Blanks collapse to a single empty-string entry. Values
 * keep their first-seen order so the dropdown matches the data's natural order;
 * numbers and text are returned as their trimmed display strings.
 *
 * @param rows the data rows (exclude the header row before calling if relevant)
 * @param colIndex the column to enumerate
 */
export function distinctColumnValues<T extends SortFilterRow>(
  rows: readonly T[],
  colIndex: number,
): string[] {
  
  // value, a cell of spaces is a blank.
  const seen = new Set<string>();
  const order: string[] = [];
  for (const row of rows) {
    const raw = row[colIndex]?.value ?? "";
    const key = filterValueKey(raw);
    if (!seen.has(key)) {
      seen.add(key);
      order.push(key === "" ? "" : filterValueText(raw));
    }
  }
  return order;
}

/**
 * Order a set of distinct display values the way Excel's AutoFilter list box
 * does (Microsoft Support, "Filter data in a range or table"): the checkbox
 * list is SORTED, not in the data's order — numbers ascending, then text A→Z,
 * then FALSE and TRUE, then errors, with "(Blanks)" always at the bottom. That
 * is exactly the ascending type order {@link compareCells} already implements,
 * so the list and a 升序 sort of the same column agree.
 *
 * `representative` supplies the authenticated key of one cell per value, so a
 * column of `$1,200` / `28%` orders by its scalar and TRUE / #N/A order as a
 * logical and an error. Values with no entry compare on their text alone.
 *
 * @param values distinct display values; "" is the blanks bucket
 * @param representative value -> one cell that displays it
 * @param options text order: case sensitivity, locale and Chinese method
 */
export function orderFilterValues(
  values: readonly string[],
  representative?: ReadonlyMap<string, SortFilterCell>,
  options?: CompareCellsOptions,
): string[] {
  const keyOf = (value: string): SortFilterCell =>
    representative?.get(value) ?? { value };
  return values
    .map((value, index) => ({ value, index }))
    .sort((left, right) => {
      const cmp = compareCells(
        keyOf(left.value),
        keyOf(right.value),
        "asc",
        true,
        options,
      );
      return cmp !== 0 ? cmp : left.index - right.index;
    })
    .map((entry) => entry.value);
}

/**
 * The distinct display values of a column IN EXCEL'S LIST ORDER — what the
 * AutoFilter dropdown shows. Same set as {@link distinctColumnValues}, sorted
 * by {@link orderFilterValues} using the first cell seen for each value as its
 * authenticated key, so numbers, logicals and errors keep their type.
 *
 * @param rows the data rows (exclude the header row before calling)
 * @param colIndex the column to enumerate
 * @param options text order: case sensitivity, locale and Chinese method
 */
export function filterListValues<T extends SortFilterRow>(
  rows: readonly T[],
  colIndex: number,
  options?: CompareCellsOptions,
): string[] {
  const representative = new Map<string, SortFilterCell>();
  const keys = new Set<string>();
  const order: string[] = [];
  for (const row of rows) {
    const cell = row[colIndex];
    const raw = cell?.value ?? "";
    const key = filterValueKey(raw);
    if (keys.has(key)) continue;
    keys.add(key);
    
    const value = key === "" ? "" : filterValueText(raw);
    representative.set(value, cell && value === raw ? cell : { ...(cell ?? {}), value });
    order.push(value);
  }
  return orderFilterValues(order, representative, options);
}

/**
 * Apply a column filter, returning the indices (into `rows`) of the rows whose
 * value in `colIndex` is a member of `allowedSet`. Values are compared by their
 * filter key ({@link filterValueKey}, matching distinctColumnValues): "alpha",
 * " alpha" and "alpha " are one value, as Excel matches them, and a
 * blank (or a cell of spaces) is the empty string "". A `null` / undefined `allowedSet` means "no filter" — every
 * row index is returned. This NEVER mutates `rows`; the caller hides the
 * excluded rows from the view.
 *
 * @returns the visible row indices, in ascending order
 */
export function applyFilter<T extends SortFilterRow>(
  rows: readonly T[],
  colIndex: number,
  allowedSet: ReadonlySet<string> | null | undefined,
): number[] {
  const visible: number[] = [];
  const keys = allowedSet ? filterValueKeys(allowedSet) : null;
  for (let i = 0; i < rows.length; i += 1) {
    if (!keys) {
      visible.push(i);
      continue;
    }
    if (keys.has(filterValueKey(rows[i][colIndex]?.value ?? ""))) visible.push(i);
  }
  return visible;
}

/**
 * Combine several per-column filters (column -> allowed value set) into the set
 * of visible row indices. A row is visible only when it passes EVERY active
 * column filter (AND semantics, like Excel's stacked AutoFilters). Columns with
 * no entry — or an entry whose allowed set is null — impose no constraint.
 *
 * @returns the visible row indices, in ascending order
 */
export function visibleRowIndices<T extends SortFilterRow>(
  rows: readonly T[],
  filters: ReadonlyMap<number, ReadonlySet<string> | null>,
): number[] {
  const active: Array<[number, ReadonlySet<string>]> = [];
  for (const [col, set] of filters) {
    if (set) active.push([col, filterValueKeys(set)]);
  }
  if (active.length === 0) return rows.map((_, i) => i);
  const visible: number[] = [];
  for (let i = 0; i < rows.length; i += 1) {
    let ok = true;
    for (const [col, set] of active) {
      const value = filterValueKey(rows[i][col]?.value ?? "");
      if (!set.has(value)) {
        ok = false;
        break;
      }
    }
    if (ok) visible.push(i);
  }
  return visible;
}

// ---------------------------------------------------------------------------
// Sort + filter BY (fill) COLOR — Excel / WPS "Sort by Color" / "Filter by
// Color" parity. These are ADDITIVE to the value-based machinery above: they
// read a cell's fill color instead of its display value, but never change how
// the value comparators / filters behave. A row whose key cell has no fill is
// grouped into a single well-defined "no fill" bucket (the empty-string color),
// so nothing is ever NaN-compared or thrown.
// ---------------------------------------------------------------------------

/**
 * A cell as far as COLOR sort/filter logic cares: just its fill (background)
 * color string, exactly as stored on `cell.style.backgroundColor` (which may be
 * any CSS color form — `#FFF`, `#ffffff`, `rgb(...)`, `transparent`, …). A
 * missing / undefined fill means "no fill".
 */
export interface ColorCell {
  /** The cell's raw fill color, e.g. `cell.style?.backgroundColor`. */
  fill?: string;
}

/** A row for color sort/filter: a (possibly sparse) array of {@link ColorCell}. */
export type ColorRow = ReadonlyArray<ColorCell | undefined>;

/** The canonical key for the "no fill / transparent" color bucket. */
export const NO_FILL_COLOR = "";

const HEX3_RE = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i;
const HEX6_RE = /^#([0-9a-f]{6})$/i;
const HEX8_RE = /^#([0-9a-f]{6})([0-9a-f]{2})$/i;
const RGB_RE =
  /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([0-9.]+)\s*)?\)$/i;

function clampByte(n: number): number {
  if (!Number.isFinite(n)) return 0;
  return Math.max(0, Math.min(255, Math.round(n)));
}

function toHex2(n: number): string {
  return clampByte(n).toString(16).padStart(2, "0");
}

/**
 * Normalize a fill color so equal colors written differently compare equal.
 *
 * Equality rule (documented contract):
 *   - `undefined`, `""`, `"transparent"`, `"none"` (any case) → {@link NO_FILL_COLOR}
 *     (the empty string), the single "no fill" bucket.
 *   - 3-digit hex `#FFF` and 6-digit hex `#FFFFFF` → lowercase 6-digit `#ffffff`
 *     (so `#FFF` === `#ffffff` === `#FFFFFF`).
 *   - 8-digit hex `#rrggbbaa`: a FULLY opaque alpha (`ff`) collapses to `#rrggbb`;
 *     any other alpha keeps the alpha suffix so semi-transparent fills stay
 *     distinct.
 *   - `rgb()` / `rgba()` → 6-digit hex `#rrggbb` (a fully opaque or absent alpha)
 *     or `#rrggbbaa` (a non-opaque alpha), matching the hex forms above.
 *   - Anything else (a named color like `red`, a gradient, …): lowercased +
 *     trimmed and returned as-is, so two identical raw strings still match.
 */
export function normalizeFillColor(raw: string | undefined | null): string {
  if (raw === undefined || raw === null) return NO_FILL_COLOR;
  const s = String(raw).trim().toLowerCase();
  if (s === "" || s === "transparent" || s === "none") return NO_FILL_COLOR;

  const hex3 = HEX3_RE.exec(s);
  if (hex3) {
    return `#${hex3[1]}${hex3[1]}${hex3[2]}${hex3[2]}${hex3[3]}${hex3[3]}`;
  }
  if (HEX6_RE.test(s)) return s;
  const hex8 = HEX8_RE.exec(s);
  if (hex8) {
    // Drop a fully-opaque alpha so #rrggbbff === #rrggbb; keep other alphas.
    return hex8[2] === "ff" ? `#${hex8[1]}` : s;
  }
  const rgb = RGB_RE.exec(s);
  if (rgb) {
    const hex = `#${toHex2(Number(rgb[1]))}${toHex2(Number(rgb[2]))}${toHex2(
      Number(rgb[3]),
    )}`;
    const alpha = rgb[4];
    if (alpha === undefined) return hex;
    const a = Number(alpha);
    if (!Number.isFinite(a) || a >= 1) return hex; // fully opaque → #rrggbb
    if (a <= 0) return NO_FILL_COLOR; // fully transparent → no fill
    return `${hex}${toHex2(a * 255)}`;
  }
  return s; // named color / gradient / unknown: compare by exact (lowercased) text
}

/** Read + normalize a color cell's fill. Missing cell → the no-fill bucket. */
export function fillColorOf(cell: ColorCell | undefined): string {
  return normalizeFillColor(cell?.fill);
}

/**
 * The ordered set of distinct NORMALIZED fill colors present in a column, for
 * the "filter by color" picker. The no-fill bucket ({@link NO_FILL_COLOR}) is
 * included when any cell in the column has no fill. Colors keep first-seen order
 * so the picker matches the data's natural top-to-bottom order; the no-fill
 * bucket — if present — is always pushed LAST so it reads like Excel's
 * "No Fill" entry at the bottom of the swatch list.
 *
 * @param rows the data rows (exclude the header row before calling if relevant)
 * @param colIndex the column to enumerate
 */
export function distinctColumnColors<T extends ColorRow>(
  rows: readonly T[],
  colIndex: number,
): string[] {
  const seen = new Set<string>();
  const order: string[] = [];
  let hasNoFill = false;
  for (const row of rows) {
    const color = fillColorOf(row[colIndex]);
    if (color === NO_FILL_COLOR) {
      hasNoFill = true;
      continue;
    }
    if (!seen.has(color)) {
      seen.add(color);
      order.push(color);
    }
  }
  if (hasNoFill) order.push(NO_FILL_COLOR);
  return order;
}

/**
 * Stable "sort by fill color": order rows so a CHOSEN color comes first, then
 * every other row keeps its original relative order. Whole rows move together
 * (the caller maps the result back through a source-index column, exactly like
 * {@link sortRows}). The chosen color is normalized, so passing `#FFF` matches a
 * cell filled `#ffffff`. Passing {@link NO_FILL_COLOR} brings the no-fill rows to
 * the top. This NEVER mutates the input and is deterministic.
 *
 * Rationale for "chosen color first" (vs. a full color ordering): Excel's
 * "Sort by Color" lets the user pick ONE color to float to the top of the
 * range; that is the common color-coded-sheet workflow. Equal-bucket rows (the
 * chosen color among themselves, and all the rest among themselves) keep their
 * relative order, so the sort is stable and a uniform-color column is a no-op.
 *
 * @param rows the rows to sort (row-major; cells may be undefined)
 * @param colIndex the key column whose fill color drives the order
 * @param chosenColor the color to bring to the top (any CSS form; normalized)
 */
export function sortRowsByColor<T extends ColorRow>(
  rows: readonly T[],
  colIndex: number,
  chosenColor: string,
): T[] {
  const target = normalizeFillColor(chosenColor);
  // Decorate-sort-undecorate so the partition is provably stable: rows compare
  // only on "is this the chosen color (0) or not (1)" and tie-break on the
  // original index, so the relative order inside each bucket is preserved.
  const decorated = rows.map((row, index) => ({
    row,
    index,
    rank: fillColorOf(row[colIndex]) === target ? 0 : 1,
  }));
  decorated.sort((a, b) => (a.rank !== b.rank ? a.rank - b.rank : a.index - b.index));
  return decorated.map((entry) => entry.row);
}

/**
 * Apply a per-column COLOR filter, returning the indices (into `rows`) of the
 * rows whose normalized fill color in `colIndex` is a member of `allowedSet`.
 * `allowedSet` holds NORMALIZED colors (see {@link normalizeFillColor}); the
 * no-fill bucket is {@link NO_FILL_COLOR}. A `null`/undefined set means "no
 * filter" — every row index is returned. NEVER mutates `rows`.
 *
 * @returns the visible row indices, in ascending order
 */
export function applyColorFilter<T extends ColorRow>(
  rows: readonly T[],
  colIndex: number,
  allowedSet: ReadonlySet<string> | null | undefined,
): number[] {
  const visible: number[] = [];
  for (let i = 0; i < rows.length; i += 1) {
    if (!allowedSet) {
      visible.push(i);
      continue;
    }
    if (allowedSet.has(fillColorOf(rows[i][colIndex]))) visible.push(i);
  }
  return visible;
}

/**
 * Combine several per-column COLOR filters (column -> allowed normalized-color
 * set) into the set of visible row indices. AND semantics, mirroring
 * {@link visibleRowIndices}: a row is visible only when it passes EVERY active
 * color filter. Columns with no entry — or a null set — impose no constraint.
 *
 * @returns the visible row indices, in ascending order
 */
export function visibleRowIndicesByColor<T extends ColorRow>(
  rows: readonly T[],
  filters: ReadonlyMap<number, ReadonlySet<string> | null>,
): number[] {
  const active: Array<[number, ReadonlySet<string>]> = [];
  for (const [col, set] of filters) {
    if (set) active.push([col, set]);
  }
  if (active.length === 0) return rows.map((_, i) => i);
  const visible: number[] = [];
  for (let i = 0; i < rows.length; i += 1) {
    let ok = true;
    for (const [col, set] of active) {
      if (!set.has(fillColorOf(rows[i][col]))) {
        ok = false;
        break;
      }
    }
    if (ok) visible.push(i);
  }
  return visible;
}
