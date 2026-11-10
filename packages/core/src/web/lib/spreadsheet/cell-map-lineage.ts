/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where an immutable cell map came from: an accepted edit records the map
 * it replaced and the keys it changed, so a value derived from a sheet's
 * cells (stored formula results, key lists) can be carried from the map
 * before and brought up to date at those keys, instead of walking every
 * stored cell again on every edit.
 *
 * Only a hint: the map before is held weakly, a chain that is broken or too
 * long answers null, and a consumer then derives from the whole map as it
 * did before.
 */
type CellMapLike = object;

const parents = new WeakMap<CellMapLike, { readonly before: WeakRef<CellMapLike>; readonly keys: readonly string[] }>();
/** Steps one lookup follows back before giving up. */
const MAX_CHAIN = 64;
/** Changed keys one lookup gathers before a whole walk is cheaper. */
const MAX_KEYS = 50_000;

/** An accepted edit replaced `before` with `after`, changing exactly `keys`. */
export function recordCellMapEdit(before: CellMapLike, after: CellMapLike, keys: readonly string[]): void {
  if (before === after || parents.has(after) || typeof WeakRef === "undefined") return;
  parents.set(after, { before: new WeakRef(before), keys: [...keys] });
}

/**
 * The keys at which `after` differs from `before`, by the cells' identity (an edit
 * makes a new object of a cell it changes and carries the others over), or null
 * past `limit`, when the two are better read whole. One walk here spares a commit
 * that does not know its keys (a format, Delete, a fill) a walk by everything it
 * feeds: the index, the extent, the formulas, the fonts, the change log.
 */
export function cellMapChangedKeys(
  before: Readonly<Record<string, unknown>>,
  after: Readonly<Record<string, unknown>>,
  limit: number,
): string[] | null {
  const keys: string[] = [];
  let kept = 0;
  for (const key in after) {
    const was = before[key];
    if (was !== undefined) kept += 1;
    if (was === after[key]) continue;
    keys.push(key);
    if (keys.length > limit) return null;
  }
  // Keys only `before` has, which the count says when there are none.
  if (kept !== Object.keys(before).length) {
    for (const key in before) {
      if (after[key] !== undefined) continue;
      keys.push(key);
      if (keys.length > limit) return null;
    }
  }
  return keys;
}

/** As {@link recordCellMapEdit}, for a map that replaces one nothing else
 * holds: it stands in the chain for the map before that, which the weak link
 * through a short-lived map would lose to the next collection. */
export function recordCellMapEditOver(before: CellMapLike, after: CellMapLike, keys: readonly string[]): void {
  const parent = parents.get(before);
  const earlier = parent?.before.deref();
  if (parent && earlier) recordCellMapEdit(earlier, after, [...parent.keys, ...keys]);
  else recordCellMapEdit(before, after, keys);
}

/**
 * A row sort's new map: the rows of `dataR1`.. moved within columns `c1`..`c2`,
 * `sourceOrder[offset]` being the row now at `dataR1 + offset`. Every cell is
 * the same object as before but those named in `relocated` (their new keys),
 * whose formula was rewritten for the row it moved to. A chain of changed
 * keys cannot say this (a sort changes nearly every key), while a value
 * derived from the map's cells can follow it from the map before.
 */
export interface CellMapRowPermutation {
  readonly dataR1: number;
  readonly c1: number;
  readonly c2: number;
  readonly sourceOrder: readonly number[];
  readonly relocated: readonly string[];
}

const permutations = new WeakMap<CellMapLike, { readonly before: WeakRef<CellMapLike>; readonly permutation: CellMapRowPermutation }>();

/** A row sort replaced `before` with `after`, as `permutation` says. An order
 * that is not exactly the rows of its span records nothing. */
export function recordCellMapRowPermutation(before: CellMapLike, after: CellMapLike, permutation: CellMapRowPermutation): void {
  if (before === after || permutations.has(after) || typeof WeakRef === "undefined") return;
  const { dataR1, sourceOrder } = permutation;
  const seen = new Uint8Array(sourceOrder.length);
  for (const row of sourceOrder) {
    const offset = row - dataR1;
    if (!Number.isInteger(offset) || offset < 0 || offset >= seen.length || seen[offset]) return;
    seen[offset] = 1;
  }
  permutations.set(after, { before: new WeakRef(before), permutation });
}

/** The row sort that made `after`, and the map it was made from while that is held. */
export function cellMapRowPermutation<M extends CellMapLike>(after: M): { before: M; permutation: CellMapRowPermutation } | null {
  const record = permutations.get(after);
  const before = record?.before.deref();
  return record && before ? { before: before as M, permutation: record.permutation } : null;
}

/** The key a permuted map's cell at `row`, `col` was read from. */
export function cellMapPermutedSourceKey(permutation: CellMapRowPermutation, row: number, col: number): string {
  return `${permutation.sourceOrder[row - permutation.dataR1]}:${col}`;
}

/** A row and column coded as one number: columns stay under this. */
const COL_CODE = 2 ** 20;

/** A row sort's move, read for what is derived from the cells: where each row
 * went, and the cells rewritten by it, which alone can answer differently. */
export interface PermutedSpan {
  readonly permutation: CellMapRowPermutation;
  readonly r2: number;
  /** destination[row - dataR1] is the row that row moved to. */
  readonly destination: Int32Array;
  /** Coded row and column of each rewritten cell, at its new place. */
  readonly relocated: ReadonlySet<number>;
  readonly relocatedAt: readonly (readonly [number, number])[];
}

const permutedSpans = new WeakMap<CellMapRowPermutation, PermutedSpan>();

export function permutedSpanOf(permutation: CellMapRowPermutation): PermutedSpan {
  const known = permutedSpans.get(permutation);
  if (known) return known;
  const { dataR1, sourceOrder } = permutation;
  const destination = new Int32Array(sourceOrder.length);
  sourceOrder.forEach((source, offset) => { destination[source - dataR1] = dataR1 + offset; });
  const relocatedAt: (readonly [number, number])[] = [];
  const relocated = new Set<number>();
  for (const key of permutation.relocated) {
    const at = key.indexOf(":");
    const row = Number(key.slice(0, at)), col = Number(key.slice(at + 1));
    relocatedAt.push([row, col]);
    relocated.add(row * COL_CODE + col);
  }
  const span: PermutedSpan = { permutation, r2: dataR1 + sourceOrder.length - 1, destination, relocated, relocatedAt };
  permutedSpans.set(permutation, span);
  return span;
}

/** A key list of the map after a row sort, in that map's own key order: the
 * keys outside the span as they stood, then the span's by new row and column.
 * The span's rewritten cells are asked again. */
export function permutedCellKeys<C>(
  keys: readonly string[],
  span: PermutedSpan,
  after: Readonly<Record<string, C>>,
  candidate: (cell: C) => boolean,
): string[] {
  const { permutation: { dataR1, c1, c2 }, r2, destination, relocated, relocatedAt } = span;
  const outside: string[] = [];
  const columns: (number[] | undefined)[] = new Array(destination.length);
  for (const key of keys) {
    const at = key.indexOf(":");
    const row = Number(key.slice(0, at)), col = Number(key.slice(at + 1));
    if (!(at > 0 && row >= dataR1 && row <= r2 && col >= c1 && col <= c2) || key !== `${row}:${col}`) {
      outside.push(key);
      continue;
    }
    const moved = destination[row - dataR1];
    if (relocated.has(moved * COL_CODE + col)) continue;
    (columns[moved - dataR1] ??= []).push(col);
  }
  for (const [row, col] of relocatedAt) {
    if (candidate(after[`${row}:${col}`])) (columns[row - dataR1] ??= []).push(col);
  }
  columns.forEach((cols, offset) => {
    if (!cols) return;
    for (let index = 1; index < cols.length; index += 1) {
      if (cols[index - 1] > cols[index]) { cols.sort((a, b) => a - b); break; }
    }
    const prefix = `${dataR1 + offset}:`;
    for (const col of cols) outside.push(prefix + col);
  });
  return outside;
}

/** The keys that changed from `before` to `after`, when the chain says. */
export function cellMapChangesSince(after: CellMapLike, before: CellMapLike): readonly string[] | null {
  return nearestKnownCellMap(after, (cells) => (cells === before ? true : undefined))?.keys ?? null;
}

/** The nearest recorded ancestor of `after` that `known` holds, with the keys changed since. */
export function nearestKnownCellMap<T, M extends CellMapLike = CellMapLike>(after: M, known: (cells: M) => T | undefined): { value: T; keys: readonly string[] } | null {
  const keys = new Set<string>();
  let cursor: CellMapLike = after;
  for (let step = 0; step < MAX_CHAIN; step += 1) {
    const parent = parents.get(cursor);
    if (!parent) return null;
    for (const key of parent.keys) keys.add(key);
    if (keys.size > MAX_KEYS) return null;
    const previous = parent.before.deref();
    if (!previous) return null;
    const value = known(previous as M);
    if (value !== undefined) return { value, keys: [...keys] };
    cursor = previous;
  }
  return null;
}
