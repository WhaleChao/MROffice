/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Transaction } from "@tiptap/pm/state";
import { DecorationSet } from "@tiptap/pm/view";

export type PositionRange = [number, number];

interface TransactionSpan {
  /** Original positions below this one are unchanged. */
  floor: number;
  /** Original positions above this one move by `delta`. */
  ceiling: number;
  delta: number;
}

const spans = new WeakMap<Transaction, TransactionSpan>();

/** Where a transaction's steps change the document, in original positions. */
export function transactionSpan(tr: Transaction): TransactionSpan {
  const known = spans.get(tr);
  if (known) return known;
  const maps = tr.mapping.maps;
  let floor = Number.POSITIVE_INFINITY;
  let ceilingAfter = Number.NEGATIVE_INFINITY;
  for (let index = 0; index < maps.length; index += 1) {
    const after = maps.length > index + 1 ? tr.mapping.slice(index + 1) : null;
    maps[index]!.forEach((oldStart, _oldEnd, _newStart, newEnd) => {
      // Each step's own positions: a position below every step's first
      // change is below it in every step's document too.
      if (oldStart < floor) floor = oldStart;
      const end = after ? after.map(newEnd, 1) : newEnd;
      if (end > ceilingAfter) ceilingAfter = end;
    });
  }
  const delta = tr.doc.content.size - tr.before.content.size;
  const span: TransactionSpan = Number.isFinite(floor)
    ? { floor, ceiling: tr.mapping.invert().map(ceilingAfter, 1), delta }
    // Nothing moves (attributes or marks only).
    : { floor: Number.POSITIVE_INFINITY, ceiling: Number.POSITIVE_INFINITY, delta: 0 };
  spans.set(tr, span);
  return span;
}

/** A position mapped through `tr` as `tr.mapping.map(pos, assoc)` maps it. */
export function mapPositionFast(tr: Transaction, pos: number, assoc: -1 | 1): number {
  const { floor, ceiling, delta } = transactionSpan(tr);
  if (pos < floor) return pos;
  if (pos > ceiling) return pos + delta;
  return tr.mapping.map(pos, assoc);
}

/** Past this many ranges a list is kept merged. */
const MERGE_ABOVE = 256;

/** Overlapping or touching ranges as one: the same positions. */
export function mergeRanges(ranges: readonly PositionRange[]): PositionRange[] {
  const sorted = [...ranges].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const merged: PositionRange[] = [];
  for (const [from, to] of sorted) {
    const last = merged[merged.length - 1];
    if (last && from <= last[1]) {
      if (to > last[1]) last[1] = to;
    } else merged.push([from, to]);
  }
  return merged;
}

/**
 * `ranges` mapped through `tr` (each start with assoc -1, each end with +1),
 * then `added`. With nothing of `ranges` at or past the transaction's first
 * change, the same entries come back without a mapping.
 */
export function mapRangesFast(
  ranges: readonly PositionRange[],
  tr: Transaction,
  added: readonly PositionRange[] = [],
  /** False for a list read by each range's start (a retry list): never merged. */
  merge = true,
): PositionRange[] {
  const { floor } = transactionSpan(tr);
  let mapped: PositionRange[];
  let allBelow = true;
  for (const range of ranges) if (range[1] >= floor) { allBelow = false; break; }
  if (allBelow) mapped = ranges.length || added.length ? [...ranges, ...added] : [];
  else {
    mapped = ranges.map(([from, to]): PositionRange => [mapPositionFast(tr, from, -1), mapPositionFast(tr, to, 1)]);
    if (added.length) mapped.push(...added);
  }
  return merge && mapped.length > MERGE_ABOVE ? mergeRanges(mapped) : mapped;
}

/** `window.__docsFastRangeMap = false` maps every range step by step (old, A/B). */
export function fastRangeMappingOn(): boolean {
  return (globalThis as { __docsFastRangeMap?: boolean }).__docsFastRangeMap !== false;
}

/**
 * A decoration set mapped through `tr`; the same set when the transaction
 * changes nothing in or after the top-level node where its first change is
 * and no decoration stands there (a part appended past every decoration).
 */
export function mapDecorationsFast(set: DecorationSet, tr: Transaction): DecorationSet {
  if (set === DecorationSet.empty) return set;
  const { floor } = transactionSpan(tr);
  const size = tr.before.content.size;
  if (floor > size) return Number.isFinite(floor) ? set.map(tr.mapping, tr.doc) : set;
  const $floor = tr.before.resolve(floor);
  const nodeStart = $floor.depth > 0 ? $floor.before(1) : floor;
  if (set.find(nodeStart, size).length === 0) return set;
  return set.map(tr.mapping, tr.doc);
}
