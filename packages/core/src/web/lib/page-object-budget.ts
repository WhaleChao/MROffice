/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a list of projected objects costs its page, and whether it has room.
 *
 * One rule, one measurement, for every host that puts a package's objects on
 * a page. The workbook learned it first — see
 * `docs/spreadsheet-page-limits.md`, "Drawing objects: a share of the page,
 * not a count": a drawing object used to be capped at a COUNT of 256, which
 * cannot be right for two objects of such different weight (a sheet shape is
 * ~430 bytes of page JSON, a preserved SmartArt diagram ~30 KB). The bound
 * that binds is the bytes; a count survives only as a bound on the importer's
 * own work.
 *
 * A deck's slide had the identical cap in the identical shape
 * (200 blocks per slide), so rather than write the rule a second time both
 * hosts bind their own two numbers to this module. The server twins are
 * `app/services/resources/object_budget.py` and the constants beside each
 * host's projection.
 *
 * The budget is asked BEFORE an object is projected — "spent", never
 * "exceeded" — so a page never holds an object it could not afford, which is
 * what lets the whole-page envelope keep the last word on save.
 */

/** One host's share, in the two currencies a page is measured in. */
export interface PageObjectBudgetLimits {
  /** A bound on the importer's own work, not the bound that binds. */
  readonly maxObjects: number;
  /** Bytes of object JSON the page will hold. */
  readonly byteBudget: number;
}

const encoder = new TextEncoder();
const measured = new WeakMap<
  readonly unknown[],
  { length: number; bytes: number }
>();

function objectBytes(value: unknown, unserializableCost: number): number {
  try {
    return encoder.encode(JSON.stringify(value) ?? "").length;
  } catch {
    // Unserializable: charge it the whole budget so it cannot slip through.
    return unserializableCost;
  }
}

/**
 * Bytes the list's objects put on the page, measured once per object.
 *
 * Memoized on the list itself: the callers ask this once per candidate
 * object while they build the list, so measuring the whole list each time
 * would be quadratic. A list that shrank is re-measured from the start.
 */
export function pageObjectBytes(
  objects: readonly unknown[],
  limits: PageObjectBudgetLimits,
): number {
  const state = measured.get(objects) ?? { length: 0, bytes: 0 };
  if (state.length > objects.length) {
    state.length = 0;
    state.bytes = 0;
  }
  for (let index = state.length; index < objects.length; index += 1) {
    state.bytes += objectBytes(objects[index], limits.byteBudget);
  }
  state.length = objects.length;
  measured.set(objects, state);
  return state.bytes;
}

/** Whether the list has used its share: no further object is projected. */
export function pageObjectBudgetSpent(
  objects: readonly unknown[],
  limits: PageObjectBudgetLimits,
): boolean {
  return (
    objects.length >= limits.maxObjects ||
    pageObjectBytes(objects, limits) >= limits.byteBudget
  );
}

/** The longest prefix of `objects` that fits the budget. */
export function pageObjectsWithinBudget<T>(
  objects: readonly T[],
  limits: PageObjectBudgetLimits,
): T[] {
  const kept: T[] = [];
  for (const object of objects) {
    if (pageObjectBudgetSpent(kept, limits)) break;
    kept.push(object);
  }
  return kept;
}
