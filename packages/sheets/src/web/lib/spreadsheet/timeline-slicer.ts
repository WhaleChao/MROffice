/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Timeline slicer (Excel "Insert ▸ Timeline" parity): a date-aware slicer that
 * filters a column to a CONTIGUOUS span of periods at a chosen granularity
 * (year / quarter / month / day). Unlike a regular slicer (arbitrary value
 * multi-select), a timeline selects a START→END range and every date in between
 * passes.
 *
 * Pure + dependency-free: it reuses {@link parseFilterDate} (the same date parser
 * the AutoFilter date tree uses) to bucket a column's display values into
 * ordered, comparable periods, and turns a selected [start,end] period range
 * into the exact set of display values that pass — which the host applies through
 * the SAME per-column filter mechanism a slicer uses. Exhaustively unit-testable.
 */

import { parseFilterDate, type FilterDate } from "@mrnote/core/web/lib/spreadsheet/date-filter-tree.ts";

export type TimelineGranularity = "year" | "quarter" | "month" | "day";

export const TIMELINE_GRANULARITIES: readonly TimelineGranularity[] = [
  "year",
  "quarter",
  "month",
  "day",
];

export interface TimelinePeriod {
   
  key: string;
  /** Human label, e.g. "2024", "Q2 2024", "May 2024", "May 14, 2024". */
  label: string;
  /** Monotonic integer for contiguous-range membership (later date ⇒ larger). */
  ordinal: number;
  /** The original column display strings that fall in this period. */
  values: string[];
}

const MONTH_ABBR = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function quarterOf(month: number): number {
  return Math.floor((month - 1) / 3) + 1;
}

/**
 * A monotonic integer for a date at the given granularity — later dates always
 * map to larger ordinals, so `start ≤ ordinal ≤ end` is exactly "in the span".
 */
export function periodOrdinal(d: FilterDate, granularity: TimelineGranularity): number {
  switch (granularity) {
    case "year":
      return d.year;
    case "quarter":
      return d.year * 4 + (quarterOf(d.month) - 1);
    case "month":
      return d.year * 12 + (d.month - 1);
    case "day":
      // YYYYMMDD is strictly monotonic with calendar order (no gap-free
      // requirement — only comparison is needed for range membership).
      return d.year * 10_000 + d.month * 100 + d.day;
  }
}

/** The stable key + display label for a date at the given granularity. */
export function periodKeyLabel(
  d: FilterDate,
  granularity: TimelineGranularity,
): { key: string; label: string } {
  switch (granularity) {
    case "year":
      return { key: `${d.year}`, label: `${d.year}` };
    case "quarter": {
      const q = quarterOf(d.month);
      return { key: `${d.year}-Q${q}`, label: `Q${q} ${d.year}` };
    }
    case "month":
      return {
        key: `${d.year}-${pad2(d.month)}`,
        label: `${MONTH_ABBR[d.month - 1]} ${d.year}`,
      };
    case "day":
      return {
        key: `${d.year}-${pad2(d.month)}-${pad2(d.day)}`,
        label: `${MONTH_ABBR[d.month - 1]} ${d.day}, ${d.year}`,
      };
  }
}

/**
 * Bucket a column's distinct display values into ordered, distinct periods at the
 * given granularity. Non-date / blank values are ignored. The result is sorted
 * ascending by ordinal so the UI can render a left-to-right timeline and the
 * caller can pick a contiguous [start,end] sub-range.
 */
export function buildTimeline(
  values: readonly string[],
  granularity: TimelineGranularity,
): TimelinePeriod[] {
  const byKey = new Map<string, TimelinePeriod>();
  for (const raw of values) {
    const date = parseFilterDate(raw);
    if (!date) continue;
    const { key, label } = periodKeyLabel(date, granularity);
    const existing = byKey.get(key);
    if (existing) {
      if (!existing.values.includes(raw)) existing.values.push(raw);
    } else {
      byKey.set(key, {
        key,
        label,
        ordinal: periodOrdinal(date, granularity),
        values: [raw],
      });
    }
  }
  return [...byKey.values()].sort((a, b) => a.ordinal - b.ordinal);
}

/**
 * The display values that pass a contiguous timeline selection — every value
 * whose period ordinal is within [min(a,b), max(a,b)] (endpoints inclusive,
 * order-agnostic). Returns a flat, de-duplicated list suitable for the column
 * filter's `selected` set. An empty/whole selection is the caller's concern;
 * this returns exactly the in-range values.
 */
export function timelineRangeValues(
  periods: readonly TimelinePeriod[],
  ordinalA: number,
  ordinalB: number,
): string[] {
  const lo = Math.min(ordinalA, ordinalB);
  const hi = Math.max(ordinalA, ordinalB);
  const out: string[] = [];
  for (const period of periods) {
    if (period.ordinal < lo || period.ordinal > hi) continue;
    for (const value of period.values) {
      if (!out.includes(value)) out.push(value);
    }
  }
  return out;
}

/**
 * Derive the currently-selected [from,to] period span FROM the persisted column
 * filter (the source of truth), mirroring how a slicer derives its selection.
 * A null/empty filter means "no constraint" ⇒ the whole span. Otherwise the span
 * is the min/max ordinal among periods whose values intersect the filter. Returns
 * null when there are no periods (nothing to select).
 */
export function deriveTimelineSpan(
  periods: readonly TimelinePeriod[],
  filter: readonly string[] | null | undefined,
): { fromOrdinal: number; toOrdinal: number } | null {
  if (periods.length === 0) return null;
  const fullSpan = {
    fromOrdinal: periods[0].ordinal,
    toOrdinal: periods[periods.length - 1].ordinal,
  };
  if (!filter || filter.length === 0) return fullSpan;
  const allowed = new Set(filter.map(String));
  let lo = Infinity;
  let hi = -Infinity;
  for (const period of periods) {
    if (!period.values.some((v) => allowed.has(v))) continue;
    if (period.ordinal < lo) lo = period.ordinal;
    if (period.ordinal > hi) hi = period.ordinal;
  }
  // No period intersects the filter (e.g. it only allows non-date values) → the
  // timeline shows nothing selected; fall back to the full span for the control.
  if (lo === Infinity) return fullSpan;
  return { fromOrdinal: lo, toOrdinal: hi };
}
