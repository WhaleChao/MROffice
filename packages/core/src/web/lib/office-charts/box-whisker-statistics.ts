/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A box plot's statistics, shared by every editor that draws one (the sheet's
 * worksheet overlay and the Docs / deck planner): the five numbers, the mean,
 * the points inside and outside the 1.5 × IQR fences, by the box's Quartile
 * Calculation (Inclusive or Exclusive median).
 */

/**
 * The five-number summary + outliers + mean for a single box-plot series,
 * computed via Excel's QUARTILE.INC method (inclusive linear interpolation).
 *
 * Excel's QUARTILE.INC formula:
 *   q(p, sorted) = sorted[floor(rank)] + frac * (sorted[ceil(rank)] - sorted[floor(rank)])
 *   where rank = p * (n - 1)  (0-indexed, 0 ≤ p ≤ 1)
 *
 * Whiskers extend to the furthest point within [Q1 - 1.5·IQR, Q3 + 1.5·IQR].
 * Points beyond that fence are outliers. Mean is the arithmetic mean of all
 * finite values (Excel shows the mean as an × marker).
 */
export interface BoxWhiskerStats {
  min: number;       // smallest value within the lower fence (whisker low)
  q1: number;        // first quartile (QUARTILE.INC, p=0.25)
  median: number;    // second quartile / median (p=0.5)
  q3: number;        // third quartile (p=0.75)
  max: number;       // largest value within the upper fence (whisker high)
  mean: number;      // arithmetic mean of finite values (Excel × marker)
  outliers: number[]; // values outside [Q1 - 1.5·IQR, Q3 + 1.5·IQR]
  inner: number[];    // the values within the fences (Show inner points)
}

/**
 * Compute QUARTILE.INC at quantile p ∈ [0, 1] over a pre-sorted numeric array.
 * Uses Excel's inclusive linear interpolation:
 *   rank = p * (n − 1); result = sorted[⌊rank⌋] + frac * (sorted[⌈rank⌉] − sorted[⌊rank⌋])
 * Requires n >= 1; caller ensures the array is sorted ascending and non-empty.
 */
export function quartileInc(sorted: ReadonlyArray<number>, p: number): number {
  const n = sorted.length;
  if (n === 0) return 0;
  if (n === 1) return sorted[0];
  const rank = p * (n - 1);
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  const frac = rank - lo;
  return sorted[lo] + frac * (sorted[hi] - sorted[lo]);
}








export function quartileExc(sorted: ReadonlyArray<number>, p: number): number {
  const n = sorted.length;
  if (n === 0) return 0;
  const rank = Math.min(Math.max(p * (n + 1), 1), n) - 1;
  const lo = Math.floor(rank);
  const hi = Math.ceil(rank);
  return sorted[lo] + (rank - lo) * (sorted[hi] - sorted[lo]);
}

/**
 * Compute the box-and-whisker five-number summary + outliers + mean for a
 * numeric series (non-finite values are dropped first). Returns null when fewer
 * than 1 finite value exists (nothing to plot). The quartiles are the box's
 * Quartile Calculation: Inclusive median (QUARTILE.INC, the default here) or
 * Exclusive median (QUARTILE.EXC); 1.5 × IQR fences mark the outliers.
 */
export function computeBoxWhiskerStats(
  values: ReadonlyArray<number>,
  quartileMethod: "inclusive" | "exclusive" = "inclusive",
): BoxWhiskerStats | null {
  const finite = values.filter((v) => Number.isFinite(v));
  if (finite.length === 0) return null;
  const sorted = [...finite].sort((a, b) => a - b);
  const n = sorted.length;
  const mean = finite.reduce((s, v) => s + v, 0) / n;
  const quartile = quartileMethod === "exclusive" ? quartileExc : quartileInc;
  const q1 = quartile(sorted, 0.25);
  const median = quartile(sorted, 0.5);
  const q3 = quartile(sorted, 0.75);
  const iqr = q3 - q1;
  const fenceLo = q1 - 1.5 * iqr;
  const fenceHi = q3 + 1.5 * iqr;
  const outliers: number[] = [];
  const inner: number[] = [];
  let whiskerLo = q1; // will be updated to the smallest non-outlier value
  let whiskerHi = q3; // will be updated to the largest non-outlier value
  // Initialize whiskerLo/Hi to the fence values; then clamp to the actual data.
  whiskerLo = sorted[0]; // start at the global min; will be bumped if it's an outlier
  whiskerHi = sorted[n - 1];
  // Walk through sorted values; separate outliers from the inner-fence set.
  // whiskerLo = the smallest non-outlier; whiskerHi = the largest non-outlier.
  let innerLo = Infinity;
  let innerHi = -Infinity;
  for (const v of sorted) {
    if (v < fenceLo || v > fenceHi) {
      outliers.push(v);
    } else {
      inner.push(v);
      if (v < innerLo) innerLo = v;
      if (v > innerHi) innerHi = v;
    }
  }
  // If every value is an outlier (extreme IQR=0 edge case), fall back to q1/q3.
  whiskerLo = Number.isFinite(innerLo) ? innerLo : q1;
  whiskerHi = Number.isFinite(innerHi) ? innerHi : q3;
  return { min: whiskerLo, q1, median, q3, max: whiskerHi, mean, outliers, inner };
}














export type BoxWhiskerCluster = Readonly<{
  categories: string[];
  slots: number[];
  boxWidth: string;
  /** Ticks and gridlines stand on the one category's edges only. */
  edgeOnly: (index: number) => boolean;
}>;

/** Slots between two neighbouring boxes' centres (even: a cluster of an even
 * count is centred between two boxes, on a whole slot). */
const BOX_WHISKER_SLOTS_PER_BOX = 20;
/** Office parts neighbouring boxes by a tenth of a box. */
const BOX_WHISKER_SERIES_GAP = 0.1;

export function boxWhiskerCluster(seriesCount: number, gapWidthPercent = 100, label = "1"): BoxWhiskerCluster {
  const count = Math.max(1, Math.floor(seriesCount));
  const gap = Math.min(500, Math.max(0, gapWidthPercent)) / 100;
  const pitch = 1 + BOX_WHISKER_SERIES_GAP;
  const cluster = count + BOX_WHISKER_SERIES_GAP * (count - 1);
  const slotWidth = pitch / BOX_WHISKER_SLOTS_PER_BOX;
  // An odd count of slots, so the category's middle is a slot's centre.
  const total = Math.max(1, 2 * Math.round((cluster * (1 + gap) / slotWidth - 1) / 2) + 1);
  const middle = (total - 1) / 2;
  const slots = Array.from({ length: count }, (_unused, index) =>
    Math.round(middle + (index - (count - 1) / 2) * BOX_WHISKER_SLOTS_PER_BOX),
  );
  const categories = Array.from({ length: total }, (_unused, index) => (index === middle ? label : ""));
  return {
    categories,
    slots,
    boxWidth: `${((100 * BOX_WHISKER_SLOTS_PER_BOX) / pitch).toFixed(3)}%`,
    edgeOnly: (index: number) => index === 0,
  };
}
