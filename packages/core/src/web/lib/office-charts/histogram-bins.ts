/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A single histogram bucket: its two edges and the values that fell inside. */
export interface HistogramBin {
  /** Lower edge. Inclusive unless the declared closed side is `r`. */
  lo: number;
  /** Upper edge. Inclusive when the declared closed side is `r`. */
  hi: number;
  /** Number of finite values that fell into this bucket. */
  count: number;
  /** The underflow / overflow bin (its cutoff is `hi` / `lo`). */
  tail?: "underflow" | "overflow";
}

/** `ST_IntervalClosedSide` ([MS-ODRAWXML] `CT_Binning`): the two legal words. */
export type HistogramIntervalClosed = "l" | "r";

/**
 * Excel's own default, and what Office writes on every automatic histogram or
 * pareto part in the corpora (31 of 31, Excel, Word and PowerPoint alike).
 */
export const DEFAULT_HISTOGRAM_INTERVAL_CLOSED: HistogramIntervalClosed = "r";

/** Hard cap on histogram bucket count (DoS bound on iteration / DOM marks). */
const MAX_HISTOGRAM_BINS = 200;

/**
 * The quotient `(v - lower) / width`, snapped when it is an integer in decimal.
 *
 * The server states this same rule in exact decimal arithmetic
 * (`office_chart_binning.py`), so a value sitting exactly ON an edge has to
 * resolve here to the same whole number it resolves to there; a binary double
 * can miss that by an ulp and throw the value one bucket over.
 */
function binQuotient(value: number, lower: number, width: number): number {
  const quotient = (value - lower) / width;
  const whole = Math.round(quotient);
  return Math.abs(quotient - whole) < 1e-9 ? whole : quotient;
}

/**
 * Office's automatic bin width: Scott's normal reference rule, cut to two
 * significant digits BEFORE any value is assigned.
 *
 * Support "Create a histogram" names Automatic bins as "calculated using
 * Scott's normal reference rule" (`3.5 * s / n^(1/3)`), and Excel's automatic
 * edges really are round numbers - rounding the axis LABELS afterwards instead
 * would move a value across an edge. Returns null when there is nothing to
 * scale (fewer than two values, or a zero-width domain).
 */
function scottBinWidth(values: ReadonlyArray<number>): number | null {
  if (values.length < 2) return null;
  const mean = values.reduce((sum, v) => sum + v, 0) / values.length;
  const variance =
    values.reduce((sum, v) => sum + (v - mean) * (v - mean), 0) / (values.length - 1);
  // `Math.pow(n, 1/3)`, not `Math.cbrt`: the server writes `n ** (1 / 3)`, and
  // the two disagree in the last bit for some n.
  const raw = (3.5 * Math.sqrt(variance)) / Math.pow(values.length, 1 / 3);
  if (!Number.isFinite(raw) || raw <= 0) return null;
  const width = Number(raw.toPrecision(2));
  return Number.isFinite(width) && width > 0 ? width : null;
}

/**
 * Bin a numeric series into equal-width buckets and count the values in each.
 *
 * This is the browser's statement of the ONE binning rule; the server states
 * the same rule in `apps/api/app/services/office_chart_binning.py`, and
 * `apps/web/lib/office-charts/histogram-binning.shared.json` pins the two
 * against each other on datasets whose values sit exactly on bucket edges.
 *
 *  - `intervalClosed` is the side the chart DECLARES (`cx:binning/@intervalClosed`),
 *    not a house style: `r` (Excel's default) means `(a, b]`, so a value on an
 *    edge joins the bucket below it and the FIRST bucket also holds the
 *    minimum; `l` means `[a, b)`, so an edge value joins the bucket above it
 *    and the LAST bucket also holds the maximum. Both sides bin the same edges.
 *  - `binCount` is Format Axis ▸ "Number of bins", clamped to
 *    [1, MAX_HISTOGRAM_BINS]. Omitted, the width is Scott's rule and the count
 *    is `ceil((max - min) / width)` - Excel's Automatic.
 *  - A zero-width domain (every value equal, or a single value) collapses to
 *    ONE bucket holding every value (no divide-by-zero).
 *
 * Non-finite values are dropped first so a blank/non-numeric cell never NaNs an
 * edge; [] comes back when no finite value is left.
 */
export function computeHistogramBins(
  values: ReadonlyArray<number>,
  options?: {
    binCount?: number;
    binWidth?: number;
    intervalClosed?: HistogramIntervalClosed;
    overflow?: number | "auto";
    underflow?: number | "auto";
  },
): HistogramBin[] {
  const intervalClosed = options?.intervalClosed ?? DEFAULT_HISTOGRAM_INTERVAL_CLOSED;
  const binCount = options?.binCount;
  const all = values.filter((v) => Number.isFinite(v));
  if (all.length === 0) return [];
  // Format Axis ▸ Bins ▸ Overflow / Underflow bin: every value past the
  // cutoff is one bucket of its own ("> 14", "≤ 11"), and the regular buckets
  // span the cutoffs (Microsoft Support, "Create a histogram"). `"auto"`
  // states the bin with no cutoff of its own: nothing is set apart.
  const overflow = typeof options?.overflow === "number" ? options.overflow : undefined;
  const underflow = typeof options?.underflow === "number" ? options.underflow : undefined;
  const beyond = (v: number) =>
    overflow !== undefined && (intervalClosed === "r" ? v > overflow : v >= overflow);
  const below = (v: number) =>
    underflow !== undefined && (intervalClosed === "r" ? v <= underflow : v < underflow);
  const finite = all.filter((v) => !beyond(v) && !below(v));
  const under = all.length - finite.length - all.filter(beyond).length;
  const over = all.filter(beyond).length;
  const withTails = (bins: HistogramBin[]): HistogramBin[] => [
    ...(underflow !== undefined
      ? [{ lo: underflow, hi: underflow, count: under, tail: "underflow" as const }]
      : []),
    ...bins,
    ...(overflow !== undefined
      ? [{ lo: overflow, hi: overflow, count: over, tail: "overflow" as const }]
      : []),
  ];
  if (finite.length === 0) return withTails([]);
  let min = underflow ?? finite[0];
  let max = overflow ?? finite[0];
  for (const v of finite) {
    if (underflow === undefined && v < min) min = v;
    if (overflow === undefined && v > max) max = v;
  }
  // Zero-width domain (all values equal / single value): one bucket holding all.
  if (max <= min) {
    return withTails([{ lo: min, hi: min, count: finite.length }]);
  }
  let bins: number;
  let width: number;
  if (binCount !== undefined && Number.isFinite(binCount)) {
    bins = Math.trunc(binCount);
    if (bins < 1) bins = 1;
    if (bins > MAX_HISTOGRAM_BINS) bins = MAX_HISTOGRAM_BINS;
    width = (max - min) / bins;
  } else if (options?.binWidth !== undefined && Number.isFinite(options.binWidth) && options.binWidth > 0) {
    // Format Axis ▸ Bins ▸ Bin width: buckets of that width from the lower
    // edge until the upper one is covered.
    width = options.binWidth;
    bins = Math.max(1, Math.ceil(binQuotient(max, min, width)));
    if (bins > MAX_HISTOGRAM_BINS) {
      bins = MAX_HISTOGRAM_BINS;
      width = (max - min) / bins;
    }
  } else {
    const scott = scottBinWidth(finite);
    if (scott === null) return [{ lo: min, hi: max, count: finite.length }];
    width = scott;
    bins = Math.ceil(binQuotient(max, min, width));
    if (bins < 1) bins = 1;
    if (bins > MAX_HISTOGRAM_BINS) {
      bins = MAX_HISTOGRAM_BINS;
      width = (max - min) / bins;
    }
  }
  const counts = new Array<number>(bins).fill(0);
  for (const v of finite) {
    const quotient = binQuotient(v, min, width);
    let idx = intervalClosed === "r" ? Math.ceil(quotient) - 1 : Math.floor(quotient);
    if (idx < 0) idx = 0;
    if (idx >= bins) idx = bins - 1;
    counts[idx] += 1;
  }
  const out: HistogramBin[] = [];
  for (let i = 0; i < bins; i += 1) {
    out.push({ lo: min + i * width, hi: min + (i + 1) * width, count: counts[i] });
  }
  return withTails(out);
}

/**
 * A bucket's axis label as Excel writes it: `[lo, hi]` for the first bucket
 * and `(lo, hi]` after it on a right-closed histogram (`[lo, hi)` / `[lo, hi]`
 * left-closed), `≤ cutoff` / `> cutoff` for the underflow / overflow bin.
 */
export function histogramBinLabel(
  bins: ReadonlyArray<HistogramBin>,
  index: number,
  intervalClosed: HistogramIntervalClosed = DEFAULT_HISTOGRAM_INTERVAL_CLOSED,
  digits = 4,
): string {
  const bin = bins[index];
  const text = (value: number) => String(Number(value.toPrecision(digits)));
  if (bin.tail === "underflow") return `${intervalClosed === "r" ? "≤" : "<"}${text(bin.hi)}`;
  if (bin.tail === "overflow") return `${intervalClosed === "r" ? ">" : "≥"}${text(bin.lo)}`;
  const regular = bins.filter((entry) => !entry.tail);
  const place = regular.indexOf(bin);
  // The first bucket also holds the minimum (right-closed), the last the
  // maximum (left-closed) -- unless a tail bin holds what lies past it.
  if (intervalClosed === "r") {
    const closedLow = place === 0 && !bins.some((entry) => entry.tail === "underflow");
    return `${closedLow ? "[" : "("}${text(bin.lo)}, ${text(bin.hi)}]`;
  }
  const closedHigh = place === regular.length - 1 && !bins.some((entry) => entry.tail === "overflow");
  return `[${text(bin.lo)}, ${text(bin.hi)}${closedHigh ? "]" : ")"}`;
}
