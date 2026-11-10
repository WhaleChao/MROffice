/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for in-cell SPARKLINES (Excel / WPS parity).
 *
 * A sparkline is a tiny chart drawn INSIDE a single cell — there is no separate
 * chart object on the sheet. It is bound to a rectangular data range and drawn
 * as a normalized line, a column (mini bar) chart, or a win/loss bar chart. The
 * sheet model stores one sparkline per cell in a `sparklines` map keyed by the
 * cell's `"row:col"` ref (the same key shape as `cells`), so a cell either has a
 * sparkline or it does not.
 *
 * Display decision (documented for the renderer): a sparkline is the cell's
 * BACKGROUND. Excel: "Because a sparkline is embedded in a cell, any text you
 * enter in the cell uses the sparkline as its background", so the cell's own
 * value is drawn over the chart exactly as it would be without one.
 *
 * This module is dependency-free and deterministic so it can be unit-tested in
 * isolation: it carries the sparkline MODEL (type + range + colors), the
 * normalize/persist coercion (a persisted value can come from an imported file,
 * an AI write, or a hand-edited JSON, so every field is clamped/dropped rather
 * than trusted), and the pure GEOMETRY (`sparklinePoints`) that turns an array
 * of numbers into the polyline points / bar rectangles the SVG draws. The React
 * layer only reads the cell values for the range, calls these helpers, and
 * renders the result — it owns no sparkline math.
 */

/** The three sparkline kinds Excel draws. */
import {
  normalizeThemeColor,
  type SpreadsheetThemeColor,
} from "./theme-colors";

export type SparklineType = "line" | "column" | "winloss";

export const SPARKLINE_TYPES: readonly SparklineType[] = [
  "line",
  "column",
  "winloss",
] as const;

/** Native x14 blank-cell plotting modes. */
export type SparklineEmptyCellDisplay = "zero" | "gap" | "span";

/** Native x14 vertical-axis scaling modes. */
export type SparklineAxisType = "individual" | "group" | "custom";

/**
 * A per-cell sparkline. `range` is an A1-style rectangle string ("A1:A12") that
 * feeds the chart; the cell values inside it are read by the host and passed to
 * `sparklinePoints`. `color` is the primary stroke/fill (#rrggbb), `negativeColor`
 * the fill used for negative columns / loss bars. Both colors are optional; the
 * renderer falls back to a default when absent.
 */
export const SPARKLINE_COLOR_FIELDS = [
  "color",
  "negativeColor",
  "axisColor",
  "markerColor",
  "firstColor",
  "lastColor",
  "highColor",
  "lowColor",
] as const;
export type SparklineColorField = (typeof SPARKLINE_COLOR_FIELDS)[number];

export interface Sparkline {
  type: SparklineType;
  /**
   * A1 data source, or "" for a sparkline without one: Excel keeps the member
   * after its source cells were deleted and draws nothing for it.
   */
  range: string;
  /** Stable editor identity for settings shared by one native sparklineGroup. */
  groupId?: string;
  /**
   * Theme-slot identity of the colours below. Excel writes its default
   * sparkline colours as theme references and keeps following the theme; a
   * reference only stays valid while its `rgb` equals the carried colour, so
   * a re-picked colour is exported as that rgb.
   */
  themeColors?: Partial<Record<SparklineColorField, SpreadsheetThemeColor>>;
  color?: string;
  negativeColor?: string;
  axisColor?: string;
  markerColor?: string;
  firstColor?: string;
  lastColor?: string;
  highColor?: string;
  lowColor?: string;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  displayXAxis?: boolean;
  displayHidden?: boolean;
  displayEmptyCellsAs?: SparklineEmptyCellDisplay;
  minAxisType?: SparklineAxisType;
  maxAxisType?: SparklineAxisType;
  manualMin?: number;
  manualMax?: number;
  /**
   * Native line weight in points (Excel Web exposes this for line sparklines);
   * absent means the x14 default, DEFAULT_SPARKLINE_LINE_WEIGHT_PT.
   */
  lineWeight?: number;
}

/** Max length of a sparkline's range string (DoS bound on parse). */
export const MAX_SPARKLINE_RANGE_LENGTH = 64;
/** Max number of sparklines a sheet may carry (DoS bound on iteration). */
export const MAX_SPARKLINES = 10_000;
/**
 * Max data points a single sparkline draws. A longer range is decimated to
 * this budget (first, last and each slice's extremes kept, in order), never
 * cut short: the whole range stays the trend the sparkline shows.
 */
export const MAX_SPARKLINE_POINTS = 1_000;
/** Excel's x14 sparklineGroup @lineWeight default, in points. */
export const DEFAULT_SPARKLINE_LINE_WEIGHT_PT = 0.75;
/** Bound imported group identities so hostile JSON cannot grow package attrs. */
export const MAX_SPARKLINE_GROUP_ID_LENGTH = 80;

/** Original x14 graph is visible but cannot be rewritten from its projection. */
export const SPARKLINE_NATIVE_GRAPH_BLOCKER = "sparkline-native-graph" as const;

/**
 * Coerce a free-form type label to one of the three kinds; unknown -> "line"
 * (the most neutral default, matching the dialog's initial selection).
 */
export function normalizeSparklineType(type: unknown): SparklineType {
  const t = typeof type === "string" ? type.trim().toLowerCase() : "";
  if (t === "column" || t === "bar") return "column";
  if (t === "winloss" || t === "win-loss" || t === "winlose") return "winloss";
  return "line";
}

// A #rgb / #rrggbb hex color (case-insensitive). Mirrors normalizeHexColor in
// conditional-format so a persisted color can never inject arbitrary CSS.
const HEX_COLOR_RE = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;

/**
 * Normalize a hex color to a lowercased #rrggbb, or undefined when it is not a
 * valid #rgb / #rrggbb string (so a render can never emit untrusted CSS).
 */
export function normalizeSparklineColor(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  const s = raw.trim();
  if (!HEX_COLOR_RE.test(s)) return undefined;
  const hex = s.slice(1).toLowerCase();
  if (hex.length === 3) {
    return `#${hex[0]}${hex[0]}${hex[1]}${hex[1]}${hex[2]}${hex[2]}`;
  }
  return `#${hex}`;
}

function normalizeSparklineBoolean(raw: unknown): true | undefined {
  return raw === true || raw === 1 || raw === "1" || raw === "true"
    ? true
    : undefined;
}

function normalizeSparklineEmptyCellDisplay(
  raw: unknown,
): SparklineEmptyCellDisplay | undefined {
  return raw === "zero" || raw === "gap" || raw === "span" ? raw : undefined;
}

function normalizeSparklineAxisType(
  raw: unknown,
): SparklineAxisType | undefined {
  return raw === "individual" || raw === "group" || raw === "custom"
    ? raw
    : undefined;
}

function normalizeSparklineManualAxis(raw: unknown): number | undefined {
  const value = typeof raw === "number" ? raw : Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function normalizeSparklineLineWeight(raw: unknown): number | undefined {
  const value = typeof raw === "number" ? raw : Number(raw);
  return Number.isFinite(value) && value >= 0.1 && value <= 20
    ? value
    : undefined;
}

function normalizeSparklineGroupId(raw: unknown): string | undefined {
  if (typeof raw !== "string") return undefined;
  const groupId = raw
    .replace(/[\u0000-\u001f\u007f-\u009f]/g, "")
    .trim()
    .slice(0, MAX_SPARKLINE_GROUP_ID_LENGTH);
  return groupId || undefined;
}

/**
 * Coerce a raw persisted entry into a safe Sparkline, or null when malformed
 * (missing/over-long range, unknown shape). The type defaults to "line"; colors
 * are dropped unless they are valid hex. Extra fields are ignored.
 */
export function normalizeSparkline(raw: unknown): Sparkline | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  if (typeof r.range !== "string") return null;
  const range = r.range.trim();
  if (range.length > MAX_SPARKLINE_RANGE_LENGTH) return null;
  const out: Sparkline = { type: normalizeSparklineType(r.type), range };
  const groupId = normalizeSparklineGroupId(r.groupId);
  if (groupId) out.groupId = groupId;
  for (const field of SPARKLINE_COLOR_FIELDS) {
    const color = normalizeSparklineColor(r[field]);
    if (color) out[field] = color;
  }
  const rawThemeColors = r.themeColors;
  if (rawThemeColors && typeof rawThemeColors === "object") {
    const themeColors: NonNullable<Sparkline["themeColors"]> = {};
    for (const field of SPARKLINE_COLOR_FIELDS) {
      const ref = normalizeThemeColor(
        (rawThemeColors as Record<string, unknown>)[field],
      );
      if (ref && ref.rgb === out[field]) themeColors[field] = ref;
    }
    if (Object.keys(themeColors).length > 0) out.themeColors = themeColors;
  }
  for (const field of [
    "markers",
    "high",
    "low",
    "first",
    "last",
    "negative",
    "displayXAxis",
    "displayHidden",
  ] as const) {
    if (normalizeSparklineBoolean(r[field])) out[field] = true;
  }
  const displayEmptyCellsAs = normalizeSparklineEmptyCellDisplay(
    r.displayEmptyCellsAs,
  );
  if (displayEmptyCellsAs) out.displayEmptyCellsAs = displayEmptyCellsAs;
  const minAxisType = normalizeSparklineAxisType(r.minAxisType);
  const maxAxisType = normalizeSparklineAxisType(r.maxAxisType);
  if (minAxisType) out.minAxisType = minAxisType;
  if (maxAxisType) out.maxAxisType = maxAxisType;
  const manualMin = normalizeSparklineManualAxis(r.manualMin);
  const manualMax = normalizeSparklineManualAxis(r.manualMax);
  if (minAxisType === "custom" && manualMin !== undefined) {
    out.manualMin = manualMin;
  }
  if (maxAxisType === "custom" && manualMax !== undefined) {
    out.manualMax = manualMax;
  }
  if (
    out.manualMin !== undefined &&
    out.manualMax !== undefined &&
    out.manualMin >= out.manualMax
  ) {
    delete out.manualMin;
    delete out.manualMax;
    delete out.minAxisType;
    delete out.maxAxisType;
  }
  const lineWeight = normalizeSparklineLineWeight(r.lineWeight);
  if (lineWeight !== undefined) out.lineWeight = lineWeight;
  return out;
}

/**
 * Normalize a raw `sparklines` map off a persisted sheet: drop malformed entries
 * and cap the entry count (DoS bound). Keys are kept as-is (they are `"row:col"`
 * cell refs, matching the cells map); a non-string key or a value that fails to
 * normalize is dropped. Returns a fresh object.
 */
export function normalizeSparklines(
  raw: unknown,
): Record<string, Sparkline> {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return {};
  const out: Record<string, Sparkline> = {};
  let count = 0;
  for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
    if (count >= MAX_SPARKLINES) break;
    const sparkline = normalizeSparkline(value);
    if (!sparkline) continue;
    out[key] = sparkline;
    count += 1;
  }
  return out;
}

/**
 * The stroke width, in CSS px at 100% zoom, that a line sparkline draws at:
 * its group's lineWeight (points, x14 default 0.75 pt) converted at 96/72.
 * `scale` is the surface's own magnification (the Page Layout page scale);
 * the grid's zoom scales the stroke with the rest of the cell.
 */
export function sparklineStrokeWidthPx(
  sparkline: Pick<Sparkline, "lineWeight">,
  scale = 1,
): number {
  const points = sparkline.lineWeight ?? DEFAULT_SPARKLINE_LINE_WEIGHT_PT;
  const factor = Number.isFinite(scale) && scale > 0 ? scale : 1;
  return (points * 96) / 72 * factor;
}

/** A line-mode point in [0, w] x [0, h] (y is already flipped for SVG). */
export interface SparklinePoint {
  x: number;
  y: number;
}

/**
 * A column/win-loss bar rectangle in SVG coordinates plus a `positive` flag.
 * For "column" bars, `positive` is true for values >= 0; for "winloss", a zero
 * value renders a thin neutral bar (`positive` is null) so wins/losses/ties are
 * all distinguishable.
 */
export interface SparklineBar {
  x: number;
  y: number;
  width: number;
  height: number;
  /** true = win/positive, false = loss/negative, null = zero/tie (winloss only). */
  positive: boolean | null;
}

/** The render model `sparklinePoints` returns, discriminated by `type`. */
export type SparklineRenderModel =
  | {
      type: "line";
      points: SparklinePoint[];
      /** One segment for span/zero; gap mode splits at every blank slot. */
      segments: SparklinePoint[][];
      pointIndexes: number[];
      pointValues: number[];
      zeroY: number | null;
      width: number;
      height: number;
    }
  | {
      type: "column";
      bars: SparklineBar[];
      barIndexes: number[];
      barValues: number[];
      zeroY: number | null;
      width: number;
      height: number;
    }
  | {
      type: "winloss";
      bars: SparklineBar[];
      barIndexes: number[];
      barValues: number[];
      zeroY: number | null;
      width: number;
      height: number;
    };

export interface SparklineRenderOptions {
  displayEmptyCellsAs?: SparklineEmptyCellDisplay;
  includeZero?: boolean;
  minimum?: number;
  maximum?: number;
  /**
   * `minimum` / `maximum` is a Custom Value axis bound (x14 minAxisType /
   * maxAxisType "custom"): column bars start from it as given instead of
   * reaching to zero, as the automatic axis's bars do.
   */
  customMinimum?: boolean;
  customMaximum?: boolean;
}

type SparklineSample = { index: number; value: number | null };

/**
 * Keep finite numbers and explicit blank slots; malformed numeric values drop.
 * `slots` is how many positions the x axis spreads over: the whole range,
 * even when a range longer than MAX_SPARKLINE_POINTS is decimated.
 */
function finiteSamples(values: readonly (number | null)[]): {
  samples: SparklineSample[];
  slots: number;
} {
  const out: SparklineSample[] = [];
  for (const value of values) {
    if (value === null) {
      out.push({ index: out.length, value: null });
    } else if (typeof value === "number" && Number.isFinite(value)) {
      out.push({ index: out.length, value });
    }
  }
  return { samples: decimateSamples(out), slots: out.length };
}

/**
 * Reduce a series longer than MAX_SPARKLINE_POINTS to the budget without
 * changing the trend it draws: the first and last samples stay, and each
 * equal slice in between keeps its lowest and highest value (and one blank
 * slot, so a gap still breaks the line), in their original order and at
 * their original x positions. A peak anywhere in the range therefore still
 * shows, and the high/low/first/last markers still find their points.
 */
function decimateSamples(samples: SparklineSample[]): SparklineSample[] {
  if (samples.length <= MAX_SPARKLINE_POINTS) return samples;
  const first = samples[0];
  const last = samples[samples.length - 1];
  const inner = samples.slice(1, -1);
  // Up to three samples per slice (min, max, one blank) plus first and last.
  const buckets = Math.floor((MAX_SPARKLINE_POINTS - 2) / 3);
  const out: SparklineSample[] = [first];
  for (let bucket = 0; bucket < buckets; bucket += 1) {
    const start = Math.floor((bucket * inner.length) / buckets);
    const end = Math.floor(((bucket + 1) * inner.length) / buckets);
    let low: SparklineSample | null = null;
    let high: SparklineSample | null = null;
    let blank: SparklineSample | null = null;
    for (let index = start; index < end; index += 1) {
      const sample = inner[index];
      if (sample.value === null) {
        blank ??= sample;
        continue;
      }
      if (low === null || sample.value < (low.value as number)) low = sample;
      if (high === null || sample.value > (high.value as number)) high = sample;
    }
    const kept = [low, high, blank].filter(
      (sample, position, list): sample is SparklineSample =>
        sample !== null && list.indexOf(sample) === position,
    );
    kept.sort((a, b) => a.index - b.index);
    out.push(...kept);
  }
  out.push(last);
  return out;
}

function renderBounds(
  values: readonly number[],
  options: SparklineRenderOptions,
): { min: number; max: number } | null {
  if (values.length === 0) return null;
  const dataMin = Math.min(...values);
  const dataMax = Math.max(...values);
  let requestedMin = Number.isFinite(options.minimum)
    ? Number(options.minimum)
    : dataMin;
  let requestedMax = Number.isFinite(options.maximum)
    ? Number(options.maximum)
    : dataMax;
  if (options.includeZero) {
    requestedMin = Math.min(0, requestedMin);
    requestedMax = Math.max(0, requestedMax);
  }
  if (requestedMin >= requestedMax && dataMin !== dataMax) {
    return { min: dataMin, max: dataMax };
  }
  return { min: requestedMin, max: requestedMax };
}

/**
 * Build the SVG render model for a sparkline from its numeric values, type and
 * canvas size (w x h, in px). Pure geometry — no DOM, no color (the renderer
 * applies the sparkline's color to the model).
 *
 *  - "line":    a polyline of the values normalized into [0, h]. The min value
 *               sits at the bottom (y = h), the max at the top (y = 0). With a
 *               single point (or all-equal values) the line is drawn flat at the
 *               vertical middle. x is spread evenly across [0, w].
 *  - "column":  one mini bar per value. Bars share a baseline; a positive value
 *               grows up from the zero line, a negative value down. The bar
 *               height is proportional to |value| over the data's magnitude span.
 *  - "winloss": one fixed-height bar per value — up (positive) for a win, down
 *               (negative) for a loss, a thin centered bar for zero (tie). The
 *               magnitude is ignored; only the sign matters (Excel parity).
 *
 * An empty value list yields an empty points/bars array (the cell renders blank).
 */
export function sparklinePoints(
  values: readonly (number | null)[],
  type: SparklineType,
  w: number,
  h: number,
  options: SparklineRenderOptions = {},
): SparklineRenderModel {
  const width = Number.isFinite(w) && w > 0 ? w : 0;
  const height = Number.isFinite(h) && h > 0 ? h : 0;
  const { samples, slots } = finiteSamples(values);
  const emptyMode = options.displayEmptyCellsAs ?? "zero";
  const plotted = samples.map((sample) => ({
    ...sample,
    value: sample.value === null && emptyMode === "zero" ? 0 : sample.value,
  }));
  const nums = plotted.flatMap((sample) =>
    sample.value === null ? [] : [sample.value],
  );
  const bounds = renderBounds(nums, options);

  if (type === "line") {
    if (!bounds || nums.length === 0) {
      return {
        type: "line",
        points: [],
        segments: [],
        pointIndexes: [],
        pointValues: [],
        zeroY: null,
        width,
        height,
      };
    }
    const { min, max } = bounds;
    const span = max - min;
    const step = slots > 1 ? width / (slots - 1) : 0;
    const points: SparklinePoint[] = [];
    const pointIndexes: number[] = [];
    const pointValues: number[] = [];
    const segments: SparklinePoint[][] = [];
    let segment: SparklinePoint[] = [];
    for (const sample of plotted) {
      if (sample.value === null) {
        if (emptyMode === "gap" && segment.length > 0) {
          segments.push(segment);
          segment = [];
        }
        continue;
      }
      const x = slots > 1 ? sample.index * step : width / 2;
      // Normalize into [0, 1] then flip so the max is at the top (y = 0). When
      // every value is equal (span = 0), draw the line through the middle.
      const norm = span === 0 ? 0.5 : (sample.value - min) / span;
      const point = {
        x,
        y: height - Math.max(0, Math.min(1, norm)) * height,
      };
      points.push(point);
      pointIndexes.push(sample.index);
      pointValues.push(sample.value);
      segment.push(point);
    }
    if (segment.length > 0) segments.push(segment);
    const zeroY =
      min <= 0 && max >= 0 && span !== 0
        ? height - ((0 - min) / span) * height
        : null;
    return {
      type: "line",
      points,
      segments,
      pointIndexes,
      pointValues,
      zeroY,
      width,
      height,
    };
  }

  // column / winloss share the bar-layout math (evenly spaced, gapped bars).
  if (!bounds || nums.length === 0 || samples.length === 0) {
    return {
      type,
      bars: [],
      barIndexes: [],
      barValues: [],
      zeroY: null,
      width,
      height,
    } as SparklineRenderModel;
  }
  // Each value gets an equal slot; the bar fills 70% of its slot, centered.
  const slot = width / slots;
  const barWidth = Math.max(1, slot * 0.7);
  const barInset = (slot - barWidth) / 2;
  const numericSamples = plotted.filter(
    (sample): sample is { index: number; value: number } =>
      sample.value !== null,
  );

  if (type === "winloss") {
    // Wins fill the top half, losses the bottom half, a tie is a thin centered
    // bar. Magnitude is ignored — only the sign matters.
    const mid = height / 2;
    const winHeight = mid;
    const tieHeight = Math.max(1, height * 0.08);
    const bars: SparklineBar[] = numericSamples.map((sample) => {
      const v = sample.value;
      const x = sample.index * slot + barInset;
      if (v > 0) {
        return { x, y: mid - winHeight, width: barWidth, height: winHeight, positive: true };
      }
      if (v < 0) {
        return { x, y: mid, width: barWidth, height: winHeight, positive: false };
      }
      return {
        x,
        y: mid - tieHeight / 2,
        width: barWidth,
        height: tieHeight,
        positive: null,
      };
    });
    return {
      type: "winloss",
      bars,
      barIndexes: numericSamples.map((sample) => sample.index),
      barValues: numericSamples.map((sample) => sample.value),
      zeroY: mid,
      width,
      height,
    };
  }

  // column: the axis runs from its minimum to its maximum, automatic ("For
  // Each Sparkline": the series' own lowest and highest value) or a Custom
  // Value, and the bars grow from the value-0 line when it is in view, else
  // from the edge zero lies beyond. So bars of all-positive data rise from
  // the series' lowest value, which keeps a hairline: measured in Excel,
  // 10, 8, 7, 7, 5 draws 28:18:10:10:3 px, not five near-equal bars from 0.
  const { min, max } = bounds;
  const lo = min;
  const hi = max;
  const span = hi - lo;
  const clampY = (y: number) => Math.max(0, Math.min(height, y));
  // y position (in SVG coords) of the value-0 baseline; a custom axis that
  // leaves zero out of view grows the bars from the edge zero lies beyond.
  const baselineY =
    span === 0 ? height : clampY(height - ((0 - lo) / span) * height);
  const zeroY =
    span === 0 ? height : lo <= 0 && hi >= 0 ? baselineY : null;
  const bars: SparklineBar[] = numericSamples.map((sample) => {
    const v = sample.value;
    const x = sample.index * slot + barInset;
    if (span === 0) {
      // All equal (and equal to 0, since lo=hi=0 only when every value is 0):
      // draw a thin bar at the baseline so the cell is not blank.
      const thin = Math.max(1, height * 0.08);
      return { x, y: height - thin, width: barWidth, height: thin, positive: v >= 0 };
    }
    const valueY = clampY(height - ((v - lo) / span) * height);
    // A zero on a drawn zero line is no bar at all, as Excel draws it
    // (S21_Sparklines, X-198); any other value keeps the 1px floor.
    const barHeight = v === 0 && zeroY !== null ? 0 : Math.max(1, Math.abs(valueY - baselineY));
    // The bar stands on the zero line and goes the value's way, the 1px
    // floor included: grown down from the value's own place, a small
    // positive (1 against 100) drew 0.9px of its pixel under the zero line
    // and read as a negative (S-113).
    const top =
      v > 0
        ? Math.max(0, baselineY - barHeight)
        : v < 0
          ? Math.min(baselineY, height - barHeight)
          : Math.min(valueY, baselineY);
    return { x, y: top, width: barWidth, height: barHeight, positive: v >= 0 };
  });
  return {
    type: "column",
    bars,
    barIndexes: numericSamples.map((sample) => sample.index),
    barValues: numericSamples.map((sample) => sample.value),
    zeroY,
    width,
    height,
  };
}
