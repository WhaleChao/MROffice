/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Forecast Sheet (数据 ▸ 预测工作表).
 *
 * Excel and WPS turn a history of dated values into a small table: the
 * timeline continues at its own step, each future point gets a forecast, and
 * the forecast is bracketed by a confidence interval so nobody reads it as a
 * promise. The maths is the workbook's existing ETS implementation
 * (`forecast-ets.ts`, the same one behind FORECAST.ETS), so a forecast here and
 * a FORECAST.ETS in a cell agree.
 *
 * The interval is FORECAST.ETS.CONFINT's, so the band the chart first draws is
 * the one the sheet's cells then compute. Where the model has no interval to
 * give, it is the normal-theory band around the residual standard error,
 * widened as the horizon grows.
 */

import { forecastEts, forecastEtsConfint } from "./forecast-ets";

export interface ForecastSheetInput {
  /** Historical x values (dates as serial numbers, or plain periods). */
  timeline: readonly number[];
  /** Historical y values, paired with `timeline` by index. */
  values: readonly (number | string | null | undefined)[];
  /** How many points to forecast past the end of the history. */
  periods: number;
  /** 0…1, defaults to 0.95. */
  confidence?: number;
  /** Passed through to ETS: 1 auto-detect, 0 none, N ≥ 2 fixed period. */
  seasonality?: number;
  /** How gaps in an otherwise regular timeline are completed. */
  missingPoints?: ForecastMissingPointPolicy;
  /** How multiple values at the same timestamp are reduced to one point. */
  duplicateAggregation?: ForecastDuplicateAggregation;
}

export type ForecastMissingPointPolicy = "interpolate" | "zero";

export type ForecastDuplicateAggregation =
  "average" | "sum" | "count" | "counta" | "min" | "max" | "median";

export interface NormalizedForecastHistory {
  timeline: number[];
  values: number[];
  step: number;
  missingPointsFilled: number;
  duplicatePointsAggregated: number;
}

export interface ForecastSheetRow {
  timeline: number;
  forecast: number;
  lower: number;
  upper: number;
}

export interface ForecastSheetResult {
  rows: ForecastSheetRow[];
  /** Step the timeline continues at (median positive gap in the history). */
  step: number;
  /** Residual standard error of the in-sample fit. */
  standardError: number;
  /** Number of inferred/blank historical points completed before fitting. */
  missingPointsFilled: number;
  /** Number of extra same-timestamp observations reduced before fitting. */
  duplicatePointsAggregated: number;
}

export type ForecastSheetError =
  | "too-few-points"
  | "mismatched-lengths"
  | "invalid-periods"
  | "invalid-confidence"
  | "invalid-seasonality"
  | "invalid-timeline"
  | "too-many-missing-points"
  | "unfillable-missing-points"
  | "not-forecastable";

/** Excel refuses below this; so do we, rather than extrapolating from noise. */
export const MIN_FORECAST_POINTS = 3;
export const MAX_FORECAST_PERIODS = 1_000;
/** Native FORECAST.ETS upper bound: hourly points across one leap-free year. */
export const MAX_FORECAST_SEASONALITY = 8_760;

/**
 * Standard normal quantile — Acklam's rational approximation, whose relative
 * error stays under 1.15e-9 across the whole range.
 *
 * This replaced a table of the common confidence levels that rounded DOWN to
 * the nearest one it knew: 0.93 was served the 90% quantile and 0.98 the 97.5%,
 * so the band came back NARROWER than the level printed beside it, and any
 * level under 0.8 got a fixed 1.036 — the 70% quantile — whatever was asked
 * for. Understating uncertainty under the user's own label is the failure this
 * module exists to avoid, so the quantile is now computed rather than looked up.
 */
function inverseNormalCdf(probability: number): number {
  const a = [
    -3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2,
    1.38357751867269e2, -3.066479806614716e1, 2.506628277459239,
  ];
  const b = [
    -5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2,
    6.680131188771972e1, -1.328068155288572e1,
  ];
  const c = [
    -7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838,
    -2.549732539343734, 4.374664141464968, 2.938163982698783,
  ];
  const d = [
    7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996,
    3.754408661907416,
  ];
  const lowTail = 0.02425;

  if (probability < lowTail) {
    const q = Math.sqrt(-2 * Math.log(probability));
    return (
      (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    );
  }
  if (probability <= 1 - lowTail) {
    const q = probability - 0.5;
    const r = q * q;
    return (
      ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) *
        q) /
      (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
    );
  }
  const q = Math.sqrt(-2 * Math.log(1 - probability));
  return (
    -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
    ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  );
}

/** Two-sided normal quantile for a confidence level in (0, 1). */
function zFor(confidence: number): number {
  return inverseNormalCdf((1 + confidence) / 2);
}

function medianStep(timeline: readonly number[]): number | null {
  const sorted = [...timeline].sort((a, b) => a - b);
  const gaps: number[] = [];
  for (let index = 1; index < sorted.length; index += 1) {
    const gap = sorted[index] - sorted[index - 1];
    if (gap > 0) gaps.push(gap);
  }
  if (gaps.length === 0) return null;
  gaps.sort((a, b) => a - b);
  const middle = Math.floor(gaps.length / 2);
  return gaps.length % 2 ? gaps[middle] : (gaps[middle - 1] + gaps[middle]) / 2;
}

function numericValue(
  value: number | string | null | undefined,
): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || value.trim() === "") return null;
  const parsed = Number(value.trim());
  return Number.isFinite(parsed) ? parsed : null;
}

function hasValue(value: number | string | null | undefined): boolean {
  return value !== null && value !== undefined && String(value).trim() !== "";
}

function aggregateDuplicates(
  values: readonly (number | string | null | undefined)[],
  aggregation: ForecastDuplicateAggregation,
): number | null {
  const numeric = values
    .map(numericValue)
    .filter((value): value is number => value !== null);
  if (aggregation === "count") return numeric.length;
  if (aggregation === "counta") return values.filter(hasValue).length;
  if (numeric.length === 0) return null;
  if (aggregation === "sum")
    return numeric.reduce((sum, value) => sum + value, 0);
  if (aggregation === "min") return Math.min(...numeric);
  if (aggregation === "max") return Math.max(...numeric);
  if (aggregation === "median") {
    const sorted = [...numeric].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2
      ? sorted[middle]
      : (sorted[middle - 1] + sorted[middle]) / 2;
  }
  return numeric.reduce((sum, value) => sum + value, 0) / numeric.length;
}

function inferredStep(timeline: readonly number[]): number | null {
  const gaps: number[] = [];
  for (let index = 1; index < timeline.length; index += 1) {
    const gap = timeline[index] - timeline[index - 1];
    if (gap > 0) gaps.push(gap);
  }
  if (gaps.length === 0) return null;
  const median = medianStep(timeline);
  const candidates = [
    ...(median === null ? [] : [median]),
    ...[...new Set(gaps)].sort((a, b) => b - a),
  ];
  return (
    candidates.find((candidate) =>
      gaps.every((gap) => {
        const intervals = Math.round(gap / candidate);
        return (
          intervals >= 1 &&
          Math.abs(gap - intervals * candidate) <=
            Math.max(1, Math.abs(gap)) * 1e-9
        );
      }),
    ) ?? null
  );
}

/**
 * Apply Forecast Sheet's explicit timeline policy before the ETS model runs.
 * Missing points are bounded to 30%; duplicate aggregation affects only actual
 * duplicate timestamps, never ordinary single observations.
 */
export function normalizeForecastHistory(
  input: Pick<
    ForecastSheetInput,
    "timeline" | "values" | "missingPoints" | "duplicateAggregation"
  >,
):
  | { ok: true; history: NormalizedForecastHistory }
  | { ok: false; error: ForecastSheetError } {
  if (input.timeline.length !== input.values.length) {
    return { ok: false, error: "mismatched-lengths" };
  }
  if (input.timeline.some((point) => !Number.isFinite(point))) {
    return { ok: false, error: "invalid-timeline" };
  }

  const aggregation = input.duplicateAggregation ?? "average";
  const groups = input.timeline
    .map((point, index) => ({ point, value: input.values[index] }))
    .sort((left, right) => left.point - right.point)
    .reduce<
      Array<{ point: number; values: Array<(typeof input.values)[number]> }>
    >((result, entry) => {
      const current = result[result.length - 1];
      if (current?.point === entry.point) current.values.push(entry.value);
      else result.push({ point: entry.point, values: [entry.value] });
      return result;
    }, []);
  const step = inferredStep(groups.map((group) => group.point));
  if (step === null || step <= 0) {
    return { ok: false, error: "invalid-timeline" };
  }
  if (groups.length < MIN_FORECAST_POINTS) {
    return { ok: false, error: "too-few-points" };
  }

  const expandedTimeline: number[] = [];
  const expandedValues: Array<number | null> = [];
  let duplicatePointsAggregated = 0;
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    expandedTimeline.push(group.point);
    expandedValues.push(
      group.values.length === 1
        ? numericValue(group.values[0])
        : aggregateDuplicates(group.values, aggregation),
    );
    duplicatePointsAggregated += Math.max(0, group.values.length - 1);
    const next = groups[index + 1];
    if (!next) continue;
    const intervals = Math.round((next.point - group.point) / step);
    for (let offset = 1; offset < intervals; offset += 1) {
      expandedTimeline.push(group.point + step * offset);
      expandedValues.push(null);
    }
  }

  const missingPointsFilled = expandedValues.filter(
    (value) => value === null,
  ).length;
  if (missingPointsFilled / expandedValues.length > 0.3) {
    return { ok: false, error: "too-many-missing-points" };
  }

  if ((input.missingPoints ?? "interpolate") === "zero") {
    for (let index = 0; index < expandedValues.length; index += 1) {
      if (expandedValues[index] === null) expandedValues[index] = 0;
    }
  } else {
    let index = 0;
    while (index < expandedValues.length) {
      if (expandedValues[index] !== null) {
        index += 1;
        continue;
      }
      const start = index;
      while (index < expandedValues.length && expandedValues[index] === null) {
        index += 1;
      }
      const before = start - 1;
      const after = index;
      if (before < 0 || after >= expandedValues.length) {
        return { ok: false, error: "unfillable-missing-points" };
      }
      const beforeValue = expandedValues[before];
      const afterValue = expandedValues[after];
      if (beforeValue === null || afterValue === null) {
        return { ok: false, error: "unfillable-missing-points" };
      }
      const span = expandedTimeline[after] - expandedTimeline[before];
      for (let fill = start; fill < after; fill += 1) {
        const weight =
          (expandedTimeline[fill] - expandedTimeline[before]) / span;
        expandedValues[fill] =
          beforeValue + (afterValue - beforeValue) * weight;
      }
    }
  }

  return {
    ok: true,
    history: {
      timeline: expandedTimeline,
      values: expandedValues as number[],
      step,
      missingPointsFilled,
      duplicatePointsAggregated,
    },
  };
}

/**
 * Build the forecast table, or say why it cannot be built. Never throws and
 * never invents a point: a horizon the model cannot produce ends the table
 * instead of repeating the last value.
 */
export function buildForecastSheet(
  input: ForecastSheetInput,
):
  | { ok: true; result: ForecastSheetResult }
  | { ok: false; error: ForecastSheetError } {
  const normalized = normalizeForecastHistory(input);
  if (!normalized.ok) return normalized;
  const { timeline, values, step } = normalized.history;
  if (
    !Number.isFinite(input.periods) ||
    input.periods < 1 ||
    input.periods > MAX_FORECAST_PERIODS
  ) {
    return { ok: false, error: "invalid-periods" };
  }
  // A confidence outside (0, 1) silently produced a band anyway — 50 read as
  // 99.5%, NaN as ~72% — so the interval came back wider than the user asked
  // for while still carrying their label. Refuse instead.
  const confidence = input.confidence ?? 0.95;
  if (!Number.isFinite(confidence) || confidence <= 0 || confidence >= 1) {
    return { ok: false, error: "invalid-confidence" };
  }
  // 1 means automatic detection, 0 means a linear/no-seasonality forecast,
  // and a fixed period needs two complete historical cycles. Refuse every
  // other value instead of letting forecastEts silently fall back to linear.
  const seasonality = input.seasonality ?? 1;
  if (
    !Number.isFinite(seasonality) ||
    !Number.isInteger(seasonality) ||
    seasonality < 0 ||
    seasonality > MAX_FORECAST_SEASONALITY ||
    (seasonality > 1 && seasonality * 2 > timeline.length)
  ) {
    return { ok: false, error: "invalid-seasonality" };
  }
  const forecastAt = (target: number): number | null =>
    forecastEts({
      target,
      values,
      timeline,
      seasonality,
    });

  // Residual standard error from one-step-ahead fits over the history, which
  // is what makes the band reflect THIS series' noise rather than a constant.
  // Starting at 2 because that is the shortest history the model can fit from.
  // Starting at MIN_FORECAST_POINTS left a three-point series with no residual
  // at all, and a standard error of zero prints lower = forecast = upper: a
  // band claiming the forecast is exact.
  let squaredError = 0;
  let residuals = 0;
  for (let index = 2; index < timeline.length; index += 1) {
    const fitted = forecastEts({
      target: timeline[index],
      values: values.slice(0, index),
      timeline: timeline.slice(0, index),
      seasonality,
    });
    if (fitted === null) continue;
    squaredError += (values[index] - fitted) ** 2;
    residuals += 1;
  }
  // No residual means no way to say how wide the band is, and a zero-width one
  // would claim certainty this model does not have.
  if (residuals === 0) return { ok: false, error: "not-forecastable" };
  const standardError = Math.sqrt(squaredError / residuals);
  const z = zFor(confidence);
  const last = Math.max(...timeline);

  const rows: ForecastSheetRow[] = [];
  for (let horizon = 1; horizon <= input.periods; horizon += 1) {
    const target = last + step * horizon;
    const forecast = forecastAt(target);
    if (forecast === null || !Number.isFinite(forecast)) break;
    const margin =
      forecastEtsConfint(target, values, timeline, confidence, seasonality) ??
      z * standardError * Math.sqrt(horizon);
    rows.push({
      timeline: target,
      forecast,
      lower: forecast - margin,
      upper: forecast + margin,
    });
  }
  if (rows.length === 0) return { ok: false, error: "not-forecastable" };
  return {
    ok: true,
    result: {
      rows,
      step,
      standardError,
      missingPointsFilled: normalized.history.missingPointsFilled,
      duplicatePointsAggregated: normalized.history.duplicatePointsAggregated,
    },
  };
}
