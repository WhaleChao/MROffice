/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface EtsForecastInput {
  /** The target point on the timeline to forecast (x). */
  target: number;
  /** Historical values (y), paired with `timeline` by index. */
  values: readonly number[];
  /** Historical timeline (x), strictly the same length as `values`. */
  timeline: readonly number[];
  /**
   * Seasonality: 1 (or omitted) ⇒ auto-detect; 0 ⇒ no seasonality (Holt linear);
   * N ≥ 2 ⇒ a fixed period of N points.
   */
  seasonality?: number;
}

/** The median positive gap between consecutive sorted timeline points. */
function medianStep(sortedX: number[]): number | null {
  const gaps: number[] = [];
  for (let i = 1; i < sortedX.length; i += 1) {
    const d = sortedX[i] - sortedX[i - 1];
    if (d > 0) gaps.push(d);
  }
  if (gaps.length === 0) return null;
  gaps.sort((a, b) => a - b);
  const mid = Math.floor(gaps.length / 2);
  return gaps.length % 2 ? gaps[mid] : (gaps[mid - 1] + gaps[mid]) / 2;
}










function detectSeasonality(values: number[]): number {
  const n = values.length;
  if (n < 4) return 1;
  const diffs: number[] = [];
  for (let i = 1; i < n; i += 1) diffs.push(values[i] - values[i - 1]);
  const m = diffs.length;
  const mean = diffs.reduce((s, v) => s + v, 0) / m;
  const centered = diffs.map((v) => v - mean);
  const denominator = centered.reduce((s, v) => s + v * v, 0);
  if (denominator <= 0) return 1;
  const threshold = 1.5 / Math.sqrt(m);
  for (let lag = 2; lag <= Math.floor(n / 2); lag += 1) {
    let sum = 0;
    for (let i = lag; i < m; i += 1) sum += centered[i] * centered[i - lag];
    if (sum / denominator > threshold) return lag;
  }
  return 1;
}

/** Ordinary least squares of `ys` on its index 0..n-1. */
function leastSquares(ys: readonly number[]): { intercept: number; slope: number } {
  const n = ys.length;
  const meanX = (n - 1) / 2;
  const meanY = ys.reduce((s, v) => s + v, 0) / n;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i += 1) {
    sxy += (i - meanX) * (ys[i] - meanY);
    sxx += (i - meanX) * (i - meanX);
  }
  const slope = sxx > 0 ? sxy / sxx : 0;
  return { intercept: meanY - slope * meanX, slope };
}

interface EtsState {
  level: number;
  trend: number;
  seasonals: number[];
}











/**
 * Excel's start trend for a series of 4, 5 or 6 points is not the
 * least-squares slope but a fixed weighting of the points (solved as an
 * overdetermined linear system from 24 short series on the local Excel,
 * residual ~1e-26): these weights, scaled by their denominators.
 */
const SHORT_SERIES_TREND_WEIGHTS: Readonly<Record<number, { weights: readonly number[]; scale: number }>> = {
  4: { weights: [-8, -3, 2, 9], scale: 28 },
  5: { weights: [-26, -13, -5, 13, 31], scale: 140 },
  6: { weights: [-90, -124, -53, 18, 89, 160], scale: 980 },
};

function initialState(ys: readonly number[], period: number): EtsState {
  const n = ys.length;
  const { intercept, slope } = leastSquares(ys);
  if (period <= 1) {
    const short = SHORT_SERIES_TREND_WEIGHTS[n];
    const trend = short
      ? short.weights.reduce((sum, weight, i) => sum + weight * ys[i], 0) / short.scale
      : slope;
    return { level: ys[0], trend, seasonals: [0] };
  }
  const sums = new Array<number>(period).fill(0);
  const counts = new Array<number>(period).fill(0);
  for (let i = 0; i < n; i += 1) {
    sums[i % period] += ys[i] - (intercept + slope * i);
    counts[i % period] += 1;
  }
  const seasonals = sums.map((sum, k) => (counts[k] > 0 ? sum / counts[k] : 0));
  const deseasonalized = ys.map((y, i) => y - seasonals[i % period]);
  return { level: ys[0] - seasonals[0], trend: leastSquares(deseasonalized).slope, seasonals };
}









function runEts(
  ys: readonly number[],
  period: number,
  alpha: number,
  beta: number,
  gamma: number,
  h: number,
): { forecast: number; sse: number; errors: number[] } {
  const n = ys.length;
  const seasonal = period > 1;
  const state = initialState(ys, period);
  let { level, trend } = state;
  const seasonals = [...state.seasonals];
  let sse = 0;
  const errors: number[] = [0];
  for (let t = 1; t < n; t += 1) {
    const season = seasonal ? seasonals[t % period] : 0;
    const e = ys[t] - (level + trend + season);
    sse += e * e;
    errors.push(e);
    level = level + trend + alpha * e;
    trend += beta * e;
    if (seasonal) seasonals[t % period] += gamma * e;
  }
  const seasonComponent = seasonal ? seasonals[(n + h - 1) % period] : 0;
  return { forecast: level + h * trend + seasonComponent, sse, errors };
}

/**
 * The seasonal model's alpha and gamma are searched on the grid
 * 0.001 + k·0.125 with beta at 0.001, minimising the in-sample one-step SSE
 * (Excel's seasonal fits land within about a percent of this).
 */
const SMOOTHING_GRID = Array.from({ length: 8 }, (_, k) => 0.001 + k * 0.125);
const BETA = 0.001;















const NON_SEASONAL_CANDIDATES: ReadonlyArray<readonly [number, number]> = [
  ...[0.002, 0.126, 0.25, 0.5, 0.75, 0.9, 0.998].map((alpha) => [alpha, 0.001] as const),
  ...[0.126, 0.25, 0.5, 2 / 3, 0.75, 0.9, 0.998].map(
    (alpha) => [alpha, Math.round((alpha - 0.001) * 1e6) / 1e6] as const,
  ),
  [0.1, 0.099],
  [2 / 3, 0.5],
  [0.998, 0.499],
  [0.5, 0.25],
  [1, 0],
];

/**
 * Forecast `target` via additive Holt–Winters. Returns null when there is too
 * little data, the timeline has no positive step, or the result is non-finite.
 * The smoothing constants are chosen by a deterministic grid search that
 * minimises in-sample one-step SSE.
 */
/** One fitted additive Holt–Winters model over a prepared series. */
export interface EtsFit {
  alpha: number;
  beta: number;
  gamma: number;
  /** Detected (or requested) season length; 1 = no seasonality. */
  period: number;
  /** Median timeline step. */
  step: number;
  lastX: number;
  ys: number[];
  /** One-step-ahead in-sample errors (from the second cycle on). */
  errors: number[];
}

function prepareSeries(
  values: readonly number[],
  timeline: readonly number[],
): { xs: number[]; ys: number[] } | null {
  if (values.length !== timeline.length || values.length < 2) return null;
  // Pair + sort by timeline; aggregate duplicate timestamps by mean (Excel's
  // default data_completion/aggregation). Non-finite points are dropped.
  const pairs: Array<{ x: number; y: number }> = [];
  for (let i = 0; i < values.length; i += 1) {
    if (Number.isFinite(values[i]) && Number.isFinite(timeline[i])) {
      pairs.push({ x: timeline[i], y: values[i] });
    }
  }
  if (pairs.length < 2) return null;
  pairs.sort((a, b) => a.x - b.x);
  const merged: Array<{ x: number; y: number }> = [];
  for (const p of pairs) {
    const last = merged[merged.length - 1];
    if (last && last.x === p.x) {
      last.y = (last.y + p.y) / 2;
    } else {
      merged.push({ ...p });
    }
  }
  if (merged.length < 2) return null;
  return { xs: merged.map((p) => p.x), ys: merged.map((p) => p.y) };
}

/**
 * Fit the model: the season (given or detected), then alpha and gamma on
 * Excel's grid by in-sample SSE. Shared by FORECAST.ETS and its
 * .SEASONALITY / .STAT / .CONFINT companions so the four functions
 * describe ONE model.
 */
export function fitEts(
  values: readonly number[],
  timeline: readonly number[],
  seasonality?: number,
): EtsFit | null {
  const prepared = prepareSeries(values, timeline);
  if (!prepared) return null;
  const { xs, ys } = prepared;
  const step = medianStep(xs);
  if (step === null || step <= 0) return null;
  let period = seasonality ?? 1;
  if (period === 1) period = detectSeasonality(ys);
  if (period < 1 || period > Math.floor(ys.length / 2)) period = 1;
  let best: EtsFit | null = null;
  let bestSse = Number.POSITIVE_INFINITY;
  const consider = (alpha: number, beta: number, gamma: number) => {
    const { forecast, sse, errors } = runEts(ys, period, alpha, beta, gamma, 1);
    if (Number.isFinite(forecast) && sse < bestSse) {
      bestSse = sse;
      best = { alpha, beta, gamma, period, step, lastX: xs[xs.length - 1], ys, errors };
    }
  };
  if (period > 1) {
    for (const alpha of SMOOTHING_GRID) {
      for (const gamma of SMOOTHING_GRID) consider(alpha, BETA, gamma);
    }
  } else {
    for (const [alpha, beta] of NON_SEASONAL_CANDIDATES) consider(alpha, beta, 0);
  }
  return best;
}

/**
 * Forecast `target` via additive Holt–Winters. Returns null when there is too
 * little data, the timeline has no usable step, or the target is not ahead of
 * the series (FORECAST.ETS only projects forward).
 */
export function forecastEts(input: EtsForecastInput): number | null {
  const fit = fitEts(input.values, input.timeline, input.seasonality);
  if (!fit) return null;
  const h = Math.round((input.target - fit.lastX) / fit.step);
  if (h < 1) return null;
  const { forecast } = runEts(fit.ys, fit.period, fit.alpha, fit.beta, fit.gamma, h);
  return Number.isFinite(forecast) ? forecast : null;
}

/** FORECAST.ETS.SEASONALITY: the detected season length, 0 when there is none. */
export function forecastEtsSeasonality(
  values: readonly number[],
  timeline: readonly number[],
): number | null {
  const fit = fitEts(values, timeline, 1);
  if (!fit) return null;
  return fit.period > 1 ? fit.period : 0;
}

/**
 * FORECAST.ETS.STAT statistic_type: 1 alpha, 2 beta, 3 gamma, 4 MASE,
 * 5 SMAPE, 6 MAE, 7 RMSE, 8 step size. Error metrics are the model's in-sample
 * one-step-ahead errors, so they describe the same fit FORECAST.ETS projects.
 */
export function forecastEtsStat(
  values: readonly number[],
  timeline: readonly number[],
  statisticType: number,
  seasonality?: number,
): number | null {
  const fit = fitEts(values, timeline, seasonality);
  if (!fit) return null;
  const { errors, ys } = fit;
  const n = errors.length;
  switch (statisticType) {
    case 1:
      return fit.alpha;
    case 2:
      return fit.beta;
    case 3:
      return fit.gamma;
    case 8:
      return fit.step;
    default:
      break;
  }
  if (n === 0) return null;
  // Excel measures the fit on a validation window: every point for a series
  // of up to 10, otherwise the last min(floor(0.4·n), 10); MASE scales by the
  // mean one-step naive error over the points before that window (all of
  
  const window = n <= 10 ? n : Math.min(Math.floor(0.4 * n), 10);
  const from = n - window;
  const held = errors.slice(from);
  const mae = held.reduce((s, e) => s + Math.abs(e), 0) / window;
  if (statisticType === 6) return mae;
  if (statisticType === 7) return Math.sqrt(held.reduce((s, e) => s + e * e, 0) / window);
  if (statisticType === 4) {
    const trainEnd = from === 0 ? n : from;
    let naive = 0;
    for (let t = 1; t < trainEnd; t += 1) naive += Math.abs(ys[t] - ys[t - 1]);
    naive /= Math.max(1, trainEnd - 1);
    return naive === 0 ? null : mae / naive;
  }
  if (statisticType === 5) {
    let total = 0;
    for (let i = from; i < n; i += 1) {
      const actual = ys[i];
      const predicted = actual - errors[i];
      const denominator = Math.abs(actual) + Math.abs(predicted);
      total += denominator === 0 ? 0 : (2 * Math.abs(errors[i])) / denominator;
    }
    return total / window;
  }
  return null;
}

/** Standard-normal quantile (Acklam's rational approximation, |err| < 1.2e-9). */
function normalQuantile(p: number): number {
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const low = 0.02425;
  if (p < low) {
    const q = Math.sqrt(-2 * Math.log(p));
    return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  if (p > 1 - low) {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  const q = p - 0.5;
  const r = q * q;
  return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
}

/**
 * FORECAST.ETS.CONFINT: the half-width of the prediction interval around the
 * forecast at `target`: z · RMSE · sqrt(1 + Σ_{j=1}^{h-1} c_j²), the AAA
 * model's h-step variance, with c_j = alpha + j·beta (+ gamma where j is a
 * whole number of seasons). With Excel's own constants this is Excel's
 * interval to the digit at every horizon (S12_tdf131380: alpha 0.998, beta
 * 0.899, ±456.13 at h=1 growing to ±25,447 at h=22; the old 1 + (h-1)·alpha²
 * left out the trend and stayed near ±1,900).
 */
export function forecastEtsConfint(
  target: number,
  values: readonly number[],
  timeline: readonly number[],
  confidence = 0.95,
  seasonality?: number,
): number | null {
  if (!(confidence > 0 && confidence < 1)) return null;
  const fit = fitEts(values, timeline, seasonality);
  if (!fit) return null;
  const h = Math.round((target - fit.lastX) / fit.step);
  if (h < 1) return null;
  const n = fit.errors.length;
  if (n === 0) return null;
  const rmse = Math.sqrt(fit.errors.reduce((s, e) => s + e * e, 0) / n);
  const z = normalQuantile(1 - (1 - confidence) / 2);
  let variance = 1;
  for (let j = 1; j < h; j += 1) {
    const c = fit.alpha + j * fit.beta + (fit.period > 1 && j % fit.period === 0 ? fit.gamma : 0);
    variance += c * c;
  }
  return z * rmse * Math.sqrt(variance);
}
