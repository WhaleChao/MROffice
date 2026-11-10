/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Estimating how far a scanned page leans.
 *
 * A scan is rarely straight, and a page that leans two degrees reads badly,
 * OCRs worse and prints crooked. WPS ships this as 纠偏 (PageCorrect); the
 * core of it is one question — by what angle are the text lines off
 * horizontal? — and that question is answerable without any DOM: rotate the
 * ink by a candidate angle, project it onto rows, and measure how sharply the
 * projection separates into lines. Straight text piles ink into narrow bands
 * with empty gaps; leaning text smears it evenly. The variance of the row
 * profile scores exactly that.
 *
 * Pure function over grayscale pixels, so the tests draw known-skew synthetic
 * pages and check the estimate — no scanner required.
 */

export interface GrayscaleImage {
  /** One byte per pixel, 0 = black ink, 255 = white paper, row-major. */
  data: Uint8Array | Uint8ClampedArray;
  width: number;
  height: number;
}

export interface SkewEstimate {
  /**
   * Clockwise lean of the text in degrees; rotate the page by the NEGATIVE of
   * this to straighten it.
   */
  angle: number;
  /**
   * 0..1. How decisively the best angle beat the field. Near zero means "this
   * page has no line structure to align" — a photo, a blank page — and a
   * caller should not rotate anything on that evidence.
   */
  confidence: number;
}

export interface EstimateSkewOptions {
  /** Search range, degrees either side of straight. */
  maxAngle?: number;
  /** Coarse sweep step in degrees. */
  coarseStep?: number;
  /** Fine sweep step around the coarse winner. */
  fineStep?: number;
  /** Pixels darker than this count as ink. */
  inkThreshold?: number;
  /** Cap on sampled ink pixels, for time. */
  maxSamples?: number;
}

const DEG = Math.PI / 180;

/** Collect (x, y) of ink pixels, evenly subsampled to the cap. */
function inkPixels(
  image: GrayscaleImage,
  threshold: number,
  maxSamples: number,
): { xs: Float64Array; ys: Float64Array; count: number } {
  const { data, width, height } = image;
  let total = 0;
  for (let i = 0; i < data.length; i += 1) {
    if (data[i] < threshold) total += 1;
  }
  const stride = Math.max(1, Math.ceil(total / maxSamples));
  const xs = new Float64Array(Math.min(total, maxSamples) + 1);
  const ys = new Float64Array(xs.length);
  let count = 0;
  let seen = 0;
  for (let y = 0; y < height; y += 1) {
    const row = y * width;
    for (let x = 0; x < width; x += 1) {
      if (data[row + x] >= threshold) continue;
      if (seen % stride === 0 && count < xs.length - 1) {
        xs[count] = x;
        ys[count] = y;
        count += 1;
      }
      seen += 1;
    }
  }
  return { xs, ys, count };
}

/**
 * Score one candidate angle: variance of the row-projection profile after
 * rotating the ink by `angle`. Sharper line separation → higher variance.
 */
function profileVariance(
  xs: Float64Array,
  ys: Float64Array,
  count: number,
  angle: number,
  bins: number,
  binSize: number,
  minY: number,
): number {
  const sin = Math.sin(angle * DEG);
  const cos = Math.cos(angle * DEG);
  const profile = new Float64Array(bins);
  for (let i = 0; i < count; i += 1) {
    // Only the projected row matters, so only y' is computed.
    const projected = ys[i] * cos - xs[i] * sin;
    const bin = Math.floor((projected - minY) / binSize);
    if (bin >= 0 && bin < bins) profile[bin] += 1;
  }
  const mean = count / bins;
  let variance = 0;
  for (let i = 0; i < bins; i += 1) {
    const diff = profile[i] - mean;
    variance += diff * diff;
  }
  return variance / bins;
}

/**
 * Estimate the skew.
 *
 * Coarse sweep over the whole range, then a fine sweep around the winner —
 * text-line spacing makes the score surface smooth enough that this finds the
 * global peak without scoring every hundredth of a degree.
 */
export function estimateSkewAngle(
  image: GrayscaleImage,
  options: EstimateSkewOptions = {},
): SkewEstimate {
  const maxAngle = options.maxAngle ?? 15;
  const coarseStep = options.coarseStep ?? 0.5;
  const fineStep = options.fineStep ?? 0.05;
  const threshold = options.inkThreshold ?? 128;
  const maxSamples = options.maxSamples ?? 60_000;

  if (!(image.width > 1) || !(image.height > 1)) {
    return { angle: 0, confidence: 0 };
  }
  const { xs, ys, count } = inkPixels(image, threshold, maxSamples);
  // Too little ink to have line structure: a blank page, a stray speck. Any
  // angle "wins" on noise like that, which is why it must not count.
  if (count < 64) return { angle: 0, confidence: 0 };

  // Bin height ~ a scanned text line's core at typical sampling; the profile
  // needs several bins per line to see gaps.
  const binSize = Math.max(2, image.height / 200);
  const slack = image.width * Math.sin(maxAngle * DEG);
  const minY = -slack;
  const bins = Math.ceil((image.height + 2 * slack) / binSize);

  const score = (angle: number) =>
    profileVariance(xs, ys, count, angle, bins, binSize, minY);

  let bestAngle = 0;
  let bestScore = -1;
  let scoreSum = 0;
  let scored = 0;
  for (let angle = -maxAngle; angle <= maxAngle + 1e-9; angle += coarseStep) {
    const value = score(angle);
    scoreSum += value;
    scored += 1;
    if (value > bestScore) {
      bestScore = value;
      bestAngle = angle;
    }
  }
  for (
    let angle = bestAngle - coarseStep;
    angle <= bestAngle + coarseStep + 1e-9;
    angle += fineStep
  ) {
    const value = score(angle);
    if (value > bestScore) {
      bestScore = value;
      bestAngle = angle;
    }
  }

  const meanScore = scoreSum / Math.max(1, scored);
  // How decisively the peak beat the average candidate. Lines give ratios in
  // the several-to-many range; structureless ink hovers near 1.
  const ratio = meanScore > 0 ? bestScore / meanScore : 1;
  const confidence = Math.max(0, Math.min(1, (ratio - 1.15) / 3));

  // `projected = y cos θ − x sin θ` IS the rotation by −θ, so the θ that
  // straightens the ink equals the text's own lean — no extra sign flip. The
  // first version of this returned −bestAngle, and the synthetic-page tests
  // caught it as an exactly mirrored estimate.
  return { angle: bestAngle, confidence };
}

/** RGBA (canvas ImageData layout) to the grayscale this module reads. */
export function grayscaleFromRgba(
  data: Uint8ClampedArray | Uint8Array,
  width: number,
  height: number,
): GrayscaleImage {
  const gray = new Uint8Array(width * height);
  for (let i = 0; i < gray.length; i += 1) {
    const at = i * 4;
    // Luma; alpha composited on white so transparent pixels read as paper.
    const alpha = data[at + 3] / 255;
    const luma =
      0.299 * data[at] + 0.587 * data[at + 1] + 0.114 * data[at + 2];
    gray[i] = Math.round(luma * alpha + 255 * (1 - alpha));
  }
  return { data: gray, width, height };
}
