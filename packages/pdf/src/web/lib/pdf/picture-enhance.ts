/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 清晰化 (WPS "sharpen/clarify") for a picture a PDF paints: a levels
 * stretch so the darkest and lightest tones reach black and white, then a
 * mild unsharp mask. The result is written back into the PDF as a native
 * picture replacement, so the change is the pixels only -- position, size,
 * angle and clip stay those of the paint.
 *
 * Pure and DOM-free: it works on RGBA bytes in place (ImageData layout), so
 * it can be tested without a canvas.
 */

export interface PictureEnhanceResult {
  /** The tones were stretched (the picture was flat). */
  readonly stretched: boolean;
  /** Pixels the sharpening changed. */
  readonly sharpenedPixels: number;
}

/** Fraction of pixels ignored at each end when finding the tone range, so a
 * speck of dust does not stand for the picture's black or white. */
const CLIP_FRACTION = 0.005;
/** A range already this wide is left alone: the picture is not flat. */
const FULL_RANGE = 235;
/** The stretch never multiplies contrast more than this (noise grows with
 * it). */
const MAX_GAIN = 2.5;
/** Unsharp mask: how much of the detail is added back, and the smallest
 * local difference treated as detail rather than noise. */
const SHARPEN_AMOUNT = 0.6;
const SHARPEN_THRESHOLD = 3;

/**
 * Enhance `data` (RGBA, `width` x `height`) in place. Transparent pixels
 * keep their alpha and do not count toward the tone range.
 */
export function enhancePicturePixels(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): PictureEnhanceResult {
  const count = width * height;
  if (!(width > 0) || !(height > 0) || data.length < count * 4) {
    return { stretched: false, sharpenedPixels: 0 };
  }
  // 1. Tone range from the luminance histogram of the visible pixels.
  const histogram = new Uint32Array(256);
  let visible = 0;
  for (let index = 0; index < count; index += 1) {
    const offset = index * 4;
    if (data[offset + 3] < 8) continue;
    const luma =
      (data[offset] * 299 + data[offset + 1] * 587 + data[offset + 2] * 114) /
      1000;
    histogram[Math.round(luma)] += 1;
    visible += 1;
  }
  let stretched = false;
  if (visible > 0) {
    const clip = Math.floor(visible * CLIP_FRACTION);
    let low = 0;
    let seen = 0;
    while (low < 255 && seen + histogram[low] <= clip) {
      seen += histogram[low];
      low += 1;
    }
    let high = 255;
    seen = 0;
    while (high > 0 && seen + histogram[high] <= clip) {
      seen += histogram[high];
      high -= 1;
    }
    const range = high - low;
    if (range > 8 && range < FULL_RANGE) {
      const gain = Math.min(MAX_GAIN, 255 / range);
      // Centre the stretched range on the original's mid tone when the gain
      // is capped, so a flat picture brightens or darkens no more than it
      // must.
      const mid = (low + high) / 2;
      const table = new Uint8ClampedArray(256);
      for (let value = 0; value < 256; value += 1) {
        table[value] = Math.round(127.5 + (value - mid) * gain);
      }
      for (let index = 0; index < count; index += 1) {
        const offset = index * 4;
        data[offset] = table[data[offset]];
        data[offset + 1] = table[data[offset + 1]];
        data[offset + 2] = table[data[offset + 2]];
      }
      stretched = true;
    }
  }
  // 2. Unsharp mask with a 3x3 binomial blur ([1 2 1] x [1 2 1] / 16).
  if (width < 3 || height < 3) return { stretched, sharpenedPixels: 0 };
  const source = Uint8ClampedArray.from(data);
  let sharpenedPixels = 0;
  for (let y = 1; y < height - 1; y += 1) {
    for (let x = 1; x < width - 1; x += 1) {
      const offset = (y * width + x) * 4;
      if (source[offset + 3] < 8) continue;
      let changed = false;
      for (let channel = 0; channel < 3; channel += 1) {
        // A transparent neighbour shows the page, not the picture: it
        // counts as the pixel itself, so a cut-out edge is not sharpened.
        const at = (dx: number, dy: number) => {
          const neighbour = ((y + dy) * width + (x + dx)) * 4;
          return source[neighbour + 3] < 8
            ? source[offset + channel]
            : source[neighbour + channel];
        };
        const blurred =
          (at(-1, -1) +
            2 * at(0, -1) +
            at(1, -1) +
            2 * at(-1, 0) +
            4 * at(0, 0) +
            2 * at(1, 0) +
            at(-1, 1) +
            2 * at(0, 1) +
            at(1, 1)) /
          16;
        const value = source[offset + channel];
        const detail = value - blurred;
        if (Math.abs(detail) < SHARPEN_THRESHOLD) continue;
        const next = Math.round(value + SHARPEN_AMOUNT * detail);
        const clamped = next < 0 ? 0 : next > 255 ? 255 : next;
        if (clamped !== value) {
          data[offset + channel] = clamped;
          changed = true;
        }
      }
      if (changed) sharpenedPixels += 1;
    }
  }
  return { stretched, sharpenedPixels };
}

/**
 * Office's picture corrections (Brightness / Contrast presets, the shared
 * 校正 gallery the deck uses) baked into the pixels, exactly as the gallery
 * previews them: CSS `brightness(b%)` then `contrast(c%)`, with `b` and `c`
 * Office's -100..100 deltas from 100%. Alpha is kept. Returns whether any
 * pixel changed.
 */
export function adjustPicturePixels(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  deltas: { readonly brightness: number; readonly contrast: number },
): boolean {
  const count = width * height;
  if (!(width > 0) || !(height > 0) || data.length < count * 4) return false;
  if (deltas.brightness === 0 && deltas.contrast === 0) return false;
  const brightness = Math.max(0, (100 + deltas.brightness) / 100);
  const contrast = Math.max(0, (100 + deltas.contrast) / 100);
  const table = new Uint8ClampedArray(256);
  for (let value = 0; value < 256; value += 1) {
    const lit = (value / 255) * brightness;
    table[value] = Math.round(((lit - 0.5) * contrast + 0.5) * 255);
  }
  let changed = false;
  for (let index = 0; index < count; index += 1) {
    const offset = index * 4;
    for (let channel = 0; channel < 3; channel += 1) {
      const next = table[data[offset + channel]];
      if (next !== data[offset + channel]) {
        data[offset + channel] = next;
        changed = true;
      }
    }
  }
  return changed;
}
