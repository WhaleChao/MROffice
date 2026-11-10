/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isSafeCssColor } from "@mrnote/core/web/lib/pdf/safe-color.ts";
import { pdfSourceTextBoxCoversInk } from "./source-text-box";
import { coverRectForEdit } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";
import type { PdfTextStyleRun } from "@mrnote/core/web/lib/pdf/text-style-runs.ts";

export const DEFAULT_TEXTBOX_BACKGROUND = "#ffffff";













export function textBoxOverlayFill(
  edit:
    | Readonly<{
        sourceText?: Parameters<typeof pdfSourceTextBoxCoversInk>[0];
        backgroundColor?: unknown;
        nativeSource?: unknown;
      }>
    | null
    | undefined,
): string {
  if (!edit) return DEFAULT_TEXTBOX_BACKGROUND;
  if (edit.sourceText) {
    return pdfSourceTextBoxCoversInk(edit.sourceText) && isSafeCssColor(edit.backgroundColor)
      ? opaqueTextBoxCoverColor(edit.backgroundColor)
      : "transparent";
  }
  if (isSafeCssColor(edit.backgroundColor)) return opaqueTextBoxCoverColor(edit.backgroundColor);
  // A box given no fill has none, as a WPS text box has none until one is
  // chosen: drawn on white, a box made of a list item (its measured height
  // past the line's) hid the two items under it, in the preview and after
  
  return "transparent";
}

/**
 * Force a persisted cover colour to be OPAQUE.
 *
 * The cover's one job is to hide the original glyphs; a translucent value
 * defeats it silently — the box renders its own text while the canvas text
 * stays visible underneath, which reads as doubled text. `isSafeCssColor`
 * deliberately admits translucent `rgba(...)` for the highlight palette, so
 * safety-validation alone does not guarantee opacity. Alpha is composited
 * over white (the export path drops alpha the same way), and anything
 * unparseable falls back to the default.
 */
export function opaqueTextBoxCoverColor(value: unknown): string {
  if (typeof value !== "string") return DEFAULT_TEXTBOX_BACKGROUND;
  const raw = value.trim();
  const hexMatch = /^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$/i.exec(raw);
  if (hexMatch) {
    let hex = hexMatch[1];
    if (hex.length === 3) {
      hex = hex
        .split("")
        .map((c) => c + c)
        .join("");
    }
    if (hex.length === 6) return `#${hex.toLowerCase()}`;
    // #rrggbbaa: composite over white.
    const r = Number.parseInt(hex.slice(0, 2), 16);
    const g = Number.parseInt(hex.slice(2, 4), 16);
    const b = Number.parseInt(hex.slice(4, 6), 16);
    const a = Number.parseInt(hex.slice(6, 8), 16) / 255;
    return `#${hexByte(r * a + 255 * (1 - a))}${hexByte(g * a + 255 * (1 - a))}${hexByte(
      b * a + 255 * (1 - a),
    )}`;
  }
  const rgbMatch =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([0-9.]+)\s*)?\)$/i.exec(
      raw,
    );
  if (rgbMatch) {
    const r = Math.min(255, Number(rgbMatch[1]));
    const g = Math.min(255, Number(rgbMatch[2]));
    const b = Math.min(255, Number(rgbMatch[3]));
    const alphaRaw = rgbMatch[4] === undefined ? 1 : Number(rgbMatch[4]);
    const a = Number.isFinite(alphaRaw) ? Math.max(0, Math.min(1, alphaRaw)) : 1;
    return `#${hexByte(r * a + 255 * (1 - a))}${hexByte(g * a + 255 * (1 - a))}${hexByte(
      b * a + 255 * (1 - a),
    )}`;
  }
  return DEFAULT_TEXTBOX_BACKGROUND;
}

type TextBoxBackgroundSample = {
  x: number;
  y: number;
  width: number;
  height: number;
};

type RgbSample = [number, number, number];

function clampPixel(value: number, maxExclusive: number): number {
  if (!Number.isFinite(value) || maxExclusive <= 0) return 0;
  return Math.min(maxExclusive - 1, Math.max(0, Math.round(value)));
}

function hexByte(value: number): string {
  return Math.min(255, Math.max(0, Math.round(value)))
    .toString(16)
    .padStart(2, "0");
}

function medianChannel(
  samples: Array<readonly [number, number, number]>,
  index: 0 | 1 | 2,
): number {
  const sorted = samples.map((sample) => sample[index]).sort((a, b) => a - b);
  return sorted[Math.floor(sorted.length / 2)] ?? 255;
}

function colorDistance(
  a: readonly [number, number, number],
  b: readonly [number, number, number],
): number {
  return Math.hypot(a[0] - b[0], a[1] - b[1], a[2] - b[2]);
}

function compositePixelOnWhite(
  data: Uint8ClampedArray,
  offset: number,
): RgbSample {
  const opacity = data[offset + 3] / 255;
  return [
    data[offset] * opacity + 255 * (1 - opacity),
    data[offset + 1] * opacity + 255 * (1 - opacity),
    data[offset + 2] * opacity + 255 * (1 - opacity),
  ];
}

/**
 * Return the dominant anti-aliasing colour cluster, not the median of a few
 * arbitrary pixels. Tight PDF glyph boxes often put strokes at every old
 * probe point; treating those strokes as the page background inverted the
 * persisted pair (white text + near-black mask) after the first keystroke.
 */
function dominantColor(
  samples: readonly RgbSample[],
): { color: RgbSample; share: number } | null {
  if (samples.length === 0) return null;
  const buckets = new Map<string, RgbSample[]>();
  for (const sample of samples) {
    // 16-level buckets keep antialiased edges together without collapsing
    // genuinely different page/text colours.
    const key = `${Math.round(sample[0] / 16)}:${Math.round(
      sample[1] / 16,
    )}:${Math.round(sample[2] / 16)}`;
    const bucket = buckets.get(key);
    if (bucket) bucket.push(sample);
    else buckets.set(key, [sample]);
  }
  const winner = Array.from(buckets.values()).sort(
    (a, b) => b.length - a.length,
  )[0];
  if (!winner) return null;
  return {
    color: [
      medianChannel(winner, 0),
      medianChannel(winner, 1),
      medianChannel(winner, 2),
    ],
    share: winner.length / samples.length,
  };
}

function outsideRingSamples(
  ctx: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  rect: TextBoxBackgroundSample,
): RgbSample[] {
  const left = rect.x * canvas.width;
  const top = rect.y * canvas.height;
  const right = (rect.x + rect.width) * canvas.width;
  const bottom = (rect.y + rect.height) * canvas.height;
  const samples: RgbSample[] = [];
  const horizontalFractions = [0.02, 0.12, 0.28, 0.5, 0.72, 0.88, 0.98];
  const verticalFractions = [0.15, 0.35, 0.5, 0.65, 0.85];
  const regionLeft = Math.max(0, Math.floor(Math.min(left, right)) - 2);
  const regionTop = Math.max(0, Math.floor(Math.min(top, bottom)) - 2);
  const regionRight = Math.min(
    canvas.width - 1,
    Math.ceil(Math.max(left, right)) + 2,
  );
  const regionBottom = Math.min(
    canvas.height - 1,
    Math.ceil(Math.max(top, bottom)) + 2,
  );
  const regionWidth = Math.max(1, regionRight - regionLeft + 1);
  const regionHeight = Math.max(1, regionBottom - regionTop + 1);
  const imageData = ctx.getImageData(
    regionLeft,
    regionTop,
    regionWidth,
    regionHeight,
  );
  const sampleAt = (x: number, y: number): RgbSample | null => {
    const px = Math.round(x);
    const py = Math.round(y);
    if (
      px < regionLeft ||
      py < regionTop ||
      px > regionRight ||
      py > regionBottom
    ) {
      return null;
    }
    const offset = ((py - regionTop) * regionWidth + (px - regionLeft)) * 4;
    return compositePixelOnWhite(imageData.data, offset);
  };

  // Probe immediately OUTSIDE the tight glyph bounds. On an ordinary page
  // this is the paper/cell fill; on a coloured block the fill normally extends
  // past the glyphs too. Two rings tolerate a one-pixel antialias fringe.
  for (const gap of [1, 2]) {
    for (const fraction of horizontalFractions) {
      const x = left + (right - left) * fraction;
      for (const y of [top - gap, bottom + gap]) {
        const sample = sampleAt(x, y);
        if (sample) samples.push(sample);
      }
    }
    for (const fraction of verticalFractions) {
      const y = top + (bottom - top) * fraction;
      for (const x of [left - gap, right + gap]) {
        const sample = sampleAt(x, y);
        if (sample) samples.push(sample);
      }
    }
  }
  return samples;
}

function interiorSamples(
  ctx: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  rect: TextBoxBackgroundSample,
): RgbSample[] {
  const left = clampPixel(rect.x * canvas.width, canvas.width);
  const top = clampPixel(rect.y * canvas.height, canvas.height);
  const right = clampPixel(
    (rect.x + rect.width) * canvas.width,
    canvas.width,
  );
  const bottom = clampPixel(
    (rect.y + rect.height) * canvas.height,
    canvas.height,
  );
  const width = Math.max(1, right - left + 1);
  const height = Math.max(1, bottom - top + 1);
  const imageData = ctx.getImageData(left, top, width, height);
  const maxSamples = 1600;
  const stride = Math.max(
    1,
    Math.floor(Math.sqrt((width * height) / maxSamples)),
  );
  const samples: RgbSample[] = [];
  for (let y = 0; y < height; y += stride) {
    for (let x = 0; x < width; x += stride) {
      samples.push(compositePixelOnWhite(imageData.data, (y * width + x) * 4));
    }
  }
  return samples;
}

export function sampleTextBoxBackgroundColor(
  canvas: HTMLCanvasElement | null,
  edit: TextBoxBackgroundSample,
): string | undefined {
  if (!canvas || canvas.width <= 0 || canvas.height <= 0) return undefined;
  const ctx = canvas.getContext("2d");
  if (!ctx) return undefined;

  const rect = coverRectForEdit(edit);
  try {
    const ring = dominantColor(outsideRingSamples(ctx, canvas, rect));
    // A uniform nearby ring is the strongest background signal because it
    // cannot be a glyph stroke. If page edges/noisy imagery make it ambiguous,
    // fall back to the dominant colour over the whole source box.
    const chosen =
      ring && ring.share >= 0.55
        ? ring
        : dominantColor(interiorSamples(ctx, canvas, rect));
    if (!chosen) return undefined;
    const [r, g, b] = chosen.color;
    return `#${hexByte(r)}${hexByte(g)}${hexByte(b)}`;
  } catch {
    return undefined;
  }
}

export function sampleTextBoxTextColor(
  canvas: HTMLCanvasElement | null,
  edit: TextBoxBackgroundSample,
): string | undefined {
  if (!canvas || canvas.width <= 0 || canvas.height <= 0) return undefined;
  const ctx = canvas.getContext("2d");
  if (!ctx) return undefined;

  const rect = coverRectForEdit(edit);
  const left = clampPixel(rect.x * canvas.width, canvas.width);
  const top = clampPixel(rect.y * canvas.height, canvas.height);
  const right = clampPixel((rect.x + rect.width) * canvas.width, canvas.width);
  const bottom = clampPixel(
    (rect.y + rect.height) * canvas.height,
    canvas.height,
  );
  const width = Math.max(1, right - left + 1);
  const height = Math.max(1, bottom - top + 1);

  let imageData: ImageData;
  try {
    imageData = ctx.getImageData(left, top, width, height);
  } catch {
    return undefined;
  }

  const backgroundHex = sampleTextBoxBackgroundColor(canvas, edit);
  const background: [number, number, number] = backgroundHex
    ? [
        Number.parseInt(backgroundHex.slice(1, 3), 16),
        Number.parseInt(backgroundHex.slice(3, 5), 16),
        Number.parseInt(backgroundHex.slice(5, 7), 16),
      ]
    : [255, 255, 255];
  const samples: Array<[number, number, number]> = [];
  const maxSamples = 1200;
  const stride = Math.max(
    1,
    Math.floor(Math.sqrt((width * height) / maxSamples)),
  );

  for (let y = 0; y < height; y += stride) {
    for (let x = 0; x < width; x += stride) {
      const offset = (y * width + x) * 4;
      const alpha = imageData.data[offset + 3] / 255;
      if (alpha < 0.25) continue;
      const rgb: [number, number, number] = [
        imageData.data[offset] * alpha + 255 * (1 - alpha),
        imageData.data[offset + 1] * alpha + 255 * (1 - alpha),
        imageData.data[offset + 2] * alpha + 255 * (1 - alpha),
      ];
      if (colorDistance(rgb, background) < 42) continue;
      samples.push(rgb);
    }
  }

  // The glyph cores, not their antialiased edges: small text has more edge
  // pixels than core ones, and the most frequent ink of a black line was a
  
  // black line). The third of the samples farthest from the background
  // carries the colour the text is painted in.
  const cores = [...samples]
    .sort((left, right) => colorDistance(right, background) - colorDistance(left, background))
    .slice(0, Math.max(1, Math.ceil(samples.length / 3)));
  const foreground = dominantColor(cores);
  if (!foreground) return undefined;
  const [r, g, b] = foreground.color;
  return `#${hexByte(r)}${hexByte(g)}${hexByte(b)}`;
}

/**
 * Sample each extracted glyph run independently so a red/black (or otherwise
 * multicolour) paragraph does not collapse to the box's dominant colour after
 * editing. Run sample geometry is live-only and is removed by
 * `persistablePdfTextStyleRuns`; only the safe sampled hex survives in JSON.
 */
export function samplePdfTextRunColors(
  canvas: HTMLCanvasElement | null,
  runs: readonly PdfTextStyleRun[] | undefined,
  fallbackColor?: string,
): PdfTextStyleRun[] | undefined {
  if (!runs) return undefined;
  if (runs.length === 1) {
    const [run] = runs;
    const color = run.color ?? fallbackColor;
    return [color ? { ...run, color } : { ...run }];
  }
  return runs.map((run) => {
    const hasSampleRect =
      Number.isFinite(run.sampleX) &&
      Number.isFinite(run.sampleY) &&
      Number.isFinite(run.sampleWidth) &&
      Number.isFinite(run.sampleHeight) &&
      (run.sampleWidth as number) > 0 &&
      (run.sampleHeight as number) > 0;
    const sampled = hasSampleRect
      ? sampleTextBoxTextColor(canvas, {
          x: run.sampleX as number,
          y: run.sampleY as number,
          width: run.sampleWidth as number,
          height: run.sampleHeight as number,
        })
      : undefined;
    const color = sampled ?? run.color ?? fallbackColor;
    return color ? { ...run, color } : { ...run };
  });
}
