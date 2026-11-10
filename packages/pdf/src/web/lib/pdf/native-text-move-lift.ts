/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { cropPixelRect } from "./canvas-crop";
import type { FractionRect } from "./embedded-image-edits";

const INK_THRESHOLD = 24;

function channelDifference(a: Uint8ClampedArray, i: number, b: Uint8ClampedArray, j: number): number {
  return Math.max(
    Math.abs(a[i] - b[j]),
    Math.abs(a[i + 1] - b[j + 1]),
    Math.abs(a[i + 2] - b[j + 2]),
    Math.abs(a[i + 3] - b[j + 3]),
  );
}

/**
 * Keeps the pixels of `page` that `clean` has not (alpha 0 elsewhere) and
 * answers how many it kept. With `width`, a pixel is the block's only when
 * it differs from the clean pixel at its place and at each of its eight
 * neighbours: a neighbour's edge drawn a pixel apart is not.
 */
export function keepNativeTextBlockInk(
  page: Uint8ClampedArray,
  clean: Uint8ClampedArray,
  threshold = INK_THRESHOLD,
  width?: number,
): number {
  let kept = 0;
  const length = Math.min(page.length, clean.length) - (Math.min(page.length, clean.length) % 4);
  const columns = width && width > 0 ? width : 0;
  const rows = columns ? Math.floor(length / 4 / columns) : 0;
  for (let i = 0; i < length; i += 4) {
    let difference = channelDifference(page, i, clean, i);
    if (columns && difference >= threshold) {
      const pixel = i / 4, x = pixel % columns, y = Math.floor(pixel / columns);
      for (let dy = -1; dy <= 1 && difference >= threshold; dy++) {
        const ny = y + dy;
        if (ny < 0 || ny >= rows) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx;
          if ((dx === 0 && dy === 0) || nx < 0 || nx >= columns) continue;
          const near = channelDifference(page, i, clean, (ny * columns + nx) * 4);
          if (near < difference) difference = near;
          if (difference < threshold) break;
        }
      }
    }
    if (difference < threshold) page[i + 3] = 0;
    else kept += 1;
  }
  for (let i = length + 3; i < page.length; i += 4) page[i] = 0;
  return kept;
}

/**
 * The block's ink against its own background: the most frequent colour on
 * the crop's edge. Keeps the pixels that differ from it (alpha 0 elsewhere)
 * and answers the background and the kept pixels' mask.
 */
export function keepNativeTextBlockInkOverBackground(
  page: Uint8ClampedArray,
  width: number,
  height: number,
  threshold = INK_THRESHOLD,
): { kept: number; background: readonly [number, number, number, number]; mask: Uint8Array } | null {
  if (!(width > 0) || !(height > 0) || page.length < width * height * 4) return null;
  const counts = new Map<number, number>();
  const tally = (x: number, y: number) => {
    const i = (y * width + x) * 4;
    // Coarse buckets: an antialiased edge pixel does not split the vote.
    const key = ((page[i] >> 3) << 15) | ((page[i + 1] >> 3) << 10) | ((page[i + 2] >> 3) << 5) | (page[i + 3] >> 3);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  };
  for (let x = 0; x < width; x++) { tally(x, 0); tally(x, height - 1); }
  for (let y = 1; y < height - 1; y++) { tally(0, y); tally(width - 1, y); }
  let best = -1, bestCount = -1;
  for (const [key, count] of counts) if (count > bestCount) { best = key; bestCount = count; }
  if (best < 0) return null;
  // The exact colour: the mean of the edge pixels in the winning bucket.
  let r = 0, g = 0, b = 0, a = 0, n = 0;
  const average = (x: number, y: number) => {
    const i = (y * width + x) * 4;
    const key = ((page[i] >> 3) << 15) | ((page[i + 1] >> 3) << 10) | ((page[i + 2] >> 3) << 5) | (page[i + 3] >> 3);
    if (key !== best) return;
    r += page[i]; g += page[i + 1]; b += page[i + 2]; a += page[i + 3]; n += 1;
  };
  for (let x = 0; x < width; x++) { average(x, 0); average(x, height - 1); }
  for (let y = 1; y < height - 1; y++) { average(0, y); average(width - 1, y); }
  const background = [Math.round(r / n), Math.round(g / n), Math.round(b / n), Math.round(a / n)] as const;
  const mask = new Uint8Array(width * height);
  let kept = 0;
  for (let p = 0, i = 0; p < width * height; p++, i += 4) {
    const difference = Math.max(
      Math.abs(page[i] - background[0]), Math.abs(page[i + 1] - background[1]),
      Math.abs(page[i + 2] - background[2]), Math.abs(page[i + 3] - background[3]),
    );
    if (difference < threshold) page[i + 3] = 0;
    else { mask[p] = 1; kept += 1; }
  }
  return { kept, background, mask };
}

/** A pixel rectangle inside the crop: [left, top, right, bottom). */
export type PdfNativeTextLiftInner = readonly [number, number, number, number];

const SPECK_PIXELS = 4;
const SPECK_RESIDUAL = 40;

/**
 * Drops the kept pixels (alpha > 0) that are not the block's: a group of
 * them (eight-connected) none of whose pixels lies in `inner` (the frame
 * itself, in crop pixels), and, with `clean`, a group of at most four
 * pixels that is no blend of the clean pixel and the block's ink colour
 * (the mean of its strongly kept pixels). Answers how many pixels stay.
 */
export function dropStrayNativeTextInk(
  page: Uint8ClampedArray,
  width: number,
  height: number,
  inner: PdfNativeTextLiftInner,
  clean?: Uint8ClampedArray,
): number {
  const count = width * height;
  if (!(width > 0) || !(height > 0) || page.length < count * 4) return 0;
  // The block's ink: the mean of the pixels kept far from the clean surface.
  let ink: [number, number, number] | null = null;
  if (clean) {
    let r = 0, g = 0, b = 0, n = 0;
    for (let p = 0, i = 0; p < count; p++, i += 4) {
      if (page[i + 3] === 0 || channelDifference(page, i, clean, i) < 160) continue;
      r += page[i]; g += page[i + 1]; b += page[i + 2]; n += 1;
    }
    if (n > 0) ink = [r / n, g / n, b / n];
  }
  const [left, top, right, bottom] = inner;
  const label = new Int32Array(count).fill(-1);
  const stack: number[] = [];
  const members: number[] = [];
  let kept = 0;
  for (let start = 0; start < count; start++) {
    if (label[start] !== -1 || page[start * 4 + 3] === 0) continue;
    label[start] = start;
    stack.length = 0;
    members.length = 0;
    stack.push(start);
    let touchesInner = false;
    while (stack.length) {
      const pixel = stack.pop()!;
      members.push(pixel);
      const x = pixel % width, y = (pixel - x) / width;
      if (x >= left && x < right && y >= top && y < bottom) touchesInner = true;
      for (let dy = -1; dy <= 1; dy++) {
        const ny = y + dy;
        if (ny < 0 || ny >= height) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx;
          if (nx < 0 || nx >= width) continue;
          const next = ny * width + nx;
          if (label[next] !== -1 || page[next * 4 + 3] === 0) continue;
          label[next] = start;
          stack.push(next);
        }
      }
    }
    let keep = touchesInner;
    if (keep && clean && ink && members.length <= SPECK_PIXELS) {
      // Each pixel must read as the clean pixel blended toward the ink.
      for (const pixel of members) {
        const i = pixel * 4;
        const toward = [ink[0] - clean[i], ink[1] - clean[i + 1], ink[2] - clean[i + 2]];
        const length = toward[0] ** 2 + toward[1] ** 2 + toward[2] ** 2;
        const offset = [page[i] - clean[i], page[i + 1] - clean[i + 1], page[i + 2] - clean[i + 2]];
        const amount = length > 0
          ? Math.min(1, Math.max(0, (offset[0] * toward[0] + offset[1] * toward[1] + offset[2] * toward[2]) / length))
          : 0;
        const residual = Math.max(
          Math.abs(offset[0] - amount * toward[0]),
          Math.abs(offset[1] - amount * toward[1]),
          Math.abs(offset[2] - amount * toward[2]),
        );
        // ...and be drawn toward it as far as ink is told from paper (a
        // shape's edge a few levels off the clean pixel is not).
        const along = amount * Math.max(Math.abs(toward[0]), Math.abs(toward[1]), Math.abs(toward[2]));
        if (residual > SPECK_RESIDUAL || along < INK_THRESHOLD) { keep = false; break; }
      }
    }
    if (keep) kept += members.length;
    else for (const pixel of members) page[pixel * 4 + 3] = 0;
  }
  return kept;
}

export type PdfNativeTextMoveLift = Readonly<{
  /** The block's ink, a transparent PNG. */
  url: string;
  /** Where that picture lies on the page, in page fractions. */
  rect: FractionRect;
  /** Cut against the clean surface (true), or against the block's own
   * background before that surface was drawn. */
  exact: boolean;
  /** Before the clean surface: the block's background over its old place. */
  cover?: Readonly<{ url: string; rect: FractionRect }>;
}>;

/** Clears the pixels of a box's own fill (`#rrggbb`): a box over a run that
 * paints nothing is drawn on the page's colour (PDF 0107), which is not ink
 * it carries -- against a clean surface showing the scanned letters under
 * it, the fill read as light letters and travelled with the moved box.
 * Answers how many pixels stay. */
function dropBoxFillPixels(data: Uint8ClampedArray, fill: string | undefined): number {
  const match = fill ? /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(fill) : null;
  let kept = 0;
  const [r, g, b] = match ? match.slice(1).map((hex) => Number.parseInt(hex, 16)) : [0, 0, 0];
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] === 0) continue;
    if (match && Math.abs(data[i] - r) <= INK_THRESHOLD / 2 && Math.abs(data[i + 1] - g) <= INK_THRESHOLD / 2 &&
      Math.abs(data[i + 2] - b) <= INK_THRESHOLD / 2) data[i + 3] = 0;
    else kept += 1;
  }
  return kept;
}

const hexRgb = (value: string | undefined): [number, number, number] | null => {
  const match = value ? /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(value) : null;
  return match ? [Number.parseInt(match[1], 16), Number.parseInt(match[2], 16), Number.parseInt(match[3], 16)] : null;
};

/** The text of a box drawn on an opaque fill, cut against that fill: each
 * pixel becomes the ink at the share of it the pixel holds (P = a·ink +
 * (1-a)·fill), so the glyphs and their antialiased edges come away clean
 * and nothing of the fill does. A box over a run that paints nothing (PDF
 * 0107) needs it: the clean surface under it shows the scanned letters of
 * the same words, so a cut against that surface left the glyphs hollow.
 * Answers how many pixels stay. */
function keyBoxTextAgainstFill(data: Uint8ClampedArray, fill: string, ink: string): number {
  const f = hexRgb(fill), k = hexRgb(ink);
  if (!f || !k) return 0;
  const toward = [k[0] - f[0], k[1] - f[1], k[2] - f[2]];
  const length = toward[0] ** 2 + toward[1] ** 2 + toward[2] ** 2;
  if (!(length > 0)) return 0;
  let kept = 0;
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] === 0) continue;
    const share = Math.min(1, Math.max(0, ((data[i] - f[0]) * toward[0] + (data[i + 1] - f[1]) * toward[1] +
      (data[i + 2] - f[2]) * toward[2]) / length));
    const alpha = Math.round(share * data[i + 3]);
    data[i] = k[0]; data[i + 1] = k[1]; data[i + 2] = k[2];
    data[i + 3] = alpha < 8 ? 0 : alpha;
    if (data[i + 3]) kept += 1;
  }
  return kept;
}

export function liftNativeTextBlockInk(
  page: HTMLCanvasElement | null | undefined,
  clean: HTMLCanvasElement | null | undefined,
  bounds: FractionRect,
  options: {
    createCanvas?: () => HTMLCanvasElement;
    frame?: FractionRect;
    /** The box's own fill (`#rrggbb`), never carried as its ink. */
    fill?: string;
    /** The frame holds the box drawn on that fill: its text is cut against
     * the fill (keyBoxTextAgainstFill), in this ink colour. */
    fillKeyInk?: string;
    /** The block's glyph boxes (page fractions): against its background
     * alone, only ink inside them is the block's (a neighbour overlapping
     * the frame is not). */
    glyphs?: readonly FractionRect[];
  } = {},
): PdfNativeTextMoveLift | null {
  if (!page) return null;
  const exact = Boolean(clean && clean.width > 0 && clean.height > 0);
  const crop = cropPixelRect(page, bounds);
  if (!crop || !(crop.width > 0) || !(crop.height > 0)) return null;
  const create = options.createCanvas ?? (() => document.createElement("canvas"));
  const ink = create();
  const surface = create();
  try {
    ink.width = surface.width = crop.width;
    ink.height = surface.height = crop.height;
    const inkContext = ink.getContext("2d", { willReadFrequently: true });
    const surfaceContext = surface.getContext("2d", { willReadFrequently: true });
    if (!inkContext || !surfaceContext) return null;
    inkContext.drawImage(page, crop.x, crop.y, crop.width, crop.height, 0, 0, crop.width, crop.height);
    const pixels = inkContext.getImageData(0, 0, crop.width, crop.height);
    const rect = Object.freeze({
      x: crop.x / page.width,
      y: crop.y / page.height,
      width: crop.width / page.width,
      height: crop.height / page.height,
    });
    // The frame itself in crop pixels (the whole crop without one).
    const frame = options.frame;
    const inner: PdfNativeTextLiftInner = frame
      ? [
          Math.max(0, Math.floor(frame.x * page.width - crop.x)),
          Math.max(0, Math.floor(frame.y * page.height - crop.y)),
          Math.min(crop.width, Math.ceil((frame.x + frame.width) * page.width - crop.x)),
          Math.min(crop.height, Math.ceil((frame.y + frame.height) * page.height - crop.y)),
        ]
      : [0, 0, crop.width, crop.height];
    if (options.fill && options.fillKeyInk) {
      // Only the frame's own pixels: what is past it is not the box's.
      for (let y = 0; y < crop.height; y++) for (let x = 0; x < crop.width; x++) {
        if (x < inner[0] || x >= inner[2] || y < inner[1] || y >= inner[3]) pixels.data[(y * crop.width + x) * 4 + 3] = 0;
      }
      if (keyBoxTextAgainstFill(pixels.data, options.fill, options.fillKeyInk) === 0) return null;
      inkContext.putImageData(pixels, 0, 0);
      return Object.freeze({ url: ink.toDataURL("image/png"), rect, exact });
    }
    if (exact) {
      // The same region of the clean surface, at the page canvas's resolution.
      const scaleX = clean!.width / page.width;
      const scaleY = clean!.height / page.height;
      surfaceContext.drawImage(
        clean!,
        crop.x * scaleX, crop.y * scaleY, crop.width * scaleX, crop.height * scaleY,
        0, 0, crop.width, crop.height,
      );
      const cleanPixels = surfaceContext.getImageData(0, 0, crop.width, crop.height);
      if (keepNativeTextBlockInk(pixels.data, cleanPixels.data, INK_THRESHOLD, crop.width) === 0) return null;
      if (options.fill && dropBoxFillPixels(pixels.data, options.fill) === 0) return null;
      if (dropStrayNativeTextInk(pixels.data, crop.width, crop.height, inner, cleanPixels.data) === 0) return null;
      inkContext.putImageData(pixels, 0, 0);
      return Object.freeze({ url: ink.toDataURL("image/png"), rect, exact: true });
    }
    const keyed = keepNativeTextBlockInkOverBackground(pixels.data, crop.width, crop.height);
    if (!keyed || keyed.kept === 0) return null;
    if (options.glyphs?.length) {
      const inside = new Uint8Array(crop.width * crop.height);
      for (const glyph of options.glyphs) {
        const left = Math.max(0, Math.floor(glyph.x * page.width - crop.x));
        const top = Math.max(0, Math.floor(glyph.y * page.height - crop.y));
        const right = Math.min(crop.width, Math.ceil((glyph.x + glyph.width) * page.width - crop.x));
        const bottom = Math.min(crop.height, Math.ceil((glyph.y + glyph.height) * page.height - crop.y));
        for (let y = top; y < bottom; y++) inside.fill(1, y * crop.width + left, Math.max(y * crop.width + left, y * crop.width + right));
      }
      for (let p = 0; p < inside.length; p++) if (!inside[p]) pixels.data[p * 4 + 3] = 0;
    }
    if (dropStrayNativeTextInk(pixels.data, crop.width, crop.height, inner) === 0) return null;
    for (let p = 0; p < keyed.mask.length; p++) if (pixels.data[p * 4 + 3] === 0) keyed.mask[p] = 0;
    inkContext.putImageData(pixels, 0, 0);
    // The cover: the background over the ink and a pixel round it (its
    // antialiased fringe), transparent elsewhere.
    const cover = surfaceContext.createImageData(crop.width, crop.height);
    const [r, g, b, a] = keyed.background;
    for (let y = 0; y < crop.height; y++) {
      for (let x = 0; x < crop.width; x++) {
        let near = false;
        for (let dy = -1; dy <= 1 && !near; dy++) {
          const ny = y + dy;
          if (ny < 0 || ny >= crop.height) continue;
          for (let dx = -1; dx <= 1; dx++) {
            const nx = x + dx;
            if (nx >= 0 && nx < crop.width && keyed.mask[ny * crop.width + nx]) { near = true; break; }
          }
        }
        if (!near) continue;
        const i = (y * crop.width + x) * 4;
        cover.data[i] = r; cover.data[i + 1] = g; cover.data[i + 2] = b; cover.data[i + 3] = Math.max(a, 255);
      }
    }
    surfaceContext.putImageData(cover, 0, 0);
    return Object.freeze({
      url: ink.toDataURL("image/png"),
      rect,
      exact: false,
      cover: Object.freeze({ url: surface.toDataURL("image/png"), rect }),
    });
  } catch {
    // A canvas that cannot be read (tainted, lost) lifts nothing.
    return null;
  } finally {
    ink.width = ink.height = 0;
    surface.width = surface.height = 0;
  }
}

/** The lift cut from the block's own layer (a transparent canvas the size of
 * the page canvas holding the block alone): `bounds` of it, as drawn. Null
 * when the layer draws nothing there. */
export function liftNativeTextBlockLayer(
  layer: HTMLCanvasElement | null | undefined,
  bounds: FractionRect,
  options: { createCanvas?: () => HTMLCanvasElement;
    /** The box's own fill (`#rrggbb`), never carried as its ink. */
    fill?: string;
    /** The layer draws the box on that fill: its text is cut against it. */
    fillKeyInk?: string } = {},
): PdfNativeTextMoveLift | null {
  if (!layer || !(layer.width > 0) || !(layer.height > 0)) return null;
  const crop = cropPixelRect(layer, bounds);
  if (!crop || !(crop.width > 0) || !(crop.height > 0)) return null;
  const create = options.createCanvas ?? (() => document.createElement("canvas"));
  const ink = create();
  try {
    ink.width = crop.width;
    ink.height = crop.height;
    const context = ink.getContext("2d", { willReadFrequently: true });
    if (!context) return null;
    context.drawImage(layer, crop.x, crop.y, crop.width, crop.height, 0, 0, crop.width, crop.height);
    const image = context.getImageData(0, 0, crop.width, crop.height);
    const alpha = image.data;
    if (options.fill) {
      if ((options.fillKeyInk ? keyBoxTextAgainstFill(alpha, options.fill, options.fillKeyInk)
        : dropBoxFillPixels(alpha, options.fill)) === 0) return null;
      context.putImageData(image, 0, 0);
    }
    let drawn = false;
    for (let i = 3; i < alpha.length; i += 4) if (alpha[i] > 0) { drawn = true; break; }
    if (!drawn) return null;
    return Object.freeze({
      url: ink.toDataURL("image/png"),
      rect: Object.freeze({
        x: crop.x / layer.width,
        y: crop.y / layer.height,
        width: crop.width / layer.width,
        height: crop.height / layer.height,
      }),
      exact: true,
    });
  } catch {
    return null;
  } finally {
    ink.width = ink.height = 0;
  }
}

