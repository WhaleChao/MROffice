/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Cutting a region out of the rendered page canvas.
 *
 * Moving an embedded image needs the image's own pixels: the mask covers where
 * it was, and something has to be drawn where it went. Pulling the original
 * XObject out of the file is possible but drags in every colour space, decode
 * array, SMask and JPX case pdf.js already handles once — for a page that has
 * just been rendered, the pixels are sitting on the canvas.
 *
 * What that gives is the image AS COMPOSITED: anything drawn over it comes
 * along, and anything drawn under it shows through where it is transparent.
 * For a figure that is what a reader would expect a move to carry. It is also
 * resolution-limited by the render scale, which is why the crop is taken from
 * the canvas at its device size rather than its CSS size.
 */

import type { FractionRect } from "./embedded-image-edits";

export interface CanvasSize {
  width: number;
  height: number;
}

export interface PixelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * A value within this many pixels of a whole one IS that whole one.
 *
 * Fractions do not survive multiplication exactly: `(0.1 + 0.2) * 1000` is
 * 300.00000000000006, and rounding that outward hands back a pixel of
 * surrounding page on every ordinary rectangle. The tolerance is far below any
 * real sub-pixel edge, so genuine overhang still rounds outward.
 */
const EDGE_EPSILON = 1e-6;

const floorEdge = (value: number) => Math.floor(value + EDGE_EPSILON);
const ceilEdge = (value: number) => Math.ceil(value - EDGE_EPSILON);

/**
 * The device-pixel rectangle a page-fraction rect covers on `canvas`.
 *
 * Rounded OUTWARD: rounding to nearest would shave a fraction of a pixel off
 * each edge, and the edge of a figure is exactly where losing a pixel shows.
 * Returns null when nothing of the rect is on the canvas, so the caller reports
 * "nothing to move" rather than drawing a 0x0 image and calling it done.
 */
export function cropPixelRect(
  canvas: CanvasSize,
  rect: FractionRect,
  options: { padding?: number } = {},
): PixelRect | null {
  if (!(canvas.width > 0) || !(canvas.height > 0)) return null;
  if (![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)) {
    return null;
  }
  if (!(rect.width > 0) || !(rect.height > 0)) return null;

  const pad = options.padding ?? 0;
  const left = floorEdge((rect.x - pad) * canvas.width);
  const top = floorEdge((rect.y - pad) * canvas.height);
  const right = ceilEdge((rect.x + rect.width + pad) * canvas.width);
  const bottom = ceilEdge((rect.y + rect.height + pad) * canvas.height);

  const x = Math.max(0, Math.min(canvas.width, left));
  const y = Math.max(0, Math.min(canvas.height, top));
  const x2 = Math.max(0, Math.min(canvas.width, right));
  const y2 = Math.max(0, Math.min(canvas.height, bottom));
  const width = x2 - x;
  const height = y2 - y;
  if (width <= 0 || height <= 0) return null;
  return { x, y, width, height };
}

type CanvasLike = HTMLCanvasElement;
export type CanvasQuarterTurns = 0 | 1 | 2 | 3;

/**
 * Copy a page-fraction region of `source` into a new canvas.
 *
 * `createCanvas` is injectable so this can be exercised without a real DOM;
 * the default is the document's own.
 */
export function cropCanvasRegion(
  source: CanvasLike | null | undefined,
  rect: FractionRect,
  options: {
    padding?: number;
    createCanvas?: () => CanvasLike;
    /** Clockwise 90-degree turns applied to the copied pixels. */
    quarterTurns?: CanvasQuarterTurns;
    flipX?: boolean;
    flipY?: boolean;
  } = {},
): { canvas: CanvasLike; rect: PixelRect } | null {
  if (!source) return null;
  const pixels = cropPixelRect(source, rect, { padding: options.padding });
  if (!pixels) return null;
  const make =
    options.createCanvas ??
    (() => document.createElement("canvas") as CanvasLike);
  const target = make();
  const quarterTurns = options.quarterTurns ?? 0;
  if (![0, 1, 2, 3].includes(quarterTurns)) return null;
  if (quarterTurns !== 0 && (options.flipX || options.flipY)) return null;
  const swapsAxes = quarterTurns % 2 === 1;
  target.width = swapsAxes ? pixels.height : pixels.width;
  target.height = swapsAxes ? pixels.width : pixels.height;
  const ctx = target.getContext("2d");
  if (!ctx) return null;
  try {
    if (options.flipX && options.flipY) {
      ctx.translate(target.width, target.height);
      ctx.scale(-1, -1);
    } else if (options.flipX) {
      ctx.translate(target.width, 0);
      ctx.scale(-1, 1);
    } else if (options.flipY) {
      ctx.translate(0, target.height);
      ctx.scale(1, -1);
    } else if (quarterTurns === 1) {
      ctx.translate(target.width, 0);
      ctx.rotate(Math.PI / 2);
    } else if (quarterTurns === 2) {
      ctx.translate(target.width, target.height);
      ctx.rotate(Math.PI);
    } else if (quarterTurns === 3) {
      ctx.translate(0, target.height);
      ctx.rotate(-Math.PI / 2);
    }
    ctx.drawImage(
      source,
      pixels.x,
      pixels.y,
      pixels.width,
      pixels.height,
      0,
      0,
      pixels.width,
      pixels.height,
    );
  } catch {
    // A tainted canvas throws here. The page canvas is same-origin, but a
    // caller could hand over one that is not, and a thrown error inside a
    // click handler would look like the button simply doing nothing.
    return null;
  }
  return { canvas: target, rect: pixels };
}

/** PNG bytes for a cropped region, or null if it could not be taken. */
export async function cropCanvasRegionToPng(
  source: CanvasLike | null | undefined,
  rect: FractionRect,
  options: {
    padding?: number;
    createCanvas?: () => CanvasLike;
    quarterTurns?: CanvasQuarterTurns;
    flipX?: boolean;
    flipY?: boolean;
  } = {},
): Promise<Uint8Array | null> {
  const cropped = cropCanvasRegion(source, rect, options);
  if (!cropped) return null;
  const blob = await new Promise<Blob | null>((resolve) => {
    try {
      cropped.canvas.toBlob((value) => resolve(value), "image/png");
    } catch {
      resolve(null);
    }
  });
  // The scratch canvas has served its purpose; release its backing store now
  // rather than at some future GC.
  cropped.canvas.width = 0;
  cropped.canvas.height = 0;
  if (!blob) return null;
  return new Uint8Array(await blob.arrayBuffer());
}
