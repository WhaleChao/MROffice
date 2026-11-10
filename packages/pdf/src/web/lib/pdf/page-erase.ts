/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * WPS 擦除 ("Erase").
 *
 * A dragged rectangle erases what the page draws inside it: the reader's text
 * items it holds whole are deleted from the page content and the embedded
 * pictures it holds whole are taken out (the native text-delete and
 * picture-removal writers); whatever else shows inside -- vector art, a
 * picture or a line of text the rectangle only cuts, anything a writer
 * declines -- is covered with the chosen colour, and the reader is told
 * (「存在无法擦除内容」). All of it is one undo step. "自动" (automatic)
 * covers in white, as WPS does. A brush stroke only covers: a round-capped
 * line of the chosen width in the chosen colour (自动 paints the colour the
 * page shows under it, sampled from the rendered page).
 *
 * A cover is page content (one path block appended to the page's content
 * stream), never a redaction: what it paints over stays in the file.
 *
 * Geometry is page-percent with a top-left origin in the page as displayed
 * (after /Rotate) -- the space every reader overlay uses. The native mapping
 * turns it into user space through the crop box and /Rotate, as it does for a
 * drawn shape (native-content-append-shapes.ts).
 */

import {
  pdfNativeContentAppendColorFromHex,
  pdfNativeContentAppendDisplayPoint,
  pdfNativeContentAppendDisplaySize,
} from "./engine/native/content-append-commands";
import type {
  PdfNativeContentAppendPageGeometry,
  PdfNativeContentAppendPathCommand,
  PdfNativeContentAppendPathContent,
} from "./engine/native/content-append-transaction";

export type PdfEraseShape = "rect" | "brush";

/** Brush widths offered in the Erase menu, in points (the page's own unit). */
export const PDF_ERASE_BRUSH_WIDTHS = [5, 10, 20] as const;
export type PdfEraseBrushWidth = (typeof PDF_ERASE_BRUSH_WIDTHS)[number];
export const PDF_ERASE_DEFAULT_BRUSH_WIDTH: PdfEraseBrushWidth = 10;

/** A drag narrower or shorter than this (page percent) is a stray click. */
export const PDF_ERASE_MIN_RECT_PERCENT = 0.3;
/** Brush samples closer than this (page percent) to the last kept one add nothing. */
export const PDF_ERASE_BRUSH_MIN_STEP_PERCENT = 0.12;
/**
 * Points kept per brush stroke. Each point is one path operator and a
 * content-append path holds at most 4,096 of them.
 */
export const PDF_ERASE_MAX_BRUSH_POINTS = 1_500;
/** Pixels read when the colour is sampled: plenty for a dominant colour. */
const MAX_COLOR_SAMPLES = 40_000;

export interface PdfErasePoint {
  readonly x: number;
  readonly y: number;
}

/** A normalized rectangle: x1 < x2, y1 < y2, all within [0, 100]. */
export interface PdfEraseRect {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

/** Displayed page size in points (width and height after /Rotate). */
export interface PdfEraseDisplaySize {
  readonly width: number;
  readonly height: number;
}

/** The colour 自动 covers an erased rectangle in (WPS paints white). */
export const PDF_ERASE_AUTOMATIC_COVER_COLOR = "#ffffff";

/** A reader text item, as the reader placed it: its text and its origin in
 * PDF user space (pdf.js `transform[4]`, `transform[5]`). */
export interface PdfEraseTextTarget {
  readonly str: string;
  readonly x: number;
  readonly y: number;
}

/**
 * What an erase rectangle holds, as the reader reads the page: the text items
 * and pictures it holds whole (taken off the page) and whether anything else
 * shows inside it (covered).
 */
export interface PdfEraseTargets {
  readonly textItems: readonly PdfEraseTextTarget[];
  /** Paint orders of the embedded pictures held whole (the picture editor's key). */
  readonly imageOrders: readonly number[];
  /** Each held picture's paint CTM as the reader read it, by paint order: the
   * host takes away the paint drawn there, never another one. */
  readonly imageMatrices?: Readonly<Record<number, readonly number[]>>;
  /** Something inside is not erased: vector art, a picture or a text item the
   * rectangle only cuts. */
  readonly cover: boolean;
}

/** What an erase did: whether the page was written, and whether part of the
 * rectangle was only covered (the reader is told). */
export interface PdfEraseOutcome {
  readonly written: boolean;
  readonly covered: boolean;
}

export type PdfEraseStroke =
  | Readonly<{
      shape: "rect";
      /** 1-based page number. */
      page: number;
      /** `#rrggbb`: the chosen colour, or white for 自动. */
      color: string;
      rect: PdfEraseRect;
      /** What the rectangle erases; absent: it only covers. */
      targets?: PdfEraseTargets;
    }>
  | Readonly<{
      shape: "brush";
      page: number;
      color: string;
      points: readonly PdfErasePoint[];
      /** Brush diameter in points. */
      width: number;
    }>;

const HEX_COLOR = /^#[0-9a-f]{6}$/i;

export function isPdfEraseColor(value: unknown): value is string {
  return typeof value === "string" && HEX_COLOR.test(value);
}

function clampPercent(value: number): number {
  return Math.min(100, Math.max(0, value));
}

/** The rectangle a drag from `start` to `end` covers, or null for a stray click. */
export function normalizePdfEraseRect(
  start: PdfErasePoint,
  end: PdfErasePoint,
): PdfEraseRect | null {
  const values = [start.x, start.y, end.x, end.y];
  if (!values.every((value) => Number.isFinite(value))) return null;
  const x1 = clampPercent(Math.min(start.x, end.x));
  const x2 = clampPercent(Math.max(start.x, end.x));
  const y1 = clampPercent(Math.min(start.y, end.y));
  const y2 = clampPercent(Math.max(start.y, end.y));
  if (
    x2 - x1 < PDF_ERASE_MIN_RECT_PERCENT ||
    y2 - y1 < PDF_ERASE_MIN_RECT_PERCENT
  ) {
    return null;
  }
  return { x1, y1, x2, y2 };
}

/**
 * Thin a brush stroke's samples: drop the ones closer than the minimum step
 * to the last kept point, then decimate evenly if the stroke is still longer
 * than a path may be. The first and the last samples are always kept, so the
 * stroke starts and ends where the pointer did.
 */
export function thinPdfEraseBrushPoints(
  points: readonly PdfErasePoint[],
): PdfErasePoint[] {
  const finitePoints = points
    .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y))
    .map((point) => ({ x: clampPercent(point.x), y: clampPercent(point.y) }));
  if (finitePoints.length === 0) return [];
  const kept: PdfErasePoint[] = [finitePoints[0]];
  for (let index = 1; index < finitePoints.length; index += 1) {
    const point = finitePoints[index];
    const last = kept[kept.length - 1];
    if (
      Math.hypot(point.x - last.x, point.y - last.y) >=
      PDF_ERASE_BRUSH_MIN_STEP_PERCENT
    ) {
      kept.push(point);
    }
  }
  const tail = finitePoints[finitePoints.length - 1];
  const lastKept = kept[kept.length - 1];
  if (kept.length > 1 && (lastKept.x !== tail.x || lastKept.y !== tail.y)) {
    kept.push(tail);
  }
  if (kept.length <= PDF_ERASE_MAX_BRUSH_POINTS) return kept;
  const step = (kept.length - 1) / (PDF_ERASE_MAX_BRUSH_POINTS - 1);
  return Array.from(
    { length: PDF_ERASE_MAX_BRUSH_POINTS },
    (_, index) => kept[Math.min(kept.length - 1, Math.round(index * step))],
  );
}

/** The page-percent box a stroke paints, its brush radius included. */
export function pdfEraseStrokeBounds(
  stroke: PdfEraseStroke,
  size: PdfEraseDisplaySize,
): PdfEraseRect {
  if (stroke.shape === "rect") return stroke.rect;
  const radiusX = size.width > 0 ? (stroke.width / 2 / size.width) * 100 : 0;
  const radiusY = size.height > 0 ? (stroke.width / 2 / size.height) * 100 : 0;
  const xs = stroke.points.map((point) => point.x);
  const ys = stroke.points.map((point) => point.y);
  return {
    x1: clampPercent(Math.min(...xs) - radiusX),
    y1: clampPercent(Math.min(...ys) - radiusY),
    x2: clampPercent(Math.max(...xs) + radiusX),
    y2: clampPercent(Math.max(...ys) + radiusY),
  };
}

/**
 * Whether a page-percent point lies under the stroke: inside the rectangle,
 * or within the brush radius of the stroke's centre line (measured in points,
 * so a brush stays round on a page that is not square).
 */
export function pdfEraseStrokeCovers(
  stroke: PdfEraseStroke,
  point: PdfErasePoint,
  size: PdfEraseDisplaySize,
): boolean {
  if (stroke.shape === "rect") {
    const { x1, y1, x2, y2 } = stroke.rect;
    return point.x >= x1 && point.x <= x2 && point.y >= y1 && point.y <= y2;
  }
  const px = (point.x / 100) * size.width;
  const py = (point.y / 100) * size.height;
  const radius = stroke.width / 2;
  const radiusSquared = radius * radius;
  const toPoints = (sample: PdfErasePoint) => ({
    x: (sample.x / 100) * size.width,
    y: (sample.y / 100) * size.height,
  });
  if (stroke.points.length === 1) {
    const only = toPoints(stroke.points[0]);
    return (px - only.x) ** 2 + (py - only.y) ** 2 <= radiusSquared;
  }
  for (let index = 1; index < stroke.points.length; index += 1) {
    const a = toPoints(stroke.points[index - 1]);
    const b = toPoints(stroke.points[index]);
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const lengthSquared = dx * dx + dy * dy;
    const t =
      lengthSquared === 0
        ? 0
        : Math.min(
            1,
            Math.max(0, ((px - a.x) * dx + (py - a.y) * dy) / lengthSquared),
          );
    const cx = a.x + t * dx;
    const cy = a.y + t * dy;
    if ((px - cx) ** 2 + (py - cy) ** 2 <= radiusSquared) return true;
  }
  return false;
}

function toHex(value: number): string {
  return Math.round(Math.min(255, Math.max(0, value)))
    .toString(16)
    .padStart(2, "0");
}

/**
 * The colour that covers most of a block of RGBA pixels: colours are bucketed
 * at 4 bits a channel (anti-aliased text edges fall into their own buckets),
 * the fullest bucket wins and its pixels are averaged for the exact shade.
 * Transparent pixels (an unpainted canvas) count for nothing. `include`
 * limits the count to the pixels a stroke covers.
 */
export function dominantPdfEraseColor(
  pixels: ArrayLike<number>,
  width: number,
  height: number,
  include?: (x: number, y: number) => boolean,
): string | null {
  if (
    !Number.isInteger(width) ||
    !Number.isInteger(height) ||
    width <= 0 ||
    height <= 0 ||
    pixels.length < width * height * 4
  ) {
    return null;
  }
  const stride = Math.max(
    1,
    Math.ceil(Math.sqrt((width * height) / MAX_COLOR_SAMPLES)),
  );
  const counts = new Map<number, { n: number; r: number; g: number; b: number }>();
  let best: { n: number; r: number; g: number; b: number } | null = null;
  for (let y = 0; y < height; y += stride) {
    for (let x = 0; x < width; x += stride) {
      if (include && !include(x, y)) continue;
      const offset = (y * width + x) * 4;
      const alpha = pixels[offset + 3];
      if (alpha < 128) continue;
      const r = pixels[offset];
      const g = pixels[offset + 1];
      const b = pixels[offset + 2];
      const key = ((r >> 4) << 8) | ((g >> 4) << 4) | (b >> 4);
      let bucket = counts.get(key);
      if (!bucket) {
        bucket = { n: 0, r: 0, g: 0, b: 0 };
        counts.set(key, bucket);
      }
      bucket.n += 1;
      bucket.r += r;
      bucket.g += g;
      bucket.b += b;
      if (!best || bucket.n > best.n) best = bucket;
    }
  }
  if (!best) return null;
  return `#${toHex(best.r / best.n)}${toHex(best.g / best.n)}${toHex(best.b / best.n)}`;
}

/**
 * 自动: the colour the rendered page shows under a stroke. Reads the page's
 * canvas over the stroke's box only. Null when the page has not painted or
 * its pixels cannot be read; the caller then asks for a chosen colour rather
 * than guessing white.
 */
export function samplePdfEraseAutoColor(
  canvas: HTMLCanvasElement | null | undefined,
  stroke: PdfEraseStroke,
  size: PdfEraseDisplaySize,
): string | null {
  if (!canvas || canvas.width <= 0 || canvas.height <= 0) return null;
  const bounds = pdfEraseStrokeBounds(stroke, size);
  const left = Math.max(0, Math.floor((bounds.x1 / 100) * canvas.width));
  const top = Math.max(0, Math.floor((bounds.y1 / 100) * canvas.height));
  const right = Math.min(
    canvas.width,
    Math.max(left + 1, Math.ceil((bounds.x2 / 100) * canvas.width)),
  );
  const bottom = Math.min(
    canvas.height,
    Math.max(top + 1, Math.ceil((bounds.y2 / 100) * canvas.height)),
  );
  const width = right - left;
  const height = bottom - top;
  if (width <= 0 || height <= 0) return null;
  let image: ImageData;
  try {
    const context = canvas.getContext("2d");
    if (!context) return null;
    image = context.getImageData(left, top, width, height);
  } catch {
    return null;
  }
  const include =
    stroke.shape === "brush"
      ? (x: number, y: number) =>
          pdfEraseStrokeCovers(
            stroke,
            {
              x: ((left + x + 0.5) / canvas.width) * 100,
              y: ((top + y + 0.5) / canvas.height) * 100,
            },
            size,
          )
      : undefined;
  return (
    dominantPdfEraseColor(image.data, width, height, include) ??
    // A stroke thinner than one sampled pixel: the box decides.
    (include ? dominantPdfEraseColor(image.data, width, height) : null)
  );
}

/**
 * The native path block that paints a stroke: a filled rectangle without an
 * outline, or a round-capped, round-joined line of the brush width without a
 * fill. Null for a colour or a geometry that cannot be painted.
 */
export function pdfNativeContentAppendPathFromErase(
  stroke: PdfEraseStroke,
  geometry: PdfNativeContentAppendPageGeometry,
): PdfNativeContentAppendPathContent | null {
  const color = pdfNativeContentAppendColorFromHex(stroke.color);
  if (!color || !isPdfEraseColor(stroke.color)) return null;
  const size = pdfNativeContentAppendDisplaySize(geometry);
  if (!(size.width > 0) || !(size.height > 0)) return null;
  const toPdf = (
    point: PdfErasePoint,
    op: "m" | "l",
  ): PdfNativeContentAppendPathCommand => {
    const mapped = pdfNativeContentAppendDisplayPoint(
      geometry,
      clampPercent(point.x) / 100,
      clampPercent(point.y) / 100,
    );
    return { op, x: mapped.x, y: mapped.y };
  };
  if (stroke.shape === "rect") {
    const rect = normalizePdfEraseRect(
      { x: stroke.rect.x1, y: stroke.rect.y1 },
      { x: stroke.rect.x2, y: stroke.rect.y2 },
    );
    if (!rect) return null;
    return Object.freeze({
      kind: "path" as const,
      stroke: null,
      fill: Object.freeze({ color }),
      opacity: null,
      fillOpacity: null,
      commands: Object.freeze([
        toPdf({ x: rect.x1, y: rect.y1 }, "m"),
        toPdf({ x: rect.x2, y: rect.y1 }, "l"),
        toPdf({ x: rect.x2, y: rect.y2 }, "l"),
        toPdf({ x: rect.x1, y: rect.y2 }, "l"),
        { op: "h" as const },
      ]),
    });
  }
  const points = thinPdfEraseBrushPoints(stroke.points);
  if (
    points.length === 0 ||
    !Number.isFinite(stroke.width) ||
    stroke.width <= 0
  ) {
    return null;
  }
  // A click without a drag is a dot: a zero-length subpath with round caps
  // paints a disc of the brush width.
  const path = points.length === 1 ? [points[0], points[0]] : points;
  return Object.freeze({
    kind: "path" as const,
    stroke: Object.freeze({
      color,
      width: stroke.width,
      dash: null,
      cap: 1 as const,
      join: 1 as const,
    }),
    fill: null,
    opacity: null,
    fillOpacity: null,
    commands: Object.freeze(
      path.map((point, index) => toPdf(point, index === 0 ? "m" : "l")),
    ),
  });
}
