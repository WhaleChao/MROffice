/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry for the PDF viewer's SNAPSHOT tool (Acrobat "Edit ▸ Take a
 * Snapshot", WPS 截图): the user drags a marquee over a rendered page and the
 * selected REGION is copied to the clipboard as an image. The component captures
 * the drag as two corner points in CSS pixels relative to the page's rendered
 * `<canvas>`, plus the canvas's CSS box size and its DEVICE-pixel backing-store
 * size (which differ by the devicePixelRatio at which pdf.js rendered). This
 * module turns that into a CLAMPED crop rectangle in the canvas's DEVICE-pixel
 * space, ready to feed `ctx.drawImage(src, sx, sy, sw, sh, …)` for a crisp crop.
 *
 * Everything here is DOM-free so it unit-tests in isolation; the actual
 * `drawImage` + `canvas.toBlob` + `navigator.clipboard.write` is thin wiring in
 * the component. Mirrors the defensive, total style of `crop-box.ts` /
 * `shapes.ts`: never throws on malformed input, returns null when the region is
 * unusable.
 */

/** A point in CSS pixels relative to the rendered canvas's top-left. */
export interface SnapshotPoint {
  x: number;
  y: number;
}

/** The rendered canvas's CSS box + device-pixel backing store dimensions. */
export interface CanvasMetrics {
  /** Canvas CSS (layout) width in pixels — the on-screen size. */
  cssWidth: number;
  /** Canvas CSS (layout) height in pixels. */
  cssHeight: number;
  /** Canvas backing-store (`canvas.width`) in device pixels. */
  deviceWidth: number;
  /** Canvas backing-store (`canvas.height`) in device pixels. */
  deviceHeight: number;
}

/** A pixel rectangle (top-left origin). Width / height are >= 0. */
export interface PixelRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Minimum CROP size (device px) for a usable snapshot — a stray click is ignored. */
export const SNAPSHOT_MIN_DEVICE_PX = 4;

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function clamp(value: number, min: number, max: number): number {
  if (value < min) return min;
  if (value > max) return max;
  return value;
}

function metricsValid(metrics: CanvasMetrics): boolean {
  return (
    isFiniteNumber(metrics.cssWidth) &&
    isFiniteNumber(metrics.cssHeight) &&
    isFiniteNumber(metrics.deviceWidth) &&
    isFiniteNumber(metrics.deviceHeight) &&
    metrics.cssWidth > 0 &&
    metrics.cssHeight > 0 &&
    metrics.deviceWidth > 0 &&
    metrics.deviceHeight > 0
  );
}

/**
 * Normalize a drag (two corners, CSS px) into a CSS-space rect clamped to the
 * canvas's CSS box (a backwards / out-of-bounds drag still yields a top-left,
 * in-bounds box). Returns null for malformed input or invalid metrics. The width
 * / height may be 0 here (size is enforced later, in device space).
 */
export function normalizeSnapshotDragCss(
  start: SnapshotPoint,
  end: SnapshotPoint,
  metrics: CanvasMetrics,
): PixelRect | null {
  if (
    !isFiniteNumber(start?.x) ||
    !isFiniteNumber(start?.y) ||
    !isFiniteNumber(end?.x) ||
    !isFiniteNumber(end?.y) ||
    !metricsValid(metrics)
  ) {
    return null;
  }
  const sx = clamp(start.x, 0, metrics.cssWidth);
  const sy = clamp(start.y, 0, metrics.cssHeight);
  const ex = clamp(end.x, 0, metrics.cssWidth);
  const ey = clamp(end.y, 0, metrics.cssHeight);
  const x = Math.min(sx, ex);
  const y = Math.min(sy, ey);
  return {
    x,
    y,
    width: Math.abs(ex - sx),
    height: Math.abs(ey - sy),
  };
}

/**
 * Map a CSS-space rect onto the canvas's DEVICE-pixel space (the backing store),
 * scaling by the per-axis ratio `device/css`. The result is clamped so it never
 * exceeds the backing store, and rounded to integer pixels (sub-pixel source
 * rects make `drawImage` blurry). Returns null on invalid metrics.
 */
export function cssRectToDevice(
  rect: PixelRect,
  metrics: CanvasMetrics,
): PixelRect | null {
  if (!metricsValid(metrics)) return null;
  if (
    !isFiniteNumber(rect?.x) ||
    !isFiniteNumber(rect?.y) ||
    !isFiniteNumber(rect?.width) ||
    !isFiniteNumber(rect?.height)
  ) {
    return null;
  }
  const scaleX = metrics.deviceWidth / metrics.cssWidth;
  const scaleY = metrics.deviceHeight / metrics.cssHeight;
  const x = clamp(Math.round(rect.x * scaleX), 0, metrics.deviceWidth);
  const y = clamp(Math.round(rect.y * scaleY), 0, metrics.deviceHeight);
  // Width / height clamped to the remaining backing store from (x, y).
  const width = clamp(
    Math.round(rect.width * scaleX),
    0,
    metrics.deviceWidth - x,
  );
  const height = clamp(
    Math.round(rect.height * scaleY),
    0,
    metrics.deviceHeight - y,
  );
  return { x, y, width, height };
}

/**
 * Full pipeline: a drag (CSS px corners) → a clamped DEVICE-pixel crop rect ready
 * for `ctx.drawImage`. Returns null when the drag is malformed OR the resulting
 * crop is smaller than {@link SNAPSHOT_MIN_DEVICE_PX} on either axis (a stray
 * click, not a real marquee).
 */
export function snapshotCropRect(
  start: SnapshotPoint,
  end: SnapshotPoint,
  metrics: CanvasMetrics,
): PixelRect | null {
  const cssRect = normalizeSnapshotDragCss(start, end, metrics);
  if (!cssRect) return null;
  const device = cssRectToDevice(cssRect, metrics);
  if (!device) return null;
  if (
    device.width < SNAPSHOT_MIN_DEVICE_PX ||
    device.height < SNAPSHOT_MIN_DEVICE_PX
  ) {
    return null;
  }
  return device;
}

/** Result of decoding a canvas data URL into raw bytes + MIME (snapshot PNG). */
export interface SnapshotImageData {
  bytes: Uint8Array;
  mime: string;
}

/**
 * Decode a `data:image/<type>;base64,<payload>` URL (what `canvas.toDataURL`
 * emits for the cropped snapshot) into raw bytes + MIME. Returns null for a
 * non-image / non-base64 / empty / invalid payload. Mirrors
 * `canvas-export.decodeImageDataUrl` but lives here so the snapshot path has no
 * cross-feature import coupling. Pure (uses `atob`, available in browser + jsdom).
 */
export function decodeSnapshotDataUrl(dataUrl: unknown): SnapshotImageData | null {
  if (typeof dataUrl !== "string") return null;
  const match = /^data:(image\/[a-z0-9.+-]+);base64,([\s\S]*)$/i.exec(dataUrl);
  if (!match) return null;
  const mime = match[1].toLowerCase();
  const payload = match[2].replace(/\s/g, "");
  if (!payload) return null;
  let binary: string;
  try {
    binary = atob(payload);
  } catch {
    return null;
  }
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i) & 0xff;
  }
  return { bytes, mime };
}

/**
 * Whether the runtime supports copying an IMAGE to the clipboard — i.e. the
 * async Clipboard API's `write` AND the `ClipboardItem` constructor both exist.
 * (Plain text clipboard via `writeText` is far more widely supported; image
 * write is gated behind these two.) The component falls back to a download when
 * this is false. Pure feature-detect; takes the candidate global so it can be
 * unit-tested without a real navigator.
 */
export function canCopyImageToClipboard(
  nav: unknown,
  clipboardItemCtor: unknown,
): boolean {
  if (typeof clipboardItemCtor !== "function") return false;
  if (!nav || typeof nav !== "object") return false;
  const clipboard = (nav as { clipboard?: unknown }).clipboard;
  if (!clipboard || typeof clipboard !== "object") return false;
  return typeof (clipboard as { write?: unknown }).write === "function";
}

/**
 * WPS' screenshot regions: a rectangle, an ellipse, a rounded rectangle, or
 * a custom outline drawn by hand (矩形 / 椭圆 / 圆角矩形 / 自定义区域截图).
 * A region other than the rectangle is copied as a PNG clear outside it.
 */
export type SnapshotShape = "rect" | "ellipse" | "rounded" | "freeform";

/** The rounded rectangle's corner in CSS px, as its marquee shows it. */
export const SNAPSHOT_ROUNDED_RADIUS_CSS = 16;

/** The bounding box of a hand-drawn outline, as a marquee's two corners. */
export function snapshotOutlineBounds(
  outline: readonly SnapshotPoint[],
): { start: SnapshotPoint; current: SnapshotPoint } | null {
  if (outline.length < 3) return null;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const point of outline) {
    minX = Math.min(minX, point.x);
    minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x);
    maxY = Math.max(maxY, point.y);
  }
  return { start: { x: minX, y: minY }, current: { x: maxX, y: maxY } };
}

type SnapshotShapePath = Pick<
  CanvasRenderingContext2D,
  "beginPath" | "ellipse" | "moveTo" | "lineTo" | "arcTo" | "closePath" | "rect"
>;

/**
 * Trace a region's outline on a canvas the size of its crop, for a
 * `destination-in` fill to keep the picture inside it. `radius` is the
 * rounded rectangle's corner and `outline` the hand-drawn points, both in the
 * crop's pixels.
 */
export function traceSnapshotShape(
  ctx: SnapshotShapePath,
  shape: SnapshotShape,
  width: number,
  height: number,
  options: { radius?: number; outline?: readonly SnapshotPoint[] } = {},
): void {
  ctx.beginPath();
  if (shape === "ellipse") {
    ctx.ellipse(width / 2, height / 2, width / 2, height / 2, 0, 0, Math.PI * 2);
  } else if (shape === "rounded") {
    const r = Math.max(0, Math.min(options.radius ?? 0, width / 2, height / 2));
    ctx.moveTo(r, 0);
    ctx.arcTo(width, 0, width, height, r);
    ctx.arcTo(width, height, 0, height, r);
    ctx.arcTo(0, height, 0, 0, r);
    ctx.arcTo(0, 0, width, 0, r);
  } else if (shape === "freeform" && options.outline && options.outline.length >= 3) {
    ctx.moveTo(options.outline[0].x, options.outline[0].y);
    for (const point of options.outline.slice(1)) ctx.lineTo(point.x, point.y);
  } else {
    ctx.rect(0, 0, width, height);
  }
  ctx.closePath();
}
