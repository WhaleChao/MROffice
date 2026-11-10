/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Ctrl/Cmd+wheel + pinch zoom math for the PDF reader, plus the scroll
 * anchoring that keeps the content point under the cursor (or the viewport
 * center for toolbar zoom) fixed across a render-scale change.
 *
 * Everything here is pure so the zoom behaviour is unit-testable without
 * jsdom layout. The wheel normalization and the two-decimal factor
 * accumulation are adapted from Mozilla pdf.js `web/app.js` (`onWheel` /
 * `_accumulateFactor`), Apache-2.0:
 * https://github.com/mozilla/pdf.js/blob/master/web/app.js
 */

// WheelEvent.deltaMode values (numeric literals so this module never touches
// the WheelEvent global — jsdom-free unit tests import it directly).
export const DOM_DELTA_PIXEL = 0;
export const DOM_DELTA_LINE = 1;
export const DOM_DELTA_PAGE = 2;

/** Pixels one DOM_DELTA_LINE wheel unit scrolls (Firefox line-mode wheels). */
export const PDF_WHEEL_LINE_PIXELS = 16;

/**
 * Pixels one wheel "notch" is worth for zooming: a notch multiplies the scale
 * by {@link PDF_WHEEL_ZOOM_STEP}. 53px is Firefox's canonical notch height;
 * a Chrome pixel-mode notch (~100px) lands slightly above one step, and
 * trackpad pans (a few px per event) produce smooth fractional steps.
 */
export const PDF_WHEEL_NOTCH_PIXELS = 53;

/** Zoom multiplier per wheel notch (pdf.js uses the same 1.1 per tick). */
export const PDF_WHEEL_ZOOM_STEP = 1.1;

/**
 * Bound on the scale change a SINGLE wheel event may apply, so an OS/driver
 * that reports one huge delta cannot jump across the whole zoom range.
 */
export const PDF_WHEEL_MAX_EVENT_FACTOR = 2;

/**
 * How long after a plain (unmodified) wheel scroll the Ctrl/Cmd+wheel zoom
 * stays disabled. Trackpad momentum keeps emitting plain wheel events after
 * the fingers lift; if the user then presses Ctrl/Cmd, those momentum events
 * suddenly carry the modifier and would zoom wildly. Same guard + duration as
 * pdf.js's zoomDisabledTimeout.
 */
export const PDF_WHEEL_ZOOM_DISABLED_MS = 1000;

/**
 * Multiplicative scale factor for one Ctrl/Cmd+wheel event. Line/page deltas
 * are normalized to pixels first; the result is clamped to
 * [1/{@link PDF_WHEEL_MAX_EVENT_FACTOR}, {@link PDF_WHEEL_MAX_EVENT_FACTOR}].
 * Scrolling up (negative deltaY) zooms in (> 1). Degenerate input → 1.
 */
export function wheelZoomScaleFactor(deltaY: number, deltaMode: number): number {
  const dy = Number(deltaY);
  if (!Number.isFinite(dy) || dy === 0) return 1;
  let pixels = dy;
  if (deltaMode === DOM_DELTA_LINE) {
    pixels = dy * PDF_WHEEL_LINE_PIXELS;
  } else if (deltaMode === DOM_DELTA_PAGE) {
    // Page-mode wheels report ±1 per event; treat one page as one notch
    // (pdf.js likewise maps one page-mode event to one zoom tick).
    pixels = dy * PDF_WHEEL_NOTCH_PIXELS;
  }
  const factor = Math.pow(
    PDF_WHEEL_ZOOM_STEP,
    -pixels / PDF_WHEEL_NOTCH_PIXELS,
  );
  return Math.min(
    PDF_WHEEL_MAX_EVENT_FACTOR,
    Math.max(1 / PDF_WHEEL_MAX_EVENT_FACTOR, factor),
  );
}

/**
 * Minimum pointer separation (px) a pinch pair must have before distance
 * ratios are trusted — below this, jitter between near-coincident touches
 * would produce wild factors.
 */
export const PDF_PINCH_MIN_DISTANCE = 8;

/**
 * Multiplicative scale factor for a two-pointer pinch step: the ratio of the
 * current finger distance to the previous one. Either distance under
 * {@link PDF_PINCH_MIN_DISTANCE} (or non-finite) → 1 (no zoom).
 */
export function pinchZoomScaleFactor(
  previousDistance: number,
  distance: number,
): number {
  const prev = Number(previousDistance);
  const next = Number(distance);
  if (
    !Number.isFinite(prev) ||
    !Number.isFinite(next) ||
    prev < PDF_PINCH_MIN_DISTANCE ||
    next < PDF_PINCH_MIN_DISTANCE
  ) {
    return 1;
  }
  return next / prev;
}

/** An {@link accumulateZoomFactor} result: the factor to apply now + carry. */
export interface AccumulatedZoomFactor {
  /** Factor to multiply the scale by NOW (quantized to 2-decimal scales). */
  value: number;
  /** Sub-quantum remainder to feed back into the next accumulation. */
  carry: number;
}

/**
 * Quantize a raw zoom factor so the resulting scale lands on a 2-decimal
 * value (what the toolbar percentage displays / `clampScale` rounds to), and
 * carry the sub-quantum remainder to the next event. Without this, trackpad
 * events (factors like 1.004) would each round to a no-op and the zoom would
 * stall. A direction flip drops the remainder.
 *
 * Adapted from pdf.js `_accumulateFactor`, with one fix: the remainder is
 * `(factor * carry) / value` (pdf.js uses `factor / value`, which lets an
 * endless run of tiny same-direction factors stall below the quantum).
 */
export function accumulateZoomFactor(
  previousScale: number,
  factor: number,
  carry: number,
): AccumulatedZoomFactor {
  let carried = Number.isFinite(carry) && carry > 0 ? carry : 1;
  if (
    !Number.isFinite(previousScale) ||
    previousScale <= 0 ||
    !Number.isFinite(factor) ||
    factor <= 0 ||
    factor === 1
  ) {
    return { value: 1, carry: carried };
  }
  if ((carried > 1 && factor < 1) || (carried < 1 && factor > 1)) {
    carried = 1;
  }
  const value =
    Math.floor(previousScale * factor * carried * 100) / (100 * previousScale);
  if (!Number.isFinite(value) || value <= 0) {
    return { value: 1, carry: 1 };
  }
  return { value, carry: (factor * carried) / value };
}

/**
 * The content point to keep stationary across a scale change, captured
 * BEFORE the new scale is applied:
 *  - `fractionX/Y`: the point's position as a fraction of the scroll
 *    content size (scale-invariant to first order — pages scale linearly,
 *    only the fixed paddings/gaps do not).
 *  - `offsetX/Y`: where that point sits inside the viewport (px from the
 *    scroll container's top-left), i.e. under the cursor / pinch center.
 */
export interface PdfZoomAnchor {
  fractionX: number;
  fractionY: number;
  offsetX: number;
  offsetY: number;
}

/** Scroll-container metrics needed to capture an anchor. */
export interface ZoomAnchorCaptureMetrics {
  scrollLeft: number;
  scrollTop: number;
  scrollWidth: number;
  scrollHeight: number;
}

/** Scroll-container metrics AFTER re-layout, needed to restore the anchor. */
export interface ZoomAnchorRestoreMetrics {
  scrollWidth: number;
  scrollHeight: number;
  clientWidth: number;
  clientHeight: number;
}

const fractionOf = (position: number, size: number): number => {
  const p = Number(position);
  const s = Number(size);
  if (!Number.isFinite(p) || !Number.isFinite(s) || s <= 0) return 0;
  return p / s;
};

/**
 * Capture the anchor for the content point `offsetX/Y` px inside the
 * viewport (cursor position relative to the scroll container, or the
 * viewport center for button zoom). Degenerate metrics → zero fractions,
 * which restore to a clamped top-left scroll (never NaN).
 */
export function captureZoomAnchor(
  view: ZoomAnchorCaptureMetrics,
  offsetX: number,
  offsetY: number,
): PdfZoomAnchor {
  const ox = Number.isFinite(offsetX) ? offsetX : 0;
  const oy = Number.isFinite(offsetY) ? offsetY : 0;
  return {
    fractionX: fractionOf(view.scrollLeft + ox, view.scrollWidth),
    fractionY: fractionOf(view.scrollTop + oy, view.scrollHeight),
    offsetX: ox,
    offsetY: oy,
  };
}

/** A restored scroll position (clamped to the scrollable range). */
export interface AnchoredScrollPosition {
  scrollLeft: number;
  scrollTop: number;
}

const clampScroll = (value: number, max: number): number => {
  if (!Number.isFinite(value)) return 0;
  return Math.min(Math.max(0, max), Math.max(0, value));
};

/**
 * The scroll position that puts the anchored content point back under the
 * captured viewport offset, given the post-relayout content size. Clamped to
 * the valid scroll range so an anchor near an edge degrades gracefully.
 */
export function anchoredScrollPosition(
  anchor: PdfZoomAnchor,
  view: ZoomAnchorRestoreMetrics,
): AnchoredScrollPosition {
  const width = Number(view.scrollWidth);
  const height = Number(view.scrollHeight);
  const sw = Number.isFinite(width) && width > 0 ? width : 0;
  const sh = Number.isFinite(height) && height > 0 ? height : 0;
  const cw = Number.isFinite(view.clientWidth) ? view.clientWidth : 0;
  const ch = Number.isFinite(view.clientHeight) ? view.clientHeight : 0;
  return {
    scrollLeft: clampScroll(anchor.fractionX * sw - anchor.offsetX, sw - cw),
    scrollTop: clampScroll(anchor.fractionY * sh - anchor.offsetY, sh - ch),
  };
}

/**
 * Cap on the backing-store pixels of a single page canvas. pdf.js uses the
 * same default (`maxCanvasPixels` = 2^25); beyond it the canvas renders at a
 * reduced device-pixel ratio and is CSS-upscaled, trading sharpness for not
 * exhausting canvas memory at 640% × dpr.
 */
export const PDF_MAX_RENDER_PIXELS = 1 << 25;

/**
 * The device-pixel-ratio multiplier to render a `cssWidth`×`cssHeight` page
 * at so its canvas stays within `maxPixels` backing pixels. Returns `dpr`
 * unchanged when the full-resolution canvas already fits; degenerate input
 * falls back to 1 (never 0 — a 0 ratio would produce an empty canvas).
 */
export function cappedRenderDpr(
  cssWidth: number,
  cssHeight: number,
  dpr: number,
  maxPixels: number = PDF_MAX_RENDER_PIXELS,
): number {
  const w = Number(cssWidth);
  const h = Number(cssHeight);
  const d = Number(dpr);
  if (!Number.isFinite(d) || d <= 0) return 1;
  if (!Number.isFinite(w) || w <= 0 || !Number.isFinite(h) || h <= 0) return d;
  const max = Number.isFinite(maxPixels) && maxPixels > 0 ? maxPixels : PDF_MAX_RENDER_PIXELS;
  const pixels = w * d * h * d;
  if (pixels <= max) return d;
  return d * Math.sqrt(max / pixels);
}
