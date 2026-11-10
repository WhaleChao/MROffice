/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * OCR text layer for SCANNED pdf pages (no embedded text stream).
 *
 * Scanned pages render a canvas bitmap but pdf.js's text layer comes out
 * EMPTY, so selection and find go dark. The viewer offers an explicit
 * "识别文字 (OCR)" affordance on such pages; the backend
 * (`/api/v1/ai/notebook/pdf-ocr-page`) returns positioned text runs
 * `{ text, bbox: [x0, y0, x1, y1] }` with bbox coordinates as PAGE FRACTIONS
 * in [0, 1] (top-left origin). This module owns the pure math + DOM helper
 * that turns those runs into invisible, absolutely-positioned spans INSIDE
 * the page's existing `div[data-pdf-text-layer]` container — the same
 * geometry model pdf.js's own TextLayer uses — so text selection, the find
 * bar and the doc-search panel all work unchanged:
 *
 *   - the container's CSS (`.pdf-viewer__text-layer > span`) already makes
 *     every child span transparent, selectable and `transform-origin: 0 0`;
 *   - the viewer's find-rect mapper measures `layer.querySelectorAll("span")`
 *     by ITEM INDEX, so injecting one span per run (in run order) into an
 *     otherwise-empty layer keeps item index == span index.
 *
 * Everything except `renderOcrTextLayer` is DOM-free so the geometry math
 * unit-tests in isolation; the DOM helper itself is jsdom-testable.
 */

/** One positioned OCR text line, bbox in page fractions (top-left origin). */
export interface PdfOcrRun {
  text: string;
  /** [x0, y0, x1, y1] page fractions in [0, 1]; x1 > x0 and y1 > y0. */
  bbox: [number, number, number, number];
}

/**
 * Mirror of the server-side caps (apps/api/app/services/pdf_ocr.py). The text
 * cap counts Unicode code points, like the server's Python `str` slice.
 */
export const MAX_OCR_RUNS_PER_PAGE = 500;
export const MAX_OCR_RUN_TEXT_CHARS = 1000;

/**
 * Long-edge cap (device px) for the page raster sent to the OCR endpoint.
 * Big enough for legible glyphs on an A4 scan, small enough that a JPEG
 * stays far under the server's decoded-byte cap.
 */
export const OCR_RASTER_MAX_EDGE = 2000;

/** Absolutely-positioned span box for one run, in text-layer CSS pixels. */
export interface OcrSpanGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
  /** Font size driving the glyph height; the caller scales X to fit width. */
  fontSize: number;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function clampFraction(value: number): number {
  if (value < 0) return 0;
  if (value > 1) return 1;
  return value;
}

/**
 * Unpaired UTF-16 surrogates. With the `u` flag a well-formed pair is ONE
 * astral code point outside this class, so only lone halves match.
 */
const LONE_SURROGATE_REGEX = /[\ud800-\udfff]/gu;

/**
 * The first `max` code points of `value`. A surrogate pair counts once and is
 * never split, so the result stays well-formed when its input is.
 */
function sliceCodePoints(value: string, max: number): string {
  if (value.length <= max) return value;
  let end = 0;
  for (let count = 0; count < max && end < value.length; count += 1) {
    const unit = value.charCodeAt(end);
    const next = value.charCodeAt(end + 1);
    end +=
      unit >= 0xd800 && unit <= 0xdbff && next >= 0xdc00 && next <= 0xdfff
        ? 2
        : 1;
  }
  return value.slice(0, end);
}

/**
 * Validate + clamp an untrusted runs payload (endpoint response OR persisted
 * annotation read back on open) into well-formed `PdfOcrRun`s. Mirrors the
 * server sanitizer: non-array input yields `[]`; malformed entries are
 * dropped; bbox fractions clamp into [0, 1]; degenerate boxes (x1 <= x0 or
 * y1 <= y0 after clamping) are dropped; text has lone surrogates replaced
 * with U+FFFD, is trimmed, and is capped to `MAX_OCR_RUN_TEXT_CHARS` code
 * points without splitting a pair; at most `MAX_OCR_RUNS_PER_PAGE` runs are
 * kept in order.
 */
export function sanitizeOcrRuns(raw: unknown): PdfOcrRun[] {
  if (!Array.isArray(raw)) return [];
  const clean: PdfOcrRun[] = [];
  for (const entry of raw) {
    if (clean.length >= MAX_OCR_RUNS_PER_PAGE) break;
    if (!entry || typeof entry !== "object") continue;
    const text = (entry as { text?: unknown }).text;
    if (typeof text !== "string") continue;
    const trimmed = sliceCodePoints(
      text.replace(LONE_SURROGATE_REGEX, "\ufffd").trim(),
      MAX_OCR_RUN_TEXT_CHARS,
    );
    if (!trimmed) continue;
    const bbox = (entry as { bbox?: unknown }).bbox;
    if (!Array.isArray(bbox) || bbox.length !== 4) continue;
    if (!bbox.every(isFiniteNumber)) continue;
    const x0 = clampFraction(bbox[0]);
    const y0 = clampFraction(bbox[1]);
    const x1 = clampFraction(bbox[2]);
    const y1 = clampFraction(bbox[3]);
    if (x1 <= x0 || y1 <= y0) continue;
    clean.push({ text: trimmed, bbox: [x0, y0, x1, y1] });
  }
  return clean;
}

/**
 * Map one run's fractional bbox onto a text layer of `layerWidth` ×
 * `layerHeight` CSS pixels. Returns null for a malformed run / non-positive
 * layer size (defensive: runs normally passed through `sanitizeOcrRuns`
 * first). `fontSize` equals the box height — with the layer's `line-height:
 * 1` the glyph row fills the box vertically, matching how pdf.js sizes its
 * own spans; horizontal fit is the caller's scaleX (see `ocrSpanScaleX`).
 */
export function ocrRunGeometry(
  run: PdfOcrRun,
  layerWidth: number,
  layerHeight: number,
): OcrSpanGeometry | null {
  if (
    !isFiniteNumber(layerWidth) ||
    !isFiniteNumber(layerHeight) ||
    layerWidth <= 0 ||
    layerHeight <= 0
  ) {
    return null;
  }
  const bbox = run?.bbox;
  if (!Array.isArray(bbox) || bbox.length !== 4 || !bbox.every(isFiniteNumber)) {
    return null;
  }
  const x0 = clampFraction(bbox[0]);
  const y0 = clampFraction(bbox[1]);
  const x1 = clampFraction(bbox[2]);
  const y1 = clampFraction(bbox[3]);
  if (x1 <= x0 || y1 <= y0) return null;
  const width = (x1 - x0) * layerWidth;
  const height = (y1 - y0) * layerHeight;
  return {
    left: x0 * layerWidth,
    top: y0 * layerHeight,
    width,
    height,
    fontSize: height,
  };
}

/**
 * Horizontal scale that stretches/squeezes a span's naturally-rendered text
 * (`measuredWidth`) onto the OCR box width (`targetWidth`) — the same trick
 * pdf.js's TextLayer uses so the invisible glyphs sit over the printed ones.
 * Returns 1 when either width is unusable (jsdom / not-yet-laid-out spans
 * measure 0 — an unscaled span is still selectable, just less aligned).
 */
export function ocrSpanScaleX(targetWidth: number, measuredWidth: number): number {
  if (
    !isFiniteNumber(targetWidth) ||
    !isFiniteNumber(measuredWidth) ||
    targetWidth <= 0 ||
    measuredWidth <= 0
  ) {
    return 1;
  }
  return targetWidth / measuredWidth;
}

/**
 * Integer raster size for the page snapshot sent to the OCR endpoint: the
 * source canvas box scaled (down only — never upscaled) so its long edge is
 * at most `maxEdge` device pixels. Returns null for a non-positive /
 * non-finite source box.
 */
export function ocrRasterDimensions(
  sourceWidth: number,
  sourceHeight: number,
  maxEdge: number = OCR_RASTER_MAX_EDGE,
): { width: number; height: number } | null {
  if (
    !isFiniteNumber(sourceWidth) ||
    !isFiniteNumber(sourceHeight) ||
    sourceWidth <= 0 ||
    sourceHeight <= 0 ||
    !isFiniteNumber(maxEdge) ||
    maxEdge <= 0
  ) {
    return null;
  }
  const longEdge = Math.max(sourceWidth, sourceHeight);
  const scale = longEdge > maxEdge ? maxEdge / longEdge : 1;
  return {
    width: Math.max(1, Math.round(sourceWidth * scale)),
    height: Math.max(1, Math.round(sourceHeight * scale)),
  };
}

/** The data attribute marking spans injected by `renderOcrTextLayer`. */
export const OCR_SPAN_ATTR = "data-pdf-ocr-span";

/**
 * Inject one invisible span per run into the (empty) pdf.js text-layer
 * container so selection / find work on a scanned page. Previously-injected
 * OCR spans are removed first (idempotent across zoom re-renders); pdf.js's
 * own spans are never touched. Spans are appended in run order so the
 * viewer's item-index → span-index find mapping holds. Returns the number of
 * spans injected.
 */
export function renderOcrTextLayer(
  container: HTMLElement,
  runs: ReadonlyArray<PdfOcrRun>,
  layerWidth: number,
  layerHeight: number,
): number {
  for (const stale of Array.from(
    container.querySelectorAll(`[${OCR_SPAN_ATTR}]`),
  )) {
    stale.remove();
  }
  let injected = 0;
  for (const run of runs) {
    const geometry = ocrRunGeometry(run, layerWidth, layerHeight);
    if (!geometry) continue;
    const span = container.ownerDocument.createElement("span");
    span.setAttribute(OCR_SPAN_ATTR, "");
    span.textContent = run.text;
    span.style.left = `${geometry.left}px`;
    span.style.top = `${geometry.top}px`;
    span.style.fontSize = `${geometry.fontSize}px`;
    span.style.fontFamily = "sans-serif";
    container.appendChild(span);
    // Fit the rendered glyph row onto the OCR box width. Measure AFTER the
    // append so layout knows the font; a 0 measurement (jsdom, display:none)
    // leaves the span unscaled rather than collapsing it.
    const measured = span.getBoundingClientRect().width;
    const scaleX = ocrSpanScaleX(geometry.width, measured);
    if (scaleX !== 1) {
      span.style.transform = `scaleX(${scaleX})`;
    }
    injected += 1;
  }
  return injected;
}
