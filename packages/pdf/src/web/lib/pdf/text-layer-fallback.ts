/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Fallback text layer for pages where pdf.js's own `TextLayer.render` throws
 * (historically Safari's ReadableStream handling with PDF.js v5; any engine
 * bug that breaks the native span pipeline).
 *
 * Instead of disabling selection/find wholesale, the viewer re-builds the
 * invisible text layer directly from `getTextContent` items: one absolutely
 * positioned span PER ITEM (in item order) inside the page's existing
 * `div[data-pdf-text-layer]` container. Keeping strict one-span-per-item
 * parity matters because the find-highlight mapper measures
 * `layer.querySelectorAll("span")` by ITEM INDEX (see
 * `buildSearchablePageText` in find-normalize.ts) — even empty items must
 * occupy a span slot.
 *
 * The geometry mirrors pdf.js's `TextLayer.#appendText` (pdfjs-dist v5,
 * src/display/text_layer.js, Apache-2.0): device-space matrix `tx =
 * viewport.transform × item.transform`, glyph row height `hypot(tx[2],
 * tx[3])`, span top at baseline minus ascent (pdf.js's 0.8 default ratio —
 * the fallback has no canvas font metrics), rotation from `atan2(tx[1],
 * tx[0])`, and a `scaleX` that stretches the naturally rendered run onto the
 * item's advance width. Everything except `renderFallbackTextLayer` is
 * DOM-free so the math unit-tests in isolation; the DOM helper is
 * jsdom-testable (mirrors `renderOcrTextLayer` in ocr-text-layer.ts).
 */

import type { PdfSearchTextItem } from "./find-normalize";
import { ocrSpanScaleX } from "./ocr-text-layer";

/** The data attribute marking nodes injected by `renderFallbackTextLayer`. */
export const FALLBACK_SPAN_ATTR = "data-pdf-fallback-text";

/**
 * Span cap per page, matching pdf.js's own MAX_TEXT_DIVS_TO_RENDER guard so
 * a pathological page cannot flood the DOM.
 */
export const MAX_FALLBACK_SPANS_PER_PAGE = 100_000;

/**
 * pdf.js's default ascent ratio (`TextLayer.#getAscent`) when canvas font
 * metrics are unavailable — the fallback never measures real font metrics.
 */
export const FALLBACK_ASCENT_RATIO = 0.8;

/** Absolutely-positioned span box for one text item, in text-layer CSS px. */
export interface FallbackSpanGeometry {
  left: number;
  top: number;
  /** Glyph row height; with the layer's `line-height: 1` it fills the row. */
  fontSize: number;
  /** Text rotation in degrees (0 for horizontal text), CSS-clockwise. */
  angleDeg: number;
  /** CSS width the run should occupy — the caller fits via scaleX. */
  targetWidth: number;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isMatrix(value: unknown): value is ReadonlyArray<number> {
  return (
    Array.isArray(value) &&
    value.length >= 6 &&
    value.slice(0, 6).every(isFiniteNumber)
  );
}

/**
 * `Util.transform(m1, m2)` from pdf.js (src/shared/util.js, Apache-2.0):
 * compose two PDF `[a, b, c, d, e, f]` matrices (apply `m2`, then `m1`).
 */
function composeTransform(
  m1: ReadonlyArray<number>,
  m2: ReadonlyArray<number>,
): [number, number, number, number, number, number] {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

/**
 * Compute the CSS-pixel span box for one `getTextContent` item on a text
 * layer positioned by `viewport.transform` (the CSS-scale viewport the canvas
 * was rendered with — its transform already folds in scale, rotation and the
 * Y flip). Returns null when the item carries no usable transform or the
 * resulting glyph height degenerates; the caller still injects an empty span
 * slot so item-index → span-index parity holds.
 */
export function fallbackSpanGeometry(
  item: Pick<PdfSearchTextItem, "transform" | "width">,
  viewportTransform: ReadonlyArray<number> | undefined,
): FallbackSpanGeometry | null {
  if (!isMatrix(viewportTransform) || !isMatrix(item?.transform)) return null;
  const tx = composeTransform(viewportTransform, item.transform);
  const fontSize = Math.hypot(tx[2], tx[3]);
  if (!isFiniteNumber(fontSize) || fontSize <= 0) return null;
  const ascent = fontSize * FALLBACK_ASCENT_RATIO;
  const angle = Math.atan2(tx[1], tx[0]);
  let left: number;
  let top: number;
  if (angle === 0) {
    left = tx[4];
    top = tx[5] - ascent;
  } else {
    // pdf.js #appendText: shift the anchor so the rotated row still hangs
    // from its baseline (transform-origin is the span's top-left corner).
    left = tx[4] + ascent * Math.sin(angle);
    top = tx[5] - ascent * Math.cos(angle);
  }
  if (!isFiniteNumber(left) || !isFiniteNumber(top)) return null;
  // The viewport scale is the length of the transform's X basis vector
  // (rotation-invariant); item.width is a user-space advance.
  const scale = Math.hypot(viewportTransform[0], viewportTransform[1]);
  const width = isFiniteNumber(item.width) && item.width > 0 ? item.width : 0;
  return {
    left,
    top,
    fontSize,
    angleDeg: (angle * 180) / Math.PI,
    targetWidth: width * scale,
  };
}

/**
 * Inject one invisible span per text item into the (cleared) text-layer
 * container. Previously-injected fallback nodes are removed first
 * (idempotent across zoom re-renders); pdf.js/OCR spans are never touched.
 * Items without usable geometry (or with empty text) still get a
 * zero-footprint span so the find mapper's item-index alignment holds.
 * `hasEOL` items are followed by a `<br>` (like pdf.js) so `innerText`
 * keeps line structure — `querySelectorAll("span")` ignores it. Returns the
 * number of spans injected.
 */
export function renderFallbackTextLayer(
  container: HTMLElement,
  items: ReadonlyArray<PdfSearchTextItem>,
  viewportTransform: ReadonlyArray<number> | undefined,
): number {
  for (const stale of Array.from(
    container.querySelectorAll(`[${FALLBACK_SPAN_ATTR}]`),
  )) {
    stale.remove();
  }
  // A page with NO text at all (scanned page) must stay span-free: the OCR
  // path injects its own spans into this container and relies on run-index
  // == span-index, so empty placeholder slots would shift its mapping.
  const hasAnyText = items.some(
    (item) => typeof item.str === "string" && item.str.length > 0,
  );
  if (!hasAnyText) return 0;
  let injected = 0;
  for (const item of items) {
    if (injected >= MAX_FALLBACK_SPANS_PER_PAGE) break;
    const span = container.ownerDocument.createElement("span");
    span.setAttribute(FALLBACK_SPAN_ATTR, "");
    span.textContent = typeof item.str === "string" ? item.str : "";
    const geometry = span.textContent
      ? fallbackSpanGeometry(item, viewportTransform)
      : null;
    if (geometry) {
      span.style.left = `${geometry.left}px`;
      span.style.top = `${geometry.top}px`;
      span.style.fontSize = `${geometry.fontSize}px`;
      span.style.fontFamily = "sans-serif";
    } else {
      // No usable geometry: keep the span slot but give it no footprint —
      // the find measurer skips zero-size rects rather than mis-painting.
      span.style.left = "0px";
      span.style.top = "0px";
      span.style.fontSize = "0px";
    }
    container.appendChild(span);
    injected += 1;
    if (geometry) {
      // Fit the rendered glyph run onto the item's advance width. Measure
      // AFTER the append so layout knows the font; a 0 measurement (jsdom,
      // hidden page) leaves the span unscaled instead of collapsing it.
      const measured = span.getBoundingClientRect().width;
      const scaleX = ocrSpanScaleX(geometry.targetWidth, measured);
      const transforms: string[] = [];
      if (geometry.angleDeg !== 0) {
        transforms.push(`rotate(${geometry.angleDeg}deg)`);
      }
      if (scaleX !== 1) {
        transforms.push(`scaleX(${scaleX})`);
      }
      // Always set: the text layer's stylesheet scales pdf.js's own spans
      // by 1 / --min-font-size (their font is set that much larger), and in
      // a Safari with a minimum font size a fallback span left to it was
      // drawn a ninth of its size. Its edit frame's pixel is unscaled too.
      span.style.transform = transforms.length > 0 ? transforms.join(" ") : "none";
      span.style.setProperty("--pdf-text-frame-px", "1px");
    }
    if (item.hasEOL === true) {
      const br = container.ownerDocument.createElement("br");
      br.setAttribute(FALLBACK_SPAN_ATTR, "");
      container.appendChild(br);
    }
  }
  return injected;
}
