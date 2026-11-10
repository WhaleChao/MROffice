/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure "which page owns the indicator" rule for the PDF reader.
 *
 * The viewer historically promoted a page to "current" when its
 * IntersectionObserver ratio crossed 0.5 — which breaks past ~2× zoom: a page
 * wrapper taller than twice the viewport can never have half of itself
 * visible, so the indicator froze while the reader scrolled through the
 * document. The replacement rule is VIEWPORT-MIDPOINT OWNERSHIP: the current
 * page is the one whose wrapper vertically spans `scrollTop + clientHeight/2`
 * (pdf.js's `getVisibleElements` "first page crossing the middle" spirit,
 * Apache-2.0). When the midpoint falls in a gap between wrappers, the nearest
 * wrapper wins so the indicator never goes blank.
 *
 * DOM-free: the component measures wrapper rects and feeds plain spans.
 */

/** A page wrapper's vertical extent in scroll-content coordinates. */
export interface PdfPageSpan {
  /** 1-based page number. */
  page: number;
  /** Wrapper top edge, px from the scroll content origin. */
  top: number;
  /** Wrapper height, px (non-positive spans are ignored — hidden pages). */
  height: number;
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/**
 * The page that OWNS the viewport midpoint, or `null` when no usable span was
 * supplied. A span containing the midpoint wins (first in input order when
 * several do — spread rows put two pages on one band); otherwise the span
 * whose nearest edge is closest to the midpoint wins (earlier input order
 * breaks ties). Zero/negative-height spans (display:none pages in single-page
 * mode) and malformed spans are skipped.
 */
export function midpointOwnedPage(
  spans: ReadonlyArray<PdfPageSpan>,
  scrollTop: number,
  clientHeight: number,
): number | null {
  if (!Array.isArray(spans) || spans.length === 0) return null;
  const top = isFiniteNumber(scrollTop) ? scrollTop : 0;
  const viewport =
    isFiniteNumber(clientHeight) && clientHeight > 0 ? clientHeight : 0;
  const midpoint = top + viewport / 2;

  let nearest: number | null = null;
  let nearestDistance = Infinity;
  for (const span of spans) {
    if (!span) continue;
    const { page, top: spanTop, height } = span;
    if (!isFiniteNumber(page) || !Number.isInteger(page) || page < 1) continue;
    if (!isFiniteNumber(spanTop) || !isFiniteNumber(height) || height <= 0) {
      continue;
    }
    const bottom = spanTop + height;
    if (midpoint >= spanTop && midpoint < bottom) return page;
    const distance =
      midpoint < spanTop ? spanTop - midpoint : midpoint - bottom;
    if (distance < nearestDistance) {
      nearestDistance = distance;
      nearest = page;
    }
  }
  return nearest;
}

/**
 * The first and last pages whose spans meet the viewport widened by `margin`
 * px above and below, or `null` when none does (no usable span). The reader
 * draws these pages in full and the others as their paper only.
 */
export function pagesAboutView(
  spans: ReadonlyArray<PdfPageSpan>,
  scrollTop: number,
  clientHeight: number,
  margin: number,
): { first: number; last: number } | null {
  if (!Array.isArray(spans) || spans.length === 0) return null;
  const viewTop = isFiniteNumber(scrollTop) ? scrollTop : 0;
  const viewHeight = isFiniteNumber(clientHeight) && clientHeight > 0 ? clientHeight : 0;
  const widen = isFiniteNumber(margin) && margin > 0 ? margin : 0;
  const top = viewTop - widen;
  const bottom = viewTop + viewHeight + widen;
  let first = Infinity;
  let last = -Infinity;
  for (const span of spans) {
    if (!span || !isFiniteNumber(span.page) || !isFiniteNumber(span.top) || !isFiniteNumber(span.height)) continue;
    if (span.height <= 0) continue;
    if (span.top + span.height < top || span.top > bottom) continue;
    first = Math.min(first, span.page);
    last = Math.max(last, span.page);
  }
  return first <= last ? { first, last } : null;
}

/**
 * The pages a document opens with in full before its pages are first
 * measured: from the first, as many as `heights` (each page's height on
 * screen, px) take to fill the view and a screen below it, as
 * {@link pagesAboutView} widens the view by a screen once they are measured.
 * Null when the view has no height yet (nothing is laid out): every page is
 * then drawn in full until it is measured, as before.
 *
 * The band used to be unknown until the frames had been laid out, so the
 * first frames of a 186-page report were all mounted in full -- every page's
 * overlays built and thrown away a moment later -- before the first one could
 * paint. The pages are counted without the gaps between them, so the guess
 * errs towards one page too many, never one too few on screen.
 */
export function pagesAboutDocumentStart(
  heights: ReadonlyArray<number>,
  clientHeight: number,
): { first: number; last: number } | null {
  if (!isFiniteNumber(clientHeight) || clientHeight <= 0 || heights.length === 0) return null;
  const reach = clientHeight * 2;
  let covered = 0;
  for (let index = 0; index < heights.length; index += 1) {
    const height = heights[index];
    // A page without a usable size: the guess stops being one.
    if (!isFiniteNumber(height) || height <= 0) return null;
    covered += height;
    if (covered >= reach) return { first: 1, last: index + 1 };
  }
  return { first: 1, last: heights.length };
}

/** Where the reader is: the page most of which is in view, and how far into
 * it the view's top edge is (a fraction of the page's height; negative when
 * the page starts below the top edge). */
export interface PdfReaderPageAnchor {
  page: number;
  into: number;
}

/**
 * The page most of which is in the viewport (the most of its height shown;
 * earlier input order breaks ties), as a {@link PdfReaderPageAnchor}, or
 * `null` when no usable span meets the viewport. A reload of the document on
 * screen puts the reader back here ({@link scrollTopForReaderPageAnchor}).
 */
export function readerPageAnchor(
  spans: ReadonlyArray<PdfPageSpan>,
  scrollTop: number,
  clientHeight: number,
): PdfReaderPageAnchor | null {
  if (!Array.isArray(spans) || spans.length === 0) return null;
  if (!isFiniteNumber(clientHeight) || clientHeight <= 0) return null;
  const viewTop = isFiniteNumber(scrollTop) ? scrollTop : 0;
  const viewBottom = viewTop + clientHeight;
  let best: PdfReaderPageAnchor | null = null;
  let bestShown = 0;
  for (const span of spans) {
    if (!span || !isFiniteNumber(span.page) || !Number.isInteger(span.page) || span.page < 1) continue;
    if (!isFiniteNumber(span.top) || !isFiniteNumber(span.height) || span.height <= 0) continue;
    const shown = Math.min(span.top + span.height, viewBottom) - Math.max(span.top, viewTop);
    if (shown > bestShown) {
      bestShown = shown;
      best = { page: span.page, into: (viewTop - span.top) / span.height };
    }
  }
  return best;
}

/** The scroll offset that shows `anchor` again: its page's top edge as far
 * above (or below) the view's top as it was. Null when that page has no
 * usable span. */
export function scrollTopForReaderPageAnchor(
  spans: ReadonlyArray<PdfPageSpan>,
  anchor: PdfReaderPageAnchor,
): number | null {
  if (!anchor || !isFiniteNumber(anchor.into)) return null;
  const span = Array.isArray(spans)
    ? spans.find((entry) => entry && entry.page === anchor.page &&
      isFiniteNumber(entry.top) && isFiniteNumber(entry.height) && entry.height > 0)
    : undefined;
  if (!span) return null;
  return Math.max(0, span.top + anchor.into * span.height);
}
