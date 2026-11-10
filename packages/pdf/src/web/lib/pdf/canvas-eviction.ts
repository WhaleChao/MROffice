/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Canvas memory eviction window for the PDF reader.
 *
 * Rendered pages each hold a full-resolution canvas backing store; on a
 * multi-hundred-page document, keeping every store alive accumulates
 * multi-GB of raster memory while scrolling. The pdf.js viewer bounds this
 * with `PDFPageViewBuffer` (web/pdf_viewer.mjs, `DEFAULT_CACHE_SIZE = 10`):
 * pages evicted from the buffer are destroyed via `PDFPageView#reset` →
 * `_resetCanvas`, which releases the backing store with
 * `canvas.width = canvas.height = 0`. This module adapts that behavior
 * (Apache-2.0, Mozilla pdf.js) to our React viewer: instead of an LRU
 * buffer we keep a distance window around the current page in VISIBLE page
 * order (deleted pages excluded, moves applied), which is the order the
 * user actually scrolls through in continuous / single / spread modes.
 *
 * The canvas backing store is evicted first; the text layer stays mounted
 * within TEXT_LAYER_EVICTION_WINDOW so search, selection, find highlights
 * and OCR spans keep working near the reader, and is released beyond it.
 * The existing lazy-render path repaints an evicted canvas, and rebuilds its
 * text layer, when it re-enters the viewport.
 */






export const CANVAS_EVICTION_WINDOW = 4;

/** Pages within this distance keep their text layer after their canvas is
 * released, so search, selection and find highlights on the pages around
 * the reader stand as before; farther ones release it too (a 111-page read
 * kept ~17k spans of pages long scrolled past) and rebuild it with their
 * canvas when they come back. */
export const TEXT_LAYER_EVICTION_WINDOW = 12;

/** Whether a page's text layer may be released: not one holding the focus,
 * the selection or an open edit (the reader is using it, wherever it is). */
export function pdfTextLayerReleasable(layer: Element, doc: Document = layer.ownerDocument): boolean {
  if (!layer.firstChild) return false;
  const active = doc.activeElement;
  if (active && active !== doc.body && layer.contains(active)) return false;
  const selection = doc.getSelection?.();
  if (selection && selection.rangeCount > 0 && !selection.isCollapsed &&
    (layer.contains(selection.anchorNode) || layer.contains(selection.focusNode))) return false;
  if (layer.querySelector("[data-pdf-native-text-edit-selected], [data-pdf-native-text-inline-input]")) return false;
  return true;
}

/**
 * The set of page numbers whose canvases must be KEPT: every page within
 * `windowSize` positions of the current page in `visiblePageOrder`.
 *
 * `visiblePageOrder` is the display order (original page numbers after
 * delete/move edits). When `currentPage` is not in the order (mid-edit
 * transient), the nearest page by page-number distance anchors the window
 * so a transient never evicts the whole neighborhood.
 */
export function computeCanvasKeepSet(
  visiblePageOrder: readonly number[],
  currentPage: number,
  windowSize: number = CANVAS_EVICTION_WINDOW,
): Set<number> {
  const keep = new Set<number>();
  if (visiblePageOrder.length === 0) return keep;
  if (!Number.isFinite(windowSize) || windowSize < 0) {
    windowSize = CANVAS_EVICTION_WINDOW;
  }
  let anchor = visiblePageOrder.indexOf(currentPage);
  if (anchor < 0) {
    let bestDistance = Infinity;
    anchor = 0;
    for (let i = 0; i < visiblePageOrder.length; i += 1) {
      const distance = Math.abs(visiblePageOrder[i] - currentPage);
      if (distance < bestDistance) {
        bestDistance = distance;
        anchor = i;
      }
    }
  }
  const from = Math.max(0, anchor - windowSize);
  const to = Math.min(visiblePageOrder.length - 1, anchor + windowSize);
  for (let i = from; i <= to; i += 1) {
    keep.add(visiblePageOrder[i]);
  }
  return keep;
}

/**
 * Filter the rendered-page set down to the pages that are safe to evict:
 * outside the keep window AND not currently painting (zeroing a canvas
 * mid-`page.render` interleaves pdf.js transform state → garbled bitmap).
 */
export function selectEvictablePages(
  renderedPages: Iterable<number>,
  keepSet: ReadonlySet<number>,
  inFlightRenders: { has(pageNumber: number): boolean },
): number[] {
  const evictable: number[] = [];
  for (const pageNumber of renderedPages) {
    if (keepSet.has(pageNumber)) continue;
    if (inFlightRenders.has(pageNumber)) continue;
    evictable.push(pageNumber);
  }
  return evictable;
}
