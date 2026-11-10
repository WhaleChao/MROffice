/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PDF viewer view-mode + zoom math for the reader window.
 *
 * Two concerns live here, both pure so they can be unit-tested without
 * pdf.js / jsdom:
 *
 *  1. Zoom math — `fitWidthScale` and `fitPageScale` compute the pdf.js render
 *     scale that makes a page exactly fit the viewport width, or fit a whole
 *     page inside the viewport (both width AND height), respectively. The
 *     component used to inline the fit-width math; it now delegates here so the
 *     same clamping rules cover the new fit-page button too.
 *
 *  2. View mode — the reader can show every page stacked (CONTINUOUS scroll,
 *     the historical behaviour) or one page at a time (SINGLE page). The mode
 *     is ephemeral UI state (like zoom / current page), so it is not persisted
 *     in the page schema; these helpers just describe the mode set + the
 *     toggle transition.
 */

/** Minimum render scale the viewer permits. Mirrors PDFViewer's MIN_SCALE. */
export const PDF_MIN_SCALE = 0.5;






export const PDF_FIT_MIN_SCALE = 0.1;
/**
 * Maximum render scale the viewer permits. Mirrors PDFViewer's MAX_SCALE.
 * 640% — the viewer anchors the scroll position across zoom changes and caps
 * canvas backing pixels (see lib/pdf/zoom-anchor.ts), so deep zoom is usable.
 */
export const PDF_MAX_SCALE = 6.4;

/**
 * Horizontal padding (px) the page list reserves inside the scroll container.
 * Subtracted before fitting so a fit-width page does not touch the edges or
 * trigger a horizontal scrollbar. Matches the CSS padding on
 * `.pdf-viewer__pages` (20px left + 16px right ≈ a 40px budget).
 */
export const PDF_FIT_WIDTH_PADDING = 40;

/**
 * Vertical padding (px) reserved when fitting a whole page into the viewport
 * height, so a fit-page page leaves a little breathing room top + bottom and
 * never clips against the toolbar / find bar. Matches the CSS top/bottom
 * padding on `.pdf-viewer__pages` (20px top + 32px bottom ≈ a 52px budget).
 */
export const PDF_FIT_PAGE_PADDING = 52;

/**
 * Horizontal gap (px) between the two pages of a spread row. Mirrors the
 * `column-gap` on `.pdf-viewer__pages--spread`; the fit-width budget must
 * subtract it before halving so two fitted pages + the gap exactly fill the
 * viewport.
 */
export const PDF_SPREAD_COLUMN_GAP = 16;

/**
 * Sticky fit modes for the reader: after "fit width" / "fit page" the fit is
 * re-applied on container resize (and layout-mode changes) until the user
 * zooms manually, which clears the mode. `null` (no active fit) is modeled at
 * the use site.
 */
export const PDF_FIT_MODES = ["width", "page"] as const;
export type PdfFitMode = (typeof PDF_FIT_MODES)[number];

/**
 * The width budget (px) available to ONE page when fitting, given the layout
 * mode: the viewport width minus the horizontal padding, halved (minus the
 * column gap) in spread mode where two pages share each row. Feed the result
 * to {@link fitWidthScale} / {@link fitPageScale} with `widthPadding = 0`
 * (the padding is already taken here). Degenerate viewport → 0 (the fit
 * helpers then fall back to the minimum scale).
 */
export function fitWidthBudget(
  viewportWidth: number,
  mode: PdfViewMode,
  padding: number = PDF_FIT_WIDTH_PADDING,
  gap: number = PDF_SPREAD_COLUMN_GAP,
): number {
  const vw = Number(viewportWidth);
  if (!Number.isFinite(vw) || vw <= 0) return 0;
  const pad = Number.isFinite(padding) && padding >= 0 ? padding : 0;
  const g = Number.isFinite(gap) && gap >= 0 ? gap : 0;
  let usable = vw - pad;
  if (mode === "spread") usable = (usable - g) / 2;
  return Math.max(0, usable);
}

/**
 * Available reader layout modes.
 *  - `continuous`: every page stacked, scroll through (the historical default).
 *  - `single`: one page at a time.
 *  - `spread`: TWO pages side-by-side per row (facing-pages / book view), the
 *    rest stacked below — complements continuous/single, matching Acrobat's
 *    "Two Page View" and WPS's 双页 layout.
 */
export const PDF_VIEW_MODES = ["continuous", "single", "spread"] as const;
export type PdfViewMode = (typeof PDF_VIEW_MODES)[number];

/** The historical / default layout: all pages stacked, scroll through. */
export const PDF_DEFAULT_VIEW_MODE: PdfViewMode = "continuous";

/**
 * Clamp a raw render scale into the viewer's permitted range and round to two
 * decimals so the displayed "%" stays stable. A non-finite / non-positive
 * input falls back to the minimum scale rather than producing a 0-area canvas.
 */
export function clampScale(
  scale: number,
  min: number = PDF_MIN_SCALE,
  max: number = PDF_MAX_SCALE,
): number {
  const value = Number(scale);
  const lo = Number.isFinite(min) ? min : PDF_MIN_SCALE;
  const hi = Number.isFinite(max) ? max : PDF_MAX_SCALE;
  if (!Number.isFinite(value) || value <= 0) return lo;
  return Math.min(hi, Math.max(lo, +value.toFixed(2)));
}

/**
 * Render scale that makes a page exactly fill the viewport width (minus
 * padding). Returns a clamped scale; a degenerate viewport / page width falls
 * back to the minimum scale so the page is still visible.
 *
 * @param viewportWidth Inner width (px) of the page scroll container.
 * @param pageWidth     Full-size page width (pdf.js viewport at scale 1.0).
 * @param padding       Horizontal padding budget to subtract before fitting.
 */
export function fitWidthScale(
  viewportWidth: number,
  pageWidth: number,
  padding: number = PDF_FIT_WIDTH_PADDING,
): number {
  const vw = Number(viewportWidth);
  const pw = Number(pageWidth);
  const pad = Number.isFinite(padding) && padding >= 0 ? padding : 0;
  if (!Number.isFinite(vw) || !Number.isFinite(pw) || pw <= 0) {
    return PDF_MIN_SCALE;
  }
  const usable = vw - pad;
  if (usable <= 0) return PDF_MIN_SCALE;
  return clampScale(usable / pw, PDF_FIT_MIN_SCALE);
}

/**
 * Render scale that fits a WHOLE page inside the viewport — constrained by both
 * the available width AND height, whichever is the tighter limit — so the page
 * is fully visible without scrolling. Returns a clamped scale; a degenerate
 * viewport / page size falls back to the minimum scale.
 *
 * @param viewportWidth  Inner width (px) of the page scroll container.
 * @param viewportHeight Inner height (px) of the page scroll container.
 * @param pageWidth      Full-size page width (pdf.js viewport at scale 1.0).
 * @param pageHeight     Full-size page height (pdf.js viewport at scale 1.0).
 * @param widthPadding   Horizontal padding budget to subtract before fitting.
 * @param heightPadding  Vertical padding budget to subtract before fitting.
 */
export function fitPageScale(
  viewportWidth: number,
  viewportHeight: number,
  pageWidth: number,
  pageHeight: number,
  widthPadding: number = PDF_FIT_WIDTH_PADDING,
  heightPadding: number = PDF_FIT_PAGE_PADDING,
): number {
  const vw = Number(viewportWidth);
  const vh = Number(viewportHeight);
  const pw = Number(pageWidth);
  const ph = Number(pageHeight);
  if (
    !Number.isFinite(vw) ||
    !Number.isFinite(vh) ||
    !Number.isFinite(pw) ||
    !Number.isFinite(ph) ||
    pw <= 0 ||
    ph <= 0
  ) {
    return PDF_MIN_SCALE;
  }
  const wPad = Number.isFinite(widthPadding) && widthPadding >= 0 ? widthPadding : 0;
  const hPad = Number.isFinite(heightPadding) && heightPadding >= 0 ? heightPadding : 0;
  const usableW = vw - wPad;
  const usableH = vh - hPad;
  if (usableW <= 0 || usableH <= 0) return PDF_MIN_SCALE;
  // The whole page must fit on both axes, so take the more restrictive scale.
  const scale = Math.min(usableW / pw, usableH / ph);
  return clampScale(scale, PDF_FIT_MIN_SCALE);
}

// --------------------------------------------------------------------------
// Zoom-percentage presets (Acrobat / WPS "Actual Size" + 50/75/100/150/200%).
//
// The viewer's render scale is a multiplier (1.0 == 100%); the toolbar shows
// `Math.round(scale * 100)%`. These helpers translate between a chosen PRESET
// percentage and the clamped render scale, and pick the preset to mark "active"
// for an arbitrary current scale (so a fit-width / step-zoomed value still shows
// the nearest standard percentage as selected).
// --------------------------------------------------------------------------

/**
 * Standard zoom percentages offered as clickable presets, in display order.
 * 100 (== "Actual Size") is the canonical reset target. All lie within the
 * [{@link PDF_MIN_SCALE}, {@link PDF_MAX_SCALE}] render-scale range (50%..200%
 * ⊂ 50%..640%).
 */
export const ZOOM_PRESET_PERCENTS = [50, 75, 100, 150, 200] as const;

/** The "Actual Size" preset percentage (100% == render scale 1.0). */
export const ZOOM_ACTUAL_SIZE_PERCENT = 100;

/**
 * Render scale (clamped) for a chosen preset percentage. `100` → scale 1.0.
 * A non-finite / non-positive percent falls back to the floor scale. The result
 * is clamped + 2-decimal-rounded by {@link clampScale} so it composes with the
 * existing zoom controls.
 */
export function zoomPresetToScale(percent: number): number {
  const p = Number(percent);
  if (!Number.isFinite(p) || p <= 0) return PDF_MIN_SCALE;
  return clampScale(p / 100);
}

/**
 * The integer percentage a render scale displays as (matching the toolbar's
 * `Math.round(scale * 100)`). Degenerate input → 0.
 */
export function scaleToPercent(scale: number): number {
  const s = Number(scale);
  if (!Number.isFinite(s) || s <= 0) return 0;
  return Math.round(s * 100);
}

/**
 * The preset percentage CLOSEST to the current render scale — used to highlight
 * the active preset in the dropdown even when the live scale is an off-preset
 * value (a fit-width result, a +/- step). Ties resolve to the SMALLER preset
 * (deterministic). Returns null only for a degenerate scale.
 */
export function nearestZoomPreset(scale: number): number | null {
  const pct = scaleToPercent(scale);
  if (pct <= 0) return null;
  let best: number = ZOOM_PRESET_PERCENTS[0];
  let bestDelta = Math.abs(pct - best);
  for (const preset of ZOOM_PRESET_PERCENTS) {
    const delta = Math.abs(pct - preset);
    if (delta < bestDelta) {
      best = preset;
      bestDelta = delta;
    }
  }
  return best;
}

/**
 * Whether a render scale EXACTLY matches a given preset percentage (within the
 * display rounding) — used to show a check / `aria-current` only on a precisely
 * selected preset, distinct from {@link nearestZoomPreset}'s fuzzy highlight.
 */
export function isExactZoomPreset(scale: number, percent: number): boolean {
  return scaleToPercent(scale) === Math.round(Number(percent));
}

/** Whether a value is one of the offered preset percentages. */
export function isZoomPreset(value: unknown): boolean {
  const n = Number(value);
  return (
    Number.isFinite(n) &&
    (ZOOM_PRESET_PERCENTS as readonly number[]).includes(Math.round(n))
  );
}

/** Whether a value is a recognised reader view mode. */
export function isPdfViewMode(value: unknown): value is PdfViewMode {
  return (
    typeof value === "string" && (PDF_VIEW_MODES as readonly string[]).includes(value)
  );
}

/**
 * The mode you land on when cycling the layout button. The order walks the full
 * {@link PDF_VIEW_MODES} ring forward (continuous → single → spread →
 * continuous), so a single control rotates through every layout. An
 * unrecognised current mode resolves to the default first so the cycle is always
 * well-defined.
 */
export function nextViewMode(current: PdfViewMode): PdfViewMode {
  const mode = isPdfViewMode(current) ? current : PDF_DEFAULT_VIEW_MODE;
  const i = PDF_VIEW_MODES.indexOf(mode);
  return PDF_VIEW_MODES[(i + 1) % PDF_VIEW_MODES.length];
}

/**
 * The `scrollIntoView` behaviour to use when navigating to a target page in a
 * given view mode.
 *
 * Single-page mode only mounts-visible the current page (the rest are
 * `display:none` until the `currentPage` state change re-renders), so a
 * navigation must (a) switch the page first and (b) let an effect perform the
 * scroll AFTER the new page paints. A smooth animation across what is, post-
 * render, the only visible page reads as a no-op jitter, so single mode uses an
 * instant (`"auto"`) scroll to bring the freshly-shown page to the top.
 *
 * Continuous AND spread modes keep every page stacked/visible (spread merely
 * lays two per row), so a `"smooth"` scroll animates the viewport to the
 * already-rendered target — the historical behaviour.
 */
export function resolveScrollBehavior(mode: PdfViewMode, restoringView = false): ScrollBehavior {
  // A restored position is not navigation through the intervening pages.
  // `auto` can inherit CSS smooth scrolling, so request a true instant jump.
  if (restoringView) return "instant";
  return mode === "single" ? "auto" : "smooth";
}

/** A deferred scroll-into-view request: a target page plus a monotonic nonce. */
export interface PdfScrollRequest {
  page: number;
  nonce: number;
}

/**
 * Reducer for clearing a consumed scroll request without dropping a newer one
 * that was queued between the scroll effect being scheduled and run.
 *
 * The PDF viewer defers `goToPage`'s scroll to an effect so it fires AFTER the
 * render reflects the new current page (single-page mode `display:none`s every
 * page but the current one, so a synchronous scroll would target a hidden, not-
 * yet-laid-out element). Once the effect scrolls it must clear the request so an
 * unrelated later re-render doesn't re-scroll — but if `goToPage` fired again in
 * the meantime (a higher nonce), that fresher request must survive so its scroll
 * isn't swallowed.
 *
 * @param current The latest request in state (may be a newer one than handled).
 * @param handledNonce The nonce of the request the effect just scrolled to.
 * @returns `null` when the handled request is still current (clear it), else the
 *   newer request untouched.
 */
export function consumeScrollRequest<T extends PdfScrollRequest>(
  current: T | null,
  handledNonce: number,
): T | null {
  return current && current.nonce === handledNonce ? null : current;
}

// --------------------------------------------------------------------------
// Two-page / facing-pages (spread) pairing.
//
// Spread mode lays the pages out two-per-row (a book / facing-pages view). The
// component renders the SAME flat page list it always does; these pure helpers
// decide which visible-list positions begin a new row so the layout can insert
// a line-break before them (a CSS `flex-basis: 100%` spacer, or a row wrapper).
// Kept here, pure, so the pairing logic is unit-testable without the pdf.js /
// flex layout — mirroring the zoom / view-mode helpers above.
// --------------------------------------------------------------------------

/**
 * The 0-based position WITHIN A ROW of the page at flat list index `index`,
 * under spread (two-up) layout.
 *
 * When `coverAlone` is true the FIRST page sits by itself on row 0 (matching how
 * a book's cover faces outward, the common reader convention), so pairs become
 * (—,1) (2,3) (4,5)…: index 0 → col 0; index 1 → col 0; index 2 → col 1; etc.
 * When false the pages pair from the start: (0,1) (2,3)… so even indexes are the
 * left page (col 0) and odd indexes the right (col 1).
 *
 * `count` is the total number of pages; it bounds the result so an out-of-range
 * index returns 0 (a degenerate fallback rather than throwing). A non-finite /
 * negative index also returns 0.
 */
export function spreadColumn(
  index: number,
  count: number,
  coverAlone = false,
): 0 | 1 {
  if (!Number.isFinite(index) || index < 0) return 0;
  const i = Math.trunc(index);
  // The cover (index 0) is always alone on its own row when coverAlone.
  if (coverAlone) {
    if (i === 0) return 0;
    // Pages 1.. pair as (1,2)(3,4)…: shift by one so parity flips correctly.
    return (i - 1) % 2 === 0 ? 0 : 1;
  }
  return i % 2 === 0 ? 0 : 1;
}

/**
 * Whether the page at flat list index `index` STARTS a new spread row (i.e. it
 * is the LEFT page of a pair, column 0). The layout inserts a row break before
 * such pages so each row holds at most two pages. Index 0 always starts a row.
 *
 * This is the predicate the renderer uses: a page that starts a row gets a
 * `data-spread-row-start` flag (CSS forces it onto a fresh flex line); the page
 * after it (column 1) flows beside it. Mirrors {@link spreadColumn} === 0.
 */
export function isSpreadRowStart(
  index: number,
  count: number,
  coverAlone = false,
): boolean {
  return spreadColumn(index, count, coverAlone) === 0;
}

/**
 * Group a flat list of items (typically the visible pages) into spread ROWS of
 * at most two, honoring the `coverAlone` convention (a lone first row when set).
 * Returns an array of rows, each a 1- or 2-length slice of the input in order.
 * An empty / non-array input yields `[]`. Pure and allocation-light; handy for a
 * renderer that prefers explicit row wrappers over CSS line-break spacers, and
 * trivially unit-testable.
 */
export function spreadRows<T>(
  items: ReadonlyArray<T>,
  coverAlone = false,
): T[][] {
  if (!Array.isArray(items) || items.length === 0) return [];
  const rows: T[][] = [];
  let i = 0;
  // A lone cover row first when requested (and there is a cover to show).
  if (coverAlone) {
    rows.push([items[0]]);
    i = 1;
  }
  for (; i < items.length; i += 2) {
    const row = items[i + 1] !== undefined ? [items[i], items[i + 1]] : [items[i]];
    rows.push(row);
  }
  return rows;
}
