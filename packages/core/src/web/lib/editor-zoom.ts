/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Editor zoom helper — shared pure logic for the Document, Spreadsheet, and
 * Presentation editor zoom controls (status-bar ± buttons + Ctrl/Cmd+scroll).
 *
 * Mirror of the CODE editor's font-scale approach (lib/code-runner/editor-prefs)
 * but expressed as a CSS `transform: scale()` multiplier rather than a font-size
 * multiplier, because the three target editors are rich canvas surfaces whose
 * entire pixel geometry must scale together.
 *
 * Framework-free + unit-testable: all exports are pure functions with no DOM
 * or React dependencies.
 */

/** Zoom level as a CSS scale multiplier.  1.0 = 100% (default). */
export type ZoomLevel = number;

export const EDITOR_ZOOM_MIN: ZoomLevel = 0.5; // 50%
export const EDITOR_ZOOM_MAX: ZoomLevel = 2.0; // 200%
export const EDITOR_ZOOM_DEFAULT: ZoomLevel = 1.0;
export const EDITOR_ZOOM_STEP: ZoomLevel = 0.1; // one step per click or wheel tick

export interface EditorZoomBounds {
  readonly min: ZoomLevel;
  readonly max: ZoomLevel;
  readonly fallback: ZoomLevel;
}

/** Default bounds retained by the Document and Spreadsheet editors. */
export const EDITOR_ZOOM_BOUNDS: EditorZoomBounds = Object.freeze({
  min: EDITOR_ZOOM_MIN,
  max: EDITOR_ZOOM_MAX,
  fallback: EDITOR_ZOOM_DEFAULT,
});

/**
 * Multiple Pages is an auto-fitted desktop view, so it may temporarily use a
 * smaller scale than the manual zoom controls. The desktop surface is sealed
 * below the mobile breakpoint; 25% still covers the narrowest mounted desktop
 * spread while leaving every manual +/- and pinch path on the 50% minimum.
 */
export const MULTIPLE_PAGES_FIT_ZOOM_BOUNDS: EditorZoomBounds = Object.freeze({
  min: 0.25,
  max: EDITOR_ZOOM_MAX,
  fallback: EDITOR_ZOOM_DEFAULT,
});

function validZoomBounds(bounds: EditorZoomBounds): boolean {
  return (
    Number.isFinite(bounds.min) &&
    bounds.min > 0 &&
    Number.isFinite(bounds.max) &&
    bounds.max >= bounds.min &&
    Number.isFinite(bounds.fallback) &&
    bounds.fallback >= bounds.min &&
    bounds.fallback <= bounds.max
  );
}

function resolvedZoomBounds(bounds: EditorZoomBounds): EditorZoomBounds {
  return validZoomBounds(bounds) ? bounds : EDITOR_ZOOM_BOUNDS;
}

/** Clamp fitted geometry without rounding upward and overflowing its viewport. */
function clampFittedZoom(
  value: number,
  bounds: EditorZoomBounds,
): ZoomLevel {
  const resolved = resolvedZoomBounds(bounds);
  if (!Number.isFinite(value)) return resolved.fallback;
  return Math.min(resolved.max, Math.max(resolved.min, value));
}

/**
 * Clamp `value` to the resolved editor bounds and snap to two decimal places
 * so repeated step arithmetic cannot accumulate float drift. Non-finite input
 * (NaN / ±Infinity) collapses to that editor's fallback.
 */
export function clampZoom(
  value: number,
  bounds: EditorZoomBounds = EDITOR_ZOOM_BOUNDS,
): ZoomLevel {
  const resolved = resolvedZoomBounds(bounds);
  if (!Number.isFinite(value)) return resolved.fallback;
  const bounded = Math.min(resolved.max, Math.max(resolved.min, value));
  // Snap to 0.01 precision so 0.5999999 reads back as 0.60.
  return Math.round(bounded * 100) / 100;
}

/**
 * Return the next zoom level after a step of `deltaSteps` increments
 * (positive = zoom in, negative = zoom out) from `current`, clamped.
 */
export function stepZoom(
  current: ZoomLevel,
  deltaSteps: number,
  bounds: EditorZoomBounds = EDITOR_ZOOM_BOUNDS,
): ZoomLevel {
  return clampZoom(current + deltaSteps * EDITOR_ZOOM_STEP, bounds);
}

/**
 * Map a wheel event's `deltaY` (with Ctrl/Cmd held) to a signed step count.
 * One standard notch (≈ ±100 px) maps to ±1 step; trackpads produce smaller
 * fractional deltas that are quantised to the nearest step (minimum ±1 for any
 * non-zero delta). The sign is inverted so "scroll up" (negative deltaY)
 * zooms IN, matching every browser zoom behaviour.
 *
 * Returns 0 when `deltaY` is exactly 0 (no-op).
 */
export function wheelDeltaToZoomSteps(deltaY: number): number {
  if (deltaY === 0) return 0;
  // 100 px ≈ one notch; scale to step units and round away from zero (so 1.5
  // → 2 not 1), but guarantee at least ±1 so even the smallest trackpad nudge
  // still moves.  Math.round(-1.5) === -1 in JS (rounds toward +∞), so we use
  // Math.floor for negatives to get the expected "half rounds away from zero"
  // semantics consistent with the positive branch.
  const raw = -(deltaY / 100);
  if (raw > 0) return Math.max(1, Math.round(raw));
  if (raw < 0) return Math.min(-1, Math.floor(raw));
  return 0;
}

/**
 * One wheel notch in CSS pixels.  Browsers report the same physical notch in
 * three different units depending on `deltaMode`, and a trackpad reports many
 * small deltas instead of one big one; normalising here is what lets a pinch
 * and a Ctrl+scroll share one code path.
 */
export function wheelDeltaPixels(deltaY: number, deltaMode = 0): number {
  if (!Number.isFinite(deltaY)) return 0;
  if (deltaMode === 1) return deltaY * 16; // DOM_DELTA_LINE, ~1 line
  if (deltaMode === 2) return deltaY * 400; // DOM_DELTA_PAGE, ~1 viewport
  return deltaY;
}

/**
 * Rate that makes one mouse notch (≈100 px) worth one `EDITOR_ZOOM_STEP`,
 * while a trackpad's small deltas move proportionally less.
 */
const PINCH_ZOOM_RATE = Math.log(1 + EDITOR_ZOOM_STEP) / 100;

/**
 * Zoom continuously from a pinch or Ctrl+scroll.
 *
 * The result is deliberately NOT snapped to 0.01: a trackpad pinch arrives as
 * a stream of sub-pixel deltas, and snapping each one would round every single
 * nudge back to where it started, so the gesture would do nothing at all. The
 * caller keeps this exact value as the live gesture level and snaps only what
 * it shows and stores.
 *
 * Growth is multiplicative, so zooming out then back in returns to the same
 * place and a step feels the same size at 50% as it does at 200%.
 */
export function pinchZoomLevel(level: ZoomLevel, deltaPixels: number): number {
  if (!Number.isFinite(level)) return EDITOR_ZOOM_DEFAULT;
  if (!Number.isFinite(deltaPixels) || deltaPixels === 0) {
    return Math.min(EDITOR_ZOOM_MAX, Math.max(EDITOR_ZOOM_MIN, level));
  }
  const next = level * Math.exp(-deltaPixels * PINCH_ZOOM_RATE);
  return Math.min(EDITOR_ZOOM_MAX, Math.max(EDITOR_ZOOM_MIN, next));
}

/** Apply a Safari pinch's absolute `event.scale` to the level it started at. */
export function gestureZoomLevel(startLevel: ZoomLevel, scale: number): number {
  if (!Number.isFinite(startLevel)) return EDITOR_ZOOM_DEFAULT;
  if (!Number.isFinite(scale) || scale <= 0) {
    return Math.min(EDITOR_ZOOM_MAX, Math.max(EDITOR_ZOOM_MIN, startLevel));
  }
  return Math.min(
    EDITOR_ZOOM_MAX,
    Math.max(EDITOR_ZOOM_MIN, startLevel * scale),
  );
}

export interface ZoomAnchorInput {
  scrollLeft: number;
  scrollTop: number;
  /** Pointer position relative to the scrolling viewport's top-left corner. */
  pointerX: number;
  pointerY: number;
  from: ZoomLevel;
  to: ZoomLevel;
}

/**
 * Scroll offsets that keep the content under the pointer under the pointer.
 *
 * Without this, zooming walks the sheet away from wherever you were looking —
 * you pinch on a cell in column M and end up staring at column A. Assumes the
 * scaled surface uses `transform-origin: top left`, which is what makes the
 * content coordinate a plain `(scroll + pointer) / level`.
 */
export function zoomAnchoredScroll(input: ZoomAnchorInput): {
  scrollLeft: number;
  scrollTop: number;
} {
  const { scrollLeft, scrollTop, pointerX, pointerY, from, to } = input;
  if (!Number.isFinite(from) || from <= 0 || !Number.isFinite(to) || to <= 0) {
    return { scrollLeft, scrollTop };
  }
  const ratio = to / from;
  return {
    scrollLeft: Math.max(0, (scrollLeft + pointerX) * ratio - pointerX),
    scrollTop: Math.max(0, (scrollTop + pointerY) * ratio - pointerY),
  };
}

/**
 * The CSS `transform` string to apply to the editor content surface for the
 * given zoom level.  At 100% the string is `"none"` so the DOM is byte-for-byte
 * unchanged and no compositing layer is created.
 */
export function zoomTransform(level: ZoomLevel): string {
  if (level === EDITOR_ZOOM_DEFAULT) return "none";
  return `scale(${level})`;
}

/**
 * Map a zoom percent onto the shared continuous zoom crown's unit interval
 * (components/ui/ContinuousZoomDial): linear across the manual range, as the
 * workbook's status-bar dial reads it.
 */
export function editorZoomPositionFromPercent(
  percent: number,
  minimum: number,
  maximum: number,
): number {
  if (!Number.isFinite(percent) || maximum <= minimum) return 0;
  return (
    (Math.min(maximum, Math.max(minimum, percent)) - minimum) /
    (maximum - minimum)
  );
}

/** The crown's unit position back to a zoom percent (two decimals). */
export function editorZoomPercentFromPosition(
  position: number,
  minimum: number,
  maximum: number,
): number {
  const bounded = Number.isFinite(position)
    ? Math.min(1, Math.max(0, position))
    : 0;
  return Number((minimum + bounded * (maximum - minimum)).toFixed(2));
}

/**
 * Format the zoom level as a human-readable percentage string (e.g. "75%").
 */
export function zoomPercent(level: ZoomLevel): string {
  return `${Math.round(level * 100)}%`;
}

function positiveFinite(value: number): number | null {
  return Number.isFinite(value) && value > 0 ? value : null;
}

export interface PhysicalPageViewportAnchor<T> {
  readonly identity: T;
  readonly index: number;
}

/**
 * Capture the selected physical page by both stable identity and source index.
 * Identity survives a CSS-only grid reflow; index is the bounded fallback when
 * React/ProseMirror has replaced the corresponding DOM node before restore.
 */
export function capturePhysicalPageViewportAnchor<T>(
  pages: readonly T[],
  selectedPage: T | null | undefined,
): PhysicalPageViewportAnchor<T> | null {
  if (pages.length === 0) return null;
  const selectedIndex = selectedPage ? pages.indexOf(selectedPage) : -1;
  const index = selectedIndex >= 0 ? selectedIndex : 0;
  const identity = pages[index];
  return identity === undefined ? null : { identity, index };
}

/** Resolve the exact captured page after layout reflow without changing focus. */
export function resolvePhysicalPageViewportAnchor<T>(
  pages: readonly T[],
  anchor: PhysicalPageViewportAnchor<T>,
): T | null {
  const identityIndex = pages.indexOf(anchor.identity);
  if (identityIndex >= 0) return pages[identityIndex] ?? null;
  return pages[anchor.index] ?? null;
}

/**
 * Fit one physical paper width inside the available editor viewport.
 *
 * Both inputs are unscaled CSS pixels. `viewportInset` reserves the canvas
 * breathing room around the paper; invalid geometry falls back to the resolved
 * editor default rather than producing an infinite/zero transform.
 */
export function fitPageWidthZoom(
  viewportWidth: number,
  pageWidth: number,
  viewportInset = 48,
  bounds: EditorZoomBounds = EDITOR_ZOOM_BOUNDS,
): ZoomLevel {
  const resolved = resolvedZoomBounds(bounds);
  const available = positiveFinite(viewportWidth - Math.max(0, viewportInset));
  const paper = positiveFinite(pageWidth);
  if (available === null || paper === null) return resolved.fallback;
  return clampZoom(available / paper, resolved);
}

/**
 * Fit one complete physical paper inside the available editor viewport.
 * Width and height use the same scale so page geometry never distorts.
 */
export function fitOnePageZoom(
  viewportWidth: number,
  viewportHeight: number,
  pageWidth: number,
  pageHeight: number,
  viewportInset = 48,
  bounds: EditorZoomBounds = EDITOR_ZOOM_BOUNDS,
): ZoomLevel {
  const resolved = resolvedZoomBounds(bounds);
  const inset = Math.max(0, viewportInset);
  const availableWidth = positiveFinite(viewportWidth - inset);
  const availableHeight = positiveFinite(viewportHeight - inset);
  const paperWidth = positiveFinite(pageWidth);
  const paperHeight = positiveFinite(pageHeight);
  if (
    availableWidth === null ||
    availableHeight === null ||
    paperWidth === null ||
    paperHeight === null
  ) {
    return resolved.fallback;
  }
  return clampZoom(
    Math.min(availableWidth / paperWidth, availableHeight / paperHeight),
    resolved,
  );
}

/**
 * Fit a two-page reading spread inside the available editor viewport.
 *
 * The paper boxes retain their physical dimensions; only the shared visual
 * scale changes. `pageGap` is the unscaled gutter between the facing pages,
 * so resizing the window cannot distort either page or silently change the
 * document model.
 */
export function fitMultiplePagesZoom(
  viewportWidth: number,
  viewportHeight: number,
  pageWidth: number,
  pageHeight: number,
  pageGap = 28,
  viewportInset = 48,
  bounds: EditorZoomBounds = MULTIPLE_PAGES_FIT_ZOOM_BOUNDS,
): ZoomLevel {
  const resolved = resolvedZoomBounds(bounds);
  const inset = Math.max(0, viewportInset);
  const availableWidth = positiveFinite(viewportWidth - inset);
  const availableHeight = positiveFinite(viewportHeight - inset);
  const paperWidth = positiveFinite(pageWidth);
  const paperHeight = positiveFinite(pageHeight);
  if (
    availableWidth === null ||
    availableHeight === null ||
    paperWidth === null ||
    paperHeight === null ||
    !Number.isFinite(pageGap) ||
    pageGap < 0
  ) {
    return resolved.fallback;
  }
  const spreadWidth = paperWidth * 2 + pageGap;
  return clampFittedZoom(
    Math.min(availableWidth / spreadWidth, availableHeight / paperHeight),
    resolved,
  );
}
