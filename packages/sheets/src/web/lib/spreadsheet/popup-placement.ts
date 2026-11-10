/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a grid popup sits so it stays inside the window.
 *
 * Excel opens its in-grid menus at the thing they belong to — the cursor for a
 * right-click menu, the smart tag's button for Paste Options / Auto Fill
 * Options / the AutoCorrect button — and then keeps the whole menu on screen:
 * near the bottom edge it opens UPWARDS, near the right edge it opens to the
 * LEFT, and it never hangs off the window with its last commands out of reach.
 *
 * One algorithm for all of them. The cell context menu (SpreadsheetWindow) and
 * the options button (SpreadsheetOptionsButton) call this, so a menu opened at
 * the bottom of the grid behaves the same wherever it came from.
 */

export interface PopupAxisPlacement {
  /** The popup's leading edge when there is room for it there. */
  start: number;
  /** How far the popup extends along this axis. */
  size: number;
  /** The window's extent along this axis. */
  viewport: number;
  /**
   * When the popup is flipped to the other side, its TRAILING edge lands here.
   * For a menu opened AT a point (a right-click) that is the point itself; for
   * one opened BELOW/BESIDE an anchor it is the anchor's opposite edge, so the
   * flipped menu sits above (or before) the button instead of over it.
   * Defaults to `start`.
   */
  flipEnd?: number;
  /** Kept clear of the window edge. Excel leaves a hair of room; we use 8px. */
  margin?: number;
}

export const POPUP_VIEWPORT_MARGIN = 8;

/**
 * The popup's leading edge on one axis: flip to the other side of the anchor
 * when it would overflow, then slide back inside the margin either way. A
 * popup larger than the window is pinned at the margin (the caller caps its
 * size with `popupMaxSize` and lets it scroll).
 */
export function placePopupAxis({
  start,
  size,
  viewport,
  flipEnd,
  margin = POPUP_VIEWPORT_MARGIN,
}: PopupAxisPlacement): number {
  const fitsForward = start + size <= viewport - margin;
  if (fitsForward) return Math.max(margin, start);
  const flipped = (flipEnd ?? start) - size;
  // Room on the other side: open there (Excel opens upwards near the bottom).
  if (flipped >= margin) return flipped;
  // Room on neither side: stay on the requested side, slid back just far
  // enough to fit, so the menu still starts near its anchor instead of
  // jumping to the far edge of the window. A popup larger than the window is
  // pinned at the margin (the caller caps its size and lets it scroll).
  return Math.max(margin, viewport - size - margin);
}

/** The largest a popup may be on one axis before it has to scroll. */
export function popupMaxSize(
  viewport: number,
  margin = POPUP_VIEWPORT_MARGIN,
): number {
  return Math.max(0, viewport - margin * 2);
}

export interface PopupPlacementInput {
  /** The popup's preferred top-left corner, in viewport coordinates. */
  x: number;
  y: number;
  width: number;
  height: number;
  viewportWidth: number;
  viewportHeight: number;
  /** Where the trailing edge lands when flipped; see `PopupAxisPlacement`. */
  flipEndX?: number;
  flipEndY?: number;
  margin?: number;
}

/** Both axes of `placePopupAxis`, in viewport coordinates. */
export function placePopupInViewport({
  x,
  y,
  width,
  height,
  viewportWidth,
  viewportHeight,
  flipEndX,
  flipEndY,
  margin = POPUP_VIEWPORT_MARGIN,
}: PopupPlacementInput): { left: number; top: number } {
  return {
    left: placePopupAxis({
      start: x,
      size: width,
      viewport: viewportWidth,
      flipEnd: flipEndX,
      margin,
    }),
    top: placePopupAxis({
      start: y,
      size: height,
      viewport: viewportHeight,
      flipEnd: flipEndY,
      margin,
    }),
  };
}
