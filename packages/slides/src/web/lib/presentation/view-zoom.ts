/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * View state of the presentation editor: view mode and live-view types, the
 * zoom bounds and slider mapping, and the view aids (ruler, gridlines, guides)
 * and where they are available.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import {
  PRESENTATION_VIEW_ZOOM_MAX_PERCENT,
  PRESENTATION_VIEW_ZOOM_MIN_PERCENT,
  type PresentationCapabilityContext,
  type PresentationViewAidKind,
} from "@mrnote/core/web/lib/presentation-feature-capabilities.ts";
import type { PresentationViewMode as LegacyPresentationViewMode } from "../presentation-sorter";
import { EDITOR_ZOOM_DEFAULT, type EditorZoomBounds } from "@mrnote/core/web/lib/editor-zoom.ts";

export type PresentationViewMode = LegacyPresentationViewMode;













export type PresentationLiveView =
  | "normal"
  | "outline"
  | "sorter"
  | "notes"
  | "reading"
  | "master";

export const PRESENTATION_VIEW_ZOOM_BOUNDS: EditorZoomBounds = Object.freeze({
  min: PRESENTATION_VIEW_ZOOM_MIN_PERCENT / 100,
  max: PRESENTATION_VIEW_ZOOM_MAX_PERCENT / 100,
  fallback: EDITOR_ZOOM_DEFAULT,
});

/** View > Zoom's fixed magnifications, PowerPoint's order (its Zoom dialog
 *  lists Fit first, then these). */
export const PRESENTATION_ZOOM_DIALOG_PRESETS = [400, 200, 100, 66, 50, 33] as const;

export const PRESENTATION_FIT_VIEWPORT_INSET_PX = 32;

const PRESENTATION_ZOOM_RATIO_PRECISION = 10_000;

export const PRESENTATION_ZOOM_KEYBOARD_STEP_PERCENT = 0.01;

export const PRESENTATION_ZOOM_KEYBOARD_PAGE_PERCENT = 10;

export type PresentationViewportZoomMode = "percent" | "fit";

export type PresentationZoomPopoverAnchor = "ribbon" | "status";

export function clampPresentationViewportZoom(zoom: number): number {
  if (!Number.isFinite(zoom)) return PRESENTATION_VIEW_ZOOM_BOUNDS.fallback;
  const bounded = Math.min(
    PRESENTATION_VIEW_ZOOM_BOUNDS.max,
    Math.max(PRESENTATION_VIEW_ZOOM_BOUNDS.min, zoom),
  );
  return (
    Math.round(bounded * PRESENTATION_ZOOM_RATIO_PRECISION) /
    PRESENTATION_ZOOM_RATIO_PRECISION
  );
}

export function presentationZoomPositionFromPercent(
  value: number,
  minimum: number,
  maximum: number,
): number {
  if (!Number.isFinite(value) || maximum <= minimum) return 0;
  return Math.min(1, Math.max(0, (value - minimum) / (maximum - minimum)));
}

export function presentationZoomPercentFromPosition(
  position: number,
  minimum: number,
  maximum: number,
): number {
  if (!Number.isFinite(position) || maximum <= minimum) return minimum;
  const value =
    minimum + Math.min(1, Math.max(0, position)) * (maximum - minimum);
  return Number(value.toFixed(2));
}

export type PresentationViewAidState = Readonly<
  Record<PresentationViewAidKind, boolean>
>;

export const INITIAL_PRESENTATION_VIEW_AIDS: PresentationViewAidState = Object.freeze({
  ruler: false,
  gridlines: false,
  guides: false,
});

// The feature's capability predicate is `always`; selection/editability are not
// authority inputs. Keep a stable neutral context so no render snapshot can
// accidentally become part of the event-time view-state decision.
export const PRESENTATION_VIEW_AID_CAPABILITY_CONTEXT: PresentationCapabilityContext =
  Object.freeze({
    editable: true,
    hasSelection: false,
    hasTextSelection: false,
    hasShapeSelection: false,
    hasClipboard: false,
    online: true,
  });

export function isPresentationViewAidNormalMode(
  viewMode: PresentationViewMode,
  presentMode: boolean,
): boolean {
  return viewMode === "rail" && !presentMode;
}

/**
 * Whether a vertical wheel of `deltaY` still scrolls a zoomed-in slide
 * canvas (PowerPoint's Normal view past Fit): it does until the canvas has
 * reached its end in that direction, and only then does the wheel turn the
 * slide. A canvas that does not overflow never scrolls.
 */
export function presentationCanvasScrollsToward(
  scrollTop: number,
  clientHeight: number,
  scrollHeight: number,
  deltaY: number,
): boolean {
  if (!(scrollHeight > clientHeight + 1) || deltaY === 0) return false;
  return deltaY > 0
    ? scrollTop + clientHeight < scrollHeight - 1
    : scrollTop > 0;
}

/** The canvas's breathing room around the slide at Fit (its padding, a side). */
export const PRESENTATION_CANVAS_PADDING_PX = PRESENTATION_FIT_VIEWPORT_INSET_PX / 2;
/** The ruler's band, fixed along the canvas's top and left edges. */
export const PRESENTATION_RULER_BAND_PX = 18;

/**
 * Fit to Window: the largest whole percent at which the whole slide (its
 * actual size, `slideWidth` x `slideHeight` CSS px) fits the canvas less its
 * room -- `insetX` / `insetY` across both sides, the ruler's band included.
 * Rounded down, so the slide at Fit never overflows the window it fits (a
 * zoom rounded up drew the slide's edges a few px under the canvas's), and
 * the status bar reads a whole percent as PowerPoint's does.
 */
export function presentationFitZoom(
  viewportWidth: number,
  viewportHeight: number,
  slideWidth: number,
  slideHeight: number,
  insetX: number,
  insetY: number,
): number {
  const width = viewportWidth - Math.max(0, insetX);
  const height = viewportHeight - Math.max(0, insetY);
  if (
    !(Number.isFinite(width) && width > 0) ||
    !(Number.isFinite(height) && height > 0) ||
    !(Number.isFinite(slideWidth) && slideWidth > 0) ||
    !(Number.isFinite(slideHeight) && slideHeight > 0)
  ) {
    return PRESENTATION_VIEW_ZOOM_BOUNDS.fallback;
  }
  const ratio = Math.min(width / slideWidth, height / slideHeight);
  const percent = Math.floor(ratio * 100 + 1e-6);
  return Math.min(
    PRESENTATION_VIEW_ZOOM_BOUNDS.max,
    Math.max(PRESENTATION_VIEW_ZOOM_BOUNDS.min, percent / 100),
  );
}

/**
 * PowerPoint's zoom in / out (the status bar's + and -, Ctrl/Cmd+plus and
 * minus): the next whole tenth up or down, within 10%–400%.
 */
export function presentationSteppedZoom(zoom: number, direction: 1 | -1): number {
  const tenths = Number.isFinite(zoom) ? zoom * 10 : 10;
  const next =
    direction > 0
      ? Math.floor(tenths + 1e-6) + 1
      : Math.ceil(tenths - 1e-6) - 1;
  return Math.min(
    PRESENTATION_VIEW_ZOOM_BOUNDS.max,
    Math.max(PRESENTATION_VIEW_ZOOM_BOUNDS.min, next / 10),
  );
}

export interface PresentationZoomScrollAnchor {
  /** The anchored point as a share of the slide frame on screen (0–1 inside it). */
  readonly fx: number;
  readonly fy: number;
  /** The anchored point from the viewport's visible content corner (inside its border), in CSS px. */
  readonly viewX: number;
  readonly viewY: number;
}

/**
 * The scroll offsets that put the slide point an anchor captured back under
 * the same spot of the viewport after the zoom changed: the viewport's
 * centre for the slider, the Zoom dialog and the keyboard, as PowerPoint
 * zooms them. `frame*` is the slide frame's box after the zoom, from the
 * scrolled content's origin (scroll offset included), in CSS px. The browser
 * clamps the result to the scroll range.
 */
export function presentationZoomAnchoredScroll(
  anchor: PresentationZoomScrollAnchor,
  frameLeft: number,
  frameTop: number,
  frameWidth: number,
  frameHeight: number,
): { scrollLeft: number; scrollTop: number } {
  return {
    scrollLeft: Math.max(0, frameLeft + anchor.fx * frameWidth - anchor.viewX),
    scrollTop: Math.max(0, frameTop + anchor.fy * frameHeight - anchor.viewY),
  };
}
