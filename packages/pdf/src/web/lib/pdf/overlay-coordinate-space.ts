/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Canonical pointer geometry for editable PDF overlays.
 *
 * Overlay records use a top-left page coordinate system (fractions or
 * percentages). PDF.js PageViewport remains the authority for turning PDF
 * user-space coordinates into the rotated/cropped display viewport; pointer
 * gestures then normalize against the viewport's *rendered* client extent.
 * Measuring that extent at gesture start keeps client-pixel deltas correct
 * when the notebook canvas applies an ancestor CSS transform.
 *
 * This is MRAI orchestration built around public PDF.js viewport semantics; it
 * does not copy PDF.js transformation code.
 */

export interface OverlayCoordinateSize {
  width: number;
  height: number;
}

export interface OverlayClientPoint {
  clientX: number;
  clientY: number;
}

export interface OverlayCoordinateDelta {
  x: number;
  y: number;
}

/** Layout/font measurements need untransformed CSS pixels. Pointer gestures
 * still use resolveOverlayCoordinateSize with the rendered client rectangle. */
export function readOverlayLocalSize(
  element: HTMLElement | null | undefined,
  fallback: OverlayCoordinateSize,
): OverlayCoordinateSize {
  const computed =
    element?.ownerDocument.defaultView?.getComputedStyle(element);
  const width =
    Number.parseFloat(computed?.width ?? "") || element?.clientWidth || 0;
  const height =
    Number.parseFloat(computed?.height ?? "") || element?.clientHeight || 0;
  return resolveOverlayCoordinateSize({ width, height }, fallback);
}

function isPositiveFinite(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

/**
 * Prefer the live rendered extent and fall back to the logical viewport size.
 * The one-pixel final fallback matches the legacy drag hook's divide-by-one
 * safety without allowing NaN/Infinity to enter persisted overlay geometry.
 */
export function resolveOverlayCoordinateSize(
  rendered: OverlayCoordinateSize | null | undefined,
  logical: OverlayCoordinateSize,
): OverlayCoordinateSize {
  if (
    rendered &&
    isPositiveFinite(rendered.width) &&
    isPositiveFinite(rendered.height)
  ) {
    return { width: rendered.width, height: rendered.height };
  }
  return {
    width: isPositiveFinite(logical.width) ? logical.width : 1,
    height: isPositiveFinite(logical.height) ? logical.height : 1,
  };
}

/** Convert a client-pixel gesture delta into canonical page units. */
export function overlayDragDelta(
  start: OverlayClientPoint,
  current: OverlayClientPoint,
  coordinateSize: OverlayCoordinateSize,
  unitsPerPage: number,
): OverlayCoordinateDelta {
  if (
    !Number.isFinite(start.clientX) ||
    !Number.isFinite(start.clientY) ||
    !Number.isFinite(current.clientX) ||
    !Number.isFinite(current.clientY) ||
    !isPositiveFinite(unitsPerPage)
  ) {
    return { x: 0, y: 0 };
  }
  const size = resolveOverlayCoordinateSize(null, coordinateSize);
  return {
    x: ((current.clientX - start.clientX) / size.width) * unitsPerPage,
    y: ((current.clientY - start.clientY) / size.height) * unitsPerPage,
  };
}
