/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure math for MOVING / RESIZING placed PDF annotations (vector shapes,
 * visual signatures, redaction boxes, sticky notes).
 *
 * All geometry is in page-relative PERCENT units [0..100] with a TOP-LEFT
 * origin — the exact space the `pdf_shapes` / `pdf_signatures` /
 * `pdf_redactions` / `pdf_note` payloads persist (see shapes.ts /
 * signatures.ts / redaction.ts / sticky-notes.ts). Framework-free so the
 * drag-delta → clamped-geometry math is unit-testable in isolation; the
 * interaction layer (PdfAnnotationFrame / PdfNoteMarker) converts pointer
 * deltas to percent and applies these.
 *
 * Mirrors `text-box-geometry.ts` (the fraction-space [0..1] equivalent used by
 * the native text editor's boxes) but adds per-caller minimum sizes (each
 * annotation family persists its own floor) and the endpoint mapping a
 * two-point shape needs (a line / arrow must keep its drawn direction).
 */

import type { ResizeHandle } from "@mrnote/core/web/lib/pdf/text-box-geometry.ts";

/** A page-relative percent box [0..100], top-left origin. */
export interface PercentBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** What a drag gesture does to a box: translate it, or drag one handle. */
export type AnnotationTransformMode = "move" | ResizeHandle;

/**
 * Default minimum box size (percent). Above every family's persisted floor
 * (shapes reject drags under 0.5, redactions under 0.2, signatures clamp to 1)
 * so a resized box always survives its type's payload builder.
 */
export const ANNOTATION_MIN_SIZE = 1;

/** Geometry deltas below this (percent) are treated as "did not change". */
export const ANNOTATION_BOX_EPSILON = 0.01;

/** Two corners' extent below this is a degenerate axis (see endpoint mapping). */
const DEGENERATE_AXIS = 1e-6;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

/** Move a box by a percent delta, clamped so it stays fully on the page. */
export function moveAnnotationBox(
  box: PercentBox,
  dx: number,
  dy: number,
): PercentBox {
  const width = clamp(box.width, 0, 100);
  const height = clamp(box.height, 0, 100);
  return {
    x: clamp(box.x + dx, 0, 100 - width),
    y: clamp(box.y + dy, 0, 100 - height),
    width,
    height,
  };
}

/**
 * Resize a box by dragging a handle. Edge handles move one side; corner
 * handles move two. The box stays within [0,100] and never shrinks below the
 * given minimum on either axis.
 */
export function resizeAnnotationBox(
  box: PercentBox,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  minWidth = ANNOTATION_MIN_SIZE,
  minHeight = ANNOTATION_MIN_SIZE,
): PercentBox {
  const minW = clamp(minWidth, 0, 100);
  const minH = clamp(minHeight, 0, 100);
  let { x, y, width, height } = box;
  const right = x + width;
  const bottom = y + height;

  if (handle.includes("e")) {
    width = clamp(width + dx, minW, 100 - x);
  }
  if (handle.includes("w")) {
    const nextX = clamp(x + dx, 0, right - minW);
    width = right - nextX;
    x = nextX;
  }
  if (handle.includes("s")) {
    height = clamp(height + dy, minH, 100 - y);
  }
  if (handle.includes("n")) {
    const nextY = clamp(y + dy, 0, bottom - minH);
    height = bottom - nextY;
    y = nextY;
  }

  return { x, y, width, height };
}









export function scaleAnnotationBoxProportionally(
  box: PercentBox,
  scale: number,
  minSize = ANNOTATION_MIN_SIZE,
): PercentBox {
  if (!Number.isFinite(scale) || scale <= 0 || !(box.width > 0) || !(box.height > 0)) return box;
  const min = clamp(minSize, 0, 100);
  const smallest = Math.max(min / box.width, min / box.height);
  const largest = Math.min(100 / box.width, 100 / box.height);
  const applied = Math.min(Math.max(scale, smallest), largest);
  const width = box.width * applied;
  const height = box.height * applied;
  return {
    x: clamp(Math.min(box.x, 100 - width), 0, 100),
    y: clamp(Math.min(box.y, 100 - height), 0, 100),
    width,
    height,
  };
}

/** A box a resize may have carried across its own opposite edge. */
export interface MirrorablePercentBox extends PercentBox {
  /** The dragged handle crossed the opposite edge on that axis. */
  mirrorX?: boolean;
  mirrorY?: boolean;
}

/**
 * Resize by a handle that may be pulled past the opposite edge, as WPS lets a
 * shape be: the box then lies on the other side of that edge (kept at the
 * minimum size there) and the axis is marked mirrored.
 */
export function resizeAnnotationBoxAcross(
  box: PercentBox,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  minWidth = ANNOTATION_MIN_SIZE,
  minHeight = ANNOTATION_MIN_SIZE,
): MirrorablePercentBox {
  const minW = clamp(minWidth, 0, 100);
  const minH = clamp(minHeight, 0, 100);
  // The fixed edge and where the dragged one went, per axis.
  const axis = (
    start: number,
    extent: number,
    delta: number,
    fromEnd: boolean,
    minimum: number,
  ): { start: number; extent: number; mirrored: boolean } => {
    const fixed = fromEnd ? start : start + extent;
    const moved = clamp((fromEnd ? start + extent : start) + delta, 0, 100);
    const mirrored = fromEnd ? moved < fixed : moved > fixed;
    const low = Math.min(fixed, moved);
    const high = Math.max(fixed, moved);
    if (high - low >= minimum) return { start: low, extent: high - low, mirrored };
    // Too thin to show: keep the minimum on the side the handle is on.
    const towardEnd = fromEnd !== mirrored;
    const at = towardEnd ? clamp(fixed, 0, 100 - minimum) : clamp(fixed - minimum, 0, 100);
    return { start: at, extent: minimum, mirrored };
  };
  let { x, y, width, height } = box;
  let mirrorX = false;
  let mirrorY = false;
  if (handle.includes("e") || handle.includes("w")) {
    const next = axis(x, width, dx, handle.includes("e"), minW);
    x = next.start;
    width = next.extent;
    mirrorX = next.mirrored;
  }
  if (handle.includes("s") || handle.includes("n")) {
    const next = axis(y, height, dy, handle.includes("s"), minH);
    y = next.start;
    height = next.extent;
    mirrorY = next.mirrored;
  }
  return {
    x,
    y,
    width,
    height,
    ...(mirrorX ? { mirrorX } : {}),
    ...(mirrorY ? { mirrorY } : {}),
  };
}

/** How much of a box (percent) lies on the page along one axis. */
function extentOnPage(start: number, extent: number): number {
  return Math.max(0, Math.min(start + extent, 100) - Math.max(start, 0));
}

/**
 * How much of a box must stay on the page along one axis when it may reach
 * past it: `keep`, or all of it when it is smaller, and never more than the
 * page showed of it before the gesture (an object already mostly outside
 * does not jump back).
 */
function keptOnPage(keep: number, extent: number, shown: number): number {
  return Math.max(0, Math.min(keep, extent, shown));
}

/**
 * Move a box that may reach past the page, as WPS lets a shape: the drag is
 * stopped where only `keepX` / `keepY` of it (percent) would still be on the
 * page -- the part a pointer takes it back by.
 */
export function moveAnnotationBoxPartlyOffPage(
  box: PercentBox,
  dx: number,
  dy: number,
  keepX: number,
  keepY: number,
): PercentBox {
  const axis = (start: number, extent: number, delta: number, keep: number) => {
    const need = keptOnPage(keep, extent, extentOnPage(start, extent));
    return clamp(
      start + delta,
      Math.min(start, need - extent),
      Math.max(start, 100 - need),
    );
  };
  return {
    x: axis(box.x, box.width, dx, keepX),
    y: axis(box.y, box.height, dy, keepY),
    width: box.width,
    height: box.height,
  };
}

/**
 * Resize by a handle where the box may reach past the page (WPS): the
 * dragged edge goes past the page edge freely, while the box keeps
 * `keepX` / `keepY` of itself on the page (all of itself when smaller). With
 * `mirror` the edge may be pulled across the opposite one, as
 * resizeAnnotationBoxAcross allows; without it the box stops at its minimum.
 */
export function resizeAnnotationBoxPartlyOffPage(
  box: PercentBox,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  minWidth: number,
  minHeight: number,
  keepX: number,
  keepY: number,
  mirror = true,
): MirrorablePercentBox {
  const axis = (
    start: number,
    extent: number,
    delta: number,
    fromEnd: boolean,
    minimum: number,
    keep: number,
  ): { start: number; extent: number; mirrored: boolean } => {
    const fixed = fromEnd ? start : start + extent;
    const from = fromEnd ? start + extent : start;
    const shown = extentOnPage(start, extent);
    const place = (moved: number) => {
      const crossed = fromEnd ? moved < fixed : moved > fixed;
      const mirrored = mirror && crossed;
      const reach = mirror ? moved : fromEnd ? Math.max(moved, fixed + minimum) : Math.min(moved, fixed - minimum);
      const low = Math.min(fixed, reach);
      const high = Math.max(fixed, reach);
      if (high - low >= minimum) return { start: low, extent: high - low, mirrored };
      // Too thin to show: keep the minimum on the side the handle is on.
      const towardEnd = fromEnd !== mirrored;
      return { start: towardEnd ? fixed : fixed - minimum, extent: minimum, mirrored };
    };
    const holds = (next: { start: number; extent: number }) =>
      extentOnPage(next.start, next.extent) >=
      keptOnPage(keep, next.extent, shown) - 1e-9;
    const wanted = place(from + delta);
    if (holds(wanted)) return wanted;
    // Stop the edge where the box would last keep its part on the page.
    let good = from;
    let bad = from + delta;
    for (let step = 0; step < 40; step += 1) {
      const middle = (good + bad) / 2;
      if (holds(place(middle))) good = middle;
      else bad = middle;
    }
    return place(good);
  };
  let { x, y, width, height } = box;
  let mirrorX = false;
  let mirrorY = false;
  if (handle.includes("e") || handle.includes("w")) {
    const next = axis(x, width, dx, handle.includes("e"), Math.max(0, minWidth), keepX);
    x = next.start;
    width = next.extent;
    mirrorX = next.mirrored;
  }
  if (handle.includes("s") || handle.includes("n")) {
    const next = axis(y, height, dy, handle.includes("s"), Math.max(0, minHeight), keepY);
    y = next.start;
    height = next.extent;
    mirrorY = next.mirrored;
  }
  return {
    x,
    y,
    width,
    height,
    ...(mirrorX ? { mirrorX } : {}),
    ...(mirrorY ? { mirrorY } : {}),
  };
}

/** Dispatch a drag gesture to the move / resize math. */
export function transformAnnotationBox(
  box: PercentBox,
  mode: AnnotationTransformMode,
  dx: number,
  dy: number,
  minWidth = ANNOTATION_MIN_SIZE,
  minHeight = ANNOTATION_MIN_SIZE,
): PercentBox {
  return mode === "move"
    ? moveAnnotationBox(box, dx, dy)
    : resizeAnnotationBox(box, mode, dx, dy, minWidth, minHeight);
}

/** True when two boxes match within {@link ANNOTATION_BOX_EPSILON}. */
export function annotationBoxesEqual(
  a: PercentBox,
  b: PercentBox,
  epsilon = ANNOTATION_BOX_EPSILON,
): boolean {
  return (
    Math.abs(a.x - b.x) < epsilon &&
    Math.abs(a.y - b.y) < epsilon &&
    Math.abs(a.width - b.width) < epsilon &&
    Math.abs(a.height - b.height) < epsilon
  );
}

/**
 * Map a two-point shape's endpoints into a transformed bounding box, keeping
 * each endpoint's RELATIVE position inside the box. A rect / ellipse persisted
 * as (min,min)→(max,max) stays normalized; a line / arrow keeps its drawn
 * direction (a right-to-left arrow stays right-to-left after a resize). On a
 * degenerate axis (an axis-aligned line has ~0 extent on one axis) both
 * endpoints share the value, so they are centered on that axis of the new box
 * — the shape stays axis-aligned instead of picking up a spurious slant.
 */
export function applyBoxToShapePoints(
  points: {
    x1: number;
    y1: number;
    x2: number;
    y2: number;
    points?: readonly { x: number; y: number }[];
  },
  next: PercentBox,
): {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  points?: Array<{ x: number; y: number }>;
} {
  const left = Math.min(points.x1, points.x2);
  const top = Math.min(points.y1, points.y2);
  const width = Math.abs(points.x2 - points.x1);
  const height = Math.abs(points.y2 - points.y1);
  const mapX = (value: number): number => {
    const t = width > DEGENERATE_AXIS ? (value - left) / width : 0.5;
    return clamp(next.x + t * next.width, 0, 100);
  };
  const mapY = (value: number): number => {
    const t = height > DEGENERATE_AXIS ? (value - top) / height : 0.5;
    return clamp(next.y + t * next.height, 0, 100);
  };
  return {
    x1: mapX(points.x1),
    y1: mapY(points.y1),
    x2: mapX(points.x2),
    y2: mapY(points.y2),
    ...(points.points
      ? {
          points: points.points.map((point) => ({
            x: mapX(point.x),
            y: mapY(point.y),
          })),
        }
      : null),
  };
}

/**
 * Move a point annotation by a percent delta, optionally keeping its visual
 * marker inset from every page edge.
 */
export function moveAnnotationPoint(
  point: { x: number; y: number },
  dx: number,
  dy: number,
  inset = 0,
): { x: number; y: number } {
  const safeInset = Number.isFinite(inset) ? clamp(inset, 0, 50) : 0;
  const originX = Number.isFinite(point.x) ? point.x : safeInset;
  const originY = Number.isFinite(point.y) ? point.y : safeInset;
  const safeDx = Number.isFinite(dx) ? dx : 0;
  const safeDy = Number.isFinite(dy) ? dy : 0;
  return {
    x: clamp(originX + safeDx, safeInset, 100 - safeInset),
    y: clamp(originY + safeDy, safeInset, 100 - safeInset),
  };
}
