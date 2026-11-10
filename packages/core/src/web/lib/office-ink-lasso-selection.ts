/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Lasso Select, after the loop is drawn: what stays on screen, the same
 * in Docs, the deck and the worksheet (components/office-ribbon/
 * OfficeInkLassoSelection draws it).
 *
 * - The loop the pen drew stays, as a dashed outline, closed: its last point
 *   is joined to its first when the pen lifts.
 * - One frame encloses every stroke the loop took (the union of their boxes),
 *   instead of a frame per stroke: the strokes were picked together and move,
 *   scale and delete as one unit.
 * - The loop follows the strokes: it is kept relative to the frame it was
 *   drawn around, so a move or a scale of the strokes carries it along.
 *
 * Geometry only; each editor measures its strokes in its own unit (slide EMU,
 * sheet pixels, client pixels) and draws in that unit.
 */

export interface InkLassoPoint {
  readonly x: number;
  readonly y: number;
}

export interface InkLassoBox {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

/** The corner grips of the shared frame. */
export const INK_LASSO_HANDLES = ["nw", "ne", "se", "sw"] as const;
export type InkLassoHandle = (typeof INK_LASSO_HANDLES)[number];

/**
 * The loop as drawn, closed: the last point joins the first. Repeated
 * points (a pen resting) are dropped; fewer than three distinct points is no
 * loop.
 */
export function closeInkLassoLoop(points: readonly InkLassoPoint[]): InkLassoPoint[] {
  const loop: InkLassoPoint[] = [];
  for (const point of points) {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
    const previous = loop[loop.length - 1];
    if (previous && previous.x === point.x && previous.y === point.y) continue;
    loop.push({ x: point.x, y: point.y });
  }
  if (loop.length < 3) return [];
  const first = loop[0]!;
  const last = loop[loop.length - 1]!;
  if (first.x !== last.x || first.y !== last.y) loop.push({ x: first.x, y: first.y });
  return loop;
}

/** The SVG path of a closed loop (`Z` joins its ends). */
export function inkLassoLoopPathData(points: readonly InkLassoPoint[]): string {
  if (points.length < 3) return "";
  const round = (value: number) => Math.round(value * 100) / 100;
  return `M${points.map((point) => `${round(point.x)} ${round(point.y)}`).join("L")}Z`;
}

/** The box enclosing every box given, or null for none. */
export function inkLassoUnionBox(boxes: readonly InkLassoBox[]): InkLassoBox | null {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const box of boxes) {
    if (![box.left, box.top, box.width, box.height].every(Number.isFinite)) continue;
    left = Math.min(left, box.left);
    top = Math.min(top, box.top);
    right = Math.max(right, box.left + box.width);
    bottom = Math.max(bottom, box.top + box.height);
  }
  return Number.isFinite(left) ? { left, top, width: right - left, height: bottom - top } : null;
}

/**
 * The loop kept relative to the frame it was drawn around: each point as a
 * fraction of that frame (beyond 0..1 where the loop runs outside it).
 */
export interface InkLassoAnchoredLoop {
  readonly points: readonly InkLassoPoint[];
}

export function anchorInkLassoLoop(
  loop: readonly InkLassoPoint[],
  box: InkLassoBox,
): InkLassoAnchoredLoop | null {
  const closed = closeInkLassoLoop(loop);
  if (closed.length === 0) return null;
  const width = Math.max(box.width, 1e-6);
  const height = Math.max(box.height, 1e-6);
  return {
    points: closed.map((point) => ({
      x: (point.x - box.left) / width,
      y: (point.y - box.top) / height,
    })),
  };
}

/** The anchored loop around the frame where the strokes are now. */
export function placeInkLassoLoop(
  anchored: InkLassoAnchoredLoop | null,
  box: InkLassoBox,
): InkLassoPoint[] {
  if (!anchored) return [];
  return anchored.points.map((point) => ({
    x: box.left + point.x * box.width,
    y: box.top + point.y * box.height,
  }));
}

/** The same ids, in any order. */
export function sameInkLassoIds<T>(a: ReadonlySet<T> | readonly T[], b: ReadonlySet<T> | readonly T[]): boolean {
  const left = a instanceof Set ? a : new Set(a as readonly T[]);
  const right = b instanceof Set ? b : new Set(b as readonly T[]);
  if (left.size !== right.size) return false;
  for (const id of left) if (!right.has(id)) return false;
  return true;
}

/** The box moved by an offset. */
export function translateInkLassoBox(box: InkLassoBox, dx: number, dy: number): InkLassoBox {
  return { left: box.left + dx, top: box.top + dy, width: box.width, height: box.height };
}

/**
 * Where a box lands when the frame it sits in goes from `from` to `to` (the
 * strokes scale together about the frame's pinned corner).
 */
export function mapInkLassoBox(box: InkLassoBox, from: InkLassoBox, to: InkLassoBox): InkLassoBox {
  const scaleX = from.width > 0 ? to.width / from.width : 1;
  const scaleY = from.height > 0 ? to.height / from.height : 1;
  return {
    left: to.left + (box.left - from.left) * scaleX,
    top: to.top + (box.top - from.top) * scaleY,
    width: box.width * scaleX,
    height: box.height * scaleY,
  };
}

/**
 * Follow one pointer from a press on the frame (its inside or a grip) until
 * it lifts: `onMove` gets the offset from the press in client pixels, `onEnd`
 * the last one (`cancelled` when the browser took the pointer or Escape was
 * pressed). The element pressed captures the pointer, so a pen or a finger
 * that leaves it keeps dragging.
 */
export function followInkLassoPointer(
  event: {
    readonly pointerId: number;
    readonly clientX: number;
    readonly clientY: number;
    readonly currentTarget: EventTarget | null;
  },
  handlers: {
    readonly onMove: (dx: number, dy: number, moveEvent: PointerEvent) => void;
    readonly onEnd: (dx: number, dy: number, cancelled: boolean) => void;
  },
): () => void {
  const { pointerId } = event;
  const startX = event.clientX;
  const startY = event.clientY;
  let dx = 0;
  let dy = 0;
  const target = event.currentTarget instanceof Element ? event.currentTarget : null;
  try {
    target?.setPointerCapture?.(pointerId);
  } catch {
    // A pointer that already lifted cannot be captured; the window still hears it.
  }
  const view = target?.ownerDocument.defaultView ?? window;
  let done = false;
  const finish = (cancelled: boolean) => {
    if (done) return;
    done = true;
    view.removeEventListener("pointermove", move, true);
    view.removeEventListener("pointerup", up, true);
    view.removeEventListener("pointercancel", cancel, true);
    view.removeEventListener("keydown", key, true);
    handlers.onEnd(dx, dy, cancelled);
  };
  const move = (moveEvent: PointerEvent) => {
    if (moveEvent.pointerId !== pointerId) return;
    dx = moveEvent.clientX - startX;
    dy = moveEvent.clientY - startY;
    handlers.onMove(dx, dy, moveEvent);
  };
  const up = (upEvent: PointerEvent) => {
    if (upEvent.pointerId !== pointerId) return;
    finish(false);
  };
  const cancel = (cancelEvent: PointerEvent) => {
    if (cancelEvent.pointerId !== pointerId) return;
    finish(true);
  };
  const key = (keyEvent: KeyboardEvent) => {
    if (keyEvent.key !== "Escape") return;
    keyEvent.preventDefault();
    keyEvent.stopPropagation();
    finish(true);
  };
  view.addEventListener("pointermove", move, true);
  view.addEventListener("pointerup", up, true);
  view.addEventListener("pointercancel", cancel, true);
  view.addEventListener("keydown", key, true);
  return () => finish(true);
}

/** The attribute a stroke picked by the lasso wears: its own frame and grips stand down. */
export const INK_LASSO_MEMBER_ATTRIBUTE = "data-ink-lasso-member";
/** The attribute on the shared overlay's elements: a host's own press handlers let them be. */
export const INK_LASSO_OVERLAY_ATTRIBUTE = "data-office-ink-lasso-selection";

/** True when a press landed on the shared frame (its inside, a grip or its bar). */
export function isInkLassoOverlayTarget(target: EventTarget | null): boolean {
  return (
    typeof Element !== "undefined" &&
    target instanceof Element &&
    target.closest(`[${INK_LASSO_OVERLAY_ATTRIBUTE}]`) !== null
  );
}

/**
 * Pen Gestures ▸ Circle to Lasso: a pen stroke that closes on itself -- its
 * ends meet and it turns once around its middle -- offers to select what it
 * encloses (the pen menu's switch, lib/office-ink-pen-gestures). A line, a
 * hook, a "C" or a scribble is not a circle, nor is a loop too small to hold
 * anything (`minSize`, in the points' unit).
 */
export function isInkCircleGesture(points: readonly InkLassoPoint[], minSize = 24): boolean {
  if (points.length < 8) return false;
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  let cx = 0;
  let cy = 0;
  for (const point of points) {
    left = Math.min(left, point.x);
    right = Math.max(right, point.x);
    top = Math.min(top, point.y);
    bottom = Math.max(bottom, point.y);
    cx += point.x;
    cy += point.y;
  }
  const width = right - left;
  const height = bottom - top;
  if (width < minSize || height < minSize) return false;
  const first = points[0]!;
  const last = points[points.length - 1]!;
  if (Math.hypot(last.x - first.x, last.y - first.y) > Math.max(minSize * 0.75, 0.3 * Math.min(width, height))) return false;
  cx /= points.length;
  cy /= points.length;
  // Once around the middle, and only once: a scribble turns back and forth.
  let turned = 0;
  let travelled = 0;
  let previous = Math.atan2(first.y - cy, first.x - cx);
  for (let index = 1; index < points.length; index += 1) {
    const point = points[index]!;
    const angle = Math.atan2(point.y - cy, point.x - cx);
    let step = angle - previous;
    if (step > Math.PI) step -= 2 * Math.PI;
    else if (step < -Math.PI) step += 2 * Math.PI;
    turned += step;
    travelled += Math.abs(step);
    previous = angle;
  }
  const turns = Math.abs(turned) / (2 * Math.PI);
  return turns >= 0.8 && turns <= 1.6 && travelled <= Math.abs(turned) * 1.35;
}
