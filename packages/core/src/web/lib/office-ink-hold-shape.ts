/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw and hold: a stroke drawn and then held still at its end snaps to a
 * clean shape, as Notability, GoodNotes and Apple Notes do it. Notability:
 * "Hold your finger or stylus down after drawing a shape to have it
 * automatically detected" (circles, ellipses, triangles, squares,
 * rectangles, pentagons, hexagons, arrows), and "Draw a line and hold to
 * automatically straighten it"; Apple Notes: lifted at once, the stroke
 * stays as drawn.
 *
 * While the pen stays down the snapped shape follows it: a line keeps its
 * start and its end goes where the pen is (longer, shorter, turned); any
 * other shape scales about its centre with the pen's distance from it, never
 * turning. The pen's width never changes. On the lift the host lands the
 * shape as a real shape (its corners editable), not as ink.
 *
 * Host-neutral: points are in any one pixel space (y down).
 */

import { fitInkStrokeToShape, type InkFitPoint, type InkShapeFit } from "./office-ink-shape-fit";
import { inkKeptSampleIndexes } from "./office-ink-smooth";

/** How long the pen rests at the stroke's end before the stroke snaps (ms). */
export const INK_HOLD_SHAPE_MS = 450;
/** A rest timer this late (ms) found the page busy: the pen's moves of that time are still queued. */
const INK_HOLD_REST_LATE_MS = 60;

/** A wait for the pen to rest; `cancel` when it moves on or lifts. */
export interface InkHoldRestWait {
  cancel(): void;
}











export function waitForInkHoldRest(onRest: () => void, ms: number = INK_HOLD_SHAPE_MS): InkHoldRestWait {
  const now = () => (typeof performance !== "undefined" ? performance.now() : Date.now());
  const nextFrame =
    typeof requestAnimationFrame === "function"
      ? (run: () => void) => requestAnimationFrame(run)
      : (run: () => void) => setTimeout(run, 16) as unknown as number;
  const dropFrame = (id: number) =>
    typeof cancelAnimationFrame === "function" ? cancelAnimationFrame(id) : clearTimeout(id);
  let cancelled = false;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let frame: number | null = null;
  const arm = () => {
    const armedAt = now();
    timer = setTimeout(() => {
      timer = null;
      if (cancelled) return;
      if (now() - armedAt - ms > INK_HOLD_REST_LATE_MS) {
        arm();
        return;
      }
      frame = nextFrame(() => {
        frame = nextFrame(() => {
          frame = null;
          if (!cancelled) onRest();
        });
      });
    }, ms);
  };
  arm();
  return {
    cancel() {
      cancelled = true;
      if (timer !== null) clearTimeout(timer);
      if (frame !== null) dropFrame(frame);
    },
  };
}

/** How far (screen px) the pen may tremble while it rests and still snap. */
export const INK_HOLD_SHAPE_STILL_PX = 4;
/** A stroke shorter than this (screen px along it) is writing: it never snaps. */
export const INK_HOLD_SHAPE_MIN_LENGTH_PX = 30;

const distance = (a: InkFitPoint, b: InkFitPoint) => Math.hypot(a.x - b.x, a.y - b.y);

/** An open stroke straightens only when its path is at most this many times its start-to-end distance. */
const MAX_LINE_WANDER = 1.5;
/** ...and its ends are at least this far apart (px). */
const MIN_LINE_CHORD_PX = 30;

function boxOf(points: readonly InkFitPoint[]) {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const point of points) {
    minX = Math.min(minX, point.x);
    minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x);
    maxY = Math.max(maxY, point.y);
  }
  return { left: minX, top: minY, width: maxX - minX, height: maxY - minY };
}

export function inkPathLength(points: readonly InkFitPoint[]): number {
  let total = 0;
  for (let index = 1; index < points.length; index += 1) total += distance(points[index - 1]!, points[index]!);
  return total;
}

/**
 * The shape a held stroke snaps to. The shape recognizer's verdict first
 * (Ink to Shape's); where it is not sure -- a stroke held is a stroke the
 * writer asked to be a shape -- an open stroke straightens into the line
 * from its start to where the pen rests, and a closed one becomes the
 * polygon of its corners.
 */
export function inkHoldShapeFit(points: readonly InkFitPoint[]): InkShapeFit | null {
  const finite = points.filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
  if (finite.length < 2) return null;
  const fit = fitInkStrokeToShape(finite);
  if (fit) return fit;
  const start = finite[0]!;
  const end = finite[finite.length - 1]!;
  const box = boxOf(finite);
  const size = Math.max(box.width, box.height);
  if (size <= 0) return null;
  const length = inkPathLength(finite);
  const closed = distance(start, end) < Math.max(12, size * 0.25) && length > size * 2;
  if (!closed) {
    // Only a stroke that runs, give or take its wavering, from its start to
    // its end straightens: a spiral, a hook or a letter the pen rested at
    // the end of stays as written (Notability straightens a line, not a curve).
    const chord = distance(start, end);
    return chord >= MIN_LINE_CHORD_PX && length <= chord * MAX_LINE_WANDER ? { kind: "line", start, end } : null;
  }
  // The corners: what is left of the path once every bump smaller than a
  // twentieth of the shape is smoothed away; the closing point joins the first.
  // Geometry alone: the pen's pressure says nothing of where the corners are.
  const kept = inkKeptSampleIndexes(finite, Math.max(4, size * 0.06), 12, 0).map((index) => finite[index]!);
  if (kept.length > 2 && distance(kept[0]!, kept[kept.length - 1]!) < Math.max(12, size * 0.25)) kept.pop();
  // A loop with no corners to speak of, or a scrawl of many, stays as written.
  if (kept.length < 3 || kept.length > 8) return null;
  return { kind: "polygon", box: boxOf(kept), points: kept };
}

function fitCenter(fit: InkShapeFit): InkFitPoint {
  if (fit.kind === "line" || fit.kind === "arrow") {
    return { x: (fit.start.x + fit.end.x) / 2, y: (fit.start.y + fit.end.y) / 2 };
  }
  return { x: fit.box.left + fit.box.width / 2, y: fit.box.top + fit.box.height / 2 };
}

/**
 * The snapped shape as the pen, still down, has moved from `anchor` (where
 * it snapped) to `pen`: a line's end goes to the pen; anything else scales
 * about its centre by how much nearer or further the pen is from it.
 */
export function inkHoldShapeAdjusted(fit: InkShapeFit, anchor: InkFitPoint, pen: InkFitPoint): InkShapeFit {
  if (fit.kind === "line" || fit.kind === "arrow") return { ...fit, end: { x: pen.x, y: pen.y } };
  const center = fitCenter(fit);
  const from = distance(anchor, center);
  const scale = from > 1 ? Math.max(0.05, distance(pen, center) / from) : 1;
  const at = (point: InkFitPoint) => ({ x: center.x + (point.x - center.x) * scale, y: center.y + (point.y - center.y) * scale });
  const box = {
    left: center.x + (fit.box.left - center.x) * scale,
    top: center.y + (fit.box.top - center.y) * scale,
    width: fit.box.width * scale,
    height: fit.box.height * scale,
  };
  if (fit.kind === "polygon") return { kind: "polygon", box, points: fit.points.map(at) };
  return { ...fit, box };
}

/** The shape's corners (or its curve, sampled) in drawing order, closed shapes ending where they began. */
function inkHoldShapeCorners(fit: InkShapeFit): InkFitPoint[] {
  if (fit.kind === "line" || fit.kind === "arrow") return [fit.start, fit.end];
  if (fit.kind === "polygon") return [...fit.points, fit.points[0]!];
  const { left, top, width, height } = fit.box;
  const right = left + width;
  const bottom = top + height;
  switch (fit.kind) {
    case "rectangle":
      return [
        { x: left, y: top },
        { x: right, y: top },
        { x: right, y: bottom },
        { x: left, y: bottom },
        { x: left, y: top },
      ];
    case "diamond":
      return [
        { x: left + width / 2, y: top },
        { x: right, y: top + height / 2 },
        { x: left + width / 2, y: bottom },
        { x: left, y: top + height / 2 },
        { x: left + width / 2, y: top },
      ];
    case "triangle":
      return [{ x: left + width / 2, y: top }, { x: right, y: bottom }, { x: left, y: bottom }, { x: left + width / 2, y: top }];
    case "right-triangle":
      return fit.flipH
        ? [{ x: right, y: top }, { x: right, y: bottom }, { x: left, y: bottom }, { x: right, y: top }]
        : [{ x: left, y: top }, { x: left, y: bottom }, { x: right, y: bottom }, { x: left, y: top }];
    case "ellipse": {
      const cx = left + width / 2;
      const cy = top + height / 2;
      const steps = 96;
      return Array.from({ length: steps + 1 }, (_, index) => {
        const angle = -Math.PI / 2 + (index / steps) * Math.PI * 2;
        return { x: cx + (width / 2) * Math.cos(angle), y: cy + (height / 2) * Math.sin(angle) };
      });
    }
  }
}

/**
 * The shape as a pen would have drawn it: its outline sampled every `spacing`
 * px, so a curve drawn through the samples runs straight along every edge
 * and turns sharply at every corner. The live stroke draws it; the landing
 * reads it back as this very shape.
 */
export function inkHoldShapeOutline(fit: InkShapeFit, spacing = 2): InkFitPoint[] {
  const corners = inkHoldShapeCorners(fit);
  const out: InkFitPoint[] = [corners[0]!];
  for (let index = 1; index < corners.length; index += 1) {
    const a = corners[index - 1]!;
    const b = corners[index]!;
    const steps = Math.max(1, Math.ceil(distance(a, b) / spacing));
    for (let step = 1; step <= steps; step += 1) {
      const t = step / steps;
      out.push({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t });
    }
  }
  return out;
}







export const INK_HOLD_SHAPE_LIFT_SLIDE_MS = 80;

/** A stroke snapped to its clean shape while the pen stays down. */
export interface InkHeldShape {
  /** The shape as it snapped, and the pen's point then. */
  readonly snapped: InkShapeFit;
  readonly anchor: InkFitPoint;
  /** The shape as it is drawn now. */
  fit: InkShapeFit;
  /** The recent shapes it took, by time (ms), for the one it lands as. */
  trail?: Array<{ readonly at: number; readonly fit: InkShapeFit }>;
}

/** The stroke snapped: the shape stands as fitted until the pen moves on. */
export function inkHeldShape(snapped: InkShapeFit, anchor: InkFitPoint, at = performance.now()): InkHeldShape {
  return { snapped, anchor, fit: snapped, trail: [{ at, fit: snapped }] };
}

/**
 * The pen moved while the shape is held: the shape follows it (a line's end,
 * anything else's size), but only once the pen has left its resting place
 * (`still`, in the points' own units): a tremble is no resize.
 */
export function inkHeldShapeFollow(held: InkHeldShape, pen: InkFitPoint, still: number, at = performance.now()): InkShapeFit {
  const moved = Math.hypot(pen.x - held.anchor.x, pen.y - held.anchor.y) > still;
  held.fit = moved ? inkHoldShapeAdjusted(held.snapped, held.anchor, pen) : held.snapped;
  const trail = held.trail ?? (held.trail = []);
  trail.push({ at, fit: held.fit });
  // Only the last moments matter to the lift.
  while (trail.length > 2 && trail[1]!.at < at - INK_HOLD_SHAPE_LIFT_SLIDE_MS * 4) trail.shift();
  return held.fit;
}

/**
 * The shape a held stroke lands as at the lift: as it stood
 * {@link INK_HOLD_SHAPE_LIFT_SLIDE_MS} before it, the pen's slide off the
 * glass dropped.
 */
export function inkHeldShapeAtLift(held: InkHeldShape, at = performance.now()): InkShapeFit {
  const trail = held.trail;
  let fit = held.snapped;
  if (trail) {
    for (const entry of trail) {
      if (entry.at > at - INK_HOLD_SHAPE_LIFT_SLIDE_MS) break;
      fit = entry.fit;
    }
  }
  held.fit = fit;
  return fit;
}
