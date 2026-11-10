/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Present-mode laser pointer + ink annotation pure helpers.
 *
 * Side-effect-free, dependency-free (no React / no DOM) logic shared by the
 * present-mode overlay's annotation layer and its unit tests. Keeping the
 * trail-decay math, stroke point simplification, hit-test erase, and tool
 * defaults here means the overlay component stays a thin render of
 * already-computed values, and the tricky geometry (decay curve, point
 * thinning, segment hit-testing) can be exercised in isolation.
 *
 * COORDINATE SPACE: every point is a SLIDE-RELATIVE fraction in `[0, 1]` on each
 * axis (x = left→right, y = top→bottom). The overlay converts pointer pixels to
 * fractions against the slide rect before calling in, and renders fractions back
 * to an SVG `viewBox="0 0 1 1"` — so a stroke stays pinned to slide content
 * regardless of how the fullscreen stage letterboxes the slide. Nothing here is
 * persisted on the deck: annotations are ephemeral per present session.
 */

import type { InkCutEnds } from "./document-shapes";
import { PRESENTATION_INK_MAX_POINTS } from "./presentation-freeform";

/** The active annotation tool. `laser` is ephemeral (never stroked); the rest draw ink. */
export type AnnotationTool = "laser" | "pen" | "pencil" | "highlighter" | "eraser";

/** The ink tools that lay down a saved stroke (everything except laser + eraser). */
export type InkTool = "pen" | "pencil" | "highlighter";

/** A single sampled point in slide-relative fractional coordinates. */
export interface AnnotationPoint {
  /** Left fraction in `[0, 1]`. */
  x: number;
  /** Top fraction in `[0, 1]`. */
  y: number;
  /** The pen's pressure at this sample (0..1), when the pen reported one. */
  p?: number;
}

/** A laser-trail sample: a point plus the wall-clock time (ms) it was captured. */
export interface TrailSample extends AnnotationPoint {
  /** `Date.now()`-style capture timestamp in ms. */
  t: number;
}

/** A finished (or in-progress) ink stroke. */
export interface InkStroke {
  /** Stable id (caller-supplied) so React keys + erase targeting are stable. */
  id: string;
  /** Pen vs highlighter — drives width + opacity + blend at render time. */
  tool: InkTool;
  /** Stroke colour (any CSS colour string the caller validated). */
  color: string;
  /** Nominal stroke width as a fraction of the slide's smaller dimension. */
  width: number;
  /** Ordered fractional points; one point renders as a dot. */
  points: AnnotationPoint[];
}

/** A laser-trail point decorated with its current render opacity in `[0, 1]`. */
export interface TrailRenderPoint extends AnnotationPoint {
  opacity: number;
}

// ── Tunables ────────────────────────────────────────────────────────────────

/** How long (ms) a laser-trail sample lives before it fully fades out. */
export const LASER_TRAIL_LIFETIME_MS = 360;

/** Hard cap on retained laser-trail samples (oldest dropped first) — bounds memory. */
export const LASER_TRAIL_MAX_POINTS = 48;

/**
 * Pen / highlighter default widths, expressed as a fraction of the slide's
 * SMALLER dimension so a stroke scales with the projected slide. A highlighter
 * is deliberately fatter (and rendered translucent) than a pen.
 */
export const PEN_DEFAULT_WIDTH = 0.006;
export const HIGHLIGHTER_DEFAULT_WIDTH = 0.03;

/** Clamp bounds for a user-tuned ink width (fraction of slide min dimension). */
export const MIN_INK_WIDTH = 0.001;
export const MAX_INK_WIDTH = 0.12;

/** Highlighter stroke opacity (pen ink is fully opaque). */
export const HIGHLIGHTER_OPACITY = 0.4;

/**
 * Minimum squared move (in fractional units) before a new point is appended to
 * the in-progress stroke. Drops jitter / sub-pixel duplicates so a stroke stays
 * light without visibly changing its shape. Squared to avoid a sqrt per move.
 */
export const STROKE_MIN_DISTANCE_SQ = 0.0000025; // (~0.0016 fraction)^2

/**
 * Distance (fraction) within which the eraser deletes a stroke: the cursor must
 * pass within this band of any stroke segment. Generous so erasing feels easy.
 */
export const ERASER_HIT_RADIUS = 0.02;

// ── Validation helpers ────────────────────────────────────────────────────────

/** A finite number, else `fallback`. */
function finiteOr(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback;
}

/**
 * Clamp `value` into `[lo, hi]`. `NaN` collapses to `lo` (a "garbage" pointer
 * coordinate snaps to the low edge); `±Infinity` clamps to the matching bound
 * (`+∞ → hi`, `-∞ → lo`) so an out-of-bounds-but-directional coordinate (e.g. a
 * divide that overflowed) still lands on the correct edge rather than flipping
 * to the opposite one.
 */
export function clampUnit(value: number, lo = 0, hi = 1): number {
  if (Number.isNaN(value)) return lo;
  if (value < lo) return lo;
  if (value > hi) return hi;
  return value;
}

/**
 * Normalize a raw pointer position (already divided by the slide rect into a
 * fraction) into a clamped, finite {@link AnnotationPoint}. Out-of-bounds drags
 * (the cursor leaving the slide) clamp to the edge rather than producing a wild
 * coordinate.
 */
export function normalizePoint(x: number, y: number): AnnotationPoint {
  return { x: clampUnit(x), y: clampUnit(y) };
}

/** Clamp an ink width into the supported band; non-finite → the pen default. */
export function clampInkWidth(width: number): number {
  const w = finiteOr(width, PEN_DEFAULT_WIDTH);
  return clampUnit(w, MIN_INK_WIDTH, MAX_INK_WIDTH);
}

/** The built-in default width for a given ink tool. */
export function defaultWidthForTool(tool: InkTool): number {
  return tool === "highlighter" ? HIGHLIGHTER_DEFAULT_WIDTH : PEN_DEFAULT_WIDTH;
}

/** Whether a tool lays down a saved ink stroke (pen / highlighter). */
export function isInkTool(tool: AnnotationTool): tool is InkTool {
  return tool === "pen" || tool === "pencil" || tool === "highlighter";
}

// ── Laser trail decay ─────────────────────────────────────────────────────────

/**
 * Append a fresh laser-trail sample, dropping anything older than the lifetime
 * relative to `now` and capping the retained count (oldest first). Returns a NEW
 * array (never mutates the input) so it slots straight into React state. The
 * appended sample is clamped + stamped with `now`.
 *
 * `lifetimeMs` defaults to {@link LASER_TRAIL_LIFETIME_MS}; a non-positive or
 * non-finite lifetime is treated as the default so a bad caller can't disable
 * pruning. The cap defaults to {@link LASER_TRAIL_MAX_POINTS}.
 */
export function pushTrailSample(
  trail: readonly TrailSample[],
  x: number,
  y: number,
  now: number,
  lifetimeMs: number = LASER_TRAIL_LIFETIME_MS,
  maxPoints: number = LASER_TRAIL_MAX_POINTS,
): TrailSample[] {
  const safeNow = finiteOr(now, 0);
  const life =
    Number.isFinite(lifetimeMs) && lifetimeMs > 0
      ? lifetimeMs
      : LASER_TRAIL_LIFETIME_MS;
  const cap =
    Number.isFinite(maxPoints) && maxPoints > 0
      ? Math.floor(maxPoints)
      : LASER_TRAIL_MAX_POINTS;
  const cutoff = safeNow - life;
  const point = normalizePoint(x, y);
  const next: TrailSample[] = [];
  for (const sample of trail) {
    // Keep only still-living samples; a non-finite timestamp is treated as dead.
    if (Number.isFinite(sample.t) && sample.t >= cutoff) next.push(sample);
  }
  next.push({ x: point.x, y: point.y, t: safeNow });
  // Cap from the FRONT (oldest) so the newest samples always survive.
  if (next.length > cap) next.splice(0, next.length - cap);
  return next;
}

/**
 * Prune a laser trail to its still-living samples relative to `now`, WITHOUT
 * appending — used by the fade ticker so the trail visibly shrinks even when the
 * cursor is still. Returns the SAME array reference when nothing was dropped (so
 * a no-op tick doesn't force a re-render); otherwise a new array.
 */
export function pruneTrail(
  trail: readonly TrailSample[],
  now: number,
  lifetimeMs: number = LASER_TRAIL_LIFETIME_MS,
): TrailSample[] {
  const safeNow = finiteOr(now, 0);
  const life =
    Number.isFinite(lifetimeMs) && lifetimeMs > 0
      ? lifetimeMs
      : LASER_TRAIL_LIFETIME_MS;
  const cutoff = safeNow - life;
  let dropped = false;
  const next: TrailSample[] = [];
  for (const sample of trail) {
    if (Number.isFinite(sample.t) && sample.t >= cutoff) next.push(sample);
    else dropped = true;
  }
  return dropped ? next : (trail as TrailSample[]);
}

/**
 * Map a laser trail to render points carrying a linear age→opacity (newest = 1,
 * about-to-expire ≈ 0). Dead samples are dropped. `now` defaults to the newest
 * sample's timestamp so callers that just want relative fade need not pass a
 * clock. Pure: never mutates the input.
 */
export function trailRenderPoints(
  trail: readonly TrailSample[],
  now?: number,
  lifetimeMs: number = LASER_TRAIL_LIFETIME_MS,
): TrailRenderPoint[] {
  if (trail.length === 0) return [];
  const life =
    Number.isFinite(lifetimeMs) && lifetimeMs > 0
      ? lifetimeMs
      : LASER_TRAIL_LIFETIME_MS;
  let clock = now;
  if (!Number.isFinite(clock as number)) {
    // Default clock = the newest finite timestamp in the trail.
    clock = 0;
    for (const sample of trail) {
      if (Number.isFinite(sample.t) && sample.t > (clock as number)) {
        clock = sample.t;
      }
    }
  }
  const safeNow = clock as number;
  const cutoff = safeNow - life;
  const out: TrailRenderPoint[] = [];
  for (const sample of trail) {
    if (!Number.isFinite(sample.t) || sample.t < cutoff) continue;
    const age = safeNow - sample.t;
    // age 0 → 1 (brightest), age == life → 0 (gone).
    const opacity = clampUnit(1 - age / life);
    out.push({ x: sample.x, y: sample.y, opacity });
  }
  return out;
}

// ── Ink stroke building + simplification ────────────────────────────────────

/**
 * Whether a new point is far enough from the last appended point to be kept,
 * using a squared-distance threshold so we avoid a sqrt. An empty stroke always
 * accepts the first point.
 */
export function shouldAppendPoint(
  points: readonly AnnotationPoint[],
  next: AnnotationPoint,
  minDistanceSq: number = STROKE_MIN_DISTANCE_SQ,
): boolean {
  const last = points[points.length - 1];
  if (!last) return true;
  const dx = next.x - last.x;
  const dy = next.y - last.y;
  const thresh =
    Number.isFinite(minDistanceSq) && minDistanceSq >= 0
      ? minDistanceSq
      : STROKE_MIN_DISTANCE_SQ;
  return dx * dx + dy * dy >= thresh;
}

/**
 * Append a clamped point to a stroke's point list IF it clears the minimum-move
 * threshold (see {@link shouldAppendPoint}). Returns a NEW array when the point
 * is kept, else the SAME reference (so a jittery move doesn't churn state).
 */
export function appendStrokePoint(
  points: readonly AnnotationPoint[],
  x: number,
  y: number,
  minDistanceSq: number = STROKE_MIN_DISTANCE_SQ,
): AnnotationPoint[] {
  const point = normalizePoint(x, y);
  if (!shouldAppendPoint(points, point, minDistanceSq)) {
    return points as AnnotationPoint[];
  }
  return [...points, point];
}

/**
 * Perpendicular distance from point `p` to the segment `a→b`, in fractional
 * units. Degenerate segments (a == b) fall back to the point-to-point distance.
 */
function pointSegmentDistance(
  p: AnnotationPoint,
  a: AnnotationPoint,
  b: AnnotationPoint,
): number {
  const abx = b.x - a.x;
  const aby = b.y - a.y;
  const apx = p.x - a.x;
  const apy = p.y - a.y;
  const lenSq = abx * abx + aby * aby;
  if (lenSq <= 0) {
    // a == b: distance to the single point.
    return Math.hypot(apx, apy);
  }
  // Projection factor of AP onto AB, clamped to the segment.
  let tNum = (apx * abx + apy * aby) / lenSq;
  if (tNum < 0) tNum = 0;
  else if (tNum > 1) tNum = 1;
  const projX = a.x + tNum * abx;
  const projY = a.y + tNum * aby;
  return Math.hypot(p.x - projX, p.y - projY);
}

/**
 * Ramer–Douglas–Peucker polyline simplification. Drops points that lie within
 * `epsilon` (fractional units) of the line between retained neighbours, keeping
 * the visual shape while bounding the point count of a saved stroke. Endpoints
 * are always preserved; strokes of ≤2 points are returned unchanged. Pure +
 * iterative (no recursion → no stack blow-up on a very long stroke).
 */
export function simplifyStroke(
  points: readonly AnnotationPoint[],
  epsilon: number,
): AnnotationPoint[] {
  const n = points.length;
  if (n <= 2 || !(epsilon > 0)) return points.slice();
  const keep = new Array<boolean>(n).fill(false);
  keep[0] = true;
  keep[n - 1] = true;
  // Iterative RDP using an explicit stack of [start, end] index pairs.
  const stack: Array<[number, number]> = [[0, n - 1]];
  while (stack.length > 0) {
    const segment = stack.pop();
    if (!segment) break;
    const [start, end] = segment;
    if (end <= start + 1) continue;
    let maxDist = -1;
    let maxIdx = -1;
    const a = points[start];
    const b = points[end];
    for (let i = start + 1; i < end; i += 1) {
      const dist = pointSegmentDistance(points[i], a, b);
      if (dist > maxDist) {
        maxDist = dist;
        maxIdx = i;
      }
    }
    if (maxDist > epsilon && maxIdx > start && maxIdx < end) {
      keep[maxIdx] = true;
      stack.push([start, maxIdx]);
      stack.push([maxIdx, end]);
    }
  }
  const out: AnnotationPoint[] = [];
  for (let i = 0; i < n; i += 1) {
    if (keep[i]) out.push(points[i]);
  }
  return out;
}

/**
 * Finalize an in-progress stroke into a saved {@link InkStroke}: clamps the
 * width, simplifies the point list (epsilon derived from the stroke width so a
 * fat highlighter thins more aggressively than a fine pen, floored so a hairline
 * pen still thins a little), and drops a zero-point stroke (returns `null`). A
 * single-point stroke is kept (renders as a dot).
 */
export function finalizeStroke(
  id: string,
  tool: InkTool,
  color: string,
  width: number,
  points: readonly AnnotationPoint[],
): InkStroke | null {
  if (points.length === 0) return null;
  const w = clampInkWidth(width);
  // Simplification tolerance scales with the stroke width (≈ a quarter of the
  // width), floored so even a hairline pen sheds redundant collinear points.
  const epsilon = Math.max(0.0008, w * 0.25);
  const simplified = simplifyStroke(points, epsilon);
  return {
    id,
    tool,
    color,
    width: w,
    points: simplified.length > 0 ? simplified : points.slice(),
  };
}

// ── Erase + clear ─────────────────────────────────────────────────────────────

/**
 * Whether an eraser pass at `point` (within `radius`) hits `stroke` — true when
 * the point lies within `radius` of any of the stroke's segments (or its single
 * point, for a dot). Used to delete whole strokes PowerPoint-style (the eraser
 * removes a stroke it touches rather than carving pixels).
 */
export function strokeHitByEraser(
  stroke: InkStroke,
  point: AnnotationPoint,
  radius: number = ERASER_HIT_RADIUS,
): boolean {
  const pts = stroke.points;
  if (pts.length === 0) return false;
  const r =
    Number.isFinite(radius) && radius >= 0 ? radius : ERASER_HIT_RADIUS;
  if (pts.length === 1) {
    return Math.hypot(point.x - pts[0].x, point.y - pts[0].y) <= r;
  }
  for (let i = 0; i < pts.length - 1; i += 1) {
    if (pointSegmentDistance(point, pts[i], pts[i + 1]) <= r) return true;
  }
  return false;
}

/**
 * Remove every stroke the eraser pass at `point` touches. Returns the SAME array
 * reference when nothing was hit (so a miss doesn't re-render), else a new array
 * with the hit strokes dropped.
 */
export function eraseAt(
  strokes: readonly InkStroke[],
  x: number,
  y: number,
  radius: number = ERASER_HIT_RADIUS,
): InkStroke[] {
  const point = normalizePoint(x, y);
  let hitAny = false;
  const kept: InkStroke[] = [];
  for (const stroke of strokes) {
    if (strokeHitByEraser(stroke, point, radius)) hitAny = true;
    else kept.push(stroke);
  }
  return hitAny ? kept : (strokes as InkStroke[]);
}

/** A persisted edit-mode ink path in slide coordinates. */
export interface PersistedInkPath {
  id: string;
  points: readonly AnnotationPoint[];
  /**
   * The line width the stroke is painted with, in the same units as
   * `points`. The ink an eraser can touch is the painted line, not only its
   * centre line: a 17pt highlighter is hit along its visible edge. Omitted, the
   * centre line alone is tested.
   */
  width?: number;
  /** The ends a point eraser already cut (flat, no round tip to reach). */
  cutEnds?: InkCutEnds;
}

const MAX_EDIT_INK_GESTURE_POINTS = 2048;









export function inkGestureWithinBudget<P extends AnnotationPoint>(
  points: readonly P[],
  limit = MAX_EDIT_INK_GESTURE_POINTS,
): AnnotationPoint[] {
  if (points.length <= limit) return points.slice();
  if (limit < 2) return points.slice(0, Math.max(0, limit));
  const along: number[] = [0];
  for (let index = 1; index < points.length; index += 1) {
    along.push(along[index - 1] + Math.hypot(points[index].x - points[index - 1].x, points[index].y - points[index - 1].y));
  }
  const total = along[along.length - 1];
  if (!(total > 0)) return [points[0]];
  const out: AnnotationPoint[] = [];
  let segment = 1;
  for (let step = 0; step < limit; step += 1) {
    const target = (total * step) / (limit - 1);
    while (segment < points.length - 1 && along[segment] < target) segment += 1;
    const a = points[segment - 1];
    const b = points[segment];
    const span = along[segment] - along[segment - 1];
    const t = span > 0 ? Math.max(0, Math.min(1, (target - along[segment - 1]) / span)) : 0;
    out.push({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t });
  }
  return out;
}

interface InkBox {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

function inkBoxOf(points: readonly AnnotationPoint[], pad: number): InkBox {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const point of points) {
    if (point.x < left) left = point.x;
    if (point.x > right) right = point.x;
    if (point.y < top) top = point.y;
    if (point.y > bottom) bottom = point.y;
  }
  return { left: left - pad, top: top - pad, right: right + pad, bottom: bottom + pad };
}

function inkBoxHolds(box: InkBox, point: AnnotationPoint): boolean {
  return point.x >= box.left && point.x <= box.right && point.y >= box.top && point.y <= box.bottom;
}

function inkBoxMeetsSegment(box: InkBox, a: AnnotationPoint, b: AnnotationPoint): boolean {
  return (
    Math.max(a.x, b.x) >= box.left &&
    Math.min(a.x, b.x) <= box.right &&
    Math.max(a.y, b.y) >= box.top &&
    Math.min(a.y, b.y) <= box.bottom
  );
}

/**
 * How far `point` is from the ink a stroke paints (0 inside it): the centre
 * line `points` widened by `half` either side, its joins round, each end the
 * pen made a round tip and each end a point eraser cut (`cutEnds`) flat
 * across the line, so nothing past a flat end is ink. The painted line is
 * the segments' butt rectangles, the outside of each bend and the round tips
 * (with every end round, exactly the centre line's distance minus `half`).
 */
export function inkPaintDistance(
  point: AnnotationPoint,
  points: readonly AnnotationPoint[],
  half: number,
  cutEnds?: InkCutEnds,
): number {
  if (points.length === 0) return Infinity;
  const segments: Array<{ a: AnnotationPoint; ux: number; uy: number; length: number }> = [];
  for (let index = 1; index < points.length; index += 1) {
    const a = points[index - 1];
    const b = points[index];
    const length = Math.hypot(b.x - a.x, b.y - a.y);
    if (length > 0) segments.push({ a, ux: (b.x - a.x) / length, uy: (b.y - a.y) / length, length });
  }
  const round = (from: AnnotationPoint) =>
    Math.max(0, Math.hypot(point.x - from.x, point.y - from.y) - half);
  // A dot: the pen's tip, all of it round.
  if (segments.length === 0) return round(points[0]);
  let best = Infinity;
  segments.forEach(({ a, ux, uy, length }, index) => {
    const along = (point.x - a.x) * ux + (point.y - a.y) * uy;
    const across = Math.abs((point.y - a.y) * ux - (point.x - a.x) * uy);
    best = Math.min(
      best,
      Math.hypot(Math.max(0, -along, along - length), Math.max(0, across - half)),
    );
    const next = segments[index + 1];
    if (!next) return;
    // The round join: the outside of the bend, ahead of this segment and
    // behind the next.
    const wx = point.x - next.a.x;
    const wy = point.y - next.a.y;
    if (wx * ux + wy * uy > 0 && wx * next.ux + wy * next.uy < 0) best = Math.min(best, round(next.a));
  });
  const first = segments[0];
  const last = segments[segments.length - 1];
  const start = points[0];
  const end = points[points.length - 1];
  if (cutEnds !== "start" && cutEnds !== "both") {
    if ((point.x - start.x) * first.ux + (point.y - start.y) * first.uy < 0) {
      best = Math.min(best, round(start));
    }
  }
  if (cutEnds !== "end" && cutEnds !== "both") {
    if ((point.x - end.x) * last.ux + (point.y - end.y) * last.uy > 0) {
      best = Math.min(best, round(end));
    }
  }
  return best;
}

function orientation(
  a: AnnotationPoint,
  b: AnnotationPoint,
  c: AnnotationPoint,
): number {
  return (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
}

function segmentsIntersect(
  a: AnnotationPoint,
  b: AnnotationPoint,
  c: AnnotationPoint,
  d: AnnotationPoint,
): boolean {
  const o1 = orientation(a, b, c);
  const o2 = orientation(a, b, d);
  const o3 = orientation(c, d, a);
  const o4 = orientation(c, d, b);
  const epsilon = 1e-9;
  if (
    Math.abs(o1) <= epsilon &&
    pointSegmentDistance(c, a, b) <= epsilon
  ) return true;
  if (
    Math.abs(o2) <= epsilon &&
    pointSegmentDistance(d, a, b) <= epsilon
  ) return true;
  if (
    Math.abs(o3) <= epsilon &&
    pointSegmentDistance(a, c, d) <= epsilon
  ) return true;
  if (
    Math.abs(o4) <= epsilon &&
    pointSegmentDistance(b, c, d) <= epsilon
  ) return true;
  return (o1 > 0) !== (o2 > 0) && (o3 > 0) !== (o4 > 0);
}

function pointInPolygon(
  point: AnnotationPoint,
  polygon: readonly AnnotationPoint[],
): boolean {
  let inside = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
    const a = polygon[i];
    const b = polygon[j];
    if (pointSegmentDistance(point, a, b) <= 1e-9) return true;
    if (
      (a.y > point.y) !== (b.y > point.y) &&
      point.x < ((b.x - a.x) * (point.y - a.y)) / (b.y - a.y) + a.x
    ) {
      inside = !inside;
    }
  }
  return inside;
}

/** Return whole persisted strokes touched by one bounded eraser gesture. */
export function inkPathIdsHitByEraser(
  paths: readonly PersistedInkPath[],
  gesturePoints: readonly AnnotationPoint[],
  radius: number,
): Set<string> {
  const hitIds = new Set<string>();
  const gesture = inkGestureWithinBudget(gesturePoints);
  const safeRadius = Number.isFinite(radius) && radius >= 0 ? radius : 0;
  if (gesture.length === 0) return hitIds;
  for (const path of paths) {
    const points = path.points.slice(0, PRESENTATION_INK_MAX_POINTS);
    if (points.length === 0) continue;
    const halfWidth =
      typeof path.width === "number" && Number.isFinite(path.width) && path.width > 0
        ? path.width / 2
        : 0;
    // Only the trace near the stroke can touch it (the whole trace is
    // measured now, so it is narrowed to the stroke's box first).
    const reachBox = inkBoxOf(points, safeRadius + halfWidth + 1);
    // The painted line, flat where a point eraser cut it: a stroke the
    // eraser passed beyond its flat end is one it missed.
    const touchedBySample = gesture.some(
      (point) => inkBoxHolds(reachBox, point) && inkPaintDistance(point, points, halfWidth, path.cutEnds) <= safeRadius,
    );
    const touchedByCrossing = !touchedBySample && gesture.some((gesturePoint, gestureIndex) =>
      gestureIndex + 1 < gesture.length && inkBoxMeetsSegment(reachBox, gesturePoint, gesture[gestureIndex + 1])
        ? points.some((pathPoint, pathIndex) =>
            pathIndex + 1 < points.length
              ? segmentsIntersect(
                  gesturePoint,
                  gesture[gestureIndex + 1],
                  pathPoint,
                  points[pathIndex + 1],
                )
              : false,
          )
        : false,
    );
    const touched = touchedBySample || touchedByCrossing;
    if (touched) hitIds.add(path.id);
  }
  return hitIds;
}

/** Return persisted strokes inside or intersecting one bounded lasso polygon. */
export function inkPathIdsInLasso(
  paths: readonly PersistedInkPath[],
  polygonPoints: readonly AnnotationPoint[],
): Set<string> {
  const selectedIds = new Set<string>();
  const polygon = inkGestureWithinBudget(polygonPoints);
  if (polygon.length < 3) return selectedIds;
  for (const path of paths) {
    const points = path.points.slice(0, PRESENTATION_INK_MAX_POINTS);
    if (points.some((point) => pointInPolygon(point, polygon))) {
      selectedIds.add(path.id);
      continue;
    }
    let intersects = false;
    for (let i = 0; i + 1 < points.length && !intersects; i += 1) {
      for (let j = 0; j < polygon.length; j += 1) {
        if (
          segmentsIntersect(
            points[i],
            points[i + 1],
            polygon[j],
            polygon[(j + 1) % polygon.length],
          )
        ) {
          intersects = true;
          break;
        }
      }
    }
    if (intersects) selectedIds.add(path.id);
  }
  return selectedIds;
}

/** The effective render opacity for an ink tool (highlighter is translucent). */
export function inkOpacityForTool(tool: InkTool): number {
  return tool === "highlighter" ? HIGHLIGHTER_OPACITY : 1;
}

/**
 * The `opacity` a committed ink shape stores for a tool, or `undefined` for
 * opaque pen ink. A slide shape's opacity is a 0-100 percentage
 * (`normalizeShapeOpacity`), so the 0-1 render fraction must be scaled: a
 * stored 0.45 rounds to 0 and the highlighter stroke vanishes (PPT-B157).
 */
export function inkShapeOpacityForTool(tool: InkTool): number | undefined {
  const opacity = inkOpacityForTool(tool);
  return opacity < 1 ? Math.round(opacity * 100) : undefined;
}

/**
 * Build the SVG path `d` string for a stroke. Points are stored as `[0,1]`
 * fractions on each axis; the X coordinate is multiplied by `xScale` so the
 * stroke can be drawn into an aspect-corrected viewBox (`0 0 xScale 1`) under
 * `preserveAspectRatio="none"`. With `xScale` = the slide's width/height ratio,
 * one X user-unit equals one Y user-unit in SCREEN pixels, so the stroke keeps a
 * uniform width and never skews with the slide's aspect ratio. `xScale` defaults
 * to 1 (a square 0..1 viewBox).
 *
 * A single point renders as a degenerate self-segment so `stroke-linecap: round`
 * paints a dot. Coordinates are rounded to 4 decimals to keep the DOM compact.
 */
export function strokePathData(stroke: InkStroke, xScale = 1): string {
  const pts = stroke.points;
  if (pts.length === 0) return "";
  const sx =
    Number.isFinite(xScale) && xScale > 0 ? xScale : 1;
  const fx = (n: number) => {
    const r = Math.round(clampUnit(n) * sx * 10000) / 10000;
    return String(r);
  };
  const fy = (n: number) => {
    const r = Math.round(clampUnit(n) * 10000) / 10000;
    return String(r);
  };
  if (pts.length === 1) {
    const x = fx(pts[0].x);
    const y = fy(pts[0].y);
    // Degenerate segment → round cap paints a dot.
    return `M ${x} ${y} L ${x} ${y}`;
  }
  let d = `M ${fx(pts[0].x)} ${fy(pts[0].y)}`;
  for (let i = 1; i < pts.length; i += 1) {
    d += ` L ${fx(pts[i].x)} ${fy(pts[i].y)}`;
  }
  return d;
}
