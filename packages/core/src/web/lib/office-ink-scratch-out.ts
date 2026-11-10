/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Scribble to erase (Goodnotes' "Scribble to Erase"): a pen scribbled back
 * and forth over ink erases the ink under it, without reaching for the
 * eraser. Goodnotes: "strokes beneath the scribbled area are removed", and
 * strokes outside the scribble (the dot of an i) stay. A scribble over
 * nothing, or over ink it only grazes, is ink like any other.
 *
 * Two questions, host-neutral (points in any one pixel space):
 * - is this stroke a scratch-out: to and fro along one direction at least
 *   {@link INK_SCRATCH_MIN_PASSES} times, or a dense zigzag of at least
 *   {@link INK_SCRATCH_MIN_ZIGZAG_LEGS} steep legs (a zigzag line, a spiral,
 *   a letter m or w are not);
 * - which strokes it covers: those with at least
 *   {@link INK_SCRATCH_COVERED_SHARE} of their samples inside the area the
 *   scribble swept (its outline, widened by the pen and a little more).
 */

export interface InkScratchPoint {
  readonly x: number;
  readonly y: number;
}

export interface InkScratchPath {
  readonly id: string;
  readonly points: readonly InkScratchPoint[];
}

/**
 * A scratch-out to and fro along one direction runs at least this many passes
 * (three turns) when it is a flat patch (breadth at most half its span); one
 * fewer than {@link INK_SCRATCH_PASSES} otherwise reads as an m or a w.
 */
export const INK_SCRATCH_MIN_PASSES = 4;
/** Passes that make a scratch-out whatever its shape (four turns). */
export const INK_SCRATCH_PASSES = 5;
/** A zigzag scratch-out (a word crossed out with WWWW) runs at least this many legs (five turns). */
export const INK_SCRATCH_MIN_ZIGZAG_LEGS = 6;
/** Its passes are at least this long (px): a tremor is not a pass. */
export const INK_SCRATCH_MIN_SPAN_PX = 12;
/** A zigzag's legs are at least this tall (px). */
export const INK_SCRATCH_MIN_ZIGZAG_PX = 8;
/** How much of a to-and-fro scratch-out's travel runs along its direction (0..1). */
export const INK_SCRATCH_MIN_ALIGNMENT = 0.72;
/** How much of a zigzag scratch-out's travel runs across its way (steep legs). */
export const INK_SCRATCH_MIN_ZIGZAG_STEEPNESS = 0.75;
/** A stroke this much under the scribble (by its samples) is erased. */
export const INK_SCRATCH_COVERED_SHARE = 0.5;

/** Whether a stroke drawn with `tool` may scratch out: a pen's may; a highlighter's and a pencil's shading may not. */
export function inkToolScratchesOut(tool: string | null | undefined): boolean {
  return tool !== "highlighter" && tool !== "pencil";
}

const distance = (a: InkScratchPoint, b: InkScratchPoint) => Math.hypot(a.x - b.x, a.y - b.y);

/** The direction the stroke mostly runs along (its points' principal axis), as a unit vector. */
function principalAxis(points: readonly InkScratchPoint[]): InkScratchPoint {
  let mx = 0;
  let my = 0;
  for (const point of points) {
    mx += point.x;
    my += point.y;
  }
  mx /= points.length;
  my /= points.length;
  let sxx = 0;
  let syy = 0;
  let sxy = 0;
  for (const point of points) {
    const dx = point.x - mx;
    const dy = point.y - my;
    sxx += dx * dx;
    syy += dy * dy;
    sxy += dx * dy;
  }
  const angle = 0.5 * Math.atan2(2 * sxy, sxx - syy);
  return { x: Math.cos(angle), y: Math.sin(angle) };
}

/**
 * The passes along one coordinate: each turn a retreat of at least `reach`
 * from the furthest point of the pass before it (hysteresis: a wobble is not
 * a turn).
 */
function passesOf(values: readonly number[], reach: number): number {
  let passes = 1;
  let direction = 0;
  let extreme = values[0]!;
  for (const value of values) {
    if (direction === 0) {
      if (Math.abs(value - values[0]!) >= reach) {
        direction = Math.sign(value - values[0]!);
        extreme = value;
      }
    } else if ((value - extreme) * direction > 0) {
      extreme = value;
    } else if ((extreme - value) * direction >= reach) {
      passes += 1;
      direction = -direction;
      extreme = value;
    }
  }
  return passes;
}

/** Loops of at least this many turns (radians) one way are circling, not scratching. */
const INK_LOOP_MIN_TURNING = 3 * Math.PI;
/** How much of a circling stroke's turning runs one way (0..1). */
const INK_LOOP_ONE_WAY_SHARE = 0.6;
/** Direction is read over steps at least this long (px), not the sample jitter. */
const INK_TURNING_STEP_PX = 2;








function keepsTurningOneWay(points: readonly InkScratchPoint[]): boolean {
  let net = 0;
  let total = 0;
  let previous: InkScratchPoint | null = null;
  let anchor = points[0]!;
  for (const point of points) {
    if (distance(point, anchor) < INK_TURNING_STEP_PX) continue;
    const step = { x: point.x - anchor.x, y: point.y - anchor.y };
    if (previous) {
      const turn = Math.atan2(previous.x * step.y - previous.y * step.x, previous.x * step.x + previous.y * step.y);
      net += turn;
      total += Math.abs(turn);
    }
    previous = step;
    anchor = point;
  }
  return Math.abs(net) >= INK_LOOP_MIN_TURNING && Math.abs(net) >= total * INK_LOOP_ONE_WAY_SHARE;
}

/**
 * Whether `points` (one stroke, in drawing order) is a scribble meant to
 * erase: to and fro along one direction (shading out a word), or a dense
 * zigzag across it (WWWW over a word). Writing, a zigzag line drawn for its
 * own sake (it travels as far as its legs are tall), a spiral, an m or a w
 * are not.
 */
export function isInkScratchOut(points: readonly InkScratchPoint[]): boolean {
  if (points.length < 12) return false;
  if (keepsTurningOneWay(points)) return false;
  const axis = principalAxis(points);
  const along = points.map((point) => point.x * axis.x + point.y * axis.y);
  const across = points.map((point) => -point.x * axis.y + point.y * axis.x);
  const span = Math.max(...along) - Math.min(...along);
  const breadth = Math.max(...across) - Math.min(...across);
  let length = 0;
  let alongTravel = 0;
  let acrossTravel = 0;
  for (let index = 1; index < points.length; index += 1) {
    const dx = points[index]!.x - points[index - 1]!.x;
    const dy = points[index]!.y - points[index - 1]!.y;
    length += Math.hypot(dx, dy);
    alongTravel += Math.abs(dx * axis.x + dy * axis.y);
    acrossTravel += Math.abs(-dx * axis.y + dy * axis.x);
  }
  if (length <= 0) return false;
  // To and fro along the stroke's own direction, compact (a scribble covers
  // a patch, it does not travel), straight passes that turn back sharply (a
  // spiral or a loop of loops turns gradually: about two thirds along).
  if (span >= INK_SCRATCH_MIN_SPAN_PX && breadth <= span * 1.5) {
    const passes = passesOf(along, span / 3);
    if (
      (passes >= INK_SCRATCH_PASSES || (passes >= INK_SCRATCH_MIN_PASSES && breadth <= span * 0.5)) &&
      length >= span * (passes - 1) * 0.8 &&
      alongTravel >= length * INK_SCRATCH_MIN_ALIGNMENT
    ) {
      return true;
    }
  }
  // A dense zigzag across its way: many steep legs, each moving on less than
  // its height (a zigzag line drawn as a line moves on as far as it rises).
  if (breadth >= INK_SCRATCH_MIN_ZIGZAG_PX) {
    const legs = passesOf(across, breadth / 3);
    if (
      legs >= INK_SCRATCH_MIN_ZIGZAG_LEGS &&
      span / legs <= breadth * 0.6 &&
      acrossTravel >= length * INK_SCRATCH_MIN_ZIGZAG_STEEPNESS
    ) {
      return true;
    }
  }
  return false;
}

/** The convex hull of `points` (counter-clockwise), for the area a scribble swept. */
function hull(points: readonly InkScratchPoint[]): InkScratchPoint[] {
  const sorted = [...points].sort((a, b) => a.x - b.x || a.y - b.y);
  if (sorted.length < 3) return sorted;
  const cross = (o: InkScratchPoint, a: InkScratchPoint, b: InkScratchPoint) =>
    (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x);
  const lower: InkScratchPoint[] = [];
  for (const point of sorted) {
    while (lower.length >= 2 && cross(lower[lower.length - 2]!, lower[lower.length - 1]!, point) <= 0) lower.pop();
    lower.push(point);
  }
  const upper: InkScratchPoint[] = [];
  for (let index = sorted.length - 1; index >= 0; index -= 1) {
    const point = sorted[index]!;
    while (upper.length >= 2 && cross(upper[upper.length - 2]!, upper[upper.length - 1]!, point) <= 0) upper.pop();
    upper.push(point);
  }
  upper.pop();
  lower.pop();
  return [...lower, ...upper];
}

function insidePolygon(point: InkScratchPoint, polygon: readonly InkScratchPoint[]): boolean {
  let inside = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
    const a = polygon[i]!;
    const b = polygon[j]!;
    if (a.y > point.y !== b.y > point.y && point.x < ((b.x - a.x) * (point.y - a.y)) / (b.y - a.y) + a.x) {
      inside = !inside;
    }
  }
  return inside;
}

function distanceToSegment(point: InkScratchPoint, a: InkScratchPoint, b: InkScratchPoint): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const length2 = dx * dx + dy * dy;
  const t = length2 > 0 ? Math.min(1, Math.max(0, ((point.x - a.x) * dx + (point.y - a.y) * dy) / length2)) : 0;
  return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
}

/**
 * The strokes a scratch-out covers: at least {@link INK_SCRATCH_COVERED_SHARE}
 * of each one's samples inside the scribble's outline or within `reachPx` of
 * it (the scribble's own half width and a margin).
 */
export function inkPathIdsUnderScratchOut(
  paths: readonly InkScratchPath[],
  scribble: readonly InkScratchPoint[],
  reachPx: number,
): Set<string> {
  const ids = new Set<string>();
  const outline = hull(scribble);
  if (outline.length < 3) return ids;
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const point of outline) {
    minX = Math.min(minX, point.x);
    minY = Math.min(minY, point.y);
    maxX = Math.max(maxX, point.x);
    maxY = Math.max(maxY, point.y);
  }
  const reach = Math.max(0, reachPx);
  const covered = (point: InkScratchPoint) => {
    if (point.x < minX - reach || point.x > maxX + reach || point.y < minY - reach || point.y > maxY + reach) return false;
    if (insidePolygon(point, outline)) return true;
    for (let index = 0; index < outline.length; index += 1) {
      if (distanceToSegment(point, outline[index]!, outline[(index + 1) % outline.length]!) <= reach) return true;
    }
    return false;
  };
  for (const path of paths) {
    if (path.points.length === 0) continue;
    let inside = 0;
    for (const point of path.points) if (covered(point)) inside += 1;
    if (inside / path.points.length >= INK_SCRATCH_COVERED_SHARE) ids.add(path.id);
  }
  return ids;
}







export function inkScratchOutEnabled(): boolean {
  return inkPenGestures().scribbleErase;
}

/**
 * Whether a scribble may take this mark: handwriting always; a shape (Ink to
 * Shape's, or a stroke drawn and held into one) and a highlighter stroke only
 * while Erase shapes and highlighter is on.
 */
export function inkScratchOutTakes(mark: { tool?: string | null; shape?: boolean }): boolean {
  if (inkPenGestures().scribbleErasesShapes) return true;
  return mark.shape !== true && mark.tool !== "highlighter";
}import { inkPenGestures } from "./office-ink-pen-gestures";

