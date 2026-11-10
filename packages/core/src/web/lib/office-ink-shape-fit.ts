/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Ink to Shape (XLS-A032): fit one hand-drawn stroke to the Office
 * shape it most plainly is, by geometry alone.
 *
 * Microsoft Support, "Convert ink to shapes in Office": with Ink to Shape on,
 * a drawn shape turns into the equivalent Office shape — rectangles and
 * squares, diamonds, parallelograms, trapezoids and other quadrilaterals,
 * pentagons, hexagons, ellipses and circles, triangles, lines and arrows.
 * Whatever this fit cannot call stays ink, as in Office.
 *
 * Host-neutral: points are in any one pixel space (y down); the host turns a
 * fit into its own shape object.
 */

export interface InkFitPoint {
  readonly x: number;
  readonly y: number;
}

export interface InkFitBox {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

type InkLineFitKind = "line" | "arrow";
type InkPresetFitKind = "rectangle" | "ellipse" | "diamond" | "triangle" | "right-triangle";

// One member per kind, so `fit.kind === …` and its negation both tell the
// members apart for every reader of a fit.
export type InkShapeFit =
  /** A straight stroke; an arrow has its head at `end`. */
  | {
      [K in InkLineFitKind]: {
        readonly kind: K;
        readonly start: InkFitPoint;
        readonly end: InkFitPoint;
      };
    }[InkLineFitKind]
  /** An axis-aligned preset filling `box` (the isosceles triangle apex up,
   * a right triangle's right angle bottom-left unless `flipH`). */
  | {
      [K in InkPresetFitKind]: {
        readonly kind: K;
        readonly box: InkFitBox;
        readonly flipH?: boolean;
      };
    }[InkPresetFitKind]
  /** Any other triangle, quadrilateral, pentagon or hexagon: its corners. */
  | { readonly kind: "polygon"; readonly box: InkFitBox; readonly points: readonly InkFitPoint[] };

/** Strokes smaller than this (either way) are writing, not shapes. */
const MIN_SHAPE_SIZE_PX = 12;
/** A stroke of fewer samples than this (a flick, a tap and a drag) stays ink. */
const MIN_SHAPE_SAMPLES = 6;
/** A stroke shorter than this (px along the path) stays ink. */
const MIN_SHAPE_LENGTH_PX = 24;

const distance = (a: InkFitPoint, b: InkFitPoint) => Math.hypot(a.x - b.x, a.y - b.y);

function pathLength(points: readonly InkFitPoint[]): number {
  let total = 0;
  for (let index = 1; index < points.length; index += 1) total += distance(points[index - 1]!, points[index]!);
  return total;
}

function boxOf(points: readonly InkFitPoint[]): InkFitBox {
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

function segmentDistance(point: InkFitPoint, a: InkFitPoint, b: InkFitPoint): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const lengthSquared = dx * dx + dy * dy;
  if (lengthSquared === 0) return distance(point, a);
  const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / lengthSquared));
  return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
}

/** Ramer–Douglas–Peucker on an open run. */
function simplify(points: readonly InkFitPoint[], epsilon: number): InkFitPoint[] {
  if (points.length <= 2) return [...points];
  let worst = 0;
  let worstIndex = 0;
  const first = points[0]!;
  const last = points[points.length - 1]!;
  for (let index = 1; index < points.length - 1; index += 1) {
    const d = segmentDistance(points[index]!, first, last);
    if (d > worst) {
      worst = d;
      worstIndex = index;
    }
  }
  if (worst <= epsilon) return [first, last];
  const left = simplify(points.slice(0, worstIndex + 1), epsilon);
  const right = simplify(points.slice(worstIndex), epsilon);
  return [...left.slice(0, -1), ...right];
}

/** The interior angle at `b`, in degrees (180 = straight on). */
function angleAt(a: InkFitPoint, b: InkFitPoint, c: InkFitPoint): number {
  const v1 = { x: a.x - b.x, y: a.y - b.y };
  const v2 = { x: c.x - b.x, y: c.y - b.y };
  const norm = Math.hypot(v1.x, v1.y) * Math.hypot(v2.x, v2.y);
  if (norm === 0) return 180;
  const cos = Math.max(-1, Math.min(1, (v1.x * v2.x + v1.y * v2.y) / norm));
  return (Math.acos(cos) * 180) / Math.PI;
}

/** The corners of a closed stroke: simplified, then near and flat corners dropped. */
function closedCorners(points: readonly InkFitPoint[], perimeter: number, size: number): InkFitPoint[] {
  // Split the loop at the point farthest from its start so both halves
  // simplify as open runs and the start is never forced to be a corner.
  const start = points[0]!;
  let farIndex = 0;
  for (let index = 1; index < points.length; index += 1) {
    if (distance(points[index]!, start) > distance(points[farIndex]!, start)) farIndex = index;
  }
  const epsilon = Math.max(2, perimeter * 0.02);
  const loop = [...points, start];
  const corners = [
    ...simplify(loop.slice(0, farIndex + 1), epsilon).slice(0, -1),
    ...simplify(loop.slice(farIndex), epsilon).slice(0, -1),
  ];
  let changed = true;
  while (changed && corners.length > 3) {
    changed = false;
    for (let index = 0; index < corners.length && corners.length > 3; index += 1) {
      const previous = corners[(index - 1 + corners.length) % corners.length]!;
      const here = corners[index]!;
      const next = corners[(index + 1) % corners.length]!;
      if (distance(previous, here) < size * 0.12 || angleAt(previous, here, next) > 155) {
        corners.splice(index, 1);
        changed = true;
        break;
      }
    }
  }
  return corners;
}

function polygonError(points: readonly InkFitPoint[], corners: readonly InkFitPoint[]): number {
  let total = 0;
  for (const point of points) {
    let best = Infinity;
    for (let index = 0; index < corners.length; index += 1) {
      best = Math.min(best, segmentDistance(point, corners[index]!, corners[(index + 1) % corners.length]!));
    }
    total += best;
  }
  return total / points.length;
}

function ellipseError(points: readonly InkFitPoint[], box: InkFitBox): number {
  const rx = box.width / 2;
  const ry = box.height / 2;
  const cx = box.left + rx;
  const cy = box.top + ry;
  let total = 0;
  for (const point of points) {
    total += Math.abs(Math.hypot((point.x - cx) / rx, (point.y - cy) / ry) - 1);
  }
  return total / points.length;
}

/** A near-horizontal or near-vertical line is straightened, as Office does. */
function straightened(start: InkFitPoint, end: InkFitPoint): { start: InkFitPoint; end: InkFitPoint } {
  const angle = (Math.atan2(Math.abs(end.y - start.y), Math.abs(end.x - start.x)) * 180) / Math.PI;
  if (angle < 6) return { start, end: { x: end.x, y: start.y } };
  if (angle > 84) return { start, end: { x: start.x, y: end.y } };
  return { start, end };
}

/** How far the stroke strays from the chord between two of its points. */
function chordDeviation(points: readonly InkFitPoint[], from: InkFitPoint, to: InkFitPoint): number {
  let worst = 0;
  for (const point of points) worst = Math.max(worst, segmentDistance(point, from, to));
  return worst;
}

/** The degrees a stroke turns through once closed back to its start (a
 * loop turns ~360; a hooked stroke that merely ends near its start, far
 * less). */
function turning(points: readonly InkFitPoint[], length: number): number {
  const path = simplify([...points, points[0]!], Math.max(2, length * 0.02)).slice(0, -1);
  if (path.length < 3) return 0;
  let total = 0;
  for (let index = 0; index < path.length; index += 1) {
    const previous = path[(index - 1 + path.length) % path.length]!;
    const next = path[(index + 1) % path.length]!;
    total += 180 - angleAt(previous, path[index]!, next);
  }
  return total;
}

/** A line is a stroke that runs straight: nearly as long as its chord and
 * never far from it. A gentle arc, or a hooked stroke of handwriting, is
 * not one even when its ends are far apart. */
const LINE_CHORD_RATIO = 0.95;
const LINE_DEVIATION_RATIO = 0.035;

function fitOpen(points: readonly InkFitPoint[], length: number): InkShapeFit | null {
  const first = points[0]!;
  const last = points[points.length - 1]!;
  if (
    distance(first, last) / length >= LINE_CHORD_RATIO &&
    chordDeviation(points, first, last) <= Math.max(2.5, length * LINE_DEVIATION_RATIO)
  ) {
    return { kind: "line", ...straightened(first, last) };
  }
  // An arrow drawn in one stroke: a straight shaft to the tip, then the head
  // drawn back from it, staying near the tip and behind it.
  // The tip is where the shaft first reaches (about) its farthest: the head's
  // second barb returns to the tip and may land a hair farther out.
  let farthest = 0;
  for (const point of points) farthest = Math.max(farthest, distance(point, first));
  let tipIndex = points.findIndex((point) => distance(point, first) >= farthest * 0.97);
  while (
    tipIndex + 1 < points.length &&
    distance(points[tipIndex + 1]!, first) > distance(points[tipIndex]!, first)
  ) {
    tipIndex += 1;
  }
  const tip = points[tipIndex]!;
  const shaft = distance(first, tip);
  const shaftPath = pathLength(points.slice(0, tipIndex + 1));
  const headPath = length - shaftPath;
  if (
    shaft === 0 ||
    shaft / shaftPath < LINE_CHORD_RATIO ||
    chordDeviation(points.slice(0, tipIndex + 1), first, tip) > Math.max(2.5, shaft * LINE_DEVIATION_RATIO) ||
    headPath < shaft * 0.12 ||
    headPath > shaft * 1.2
  ) {
    return null;
  }
  const direction = { x: (tip.x - first.x) / shaft, y: (tip.y - first.y) / shaft };
  const head = points.slice(tipIndex);
  const headIsBehindTip = head.every(
    (point) =>
      distance(point, tip) <= shaft * 0.45 &&
      (point.x - tip.x) * direction.x + (point.y - tip.y) * direction.y <= shaft * 0.05,
  );
  // A head has both barbs: it reaches out on either side of the shaft.
  const sides = head.map((point) => (point.x - tip.x) * direction.y - (point.y - tip.y) * direction.x);
  const reach = shaft * 0.04;
  if (!headIsBehindTip || !sides.some((side) => side > reach) || !sides.some((side) => side < -reach)) return null;
  return { kind: "arrow", ...straightened(first, tip) };
}

function near(point: InkFitPoint, target: InkFitPoint, tolerance: number): boolean {
  return distance(point, target) <= tolerance;
}

function fitPolygon(corners: readonly InkFitPoint[]): InkShapeFit | null {
  const box = boxOf(corners);
  const size = Math.max(box.width, box.height);
  const tolerance = size * 0.18;
  const { left, top, width, height } = box;
  const right = left + width;
  const bottom = top + height;
  const all = (targets: readonly InkFitPoint[]) =>
    targets.every((target) => corners.some((corner) => near(corner, target, tolerance)));
  if (corners.length === 4) {
    if (all([{ x: left, y: top }, { x: right, y: top }, { x: right, y: bottom }, { x: left, y: bottom }])) {
      return { kind: "rectangle", box };
    }
    const midX = left + width / 2;
    const midY = top + height / 2;
    if (all([{ x: midX, y: top }, { x: right, y: midY }, { x: midX, y: bottom }, { x: left, y: midY }])) {
      return { kind: "diamond", box };
    }
  }
  if (corners.length === 3) {
    const midX = left + width / 2;
    if (all([{ x: midX, y: top }, { x: left, y: bottom }, { x: right, y: bottom }])) {
      return { kind: "triangle", box };
    }
    if (all([{ x: left, y: top }, { x: left, y: bottom }, { x: right, y: bottom }])) {
      return { kind: "right-triangle", box };
    }
    if (all([{ x: right, y: top }, { x: right, y: bottom }, { x: left, y: bottom }])) {
      return { kind: "right-triangle", box, flipH: true };
    }
    // Any other three-cornered loop is a triangle too (Office snaps it to
    // the standard one in its box), never a wedge of its own.
    return { kind: "triangle", box };
  }
  // A loop with more corners that is none of the standard shapes stays ink:
  // Office does not make a shape of a stroke it does not recognise.
  return null;
}

/**
 * A STORED stroke read back at the density a pen gives (a sample at least
 * every `spacing` px along it): a landed stroke keeps only the samples its
 * shape needs (a rectangle is little more than its corners), and the fitter
 * reads a stroke as the pen drew it. A stroke just drawn is fitted as it is.
 */
export function inkSamplesForFitting(points: readonly InkFitPoint[], spacing = 2): InkFitPoint[] {
  if (points.length < 2) return points.slice();
  const out: InkFitPoint[] = [points[0]!];
  for (let index = 1; index < points.length; index += 1) {
    const a = points[index - 1]!;
    const b = points[index]!;
    const steps = Math.ceil(distance(a, b) / spacing);
    for (let step = 1; step < steps; step += 1) {
      const t = step / steps;
      out.push({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t });
    }
    out.push(b);
  }
  return out;
}

/**
 * The shape one stroke is, or null when it is not plainly one (writing, a
 * squiggle, a tap): that stroke stays ink.
 */
export function fitInkStrokeToShape(samples: readonly InkFitPoint[]): InkShapeFit | null {
  const points = samples.filter(
    (point, index) =>
      Number.isFinite(point.x) &&
      Number.isFinite(point.y) &&
      (index === 0 || distance(point, samples[index - 1]!) > 0),
  );
  if (points.length < 3) return null;
  const box = boxOf(points);
  const size = Math.max(box.width, box.height);
  if (size < MIN_SHAPE_SIZE_PX) return null;
  const length = pathLength(points);
  // A shape is drawn deliberately: a flick of a few samples, or a stroke
  // too short to have been meant as one, stays ink. Office converts only a
  // stroke it recognises with confidence; the rest of the checks below are
  // its shape tests, these are the pen's.
  if (points.length < MIN_SHAPE_SAMPLES || length < MIN_SHAPE_LENGTH_PX) return null;
  const gap = distance(points[0]!, points[points.length - 1]!);
  if (gap > length * 0.18 || gap > size * 0.35) return fitOpen(points, length);
  // A closed stroke too thin to enclose anything is a line drawn back on itself.
  if (Math.min(box.width, box.height) < size * 0.12) return null;
  // A loop goes once around (it turns through about a full circle); a
  // stroke of writing that merely ends near where it began -- a hook, a
  // tick drawn back on itself -- turns far less and stays ink.
  if (turning(points, length) < 300) return null;
  const roundness = ellipseError(points, box);
  const corners = closedCorners(points, length, size);
  const cornerError = corners.length >= 3 ? polygonError(points, corners) / size : Infinity;
  if (roundness < 0.1 && (corners.length > 6 || cornerError > roundness * 0.35)) {
    return { kind: "ellipse", box };
  }
  if (corners.length >= 3 && corners.length <= 6 && cornerError < 0.04) return fitPolygon(corners);
  if (roundness < 0.14) return { kind: "ellipse", box };
  return null;
}
