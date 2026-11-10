/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A pen stroke as a pen draws it: one curve through its samples.
 *
 * Every ink surface drew a stroke as the polyline of its samples. A pointer
 * reports a sample every few pixels (sixty to a second; a quick loop of a
 * letter is six or eight of them), so handwriting came out as polygons: the
 * round of an `o` had corners, and at a glance the ink read as low
 * resolution. Office and every note-taking app draw ink as a curve.
 *
 * The curve is a centripetal Catmull-Rom spline as cubic Béziers:
 * - it passes through every sample, so what the erasers and the lasso test
 *   (the samples) is what is drawn, and a stroke keeps its size;
 * - centripetal, so unevenly spaced samples (a pen slowing into a turn) do
 *   not overshoot or loop, as the uniform spline does;
 * - a sharp turn stays sharp: the spline is broken at a sample where the
 *   pen turned by more than {@link INK_CORNER_DEGREES} (the points of an M,
 *   a V, a 7), instead of rounding it into a loop;
 * - the ends leave along the first and last segments, so a cut end's flat
 *   cap and a tip drawn there sit as they did on the polyline.
 *
 * DOM-free; the samples, the file formats and every hit test stay as they
 * are. Only the paint changes.
 */

/** The pen turned by more than this at a sample: a corner, kept sharp. */
export const INK_CORNER_DEGREES = 80;

/**
 * The corner angle per tool: a pen keeps its angles (the point of an M, a 7)
 * from a smaller turn, so handwriting is not rounded into a lazy curve; a
 * highlighter's broad band hides a corner, so it is rounded further.
 */
export function inkCornerDegreesForTool(tool: string | null | undefined): number {
  switch (tool) {
    case "pen":
      return 62;
    case "pencil":
      return 70;
    default:
      return INK_CORNER_DEGREES;
  }
}

export interface InkSmoothPoint {
  readonly x: number;
  readonly y: number;
}

/** A sample with the pen's pressure (0..1, 0.5 for a mouse) when it is known. */
export interface InkPressurePoint extends InkSmoothPoint {
  readonly p?: number;
}

export interface InkSmoothOptions {
  /**
   * The size of one unit of each axis on screen, when the points are in a
   * stretched space (a stroke's own 0..100 box): the corner test measures
   * angles as drawn. Default 1 × 1.
   */
  readonly unit?: { readonly x: number; readonly y: number };
  /** Decimals kept in the path data (default 2; a 0..100 box wants 3). */
  readonly precision?: number;
  /** The turn (degrees) past which a sample is a corner; default {@link INK_CORNER_DEGREES}. */
  readonly corner?: number;
  /** An outline's ends a point eraser cut: those stop flat, not in a round tip. */
  readonly flatStart?: boolean;
  readonly flatEnd?: boolean;
}

function format(value: number, precision: number): string {
  const rounded = Number(value.toFixed(precision));
  return String(Object.is(rounded, -0) ? 0 : rounded);
}

/** Samples that are not on top of the one before (a still pen repeats itself). */
function distinct<P extends InkSmoothPoint>(points: readonly P[]): P[] {
  const kept: P[] = [];
  for (const point of points) {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
    const previous = kept[kept.length - 1];
    if (previous && previous.x === point.x && previous.y === point.y) continue;
    kept.push(point);
  }
  return kept;
}

/** The two Bézier controls of the curve from `p1` to `p2`, its neighbours
 * `p0` / `p3` when the curve follows them (null across a corner or past an
 * end). `unit` is the size of one unit on screen (see {@link InkSmoothOptions}). */
export function inkCurveControls(
  p0: InkSmoothPoint | null,
  p1: InkSmoothPoint,
  p2: InkSmoothPoint,
  p3: InkSmoothPoint | null,
  unitX = 1,
  unitY = 1,
): { c1x: number; c1y: number; c2x: number; c2y: number } {
  const chordOf = (a: InkSmoothPoint, b: InkSmoothPoint) =>
    Math.hypot((b.x - a.x) * unitX, (b.y - a.y) * unitY);
  const d2 = Math.sqrt(chordOf(p1, p2));
  const d1 = p0 ? Math.sqrt(chordOf(p0, p1)) : 0;
  const d3 = p3 ? Math.sqrt(chordOf(p2, p3)) : 0;
  let c1x = p1.x + (p2.x - p1.x) / 3;
  let c1y = p1.y + (p2.y - p1.y) / 3;
  let c2x = p2.x - (p2.x - p1.x) / 3;
  let c2y = p2.y - (p2.y - p1.y) / 3;
  if (p0 && d1 > 0 && d2 > 0) {
    const scale = 3 * d1 * (d1 + d2);
    c1x = (d1 * d1 * p2.x - d2 * d2 * p0.x + (2 * d1 * d1 + 3 * d1 * d2 + d2 * d2) * p1.x) / scale;
    c1y = (d1 * d1 * p2.y - d2 * d2 * p0.y + (2 * d1 * d1 + 3 * d1 * d2 + d2 * d2) * p1.y) / scale;
  }
  if (p3 && d3 > 0 && d2 > 0) {
    const scale = 3 * d3 * (d3 + d2);
    c2x = (d3 * d3 * p1.x - d2 * d2 * p3.x + (2 * d3 * d3 + 3 * d3 * d2 + d2 * d2) * p2.x) / scale;
    c2y = (d3 * d3 * p1.y - d2 * d2 * p3.y + (2 * d3 * d3 + 3 * d3 * d2 + d2 * d2) * p2.y) / scale;
  }
  return { c1x, c1y, c2x, c2y };
}

/** Whether the pen turned by more than `cornerDegrees` at `at` (its
 * neighbours `before` / `after`), measured as drawn. */
export function inkTurnIsCorner(
  before: InkSmoothPoint,
  at: InkSmoothPoint,
  after: InkSmoothPoint,
  cornerDegrees = INK_CORNER_DEGREES,
  unitX = 1,
  unitY = 1,
): boolean {
  const ax = (at.x - before.x) * unitX;
  const ay = (at.y - before.y) * unitY;
  const bx = (after.x - at.x) * unitX;
  const by = (after.y - at.y) * unitY;
  const la = Math.hypot(ax, ay);
  const lb = Math.hypot(bx, by);
  if (la === 0 || lb === 0) return false;
  return (ax * bx + ay * by) / (la * lb) < Math.cos((cornerDegrees * Math.PI) / 180);
}

/** The corner flags of `samples` (distinct points), one per sample. */
function cornersOf(samples: readonly InkSmoothPoint[], cornerDegrees: number, unitX: number, unitY: number): boolean[] {
  const corner: boolean[] = new Array<boolean>(samples.length).fill(false);
  for (let index = 1; index < samples.length - 1; index += 1) {
    corner[index] = inkTurnIsCorner(samples[index - 1]!, samples[index]!, samples[index + 1]!, cornerDegrees, unitX, unitY);
  }
  return corner;
}

/**
 * SVG path data for the stroke through `points`. One or two points are the
 * dot or the line they were; more are the curve.
 */
export function inkStrokePathData(
  points: readonly InkSmoothPoint[],
  options: InkSmoothOptions = {},
): string {
  const precision = options.precision ?? 2;
  const unitX = options.unit?.x && options.unit.x > 0 ? options.unit.x : 1;
  const unitY = options.unit?.y && options.unit.y > 0 ? options.unit.y : 1;
  const at = (point: InkSmoothPoint) => `${format(point.x, precision)} ${format(point.y, precision)}`;
  const samples = distinct(points);
  if (samples.length === 0) {
    const only = points.find((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
    return only ? `M${at(only)}L${at(only)}` : "";
  }
  if (samples.length === 1) return `M${at(samples[0]!)}L${at(samples[0]!)}`;
  if (samples.length === 2) return `M${at(samples[0]!)}L${at(samples[1]!)}`;

  const count = samples.length;
  const corner = cornersOf(samples, options.corner ?? INK_CORNER_DEGREES, unitX, unitY);
  let data = `M${at(samples[0]!)}`;
  for (let index = 0; index < count - 1; index += 1) {
    const p1 = samples[index]!;
    const p2 = samples[index + 1]!;
    // A neighbour across a corner, or past an end, is not followed: the
    // curve leaves the corner along its own segment.
    const p0 = index > 0 && !corner[index] ? samples[index - 1]! : null;
    const p3 = index + 2 < count && !corner[index + 1] ? samples[index + 2]! : null;
    const { c1x, c1y, c2x, c2y } = inkCurveControls(p0, p1, p2, p3, unitX, unitY);
    data += `C${format(c1x, precision)} ${format(c1y, precision)} ${format(c2x, precision)} ${format(c2y, precision)} ${at(p2)}`;
  }
  return data;
}

/**
 * How wide the pen draws at a pressure, as a factor of its width: a mouse
 * (pressure 0.5, the only value it reports) draws the width as stated, a
 * light pen touch thinner, a firm one wider. Bounded so no pressure doubles
 * or vanishes a stroke; without a pressure the width is the stated one.
 */
export function inkPressureWidthFactor(pressure: number | undefined): number {
  if (pressure === undefined || !Number.isFinite(pressure)) return 1;
  const clamped = Math.min(1, Math.max(0, pressure));
  return 0.6 + 0.8 * clamped;
}

/** Whether a stroke's pressures would draw it other than at one width. */
export function inkPressuresVary(pressures: readonly (number | undefined)[] | null | undefined): boolean {
  if (!pressures || pressures.length < 2) return false;
  let min = Infinity;
  let max = -Infinity;
  for (const pressure of pressures) {
    const factor = inkPressureWidthFactor(pressure);
    min = Math.min(min, factor);
    max = Math.max(max, factor);
  }
  return max - min > 0.04;
}

/** A point on the stroke's curve, with the width the pen drew there. */
interface InkOutlineSample {
  x: number;
  y: number;
  /** Half the width, in the same units as x. */
  half: number;
}

function cubicAt(p1: number, c1: number, c2: number, p2: number, t: number): number {
  const u = 1 - t;
  return u * u * u * p1 + 3 * u * u * t * c1 + 3 * u * t * t * c2 + t * t * t * p2;
}

/**
 * The stroke as a FILLED outline whose width follows the pen's pressure:
 * the same curve {@link inkStrokePathData} draws, walked at a few steps per
 * segment, with the pen's half width on each side along the curve's normal
 * and a round tip at each end. The outline is in the points' own units
 * (`unit` scales the width per axis for a stretched space); `widthPx` is the
 * pen's stated width in SCREEN pixels, each sample's `p` the pressure there.
 * Pressures missing on a sample draw the stated width.
 */
export function inkVariableWidthOutlinePathData(
  points: readonly InkPressurePoint[],
  widthPx: number,
  options: InkSmoothOptions = {},
): string {
  const precision = options.precision ?? 2;
  const unitX = options.unit?.x && options.unit.x > 0 ? options.unit.x : 1;
  const unitY = options.unit?.y && options.unit.y > 0 ? options.unit.y : 1;
  const samples = distinct(points);
  if (samples.length === 0) return "";
  const corner = cornersOf(samples, options.corner ?? INK_CORNER_DEGREES, unitX, unitY);
  // The curve, sampled: each step carries the pressure interpolated between
  // its segment's ends.
  const walk: InkOutlineSample[] = [];
  const half = (pressure: number | undefined) => (widthPx * inkPressureWidthFactor(pressure)) / 2;
  if (samples.length === 1) {
    walk.push({ x: samples[0]!.x, y: samples[0]!.y, half: half(samples[0]!.p) });
  } else {
    for (let index = 0; index < samples.length - 1; index += 1) {
      const p1 = samples[index]!;
      const p2 = samples[index + 1]!;
      const p0 = index > 0 && !corner[index] ? samples[index - 1]! : null;
      const p3 = index + 2 < samples.length && !corner[index + 1] ? samples[index + 2]! : null;
      const { c1x, c1y, c2x, c2y } = inkCurveControls(p0, p1, p2, p3, unitX, unitY);
      const chord = Math.hypot((p2.x - p1.x) * unitX, (p2.y - p1.y) * unitY);
      // A step every pixel and a half of curve: the outline follows the
      // curve, not chords of it (two or three chords a segment read as a
      // polygon once the page was zoomed in).
      const steps = Math.max(2, Math.min(OUTLINE_MAX_STEPS, Math.ceil(chord / OUTLINE_STEP_PX)));
      const h1 = half(p1.p);
      const h2 = half(p2.p);
      for (let step = index === 0 ? 0 : 1; step <= steps; step += 1) {
        const t = step / steps;
        walk.push({
          x: cubicAt(p1.x, c1x, c2x, p2.x, t),
          y: cubicAt(p1.y, c1y, c2y, p2.y, t),
          half: h1 + (h2 - h1) * t,
        });
      }
    }
  }
  // The outline: the left side forward, a round tip, the right side back,
  // a round tip home. Normals are measured as drawn and put back into the
  // points' own units, so the width is the pen's on screen.
  const fx = (value: number) => format(value, precision);
  const left: OutlinePoint[] = [];
  const right: OutlinePoint[] = [];
  const n = walk.length;
  for (let index = 0; index < n; index += 1) {
    const here = walk[index]!;
    const prev = walk[Math.max(0, index - 1)]!;
    const next = walk[Math.min(n - 1, index + 1)]!;
    let dx = (next.x - prev.x) * unitX;
    let dy = (next.y - prev.y) * unitY;
    const length = Math.hypot(dx, dy);
    if (length === 0) {
      dx = 1;
      dy = 0;
    } else {
      dx /= length;
      dy /= length;
    }
    // The normal, in screen pixels, scaled back into each axis's units.
    const nx = (-dy * here.half) / unitX;
    const ny = (dx * here.half) / unitY;
    left.push([here.x + nx, here.y + ny]);
    right.push([here.x - nx, here.y - ny]);
  }
  const at = ([x, y]: OutlinePoint) => `${fx(x)} ${fx(y)}`;
  const tip = (at: InkOutlineSample, to: string, flat: boolean | undefined) => {
    // A round tip: an arc of the pen's half width (an ellipse in a
    // stretched space, as the width is); a cut end stops flat. The left side
    // is the normal (-dy, dx) on y-down axes, so turning from it to the right
    // side ahead of the pen is counter-clockwise on screen: sweep 0. Sweep 1
    // swung the arc back into the stroke and its winding cut a half-moon bite
    
    if (flat) return `L${to}`;
    const rx = at.half / unitX;
    const ry = at.half / unitY;
    return `A${fx(rx)} ${fx(ry)} 0 0 0 ${to}`;
  };
  const first = walk[0]!;
  const last = walk[n - 1]!;
  const backward = right.slice().reverse();
  let data = `M${at(left[0]!)}${smoothSide(left, at)}`;
  data += tip(last, at(right[n - 1]!), options.flatEnd);
  data += smoothSide(backward, at);
  data += tip(first, at(left[0]!), options.flatStart);
  return `${data}Z`;
}

type OutlinePoint = [number, number];

/** How far apart (screen px at 100 %) the outline is sampled along the curve. */
const OUTLINE_STEP_PX = 1.5;
/** The most samples one segment takes (a long jump of the pen). */
const OUTLINE_MAX_STEPS = 48;

/**
 * One side of the outline from its first point (already the current point)
 * to its last: a quadratic curve through each sample's midpoints with the
 * sample as its control, as perfect-freehand and the note apps draw a
 * pressure stroke's edge. The side passes smoothly through every turn
 * instead of breaking at each sample; its ends are the samples themselves,
 * where the tips join.
 */
function smoothSide(side: readonly OutlinePoint[], at: (point: OutlinePoint) => string): string {
  const n = side.length;
  if (n < 2) return "";
  if (n === 2) return `L${at(side[1]!)}`;
  let data = "";
  for (let index = 1; index < n - 1; index += 1) {
    const here = side[index]!;
    const next = side[index + 1]!;
    const mid: OutlinePoint = index === n - 2 ? next : [(here[0] + next[0]) / 2, (here[1] + next[1]) / 2];
    data += `Q${at(here)} ${at(mid)}`;
  }
  return data;
}

/** The points of an SVG `points` list ("x,y x,y ..."), the ones that are numbers. */
export function inkPointsOfList(list: string | null | undefined): InkSmoothPoint[] {
  if (!list) return [];
  const points: InkSmoothPoint[] = [];
  for (const pair of list.trim().split(/\s+/)) {
    const [x, y] = pair.split(",").map(Number);
    if (Number.isFinite(x) && Number.isFinite(y)) points.push({ x: x!, y: y! });
  }
  return points;
}

/**
 * Path data for a stroke stored as a `points` list in its own 0..100 box,
 * drawn in a box `width` × `height` on screen (the corner test needs the
 * box's shape; without a size the box is taken as square).
 */
export function inkStrokePathDataOfList(
  list: string | null | undefined,
  box?: { readonly width: number; readonly height: number } | null,
  /** The drawing tool's corner rule ({@link inkCornerDegreesForTool}): the one the stroke was drawn with while the pen was down. */
  corner?: number,
): string {
  const unit =
    box && box.width > 0 && box.height > 0 ? { x: box.width / 100, y: box.height / 100 } : undefined;
  return inkStrokePathData(inkPointsOfList(list), { unit, precision: 3, ...(corner === undefined ? {} : { corner }) });
}

/**
 * How many samples a stroke being drawn may hold before it is thinned, and
 * how it is thinned: every other sample goes (the first and last stay) and
 * from then on samples are taken twice as far apart. A surface that stopped
 * taking samples at its limit cut a long stroke short while the pen was
 * still down.
 */
export interface InkSampleBudget {
  /** The least distance between two samples, as it stands. */
  spacing: number;
}

export function thinInkSamples<P>(samples: P[], limit: number, budget: InkSampleBudget): void {
  if (samples.length < limit || samples.length < 4) return;
  const last = samples[samples.length - 1]!;
  let write = 1;
  for (let read = 2; read < samples.length - 1; read += 2) {
    samples[write] = samples[read]!;
    write += 1;
  }
  samples[write] = last;
  samples.length = write + 1;
  budget.spacing *= 2;
}

/**
 * The width factor a stroke whose pressure held steady is drawn at: its mean
 * pressure's (1 when it has none, or when its pressure varies -- that stroke
 * is drawn along its pressures instead). The running stroke is drawn at each
 * sample's pressure; a steady stroke landed at its stated width, so a firm
 * hand's lines thinned as the pen lifted (a scribble's gaps opened) and a
 * light one's thickened. Landed at this factor, it stays as it was drawn.
 */
export function inkSteadyPressureWidthFactor(
  pressures: readonly (number | undefined)[] | null | undefined,
): number {
  if (!pressures || pressures.length === 0 || inkPressuresVary(pressures)) return 1;
  const known = pressures.filter((pressure): pressure is number => typeof pressure === "number" && Number.isFinite(pressure));
  if (known.length === 0) return 1;
  return inkPressureWidthFactor(known.reduce((sum, pressure) => sum + pressure, 0) / known.length);
}

/**
 * How many samples a stroke being drawn may hold before it is thinned. The
 * stored stroke is simplified once, as it lands ({@link inkKeptSampleIndexes});
 * thinning while the pen was down (every other sample, at 512) cut the tips
 * off a long scribble's reversals and the stroke drew in on the lift.
 */
export const INK_CAPTURE_MAX_POINTS = 8192;
























export function inkKeptSampleIndexes(
  points: readonly (InkSmoothPoint & { readonly p?: number })[],
  tolerance: number,
  limit: number,
  /** A sample whose pressure is further than this from its neighbours' run stays too (the width it drew). */
  pressureTolerance = 0.025,
  options: { readonly fillToLimit?: boolean } = {},
): number[] {
  const n = points.length;
  const cap = Math.max(2, Math.floor(limit));
  if (n <= 2) return points.map((_, index) => index);
  const tol = Math.max(0, tolerance);
  const pressureScale = pressureTolerance > 0 ? (tol > 0 ? tol : 1) / pressureTolerance : 0;
  // Each sample scanned costs about the same: past this many, the stroke is
  // filled in evenly (never reached by a stroke a hand draws; a guard).
  let budget = Math.max(64 * n, 200_000);
  // The worst sample of the run a..b: [index, how far off] ([-1, 0] if none).
  const worst = (a: number, b: number): [number, number] => {
    const pa = points[a]!;
    const pb = points[b]!;
    const dx = pb.x - pa.x;
    const dy = pb.y - pa.y;
    const length2 = dx * dx + dy * dy;
    const pressA = pa.p;
    const pressB = pb.p;
    const pressured = pressureScale > 0 && typeof pressA === "number" && typeof pressB === "number";
    const middle = (a + b) / 2;
    let best = -1;
    let bestError = 0;
    budget -= b - a;
    for (let index = a + 1; index < b; index += 1) {
      const point = points[index]!;
      const t = length2 > 0 ? Math.min(1, Math.max(0, ((point.x - pa.x) * dx + (point.y - pa.y) * dy) / length2)) : 0;
      let error = Math.hypot(point.x - (pa.x + t * dx), point.y - (pa.y + t * dy));
      if (pressured && typeof point.p === "number") {
        const along = (index - a) / (b - a);
        const off = Math.abs(point.p - (pressA! + (pressB! - pressA!) * along)) * pressureScale;
        if (off > error) error = off;
      }
      // The first of equal misses: along a line drawn out and back, each
      // reversal's tip is off the run only until an earlier one is kept.
      if (error > bestError) {
        best = index;
        bestError = error;
      }
    }
    // A run with nothing off it (samples on their segment): its middle, for
    // a stroke filled to its limit -- dropped or kept, it draws the same.
    if (best < 0) best = Math.floor(middle);
    return [best, bestError];
  };
  // A max-heap of runs by their worst miss.
  const runs: Array<{ a: number; b: number; at: number; error: number }> = [];
  const push = (a: number, b: number) => {
    if (b - a < 2) return;
    const [at, error] = worst(a, b);
    if (at < 0) return;
    const run = { a, b, at, error };
    runs.push(run);
    let child = runs.length - 1;
    while (child > 0) {
      const parent = (child - 1) >> 1;
      if (runs[parent]!.error >= run.error) break;
      runs[child] = runs[parent]!;
      child = parent;
    }
    runs[child] = run;
  };
  const pop = () => {
    const top = runs[0]!;
    const last = runs.pop()!;
    if (runs.length > 0) {
      let parent = 0;
      for (;;) {
        const left = parent * 2 + 1;
        if (left >= runs.length) break;
        const right = left + 1;
        const child = right < runs.length && runs[right]!.error > runs[left]!.error ? right : left;
        if (runs[child]!.error <= last.error) break;
        runs[parent] = runs[child]!;
        parent = child;
      }
      runs[parent] = last;
    }
    return top;
  };
  const keep = new Uint8Array(n);
  keep[0] = 1;
  keep[n - 1] = 1;
  let kept = 2;
  push(0, n - 1);
  while (runs.length > 0 && kept < cap && budget > 0) {
    if (!options.fillToLimit && runs[0]!.error <= tol) break;
    const run = pop();
    keep[run.at] = 1;
    kept += 1;
    push(run.a, run.at);
    push(run.at, run.b);
  }
  if (budget <= 0 && kept < cap && runs.length > 0) {
    // Past the budget: the rest evenly from the samples of the runs still open.
    const open: number[] = [];
    for (const run of runs) for (let index = run.a + 1; index < run.b; index += 1) if (!keep[index]) open.push(index);
    open.sort((x, y) => x - y);
    const want = Math.min(cap - kept, open.length);
    for (let step = 0; step < want; step += 1) keep[open[Math.floor(((step + 0.5) * open.length) / want)]!] = 1;
    kept += want;
  }
  const indexes: number[] = [];
  for (let index = 0; index < n; index += 1) if (keep[index]) indexes.push(index);
  return indexes;
}

/**
 * Which samples a stroke keeps as it lands: every one, when they fit the
 * stored stroke's limit. The running stroke and the stroke shown on the lift
 * are drawn through every sample; the landed one drawn through fewer is a
 * different curve (a sample dropped for lying within a fifth of a pixel of
 * the segment its neighbours draw still bends the curve through them: up to
 * 1.3px off, 2.6px at 200 %), and the stroke moved as it landed. A stroke
 * past the limit is simplified ({@link inkKeptSampleIndexes}) to fit it.
 */
export function inkLandedSampleIndexes(
  points: readonly (InkSmoothPoint & { readonly p?: number })[],
  tolerance: number,
  limit: number,
): number[] {
  if (points.length <= Math.max(2, limit)) return points.map((_, index) => index);
  // Over the limit: the samples it misses least go, and only as many as
  // must -- simplified to a tolerance instead, it dropped far more and the
  // landed curve moved off the drawn one (a release review: 1.3px on 1,104
  // samples).
  return inkKeptSampleIndexes(points, tolerance, limit, undefined, { fillToLimit: true });
}
