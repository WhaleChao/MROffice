/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Snapping for the measuring tool (WPS's 测量 tab: snap to end points,
 * paths, mid points and intersections; Acrobat's Snap to paths / end points /
 * mid points / intersections).
 *
 * The page's line art is read from pdf.js's operator list (each
 * `constructPath` carries its path as pdf.js's DrawOPS codes in the user
 * space of the moment, under the transforms and form matrices around it)
 * and kept as straight segments in the measurement's own space: percent of
 * the page as it shows (rotation applied), the space measurement points are
 * clicked in. Curves are kept as a few chords: enough to snap along them.
 */

export type MeasureSnapKind = "endpoint" | "path" | "midpoint" | "intersection";

export const MEASURE_SNAP_KINDS: readonly MeasureSnapKind[] = [
  "endpoint",
  "path",
  "midpoint",
  "intersection",
];

/** A straight piece of the page's line art, in page percent. */
export interface MeasureSnapSegment {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

export interface MeasureSnapPoint {
  readonly x: number;
  readonly y: number;
}

type Matrix = readonly [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
/** pdf.js DrawOPS: the codes a constructPath's path data is written in. */
const DRAW_MOVE = 0;
const DRAW_LINE = 1;
const DRAW_CURVE = 2;
const DRAW_QUADRATIC = 3;
const DRAW_CLOSE = 4;
const CURVE_CHORDS = 6;
/** A page of dense line art (a map, a scan's vector trace) stops here. */
export const MEASURE_SNAP_MAX_SEGMENTS = 20_000;

function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ];
}

function apply(m: Matrix, x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

function asMatrix(value: unknown): Matrix | null {
  if (!value || typeof (value as ArrayLike<number>).length !== "number") return null;
  const list = Array.from(value as ArrayLike<number>);
  if (list.length < 6 || !list.slice(0, 6).every(Number.isFinite)) return null;
  return [list[0], list[1], list[2], list[3], list[4], list[5]];
}

/** The pdf.js operator codes the reader needs (`pdfjs.OPS`). */
export interface MeasureSnapOps {
  readonly save: number;
  readonly restore: number;
  readonly transform: number;
  readonly constructPath: number;
  readonly paintFormXObjectBegin: number;
  readonly paintFormXObjectEnd: number;
}

/**
 * The straight segments of a page's line art, in page percent.
 *
 * `viewportTransform` and `viewportSize` are those of the page's viewport at
 * scale 1 with the rotation it shows (`page.getViewport({ scale: 1, rotation })`).
 */
export function measureSnapSegmentsFromOperatorList(
  operatorList: { readonly fnArray: ArrayLike<number>; readonly argsArray: ArrayLike<unknown> },
  ops: MeasureSnapOps,
  viewportTransform: ArrayLike<number>,
  viewportSize: { readonly width: number; readonly height: number },
): MeasureSnapSegment[] {
  const viewport = asMatrix(viewportTransform);
  if (!viewport || !(viewportSize.width > 0) || !(viewportSize.height > 0)) return [];
  const toPercent = (m: Matrix, x: number, y: number): [number, number] => {
    const [ux, uy] = apply(m, x, y);
    const [vx, vy] = apply(viewport, ux, uy);
    return [(vx / viewportSize.width) * 100, (vy / viewportSize.height) * 100];
  };
  const segments: MeasureSnapSegment[] = [];
  const push = (from: [number, number], to: [number, number]) => {
    if (segments.length >= MEASURE_SNAP_MAX_SEGMENTS) return;
    if (from[0] === to[0] && from[1] === to[1]) return;
    if (![...from, ...to].every(Number.isFinite)) return;
    segments.push({ x1: from[0], y1: from[1], x2: to[0], y2: to[1] });
  };
  let ctm: Matrix = IDENTITY;
  const stack: Matrix[] = [];
  const { fnArray, argsArray } = operatorList;
  for (let index = 0; index < fnArray.length; index += 1) {
    if (segments.length >= MEASURE_SNAP_MAX_SEGMENTS) break;
    const fn = fnArray[index];
    const args = argsArray[index] as unknown[] | null | undefined;
    if (fn === ops.save) {
      stack.push(ctm);
    } else if (fn === ops.restore) {
      ctm = stack.pop() ?? IDENTITY;
    } else if (fn === ops.transform) {
      const m = asMatrix(args);
      if (m) ctm = multiply(m, ctm);
    } else if (fn === ops.paintFormXObjectBegin) {
      stack.push(ctm);
      const m = asMatrix(args?.[0]);
      if (m) ctm = multiply(m, ctm);
    } else if (fn === ops.paintFormXObjectEnd) {
      ctm = stack.pop() ?? IDENTITY;
    } else if (fn === ops.constructPath) {
      const holder = args?.[1] as ArrayLike<unknown> | undefined;
      const data = holder?.[0] as ArrayLike<number> | undefined;
      if (!data || typeof data.length !== "number") continue;
      let start: [number, number] | null = null;
      let current: [number, number] | null = null;
      let cursor = 0;
      while (cursor < data.length) {
        const code = data[cursor];
        if (code === DRAW_MOVE) {
          current = start = toPercent(ctm, data[cursor + 1], data[cursor + 2]);
          cursor += 3;
        } else if (code === DRAW_LINE) {
          const next = toPercent(ctm, data[cursor + 1], data[cursor + 2]);
          if (current) push(current, next);
          current = next;
          cursor += 3;
        } else if (code === DRAW_CURVE || code === DRAW_QUADRATIC) {
          const cubic = code === DRAW_CURVE;
          const values = Array.from({ length: cubic ? 6 : 4 }, (_, k) => data[cursor + 1 + k]);
          const end = toPercent(ctm, values[values.length - 2], values[values.length - 1]);
          if (current) {
            const p0 = current;
            const c1 = toPercent(ctm, values[0], values[1]);
            const c2 = cubic ? toPercent(ctm, values[2], values[3]) : c1;
            let previous = p0;
            for (let step = 1; step <= CURVE_CHORDS; step += 1) {
              const t = step / CURVE_CHORDS;
              const u = 1 - t;
              const point: [number, number] = cubic
                ? [
                    u * u * u * p0[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t * t * t * end[0],
                    u * u * u * p0[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t * t * t * end[1],
                  ]
                : [
                    u * u * p0[0] + 2 * u * t * c1[0] + t * t * end[0],
                    u * u * p0[1] + 2 * u * t * c1[1] + t * t * end[1],
                  ];
              push(previous, point);
              previous = point;
            }
          }
          current = end;
          cursor += cubic ? 7 : 5;
        } else if (code === DRAW_CLOSE) {
          if (current && start) push(current, start);
          current = start;
          cursor += 1;
        } else {
          break;
        }
      }
    }
  }
  return segments;
}

export interface MeasureSnapResult extends MeasureSnapPoint {
  readonly kind: MeasureSnapKind;
}

/**
 * The point a click snaps to, or null. Distances are measured on the page as
 * it shows (`pageWidth` × `pageHeight` in any one unit, e.g. CSS pixels), so
 * `tolerance` is in that unit. End points win over intersections, those over
 * mid points, those over the nearest point along a path.
 */
export function snapMeasurePoint(options: {
  readonly point: MeasureSnapPoint;
  readonly segments: readonly MeasureSnapSegment[];
  readonly kinds: ReadonlySet<MeasureSnapKind>;
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly tolerance: number;
}): MeasureSnapResult | null {
  const { point, segments, kinds, pageWidth, pageHeight, tolerance } = options;
  if (kinds.size === 0 || !(pageWidth > 0) || !(pageHeight > 0)) return null;
  const sx = pageWidth / 100;
  const sy = pageHeight / 100;
  const distance = (x: number, y: number) => Math.hypot((x - point.x) * sx, (y - point.y) * sy);
  // Only the segments passing near the click take part.
  const near = segments.filter((segment) => {
    const minX = Math.min(segment.x1, segment.x2);
    const maxX = Math.max(segment.x1, segment.x2);
    const minY = Math.min(segment.y1, segment.y2);
    const maxY = Math.max(segment.y1, segment.y2);
    const dx = Math.max(minX - point.x, 0, point.x - maxX) * sx;
    const dy = Math.max(minY - point.y, 0, point.y - maxY) * sy;
    return Math.hypot(dx, dy) <= tolerance;
  });
  if (near.length === 0) return null;
  const best = (candidates: MeasureSnapPoint[], kind: MeasureSnapKind): MeasureSnapResult | null => {
    let found: MeasureSnapResult | null = null;
    let foundDistance = tolerance;
    for (const candidate of candidates) {
      const d = distance(candidate.x, candidate.y);
      if (d <= foundDistance) {
        found = { x: candidate.x, y: candidate.y, kind };
        foundDistance = d;
      }
    }
    return found;
  };
  if (kinds.has("endpoint")) {
    const hit = best(near.flatMap((s) => [{ x: s.x1, y: s.y1 }, { x: s.x2, y: s.y2 }]), "endpoint");
    if (hit) return hit;
  }
  if (kinds.has("intersection")) {
    const crossings: MeasureSnapPoint[] = [];
    const limit = Math.min(near.length, 400);
    for (let i = 0; i < limit; i += 1) {
      for (let j = i + 1; j < limit; j += 1) {
        const crossing = segmentIntersection(near[i], near[j], sx, sy);
        if (crossing) crossings.push(crossing);
      }
    }
    const hit = best(crossings, "intersection");
    if (hit) return hit;
  }
  if (kinds.has("midpoint")) {
    const hit = best(near.map((s) => ({ x: (s.x1 + s.x2) / 2, y: (s.y1 + s.y2) / 2 })), "midpoint");
    if (hit) return hit;
  }
  if (kinds.has("path")) {
    const projections = near.map((s) => {
      const ax = s.x1 * sx;
      const ay = s.y1 * sy;
      const bx = s.x2 * sx;
      const by = s.y2 * sy;
      const px = point.x * sx;
      const py = point.y * sy;
      const lengthSquared = (bx - ax) ** 2 + (by - ay) ** 2;
      const t = lengthSquared > 0
        ? Math.max(0, Math.min(1, ((px - ax) * (bx - ax) + (py - ay) * (by - ay)) / lengthSquared))
        : 0;
      return { x: (ax + t * (bx - ax)) / sx, y: (ay + t * (by - ay)) / sy };
    });
    return best(projections, "path");
  }
  return null;
}

function segmentIntersection(
  a: MeasureSnapSegment,
  b: MeasureSnapSegment,
  sx: number,
  sy: number,
): MeasureSnapPoint | null {
  const x1 = a.x1 * sx, y1 = a.y1 * sy, x2 = a.x2 * sx, y2 = a.y2 * sy;
  const x3 = b.x1 * sx, y3 = b.y1 * sy, x4 = b.x2 * sx, y4 = b.y2 * sy;
  const denominator = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
  if (Math.abs(denominator) < 1e-9) return null;
  const t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denominator;
  const u = -((x1 - x2) * (y1 - y3) - (y1 - y2) * (x1 - x3)) / denominator;
  if (t < 0 || t > 1 || u < 0 || u > 1) return null;
  return { x: (x1 + t * (x2 - x1)) / sx, y: (y1 + t * (y2 - y1)) / sy };
}
