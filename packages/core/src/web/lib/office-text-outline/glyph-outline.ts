/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A glyph's outline for a WordArt text outline, as Office draws one: the
 * band is a centred, round-joined stroke along the glyph's contours MERGED,
 * so no line is drawn inside a letter where its contours overlap (most CJK
 * glyphs, accented letters built from components, some Latin faces). A
 * glyph whose contours do not overlap (most Latin glyphs) is stroked along
 * its contours as they stand; one whose contours overlap is first unioned
 * (Skia's PathOps through PathKit, as the server's PDFs union with
 * skia-pathops: app.services.office_text_outline).
 *
 * Pure geometry over the sfnt reader's glyphs (font units, y up); the
 * outline worker (office-text-outline.worker.ts) runs it off the page's
 * thread.
 */
import {
  COMPOSITE_ARGS_ARE_XY_VALUES,
  COMPOSITE_WE_HAVE_A_SCALE,
  COMPOSITE_WE_HAVE_A_TWO_BY_TWO,
  COMPOSITE_WE_HAVE_AN_X_AND_Y_SCALE,
  type SfntGlyph,
  type SfntProgram,
} from "../pdf/engine/native/sfnt/sfnt-reader";

type Point = readonly [number, number];
/** One closed contour: a start point and its segments (a line to a point,
 * or a quadratic through a control point to a point). */
export interface GlyphContour {
  readonly start: Point;
  readonly segments: readonly (readonly [Point] | readonly [Point, Point])[];
}

/** An affine map of font units (a b c d e f, as SVG's matrix()). */
type Affine = readonly [number, number, number, number, number, number];
const IDENTITY: Affine = [1, 0, 0, 1, 0, 0];
const apply = (m: Affine, x: number, y: number): Point => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const compose = (outer: Affine, inner: Affine): Affine => [
  outer[0] * inner[0] + outer[2] * inner[1],
  outer[1] * inner[0] + outer[3] * inner[1],
  outer[0] * inner[2] + outer[2] * inner[3],
  outer[1] * inner[2] + outer[3] * inner[3],
  outer[0] * inner[4] + outer[2] * inner[5] + outer[4],
  outer[1] * inner[4] + outer[3] * inner[5] + outer[5],
];

const F2DOT14 = 1 / 16384;

/**
 * The contours of glyph `glyphId`, components placed (TrueType composites:
 * offsets and their scale or 2x2); null when a component is placed by point
 * matching, which this does not follow (the caller draws no band of its own
 * for that glyph).
 */
export function glyphContours(program: SfntProgram, glyphId: number): GlyphContour[] | null {
  const out: GlyphContour[] = [];
  const walk = (glyph: SfntGlyph | null, m: Affine, depth: number): boolean => {
    if (!glyph) return true;
    if (depth > 8) return false;
    if (glyph.kind === "simple") {
      let first = 0;
      for (const last of glyph.endPoints) {
        const contour = simpleContour(glyph.xs, glyph.ys, glyph.onCurve, first, last, m);
        if (contour) out.push(contour);
        first = last + 1;
      }
      return true;
    }
    for (const component of glyph.components) {
      if (!(component.flags & COMPOSITE_ARGS_ARE_XY_VALUES)) return false;
      const t = component.transform.map((value) => value * F2DOT14);
      let linear: Affine = IDENTITY;
      if (component.flags & COMPOSITE_WE_HAVE_A_SCALE) linear = [t[0], 0, 0, t[0], 0, 0];
      else if (component.flags & COMPOSITE_WE_HAVE_AN_X_AND_Y_SCALE) linear = [t[0], 0, 0, t[1], 0, 0];
      else if (component.flags & COMPOSITE_WE_HAVE_A_TWO_BY_TWO) linear = [t[0], t[1], t[2], t[3], 0, 0];
      // The offset is not scaled (Microsoft's default, SCALED_COMPONENT_OFFSET unset).
      const placed: Affine = [linear[0], linear[1], linear[2], linear[3], component.arg1, component.arg2];
      let child: SfntGlyph | null;
      try {
        child = program.glyph(component.glyphIndex);
      } catch {
        return false;
      }
      if (!walk(child, compose(m, placed), depth + 1)) return false;
    }
    return true;
  };
  let glyph: SfntGlyph | null;
  try {
    glyph = program.glyph(glyphId);
  } catch {
    return null;
  }
  return walk(glyph, IDENTITY, 0) ? out : null;
}

function simpleContour(
  xs: readonly number[], ys: readonly number[], on: readonly boolean[], first: number, last: number, m: Affine,
): GlyphContour | null {
  const count = last - first + 1;
  if (count < 2) return null;
  const at = (index: number): Point => {
    const i = first + ((index % count) + count) % count;
    return apply(m, xs[i], ys[i]);
  };
  const isOn = (index: number) => on[first + ((index % count) + count) % count];
  // Start on an on-curve point; with none, midway between the first two.
  let startIndex = -1;
  for (let index = 0; index < count; index += 1) {
    if (isOn(index)) { startIndex = index; break; }
  }
  const mid = (a: Point, b: Point): Point => [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
  let start: Point;
  let begin: number;
  if (startIndex < 0) {
    start = mid(at(0), at(1));
    begin = 1;
  } else {
    start = at(startIndex);
    begin = startIndex + 1;
  }
  const segments: (readonly [Point] | readonly [Point, Point])[] = [];
  let control: Point | null = null;
  for (let step = 0; step < count; step += 1) {
    const index = begin + step;
    const point = at(index);
    const reachedStart = step === count - 1 && startIndex >= 0;
    if (isOn(index)) {
      if (control) segments.push([control, point]);
      else segments.push([point]);
      control = null;
    } else if (control) {
      const implied = mid(control, point);
      segments.push([control, implied]);
      control = point;
    } else {
      control = point;
    }
    if (reachedStart) break;
  }
  if (control) segments.push([control, start]);
  return { start, segments };
}

/** The contours as SVG path data in font units (y up), `digits` decimals. */
export function contoursPathData(contours: readonly GlyphContour[], digits = 2): string {
  const f = (value: number) => {
    const rounded = Number(value.toFixed(digits));
    return Object.is(rounded, -0) ? "0" : String(rounded);
  };
  let d = "";
  for (const contour of contours) {
    d += `M${f(contour.start[0])} ${f(contour.start[1])}`;
    for (const segment of contour.segments) {
      if (segment.length === 1) d += `L${f(segment[0][0])} ${f(segment[0][1])}`;
      else d += `Q${f(segment[0][0])} ${f(segment[0][1])} ${f(segment[1][0])} ${f(segment[1][1])}`;
    }
    d += "Z";
  }
  return d;
}

/** Each contour as a closed polygon (quadratics flattened), x,y pairs. */
function polygons(contours: readonly GlyphContour[]): Float64Array[] {
  return contours.map((contour) => {
    const points: number[] = [contour.start[0], contour.start[1]];
    let [px, py] = contour.start;
    for (const segment of contour.segments) {
      if (segment.length === 1) {
        [px, py] = segment[0];
        points.push(px, py);
        continue;
      }
      const [[cx, cy], [ex, ey]] = segment;
      for (let step = 1; step <= 8; step += 1) {
        const t = step / 8;
        const u = 1 - t;
        points.push(u * u * px + 2 * u * t * cx + t * t * ex, u * u * py + 2 * u * t * cy + t * t * ey);
      }
      [px, py] = [ex, ey];
    }
    return Float64Array.from(points);
  });
}

function signedArea(polygon: Float64Array): number {
  let area = 0;
  const n = polygon.length / 2;
  for (let i = 0, j = n - 1; i < n; j = i, i += 1) {
    area += polygon[j * 2] * polygon[i * 2 + 1] - polygon[i * 2] * polygon[j * 2 + 1];
  }
  return area / 2;
}

function bounds(polygon: Float64Array): [number, number, number, number] {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (let i = 0; i < polygon.length; i += 2) {
    const x = polygon[i], y = polygon[i + 1];
    if (x < x0) x0 = x;
    if (x > x1) x1 = x;
    if (y < y0) y0 = y;
    if (y > y1) y1 = y;
  }
  return [x0, y0, x1, y1];
}

/** Whether segments p1-p2 and p3-p4 cross or touch (collinear overlap counts). */
function segmentsMeet(
  x1: number, y1: number, x2: number, y2: number, x3: number, y3: number, x4: number, y4: number,
): boolean {
  if (Math.max(x1, x2) < Math.min(x3, x4) || Math.max(x3, x4) < Math.min(x1, x2)) return false;
  if (Math.max(y1, y2) < Math.min(y3, y4) || Math.max(y3, y4) < Math.min(y1, y2)) return false;
  const d1 = (x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3);
  const d2 = (x4 - x3) * (y2 - y3) - (y4 - y3) * (x2 - x3);
  const d3 = (x2 - x1) * (y3 - y1) - (y2 - y1) * (x3 - x1);
  const d4 = (x2 - x1) * (y4 - y1) - (y2 - y1) * (x4 - x1);
  if (((d1 > 0 && d2 < 0) || (d1 < 0 && d2 > 0)) && ((d3 > 0 && d4 < 0) || (d3 < 0 && d4 > 0))) return true;
  // Collinear and overlapping: the contours share an edge.
  return d1 === 0 && d2 === 0 && d3 === 0 && d4 === 0;
}

function pointInPolygon(x: number, y: number, polygon: Float64Array): boolean {
  let inside = false;
  const n = polygon.length / 2;
  for (let i = 0, j = n - 1; i < n; j = i, i += 1) {
    const xi = polygon[i * 2], yi = polygon[i * 2 + 1], xj = polygon[j * 2], yj = polygon[j * 2 + 1];
    if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

/**
 * Whether any point of the glyph is covered twice under the non-zero rule
 * TrueType and CFF fill by: two contours crossing (or sharing an edge), a
 * contour crossing itself, or a contour inside another turning the same way
 * (a counter turns the other way and is a hole, not an overlap).
 */
export function contoursOverlap(contours: readonly GlyphContour[]): boolean {
  const shapes = polygons(contours).filter((polygon) => polygon.length >= 6);
  const boxes = shapes.map(bounds);
  const areas = shapes.map(signedArea);
  for (let a = 0; a < shapes.length; a += 1) {
    const pa = shapes[a];
    const na = pa.length / 2;
    // A contour crossing itself.
    for (let i = 0; i < na; i += 1) {
      const i2 = (i + 1) % na;
      for (let j = i + 2; j < na; j += 1) {
        const j2 = (j + 1) % na;
        if (j2 === i) continue;
        if (segmentsMeet(pa[i * 2], pa[i * 2 + 1], pa[i2 * 2], pa[i2 * 2 + 1], pa[j * 2], pa[j * 2 + 1], pa[j2 * 2], pa[j2 * 2 + 1])) {
          // Adjacent flattening steps of one curve can touch at a shared end only.
          return true;
        }
      }
    }
    for (let b = a + 1; b < shapes.length; b += 1) {
      const ba = boxes[a], bb = boxes[b];
      if (ba[2] < bb[0] || bb[2] < ba[0] || ba[3] < bb[1] || bb[3] < ba[1]) continue;
      const pb = shapes[b];
      const nb = pb.length / 2;
      for (let i = 0; i < na; i += 1) {
        const i2 = (i + 1) % na;
        const sx0 = Math.min(pa[i * 2], pa[i2 * 2]), sx1 = Math.max(pa[i * 2], pa[i2 * 2]);
        const sy0 = Math.min(pa[i * 2 + 1], pa[i2 * 2 + 1]), sy1 = Math.max(pa[i * 2 + 1], pa[i2 * 2 + 1]);
        if (sx1 < bb[0] || sx0 > bb[2] || sy1 < bb[1] || sy0 > bb[3]) continue;
        for (let j = 0; j < nb; j += 1) {
          const j2 = (j + 1) % nb;
          if (segmentsMeet(pa[i * 2], pa[i * 2 + 1], pa[i2 * 2], pa[i2 * 2 + 1], pb[j * 2], pb[j * 2 + 1], pb[j2 * 2], pb[j2 * 2 + 1])) return true;
        }
      }
      // No crossing: one inside the other turning the same way is covered twice.
      if (Math.sign(areas[a]) === Math.sign(areas[b])) {
        if (ba[0] >= bb[0] && ba[2] <= bb[2] && ba[1] >= bb[1] && ba[3] <= bb[3] && pointInPolygon(pa[0], pa[1], pb)) return true;
        if (bb[0] >= ba[0] && bb[2] <= ba[2] && bb[1] >= ba[1] && bb[3] <= ba[3] && pointInPolygon(pb[0], pb[1], pa)) return true;
      }
    }
  }
  return false;
}

/** What the union needs of PathKit (pathkit-wasm, Skia's PathOps). */
export interface PathUnion {
  /** `d` (non-zero) with its overlaps merged into one outline; null when it could not be. */
  union(d: string): string | null;
}

export interface GlyphOutline {
  /** SVG path data in font units, y up; "" for an empty glyph. */
  readonly d: string;
  /** Whether the contours overlap (and so are to be merged). */
  readonly overlaps: boolean;
  /** Whether `d` is the merged outline. */
  readonly merged: boolean;
  /** Font-unit bounds [xMin, yMin, xMax, yMax]; null when empty. */
  readonly bbox: readonly [number, number, number, number] | null;
}

/** The outline a glyph's band follows; null when the glyph cannot be read. */
export function glyphOutline(program: SfntProgram, glyphId: number, pathUnion: PathUnion | null): GlyphOutline | null {
  const contours = glyphContours(program, glyphId);
  if (!contours) return null;
  if (!contours.length) return { d: "", overlaps: false, merged: false, bbox: null };
  const d = contoursPathData(contours);
  const shapes = polygons(contours);
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
  for (const shape of shapes) {
    const [a, b, c, e] = bounds(shape);
    x0 = Math.min(x0, a); y0 = Math.min(y0, b); x1 = Math.max(x1, c); y1 = Math.max(y1, e);
  }
  const bbox = [x0, y0, x1, y1] as const;
  if (!contoursOverlap(contours)) return { d, overlaps: false, merged: false, bbox };
  const merged = pathUnion?.union(d) ?? null;
  // No union to hand, or PathOps declined it: the contours as they stand
  // (the band then draws the lines inside where they overlap).
  return merged === null ? { d, overlaps: true, merged: false, bbox } : { d: merged, overlaps: true, merged: true, bbox };
}
