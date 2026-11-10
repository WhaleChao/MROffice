/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Edit Points on the ONE custom-geometry model (`a:custGeom`, stored as
 * `PresentationCustomPath` on the shared 0..100000 grid) that Docs, the deck
 * and the worksheet persist.
 *
 * Office semantics (Microsoft Support, "Draw or edit a freeform shape"; Shape
 * Format > Edit Shape > Edit Points in Word, Excel and PowerPoint):
 *  - Edit Points shows the outline's vertices as black squares; dragging one
 *    moves it, and its two Bezier handles travel with it.
 *  - Ctrl+click on the outline adds a vertex; Ctrl+click on a vertex deletes it.
 *  - Dragging a handle with Shift makes a smooth point (both handles on one
 *    line, equal length), with Ctrl a straight point (one line, lengths kept),
 *    with Alt a corner point (the handles move independently).
 *  - The vertex menu offers Add Point, Delete Point, Open Path / Close Path,
 *    Smooth / Straight / Corner Point; the segment menu Straight Segment /
 *    Curved Segment.
 *  - Esc leaves Edit Points; an outline that grows past its box grows the box.
 *  - Edit Points on a preset shape first turns it into a freeform: Office
 *    writes the preset's current outline as `a:custGeom` (the file then holds
 *    no `a:prstGeom`), so the conversion below evaluates the same catalog
 *    program the canvas draws, with the shape's own adjust values.
 *
 * Coordinates here are the shape's own unrotated, unflipped box in CSS pixels;
 * {@link officeCustomPathFromEdit} returns the new box relative to the old one
 * and the path re-gridded to it.
 */

import { evaluateDrawingmlDefaultPreset } from "./drawingml-default-preset-program";
import { DEFAULT_PRESET_PROGRAMS } from "./presentation-default-preset-programs.generated";
import {
  normalizePresentationCustomPath,
  presentationCustomPathArcEnd,
  PRESENTATION_CUSTOM_PATH_GRID,
  PRESENTATION_CUSTOM_PATH_MAX_COMMANDS,
  PRESENTATION_SCRIBBLE_MAX_POINTS,
  type PresentationCustomPath,
  type PresentationCustomPathCommand,
} from "./presentation-freeform";

export interface OfficePathPoint {
  readonly x: number;
  readonly y: number;
}

export type OfficeEditSegment =
  | { readonly kind: "line"; readonly to: OfficePathPoint }
  | {
      readonly kind: "cubic";
      readonly c1: OfficePathPoint;
      readonly c2: OfficePathPoint;
      readonly to: OfficePathPoint;
    };

export interface OfficeEditFigure {
  readonly closed: boolean;
  readonly start: OfficePathPoint;
  readonly segments: readonly OfficeEditSegment[];
}

export interface OfficeEditPath {
  readonly figures: readonly OfficeEditFigure[];
}

export type OfficeEditPointType = "smooth" | "straight" | "corner";

/**
 * The shared vertex limit of one authored outline: the deck's Freeform /
 * Scribble limit (512 points), applied by every host, so a path drawn in one
 * editor can be edited in any other.
 */
export const OFFICE_FREEFORM_MAX_POINTS = PRESENTATION_SCRIBBLE_MAX_POINTS;

const GRID = PRESENTATION_CUSTOM_PATH_GRID;

function lerp(a: OfficePathPoint, b: OfficePathPoint, t: number): OfficePathPoint {
  return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
}

/** The stored outline in box pixels, as editable figures of lines and cubics. */
export function officeEditPathFromCustomPath(
  value: unknown,
  width: number,
  height: number,
): OfficeEditPath | null {
  const path = normalizePresentationCustomPath(value);
  if (!path || !(width > 0) || !(height > 0)) return null;
  const sx = width / GRID;
  const sy = height / GRID;
  const px = (x: number, y: number): OfficePathPoint => ({ x: x * sx, y: y * sy });
  const figures: OfficeEditFigure[] = [];
  let start: OfficePathPoint | null = null;
  let segments: OfficeEditSegment[] = [];
  let pen = { x: 0, y: 0 }; // grid units
  const flush = (closed: boolean) => {
    if (start && segments.length > 0) figures.push({ closed, start, segments });
    start = null;
    segments = [];
  };
  for (const command of path.commands) {
    switch (command.type) {
      case "moveTo":
        flush(false);
        start = px(command.x, command.y);
        pen = { x: command.x, y: command.y };
        break;
      case "lineTo":
        segments.push({ kind: "line", to: px(command.x, command.y) });
        pen = { x: command.x, y: command.y };
        break;
      case "quadBezierTo": {
        const p0 = px(pen.x, pen.y);
        const q = px(command.x1, command.y1);
        const to = px(command.x, command.y);
        segments.push({ kind: "cubic", c1: lerp(p0, q, 2 / 3), c2: lerp(to, q, 2 / 3), to });
        pen = { x: command.x, y: command.y };
        break;
      }
      case "cubicBezierTo":
        segments.push({
          kind: "cubic",
          c1: px(command.x1, command.y1),
          c2: px(command.x2, command.y2),
          to: px(command.x, command.y),
        });
        pen = { x: command.x, y: command.y };
        break;
      case "arcTo": {
        const end = presentationCustomPathArcEnd(pen, command);
        const cubics = arcToCubics(pen, command);
        for (const cubic of cubics) {
          segments.push({ kind: "cubic", c1: px(cubic.c1.x, cubic.c1.y), c2: px(cubic.c2.x, cubic.c2.y), to: px(cubic.to.x, cubic.to.y) });
        }
        pen = end;
        break;
      }
      case "close": {
        const figureStart = start;
        flush(true);
        if (figureStart) pen = { x: figureStart.x / sx, y: figureStart.y / sy };
        break;
      }
    }
  }
  flush(path.closed);
  return figures.length > 0 ? { figures } : null;
}

/** A DrawingML `a:arcTo` from `from` (grid units), as cubic Beziers. */
function arcToCubics(
  from: OfficePathPoint,
  arc: { readonly wR: number; readonly hR: number; readonly stAng: number; readonly swAng: number },
): { c1: OfficePathPoint; c2: OfficePathPoint; to: OfficePathPoint }[] {
  const turn = 21_600_000;
  const start = (arc.stAng / turn) * 2 * Math.PI;
  const sweep = (arc.swAng / turn) * 2 * Math.PI;
  const cx = from.x - arc.wR * Math.cos(start);
  const cy = from.y - arc.hR * Math.sin(start);
  return ellipseCubics(cx, cy, arc.wR, arc.hR, start, sweep);
}

function ellipseCubics(
  cx: number, cy: number, rx: number, ry: number, start: number, sweep: number,
): { c1: OfficePathPoint; c2: OfficePathPoint; to: OfficePathPoint }[] {
  const count = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  const step = sweep / count;
  const k = (4 / 3) * Math.tan(step / 4);
  const out: { c1: OfficePathPoint; c2: OfficePathPoint; to: OfficePathPoint }[] = [];
  for (let index = 0; index < count; index += 1) {
    const a0 = start + step * index;
    const a1 = a0 + step;
    const p0 = { x: cx + rx * Math.cos(a0), y: cy + ry * Math.sin(a0) };
    const p1 = { x: cx + rx * Math.cos(a1), y: cy + ry * Math.sin(a1) };
    out.push({
      c1: { x: p0.x - k * rx * Math.sin(a0), y: p0.y + k * ry * Math.cos(a0) },
      c2: { x: p1.x + k * rx * Math.sin(a1), y: p1.y - k * ry * Math.cos(a1) },
      to: p1,
    });
  }
  return out;
}

/** SVG endpoint arc (the evaluator's `A rx ry 0 large sweep x y`) as cubics. */
function svgArcToCubics(
  from: OfficePathPoint, rx: number, ry: number, large: boolean, sweepFlag: boolean, to: OfficePathPoint,
): { c1: OfficePathPoint; c2: OfficePathPoint; to: OfficePathPoint }[] {
  if (rx === 0 || ry === 0 || (from.x === to.x && from.y === to.y)) {
    return [{ c1: from, c2: to, to }];
  }
  const dx = (from.x - to.x) / 2;
  const dy = (from.y - to.y) / 2;
  let rxs = Math.abs(rx);
  let rys = Math.abs(ry);
  const lambda = (dx * dx) / (rxs * rxs) + (dy * dy) / (rys * rys);
  if (lambda > 1) {
    rxs *= Math.sqrt(lambda);
    rys *= Math.sqrt(lambda);
  }
  const numerator = rxs * rxs * rys * rys - rxs * rxs * dy * dy - rys * rys * dx * dx;
  const denominator = rxs * rxs * dy * dy + rys * rys * dx * dx;
  let factor = denominator === 0 ? 0 : Math.sqrt(Math.max(0, numerator / denominator));
  if (large === sweepFlag) factor = -factor;
  const cxp = (factor * rxs * dy) / rys;
  const cyp = (-factor * rys * dx) / rxs;
  const cx = cxp + (from.x + to.x) / 2;
  const cy = cyp + (from.y + to.y) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number) => Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
  const start = angle(1, 0, (dx - cxp) / rxs, (dy - cyp) / rys);
  let sweep = angle((dx - cxp) / rxs, (dy - cyp) / rys, (-dx - cxp) / rxs, (-dy - cyp) / rys);
  if (!sweepFlag && sweep > 0) sweep -= 2 * Math.PI;
  if (sweepFlag && sweep < 0) sweep += 2 * Math.PI;
  const cubics = ellipseCubics(cx, cy, rxs, rys, start, sweep);
  if (cubics.length) cubics[cubics.length - 1] = { ...cubics[cubics.length - 1], to };
  return cubics;
}

/**
 * A preset's current outline (with the shape's own adjust values) as an
 * editable path -- what Office does to a preset when Edit Points starts.
 * Only a preset drawn as ONE path converts; a preset whose look is several
 * stacked paths (a cube's shaded faces, a can's lid) has no single outline to
 * edit, and the caller says so instead of flattening it.
 */
export function officeEditPathFromPreset(
  drawingPreset: string,
  width: number,
  height: number,
  adjustments?: Readonly<Record<string, number>> | null,
): OfficeEditPath | null {
  if (!(width > 0) || !(height > 0)) return null;
  const entry = Object.prototype.hasOwnProperty.call(DEFAULT_PRESET_PROGRAMS, drawingPreset)
    ? (DEFAULT_PRESET_PROGRAMS[drawingPreset] as { program?: unknown })
    : null;
  if (!entry?.program) return null;
  const geometry = evaluateDrawingmlDefaultPreset(entry.program, width * 9525, height * 9525, adjustments ?? null);
  if (!geometry || geometry.paths.length !== 1) return null;
  const [only] = geometry.paths;
  const [vw, vh] = geometry.viewBox ?? [width, height];
  const sx = width / vw;
  const sy = height / vh;
  const tokens = only.d.match(/[MLQCAZ]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?/gi) ?? [];
  const figures: OfficeEditFigure[] = [];
  let start: OfficePathPoint | null = null;
  let segments: OfficeEditSegment[] = [];
  let pen: OfficePathPoint = { x: 0, y: 0 };
  let index = 0;
  const number = () => Number(tokens[index++]);
  const point = () => {
    const x = number() * sx;
    const y = number() * sy;
    return { x, y };
  };
  const flush = (closed: boolean) => {
    if (start && segments.length > 0) figures.push({ closed, start, segments });
    start = null;
    segments = [];
  };
  while (index < tokens.length) {
    const op = tokens[index++].toUpperCase();
    if (op === "M") {
      flush(false);
      start = point();
      pen = start;
    } else if (op === "L") {
      const to = point();
      segments.push({ kind: "line", to });
      pen = to;
    } else if (op === "Q") {
      const q = point();
      const to = point();
      segments.push({ kind: "cubic", c1: lerp(pen, q, 2 / 3), c2: lerp(to, q, 2 / 3), to });
      pen = to;
    } else if (op === "C") {
      const c1 = point();
      const c2 = point();
      const to = point();
      segments.push({ kind: "cubic", c1, c2, to });
      pen = to;
    } else if (op === "A") {
      const rx = number() * sx;
      const ry = number() * sy;
      number(); // x-axis rotation is always 0 in DrawingML
      const large = number() === 1;
      const sweep = number() === 1;
      const to = point();
      for (const cubic of svgArcToCubics(pen, rx, ry, large, sweep, to)) {
        segments.push({ kind: "cubic", ...cubic });
      }
      pen = to;
    } else if (op === "Z") {
      const figureStart = start;
      flush(true);
      if (figureStart) pen = figureStart;
    } else {
      return null;
    }
  }
  flush(false);
  if (only.fill === "none") {
    // An open preset (an arc, a bracket): its figures stay open.
    return figures.length ? { figures: figures.map((figure) => ({ ...figure, closed: false })) } : null;
  }
  return figures.length ? { figures } : null;
}

/** Every vertex of the path, in order, with its figure and index. */
export function officeEditPathVertices(path: OfficeEditPath): {
  figure: number;
  index: number;
  point: OfficePathPoint;
}[] {
  const out: { figure: number; index: number; point: OfficePathPoint }[] = [];
  path.figures.forEach((figure, figureIndex) => {
    const count = vertexCount(figure);
    for (let index = 0; index < count; index += 1) {
      out.push({ figure: figureIndex, index, point: vertexAt(figure, index) });
    }
  });
  return out;
}

/** A closed figure's last segment usually returns to its start; that end is the start vertex, not a second one. */
function closesOnStart(figure: OfficeEditFigure): boolean {
  const last = figure.segments[figure.segments.length - 1];
  return figure.closed && !!last && Math.hypot(last.to.x - figure.start.x, last.to.y - figure.start.y) < 1e-6;
}

function vertexCount(figure: OfficeEditFigure): number {
  return figure.segments.length + (closesOnStart(figure) ? 0 : 1);
}

function vertexAt(figure: OfficeEditFigure, index: number): OfficePathPoint {
  return index === 0 ? figure.start : figure.segments[index - 1].to;
}

/** The Bezier handles of one vertex: the incoming segment's `c2`, the outgoing one's `c1`. */
export function officeEditVertexHandles(
  path: OfficeEditPath,
  figureIndex: number,
  index: number,
): { incoming?: OfficePathPoint; outgoing?: OfficePathPoint } {
  const figure = path.figures[figureIndex];
  if (!figure) return {};
  const incoming = incomingSegmentIndex(figure, index);
  const outgoing = outgoingSegmentIndex(figure, index);
  const inSegment = incoming === null ? null : figure.segments[incoming];
  const outSegment = outgoing === null ? null : figure.segments[outgoing];
  return {
    ...(inSegment?.kind === "cubic" ? { incoming: inSegment.c2 } : {}),
    ...(outSegment?.kind === "cubic" ? { outgoing: outSegment.c1 } : {}),
  };
}

function incomingSegmentIndex(figure: OfficeEditFigure, index: number): number | null {
  if (index > 0) return index - 1;
  return figure.closed ? figure.segments.length - 1 : null;
}

function outgoingSegmentIndex(figure: OfficeEditFigure, index: number): number | null {
  if (index < figure.segments.length) {
    // In a figure whose last segment lands on the start, vertex N-1's outgoing
    // is the last segment; an open figure's last vertex has none.
    if (!figure.closed && index === figure.segments.length) return null;
    return index;
  }
  return null;
}

function withVertex(figure: OfficeEditFigure, index: number, point: OfficePathPoint): OfficeEditFigure {
  const segments = [...figure.segments];
  let start = figure.start;
  if (index === 0) {
    start = point;
    if (closesOnStart(figure)) {
      const last = segments[segments.length - 1];
      segments[segments.length - 1] = { ...last, to: point };
    }
  } else {
    segments[index - 1] = { ...segments[index - 1], to: point };
  }
  return { ...figure, start, segments };
}

function replaceFigure(path: OfficeEditPath, figureIndex: number, figure: OfficeEditFigure): OfficeEditPath {
  return { figures: path.figures.map((candidate, index) => (index === figureIndex ? figure : candidate)) };
}

/** Drag a vertex: it moves, and its two handles travel with it. */
export function officeEditMoveVertex(
  path: OfficeEditPath,
  figureIndex: number,
  index: number,
  to: OfficePathPoint,
): OfficeEditPath {
  const figure = path.figures[figureIndex];
  if (!figure) return path;
  const from = vertexAt(figure, index);
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  let next = withVertex(figure, index, to);
  const segments = [...next.segments];
  const incoming = incomingSegmentIndex(figure, index);
  const outgoing = outgoingSegmentIndex(figure, index);
  if (incoming !== null) {
    const segment = segments[incoming];
    if (segment.kind === "cubic") segments[incoming] = { ...segment, c2: { x: segment.c2.x + dx, y: segment.c2.y + dy } };
  }
  if (outgoing !== null) {
    const segment = segments[outgoing];
    if (segment.kind === "cubic") segments[outgoing] = { ...segment, c1: { x: segment.c1.x + dx, y: segment.c1.y + dy } };
  }
  next = { ...next, segments };
  return replaceFigure(path, figureIndex, next);
}

/**
 * Drag one Bezier handle of a vertex. `type` is the modifier Office reads while
 * dragging: smooth (Shift) mirrors the other handle at equal length, straight
 * (Ctrl) keeps it on the same line at its own length, corner (Alt, and the
 * default) leaves it alone.
 */
export function officeEditMoveHandle(
  path: OfficeEditPath,
  figureIndex: number,
  index: number,
  side: "incoming" | "outgoing",
  to: OfficePathPoint,
  type: OfficeEditPointType = "corner",
): OfficeEditPath {
  const figure = path.figures[figureIndex];
  if (!figure) return path;
  const vertex = vertexAt(figure, index);
  const incoming = incomingSegmentIndex(figure, index);
  const outgoing = outgoingSegmentIndex(figure, index);
  const segments = [...figure.segments];
  const own = side === "incoming" ? incoming : outgoing;
  const other = side === "incoming" ? outgoing : incoming;
  if (own === null || segments[own].kind !== "cubic") return path;
  const ownSegment = segments[own] as Extract<OfficeEditSegment, { kind: "cubic" }>;
  segments[own] = side === "incoming" ? { ...ownSegment, c2: to } : { ...ownSegment, c1: to };
  if (type !== "corner" && other !== null && segments[other].kind === "cubic") {
    const otherSegment = segments[other] as Extract<OfficeEditSegment, { kind: "cubic" }>;
    const otherHandle = side === "incoming" ? otherSegment.c1 : otherSegment.c2;
    const dx = to.x - vertex.x;
    const dy = to.y - vertex.y;
    const length = Math.hypot(dx, dy);
    const otherLength = type === "smooth" ? length : Math.hypot(otherHandle.x - vertex.x, otherHandle.y - vertex.y);
    const mirrored = length > 0
      ? { x: vertex.x - (dx / length) * otherLength, y: vertex.y - (dy / length) * otherLength }
      : otherHandle;
    segments[other] = side === "incoming" ? { ...otherSegment, c1: mirrored } : { ...otherSegment, c2: mirrored };
  }
  return replaceFigure(path, figureIndex, { ...figure, segments });
}

/** Make a vertex smooth / straight / corner (the vertex menu's point types). */
export function officeEditSetPointType(
  path: OfficeEditPath,
  figureIndex: number,
  index: number,
  type: OfficeEditPointType,
): OfficeEditPath {
  if (type === "corner") return path;
  const stated = path.figures[figureIndex];
  if (!stated) return path;
  // A closed figure whose close is implied (no segment back to its start)
  // has its closing edge drawn as one: its last vertex had no outgoing
  // segment and its first an incoming one that was not its own, so Smooth
  
  const figure =
    stated.closed && !closesOnStart(stated) && stated.segments.length > 0
      ? { ...stated, segments: [...stated.segments, { kind: "line" as const, to: stated.start }] }
      : stated;
  const vertex = vertexAt(figure, index);
  const incoming = incomingSegmentIndex(figure, index);
  const outgoing = outgoingSegmentIndex(figure, index);
  if (incoming === null || outgoing === null) return path;
  let next = figure;
  // A point type is about handles, so both neighbouring segments become curves.
  for (const segmentIndex of [incoming, outgoing]) next = curvedSegment(next, segmentIndex);
  const segments = [...next.segments];
  const inSegment = segments[incoming] as Extract<OfficeEditSegment, { kind: "cubic" }>;
  const outSegment = segments[outgoing] as Extract<OfficeEditSegment, { kind: "cubic" }>;
  const previous = incoming === null ? vertex : segmentStart(next, incoming);
  const following = outSegment.to;
  // The tangent through the vertex runs from its previous neighbour to its next.
  let tx = following.x - previous.x;
  let ty = following.y - previous.y;
  const tangent = Math.hypot(tx, ty) || 1;
  tx /= tangent;
  ty /= tangent;
  const inLength = Math.hypot(inSegment.c2.x - vertex.x, inSegment.c2.y - vertex.y) || Math.hypot(vertex.x - previous.x, vertex.y - previous.y) / 3;
  const outLength = Math.hypot(outSegment.c1.x - vertex.x, outSegment.c1.y - vertex.y) || Math.hypot(following.x - vertex.x, following.y - vertex.y) / 3;
  const shared = (inLength + outLength) / 2;
  const inL = type === "smooth" ? shared : inLength;
  const outL = type === "smooth" ? shared : outLength;
  segments[incoming] = { ...inSegment, c2: { x: vertex.x - tx * inL, y: vertex.y - ty * inL } };
  segments[outgoing] = { ...outSegment, c1: { x: vertex.x + tx * outL, y: vertex.y + ty * outL } };
  return replaceFigure(path, figureIndex, { ...next, segments });
}

function segmentStart(figure: OfficeEditFigure, segmentIndex: number): OfficePathPoint {
  return segmentIndex === 0 ? figure.start : figure.segments[segmentIndex - 1].to;
}

function curvedSegment(figure: OfficeEditFigure, segmentIndex: number): OfficeEditFigure {
  const segment = figure.segments[segmentIndex];
  if (!segment || segment.kind === "cubic") return figure;
  const from = segmentStart(figure, segmentIndex);
  const segments = [...figure.segments];
  segments[segmentIndex] = { kind: "cubic", c1: lerp(from, segment.to, 1 / 3), c2: lerp(from, segment.to, 2 / 3), to: segment.to };
  return { ...figure, segments };
}

/** Straight Segment / Curved Segment on the segment menu. */
export function officeEditSetSegmentKind(
  path: OfficeEditPath,
  figureIndex: number,
  segmentIndex: number,
  kind: "line" | "cubic",
): OfficeEditPath {
  const figure = path.figures[figureIndex];
  const segment = figure?.segments[segmentIndex];
  if (!figure || !segment || segment.kind === kind) return path;
  if (kind === "cubic") return replaceFigure(path, figureIndex, curvedSegment(figure, segmentIndex));
  const segments = [...figure.segments];
  segments[segmentIndex] = { kind: "line", to: segment.to };
  return replaceFigure(path, figureIndex, { ...figure, segments });
}

function segmentPoint(from: OfficePathPoint, segment: OfficeEditSegment, t: number): OfficePathPoint {
  if (segment.kind === "line") return lerp(from, segment.to, t);
  const u = 1 - t;
  return {
    x: u * u * u * from.x + 3 * u * u * t * segment.c1.x + 3 * u * t * t * segment.c2.x + t * t * t * segment.to.x,
    y: u * u * u * from.y + 3 * u * u * t * segment.c1.y + 3 * u * t * t * segment.c2.y + t * t * t * segment.to.y,
  };
}

/** The segment nearest `point` within `tolerance` pixels, with where along it. */
export function officeEditHitSegment(
  path: OfficeEditPath,
  point: OfficePathPoint,
  tolerance: number,
): { figure: number; segment: number; t: number } | null {
  let best: { figure: number; segment: number; t: number; distance: number } | null = null;
  path.figures.forEach((figure, figureIndex) => {
    figure.segments.forEach((segment, segmentIndex) => {
      const from = segmentStart(figure, segmentIndex);
      for (let step = 0; step <= 48; step += 1) {
        const t = step / 48;
        const at = segmentPoint(from, segment, t);
        const distance = Math.hypot(at.x - point.x, at.y - point.y);
        if (distance <= tolerance && (!best || distance < best.distance)) {
          best = { figure: figureIndex, segment: segmentIndex, t, distance };
        }
      }
    });
  });
  if (!best) return null;
  const { figure, segment, t } = best;
  return { figure, segment, t };
}

/** Add Point: split one segment at `t` (de Casteljau for a curve). */
export function officeEditAddVertex(
  path: OfficeEditPath,
  figureIndex: number,
  segmentIndex: number,
  t: number,
): OfficeEditPath {
  const figure = path.figures[figureIndex];
  const segment = figure?.segments[segmentIndex];
  if (!figure || !segment || t <= 0 || t >= 1) return path;
  if (officeEditPathVertices(path).length >= OFFICE_FREEFORM_MAX_POINTS) return path;
  const from = segmentStart(figure, segmentIndex);
  let pieces: OfficeEditSegment[];
  if (segment.kind === "line") {
    const mid = lerp(from, segment.to, t);
    pieces = [{ kind: "line", to: mid }, { kind: "line", to: segment.to }];
  } else {
    const p01 = lerp(from, segment.c1, t);
    const p12 = lerp(segment.c1, segment.c2, t);
    const p23 = lerp(segment.c2, segment.to, t);
    const p012 = lerp(p01, p12, t);
    const p123 = lerp(p12, p23, t);
    const mid = lerp(p012, p123, t);
    pieces = [
      { kind: "cubic", c1: p01, c2: p012, to: mid },
      { kind: "cubic", c1: p123, c2: p23, to: segment.to },
    ];
  }
  const segments = [...figure.segments];
  segments.splice(segmentIndex, 1, ...pieces);
  return replaceFigure(path, figureIndex, { ...figure, segments });
}

/**
 * Delete Point. A figure keeps at least two vertices when open and three when
 * closed, as Office refuses to delete below that.
 */
export function officeEditDeleteVertex(
  path: OfficeEditPath,
  figureIndex: number,
  index: number,
): OfficeEditPath {
  const figure = path.figures[figureIndex];
  if (!figure) return path;
  const count = vertexCount(figure);
  if (count <= (figure.closed ? 3 : 2)) return path;
  const segments = [...figure.segments];
  if (index === 0) {
    const first = segments.shift();
    if (!first) return path;
    let start = first.to;
    if (closesOnStart(figure)) {
      // The closing segment now returns to the new start.
      const last = segments[segments.length - 1];
      segments[segments.length - 1] = mergeSegments(last, first);
      start = first.to;
    }
    return replaceFigure(path, figureIndex, { ...figure, start, segments });
  }
  const incoming = segments[index - 1];
  const outgoing = segments[index];
  if (!outgoing) {
    segments.pop();
    return replaceFigure(path, figureIndex, { ...figure, segments });
  }
  segments.splice(index - 1, 2, mergeSegments(incoming, outgoing));
  return replaceFigure(path, figureIndex, { ...figure, segments });
}

function mergeSegments(incoming: OfficeEditSegment, outgoing: OfficeEditSegment): OfficeEditSegment {
  if (incoming.kind === "line" && outgoing.kind === "line") return { kind: "line", to: outgoing.to };
  return {
    kind: "cubic",
    c1: incoming.kind === "cubic" ? incoming.c1 : incoming.to,
    c2: outgoing.kind === "cubic" ? outgoing.c2 : outgoing.to,
    to: outgoing.to,
  };
}

/** Open Path (at a vertex of a closed figure) / Close Path (an open one). */
export function officeEditToggleClosed(
  path: OfficeEditPath,
  figureIndex: number,
  index = 0,
): OfficeEditPath {
  const figure = path.figures[figureIndex];
  if (!figure) return path;
  if (!figure.closed) {
    return replaceFigure(path, figureIndex, { ...figure, closed: true });
  }
  // Open at `index`: the figure now starts and ends there.
  const vertices = vertexCount(figure);
  const ring: OfficeEditSegment[] = [...figure.segments];
  if (!closesOnStart(figure)) ring.push({ kind: "line", to: figure.start });
  const rotated = [...ring.slice(index), ...ring.slice(0, index)];
  const start = vertexAt(figure, index % vertices);
  return replaceFigure(path, figureIndex, { closed: false, start, segments: rotated });
}

function cubicExtent(from: OfficePathPoint, segment: OfficeEditSegment): OfficePathPoint[] {
  if (segment.kind === "line") return [segment.to];
  const points: OfficePathPoint[] = [];
  for (let step = 1; step <= 24; step += 1) points.push(segmentPoint(from, segment, step / 24));
  return points;
}

export interface OfficeCustomPathEditResult {
  /** The new box, relative to the old box's top-left, in box pixels. */
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly customPath: PresentationCustomPath;
}

/**
 * The edited path re-gridded onto its own extent. The box follows the drawn
 * outline (Office grows a freeform's box with it), so vertices land on the
 * 0..100000 grid and Bezier handles may overhang within the envelope.
 */
export function officeCustomPathFromEdit(path: OfficeEditPath): OfficeCustomPathEditResult | null {
  const extent: OfficePathPoint[] = [];
  for (const figure of path.figures) {
    extent.push(figure.start);
    figure.segments.forEach((segment, index) => extent.push(...cubicExtent(segmentStart(figure, index), segment)));
  }
  if (extent.length < 2 || extent.some(({ x, y }) => !Number.isFinite(x) || !Number.isFinite(y))) return null;
  let minX = Math.min(...extent.map(({ x }) => x));
  let maxX = Math.max(...extent.map(({ x }) => x));
  let minY = Math.min(...extent.map(({ y }) => y));
  let maxY = Math.max(...extent.map(({ y }) => y));
  // A straight open stroke still needs a box a selection can grab.
  if (maxX - minX < 1) { minX -= 2.5; maxX += 2.5; }
  if (maxY - minY < 1) { minY -= 2.5; maxY += 2.5; }
  const width = maxX - minX;
  const height = maxY - minY;
  const anchor = (value: number, min: number, size: number) =>
    Math.max(0, Math.min(GRID, Math.round(((value - min) / size) * GRID)));
  const control = (value: number, min: number, size: number) =>
    Math.max(-4 * GRID, Math.min(4 * GRID, Math.round(((value - min) / size) * GRID)));
  const commands: PresentationCustomPathCommand[] = [];
  path.figures.forEach((figure, figureIndex) => {
    commands.push({ type: "moveTo", x: anchor(figure.start.x, minX, width), y: anchor(figure.start.y, minY, height) });
    for (const segment of figure.segments) {
      const x = anchor(segment.to.x, minX, width);
      const y = anchor(segment.to.y, minY, height);
      if (segment.kind === "line") commands.push({ type: "lineTo", x, y });
      else commands.push({
        type: "cubicBezierTo",
        x1: control(segment.c1.x, minX, width), y1: control(segment.c1.y, minY, height),
        x2: control(segment.c2.x, minX, width), y2: control(segment.c2.y, minY, height),
        x, y,
      });
    }
    const last = figureIndex === path.figures.length - 1;
    if (figure.closed && !last) commands.push({ type: "close" });
  });
  if (commands.length > PRESENTATION_CUSTOM_PATH_MAX_COMMANDS) return null;
  const closed = path.figures[path.figures.length - 1]?.closed ?? false;
  const customPath = normalizePresentationCustomPath({ version: 1, closed, commands });
  if (!customPath) return null;
  return { left: minX, top: minY, width, height, customPath };
}

/**
 * Where the re-boxed shape lands on its page: `box` is the shape's current
 * position and size in page units, `result` is in the same units relative to
 * it. Rotation and flips are about the box centre, so the new centre is the
 * old centre plus the rotated, mirrored offset -- the outline stays exactly
 * where the user dragged it.
 */
export function officeCustomPathEditBox(
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  result: { readonly left: number; readonly top: number; readonly width: number; readonly height: number },
  rotationDeg = 0,
  flipH = false,
  flipV = false,
): { x: number; y: number; width: number; height: number } {
  let ox = result.left + result.width / 2 - box.width / 2;
  let oy = result.top + result.height / 2 - box.height / 2;
  if (flipH) ox = -ox;
  if (flipV) oy = -oy;
  const radians = ((rotationDeg || 0) * Math.PI) / 180;
  const rx = ox * Math.cos(radians) - oy * Math.sin(radians);
  const ry = ox * Math.sin(radians) + oy * Math.cos(radians);
  const cx = box.x + box.width / 2 + rx;
  const cy = box.y + box.height / 2 + ry;
  return { x: cx - result.width / 2, y: cy - result.height / 2, width: result.width, height: result.height };
}

/** The edit path as an SVG `d` in box pixels, for the live outline. */
export function officeEditPathToSvg(path: OfficeEditPath): string {
  return path.figures
    .map((figure) => {
      const parts = [`M${figure.start.x} ${figure.start.y}`];
      for (const segment of figure.segments) {
        parts.push(segment.kind === "line"
          ? `L${segment.to.x} ${segment.to.y}`
          : `C${segment.c1.x} ${segment.c1.y} ${segment.c2.x} ${segment.c2.y} ${segment.to.x} ${segment.to.y}`);
      }
      if (figure.closed) parts.push("Z");
      return parts.join(" ");
    })
    .join(" ");
}
