/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Merge Shapes on a PDF page's own vectors (PowerPoint's Shape Format >
 * Merge Shapes; ONLYOFFICE's 合并形状): the selected objects' outlines, read
 * from the file, handed to the merge every editor uses (the engine's
 * `shapes.merge`, lib/office-shape-merge-client.ts), and its results as the
 * page paths that replace them (PdfNativeVectorMergeCommitRequest).
 *
 * Lengths stay page points throughout: the merge's canvas is the page's user
 * space, whose y runs up, which a boolean does not mind.
 */
import type { PdfNativeVectorSubpathCommand } from "./engine/native/vector-content-transaction";
import type { PdfNativeVectorWriteSelection } from "./native-vector-selection-edit";
import { PRESENTATION_CUSTOM_PATH_GRID } from "@mrnote/core/web/lib/presentation-freeform.ts";
import type { MergeableShape, MergedShapeResult } from "@mrnote/core/web/lib/presentation-shapes.ts";

type Point = Readonly<{ x: number; y: number }>;

/** The merge takes a polygon of at most this many points (the custom
 * polygon's bound, MAX_CUSTOM_POLYGON_POINTS). */
export const PDF_VECTOR_MERGE_MAX_POINTS = 256;
/** Segments a cubic curve is drawn as. */
const CURVE_STEPS = 12;

/** Segments an ellipse is drawn as. */
const ELLIPSE_STEPS = 64;

/**
 * The path (page space) a selected object draws, as the reader proves it: a
 * rectangle's four corners, an ellipse's outline from its centre and axes, a
 * subpath's own commands. A line has no area (null).
 */
export function pdfVectorSelectionOutline(
  selection: PdfNativeVectorWriteSelection,
): readonly PdfNativeVectorSubpathCommand[] {
  if (selection.proofFamily === "subpath") return selection.commands;
  const geometry = selection.geometry;
  const ring: Point[] = geometry.kind === "rectangle"
    ? [...geometry.corners]
    : geometry.kind === "ellipse"
      ? Array.from({ length: ELLIPSE_STEPS }, (_, step) => {
        const angle = (step / ELLIPSE_STEPS) * Math.PI * 2;
        return {
          x: geometry.center.x + Math.cos(angle) * geometry.axisX.x + Math.sin(angle) * geometry.axisY.x,
          y: geometry.center.y + Math.cos(angle) * geometry.axisX.y + Math.sin(angle) * geometry.axisY.y,
        };
      })
      : [];
  if (ring.length < 3) return [];
  return [
    ...ring.map((point, index): PdfNativeVectorSubpathCommand => ({ operator: index === 0 ? "m" : "l", point })),
    { operator: "h" },
  ];
}

/** A path's subpaths as point rings (curves drawn as short lines), or null
 * where it is not one closed area: a line, or a path drawn open. */
export function pdfVectorOutlineRings(
  commands: readonly PdfNativeVectorSubpathCommand[],
): Point[][] | null {
  const rings: Point[][] = [];
  let ring: Point[] = [];
  let current: Point | null = null;
  const close = () => {
    if (ring.length >= 3) rings.push(ring);
    ring = [];
  };
  for (const command of commands) {
    if (command.operator === "m") {
      close();
      current = command.point;
      ring.push(current);
    } else if (command.operator === "l") {
      current = command.point;
      ring.push(current);
    } else if (command.operator === "h") {
      close();
      current = null;
    } else if (command.operator === "c" || command.operator === "v" || command.operator === "y") {
      if (!current) return null;
      const start: Point = current;
      // `v` takes its first control from the current point, `y` its second
      // from the end (PDF 32000-1 8.5.2.2).
      const c1 = command.operator === "v" ? start : command.control1;
      const c2 = command.operator === "y" ? command.end : command.control2;
      for (let step = 1; step <= CURVE_STEPS; step += 1) {
        const t = step / CURVE_STEPS;
        const u = 1 - t;
        ring.push({
          x: u * u * u * start.x + 3 * u * u * t * c1.x + 3 * u * t * t * c2.x + t * t * t * command.end.x,
          y: u * u * u * start.y + 3 * u * u * t * c1.y + 3 * u * t * t * c2.y + t * t * t * command.end.y,
        });
      }
      current = command.end;
    }
  }
  close();
  return rings.length > 0 ? rings.map(withoutRepeats).filter((points) => points.length >= 3) : null;
}

function withoutRepeats(ring: Point[]): Point[] {
  const out: Point[] = [];
  for (const point of ring) {
    const last = out[out.length - 1];
    if (!last || Math.abs(last.x - point.x) > 1e-6 || Math.abs(last.y - point.y) > 1e-6) out.push(point);
  }
  while (out.length > 1 && Math.abs(out[0].x - out[out.length - 1].x) <= 1e-6 &&
    Math.abs(out[0].y - out[out.length - 1].y) <= 1e-6) out.pop();
  return out;
}

/**
 * One object's outline as the merge takes it: a single closed area as a
 * custom polygon over its own box (0-100), its box in page points. Null for
 * what cannot take part -- a line, an open path, a path of several areas (the
 * merge takes one outline a shape, as in the other editors) or one of more
 * points than a polygon holds.
 */
export function pdfVectorMergeableShape(
  commands: readonly PdfNativeVectorSubpathCommand[],
): MergeableShape | null {
  const rings = pdfVectorOutlineRings(commands);
  if (!rings || rings.length !== 1 || rings[0].length > PDF_VECTOR_MERGE_MAX_POINTS) return null;
  const ring = rings[0];
  const xs = ring.map((point) => point.x);
  const ys = ring.map((point) => point.y);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const width = Math.max(...xs) - left;
  const height = Math.max(...ys) - top;
  if (!(width > 0) || !(height > 0)) return null;
  const percent = (value: number) => Number(value.toFixed(2));
  return {
    shape: "rect",
    points: ring.map((point) => `${percent(((point.x - left) / width) * 100)},${percent(((point.y - top) / height) * 100)}`).join(" "),
    left,
    top,
    width,
    height,
  };
}

/** A merge result's regions as point rings, page points. */
function resultRings(result: MergedShapeResult): Point[][] {
  const at = (x: number, y: number): Point => ({
    x: result.left + x * result.width,
    y: result.top + y * result.height,
  });
  if (result.points) {
    return [result.points.trim().split(/\s+/).map((pair) => {
      const [x, y] = pair.split(",").map(Number);
      return at(x / 100, y / 100);
    })];
  }
  const rings: Point[][] = [];
  let ring: Point[] = [];
  for (const command of result.customPath?.commands ?? []) {
    if (command.type === "moveTo") {
      if (ring.length >= 3) rings.push(ring);
      ring = [at(command.x / PRESENTATION_CUSTOM_PATH_GRID, command.y / PRESENTATION_CUSTOM_PATH_GRID)];
    } else if (command.type === "lineTo") {
      ring.push(at(command.x / PRESENTATION_CUSTOM_PATH_GRID, command.y / PRESENTATION_CUSTOM_PATH_GRID));
    } else if (command.type === "close") {
      if (ring.length >= 3) rings.push(ring);
      ring = [];
    }
  }
  if (ring.length >= 3) rings.push(ring);
  return rings.map(withoutRepeats).filter((points) => points.length >= 3);
}

function signedArea(ring: readonly Point[]): number {
  let area = 0;
  for (let index = 0; index < ring.length; index += 1) {
    const a = ring[index];
    const b = ring[(index + 1) % ring.length];
    area += a.x * b.y - b.x * a.y;
  }
  return area / 2;
}

function inside(point: Point, ring: readonly Point[]): boolean {
  let hit = false;
  for (let index = 0, previous = ring.length - 1; index < ring.length; previous = index, index += 1) {
    const a = ring[index];
    const b = ring[previous];
    if ((a.y > point.y) !== (b.y > point.y) &&
      point.x < ((b.x - a.x) * (point.y - a.y)) / (b.y - a.y) + a.x) hit = !hit;
  }
  return hit;
}

/**
 * One merge result as the path that is written: a subpath a region, each
 * hole turned the other way from the outline it lies in (the number of rings
 * around it odd), so the page fills the result the same under the nonzero and
 * the even-odd rule -- whichever paint operator the first object carries.
 */
export function pdfVectorMergedPath(result: MergedShapeResult): PdfNativeVectorSubpathCommand[] {
  const rings = resultRings(result);
  const commands: PdfNativeVectorSubpathCommand[] = [];
  for (const [index, ring] of rings.entries()) {
    const depth = rings.filter((other, at) => at !== index && inside(ring[0], other)).length;
    const counterClockwise = signedArea(ring) > 0;
    // Outlines counter-clockwise, holes clockwise.
    const ordered = counterClockwise === (depth % 2 === 0) ? ring : [...ring].reverse();
    ordered.forEach((point, at) => {
      commands.push(Object.freeze({ operator: at === 0 ? "m" : "l", point: Object.freeze({ x: point.x, y: point.y }) }));
    });
    commands.push(Object.freeze({ operator: "h" }));
  }
  return commands;
}
