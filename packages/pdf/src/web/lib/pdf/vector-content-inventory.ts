/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { multiplyMatrix, type ImageOperatorIds } from "./embedded-images";

export interface PdfVectorContentOperatorIds extends
  Pick<ImageOperatorIds, "save" | "restore" | "transform"> {
  readonly constructPath: number | undefined;
  readonly strokeOps: ReadonlySet<number>;
  readonly fillOps: ReadonlySet<number>;
  /** pdf.js `setLineWidth` and `setGState` (whose `LW` entry sets it too):
   * with them each stroked item carries the width the page paints it at. */
  readonly setLineWidth?: number;
  readonly setGState?: number;
  /** pdf.js annotation block markers. An annotation's appearance is drawn in
   * its own space, which the semantic appearance reconciliation maps to the
   * page from the block's own transform: its vectors keep the user-space
   * reading, without the page viewport. */
  readonly beginAnnotation?: number;
  readonly endAnnotation?: number;
  /** pdf.js Form XObject markers: the renderer saves, applies the form's
   * /Matrix (the first argument) and restores at the end, so the form's
   * paths are drawn under that matrix too. */
  readonly paintFormXObjectBegin?: number;
  readonly paintFormXObjectEnd?: number;
  /** pdf.js clip operators (`W` / `W*`); the `constructPath` that follows
   * defines the clip. With them an item a clip cuts carries `clippedBounds`. */
  readonly clip?: number;
  readonly eoClip?: number;
}

export type PdfVectorPaintMode = "stroke" | "fill" | "fill-stroke";
export type PdfVectorContentKind = "rule" | "rect-like" | "generic-path";

export interface PdfVectorSimpleLineEndpoints {
  /** Directed renderer start point in page fractions, top-left origin. */
  readonly start: Readonly<{ x: number; y: number }>;
  /** Directed renderer end point in page fractions, top-left origin. */
  readonly end: Readonly<{ x: number; y: number }>;
}

export type PdfVectorSimplePathCommand = Readonly<
  | {
      kind: "moveTo" | "lineTo";
      point: Readonly<{ x: number; y: number }>;
    }
  | {
      kind: "curveTo";
      control1: Readonly<{ x: number; y: number }>;
      control2: Readonly<{ x: number; y: number }>;
      end: Readonly<{ x: number; y: number }>;
    }
  | {
      kind: "quadraticCurveTo";
      control: Readonly<{ x: number; y: number }>;
      end: Readonly<{ x: number; y: number }>;
    }
  | { kind: "closePath" }
>;

export interface PdfVectorSimpleSubpath {
  readonly commands: readonly PdfVectorSimplePathCommand[];
  readonly closed: boolean;
}

export interface PdfVectorContentItem {
  readonly kind: PdfVectorContentKind;
  readonly paintMode: PdfVectorPaintMode;
  readonly operatorIndex: number;
  /** Page fractions with a top-left origin, clipped to the visible page. */
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly angle: number;
  readonly rotatedOrSkewed: boolean;
  readonly pathValueCount: number;
  readonly simplePathEvidence: boolean;
  /** Exact renderer payload proof only; never grants writer authority alone. */
  readonly simpleLineEndpoints?: PdfVectorSimpleLineEndpoints;
  /** One exact directed renderer subpath; evidence only, never writer authority. */
  readonly simpleSubpath?: PdfVectorSimpleSubpath;
  /** The whole path's box in page fractions, top-left origin: not clipped to
   * the page (a shape the page cuts reaches past it) and not widened on an
   * axis it has no extent on. What a selected shape's frame stands on. */
  readonly extent?: Readonly<{ x: number; y: number; width: number; height: number }>;
  /** The stroke's width in page points as the page paints it: the `w` in
   * effect times the scale of the transform it is drawn under. Present only
   * for a stroked item whose operator list names the line-width operators. */
  readonly strokeWidth?: number;
  /** How far the painted stroke reaches past the path along the page's x and
   * y axes, in page fractions (of its width and of its height): half the `w`
   * carried through the transform per axis. A rule scaled into a bar (WPS's thickness grip: a `cm` stretching
   * one axis) paints its stroke tall, not wide; one mean width would measure
   * it as neither. Present with `strokeWidth`. */
  readonly strokeReach?: Readonly<{ x: number; y: number }>;
  /** The part of `x`/`y`/`width`/`height` the clips in effect leave
   * visible (their boxes intersected), present only when a clip cuts the
   * item. Page fractions, top-left origin; empty when it is clipped away.
   * Present only when the list names the clip operators. */
  readonly clippedBounds?: Readonly<{ x: number; y: number; width: number; height: number }>;
}

type Matrix = [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
const MAX_VECTOR_CONTENT_ITEMS = 100_000;
const MAX_PATH_VALUE_COUNT = 100_000;
const SIMPLE_PATH_VALUE_COUNT = 64;
const RULE_MAX_THICKNESS_PT = 2.5;
const PDFJS_DRAW_PATH_MOVE_TO = 0;
const PDFJS_DRAW_PATH_LINE_TO = 1;
const PDFJS_DRAW_PATH_CURVE_TO = 2;
const PDFJS_DRAW_PATH_QUADRATIC_CURVE_TO = 3;
const PDFJS_DRAW_PATH_CLOSE = 4;
const MAX_SIMPLE_SUBPATH_COMMANDS = 4096;

function finiteMatrix(value: unknown): Matrix | null {
  const raw = Array.isArray(value) && value.length === 1 ? value[0] : value;
  return Array.isArray(raw) &&
    raw.length >= 6 &&
    raw.slice(0, 6).every((entry) => Number.isFinite(Number(entry)))
    ? (raw.slice(0, 6).map(Number) as Matrix)
    : null;
}

function pathValueCount(value: unknown, budget = MAX_PATH_VALUE_COUNT): number {
  const stack: unknown[] = [value];
  let count = 0;
  while (stack.length > 0) {
    const current = stack.pop();
    if (Array.isArray(current) || ArrayBuffer.isView(current)) {
      const values = Array.from(current as ArrayLike<unknown>);
      count += values.length;
      if (count > budget) return count;
      stack.push(...values.filter((entry) => typeof entry === "object"));
    }
  }
  return count;
}

function mapPoint(matrix: Matrix, x: number, y: number) {
  return Object.freeze({
    x: matrix[0] * x + matrix[2] * y + matrix[4],
    y: matrix[1] * x + matrix[3] * y + matrix[5],
  });
}

function normalizedPathPoint(
  matrix: Matrix,
  page: Readonly<{ width: number; height: number }>,
  x: number,
  y: number,
): Readonly<{ x: number; y: number }> | null {
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  const mapped = mapPoint(matrix, x, y);
  const point = Object.freeze({
    x: mapped.x / page.width,
    y: (page.height - mapped.y) / page.height,
  });
  return Number.isFinite(point.x) && Number.isFinite(point.y) ? point : null;
}

function pointDistance(
  left: Readonly<{ x: number; y: number }>,
  right: Readonly<{ x: number; y: number }>,
): number {
  return Math.hypot(left.x - right.x, left.y - right.y);
}

function normalizedDirectedSimpleSubpath(
  value: unknown,
  matrix: Matrix,
  page: Readonly<{ width: number; height: number }>,
): PdfVectorSimpleSubpath | null {
  if (!Array.isArray(value) || value.length !== 1) return null;
  const compactPath = value[0];
  if (!ArrayBuffer.isView(compactPath) || compactPath instanceof DataView) {
    return null;
  }
  const compactValues = compactPath as unknown as ArrayLike<unknown>;
  if (
    compactValues.length < 6 ||
    compactValues.length > 1 + MAX_SIMPLE_SUBPATH_COMMANDS * 7
  ) {
    return null;
  }
  const values = Array.from(compactValues, Number);
  const commands: PdfVectorSimplePathCommand[] = [];
  let cursor = 0;
  let currentPoint: Readonly<{ x: number; y: number }> | null = null;
  let startPoint: Readonly<{ x: number; y: number }> | null = null;
  // One path of one or more subpaths (a ring, a framed panel, a glyph-like
  // icon): a move opens each subpath after the first, a close may end any of
  // them, and a segment straight after a close goes on from the closed
  // subpath's start (PDF 32000-1 8.5.2.1), exactly as the byte engine reads
  // the same operators.
  let segments = 0;
  let closed = false;
  const readPoint = () => {
    if (cursor + 1 >= values.length) return null;
    const point = normalizedPathPoint(
      matrix,
      page,
      values[cursor],
      values[cursor + 1],
    );
    cursor += 2;
    return point;
  };
  while (cursor < values.length) {
    if (commands.length >= MAX_SIMPLE_SUBPATH_COMMANDS) return null;
    const opcode = values[cursor];
    cursor += 1;
    if (opcode === PDFJS_DRAW_PATH_MOVE_TO) {
      const point = readPoint();
      if (!point) return null;
      // A move that ends the path draws nothing (Word moves back to the
      // start after its close); the byte engine leaves it out, and so here.
      if (cursor >= values.length && commands.length > 0) break;
      // A move straight after a move draws nothing either; the byte engine
      // does not read such a path.
      if (commands.length !== 0 && segments === 0 && !closed) return null;
      currentPoint = point;
      startPoint = point;
      segments = 0;
      closed = false;
      commands.push(Object.freeze({ kind: "moveTo", point }));
      continue;
    }
    if (!currentPoint || !startPoint) return null;
    if (opcode === PDFJS_DRAW_PATH_LINE_TO) {
      // A segment of no length is part of the path as written (the byte
      // engine keeps it too).
      const point = readPoint();
      if (!point) return null;
      commands.push(Object.freeze({ kind: "lineTo", point }));
      currentPoint = point;
      segments += 1;
      closed = false;
      continue;
    }
    if (opcode === PDFJS_DRAW_PATH_CURVE_TO) {
      const control1 = readPoint();
      const control2 = readPoint();
      const end = readPoint();
      if (!control1 || !control2 || !end) return null;
      commands.push(
        Object.freeze({ kind: "curveTo", control1, control2, end }),
      );
      currentPoint = end;
      segments += 1;
      closed = false;
      continue;
    }
    if (opcode === PDFJS_DRAW_PATH_QUADRATIC_CURVE_TO) {
      const control = readPoint();
      const end = readPoint();
      if (
        !control ||
        !end ||
        Math.max(
          pointDistance(currentPoint, control),
          pointDistance(currentPoint, end),
        ) <= 0.000_000_001
      ) {
        return null;
      }
      commands.push(Object.freeze({ kind: "quadraticCurveTo", control, end }));
      currentPoint = end;
      segments += 1;
      closed = false;
      continue;
    }
    if (opcode === PDFJS_DRAW_PATH_CLOSE) {
      // A second close does nothing (an `h` the closing paint repeats);
      // `m h` closes a subpath of one point.
      if (closed) continue;
      commands.push(Object.freeze({ kind: "closePath" }));
      currentPoint = startPoint;
      closed = true;
      continue;
    }
    return null;
  }
  if (commands.length < 2 || commands[0].kind !== "moveTo") return null;
  if (segments === 0 && !closed) return null;
  return Object.freeze({ commands: Object.freeze(commands), closed });
}

function simpleLineEndpointsFromSubpath(
  subpath: PdfVectorSimpleSubpath | null,
): PdfVectorSimpleLineEndpoints | null {
  if (
    !subpath ||
    subpath.closed ||
    subpath.commands.length !== 2 ||
    subpath.commands[0].kind !== "moveTo" ||
    subpath.commands[1].kind !== "lineTo"
  ) {
    return null;
  }
  return Object.freeze({
    start: subpath.commands[0].point,
    end: subpath.commands[1].point,
  });
}

function normalizedVisibleBounds(
  matrix: Matrix,
  minMax: unknown,
  page: Readonly<{ width: number; height: number }>,
): Readonly<{
  x: number;
  y: number;
  width: number;
  height: number;
  extent: Readonly<{ x: number; y: number; width: number; height: number }>;
}> | null {
  if (!Array.isArray(minMax) && !ArrayBuffer.isView(minMax)) return null;
  const values = Array.from(minMax as ArrayLike<unknown>, Number);
  if (
    values.length < 4 ||
    values.slice(0, 4).some((value) => !Number.isFinite(value)) ||
    values[2] < values[0] ||
    values[3] < values[1]
  ) {
    return null;
  }
  const corners = [
    mapPoint(matrix, values[0], values[1]),
    mapPoint(matrix, values[2], values[1]),
    mapPoint(matrix, values[2], values[3]),
    mapPoint(matrix, values[0], values[3]),
  ];
  let rawLeft = Math.min(...corners.map((point) => point.x));
  let rawRight = Math.max(...corners.map((point) => point.x));
  let rawBottom = Math.min(...corners.map((point) => point.y));
  let rawTop = Math.max(...corners.map((point) => point.y));
  // The whole path as drawn, before the page cuts it or a flat axis is given
  // a hair of depth to be found by.
  const extent = Object.freeze({
    x: rawLeft / page.width,
    y: (page.height - rawTop) / page.height,
    width: (rawRight - rawLeft) / page.width,
    height: (rawTop - rawBottom) / page.height,
  });
  if (rawRight <= rawLeft && rawTop > rawBottom) {
    rawLeft -= 0.25;
    rawRight += 0.25;
  } else if (rawTop <= rawBottom && rawRight > rawLeft) {
    rawBottom -= 0.25;
    rawTop += 0.25;
  }
  const left = Math.max(0, rawLeft);
  const right = Math.min(
    page.width,
    rawRight,
  );
  const bottom = Math.max(0, rawBottom);
  const top = Math.min(
    page.height,
    rawTop,
  );
  if (right <= left || top <= bottom) return null;
  return Object.freeze({
    x: left / page.width,
    y: (page.height - top) / page.height,
    width: (right - left) / page.width,
    height: (top - bottom) / page.height,
    extent,
  });
}

/** The `LW` a pdf.js `setGState` sets, or null when it sets none. */
function graphicsStateLineWidth(args: unknown): number | null {
  const entries = Array.isArray(args) ? args[0] : null;
  if (!Array.isArray(entries)) return null;
  for (const entry of entries) {
    if (Array.isArray(entry) && entry[0] === "LW") {
      const value = Number(entry[1]);
      return Number.isFinite(value) && value >= 0 ? value : null;
    }
  }
  return null;
}

function paintMode(
  paintOperator: number,
  ids: PdfVectorContentOperatorIds,
): PdfVectorPaintMode | null {
  const stroke = ids.strokeOps.has(paintOperator);
  const fill = ids.fillOps.has(paintOperator);
  return stroke && fill
    ? "fill-stroke"
    : stroke
      ? "stroke"
      : fill
        ? "fill"
        : null;
}

function normalizedAngle(matrix: Matrix): number {
  const raw = (Math.atan2(matrix[1], matrix[0]) * 180) / Math.PI;
  const normalized = Math.abs(raw) < 0.000_001 ? 0 : raw;
  return Number(normalized.toFixed(6));
}

function vectorKind(
  bounds: Readonly<{ width: number; height: number }>,
  page: Readonly<{ width: number; height: number }>,
  values: number,
): PdfVectorContentKind {
  const widthPt = bounds.width * page.width;
  const heightPt = bounds.height * page.height;
  if (
    Math.min(widthPt, heightPt) <= RULE_MAX_THICKNESS_PT &&
    Math.max(widthPt, heightPt) >= 12
  ) {
    return "rule";
  }
  if (values <= 24 && widthPt > 2.5 && heightPt > 2.5) return "rect-like";
  return "generic-path";
}

/**
 * Inventory every painted pdf.js constructPath operation without pretending
 * the renderer's path payload is a source-byte writer authority. The returned
 * items are selection/recognition evidence only; native editing still needs an
 * authenticated content-stream operator span from the byte engine.
 */
/**
 * The renderer's own page transform (pdf.js `viewport.transform` at scale 1),
 * folded into the inventory's y-up page frame. It moves the CropBox corner to
 * the origin and applies /Rotate: the operator list is in user space, and a
 * page whose CropBox does not start at 0,0 (a print-ready page with a bleed)
 * or that is rotated shows its vectors somewhere else than user space says.
 * Without it such a page's vectors never meet the byte engine's proofs,
 * which are projected through the CropBox and /Rotate.
 */
function pageBaseMatrix(
  page: Readonly<{ width: number; height: number; transform?: ArrayLike<number> }>,
): Matrix {
  const transform = page.transform;
  if (!transform) return [...IDENTITY];
  const values = Array.from(transform, Number);
  if (values.length !== 6 || values.some((value) => !Number.isFinite(value))) {
    return [...IDENTITY];
  }
  return multiplyMatrix([1, 0, 0, -1, 0, page.height], values as Matrix) as Matrix;
}

/** A clip edge closer than this (page fraction) to the item's own edge does
 * not count as cutting it: a layout program's frame clip sits a rounding gap
 * outside its fill. */
const CLIP_CUT_EPSILON = 0.000_5;

function intersectBoxes(
  a: Readonly<{ x: number; y: number; width: number; height: number }>,
  b: Readonly<{ x: number; y: number; width: number; height: number }>,
): Readonly<{ x: number; y: number; width: number; height: number }> {
  const left = Math.max(a.x, b.x);
  const top = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  return Object.freeze({
    x: left,
    y: top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  });
}

export function pdfVectorContentFromOperatorList(
  operatorList: Readonly<{
    fnArray: ArrayLike<number>;
    argsArray: ArrayLike<unknown>;
  }>,
  ids: PdfVectorContentOperatorIds,
  page: Readonly<{
    width: number;
    height: number;
    /** pdf.js `viewport.transform` at scale 1 (the viewport the page is shown in). */
    transform?: ArrayLike<number>;
  }>,
): readonly PdfVectorContentItem[] {
  if (!(page.width > 0) || !(page.height > 0)) return Object.freeze([]);
  const count = Math.min(
    operatorList.fnArray.length,
    operatorList.argsArray.length,
  );
  const base = pageBaseMatrix(page);
  const stack: Matrix[] = [];
  // The user-space transform alone: a vector's angle is its own, not the
  // page's /Rotate.
  const userStack: Matrix[] = [];
  const knownStack: boolean[] = [];
  // The line width in effect (PDF's default is 1), saved and restored with
  // the transform; only read when the list names the operators that set it.
  const tracksLineWidth = ids.setLineWidth !== undefined;
  const lineWidthStack: number[] = [];
  let lineWidth = 1;
  let matrix: Matrix = [...base];
  let userMatrix: Matrix = [...IDENTITY];
  let matrixKnown = true;
  let annotationDepth = 0;
  // The clip in effect as one box (the intersection of every clip's box),
  // saved and restored with the transform. pdf.js emits `clip` before the
  // path that defines it, so a clip is pending until that `constructPath`.
  const tracksClip = ids.clip !== undefined || ids.eoClip !== undefined;
  type ClipBox = Readonly<{ x: number; y: number; width: number; height: number }>;
  const clipStack: (ClipBox | null)[] = [];
  let clipBox: ClipBox | null = null;
  let pendingClip = false;
  const result: PdfVectorContentItem[] = [];
  for (let index = 0; index < count; index += 1) {
    const fn = operatorList.fnArray[index];
    const args = operatorList.argsArray[index];
    if (ids.beginAnnotation !== undefined && fn === ids.beginAnnotation) {
      annotationDepth += 1;
      continue;
    }
    if (ids.endAnnotation !== undefined && fn === ids.endAnnotation) {
      annotationDepth = Math.max(0, annotationDepth - 1);
      continue;
    }
    if (tracksLineWidth && fn === ids.setLineWidth) {
      const value = Array.isArray(args) ? Number(args[0]) : Number.NaN;
      if (Number.isFinite(value) && value >= 0) lineWidth = value;
      continue;
    }
    if (tracksLineWidth && ids.setGState !== undefined && fn === ids.setGState) {
      lineWidth = graphicsStateLineWidth(args) ?? lineWidth;
      continue;
    }
    if (fn === ids.save || (ids.paintFormXObjectBegin !== undefined && fn === ids.paintFormXObjectBegin)) {
      stack.push([...matrix]);
      userStack.push([...userMatrix]);
      knownStack.push(matrixKnown);
      lineWidthStack.push(lineWidth);
      clipStack.push(clipBox);
      if (fn !== ids.save) {
        const raw = Array.isArray(args) ? (args[0] as unknown) : null;
        // pdf.js hands the form's /Matrix over as a Float32Array.
        const formMatrix =
          raw && (Array.isArray(raw) || ArrayBuffer.isView(raw))
            ? finiteMatrix(Array.from(raw as ArrayLike<number>))
            : null;
        if (formMatrix) {
          matrix = multiplyMatrix(matrix, formMatrix) as Matrix;
          userMatrix = multiplyMatrix(userMatrix, formMatrix) as Matrix;
        } else if (raw) {
          matrixKnown = false;
        }
      }
      continue;
    }
    if (fn === ids.restore || (ids.paintFormXObjectEnd !== undefined && fn === ids.paintFormXObjectEnd)) {
      matrix = stack.pop() ?? [...base];
      userMatrix = userStack.pop() ?? [...IDENTITY];
      matrixKnown = knownStack.pop() ?? false;
      lineWidth = lineWidthStack.pop() ?? 1;
      clipBox = clipStack.pop() ?? null;
      pendingClip = false;
      continue;
    }
    if (tracksClip && (fn === ids.clip || fn === ids.eoClip)) {
      pendingClip = true;
      continue;
    }
    if (fn === ids.transform) {
      const transform = finiteMatrix(args);
      if (transform) {
        matrix = multiplyMatrix(matrix, transform) as Matrix;
        userMatrix = multiplyMatrix(userMatrix, transform) as Matrix;
        matrixKnown =
          matrixKnown && matrix.every((entry) => Number.isFinite(entry));
      } else {
        matrixKnown = false;
      }
      continue;
    }
    if (ids.constructPath === undefined || fn !== ids.constructPath) continue;
    if (!Array.isArray(args) || args.length < 3) continue;
    const placed = annotationDepth > 0 ? userMatrix : matrix;
    if (pendingClip) {
      pendingClip = false;
      const region = matrixKnown ? normalizedVisibleBounds(placed, args[2], page) : null;
      clipBox = region
        ? intersectBoxes(clipBox ?? region, region)
        : clipBox;
    }
    const mode = paintMode(Number(args[0]), ids);
    if (!mode) continue;
    const measured = normalizedVisibleBounds(placed, args[2], page);
    if (!measured) continue;
    const { extent, ...bounds } = measured;
    const clipped = clipBox ? intersectBoxes(bounds, clipBox) : null;
    const clippedBounds =
      clipped &&
      (clipped.x - bounds.x > CLIP_CUT_EPSILON ||
        clipped.y - bounds.y > CLIP_CUT_EPSILON ||
        bounds.x + bounds.width - (clipped.x + clipped.width) > CLIP_CUT_EPSILON ||
        bounds.y + bounds.height - (clipped.y + clipped.height) > CLIP_CUT_EPSILON)
        ? clipped
        : null;
    const values = pathValueCount(args[1]);
    if (values > MAX_PATH_VALUE_COUNT) continue;
    const simpleSubpath = matrixKnown
      ? normalizedDirectedSimpleSubpath(args[1], placed, page)
      : null;
    const simpleLineEndpoints = simpleLineEndpointsFromSubpath(simpleSubpath);
    const rotatedOrSkewed =
      Math.abs(userMatrix[1]) > 0.000_001 ||
      Math.abs(userMatrix[2]) > 0.000_001;
    // The width as painted: the user-space width under the transform's
    // scale (a uniform one; a skewed transform reads its mean).
    const strokeWidth =
      tracksLineWidth && mode !== "fill" && matrixKnown
        ? lineWidth * Math.sqrt(Math.abs(placed[0] * placed[3] - placed[1] * placed[2]))
        : null;
    // The pen's round tip mapped through the transform: an ellipse whose
    // half-extents on the page axes are these.
    const strokeReach =
      strokeWidth !== null && Number.isFinite(strokeWidth)
        ? Object.freeze({
            x: ((lineWidth / 2) * Math.hypot(placed[0], placed[2])) / page.width,
            y: ((lineWidth / 2) * Math.hypot(placed[1], placed[3])) / page.height,
          })
        : null;
    result.push(
      Object.freeze({
        kind: vectorKind(bounds, page, values),
        paintMode: mode,
        operatorIndex: index,
        ...bounds,
        angle: normalizedAngle(userMatrix),
        rotatedOrSkewed,
        pathValueCount: values,
        simplePathEvidence: values <= SIMPLE_PATH_VALUE_COUNT,
        ...(simpleLineEndpoints ? { simpleLineEndpoints } : {}),
        ...(simpleSubpath ? { simpleSubpath } : {}),
        extent,
        ...(strokeWidth !== null && Number.isFinite(strokeWidth) ? { strokeWidth } : {}),
        ...(strokeReach && Number.isFinite(strokeReach.x) && Number.isFinite(strokeReach.y)
          ? { strokeReach }
          : {}),
        ...(clippedBounds ? { clippedBounds } : {}),
      }),
    );
    if (result.length >= MAX_VECTOR_CONTENT_ITEMS) break;
  }
  return Object.freeze(result);
}
