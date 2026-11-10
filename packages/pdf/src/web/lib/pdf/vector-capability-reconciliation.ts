/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfVectorObjectDiscovery } from "./engine/content-object-discovery";
import type { PdfContentObjectBounds } from "./engine/content-object-graph";
import type {
  PdfNativeVectorContentShape,
  PdfNativeVectorSubpathProof,
} from "./engine/native/vector-content-transaction";
import type {
  PdfVectorContentItem,
  PdfVectorSimpleLineEndpoints,
  PdfVectorSimplePathCommand,
  PdfVectorSimpleSubpath,
} from "./vector-content-inventory";

export const PDF_VECTOR_CAPABILITY_RECONCILIATION_SCHEMA_VERSION =
  "mrai.pdf-vector-capability-reconciliation.v1" as const;

export interface PdfVectorPageGeometry {
  readonly cropBox: Readonly<{
    xMin: number;
    yMin: number;
    xMax: number;
    yMax: number;
  }>;
  readonly rotation: 0 | 90 | 180 | 270;
}

export interface PdfVectorCapabilityReconciliationInput {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly visibleVectors: readonly PdfVectorContentItem[];
  readonly nativeProofs: readonly PdfNativeVectorContentShape[];
  readonly nativeSubpathProofs?: readonly PdfNativeVectorSubpathProof[];
}

export interface PdfVectorCapabilityReconciliationResult {
  readonly schemaVersion: typeof PDF_VECTOR_CAPABILITY_RECONCILIATION_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly discoveries: readonly PdfVectorObjectDiscovery[];
  readonly matches: readonly Readonly<{
    visibleOperatorIndex: number;
    nativeProofId: string;
    contentStreamObjectNumber: number;
    operatorStart: number;
    operatorEnd: number;
  }>[];
  readonly diagnostics: Readonly<{
    visibleCount: number;
    nativeProofCount: number;
    /** Legacy strict line/rectangle/ellipse match count. */
    matchedCount: number;
    /** Legacy count not matched by the strict shape family. */
    preserveOnlyCount: number;
    ambiguousVisibleCount: number;
    ambiguousProofCount: number;
    sourceOrPageMismatchCount: number;
    nonEditableProofCount: number;
    invalidProofCount: number;
    unmatchedProofCount: number;
    nativeSubpathProofCount: number;
    subpathMatchedCount: number;
    /** Unique matches across the strict-shape and generic-subpath families. */
    combinedMatchedCount: number;
    /** Visible vectors left preserve-only after both proof families. */
    combinedPreserveOnlyCount: number;
    ambiguousSubpathVisibleCount: number;
    ambiguousSubpathProofCount: number;
    ambiguousCombinedVisibleCount: number;
    ambiguousCombinedProofCount: number;
    subpathSourceOrPageMismatchCount: number;
    nonEditableSubpathProofCount: number;
    invalidSubpathProofCount: number;
    unmatchedSubpathProofCount: number;
  }>;
}

interface ProjectedProof {
  readonly proof: PdfNativeVectorContentShape;
  readonly bounds: PdfContentObjectBounds;
  readonly lineEndpoints: PdfVectorSimpleLineEndpoints | null;
  /** The renderer path points a rectangle or ellipse proof predicts. */
  readonly shapePoints: readonly Readonly<{ x: number; y: number }>[] | null;
}

interface ProjectedSubpathProof {
  readonly proof: PdfNativeVectorSubpathProof;
  readonly bounds: PdfContentObjectBounds;
  readonly subpath: PdfVectorSimpleSubpath;
}

type CombinedProjectedProof = Readonly<
  | { family: "shape"; projected: ProjectedProof }
  | { family: "subpath"; projected: ProjectedSubpathProof }
>;

const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const GEOMETRY_TOLERANCE = 0.004;
const LINE_ENDPOINT_TOLERANCE = 0.000_5;
const SUBPATH_POINT_TOLERANCE = 0.000_5;
/** Renderer points agree with a proof's to this (page fractions, about a
 * thousandth of a point) when they are the same path: pdf.js keeps path
 * coordinates as Float32, the byte engine reads them as written. */
const TIGHT_POINT_TOLERANCE = 0.000_002;
const SUBPATH_BOUNDS_TOLERANCE = 0.004;
const PAGE_EPSILON = 0.000_001;
/** A wiring diagram or an exploded parts view draws fifty thousand paths
 * and more on one page (mirror of the inventory's and the byte engine's
 * per-page limits). */
const MAX_VECTORS_PER_PAGE = 100_000;
const MAX_SUBPATH_COMMANDS = 4096;
/** Cells of the grids that find a vector's candidate proofs; every grid is
 * read over the vector's cell and the ring(s) around it, which hold every
 * proof that can match. A rectangle or ellipse proof can match only a
 * vector whose center lies within GEOMETRY_TOLERANCE of its own (two rings
 * of CENTER_CELL); a line proof only a vector whose start lies within
 * LINE_ENDPOINT_TOLERANCE of its own, and a subpath proof only one of as
 * many commands whose first point lies within SUBPATH_POINT_TOLERANCE of
 * its own (one ring of POINT_CELL). */
const CENTER_CELL = 0.0021;
const CENTER_RING = 2;
const POINT_CELL = 0.0006;
const POINT_RING = 1;

function rounded(value: number): number {
  const result = Number(value.toFixed(6));
  return Object.is(result, -0) ? 0 : result;
}

function exactSourceSha256(value: string): string {
  const normalized = value.trim().toLowerCase();
  if (!SHA256_PATTERN.test(normalized)) {
    throw new Error("PDF vector reconciliation requires a source SHA-256.");
  }
  return normalized;
}

function exactPageNumber(value: number): number {
  if (!Number.isSafeInteger(value) || value < 1) {
    throw new Error("PDF vector reconciliation requires a page number.");
  }
  return value;
}

function exactPageGeometry(
  geometry: PdfVectorPageGeometry,
): PdfVectorPageGeometry {
  const crop = geometry.cropBox;
  if (
    ![crop.xMin, crop.yMin, crop.xMax, crop.yMax].every(Number.isFinite) ||
    crop.xMax <= crop.xMin ||
    crop.yMax <= crop.yMin ||
    ![0, 90, 180, 270].includes(geometry.rotation)
  ) {
    throw new Error("PDF vector reconciliation page geometry is invalid.");
  }
  return Object.freeze({
    cropBox: Object.freeze({ ...crop }),
    rotation: geometry.rotation,
  });
}

function exactVisibleBounds(
  vector: PdfVectorContentItem,
): PdfContentObjectBounds | null {
  if (
    ![vector.x, vector.y, vector.width, vector.height].every(Number.isFinite) ||
    vector.x < 0 ||
    vector.y < 0 ||
    vector.width <= 0 ||
    vector.height <= 0 ||
    vector.x + vector.width > 1 + PAGE_EPSILON ||
    vector.y + vector.height > 1 + PAGE_EPSILON
  ) {
    return null;
  }
  return Object.freeze({
    x: rounded(vector.x),
    y: rounded(vector.y),
    width: rounded(vector.width),
    height: rounded(vector.height),
    ...(Number.isFinite(vector.angle) ? { angle: rounded(vector.angle) } : {}),
  });
}

function projectedPoint(
  x: number,
  y: number,
  geometry: PdfVectorPageGeometry,
): Readonly<{ x: number; y: number }> {
  const crop = geometry.cropBox;
  if (geometry.rotation === 0) {
    return Object.freeze({ x: x - crop.xMin, y: crop.yMax - y });
  }
  if (geometry.rotation === 90) {
    return Object.freeze({ x: y - crop.yMin, y: x - crop.xMin });
  }
  if (geometry.rotation === 180) {
    return Object.freeze({ x: crop.xMax - x, y: y - crop.yMin });
  }
  return Object.freeze({ x: crop.yMax - y, y: crop.xMax - x });
}

function projectedPageSize(
  geometry: PdfVectorPageGeometry,
): Readonly<{ width: number; height: number }> {
  const cropWidth = geometry.cropBox.xMax - geometry.cropBox.xMin;
  const cropHeight = geometry.cropBox.yMax - geometry.cropBox.yMin;
  return Object.freeze({
    width:
      geometry.rotation === 0 || geometry.rotation === 180
        ? cropWidth
        : cropHeight,
    height:
      geometry.rotation === 0 || geometry.rotation === 180
        ? cropHeight
        : cropWidth,
  });
}

export function projectPdfNativeVectorLineEndpoints(
  proof: Pick<PdfNativeVectorContentShape, "kind" | "geometry">,
  geometryInput: PdfVectorPageGeometry,
): PdfVectorSimpleLineEndpoints | null {
  const geometry = exactPageGeometry(geometryInput);
  if (proof.kind !== "line" || proof.geometry.kind !== "line") return null;
  const page = projectedPageSize(geometry);
  const start = projectedPoint(
    proof.geometry.start.x,
    proof.geometry.start.y,
    geometry,
  );
  const end = projectedPoint(
    proof.geometry.end.x,
    proof.geometry.end.y,
    geometry,
  );
  const normalizedStart = Object.freeze({
    x: start.x / page.width,
    y: start.y / page.height,
  });
  const normalizedEnd = Object.freeze({
    x: end.x / page.width,
    y: end.y / page.height,
  });
  if (
    ![
      normalizedStart.x,
      normalizedStart.y,
      normalizedEnd.x,
      normalizedEnd.y,
    ].every(Number.isFinite) ||
    Math.hypot(
      normalizedEnd.x - normalizedStart.x,
      normalizedEnd.y - normalizedStart.y,
    ) <= 0.000_000_001
  ) {
    return null;
  }
  return Object.freeze({ start: normalizedStart, end: normalizedEnd });
}

function normalizedProjectedPoint(
  value: Readonly<{ x: number; y: number }>,
  geometry: PdfVectorPageGeometry,
): Readonly<{ x: number; y: number }> | null {
  if (!Number.isFinite(value.x) || !Number.isFinite(value.y)) return null;
  const projected = projectedPoint(value.x, value.y, geometry);
  const page = projectedPageSize(geometry);
  const normalized = Object.freeze({
    x: projected.x / page.width,
    y: projected.y / page.height,
  });
  return Number.isFinite(normalized.x) && Number.isFinite(normalized.y)
    ? normalized
    : null;
}

function nativeSubpathPaintMode(
  operator: PdfNativeVectorSubpathProof["paintOperator"],
): PdfNativeVectorSubpathProof["paintMode"] | null {
  if (operator === "S" || operator === "s") return "stroke";
  if (operator === "f" || operator === "F" || operator === "f*") {
    return "fill";
  }
  if (["B", "B*", "b", "b*"].includes(operator)) return "fill-stroke";
  return null;
}

function closingNativePaint(
  operator: PdfNativeVectorSubpathProof["paintOperator"],
): boolean {
  return operator === "s" || operator === "b" || operator === "b*";
}

function pointDistance(
  left: Readonly<{ x: number; y: number }>,
  right: Readonly<{ x: number; y: number }>,
): number {
  return Math.hypot(left.x - right.x, left.y - right.y);
}

export function projectPdfNativeVectorSubpath(
  proof: Pick<
    PdfNativeVectorSubpathProof,
    "commands" | "closed" | "paintMode" | "paintOperator"
  >,
  geometryInput: PdfVectorPageGeometry,
): PdfVectorSimpleSubpath | null {
  const geometry = exactPageGeometry(geometryInput);
  if (
    proof.commands.length < 2 ||
    proof.commands.length > MAX_SUBPATH_COMMANDS ||
    proof.commands[0]?.operator !== "m" ||
    nativeSubpathPaintMode(proof.paintOperator) !== proof.paintMode
  ) {
    return null;
  }
  // One path of one or more subpaths, as the renderer lists it: a move opens
  // each subpath after the first (once the one before drew something or was
  // closed), a close may end any of them, and a segment straight after a
  // close starts the next subpath at the closed one's start.
  const commands: PdfVectorSimplePathCommand[] = [];
  let currentPoint: Readonly<{ x: number; y: number }> | null = null;
  let startPoint: Readonly<{ x: number; y: number }> | null = null;
  let segments = 0;
  let explicitClose = false;
  for (const command of proof.commands) {
    if (command.operator === "m") {
      if (currentPoint && segments === 0 && !explicitClose) return null;
      const projected = normalizedProjectedPoint(command.point, geometry);
      if (!projected) return null;
      commands.push(Object.freeze({ kind: "moveTo", point: projected }));
      currentPoint = command.point;
      startPoint = command.point;
      segments = 0;
      explicitClose = false;
      continue;
    }
    if (!currentPoint) return null;
    if (command.operator === "h") {
      if (explicitClose) return null;
      commands.push(Object.freeze({ kind: "closePath" }));
      currentPoint = startPoint;
      explicitClose = true;
      continue;
    }
    explicitClose = false;
    if (command.operator === "l") {
      const projected = normalizedProjectedPoint(command.point, geometry);
      if (!projected) return null;
      commands.push(Object.freeze({ kind: "lineTo", point: projected }));
      currentPoint = command.point;
      segments += 1;
      continue;
    }
    if (!("control1" in command) || !currentPoint) return null;
    if (
      command.operator === "v" &&
      pointDistance(command.control1, currentPoint) > PAGE_EPSILON
    ) {
      return null;
    }
    if (
      command.operator === "y" &&
      pointDistance(command.control2, command.end) > PAGE_EPSILON
    ) {
      return null;
    }
    const control1 = normalizedProjectedPoint(command.control1, geometry);
    const control2 = normalizedProjectedPoint(command.control2, geometry);
    const end = normalizedProjectedPoint(command.end, geometry);
    if (!control1 || !control2 || !end) return null;
    commands.push(Object.freeze({ kind: "curveTo", control1, control2, end }));
    currentPoint = command.end;
    segments += 1;
  }
  if (segments === 0 && !explicitClose) return null;
  // A closing paint (s, b, b*) closes the last subpath; the renderer lists
  // that close as a command of its own.
  const closingPaint = closingNativePaint(proof.paintOperator);
  if (closingPaint && !explicitClose) {
    commands.push(Object.freeze({ kind: "closePath" }));
  }
  const closed = explicitClose || closingPaint;
  if (proof.closed !== closed) return null;
  return Object.freeze({ commands: Object.freeze(commands), closed });
}

/**
 * Project a byte-engine user-space rectangle through the authoritative page
 * CropBox/Rotate contract into the same top-left page fractions used by
 * pdf.js overlays. Partial visibility is clipped exactly like the renderer.
 */
export function projectPdfNativeVectorBounds(
  bounds: PdfNativeVectorContentShape["bounds"],
  geometryInput: PdfVectorPageGeometry,
): PdfContentObjectBounds | null {
  const geometry = exactPageGeometry(geometryInput);
  if (
    ![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite) ||
    bounds.width < 0 ||
    bounds.height < 0
  ) {
    return null;
  }
  const corners = [
    projectedPoint(bounds.x, bounds.y, geometry),
    projectedPoint(bounds.x + bounds.width, bounds.y, geometry),
    projectedPoint(bounds.x + bounds.width, bounds.y + bounds.height, geometry),
    projectedPoint(bounds.x, bounds.y + bounds.height, geometry),
  ];
  const page = projectedPageSize(geometry);
  const pageWidth = page.width;
  const pageHeight = page.height;
  const rawLeft = Math.min(...corners.map((corner) => corner.x));
  const rawTop = Math.min(...corners.map((corner) => corner.y));
  const rawRight = Math.max(...corners.map((corner) => corner.x));
  const rawBottom = Math.max(...corners.map((corner) => corner.y));
  let left = Math.max(0, rawLeft);
  let top = Math.max(0, rawTop);
  let right = Math.min(pageWidth, rawRight);
  let bottom = Math.min(pageHeight, rawBottom);
  if (right <= left && rawRight === rawLeft && left >= 0 && left <= pageWidth) {
    left = Math.max(0, left - 0.125);
    right = Math.min(pageWidth, right + 0.125);
  }
  if (bottom <= top && rawBottom === rawTop && top >= 0 && top <= pageHeight) {
    top = Math.max(0, top - 0.125);
    bottom = Math.min(pageHeight, bottom + 0.125);
  }
  if (right <= left || bottom <= top) return null;
  return Object.freeze({
    x: rounded(left / pageWidth),
    y: rounded(top / pageHeight),
    width: rounded((right - left) / pageWidth),
    height: rounded((bottom - top) / pageHeight),
  });
}

function kindCompatible(
  visible: PdfVectorContentItem,
  proof: PdfNativeVectorContentShape,
): boolean {
  if (proof.kind === "line") {
    return (
      visible.kind === "rule" ||
      visible.kind === "rect-like" ||
      visible.kind === "generic-path"
    );
  }
  if (proof.kind === "rectangle") return visible.kind === "rect-like";
  return visible.kind === "rect-like" || visible.kind === "generic-path";
}

function exactVisibleLineEndpoints(
  visible: PdfVectorContentItem,
): PdfVectorSimpleLineEndpoints | null {
  if (
    visible.paintMode !== "stroke" ||
    !visible.simplePathEvidence ||
    visible.pathValueCount > 8
  ) {
    return null;
  }
  const endpoints = visible.simpleLineEndpoints;
  if (!endpoints) return null;
  const values = [
    endpoints.start?.x,
    endpoints.start?.y,
    endpoints.end?.x,
    endpoints.end?.y,
  ];
  if (
    values.some(
      (value) => typeof value !== "number" || !Number.isFinite(value),
    ) ||
    Math.hypot(
      endpoints.end.x - endpoints.start.x,
      endpoints.end.y - endpoints.start.y,
    ) <= 0.000_000_001
  ) {
    return null;
  }
  return endpoints;
}

function exactNormalizedCommandPoint(
  value: Readonly<{ x: number; y: number }> | undefined,
): Readonly<{ x: number; y: number }> | null {
  if (
    !value ||
    !Number.isFinite(value.x) ||
    !Number.isFinite(value.y) ||
    Math.abs(value.x) > 16 ||
    Math.abs(value.y) > 16
  ) {
    return null;
  }
  return Object.freeze({ x: value.x, y: value.y });
}

function exactVisibleSubpath(
  visible: PdfVectorContentItem,
): PdfVectorSimpleSubpath | null {
  const subpath = visible.simpleSubpath;
  if (
    !subpath ||
    subpath.commands.length < 2 ||
    subpath.commands.length > MAX_SUBPATH_COMMANDS ||
    subpath.commands[0]?.kind !== "moveTo"
  ) {
    return null;
  }
  // The renderer's path, one or more subpaths (see the inventory): a move
  // opens each, a close may end any, a segment after a close goes on from
  // the closed subpath's start.
  const commands: PdfVectorSimplePathCommand[] = [];
  let started = false;
  let segments = 0;
  let closed = false;
  for (const command of subpath.commands) {
    if (command.kind === "quadraticCurveTo") return null;
    if (command.kind === "moveTo") {
      if (started && segments === 0 && !closed) return null;
      const point = exactNormalizedCommandPoint(command.point);
      if (!point) return null;
      commands.push(Object.freeze({ kind: "moveTo", point }));
      started = true;
      segments = 0;
      closed = false;
      continue;
    }
    if (!started) return null;
    if (command.kind === "closePath") {
      if (closed) return null;
      commands.push(Object.freeze({ kind: "closePath" }));
      closed = true;
      continue;
    }
    closed = false;
    segments += 1;
    if (command.kind === "lineTo") {
      const point = exactNormalizedCommandPoint(command.point);
      if (!point) return null;
      commands.push(Object.freeze({ kind: "lineTo", point }));
      continue;
    }
    if (!("control1" in command)) return null;
    const control1 = exactNormalizedCommandPoint(command.control1);
    const control2 = exactNormalizedCommandPoint(command.control2);
    const end = exactNormalizedCommandPoint(command.end);
    if (!control1 || !control2 || !end) return null;
    commands.push(Object.freeze({ kind: "curveTo", control1, control2, end }));
  }
  if (subpath.closed !== closed) return null;
  return Object.freeze({ commands: Object.freeze(commands), closed });
}

function subpathBoundsCompatible(
  visible: PdfContentObjectBounds,
  proof: PdfContentObjectBounds,
): boolean {
  return (
    Math.abs(visible.x - proof.x) <= SUBPATH_BOUNDS_TOLERANCE &&
    Math.abs(visible.y - proof.y) <= SUBPATH_BOUNDS_TOLERANCE &&
    Math.abs(visible.width - proof.width) <= SUBPATH_BOUNDS_TOLERANCE &&
    Math.abs(visible.height - proof.height) <= SUBPATH_BOUNDS_TOLERANCE
  );
}

function commandPointCompatible(
  left: Readonly<{ x: number; y: number }>,
  right: Readonly<{ x: number; y: number }>,
  tolerance = SUBPATH_POINT_TOLERANCE,
): boolean {
  return pointDistance(left, right) <= tolerance;
}

function subpathCommandsCompatible(
  visible: PdfVectorSimpleSubpath,
  proof: PdfVectorSimpleSubpath,
  tolerance = SUBPATH_POINT_TOLERANCE,
): boolean {
  if (
    visible.closed !== proof.closed ||
    visible.commands.length !== proof.commands.length
  ) {
    return false;
  }
  for (let index = 0; index < visible.commands.length; index += 1) {
    const left = visible.commands[index];
    const right = proof.commands[index];
    if (left.kind !== right.kind) return false;
    if (left.kind === "closePath" || right.kind === "closePath") {
      if (left.kind !== "closePath" || right.kind !== "closePath") return false;
      continue;
    }
    if ("point" in left || "point" in right) {
      if (!("point" in left) || !("point" in right)) return false;
      if (!commandPointCompatible(left.point, right.point, tolerance)) return false;
      continue;
    }
    if (
      !("control1" in left) ||
      !("control1" in right) ||
      !commandPointCompatible(left.control1, right.control1, tolerance) ||
      !commandPointCompatible(left.control2, right.control2, tolerance) ||
      !commandPointCompatible(left.end, right.end, tolerance)
    ) {
      return false;
    }
  }
  return true;
}

/**
 * The path points a rectangle or ellipse proof predicts the renderer draws,
 * in the page fractions the renderer's own points are read in: a
 * rectangle's four corners in drawing order (`re` and `m l l l h` both start
 * at the first), an ellipse's start and its four curve ends.
 */
function projectedShapePoints(
  proof: PdfNativeVectorContentShape,
  geometry: PdfVectorPageGeometry,
): readonly Readonly<{ x: number; y: number }>[] | null {
  const shape = proof.geometry;
  let points: readonly Readonly<{ x: number; y: number }>[];
  if (shape.kind === "rectangle") {
    points = shape.corners;
  } else if (shape.kind === "ellipse") {
    const { center, axisX, axisY } = shape;
    points = [
      { x: center.x + axisX.x, y: center.y + axisX.y },
      { x: center.x + axisY.x, y: center.y + axisY.y },
      { x: center.x - axisX.x, y: center.y - axisX.y },
      { x: center.x - axisY.x, y: center.y - axisY.y },
      { x: center.x + axisX.x, y: center.y + axisX.y },
    ];
  } else {
    return null;
  }
  const projected = points.map((point) =>
    normalizedProjectedPoint(point, geometry),
  );
  return projected.every((point) => point !== null)
    ? Object.freeze(projected as Readonly<{ x: number; y: number }>[])
    : null;
}

/**
 * Whether the renderer's own path is the one the shape proof names, point
 * for point: a rectangle as move, three lines and a close through its four
 * corners; an ellipse as a move and four curves through its ends. Such a
 * match needs no size class -- a two-point filled cell, a hairline rule
 * drawn as a thin rectangle, a bullet dot -- and it tells apart shapes that
 * lie closer together than the bounds tolerance.
 */
function shapePointsExact(
  visible: PdfVectorContentItem,
  projected: ProjectedProof,
  tolerance = SUBPATH_POINT_TOLERANCE,
): boolean {
  const points = projected.shapePoints;
  const commands = visible.simpleSubpath?.commands;
  if (!points || !commands) return false;
  const near = (
    point: Readonly<{ x: number; y: number }>,
    expected: Readonly<{ x: number; y: number }>,
  ) => pointDistance(point, expected) <= tolerance;
  if (projected.proof.kind === "rectangle") {
    if (commands.length !== 5 || commands[4].kind !== "closePath") return false;
    return [0, 1, 2, 3].every((index) => {
      const command = commands[index];
      return (
        command.kind === (index === 0 ? "moveTo" : "lineTo") &&
        "point" in command &&
        near(command.point, points[index])
      );
    });
  }
  if (
    (commands.length !== 5 && commands.length !== 6) ||
    (commands.length === 6 && commands[5].kind !== "closePath")
  ) {
    return false;
  }
  const first = commands[0];
  if (first.kind !== "moveTo" || !near(first.point, points[0])) return false;
  return [1, 2, 3, 4].every((index) => {
    const command = commands[index];
    return command.kind === "curveTo" && near(command.end, points[index]);
  });
}

function lineEndpointsCompatible(
  visible: PdfVectorContentItem,
  proof: ProjectedProof,
  tolerance = LINE_ENDPOINT_TOLERANCE,
): boolean {
  if (proof.proof.kind !== "line") return true;
  const visibleEndpoints = exactVisibleLineEndpoints(visible);
  const nativeEndpoints = proof.lineEndpoints;
  if (!visibleEndpoints || !nativeEndpoints) return false;
  return (
    Math.hypot(
      visibleEndpoints.start.x - nativeEndpoints.start.x,
      visibleEndpoints.start.y - nativeEndpoints.start.y,
    ) <= tolerance &&
    Math.hypot(
      visibleEndpoints.end.x - nativeEndpoints.end.x,
      visibleEndpoints.end.y - nativeEndpoints.end.y,
    ) <= tolerance
  );
}

function dimensionCompatible(left: number, right: number): boolean {
  const minimum = Math.min(left, right);
  const maximum = Math.max(left, right);
  if (minimum <= GEOMETRY_TOLERANCE) {
    return maximum <= GEOMETRY_TOLERANCE * 2;
  }
  return (
    Math.abs(left - right) <= Math.max(GEOMETRY_TOLERANCE, minimum * 0.025)
  );
}

function boundsCompatible(
  visible: PdfContentObjectBounds,
  proof: PdfContentObjectBounds,
): boolean {
  const visibleCenterX = visible.x + visible.width / 2;
  const visibleCenterY = visible.y + visible.height / 2;
  const proofCenterX = proof.x + proof.width / 2;
  const proofCenterY = proof.y + proof.height / 2;
  return (
    Math.abs(visibleCenterX - proofCenterX) <= GEOMETRY_TOLERANCE &&
    Math.abs(visibleCenterY - proofCenterY) <= GEOMETRY_TOLERANCE &&
    dimensionCompatible(visible.width, proof.width) &&
    dimensionCompatible(visible.height, proof.height)
  );
}

function proofIsValid(
  proof: PdfNativeVectorContentShape,
  sourceSha256: string,
  pageNumber: number,
): boolean {
  return (
    proof.sourceSha256 === sourceSha256 &&
    proof.pageNumber === pageNumber &&
    proof.editable &&
    proof.preserveReason === null &&
    SHA256_PATTERN.test(proof.proofDigest) &&
    Number.isSafeInteger(proof.contentStream.objectNumber) &&
    proof.contentStream.objectNumber >= 1 &&
    Number.isSafeInteger(proof.contentStream.generation) &&
    proof.contentStream.generation >= 0 &&
    Number.isSafeInteger(proof.operatorSpan.start) &&
    proof.operatorSpan.start >= 0 &&
    Number.isSafeInteger(proof.operatorSpan.end) &&
    proof.operatorSpan.end > proof.operatorSpan.start &&
    SHA256_PATTERN.test(proof.operatorSpan.expectedSha256) &&
    proof.geometry.kind === proof.kind
  );
}

function subpathProofIsValid(
  proof: PdfNativeVectorSubpathProof,
  sourceSha256: string,
  pageNumber: number,
): boolean {
  return (
    proof.schemaVersion === "mrai.native-vector-subpath.v1" &&
    proof.sourceSha256 === sourceSha256 &&
    proof.pageNumber === pageNumber &&
    Number.isSafeInteger(proof.pageSubpathIndex) &&
    proof.pageSubpathIndex >= 0 &&
    proof.editable &&
    proof.preserveReason === null &&
    SHA256_PATTERN.test(proof.proofDigest) &&
    Number.isSafeInteger(proof.contentStream.objectNumber) &&
    proof.contentStream.objectNumber >= 1 &&
    Number.isSafeInteger(proof.contentStream.generation) &&
    proof.contentStream.generation >= 0 &&
    Number.isSafeInteger(proof.operatorSpan.start) &&
    proof.operatorSpan.start >= 0 &&
    Number.isSafeInteger(proof.operatorSpan.end) &&
    proof.operatorSpan.end > proof.operatorSpan.start &&
    SHA256_PATTERN.test(proof.operatorSpan.expectedSha256) &&
    [
      proof.bounds.x,
      proof.bounds.y,
      proof.bounds.width,
      proof.bounds.height,
    ].every(Number.isFinite) &&
    proof.bounds.width >= 0 &&
    proof.bounds.height >= 0 &&
    Math.max(proof.bounds.width, proof.bounds.height) > PAGE_EPSILON
  );
}

async function sha256Text(value: string): Promise<string> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle)
    throw new Error("PDF vector reconciliation SHA-256 is unavailable.");
  const digest = await subtle.digest(
    "SHA-256",
    new TextEncoder().encode(value),
  );
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}

async function visibleRegionDigest(
  sourceSha256: string,
  pageNumber: number,
  vector: PdfVectorContentItem,
  bounds: PdfContentObjectBounds,
): Promise<string> {
  return sha256Text(
    JSON.stringify({
      schemaVersion: PDF_VECTOR_CAPABILITY_RECONCILIATION_SCHEMA_VERSION,
      sourceSha256,
      pageNumber,
      operatorIndex: vector.operatorIndex,
      kind: vector.kind,
      paintMode: vector.paintMode,
      bounds: {
        x: rounded(bounds.x),
        y: rounded(bounds.y),
        width: rounded(bounds.width),
        height: rounded(bounds.height),
      },
    }),
  );
}

export async function reconcilePdfVectorCapabilities(
  input: PdfVectorCapabilityReconciliationInput,
): Promise<PdfVectorCapabilityReconciliationResult> {
  const sourceSha256 = exactSourceSha256(input.sourceSha256);
  const pageNumber = exactPageNumber(input.pageNumber);
  const pageGeometry = exactPageGeometry(input.pageGeometry);
  const nativeSubpathProofs = input.nativeSubpathProofs ?? [];
  if (
    input.visibleVectors.length > MAX_VECTORS_PER_PAGE ||
    input.nativeProofs.length + nativeSubpathProofs.length >
      MAX_VECTORS_PER_PAGE
  ) {
    throw new Error("PDF vector reconciliation page capacity is exceeded.");
  }

  const visible = input.visibleVectors
    .map((vector, inputIndex) => ({
      vector,
      inputIndex,
      bounds: exactVisibleBounds(vector),
    }))
    .filter(
      (
        item,
      ): item is {
        vector: PdfVectorContentItem;
        inputIndex: number;
        bounds: PdfContentObjectBounds;
      } => item.bounds !== null,
    );

  let sourceOrPageMismatchCount = 0;
  let nonEditableProofCount = 0;
  let invalidProofCount = 0;
  const projectedProofs: ProjectedProof[] = [];
  for (const proof of input.nativeProofs) {
    if (
      proof.sourceSha256 !== sourceSha256 ||
      proof.pageNumber !== pageNumber
    ) {
      sourceOrPageMismatchCount += 1;
      continue;
    }
    if (!proof.editable || proof.preserveReason !== null) {
      nonEditableProofCount += 1;
      continue;
    }
    if (!proofIsValid(proof, sourceSha256, pageNumber)) {
      invalidProofCount += 1;
      continue;
    }
    const bounds = projectPdfNativeVectorBounds(proof.bounds, pageGeometry);
    const lineEndpoints =
      proof.kind === "line"
        ? projectPdfNativeVectorLineEndpoints(proof, pageGeometry)
        : null;
    if (!bounds || (proof.kind === "line" && !lineEndpoints)) {
      invalidProofCount += 1;
      continue;
    }
    projectedProofs.push(
      Object.freeze({
        proof,
        bounds,
        lineEndpoints,
        shapePoints: projectedShapePoints(proof, pageGeometry),
      }),
    );
  }

  let subpathSourceOrPageMismatchCount = 0;
  let nonEditableSubpathProofCount = 0;
  let invalidSubpathProofCount = 0;
  const projectedSubpathProofs: ProjectedSubpathProof[] = [];
  for (const proof of nativeSubpathProofs) {
    if (
      proof.sourceSha256 !== sourceSha256 ||
      proof.pageNumber !== pageNumber
    ) {
      subpathSourceOrPageMismatchCount += 1;
      continue;
    }
    if (!proof.editable || proof.preserveReason !== null) {
      nonEditableSubpathProofCount += 1;
      continue;
    }
    if (!subpathProofIsValid(proof, sourceSha256, pageNumber)) {
      invalidSubpathProofCount += 1;
      continue;
    }
    const bounds = projectPdfNativeVectorBounds(proof.bounds, pageGeometry);
    const subpath = projectPdfNativeVectorSubpath(proof, pageGeometry);
    if (!bounds || !subpath) {
      invalidSubpathProofCount += 1;
      continue;
    }
    projectedSubpathProofs.push(Object.freeze({ proof, bounds, subpath }));
  }

  // Candidate pairs by paint mode, bounds and path evidence. A pair whose
  // renderer points coincide with the proof's (a line's endpoints, a
  // rectangle's corners, an ellipse's curve ends, a subpath's commands) is
  // exact; where a vector or a proof has an exact partner, only exact pairs
  // stand for it. Two shapes closer together than the bounds tolerance --
  // dots of a dotted rule, cells of a dense table -- then stay apart.
  const visibleSubpaths = visible.map((item) =>
    exactVisibleSubpath(item.vector),
  );
  // Proofs by paint mode and grid cell: a rectangle or an ellipse by its
  // center, a line by its start, a subpath by its first point (and its
  // command count). A dense page -- fifty thousand paths, a schematic's
  // outlined lettering -- then compares each vector with the few proofs that
  // can match it, not with every proof on the page. The candidates found are
  // exactly the ones a full scan finds, listed in the same ascending order.
  type Grid = Map<string, number[]>;
  const addToGrid = (grid: Grid, key: string, index: number) => {
    const cell = grid.get(key);
    if (cell) cell.push(index);
    else grid.set(key, [index]);
  };
  const around = (
    grid: Grid,
    prefix: string,
    x: number,
    y: number,
    size: number,
    ring: number,
  ): number[] => {
    const column = Math.floor(x / size);
    const row = Math.floor(y / size);
    const found: number[] = [];
    for (let dx = -ring; dx <= ring; dx += 1) {
      for (let dy = -ring; dy <= ring; dy += 1) {
        const cell = grid.get(`${prefix}:${column + dx}:${row + dy}`);
        if (cell) for (const index of cell) found.push(index);
      }
    }
    return found;
  };
  const cellKey = (prefix: string, x: number, y: number, size: number) =>
    `${prefix}:${Math.floor(x / size)}:${Math.floor(y / size)}`;
  const ascending = (left: number, right: number) => left - right;
  const lineGrid: Grid = new Map();
  const areaGrid: Grid = new Map();
  for (const [index, projected] of projectedProofs.entries()) {
    const mode = projected.proof.paintMode;
    if (projected.proof.kind === "line") {
      // A line proof always carries its endpoints (see above).
      const start = projected.lineEndpoints!.start;
      addToGrid(lineGrid, cellKey(mode, start.x, start.y, POINT_CELL), index);
    } else {
      const bounds = projected.bounds;
      addToGrid(
        areaGrid,
        cellKey(
          mode,
          bounds.x + bounds.width / 2,
          bounds.y + bounds.height / 2,
          CENTER_CELL,
        ),
        index,
      );
    }
  }
  const subpathGrid: Grid = new Map();
  for (const [index, projected] of projectedSubpathProofs.entries()) {
    const first = projected.subpath.commands[0];
    if (first?.kind !== "moveTo") continue;
    addToGrid(
      subpathGrid,
      cellKey(
        `${projected.proof.paintMode}:${projected.subpath.commands.length}`,
        first.point.x,
        first.point.y,
        POINT_CELL,
      ),
      index,
    );
  }
  // The centers boundsCompatible compares, computed the same way.
  const proofCenters = new Float64Array(projectedProofs.length * 2);
  for (const [index, projected] of projectedProofs.entries()) {
    proofCenters[index * 2] = projected.bounds.x + projected.bounds.width / 2;
    proofCenters[index * 2 + 1] =
      projected.bounds.y + projected.bounds.height / 2;
  }
  const shapeCandidatesOf = (item: (typeof visible)[number]): number[] => {
    const mode = item.vector.paintMode;
    const bounds = item.bounds;
    const centerX = bounds.x + bounds.width / 2;
    const centerY = bounds.y + bounds.height / 2;
    const found = around(
      areaGrid,
      mode,
      centerX,
      centerY,
      CENTER_CELL,
      CENTER_RING,
    ).filter(
      (index) =>
        Math.abs(centerX - proofCenters[index * 2]) <= GEOMETRY_TOLERANCE &&
        Math.abs(centerY - proofCenters[index * 2 + 1]) <= GEOMETRY_TOLERANCE,
    );
    const endpoints = exactVisibleLineEndpoints(item.vector);
    if (endpoints) {
      found.push(
        ...around(
          lineGrid,
          mode,
          endpoints.start.x,
          endpoints.start.y,
          POINT_CELL,
          POINT_RING,
        ),
      );
    }
    return found.sort(ascending);
  };
  const subpathCandidatesOf = (
    item: (typeof visible)[number],
    subpath: PdfVectorSimpleSubpath,
  ): number[] => {
    const first = subpath.commands[0];
    if (first?.kind !== "moveTo") return [];
    return around(
      subpathGrid,
      `${item.vector.paintMode}:${subpath.commands.length}`,
      first.point.x,
      first.point.y,
      POINT_CELL,
      POINT_RING,
    ).sort(ascending);
  };
  const proofCandidatesByVisible: number[][] = [];
  const subpathCandidatesByVisible: number[][] = [];
  const exactByVisible: Set<number>[] = [];
  const tightByVisible: Set<number>[] = [];
  const subpathOffset = projectedProofs.length;
  for (const [visibleIndex, item] of visible.entries()) {
    const shapes: number[] = [];
    const subpaths: number[] = [];
    const exact = new Set<number>();
    const tight = new Set<number>();
    for (const proofIndex of shapeCandidatesOf(item)) {
      const projected = projectedProofs[proofIndex];
      if (
        item.vector.paintMode !== projected.proof.paintMode ||
        !boundsCompatible(item.bounds, projected.bounds) ||
        !lineEndpointsCompatible(item.vector, projected)
      ) {
        continue;
      }
      const pointsExact =
        projected.proof.kind === "line" ||
        shapePointsExact(item.vector, projected);
      if (!pointsExact && !kindCompatible(item.vector, projected.proof)) {
        continue;
      }
      shapes.push(proofIndex);
      if (pointsExact) exact.add(proofIndex);
      if (
        projected.proof.kind === "line"
          ? lineEndpointsCompatible(item.vector, projected, TIGHT_POINT_TOLERANCE)
          : shapePointsExact(item.vector, projected, TIGHT_POINT_TOLERANCE)
      ) {
        tight.add(proofIndex);
      }
    }
    const visibleSubpath = visibleSubpaths[visibleIndex];
    if (visibleSubpath) {
      for (const proofIndex of subpathCandidatesOf(item, visibleSubpath)) {
        const projected = projectedSubpathProofs[proofIndex];
        if (
          item.vector.paintMode === projected.proof.paintMode &&
          subpathBoundsCompatible(item.bounds, projected.bounds) &&
          subpathCommandsCompatible(visibleSubpath, projected.subpath)
        ) {
          subpaths.push(proofIndex);
          exact.add(subpathOffset + proofIndex);
          if (subpathCommandsCompatible(visibleSubpath, projected.subpath, TIGHT_POINT_TOLERANCE)) {
            tight.add(subpathOffset + proofIndex);
          }
        }
      }
    }
    proofCandidatesByVisible.push(shapes);
    subpathCandidatesByVisible.push(subpaths);
    exactByVisible.push(exact);
    tightByVisible.push(tight);
  }
  const invert = (lists: readonly (readonly number[])[], size: number) => {
    const inverse: number[][] = Array.from({ length: size }, () => []);
    for (const [visibleIndex, list] of lists.entries()) {
      for (const proofIndex of list) inverse[proofIndex].push(visibleIndex);
    }
    return inverse;
  };
  const visibleCandidatesByProof = invert(
    proofCandidatesByVisible,
    projectedProofs.length,
  );
  const visibleCandidatesBySubpathProof = invert(
    subpathCandidatesByVisible,
    projectedSubpathProofs.length,
  );

  const combinedProjectedProofs: CombinedProjectedProof[] = [
    ...projectedProofs.map((projected) =>
      Object.freeze({ family: "shape" as const, projected }),
    ),
    ...projectedSubpathProofs.map((projected) =>
      Object.freeze({ family: "subpath" as const, projected }),
    ),
  ];
  const rawCombinedCandidatesByVisible = visible.map((_item, visibleIndex) => [
    ...proofCandidatesByVisible[visibleIndex],
    ...subpathCandidatesByVisible[visibleIndex].map(
      (proofIndex) => subpathOffset + proofIndex,
    ),
  ]);
  // Three levels of evidence: the same points to renderer precision, the
  // same points to the drawing tolerance, bounds alone. A vector or a proof
  // with a partner at a level keeps only partners at that level -- hatching
  // of tenth-of-a-point strokes lies well inside the drawing tolerance.
  const proofHasExact = new Set<number>();
  const proofHasTight = new Set<number>();
  for (const exact of exactByVisible) {
    for (const proofIndex of exact) proofHasExact.add(proofIndex);
  }
  for (const tight of tightByVisible) {
    for (const proofIndex of tight) proofHasTight.add(proofIndex);
  }
  const combinedCandidatesByVisible = rawCombinedCandidatesByVisible.map(
    (candidates, visibleIndex) => {
      const exact = exactByVisible[visibleIndex];
      const tight = tightByVisible[visibleIndex];
      return candidates.filter(
        (proofIndex) =>
          tight.has(proofIndex) ||
          (tight.size === 0 &&
            !proofHasTight.has(proofIndex) &&
            (exact.has(proofIndex) ||
              (exact.size === 0 && !proofHasExact.has(proofIndex)))),
      );
    },
  );
  const visibleCandidatesByCombinedProof = invert(
    combinedCandidatesByVisible,
    combinedProjectedProofs.length,
  );

  const uniqueMatches = new Map<number, number>();
  for (
    let visibleIndex = 0;
    visibleIndex < combinedCandidatesByVisible.length;
    visibleIndex += 1
  ) {
    const proofCandidates = combinedCandidatesByVisible[visibleIndex];
    if (proofCandidates.length !== 1) continue;
    const proofIndex = proofCandidates[0];
    if (visibleCandidatesByCombinedProof[proofIndex].length !== 1) continue;
    uniqueMatches.set(visibleIndex, proofIndex);
  }
  // The same path painted more than once in one stream (a table corner a
  // producer strokes twice, a fill laid over its own copy): n renderer paths
  // and n proofs that match each other exactly and nothing else. Both sides
  // list them in drawing order -- the renderer's operators and the stream's
  // bytes -- so the k-th paint is the k-th proof.
  const groups = new Map<string, number[]>();
  for (const [visibleIndex, candidates] of combinedCandidatesByVisible.entries()) {
    if (candidates.length < 2 || uniqueMatches.has(visibleIndex)) continue;
    const key = [...candidates].sort((left, right) => left - right).join(",");
    const group = groups.get(key);
    if (group) group.push(visibleIndex);
    else groups.set(key, [visibleIndex]);
  }
  for (const [key, visibleIndexes] of groups) {
    const proofIndexes = key.split(",").map(Number);
    const visibleKey = [...visibleIndexes].sort((left, right) => left - right).join(",");
    const proofs = proofIndexes.map((proofIndex) => combinedProjectedProofs[proofIndex].projected.proof);
    const stream = proofs[0].contentStream;
    if (
      proofIndexes.length !== visibleIndexes.length ||
      visibleIndexes.some((visibleIndex) =>
        proofIndexes.some((proofIndex) => !exactByVisible[visibleIndex].has(proofIndex)),
      ) ||
      proofIndexes.some(
        (proofIndex) =>
          [...visibleCandidatesByCombinedProof[proofIndex]].sort((left, right) => left - right).join(",") !== visibleKey,
      ) ||
      proofs.some(
        (proof) =>
          proof.contentStream.objectNumber !== stream.objectNumber ||
          proof.contentStream.generation !== stream.generation,
      )
    ) {
      continue;
    }
    const byDrawing = [...visibleIndexes].sort(
      (left, right) => visible[left].vector.operatorIndex - visible[right].vector.operatorIndex,
    );
    const byBytes = [...proofIndexes].sort(
      (left, right) =>
        combinedProjectedProofs[left].projected.proof.operatorSpan.start -
        combinedProjectedProofs[right].projected.proof.operatorSpan.start,
    );
    for (const [order, visibleIndex] of byDrawing.entries()) {
      uniqueMatches.set(visibleIndex, byBytes[order]);
    }
  }

  const ambiguousVisibleCount = proofCandidatesByVisible.filter(
    (candidates) => candidates.length > 1,
  ).length;
  const ambiguousProofCount = visibleCandidatesByProof.filter(
    (candidates) => candidates.length > 1,
  ).length;
  const ambiguousSubpathVisibleCount = subpathCandidatesByVisible.filter(
    (candidates) => candidates.length > 1,
  ).length;
  const ambiguousSubpathProofCount = visibleCandidatesBySubpathProof.filter(
    (candidates) => candidates.length > 1,
  ).length;
  const ambiguousCombinedVisibleCount = combinedCandidatesByVisible.filter(
    (candidates) => candidates.length > 1,
  ).length;
  const ambiguousCombinedProofCount = visibleCandidatesByCombinedProof.filter(
    (candidates) => candidates.length > 1,
  ).length;
  const matchedShapeProofIndexes = new Set(
    [...uniqueMatches.values()].filter(
      (proofIndex) => proofIndex < subpathOffset,
    ),
  );
  const matchedSubpathProofIndexes = new Set(
    [...uniqueMatches.values()]
      .filter((proofIndex) => proofIndex >= subpathOffset)
      .map((proofIndex) => proofIndex - subpathOffset),
  );
  const shapeMatchedCount = matchedShapeProofIndexes.size;
  const subpathMatchedCount = matchedSubpathProofIndexes.size;
  const matches: Array<
    PdfVectorCapabilityReconciliationResult["matches"][number]
  > = [];
  const discoveries = await Promise.all(
    visible.map(
      async (item, visibleIndex): Promise<PdfVectorObjectDiscovery> => {
        const proofIndex = uniqueMatches.get(visibleIndex);
        if (proofIndex !== undefined) {
          const proof = combinedProjectedProofs[proofIndex].projected.proof;
          matches.push(
            Object.freeze({
              visibleOperatorIndex: item.vector.operatorIndex,
              nativeProofId: proof.id,
              contentStreamObjectNumber: proof.contentStream.objectNumber,
              operatorStart: proof.operatorSpan.start,
              operatorEnd: proof.operatorSpan.end,
            }),
          );
          return Object.freeze({
            kind: "vector",
            pageNumber,
            bounds: item.bounds,
            locator: Object.freeze({
              contentStreamObjectNumber: proof.contentStream.objectNumber,
              contentStreamGeneration: proof.contentStream.generation,
              operatorStart: proof.operatorSpan.start,
              operatorEnd: proof.operatorSpan.end,
            }),
            proofDigest: proof.proofDigest,
            isolatedOperatorSpan: true,
            visualRecomposeAvailable: false,
          });
        }
        return Object.freeze({
          kind: "vector",
          pageNumber,
          bounds: item.bounds,
          locator: Object.freeze({
            regionDigest: await visibleRegionDigest(
              sourceSha256,
              pageNumber,
              item.vector,
              item.bounds,
            ),
          }),
          proofDigest: null,
          isolatedOperatorSpan: false,
          visualRecomposeAvailable: false,
        });
      },
    ),
  );
  matches.sort(
    (left, right) => left.visibleOperatorIndex - right.visibleOperatorIndex,
  );

  return Object.freeze({
    schemaVersion: PDF_VECTOR_CAPABILITY_RECONCILIATION_SCHEMA_VERSION,
    sourceSha256,
    pageNumber,
    discoveries: Object.freeze(discoveries),
    matches: Object.freeze(matches),
    diagnostics: Object.freeze({
      visibleCount: input.visibleVectors.length,
      nativeProofCount: input.nativeProofs.length,
      matchedCount: shapeMatchedCount,
      preserveOnlyCount: discoveries.length - shapeMatchedCount,
      ambiguousVisibleCount,
      ambiguousProofCount,
      sourceOrPageMismatchCount,
      nonEditableProofCount,
      invalidProofCount,
      unmatchedProofCount:
        projectedProofs.length - matchedShapeProofIndexes.size,
      nativeSubpathProofCount: nativeSubpathProofs.length,
      subpathMatchedCount,
      combinedMatchedCount: uniqueMatches.size,
      combinedPreserveOnlyCount: discoveries.length - uniqueMatches.size,
      ambiguousSubpathVisibleCount,
      ambiguousSubpathProofCount,
      ambiguousCombinedVisibleCount,
      ambiguousCombinedProofCount,
      subpathSourceOrPageMismatchCount,
      nonEditableSubpathProofCount,
      invalidSubpathProofCount,
      unmatchedSubpathProofCount:
        projectedSubpathProofs.length - matchedSubpathProofIndexes.size,
    }),
  });
}
