/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MAX_SHAPE_COMMENT,
  normalizePolygonPoints,
  normalizePolylinePoints,
  resolveShapeCornerRadius,
  resolveShapeFill,
  resolveShapeLineEndings,
  resolveShapeLineStyle,
  resolveShapeOpacity,
  sanitizeShapePayload,
  shapeKindSupportsLineEndings,
  type ShapeAnnotation,
  type ShapeKind,
  type ShapeLineEnding,
  type ShapeLineStyle,
} from "@mrnote/core/web/lib/pdf/shapes.ts";
import {
  NATIVE_PDF_SHAPE_SOURCE_KIND,
  nativePdfShapeRecordIsEdited,
  sanitizeNativePdfShapeAnnotationSource,
  type NativePdfShapeAnnotation,
  type NativePdfShapeRecord,
  type NativePdfShapeSourceStyle,
  type NativePdfShapeSubtype,
} from "./native-shape-annotations";

export const MANAGED_NATIVE_SHAPE_PREFIX = "[[MRNOTE_NATIVE_SHAPE_V1|";
export const MANAGED_NATIVE_SHAPE_V2_PREFIX = "[[MRNOTE_NATIVE_SHAPE_V2|";
export const MANAGED_NATIVE_SHAPE_V3_PREFIX = "[[MRNOTE_NATIVE_SHAPE_V3|";
export const IMPORTED_NATIVE_SHAPE_PREFIX = "[[MRNOTE_NATIVE_SHAPE_SOURCE_V1|";
/** A source drawing restyled (and maybe moved): see NativePdfShapeSourceStyle
 * and _IMPORTED_SHAPE_STYLE_RE in native_edit_primitives.py. */
export const IMPORTED_NATIVE_SHAPE_V2_PREFIX = "[[MRNOTE_NATIVE_SHAPE_SOURCE_V2|";
export const MAX_MANAGED_NATIVE_SHAPE_PATCHES = 400;

const MANAGED_NATIVE_SHAPE_PATTERN =
  /^\[\[MRNOTE_NATIVE_SHAPE_V1\|(upsert|delete)\|([A-Za-z0-9._~%\-]{1,384})(?:\|(rect|line|arrow|ellipse|cloud|polygon|polyline)\|([0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(none|[0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(auto|0|90|180|270)\|([0-9.,-]{7,16384}))?\]\]$/;
const MANAGED_NATIVE_SHAPE_V2_PATTERN =
  /^\[\[MRNOTE_NATIVE_SHAPE_V2\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(rect|line|arrow|ellipse|cloud|polygon|polyline)\|([0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(none|[0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(solid|dashed|dotted)\|(none|openArrow|closedArrow|circle|square)\|(none|openArrow|closedArrow|circle|square)\|(auto|0|90|180|270)\|([0-9.,-]{7,16384})\]\]$/;
const MANAGED_NATIVE_SHAPE_V3_PATTERN =
  /^\[\[MRNOTE_NATIVE_SHAPE_V3\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(rect|line|arrow|ellipse|cloud|polygon|polyline)\|([0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(none|[0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(solid|dashed|dotted)\|(none|openArrow|closedArrow|circle|square)\|(none|openArrow|closedArrow|circle|square)\|([0-9.]{1,20})\|([0-9.]{1,20})\|(auto|0|90|180|270)\|([0-9.,-]{7,16384})\]\]$/;
const IMPORTED_NATIVE_SHAPE_PATTERN =
  /^\[\[MRNOTE_NATIVE_SHAPE_SOURCE_V1\|(transform|delete)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(Square|Circle|Line|Polygon|PolyLine)(?:\|(auto|0|90|180|270)\|([0-9.,-]{7,256}))?\]\]$/;
const IMPORTED_NATIVE_SHAPE_V2_PATTERN =
  /^\[\[MRNOTE_NATIVE_SHAPE_SOURCE_V2\|style\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(Square|Circle|Line|Polygon|PolyLine)\|(auto|0|90|180|270)\|(move|keep)\|([0-9.,-]{7,256})\|([0-9A-Fa-f]{6})\|([0-9.]{1,20})\|(solid|dashed|dotted)\|(keep|none|[0-9A-Fa-f]{6})\|([0-9.]{1,20})\]\]$/;

export type ManagedNativeShapeRecord =
  | (ShapeAnnotation & { deleted?: false })
  | {
      id: string;
      /** One-based original PDF page number. */
      page: number;
      deleted: true;
    };

export interface ManagedNativeShapeAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

export interface ManagedNativeShapePlan {
  annotations: ManagedNativeShapeAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_shape" | "duplicate_id" | "patch_limit";
}

export interface ImportedNativeShapePlan {
  annotations: ManagedNativeShapeAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_shape" | "duplicate_id" | "patch_limit";
}

export interface ParsedManagedNativeShapeDirective {
  operation: "upsert" | "delete";
  id: string;
  nativeId: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  shape?: Omit<ShapeAnnotation, "id" | "page">;
}

export interface ParsedImportedNativeShapeDirective {
  operation: "transform" | "delete" | "style";
  objectRef: string;
  subtype: NativePdfShapeSubtype;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  geometry?: Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2">;
  /** A restyle that also moves the drawing to `geometry`. */
  moved?: boolean;
  style?: NativePdfShapeSourceStyle;
}

/** Bounds of a restyle the server takes (native_edit_primitives.py). */
const SOURCE_STYLE_MIN_WIDTH = 0.25;
const SOURCE_STYLE_MAX_WIDTH = 72;

function sanitizeNativePdfShapeSourceStyle(
  style: NativePdfShapeSourceStyle,
): NativePdfShapeSourceStyle | null {
  const hex = /^#[0-9A-Fa-f]{6}$/;
  if (
    !hex.test(style.color) ||
    (typeof style.fillColor === "string" && !hex.test(style.fillColor)) ||
    !Number.isFinite(style.widthPoints) ||
    style.widthPoints < SOURCE_STYLE_MIN_WIDTH ||
    style.widthPoints > SOURCE_STYLE_MAX_WIDTH ||
    !Number.isFinite(style.opacity) ||
    style.opacity < 0 ||
    style.opacity > 1 ||
    !["solid", "dashed", "dotted"].includes(style.lineStyle)
  ) {
    return null;
  }
  return {
    color: style.color.toUpperCase(),
    widthPoints: Number(style.widthPoints.toFixed(3)),
    lineStyle: style.lineStyle,
    ...(style.fillColor === undefined
      ? {}
      : { fillColor: style.fillColor === null ? null : style.fillColor.toUpperCase() }),
    opacity: Number(style.opacity.toFixed(3)),
  };
}

function encodeOpaqueShapeId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `shape.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeOpaqueShapeId(value: string): string | null {
  if (!value.startsWith("shape.")) return null;
  try {
    const decoded = decodeURIComponent(value.slice("shape.".length));
    return decoded.trim() ? decoded : null;
  } catch {
    return null;
  }
}

function numberToken(value: number): string {
  return Number(value.toFixed(6)).toString();
}

function isValidPage(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 10_000
  );
}

const NATIVE_SHAPE_SYNC_EPSILON = 0.0001;

function nearlyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= NATIVE_SHAPE_SYNC_EPSILON;
}

/**
 * Confirm that a reader annotation row and its native shadow describe the same
 * visible object. Matching only the id is unsafe: the annotation API can accept
 * a move/delete while original-package autosave fails, leaving an older native
 * object under that same id.
 */
export function isManagedNativeShapeEquivalent(
  readerShape: ShapeAnnotation,
  nativeShape: ShapeAnnotation,
): boolean {
  if (
    readerShape.id !== nativeShape.id ||
    readerShape.page !== nativeShape.page ||
    readerShape.kind !== nativeShape.kind ||
    readerShape.color.toLowerCase() !== nativeShape.color.toLowerCase() ||
    !nearlyEqual(readerShape.width, nativeShape.width) ||
    !nearlyEqual(readerShape.x1, nativeShape.x1) ||
    !nearlyEqual(readerShape.y1, nativeShape.y1) ||
    !nearlyEqual(readerShape.x2, nativeShape.x2) ||
    !nearlyEqual(readerShape.y2, nativeShape.y2)
  ) {
    return false;
  }
  const readerEndings = resolveShapeLineEndings(readerShape);
  const nativeEndings = resolveShapeLineEndings(nativeShape);
  if (
    resolveShapeLineStyle(readerShape) !== resolveShapeLineStyle(nativeShape) ||
    readerEndings.start !== nativeEndings.start ||
    readerEndings.end !== nativeEndings.end ||
    !nearlyEqual(
      resolveShapeOpacity(readerShape.opacity),
      resolveShapeOpacity(nativeShape.opacity),
    ) ||
    !nearlyEqual(
      resolveShapeCornerRadius(readerShape.kind, readerShape.cornerRadius),
      resolveShapeCornerRadius(nativeShape.kind, nativeShape.cornerRadius),
    )
  ) {
    return false;
  }
  if (readerShape.kind === "polygon" || readerShape.kind === "polyline") {
    const normalize =
      readerShape.kind === "polygon"
        ? normalizePolygonPoints
        : normalizePolylinePoints;
    const readerPoints = normalize(readerShape.points);
    const nativePoints = normalize(nativeShape.points);
    if (
      !readerPoints ||
      !nativePoints ||
      readerPoints.length !== nativePoints.length ||
      readerPoints.some(
        (point, index) =>
          !nearlyEqual(point.x, nativePoints[index].x) ||
          !nearlyEqual(point.y, nativePoints[index].y),
      )
    ) {
      return false;
    }
  }
  const readerFill = resolveShapeFill(readerShape);
  const nativeFill = resolveShapeFill(nativeShape);
  if (!readerFill || !nativeFill) return readerFill === nativeFill;
  return (
    readerFill.color.toLowerCase() === nativeFill.color.toLowerCase() &&
    nearlyEqual(readerFill.opacity, nativeFill.opacity)
  );
}

function nativeShapeDirective(
  record: ManagedNativeShapeRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const directive = nativeShapeDirectiveHead(record, rotation);
  if (!directive || record.deleted) return directive;
  
  const comment = (record as { comment?: unknown }).comment;
  return typeof comment === "string" && comment.trim()
    ? `${directive}\n${comment.slice(0, MAX_SHAPE_COMMENT)}`
    : directive;
}

function nativeShapeDirectiveHead(
  record: ManagedNativeShapeRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const nativeId = encodeOpaqueShapeId(record.id);
  if (!nativeId || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return `${MANAGED_NATIVE_SHAPE_PREFIX}delete|${nativeId}]]`;
  }
  const shape = sanitizeShapePayload(
    record.id,
    record as unknown as Record<string, unknown>,
  );
  if (!shape) return null;
  const fill = resolveShapeFill(shape);
  const pathKind = shape.kind === "polygon" || shape.kind === "polyline";
  const points = (
    pathKind && shape.points
      ? shape.points.flatMap((point) => [point.x, point.y])
      : [shape.x1, shape.y1, shape.x2, shape.y2]
  )
    .map((coordinate) => numberToken(coordinate / 100))
    .join(",");
  const usesV2 =
    shape.lineStyle !== undefined ||
    shape.lineStart !== undefined ||
    shape.lineEnd !== undefined;
  const usesV3 =
    shape.opacity !== undefined || shape.cornerRadius !== undefined;
  if (usesV3) {
    const endings = resolveShapeLineEndings(shape);
    return (
      `${MANAGED_NATIVE_SHAPE_V3_PREFIX}upsert|${nativeId}|${shape.kind}|` +
      `${shape.color.slice(1).toUpperCase()}|${numberToken(shape.width)}|` +
      `${fill ? fill.color.slice(1).toUpperCase() : "none"}|` +
      `${numberToken(fill?.opacity ?? 0)}|${resolveShapeLineStyle(shape)}|` +
      `${endings.start}|${endings.end}|` +
      `${numberToken(resolveShapeOpacity(shape.opacity))}|` +
      `${numberToken(resolveShapeCornerRadius(shape.kind, shape.cornerRadius))}|` +
      `${rotation}|${points}]]`
    );
  }
  if (usesV2) {
    const endings = resolveShapeLineEndings(shape);
    return (
      `${MANAGED_NATIVE_SHAPE_V2_PREFIX}upsert|${nativeId}|${shape.kind}|` +
      `${shape.color.slice(1).toUpperCase()}|${numberToken(shape.width)}|` +
      `${fill ? fill.color.slice(1).toUpperCase() : "none"}|` +
      `${numberToken(fill?.opacity ?? 0)}|${resolveShapeLineStyle(shape)}|` +
      `${endings.start}|${endings.end}|${rotation}|${points}]]`
    );
  }
  return (
    `${MANAGED_NATIVE_SHAPE_PREFIX}upsert|${nativeId}|${shape.kind}|` +
    `${shape.color.slice(1).toUpperCase()}|${numberToken(shape.width)}|` +
    `${fill ? fill.color.slice(1).toUpperCase() : "none"}|` +
    `${numberToken(fill?.opacity ?? 0)}|${rotation}|${points}]]`
  );
}

function importedNativeShapeDirective(
  record: NativePdfShapeRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const source = sanitizeNativePdfShapeAnnotationSource(record.nativeSource);
  if (!source || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return (
      `${IMPORTED_NATIVE_SHAPE_PREFIX}delete|${source.objectRef}|` +
      `${source.subtype}]]`
    );
  }
  // An untouched source object already exists in the immutable input package.
  if (record.sourceAction !== "transform" && record.sourceStyle === undefined) {
    return null;
  }
  const shape = sanitizeShapePayload(
    record.id,
    record as unknown as Record<string, unknown>,
  );
  if (!shape) return null;
  const geometry = [shape.x1, shape.y1, shape.x2, shape.y2]
    .map((coordinate) => numberToken(coordinate / 100))
    .join(",");
  if (record.sourceStyle !== undefined) {
    const style = sanitizeNativePdfShapeSourceStyle(record.sourceStyle);
    if (!style) return null;
    return (
      `${IMPORTED_NATIVE_SHAPE_V2_PREFIX}style|${source.objectRef}|` +
      `${source.subtype}|${rotation}|` +
      `${record.sourceAction === "transform" ? "move" : "keep"}|${geometry}|` +
      `${style.color.slice(1)}|${numberToken(style.widthPoints)}|${style.lineStyle}|` +
      `${style.fillColor === undefined ? "keep" : style.fillColor === null ? "none" : style.fillColor.slice(1)}|` +
      `${numberToken(style.opacity)}]]`
    );
  }
  return (
    `${IMPORTED_NATIVE_SHAPE_PREFIX}transform|${source.objectRef}|` +
    `${source.subtype}|${rotation}|${geometry}]]`
  );
}

/**
 * Build invisible annotation rows consumed by the existing editable-page PDF
 * bridge. The plan is deliberately all-or-nothing: malformed geometry,
 * duplicate ids, or an oversized batch yields no native directives.
 */
export function planManagedNativeShapeAnnotations(
  records: readonly ManagedNativeShapeRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeShapePlan {
  if (records.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const annotations: ManagedNativeShapeAnnotationShadow[] = [];
  for (const record of records) {
    if (ids.has(record.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    const text = nativeShapeDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      return { annotations: [], supported: false, reason: "invalid_shape" };
    }
    annotations.push({
      id: `__mrnote_native_shape__${record.id}`,
      text,
      page: record.page,
      // The native primitive derives its exact /Rect from the shape geometry.
      // This bounded placeholder only satisfies the shared annotation transport.
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_SHAPE_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

/**
 * Build original-object transform/delete rows for third-party drawing markups.
 * Untouched source objects intentionally emit no directive. Invalid input
 * fails the whole plan closed so a save never applies only a prefix.
 */
export function planImportedNativeShapeAnnotations(
  records: readonly NativePdfShapeRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ImportedNativeShapePlan {
  if (records.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const objectRefs = new Set<string>();
  const annotations: ManagedNativeShapeAnnotationShadow[] = [];
  for (const record of records) {
    const source = sanitizeNativePdfShapeAnnotationSource(record.nativeSource);
    if (!source) {
      return { annotations: [], supported: false, reason: "invalid_shape" };
    }
    if (ids.has(record.id) || objectRefs.has(source.objectRef)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    objectRefs.add(source.objectRef);
    const text = importedNativeShapeDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      if (!nativePdfShapeRecordIsEdited(record)) continue;
      return { annotations: [], supported: false, reason: "invalid_shape" };
    }
    annotations.push({
      id: `__mrnote_native_shape_source__${source.objectRef}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_SHAPE_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

export function isManagedNativeShapeAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    (value.startsWith(MANAGED_NATIVE_SHAPE_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_SHAPE_V2_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_SHAPE_V3_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_SHAPE_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_SHAPE_V2_PREFIX))
  );
}

export function parseImportedNativeShapeDirective(
  value: unknown,
): ParsedImportedNativeShapeDirective | null {
  if (typeof value !== "string") return null;
  const styled = IMPORTED_NATIVE_SHAPE_V2_PATTERN.exec(value);
  if (styled) {
    const points = styled[5].split(",").map(Number);
    const style = sanitizeNativePdfShapeSourceStyle({
      color: `#${styled[6]}`,
      widthPoints: Number(styled[7]),
      lineStyle: styled[8] as ShapeLineStyle,
      ...(styled[9] === "keep"
        ? {}
        : { fillColor: styled[9] === "none" ? null : `#${styled[9]}` }),
      opacity: Number(styled[10]),
    });
    if (
      !style ||
      points.length !== 4 ||
      points.some(
        (coordinate) =>
          !Number.isFinite(coordinate) || coordinate < 0 || coordinate > 1,
      ) ||
      points[2] - points[0] < 0.000_001 ||
      points[3] - points[1] < 0.000_001
    ) {
      return null;
    }
    return {
      operation: "style",
      objectRef: styled[1],
      subtype: styled[2] as NativePdfShapeSubtype,
      rotation:
        styled[3] === "auto" ? "auto" : (Number(styled[3]) as 0 | 90 | 180 | 270),
      geometry: {
        x1: points[0] * 100,
        y1: points[1] * 100,
        x2: points[2] * 100,
        y2: points[3] * 100,
      },
      moved: styled[4] === "move",
      style,
    };
  }
  const match = IMPORTED_NATIVE_SHAPE_PATTERN.exec(value);
  if (!match) return null;
  const operation = match[1] as "transform" | "delete";
  const objectRef = match[2];
  const subtype = match[3] as NativePdfShapeSubtype;
  if (operation === "delete") {
    if (match[4] !== undefined || match[5] !== undefined) return null;
    return { operation, objectRef, subtype };
  }
  if (!match[4] || !match[5]) return null;
  const rotation =
    match[4] === "auto" ? "auto" : (Number(match[4]) as 0 | 90 | 180 | 270);
  const points = match[5].split(",").map(Number);
  if (
    points.length !== 4 ||
    points.some(
      (coordinate) =>
        !Number.isFinite(coordinate) || coordinate < 0 || coordinate > 1,
    ) ||
    points[2] - points[0] < 0.000_001 ||
    points[3] - points[1] < 0.000_001
  ) {
    return null;
  }
  return {
    operation,
    objectRef,
    subtype,
    rotation,
    geometry: {
      x1: points[0] * 100,
      y1: points[1] * 100,
      x2: points[2] * 100,
      y2: points[3] * 100,
    },
  };
}

export function parseManagedNativeShapeDirective(
  input: unknown,
): ParsedManagedNativeShapeDirective | null {
  if (typeof input !== "string") return null;
  
  const directiveEnd = input.indexOf("]]");
  const hasComment =
    directiveEnd >= 0 && input.slice(directiveEnd + 2).startsWith("\n");
  const value = hasComment ? input.slice(0, directiveEnd + 2) : input;
  const comment = hasComment ? input.slice(directiveEnd + 3) : "";
  const parsed = parseManagedNativeShapeDirectiveHead(value);
  if (!parsed || !comment.trim() || !parsed.shape) return parsed;
  return {
    ...parsed,
    shape: { ...parsed.shape, comment: comment.slice(0, MAX_SHAPE_COMMENT) },
  };
}

function parseManagedNativeShapeDirectiveHead(
  value: string,
): ParsedManagedNativeShapeDirective | null {
  const v3Match = MANAGED_NATIVE_SHAPE_V3_PATTERN.exec(value);
  const v2Match = MANAGED_NATIVE_SHAPE_V2_PATTERN.exec(value);
  const v1Match = MANAGED_NATIVE_SHAPE_PATTERN.exec(value);
  if (!v3Match && !v2Match && !v1Match) return null;
  const operation =
    v3Match || v2Match
      ? ("upsert" as const)
      : (v1Match![1] as "upsert" | "delete");
  const nativeId = v3Match?.[1] ?? v2Match?.[1] ?? v1Match![2];
  const id = decodeOpaqueShapeId(nativeId);
  if (!id) return null;
  if (operation === "delete") {
    if (v1Match![3] !== undefined) return null;
    return { operation, id, nativeId };
  }
  const kind = (v3Match?.[2] ?? v2Match?.[2] ?? v1Match?.[3]) as
    ShapeKind | undefined;
  const colorToken = v3Match?.[3] ?? v2Match?.[3] ?? v1Match?.[4];
  const widthToken = v3Match?.[4] ?? v2Match?.[4] ?? v1Match?.[5];
  const fillToken = v3Match?.[5] ?? v2Match?.[5] ?? v1Match?.[6];
  const fillOpacityToken = v3Match?.[6] ?? v2Match?.[6] ?? v1Match?.[7];
  const lineStyle = (v3Match?.[7] ?? v2Match?.[7] ?? "solid") as ShapeLineStyle;
  const lineStart = (v3Match?.[8] ?? v2Match?.[8] ?? "none") as ShapeLineEnding;
  const lineEnd = (v3Match?.[9] ??
    v2Match?.[9] ??
    (kind === "arrow" ? "openArrow" : "none")) as ShapeLineEnding;
  const opacityToken = v3Match?.[10] ?? "1";
  const cornerRadiusToken = v3Match?.[11] ?? "0";
  const rotationToken = v3Match?.[12] ?? v2Match?.[10] ?? v1Match?.[8];
  const pointsToken = v3Match?.[13] ?? v2Match?.[11] ?? v1Match?.[9];
  if (
    !kind ||
    !colorToken ||
    !widthToken ||
    !fillToken ||
    !fillOpacityToken ||
    !rotationToken ||
    !pointsToken
  ) {
    return null;
  }
  if (
    !shapeKindSupportsLineEndings(kind) &&
    (lineStart !== "none" || lineEnd !== "none")
  ) {
    return null;
  }
  const width = Number(widthToken);
  const fillOpacity = Number(fillOpacityToken);
  const opacity = Number(opacityToken);
  const cornerRadius = Number(cornerRadiusToken);
  const rotation =
    rotationToken === "auto"
      ? "auto"
      : (Number(rotationToken) as 0 | 90 | 180 | 270);
  const points = pointsToken.split(",").map(Number);
  if (
    !Number.isFinite(width) ||
    width < 0.1 ||
    width > 12 ||
    !Number.isFinite(fillOpacity) ||
    fillOpacity < 0 ||
    fillOpacity > 1 ||
    !Number.isFinite(opacity) ||
    opacity < 0 ||
    opacity > 1 ||
    !Number.isFinite(cornerRadius) ||
    cornerRadius < 0 ||
    cornerRadius > 50 ||
    (kind !== "rect" && cornerRadius !== 0) ||
    (kind === "polygon" || kind === "polyline"
      ? points.length < (kind === "polygon" ? 6 : 4) ||
        points.length > 400 ||
        points.length % 2 !== 0
      : points.length !== 4) ||
    points.some(
      (coordinate) =>
        !Number.isFinite(coordinate) || coordinate < 0 || coordinate > 1,
    )
  ) {
    return null;
  }
  const pathPoints =
    kind === "polygon" || kind === "polyline"
      ? Array.from({ length: points.length / 2 }, (_, index) => ({
          x: points[index * 2] * 100,
          y: points[index * 2 + 1] * 100,
        }))
      : undefined;
  const pointXs = pathPoints?.map((point) => point.x);
  const pointYs = pathPoints?.map((point) => point.y);
  const fillColor = fillToken === "none" ? null : `#${fillToken.toUpperCase()}`;
  const shape = sanitizeShapePayload(id, {
    page: 1,
    kind,
    x1: pointXs ? Math.min(...pointXs) : points[0] * 100,
    y1: pointYs ? Math.min(...pointYs) : points[1] * 100,
    x2: pointXs ? Math.max(...pointXs) : points[2] * 100,
    y2: pointYs ? Math.max(...pointYs) : points[3] * 100,
    ...(pathPoints ? { points: pathPoints } : null),
    color: `#${colorToken.toUpperCase()}`,
    width,
    fillColor,
    fillOpacity,
    ...(v3Match || v2Match ? { lineStyle, lineStart, lineEnd } : null),
    ...(v3Match ? { opacity, cornerRadius } : null),
  });
  if (!shape) return null;
  const shapeWithoutIdentity: Omit<ShapeAnnotation, "id" | "page"> = {
    kind: shape.kind,
    x1: shape.x1,
    y1: shape.y1,
    x2: shape.x2,
    y2: shape.y2,
    color: shape.color,
    width: shape.width,
    fillColor: shape.fillColor,
    fillOpacity: shape.fillOpacity,
    ...(v3Match || v2Match
      ? { lineStyle, lineStart, lineEnd }
      : kind === "arrow"
        ? { lineEnd: "openArrow" as const }
        : null),
    ...(shape.opacity !== undefined ? { opacity: shape.opacity } : null),
    ...(shape.cornerRadius !== undefined
      ? { cornerRadius: shape.cornerRadius }
      : null),
    ...(shape.points ? { points: shape.points } : null),
  };
  return {
    operation,
    id,
    nativeId,
    rotation,
    shape: shapeWithoutIdentity,
  };
}

export function parseManagedNativeShapeShadow(
  value: unknown,
): ManagedNativeShapeRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseManagedNativeShapeDirective(row.text ?? row.comment);
  if (!parsed) return null;
  if (parsed.operation === "delete") {
    return { id: parsed.id, page, deleted: true };
  }
  if (!parsed.shape) return null;
  return { id: parsed.id, page, ...parsed.shape };
}

export function parseImportedNativeShapeShadow(
  value: unknown,
): NativePdfShapeRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseImportedNativeShapeDirective(row.text ?? row.comment);
  if (!parsed) return null;
  const nativeSource = {
    kind: NATIVE_PDF_SHAPE_SOURCE_KIND,
    objectRef: parsed.objectRef,
    subtype: parsed.subtype,
  } as const;
  const id = `pdf-native-shape:${page}:${parsed.objectRef}`;
  if (parsed.operation === "delete") {
    return { id, page, deleted: true, nativeSource };
  }
  if (!parsed.geometry) return null;
  const kind: ShapeKind =
    parsed.subtype === "Square" || parsed.subtype === "Polygon"
      ? "rect"
      : parsed.subtype === "Circle"
        ? "ellipse"
        : "line";
  const style = parsed.style;
  const record: NativePdfShapeAnnotation = {
    id,
    page,
    kind,
    ...parsed.geometry,
    color: style?.color ?? "#111827",
    width: 0.6,
    nativeSource,
    ...(parsed.operation === "transform" || parsed.moved
      ? { sourceAction: "transform" as const }
      : {}),
    ...(style
      ? {
          sourceStyle: style,
          lineWidthPoints: style.widthPoints,
          ...(style.lineStyle !== "solid" ? { lineStyle: style.lineStyle } : {}),
          ...(style.fillColor ? { sourceFillColor: style.fillColor } : {}),
          ...(style.opacity !== 1 ? { opacity: style.opacity } : {}),
        }
      : {}),
  };
  return record;
}
