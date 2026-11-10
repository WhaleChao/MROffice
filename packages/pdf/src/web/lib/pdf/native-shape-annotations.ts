/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  DEFAULT_SHAPE_COLOR,
  DEFAULT_SHAPE_WIDTH,
  MAX_SHAPE_COMMENT,
  clampShapeWidth,
  type ShapeAnnotation,
  type ShapeKind,
  type ShapeLineEnding,
  type ShapeLineStyle,
} from "@mrnote/core/web/lib/pdf/shapes.ts";
import {
  readNativePdfAnnotationFlags,
  sanitizePdfText,
  type NativePdfAnnotationFlags,
} from "@mrnote/core/web/lib/pdf/native-text-annotations.ts";
import type {
  NativePdfShapeSourceDetails,
  NativePdfShapeSourceDetailsIndex,
} from "./native-shape-source-details";

/**
 * Editable third-party drawing markups discovered through pdf.js.
 *
 * The source page canvas remains the visual authority. These records provide a
 * transparent selection target plus the exact indirect-object identity needed
 * to transform/delete the original annotation without regenerating `/AP`.
 */

export const NATIVE_PDF_SHAPE_SOURCE_KIND = "pdf_shape_annotation";
export const MAX_NATIVE_PDF_SHAPE_ANNOTATIONS_PER_PAGE = 1_024;

const PDF_OBJECT_REF_RE = /^[1-9][0-9]{0,9}R(?:[0-9]{1,5})?$/;
const COORDINATE_EPSILON = 0.000_001;

export type NativePdfShapeSubtype =
  | "Square"
  | "Circle"
  | "Line"
  | "Polygon"
  | "PolyLine";

export interface NativePdfShapeAnnotationSource {
  kind: typeof NATIVE_PDF_SHAPE_SOURCE_KIND;
  objectRef: string;
  subtype: NativePdfShapeSubtype;
  /** The source's /F flags; absent on records written before they were read. */
  flags?: NativePdfAnnotationFlags;
  




  group?: { ref: string; members: string[] };
}

export interface NativePdfShapeAnnotation extends ShapeAnnotation {
  nativeSource: NativePdfShapeAnnotationSource;
  




  lineWidthPoints?: number;
  






  sourceLineInBox?: Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2">;
   
  sourceFillColor?: string;
  

  text?: string;
  author?: string;
  




  sourceFigure?: {
    kind: "polygon" | "polyline";
    points: { x: number; y: number }[];
  };
  /**
   * Absent for an untouched source object. A transformed record is persisted
   * as an original-object patch directive until the patched preview is ready.
   */
  sourceAction?: "transform";
  /**
   * The appearance this drawing is restyled to (Acrobat's and WPS's
   * properties): written to its /C, /IC, /BS and /CA with a new /AP, the
   * rest of the annotation kept. Absent while its own style stands.
   */
  sourceStyle?: NativePdfShapeSourceStyle;
}

export interface NativePdfShapeSourceStyle {
  /** The outline colour (/C), `#RRGGBB`. */
  color: string;
  /** The border width (/BS /W) in points. */
  widthPoints: number;
  lineStyle: ShapeLineStyle;
  /** The interior colour (/IC), `#RRGGBB`; null for no fill; absent keeps
   * the file's own (the page may not have told the reader its /IC). */
  fillColor?: string | null;
  /** The annotation's opacity (/CA), 0..1. */
  opacity: number;
}

/** The style a drawing stands in now: the one it was restyled to, else the
 * one read from the file (a border the writer cannot draw thinner than a
 * quarter point is read as that). */
export function nativePdfShapeSourceStyleOf(
  record: NativePdfShapeAnnotation,
): NativePdfShapeSourceStyle {
  if (record.sourceStyle) return record.sourceStyle;
  return {
    color: record.color.toUpperCase(),
    widthPoints: Math.min(72, Math.max(0.25, record.lineWidthPoints ?? 1)),
    lineStyle: record.lineStyle ?? "solid",
    ...(record.sourceFillColor ? { fillColor: record.sourceFillColor.toUpperCase() } : {}),
    opacity: record.opacity ?? 1,
  };
}

/** A drawing whose source object an edit names: kept over a fresh reading. */
export function nativePdfShapeRecordIsEdited(record: NativePdfShapeRecord): boolean {
  return (
    record.deleted === true ||
    record.sourceAction === "transform" ||
    record.sourceStyle !== undefined
  );
}

export interface NativePdfShapeTombstone {
  id: string;
  /** One-based original PDF page number. */
  page: number;
  deleted: true;
  nativeSource: NativePdfShapeAnnotationSource;
}

export type NativePdfShapeRecord =
  (NativePdfShapeAnnotation & { deleted?: false }) | NativePdfShapeTombstone;

export type NativePdfShapeImportLossReason =
  | "invalid_object_ref"
  | "duplicate_object_ref"
  | "invalid_page_geometry"
  | "invalid_shape_geometry"
  | "unsupported_opacity"
  | "unsupported_border_style"
  | "unsupported_line_ending"
  | "annotation_limit"
  

  | "unsupported_group";

export interface NativePdfShapeImportLoss {
  objectRef?: string;
  subtype?: NativePdfShapeSubtype;
  reason: NativePdfShapeImportLossReason;
}

export interface NativePdfShapeImportResult {
  shapes: NativePdfShapeAnnotation[];
  losses: NativePdfShapeImportLoss[];
}

export interface NativePdfShapeViewport {
  width: number;
  height: number;
  





  pageBox?: { width: number; height: number };
  /** pdf.js `page.userUnit` (PDF 1.6 /UserUnit); 1 when absent. */
  userUnit?: number;
  /** The viewport's rotation in degrees (pdf.js `viewport.rotation`). */
  rotation?: number;
  /** pdf.js `viewport.convertToViewportPoint`, for NoRotate anchoring. */
  convertToViewportPoint?: (x: number, y: number) => readonly number[];
  convertToViewportRectangle: (
    rect: readonly number[],
  ) => readonly [number, number, number, number];
}

interface RawNativePdfShapeAnnotation {
  subtype?: unknown;
  id?: unknown;
  /** pdf.js `annotationFlags`: the annotation's /F. */
  annotationFlags?: unknown;
  rect?: unknown;
  /** pdf.js Line `lineCoordinates`: /L normalized into a min/max box. */
  lineCoordinates?: unknown;
  /** pdf.js Polygon/PolyLine `vertices`: /Vertices as x, y pairs. */
  vertices?: unknown;
  /** pdf.js `inReplyTo` (the /IRT ref) and `replyType` ("Group" or "R"). */
  inReplyTo?: unknown;
  replyType?: unknown;
  /** The drawing's /Contents comment and /T author, as pdf.js reports them. */
  contentsObj?: unknown;
  titleObj?: unknown;
  color?: unknown;
  opacity?: unknown;
  lineEndings?: unknown;
  borderStyle?: {
    width?: unknown;
    rawWidth?: unknown;
    style?: unknown;
    dashArray?: unknown;
    horizontalCornerRadius?: unknown;
    verticalCornerRadius?: unknown;
  };
}

function isPdfObjectRef(value: unknown): value is string {
  return typeof value === "string" && PDF_OBJECT_REF_RE.test(value);
}

function isNativePdfShapeSubtype(
  value: unknown,
): value is NativePdfShapeSubtype {
  return (
    value === "Square" ||
    value === "Circle" ||
    value === "Line" ||
    value === "Polygon" ||
    value === "PolyLine"
  );
}

function shapeKindForSubtype(subtype: NativePdfShapeSubtype): ShapeKind {
  if (subtype === "Square") return "rect";
  if (subtype === "Circle") return "ellipse";
  if (subtype === "Polygon") return "rect";
  return "line";
}

function finiteNumber(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function numericRect(value: unknown): [number, number, number, number] | null {
  if (!Array.isArray(value) && !ArrayBuffer.isView(value)) return null;
  const values = Array.from(value as ArrayLike<unknown>).map(finiteNumber);
  if (values.length !== 4 || values.some((item) => item === null)) return null;
  return values as [number, number, number, number];
}








function colorComponents(values: readonly unknown[]): number[] | null {
  const numbers = values.map(finiteNumber);
  if (numbers.some((number) => number === null || number < 0 || number > 255)) {
    return null;
  }
  const finite = numbers as number[];
  if (finite.every(Number.isInteger)) return finite;
  if (finite.some((number) => number > 1)) {
    return finite.map((number) => Math.round(number));
  }
  return finite.map((number) => Math.round(number * 255));
}

function nativeShapeColor(value: unknown): string {
  if (!Array.isArray(value) && !ArrayBuffer.isView(value)) {
    return DEFAULT_SHAPE_COLOR;
  }
  const components = Array.from(value as ArrayLike<unknown>);
  if (components.length === 1) {
    const gray = colorComponents(components)?.[0] ?? null;
    if (gray === null) return DEFAULT_SHAPE_COLOR;
    const token = gray.toString(16).padStart(2, "0");
    return `#${token}${token}${token}`.toUpperCase();
  }
  if (components.length < 3) return DEFAULT_SHAPE_COLOR;
  const rgb = colorComponents(components.slice(0, 3));
  if (!rgb) return DEFAULT_SHAPE_COLOR;
  return `#${rgb
    .map((component) => component.toString(16).padStart(2, "0"))
    .join("")}`.toUpperCase();
}

function nativeShapeLineStyle(
  annotation: RawNativePdfShapeAnnotation,
): ShapeLineStyle | null {
  const border = annotation.borderStyle;
  if (!border) return "solid";
  if (border.style === undefined || border.style === 1) return "solid";
  if (border.style !== 2) return null;
  if (
    !Array.isArray(border.dashArray) &&
    !ArrayBuffer.isView(border.dashArray)
  ) {
    return null;
  }
  const dash = Array.from(border.dashArray as ArrayLike<unknown>).map(
    finiteNumber,
  );
  if (
    dash.length < 1 ||
    dash.length > 16 ||
    dash.some((value) => value === null || value <= 0 || value > 1_000)
  ) {
    return null;
  }
  const sourceWidth = Math.max(
    0.000_001,
    finiteNumber(border.width) ?? finiteNumber(border.rawWidth) ?? 1,
  );
  return (dash[0] as number) <= 1.5 * sourceWidth ? "dotted" : "dashed";
}

function nativeShapeOpacity(
  annotation: RawNativePdfShapeAnnotation,
): number | null {
  if (annotation.opacity == null) return 1;
  const opacity = finiteNumber(annotation.opacity);
  return opacity !== null && opacity >= 0 && opacity <= 1 ? opacity : null;
}

function nativeShapeCornerRadius(
  annotation: RawNativePdfShapeAnnotation,
  sourceRect: [number, number, number, number],
): number | null {
  const horizontalRaw = annotation.borderStyle?.horizontalCornerRadius;
  const verticalRaw = annotation.borderStyle?.verticalCornerRadius;
  const horizontal = horizontalRaw == null ? 0 : finiteNumber(horizontalRaw);
  const vertical = verticalRaw == null ? 0 : finiteNumber(verticalRaw);
  if (
    horizontal === null ||
    vertical === null ||
    horizontal < 0 ||
    vertical < 0
  ) {
    return null;
  }
  if (horizontal === 0 && vertical === 0) return 0;
  if (annotation.subtype !== "Square") return null;
  const shortSide = Math.min(
    Math.abs(sourceRect[2] - sourceRect[0]),
    Math.abs(sourceRect[3] - sourceRect[1]),
  );
  if (!Number.isFinite(shortSide) || shortSide <= COORDINATE_EPSILON)
    return null;
  const radiusEpsilon = Math.max(
    COORDINATE_EPSILON,
    shortSide * COORDINATE_EPSILON,
  );
  if (Math.abs(horizontal - vertical) > radiusEpsilon) return null;
  const cornerRadius = ((horizontal + vertical) / 2 / shortSide) * 100;
  return Number.isFinite(cornerRadius) && cornerRadius <= 50
    ? cornerRadius
    : null;
}

const NATIVE_LINE_ENDINGS: Readonly<Record<string, ShapeLineEnding>> = {
  None: "none",
  OpenArrow: "openArrow",
  ClosedArrow: "closedArrow",
  Circle: "circle",
  Square: "square",
};

function nativeShapeLineEndings(
  annotation: RawNativePdfShapeAnnotation,
): { start: ShapeLineEnding; end: ShapeLineEnding } | null {
  if (
    (annotation.subtype !== "Line" && annotation.subtype !== "PolyLine") ||
    annotation.lineEndings == null
  ) {
    return { start: "none", end: "none" };
  }
  if (
    !Array.isArray(annotation.lineEndings) &&
    !ArrayBuffer.isView(annotation.lineEndings)
  ) {
    return null;
  }
  const endings = Array.from(annotation.lineEndings as ArrayLike<unknown>);
  if (endings.length !== 2) return null;
  const normalized = endings.map((ending) =>
    ending == null ? "none" : NATIVE_LINE_ENDINGS[String(ending)],
  );
  return normalized.every(Boolean)
    ? {
        start: normalized[0] as ShapeLineEnding,
        end: normalized[1] as ShapeLineEnding,
      }
    : null;
}

function nativeShapeSourceWidth(
  annotation: RawNativePdfShapeAnnotation,
): number | null {
  const sourceWidth =
    finiteNumber(annotation.borderStyle?.width) ??
    finiteNumber(annotation.borderStyle?.rawWidth);
  return sourceWidth === null ? null : Math.max(0, sourceWidth);
}

function nativeShapePageMinUserUnits(
  viewport: NativePdfShapeViewport,
): number {
  const box = viewport.pageBox;
  if (
    box &&
    Number.isFinite(box.width) &&
    Number.isFinite(box.height) &&
    box.width > 0 &&
    box.height > 0
  ) {
    return Math.min(box.width, box.height);
  }
  return Math.max(1, Math.min(viewport.width, viewport.height));
}

function nativeShapeWidth(
  annotation: RawNativePdfShapeAnnotation,
  viewport: NativePdfShapeViewport,
): number {
  const sourceWidth = nativeShapeSourceWidth(annotation) ?? DEFAULT_SHAPE_WIDTH;
  // The same page-relative unit the managed-shape writer uses: a share of
  // the page box's shorter side in user units, never of the zoomed viewport.
  return clampShapeWidth(
    (sourceWidth / Math.max(1, nativeShapePageMinUserUnits(viewport))) * 100,
  );
}

function nativeShapeUserUnit(viewport: NativePdfShapeViewport): number {
  const unit = viewport.userUnit;
  return typeof unit === "number" && Number.isFinite(unit) && unit > 0
    ? unit
    : 1;
}









function noRotateViewportRect(
  sourceRect: [number, number, number, number],
  viewport: NativePdfShapeViewport,
): Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2"> | null {
  const rotation = ((Math.round(viewport.rotation ?? 0) % 360) + 360) % 360;
  if (rotation === 0 || typeof viewport.convertToViewportPoint !== "function") {
    return null;
  }
  const [x1, y1, x2, y2] = sourceRect;
  const left = Math.min(x1, x2);
  const top = Math.max(y1, y2);
  const width = Math.abs(x2 - x1);
  const height = Math.abs(y2 - y1);
  let anchor: readonly number[];
  let unit: readonly number[];
  try {
    anchor = viewport.convertToViewportPoint(left, top);
    unit = viewport.convertToViewportPoint(left + 1, top);
  } catch {
    return null;
  }
  const [ax, ay] = anchor;
  const scale = Math.hypot(unit[0] - ax, unit[1] - ay);
  if (![ax, ay, scale].every(Number.isFinite) || scale <= 0) return null;
  const rect = numericRect([ax, ay, ax + width * scale, ay + height * scale]);
  if (!rect) return null;
  const l = rect[0] / viewport.width;
  const t = rect[1] / viewport.height;
  const r = rect[2] / viewport.width;
  const b = rect[3] / viewport.height;
  if (r - l <= COORDINATE_EPSILON || b - t <= COORDINATE_EPSILON) return null;
  return {
    x1: Math.min(1, Math.max(0, l)) * 100,
    y1: Math.min(1, Math.max(0, t)) * 100,
    x2: Math.min(1, Math.max(0, r)) * 100,
    y2: Math.min(1, Math.max(0, b)) * 100,
  };
}

function normalizeViewportRect(
  sourceRect: [number, number, number, number],
  viewport: NativePdfShapeViewport,
): Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2"> | null {
  let converted: readonly number[];
  try {
    converted = viewport.convertToViewportRectangle(sourceRect);
  } catch {
    return null;
  }
  const rect = numericRect(converted);
  if (!rect) return null;
  const left = Math.min(rect[0], rect[2]) / viewport.width;
  const right = Math.max(rect[0], rect[2]) / viewport.width;
  const top = Math.min(rect[1], rect[3]) / viewport.height;
  const bottom = Math.max(rect[1], rect[3]) / viewport.height;
  if (
    left < -COORDINATE_EPSILON ||
    right > 1 + COORDINATE_EPSILON ||
    top < -COORDINATE_EPSILON ||
    bottom > 1 + COORDINATE_EPSILON ||
    right - left <= COORDINATE_EPSILON ||
    bottom - top <= COORDINATE_EPSILON
  ) {
    return null;
  }
  return {
    x1: Math.min(1, Math.max(0, left)) * 100,
    y1: Math.min(1, Math.max(0, top)) * 100,
    x2: Math.min(1, Math.max(0, right)) * 100,
    y2: Math.min(1, Math.max(0, bottom)) * 100,
  };
}

function viewportPointPercent(
  x: number,
  y: number,
  viewport: NativePdfShapeViewport,
): { x: number; y: number } | null {
  let point: readonly number[];
  try {
    point =
      typeof viewport.convertToViewportPoint === "function"
        ? viewport.convertToViewportPoint(x, y)
        : viewport.convertToViewportRectangle([x, y, x, y]).slice(0, 2);
  } catch {
    return null;
  }
  const px = finiteNumber(point?.[0]);
  const py = finiteNumber(point?.[1]);
  if (px === null || py === null) return null;
  const fx = px / viewport.width;
  const fy = py / viewport.height;
  if (
    fx < -COORDINATE_EPSILON ||
    fx > 1 + COORDINATE_EPSILON ||
    fy < -COORDINATE_EPSILON ||
    fy > 1 + COORDINATE_EPSILON
  ) {
    return null;
  }
  return {
    x: Math.min(1, Math.max(0, fx)) * 100,
    y: Math.min(1, Math.max(0, fy)) * 100,
  };
}

/**
 * A Line's real segment, start to end, in page percent. The source's own /L
 * gives the direction (which end carries which /LE); pdf.js only reports it
 * normalized into a min/max box, which is still the right length and axis
 * when the raw /L could not be read. The raw /L is trusted only when its box
 * is the one pdf.js read for this object.
 */
function nativeShapeSourceLine(
  annotation: RawNativePdfShapeAnnotation,
  details: NativePdfShapeSourceDetails | undefined,
  viewport: NativePdfShapeViewport,
): Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2"> | undefined {
  if (annotation.subtype !== "Line") return undefined;
  const normalized = numericRect(annotation.lineCoordinates);
  let line: readonly number[] | null = details?.line ?? null;
  if (line && normalized) {
    const box = [
      Math.min(line[0], line[2]),
      Math.min(line[1], line[3]),
      Math.max(line[0], line[2]),
      Math.max(line[1], line[3]),
    ];
    if (box.some((value, index) => Math.abs(value - normalized[index]) > 0.01)) {
      line = null;
    }
  }
  const endpoints = line ?? normalized;
  if (!endpoints) return undefined;
  const start = viewportPointPercent(endpoints[0], endpoints[1], viewport);
  const end = viewportPointPercent(endpoints[2], endpoints[3], viewport);
  if (!start || !end) return undefined;
  if (
    Math.abs(start.x - end.x) <= COORDINATE_EPSILON &&
    Math.abs(start.y - end.y) <= COORDINATE_EPSILON
  ) {
    return undefined;
  }
  return { x1: start.x, y1: start.y, x2: end.x, y2: end.y };
}



function nativeShapeSourceFigure(
  annotation: RawNativePdfShapeAnnotation,
  viewport: NativePdfShapeViewport,
): { kind: "polygon" | "polyline"; points: { x: number; y: number }[] } | undefined {
  if (annotation.subtype !== "Polygon" && annotation.subtype !== "PolyLine") {
    return undefined;
  }
  const raw = annotation.vertices;
  if (!Array.isArray(raw) && !ArrayBuffer.isView(raw)) return undefined;
  const values = Array.from(raw as ArrayLike<unknown>).map(Number);
  if (values.length < 4 || values.length % 2 !== 0 || values.length > 2_000) {
    return undefined;
  }
  const points: { x: number; y: number }[] = [];
  for (let index = 0; index < values.length; index += 2) {
    const point = viewportPointPercent(values[index], values[index + 1], viewport);
    if (!point) return undefined;
    points.push({ x: point.x, y: point.y });
  }
  const kind = annotation.subtype === "Polygon" ? "polygon" : "polyline";
  return points.length >= (kind === "polygon" ? 3 : 2) ? { kind, points } : undefined;
}

function nativeShapeSourceFillColor(
  annotation: RawNativePdfShapeAnnotation,
  details: NativePdfShapeSourceDetails | undefined,
): string | undefined {
  if (
    !details?.interiorColor ||
    (annotation.subtype !== "Square" &&
      annotation.subtype !== "Circle" &&
      annotation.subtype !== "Polygon")
  ) {
    return undefined;
  }
  const bytes = details.interiorColor.map((component) =>
    Math.round(Math.min(1, Math.max(0, component)) * 255),
  );
  return `#${bytes
    .map((component) => component.toString(16).padStart(2, "0"))
    .join("")}`.toUpperCase();
}

export function sanitizeNativePdfShapeAnnotationSource(
  value: unknown,
): NativePdfShapeAnnotationSource | undefined {
  if (!value || typeof value !== "object") return undefined;
  const source = value as Record<string, unknown>;
  const objectRef = source.objectRef ?? source.object_ref;
  if (
    source.kind !== NATIVE_PDF_SHAPE_SOURCE_KIND ||
    !isPdfObjectRef(objectRef) ||
    !isNativePdfShapeSubtype(source.subtype)
  ) {
    return undefined;
  }
  const flags = sanitizeNativePdfShapeFlags(source.flags);
  const group = sanitizeNativePdfShapeGroup(source.group, objectRef);
  return {
    kind: NATIVE_PDF_SHAPE_SOURCE_KIND,
    objectRef,
    subtype: source.subtype,
    ...(flags ? { flags } : {}),
    ...(group ? { group } : {}),
  };
}

/** Members a source group may have and still be handled as one entity. */
export const MAX_NATIVE_PDF_SHAPE_GROUP_MEMBERS = 64;
const MAX_NATIVE_PDF_SHAPE_GROUP_DEPTH = 32;

function sanitizeNativePdfShapeGroup(
  value: unknown,
  objectRef: string,
): { ref: string; members: string[] } | undefined {
  if (!value || typeof value !== "object") return undefined;
  const raw = value as Record<string, unknown>;
  if (!isPdfObjectRef(raw.ref) || !Array.isArray(raw.members)) return undefined;
  const members = raw.members.filter(isPdfObjectRef);
  if (
    members.length !== raw.members.length ||
    members.length < 2 ||
    members.length > MAX_NATIVE_PDF_SHAPE_GROUP_MEMBERS ||
    new Set(members).size !== members.length ||
    !members.includes(objectRef) ||
    !members.includes(raw.ref)
  ) {
    return undefined;
  }
  return { ref: raw.ref, members: [...members].sort() };
}

/**
 * The page's annotation groups (ISO 32000-1 12.5.6.2): each primary's ref
 * with its members' (itself included), and the refs whose group cannot be
 * trusted -- a reference that leaves the page, loops or runs too deep.
 */
export function nativePdfAnnotationGroups(rawAnnotations: readonly unknown[]): {
  groups: Map<string, string[]>;
  broken: Set<string>;
} {
  const parents = new Map<string, string | null>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfShapeAnnotation;
    if (!isPdfObjectRef(annotation.id)) continue;
    parents.set(
      annotation.id,
      annotation.replyType === "Group" && isPdfObjectRef(annotation.inReplyTo)
        ? annotation.inReplyTo
        : annotation.replyType === "Group" && annotation.inReplyTo != null
          ? ""
          : null,
    );
  }
  const heads = new Map<string, string | null>();
  const chains = new Map<string, string[]>();
  for (const id of parents.keys()) {
    const path: string[] = [];
    let current: string | null = id;
    let head: string | null = null;
    while (current !== null) {
      if (path.includes(current) || path.length > MAX_NATIVE_PDF_SHAPE_GROUP_DEPTH) break;
      path.push(current);
      const parent = parents.get(current);
      if (parent === undefined || parent === "") break;
      if (parent === null) {
        head = current;
        break;
      }
      current = parent;
    }
    heads.set(id, head);
    chains.set(id, path);
  }
  const groups = new Map<string, string[]>();
  for (const [id, head] of heads) {
    if (head === null) continue;
    groups.set(head, [...(groups.get(head) ?? []), id]);
  }
  for (const [head, members] of groups) {
    if (members.length < 2) groups.delete(head);
  }
  const broken = new Set<string>();
  for (const [id, head] of heads) {
    if (head !== null) continue;
    broken.add(id);
    // A broken chain that runs through a group spoils that group.
    for (const step of chains.get(id) ?? []) {
      const stepHead = heads.get(step);
      if (stepHead) for (const member of groups.get(stepHead) ?? []) broken.add(member);
    }
  }
  for (const members of groups.values()) {
    if (members.length > MAX_NATIVE_PDF_SHAPE_GROUP_MEMBERS) {
      for (const member of members) broken.add(member);
    }
  }
  return { groups, broken };
}

function sanitizeNativePdfShapeFlags(
  value: unknown,
): NativePdfAnnotationFlags | undefined {
  if (!value || typeof value !== "object") return undefined;
  const raw = value as Record<string, unknown>;
  const flag = (key: string) => raw[key] === true;
  return {
    hidden: flag("hidden"),
    noView: flag("noView"),
    readOnly: flag("readOnly"),
    locked: flag("locked"),
    lockedContents: flag("lockedContents"),
    noZoom: flag("noZoom"),
    noRotate: flag("noRotate"),
  };
}









export function nativeShapeDisplayGeometry<T extends ShapeAnnotation>(
  shape: T,
  scale: number,
): T {
  const source = (shape as { nativeSource?: NativePdfShapeAnnotationSource })
    .nativeSource;
  if (
    !source?.flags?.noZoom ||
    !Number.isFinite(scale) ||
    scale <= 0 ||
    Math.abs(scale - 1) < 1e-6
  ) {
    return shape;
  }
  const factor = 1 / scale;
  const points = shape.points;
  const xs = [shape.x1, shape.x2, ...(points?.map((point) => point.x) ?? [])];
  const ys = [shape.y1, shape.y2, ...(points?.map((point) => point.y) ?? [])];
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const px = (x: number) => left + (x - left) * factor;
  const py = (y: number) => top + (y - top) * factor;
  return {
    ...shape,
    x1: px(shape.x1),
    y1: py(shape.y1),
    x2: px(shape.x2),
    y2: py(shape.y2),
    ...(points
      ? { points: points.map((point) => ({ ...point, x: px(point.x), y: py(point.y) })) }
      : {}),
  };
}






export function nativeShapeSourceLineGeometry(shape: {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  sourceLineInBox?: Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2">;
}): Pick<ShapeAnnotation, "x1" | "y1" | "x2" | "y2"> | null {
  const line = shape.sourceLineInBox;
  if (!line) return null;
  const left = Math.min(shape.x1, shape.x2);
  const top = Math.min(shape.y1, shape.y2);
  const width = Math.abs(shape.x2 - shape.x1);
  const height = Math.abs(shape.y2 - shape.y1);
  const geometry = {
    x1: left + line.x1 * width,
    y1: top + line.y1 * height,
    x2: left + line.x2 * width,
    y2: top + line.y2 * height,
  };
  return Object.values(geometry).every(Number.isFinite) ? geometry : null;
}

export function isNativePdfShapeRecord(value: {
  nativeSource?: unknown;
}): value is NativePdfShapeRecord {
  return Boolean(sanitizeNativePdfShapeAnnotationSource(value.nativeSource));
}

/**
 * Normalize standard drawing markup rectangles into rotation-aware page
 * percentages. Unsupported annotations stay painted by pdf.js and are
 * returned as explicit losses.
 */
export function normalizeNativePdfShapeAnnotations(
  rawAnnotations: unknown,
  pageNumber: number,
  viewport: NativePdfShapeViewport,
  /** Raw /L and /IC by object ref, read from the same document's bytes. */
  sourceDetails?: NativePdfShapeSourceDetailsIndex | null,
): NativePdfShapeImportResult {
  const shapes: NativePdfShapeAnnotation[] = [];
  const losses: NativePdfShapeImportLoss[] = [];
  if (
    !Array.isArray(rawAnnotations) ||
    !Number.isInteger(pageNumber) ||
    pageNumber < 1 ||
    !Number.isFinite(viewport.width) ||
    !Number.isFinite(viewport.height) ||
    viewport.width <= 0 ||
    viewport.height <= 0
  ) {
    return { shapes, losses: [{ reason: "invalid_page_geometry" }] };
  }

  const objectRefCounts = new Map<string, number>();
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfShapeAnnotation;
    if (
      isNativePdfShapeSubtype(annotation.subtype) &&
      isPdfObjectRef(annotation.id)
    ) {
      objectRefCounts.set(
        annotation.id,
        (objectRefCounts.get(annotation.id) ?? 0) + 1,
      );
    }
  }
  const reportedDuplicateRefs = new Set<string>();
  let shapeCount = 0;
  for (const raw of rawAnnotations) {
    if (!raw || typeof raw !== "object") continue;
    const annotation = raw as RawNativePdfShapeAnnotation;
    if (!isNativePdfShapeSubtype(annotation.subtype)) continue;
    shapeCount += 1;
    const objectRef = isPdfObjectRef(annotation.id) ? annotation.id : undefined;
    const lossIdentity = {
      objectRef,
      subtype: annotation.subtype,
    };
    if (shapeCount > MAX_NATIVE_PDF_SHAPE_ANNOTATIONS_PER_PAGE) {
      losses.push({ ...lossIdentity, reason: "annotation_limit" });
      continue;
    }
    if (!objectRef) {
      losses.push({
        subtype: annotation.subtype,
        reason: "invalid_object_ref",
      });
      continue;
    }
    if ((objectRefCounts.get(objectRef) ?? 0) > 1) {
      if (!reportedDuplicateRefs.has(objectRef)) {
        losses.push({
          ...lossIdentity,
          reason: "duplicate_object_ref",
        });
        reportedDuplicateRefs.add(objectRef);
      }
      continue;
    }
    const sourceRect = numericRect(annotation.rect);
    const flags =
      annotation.annotationFlags !== undefined
        ? readNativePdfAnnotationFlags(annotation.annotationFlags)
        : undefined;
    const geometry = sourceRect
      ? (flags?.noRotate ? noRotateViewportRect(sourceRect, viewport) : null) ??
        normalizeViewportRect(sourceRect, viewport)
      : null;
    if (!geometry) {
      losses.push({ ...lossIdentity, reason: "invalid_shape_geometry" });
      continue;
    }
    const opacity = nativeShapeOpacity(annotation);
    if (opacity === null) {
      losses.push({ ...lossIdentity, reason: "unsupported_opacity" });
      continue;
    }
    const lineStyle = nativeShapeLineStyle(annotation);
    const cornerRadius = nativeShapeCornerRadius(annotation, sourceRect!);
    if (!lineStyle || cornerRadius === null) {
      losses.push({ ...lossIdentity, reason: "unsupported_border_style" });
      continue;
    }
    const lineEndings = nativeShapeLineEndings(annotation);
    if (!lineEndings) {
      losses.push({ ...lossIdentity, reason: "unsupported_line_ending" });
      continue;
    }
    const details = sourceDetails?.get(objectRef);
    // A NoRotate line on a rotated page is drawn unrotated from its corner;
    // its page-space /L is not where it shows, so no segment is claimed.
    const sourceLine = flags?.noRotate && (viewport.rotation ?? 0) % 360 !== 0
      ? undefined
      : nativeShapeSourceLine(annotation, details, viewport);
    const boxWidth = geometry.x2 - geometry.x1;
    const boxHeight = geometry.y2 - geometry.y1;
    const sourceLineInBox =
      sourceLine && boxWidth > COORDINATE_EPSILON && boxHeight > COORDINATE_EPSILON
        ? {
            x1: (sourceLine.x1 - geometry.x1) / boxWidth,
            y1: (sourceLine.y1 - geometry.y1) / boxHeight,
            x2: (sourceLine.x2 - geometry.x1) / boxWidth,
            y2: (sourceLine.y2 - geometry.y1) / boxHeight,
          }
        : undefined;
    const sourceFillColor = nativeShapeSourceFillColor(annotation, details);
    const comment = sanitizePdfText(annotation.contentsObj, MAX_SHAPE_COMMENT).trim();
    const author = sanitizePdfText(annotation.titleObj, 500).trim();
    const figure = nativeShapeSourceFigure(annotation, viewport);
    const sourceFigure =
      figure && boxWidth > COORDINATE_EPSILON && boxHeight > COORDINATE_EPSILON
        ? {
            kind: figure.kind,
            points: figure.points.map((point) => ({
              x: (point.x - geometry.x1) / boxWidth,
              y: (point.y - geometry.y1) / boxHeight,
            })),
          }
        : undefined;
    shapes.push({
      id: `pdf-native-shape:${pageNumber}:${objectRef}`,
      page: pageNumber,
      kind: shapeKindForSubtype(annotation.subtype),
      ...geometry,
      color: nativeShapeColor(annotation.color),
      width: nativeShapeWidth(annotation, viewport),
      lineWidthPoints:
        (nativeShapeSourceWidth(annotation) ?? 1) *
        nativeShapeUserUnit(viewport),
      ...(lineStyle !== "solid" ? { lineStyle } : null),
      ...(opacity !== 1 ? { opacity } : null),
      ...(cornerRadius !== 0 ? { cornerRadius } : null),
      ...(lineEndings.start !== "none"
        ? { lineStart: lineEndings.start }
        : null),
      ...(lineEndings.end !== "none" ? { lineEnd: lineEndings.end } : null),
      ...(sourceLineInBox ? { sourceLineInBox } : null),
      ...(sourceFillColor ? { sourceFillColor } : null),
      ...(sourceFigure ? { sourceFigure } : null),
      ...(comment ? { text: comment } : null),
      ...(author ? { author } : null),
      nativeSource: {
        kind: NATIVE_PDF_SHAPE_SOURCE_KIND,
        objectRef,
        subtype: annotation.subtype,
        ...(annotation.annotationFlags !== undefined
          ? { flags: readNativePdfAnnotationFlags(annotation.annotationFlags) }
          : {}),
      },
    });
  }
  // A grouped drawing is editable only with its whole group: every member
  // a drawing read here, its references sound. Otherwise each stays painted
  // by the page and is reported, rather than one member moved alone
  
  const { groups, broken } = nativePdfAnnotationGroups(rawAnnotations);
  if (groups.size === 0 && broken.size === 0) return { shapes, losses };
  const groupOf = new Map<string, string>();
  for (const [head, members] of groups) {
    for (const member of members) groupOf.set(member, head);
  }
  const editable = new Set(shapes.map((shape) => shape.nativeSource.objectRef));
  const kept: NativePdfShapeAnnotation[] = [];
  for (const shape of shapes) {
    const ref = shape.nativeSource.objectRef;
    const head = groupOf.get(ref);
    if (broken.has(ref)) {
      losses.push({ objectRef: ref, subtype: shape.nativeSource.subtype, reason: "unsupported_group" });
      continue;
    }
    if (!head) {
      kept.push(shape);
      continue;
    }
    const members = groups.get(head) ?? [];
    if (!members.every((member) => editable.has(member) && !broken.has(member))) {
      losses.push({ objectRef: ref, subtype: shape.nativeSource.subtype, reason: "unsupported_group" });
      continue;
    }
    kept.push({
      ...shape,
      nativeSource: { ...shape.nativeSource, group: { ref: head, members: [...members].sort() } },
    });
  }
  return { shapes: kept, losses };
}

/** `record` (its box, identity and edit) drawn as `drawn` reads; the same
 * object when nothing about its look differs. */
function withDrawnLook(
  record: NativePdfShapeAnnotation,
  drawn: NativePdfShapeAnnotation,
): NativePdfShapeAnnotation {
  const keys = [
    "color",
    "width",
    "lineWidthPoints",
    "lineStyle",
    "opacity",
    "cornerRadius",
    "lineStart",
    "lineEnd",
    "sourceFillColor",
  ] as const;
  if (keys.every((key) => record[key] === drawn[key])) return record;
  const next: NativePdfShapeAnnotation = { ...record };
  for (const key of keys) {
    if (drawn[key] === undefined) delete next[key];
    else (next as unknown as Record<string, unknown>)[key] = drawn[key];
  }
  return next;
}

/**
 * Merge one freshly-discovered page into host state without overwriting
 * pending transforms or deletion tombstones that still target the immutable
 * source object's reference.
 */
export function mergeNativePdfShapeAnnotations(
  current: readonly NativePdfShapeRecord[],
  incoming: readonly NativePdfShapeAnnotation[],
  pageNumber: number,
): NativePdfShapeRecord[] {
  const currentPage = current.filter((record) => record.page === pageNumber);
  const protectedByRef = new Map(
    currentPage
      .filter(nativePdfShapeRecordIsEdited)
      .map((record) => [record.nativeSource.objectRef, record]),
  );
  // An edited record read back from its directive alone knows its box but
  // not its look (a moved drawing parsed again after a reload carried a
  // placeholder colour and width): the patched page's own reading of the
  // same object says how it is drawn now.
  const looks = new Map(
    incoming.map((discovered) => [discovered.nativeSource.objectRef, discovered]),
  );
  const nextPage: NativePdfShapeRecord[] = [...protectedByRef.values()].map(
    (record) => {
      const drawn = looks.get(record.nativeSource.objectRef);
      return record.deleted || !drawn || record.sourceStyle !== undefined
        ? record
        : withDrawnLook(record, drawn);
    },
  );
  for (const discovered of incoming) {
    if (protectedByRef.has(discovered.nativeSource.objectRef)) continue;
    // pikepdf normally preserves source object numbers. If a save rewrites an
    // indirect ref, join the patched preview back to the pending transform by
    // exact subtype + geometry rather than exposing a duplicate hit target.
    const matchesProtectedTransform = [...protectedByRef.values()].some(
      (record) =>
        !record.deleted &&
        nativePdfShapeRecordIsEdited(record) &&
        record.nativeSource.subtype === discovered.nativeSource.subtype &&
        Math.abs(record.x1 - discovered.x1) <= 0.0001 &&
        Math.abs(record.y1 - discovered.y1) <= 0.0001 &&
        Math.abs(record.x2 - discovered.x2) <= 0.0001 &&
        Math.abs(record.y2 - discovered.y2) <= 0.0001,
    );
    if (matchesProtectedTransform) continue;
    nextPage.push(discovered);
  }

  const unchanged =
    currentPage.length === nextPage.length &&
    nextPage.every((record, index) => record === currentPage[index]);
  if (unchanged) return current as NativePdfShapeRecord[];

  const merged: NativePdfShapeRecord[] = [];
  let insertedPage = false;
  for (const record of current) {
    if (record.page === pageNumber) {
      if (!insertedPage) {
        merged.push(...nextPage);
        insertedPage = true;
      }
      continue;
    }
    merged.push(record);
  }
  if (!insertedPage) merged.push(...nextPage);
  return merged;
}
