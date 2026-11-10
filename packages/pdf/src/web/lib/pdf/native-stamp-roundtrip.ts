/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  buildStampPayload,
  sanitizeStampPayload,
  type StampAnnotation,
} from "./stamps-annotation";
import {
  NATIVE_PDF_STAMP_SOURCE_KIND,
  sanitizeNativePdfStampAnnotationSource,
  type NativePdfStampAnnotation,
  type NativePdfStampAnnotationSource,
  type NativePdfStampRecord,
} from "./native-stamp-annotations";

export const MANAGED_NATIVE_STAMP_PREFIX = "[[MRNOTE_NATIVE_STAMP_V1|";
export const MANAGED_NATIVE_IMAGE_STAMP_PREFIX =
  "[[MRNOTE_NATIVE_STAMP_IMAGE_V1|";
export const IMPORTED_NATIVE_STAMP_PREFIX = "[[MRNOTE_NATIVE_STAMP_SOURCE_V1|";






export const MANAGED_NATIVE_DECORATION_PREFIX = "[[MRNOTE_NATIVE_DECORATION_V1|";
export const MAX_MANAGED_NATIVE_STAMP_PATCHES = 400;
/**
 * Keep the hidden original-package patch plan below the API's 20 MiB JSON
 * request ceiling. One uploaded image is already capped at 1 MiB decoded by
 * sanitizeSignatureImage; this aggregate encoded-character limit still allows
 * many normal stamp assets while refusing an unbounded duplicated blob.
 */
export const MAX_MANAGED_NATIVE_STAMP_IMAGE_TOTAL_CHARS = 12 * 1024 * 1024;
const MAX_MANAGED_NATIVE_STAMP_IMAGE_DIRECTIVE_CHARS = 1_400_000;
export const MAX_MANAGED_NATIVE_DECORATION_LAYOUT_CHARS = 256 * 1024;
const MAX_MANAGED_NATIVE_DECORATION_RASTER_CHARS = 1_398_104;
const MAX_MANAGED_NATIVE_DECORATION_DIRECTIVE_CHARS =
  MAX_MANAGED_NATIVE_DECORATION_LAYOUT_CHARS +
  MAX_MANAGED_NATIVE_DECORATION_RASTER_CHARS +
  512;
const BASE64_PATTERN = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/;
const MANAGED_NATIVE_DECORATION_PATTERN =
  /^\[\[MRNOTE_NATIVE_DECORATION_V1\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(auto|0|90|180|270)\|([A-Za-z0-9+/]+={0,2})\|([A-Za-z0-9+/]*={0,2})\]\]$/;

const MANAGED_NATIVE_STAMP_PATTERN =
  /^\[\[MRNOTE_NATIVE_STAMP_V1\|(upsert|delete)\|([A-Za-z0-9._~%\-]{1,384})(?:\|([a-z][a-z0-9_-]{0,39})\|(auto|0|90|180|270)\|([0-9.,-]{7,256}))?\]\]$/;
const MANAGED_NATIVE_IMAGE_STAMP_PATTERN =
  /^\[\[MRNOTE_NATIVE_STAMP_IMAGE_V1\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(png|jpeg)\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/i;
const IMPORTED_NATIVE_STAMP_PATTERN =
  /^\[\[MRNOTE_NATIVE_STAMP_SOURCE_V1\|(transform|delete)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)(?:\|(auto|0|90|180|270)\|([0-9.,-]{7,256}))?\]\]$/;

/**
 * One page's document decorations as the vector layout the server draws into
 * the page content. `layout` is base64 JSON (rendered-page points, y down);
 * `raster` is an optional base64 PNG layer for an image watermark only.
 */
export interface ManagedNativeDecorationRecord {
  id: string;
  /** One-based original PDF page number. */
  page: number;
  kind: "decoration";
  layout: string;
  raster: string;
  deleted?: false;
}

export type ManagedNativeStampRecord =
  | (StampAnnotation & { deleted?: false })
  | ManagedNativeDecorationRecord
  | {
      id: string;
      /** One-based original PDF page number. */
      page: number;
      deleted: true;
    };

export interface ManagedNativeStampAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

export interface ManagedNativeStampPlan {
  annotations: ManagedNativeStampAnnotationShadow[];
  supported: boolean;
  reason?:
    | "invalid_stamp"
    | "image_stamp_total_limit"
    | "duplicate_id"
    | "patch_limit";
}

export interface ImportedNativeStampPlan {
  annotations: ManagedNativeStampAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_stamp" | "duplicate_id" | "patch_limit";
}

export interface ParsedManagedNativeStampDirective {
  operation: "upsert" | "delete";
  id: string;
  nativeId: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  stamp?: Omit<StampAnnotation, "id" | "page">;
  decoration?: Pick<ManagedNativeDecorationRecord, "layout" | "raster">;
}

export interface ParsedImportedNativeStampDirective {
  operation: "transform" | "delete";
  objectRef: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  geometry?: Pick<NativePdfStampAnnotation, "x" | "y" | "width" | "height">;
}

function encodeOpaqueStampId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `stamp.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeOpaqueStampId(value: string): string | null {
  if (!value.startsWith("stamp.")) return null;
  try {
    const decoded = decodeURIComponent(value.slice("stamp.".length));
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

function normalizedGeometryTokens(value: {
  x: number;
  y: number;
  width: number;
  height: number;
}): string {
  return [value.x, value.y, value.width, value.height]
    .map((coordinate) => numberToken(coordinate / 100))
    .join(",");
}

function parseGeometry(
  token: string,
): Pick<NativePdfStampAnnotation, "x" | "y" | "width" | "height"> | null {
  const values = token.split(",").map(Number);
  if (
    values.length !== 4 ||
    values.some(
      (coordinate) =>
        !Number.isFinite(coordinate) || coordinate < 0 || coordinate > 1,
    )
  ) {
    return null;
  }
  const [x, y, width, height] = values;
  if (
    width < 0.000_001 ||
    height < 0.000_001 ||
    x + width > 1.000_001 ||
    y + height > 1.000_001
  ) {
    return null;
  }
  return {
    x: x * 100,
    y: y * 100,
    width: width * 100,
    height: height * 100,
  };
}

const NATIVE_STAMP_SYNC_EPSILON = 0.0001;

function nearlyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= NATIVE_STAMP_SYNC_EPSILON;
}

export function isManagedNativeStampEquivalent(
  readerStamp: StampAnnotation,
  nativeStamp: StampAnnotation,
): boolean {
  const contentMatches =
    readerStamp.kind === nativeStamp.kind &&
    (readerStamp.kind === "preset"
      ? readerStamp.preset === nativeStamp.preset
      : readerStamp.image === nativeStamp.image);
  return (
    readerStamp.id === nativeStamp.id &&
    readerStamp.page === nativeStamp.page &&
    contentMatches &&
    nearlyEqual(readerStamp.x, nativeStamp.x) &&
    nearlyEqual(readerStamp.y, nativeStamp.y) &&
    nearlyEqual(readerStamp.width, nativeStamp.width) &&
    nearlyEqual(readerStamp.height, nativeStamp.height)
  );
}

function managedStampDirective(
  record: ManagedNativeStampRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const nativeId = encodeOpaqueStampId(record.id);
  if (!nativeId || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return `${MANAGED_NATIVE_STAMP_PREFIX}delete|${nativeId}]]`;
  }
  if (record.kind === "decoration") {
    if (
      !isBoundedBase64(record.layout, MAX_MANAGED_NATIVE_DECORATION_LAYOUT_CHARS) ||
      record.layout.length === 0 ||
      !isBoundedBase64(record.raster, MAX_MANAGED_NATIVE_DECORATION_RASTER_CHARS)
    ) {
      return null;
    }
    return (
      `${MANAGED_NATIVE_DECORATION_PREFIX}upsert|${nativeId}|${rotation}|` +
      `${record.layout}|${record.raster}]]`
    );
  }
  const stamp = sanitizeStampPayload(
    record.id,
    record as unknown as Record<string, unknown>,
  );
  if (!stamp) return null;
  if (stamp.kind === "preset") {
    return (
      `${MANAGED_NATIVE_STAMP_PREFIX}upsert|${nativeId}|${stamp.preset}|` +
      `${rotation}|${normalizedGeometryTokens(stamp)}]]`
    );
  }
  const imageMatch =
    /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/i.exec(stamp.image);
  if (!imageMatch) return null;
  const mime = imageMatch[1].toLowerCase();
  const payload = imageMatch[2];
  return (
    `${MANAGED_NATIVE_IMAGE_STAMP_PREFIX}upsert|${nativeId}|${mime}|` +
    `${payload}|${rotation}|${normalizedGeometryTokens(stamp)}]]`
  );
}

function isBoundedBase64(value: unknown, maxLength: number): value is string {
  return (
    typeof value === "string" &&
    value.length <= maxLength &&
    value.length % 4 === 0 &&
    BASE64_PATTERN.test(value)
  );
}

function importedStampDirective(
  record: NativePdfStampRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const source = sanitizeNativePdfStampAnnotationSource(record.nativeSource);
  if (!source || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return `${IMPORTED_NATIVE_STAMP_PREFIX}delete|${source.objectRef}]]`;
  }
  if (record.sourceAction !== "transform") return null;
  return (
    `${IMPORTED_NATIVE_STAMP_PREFIX}transform|${source.objectRef}|${rotation}|` +
    `${normalizedGeometryTokens(record)}]]`
  );
}

/**
 * Build invisible page annotation rows consumed by the editable-page PDF
 * bridge. Presets and strict PNG/JPEG image stamps both become original-package
 * `/Stamp` annotations; malformed or aggregate-oversized image plans refuse.
 */
export function planManagedNativeStampAnnotations(
  records: readonly ManagedNativeStampRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeStampPlan {
  if (records.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const annotations: ManagedNativeStampAnnotationShadow[] = [];
  let imageDirectiveCharacters = 0;
  for (const record of records) {
    if (ids.has(record.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    const text = managedStampDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      return { annotations: [], supported: false, reason: "invalid_stamp" };
    }
    if (
      text.startsWith(MANAGED_NATIVE_IMAGE_STAMP_PREFIX) ||
      text.startsWith(MANAGED_NATIVE_DECORATION_PREFIX)
    ) {
      imageDirectiveCharacters += text.length;
      if (
        imageDirectiveCharacters > MAX_MANAGED_NATIVE_STAMP_IMAGE_TOTAL_CHARS
      ) {
        return {
          annotations: [],
          supported: false,
          reason: "image_stamp_total_limit",
        };
      }
    }
    annotations.push({
      id: `__mrnote_native_stamp__${record.id}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_STAMP_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

export function planImportedNativeStampAnnotations(
  records: readonly NativePdfStampRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ImportedNativeStampPlan {
  if (records.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const objectRefs = new Set<string>();
  const annotations: ManagedNativeStampAnnotationShadow[] = [];
  for (const record of records) {
    const source = sanitizeNativePdfStampAnnotationSource(record.nativeSource);
    if (!source) {
      return { annotations: [], supported: false, reason: "invalid_stamp" };
    }
    if (ids.has(record.id) || objectRefs.has(source.objectRef)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    objectRefs.add(source.objectRef);
    const text = importedStampDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      if (!record.deleted && record.sourceAction === undefined) continue;
      return { annotations: [], supported: false, reason: "invalid_stamp" };
    }
    annotations.push({
      id: `__mrnote_native_stamp_source__${source.objectRef}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_STAMP_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

export function isManagedNativeStampAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    (value.startsWith(MANAGED_NATIVE_STAMP_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_IMAGE_STAMP_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_DECORATION_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_STAMP_PREFIX))
  );
}

export function parseManagedNativeStampDirective(
  value: unknown,
): ParsedManagedNativeStampDirective | null {
  if (typeof value !== "string") return null;
  if (value.startsWith(MANAGED_NATIVE_DECORATION_PREFIX)) {
    if (value.length > MAX_MANAGED_NATIVE_DECORATION_DIRECTIVE_CHARS) {
      return null;
    }
    const decorationMatch = MANAGED_NATIVE_DECORATION_PATTERN.exec(value);
    if (!decorationMatch) return null;
    const nativeId = decorationMatch[1];
    const id = decodeOpaqueStampId(nativeId);
    const layout = decorationMatch[3];
    const raster = decorationMatch[4];
    if (
      !id ||
      !isBoundedBase64(layout, MAX_MANAGED_NATIVE_DECORATION_LAYOUT_CHARS) ||
      !isBoundedBase64(raster, MAX_MANAGED_NATIVE_DECORATION_RASTER_CHARS)
    ) {
      return null;
    }
    return {
      operation: "upsert",
      id,
      nativeId,
      rotation:
        decorationMatch[2] === "auto"
          ? "auto"
          : (Number(decorationMatch[2]) as 0 | 90 | 180 | 270),
      decoration: { layout, raster },
    };
  }
  if (value.startsWith(MANAGED_NATIVE_IMAGE_STAMP_PREFIX)) {
    if (value.length > MAX_MANAGED_NATIVE_STAMP_IMAGE_DIRECTIVE_CHARS) {
      return null;
    }
    const imageMatch = MANAGED_NATIVE_IMAGE_STAMP_PATTERN.exec(value);
    if (!imageMatch) return null;
    const nativeId = imageMatch[1];
    const id = decodeOpaqueStampId(nativeId);
    const geometry = parseGeometry(imageMatch[5]);
    if (!id || !geometry) return null;
    const mime = imageMatch[2].toLowerCase();
    const stamp = buildStampPayload({
      page: 1,
      ...geometry,
      kind: "image",
      image: `data:image/${mime};base64,${imageMatch[3]}`,
    });
    if (!stamp) return null;
    return {
      operation: "upsert",
      id,
      nativeId,
      rotation:
        imageMatch[4] === "auto"
          ? "auto"
          : (Number(imageMatch[4]) as 0 | 90 | 180 | 270),
      stamp: {
        x: stamp.x,
        y: stamp.y,
        width: stamp.width,
        height: stamp.height,
        kind: stamp.kind,
        preset: stamp.preset,
        image: stamp.image,
      },
    };
  }
  const match = MANAGED_NATIVE_STAMP_PATTERN.exec(value);
  if (!match) return null;
  const operation = match[1] as "upsert" | "delete";
  const nativeId = match[2];
  const id = decodeOpaqueStampId(nativeId);
  if (!id) return null;
  if (operation === "delete") {
    if (match[3] !== undefined) return null;
    return { operation, id, nativeId };
  }
  if (!match[3] || !match[4] || !match[5]) return null;
  const geometry = parseGeometry(match[5]);
  if (!geometry) return null;
  const stamp = buildStampPayload({
    page: 1,
    ...geometry,
    kind: "preset",
    preset: match[3],
  });
  if (!stamp) return null;
  const rotation =
    match[4] === "auto" ? "auto" : (Number(match[4]) as 0 | 90 | 180 | 270);
  return {
    operation,
    id,
    nativeId,
    rotation,
    stamp: {
      x: stamp.x,
      y: stamp.y,
      width: stamp.width,
      height: stamp.height,
      kind: stamp.kind,
      preset: stamp.preset,
      image: stamp.image,
    },
  };
}

export function parseImportedNativeStampDirective(
  value: unknown,
): ParsedImportedNativeStampDirective | null {
  if (typeof value !== "string") return null;
  const match = IMPORTED_NATIVE_STAMP_PATTERN.exec(value);
  if (!match) return null;
  const operation = match[1] as "transform" | "delete";
  const objectRef = match[2];
  if (operation === "delete") {
    if (match[3] !== undefined || match[4] !== undefined) return null;
    return { operation, objectRef };
  }
  if (!match[3] || !match[4]) return null;
  const geometry = parseGeometry(match[4]);
  if (!geometry) return null;
  return {
    operation,
    objectRef,
    rotation:
      match[3] === "auto" ? "auto" : (Number(match[3]) as 0 | 90 | 180 | 270),
    geometry,
  };
}

export function parseManagedNativeStampShadow(
  value: unknown,
): ManagedNativeStampRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseManagedNativeStampDirective(row.text);
  if (!parsed) return null;
  if (parsed.operation === "delete") {
    return { id: parsed.id, page, deleted: true };
  }
  if (parsed.decoration) {
    return { id: parsed.id, page, kind: "decoration", ...parsed.decoration };
  }
  if (!parsed.stamp) return null;
  return { id: parsed.id, page, ...parsed.stamp };
}

export function parseImportedNativeStampShadow(
  value: unknown,
): NativePdfStampRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseImportedNativeStampDirective(row.text);
  if (!parsed) return null;
  const nativeSource: NativePdfStampAnnotationSource = {
    kind: NATIVE_PDF_STAMP_SOURCE_KIND,
    objectRef: parsed.objectRef,
    subtype: "Stamp",
  };
  const id = `pdf-native-stamp:${page}:${parsed.objectRef}`;
  if (parsed.operation === "delete") {
    return { id, page, deleted: true, nativeSource };
  }
  if (!parsed.geometry) return null;
  return {
    id,
    page,
    ...parsed.geometry,
    nativeSource,
    sourceAction: "transform",
  };
}
