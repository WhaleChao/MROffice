/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { buildLinkOverlayPayload, type PdfLinkOverlay } from "@mrnote/core/web/lib/pdf/links.ts";
import {
  NATIVE_PDF_LINK_SOURCE_KIND,
  sanitizeNativePdfLinkAnnotationSource,
  type NativePdfLinkAnnotation,
  type NativePdfLinkRecord,
} from "./native-link-annotations";

export const MANAGED_NATIVE_LINK_PREFIX = "[[MRNOTE_NATIVE_LINK_V1|";
export const IMPORTED_NATIVE_LINK_PREFIX = "[[MRNOTE_NATIVE_LINK_SOURCE_V1|";
export const MAX_MANAGED_NATIVE_LINK_PATCHES = 400;
const MAX_MANAGED_NATIVE_LINK_URL_TOKEN_CHARS = 8_192;
const MAX_MANAGED_NATIVE_LINK_DIRECTIVE_CHARS =
  MAX_MANAGED_NATIVE_LINK_URL_TOKEN_CHARS + 1_024;

const UPSERT_URI_PATTERN =
  /^\[\[MRNOTE_NATIVE_LINK_V1\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|uri\|([A-Za-z0-9_-]{1,8192})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/u;
const UPSERT_PAGE_PATTERN =
  /^\[\[MRNOTE_NATIVE_LINK_V1\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|page\|([1-9][0-9]{0,4})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/u;
const DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_LINK_V1\|delete\|([A-Za-z0-9._~%\-]{1,384})\]\]$/u;
// `geometry` moves the source box only: the server keeps /A, /AA, /Dest and

const IMPORTED_URI_PATTERN =
  /^\[\[MRNOTE_NATIVE_LINK_SOURCE_V1\|(upsert|geometry)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|uri\|([A-Za-z0-9_-]{1,8192})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/u;
const IMPORTED_PAGE_PATTERN =
  /^\[\[MRNOTE_NATIVE_LINK_SOURCE_V1\|(upsert|geometry)\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|page\|([1-9][0-9]{0,4})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/u;
const IMPORTED_DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_LINK_SOURCE_V1\|delete\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\]\]$/u;

export type ManagedNativeLinkRecord =
  | (PdfLinkOverlay & { deleted?: false })
  | {
      id: string;
      /** One-based source PDF page number. */
      page: number;
      deleted: true;
    };

export interface ManagedNativeLinkAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

export interface ManagedNativeLinkPlan {
  annotations: ManagedNativeLinkAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_link" | "duplicate_id" | "patch_limit";
}

export interface ImportedNativeLinkPlan {
  annotations: ManagedNativeLinkAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_link" | "duplicate_id" | "patch_limit";
}

export interface ParsedManagedNativeLinkDirective {
  operation: "upsert" | "delete";
  id: string;
  nativeId: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  link?: Omit<PdfLinkOverlay, "id" | "page">;
}

export interface ParsedImportedNativeLinkDirective {
  operation: "upsert" | "delete";
  objectRef: string;
  /** Only the box moved; the source destination is left untouched. */
  geometryOnly?: boolean;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  link?: Omit<PdfLinkOverlay, "id" | "page">;
}

function isValidPage(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 10_000
  );
}

function encodeOpaqueLinkId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `link.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeOpaqueLinkId(value: string): string | null {
  if (!value.startsWith("link.")) return null;
  try {
    const decoded = decodeURIComponent(value.slice("link.".length));
    return decoded.trim() ? decoded : null;
  } catch {
    return null;
  }
}

function bytesToBase64Url(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return globalThis
    .btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/u, "");
}

function encodeUrlToken(url: string): string | null {
  try {
    const token = bytesToBase64Url(new TextEncoder().encode(url));
    return token.length > 0 &&
      token.length <= MAX_MANAGED_NATIVE_LINK_URL_TOKEN_CHARS
      ? token
      : null;
  } catch {
    return null;
  }
}

function decodeUrlToken(token: string): string | null {
  if (
    !token ||
    token.length > MAX_MANAGED_NATIVE_LINK_URL_TOKEN_CHARS ||
    !/^[A-Za-z0-9_-]+$/u.test(token)
  ) {
    return null;
  }
  const padded = `${token.replaceAll("-", "+").replaceAll("_", "/")}${"=".repeat(
    (4 - (token.length % 4)) % 4,
  )}`;
  try {
    const binary = globalThis.atob(padded);
    const url = new TextDecoder("utf-8", { fatal: true }).decode(
      Uint8Array.from(binary, (character) => character.charCodeAt(0)),
    );
    return encodeUrlToken(url) === token ? url : null;
  } catch {
    return null;
  }
}

function numberToken(value: number): string {
  return Number(value.toFixed(6)).toString();
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
): Pick<PdfLinkOverlay, "x" | "y" | "width" | "height"> | null {
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

const LINK_SYNC_EPSILON = 0.000_1;

function nearlyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= LINK_SYNC_EPSILON;
}

export function isManagedNativeLinkEquivalent(
  readerLink: PdfLinkOverlay,
  nativeLink: PdfLinkOverlay,
): boolean {
  return (
    readerLink.id === nativeLink.id &&
    readerLink.page === nativeLink.page &&
    readerLink.url === nativeLink.url &&
    readerLink.goToPage === nativeLink.goToPage &&
    nearlyEqual(readerLink.x, nativeLink.x) &&
    nearlyEqual(readerLink.y, nativeLink.y) &&
    nearlyEqual(readerLink.width, nativeLink.width) &&
    nearlyEqual(readerLink.height, nativeLink.height)
  );
}

function managedLinkDirective(
  record: ManagedNativeLinkRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const nativeId = encodeOpaqueLinkId(record.id);
  if (!nativeId || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return `${MANAGED_NATIVE_LINK_PREFIX}delete|${nativeId}]]`;
  }
  const link = buildLinkOverlayPayload(record);
  if (!link) return null;
  const destination =
    link.url != null
      ? (() => {
          const token = encodeUrlToken(link.url);
          return token ? `uri|${token}` : null;
        })()
      : link.goToPage != null
        ? `page|${link.goToPage}`
        : null;
  if (!destination) return null;
  return (
    `${MANAGED_NATIVE_LINK_PREFIX}upsert|${nativeId}|${destination}|` +
    `${rotation}|${normalizedGeometryTokens(link)}]]`
  );
}

function importedLinkDirective(
  record: NativePdfLinkRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const source = sanitizeNativePdfLinkAnnotationSource(record.nativeSource);
  if (!source || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return `${IMPORTED_NATIVE_LINK_PREFIX}delete|${source.objectRef}]]`;
  }
  // An untouched source object already exists in the immutable input package.
  if (
    record.sourceAction !== "destination" &&
    record.sourceAction !== "geometry"
  ) {
    return null;
  }
  const link = buildLinkOverlayPayload(record);
  if (!link) return null;
  // A moved or resized box must not re-encode the destination: pdf.js only
  // reports a page number, so an /XYZ 200% target would come back as /Fit.
  const operation = record.sourceAction === "geometry" ? "geometry" : "upsert";
  if (link.url != null) {
    const token = encodeUrlToken(link.url);
    return token
      ? `${IMPORTED_NATIVE_LINK_PREFIX}${operation}|${source.objectRef}|uri|${token}|${rotation}|${normalizedGeometryTokens(link)}]]`
      : null;
  }
  return link.goToPage != null
    ? `${IMPORTED_NATIVE_LINK_PREFIX}${operation}|${source.objectRef}|page|${link.goToPage}|${rotation}|${normalizedGeometryTokens(link)}]]`
    : null;
}

/**
 * Build inert page-annotation rows consumed by the editable-page PDF bridge.
 * Every accepted row becomes an invisible, standard `/Link` annotation in the
 * original package. Unsupported destinations refuse the entire plan.
 */
export function planManagedNativeLinkAnnotations(
  records: readonly ManagedNativeLinkRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeLinkPlan {
  if (records.length === 0) return { annotations: [], supported: true };
  const ids = new Set<string>();
  const annotations: ManagedNativeLinkAnnotationShadow[] = [];
  for (const record of records) {
    if (ids.has(record.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    const text = managedLinkDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text || text.length > MAX_MANAGED_NATIVE_LINK_DIRECTIVE_CHARS) {
      return { annotations: [], supported: false, reason: "invalid_link" };
    }
    annotations.push({
      id: `__mrnote_native_link__${record.id}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_LINK_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

/**
 * Build exact-object destination/delete rows for third-party `/Link`
 * annotations. Untouched source objects emit no directive.
 */
export function planImportedNativeLinkAnnotations(
  records: readonly NativePdfLinkRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ImportedNativeLinkPlan {
  if (records.length === 0) return { annotations: [], supported: true };
  const ids = new Set<string>();
  const objectRefs = new Set<string>();
  const annotations: ManagedNativeLinkAnnotationShadow[] = [];
  for (const record of records) {
    const source = sanitizeNativePdfLinkAnnotationSource(record.nativeSource);
    if (!source) {
      return { annotations: [], supported: false, reason: "invalid_link" };
    }
    if (ids.has(record.id) || objectRefs.has(source.objectRef)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    objectRefs.add(source.objectRef);
    const text = importedLinkDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      if (!record.deleted && record.sourceAction === undefined) continue;
      return { annotations: [], supported: false, reason: "invalid_link" };
    }
    if (text.length > MAX_MANAGED_NATIVE_LINK_DIRECTIVE_CHARS) {
      return { annotations: [], supported: false, reason: "invalid_link" };
    }
    annotations.push({
      id: `__mrnote_native_link_source__${source.objectRef}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_LINK_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

export function isManagedNativeLinkAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    (value.startsWith(MANAGED_NATIVE_LINK_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_LINK_PREFIX))
  );
}

export function parseImportedNativeLinkDirective(
  value: unknown,
): ParsedImportedNativeLinkDirective | null {
  if (
    typeof value !== "string" ||
    value.length > MAX_MANAGED_NATIVE_LINK_DIRECTIVE_CHARS
  ) {
    return null;
  }
  const deleted = IMPORTED_DELETE_PATTERN.exec(value);
  if (deleted) {
    return { operation: "delete", objectRef: deleted[1] };
  }
  const uriMatch = IMPORTED_URI_PATTERN.exec(value);
  const pageMatch = uriMatch ? null : IMPORTED_PAGE_PATTERN.exec(value);
  const match = uriMatch ?? pageMatch;
  if (!match) return null;
  const geometry = parseGeometry(match[5]);
  const rotation =
    match[4] === "auto" ? "auto" : (Number(match[4]) as 0 | 90 | 180 | 270);
  if (!geometry) return null;
  const built = buildLinkOverlayPayload({
    page: 1,
    ...geometry,
    url: uriMatch ? decodeUrlToken(match[3]) : null,
    goToPage: pageMatch ? Number(match[3]) : null,
  });
  if (!built) return null;
  return {
    operation: "upsert",
    objectRef: match[2],
    ...(match[1] === "geometry" ? { geometryOnly: true } : {}),
    rotation,
    link: {
      x: built.x,
      y: built.y,
      width: built.width,
      height: built.height,
      url: built.url,
      goToPage: built.goToPage,
    },
  };
}

export function parseManagedNativeLinkDirective(
  value: unknown,
): ParsedManagedNativeLinkDirective | null {
  if (
    typeof value !== "string" ||
    value.length > MAX_MANAGED_NATIVE_LINK_DIRECTIVE_CHARS
  ) {
    return null;
  }
  const deleted = DELETE_PATTERN.exec(value);
  if (deleted) {
    const nativeId = deleted[1];
    const id = decodeOpaqueLinkId(nativeId);
    return id ? { operation: "delete", id, nativeId } : null;
  }
  const uriMatch = UPSERT_URI_PATTERN.exec(value);
  const pageMatch = uriMatch ? null : UPSERT_PAGE_PATTERN.exec(value);
  const match = uriMatch ?? pageMatch;
  if (!match) return null;
  const nativeId = match[1];
  const id = decodeOpaqueLinkId(nativeId);
  const geometry = parseGeometry(match[4]);
  const rotation =
    match[3] === "auto" ? "auto" : (Number(match[3]) as 0 | 90 | 180 | 270);
  if (!id || !geometry) return null;
  const built = buildLinkOverlayPayload({
    page: 1,
    ...geometry,
    url: uriMatch ? decodeUrlToken(match[2]) : null,
    goToPage: pageMatch ? Number(match[2]) : null,
  });
  if (!built) return null;
  return {
    operation: "upsert",
    id,
    nativeId,
    rotation,
    link: {
      x: built.x,
      y: built.y,
      width: built.width,
      height: built.height,
      url: built.url,
      goToPage: built.goToPage,
    },
  };
}

export function parseManagedNativeLinkShadow(
  value: unknown,
): ManagedNativeLinkRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseManagedNativeLinkDirective(row.text);
  if (!parsed) return null;
  if (parsed.operation === "delete") {
    return { id: parsed.id, page, deleted: true };
  }
  return parsed.link ? { id: parsed.id, page, ...parsed.link } : null;
}

export function parseImportedNativeLinkShadow(
  value: unknown,
): NativePdfLinkRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseImportedNativeLinkDirective(row.text ?? row.comment);
  if (!parsed) return null;
  const nativeSource = {
    kind: NATIVE_PDF_LINK_SOURCE_KIND,
    objectRef: parsed.objectRef,
    subtype: "Link",
  } as const;
  const id = `pdf-native-link:${page}:${parsed.objectRef}`;
  if (parsed.operation === "delete") {
    return { id, page, deleted: true, nativeSource };
  }
  if (!parsed.link) return null;
  const record: NativePdfLinkAnnotation = {
    id,
    page,
    ...parsed.link,
    nativeSource,
    sourceAction: parsed.geometryOnly ? "geometry" : "destination",
  };
  return record;
}
