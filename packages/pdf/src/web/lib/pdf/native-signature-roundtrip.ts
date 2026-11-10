/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  DEFAULT_SIGNATURE_FONT,
  buildSignaturePayload,
  normalizeSignatureObjectRotation,
  sanitizeSignatureImage,
  type SignatureAnnotation,
} from "@mrnote/core/web/lib/pdf/signatures.ts";

export const MANAGED_NATIVE_SIGNATURE_PREFIX = "[[MRNOTE_NATIVE_SIGNATURE_V1|";
export const MANAGED_NATIVE_SIGNATURE_V2_PREFIX =
  "[[MRNOTE_NATIVE_SIGNATURE_V2|";
export const MANAGED_NATIVE_SIGNATURE_V3_PREFIX =
  "[[MRNOTE_NATIVE_SIGNATURE_V3|";
export const MAX_MANAGED_NATIVE_SIGNATURE_PATCHES = 400;
export const MAX_MANAGED_NATIVE_SIGNATURE_IMAGE_TOTAL_CHARS = 12 * 1024 * 1024;
const MAX_MANAGED_NATIVE_SIGNATURE_DIRECTIVE_CHARS = 1_402_000;
const MAX_MANAGED_NATIVE_SIGNATURE_TEXT_TOKEN_CHARS = 1024;

const MANAGED_NATIVE_SIGNATURE_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_SIGNATURE_V1\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(drawn|typed)\|(none|cursive|serif|sans|mono)\|(none|[A-Za-z0-9_-]{1,1024})\|(png|jpeg)\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/i;
const MANAGED_NATIVE_SIGNATURE_V2_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_SIGNATURE_V2\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(image)\|(none)\|(none)\|(png|jpeg)\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/i;
const MANAGED_NATIVE_SIGNATURE_V3_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_SIGNATURE_V3\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(drawn|image|typed)\|(none|cursive|serif|sans|mono)\|(none|[A-Za-z0-9_-]{1,1024})\|(png|jpeg)\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9]{1,3})\|([0-9.,-]{7,256})\]\]$/i;
const MANAGED_NATIVE_SIGNATURE_DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_SIGNATURE_V1\|delete\|([A-Za-z0-9._~%\-]{1,384})\]\]$/;

export interface ManagedNativeSignatureAnnotation extends SignatureAnnotation {
  /**
   * Strict PNG/JPEG raster used by the native `/Stamp` appearance. Drawn
   * signatures reuse their source image; typed signatures keep a deterministic
   * browser raster so their selected script face remains identical in WPS,
   * Acrobat, Preview, and Poppler.
   */
  nativeAppearanceImage: string;
  deleted?: false;
}

export interface ManagedNativeSignatureTombstone {
  id: string;
  /** One-based original PDF page number. */
  page: number;
  deleted: true;
}

export type ManagedNativeSignatureRecord =
  ManagedNativeSignatureAnnotation | ManagedNativeSignatureTombstone;

export interface ManagedNativeSignatureAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

export interface ManagedNativeSignaturePlan {
  annotations: ManagedNativeSignatureAnnotationShadow[];
  supported: boolean;
  reason?:
    "invalid_signature" | "image_total_limit" | "duplicate_id" | "patch_limit";
}

export interface ParsedManagedNativeSignatureDirective {
  operation: "upsert" | "delete";
  id: string;
  nativeId: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  signature?: Omit<ManagedNativeSignatureAnnotation, "id" | "page">;
}

function encodeOpaqueSignatureId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `signature.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeOpaqueSignatureId(value: string): string | null {
  if (!value.startsWith("signature.")) return null;
  try {
    const decoded = decodeURIComponent(value.slice("signature.".length));
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
): Pick<SignatureAnnotation, "x" | "y" | "width" | "height"> | null {
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

function encodeTextToken(text: string): string | null {
  try {
    const token = bytesToBase64Url(new TextEncoder().encode(text));
    return token.length > 0 &&
      token.length <= MAX_MANAGED_NATIVE_SIGNATURE_TEXT_TOKEN_CHARS
      ? token
      : null;
  } catch {
    return null;
  }
}

function decodeTextToken(token: string): string | null {
  if (
    !token ||
    token.length > MAX_MANAGED_NATIVE_SIGNATURE_TEXT_TOKEN_CHARS ||
    !/^[A-Za-z0-9_-]+$/u.test(token)
  ) {
    return null;
  }
  const padded = `${token.replaceAll("-", "+").replaceAll("_", "/")}${"=".repeat(
    (4 - (token.length % 4)) % 4,
  )}`;
  try {
    const binary = globalThis.atob(padded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return encodeTextToken(text) === token ? text : null;
  } catch {
    return null;
  }
}

const NATIVE_SIGNATURE_SYNC_EPSILON = 0.0001;

function nearlyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= NATIVE_SIGNATURE_SYNC_EPSILON;
}

export function isManagedNativeSignatureEquivalent(
  readerSignature: SignatureAnnotation,
  nativeSignature: ManagedNativeSignatureAnnotation,
): boolean {
  const readerAppearance = sanitizeSignatureImage(
    readerSignature.nativeAppearanceImage,
  );
  return (
    readerAppearance !== null &&
    readerSignature.id === nativeSignature.id &&
    readerSignature.page === nativeSignature.page &&
    readerSignature.kind === nativeSignature.kind &&
    readerSignature.image === nativeSignature.image &&
    readerSignature.text === nativeSignature.text &&
    readerSignature.font === nativeSignature.font &&
    readerAppearance === nativeSignature.nativeAppearanceImage &&
    nearlyEqual(readerSignature.x, nativeSignature.x) &&
    nearlyEqual(readerSignature.y, nativeSignature.y) &&
    nearlyEqual(readerSignature.width, nativeSignature.width) &&
    nearlyEqual(readerSignature.height, nativeSignature.height) &&
    normalizeSignatureObjectRotation(readerSignature.objectRotation) ===
      normalizeSignatureObjectRotation(nativeSignature.objectRotation)
  );
}

function managedSignatureDirective(
  record: ManagedNativeSignatureRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const nativeId = encodeOpaqueSignatureId(record.id);
  if (!nativeId || !isValidPage(record.page)) return null;
  if (record.deleted) {
    return `${MANAGED_NATIVE_SIGNATURE_PREFIX}delete|${nativeId}]]`;
  }
  const signature = buildSignaturePayload(record);
  const appearance = sanitizeSignatureImage(record.nativeAppearanceImage);
  if (!signature || !appearance) return null;
  const imageMatch =
    /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/iu.exec(appearance);
  if (!imageMatch) return null;
  const textToken =
    signature.kind === "typed" ? encodeTextToken(signature.text) : "none";
  if (!textToken) return null;
  const font = signature.kind === "typed" ? signature.font : ("none" as const);
  const objectRotation = normalizeSignatureObjectRotation(
    signature.objectRotation,
  );
  const prefix =
    objectRotation !== 0
      ? MANAGED_NATIVE_SIGNATURE_V3_PREFIX
      : signature.kind === "image"
        ? MANAGED_NATIVE_SIGNATURE_V2_PREFIX
        : MANAGED_NATIVE_SIGNATURE_PREFIX;
  return (
    `${prefix}upsert|${nativeId}|${signature.kind}|` +
    `${font}|${textToken}|${imageMatch[1].toLowerCase()}|${imageMatch[2]}|` +
    `${rotation}|${objectRotation === 0 ? "" : `${objectRotation}|`}${normalizedGeometryTokens(signature)}]]`
  );
}

/**
 * Build invisible page annotation rows consumed by the editable-page bridge.
 * These rows become standard self-contained `/Stamp` annotations explicitly
 * marked as non-cryptographic visual signatures. No `/FT /Sig` value is ever
 * created or mutated.
 */
export function planManagedNativeSignatureAnnotations(
  records: readonly ManagedNativeSignatureRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeSignaturePlan {
  if (records.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const annotations: ManagedNativeSignatureAnnotationShadow[] = [];
  let imageDirectiveCharacters = 0;
  for (const record of records) {
    if (ids.has(record.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    const text = managedSignatureDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      return {
        annotations: [],
        supported: false,
        reason: "invalid_signature",
      };
    }
    imageDirectiveCharacters += text.length;
    if (
      imageDirectiveCharacters > MAX_MANAGED_NATIVE_SIGNATURE_IMAGE_TOTAL_CHARS
    ) {
      return {
        annotations: [],
        supported: false,
        reason: "image_total_limit",
      };
    }
    annotations.push({
      id: `__mrnote_native_signature__${record.id}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_SIGNATURE_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

export function isManagedNativeSignatureAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    (value.startsWith(MANAGED_NATIVE_SIGNATURE_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_SIGNATURE_V2_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_SIGNATURE_V3_PREFIX))
  );
}

export function parseManagedNativeSignatureDirective(
  value: unknown,
): ParsedManagedNativeSignatureDirective | null {
  if (
    typeof value !== "string" ||
    value.length > MAX_MANAGED_NATIVE_SIGNATURE_DIRECTIVE_CHARS
  ) {
    return null;
  }
  const deleteMatch = MANAGED_NATIVE_SIGNATURE_DELETE_PATTERN.exec(value);
  if (deleteMatch) {
    const nativeId = deleteMatch[1];
    const id = decodeOpaqueSignatureId(nativeId);
    return id ? { operation: "delete", id, nativeId } : null;
  }
  const v3Match = MANAGED_NATIVE_SIGNATURE_V3_UPSERT_PATTERN.exec(value);
  const match =
    v3Match ??
    MANAGED_NATIVE_SIGNATURE_UPSERT_PATTERN.exec(value) ??
    MANAGED_NATIVE_SIGNATURE_V2_UPSERT_PATTERN.exec(value);
  if (!match) return null;
  const nativeId = match[1];
  const id = decodeOpaqueSignatureId(nativeId);
  const kind = match[2].toLowerCase() as "drawn" | "image" | "typed";
  const fontToken = match[3].toLowerCase();
  const textToken = match[4];
  const mime = match[5].toLowerCase();
  const appearance = sanitizeSignatureImage(
    `data:image/${mime};base64,${match[6]}`,
  );
  const objectRotation = v3Match
    ? normalizeSignatureObjectRotation(match[8])
    : 0;
  const geometry = parseGeometry(v3Match ? match[9] : match[8]);
  if (
    !id ||
    !appearance ||
    !geometry ||
    (v3Match && (objectRotation === 0 || objectRotation !== Number(match[8])))
  ) {
    return null;
  }
  const text = kind === "typed" ? decodeTextToken(textToken) : "";
  if (
    (kind !== "typed" &&
      (fontToken !== "none" || textToken.toLowerCase() !== "none")) ||
    (kind === "typed" &&
      (fontToken === "none" || !text || textToken.toLowerCase() === "none"))
  ) {
    return null;
  }
  const signature = buildSignaturePayload({
    page: 1,
    ...geometry,
    kind,
    image: kind !== "typed" ? appearance : "",
    text: kind === "typed" ? (text ?? "") : "",
    font: kind === "typed" ? fontToken : DEFAULT_SIGNATURE_FONT,
    nativeAppearanceImage: appearance,
    objectRotation,
  });
  if (!signature?.nativeAppearanceImage) return null;
  return {
    operation: "upsert",
    id,
    nativeId,
    rotation:
      match[7] === "auto" ? "auto" : (Number(match[7]) as 0 | 90 | 180 | 270),
    signature: {
      x: signature.x,
      y: signature.y,
      width: signature.width,
      height: signature.height,
      kind: signature.kind,
      image: signature.image,
      text: signature.text,
      font: signature.font,
      nativeAppearanceImage: signature.nativeAppearanceImage,
      ...(signature.objectRotation === undefined
        ? {}
        : { objectRotation: signature.objectRotation }),
    },
  };
}

export function parseManagedNativeSignatureShadow(
  value: unknown,
): ManagedNativeSignatureRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseManagedNativeSignatureDirective(row.text);
  if (!parsed) return null;
  if (parsed.operation === "delete") {
    return { id: parsed.id, page, deleted: true };
  }
  return parsed.signature ? { id: parsed.id, page, ...parsed.signature } : null;
}
