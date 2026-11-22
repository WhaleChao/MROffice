/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { DEFAULT_NEW_BOX_FONT_SIZE } from "@mrnote/core/web/lib/pdf/text-box-create.ts";
import { fontClassFromFamily, type TextBoxEditLike } from "./text-box-flatten";
import { effectiveTextBoxKind } from "@mrnote/core/web/lib/pdf/text-box-kind.ts";
import { sanitizePdfTextBoxEditGeometry } from "@mrnote/core/web/lib/pdf/text-box-rect.ts";
import { pdfFreeTextRectangleDifferencesV2Enabled } from "./pdf-flags";
import { MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS, pdfSourceTextBoxCoversInk, pdfSourceTextBoxInkCoverRect, readPdfSourceTextBoxBinding, type PdfSourceTextBoxBinding } from "./source-text-box";
import { PDF_PORTABLE_LAYOUT_POINT_TOLERANCE, readUnverifiedPdfPortableTextLayout, type PdfPortableTextLayout } from "./portable-text-layout";
import {
  decodePdfTextBoxTypographyField,
  encodePdfTextBoxTypographyField,
  hasPdfTextBoxTypography,
  readPdfTextBoxTypography,
  type PdfTextBoxTypographyFields,
} from "@mrnote/core/web/lib/pdf/text-box-typography.ts";
import {
  decodePdfTextBoxRunsField,
  encodePdfTextBoxRunsField,
  isPdfBundledCjkCharacter,
  isPdfWinAnsiCharacter,
  pdfTextBoxDrawnRunsNeedRunWriter,
  pdfTextBoxDrawnRunsPortable,
  resolvePdfTextBoxDrawnRuns,
  MAX_PDF_TEXT_BOX_RUNS_FIELD_CHARS,
  type PdfTextBoxDrawnRun,
} from "./text-box-runs";
import {
  NATIVE_PDF_FREE_TEXT_SOURCE_KIND,
  isNativePdfFreeTextRecord,
  sanitizeNativePdfFreeTextAnnotationSource,
  type NativePdfFreeTextAnnotation,
} from "./native-free-text-annotations";
import { pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";

/** A frame grown to `height` (page fractions) to hold its rows: down from
 * where it is while the page has room, moved up just enough to stay on the
 * page when it would pass the foot of it. */
export function pdfTextBoxGrownFrame(edit: Readonly<{ y?: unknown }>, height: number): { height: number; y?: number } {
  const y = Number(edit.y);
  if (!Number.isFinite(y) || y + height <= 1) return { height };
  return { height, y: Math.max(0, 1 - height) };
}

export const MANAGED_NATIVE_TEXT_BOX_PREFIX = "[[MRNOTE_NATIVE_TEXT_BOX_V1|";
export const MANAGED_NATIVE_TEXT_BOX_V2_PREFIX =
  "[[MRNOTE_NATIVE_TEXT_BOX_V2|";
export const MANAGED_NATIVE_TEXT_BOX_V3_PREFIX =
  "[[MRNOTE_NATIVE_TEXT_BOX_V3|";
export const IMPORTED_NATIVE_TEXT_BOX_PREFIX =
  "[[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V1|";
export const IMPORTED_NATIVE_TEXT_BOX_RD_PREFIX =
  "[[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V2|";
export const IMPORTED_NATIVE_TEXT_BOX_ROTATION_PREFIX =
  "[[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V3|";
export const IMPORTED_NATIVE_TEXT_BOX_RD_ROTATION_PREFIX =
  "[[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V4|";
export const MAX_MANAGED_NATIVE_TEXT_BOX_PATCHES = 400;
export const MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_TOTAL_CHARS = 12 * 1024 * 1024;
const MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_BYTES = 1024 * 1024;
const MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_CHARS =
  Math.ceil(MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_BYTES / 3) * 4;
const MAX_MANAGED_NATIVE_TEXT_BOX_TEXT_CHARS = 20_000;
const MAX_MANAGED_NATIVE_TEXT_BOX_TEXT_TOKEN_CHARS = 80_000;
const MAX_MANAGED_NATIVE_TEXT_BOX_DIRECTIVE_CHARS =
  MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_CHARS +
  MAX_MANAGED_NATIVE_TEXT_BOX_TEXT_TOKEN_CHARS +
  // (The source binding's token and its wrapper's name.)
  MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS + 64 + 262_144 + MAX_PDF_TEXT_BOX_RUNS_FIELD_CHARS;

const UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_V1\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(left|center|right|justify)\|(serif|sans|mono)\|([01])\|([01])\|([0-3])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/;
const V2_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_V2\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(left|center|right|justify)\|(serif|sans|mono)\|([01])\|([01])\|([0-3])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.]{1,20})\|([0-9.,-]{7,256})\]\]$/;
const V3_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_V3\|upsert\|([A-Za-z0-9._~%\-]{1,384})\|(left|center|right|justify)\|(serif|sans|mono)\|([01])\|([01])\|([0-3])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.]{1,20})\|([0-9.,-]{7,256})\]\]$/;
const DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_V1\|delete\|([A-Za-z0-9._~%\-]{1,384})\]\]$/;
// `source` = the author's `/Q` still governs the object (see `alignFromSource`).
const IMPORTED_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V1\|upsert\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(left|center|right|source)\|(serif|sans|mono)\|([01])\|([01])\|([01])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/;
const IMPORTED_DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V1\|delete\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\]\]$/;
const IMPORTED_RD_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V2\|upsert\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(left|center|right|source)\|(serif|sans|mono)\|([01])\|([01])\|([01])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.,-]{7,256})\]\]$/;
const IMPORTED_ROTATION_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V3\|upsert\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(left|center|right|source)\|(serif|sans|mono)\|([01])\|([01])\|([01])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.]{1,20})\|([0-9.,-]{7,256})\]\]$/;
const IMPORTED_RD_ROTATION_UPSERT_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V4\|upsert\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\|(left|center|right|source)\|(serif|sans|mono)\|([01])\|([01])\|([01])\|([0-9.]{1,20})\|([0-9A-Fa-f]{6})\|([0-9A-Fa-f]{16})\|([A-Za-z0-9_-]{1,80000})\|png\|([A-Za-z0-9+/]+={0,2})\|(auto|0|90|180|270)\|([0-9.]{1,20})\|([0-9.,-]{7,256})\]\]$/;
const IMPORTED_RD_DELETE_PATTERN =
  /^\[\[MRNOTE_NATIVE_TEXT_BOX_SOURCE_V2\|delete\|([1-9][0-9]{0,9}R(?:[0-9]{1,5})?)\]\]$/;
const PNG_DATA_URL_PATTERN = /^data:image\/png;base64,([A-Za-z0-9+/]+={0,2})$/u;

export interface ManagedNativeTextBoxAnnotation {
  portableLayout?: PdfPortableTextLayout;
  sourceText?: PdfSourceTextBoxBinding;
  restoreSource?: boolean;
  id: string;
  /** One-based source PDF page number. */
  page: number;
  /** Page-relative fractions with a top-left origin. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** Clockwise object rotation around the box centre; absent means upright. */
  angle?: number;
  text: string;
  fontSize: number;
  color: string;
  backgroundColor?: string;
  /** Where the fill of a box over a run that paints nothing stays once the
   * frame no longer holds the line it replaced (PDF 0107,
   * pdfSourceTextBoxInkCoverRect; page fractions, top-left origin). The
   * annotation is the union of it and the box: the fill is painted only
   * here, the text at the box, the rest clear. Sent as four more numbers
   * after the box's in the V3 geometry field. */
  cover?: { x: number; y: number; width: number; height: number };
  align: "left" | "center" | "right";
  font: "serif" | "sans" | "mono";
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /**
   * Word's Strikethrough and Justify for the whole box. On the wire they ride
   * the existing fields: the underline flag is a 2-bit decoration code
   * (1 underline, 2 strikethrough) and Justify is the alignment `justify`,
   * so a writer that predates them refuses the directive instead of
   * silently dropping the format.
   */
  strikethrough?: boolean;
  justify?: boolean;
  /** The remaining whole-box Word formats (text-box-typography.ts); sent as
   * one optional `~` field before `png`. */
  typography?: PdfTextBoxTypographyFields;
  


  formula?: true;
  /**
   * The text as drawn runs (text-box-runs.ts) when it is not drawn in the
   * box's one style: part of it formatted, or set in an Office family. Sent
   * as the `^` field after the text token; such a box is always written with
   * its row plan (`portableLayout`).
   */
  runs?: PdfTextBoxDrawnRun[];
  /** Detects source-font-family changes that alter the browser raster. */
  styleKey: string;
  nativeAppearanceImage: string;
  deleted?: false;
}

export interface ManagedNativeTextBoxTombstone {
  sourceText?: PdfSourceTextBoxBinding;
  restoreSource?: boolean;
  id: string;
  /** One-based source PDF page number. */
  page: number;
  deleted: true;
}

export type ManagedNativeTextBoxRecord =
  ManagedNativeTextBoxAnnotation | ManagedNativeTextBoxTombstone;

export interface ManagedNativeTextBoxAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

export interface ManagedNativeTextBoxPlan {
  annotations: ManagedNativeTextBoxAnnotationShadow[];
  supported: boolean;
  reason?:
    "invalid_text_box" | "image_total_limit" | "duplicate_id" | "patch_limit";
}

export interface ImportedNativeTextBoxPlan {
  annotations: ManagedNativeTextBoxAnnotationShadow[];
  supported: boolean;
  reason?:
    "invalid_text_box" | "image_total_limit" | "duplicate_id" | "patch_limit";
}

export interface ParsedManagedNativeTextBoxDirective {
  sourceText?: PdfSourceTextBoxBinding;
  restoreSource?: boolean;
  operation: "upsert" | "delete";
  id: string;
  nativeId: string;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  textBox?: Omit<ManagedNativeTextBoxAnnotation, "id" | "page">;
}

export interface ParsedImportedNativeTextBoxDirective {
  operation: "upsert" | "delete";
  objectRef: string;
  sourceVersion: 1 | 2 | 3 | 4;
  rotation?: 0 | 90 | 180 | 270 | "auto";
  textBox?: Omit<
    NativePdfFreeTextAnnotation,
    | "id"
    | "page"
    | "created"
    | "originX"
    | "originY"
    | "originWidth"
    | "originHeight"
    | "nativeSource"
    | "sourceAction"
  >;
}

export interface ImportedNativeTextBoxAppearanceResult<
  T extends TextBoxEditLike,
> {
  records: T[];
  appearanceFailureIds: string[];
}

export interface ManagedNativeTextBoxReconcileResult {
  records: ManagedNativeTextBoxRecord[];
  appearanceFailureIds: string[];
  duplicateIds: string[];
  /** Boxes whose frame grew down to hold their rows (page fractions). */
  grownHeights?: ReadonlyMap<string, number>;
}

function isValidPage(value: unknown): value is number {
  return (
    typeof value === "number" &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 10_000
  );
}

function encodeOpaqueTextBoxId(value: string): string | null {
  const trimmed = value.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(
    /[!'()*]/g,
    (character) =>
      `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
  const token = `textbox.${encoded}`;
  return token.length <= 384 ? token : null;
}

export function managedNativeTextBoxAnnotationName(id: string): string | null {
  const encoded = encodeOpaqueTextBoxId(id);
  return encoded ? `MRNote:${encoded}` : null;
}

/** Read-only preview identity; never an annotation writer or a bounds match. */
export async function resolveManagedTextBoxPreviewAnnotationId(
  sourceBytes: Uint8Array,
  pageNumber: number,
  textBoxId: string,
): Promise<string | null> {
  const nativeId = encodeOpaqueTextBoxId(textBoxId);
  if (!nativeId || !Number.isSafeInteger(pageNumber) || pageNumber < 1) return null;
  const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRef, PDFBool, PDFString, PDFHexString } =
    await import("pdf-lib");
  const document = await PDFDocument.load(sourceBytes, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
  if (document.isEncrypted || document.getPageCount() > 10_000) return null;
  let match: string | null = null;
  let count = 0;
  for (const [index, page] of document.getPages().entries()) {
    const annotations = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    if (!annotations) continue;
    count += annotations.size();
    if (count > MAX_MANAGED_NATIVE_TEXT_BOX_PATCHES * 10) return null;
    for (let i = 0; i < annotations.size(); i += 1) {
      const reference = annotations.get(i);
      const annotation = document.context.lookup(reference);
      if (!(annotation instanceof PDFDict)) continue;
      const name = annotation.lookup(PDFName.of("NM"));
      if (!(name instanceof PDFString || name instanceof PDFHexString) ||
        name.decodeText() !== `MRNote:${nativeId}`) continue;
      if (match || index + 1 !== pageNumber || !(reference instanceof PDFRef) ||
        annotation.get(PDFName.of("Subtype")) !== PDFName.of("FreeText") ||
        annotation.get(PDFName.of("MRNoteManaged")) !== PDFBool.True ||
        annotation.get(PDFName.of("MRNoteTextBox")) !== PDFBool.True) return null;
      match = `${reference.objectNumber}R${reference.generationNumber || ""}`;
    }
  }
  return match;
}

function decodeOpaqueTextBoxId(value: string): string | null {
  if (!value.startsWith("textbox.")) return null;
  try {
    const decoded = decodeURIComponent(value.slice("textbox.".length));
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

function encodeTextToken(text: string): string | null {
  if (
    !text.trim() ||
    text.length > MAX_MANAGED_NATIVE_TEXT_BOX_TEXT_CHARS ||
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/u.test(text)
  ) {
    return null;
  }
  try {
    const token = bytesToBase64Url(new TextEncoder().encode(text));
    return token.length <= MAX_MANAGED_NATIVE_TEXT_BOX_TEXT_TOKEN_CHARS
      ? token
      : null;
  } catch {
    return null;
  }
}

function decodeTextToken(token: string): string | null {
  if (
    !token ||
    token.length > MAX_MANAGED_NATIVE_TEXT_BOX_TEXT_TOKEN_CHARS ||
    !/^[A-Za-z0-9_-]+$/u.test(token)
  ) {
    return null;
  }
  const padded = `${token.replaceAll("-", "+").replaceAll("_", "/")}${"=".repeat(
    (4 - (token.length % 4)) % 4,
  )}`;
  try {
    const binary = globalThis.atob(padded);
    const text = new TextDecoder("utf-8", { fatal: true }).decode(
      Uint8Array.from(binary, (character) => character.charCodeAt(0)),
    );
    return encodeTextToken(text) === token ? text : null;
  } catch {
    return null;
  }
}

/** The key a box is saved under: this page's fingerprint of the family, the
 * face name and the fill it sets the box in. */
export function managedNativeTextBoxStyleKey(edit: TextBoxEditLike): string {
  return styleKeyForEdit(edit);
}

function styleKeyForEdit(edit: TextBoxEditLike): string {
  const input = [
    edit.fontFamily ?? "",
    edit.sourceFontName ?? "",
    edit.backgroundColor ?? "",
  ].join("\u0000");
  let first = 0x811c9dc5;
  let second = 0x9e3779b9;
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    first ^= code;
    first = Math.imul(first, 0x01000193);
    second = Math.imul(second ^ code, 33) ^ (second >>> 13);
  }
  return `${(first >>> 0).toString(16).padStart(8, "0")}${(second >>> 0)
    .toString(16)
    .padStart(8, "0")}`;
}

function normalizeColor(value: unknown): string | null {
  if (value === undefined || value === null || value === "") return "#000000";
  if (typeof value !== "string" || !/^#[0-9A-Fa-f]{6}$/u.test(value)) {
    return null;
  }
  return value.toUpperCase();
}

const isWinAnsiCharacter = isPdfWinAnsiCharacter;
const isBundledCjkCharacter = isPdfBundledCjkCharacter;

/** Letters past Windows-1252 that the base-14 class's metric twin (Arimo,
 * Tinos, Cousine) draws, in any style: Latin Extended-A/B and Additional
 * (Turkish ı ğ ş İ, Polish ł, Czech ř, Vietnamese), combining marks, Greek
 * and Cyrillic. The writer sets them in the twin, embedded
 * (native_edit_primitives._latin_twin_face), and checks each glyph; they
 * were refused, and a box retyped 「Onaylandı」 came back 22 times (AI70H
 * H205 B18). */
export function isPdfLatinTwinCharacter(character: string): boolean {
  const code = character.codePointAt(0) ?? 0;
  return (code >= 0x0100 && code <= 0x024f) || (code >= 0x0300 && code <= 0x036f) ||
    (code >= 0x0370 && code <= 0x03ff) || (code >= 0x0400 && code <= 0x052f) ||
    (code >= 0x1e00 && code <= 0x1eff);
}

export function isPortableNativeTextBoxTypography(
  text: string,
  font: "serif" | "sans" | "mono",
  bold: boolean,
  italic: boolean,
): boolean {
  if ([...text].every(isWinAnsiCharacter)) return true;
  // Whatever the box's face and style: its Latin is set in that base-14 face
  // and every other character in the bundled Noto Sans SC, upright and
  // regular (as Word and WPS fall back for a face with no such glyph). Only
  // a character neither face holds keeps the box off the native path.
  // (Refusing a bold, italic, serif or mono box with CJK left the page data
  // saved and the PDF unwritten: a ghost box over the still-live source.)
  void font; void bold; void italic;
  return [...text].every(
    (character) =>
      isWinAnsiCharacter(character) || isPdfLatinTwinCharacter(character) ||
      isBundledCjkCharacter(character),
  );
}

function numberToken(value: number): string {
  return Number(value.toFixed(6)).toString();
}

function normalizedObjectRotation(value: unknown): number | null {
  if (value === undefined || value === null) return 0;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return null;
  const normalized = ((parsed % 360) + 360) % 360;
  const rounded = Number(numberToken(normalized));
  return rounded === 360 ? 0 : rounded;
}

function normalizedGeometryTokens(value: {
  x: number;
  y: number;
  width: number;
  height: number;
}): string {
  return [value.x, value.y, value.width, value.height]
    .map(numberToken)
    .join(",");
}

function parseGeometry(token: string): {
  x: number;
  y: number;
  width: number;
  height: number;
} | null {
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
  return { x, y, width, height };
}

export function sanitizeManagedNativeTextBoxImage(
  value: unknown,
): string | null {
  if (typeof value !== "string") return null;
  const match = PNG_DATA_URL_PATTERN.exec(value);
  if (
    !match ||
    match[1].length === 0 ||
    match[1].length > MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_CHARS ||
    match[1].length % 4 !== 0
  ) {
    return null;
  }
  try {
    const binary = globalThis.atob(match[1]);
    if (
      binary.length === 0 ||
      binary.length > MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_BYTES
    ) {
      return null;
    }
    const signature = Array.from(binary.slice(0, 8), (character) =>
      character.charCodeAt(0),
    );
    if (
      signature.join(",") !==
      [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].join(",")
    ) {
      return null;
    }
    return value;
  } catch {
    return null;
  }
}

function textBoxPropertiesFromEdit(
  edit: TextBoxEditLike,
  options: { allowImportedSource?: boolean } = {},
): Omit<
  ManagedNativeTextBoxAnnotation,
  "id" | "page" | "nativeAppearanceImage"
> | null {
  const geometry = sanitizePdfTextBoxEditGeometry(edit);
  const id = typeof edit.id === "string" ? edit.id : "";
  const page = Number(edit.page);
  const fontSize =
    edit.fontSize === undefined
      ? DEFAULT_NEW_BOX_FONT_SIZE
      : Number(edit.fontSize);
  const color = normalizeColor(edit.color);
  const backgroundColor =
    edit.backgroundColor === undefined ||
    edit.backgroundColor === null ||
    edit.backgroundColor === ""
      ? undefined
      : normalizeColor(edit.backgroundColor);
  // A formula box (插入 › 公式) is written as its rendered equation, its
  // LaTeX source the annotation's text: searchable, and two formulas never
  // read as the same box (their fallback text may).
  const formula = effectiveTextBoxKind(edit) === "formula";
  const latex = formula && typeof edit.latex === "string" ? edit.latex.trim() : "";
  const text = formula ? latex : typeof edit.text === "string" ? edit.text : "";
  let font = fontClassFromFamily(edit.sourceFontName || edit.fontFamily);
  let bold = Boolean(edit.bold);
  let italic = Boolean(edit.italic);
  let underline = Boolean(edit.underline);
  let strikethrough = Boolean(edit.strikethrough);
  let runColor = color;
  const angle = normalizedObjectRotation(edit.angle);
  const sourceText = edit.sourceText === undefined ? null : readPdfSourceTextBoxBinding(edit.sourceText);
  const boxTypography = readPdfTextBoxTypography(edit, fontSize);
  // Part of the text formatted (or the box set in an Office family): its
  // drawn runs go to the run writer, which sets each in its own face and
  // size on the box's row plan. Its superscript/subscript is the runs' own.
  const drawn = formula ? null : resolvePdfTextBoxDrawnRuns(edit);
  let runWriter = drawn !== null && pdfTextBoxDrawnRunsNeedRunWriter(edit, drawn);
  const typography: PdfTextBoxTypographyFields = { ...boxTypography };
  if (runWriter) delete typography.textScript;
  // One style over the whole text (bold pressed before the first word sets
  // it on the text, not the box) is the box's style. Where the row plan has
  // no place for the box's formats (竖排, 分栏, 艺术字), the box is written
  // in it as a box of that style; elsewhere the run writer keeps the box
  // it has always drawn (an unchanged box is not redrawn).
  // (Its size the box's too: the box's formats are measured in it.)
  const only = runWriter && drawn!.length === 1 && !drawn![0].officeFamily && !edit.sourceText &&
    Math.abs(drawn![0].fontSize - fontSize) < 1e-9 ? drawn![0] : null;
  if (only && !pdfSourceTextBoxTypographyPortable(typography)) {
    runWriter = false;
    font = only.font;
    bold = only.bold;
    italic = only.italic;
    underline = only.underline;
    strikethrough = only.strikethrough;
    runColor = normalizeColor(only.color) ?? color;
    if (only.script) typography.textScript = only.script;
    else delete typography.textScript;
  }
  // Dragged off the scanned words it replaced, an OCR line's box leaves its
  // fill where they are and is clear where it stands (PDF 0107).
  const cover = sourceText && backgroundColor ? pdfSourceTextBoxInkCoverRect({ ...edit, ...geometry, sourceText }) : null;
  if (
    (!options.allowImportedSource && isNativePdfFreeTextRecord(edit)) ||
    edit.created !== true ||
    edit.deleted === true ||
    // A formula is a box the user made, with a source to draw.
    (formula && (!latex || edit.sourceText !== undefined)) ||
    !geometry ||
    !id.trim() ||
    id.length > 120 ||
    !Number.isInteger(page) ||
    page < 0 ||
    page >= 10_000 ||
    !encodeTextToken(text) ||
    !Number.isFinite(fontSize) ||
    fontSize <= 0 ||
    fontSize > 1 ||
    !color ||
    !runColor ||
    backgroundColor === null ||
    angle === null ||
    // Vertical writing is the writer's only for a box made in MRAI (Word's
    // 竖排, in its typography); an imported vertical line keeps its own.
    (edit.writingMode === "vertical-rl" && !edit.created) ||
    (!formula && !drawn) ||
    (Array.isArray(edit.lineLayout) && edit.lineLayout.length > 0) ||
    // A row plan carries the runs and the spacing it was measured with
    // (Tc, Tw, Tz), and no other whole-box format. A box of one style is
    // written whatever its characters: one its faces cannot draw (Polish,
    // Czech...) is drawn by the appearance the editor rendered of it
    
    (runWriter && drawn && (!pdfTextBoxDrawnRunsPortable(drawn) || !pdfSourceTextBoxTypographyPortable(typography)))
    // A source-bound box is clear (its run is hidden in the bytes), unless
    // the run paints nothing: then the box covers the scan under it with
    // the page's colour (PDF 0107; the server proves the run's mode).
    || (edit.sourceText !== undefined && (!sourceText || (backgroundColor && !pdfSourceTextBoxCoversInk(sourceText))))
    // A source-bound box is laid out by the portable row plan, which carries
    // the run's character spacing, word spacing and horizontal scale (its
    // rows are measured and drawn with them) and has no place for any other
    // of these formats.
    || (sourceText && !pdfSourceTextBoxTypographyPortable(typography))
  ) {
    return null;
  }
  return {
    x: geometry.x,
    y: geometry.y,
    width: geometry.width,
    height: geometry.height,
    text,
    ...(formula ? { formula: true as const } : {}),
    fontSize,
    color: runColor,
    align:
      edit.align === "center" || edit.align === "right" ? edit.align : "left",
    font,
    bold,
    italic,
    underline,
    ...(strikethrough ? { strikethrough: true } : {}),
    ...(edit.justify ? { justify: true } : {}),
    // A formula is set as its equation: no Word format has a place in it.
    ...(!formula && hasPdfTextBoxTypography(typography) ? { typography } : {}),
    ...(runWriter && drawn ? { runs: drawn } : {}),
    styleKey: styleKeyForEdit(edit),
    ...(backgroundColor ? { backgroundColor } : {}),
    ...(cover ? { cover } : {}),
    ...(angle ? { angle } : {}),
    ...(sourceText ? { sourceText } : {}),
  };
}

/** The formats a source-bound box's row plan carries: Tc, Tw and Tz. */
export function pdfSourceTextBoxTypographyPortable(
  typography: ReturnType<typeof readPdfTextBoxTypography>,
): boolean {
  // Align Text moves the plan's rows where they stand: the plan carries it.
  return Object.keys(typography).every((key) =>
    key === "characterSpacing" || key === "wordSpacing" || key === "horizontalScale" ||
    key === "verticalAlign");
}

export function isManagedNativeTextBoxEditEligible(
  edit: TextBoxEditLike,
): boolean {
  return textBoxPropertiesFromEdit(edit) !== null;
}

/** A box the writer draws on its row plan: one bound to source text, and
 * one drawn in runs (text-box-runs.ts: part of its text formatted, or an
 * Office family), whose rows the writer re-measures. */
export function managedNativeTextBoxNeedsRowPlan(edit: TextBoxEditLike): boolean {
  if (edit.sourceText) return true;
  // A formula is drawn as its picture: no row plan.
  if (effectiveTextBoxKind(edit) === "formula") return false;
  // Nor is a box set vertically: its rows are columns, which the writer
  // lays out itself in each run's format (_build_vertical_run_text_box_appearance).
  if ((edit as { writingMode?: string }).writingMode === "vertical-rl") return false;
  const drawn = resolvePdfTextBoxDrawnRuns(edit);
  return drawn !== null && pdfTextBoxDrawnRunsNeedRunWriter(edit, drawn);
}





export function isImportedNativeTextBoxEditEligible(
  edit: TextBoxEditLike,
): boolean {
  return textBoxPropertiesFromEdit(edit, { allowImportedSource: true }) !== null;
}

export function buildManagedNativeTextBoxRecord(
  edit: TextBoxEditLike,
  nativeAppearanceImage: string,
  portableLayout?: PdfPortableTextLayout,
): ManagedNativeTextBoxAnnotation | null {
  const id = typeof edit.id === "string" ? edit.id : "";
  const page = Number(edit.page);
  const properties = textBoxPropertiesFromEdit(edit);
  const image = sanitizeManagedNativeTextBoxImage(nativeAppearanceImage);
  if (!properties || !image || !encodeOpaqueTextBoxId(id)) return null;
  return { id, page: page + 1, ...properties, nativeAppearanceImage: image,
    ...(portableLayout ? {portableLayout} : {}) };
}

/**
 * Reconcile persisted native shadows with the latest text-box edit list.
 * Unsupported/source edits never become `/FreeText`; an existing managed box
 * that becomes unsupported is explicitly tombstoned so an old `/AP` cannot
 * remain visible behind the flattened fallback.
 */
/**
 * Whether a record's row plan is the plan of its frame on the page as it is
 * turned. A box's frame is a share of the page as it is shown and its letters
 * a share of that page's height (native-text-box-appearance.ts), so the plan
 * -- its width, height and letter size, in points -- is that of one turn of
 * the page: turned a quarter, a box on an A4 page is 1.41 times as wide and
 * its letters 0.71 times as tall. The writer checks the plan against the
 * frame of the turn the row names and refuses the row otherwise
 * (`pdf_portable_text_layout_invalid`), the whole page's write with it. A
 * record with no plan (the writer sets its rows itself) always fits.
 */
export function managedNativeTextBoxRowPlanFitsPage(
  record: ManagedNativeTextBoxRecord,
  pageSize: Readonly<{ width: number; height: number; rotation?: number }>,
): boolean {
  if (record.deleted || !record.portableLayout) return true;
  const quarter = ((Number(pageSize.rotation ?? 0) % 180) + 180) % 180 === 90;
  const shownWidth = quarter ? pageSize.height : pageSize.width;
  const shownHeight = quarter ? pageSize.width : pageSize.height;
  if (!(shownWidth > 0) || !(shownHeight > 0)) return true;
  const plan = record.portableLayout;
  const near = (left: number, right: number) => Math.abs(left - right) <= PDF_PORTABLE_LAYOUT_POINT_TOLERANCE;
  return near(plan.width, record.width * shownWidth) && near(plan.height, record.height * shownHeight) &&
    near(plan.fontSize, record.fontSize * shownHeight);
}

export async function reconcileManagedNativeTextBoxRecords(
  previous: readonly ManagedNativeTextBoxRecord[],
  edits: readonly TextBoxEditLike[],
  resolveAppearance: (edit: TextBoxEditLike) => Promise<string | {image:string; portableLayout:PdfPortableTextLayout; height?: number} | null>,
  options: {
    restoreMissingSourceText?: boolean;
    /** Whether a record is still the one its page is written from as the
     * page is now (managedNativeTextBoxRowPlanFitsPage): a record that is
     * not is drawn again though its box reads the same. */
    recordFits?: (record: ManagedNativeTextBoxAnnotation) => boolean;
  } = {},
): Promise<ManagedNativeTextBoxReconcileResult> {
  const appearanceFailureIds: string[] = [];
  const duplicateIds: string[] = [];
  const grownHeights = new Map<string, number>();
  const seenEditIds = new Set<string>();
  const eligible: TextBoxEditLike[] = [];
  for (const edit of edits) {
    const id = typeof edit.id === "string" ? edit.id : "";
    if (!id || !isManagedNativeTextBoxEditEligible(edit)) continue;
    if (seenEditIds.has(id)) {
      duplicateIds.push(id);
      continue;
    }
    seenEditIds.add(id);
    eligible.push(edit);
  }
  const previousById = new Map(previous.map((record) => [record.id, record]));
  const nextById = new Map<string, ManagedNativeTextBoxRecord>();
  for (const edit of eligible) {
    const id = edit.id as string;
    const existing = previousById.get(id);
    if (isManagedNativeTextBoxEquivalent(existing, edit) && (options.recordFits?.(existing) ?? true)) {
      nextById.set(id, existing);
      continue;
    }
    const appearance = await resolveAppearance(edit);
    // A frame grown down to hold its rows (a bolder or larger line that
    // wraps): the record is the box at that height, as WPS grows it.
    const grown = appearance && typeof appearance !== "string" && appearance.height !== undefined
      ? appearance.height : undefined;
    if (grown !== undefined) grownHeights.set(id, grown);
    const record = appearance
      ? buildManagedNativeTextBoxRecord(grown !== undefined ? { ...edit, ...pdfTextBoxGrownFrame(edit, grown) } : edit,
          typeof appearance === "string" ? appearance : appearance.image,
          typeof appearance === "string" ? undefined : appearance.portableLayout)
      : null;
    if (record) {
      nextById.set(id, options.restoreMissingSourceText && existing?.sourceText && !edit.sourceText
        ? {...record, sourceText:existing.sourceText, restoreSource:true} : record);
    } else {
      appearanceFailureIds.push(id);
      // Failure is not a delete command. Source-bound edits cannot fall back
      // to a flatten overlay; retain the current native owner for retry.
      if (existing && (edit.sourceText || existing.sourceText)) nextById.set(id, existing);
    }
  }
  for (const record of previous) {
    if (nextById.has(record.id)) continue;
    nextById.set(
      record.id,
      record.deleted
        ? record
        : { id: record.id, page: record.page, deleted: true,
            ...(record.sourceText ? {sourceText: record.sourceText,
              restoreSource: options.restoreMissingSourceText === true && !edits.some(edit => edit.id === record.id)} : {}),
          },
    );
  }
  return {
    records: [...nextById.values()],
    appearanceFailureIds,
    duplicateIds,
    ...(grownHeights.size ? { grownHeights } : {}),
  };
}

/**
 * Ensure every changed third-party `/FreeText` record carries the exact
 * appearance raster used as an integrity-bound input to native patchback.
 * Untouched sources and tombstones need no raster.
 */
export async function prepareImportedNativeTextBoxAppearances<
  T extends TextBoxEditLike,
>(
  records: readonly T[],
  resolveAppearance: (edit: T) => Promise<string | null>,
): Promise<ImportedNativeTextBoxAppearanceResult<T>> {
  const appearanceFailureIds: string[] = [];
  let changed = false;
  const next = await Promise.all(
    records.map(async (record) => {
      if (
        !isNativePdfFreeTextRecord(record) ||
        record.sourceAction !== "content_style_geometry"
      ) {
        return record;
      }
      if (
        sanitizeManagedNativeTextBoxImage(record.nativeAppearanceImage) !== null
      ) {
        return record;
      }
      const appearance = await resolveAppearance(record);
      const sanitized = sanitizeManagedNativeTextBoxImage(appearance);
      if (!sanitized) {
        appearanceFailureIds.push(String(record.id ?? ""));
        return record;
      }
      changed = true;
      return { ...record, nativeAppearanceImage: sanitized } as T;
    }),
  );
  return {
    records: changed ? next : (records as T[]),
    appearanceFailureIds,
  };
}

const EQUIVALENCE_EPSILON = 0.000_1;

/** JSON with every object's keys sorted: one spelling per value. */
function canonicalJsonValue(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value) ?? "null";
  if (Array.isArray(value)) return `[${value.map((item) => (item === undefined ? "null" : canonicalJsonValue(item))).join(",")}]`;
  const object = value as Record<string, unknown>;
  const keys = Object.keys(object).filter((key) => object[key] !== undefined).sort();
  return `{${keys.map((key) => `${JSON.stringify(key)}:${canonicalJsonValue(object[key])}`).join(",")}}`;
}

function sameJsonValue(left: unknown, right: unknown): boolean {
  return canonicalJsonValue(left) === canonicalJsonValue(right);
}

/** Fields a box carries for the live reader only, never saved. */
const RENDER_ONLY_BOX_FIELDS = new Set(["displayFontFamily", "sampleX", "sampleY", "sampleWidth", "sampleHeight"]);
function persistedShape(value: unknown): unknown {
  if (value === null || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map(persistedShape);
  return Object.fromEntries(Object.entries(value as Record<string, unknown>)
    .filter(([key]) => !RENDER_ONLY_BOX_FIELDS.has(key))
    .map(([key, entry]) => [key, persistedShape(entry)]));
}

/** Whether a box is the one the page saved, by value: what the database
 * gives back has its keys reordered (JSONB), and the live box carries
 * render-only fields the save strips. */
export function samePersistedPdfTextBoxEdit(saved: unknown, live: unknown): boolean {
  return saved !== undefined && sameJsonValue(persistedShape(saved), persistedShape(live));
}

/** The stationary covers of two boxes (PDF 0107): none on both, or one rect. */
function sameCover(
  left: ManagedNativeTextBoxAnnotation["cover"],
  right: ManagedNativeTextBoxAnnotation["cover"],
): boolean {
  if (!left || !right) return !left && !right;
  return sameFrameValue(left.x, right.x) && sameFrameValue(left.y, right.y) &&
    sameFrameValue(left.width, right.width) && sameFrameValue(left.height, right.height);
}

function nearlyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= EQUIVALENCE_EPSILON;
}






const FRAME_EQUIVALENCE_EPSILON = 0.000_5;

function sameFrameValue(left: number, right: number): boolean {
  return Math.abs(left - right) <= FRAME_EQUIVALENCE_EPSILON;
}






function sourceTextIdentity(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) return value ?? null;
  return Object.fromEntries(Object.entries(value as Record<string, unknown>)
    .filter(([key]) => key !== "expectedAppearance"));
}

export function isManagedNativeTextBoxEquivalent(
  record: ManagedNativeTextBoxRecord | undefined,
  edit: TextBoxEditLike,
): record is ManagedNativeTextBoxAnnotation {
  if (!record || record.deleted) return false;
  const properties = textBoxPropertiesFromEdit(edit);
  if (!properties) return false;
  return (
    record.id === edit.id &&
    record.page === Number(edit.page) + 1 &&
    record.text === properties.text &&
    Boolean(record.formula) === Boolean(properties.formula) &&
    record.color === properties.color &&
    (record.backgroundColor ?? "") === (properties.backgroundColor ?? "") &&
    sameCover(record.cover, properties.cover) &&
    record.align === properties.align &&
    record.font === properties.font &&
    record.bold === properties.bold &&
    record.italic === properties.italic &&
    record.underline === properties.underline &&
    Boolean(record.strikethrough) === Boolean(properties.strikethrough) &&
    Boolean(record.justify) === Boolean(properties.justify) &&
    encodePdfTextBoxTypographyField(record.typography ?? {}, record.fontSize) ===
      encodePdfTextBoxTypographyField(properties.typography ?? {}, properties.fontSize) &&
    (record.runs ? encodePdfTextBoxRunsField(record.runs) : null) ===
      (properties.runs ? encodePdfTextBoxRunsField(properties.runs) : null) &&
    record.styleKey === properties.styleKey &&
    // By value, not spelling: the page's edits come back from the database
    // with their nested keys reordered (JSONB), the record keeps its own
    // order, and every source-bound box read as changed after a reload. All
    // were redrawn on the next save, one without its face wrapped, and that
    
    sameJsonValue(sourceTextIdentity(record.sourceText), sourceTextIdentity(properties.sourceText)) &&
    nearlyEqual(record.angle ?? 0, properties.angle ?? 0) &&
    sameFrameValue(record.x, properties.x) &&
    sameFrameValue(record.y, properties.y) &&
    sameFrameValue(record.width, properties.width) &&
    sameFrameValue(record.height, properties.height) &&
    nearlyEqual(record.fontSize, properties.fontSize) &&
    sanitizeManagedNativeTextBoxImage(record.nativeAppearanceImage) !== null
  );
}

/** Development: the fields on which a box differs from its managed record
 * (isManagedNativeTextBoxEquivalent's terms). Never throws. */
export function managedNativeTextBoxDifferences(
  record: ManagedNativeTextBoxRecord | undefined,
  edit: TextBoxEditLike,
): string[] {
  try {
    if (!record) return ["record"];
    if (record.deleted) return ["deleted"];
    const properties = textBoxPropertiesFromEdit(edit);
    if (!properties) return ["properties"];
    const out: string[] = [];
    const left = record as unknown as Record<string, unknown>;
    const right = properties as unknown as Record<string, unknown>;
    for (const key of ["text", "formula", "color", "backgroundColor", "align", "font", "bold", "italic", "underline",
      "strikethrough", "justify", "styleKey", "angle", "x", "y", "width", "height", "fontSize"]) {
      const a = left[key];
      const b = right[key];
      if (typeof a === "number" && typeof b === "number"
        ? !(["x", "y", "width", "height"].includes(key) ? sameFrameValue(a, b) : nearlyEqual(a, b)) : (a ?? null) !== (b ?? null) &&
        !(a === undefined && b === false) && !(a === false && b === undefined)) {
        out.push(`${key}:${JSON.stringify(a)}≠${JSON.stringify(b)}`.slice(0, 160));
      }
    }
    if (!sameCover(record.cover, properties.cover)) {
      out.push(`cover:${JSON.stringify(record.cover ?? null)}≠${JSON.stringify(properties.cover ?? null)}`.slice(0, 200));
    }
    if (!sameJsonValue(sourceTextIdentity(record.sourceText), sourceTextIdentity(properties.sourceText))) {
      out.push(`sourceText:${JSON.stringify(record.sourceText ?? null).slice(0, 400)}≠${JSON.stringify(properties.sourceText ?? null).slice(0, 400)}`);
    }
    if (encodePdfTextBoxTypographyField(record.typography ?? {}, record.fontSize) !==
      encodePdfTextBoxTypographyField(properties.typography ?? {}, properties.fontSize)) out.push("typography");
    if ((record.runs ? encodePdfTextBoxRunsField(record.runs) : null) !==
      (properties.runs ? encodePdfTextBoxRunsField(properties.runs) : null)) out.push("runs");
    if (sanitizeManagedNativeTextBoxImage(record.nativeAppearanceImage) === null) out.push("appearance");
    return out;
  } catch (error) {
    return [`failed:${String(error)}`];
  }
}

/** The lowest first baseline a box's row plan may have: a box set in runs
 * may start with a row of runs smaller than the box's own size. */
export function pdfTextBoxRunsBaselineOptions(
  record: { runs?: readonly PdfTextBoxDrawnRun[]; fontSize: number;
    typography?: { columnCount?: number; columnSpacing?: number } },
): { minimumBaselineScale?: number; columns?: { count: number; gapEm: number } } {
  if (!(record.fontSize > 0)) return {};
  // A box in columns: its rows start again at the top of each column.
  const count = record.typography?.columnCount ?? 1;
  const columns = count > 1 ? { columns: { count, gapEm: (record.typography?.columnSpacing ?? 0) / record.fontSize } } : {};
  if (!record.runs?.length) return columns;
  return {
    minimumBaselineScale: Math.min(...record.runs.map((run) => run.fontSize)) / record.fontSize,
    ...columns,
  };
}

/** The token a record's source binding is carried in; null for a record
 * with none, or one that does not read as a binding. */
function sourceBindingToken(record: Pick<ManagedNativeTextBoxRecord, "sourceText" | "restoreSource">): string | null {
  const binding = record.sourceText ? readPdfSourceTextBoxBinding(record.sourceText) : null;
  return binding ? bytesToBase64Url(new TextEncoder().encode(JSON.stringify({
    ...binding, ...(record.restoreSource ? {restore: true} : {}),
  }))) : null;
}









export function managedNativeTextBoxSourceBindingTooLong(
  record: Pick<ManagedNativeTextBoxRecord, "sourceText" | "restoreSource">,
): boolean {
  const token = sourceBindingToken(record);
  return token !== null && token.length > MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS;
}

function nativeTextBoxDirective(
  record: ManagedNativeTextBoxRecord,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const nativeId = encodeOpaqueTextBoxId(record.id);
  if (!nativeId || !isValidPage(record.page)) return null;
  const binding = record.sourceText ? readPdfSourceTextBoxBinding(record.sourceText) : null;
  if (record.sourceText && !binding) return null;
  const sourceToken = sourceBindingToken(record);
  // Refuse before HTTP instead of sending a body the source writer cannot
  // decode, then reporting its generic identity mismatch after JSON saved.
  if (sourceToken && sourceToken.length > MAX_PDF_SOURCE_TEXT_BOX_TOKEN_CHARS) return null;
  if (record.deleted) {
    const legacy = `${MANAGED_NATIVE_TEXT_BOX_PREFIX}delete|${nativeId}]]`;
    return sourceToken
      ? `[[MRNOTE_NATIVE_TEXT_BOX_V4|${sourceToken}|${legacy.slice(2)}`
      : legacy;
  }
  const textToken = encodeTextToken(record.text);
  const image = sanitizeManagedNativeTextBoxImage(record.nativeAppearanceImage);
  const imageMatch = image ? PNG_DATA_URL_PATTERN.exec(image) : null;
  if (
    !textToken ||
    !imageMatch ||
    !/^[0-9A-Fa-f]{16}$/u.test(record.styleKey)
  ) {
    return null;
  }
  const angle = normalizedObjectRotation(record.angle);
  if (angle === null) return null;
  const backgroundColor = record.backgroundColor
    ? normalizeColor(record.backgroundColor)
    : undefined;
  if (record.backgroundColor && !backgroundColor) return null;
  // A stationary cover is a filled box's over a run that paints nothing
  // only (PDF 0107); the writer proves the run again.
  const cover = record.cover ? parseGeometry(normalizedGeometryTokens(record.cover)) : null;
  if (record.cover && (!cover || !backgroundColor || !binding || record.restoreSource ||
    !pdfSourceTextBoxCoversInk(binding))) return null;
  const version = backgroundColor ? 3 : angle ? 2 : 1;
  const typographyField = record.typography
    ? encodePdfTextBoxTypographyField(record.typography, record.fontSize)
    : null;
  if (record.typography && hasPdfTextBoxTypography(record.typography) && !typographyField) return null;
  // A box drawn in runs is written only with its row plan: the runs' rows
  // are the window's measure of them, which the writer re-measures.
  const runsField = record.runs ? encodePdfTextBoxRunsField(record.runs) : null;
  // A box set vertically is laid out by the writer, in columns: no plan.
  const vertical = record.typography?.writingMode === "vertical-rl";
  if (record.runs && (!runsField || (!record.portableLayout && !vertical))) return null;
  // A formula is one equation's LaTeX: nothing else sets it.
  if (record.formula && (record.runs || typographyField || binding || record.portableLayout)) return null;
  const prefix =
    version === 3
      ? MANAGED_NATIVE_TEXT_BOX_V3_PREFIX
      : version === 2
        ? MANAGED_NATIVE_TEXT_BOX_V2_PREFIX
        : MANAGED_NATIVE_TEXT_BOX_PREFIX;
  const directive = (
    `${prefix}upsert|${nativeId}|${record.justify ? "justify" : record.align}|` +
    `${record.font}|${record.bold ? 1 : 0}|${record.italic ? 1 : 0}|` +
    `${(record.underline ? 1 : 0) | (record.strikethrough ? 2 : 0)}|${numberToken(record.fontSize)}|` +
    `${record.color.slice(1)}|${backgroundColor ? `${backgroundColor.slice(1)}|` : ""}` +
    `${record.styleKey}|${textToken}|${runsField ? `${runsField}|` : ""}` +
    `${typographyField ? `${typographyField}|` : ""}${record.formula ? "$tex|" : ""}` +
    `png|${imageMatch[1]}|${rotation}|` +
    `${version > 1 ? `${numberToken(angle)}|` : ""}` +
    `${normalizedGeometryTokens(record)}${cover ? `,${normalizedGeometryTokens(cover)}` : ""}]]`
  );
  const sourceDirective = sourceToken
    ? `[[MRNOTE_NATIVE_TEXT_BOX_V4|${sourceToken}|${directive.slice(2)}`
    : directive;
  if (!record.portableLayout) return sourceDirective;
  const layout = readUnverifiedPdfPortableTextLayout(record.portableLayout, record.text,
    pdfTextBoxRunsBaselineOptions(record));
  if (!layout) return null;
  const token = bytesToBase64Url(new TextEncoder().encode(JSON.stringify(layout)));
  return token.length <= 262_144 ? `[[MRNOTE_NATIVE_TEXT_BOX_V5|${token}|${sourceDirective.slice(2)}` : null;
}

function importedNativeTextBoxDirective(
  record: TextBoxEditLike,
  rotation: 0 | 90 | 180 | 270 | "auto",
): string | null {
  const rectangleDifferences = pdfFreeTextRectangleDifferencesV2Enabled();
  const source = sanitizeNativePdfFreeTextAnnotationSource(record.nativeSource);
  const page = Number(record.page);
  if (!source || !Number.isInteger(page) || page < 0 || page >= 10_000) {
    return null;
  }
  if (record.deleted === true) {
    const deletePrefix = rectangleDifferences
      ? IMPORTED_NATIVE_TEXT_BOX_RD_PREFIX
      : IMPORTED_NATIVE_TEXT_BOX_PREFIX;
    return `${deletePrefix}delete|${source.objectRef}]]`;
  }
  // Untouched source objects already exist in the immutable input package.
  if (record.sourceAction !== "content_style_geometry") return null;
  const properties = textBoxPropertiesFromEdit(record, {
    allowImportedSource: true,
  });
  const image = sanitizeManagedNativeTextBoxImage(record.nativeAppearanceImage);
  const imageMatch = image ? PNG_DATA_URL_PATTERN.exec(image) : null;
  const textToken = properties ? encodeTextToken(properties.text) : null;
  const angle = properties ? normalizedObjectRotation(properties.angle) : null;
  if (
    !properties ||
    !imageMatch ||
    !textToken ||
    angle === null ||
    angle % 90 !== 0 ||
    // The source-object writer has no strikethrough or justify; refusing
    // keeps a format the reader chose from being dropped on save.
    properties.strikethrough ||
    properties.justify ||
    properties.typography ||
    properties.runs
  ) {
    return null;
  }
  const prefix = angle
    ? rectangleDifferences
      ? IMPORTED_NATIVE_TEXT_BOX_RD_ROTATION_PREFIX
      : IMPORTED_NATIVE_TEXT_BOX_ROTATION_PREFIX
    : rectangleDifferences
      ? IMPORTED_NATIVE_TEXT_BOX_RD_PREFIX
      : IMPORTED_NATIVE_TEXT_BOX_PREFIX;
  // Never persist the overlay's fallback alignment as if the user chose it.
  const align = record.alignFromSource === true ? "source" : properties.align;
  return (
    `${prefix}upsert|${source.objectRef}|` +
    `${align}|${properties.font}|${properties.bold ? 1 : 0}|` +
    `${properties.italic ? 1 : 0}|${properties.underline ? 1 : 0}|` +
    `${numberToken(properties.fontSize)}|${properties.color.slice(1)}|` +
    `${properties.styleKey}|${textToken}|png|${imageMatch[1]}|${rotation}|` +
    `${angle ? `${numberToken(angle)}|` : ""}` +
    `${normalizedGeometryTokens(properties)}]]`
  );
}

export function planManagedNativeTextBoxAnnotations(
  records: readonly ManagedNativeTextBoxRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeTextBoxPlan {
  if (records.length === 0) return { annotations: [], supported: true };
  const ids = new Set<string>();
  const annotations: ManagedNativeTextBoxAnnotationShadow[] = [];
  let directiveCharacters = 0;
  for (const record of records) {
    if (ids.has(record.id)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(record.id);
    const text = nativeTextBoxDirective(
      record,
      rotations.get(record.page) ?? "auto",
    );
    if (!text) {
      return { annotations: [], supported: false, reason: "invalid_text_box" };
    }
    directiveCharacters += text.length;
    if (directiveCharacters > MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_TOTAL_CHARS) {
      return {
        annotations: [],
        supported: false,
        reason: "image_total_limit",
      };
    }
    annotations.push({
      id: `__mrnote_native_text_box__${record.id}`,
      text,
      page: record.page,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_TEXT_BOX_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

/**
 * Build exact-object update/delete directives for third-party `/FreeText`
 * annotations. Untouched discovery records intentionally emit no directive.
 */
export function planImportedNativeTextBoxAnnotations(
  records: readonly TextBoxEditLike[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ImportedNativeTextBoxPlan {
  const sourceRecords = records.filter(
    (record) => record.nativeSource !== undefined,
  );
  if (sourceRecords.length === 0) {
    return { annotations: [], supported: true };
  }
  const ids = new Set<string>();
  const objectRefs = new Set<string>();
  const annotations: ManagedNativeTextBoxAnnotationShadow[] = [];
  let directiveCharacters = 0;
  for (const record of sourceRecords) {
    const source = sanitizeNativePdfFreeTextAnnotationSource(
      record.nativeSource,
    );
    const id = typeof record.id === "string" ? record.id : "";
    const pageIndex = Number(record.page);
    if (
      !source ||
      !id ||
      !Number.isInteger(pageIndex) ||
      pageIndex < 0 ||
      pageIndex >= 10_000
    ) {
      return {
        annotations: [],
        supported: false,
        reason: "invalid_text_box",
      };
    }
    if (ids.has(id) || objectRefs.has(source.objectRef)) {
      return { annotations: [], supported: false, reason: "duplicate_id" };
    }
    ids.add(id);
    objectRefs.add(source.objectRef);
    const pageNumber = pageIndex + 1;
    const text = importedNativeTextBoxDirective(
      record,
      rotations.get(pageNumber) ?? "auto",
    );
    if (!text) {
      if (record.deleted !== true && record.sourceAction === undefined) {
        continue;
      }
      return {
        annotations: [],
        supported: false,
        reason: "invalid_text_box",
      };
    }
    directiveCharacters += text.length;
    if (directiveCharacters > MAX_MANAGED_NATIVE_TEXT_BOX_IMAGE_TOTAL_CHARS) {
      return {
        annotations: [],
        supported: false,
        reason: "image_total_limit",
      };
    }
    annotations.push({
      id: `__mrnote_native_text_box_source__${source.objectRef}`,
      text,
      page: pageNumber,
      rect: [24, 24, 44, 44],
    });
    if (annotations.length > MAX_MANAGED_NATIVE_TEXT_BOX_PATCHES) {
      return { annotations: [], supported: false, reason: "patch_limit" };
    }
  }
  return { annotations, supported: true };
}

export function isManagedNativeTextBoxAnnotationText(
  value: unknown,
): value is string {
  return (
    typeof value === "string" &&
    (value.startsWith(MANAGED_NATIVE_TEXT_BOX_PREFIX) ||
      value.startsWith("[[MRNOTE_NATIVE_TEXT_BOX_V4|") ||
      value.startsWith("[[MRNOTE_NATIVE_TEXT_BOX_V5|") ||
      value.startsWith(MANAGED_NATIVE_TEXT_BOX_V2_PREFIX) ||
      value.startsWith(MANAGED_NATIVE_TEXT_BOX_V3_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_TEXT_BOX_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_TEXT_BOX_RD_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_TEXT_BOX_ROTATION_PREFIX) ||
      value.startsWith(IMPORTED_NATIVE_TEXT_BOX_RD_ROTATION_PREFIX))
  );
}

export function parseManagedNativeTextBoxDirective(
  value: unknown,
  /** A stationary cover is read only inside its source binding (V4), whose
   * run paints nothing (PDF 0107). */
  options: { allowCover?: boolean } = {},
): ParsedManagedNativeTextBoxDirective | null {
  if (
    typeof value !== "string" ||
    value.length > MAX_MANAGED_NATIVE_TEXT_BOX_DIRECTIVE_CHARS
  ) {
    return null;
  }
  if (value.startsWith("[[MRNOTE_NATIVE_TEXT_BOX_V5|")) {
    const start = "[[MRNOTE_NATIVE_TEXT_BOX_V5|".length;
    const end = value.indexOf("|", start);
    const token = value.slice(start, end);
    const legacy = `[[${value.slice(end + 1)}`;
    if (end < start || token.length > 262_144 || !/^[A-Za-z0-9_-]+$/.test(token) ||
      ![MANAGED_NATIVE_TEXT_BOX_PREFIX,MANAGED_NATIVE_TEXT_BOX_V2_PREFIX,MANAGED_NATIVE_TEXT_BOX_V3_PREFIX,"[[MRNOTE_NATIVE_TEXT_BOX_V4|"].some(prefix=>legacy.startsWith(prefix))) return null;
    try {
      const bytes = Uint8Array.from(atob(token.replaceAll("-", "+").replaceAll("_", "/")), ch=>ch.charCodeAt(0));
      if (bytesToBase64Url(bytes) !== token) return null;
      const parsed = parseManagedNativeTextBoxDirective(legacy, options);
      if (!parsed?.textBox || parsed.operation !== "upsert") return null;
      const layout = readUnverifiedPdfPortableTextLayout(JSON.parse(new TextDecoder("utf-8",{fatal:true}).decode(bytes)), parsed.textBox.text,
        pdfTextBoxRunsBaselineOptions(parsed.textBox));
      return layout ? {...parsed,textBox:{...parsed.textBox,portableLayout:layout}} : null;
    } catch { return null; }
  }
  if (value.startsWith("[[MRNOTE_NATIVE_TEXT_BOX_V4|")) {
    const prefixLength = "[[MRNOTE_NATIVE_TEXT_BOX_V4|".length;
    const end = value.indexOf("|", prefixLength);
    if (end < 0 || end > 65000) return null;
    try {
      const token = value.slice(prefixLength, end);
      const payload = JSON.parse(new TextDecoder().decode(
        Uint8Array.from(atob(token.replaceAll("-", "+").replaceAll("_", "/")), ch => ch.charCodeAt(0)),
      ));
      const binding = readPdfSourceTextBoxBinding(payload);
      const legacy = `[[${value.slice(end + 1)}`;
      if (!binding || !legacy.startsWith(MANAGED_NATIVE_TEXT_BOX_PREFIX) &&
        !legacy.startsWith(MANAGED_NATIVE_TEXT_BOX_V2_PREFIX) &&
        // A filled (V3) box over source text: one restoring its run, or one
        // covering the scan under a run that paints nothing (PDF 0107).
        !((payload.restore === true || pdfSourceTextBoxCoversInk(binding)) &&
          legacy.startsWith(MANAGED_NATIVE_TEXT_BOX_V3_PREFIX))) return null;
      const parsed = parseManagedNativeTextBoxDirective(legacy, { allowCover: payload.restore !== true &&
        pdfSourceTextBoxCoversInk(binding) });
      return parsed?.textBox ? { ...parsed, textBox: { ...parsed.textBox, ...(payload.restore === true ? {} : {sourceText: binding}) } }
        : parsed?.operation === "delete" ? {...parsed, sourceText: binding, restoreSource: payload.restore === true} : null;
    } catch { return null; }
  }
  const deleted = DELETE_PATTERN.exec(value);
  if (deleted) {
    const nativeId = deleted[1];
    const id = decodeOpaqueTextBoxId(nativeId);
    return id ? { operation: "delete", id, nativeId } : null;
  }
  // The optional runs field (`^`), typography field (`~`) and formula field
  // (`$tex`) sit, in that order, between the text token and `png`; no sigil
  // can occur in a neighbour, so each is lifted out unambiguously.
  const formulaMatch = /\|\$tex\|png\|/u.exec(value);
  const formulaless = formulaMatch
    ? `${value.slice(0, formulaMatch.index)}|png|${value.slice(formulaMatch.index + formulaMatch[0].length)}`
    : value;
  const runsMatch = /\|(\^[A-Za-z0-9_-]+)\|(?=(?:~[A-Za-z0-9_-]{1,2048}\|)?png\|)/u.exec(formulaless);
  const runlessValue = runsMatch
    ? `${formulaless.slice(0, runsMatch.index)}${formulaless.slice(runsMatch.index + runsMatch[0].length - 1)}`
    : formulaless;
  const typographyMatch = /\|(~[A-Za-z0-9_-]{1,2048})\|png\|/u.exec(runlessValue);
  const legacyValue = typographyMatch
    ? `${runlessValue.slice(0, typographyMatch.index)}|png|${runlessValue.slice(typographyMatch.index + typographyMatch[0].length)}`
    : runlessValue;
  const v3Match = V3_UPSERT_PATTERN.exec(legacyValue);
  const v2Match = V2_UPSERT_PATTERN.exec(legacyValue);
  const match = v3Match ?? v2Match ?? UPSERT_PATTERN.exec(legacyValue);
  if (!match) return null;
  const nativeId = match[1];
  const id = decodeOpaqueTextBoxId(nativeId);
  const fontSize = Number(match[7]);
  const typography = typographyMatch
    ? decodePdfTextBoxTypographyField(typographyMatch[1], fontSize)
    : null;
  if (typographyMatch && !typography) return null;
  const styleKeyIndex = v3Match ? 10 : 9;
  const textIndex = v3Match ? 11 : 10;
  const imageIndex = v3Match ? 12 : 11;
  const rotationIndex = v3Match ? 13 : 12;
  const angleIndex = v3Match ? 14 : v2Match ? 13 : null;
  const geometryIndex = v3Match ? 15 : v2Match ? 14 : 13;
  const text = decodeTextToken(match[textIndex]);
  const image = sanitizeManagedNativeTextBoxImage(
    `data:image/png;base64,${match[imageIndex]}`,
  );
  const angle = angleIndex ? normalizedObjectRotation(match[angleIndex]) : 0;
  // A V3 box over a run that paints nothing may carry its stationary cover
  // as four more numbers after its own (PDF 0107).
  const geometryTokens = match[geometryIndex].split(",");
  const covered = v3Match && options.allowCover === true && geometryTokens.length === 8;
  const geometry = parseGeometry(covered ? geometryTokens.slice(0, 4).join(",") : match[geometryIndex]);
  const cover = covered ? parseGeometry(geometryTokens.slice(4).join(",")) : null;
  if (covered && !cover) return null;
  const runs = runsMatch && text ? decodePdfTextBoxRunsField(runsMatch[1], text) : null;
  if (
    (runsMatch && !runs) ||
    (formulaMatch && (runsMatch || typographyMatch)) ||
    !id ||
    !text ||
    !image ||
    !geometry ||
    angle === null ||
    !Number.isFinite(fontSize) ||
    fontSize <= 0 ||
    fontSize > 1
  ) {
    return null;
  }
  return {
    operation: "upsert",
    id,
    nativeId,
    rotation:
      match[rotationIndex] === "auto"
        ? "auto"
        : (Number(match[rotationIndex]) as 0 | 90 | 180 | 270),
    textBox: {
    ...geometry,
      text,
      fontSize,
      color: `#${match[8].toUpperCase()}`,
      ...(v3Match
        ? { backgroundColor: `#${match[9].toUpperCase()}` }
        : {}),
      ...(cover ? { cover } : {}),
      align: match[2] === "justify" ? "left" : (match[2] as "left" | "center" | "right"),
      ...(match[2] === "justify" ? { justify: true } : {}),
      font: match[3] as "serif" | "sans" | "mono",
      bold: match[4] === "1",
      italic: match[5] === "1",
      underline: (Number(match[6]) & 1) === 1,
      ...((Number(match[6]) & 2) === 2 ? { strikethrough: true } : {}),
      ...(typography ? { typography } : {}),
      ...(runs ? { runs } : {}),
      ...(formulaMatch ? { formula: true as const } : {}),
      styleKey: match[styleKeyIndex].toLowerCase(),
      nativeAppearanceImage: image,
      ...(angle ? { angle } : {}),
    },
  };
}

export function parseManagedNativeTextBoxShadow(
  value: unknown,
): ManagedNativeTextBoxRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseManagedNativeTextBoxDirective(row.text);
  if (!parsed) return null;
  if (parsed.operation === "delete") {
    return { id: parsed.id, page, deleted: true,
      ...(parsed.sourceText ? {sourceText: parsed.sourceText, restoreSource: parsed.restoreSource} : {}),
    };
  }
  return parsed.textBox ? { id: parsed.id, page, ...parsed.textBox } : null;
}

export function parseImportedNativeTextBoxDirective(
  value: unknown,
): ParsedImportedNativeTextBoxDirective | null {
  if (
    typeof value !== "string" ||
    value.length > MAX_MANAGED_NATIVE_TEXT_BOX_DIRECTIVE_CHARS
  ) {
    return null;
  }
  const rdDeleted = IMPORTED_RD_DELETE_PATTERN.exec(value);
  const deleted = rdDeleted ?? IMPORTED_DELETE_PATTERN.exec(value);
  if (deleted) {
    return {
      operation: "delete",
      objectRef: deleted[1],
      sourceVersion: rdDeleted ? 2 : 1,
    };
  }
  const rdRotationMatch = IMPORTED_RD_ROTATION_UPSERT_PATTERN.exec(value);
  const rotationMatch = IMPORTED_ROTATION_UPSERT_PATTERN.exec(value);
  const rdMatch = IMPORTED_RD_UPSERT_PATTERN.exec(value);
  const match =
    rdRotationMatch ??
    rotationMatch ??
    rdMatch ??
    IMPORTED_UPSERT_PATTERN.exec(value);
  if (!match) return null;
  const sourceVersion: 1 | 2 | 3 | 4 = rdRotationMatch
    ? 4
    : rotationMatch
      ? 3
      : rdMatch
        ? 2
        : 1;
  const fontSize = Number(match[7]);
  const text = decodeTextToken(match[10]);
  const image = sanitizeManagedNativeTextBoxImage(
    `data:image/png;base64,${match[11]}`,
  );
  const angle = sourceVersion >= 3 ? normalizedObjectRotation(match[13]) : 0;
  const geometry = parseGeometry(match[sourceVersion >= 3 ? 14 : 13]);
  if (
    !text ||
    !image ||
    !geometry ||
    angle === null ||
    !Number.isFinite(fontSize) ||
    fontSize <= 0 ||
    fontSize > 1
  ) {
    return null;
  }
  const font = match[3] as "serif" | "sans" | "mono";
  const inheritsAlign = match[2] === "source";
  return {
    operation: "upsert",
    objectRef: match[1],
    sourceVersion,
    rotation:
      match[12] === "auto" ? "auto" : (Number(match[12]) as 0 | 90 | 180 | 270),
    textBox: {
      ...geometry,
      text,
      fontSize,
      color: `#${match[8].toUpperCase()}`,
      // An inherited `/Q` reopens as the same overlay fallback + flag, so the
      // next save carries the author's alignment through again.
      align: inheritsAlign ? "left" : (match[2] as "left" | "center" | "right"),
      ...(inheritsAlign ? { alignFromSource: true as const } : {}),
      fontFamily:
        font === "mono"
          ? "monospace"
          : font === "serif"
            ? "serif"
            : "sans-serif",
      sourceFontName:
        font === "mono"
          ? "Courier"
          : font === "serif"
            ? "Times-Roman"
            : "Helvetica",
      bold: match[4] === "1",
      italic: match[5] === "1",
      underline: match[6] === "1",
      nativeAppearanceImage: image,
      ...(angle ? { angle } : {}),
    },
  };
}

/**
 * Recover an edited source record (including its integrity-bound appearance)
 * from an inert annotation row after a page reload. Delete rows are restored
 * from `textBoxEdits`, which retains their last geometry for undo/UI history.
 */
export function parseImportedNativeTextBoxShadow(
  value: unknown,
): NativePdfFreeTextAnnotation | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!isValidPage(page)) return null;
  const parsed = parseImportedNativeTextBoxDirective(row.text ?? row.comment);
  if (!parsed || parsed.operation !== "upsert" || !parsed.textBox) return null;
  const { objectRef, textBox } = parsed;
  const id = `pdf-native-freetext:${page}:${objectRef}`;
  return {
    id,
    page: page - 1,
    ...textBox,
    created: true,
    originX: textBox.x,
    originY: textBox.y,
    originWidth: textBox.width,
    originHeight: textBox.height,
    nativeSource: {
      kind: NATIVE_PDF_FREE_TEXT_SOURCE_KIND,
      objectRef,
      subtype: "FreeText",
    },
    sourceAction: "content_style_geometry",
  };
}
