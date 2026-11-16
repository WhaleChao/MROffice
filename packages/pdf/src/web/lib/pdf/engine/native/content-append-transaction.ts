/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Native content append: insert text, an image or a vector path as real page
 * content by appending one self-contained `q … Q` block to the tail of the
 * page's single content stream. Resources the block names are installed
 * through the shared clone-on-write Resources path of the content-stream
 * transaction; undo removes the exact appended bytes, redo re-appends them
 * against the resources that already resolve on the page.
 */

import { LATIN1_EXACT } from "./latin1";
import { MAX_TEXT_CODE_UNITS, PDF_NATIVE_CONTENT_APPEND_CONTROL_IDS, PDF_NATIVE_CONTENT_APPEND_FEATURE_IDS, TEXT_ROTATION_MATRICES, buildPdfNativeContentAppendCommand, createPdfNativeContentAppendOwnershipManifest, invalid, pdfNativeContentAppendColorFromHex, pdfNativeContentAppendDisplayPoint, pdfNativeContentAppendDisplayRectMatrix, pdfNativeContentAppendDisplaySize, pdfNativeContentAppendFeatureId, pdfNativeContentAppendPayloadKind, readMatrix, readPdfNativeContentAppendPayload } from "./content-append-commands";
export { PDF_NATIVE_CONTENT_APPEND_CONTROL_IDS, PDF_NATIVE_CONTENT_APPEND_FEATURE_IDS, buildPdfNativeContentAppendCommand, createPdfNativeContentAppendOwnershipManifest, pdfNativeContentAppendColorFromHex, pdfNativeContentAppendDisplayPoint, pdfNativeContentAppendDisplayRectMatrix, pdfNativeContentAppendDisplaySize, pdfNativeContentAppendFeatureId, pdfNativeContentAppendPayloadKind, readPdfNativeContentAppendPayload };
import {
  MRAI_NATIVE_CONTENT_APPEND_ACTION,
} from "./vocabulary";
import { MRAI_NATIVE_BOX_MARK_TAG } from "./box-mark";
import { PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM } from "../../native-synthetic-style";
import { buildPdfNativeSfntSubset, type PdfNativeSfntGlyphCopy } from "./sfnt/sfnt-subset";
import { readSfntProgram } from "@mrnote/core/web/lib/pdf/engine/native/sfnt/sfnt-reader.ts";
import { PDF_COMPLEX_SCRIPT_FONTS } from "../../unicode-font-asset";
export {
  MRAI_NATIVE_CONTENT_APPEND_ACTION,
};
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { readPdfCommandOwnershipManifest } from "../command-ownership";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "../contracts";
import {
  MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
  inspectMraiNativeContentAppendPage,
  inspectMraiNativeEmbeddedImageReplacement,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  skipMraiNativeInlineImage,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentAppendAddition,
  type PdfNativeContentAppendCategory,
  type PdfNativeContentAppendKind,
  type PdfNativeContentAppendPageInspection,
  type PdfNativeContentAppendPlan,
  type PdfNativeContentAppendReusedResource,
  type PdfNativePreparedContentStreamTransaction,
} from "./content-stream-transaction";
import {
  buildPdfNativeFontExtensionPackage,
  buildPdfNativeShapedFontPackage,
  measurePdfNativeFontExtensionText,
  serializePdfNativeFontExtensionPackage,
  type PdfNativeFontExtensionWirePackage,
} from "./font-extension-package";

export { MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION };


const MAX_BLOCK_BYTES = 60 * 1024;
const MAX_RESTORED_DEPTH = 64;
const NUMBER_PATTERN = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
const IDENTITY: PdfNativeContentAppendMatrix = Object.freeze([1, 0, 0, 1, 0, 0]);
/** Unit vector, in PDF user space, that points down the reader's screen. */
const SCREEN_DOWN_DIRECTIONS: Readonly<
  Record<0 | 90 | 180 | 270, readonly [number, number]>
> = Object.freeze({
  0: [0, -1],
  90: [1, 0],
  180: [0, 1],
  270: [-1, 0],
});

export type PdfNativeContentAppendMatrix = readonly [
  number,
  number,
  number,
  number,
  number,
  number,
];

export interface PdfNativeContentAppendColor {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  /** Said as DeviceCMYK ink instead (a pasted copy of a CMYK paint, whose
   * RGB above is only how it reads): `r`/`g`/`b` stay its naive RGB. */
  readonly cmyk?: readonly [number, number, number, number];
}

export type PdfNativeContentAppendRotation = 0 | 90 | 180 | 270;

export interface PdfNativeContentAppendTextContent {
  readonly kind: "text";
  /** Lines separated by `\n`; each line becomes one Tj on its own baseline. */
  readonly text: string;
  readonly fontSizePoints: number;
  readonly color: PdfNativeContentAppendColor;
  /**
   * PDF user-space point of the first line box's top-left corner as the
   * reader sees it; `rotation` is the page's /Rotate so the text stays
   * upright on screen (0 keeps the plain `1 0 0 1 x y Tm` form).
   */
  readonly x: number;
  readonly top: number;
  readonly rotation?: PdfNativeContentAppendRotation;
  readonly lineHeightPoints: number;
  /**
   * The words set bold: the bundled face has one weight, so its outlines
   * are stroked in the text's colour as well as filled, as a line's own
   * bold is drawn in place where its font has no bold face
   * (PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM).
   */
  readonly bold?: boolean;
  /**
   * Each line as it prints, left to right, where a line holds words of a
   * complex script (Arabic, Hebrew, Thai, Devanagari, Bengali): the API's
   * shaping of it (`/api/v1/pdf-text-shape`, app.services.pdf_text_shaping),
   * a plain piece set in the bundled face, a shaped one in its packaged face
   * by HarfBuzz's glyphs. Absent: every line is set in the bundled face.
   */
  readonly shapedLines?: readonly (readonly PdfNativeShapedTextPiece[])[];
}

/** One piece of a shaped line (the API's wire shape). */
export interface PdfNativeShapedTextPiece {
  readonly text: string;
  readonly rtl: boolean;
  readonly kind: "plain" | "shaped";
  readonly sourceId?: string;
  readonly fontSha256?: string;
  readonly unitsPerEm?: number;
  /** [glyph id, cluster, x advance, y advance, x offset, y offset], font units, visual order. */
  readonly glyphs?: readonly (readonly [number, number, number, number, number, number])[];
}

export interface PdfNativeContentAppendImageContent {
  readonly kind: "image";
  readonly mimeType: "image/png" | "image/jpeg";
  readonly imageBase64: string;
  /** Maps the image unit square onto PDF user space. */
  readonly matrix: PdfNativeContentAppendMatrix;
}

export type PdfNativeContentAppendPathCommand = Readonly<
  | { op: "m" | "l"; x: number; y: number }
  | { op: "c"; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: "re"; x: number; y: number; width: number; height: number }
  | { op: "h" }
>;

export interface PdfNativeContentAppendPathContent {
  readonly kind: "path";
  readonly stroke: Readonly<{
    color: PdfNativeContentAppendColor;
    width: number;
    dash: readonly number[] | null;
    cap: 0 | 1 | 2;
    join: 0 | 1 | 2;
  }> | null;
  readonly fill: Readonly<{ color: PdfNativeContentAppendColor }> | null;
  /** Constant alpha for stroke (and fill unless overridden); null keeps the page default. */
  readonly opacity: number | null;
  /** Separate fill alpha; null falls back to `opacity`. */
  readonly fillOpacity?: number | null;
  readonly commands: readonly PdfNativeContentAppendPathCommand[];
}

export type PdfNativeContentAppendContent =
  | PdfNativeContentAppendTextContent
  | PdfNativeContentAppendImageContent
  | PdfNativeContentAppendPathContent;

export interface PdfNativeContentAppendSpan {
  readonly start: number;
  readonly end: number;
  readonly expectedSha256: string;
}

export type PdfNativeContentAppendRequest = Readonly<
  | { direction: "append"; content: PdfNativeContentAppendContent }
  | {
      direction: "append";
      appendKind: PdfNativeContentAppendKind;
      blockBase64: string;
      reuse: readonly PdfNativeContentAppendReusedResource[];
    }
  | {
      direction: "remove";
      appendKind: PdfNativeContentAppendKind;
      span: PdfNativeContentAppendSpan;
    }
>;

export interface PdfNativeContentAppendPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly request: PdfNativeContentAppendRequest;
}

export type PdfNativeContentAppendWireAddition = Readonly<
  | {
      category: "Font";
      resourceName: string;
      text: string;
      fontPackage: PdfNativeFontExtensionWirePackage;
    }
  | Exclude<PdfNativeContentAppendAddition, { category: "Font" }>
>;

export interface PdfNativeContentAppendEvidence
  extends Omit<PdfNativeContentAppendPlan, "additions"> {
  readonly additions: readonly PdfNativeContentAppendWireAddition[];
  /** SHA-256 of the exact bytes appended to, or removed from, the stream. */
  readonly blockSha256: string;
  /** Block span inside the resulting decoded stream (append) or the source (remove). */
  readonly blockStart: number;
  readonly blockEnd: number;
}

export interface PdfNativeContentAppendHistorySeed {
  readonly appendKind: PdfNativeContentAppendKind;
  readonly pageNumber: number;
  /** Exact bytes to remove from the candidate stream tail to undo. */
  readonly removeSpan: PdfNativeContentAppendSpan;
  /** Block without its whitespace prefix, re-appended on redo. */
  readonly blockBase64: string;
  readonly reuse: readonly PdfNativeContentAppendReusedResource[];
}

export interface PdfNativeContentAppendPlanResult {
  readonly payload: PdfNativeContentAppendPayload;
  readonly plan: PdfNativeContentAppendPlan;
  readonly evidence: PdfNativeContentAppendEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly history: PdfNativeContentAppendHistorySeed;
}















export function formatPdfNativeContentAppendNumber(value: number): string {
  if (!Number.isFinite(value)) invalid("Content append number is invalid.");
  const normalized = Math.abs(value) < 0.0000005 ? 0 : value;
  const formatted = normalized
    .toFixed(6)
    .replace(/\.(\d*?)0+$/, (_, kept: string) => (kept ? `.${kept}` : "")) // no lookbehind: Safari < 16.4
    .replace(/\.$/, "");
  return formatted === "-0" ? "0" : formatted;
}

const fmt = formatPdfNativeContentAppendNumber;

function multiply(
  left: PdfNativeContentAppendMatrix,
  right: PdfNativeContentAppendMatrix,
): PdfNativeContentAppendMatrix {
  return Object.freeze([
    left[0] * right[0] + left[2] * right[1],
    left[1] * right[0] + left[3] * right[1],
    left[0] * right[2] + left[2] * right[3],
    left[1] * right[2] + left[3] * right[3],
    left[0] * right[4] + left[2] * right[5] + left[4],
    left[1] * right[4] + left[3] * right[5] + left[5],
  ] as const);
}

function inverse(matrix: PdfNativeContentAppendMatrix): PdfNativeContentAppendMatrix {
  const [a, b, c, d, e, f] = matrix;
  const determinant = a * d - b * c;
  if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-9) {
    invalid("Content stream leaves a singular transform at its tail.");
  }
  return Object.freeze([
    d / determinant,
    -b / determinant,
    -c / determinant,
    a / determinant,
    (c * f - d * e) / determinant,
    (b * e - a * f) / determinant,
  ] as const);
}

function isIdentity(matrix: PdfNativeContentAppendMatrix): boolean {
  return matrix.every((value, index) => Math.abs(value - IDENTITY[index]) < 1e-9);
}

export interface PdfNativeContentAppendTailState {
  readonly restoredDepth: number;
  readonly effectiveMatrix: PdfNativeContentAppendMatrix;
  readonly endsWithWhitespace: boolean;
}

const WHITESPACE = new Set([" ", "\n", "\r", "\t", "\f", "\0"]);
const DELIMITERS = new Set(["(", ")", "<", ">", "[", "]", "{", "}", "/", "%"]);

/**
 * Walk the whole stream with a graphics-state tokenizer and report the state
 * an appended block inherits: how many `q` levels are left open and which
 * CTM is effective once they are restored. Mirrors the server's prefix walk.
 */
export function analyzePdfNativeContentStreamTail(
  streamBytes: Uint8Array,
): PdfNativeContentAppendTailState {
  const source = LATIN1_EXACT.decode(streamBytes);
  const length = source.length;
  const stack: PdfNativeContentAppendMatrix[] = [];
  let matrix: PdfNativeContentAppendMatrix = IDENTITY;
  let inText = false;
  let containerDepth = 0;
  let operands: string[] = [];
  let index = 0;
  while (index < length) {
    const character = source[index];
    if (WHITESPACE.has(character)) {
      index += 1;
      continue;
    }
    if (character === "%") {
      while (index < length && source[index] !== "\n" && source[index] !== "\r") {
        index += 1;
      }
      if (index >= length) {
        invalid("Content stream ends inside a comment.");
      }
      continue;
    }
    if (character === "(") {
      let depth = 1;
      index += 1;
      while (index < length && depth > 0) {
        const current = source[index];
        if (current === "\\") {
          index += 2;
          continue;
        }
        if (current === "(") depth += 1;
        else if (current === ")") depth -= 1;
        index += 1;
      }
      if (depth !== 0) invalid("Content stream literal string is unterminated.");
      if (containerDepth === 0) operands.push("(");
      continue;
    }
    if (character === "<") {
      if (source[index + 1] === "<") {
        containerDepth += 1;
        index += 2;
        continue;
      }
      const end = source.indexOf(">", index);
      if (end < 0) invalid("Content stream hex string is unterminated.");
      index = end + 1;
      if (containerDepth === 0) operands.push("<");
      continue;
    }
    if (character === ">") {
      if (source[index + 1] !== ">" || containerDepth === 0) {
        invalid("Content stream dictionary delimiter is unbalanced.");
      }
      containerDepth -= 1;
      index += 2;
      continue;
    }
    if (character === "[") {
      containerDepth += 1;
      index += 1;
      continue;
    }
    if (character === "]") {
      if (containerDepth === 0) invalid("Content stream array is unbalanced.");
      containerDepth -= 1;
      index += 1;
      continue;
    }
    if (character === "{" || character === "}") {
      invalid("Content stream contains PostScript procedures.");
    }
    if (character === "/") {
      let end = index + 1;
      while (end < length && !WHITESPACE.has(source[end]) && !DELIMITERS.has(source[end])) {
        end += 1;
      }
      if (containerDepth === 0) operands.push(source.slice(index, end));
      index = end;
      continue;
    }
    let end = index;
    while (end < length && !WHITESPACE.has(source[end]) && !DELIMITERS.has(source[end])) {
      end += 1;
    }
    if (end === index) invalid("Content stream contains an unreadable token.");
    const token = source.slice(index, end);
    index = end;
    if (containerDepth > 0) continue;
    if (NUMBER_PATTERN.test(token)) {
      operands.push(token);
      continue;
    }
    if (token === "BI") {
      // One opaque painting -- its dictionary, `ID`, data that may hold any
      // byte, `EI` -- that sets no transform and saves or restores nothing
      // (the server's prefix walk reads it the same way). Its data is
      // stepped over by the count its dictionary states, else pdf.js's rule.
      if (operands.length !== 0) invalid("Content stream inline image carries operands.");
      index = skipMraiNativeInlineImage(source, index);
      operands = [];
      continue;
    }
    if (token === "ID" || token === "EI") {
      invalid("Content stream inline image delimiter stands alone.");
    } else if (token === "BT") {
      if (inText) invalid("Content stream nests text objects.");
      inText = true;
    } else if (token === "ET") {
      if (!inText) invalid("Content stream text object is unbalanced.");
      inText = false;
    } else if (!inText && token === "q") {
      if (operands.length !== 0) invalid("Content stream q carries operands.");
      stack.push(matrix);
      if (stack.length > MAX_RESTORED_DEPTH) {
        invalid("Content stream graphics stack is too deep.");
      }
    } else if (!inText && token === "Q") {
      if (operands.length !== 0 || stack.length === 0) {
        invalid("Content stream restore is unmatched.");
      }
      matrix = stack.pop()!;
    } else if (!inText && token === "cm") {
      if (operands.length !== 6 || !operands.every((item) => NUMBER_PATTERN.test(item))) {
        invalid("Content stream cm operands are invalid.");
      }
      const local = readMatrix(
        operands.map((item) => Number(item)),
        "content stream cm",
      );
      matrix = multiply(matrix, local);
    }
    operands = [];
  }
  if (containerDepth !== 0 || inText || operands.length !== 0) {
    invalid("Content stream ends at an ambiguous boundary.");
  }
  const restoredDepth = stack.length;
  return Object.freeze({
    restoredDepth,
    effectiveMatrix: restoredDepth > 0 ? stack[0] : matrix,
    endsWithWhitespace:
      length === 0 || WHITESPACE.has(source[length - 1]),
  });
}

function nextResourceName(
  used: ReadonlySet<string>,
  prefix: string,
): string {
  for (let index = 0; index < 10_000; index += 1) {
    const candidate = `${prefix}${index}`;
    if (!used.has(candidate)) return candidate;
  }
  return invalid("Content append resource namespace is exhausted.");
}

function usedNames(
  page: PdfNativeContentAppendPageInspection,
  category: PdfNativeContentAppendCategory,
): Set<string> {
  return new Set(page.categories[category].entries.map((entry) => entry.resourceName));
}

function base64ToBytes(value: string): Uint8Array {
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(value) || value.length % 4 !== 0) {
    invalid("Content append base64 payload is invalid.");
  }
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary);
}

const ascii = (value: string): Uint8Array => new TextEncoder().encode(value);
const latin1 = (bytes: Uint8Array): string => LATIN1_EXACT.decode(bytes);

interface BuiltBlock {
  /** Block body between `q` and `Q`, without the CTM restore prefix. */
  readonly inner: string;
  readonly additions: readonly PdfNativeContentAppendAddition[];
  readonly appendKind: PdfNativeContentAppendKind;
}

async function buildTextBlock(
  content: PdfNativeContentAppendTextContent,
  page: PdfNativeContentAppendPageInspection,
  fontBytes: Uint8Array,
): Promise<BuiltBlock> {
  const lines = content.text.split(/\r?\n/);
  const joined = lines.join("");
  if (!joined.trim() || joined.length > MAX_TEXT_CODE_UNITS) {
    invalid("Content append text is empty or too long.");
  }
  const resourceName = nextResourceName(usedNames(page, "Font"), "MRAIFont");
  const metrics = await measurePdfNativeFontExtensionText({ fontBytes, text: joined });
  const fontPackage = await buildPdfNativeFontExtensionPackage({
    fontBytes,
    text: joined,
    resourceName,
    firstObjectNumber: page.firstNewObjectNumber,
    subset: true,
  });
  if (fontPackage.encodedHex.length !== metrics.glyphs.length * 4) {
    invalid("Content append glyph encoding does not match its measurement.");
  }
  const lineHex: string[] = [];
  let glyphCursor = 0;
  for (const line of lines) {
    const start = glyphCursor;
    let consumed = 0;
    while (consumed < line.length) {
      const glyph = metrics.glyphs[glyphCursor];
      if (!glyph) invalid("Content append glyph run ended early.");
      consumed += glyph.text.length;
      glyphCursor += 1;
    }
    if (consumed !== line.length) {
      invalid("Content append text shaping spans a line break.");
    }
    lineHex.push(fontPackage.encodedHex.slice(start * 4, glyphCursor * 4));
  }
  if (glyphCursor !== metrics.glyphs.length) {
    invalid("Content append glyph run has unassigned glyphs.");
  }
  const ascent = (metrics.ascent / metrics.unitsPerEm) * content.fontSizePoints;
  // Text space is rotated with the page so the run reads upright on screen;
  // the baseline sits one ascent below the box top along the screen-down axis.
  const rotation = content.rotation ?? 0;
  const [a, b, c, d] = TEXT_ROTATION_MATRICES[rotation];
  const down = SCREEN_DOWN_DIRECTIONS[rotation];
  const baselineX = content.x + down[0] * ascent;
  const baselineY = content.top + down[1] * ascent;
  // The text box is wrapped in a box mark of its own (box-mark.ts): a box
  // inserted under or beside another stays its own box when the page is read
  
  const operators = [
    `/${MRAI_NATIVE_BOX_MARK_TAG} BMC`,
    "BT",
    `/${resourceName} ${fmt(content.fontSizePoints)} Tf`,
    `${fmt(content.color.r)} ${fmt(content.color.g)} ${fmt(content.color.b)} rg`,
    ...syntheticBoldOperators(content),
    `${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(baselineX)} ${fmt(baselineY)} Tm`,
    `${fmt(content.lineHeightPoints)} TL`,
  ];
  lineHex.forEach((hex, index) => {
    if (index > 0) operators.push("T*");
    if (hex) operators.push(`<${hex}> Tj`);
  });
  operators.push("ET", "EMC");
  return Object.freeze({
    inner: operators.join("\n"),
    additions: Object.freeze([
      Object.freeze({
        category: "Font" as const,
        resourceName,
        text: joined,
        fontPackage,
      }),
    ]),
    appendKind: "text" as const,
  });
}

/** A packaged face's name in the PDF (its subset tag is added before it). */
const PDF_COMPLEX_SCRIPT_BASE_NAMES: Readonly<Record<string, string>> = Object.freeze({
  "mrnote.noto-sans-arabic.v1": "NotoSansArabic-Regular",
  "mrnote.noto-sans-hebrew.v1": "NotoSansHebrew-Regular",
  "mrnote.noto-sans-thai.v1": "NotoSansThai-Regular",
  "mrnote.noto-sans-devanagari.v1": "NotoSansDevanagari-Regular",
  "mrnote.noto-sans-bengali.v1": "NotoSansBengali-Regular",
});

function utf16BeHex(text: string): string {
  let hex = "FEFF";
  for (let index = 0; index < text.length; index += 1) hex += text.charCodeAt(index).toString(16).toUpperCase().padStart(4, "0");
  return hex;
}

/**
 * Lines holding words of a complex script, as the API shaped them
 * (`content.shapedLines`): each line's pieces left to right, a plain piece
 * in the bundled face (one package for all of them, as a plain block), a
 * shaped piece glyph by glyph in its packaged face (one package per face),
 * each glyph placed at its pen position plus its offsets (marks over their
 * bases, Thai's above and below vowels), the pen moved by its advance. A
 * shaped piece is a marked span whose ActualText is its words in reading
 * order, so a reader copies and searches them as typed.
 */
async function buildShapedTextBlock(
  content: PdfNativeContentAppendTextContent,
  page: PdfNativeContentAppendPageInspection,
  fontBytes: Uint8Array,
  loadComplexFont: (sourceId: string) => Promise<Uint8Array | ArrayBuffer | null>,
): Promise<BuiltBlock> {
  const shapedLines = content.shapedLines!;
  const fontNames = usedNames(page, "Font");
  const additions: PdfNativeContentAppendAddition[] = [];
  let nextObject = page.firstNewObjectNumber;
  const size = content.fontSizePoints;
  // The bundled face: every plain piece's characters, in drawing order.
  const plainText = shapedLines.flatMap((line) => line.filter((piece) => piece.kind === "plain").map((piece) => piece.text)).join("");
  const plainSlices: Array<{ hex: string; advance: number }> = [];
  let plainResource: string | null = null;
  let ascentPoints = 0;
  if (plainText.replace(/\s/gu, "")) {
    plainResource = nextResourceName(fontNames, "MRAIFont");
    fontNames.add(plainResource);
    const metrics = await measurePdfNativeFontExtensionText({ fontBytes, text: plainText });
    const fontPackage = await buildPdfNativeFontExtensionPackage({
      fontBytes, text: plainText, resourceName: plainResource, firstObjectNumber: nextObject, subset: true,
    });
    if (fontPackage.encodedHex.length !== metrics.glyphs.length * 4) {
      invalid("Content append glyph encoding does not match its measurement.");
    }
    nextObject += fontPackage.objects.length;
    additions.push(Object.freeze({ category: "Font" as const, resourceName: plainResource, text: plainText, fontPackage }));
    ascentPoints = Math.max(ascentPoints, (metrics.ascent / metrics.unitsPerEm) * size);
    let cursor = 0;
    for (const line of shapedLines) {
      for (const piece of line) {
        if (piece.kind !== "plain") continue;
        const start = cursor;
        let consumed = 0;
        let advance = 0;
        while (consumed < piece.text.length) {
          const glyph = metrics.glyphs[cursor];
          if (!glyph) invalid("Content append glyph run ended early.");
          consumed += glyph.text.length;
          advance += glyph.advanceAt1000;
          cursor += 1;
        }
        if (consumed !== piece.text.length) invalid("Content append shaping spans a piece.");
        plainSlices.push({ hex: fontPackage.encodedHex.slice(start * 4, cursor * 4), advance: (advance / 1000) * size });
      }
    }
    if (cursor !== metrics.glyphs.length) invalid("Content append glyph run has unassigned glyphs.");
  } else {
    // White space alone: drawn as nothing, its advance the face's.
    for (const line of shapedLines) {
      for (const piece of line) if (piece.kind === "plain") plainSlices.push({ hex: "", advance: piece.text.length * size * 0.26 });
    }
  }
  // Each packaged face: its shaped glyphs, in drawing order, in one subset.
  //
  // What each glyph reads as, for a reader that rebuilds the words from the
  // shows (pdf.js's text layer, copy and search; Acrobat reads the line's
  // ActualText): each cluster is drawn as one glyph that reads as exactly
  // its letters -- a cluster of several glyphs (a letter and the dots Noto
  // Sans Arabic draws apart, a conjunct and its vowel sign) as a composite
  // of them placed as the shaper placed them. A glyph left unmapped read as
  // its code (pdf.js: U+0006 in a word), and a joiner given to it was copied
  // with the words. Up to its first combining mark (Mn) the cluster's
  // letters go to that glyph: pdf.js takes a code whose letters hold one for
  // a zero-width diacritic and hangs it on the letters before it (दु after
  // a line break read with the line before); the marks and what follows
  // them go to a carrier drawn for them (an empty glyph of no advance, after
  // the cluster in reading order). A glyph that reads otherwise elsewhere is
  // drawn from a copy of its own: a code reads as one thing.
  type Show =
    | { readonly kind: "glyph"; readonly glyph: number; readonly text: string; readonly dx: number; readonly dy: number }
    | { readonly kind: "empty"; readonly glyph: number; readonly text: string; readonly dx: number; readonly dy: 0 }
    | { readonly kind: "composite"; readonly components: readonly { glyph: number; x: number; y: number }[];
      readonly advance: number; readonly text: string; readonly dx: number; readonly dy: 0 };
  const showsOf = new Map<PdfNativeShapedTextPiece, Show[]>();
  const faces = new Map<string, { resource: string; unitsPerEm: number; codes: string[]; widths: Map<string, number> }>();
  // The white space between two words of one face is drawn in that face (its
  // own space): a reader that ends its item where the font changes (pdf.js)
  // keeps a right-to-left phrase in one item, read in order, not word by word
  // as drawn ("بالعالم مرحبا").
  const spaceFace = new Map<PdfNativeShapedTextPiece, string>();
  for (const line of shapedLines) {
    line.forEach((piece, index) => {
      const before = line[index - 1];
      const after = line[index + 1];
      if (piece.kind === "plain" && /^ +$/u.test(piece.text) && before?.kind === "shaped" && after?.kind === "shaped" &&
        before.sourceId === after.sourceId) spaceFace.set(piece, before.sourceId!);
    });
  }
  const bySource = new Map<string, PdfNativeShapedTextPiece[]>();
  for (const line of shapedLines) for (const piece of line) {
    const sourceId = piece.kind === "shaped" ? piece.sourceId! : spaceFace.get(piece);
    if (sourceId) bySource.set(sourceId, [...(bySource.get(sourceId) ?? []), piece]);
  }
  for (const [sourceId, listed] of bySource) {
    const face = PDF_COMPLEX_SCRIPT_FONTS[sourceId];
    const loaded = face ? await loadComplexFont(sourceId) : null;
    if (!face || !loaded) invalid("Content append complex-script face is unavailable.");
    const program = await readSfntProgram(loaded instanceof Uint8Array ? loaded : new Uint8Array(loaded));
    const space = program.cmap.get(0x20);
    const pieces = listed.filter((piece) => piece.kind === "shaped" || (space !== undefined && space > 0));
    for (const piece of listed) if (piece.kind === "plain" && !pieces.includes(piece)) spaceFace.delete(piece);
    for (const piece of pieces) {
      if (piece.kind === "plain") {
        showsOf.set(piece, [...piece.text].map((): Show => ({ kind: "glyph", glyph: space!, text: " ", dx: 0, dy: 0 })));
        continue;
      }
      const glyphs = piece.glyphs!;
      const starts = [...new Set(glyphs.map((glyph) => glyph[1]))].sort((left, right) => left - right);
      const shows: Show[] = [];
      let pen = 0;
      for (let index = 0; index < glyphs.length;) {
        // The cluster: its glyphs side by side in drawing order.
        const cluster = glyphs[index][1];
        let last = index;
        while (last + 1 < glyphs.length && glyphs[last + 1][1] === cluster) last += 1;
        const group = glyphs.slice(index, last + 1);
        const end = starts.find((start) => start > cluster) ?? piece.text.length;
        const letters = piece.text.slice(cluster, end);
        const advance = group.reduce((total, glyph) => total + glyph[2], 0);
        const mark = letters.search(/\p{Mn}/u);
        const own = advance !== 0 && mark > 0 ? letters.slice(0, mark) : letters;
        const carried = own.length < letters.length ? letters.slice(own.length) : "";
        const at = pen;
        if (carried && piece.rtl) shows.push({ kind: "empty", glyph: group[0][0], text: carried, dx: at, dy: 0 });
        if (group.length === 1) {
          shows.push({ kind: "glyph", glyph: group[0][0], text: own, dx: at + group[0][4], dy: group[0][5] });
        } else {
          let x = 0;
          const components = group.map((glyph) => {
            const placed = { glyph: glyph[0], x: Math.round(x + glyph[4]), y: Math.round(glyph[5]) };
            x += glyph[2];
            return placed;
          });
          shows.push({ kind: "composite", components, advance: Math.max(0, Math.round(advance)), text: own, dx: at, dy: 0 });
        }
        if (carried && !piece.rtl) shows.push({ kind: "empty", glyph: group[0][0], text: carried, dx: at + advance, dy: 0 });
        pen += advance;
        index = last + 1;
      }
      showsOf.set(piece, shows);
    }
    // Each glyph's first reading is the glyph itself; another reading, every
    // carrier and every composite, a glyph of its own after the subset's.
    const all = pieces.flatMap((piece) => showsOf.get(piece)!);
    const first = new Map<number, string>();
    const copyKeys: string[] = [];
    const copies: PdfNativeSfntGlyphCopy[] = [];
    const keyOf = (show: Show) => show.kind === "composite"
      ? `composite\u0000${JSON.stringify([show.components, show.advance])}\u0000${show.text}`
      : `${show.kind === "empty" ? "carrier" : show.glyph}\u0000${show.text}`;
    const primary = (show: Show) => show.kind === "glyph" && first.get(show.glyph) === show.text;
    for (const show of all) {
      if (show.kind === "glyph" && !first.has(show.glyph)) { first.set(show.glyph, show.text); continue; }
      if (primary(show)) continue;
      const key = keyOf(show);
      if (copyKeys.includes(key)) continue;
      copyKeys.push(key);
      copies.push(show.kind === "composite" ? { components: show.components, advance: show.advance }
        : { glyph: show.glyph, empty: show.kind === "empty" });
    }
    const subset = buildPdfNativeSfntSubset({
      program, text: "",
      glyphIds: all.flatMap((show) => show.kind === "composite" ? show.components.map((component) => component.glyph) : [show.glyph]),
      glyphCopies: copies,
    });
    const cidOf = (show: Show) => primary(show) && show.kind === "glyph"
      ? subset.shapedGlyphs!.get(show.glyph)!.cid
      : subset.glyphCopies![copyKeys.indexOf(keyOf(show))].cid;
    const resource = nextResourceName(fontNames, "MRAIFont");
    fontNames.add(resource);
    const fontPackage = await buildPdfNativeShapedFontPackage({
      subset, programSha256: face.sha256, baseName: PDF_COMPLEX_SCRIPT_BASE_NAMES[sourceId] ?? "NotoSans-Regular",
      glyphs: all.map((show) => ({ cid: cidOf(show), text: show.text })), resourceName: resource, firstObjectNumber: nextObject,
    });
    nextObject += fontPackage.objects.length;
    additions.push(Object.freeze({ category: "Font" as const, resourceName: resource,
      text: pieces.filter((piece) => piece.kind === "shaped").map((piece) => piece.text).join(""), fontPackage }));
    const codes: string[] = [];
    for (let index = 0; index < fontPackage.encodedHex.length; index += 4) codes.push(fontPackage.encodedHex.slice(index, index + 4));
    // Each code's advance as the font's /W states it (at 1000 an em): where
    // a viewer's pen stands after the code is drawn.
    const widths = new Map<string, number>();
    for (const entry of [...subset.shapedGlyphs!.values(), ...(subset.glyphCopies ?? [])]) {
      widths.set(entry.cid.toString(16).toUpperCase().padStart(4, "0"), Math.round((entry.advance * 1000) / subset.unitsPerEm));
    }
    faces.set(sourceId, { resource, unitsPerEm: subset.unitsPerEm, codes, widths });
    ascentPoints = Math.max(ascentPoints, (subset.ascent / subset.unitsPerEm) * size);
  }
  const rotation = content.rotation ?? 0;
  const [a, b, c, d] = TEXT_ROTATION_MATRICES[rotation];
  const down = SCREEN_DOWN_DIRECTIONS[rotation];
  // A point of the line (u along it, v up from its baseline) on the page.
  const at = (lineIndex: number, u: number, v: number) => {
    const baseline = ascentPoints + lineIndex * content.lineHeightPoints;
    return [content.x + down[0] * baseline + a * u + c * v, content.top + down[1] * baseline + b * u + d * v] as const;
  };
  const operators = [
    `/${MRAI_NATIVE_BOX_MARK_TAG} BMC`,
    "BT",
    `${fmt(content.color.r)} ${fmt(content.color.g)} ${fmt(content.color.b)} rg`,
    ...syntheticBoldOperators(content),
  ];
  let current: string | null = null;
  const place = (x: number, y: number) => operators.push(`${fmt(a)} ${fmt(b)} ${fmt(c)} ${fmt(d)} ${fmt(x)} ${fmt(y)} Tm`);
  let plainIndex = 0;
  const faceCursor = new Map<string, number>();
  const lineTexts = content.text.split(/\r?\n/u);
  shapedLines.forEach((line, lineIndex) => {
    // One ActualText a line, its words in reading order: a reader that
    // joins the line's glyphs (pdf.js ends an item at marked content) reads
    // its right-to-left words in order, not word by word as drawn.
    const spanned = line.some((piece) => piece.kind === "shaped");
    if (spanned) operators.push(`/Span <</ActualText <${utf16BeHex(lineTexts[lineIndex] ?? "")}>>> BDC`);
    // The line placed once, its codes then moved along it by the numbers
    // of its shows (and raised by Ts): a viewer that walks the shows keeps
    // its own pen, which runs short of the drawing after a code it counts
    // as a zero-width diacritic (स्), and a glyph placed by a Tm of its own
    // after one read as a new word ("नमस् ते").
    place(...at(lineIndex, 0, 0));
    let pen = 0;
    let drawnTo = 0;
    let rise = 0;
    let shown: string[] = [];
    const flush = () => {
      if (shown.length) operators.push(`[${shown.join(" ")}] TJ`);
      shown = [];
    };
    const show = (resource: string, code: string, x: number, y: number, widthAt1000: number) => {
      if (current !== resource) {
        flush();
        operators.push(`/${resource} ${fmt(size)} Tf`);
        current = resource;
      }
      if (Math.abs(y - rise) > 1e-6) {
        flush();
        operators.push(`${fmt(y)} Ts`);
        rise = y;
      }
      const gap = x - drawnTo;
      if (Math.abs(gap) > 1e-6) shown.push(fmt((-gap * 1000) / size));
      shown.push(`<${code}>`);
      drawnTo = x + (widthAt1000 * size) / 1000;
    };
    for (const piece of line) {
      const spacedIn = spaceFace.get(piece);
      if (piece.kind === "plain" && spacedIn) {
        // A space of the face around it (its advance the line's, as laid).
        const slice = plainSlices[plainIndex++];
        const face = faces.get(spacedIn)!;
        let cursor = faceCursor.get(spacedIn) ?? 0;
        const step = slice.advance / Math.max(1, piece.text.length);
        for (let index = 0; index < piece.text.length; index += 1) {
          const code = face.codes[cursor];
          show(face.resource, code, pen + index * step, 0, face.widths.get(code) ?? 0);
          cursor += 1;
        }
        faceCursor.set(spacedIn, cursor);
        pen += slice.advance;
        continue;
      }
      if (piece.kind === "plain") {
        const slice = plainSlices[plainIndex++];
        if (slice.hex && plainResource) show(plainResource, slice.hex, pen, 0, (slice.advance * 1000) / size);
        pen += slice.advance;
        continue;
      }
      const face = faces.get(piece.sourceId!)!;
      const unit = size / face.unitsPerEm;
      let cursor = faceCursor.get(piece.sourceId!) ?? 0;
      for (const entry of showsOf.get(piece)!) {
        const code = face.codes[cursor];
        show(face.resource, code, pen + entry.dx * unit, entry.dy * unit, face.widths.get(code) ?? 0);
        cursor += 1;
      }
      faceCursor.set(piece.sourceId!, cursor);
      pen += piece.glyphs!.reduce((total, glyph) => total + glyph[2], 0) * unit;
    }
    flush();
    if (rise) {
      operators.push("0 Ts");
      rise = 0;
    }
    if (spanned) operators.push("EMC");
  });
  operators.push("ET", "EMC");
  if (!additions.length) invalid("Content append text is empty.");
  return Object.freeze({ inner: operators.join("\n"), additions: Object.freeze(additions), appendKind: "text" as const });
}

async function buildImageBlock(
  content: PdfNativeContentAppendImageContent,
  page: PdfNativeContentAppendPageInspection,
): Promise<BuiltBlock> {
  const image = await inspectMraiNativeEmbeddedImageReplacement(
    base64ToBytes(content.imageBase64),
    content.mimeType,
  );
  const resourceName = nextResourceName(usedNames(page, "XObject"), "MRAIImg");
  const reference = Object.freeze({
    objectNumber: page.firstNewObjectNumber,
    generation: 0 as const,
  });
  const softMask = image.softMask
    ? Object.freeze({ objectNumber: page.firstNewObjectNumber + 1, generation: 0 as const })
    : null;
  const inner = [
    `${content.matrix.map((value) => fmt(value)).join(" ")} cm`,
    `/${resourceName} Do`,
  ].join("\n");
  return Object.freeze({
    inner,
    additions: Object.freeze([
      Object.freeze({
        category: "XObject" as const,
        resourceName,
        reference,
        softMask,
        image,
      }),
    ]),
    appendKind: "image" as const,
  });
}

function buildPathBlock(
  content: PdfNativeContentAppendPathContent,
  page: PdfNativeContentAppendPageInspection,
): BuiltBlock {
  const operators: string[] = [];
  const additions: PdfNativeContentAppendAddition[] = [];
  const fillAlpha = content.fill ? (content.fillOpacity ?? content.opacity) : null;
  const strokeAlpha = content.stroke ? content.opacity : null;
  if (fillAlpha !== null || strokeAlpha !== null) {
    const resourceName = nextResourceName(usedNames(page, "ExtGState"), "MRAIGs");
    additions.push(
      Object.freeze({
        category: "ExtGState" as const,
        resourceName,
        reference: Object.freeze({
          objectNumber: page.firstNewObjectNumber,
          generation: 0 as const,
        }),
        fillOpacity: fillAlpha,
        strokeOpacity: strokeAlpha,
      }),
    );
    operators.push(`/${resourceName} gs`);
  }
  if (content.stroke) {
    const stroke = content.stroke;
    operators.push(`${fmt(stroke.width)} w`);
    if (stroke.cap !== 0) operators.push(`${stroke.cap} J`);
    if (stroke.join !== 0) operators.push(`${stroke.join} j`);
    if (stroke.dash) {
      operators.push(`[${stroke.dash.map((value) => fmt(value)).join(" ")}] 0 d`);
    }
    operators.push(pathColorOperator(stroke.color, true));
  }
  if (content.fill) {
    operators.push(pathColorOperator(content.fill.color, false));
  }
  for (const command of content.commands) {
    switch (command.op) {
      case "m":
      case "l":
        operators.push(`${fmt(command.x)} ${fmt(command.y)} ${command.op}`);
        break;
      case "c":
        operators.push(
          `${fmt(command.x1)} ${fmt(command.y1)} ${fmt(command.x2)} ${fmt(command.y2)} ${fmt(command.x)} ${fmt(command.y)} c`,
        );
        break;
      case "re":
        operators.push(
          `${fmt(command.x)} ${fmt(command.y)} ${fmt(command.width)} ${fmt(command.height)} re`,
        );
        break;
      case "h":
        operators.push("h");
        break;
    }
  }
  operators.push(
    content.fill && content.stroke ? "B" : content.fill ? "f" : "S",
  );
  return Object.freeze({
    inner: operators.join("\n"),
    additions: Object.freeze(additions),
    appendKind: "path" as const,
  });
}

/** Assemble the exact block bytes: restore prefix, `q`, tail CTM undo, body, `Q`. */
function assembleBlock(inner: string, tail: PdfNativeContentAppendTailState): string {
  const parts: string[] = [];
  for (let index = 0; index < tail.restoredDepth; index += 1) parts.push("Q");
  parts.push("q");
  if (!isIdentity(tail.effectiveMatrix)) {
    parts.push(
      `${inverse(tail.effectiveMatrix)
        .map((value) => fmt(value))
        .join(" ")} cm`,
    );
  }
  parts.push(inner, "Q");
  return `${parts.join("\n")}\n`;
}

interface ParsedBlock {
  readonly restoredDepth: number;
  readonly tailCm: string | null;
  readonly inner: string;
}

/**
 * Split one appended block into its restore prefix, optional tail-CTM undo
 * line and body. The tail line is only present when the stream leaves a
 * non-identity CTM, so the caller decides whether the first body line is it.
 */
function parseBlock(block: string, expectTailCm: boolean): ParsedBlock | null {
  if (!block.endsWith("\nQ\n")) return null;
  const lines = block.slice(0, -"\nQ\n".length).split("\n");
  let index = 0;
  while (index < lines.length && lines[index] === "Q") index += 1;
  const restoredDepth = index;
  if (lines[index] !== "q") return null;
  index += 1;
  let tailCm: string | null = null;
  if (expectTailCm) {
    const line = lines[index];
    if (!line || !/^[-+.\d ]+ cm$/.test(line)) return null;
    tailCm = line.slice(0, -" cm".length);
    index += 1;
  }
  const inner = lines.slice(index).join("\n");
  if (!inner) return null;
  return Object.freeze({ restoredDepth, tailCm, inner });
}

function referencedNames(inner: string): ReadonlySet<string> {
  const names = new Set<string>();
  for (const match of inner.matchAll(/\/([A-Za-z][A-Za-z0-9_.+\-]{0,63})\s+(?:[-+.\d]+\s+Tf|Do|gs)\b/g)) {
    names.add(match[1]);
  }
  return names;
}

function wireAdditions(
  additions: readonly PdfNativeContentAppendAddition[],
): readonly PdfNativeContentAppendWireAddition[] {
  return Object.freeze(
    additions.map((addition) =>
      addition.category === "Font"
        ? Object.freeze({
            category: "Font" as const,
            resourceName: addition.resourceName,
            text: addition.text,
            fontPackage: serializePdfNativeFontExtensionPackage(addition.fontPackage),
          })
        : addition,
    ),
  );
}

export interface PdfNativeContentAppendPlanInput {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeContentAppendPayload;
  /** Verified unicode fallback font bytes; required for text appends. */
  readonly loadFont?: () =>
    | Uint8Array
    | ArrayBuffer
    | Promise<Uint8Array | ArrayBuffer>;
  /** A verified packaged complex-script face (`PDF_COMPLEX_SCRIPT_FONTS`),
   * for a text append whose lines hold shaped pieces. */
  readonly loadComplexFont?: (sourceId: string) => Promise<Uint8Array | ArrayBuffer | null>;
}

export async function planPdfNativeContentAppend(
  input: PdfNativeContentAppendPlanInput,
): Promise<PdfNativeContentAppendPlanResult> {
  const payload = readPdfNativeContentAppendPayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (sourceSha256 !== payload.sourceSha256) {
    invalid("Content append source changed.");
  }
  const page = await inspectMraiNativeContentAppendPage(
    input.sourceBytes,
    payload.pageNumber,
  );
  const stream = page.contentStream;
  const streamLength = stream.streamBytes.length;
  // The block lands at the end of the page drawing, and a page drawn from
  // several streams leaves the graphics state the block restores across all
  // of them; the span it writes is still an offset into the last stream.
  const tail = analyzePdfNativeContentStreamTail(
    page.drawingPrefix.length === 0
      ? stream.streamBytes
      : Uint8Array.from([
          ...page.drawingPrefix,
          0x0a,
          ...stream.streamBytes,
        ]),
  );
  const request = payload.request;

  let plan: PdfNativeContentAppendPlan;
  let operatorSpan: PdfNativeContentAppendSpan;
  let replacement: Uint8Array;
  let blockBytes: Uint8Array;
  let blockStart: number;
  let blockEnd: number;
  let history: Omit<PdfNativeContentAppendHistorySeed, "removeSpan">;
  const baseAuthority = Object.freeze({
    schemaVersion: MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
    pageNumber: page.pageNumber,
    page: page.page,
    resourceOwner: page.resourceOwner,
  });

  if (request.direction === "remove") {
    const span = request.span;
    if (span.end !== streamLength || span.start >= span.end) {
      invalid("Content removal must target the exact stream tail block.");
    }
    blockBytes = stream.streamBytes.slice(span.start, span.end);
    if ((await sha256PdfBytes(blockBytes)) !== span.expectedSha256) {
      invalid("Content removal block changed.");
    }
    const parsedBlock = parseBlock(
      latin1(blockBytes).replace(/^\n/, ""),
      !isIdentity(
        analyzePdfNativeContentStreamTail(stream.streamBytes.slice(0, span.start))
          .effectiveMatrix,
      ),
    );
    if (!parsedBlock) invalid("Content removal span is not one appended block.");
    replacement = new Uint8Array();
    operatorSpan = span;
    blockStart = span.start;
    blockEnd = span.end;
    plan = Object.freeze({
      ...baseAuthority,
      direction: "remove" as const,
      appendKind: request.appendKind,
      categories: Object.freeze({}),
      additions: Object.freeze([]),
      reusedResources: Object.freeze([]),
      restoredDepth: parsedBlock.restoredDepth,
    });
    history = Object.freeze({
      appendKind: request.appendKind,
      pageNumber: page.pageNumber,
      blockBase64: bytesToBase64(ascii(latin1(blockBytes).replace(/^\n/, ""))),
      reuse: Object.freeze([]),
    });
  } else {
    let block: string;
    let additions: readonly PdfNativeContentAppendAddition[];
    let reuse: readonly PdfNativeContentAppendReusedResource[];
    let appendKind: PdfNativeContentAppendKind;
    if ("blockBase64" in request) {
      block = latin1(base64ToBytes(request.blockBase64));
      const tailIsIdentity = isIdentity(tail.effectiveMatrix);
      const parsedBlock = parseBlock(block, !tailIsIdentity);
      if (!parsedBlock) invalid("Content re-append block is not one appended block.");
      if (parsedBlock.restoredDepth !== tail.restoredDepth) {
        invalid("Content re-append restore depth no longer matches the stream tail.");
      }
      const expectedCm = tailIsIdentity
        ? null
        : inverse(tail.effectiveMatrix).map((value) => fmt(value)).join(" ");
      if (parsedBlock.tailCm !== expectedCm) {
        invalid("Content re-append tail transform no longer matches.");
      }
      const names = referencedNames(parsedBlock.inner);
      const reuseNames = new Set(request.reuse.map((item) => item.resourceName));
      for (const name of names) {
        if (!reuseNames.has(name)) {
          invalid("Content re-append references a resource outside its reuse list.");
        }
      }
      additions = Object.freeze([]);
      reuse = request.reuse;
      appendKind = request.appendKind;
    } else {
      const content = request.content;
      let built: BuiltBlock;
      if (content.kind === "text") {
        if (!input.loadFont) invalid("Content append text requires the fallback font.");
        const loaded = await input.loadFont();
        const fontBytes =
          loaded instanceof Uint8Array ? loaded : new Uint8Array(loaded);
        built = content.shapedLines?.some((line) => line.some((piece) => piece.kind === "shaped"))
          ? await buildShapedTextBlock(content, page, fontBytes, input.loadComplexFont ??
            (() => invalid("Content append text requires its complex-script face.")))
          : await buildTextBlock(content, page, fontBytes);
      } else if (content.kind === "image") {
        built = await buildImageBlock(content, page);
      } else {
        built = buildPathBlock(content, page);
      }
      block = assembleBlock(built.inner, tail);
      additions = built.additions;
      reuse = Object.freeze([]);
      appendKind = built.appendKind;
    }
    if (block.length > MAX_BLOCK_BYTES) invalid("Content append block is too large.");
    const prefix = tail.endsWithWhitespace ? "" : "\n";
    replacement = ascii(`${prefix}${block}`);
    blockBytes = replacement;
    operatorSpan = Object.freeze({
      start: streamLength,
      end: streamLength,
      expectedSha256: await sha256PdfBytes(new Uint8Array()),
    });
    blockStart = streamLength;
    blockEnd = streamLength + replacement.length;
    const categories: Partial<
      Record<PdfNativeContentAppendCategory, PdfNativeContentAppendPageInspection["categories"]["Font"]["state"]>
    > = {};
    for (const addition of additions) {
      categories[addition.category] = page.categories[addition.category].state;
    }
    plan = Object.freeze({
      ...baseAuthority,
      direction: "append" as const,
      appendKind,
      categories: Object.freeze(categories),
      additions,
      reusedResources: reuse,
      restoredDepth: tail.restoredDepth,
    });
    history = Object.freeze({
      appendKind,
      pageNumber: page.pageNumber,
      blockBase64: bytesToBase64(ascii(block)),
      reuse,
    });
  }

  const prepared = await prepareMraiNativeContentStreamTransaction(
    input.sourceBytes,
    {
      identity: stream.identity,
      expectedObjectSha256: stream.objectSha256,
      expectedStreamSha256: stream.streamSha256,
      operatorSpan,
      replacement,
      untouchedObjects: [],
      contentAppend: plan,
    },
  );
  await verifyMraiNativeContentStreamCandidate(input.sourceBytes, prepared);
  const receipt = prepared.contentAppend;
  if (!receipt) invalid("Content append receipt is missing.");
  const blockSha256 = await sha256PdfBytes(blockBytes);
  const reuseAfter: PdfNativeContentAppendReusedResource[] = [...plan.reusedResources];
  for (const addition of plan.additions) {
    const reference =
      addition.category === "Font" ? addition.fontPackage.font : addition.reference;
    const object = receipt.objects.find(
      (item) =>
        item.objectNumber === reference.objectNumber &&
        item.generation === reference.generation,
    );
    if (!object) invalid("Content append addition receipt is missing.");
    reuseAfter.push(
      Object.freeze({
        category: addition.category,
        resourceName: addition.resourceName,
        reference: Object.freeze({ ...reference }),
        expectedObjectSha256: object.sha256,
      }),
    );
  }
  return Object.freeze({
    payload,
    plan,
    evidence: Object.freeze({
      ...plan,
      additions: wireAdditions(plan.additions),
      blockSha256,
      blockStart,
      blockEnd,
    }),
    prepared,
    history: Object.freeze({
      ...history,
      reuse: Object.freeze(reuseAfter),
      removeSpan: Object.freeze({
        start: blockStart,
        end: blockEnd,
        expectedSha256: blockSha256,
      }),
    }),
  });
}

export interface PdfNativeContentAppendPageGeometry {
  readonly cropBox: Readonly<{ xMin: number; yMin: number; xMax: number; yMax: number }>;
  readonly rotation: PdfNativeContentAppendRotation;
}

function pathColorOperator(color: PdfNativeContentAppendColor, stroke: boolean): string {
  if (color.cmyk) {
    return `${color.cmyk.map((component) => fmt(component)).join(" ")} ${stroke ? "K" : "k"}`;
  }
  return `${fmt(color.r)} ${fmt(color.g)} ${fmt(color.b)} ${stroke ? "RG" : "rg"}`;
}

/** Bold in the one weight the bundled face has: its outlines stroked in the
 * text's own colour over their fill (text render mode 2), the stroke as
 * wide as a bold set in place is. Inside the block's own q ... Q. */
function syntheticBoldOperators(content: PdfNativeContentAppendTextContent): string[] {
  if (content.bold !== true) return [];
  return [
    `${fmt(content.color.r)} ${fmt(content.color.g)} ${fmt(content.color.b)} RG`,
    `${fmt(PDF_NATIVE_SYNTHETIC_BOLD_STROKE_EM * content.fontSizePoints)} w`,
    "2 Tr",
  ];
}
