/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
} from "./vocabulary";
export {
  MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
};
import {
  inspectMraiNativeContentStream,
  inspectMraiNativePageContentGraph,
  inspectMraiNativeVectorFormHosts,
  planMraiNativePageExtGStateMutation,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamInspection,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativePreparedContentStreamTransaction,
  type PdfNativePageExtGStateEntryInspection,
} from "./content-stream-transaction";
import {
  PDF_STANDARD_BLEND_MODES,
  type PdfStandardBlendMode,
} from "../../blend-modes";
import {
  inspectMraiNativeVectorDocumentAuthority,
  inspectMraiNativeVectorContent,
  inspectMraiNativeVectorSubpaths,
  MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE,
  pdfNativeVectorSelectionClipAllows,
  pdfNativeVectorSelectionMatrix,
  type PdfNativeVectorContentShape,
  type PdfNativeVectorDocumentAuthorityInspection,
  type PdfNativeVectorGeometry,
  type PdfNativeVectorPaintOperator,
  type PdfNativeVectorSubpathProof,
} from "./vector-content-transaction";

export const MRAI_NATIVE_VECTOR_STYLE_SCHEMA_VERSION =
  "mrai.native-vector-style.v1" as const;

export interface PdfNativeVectorRgbColor {
  readonly r: number;
  readonly g: number;
  readonly b: number;
}

export interface PdfNativeVectorStyle {
  readonly fillColor: PdfNativeVectorRgbColor | null;
  readonly strokeColor: PdfNativeVectorRgbColor | null;
  readonly lineWidth: number | null;
}

export interface PdfNativeVectorStyleProof {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_STYLE_SCHEMA_VERSION;
  readonly id: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageVectorIndex: number;
  readonly selectionId: string;
  readonly selectionProofDigest: string;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly envelopeSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly paintMode: PdfNativeVectorContentShape["paintMode"];
  readonly style: PdfNativeVectorStyle;
  readonly proofDigest: string;
}

export interface PdfNativeVectorStyleInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly proofs: readonly PdfNativeVectorStyleProof[];
  readonly unsupportedSelectionIds: readonly string[];
}

export type PdfNativeVectorCmyk = readonly [number, number, number, number];

export interface PdfNativeVectorDeviceCmyk {
  readonly fill: PdfNativeVectorCmyk | null;
  readonly stroke: PdfNativeVectorCmyk | null;
}

export type PdfNativeVectorObjectStyleProofFamily = "shape" | "subpath";
export type PdfNativeVectorObjectStyleLineCap = "butt" | "round" | "square";
export type PdfNativeVectorObjectStyleLineJoin = "miter" | "round" | "bevel";

export interface PdfNativeVectorObjectStyle {
  readonly fillColor: PdfNativeVectorRgbColor | null;
  readonly strokeColor: PdfNativeVectorRgbColor | null;
  readonly lineWidth: number | null;
  readonly dashArray: readonly number[] | null;
  readonly dashPhase: number | null;
  readonly lineCap: PdfNativeVectorObjectStyleLineCap | null;
  readonly lineJoin: PdfNativeVectorObjectStyleLineJoin | null;
  readonly miterLimit: number | null;
  readonly opacity: number | null;
  readonly blendMode: PdfStandardBlendMode | null;
  readonly hasExternalGraphicsState: boolean;
  /** True when the fill is a pattern or a named colour space (not rewritable). */
  readonly fillUnreadable?: boolean;
}

export interface PdfNativeVectorObjectStyleProof {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION;
  readonly id: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly proofFamily: PdfNativeVectorObjectStyleProofFamily;
  readonly pageProofIndex: number;
  readonly selectionId: string;
  readonly selectionProofDigest: string;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly pathSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly paintSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly paintMode: PdfNativeVectorContentShape["paintMode"];
  readonly paintOperator: string;
  readonly style: PdfNativeVectorObjectStyle;
  /** The paint's own DeviceCMYK components, where it was set by `k`/`K`
   * (`style` holds them as the reader's naive RGB, which loses the ink): a
   * copy drawn again elsewhere says them again (user audit PDF 0167). Not
   * part of the digest, which covers `style`. */
  readonly deviceCmyk?: PdfNativeVectorDeviceCmyk;
  readonly proofDigest: string;
}

export interface PdfNativeVectorObjectStylePageInspection {
  readonly pageNumber: number;
  readonly proofs: readonly PdfNativeVectorObjectStyleProof[];
  readonly unsupportedSelectionIds: readonly string[];
}

export interface PdfNativeVectorObjectStyleDocumentInspection {
  readonly sourceSha256: string;
  readonly pageCount: number;
  readonly pages: readonly PdfNativeVectorObjectStylePageInspection[];
}

export interface PdfNativeVectorObjectStyleChanges {
  readonly fillColor?: PdfNativeVectorRgbColor;
  readonly strokeColor?: PdfNativeVectorRgbColor;
  readonly lineWidth?: number;
  readonly dashArray?: readonly number[];
  readonly dashPhase?: number;
  readonly lineCap?: PdfNativeVectorObjectStyleLineCap;
  readonly lineJoin?: PdfNativeVectorObjectStyleLineJoin;
  readonly miterLimit?: number;
  readonly opacity?: number;
  readonly blendMode?: PdfStandardBlendMode;
  /**
   * Repaint the same outline with a different paint operator family (fill
   * on/off). The path bytes stay verbatim; only the adjacent paint operator
   * is rewritten, keeping the source's closure and even-odd choices.
   */
  readonly paintMode?: PdfNativeVectorContentShape["paintMode"];
  /**
   * A shift of a straight line that comes with its new `lineWidth`: a drag
   * of the frame's edge across the line, which WPS reads as a thicker line
   * whose other edge stays put. In page space (PDF default user space) as
   * requested; the plan writes it as one `1 0 0 1 x y cm` in the line's own
   * user space (see `writtenChanges`), and the path bytes stay verbatim.
   */
  readonly translate?: Readonly<{ x: number; y: number }>;
  /**
   * A resize of the whole object, as WPS resizes a shape: its appearance --
   * stroke, dash, corners -- scales with its box, while the width the stream
   * sets (the ribbon's 线宽) stays what it was. In page space (PDF default user
   * space) as requested; written as one `a b c d e f cm` in the object's own
   * user space first in the style's `q…Q` (see `writtenChanges`), the path
   * bytes verbatim. It comes alone.
   */
  readonly transform?: PdfNativeVectorStyleAffine;
}

/** An affine in PDF's convention: x' = a x + c y + e, y' = b x + d y + f. */
export interface PdfNativeVectorStyleAffine {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export interface PdfNativeVectorObjectStyleEditPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION;
  readonly proof: PdfNativeVectorObjectStyleProof;
  readonly selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
  readonly beforeStyle: PdfNativeVectorObjectStyle;
  readonly expectedStyle: PdfNativeVectorObjectStyle;
  /** Paint mode after the edit; equals the proof's mode unless repainted. */
  readonly paintModeAfter: PdfNativeVectorContentShape["paintMode"];
  /** Paint operator written after the path; null when the paint is kept. */
  readonly paintOperatorAfter: string | null;
  readonly beforeSelectionFingerprint: string;
  readonly beforeSiblingFingerprints: readonly string[];
  readonly request: PdfNativeContentStreamTransactionRequest;
  /** The changes as written: a translation in the line's own user space,
   * rounded as the operator spells it. What the record carries. */
  readonly writtenChanges: PdfNativeVectorObjectStyleChanges;
  /** The page-space shift the written translation makes; null without one. */
  readonly pageShift: Readonly<{ x: number; y: number }> | null;
  /** The page-space transform the written `cm` makes; null without one. */
  readonly pageTransform: PdfNativeVectorStyleAffine | null;
}

export interface PdfNativeVectorObjectStyleVerification {
  readonly verified: true;
  readonly candidateSha256: string;
  readonly reopenedProof: PdfNativeVectorObjectStyleProof;
  readonly reopenedSelection:
    PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
}

export interface PdfNativeVectorStyleChanges {
  readonly fillColor?: PdfNativeVectorRgbColor;
  readonly strokeColor?: PdfNativeVectorRgbColor;
  readonly lineWidth?: number;
}

export interface PdfNativeVectorStyleEditPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_STYLE_SCHEMA_VERSION;
  readonly proof: PdfNativeVectorStyleProof;
  readonly beforeGeometry: PdfNativeVectorGeometry;
  readonly expectedStyle: PdfNativeVectorStyle;
  readonly request: PdfNativeContentStreamTransactionRequest;
}

export interface PdfNativeVectorStyleVerification {
  readonly verified: true;
  readonly candidateSha256: string;
  readonly reopenedProof: PdfNativeVectorStyleProof;
}

interface ContentToken {
  readonly kind: "number" | "name" | "value" | "operator";
  readonly value: string;
  readonly start: number;
  readonly end: number;
}

interface ContentOperation {
  readonly operator: string;
  readonly operands: readonly ContentToken[];
  readonly start: number;
  readonly end: number;
}

interface StyleEnvelope {
  readonly start: number;
  readonly end: number;
  readonly fillOperation: ContentOperation | null;
  readonly strokeOperation: ContentOperation | null;
  readonly lineWidthOperation: ContentOperation | null;
  readonly style: PdfNativeVectorStyle;
}

interface InternalStyleProof {
  readonly visible: PdfNativeVectorStyleProof;
  readonly selection: PdfNativeVectorContentShape;
  readonly streamInspection: PdfNativeContentStreamInspection;
  readonly operations: readonly ContentOperation[];
  readonly envelope: StyleEnvelope;
}

const NUMBER_PATTERN = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
/** A page's whole drawing, read as one scan: a cover illustration runs to
 * hundreds of thousands of operands (an annual report's front page is two
 * megabytes of paths), and the scan is linear. */
const MAX_CONTENT_TOKENS = 4_000_000;
const MAX_ABS_NUMBER = 10_000_000;
const MAX_LINE_WIDTH = 1_000;
const RESOURCE_NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_.+\-]{0,63}$/;
const PATH_OPERATORS = new Set(["m", "l", "c", "v", "y", "h", "re"]);
const PAINT_OPERATORS = new Set([
  "S",
  "s",
  "f",
  "F",
  "f*",
  "B",
  "B*",
  "b",
  "b*",
]);
const ALLOWED_GROUP_OPERATORS = new Set([
  "q",
  "Q",
  "cm",
  "g",
  "rg",
  "k",
  "G",
  "RG",
  "K",
  "w",
  "d",
  "J",
  "j",
  "M",
  "ri",
  "i",
  "n",
  ...PATH_OPERATORS,
  ...PAINT_OPERATORS,
]);

function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}

function invalid(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

function latin1(bytes: Uint8Array): string {
  let value = "";
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    value += String.fromCharCode(...bytes.subarray(offset, offset + 32_768));
  }
  return value;
}

function ascii(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function isWhitespace(character: string): boolean {
  return character === "\0" || /[\t\n\f\r ]/.test(character);
}

function isDelimiter(character: string): boolean {
  return "()<>[]{}/%".includes(character);
}

function skipWhitespaceAndComments(source: string, start: number): number {
  let cursor = start;
  while (cursor < source.length) {
    if (isWhitespace(source[cursor])) {
      cursor += 1;
      continue;
    }
    if (source[cursor] === "%") {
      while (
        cursor < source.length &&
        source[cursor] !== "\n" &&
        source[cursor] !== "\r"
      ) {
        cursor += 1;
      }
      continue;
    }
    break;
  }
  return cursor;
}

function literalEnd(source: string, start: number): number {
  let cursor = start + 1;
  let depth = 1;
  while (cursor < source.length) {
    if (source[cursor] === "\\") {
      cursor += 2;
      continue;
    }
    if (source[cursor] === "(") depth += 1;
    else if (source[cursor] === ")") {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
    cursor += 1;
  }
  refuse("PDF vector style literal string is unterminated.");
}

function hexEnd(source: string, start: number): number {
  const end = source.indexOf(">", start + 1);
  if (end < 0) refuse("PDF vector style hex string is unterminated.");
  return end + 1;
}

function containerEnd(
  source: string,
  start: number,
  opening: "[" | "<<",
): number {
  const stack: Array<"[" | "<<"> = [opening];
  let cursor = start + opening.length;
  while (cursor < source.length) {
    cursor = skipWhitespaceAndComments(source, cursor);
    if (cursor >= source.length) break;
    if (source[cursor] === "(") {
      cursor = literalEnd(source, cursor);
      continue;
    }
    if (source[cursor] === "<" && source[cursor + 1] !== "<") {
      cursor = hexEnd(source, cursor);
      continue;
    }
    const pair = source.slice(cursor, cursor + 2);
    if (pair === "<<") {
      stack.push("<<");
      cursor += 2;
      continue;
    }
    if (pair === ">>") {
      if (stack.pop() !== "<<")
        refuse("PDF vector style dictionary is unbalanced.");
      cursor += 2;
      if (stack.length === 0) return cursor;
      continue;
    }
    if (source[cursor] === "[") {
      stack.push("[");
      cursor += 1;
      continue;
    }
    if (source[cursor] === "]") {
      if (stack.pop() !== "[") refuse("PDF vector style array is unbalanced.");
      cursor += 1;
      if (stack.length === 0) return cursor;
      continue;
    }
    cursor += 1;
  }
  refuse("PDF vector style container is unterminated.");
}

function tokenEnd(source: string, start: number): number {
  let cursor = start;
  while (
    cursor < source.length &&
    !isWhitespace(source[cursor]) &&
    !isDelimiter(source[cursor])
  ) {
    cursor += 1;
  }
  if (cursor === start) refuse("PDF vector style token is malformed.");
  return cursor;
}

function nameEnd(source: string, start: number): number {
  const end = tokenEnd(source, start + 1);
  if (end === start + 1) refuse("PDF vector style name is empty.");
  return end;
}

function tokenize(source: string): readonly ContentToken[] {
  const tokens: ContentToken[] = [];
  let cursor = 0;
  while (cursor < source.length) {
    cursor = skipWhitespaceAndComments(source, cursor);
    if (cursor >= source.length) break;
    if (tokens.length >= MAX_CONTENT_TOKENS) {
      refuse("PDF vector style content has too many tokens.");
    }
    const start = cursor;
    const character = source[cursor];
    if (character === "(") {
      cursor = literalEnd(source, cursor);
      tokens.push({
        kind: "value",
        value: source.slice(start, cursor),
        start,
        end: cursor,
      });
      continue;
    }
    if (character === "[") {
      cursor = containerEnd(source, cursor, "[");
      tokens.push({
        kind: "value",
        value: source.slice(start, cursor),
        start,
        end: cursor,
      });
      continue;
    }
    if (character === "<" && source[cursor + 1] === "<") {
      cursor = containerEnd(source, cursor, "<<");
      tokens.push({
        kind: "value",
        value: source.slice(start, cursor),
        start,
        end: cursor,
      });
      continue;
    }
    if (character === "<") {
      cursor = hexEnd(source, cursor);
      tokens.push({
        kind: "value",
        value: source.slice(start, cursor),
        start,
        end: cursor,
      });
      continue;
    }
    if (character === "/") {
      cursor = nameEnd(source, cursor);
      tokens.push({
        kind: "name",
        value: source.slice(start, cursor),
        start,
        end: cursor,
      });
      continue;
    }
    if (isDelimiter(character))
      refuse("PDF vector style delimiter is unsupported.");
    cursor = tokenEnd(source, cursor);
    const value = source.slice(start, cursor);
    tokens.push({
      kind: NUMBER_PATTERN.test(value)
        ? "number"
        : value === "true" || value === "false" || value === "null"
          ? "value"
          : "operator",
      value,
      start,
      end: cursor,
    });
  }
  return Object.freeze(tokens.map((token) => Object.freeze(token)));
}

function operations(source: string): readonly ContentOperation[] {
  const result: ContentOperation[] = [];
  let operands: ContentToken[] = [];
  for (const token of tokenize(source)) {
    if (token.kind !== "operator") {
      operands.push(token);
      continue;
    }
    if (["BI", "ID", "EI"].includes(token.value)) {
      refuse("Inline images make vector style ownership ambiguous.");
    }
    result.push(
      Object.freeze({
        operator: token.value,
        operands: Object.freeze(operands),
        start: operands[0]?.start ?? token.start,
        end: token.end,
      }),
    );
    operands = [];
  }
  if (operands.length !== 0)
    refuse("PDF vector style ends inside an operation.");
  return Object.freeze(result);
}

function numericOperands(operation: ContentOperation, count: number): number[] {
  if (
    operation.operands.length !== count ||
    operation.operands.some((operand) => operand.kind !== "number")
  ) {
    refuse(`PDF ${operation.operator} vector style operands are invalid.`);
  }
  const values = operation.operands.map((operand) => Number(operand.value));
  if (
    values.some(
      (value) => !Number.isFinite(value) || Math.abs(value) > MAX_ABS_NUMBER,
    )
  ) {
    refuse(`PDF ${operation.operator} vector style values are unsupported.`);
  }
  return values;
}

function exactComponent(value: number, label: string): number {
  if (!Number.isFinite(value) || value < 0 || value > 1) {
    invalid(`Invalid PDF vector ${label} component.`);
  }
  return value;
}

function exactColor(
  value: PdfNativeVectorRgbColor,
  label: string,
): PdfNativeVectorRgbColor {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    invalid(`Invalid PDF vector ${label} color.`);
  }
  return Object.freeze({
    r: exactComponent(value.r, label),
    g: exactComponent(value.g, label),
    b: exactComponent(value.b, label),
  });
}

function colorFromOperation(
  operation: ContentOperation,
): PdfNativeVectorRgbColor {
  if (operation.operator === "g" || operation.operator === "G") {
    const [gray] = numericOperands(operation, 1);
    const value = exactComponent(gray, "gray");
    return Object.freeze({ r: value, g: value, b: value });
  }
  if (operation.operator === "rg" || operation.operator === "RG") {
    const [r, g, b] = numericOperands(operation, 3);
    return exactColor({ r, g, b }, "RGB");
  }
  const [c, m, y, k] = numericOperands(operation, 4).map((component) =>
    exactComponent(component, "CMYK"),
  );
  return Object.freeze({
    r: 1 - Math.min(1, c + k),
    g: 1 - Math.min(1, m + k),
    b: 1 - Math.min(1, y + k),
  });
}

type PdfNativeVectorObjectStyleSelection =
  PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;

/**
 * Colour space a channel was last set in. PDF's initial colour space for both
 * channels is DeviceGray, so that is the default; `cs`/`scn` selections leave
 * the channel unknown and read-only.
 */
type ObjectStyleColorSpace = "gray" | "rgb" | "cmyk" | null;

interface EffectiveObjectStyleState {
  readonly fillColor: PdfNativeVectorRgbColor | null;
  readonly strokeColor: PdfNativeVectorRgbColor | null;
  readonly fillColorSpace: ObjectStyleColorSpace;
  readonly strokeColorSpace: ObjectStyleColorSpace;
  /** The `k`/`K` components themselves, while that space holds. */
  readonly fillCmyk: PdfNativeVectorCmyk | null;
  readonly strokeCmyk: PdfNativeVectorCmyk | null;
  readonly lineWidth: number | null;
  readonly dashArray: readonly number[] | null;
  readonly dashPhase: number | null;
  readonly lineCap: PdfNativeVectorObjectStyleLineCap | null;
  readonly lineJoin: PdfNativeVectorObjectStyleLineJoin | null;
  readonly miterLimit: number | null;
  readonly fillOpacity: number | null;
  readonly strokeOpacity: number | null;
  readonly blendMode: PdfStandardBlendMode | null;
  readonly hasExternalGraphicsState: boolean;
  /** The fill was last set through `cs`/`sc`/`scn` (a pattern, a named
   *  colour space): not a colour this editor reads or rewrites. */
  readonly fillUnreadable: boolean;
}

function initialObjectStyleState(): EffectiveObjectStyleState {
  return Object.freeze({
    fillColor: Object.freeze({ r: 0, g: 0, b: 0 }),
    strokeColor: Object.freeze({ r: 0, g: 0, b: 0 }),
    fillColorSpace: "gray",
    strokeColorSpace: "gray",
    fillCmyk: null,
    strokeCmyk: null,
    lineWidth: 1,
    dashArray: Object.freeze([]),
    dashPhase: 0,
    lineCap: "butt",
    lineJoin: "miter",
    miterLimit: 10,
    fillOpacity: 1,
    strokeOpacity: 1,
    blendMode: "normal",
    hasExternalGraphicsState: false,
    fillUnreadable: false,
  });
}

function updateObjectStyleState(
  state: EffectiveObjectStyleState,
  change: Partial<EffectiveObjectStyleState>,
): EffectiveObjectStyleState {
  return Object.freeze({ ...state, ...change });
}

function colorSpaceForOperator(operator: string): ObjectStyleColorSpace {
  if (operator === "g" || operator === "G") return "gray";
  if (operator === "k" || operator === "K") return "cmyk";
  return "rgb";
}

function cmykFromOperation(operation: ContentOperation): PdfNativeVectorCmyk | null {
  try {
    const [c, m, y, k] = numericOperands(operation, 4).map((component) =>
      exactComponent(component, "CMYK"),
    );
    return Object.freeze([c, m, y, k] as const);
  } catch {
    return null;
  }
}

function safeColorFromOperation(
  operation: ContentOperation,
): PdfNativeVectorRgbColor | null {
  try {
    return colorFromOperation(operation);
  } catch {
    return null;
  }
}

function numericArray(value: string): readonly number[] | null {
  if (!value.startsWith("[") || !value.endsWith("]")) return null;
  const body = value
    .slice(1, -1)
    .replace(/%[^\r\n]*/g, " ")
    .trim();
  if (!body) return Object.freeze([]);
  const tokens = body.split(/[\t\n\f\r ]+/);
  if (
    tokens.length > 64 ||
    tokens.some((token) => !NUMBER_PATTERN.test(token))
  ) {
    return null;
  }
  const values = tokens.map(Number);
  return values.every(
    (value) => Number.isFinite(value) && value >= 0 && value <= MAX_ABS_NUMBER,
  )
    ? Object.freeze(values)
    : null;
}

function graphicsStateResourceName(operation: ContentOperation): string | null {
  if (
    operation.operands.length !== 1 ||
    operation.operands[0].kind !== "name"
  ) {
    return null;
  }
  const raw = operation.operands[0].value.slice(1);
  let decoded = "";
  for (let index = 0; index < raw.length; index += 1) {
    if (raw[index] !== "#") {
      decoded += raw[index];
      continue;
    }
    const escaped = raw.slice(index + 1, index + 3);
    if (!/^[0-9a-f]{2}$/i.test(escaped)) return null;
    decoded += String.fromCharCode(Number.parseInt(escaped, 16));
    index += 2;
  }
  return RESOURCE_NAME_PATTERN.test(decoded) ? decoded : null;
}

/** One operation's effect on the graphics state, exactly as the style walk
 * applied it: `stack` is the q/Q stack, mutated in place. */
function stepObjectStyleState(
  state: EffectiveObjectStyleState,
  stack: EffectiveObjectStyleState[],
  operation: ContentOperation,
  graphicsStates: ReadonlyMap<string, PdfNativePageExtGStateEntryInspection>,
): EffectiveObjectStyleState {
  const operator = operation.operator;
  if (operator === "q") {
    if (operation.operands.length !== 0)
      refuse("PDF q style operands are invalid.");
    stack.push(state);
    return state;
  }
  if (operator === "Q") {
    if (operation.operands.length !== 0 || stack.length === 0) {
      refuse("PDF vector style graphics-state restore is unmatched.");
    }
    state = stack.pop()!;
    return state;
  }
  if (["g", "rg", "k"].includes(operator)) {
    const fillColor = safeColorFromOperation(operation);
    state = updateObjectStyleState(state, {
      fillColor,
      fillColorSpace: fillColor ? colorSpaceForOperator(operator) : null,
      fillCmyk: fillColor && operator === "k" ? cmykFromOperation(operation) : null,
      fillUnreadable: false,
    });
    return state;
  }
  if (["G", "RG", "K"].includes(operator)) {
    const strokeColor = safeColorFromOperation(operation);
    state = updateObjectStyleState(state, {
      strokeColor,
      strokeColorSpace: strokeColor ? colorSpaceForOperator(operator) : null,
      strokeCmyk: strokeColor && operator === "K" ? cmykFromOperation(operation) : null,
    });
    return state;
  }
  if (operator === "cs" || operator === "sc" || operator === "scn") {
    // A gradient (a Pattern fill) read as 「无填充」 and a colour chosen for
    // it was accepted and written nowhere (user audit PDF 0145): the fill
    // is named unreadable, and the fill controls say so.
    state = updateObjectStyleState(state, {
      fillColor: null,
      fillColorSpace: null,
      fillCmyk: null,
      fillUnreadable: true,
    });
    return state;
  }
  if (operator === "CS" || operator === "SC" || operator === "SCN") {
    state = updateObjectStyleState(state, {
      strokeColor: null,
      strokeColorSpace: null,
      strokeCmyk: null,
    });
    return state;
  }
  if (operator === "w") {
    let lineWidth: number | null = null;
    try {
      const value = numericOperands(operation, 1)[0];
      lineWidth = value >= 0 && value <= MAX_LINE_WIDTH ? value : null;
    } catch {
      lineWidth = null;
    }
    state = updateObjectStyleState(state, { lineWidth });
    return state;
  }
  if (operator === "d") {
    const array =
      operation.operands.length === 2 &&
      operation.operands[0].kind === "value"
        ? numericArray(operation.operands[0].value)
        : null;
    const phase =
      operation.operands.length === 2 &&
      operation.operands[1].kind === "number"
        ? Number(operation.operands[1].value)
        : Number.NaN;
    state = updateObjectStyleState(state, {
      dashArray: array,
      dashPhase:
        Number.isFinite(phase) && Math.abs(phase) <= MAX_ABS_NUMBER
          ? phase
          : null,
    });
    return state;
  }
  if (operator === "J" || operator === "j") {
    const value =
      operation.operands.length === 1 &&
      operation.operands[0].kind === "number"
        ? Number(operation.operands[0].value)
        : Number.NaN;
    const mapped =
      value === 0
        ? operator === "J"
          ? "butt"
          : "miter"
        : value === 1
          ? "round"
          : value === 2
            ? operator === "J"
              ? "square"
              : "bevel"
            : null;
    state =
      operator === "J"
        ? updateObjectStyleState(state, {
            lineCap:
              mapped === "butt" || mapped === "round" || mapped === "square"
                ? mapped
                : null,
          })
        : updateObjectStyleState(state, {
            lineJoin:
              mapped === "miter" || mapped === "round" || mapped === "bevel"
                ? mapped
                : null,
          });
    return state;
  }
  if (operator === "M") {
    const value =
      operation.operands.length === 1 &&
      operation.operands[0].kind === "number"
        ? Number(operation.operands[0].value)
        : Number.NaN;
    state = updateObjectStyleState(state, {
      miterLimit:
        Number.isFinite(value) && value >= 1 && value <= MAX_ABS_NUMBER
          ? value
          : null,
    });
    return state;
  }
  if (operator === "gs") {
    const name = graphicsStateResourceName(operation);
    const graphicsState = name ? graphicsStates.get(name) : undefined;
    if (graphicsState?.supportedAppearanceOnly) {
      state = updateObjectStyleState(state, {
        ...(graphicsState.fillOpacity === null
          ? {}
          : { fillOpacity: graphicsState.fillOpacity }),
        ...(graphicsState.strokeOpacity === null
          ? {}
          : { strokeOpacity: graphicsState.strokeOpacity }),
        ...(graphicsState.blendMode === null
          ? {}
          : { blendMode: graphicsState.blendMode }),
        hasExternalGraphicsState: true,
      });
      return state;
    }
    state = updateObjectStyleState(state, {
      lineWidth: null,
      dashArray: null,
      dashPhase: null,
      lineCap: null,
      lineJoin: null,
      miterLimit: null,
      fillOpacity: null,
      strokeOpacity: null,
      blendMode: null,
      hasExternalGraphicsState: true,
    });
  }
  return state;
}

/** The style state before each operation of one content stream, walked once
 * and read by index. Every shape on a page used to replay the stream from its
 * first operation, so a page of five thousand rules cost five thousand walks
 * (the whole document's style pass took longer than its content pass).
 * Reading index k yields what a fresh walk over operations [0, k) yielded,
 * refusals included: an operation the walk refuses (an unmatched Q, bad q
 * operands) refuses every read past it, and only those. */
class ObjectStyleStateIndex {
  readonly #operations: readonly ContentOperation[];
  readonly #graphicsStates: ReadonlyMap<string, PdfNativePageExtGStateEntryInspection>;
  /** `#states[i]` is the state before operation `i`. */
  readonly #states: EffectiveObjectStyleState[] = [initialObjectStyleState()];
  readonly #stack: EffectiveObjectStyleState[] = [];
  #refusal: { readonly index: number; readonly error: unknown } | null = null;

  constructor(
    operations: readonly ContentOperation[],
    graphicsStates: ReadonlyMap<string, PdfNativePageExtGStateEntryInspection>,
  ) {
    this.#operations = operations;
    this.#graphicsStates = graphicsStates;
  }

  before(endIndex: number): EffectiveObjectStyleState {
    while (this.#states.length <= endIndex) {
      const index = this.#states.length - 1;
      if (this.#refusal && this.#refusal.index <= index) throw this.#refusal.error;
      if (index >= this.#operations.length) break;
      let next: EffectiveObjectStyleState;
      try {
        next = stepObjectStyleState(
          this.#states[index],
          this.#stack,
          this.#operations[index],
          this.#graphicsStates,
        );
      } catch (error) {
        this.#refusal = { index, error };
        throw error;
      }
      this.#states.push(next);
    }
    return this.#states[Math.min(endIndex, this.#states.length - 1)];
  }
}

function styleStateBefore(
  parsedOperations: readonly ContentOperation[],
  endIndex: number,
  graphicsStates: ReadonlyMap<string, PdfNativePageExtGStateEntryInspection>,
): EffectiveObjectStyleState {
  return new ObjectStyleStateIndex(parsedOperations, graphicsStates).before(endIndex);
}

function paintModeForOperator(
  operator: string,
): PdfNativeVectorContentShape["paintMode"] | null {
  if (["S", "s"].includes(operator)) return "stroke";
  if (["f", "F", "f*"].includes(operator)) return "fill";
  if (["B", "B*", "b", "b*"].includes(operator)) return "fill-stroke";
  return null;
}

function publicObjectStyle(
  state: EffectiveObjectStyleState,
  paintMode: PdfNativeVectorContentShape["paintMode"],
): PdfNativeVectorObjectStyle {
  const strokes = paintMode === "stroke" || paintMode === "fill-stroke";
  const fills = paintMode === "fill" || paintMode === "fill-stroke";
  const opacity =
    fills && strokes
      ? state.fillOpacity !== null &&
        state.strokeOpacity !== null &&
        Math.abs(state.fillOpacity - state.strokeOpacity) <= 0.000_001
        ? state.fillOpacity
        : null
      : fills
        ? state.fillOpacity
        : strokes
          ? state.strokeOpacity
          : null;
  return Object.freeze({
    fillColor: fills ? state.fillColor : null,
    strokeColor: strokes ? state.strokeColor : null,
    lineWidth: strokes ? state.lineWidth : null,
    dashArray: strokes ? state.dashArray : null,
    dashPhase: strokes ? state.dashPhase : null,
    lineCap: strokes ? state.lineCap : null,
    lineJoin: strokes ? state.lineJoin : null,
    miterLimit: strokes ? state.miterLimit : null,
    opacity,
    blendMode: state.blendMode,
    hasExternalGraphicsState: state.hasExternalGraphicsState,
    fillUnreadable: fills ? state.fillUnreadable : false,
  });
}

function qPairs(ops: readonly ContentOperation[]): ReadonlyMap<number, number> {
  const stack: number[] = [];
  const pairs = new Map<number, number>();
  ops.forEach((operation, index) => {
    if (operation.operator === "q") {
      if (operation.operands.length !== 0)
        refuse("PDF q style operands are invalid.");
      stack.push(index);
    } else if (operation.operator === "Q") {
      if (operation.operands.length !== 0 || stack.length === 0) {
        refuse("PDF vector style graphics-state stack is unbalanced.");
      }
      pairs.set(stack.pop()!, index);
    }
  });
  if (stack.length !== 0)
    refuse("PDF vector style graphics-state stack is unbalanced.");
  return pairs;
}

function styleEnvelope(
  streamBytes: Uint8Array,
  selection: PdfNativeVectorContentShape,
): Readonly<{
  operations: readonly ContentOperation[];
  envelope: StyleEnvelope;
}> {
  const source = latin1(streamBytes);
  const ops = operations(source);
  const firstPathIndex = ops.findIndex(
    (operation) =>
      PATH_OPERATORS.has(operation.operator) &&
      operation.start === selection.operatorSpan.start,
  );
  if (firstPathIndex < 0)
    refuse("Selected PDF vector path start is unavailable.");
  let lastPathIndex = firstPathIndex;
  while (
    lastPathIndex + 1 < ops.length &&
    PATH_OPERATORS.has(ops[lastPathIndex + 1].operator) &&
    ops[lastPathIndex + 1].end <= selection.operatorSpan.end
  ) {
    lastPathIndex += 1;
  }
  if (ops[lastPathIndex].end !== selection.operatorSpan.end) {
    refuse("Selected PDF vector path boundary changed.");
  }
  const pairs = qPairs(ops);
  const qStack: number[] = [];
  for (let index = 0; index < firstPathIndex; index += 1) {
    if (ops[index].operator === "q") qStack.push(index);
    else if (ops[index].operator === "Q") qStack.pop();
  }
  for (let stackIndex = qStack.length - 1; stackIndex >= 0; stackIndex -= 1) {
    const qIndex = qStack[stackIndex];
    const closeIndex = pairs.get(qIndex);
    if (closeIndex === undefined || closeIndex <= lastPathIndex) continue;
    const group = ops.slice(qIndex + 1, closeIndex);
    if (
      group.some(
        (operation) => !ALLOWED_GROUP_OPERATORS.has(operation.operator),
      )
    ) {
      continue;
    }
    const paints = group
      .map((operation, offset) => ({ operation, index: qIndex + 1 + offset }))
      .filter(({ operation }) => PAINT_OPERATORS.has(operation.operator));
    if (paints.length !== 1 || paints[0].index <= lastPathIndex) continue;
    const outsideSelectedPath = group.filter(
      (operation) =>
        PATH_OPERATORS.has(operation.operator) &&
        (operation.start < selection.operatorSpan.start ||
          operation.end > selection.operatorSpan.end),
    );
    if (outsideSelectedPath.length !== 0) continue;
    const betweenPathAndPaint = ops.slice(lastPathIndex + 1, paints[0].index);
    if (betweenPathAndPaint.some((operation) => operation.operator !== "Q")) {
      continue;
    }
    const beforePath = ops.slice(qIndex + 1, firstPathIndex);
    if (beforePath.some((operation) => operation.operator === "gs")) continue;
    const fills = beforePath.filter((operation) =>
      ["g", "rg", "k"].includes(operation.operator),
    );
    const strokes = beforePath.filter((operation) =>
      ["G", "RG", "K"].includes(operation.operator),
    );
    const widths = beforePath.filter((operation) => operation.operator === "w");
    const needsFill =
      selection.paintMode === "fill" || selection.paintMode === "fill-stroke";
    const needsStroke =
      selection.paintMode === "stroke" || selection.paintMode === "fill-stroke";
    if (
      (needsFill && fills.length !== 1) ||
      (!needsFill && fills.length > 1) ||
      (needsStroke && (strokes.length !== 1 || widths.length !== 1)) ||
      (!needsStroke && (strokes.length > 1 || widths.length > 1))
    ) {
      continue;
    }
    const fillOperation = needsFill ? fills[0] : null;
    const strokeOperation = needsStroke ? strokes[0] : null;
    const lineWidthOperation = needsStroke ? widths[0] : null;
    const lineWidth = lineWidthOperation
      ? numericOperands(lineWidthOperation, 1)[0]
      : null;
    if (
      lineWidth !== null &&
      (!(lineWidth > 0) || lineWidth > MAX_LINE_WIDTH)
    ) {
      continue;
    }
    return Object.freeze({
      operations: ops,
      envelope: Object.freeze({
        start: ops[qIndex].end,
        end: selection.operatorSpan.start,
        fillOperation,
        strokeOperation,
        lineWidthOperation,
        style: Object.freeze({
          fillColor: fillOperation ? colorFromOperation(fillOperation) : null,
          strokeColor: strokeOperation
            ? colorFromOperation(strokeOperation)
            : null,
          lineWidth,
        }),
      }),
    });
  }
  refuse("Selected PDF vector style is inherited or structurally ambiguous.");
}

function rounded(value: number): number {
  const result = Number(value.toFixed(6));
  return Object.is(result, -0) ? 0 : result;
}

function canonicalStyle(style: PdfNativeVectorStyle): unknown {
  const color = (value: PdfNativeVectorRgbColor | null) =>
    value
      ? { r: rounded(value.r), g: rounded(value.g), b: rounded(value.b) }
      : null;
  return {
    fillColor: color(style.fillColor),
    strokeColor: color(style.strokeColor),
    lineWidth: style.lineWidth === null ? null : rounded(style.lineWidth),
  };
}

function geometryFingerprint(geometry: PdfNativeVectorGeometry): string {
  return JSON.stringify(geometry, (_key, value) =>
    typeof value === "number" ? rounded(value) : value,
  );
}

async function visibleProof(
  selection: PdfNativeVectorContentShape,
  envelope: StyleEnvelope,
  streamInspection: PdfNativeContentStreamInspection,
): Promise<PdfNativeVectorStyleProof> {
  const bytes = streamInspection.streamBytes.slice(
    envelope.start,
    envelope.end,
  );
  const expectedSha256 = await sha256PdfBytes(bytes);
  const proofDigest = await sha256PdfBytes(
    ascii(
      JSON.stringify({
        schemaVersion: MRAI_NATIVE_VECTOR_STYLE_SCHEMA_VERSION,
        sourceSha256: selection.sourceSha256,
        pageNumber: selection.pageNumber,
        pageVectorIndex: selection.pageVectorIndex,
        selectionId: selection.id,
        selectionProofDigest: selection.proofDigest,
        contentStream: selection.contentStream,
        envelopeSpan: {
          start: envelope.start,
          end: envelope.end,
          expectedSha256,
        },
        paintMode: selection.paintMode,
        style: canonicalStyle(envelope.style),
      }),
    ),
  );
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_STYLE_SCHEMA_VERSION,
    id: `pdf-native-vector-style:${selection.sourceSha256}:p${selection.pageNumber}:v${selection.pageVectorIndex}:${proofDigest}`,
    sourceSha256: selection.sourceSha256,
    pageNumber: selection.pageNumber,
    pageVectorIndex: selection.pageVectorIndex,
    selectionId: selection.id,
    selectionProofDigest: selection.proofDigest,
    contentStream: selection.contentStream,
    envelopeSpan: Object.freeze({
      start: envelope.start,
      end: envelope.end,
      expectedSha256,
    }),
    paintMode: selection.paintMode,
    style: envelope.style,
    proofDigest,
  });
}

function canonicalObjectStyle(style: PdfNativeVectorObjectStyle): unknown {
  const canonicalColor = (value: PdfNativeVectorRgbColor | null) =>
    value
      ? { r: rounded(value.r), g: rounded(value.g), b: rounded(value.b) }
      : null;
  return {
    fillColor: canonicalColor(style.fillColor),
    strokeColor: canonicalColor(style.strokeColor),
    lineWidth: style.lineWidth === null ? null : rounded(style.lineWidth),
    dashArray: style.dashArray?.map(rounded) ?? null,
    dashPhase: style.dashPhase === null ? null : rounded(style.dashPhase),
    lineCap: style.lineCap,
    lineJoin: style.lineJoin,
    miterLimit: style.miterLimit === null ? null : rounded(style.miterLimit),
    opacity: style.opacity === null ? null : rounded(style.opacity),
    blendMode: style.blendMode,
    hasExternalGraphicsState: style.hasExternalGraphicsState,
  };
}

/** The index of the first path operation starting at each byte offset, built
 * once per stream so a selection is found without scanning the whole stream. */
function pathOperationIndexByStart(
  parsedOperations: readonly ContentOperation[],
): ReadonlyMap<number, number> {
  const byStart = new Map<number, number>();
  parsedOperations.forEach((operation, index) => {
    if (PATH_OPERATORS.has(operation.operator) && !byStart.has(operation.start)) {
      byStart.set(operation.start, index);
    }
  });
  return byStart;
}

async function objectStyleProof(
  streamInspection: PdfNativeContentStreamInspection,
  parsedOperations: readonly ContentOperation[],
  selection: PdfNativeVectorObjectStyleSelection,
  graphicsStates: ReadonlyMap<string, PdfNativePageExtGStateEntryInspection>,
  shared?: {
    readonly states: ObjectStyleStateIndex;
    readonly pathIndexByStart: ReadonlyMap<number, number>;
  },
): Promise<PdfNativeVectorObjectStyleProof | null> {
  if (
    !selection.editable ||
    selection.preserveReason !== null ||
    selection.contentStream.objectNumber !==
      streamInspection.identity.objectNumber ||
    selection.contentStream.generation !== streamInspection.identity.generation
  ) {
    return null;
  }
  const firstPathIndex = shared
    ? (shared.pathIndexByStart.get(selection.operatorSpan.start) ?? -1)
    : parsedOperations.findIndex(
        (operation) =>
          PATH_OPERATORS.has(operation.operator) &&
          operation.start === selection.operatorSpan.start,
      );
  if (firstPathIndex < 0) return null;
  let lastPathIndex = firstPathIndex;
  while (
    lastPathIndex + 1 < parsedOperations.length &&
    PATH_OPERATORS.has(parsedOperations[lastPathIndex + 1].operator) &&
    parsedOperations[lastPathIndex + 1].end <= selection.operatorSpan.end
  ) {
    lastPathIndex += 1;
  }
  if (parsedOperations[lastPathIndex].end !== selection.operatorSpan.end) {
    return null;
  }
  const paint = parsedOperations[lastPathIndex + 1];
  if (
    !paint ||
    paint.operands.length !== 0 ||
    paintModeForOperator(paint.operator) !== selection.paintMode ||
    ("paintOperator" in selection && paint.operator !== selection.paintOperator)
  ) {
    return null;
  }
  const state = shared
    ? shared.states.before(firstPathIndex)
    : styleStateBefore(parsedOperations, firstPathIndex, graphicsStates);
  const style = publicObjectStyle(state, selection.paintMode);
  const fillCmyk = style.fillColor ? state.fillCmyk : null;
  const strokeCmyk = style.strokeColor ? state.strokeCmyk : null;
  const paintBytes = streamInspection.streamBytes.slice(paint.start, paint.end);
  const paintSha256 = await sha256PdfBytes(paintBytes);
  const proofFamily = "kind" in selection ? "shape" : "subpath";
  const pageProofIndex =
    proofFamily === "shape"
      ? (selection as PdfNativeVectorContentShape).pageVectorIndex
      : (selection as PdfNativeVectorSubpathProof).pageSubpathIndex;
  const proofDigest = await sha256PdfBytes(
    ascii(
      JSON.stringify({
        schemaVersion: MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
        sourceSha256: selection.sourceSha256,
        pageNumber: selection.pageNumber,
        proofFamily,
        pageProofIndex,
        selectionId: selection.id,
        selectionProofDigest: selection.proofDigest,
        contentStream: selection.contentStream,
        pathSpan: selection.operatorSpan,
        paintSpan: {
          start: paint.start,
          end: paint.end,
          expectedSha256: paintSha256,
        },
        paintMode: selection.paintMode,
        paintOperator: paint.operator,
        style: canonicalObjectStyle(style),
      }),
    ),
  );
  const proof: PdfNativeVectorObjectStyleProof = Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
    id: `pdf-native-vector-object-style:${selection.sourceSha256}:p${selection.pageNumber}:${proofFamily}:${pageProofIndex}:${proofDigest}`,
    sourceSha256: selection.sourceSha256,
    pageNumber: selection.pageNumber,
    proofFamily,
    pageProofIndex,
    selectionId: selection.id,
    selectionProofDigest: selection.proofDigest,
    contentStream: selection.contentStream,
    pathSpan: selection.operatorSpan,
    paintSpan: Object.freeze({
      start: paint.start,
      end: paint.end,
      expectedSha256: paintSha256,
    }),
    paintMode: selection.paintMode,
    paintOperator: paint.operator,
    style,
    ...(fillCmyk || strokeCmyk
      ? { deviceCmyk: Object.freeze({ fill: fillCmyk, stroke: strokeCmyk }) }
      : {}),
    proofDigest,
  });
  proofStates.set(proof, state);
  return proof;
}

async function objectStylePageFromAuthority(
  nativePage: PdfNativeVectorDocumentAuthorityInspection["document"]["pages"][number],
  authorityPage: PdfNativeVectorDocumentAuthorityInspection["pages"][number],
  graphicsStates: readonly PdfNativePageExtGStateEntryInspection[],
): Promise<PdfNativeVectorObjectStylePageInspection> {
  const selections: PdfNativeVectorObjectStyleSelection[] = [
    ...nativePage.shapes,
    ...nativePage.subpaths,
  ];
  const streamInspection = authorityPage.contentStream;
  // A form other pages reference too is edited through a clone of it; a
  // style edit does not write one, so such a form's vectors keep their style.
  if (
    authorityPage.pageNumber !== nativePage.pageNumber ||
    !streamInspection ||
    "formInstanceIsolation" in streamInspection
  ) {
    return Object.freeze({
      pageNumber: nativePage.pageNumber,
      proofs: Object.freeze([]),
      unsupportedSelectionIds: Object.freeze(
        selections.map((selection) => selection.id),
      ),
    });
  }
  let parsedOperations: readonly ContentOperation[];
  try {
    parsedOperations = operations(latin1(streamInspection.streamBytes));
  } catch {
    return Object.freeze({
      pageNumber: nativePage.pageNumber,
      proofs: Object.freeze([]),
      unsupportedSelectionIds: Object.freeze(
        selections.map((selection) => selection.id),
      ),
    });
  }
  const graphicsStateByName = new Map(
    graphicsStates.map((entry) => [entry.resourceName, entry]),
  );
  const shared = {
    states: new ObjectStyleStateIndex(parsedOperations, graphicsStateByName),
    pathIndexByStart: pathOperationIndexByStart(parsedOperations),
  };
  const resolved = await Promise.all(
    selections.map((selection) =>
      objectStyleProof(
        streamInspection,
        parsedOperations,
        selection,
        graphicsStateByName,
        shared,
      ).catch(() => null),
    ),
  );
  const proofs = resolved.filter(
    (proof): proof is PdfNativeVectorObjectStyleProof => proof !== null,
  );
  const supportedIds = new Set(proofs.map((proof) => proof.selectionId));
  return Object.freeze({
    pageNumber: nativePage.pageNumber,
    proofs: Object.freeze(proofs),
    unsupportedSelectionIds: Object.freeze(
      selections
        .filter((selection) => !supportedIds.has(selection.id))
        .map((selection) => selection.id),
    ),
  });
}

export async function inspectMraiNativeVectorObjectStyleDocument(
  sourceBytes: Uint8Array,
  suppliedAuthority?: PdfNativeVectorDocumentAuthorityInspection,
): Promise<PdfNativeVectorObjectStyleDocumentInspection> {
  const authority =
    suppliedAuthority ??
    (await inspectMraiNativeVectorDocumentAuthority(sourceBytes));
  const sourceSha256 = await sha256PdfBytes(sourceBytes);
  const pageGraphicsStates = authority.pages.map(
    (page) => page.pageGraph.extGStates,
  );
  if (
    sourceSha256 !== authority.document.sourceSha256 ||
    authority.pages.length !== authority.document.pageCount ||
    pageGraphicsStates.length !== authority.document.pageCount ||
    pageGraphicsStates.some(
      (page, index) =>
        !page ||
        page.sourceSha256 !== sourceSha256 ||
        page.pageNumber !== authority.document.pages[index]?.pageNumber,
    )
  ) {
    invalid("PDF vector object-style source authority is inconsistent.");
  }
  const pages = await Promise.all(
    authority.document.pages.map((page, index) =>
      objectStylePageFromAuthority(
        page,
        authority.pages[index],
        pageGraphicsStates[index]!.entries,
      ),
    ),
  );
  return Object.freeze({
    sourceSha256,
    pageCount: pages.length,
    pages: Object.freeze(pages),
  });
}

const OBJECT_STYLE_CHANGE_KEYS = new Set([
  "fillColor",
  "strokeColor",
  "lineWidth",
  "dashArray",
  "dashPhase",
  "lineCap",
  "lineJoin",
  "miterLimit",
  "opacity",
  "blendMode",
  "paintMode",
  "translate",
  "transform",
]);
const PAINT_MODES: ReadonlySet<string> = new Set([
  "stroke",
  "fill",
  "fill-stroke",
]);

/**
 * Effective-state registry for proofs built in this module. Repainting an
 * object needs the channels its current paint mode hides (a fill-only shape
 * still carries a stroke colour in the graphics state), so the plan looks the
 * pre-projection state up here instead of widening the public proof contract.
 */
const proofStates = new WeakMap<
  PdfNativeVectorObjectStyleProof,
  EffectiveObjectStyleState
>();

/**
 * Paint operator for a new paint mode, keeping the source's closure and
 * even-odd decisions. A fill always closes its outline, so repainting a
 * filled path as a stroke closes it explicitly (`s`) to keep the same
 * silhouette. Mirrored byte-for-byte by the server validator.
 */
export function pdfNativeVectorPaintOperatorForMode(
  currentOperator: string,
  currentMode: PdfNativeVectorContentShape["paintMode"],
  mode: PdfNativeVectorContentShape["paintMode"],
): PdfNativeVectorPaintOperator {
  const closed =
    currentOperator === "s" ||
    currentOperator === "b" ||
    currentOperator === "b*" ||
    currentMode === "fill";
  const evenOdd =
    currentOperator === "f*" ||
    currentOperator === "B*" ||
    currentOperator === "b*";
  if (mode === "stroke") return closed ? "s" : "S";
  if (mode === "fill") return evenOdd ? "f*" : "f";
  if (closed) return evenOdd ? "b*" : "b";
  return evenOdd ? "B*" : "B";
}
const BLEND_MODE_SET: ReadonlySet<string> = new Set(PDF_STANDARD_BLEND_MODES);

function exactObjectStyleChanges(
  value: PdfNativeVectorObjectStyleChanges,
): PdfNativeVectorObjectStyleChanges {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    invalid("Invalid PDF vector object-style changes.");
  }
  const descriptors = Object.getOwnPropertyDescriptors(value);
  const keys = Reflect.ownKeys(value);
  if (
    keys.length === 0 ||
    keys.some(
      (key) =>
        typeof key !== "string" ||
        !OBJECT_STYLE_CHANGE_KEYS.has(key) ||
        !descriptors[key]?.enumerable ||
        !("value" in descriptors[key]),
    )
  ) {
    invalid("Invalid PDF vector object-style changes.");
  }
  const changes = value as Record<string, unknown>;
  const result: {
    fillColor?: PdfNativeVectorRgbColor;
    strokeColor?: PdfNativeVectorRgbColor;
    lineWidth?: number;
    dashArray?: readonly number[];
    dashPhase?: number;
    lineCap?: PdfNativeVectorObjectStyleLineCap;
    lineJoin?: PdfNativeVectorObjectStyleLineJoin;
    miterLimit?: number;
    opacity?: number;
    blendMode?: PdfStandardBlendMode;
    paintMode?: PdfNativeVectorContentShape["paintMode"];
    translate?: Readonly<{ x: number; y: number }>;
    transform?: PdfNativeVectorStyleAffine;
  } = {};
  if (Object.hasOwn(changes, "transform")) {
    // A resize is the object's whole appearance scaled: no style comes with
    // it, and it must be invertible (Undo writes the bytes back).
    const transform = changes.transform as Record<string, unknown> | null;
    const entries = transform && typeof transform === "object"
      ? (["a", "b", "c", "d", "e", "f"] as const).map((key) => transform[key])
      : [];
    if (
      entries.length !== 6 ||
      entries.some(
        (entry) =>
          typeof entry !== "number" ||
          !Number.isFinite(entry) ||
          Math.abs(entry) > MAX_ABS_NUMBER,
      ) ||
      Object.keys(changes).length !== 1
    ) {
      invalid("Invalid PDF vector transform.");
    }
    const [a, b, c, d, e, f] = entries as number[];
    if (Math.abs(a * d - b * c) <= 1e-12) invalid("Invalid PDF vector transform.");
    result.transform = Object.freeze({ a, b, c, d, e, f });
  }
  if (Object.hasOwn(changes, "translate")) {
    const translate = changes.translate as { x?: unknown; y?: unknown } | null;
    if (
      !translate ||
      typeof translate !== "object" ||
      typeof translate.x !== "number" ||
      typeof translate.y !== "number" ||
      !Number.isFinite(translate.x) ||
      !Number.isFinite(translate.y) ||
      Math.abs(translate.x) > MAX_ABS_NUMBER ||
      Math.abs(translate.y) > MAX_ABS_NUMBER ||
      !Object.hasOwn(changes, "lineWidth")
    ) {
      invalid("Invalid PDF vector translation.");
    }
    result.translate = Object.freeze({ x: translate.x, y: translate.y });
  }
  if (Object.hasOwn(changes, "paintMode")) {
    const paintMode = changes.paintMode;
    if (typeof paintMode !== "string" || !PAINT_MODES.has(paintMode)) {
      invalid("Invalid PDF vector paint mode.");
    }
    result.paintMode = paintMode as PdfNativeVectorContentShape["paintMode"];
  }
  if (Object.hasOwn(changes, "fillColor")) {
    result.fillColor = exactColor(
      changes.fillColor as PdfNativeVectorRgbColor,
      "fill",
    );
  }
  if (Object.hasOwn(changes, "strokeColor")) {
    result.strokeColor = exactColor(
      changes.strokeColor as PdfNativeVectorRgbColor,
      "stroke",
    );
  }
  if (Object.hasOwn(changes, "lineWidth")) {
    const lineWidth = changes.lineWidth;
    if (
      typeof lineWidth !== "number" ||
      !Number.isFinite(lineWidth) ||
      lineWidth <= 0 ||
      lineWidth > MAX_LINE_WIDTH
    ) {
      invalid("Invalid PDF vector line width.");
    }
    result.lineWidth = lineWidth;
  }
  if (Object.hasOwn(changes, "dashArray")) {
    const dashArray = changes.dashArray;
    if (
      !Array.isArray(dashArray) ||
      dashArray.length > 64 ||
      dashArray.some(
        (entry) =>
          typeof entry !== "number" ||
          !Number.isFinite(entry) ||
          entry < 0 ||
          entry > MAX_ABS_NUMBER,
      ) ||
      (dashArray.length > 0 && dashArray.every((entry) => entry === 0))
    ) {
      invalid("Invalid PDF vector dash pattern.");
    }
    result.dashArray = Object.freeze([...dashArray]);
  }
  if (Object.hasOwn(changes, "dashPhase")) {
    const dashPhase = changes.dashPhase;
    if (
      typeof dashPhase !== "number" ||
      !Number.isFinite(dashPhase) ||
      Math.abs(dashPhase) > MAX_ABS_NUMBER
    ) {
      invalid("Invalid PDF vector dash phase.");
    }
    result.dashPhase = dashPhase;
  }
  if (Object.hasOwn(changes, "lineCap")) {
    const lineCap = changes.lineCap;
    if (lineCap !== "butt" && lineCap !== "round" && lineCap !== "square") {
      invalid("Invalid PDF vector line cap.");
    }
    result.lineCap = lineCap;
  }
  if (Object.hasOwn(changes, "lineJoin")) {
    const lineJoin = changes.lineJoin;
    if (lineJoin !== "miter" && lineJoin !== "round" && lineJoin !== "bevel") {
      invalid("Invalid PDF vector line join.");
    }
    result.lineJoin = lineJoin;
  }
  if (Object.hasOwn(changes, "miterLimit")) {
    const miterLimit = changes.miterLimit;
    if (
      typeof miterLimit !== "number" ||
      !Number.isFinite(miterLimit) ||
      miterLimit < 1 ||
      miterLimit > MAX_ABS_NUMBER
    ) {
      invalid("Invalid PDF vector miter limit.");
    }
    result.miterLimit = miterLimit;
  }
  if (Object.hasOwn(changes, "opacity")) {
    const opacity = changes.opacity;
    if (
      typeof opacity !== "number" ||
      !Number.isFinite(opacity) ||
      opacity < 0 ||
      opacity > 1
    ) {
      invalid("Invalid PDF vector opacity.");
    }
    result.opacity = opacity;
  }
  if (Object.hasOwn(changes, "blendMode")) {
    const blendMode = changes.blendMode;
    if (typeof blendMode !== "string" || !BLEND_MODE_SET.has(blendMode)) {
      invalid("Invalid PDF vector blend mode.");
    }
    result.blendMode = blendMode as PdfStandardBlendMode;
  }
  return Object.freeze(result);
}

function expectedObjectStyle(
  current: PdfNativeVectorObjectStyle,
  changes: PdfNativeVectorObjectStyleChanges,
  options?: Readonly<{ repaint?: boolean }>,
): PdfNativeVectorObjectStyle {
  // While repainting, an explicit colour operator defines a channel whose
  // source authority was unknown (non-Device colour space); the prefix writes
  // it, so the after-state is fully determined.
  const repaint = options?.repaint === true;
  const strokes = current.strokeColor !== null || (repaint && !!changes.strokeColor);
  const fills = current.fillColor !== null || (repaint && !!changes.fillColor);
  if (
    (changes.fillColor && !fills) ||
    ((changes.strokeColor ||
      changes.lineWidth !== undefined ||
      changes.dashArray !== undefined ||
      changes.dashPhase !== undefined ||
      changes.lineCap !== undefined ||
      changes.lineJoin !== undefined ||
      changes.miterLimit !== undefined) &&
      !strokes)
  ) {
    invalid("PDF vector style channel is not available on this object.");
  }
  if (
    (changes.dashArray !== undefined || changes.dashPhase !== undefined) &&
    (current.dashArray === null || current.dashPhase === null)
  ) {
    invalid("PDF vector dash authority is unknown.");
  }
  if (changes.lineCap !== undefined && current.lineCap === null) {
    invalid("PDF vector line-cap authority is unknown.");
  }
  if (changes.lineJoin !== undefined && current.lineJoin === null) {
    invalid("PDF vector line-join authority is unknown.");
  }
  if (changes.miterLimit !== undefined && current.miterLimit === null) {
    invalid("PDF vector miter authority is unknown.");
  }
  if (changes.opacity !== undefined && current.opacity === null) {
    invalid("PDF vector opacity authority is unknown.");
  }
  if (changes.blendMode !== undefined && current.blendMode === null) {
    invalid("PDF vector blend-mode authority is unknown.");
  }
  return Object.freeze({
    fillColor: changes.fillColor ?? current.fillColor,
    strokeColor: changes.strokeColor ?? current.strokeColor,
    lineWidth: changes.lineWidth ?? current.lineWidth,
    dashArray: changes.dashArray ?? current.dashArray,
    dashPhase: changes.dashPhase ?? current.dashPhase,
    lineCap: changes.lineCap ?? current.lineCap,
    lineJoin: changes.lineJoin ?? current.lineJoin,
    miterLimit: changes.miterLimit ?? current.miterLimit,
    opacity: changes.opacity ?? current.opacity,
    blendMode: changes.blendMode ?? current.blendMode,
    hasExternalGraphicsState:
      current.hasExternalGraphicsState ||
      changes.opacity !== undefined ||
      changes.blendMode !== undefined,
  });
}

function selectionFingerprint(
  selection: PdfNativeVectorObjectStyleSelection,
): string {
  return JSON.stringify(
    "kind" in selection
      ? {
          proofFamily: "shape",
          pageNumber: selection.pageNumber,
          pageProofIndex: selection.pageVectorIndex,
          nativeKind: selection.kind,
          paintMode: selection.paintMode,
          geometry: selection.geometry,
        }
      : {
          proofFamily: "subpath",
          pageNumber: selection.pageNumber,
          pageProofIndex: selection.pageSubpathIndex,
          paintMode: selection.paintMode,
          paintOperator: selection.paintOperator,
          commands: selection.commands,
          closed: selection.closed,
        },
    (_key, entry) =>
      typeof entry === "number" ? Number(entry.toFixed(6)) : entry,
  );
}

function documentSelectionFingerprints(
  authority: PdfNativeVectorDocumentAuthorityInspection,
  excluded: PdfNativeVectorObjectStyleProof,
): readonly string[] {
  return Object.freeze(
    authority.document.pages.flatMap((page) =>
      [...page.shapes, ...page.subpaths]
        .filter(
          (selection) =>
            !(
              selection.pageNumber === excluded.pageNumber &&
              ("kind" in selection ? "shape" : "subpath") ===
                excluded.proofFamily &&
              ("kind" in selection
                ? selection.pageVectorIndex
                : selection.pageSubpathIndex) === excluded.pageProofIndex
            ),
        )
        .map(selectionFingerprint),
    ),
  );
}

function lineCapOperator(value: PdfNativeVectorObjectStyleLineCap): number {
  return value === "butt" ? 0 : value === "round" ? 1 : 2;
}

function lineJoinOperator(value: PdfNativeVectorObjectStyleLineJoin): number {
  return value === "miter" ? 0 : value === "round" ? 1 : 2;
}

function concatBytes(parts: readonly Uint8Array[]): Uint8Array {
  const length = parts.reduce((total, part) => total + part.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

/**
 * Graphics-state prefix for a style edit. A paint-mode-only repaint writes no
 * state at all, so it gets no `q…Q` wrapper either: the object keeps its own
 * place in the stream (and stays reorderable) and only its paint operator
 * changes. Mirrored byte-for-byte by the server validator.
 */
function objectStyleOverridePrefix(
  changes: PdfNativeVectorObjectStyleChanges,
  expected: PdfNativeVectorObjectStyle,
  appearanceResourceName: string | null,
  spaces?: Readonly<{
    fill: ObjectStyleColorSpace;
    stroke: ObjectStyleColorSpace;
  }>,
): Uint8Array {
  const operators: string[] = [];
  if (changes.transform) {
    const { a, b, c, d, e, f } = changes.transform;
    operators.push(`${[a, b, c, d, e, f].map(formatNumber).join(" ")} cm`);
  }
  if (changes.translate) {
    operators.push(
      `1 0 0 1 ${formatNumber(changes.translate.x)} ${formatNumber(changes.translate.y)} cm`,
    );
  }
  if (changes.fillColor)
    operators.push(colorOperator(changes.fillColor, false, spaces?.fill));
  if (changes.strokeColor)
    operators.push(colorOperator(changes.strokeColor, true, spaces?.stroke));
  if (changes.lineWidth !== undefined) {
    operators.push(`${formatNumber(changes.lineWidth)} w`);
  }
  if (changes.dashArray !== undefined || changes.dashPhase !== undefined) {
    if (expected.dashArray === null || expected.dashPhase === null) {
      invalid("PDF vector dash authority is unknown.");
    }
    operators.push(
      `[${expected.dashArray.map(formatNumber).join(" ")}] ${formatNumber(expected.dashPhase)} d`,
    );
  }
  if (changes.lineCap !== undefined) {
    operators.push(`${lineCapOperator(changes.lineCap)} J`);
  }
  if (changes.lineJoin !== undefined) {
    operators.push(`${lineJoinOperator(changes.lineJoin)} j`);
  }
  if (changes.miterLimit !== undefined) {
    operators.push(`${formatNumber(changes.miterLimit)} M`);
  }
  if (changes.opacity !== undefined || changes.blendMode !== undefined) {
    if (!appearanceResourceName) {
      invalid("PDF vector appearance resource authority is missing.");
    }
    operators.push(`/${appearanceResourceName} gs`);
  }
  return operators.length === 0 ? new Uint8Array(0) : ascii(`q\n${operators.join("\n")}\n`);
}

async function appearanceResourceName(
  paintMode: PdfNativeVectorContentShape["paintMode"],
  changes: Readonly<{ opacity?: number; blendMode?: PdfStandardBlendMode }>,
): Promise<string> {
  const opacity = changes.opacity;
  if (opacity !== undefined && changes.blendMode === undefined) {
    const digest = await sha256PdfBytes(
      ascii(`mrai.vector-opacity.v1|${paintMode}|${formatNumber(opacity)}`),
    );
    return `MRAI_VO_${digest.slice(0, 24)}`;
  }
  const digest = await sha256PdfBytes(
    ascii(
      `mrai.vector-appearance.v1|${paintMode}|${opacity === undefined ? "-" : formatNumber(opacity)}|${changes.blendMode ?? "-"}`,
    ),
  );
  return `MRAI_VA_${digest.slice(0, 24)}`;
}

export async function planMraiNativeVectorObjectStyleEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly proofId: string;
  readonly changes: PdfNativeVectorObjectStyleChanges;
}): Promise<PdfNativeVectorObjectStyleEditPlan> {
  const changes = exactObjectStyleChanges(input.changes);
  const authority = await inspectMraiNativeVectorDocumentAuthority(
    input.sourceBytes,
  );
  const styles = await inspectMraiNativeVectorObjectStyleDocument(
    input.sourceBytes,
    authority,
  );
  const page = styles.pages[input.pageNumber - 1];
  const nativePage = authority.document.pages[input.pageNumber - 1];
  const authorityPage = authority.pages[input.pageNumber - 1];
  if (
    !page ||
    page.pageNumber !== input.pageNumber ||
    !nativePage ||
    !authorityPage?.contentStream
  ) {
    invalid("Selected PDF vector style page is unavailable.");
  }
  const matches = page.proofs.filter((proof) => proof.id === input.proofId);
  if (matches.length !== 1) {
    invalid("Selected PDF vector style proof is stale or unavailable.");
  }
  const proof = matches[0];
  const selection =
    proof.proofFamily === "shape"
      ? nativePage.shapes[proof.pageProofIndex]
      : nativePage.subpaths[proof.pageProofIndex];
  if (
    !selection ||
    selection.id !== proof.selectionId ||
    selection.proofDigest !== proof.selectionProofDigest
  ) {
    invalid("Selected PDF vector style selection is inconsistent.");
  }
  const {
    paintMode: requestedPaintMode,
    translate: requestedShift,
    transform: requestedTransform,
    ...styleChanges
  } = changes;
  const paintModeAfter = requestedPaintMode ?? proof.paintMode;
  // A resize as the object's own `cm`: the page transform T conjugated into
  // its user space (cm = M T M^-1, M the transform the path is drawn under),
  // rounded as the operator spells it; then the page transform those
  // numbers make, which the reopened object is held to.
  let localTransform: PdfNativeVectorStyleAffine | null = null;
  let pageTransform: PdfNativeVectorStyleAffine | null = null;
  if (requestedTransform) {
    const matrix = pdfNativeVectorSelectionMatrix(selection);
    const drawn = matrix ? styleAffine(matrix) : null;
    if (!drawn || Math.abs(drawn.a * drawn.d - drawn.b * drawn.c) < 1e-12) {
      invalid("PDF vector transform authority is unavailable.");
    }
    const inverse = invertStyleAffine(drawn);
    const exact = multiplyStyleAffine(multiplyStyleAffine(drawn, requestedTransform), inverse);
    const spell = (value: number) => Number(formatNumber(value));
    localTransform = Object.freeze({
      a: spell(exact.a),
      b: spell(exact.b),
      c: spell(exact.c),
      d: spell(exact.d),
      e: spell(exact.e),
      f: spell(exact.f),
    });
    if (Math.abs(localTransform.a * localTransform.d - localTransform.b * localTransform.c) <= 1e-12) {
      invalid("Invalid PDF vector transform.");
    }
    pageTransform = multiplyStyleAffine(multiplyStyleAffine(inverse, localTransform), drawn);
  }
  // The page-space shift as the line's own `cm` operands, rounded as the
  // operator spells them; the shift it then makes on the page.
  let localShift: Readonly<{ x: number; y: number }> | null = null;
  let pageShift: Readonly<{ x: number; y: number }> | null = null;
  if (requestedShift) {
    if (!("kind" in selection) || selection.kind !== "line") {
      invalid("Only a straight PDF line moves with its line width.");
    }
    const matrix = pdfNativeVectorSelectionMatrix(selection);
    const determinant = matrix ? matrix[0] * matrix[3] - matrix[1] * matrix[2] : 0;
    if (!matrix || Math.abs(determinant) < 1e-12) {
      invalid("PDF vector transform authority is unavailable.");
    }
    const spell = (value: number) => Number(formatNumber(value));
    localShift = Object.freeze({
      x: spell((matrix[3] * requestedShift.x - matrix[2] * requestedShift.y) / determinant),
      y: spell((-matrix[1] * requestedShift.x + matrix[0] * requestedShift.y) / determinant),
    });
    pageShift = Object.freeze({
      x: localShift.x * matrix[0] + localShift.y * matrix[2],
      y: localShift.x * matrix[1] + localShift.y * matrix[3],
    });
  }
  let baseStyle = proof.style;
  if (requestedPaintMode !== undefined) {
    if (requestedPaintMode === proof.paintMode) {
      invalid("PDF vector paint mode is unchanged.");
    }
    if ("kind" in selection && selection.kind === "line") {
      invalid("PDF vector lines have no fill area to repaint.");
    }
    if (
      !("kind" in selection) &&
      selection.commands.length === 2 &&
      !selection.closed
    ) {
      // An open two-point subpath repainted as a stroke would be reclassified
      // as a line shape, moving it across proof families.
      invalid("PDF vector open segment cannot be repainted.");
    }
    const state = proofStates.get(proof);
    if (!state) invalid("PDF vector object-style state authority is missing.");
    baseStyle = publicObjectStyle(state, paintModeAfter);
  }
  const expectedStyle = expectedObjectStyle(baseStyle, styleChanges, {
    repaint: requestedPaintMode !== undefined,
  });
  if (
    requestedPaintMode === undefined &&
    !requestedTransform &&
    JSON.stringify(
      canonicalObjectStyle({
        ...expectedStyle,
        hasExternalGraphicsState: proof.style.hasExternalGraphicsState,
      }),
    ) === JSON.stringify(canonicalObjectStyle(proof.style))
  ) {
    invalid("PDF vector object-style edit is unchanged.");
  }
  const stream = authorityPage.contentStream.streamBytes;
  if (
    proof.paintSpan.start < proof.pathSpan.end ||
    proof.paintSpan.end > stream.length
  ) {
    invalid("PDF vector object-style span is invalid.");
  }
  const original = stream.slice(proof.pathSpan.start, proof.paintSpan.end);
  const expectedSha256 = await sha256PdfBytes(original);
  const hasAppearanceChange =
    changes.opacity !== undefined || changes.blendMode !== undefined;
  const graphicsStateName = !hasAppearanceChange
    ? null
    : await appearanceResourceName(paintModeAfter, changes);
  const pageExtGStateMutation = !hasAppearanceChange
    ? null
    : await planMraiNativePageExtGStateMutation({
        sourceBytes: input.sourceBytes,
        pageNumber: input.pageNumber,
        resourceName: graphicsStateName!,
        fillOpacity:
          changes.opacity === undefined || paintModeAfter === "stroke"
            ? null
            : changes.opacity,
        strokeOpacity:
          changes.opacity === undefined || paintModeAfter === "fill"
            ? null
            : changes.opacity,
        blendMode: changes.blendMode ?? null,
      });
  const paintOperatorAfter =
    requestedPaintMode === undefined
      ? null
      : pdfNativeVectorPaintOperatorForMode(
          proof.paintOperator,
          proof.paintMode,
          requestedPaintMode,
        );
  const proofState = proofStates.get(proof);
  const overridePrefix = objectStyleOverridePrefix(
    localTransform
      ? { ...styleChanges, transform: localTransform }
      : localShift
        ? { ...styleChanges, translate: localShift }
        : styleChanges,
    expectedStyle,
    graphicsStateName,
    Object.freeze({
      fill: proofState?.fillColorSpace ?? "rgb",
      stroke: proofState?.strokeColorSpace ?? "rgb",
    }),
  );
  if (overridePrefix.length === 0 && paintOperatorAfter === null) {
    invalid("PDF vector object-style edit writes nothing.");
  }
  const replacement = concatBytes([
    overridePrefix,
    paintOperatorAfter === null
      ? original
      : stream.slice(proof.pathSpan.start, proof.pathSpan.end),
    ...(paintOperatorAfter === null ? [] : [ascii(`\n${paintOperatorAfter}`)]),
    ...(overridePrefix.length === 0 ? [] : [ascii("\nQ")]),
  ]);
  const mutableResourceIdentities = new Set(
    pageExtGStateMutation
      ? [
          `${pageExtGStateMutation.page.objectNumber}:${pageExtGStateMutation.page.generation}`,
          `${pageExtGStateMutation.resourceOwner.objectNumber}:${pageExtGStateMutation.resourceOwner.generation}`,
          ...(pageExtGStateMutation.category.kind === "indirect"
            ? [
                `${pageExtGStateMutation.category.target.objectNumber}:${pageExtGStateMutation.category.target.generation}`,
              ]
            : []),
        ]
      : [],
  );
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: authorityPage.contentStream.identity,
    expectedObjectSha256: authorityPage.contentStream.objectSha256,
    expectedStreamSha256: authorityPage.contentStream.streamSha256,
    operatorSpan: Object.freeze({
      start: proof.pathSpan.start,
      end: proof.paintSpan.end,
      expectedSha256,
    }),
    replacement,
    untouchedObjects: Object.freeze(
      authorityPage.contentStream.untouchedObjects.filter(
        (item) =>
          !mutableResourceIdentities.has(
            `${item.objectNumber}:${item.generation}`,
          ),
      ),
    ),
    ...(pageExtGStateMutation ? { pageExtGStateMutation } : {}),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
    proof,
    selection,
    beforeStyle: proof.style,
    expectedStyle,
    paintModeAfter,
    paintOperatorAfter,
    beforeSelectionFingerprint: selectionFingerprint(selection),
    beforeSiblingFingerprints: documentSelectionFingerprints(authority, proof),
    request,
    writtenChanges: localTransform
      ? Object.freeze({ ...changes, transform: localTransform })
      : localShift
        ? Object.freeze({ ...changes, translate: localShift })
        : changes,
    pageShift,
    pageTransform,
  });
}

export async function prepareMraiNativeVectorObjectStyleEdit(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorObjectStyleEditPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

/**
 * Project a reopened selection back onto the pre-edit paint identity so the
 * outline (geometry / commands) can be compared verbatim across a repaint.
 */
function repaintedSelectionForFingerprint(
  selection: PdfNativeVectorObjectStyleSelection,
  plan: PdfNativeVectorObjectStyleEditPlan,
): PdfNativeVectorObjectStyleSelection {
  if (plan.paintOperatorAfter === null) return selection;
  const before = plan.selection;
  if ("kind" in selection) {
    return Object.freeze({ ...selection, paintMode: before.paintMode });
  }
  if ("kind" in before) return selection;
  return Object.freeze({
    ...selection,
    paintMode: before.paintMode,
    paintOperator: before.paintOperator,
    closed: before.closed,
  });
}

function styleAffine(
  matrix: readonly [number, number, number, number, number, number],
): PdfNativeVectorStyleAffine {
  const [a, b, c, d, e, f] = matrix;
  return { a, b, c, d, e, f };
}

/** `left` then `right`, in PDF's row-vector convention (p' = p · left · right). */
function multiplyStyleAffine(
  left: PdfNativeVectorStyleAffine,
  right: PdfNativeVectorStyleAffine,
): PdfNativeVectorStyleAffine {
  return Object.freeze({
    a: left.a * right.a + left.b * right.c,
    b: left.a * right.b + left.b * right.d,
    c: left.c * right.a + left.d * right.c,
    d: left.c * right.b + left.d * right.d,
    e: left.e * right.a + left.f * right.c + right.e,
    f: left.e * right.b + left.f * right.d + right.f,
  });
}

function invertStyleAffine(
  affine: PdfNativeVectorStyleAffine,
): PdfNativeVectorStyleAffine {
  const determinant = affine.a * affine.d - affine.b * affine.c;
  return Object.freeze({
    a: affine.d / determinant,
    b: -affine.b / determinant,
    c: -affine.c / determinant,
    d: affine.a / determinant,
    e: (affine.c * affine.f - affine.d * affine.e) / determinant,
    f: (affine.b * affine.e - affine.a * affine.f) / determinant,
  });
}

/**
 * The reopened object is the planned one carried by `transform` on the page:
 * a line's ends, a rectangle's corners, an ellipse's centre and axes, a
 * subpath's every point -- in the order the path draws them.
 */
function selectionTransformedBy(
  before: PdfNativeVectorObjectStyleSelection,
  after: PdfNativeVectorObjectStyleSelection,
  transform: PdfNativeVectorStyleAffine,
): boolean {
  const map = (point: Readonly<{ x: number; y: number }>) => ({
    x: transform.a * point.x + transform.c * point.y + transform.e,
    y: transform.b * point.x + transform.d * point.y + transform.f,
  });
  const turn = (vector: Readonly<{ x: number; y: number }>) => ({
    x: transform.a * vector.x + transform.c * vector.y,
    y: transform.b * vector.x + transform.d * vector.y,
  });
  // A path's numbers go through the written `cm`'s six decimals: the
  // reopened point is the planned one to within that rounding's reach.
  const near = (
    point: Readonly<{ x: number; y: number }>,
    expected: Readonly<{ x: number; y: number }>,
  ) =>
    Math.abs(point.x - expected.x) <= 1e-4 &&
    Math.abs(point.y - expected.y) <= 1e-4;
  if (
    after.pageNumber !== before.pageNumber ||
    after.paintMode !== before.paintMode ||
    ("kind" in before) !== ("kind" in after)
  ) {
    return false;
  }
  if ("kind" in before && "kind" in after) {
    if (before.pageVectorIndex !== after.pageVectorIndex) return false;
    const from = before.geometry;
    const to = after.geometry;
    if (from.kind === "line" && to.kind === "line") {
      return near(to.start, map(from.start)) && near(to.end, map(from.end));
    }
    if (from.kind === "rectangle" && to.kind === "rectangle") {
      return (
        from.corners.length === to.corners.length &&
        from.corners.every((corner, index) => near(to.corners[index], map(corner)))
      );
    }
    if (from.kind === "ellipse" && to.kind === "ellipse") {
      return (
        near(to.center, map(from.center)) &&
        near(to.axisX, turn(from.axisX)) &&
        near(to.axisY, turn(from.axisY))
      );
    }
    return false;
  }
  if ("kind" in before || "kind" in after) return false;
  if (
    before.pageSubpathIndex !== after.pageSubpathIndex ||
    before.paintOperator !== after.paintOperator ||
    before.closed !== after.closed ||
    before.commands.length !== after.commands.length
  ) {
    return false;
  }
  return before.commands.every((command, index) => {
    const other = after.commands[index];
    if (command.operator !== other.operator) return false;
    if (command.operator === "h") return true;
    if ("point" in command) {
      return "point" in other && near(other.point, map(command.point));
    }
    return (
      "control1" in other &&
      near(other.control1, map(command.control1)) &&
      near(other.control2, map(command.control2)) &&
      near(other.end, map(command.end))
    );
  });
}

/** The reopened line is the planned one, shifted on the page by `shift`. */
function lineMovedBy(
  before: PdfNativeVectorObjectStyleSelection,
  after: PdfNativeVectorObjectStyleSelection,
  shift: Readonly<{ x: number; y: number }>,
): boolean {
  if (
    !("kind" in before) ||
    !("kind" in after) ||
    before.geometry.kind !== "line" ||
    after.geometry.kind !== "line" ||
    after.paintMode !== before.paintMode ||
    after.pageNumber !== before.pageNumber ||
    after.pageVectorIndex !== before.pageVectorIndex
  ) {
    return false;
  }
  const near = (
    point: Readonly<{ x: number; y: number }>,
    from: Readonly<{ x: number; y: number }>,
  ) =>
    Math.abs(point.x - (from.x + shift.x)) <= 1e-6 &&
    Math.abs(point.y - (from.y + shift.y)) <= 1e-6;
  return (
    near(after.geometry.start, before.geometry.start) &&
    near(after.geometry.end, before.geometry.end)
  );
}

export async function verifyMraiNativeVectorObjectStyleCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorObjectStyleEditPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorObjectStyleVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const authority = await inspectMraiNativeVectorDocumentAuthority(
    prepared.candidateBytes,
  );
  const styles = await inspectMraiNativeVectorObjectStyleDocument(
    prepared.candidateBytes,
    authority,
  );
  const page = styles.pages[plan.proof.pageNumber - 1];
  const nativePage = authority.document.pages[plan.proof.pageNumber - 1];
  const reopened = page?.proofs.find(
    (proof) =>
      proof.proofFamily === plan.proof.proofFamily &&
      proof.pageProofIndex === plan.proof.pageProofIndex,
  );
  const selection = reopened
    ? reopened.proofFamily === "shape"
      ? nativePage?.shapes[reopened.pageProofIndex]
      : nativePage?.subpaths[reopened.pageProofIndex]
    : null;
  // A resized shape answers to the clip it is drawn under, as a moved one
  // does: its reopened outline must stay inside it (the replay holds the
  // record to the same rule).
  if (
    reopened &&
    selection &&
    plan.pageTransform &&
    !pdfNativeVectorSelectionClipAllows(selection, selection.bounds)
  ) {
    prepared.candidateBytes.fill(0);
    refuse(MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE);
  }
  if (
    !reopened ||
    !selection ||
    (plan.pageTransform
      ? !selectionTransformedBy(plan.selection, selection, plan.pageTransform)
      : plan.pageShift
        ? !lineMovedBy(plan.selection, selection, plan.pageShift)
        : selectionFingerprint(
            repaintedSelectionForFingerprint(selection, plan),
          ) !== plan.beforeSelectionFingerprint) ||
    reopened.paintMode !== plan.paintModeAfter ||
    (plan.paintOperatorAfter !== null &&
      reopened.paintOperator !== plan.paintOperatorAfter) ||
    JSON.stringify(canonicalObjectStyle(reopened.style)) !==
      JSON.stringify(canonicalObjectStyle(plan.expectedStyle)) ||
    JSON.stringify(documentSelectionFingerprints(authority, reopened)) !==
      JSON.stringify(plan.beforeSiblingFingerprints)
  ) {
    prepared.candidateBytes.fill(0);
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector object style does not match the requested edit.",
    );
  }
  return Object.freeze({
    verified: true,
    candidateSha256: prepared.candidateSha256,
    reopenedProof: reopened,
    reopenedSelection: selection,
  });
}

async function inspectInternal(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<
  Readonly<{
    sourceSha256: string;
    pageNumber: number;
    proofs: readonly InternalStyleProof[];
    unsupportedSelectionIds: readonly string[];
  }>
> {
  const vectors = await inspectMraiNativeVectorContent(sourceBytes, pageNumber);
  const page = await inspectMraiNativePageContentGraph(sourceBytes, pageNumber);
  const streamCache = new Map<string, PdfNativeContentStreamInspection>();
  const proofs: InternalStyleProof[] = [];
  const unsupportedSelectionIds: string[] = [];
  // A form other pages reference too is edited through a clone of it; a
  // style edit does not write one, so such a form's vectors keep their style.
  const isolatedForms = new Set(
    (await inspectMraiNativeVectorFormHosts(sourceBytes, pageNumber))
      .filter((host) => host.formInstanceIsolation)
      .map((host) => `${host.form.objectNumber}:${host.form.generation}`),
  );
  for (const selection of vectors.shapes) {
    if (
      !selection.editable ||
      isolatedForms.has(`${selection.contentStream.objectNumber}:${selection.contentStream.generation}`)
    ) {
      unsupportedSelectionIds.push(selection.id);
      continue;
    }
    const key = `${selection.contentStream.objectNumber}:${selection.contentStream.generation}`;
    let streamInspection = streamCache.get(key);
    if (!streamInspection) {
      streamInspection = await inspectMraiNativeContentStream(
        sourceBytes,
        selection.contentStream,
        [page.page],
      );
      streamCache.set(key, streamInspection);
    }
    try {
      const parsed = styleEnvelope(streamInspection.streamBytes, selection);
      const visible = await visibleProof(
        selection,
        parsed.envelope,
        streamInspection,
      );
      proofs.push(
        Object.freeze({
          visible,
          selection,
          streamInspection,
          operations: parsed.operations,
          envelope: parsed.envelope,
        }),
      );
    } catch (error) {
      if (
        error instanceof PdfNativeTransactionError &&
        error.code === "unsupported-document"
      ) {
        unsupportedSelectionIds.push(selection.id);
        continue;
      }
      throw error;
    }
  }
  return Object.freeze({
    sourceSha256: vectors.sourceSha256,
    pageNumber,
    proofs: Object.freeze(proofs),
    unsupportedSelectionIds: Object.freeze(unsupportedSelectionIds),
  });
}

export async function inspectMraiNativeVectorStyles(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativeVectorStyleInspection> {
  const inspection = await inspectInternal(sourceBytes, pageNumber);
  return Object.freeze({
    sourceSha256: inspection.sourceSha256,
    pageNumber,
    proofs: Object.freeze(inspection.proofs.map((proof) => proof.visible)),
    unsupportedSelectionIds: inspection.unsupportedSelectionIds,
  });
}

function formatNumber(value: number): string {
  const normalized = rounded(value);
  return normalized
    .toFixed(6)
    .replace(/\.(\d*?)0+$/, (_, kept: string) => (kept ? `.${kept}` : "")) // no lookbehind: Safari < 16.4
    .replace(/\.$/, "");
}

/**
 * Write a colour in the space the object already used. DeviceGray keeps `g`
 * when the requested colour is achromatic, DeviceCMYK converts exactly back
 * through the same naive transform the reader applies, and everything else
 * stays DeviceRGB. Mirrored byte-for-byte by the server writer.
 */
function colorOperator(
  color: PdfNativeVectorRgbColor,
  stroke: boolean,
  space: ObjectStyleColorSpace = "rgb",
): string {
  const r = rounded(color.r);
  const g = rounded(color.g);
  const b = rounded(color.b);
  if (space === "gray" && r === g && g === b) {
    return `${formatNumber(r)} ${stroke ? "G" : "g"}`;
  }
  if (space === "cmyk") {
    const k = 1 - Math.max(r, g, b);
    const maximum = 1 - k;
    return `${formatNumber(maximum - r)} ${formatNumber(maximum - g)} ${formatNumber(maximum - b)} ${formatNumber(k)} ${stroke ? "K" : "k"}`;
  }
  return `${formatNumber(r)} ${formatNumber(g)} ${formatNumber(b)} ${stroke ? "RG" : "rg"}`;
}

function patchedEnvelope(
  streamBytes: Uint8Array,
  envelope: StyleEnvelope,
  changes: PdfNativeVectorStyleChanges,
): Uint8Array {
  const source = latin1(streamBytes);
  const replacements: Array<{
    start: number;
    end: number;
    value: string;
  }> = [];
  if (changes.fillColor) {
    if (!envelope.fillOperation)
      invalid("Selected vector has no direct fill style.");
    replacements.push({
      start: envelope.fillOperation.start,
      end: envelope.fillOperation.end,
      value: colorOperator(exactColor(changes.fillColor, "fill"), false),
    });
  }
  if (changes.strokeColor) {
    if (!envelope.strokeOperation)
      invalid("Selected vector has no direct stroke style.");
    replacements.push({
      start: envelope.strokeOperation.start,
      end: envelope.strokeOperation.end,
      value: colorOperator(exactColor(changes.strokeColor, "stroke"), true),
    });
  }
  if (changes.lineWidth !== undefined) {
    if (!envelope.lineWidthOperation)
      invalid("Selected vector has no direct line width.");
    if (
      !Number.isFinite(changes.lineWidth) ||
      changes.lineWidth <= 0 ||
      changes.lineWidth > MAX_LINE_WIDTH
    ) {
      invalid("Invalid PDF vector line width.");
    }
    replacements.push({
      start: envelope.lineWidthOperation.start,
      end: envelope.lineWidthOperation.end,
      value: `${formatNumber(changes.lineWidth)} w`,
    });
  }
  if (replacements.length === 0) invalid("PDF vector style edit is empty.");
  replacements.sort((left, right) => left.start - right.start);
  let cursor = envelope.start;
  let replacement = "";
  for (const item of replacements) {
    if (item.start < cursor || item.end > envelope.end) {
      invalid("PDF vector style replacement spans overlap.");
    }
    replacement += source.slice(cursor, item.start);
    replacement += item.value;
    cursor = item.end;
  }
  replacement += source.slice(cursor, envelope.end);
  return ascii(replacement);
}

function expectedStyle(
  current: PdfNativeVectorStyle,
  changes: PdfNativeVectorStyleChanges,
): PdfNativeVectorStyle {
  return Object.freeze({
    fillColor: changes.fillColor
      ? exactColor(changes.fillColor, "fill")
      : current.fillColor,
    strokeColor: changes.strokeColor
      ? exactColor(changes.strokeColor, "stroke")
      : current.strokeColor,
    lineWidth:
      changes.lineWidth !== undefined ? changes.lineWidth : current.lineWidth,
  });
}

export async function planMraiNativeVectorStyleEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly proofId: string;
  readonly changes: PdfNativeVectorStyleChanges;
}): Promise<PdfNativeVectorStyleEditPlan> {
  const inspection = await inspectInternal(input.sourceBytes, input.pageNumber);
  const proof = inspection.proofs.find(
    (item) => item.visible.id === input.proofId,
  );
  if (!proof)
    invalid("Selected PDF vector style proof is stale or unavailable.");
  const replacement = patchedEnvelope(
    proof.streamInspection.streamBytes,
    proof.envelope,
    input.changes,
  );
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: proof.streamInspection.identity,
    expectedObjectSha256: proof.streamInspection.objectSha256,
    expectedStreamSha256: proof.streamInspection.streamSha256,
    operatorSpan: proof.visible.envelopeSpan,
    replacement,
    untouchedObjects: proof.streamInspection.untouchedObjects,
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_STYLE_SCHEMA_VERSION,
    proof: proof.visible,
    beforeGeometry: proof.selection.geometry,
    expectedStyle: expectedStyle(proof.visible.style, input.changes),
    request,
  });
}

export async function prepareMraiNativeVectorStyleEdit(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorStyleEditPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeVectorStyleCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorStyleEditPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorStyleVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const inspection = await inspectInternal(
    prepared.candidateBytes,
    plan.proof.pageNumber,
  );
  const reopened = inspection.proofs.find(
    (item) => item.visible.pageVectorIndex === plan.proof.pageVectorIndex,
  );
  if (
    !reopened ||
    geometryFingerprint(reopened.selection.geometry) !==
      geometryFingerprint(plan.beforeGeometry) ||
    JSON.stringify(canonicalStyle(reopened.visible.style)) !==
      JSON.stringify(canonicalStyle(plan.expectedStyle))
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector style does not match the requested edit.",
    );
  }
  return Object.freeze({
    verified: true,
    candidateSha256: prepared.candidateSha256,
    reopenedProof: reopened.visible,
  });
}
