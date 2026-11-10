/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK, MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT, MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED, pdfNativeEmbeddedImageIsAlreadyArranged } from "./embedded-image-arrange-model";
export { MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK, MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT, MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED, pdfNativeEmbeddedImageIsAlreadyArranged };
import { PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN, pdfNativeVectorProofId } from "./vector-identity";
import { decodePdfSourceName, pdfSourceNamesRead } from "./pdf-source-name";
export { PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN, pdfNativeVectorProofId };
import {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
} from "./vocabulary";
export {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
};
import {
  countMraiNativeIndirectObjectReferences,
  countMraiNativeIndirectObjectReferencesBatch,
  expandPdfNativeReplacementPieces,
  inspectMraiNativeDocumentPageContentGraphs,
  inspectMraiNativePageContentGraph,
  inspectMraiNativePageContentStreams,
  inspectMraiNativeVectorFormHosts,
  type PdfNativeVectorFormHost,
  joinNativePageDrawing,
  mraiNativeInlineImageDataEnd,
  PdfNativeTransactionError,
  prepareMraiNativeContentStreamTransaction,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  type PdfNativeContentStreamInspection,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeFormInstanceIsolationPlan,
  type PdfNativeReplacementPiece,
  type PdfNativePageContentGraphInspection,
  type PdfNativePageContentStreamInspection,
  type PdfNativePreparedContentStreamTransaction,
  inspectMraiNativeEmbeddedImagePagePaints,
  inspectMraiNativeConsolidatedContentStream,
  inspectMraiNativeObject,
  planMraiNativeContentStreamSpanRevision,
  parseMraiNativeExtGStateDictionaryEntries,
  parseMraiNativeResourceCategoryEntries,
  nextMraiNativeIndirectObjectNumber,
  type PdfNativePageContentConsolidationPlan,
  type PdfNativeEmbeddedImagePagePaints,
  type PdfNativeEmbeddedImagePageStream,
  type PdfNativeEmbeddedImagePaintBlock,
} from "./content-stream-transaction";

export const MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION =
  "mrai.native-vector-geometry.v1" as const;
export const MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION =
  "mrai.native-vector-subpath.v1" as const;

export type PdfNativeVectorKind = "line" | "rectangle" | "ellipse";
export type PdfNativeVectorPaintMode = "stroke" | "fill" | "fill-stroke";

export interface PdfNativeVectorPoint {
  readonly x: number;
  readonly y: number;
}

export type PdfNativeVectorGeometry = Readonly<
  | {
      kind: "line";
      start: PdfNativeVectorPoint;
      end: PdfNativeVectorPoint;
    }
  | {
      kind: "rectangle";
      corners: readonly [
        PdfNativeVectorPoint,
        PdfNativeVectorPoint,
        PdfNativeVectorPoint,
        PdfNativeVectorPoint,
      ];
    }
  | {
      kind: "ellipse";
      center: PdfNativeVectorPoint;
      axisX: PdfNativeVectorPoint;
      axisY: PdfNativeVectorPoint;
    }
>;

export interface PdfNativeVectorBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type PdfNativeVectorPreserveReason =
  "shared-content-stream" | "unsupported-content-stream";

export interface PdfNativeVectorContentShape {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION;
  readonly id: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageVectorIndex: number;
  readonly kind: PdfNativeVectorKind;
  readonly paintMode: PdfNativeVectorPaintMode;
  readonly geometry: PdfNativeVectorGeometry;
  readonly bounds: PdfNativeVectorBounds;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly proofDigest: string;
  readonly editable: boolean;
  readonly preserveReason: PdfNativeVectorPreserveReason | null;
  /** Which drawing of a Form the page shows more than once this proof reads
   * (the `Do` offsets on its way): part of its id. */
  readonly formInstance?: string;
}



export interface PdfNativeVectorContentInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly shapes: readonly PdfNativeVectorContentShape[];
  readonly losses: readonly Readonly<{
    contentStream: Readonly<{ objectNumber: number; generation: number }>;
    reason: "unsupported-content-stream";
  }>[];
}

export interface PdfNativeVectorGeometryEditPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorContentShape;
  readonly replacement: PdfNativeVectorGeometry;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly beforeShapeFingerprints: readonly string[];
}

export interface PdfNativeVectorGeometryVerification {
  readonly verified: true;
  readonly pageNumber: number;
  readonly pageVectorIndex: number;
  readonly candidateSha256: string;
  readonly reopenedShape: PdfNativeVectorContentShape;
}

export type PdfNativeVectorSubpathOperator = "m" | "l" | "c" | "v" | "y" | "h";
export type PdfNativeVectorPaintOperator =
  "S" | "s" | "f" | "F" | "f*" | "B" | "B*" | "b" | "b*";

export type PdfNativeVectorSubpathCommand = Readonly<
  | {
      operator: "m" | "l";
      point: PdfNativeVectorPoint;
    }
  | {
      operator: "c" | "v" | "y";
      control1: PdfNativeVectorPoint;
      control2: PdfNativeVectorPoint;
      end: PdfNativeVectorPoint;
    }
  | { operator: "h" }
>;

export interface PdfNativeVectorSubpathProof {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION;
  readonly id: string;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly pageSubpathIndex: number;
  readonly paintMode: PdfNativeVectorPaintMode;
  readonly paintOperator: PdfNativeVectorPaintOperator;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
  readonly closed: boolean;
  readonly bounds: PdfNativeVectorBounds;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly proofDigest: string;
  readonly editable: boolean;
  readonly preserveReason: PdfNativeVectorPreserveReason | null;
  /** See PdfNativeVectorContentShape.formInstance. */
  readonly formInstance?: string;
}

export interface PdfNativeVectorSubpathInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly subpaths: readonly PdfNativeVectorSubpathProof[];
  readonly excludedExistingShapeCount: number;
  readonly losses: PdfNativeVectorContentInspection["losses"];
}

export interface PdfNativeVectorDocumentPageInspection {
  readonly pageNumber: number;
  readonly shapes: readonly PdfNativeVectorContentShape[];
  readonly subpaths: readonly PdfNativeVectorSubpathProof[];
  readonly excludedExistingShapeCount: number;
  readonly shapeLosses: PdfNativeVectorContentInspection["losses"];
  readonly subpathLosses: PdfNativeVectorSubpathInspection["losses"];
}

export interface PdfNativeVectorDocumentInspection {
  readonly sourceSha256: string;
  readonly pageCount: number;
  readonly pages: readonly PdfNativeVectorDocumentPageInspection[];
}

/** Worker-local source authority. Raw stream bytes must never cross into UI. */
export interface PdfNativeVectorDocumentPageAuthority {
  readonly pageNumber: number;
  readonly contentStream: PdfNativeContentStreamInspection | null;
  readonly pageGraph: PdfNativePageContentGraphInspection;
}

export interface PdfNativeVectorDocumentAuthorityInspection {
  readonly document: PdfNativeVectorDocumentInspection;
  readonly pages: readonly PdfNativeVectorDocumentPageAuthority[];
}

export interface PdfNativeVectorAffineTransform {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export interface PdfNativeVectorSubpathAffineEditPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorSubpathProof;
  readonly transform: PdfNativeVectorAffineTransform;
  readonly transformedCommands: readonly PdfNativeVectorSubpathCommand[];
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly beforeSubpathFingerprints: readonly string[];
  readonly beforeShapeFingerprints: readonly string[];
}

export interface PdfNativeVectorSubpathVerification {
  readonly verified: true;
  readonly pageNumber: number;
  readonly pageSubpathIndex: number;
  readonly candidateSha256: string;
  readonly reopenedSubpath: PdfNativeVectorSubpathProof;
}

export type PdfNativeVectorDeleteSelection =
  | PdfNativeVectorContentShape
  | PdfNativeVectorSubpathProof;

export interface PdfNativeVectorDeletePlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorDeleteSelection;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly originalOperatorBytes: Uint8Array;
  readonly beforeShapeFingerprints: readonly string[];
  readonly beforeSubpathFingerprints: readonly string[];
}

export interface PdfNativeVectorDeleteVerification {
  readonly verified: true;
  readonly pageNumber: number;
  readonly deletedProofFamily: "shape" | "subpath";
  readonly deletedProofIndex: number;
  readonly candidateSha256: string;
}

export type PdfNativeVectorDuplicateDirection = "insert" | "remove";

/**
 * Duplicate a path object in place: the copy is emitted right after its
 * source (same paint operator, same graphics state) translated by `offset`
 * in PDF user space. `remove` deletes exactly the inserted bytes again.
 */
export interface PdfNativeVectorDuplicatePlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION;
  readonly direction: PdfNativeVectorDuplicateDirection;
  /** Insert: the source object. Remove: the copy being removed. */
  readonly selection: PdfNativeVectorDeleteSelection;
  readonly offset: Readonly<{ dx: number; dy: number }>;
  readonly paintOperator: string;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly originalOperatorBytes: Uint8Array;
  /** `\n<copy path>\n<paint>` exactly as appended (insert) or removed. */
  readonly insertedBytes: Uint8Array;
  readonly expectedCopyFingerprint: string;
  readonly beforeShapeFingerprints: readonly string[];
  readonly beforeSubpathFingerprints: readonly string[];
}

export interface PdfNativeVectorDuplicateVerification {
  readonly verified: true;
  readonly pageNumber: number;
  readonly proofFamily: "shape" | "subpath";
  /** Index of the copy in its family after an insert; the removed index otherwise. */
  readonly copyProofIndex: number;
  readonly candidateSha256: string;
  readonly reopenedCopy: PdfNativeVectorDeleteSelection | null;
}

/** One step past the next object (Bring Forward / Send Backward) or all the
 * way to that end of the drawing in one write (Bring to Front / Send to
 * Back). */
export type PdfNativeVectorReorderDirection = "forward" | "backward" | "front" | "back";

/**
 * Swap a path object with its stream-adjacent neighbour so it paints later
 * (`forward`, on top) or earlier (`backward`, underneath). The pair
 * `[A][gap][B]` becomes `q\n[gap]\n[B]\nQ\n[A]\n[gap]`: B keeps the
 * graphics state it was painted with, A keeps its own, and content after
 * the pair still sees the gap's state. `restore` swaps the bytes back.
 */
export interface PdfNativeVectorReorderPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION;
  readonly direction: PdfNativeVectorReorderDirection | "restore";
  readonly selection: PdfNativeVectorDeleteSelection;
  /** The path object swapped with, or null when the object passed other
   * content (`passed`). */
  readonly neighbour: PdfNativeVectorDeleteSelection | null;
  /** The non-path content the object was moved past, in stream offsets. */
  readonly passed: PdfNativeVectorPassedUnit | null;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly originalOperatorBytes: Uint8Array;
  readonly reorderedBytes: Uint8Array;
  readonly beforeShapeFingerprints: readonly string[];
  readonly beforeSubpathFingerprints: readonly string[];
  readonly expectedShapeFingerprints: readonly string[];
  readonly expectedSubpathFingerprints: readonly string[];
  /** Where the object stands in its own family once reopened. */
  readonly expectedSelectionIndex: number;
}

/**
 * Content a path object can be moved past in one step when it is not another
 * path object: one text object (`BT`…`ET`), one balanced `q`…`Q` group, one
 * XObject (`Do`) or one shading (`sh`). The object is then written inside its
 * own `q`…`Q` naming every graphics state the content it passes changes, so
 * it paints as it did and the content keeps the state it leaves.
 */
export type PdfNativeVectorPassedUnitKind =
  | "path"
  | "text"
  | "group"
  | "xobject"
  | "shading"
  /** A layer move: everything between the object's place and where it
   * lands, which may be any content (see layerMove). */
  | "content";

export interface PdfNativeVectorPassedUnit {
  readonly kind: PdfNativeVectorPassedUnitKind;
  readonly start: number;
  readonly end: number;
}

export interface PdfNativeVectorReorderVerification {
  readonly verified: true;
  readonly pageNumber: number;
  readonly candidateSha256: string;
  readonly reopenedSelection: PdfNativeVectorDeleteSelection;
}

type Matrix = readonly [number, number, number, number, number, number];

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
  readonly operatorStart: number;
  readonly end: number;
}

interface CubicSegment {
  readonly start: PdfNativeVectorPoint;
  readonly control1: PdfNativeVectorPoint;
  readonly control2: PdfNativeVectorPoint;
  readonly end: PdfNativeVectorPoint;
}

interface ParsedPath {
  readonly start: number;
  readonly end: number;
  readonly matrix: Matrix;
  readonly moves: readonly PdfNativeVectorPoint[];
  readonly lines: readonly Readonly<{
    start: PdfNativeVectorPoint;
    end: PdfNativeVectorPoint;
  }>[];
  readonly cubics: readonly CubicSegment[];
  readonly rectangles: readonly (readonly [
    PdfNativeVectorPoint,
    PdfNativeVectorPoint,
    PdfNativeVectorPoint,
    PdfNativeVectorPoint,
  ])[];
  readonly closed: boolean;
  readonly invalid: boolean;
  /** A move came after the path's segments. Nothing may follow it but more
   * moves: a subpath of one point paints nothing, while one that goes on
   * draws another shape than the corners read. */
  readonly reopened: boolean;
}

interface ScannedShape {
  readonly kind: PdfNativeVectorKind;
  readonly paintMode: PdfNativeVectorPaintMode;
  readonly paintOperator: string;
  readonly geometry: PdfNativeVectorGeometry;
  readonly bounds: PdfNativeVectorBounds;
  readonly matrix: Matrix;
  readonly span: Readonly<{ start: number; end: number }>;
  /** Byte span of the paint operator that closes this path object. */
  readonly paintSpan: Readonly<{ start: number; end: number }>;
  /** The clips the shape lies inside, or null when unclipped. */
  readonly clip: readonly ClipRegion[] | null;
}

interface InternalShape extends PdfNativeVectorContentShape {
  readonly streamInspection: PdfNativeContentStreamInspection;
  readonly matrix: Matrix;
  readonly paintSpan: Readonly<{ start: number; end: number }>;
  readonly paintOperatorToken: string;
  readonly clip: readonly ClipRegion[] | null;
}

interface ScannedSubpath {
  readonly paintMode: PdfNativeVectorPaintMode;
  readonly paintOperator: PdfNativeVectorPaintOperator;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
  readonly closed: boolean;
  readonly bounds: PdfNativeVectorBounds;
  readonly matrix: Matrix;
  readonly span: Readonly<{ start: number; end: number }>;
  readonly paintSpan: Readonly<{ start: number; end: number }>;
  readonly clip: readonly ClipRegion[] | null;
}

interface InternalSubpath extends PdfNativeVectorSubpathProof {
  readonly streamInspection: PdfNativeContentStreamInspection;
  readonly matrix: Matrix;
  readonly paintSpan: Readonly<{ start: number; end: number }>;
  readonly clip: readonly ClipRegion[] | null;
}

interface InternalCombinedPageInspection {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly shapes: readonly InternalShape[];
  readonly subpaths: readonly InternalSubpath[];
  readonly excludedExistingShapeCount: number;
  readonly shapeLosses: PdfNativeVectorContentInspection["losses"];
  readonly subpathLosses: PdfNativeVectorSubpathInspection["losses"];
  readonly subpathError: unknown | null;
}

interface PendingSubpath {
  readonly start: number;
  readonly end: number;
  readonly matrix: Matrix;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
  readonly currentPoint: PdfNativeVectorPoint | null;
  readonly startPoint: PdfNativeVectorPoint | null;
  /** The subpath the pen is on was closed by `h`; only `m` may follow. */
  readonly explicitlyClosed: boolean;
  /** Segments (l, c, v, y) drawn in the subpath the pen is on. */
  readonly segments: number;
  readonly invalid: boolean;
}

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
const NUMBER_PATTERN = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const KAPPA = 0.5522847498307936;
/** A page's whole drawing, read as one scan: a cover illustration runs to
 * hundreds of thousands of operands (an annual report's front page is two
 * megabytes of paths), and the scan is linear. */
const MAX_CONTENT_TOKENS = 4_000_000;
/** A wiring diagram or an exploded parts view draws fifty thousand paths
 * on one page (an air conditioner manual's schematic sheet: 52,000). Every
 * one is proved: a page past the limit loses its subpath proofs whole, and
 * the paths nobody can prove then leave their neighbours ambiguous too. The
 * scan is linear and the viewer runs it in its inspection worker (mirror of
 * the inventory's and the reconciliation's per-page limits). */
const MAX_VECTOR_SHAPES_PER_PAGE = 100_000;
// Outlined text writes a line of glyphs, and traced artwork a whole shape, as
// one path of a thousand commands and more (mirror of the server contract).
const MAX_NATIVE_SUBPATH_COMMANDS = 4096;
const MAX_ABS_COORDINATE = 10_000_000;
const EPSILON = 0.000_01;

/** A form's stream as a vector edit writes it: when other pages reference
 * the form too, the clone-on-write chain the transaction must carry. */
type VectorStreamInspection = PdfNativeContentStreamInspection &
  Readonly<{ formInstanceIsolation?: PdfNativeFormInstanceIsolationPlan }>;

/** The clone-on-write a request on this stream must carry, if any. */
function streamRequestIsolation(
  inspection: PdfNativeContentStreamInspection,
): Readonly<{ formInstanceIsolation?: PdfNativeFormInstanceIsolationPlan }> {
  const isolation = (inspection as VectorStreamInspection).formInstanceIsolation;
  return isolation ? { formInstanceIsolation: isolation } : {};
}

function refuse(message: string): never {
  throw new PdfNativeTransactionError("unsupported-document", message);
}

function invalid(message: string): never {
  throw new PdfNativeTransactionError("invalid-request", message);
}

function latin1(bytes: Uint8Array): string {
  let value = "";
  const chunk = 32_768;
  for (let offset = 0; offset < bytes.length; offset += chunk) {
    value += String.fromCharCode(...bytes.subarray(offset, offset + chunk));
  }
  return value;
}

function ascii(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

/** The bytes a latin1 string was read from. */
function latin1Bytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0) & 0xff);
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

function literalStringEnd(source: string, start: number): number {
  let cursor = start + 1;
  let depth = 1;
  while (cursor < source.length) {
    const character = source[cursor];
    if (character === "\\") {
      cursor += 2;
      continue;
    }
    if (character === "(") depth += 1;
    else if (character === ")") {
      depth -= 1;
      if (depth === 0) return cursor + 1;
    }
    cursor += 1;
  }
  refuse("PDF vector content literal string is unterminated.");
}

function hexStringEnd(source: string, start: number): number {
  const end = source.indexOf(">", start + 1);
  if (end < 0) refuse("PDF vector content hex string is unterminated.");
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
      cursor = literalStringEnd(source, cursor);
      continue;
    }
    if (source[cursor] === "<" && source[cursor + 1] !== "<") {
      cursor = hexStringEnd(source, cursor);
      continue;
    }
    const pair = source.slice(cursor, cursor + 2);
    if (pair === "<<") {
      stack.push("<<");
      cursor += 2;
      continue;
    }
    if (pair === ">>") {
      if (stack.pop() !== "<<") {
        refuse("PDF vector content dictionary is unbalanced.");
      }
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
      if (stack.pop() !== "[") {
        refuse("PDF vector content array is unbalanced.");
      }
      cursor += 1;
      if (stack.length === 0) return cursor;
      continue;
    }
    cursor += 1;
  }
  refuse("PDF vector content container is unterminated.");
}

function nameEnd(source: string, start: number): number {
  let cursor = start + 1;
  while (
    cursor < source.length &&
    !isWhitespace(source[cursor]) &&
    !isDelimiter(source[cursor])
  ) {
    cursor += 1;
  }
  if (cursor === start + 1) refuse("PDF vector content name is empty.");
  return cursor;
}

function wordEnd(source: string, start: number): number {
  let cursor = start;
  while (
    cursor < source.length &&
    !isWhitespace(source[cursor]) &&
    !isDelimiter(source[cursor])
  ) {
    cursor += 1;
  }
  if (cursor === start) refuse("PDF vector content token is malformed.");
  return cursor;
}

/** ``inlineImages``: read each inline image as one opaque painting -- its
 * dictionary, `ID`, its data and `EI` -- as a single `BI` operator, its data
 * ending where its own dictionary says. Only a route that moves paints as
 * wholes asks for it; one that walks graphics state still refuses them.
 * Mirrors replay's `_vector_style_operations(..., inline_images=True)`. */
function tokenize(source: string, inlineImages = false): readonly ContentToken[] {
  const tokens: ContentToken[] = [];
  let cursor = 0;
  while (cursor < source.length) {
    cursor = skipWhitespaceAndComments(source, cursor);
    if (cursor >= source.length) break;
    if (tokens.length >= MAX_CONTENT_TOKENS) {
      refuse("PDF vector content has too many tokens.");
    }
    const start = cursor;
    const character = source[cursor];
    if (character === "(") {
      cursor = literalStringEnd(source, cursor);
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
      cursor = hexStringEnd(source, cursor);
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
    if (isDelimiter(character)) {
      refuse("PDF vector content delimiter is unsupported.");
    }
    cursor = wordEnd(source, cursor);
    const value = source.slice(start, cursor);
    if (inlineImages && value === "ID") {
      let begin = tokens.length - 1;
      while (begin >= 0 && !(tokens[begin].kind === "operator" && tokens[begin].value === "BI")) {
        begin -= 1;
      }
      if (begin < 0 || tokens.slice(begin + 1).some((token) => token.kind === "operator")) {
        refuse("Inline PDF image dictionary is malformed.");
      }
      const opening = tokens[begin];
      const dataEnd = mraiNativeInlineImageDataEnd(
        source,
        cursor + 1,
        source.slice(opening.end, start),
      );
      if (dataEnd === null) refuse("Inline PDF image data is unterminated.");
      tokens.splice(begin, tokens.length - begin, {
        kind: "operator",
        value: "BI",
        start: opening.start,
        end: dataEnd,
      });
      cursor = dataEnd;
      continue;
    }
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

function operations(source: string, inlineImages = false): readonly ContentOperation[] {
  const result: ContentOperation[] = [];
  let operands: ContentToken[] = [];
  for (const token of tokenize(source, inlineImages)) {
    if (token.kind !== "operator") {
      operands.push(token);
      continue;
    }
    if (
      (token.value === "BI" && (!inlineImages || operands.length !== 0 || token.end - token.start === 2)) ||
      token.value === "ID" ||
      token.value === "EI"
    ) {
      refuse("Inline PDF images make vector operator ownership ambiguous.");
    }
    result.push(
      Object.freeze({
        operator: token.value,
        operands: Object.freeze(operands),
        start: operands[0]?.start ?? token.start,
        operatorStart: token.start,
        end: token.end,
      }),
    );
    operands = [];
  }
  if (operands.length !== 0) {
    refuse("PDF vector content ends inside an operation.");
  }
  return Object.freeze(result);
}

function numberOperands(
  operation: ContentOperation,
  count: number,
): readonly number[] {
  if (
    operation.operands.length !== count ||
    operation.operands.some((operand) => operand.kind !== "number")
  ) {
    refuse(`PDF ${operation.operator} vector operands are invalid.`);
  }
  const values = operation.operands.map((operand) => Number(operand.value));
  if (
    values.some(
      (value) =>
        !Number.isFinite(value) || Math.abs(value) > MAX_ABS_COORDINATE,
    )
  ) {
    refuse(`PDF ${operation.operator} vector coordinates are unsupported.`);
  }
  return values;
}

function multiply(left: Matrix, rightMatrix: Matrix): Matrix {
  return [
    left[0] * rightMatrix[0] + left[2] * rightMatrix[1],
    left[1] * rightMatrix[0] + left[3] * rightMatrix[1],
    left[0] * rightMatrix[2] + left[2] * rightMatrix[3],
    left[1] * rightMatrix[2] + left[3] * rightMatrix[3],
    left[0] * rightMatrix[4] + left[2] * rightMatrix[5] + left[4],
    left[1] * rightMatrix[4] + left[3] * rightMatrix[5] + left[5],
  ];
}

function point(matrix: Matrix, x: number, y: number): PdfNativeVectorPoint {
  return Object.freeze({
    x: matrix[0] * x + matrix[2] * y + matrix[4],
    y: matrix[1] * x + matrix[3] * y + matrix[5],
  });
}

const PATH_POINT_OPERATORS = new Set(["m", "l", "c", "v", "y", "re"]);

/**
 * Whether a path operator's numbers lie beyond what a proof may carry -- a
 * coordinate past MAX_ABS_COORDINATE in the operator or on the page. Such a
 * path is not read (it cannot be a proof), but it is one path: plotting
 * programs clip to the whole plane with `-2147483648 ... m ... W n`, and
 * refusing the drawing over it lost every vector drawn under that clip.
 * Operands that are not numbers are still malformed content.
 */
function pathOperationOutOfRange(operation: ContentOperation, matrix: Matrix): boolean {
  if (
    !PATH_POINT_OPERATORS.has(operation.operator) ||
    operation.operands.some((operand) => operand.kind !== "number")
  ) {
    return false;
  }
  const values = operation.operands.map((operand) => Number(operand.value));
  if (values.some((value) => !Number.isFinite(value) || Math.abs(value) > MAX_ABS_COORDINATE)) {
    return true;
  }
  const pairs: (readonly [number, number])[] = [];
  if (operation.operator === "re" && values.length === 4) {
    const [x, y, width, height] = values;
    pairs.push([x, y], [x + width, y], [x + width, y + height], [x, y + height]);
  } else {
    for (let index = 0; index + 1 < values.length; index += 2) pairs.push([values[index], values[index + 1]]);
  }
  return pairs.some(([x, y]) => {
    const mapped = point(matrix, x, y);
    return (
      !Number.isFinite(mapped.x) ||
      !Number.isFinite(mapped.y) ||
      Math.abs(mapped.x) > MAX_ABS_COORDINATE ||
      Math.abs(mapped.y) > MAX_ABS_COORDINATE
    );
  });
}

function boundedSourcePathPoint(
  matrix: Matrix,
  x: number,
  y: number,
): PdfNativeVectorPoint {
  const mapped = point(matrix, x, y);
  if (
    !Number.isFinite(mapped.x) ||
    !Number.isFinite(mapped.y) ||
    Math.abs(mapped.x) > MAX_ABS_COORDINATE ||
    Math.abs(mapped.y) > MAX_ABS_COORDINATE
  ) {
    refuse("PDF vector subpath page coordinates are unsupported.");
  }
  return mapped;
}

function inverse(matrix: Matrix): Matrix {
  const determinant = matrix[0] * matrix[3] - matrix[1] * matrix[2];
  if (!Number.isFinite(determinant) || Math.abs(determinant) <= 1e-12) {
    refuse("PDF vector transform is singular.");
  }
  return [
    matrix[3] / determinant,
    -matrix[1] / determinant,
    -matrix[2] / determinant,
    matrix[0] / determinant,
    (matrix[2] * matrix[5] - matrix[3] * matrix[4]) / determinant,
    (matrix[1] * matrix[4] - matrix[0] * matrix[5]) / determinant,
  ];
}

function distance(
  left: PdfNativeVectorPoint,
  rightPoint: PdfNativeVectorPoint,
): number {
  return Math.hypot(left.x - rightPoint.x, left.y - rightPoint.y);
}

function vectorLength(value: PdfNativeVectorPoint): number {
  return Math.hypot(value.x, value.y);
}

function add(
  left: PdfNativeVectorPoint,
  rightPoint: PdfNativeVectorPoint,
  scale = 1,
): PdfNativeVectorPoint {
  return Object.freeze({
    x: left.x + rightPoint.x * scale,
    y: left.y + rightPoint.y * scale,
  });
}

function subtract(
  left: PdfNativeVectorPoint,
  rightPoint: PdfNativeVectorPoint,
): PdfNativeVectorPoint {
  return Object.freeze({ x: left.x - rightPoint.x, y: left.y - rightPoint.y });
}

function midpoint(
  left: PdfNativeVectorPoint,
  rightPoint: PdfNativeVectorPoint,
): PdfNativeVectorPoint {
  return Object.freeze({
    x: (left.x + rightPoint.x) / 2,
    y: (left.y + rightPoint.y) / 2,
  });
}

function closePoint(
  left: PdfNativeVectorPoint,
  rightPoint: PdfNativeVectorPoint,
  scale: number,
): boolean {
  return distance(left, rightPoint) <= Math.max(EPSILON, scale * 0.005);
}

function dot(
  left: PdfNativeVectorPoint,
  rightPoint: PdfNativeVectorPoint,
): number {
  return left.x * rightPoint.x + left.y * rightPoint.y;
}

function paintMode(operator: string): PdfNativeVectorPaintMode | null {
  if (operator === "S" || operator === "s") return "stroke";
  if (operator === "f" || operator === "F" || operator === "f*") return "fill";
  if (["B", "B*", "b", "b*"].includes(operator)) return "fill-stroke";
  return null;
}

function paintOperator(operator: string): PdfNativeVectorPaintOperator | null {
  return paintMode(operator)
    ? (operator as PdfNativeVectorPaintOperator)
    : null;
}

function paintClosesPath(operator: PdfNativeVectorPaintOperator): boolean {
  return operator === "s" || operator === "b" || operator === "b*";
}

function emptyPath(): ParsedPath {
  return {
    start: -1,
    end: -1,
    matrix: IDENTITY,
    moves: [],
    lines: [],
    cubics: [],
    rectangles: [],
    closed: false,
    invalid: false,
    reopened: false,
  };
}

function withPath(path: ParsedPath, values: Partial<ParsedPath>): ParsedPath {
  return { ...path, ...values };
}

/** `[x0, y0, x1, y1]` of the page the drawing belongs to (see the graph). */
type PageBox = readonly [number, number, number, number];

/**
 * One clip in effect: its page-space box and, unless it is exactly that
 * axis-aligned box, the polygons of its path (every subpath flattened, a
 * cubic as sixteen chords) with the fill rule that made it. The clip a
 * drawing runs under is the intersection of every clip in effect, or
 * "opaque" once one of them could not be read (an empty path, an empty
 * intersection). A shape lies inside a clip when its box lies inside the
 * clip's box and, for a polygon clip, its four corners are inside by the
 * fill rule and no edge of the clip runs through the box: such a shape is
 * drawn exactly as its bytes say, and may be edited as long as the edit
 * keeps it there. Mirror of the server's _vector_prefix_state clip state.
 */
interface ClipRegion {
  readonly box: PageBox;
  readonly polygons: readonly (readonly PdfNativeVectorPoint[])[] | null;
  readonly evenOdd: boolean;
}
type ClipState = readonly ClipRegion[] | "opaque" | null;

/** How far (pt) a shape may cross its clip and still count inside: a tenth
 * of a point, the rounding gap between the clip and the fill a layout
 * program writes for one frame (InDesign: a clip at y 0.11 over a fill from
 * 0.06); an edit that keeps the shape so reveals no more (mirror of the
 * server). */
const CLIP_TOLERANCE = 0.1;
/** How far (pt) a page clip may fall inside the page and still count as it. */
const PAGE_CLIP_TOLERANCE = 1;
/** Chords one cubic of a clip path is flattened into. */
const CLIP_CURVE_CHORDS = 16;

/** The one refusal an edit under a clip can earn (mirrored). */
export const MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE =
  "PDF vector edit would leave the clip it is drawn under.";

/** The polygon of one flattened clip subpath, page space. */
type ClipPolygon = readonly PdfNativeVectorPoint[];

interface PendingClipPath {
  readonly polygons: ClipPolygon[];
  readonly current: PdfNativeVectorPoint[] | null;
  readonly startPoint: PdfNativeVectorPoint | null;
  readonly invalid: boolean;
}

function boxOfPolygons(polygons: readonly ClipPolygon[]): PageBox | null {
  const xs: number[] = [];
  const ys: number[] = [];
  for (const polygon of polygons) {
    for (const vertex of polygon) {
      xs.push(vertex.x);
      ys.push(vertex.y);
    }
  }
  if (xs.length === 0) return null;
  const box: PageBox = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
  return box[2] - box[0] <= 0 || box[3] - box[1] <= 0 ? null : box;
}

/** The box of four corners that stand on two vertical and two horizontal
 * lines (an `re`, or a rectangle drawn with `m l l l h`, under any transform
 * that keeps it axis-aligned); null for anything else. */
function rectangleClipBox(corners: readonly PdfNativeVectorPoint[]): PageBox | null {
  if (corners.length !== 4) return null;
  const box = boxOfPolygons([corners]);
  if (!box) return null;
  // Four corners on the box's edges are not enough: a bow tie whose vertices
  // are the box's corners passes that test and covers only two triangles of
  // it. A rectangle is walked edge by edge, each edge along one axis and
  // the next along the other.
  const along = (a: PdfNativeVectorPoint, b: PdfNativeVectorPoint): "x" | "y" | null => {
    const dx = Math.abs(b.x - a.x);
    const dy = Math.abs(b.y - a.y);
    if (dx <= 1e-6 && dy > 1e-6) return "y";
    if (dy <= 1e-6 && dx > 1e-6) return "x";
    return null;
  };
  const first = along(corners[0], corners[1]);
  if (!first) return null;
  for (let index = 1; index < 4; index += 1) {
    const axis = along(corners[index], corners[(index + 1) % 4]);
    if (!axis || axis === along(corners[index - 1], corners[index])) return null;
  }
  return box;
}

/** The region a clip path makes: the axis-aligned box when the path is one
 * such rectangle (closed by `h`, by a line back to its start, or by the clip
 * itself), else its polygons. Null when the path makes no region. */
function clipRegionOf(path: PendingClipPath, evenOdd: boolean): ClipRegion | null {
  if (path.invalid) return null;
  const polygons = path.current && path.current.length >= 2 ? [...path.polygons, path.current] : path.polygons;
  if (polygons.length === 0) return null;
  if (polygons.length === 1) {
    let vertices = polygons[0];
    if (
      vertices.length === 5 &&
      Math.abs(vertices[0].x - vertices[4].x) <= 1e-6 &&
      Math.abs(vertices[0].y - vertices[4].y) <= 1e-6
    ) {
      vertices = vertices.slice(0, 4);
    }
    const box = rectangleClipBox(vertices);
    if (box) return Object.freeze({ box, polygons: null, evenOdd });
  }
  const box = boxOfPolygons(polygons);
  if (!box) return null;
  return Object.freeze({ box, polygons: Object.freeze(polygons.map((polygon) => Object.freeze([...polygon]))), evenOdd });
}

function intersectClip(clip: ClipState, region: ClipRegion | null): ClipState {
  if (clip === "opaque" || region === null) return "opaque";
  if (clip) {
    for (const existing of clip) {
      if (
        Math.min(existing.box[2], region.box[2]) - Math.max(existing.box[0], region.box[0]) <= 0 ||
        Math.min(existing.box[3], region.box[3]) - Math.max(existing.box[1], region.box[1]) <= 0
      ) {
        return "opaque";
      }
    }
  }
  return Object.freeze([...(clip ?? []), region]);
}

/** Whether the point is inside the polygons by the fill rule: the winding
 * number for nonzero, the crossing parity for even-odd (the ray runs to +x). */
function pointInsidePolygons(polygons: readonly ClipPolygon[], evenOdd: boolean, x: number, y: number): boolean {
  let winding = 0;
  let crossings = 0;
  for (const polygon of polygons) {
    const count = polygon.length;
    for (let index = 0; index < count; index += 1) {
      const a = polygon[index];
      const b = polygon[(index + 1) % count];
      if (a.y <= y ? b.y > y : b.y <= y) {
        const cross = (b.x - a.x) * (y - a.y) - (x - a.x) * (b.y - a.y);
        if (a.y <= y) {
          if (cross > 0) winding += 1;
        } else if (cross < 0) {
          winding -= 1;
        }
        if (cross !== 0 && (cross > 0) === (a.y <= y)) crossings += 1;
      }
    }
  }
  return evenOdd ? crossings % 2 === 1 : winding !== 0;
}

/** Whether the segment runs through the open box (a touch of its edge is not
 * through it): Liang-Barsky against the closed box, then the clipped
 * segment's midpoint strictly inside. */
function segmentEntersBox(a: PdfNativeVectorPoint, b: PdfNativeVectorPoint, box: PageBox): boolean {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  let t0 = 0;
  let t1 = 1;
  const edges: readonly (readonly [number, number])[] = [
    [-dx, a.x - box[0]],
    [dx, box[2] - a.x],
    [-dy, a.y - box[1]],
    [dy, box[3] - a.y],
  ];
  for (const [p, q] of edges) {
    if (p === 0) {
      if (q < 0) return false;
      continue;
    }
    const t = q / p;
    if (p < 0) {
      if (t > t1) return false;
      if (t > t0) t0 = t;
    } else {
      if (t < t0) return false;
      if (t < t1) t1 = t;
    }
  }
  if (t0 > t1) return false;
  const mid = (t0 + t1) / 2;
  const mx = a.x + dx * mid;
  const my = a.y + dy * mid;
  return mx > box[0] && mx < box[2] && my > box[1] && my < box[3];
}

function regionAllows(region: ClipRegion, bounds: PdfNativeVectorBounds): boolean {
  const x0 = bounds.x;
  const y0 = bounds.y;
  const x1 = bounds.x + bounds.width;
  const y1 = bounds.y + bounds.height;
  if (
    x0 < region.box[0] - CLIP_TOLERANCE ||
    y0 < region.box[1] - CLIP_TOLERANCE ||
    x1 > region.box[2] + CLIP_TOLERANCE ||
    y1 > region.box[3] + CLIP_TOLERANCE
  ) {
    return false;
  }
  if (!region.polygons) return true;
  const corners: readonly (readonly [number, number])[] = [[x0, y0], [x1, y0], [x1, y1], [x0, y1]];
  if (!corners.every(([x, y]) => pointInsidePolygons(region.polygons!, region.evenOdd, x, y))) {
    return false;
  }
  const inner: PageBox = [x0 + CLIP_TOLERANCE, y0 + CLIP_TOLERANCE, x1 - CLIP_TOLERANCE, y1 - CLIP_TOLERANCE];
  if (inner[2] - inner[0] <= 0 || inner[3] - inner[1] <= 0) return true;
  for (const polygon of region.polygons) {
    const count = polygon.length;
    for (let index = 0; index < count; index += 1) {
      if (segmentEntersBox(polygon[index], polygon[(index + 1) % count], inner)) return false;
    }
  }
  return true;
}

/** Whether `bounds` lie inside every clip in effect (a null clip allows all). */
function clipAllows(clip: readonly ClipRegion[] | null, bounds: PdfNativeVectorBounds): boolean {
  return !clip || clip.every((region) => regionAllows(region, bounds));
}

/**
 * The clip in effect at every operation of a drawing (indexed like the
 * operations), read once for both scanners: q/Q stack the state, the path
 * drawn since the last paint is kept in page space (a cubic as sixteen
 * chords, an `re` as its four corners), W/W* intersect it in, and a paint or
 * `n` drops it. A page-sized rectangle restricts nothing and is not kept.
 */
function computeClipStates(
  parsedOperations: readonly ContentOperation[],
  initialMatrix: Matrix,
  pageBox: PageBox | null,
  initialClip: ClipState = null,
): Readonly<{ clips: readonly ClipState[]; matrices: readonly Matrix[] }> {
  const states: ClipState[] = [];
  const matrices: Matrix[] = [];
  let matrix: Matrix = initialMatrix;
  const matrixStack: Matrix[] = [];
  let clip: ClipState = initialClip;
  const clipStack: ClipState[] = [];
  const emptyPath = (): PendingClipPath => ({ polygons: [], current: null, startPoint: null, invalid: false });
  let path = emptyPath();
  const invalidate = () => {
    path = { ...path, invalid: true };
  };
  // A clip may reach past the page by any finite amount (a whole-plane clip
  // at +-2^31 restricts nothing); a transform keeps the proof bound (mirror
  // of the server's _vector_prefix_state).
  const numbers = (operation: ContentOperation, count: number, bounded = false): readonly number[] | null => {
    if (operation.operands.length !== count || operation.operands.some((operand) => operand.kind !== "number")) {
      return null;
    }
    const values = operation.operands.map((operand) => Number(operand.value));
    return values.every((value) => Number.isFinite(value) && (!bounded || Math.abs(value) <= MAX_ABS_COORDINATE))
      ? values
      : null;
  };
  for (const operation of parsedOperations) {
    states.push(clip);
    matrices.push(matrix);
    const operator = operation.operator;
    if (operator === "q") {
      matrixStack.push(matrix);
      clipStack.push(clip);
      continue;
    }
    if (operator === "Q") {
      if (matrixStack.length === 0) continue;
      matrix = matrixStack.pop()!;
      clip = clipStack.pop()!;
      continue;
    }
    if (operator === "cm") {
      const values = numbers(operation, 6, true);
      if (!values) continue;
      if (path.current || path.polygons.length) invalidate();
      matrix = multiply(matrix, values as unknown as Matrix);
      continue;
    }
    if (operator === "m") {
      const values = numbers(operation, 2);
      if (!values) {
        invalidate();
        continue;
      }
      const mapped = point(matrix, values[0], values[1]);
      // A subpath of one point adds nothing to a clip: Illustrator ends its
      // clip rectangles with a move back to the corner (`re m W n`).
      path = {
        ...path,
        polygons: path.current && path.current.length >= 2 ? [...path.polygons, path.current] : path.polygons,
        current: [mapped],
        startPoint: mapped,
      };
      continue;
    }
    if (operator === "l" || operator === "c" || operator === "v" || operator === "y") {
      const values = numbers(operation, operator === "l" ? 2 : operator === "c" ? 6 : 4);
      if (!values || !path.startPoint) {
        invalidate();
        continue;
      }
      // After a close the pen is at the closed subpath's start: a segment
      // there opens the next subpath (mirror of the server's prefix state).
      const current = path.current ?? [path.startPoint];
      const from = current[current.length - 1];
      if (operator === "l") {
        path = { ...path, current: [...current, point(matrix, values[0], values[1])] };
        continue;
      }
      const control1 = operator === "v" ? from : point(matrix, values[0], values[1]);
      const control2 =
        operator === "y" ? point(matrix, values[2], values[3]) : point(matrix, values[values.length - 4], values[values.length - 3]);
      const end = point(matrix, values[values.length - 2], values[values.length - 1]);
      const chords: PdfNativeVectorPoint[] = [];
      for (let index = 1; index <= CLIP_CURVE_CHORDS; index += 1) {
        chords.push(cubicPointAt(from, control1, control2, end, index / CLIP_CURVE_CHORDS));
      }
      path = { ...path, current: [...current, ...chords] };
      continue;
    }
    if (operator === "h") {
      if (!path.startPoint) {
        invalidate();
        continue;
      }
      // The close draws the segment back to the start: the polygon holds it
      // implicitly, and the pen stays at the start for what follows. A
      // subpath already closed (an `re`) stays as it is.
      if (!path.current) continue;
      path = { ...path, polygons: path.current.length >= 2 ? [...path.polygons, path.current] : path.polygons, current: null };
      continue;
    }
    if (operator === "re") {
      const values = numbers(operation, 4);
      if (!values) {
        invalidate();
        continue;
      }
      const [x, y, width, height] = values;
      const corners = [
        point(matrix, x, y),
        point(matrix, x + width, y),
        point(matrix, x + width, y + height),
        point(matrix, x, y + height),
      ];
      // `re` appends a closed subpath and leaves the pen at its corner
      // (PDF 32000-1 8.5.2.1: it is `m l l l h`).
      path = {
        ...path,
        polygons: [...(path.current && path.current.length >= 2 ? [...path.polygons, path.current] : path.polygons), corners],
        current: null,
        startPoint: corners[0],
      };
      continue;
    }
    if (operator === "W" || operator === "W*") {
      const region = clipRegionOf(path, operator === "W*");
      if (!(region && region.polygons === null && rectangleCoversPage(region.box, pageBox))) {
        clip = intersectClip(clip, region);
      }
      path = emptyPath();
      continue;
    }
    if (operator === "n" || paintMode(operator)) {
      path = emptyPath();
      continue;
    }
  }
  return Object.freeze({ clips: Object.freeze(states), matrices: Object.freeze(matrices) });
}

/**
 * Whether one `re` rectangle, drawn under `matrix`, covers the whole visible
 * page. Chromium (Skia), Cairo and other producers open every page with
 * `0 0 w h re W n` under the page transform: that clip restricts nothing a
 * page can show, so a shape under it is as free as one under no clip at all.
 * Only an unrotated, unskewed transform qualifies; the server mirrors this
 * test in `_vector_prefix_state`.
 */
function rectangleCoversPage(
  box: PageBox | null,
  pageBox: PageBox | null,
): boolean {
  if (!pageBox || !box) return false;
  return (
    box[0] <= pageBox[0] + PAGE_CLIP_TOLERANCE &&
    box[1] <= pageBox[1] + PAGE_CLIP_TOLERANCE &&
    box[2] >= pageBox[2] - PAGE_CLIP_TOLERANCE &&
    box[3] >= pageBox[3] - PAGE_CLIP_TOLERANCE
  );
}

function sameMatrix(left: Matrix, rightMatrix: Matrix): boolean {
  return left.every(
    (value, index) => Math.abs(value - rightMatrix[index]) <= EPSILON,
  );
}

function boundsForGeometry(
  geometry: PdfNativeVectorGeometry,
): PdfNativeVectorBounds {
  let points: readonly PdfNativeVectorPoint[];
  if (geometry.kind === "line") {
    points = [geometry.start, geometry.end];
  } else if (geometry.kind === "rectangle") {
    points = geometry.corners;
  } else {
    const xRadius = Math.hypot(geometry.axisX.x, geometry.axisY.x);
    const yRadius = Math.hypot(geometry.axisX.y, geometry.axisY.y);
    return Object.freeze({
      x: geometry.center.x - xRadius,
      y: geometry.center.y - yRadius,
      width: xRadius * 2,
      height: yRadius * 2,
    });
  }
  const x = Math.min(...points.map((entry) => entry.x));
  const y = Math.min(...points.map((entry) => entry.y));
  const x2 = Math.max(...points.map((entry) => entry.x));
  const y2 = Math.max(...points.map((entry) => entry.y));
  return Object.freeze({ x, y, width: x2 - x, height: y2 - y });
}

function exactPoint(
  value: PdfNativeVectorPoint,
  label: string,
): PdfNativeVectorPoint {
  if (
    !Number.isFinite(value.x) ||
    !Number.isFinite(value.y) ||
    Math.abs(value.x) > MAX_ABS_COORDINATE ||
    Math.abs(value.y) > MAX_ABS_COORDINATE
  ) {
    invalid(`Invalid PDF vector ${label}.`);
  }
  return Object.freeze({ x: value.x, y: value.y });
}

function exactGeometry(
  geometry: PdfNativeVectorGeometry,
  expectedKind?: PdfNativeVectorKind,
): PdfNativeVectorGeometry {
  if (expectedKind && geometry.kind !== expectedKind) {
    invalid("PDF vector replacement kind changed.");
  }
  if (geometry.kind === "line") {
    const start = exactPoint(geometry.start, "line start");
    const end = exactPoint(geometry.end, "line end");
    if (distance(start, end) <= EPSILON) invalid("PDF vector line is empty.");
    return Object.freeze({ kind: "line", start, end });
  }
  if (geometry.kind === "rectangle") {
    if (geometry.corners.length !== 4)
      invalid("PDF vector rectangle is invalid.");
    const mapped = geometry.corners.map((entry, index) =>
      exactPoint(entry, `rectangle corner ${index + 1}`),
    );
    const corners: [
      PdfNativeVectorPoint,
      PdfNativeVectorPoint,
      PdfNativeVectorPoint,
      PdfNativeVectorPoint,
    ] = [mapped[0], mapped[1], mapped[2], mapped[3]];
    const sideA = subtract(corners[1], corners[0]);
    const sideB = subtract(corners[3], corners[0]);
    const scale = Math.max(vectorLength(sideA), vectorLength(sideB));
    if (
      scale <= EPSILON ||
      Math.abs(dot(sideA, sideB)) >
        Math.max(EPSILON, vectorLength(sideA) * vectorLength(sideB) * 0.01) ||
      !closePoint(add(corners[0], add(sideA, sideB)), corners[2], scale)
    ) {
      invalid("PDF vector rectangle geometry is not rectangular.");
    }
    return Object.freeze({
      kind: "rectangle",
      corners: Object.freeze(corners),
    });
  }
  const center = exactPoint(geometry.center, "ellipse center");
  const axisX = exactPoint(geometry.axisX, "ellipse x-axis");
  const axisY = exactPoint(geometry.axisY, "ellipse y-axis");
  const xLength = vectorLength(axisX);
  const yLength = vectorLength(axisY);
  if (
    xLength <= EPSILON ||
    yLength <= EPSILON ||
    Math.abs(dot(axisX, axisY)) > xLength * yLength * 0.01
  ) {
    invalid("PDF vector ellipse axes are invalid.");
  }
  return Object.freeze({ kind: "ellipse", center, axisX, axisY });
}

/**
 * Classification asks a question, it does not make a promise: a path the
 * strict families cannot name is simply not one of them, and the page goes on
 * being read. `exactGeometry` is the promise-keeper used on the write side,
 * where a geometry that fails its own checks must stop the edit, so every
 * classifier that asks it a question has to turn that refusal back into "no".
 * Letting one escape ends the whole page's inspection over a single path the
 * producer drew degenerate -- a zero-length stroked rule, an ellipse whose
 * axes collapsed -- and with it every other shape on that page.
 */
function classified(
  build: () => PdfNativeVectorGeometry,
): PdfNativeVectorGeometry | null {
  try {
    return build();
  } catch (error) {
    if (
      error instanceof PdfNativeTransactionError &&
      error.code === "invalid-request"
    ) {
      return null;
    }
    throw error;
  }
}

function classifyRectangle(path: ParsedPath): PdfNativeVectorGeometry | null {
  if (
    path.rectangles.length === 1 &&
    path.lines.length === 0 &&
    path.cubics.length === 0
  ) {
    return classified(() =>
      exactGeometry({
        kind: "rectangle",
        corners: path.rectangles[0],
      }),
    );
  }
  if (
    path.cubics.length !== 0 ||
    path.rectangles.length !== 0 ||
    !path.closed
  ) {
    return null;
  }
  const distinctMoves = path.moves.filter(
    (entry, index, values) =>
      index ===
      values.findIndex((candidate) => closePoint(candidate, entry, 1)),
  );
  if (distinctMoves.length !== 1 || path.lines.length !== 3) return null;
  const corners = [
    distinctMoves[0],
    path.lines[0].end,
    path.lines[1].end,
    path.lines[2].end,
  ] as const;
  return classified(() => exactGeometry({ kind: "rectangle", corners }));
}

function classifyLine(path: ParsedPath): PdfNativeVectorGeometry | null {
  if (
    path.cubics.length !== 0 ||
    path.rectangles.length !== 0 ||
    path.closed ||
    path.lines.length !== 1
  ) {
    return null;
  }
  const move = path.moves.at(-1);
  const scale = distance(path.lines[0].start, path.lines[0].end);
  if (
    !move ||
    path.moves.some(
      (candidate) => !closePoint(candidate, path.lines[0].start, scale),
    )
  ) {
    return null;
  }
  return classified(() =>
    exactGeometry({
      kind: "line",
      start: path.lines[0].start,
      end: path.lines[0].end,
    }),
  );
}

function classifyEllipse(path: ParsedPath): PdfNativeVectorGeometry | null {
  if (
    path.lines.length !== 0 ||
    path.rectangles.length !== 0 ||
    path.cubics.length !== 4 ||
    path.moves.length !== 1
  ) {
    return null;
  }
  const [first, second, third, fourth] = path.cubics;
  const scale = Math.max(
    ...path.cubics.flatMap((segment) => [
      distance(segment.start, segment.end),
      distance(segment.control1, segment.control2),
    ]),
  );
  if (
    !closePoint(first.start, path.moves[0], scale) ||
    !closePoint(first.end, second.start, scale) ||
    !closePoint(second.end, third.start, scale) ||
    !closePoint(third.end, fourth.start, scale) ||
    !closePoint(fourth.end, first.start, scale)
  ) {
    return null;
  }
  const centerA = midpoint(first.start, second.end);
  const centerB = midpoint(first.end, third.end);
  const center = midpoint(centerA, centerB);
  if (!closePoint(centerA, centerB, scale)) return null;
  const axisX = subtract(first.start, center);
  const axisY = subtract(first.end, center);
  const xLength = vectorLength(axisX);
  const yLength = vectorLength(axisY);
  if (
    xLength <= EPSILON ||
    yLength <= EPSILON ||
    Math.abs(dot(axisX, axisY)) > xLength * yLength * 0.01
  ) {
    return null;
  }
  const expectedControls = [
    [add(first.start, axisY, KAPPA), add(first.end, axisX, KAPPA)],
    [add(first.end, axisX, -KAPPA), add(second.end, axisY, KAPPA)],
    [add(second.end, axisY, -KAPPA), add(third.end, axisX, -KAPPA)],
    [add(third.end, axisX, KAPPA), add(fourth.end, axisY, -KAPPA)],
  ];
  if (
    path.cubics.some(
      (segment, index) =>
        !closePoint(segment.control1, expectedControls[index][0], scale) ||
        !closePoint(segment.control2, expectedControls[index][1], scale),
    )
  ) {
    return null;
  }
  return classified(() =>
    exactGeometry({ kind: "ellipse", center, axisX, axisY }),
  );
}

function classifyPath(
  path: ParsedPath,
  mode: PdfNativeVectorPaintMode,
): Readonly<{
  kind: PdfNativeVectorKind;
  geometry: PdfNativeVectorGeometry;
}> | null {
  if (path.invalid || path.start < 0 || path.end <= path.start) return null;
  if (mode === "stroke") {
    const line = classifyLine(path);
    if (line) return { kind: "line", geometry: line };
  }
  const rectangle = classifyRectangle(path);
  if (rectangle) return { kind: "rectangle", geometry: rectangle };
  const ellipse = classifyEllipse(path);
  if (ellipse) return { kind: "ellipse", geometry: ellipse };
  return null;
}

function scanShapes(
  parsedOperations: readonly ContentOperation[],
  initialMatrix: Matrix = IDENTITY,
  pageBox: PageBox | null = null,
  initialClip: ClipState = null,
): readonly ScannedShape[] {
  let matrix: Matrix = initialMatrix;
  const matrixStack: Matrix[] = [];
  const clipStates = computeClipStates(parsedOperations, initialMatrix, pageBox, initialClip).clips;
  let path = emptyPath();
  let currentPoint: PdfNativeVectorPoint | null = null;
  let subpathStart: PdfNativeVectorPoint | null = null;
  const shapes: ScannedShape[] = [];
  for (const [index, operation] of parsedOperations.entries()) {
    const operator = operation.operator;
    const mode = paintMode(operator);
    if (operator === "q") {
      if (operation.operands.length !== 0)
        refuse("PDF q operands are invalid.");
      matrixStack.push(matrix);
      continue;
    }
    if (operator === "Q") {
      if (operation.operands.length !== 0 || matrixStack.length === 0) {
        refuse("PDF vector graphics-state restore is unmatched.");
      }
      matrix = matrixStack.pop()!;
      continue;
    }
    if (operator === "cm") {
      if (path.start >= 0) path = withPath(path, { invalid: true });
      const values = numberOperands(operation, 6) as Matrix;
      matrix = multiply(matrix, values);
      continue;
    }
    if (pathOperationOutOfRange(operation, matrix)) {
      path =
        path.start < 0
          ? withPath(path, { start: operation.start, end: operation.end, matrix, invalid: true })
          : withPath(path, { end: operation.end, invalid: true });
      currentPoint = null;
      subpathStart = null;
      continue;
    }
    if (operator === "m") {
      const [x, y] = numberOperands(operation, 2);
      const mapped = point(matrix, x, y);
      if (path.start < 0) {
        path = withPath(path, {
          start: operation.start,
          end: operation.end,
          matrix,
        });
      } else if (!sameMatrix(path.matrix, matrix)) {
        path = withPath(path, { invalid: true });
      }
      path = withPath(path, {
        end: operation.end,
        moves: [...path.moves, mapped],
        reopened:
          path.reopened ||
          path.lines.length + path.cubics.length + path.rectangles.length > 0,
      });
      currentPoint = mapped;
      subpathStart = mapped;
      continue;
    }
    if (path.reopened && ["l", "c", "v", "y", "re", "h"].includes(operator)) {
      path = withPath(path, { invalid: true });
      continue;
    }
    if (operator === "l") {
      const [x, y] = numberOperands(operation, 2);
      if (!currentPoint || path.start < 0 || !sameMatrix(path.matrix, matrix)) {
        path = withPath(path, { invalid: true });
        continue;
      }
      const mapped = point(matrix, x, y);
      path = withPath(path, {
        end: operation.end,
        lines: [...path.lines, { start: currentPoint, end: mapped }],
      });
      currentPoint = mapped;
      continue;
    }
    if (["c", "v", "y"].includes(operator)) {
      if (!currentPoint || path.start < 0 || !sameMatrix(path.matrix, matrix)) {
        path = withPath(path, { invalid: true });
        continue;
      }
      const count = operator === "c" ? 6 : 4;
      const values = numberOperands(operation, count);
      const control1 =
        operator === "v" ? currentPoint : point(matrix, values[0], values[1]);
      const control2 =
        operator === "y"
          ? point(matrix, values[values.length - 2], values[values.length - 1])
          : point(matrix, values[values.length - 4], values[values.length - 3]);
      const end = point(
        matrix,
        values[values.length - 2],
        values[values.length - 1],
      );
      path = withPath(path, {
        end: operation.end,
        cubics: [
          ...path.cubics,
          { start: currentPoint, control1, control2, end },
        ],
      });
      currentPoint = end;
      continue;
    }
    if (operator === "re") {
      const [x, y, width, height] = numberOperands(operation, 4);
      const corners = Object.freeze([
        point(matrix, x, y),
        point(matrix, x + width, y),
        point(matrix, x + width, y + height),
        point(matrix, x, y + height),
      ]) as readonly [
        PdfNativeVectorPoint,
        PdfNativeVectorPoint,
        PdfNativeVectorPoint,
        PdfNativeVectorPoint,
      ];
      if (path.start < 0) {
        path = withPath(path, {
          start: operation.start,
          end: operation.end,
          matrix,
        });
      } else if (!sameMatrix(path.matrix, matrix)) {
        path = withPath(path, { invalid: true });
      }
      path = withPath(path, {
        end: operation.end,
        rectangles: [...path.rectangles, corners],
      });
      currentPoint = corners[0];
      subpathStart = corners[0];
      continue;
    }
    if (operator === "h") {
      if (
        operation.operands.length !== 0 ||
        !currentPoint ||
        !subpathStart ||
        path.start < 0
      ) {
        path = withPath(path, { invalid: true });
        continue;
      }
      path = withPath(path, { end: operation.end, closed: true });
      currentPoint = subpathStart;
      continue;
    }
    if (mode) {
      if (operation.operands.length !== 0) {
        path = withPath(path, { invalid: true });
      }
      const classified = classifyPath(path, mode);
      const determinant =
        path.matrix[0] * path.matrix[3] - path.matrix[1] * path.matrix[2];
      const bounds = classified ? boundsForGeometry(classified.geometry) : null;
      const clip = clipStates[index];
      if (
        classified &&
        bounds &&
        clip !== "opaque" &&
        clipAllows(clip, bounds) &&
        Number.isFinite(determinant) &&
        Math.abs(determinant) > 1e-12
      ) {
        shapes.push(
          Object.freeze({
            kind: classified.kind,
            paintMode: mode,
            paintOperator: operator,
            geometry: classified.geometry,
            bounds,
            matrix: path.matrix,
            span: Object.freeze({ start: path.start, end: path.end }),
            paintSpan: Object.freeze({
              start: operation.operatorStart,
              end: operation.end,
            }),
            clip,
          }),
        );
      }
      path = emptyPath();
      currentPoint = null;
      subpathStart = null;
      if (shapes.length > MAX_VECTOR_SHAPES_PER_PAGE) {
        refuse("PDF page has too many native vector shapes.");
      }
      continue;
    }
    if (operator === "W" || operator === "W*") {
      // The clip itself is read by computeClipStates; here the path ends.
      path = emptyPath();
      currentPoint = null;
      subpathStart = null;
      continue;
    }
    if (operator === "n") {
      path = emptyPath();
      currentPoint = null;
      subpathStart = null;
      continue;
    }
    if (path.start >= 0 && operator !== "Q") {
      path = withPath(path, { invalid: true });
    }
  }
  if (matrixStack.length !== 0) {
    refuse("PDF vector graphics-state stack is unbalanced.");
  }
  // A transform still active when the drawing ends (Chromium's page flip,
  // set once outside q/Q) maps nothing after it: every shape above carried
  // its own matrix, so no refusal is owed for it.
  return Object.freeze(shapes);
}

function emptyPendingSubpath(): PendingSubpath | null {
  return null;
}

function invalidPendingSubpath(
  operation: ContentOperation,
  matrix: Matrix,
  previous: PendingSubpath | null,
): PendingSubpath {
  return {
    start: previous?.start ?? operation.start,
    end: operation.end,
    matrix: previous?.matrix ?? matrix,
    commands: previous?.commands ?? Object.freeze([]),
    currentPoint: previous?.currentPoint ?? null,
    startPoint: previous?.startPoint ?? null,
    explicitlyClosed: previous?.explicitlyClosed ?? false,
    segments: previous?.segments ?? 0,
    invalid: true,
  };
}

function cubicCoordinateAt(
  start: number,
  control1: number,
  control2: number,
  end: number,
  time: number,
): number {
  const first = start + (control1 - start) * time;
  const second = control1 + (control2 - control1) * time;
  const third = control2 + (end - control2) * time;
  const fourth = first + (second - first) * time;
  const fifth = second + (third - second) * time;
  return fourth + (fifth - fourth) * time;
}

function cubicDerivativeRoots(
  start: number,
  control1: number,
  control2: number,
  end: number,
): readonly number[] {
  const cubic = -start + 3 * control1 - 3 * control2 + end;
  const quadratic = 3 * start - 6 * control1 + 3 * control2;
  const linear = -3 * start + 3 * control1;
  const a = 3 * cubic;
  const b = 2 * quadratic;
  const c = linear;
  const coefficientScale = Math.max(1, Math.abs(a), Math.abs(b), Math.abs(c));
  const coefficientTolerance = Number.EPSILON * 64 * coefficientScale;
  const roots: number[] = [];
  const include = (value: number) => {
    if (
      Number.isFinite(value) &&
      value > 0 &&
      value < 1 &&
      roots.every(
        (candidate) => Math.abs(candidate - value) > Number.EPSILON * 128,
      )
    ) {
      roots.push(value);
    }
  };
  if (Math.abs(a) <= coefficientTolerance) {
    if (Math.abs(b) > coefficientTolerance) include(-c / b);
    return Object.freeze(roots.sort((left, right) => left - right));
  }
  const discriminant = b * b - 4 * a * c;
  const discriminantTolerance =
    Number.EPSILON * 128 * coefficientScale * coefficientScale;
  if (discriminant < -discriminantTolerance) return Object.freeze([]);
  if (Math.abs(discriminant) <= discriminantTolerance) {
    include(-b / (2 * a));
    return Object.freeze(roots);
  }
  const squareRoot = Math.sqrt(discriminant);
  const q = -0.5 * (b + (b < 0 ? -squareRoot : squareRoot));
  if (Math.abs(q) <= coefficientTolerance) {
    include((-b + squareRoot) / (2 * a));
    include((-b - squareRoot) / (2 * a));
  } else {
    include(q / a);
    include(c / q);
  }
  return Object.freeze(roots.sort((left, right) => left - right));
}

function cubicPointAt(
  start: PdfNativeVectorPoint,
  control1: PdfNativeVectorPoint,
  control2: PdfNativeVectorPoint,
  end: PdfNativeVectorPoint,
  time: number,
): PdfNativeVectorPoint {
  return Object.freeze({
    x: cubicCoordinateAt(start.x, control1.x, control2.x, end.x, time),
    y: cubicCoordinateAt(start.y, control1.y, control2.y, end.y, time),
  });
}

function boundsForSubpath(
  commands: readonly PdfNativeVectorSubpathCommand[],
): PdfNativeVectorBounds {
  const points: PdfNativeVectorPoint[] = [];
  let currentPoint: PdfNativeVectorPoint | null = null;
  let startPoint: PdfNativeVectorPoint | null = null;
  let segments = 0;
  let closed = false;
  for (const command of commands) {
    if (command.operator === "m") {
      // The next subpath of a compound path; a move straight after a move
      // draws nothing and is not a subpath.
      if (currentPoint && segments === 0 && !closed) refuse("PDF vector subpath has a lone move.");
      currentPoint = command.point;
      startPoint = command.point;
      segments = 0;
      closed = false;
      continue;
    }
    if (!currentPoint || !startPoint) {
      refuse("PDF vector subpath does not start with a move.");
    }
    if (command.operator === "l") {
      // After a close the pen is at the closed subpath's start and a
      // segment opens the next subpath there.
      points.push(currentPoint, command.point);
      currentPoint = command.point;
      segments += 1;
      closed = false;
      continue;
    }
    if (command.operator === "h") {
      // `m h` closes a subpath of one point: it paints nothing but a cap.
      if (closed) refuse("PDF vector subpath closes nothing.");
      points.push(currentPoint, startPoint);
      currentPoint = startPoint;
      closed = true;
      continue;
    }
    if (!("control1" in command)) {
      refuse("PDF vector subpath curve topology is invalid.");
    }
    points.push(currentPoint, command.end);
    const roots = new Set([
      ...cubicDerivativeRoots(
        currentPoint.x,
        command.control1.x,
        command.control2.x,
        command.end.x,
      ),
      ...cubicDerivativeRoots(
        currentPoint.y,
        command.control1.y,
        command.control2.y,
        command.end.y,
      ),
    ]);
    for (const time of roots) {
      points.push(
        cubicPointAt(
          currentPoint,
          command.control1,
          command.control2,
          command.end,
          time,
        ),
      );
    }
    currentPoint = command.end;
    segments += 1;
    closed = false;
  }
  if (currentPoint && segments === 0 && !closed) refuse("PDF vector subpath has a lone move.");
  if (points.length < 2) refuse("PDF vector subpath has no visible segment.");
  const x = Math.min(...points.map((entry) => entry.x));
  const y = Math.min(...points.map((entry) => entry.y));
  const x2 = Math.max(...points.map((entry) => entry.x));
  const y2 = Math.max(...points.map((entry) => entry.y));
  if (
    ![x, y, x2, y2].every(Number.isFinite) ||
    Math.max(x2 - x, y2 - y) <= EPSILON
  ) {
    refuse("PDF vector subpath bounds are degenerate.");
  }
  return Object.freeze({ x, y, width: x2 - x, height: y2 - y });
}

/** A painted path's bounds, or null when it covers no area to select: a
 * path of `m h` point subpaths alone paints nothing a reader could pick
 * (it stays in the drawing, and only this page's proofs skip it). */
function paintedSubpathBounds(
  commands: readonly PdfNativeVectorSubpathCommand[],
): PdfNativeVectorBounds | null {
  try {
    return boundsForSubpath(commands);
  } catch (error) {
    if (error instanceof PdfNativeTransactionError && error.code === "unsupported-document") {
      return null;
    }
    throw error;
  }
}

function scanSubpaths(
  parsedOperations: readonly ContentOperation[],
  initialMatrix: Matrix = IDENTITY,
  pageBox: PageBox | null = null,
  initialClip: ClipState = null,
): readonly ScannedSubpath[] {
  let matrix: Matrix = initialMatrix;
  const matrixStack: Matrix[] = [];
  const clipStates = computeClipStates(parsedOperations, initialMatrix, pageBox, initialClip).clips;
  let path = emptyPendingSubpath();
  const subpaths: ScannedSubpath[] = [];
  for (const [index, operation] of parsedOperations.entries()) {
    const operator = operation.operator;
    const mode = paintMode(operator);
    const supportedPaint = paintOperator(operator);
    if (operator === "q") {
      if (operation.operands.length !== 0)
        refuse("PDF q operands are invalid.");
      if (path) path = invalidPendingSubpath(operation, matrix, path);
      matrixStack.push(matrix);
      continue;
    }
    if (operator === "Q") {
      if (operation.operands.length !== 0 || matrixStack.length === 0) {
        refuse("PDF vector graphics-state restore is unmatched.");
      }
      if (path) path = invalidPendingSubpath(operation, matrix, path);
      matrix = matrixStack.pop()!;
      continue;
    }
    if (operator === "cm") {
      if (path) path = invalidPendingSubpath(operation, matrix, path);
      const values = numberOperands(operation, 6) as Matrix;
      matrix = multiply(matrix, values);
      continue;
    }
    if (pathOperationOutOfRange(operation, matrix)) {
      path = invalidPendingSubpath(operation, matrix, path);
      continue;
    }
    if (operator === "m") {
      const [x, y] = numberOperands(operation, 2);
      const mapped = boundedSourcePathPoint(matrix, x, y);
      if (path) {
        // A move after a drawn subpath opens the next subpath of the same
        // path (a ring, an icon, a letter's outline): one compound path,
        // painted as one. A move straight after a move draws nothing and is
        // not read (mirror of the server's _parse_native_vector_path).
        if (
          path.invalid ||
          (path.segments === 0 && !path.explicitlyClosed) ||
          !sameMatrix(path.matrix, matrix) ||
          path.commands.length >= MAX_NATIVE_SUBPATH_COMMANDS
        ) {
          path = invalidPendingSubpath(operation, matrix, {
            ...path,
            currentPoint: mapped,
            startPoint: mapped,
          });
          continue;
        }
        path = {
          ...path,
          end: operation.end,
          commands: Object.freeze([
            ...path.commands,
            Object.freeze({ operator: "m" as const, point: mapped }),
          ]),
          currentPoint: mapped,
          startPoint: mapped,
          explicitlyClosed: false,
          segments: 0,
        };
        continue;
      }
      path = {
        start: operation.start,
        end: operation.end,
        matrix,
        commands: Object.freeze([
          Object.freeze({ operator: "m" as const, point: mapped }),
        ]),
        currentPoint: mapped,
        startPoint: mapped,
        explicitlyClosed: false,
        segments: 0,
        invalid: false,
      };
      continue;
    }
    if (operator === "l") {
      const [x, y] = numberOperands(operation, 2);
      // A segment straight after `h` starts the next subpath at the closed
      // one's start, where `h` left the pen (PDF 32000-1 8.5.2.1): a frame
      // drawn as `re`-like squares one after another without a move.
      if (
        !path ||
        !path.currentPoint ||
        !sameMatrix(path.matrix, matrix)
      ) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      if (path.commands.length >= MAX_NATIVE_SUBPATH_COMMANDS) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      // A segment of no length draws nothing but is part of the path
      // (InDesign opens shapes with `0 0 m 0 0 l`); it moves with the rest.
      const mapped = boundedSourcePathPoint(matrix, x, y);
      path = {
        ...path,
        end: operation.end,
        commands: Object.freeze([
          ...path.commands,
          Object.freeze({ operator: "l" as const, point: mapped }),
        ]),
        currentPoint: mapped,
        explicitlyClosed: false,
        segments: path.segments + 1,
      };
      continue;
    }
    if (operator === "c" || operator === "v" || operator === "y") {
      const values = numberOperands(operation, operator === "c" ? 6 : 4);
      if (
        !path ||
        !path.currentPoint ||
        !sameMatrix(path.matrix, matrix)
      ) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      if (path.commands.length >= MAX_NATIVE_SUBPATH_COMMANDS) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      const control1 =
        operator === "v"
          ? path.currentPoint
          : boundedSourcePathPoint(matrix, values[0], values[1]);
      const control2 =
        operator === "y"
          ? boundedSourcePathPoint(
              matrix,
              values[values.length - 2],
              values[values.length - 1],
            )
          : boundedSourcePathPoint(
              matrix,
              values[values.length - 4],
              values[values.length - 3],
            );
      const end = boundedSourcePathPoint(
        matrix,
        values[values.length - 2],
        values[values.length - 1],
      );
      path = {
        ...path,
        end: operation.end,
        commands: Object.freeze([
          ...path.commands,
          Object.freeze({ operator, control1, control2, end }),
        ]),
        currentPoint: end,
        explicitlyClosed: false,
        segments: path.segments + 1,
      };
      continue;
    }
    if (operator === "h") {
      if (
        operation.operands.length !== 0 ||
        !path ||
        !path.currentPoint ||
        !path.startPoint ||
        path.explicitlyClosed ||
        !sameMatrix(path.matrix, matrix)
      ) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      if (path.commands.length >= MAX_NATIVE_SUBPATH_COMMANDS) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      path = {
        ...path,
        end: operation.end,
        commands: Object.freeze([
          ...path.commands,
          Object.freeze({ operator: "h" as const }),
        ]),
        currentPoint: path.startPoint,
        explicitlyClosed: true,
      };
      continue;
    }
    if (operator === "re") {
      const [x, y, width, height] = numberOperands(operation, 4);
      // A rectangle in a path of several subpaths (a frame with its inner
      // panel, a table of cells painted at once) is the closed subpath of
      
      
      const corners = [
        point(matrix, x, y),
        point(matrix, x + width, y),
        point(matrix, x + width, y + height),
        point(matrix, x, y + height),
      ];
      const rectangleCommands = Object.freeze([
        Object.freeze({ operator: "m" as const, point: corners[0] }),
        Object.freeze({ operator: "l" as const, point: corners[1] }),
        Object.freeze({ operator: "l" as const, point: corners[2] }),
        Object.freeze({ operator: "l" as const, point: corners[3] }),
        Object.freeze({ operator: "h" as const }),
      ]);
      if (
        !(width !== 0 && height !== 0) ||
        (path &&
          (path.invalid ||
            path.segments === 0 ||
            !sameMatrix(path.matrix, matrix) ||
            path.commands.length + rectangleCommands.length > MAX_NATIVE_SUBPATH_COMMANDS))
      ) {
        path = invalidPendingSubpath(operation, matrix, path);
        continue;
      }
      path = path
        ? {
            ...path,
            end: operation.end,
            commands: Object.freeze([...path.commands, ...rectangleCommands]),
            currentPoint: corners[0],
            startPoint: corners[0],
            explicitlyClosed: true,
            segments: 3,
          }
        : {
            start: operation.start,
            end: operation.end,
            matrix,
            commands: rectangleCommands,
            currentPoint: corners[0],
            startPoint: corners[0],
            explicitlyClosed: true,
            segments: 3,
            invalid: false,
          };
      continue;
    }
    if (mode && supportedPaint) {
      if (operation.operands.length !== 0 && path) {
        path = invalidPendingSubpath(operation, matrix, path);
      }
      if (
        path &&
        !path.invalid &&
        path.segments === 0 &&
        !path.explicitlyClosed &&
        path.commands.length >= 2
      ) {
        // A move after the last subpath draws nothing (Word moves back to the
        // start after its close); it is not part of the path.
        path = {
          ...path,
          commands: Object.freeze(path.commands.slice(0, -1)),
          explicitlyClosed: path.commands.at(-2)?.operator === "h",
          segments: 1,
        };
      }
      const clip = clipStates[index];
      if (path) {
        const determinant =
          path.matrix[0] * path.matrix[3] - path.matrix[1] * path.matrix[2];
        const subpathBounds =
          !path.invalid &&
          clip !== "opaque" &&
          path.commands.length >= 2 &&
          path.commands.length <= MAX_NATIVE_SUBPATH_COMMANDS &&
          path.commands[0].operator === "m"
            ? paintedSubpathBounds(path.commands)
            : null;
        if (
          subpathBounds &&
          clip !== "opaque" &&
          clipAllows(clip, subpathBounds) &&
          Number.isFinite(determinant) &&
          Math.abs(determinant) > 1e-12
        ) {
          subpaths.push(
            Object.freeze({
              paintMode: mode,
              paintOperator: supportedPaint,
              commands: Object.freeze(path.commands),
              closed: path.explicitlyClosed || paintClosesPath(supportedPaint),
              bounds: subpathBounds,
              matrix: path.matrix,
              span: Object.freeze({ start: path.start, end: path.end }),
              clip,
              paintSpan: Object.freeze({
                start: operation.operatorStart,
                end: operation.end,
              }),
            }),
          );
        }
      }
      path = emptyPendingSubpath();
      if (subpaths.length > MAX_VECTOR_SHAPES_PER_PAGE) {
        refuse("PDF page has too many native vector subpaths.");
      }
      continue;
    }
    if (operator === "W" || operator === "W*") {
      if (operation.operands.length !== 0) {
        refuse(`PDF ${operator} operands are invalid.`);
      }
      // The clip itself is read by computeClipStates; here the path ends.
      path = emptyPendingSubpath();
      continue;
    }
    if (operator === "n") {
      if (operation.operands.length !== 0)
        refuse("PDF n operands are invalid.");
      path = emptyPendingSubpath();
      continue;
    }
    if (path) path = invalidPendingSubpath(operation, matrix, path);
  }
  if (matrixStack.length !== 0) {
    refuse("PDF vector graphics-state stack is unbalanced.");
  }
  // See scanShapes: a transform left active at the end maps nothing.
  return Object.freeze(subpaths);
}

function rounded(value: number): number {
  const result = Number(value.toFixed(6));
  return Object.is(result, -0) ? 0 : result;
}

function canonicalGeometry(geometry: PdfNativeVectorGeometry): unknown {
  if (geometry.kind === "line") {
    return {
      kind: geometry.kind,
      start: { x: rounded(geometry.start.x), y: rounded(geometry.start.y) },
      end: { x: rounded(geometry.end.x), y: rounded(geometry.end.y) },
    };
  }
  if (geometry.kind === "rectangle") {
    return {
      kind: geometry.kind,
      corners: geometry.corners.map((entry) => ({
        x: rounded(entry.x),
        y: rounded(entry.y),
      })),
    };
  }
  return {
    kind: geometry.kind,
    center: { x: rounded(geometry.center.x), y: rounded(geometry.center.y) },
    axisX: { x: rounded(geometry.axisX.x), y: rounded(geometry.axisX.y) },
    axisY: { x: rounded(geometry.axisY.x), y: rounded(geometry.axisY.y) },
  };
}

function shapeFingerprint(
  shape: Pick<PdfNativeVectorContentShape, "kind" | "paintMode" | "geometry">,
): string {
  return JSON.stringify({
    kind: shape.kind,
    paintMode: shape.paintMode,
    geometry: canonicalGeometry(shape.geometry),
  });
}

function canonicalSubpathCommand(
  command: PdfNativeVectorSubpathCommand,
): unknown {
  if ("point" in command) {
    return {
      operator: command.operator,
      point: { x: rounded(command.point.x), y: rounded(command.point.y) },
    };
  }
  if (command.operator === "h") return { operator: command.operator };
  return {
    operator: command.operator,
    control1: {
      x: rounded(command.control1.x),
      y: rounded(command.control1.y),
    },
    control2: {
      x: rounded(command.control2.x),
      y: rounded(command.control2.y),
    },
    end: { x: rounded(command.end.x), y: rounded(command.end.y) },
  };
}

function canonicalSubpathCommands(
  commands: readonly PdfNativeVectorSubpathCommand[],
): readonly unknown[] {
  return commands.map(canonicalSubpathCommand);
}

function subpathFingerprint(
  subpath: Pick<
    PdfNativeVectorSubpathProof,
    "paintMode" | "paintOperator" | "commands" | "closed"
  >,
): string {
  return JSON.stringify({
    paintMode: subpath.paintMode,
    paintOperator: subpath.paintOperator,
    commands: canonicalSubpathCommands(subpath.commands),
    closed: subpath.closed,
  });
}

async function proofDigest(input: {
  sourceSha256: string;
  pageNumber: number;
  stream: Readonly<{ objectNumber: number; generation: number }>;
  span: Readonly<{ start: number; end: number; expectedSha256: string }>;
  kind: PdfNativeVectorKind;
  paintMode: PdfNativeVectorPaintMode;
  geometry: PdfNativeVectorGeometry;
}): Promise<string> {
  return sha256PdfBytes(
    ascii(
      JSON.stringify({
        schemaVersion: MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION,
        sourceSha256: input.sourceSha256,
        pageNumber: input.pageNumber,
        stream: input.stream,
        span: input.span,
        kind: input.kind,
        paintMode: input.paintMode,
        geometry: canonicalGeometry(input.geometry),
      }),
    ),
  );
}

async function subpathProofDigest(input: {
  sourceSha256: string;
  pageNumber: number;
  stream: Readonly<{ objectNumber: number; generation: number }>;
  span: Readonly<{ start: number; end: number; expectedSha256: string }>;
  paintMode: PdfNativeVectorPaintMode;
  paintOperator: PdfNativeVectorPaintOperator;
  commands: readonly PdfNativeVectorSubpathCommand[];
  closed: boolean;
}): Promise<string> {
  return sha256PdfBytes(
    ascii(
      JSON.stringify({
        schemaVersion: MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
        sourceSha256: input.sourceSha256,
        pageNumber: input.pageNumber,
        stream: input.stream,
        span: input.span,
        paintMode: input.paintMode,
        paintOperator: input.paintOperator,
        commands: canonicalSubpathCommands(input.commands),
        closed: input.closed,
      }),
    ),
  );
}

async function inspectCombinedPage(
  sourceBytes: Uint8Array,
  page: PdfNativePageContentGraphInspection,
  mode: "shapes" | "subpaths" | "combined",
  knownReferenceCount?: number,
  maxDrawingBytes?: number,
): Promise<InternalCombinedPageInspection> {
  const shapes: InternalShape[] = [];
  const subpaths: InternalSubpath[] = [];
  const shapeLosses: Array<{
    contentStream: Readonly<{ objectNumber: number; generation: number }>;
    reason: "unsupported-content-stream";
  }> = [];
  const subpathLosses: Array<{
    contentStream: Readonly<{ objectNumber: number; generation: number }>;
    reason: "unsupported-content-stream";
  }> = [];
  let subpathError: unknown | null = null;
  let excludedExistingShapeCount = 0;
  // A page drawn from several streams is one drawing: the shapes are scanned
  // across the whole of it, so a graphics state opened in one stream still
  // maps the geometry a later stream draws, and every shape is then attributed
  // to the stream that holds its own operators.
  let inspections: readonly PdfNativePageContentStreamInspection[];
  try {
    inspections = await inspectMraiNativePageContentStreams(
      sourceBytes,
      page.pageNumber,
      { pageRole: "untouched-object" },
    );
  } catch (error) {
    if (
      error instanceof PdfNativeTransactionError &&
      (error.code === "unsupported-document" ||
        error.code === "precondition-mismatch")
    ) {
      const losses = Object.freeze(
        page.contentStreams.map((contentStream) =>
          Object.freeze({
            contentStream: Object.freeze({ ...contentStream }),
            reason: "unsupported-content-stream" as const,
          }),
        ),
      );
      return Object.freeze({
        sourceSha256: page.sourceSha256,
        pageNumber: page.pageNumber,
        shapes: Object.freeze([]),
        subpaths: Object.freeze([]),
        excludedExistingShapeCount,
        shapeLosses: losses,
        subpathLosses: losses,
        subpathError,
      });
    }
    throw error;
  }
  const drawingBytes = joinNativePageDrawing(inspections);
  const ownerOf = (span: Readonly<{ start: number; end: number }>) =>
    inspections.find(
      (inspection) =>
        span.start >= inspection.drawingOffset &&
        span.end <= inspection.drawingOffset + inspection.streamBytes.length,
    ) ?? null;
  const unsupportedLoss = () =>
    Object.freeze({
      contentStream: Object.freeze({ ...inspections[0].contentStream }),
      reason: "unsupported-content-stream" as const,
    });
  const emptyResult = () =>
    Object.freeze({
      sourceSha256: page.sourceSha256,
      pageNumber: page.pageNumber,
      shapes: Object.freeze([]),
      subpaths: Object.freeze([]),
      excludedExistingShapeCount,
      shapeLosses: Object.freeze(shapeLosses),
      subpathLosses: Object.freeze(subpathLosses),
      subpathError,
    });
  // A page drawn by megabytes of operators (a datasheet's 5.5 MB diagram
  // page took 15 s alone) is not read for a bounded caller: its vectors are
  // left as one unsupported stream, as an unreadable stream is.
  if (maxDrawingBytes !== undefined && drawingBytes.length > maxDrawingBytes) {
    shapeLosses.push(unsupportedLoss());
    subpathLosses.push(unsupportedLoss());
    return emptyResult();
  }
  let scannedShapes: readonly ScannedShape[] = [];
  let scannedSubpaths: readonly ScannedSubpath[] = [];
  let parsedOperations: readonly ContentOperation[];
  try {
    // An inline image is one opaque painting: it sets no transform, no clip
    // and draws no path, so the vectors around it are read as on any page
    // (a scanned logo or a checkbox glyph written as `BI ... EI` used to
    // take every vector of its page down with it). Mirror of the server's
    // _vector_prefix_state.
    parsedOperations = operations(latin1(drawingBytes), true);
  } catch (error) {
    if (
      error instanceof PdfNativeTransactionError &&
      error.code === "unsupported-document"
    ) {
      if (mode !== "subpaths") shapeLosses.push(unsupportedLoss());
      if (mode !== "shapes") subpathLosses.push(unsupportedLoss());
      return emptyResult();
    }
    throw error;
  }
  if (mode === "subpaths") {
    try {
      scannedSubpaths = scanSubpaths(parsedOperations, IDENTITY, page.visibleBox);
    } catch (error) {
      if (
        error instanceof PdfNativeTransactionError &&
        error.code === "unsupported-document"
      ) {
        subpathLosses.push(unsupportedLoss());
        return emptyResult();
      }
      throw error;
    }
  }
  try {
    scannedShapes = scanShapes(parsedOperations, IDENTITY, page.visibleBox);
  } catch (error) {
    if (
      error instanceof PdfNativeTransactionError &&
      error.code === "unsupported-document"
    ) {
      if (mode !== "subpaths") shapeLosses.push(unsupportedLoss());
      if (mode !== "shapes") subpathLosses.push(unsupportedLoss());
      return emptyResult();
    }
    throw error;
  }
  if (mode === "combined") {
    try {
      scannedSubpaths = scanSubpaths(parsedOperations, IDENTITY, page.visibleBox);
    } catch (error) {
      if (
        error instanceof PdfNativeTransactionError &&
        error.code === "unsupported-document"
      ) {
        subpathLosses.push(unsupportedLoss());
      } else {
        subpathError = error;
      }
    }
  }
  const streamReferenceCounts = new Map<number, number>();
  for (const inspection of inspections) {
    const count =
      knownReferenceCount === undefined || inspections.length !== 1
        ? countMraiNativeIndirectObjectReferences(
            sourceBytes,
            inspection.contentStream,
          )
        : knownReferenceCount;
    if (!Number.isSafeInteger(count) || count < 0) {
      invalid("PDF content-stream reference authority is invalid.");
    }
    streamReferenceCounts.set(inspection.streamIndex, count);
  }
  const attributeShape = async (
    scannedShape: ScannedShape,
    streamInspection: PdfNativeContentStreamInspection,
    contentStream: Readonly<{ objectNumber: number; generation: number }>,
    drawingOffset: number,
    referenceCount: number,
    formInstance?: string,
  ): Promise<void> => {
    const spanBytes = streamInspection.streamBytes.slice(
      scannedShape.span.start - drawingOffset,
      scannedShape.span.end - drawingOffset,
    );
    const expectedSha256 = await sha256PdfBytes(spanBytes);
    const pageVectorIndex = shapes.length;
    const locator = Object.freeze({
      start: scannedShape.span.start - drawingOffset,
      end: scannedShape.span.end - drawingOffset,
      expectedSha256,
    });
    const digest = await proofDigest({
      sourceSha256: page.sourceSha256,
      pageNumber: page.pageNumber,
      stream: contentStream,
      span: locator,
      kind: scannedShape.kind,
      paintMode: scannedShape.paintMode,
      geometry: scannedShape.geometry,
    });
    shapes.push(
      Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION,
        id: pdfNativeVectorProofId({
          family: "shape",
          sourceSha256: page.sourceSha256,
          pageNumber: page.pageNumber,
          objectNumber: contentStream.objectNumber,
          generation: contentStream.generation,
          start: locator.start,
          end: locator.end,
          formInstance,
        }),
        sourceSha256: page.sourceSha256,
        pageNumber: page.pageNumber,
        pageVectorIndex,
        kind: scannedShape.kind,
        paintMode: scannedShape.paintMode,
        geometry: scannedShape.geometry,
        bounds: scannedShape.bounds,
        contentStream: Object.freeze({ ...contentStream }),
        operatorSpan: locator,
        proofDigest: digest,
        editable: referenceCount === 1,
        preserveReason: referenceCount === 1 ? null : "shared-content-stream",
        streamInspection,
        matrix: scannedShape.matrix,
        clip: scannedShape.clip,
        paintSpan: Object.freeze({
          start: scannedShape.paintSpan.start - drawingOffset,
          end: scannedShape.paintSpan.end - drawingOffset,
        }),
        paintOperatorToken: scannedShape.paintOperator,
        ...(formInstance !== undefined ? { formInstance } : {}),
      }),
    );
  };
  const attributeSubpath = async (
    scannedSubpath: ScannedSubpath,
    streamInspection: PdfNativeContentStreamInspection,
    contentStream: Readonly<{ objectNumber: number; generation: number }>,
    drawingOffset: number,
    referenceCount: number,
    formInstance?: string,
  ): Promise<void> => {
    const spanBytes = streamInspection.streamBytes.slice(
      scannedSubpath.span.start - drawingOffset,
      scannedSubpath.span.end - drawingOffset,
    );
    const expectedSha256 = await sha256PdfBytes(spanBytes);
    const pageSubpathIndex = subpaths.length;
    const locator = Object.freeze({
      start: scannedSubpath.span.start - drawingOffset,
      end: scannedSubpath.span.end - drawingOffset,
      expectedSha256,
    });
    const digest = await subpathProofDigest({
      sourceSha256: page.sourceSha256,
      pageNumber: page.pageNumber,
      stream: contentStream,
      span: locator,
      paintMode: scannedSubpath.paintMode,
      paintOperator: scannedSubpath.paintOperator,
      commands: scannedSubpath.commands,
      closed: scannedSubpath.closed,
    });
    subpaths.push(
      Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
        id: pdfNativeVectorProofId({
          family: "subpath",
          sourceSha256: page.sourceSha256,
          pageNumber: page.pageNumber,
          objectNumber: contentStream.objectNumber,
          generation: contentStream.generation,
          start: locator.start,
          end: locator.end,
          formInstance,
        }),
        sourceSha256: page.sourceSha256,
        pageNumber: page.pageNumber,
        pageSubpathIndex,
        paintMode: scannedSubpath.paintMode,
        paintOperator: scannedSubpath.paintOperator,
        commands: scannedSubpath.commands,
        closed: scannedSubpath.closed,
        bounds: scannedSubpath.bounds,
        contentStream: Object.freeze({ ...contentStream }),
        operatorSpan: locator,
        proofDigest: digest,
        editable: referenceCount === 1,
        preserveReason: referenceCount === 1 ? null : "shared-content-stream",
        streamInspection,
        matrix: scannedSubpath.matrix,
        clip: scannedSubpath.clip,
        paintSpan: Object.freeze({
          start: scannedSubpath.paintSpan.start - drawingOffset,
          end: scannedSubpath.paintSpan.end - drawingOffset,
        }),
        ...(formInstance !== undefined ? { formInstance } : {}),
      }),
    );
  };
  if (mode !== "subpaths") {
    for (const scannedShape of scannedShapes) {
      // A shape whose operators run across a stream boundary belongs to no
      // single object, so it is preserved rather than offered for editing.
      const streamInspection = ownerOf({
        start: scannedShape.span.start,
        end: Math.max(scannedShape.span.end, scannedShape.paintSpan.end),
      });
      if (!streamInspection) {
        shapeLosses.push(unsupportedLoss());
        continue;
      }
      await attributeShape(
        scannedShape,
        streamInspection,
        streamInspection.contentStream,
        streamInspection.drawingOffset,
        streamReferenceCounts.get(streamInspection.streamIndex)!,
      );
    }
  }
  const subpathInspectionAvailable =
    subpathLosses.length === 0 && subpathError === null;
  // A subpath that is one of the shapes above is that shape's; the key is
  // the stream and the span, since a form's offsets start again at zero.
  const spanKey = (
    stream: Readonly<{ objectNumber: number; generation: number }>,
    span: Readonly<{ start: number; end: number }>,
  ) => `${stream.objectNumber}:${stream.generation}:${span.start}:${span.end}`;
  const existingSpanKeys = new Set<string>();
  if (subpathInspectionAvailable) {
    for (const shape of scannedShapes) {
      const owner = ownerOf({ start: shape.span.start, end: Math.max(shape.span.end, shape.paintSpan.end) });
      if (owner) existingSpanKeys.add(spanKey(owner.contentStream, { start: shape.span.start - owner.drawingOffset, end: shape.span.end - owner.drawingOffset }));
    }
  }
  excludedExistingShapeCount = subpathInspectionAvailable
    ? scannedShapes.length
    : 0;
  for (const scannedSubpath of subpathInspectionAvailable
    ? scannedSubpaths
    : []) {
    const streamInspection = ownerOf({
      start: scannedSubpath.span.start,
      end: Math.max(scannedSubpath.span.end, scannedSubpath.paintSpan.end),
    });
    if (!streamInspection) {
      subpathLosses.push(unsupportedLoss());
      continue;
    }
    if (
      existingSpanKeys.has(
        spanKey(streamInspection.contentStream, {
          start: scannedSubpath.span.start - streamInspection.drawingOffset,
          end: scannedSubpath.span.end - streamInspection.drawingOffset,
        }),
      )
    ) {
      continue;
    }
    await attributeSubpath(
      scannedSubpath,
      streamInspection,
      streamInspection.contentStream,
      streamInspection.drawingOffset,
      streamReferenceCounts.get(streamInspection.streamIndex)!,
    );
  }
  // Vectors inside a Form XObject the page draws once and nothing else
  // references: the form's stream is scanned as a drawing of its own, under
  // the graphics state its `Do` inherits (the transform there times the
  // form's Matrix, the clip there cut to its BBox), and every shape found
  // is attributed to the form's stream, which the writer rewrites. A form
  // drawn twice, or one reached through another form, is not read here
  // (mirror of the server's form target in _vector_page_authority).
  let hosts: readonly PdfNativeVectorFormHost[] = [];
  try {
    hosts = await inspectMraiNativeVectorFormHosts(sourceBytes, page.pageNumber);
  } catch (error) {
    if (!(error instanceof PdfNativeTransactionError)) throw error;
  }
  if (hosts.length > 0) {
    const pageStates = computeClipStates(parsedOperations, IDENTITY, page.visibleBox);
    const pageDoIndexes = new Map<number, number>();
    parsedOperations.forEach((operation, index) => {
      if (operation.operator === "Do") pageDoIndexes.set(operation.start, index);
    });
    // The state a Form is entered in: the transform at its `Do` times its
    // own Matrix, the clip there cut to its BBox; null once nothing shows.
    const enter = (
      outerMatrix: Matrix,
      outerClip: ClipState,
      level: PdfNativeVectorFormHost["path"][number],
    ): Readonly<{ matrix: Matrix; clip: ClipState }> | null => {
      if (outerClip === "opaque") return null;
      const formMatrix = multiply(outerMatrix, level.matrix as unknown as Matrix);
      let initialClip: ClipState = outerClip;
      if (level.bbox) {
        const [x0, y0, x1, y1] = level.bbox;
        const corners = [
          point(formMatrix, x0, y0),
          point(formMatrix, x1, y0),
          point(formMatrix, x1, y1),
          point(formMatrix, x0, y1),
        ];
        const region = clipRegionOf({ polygons: [corners], current: null, startPoint: null, invalid: false }, false);
        // A BBox that covers the whole visible page (the form is the page's
        // size, as producers write a page-sized form) restricts nothing a
        // page can show, like the page-sized `re W n` a page opens with
        // (rectangleCoversPage): a shape that runs past the page's edge is
        // read as drawn, cut by the page itself. Cut by the BBox instead,
        // such a shape was not read at all (the IPCC report's bleed section
        
        if (!(region && region.polygons === null && rectangleCoversPage(region.box, page.visibleBox))) {
          initialClip = intersectClip(outerClip, region);
          if (initialClip === "opaque") return null;
        }
      }
      return Object.freeze({ matrix: formMatrix, clip: initialClip });
    };
    // A Form's operators, read once however often it is drawn (a chart's
    // marker is drawn thousands of times), and the states inside a Form
    // entered one way, shared by every drawing that goes on from there.
    const formOperationsByStream = new Map<string, readonly ContentOperation[] | null>();
    const formOperationsOf = (stream: PdfNativeContentStreamInspection, form: Readonly<{ objectNumber: number; generation: number }>) => {
      const key = `${form.objectNumber}:${form.generation}`;
      if (!formOperationsByStream.has(key)) {
        try {
          formOperationsByStream.set(key, operations(latin1(stream.streamBytes), true));
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError && error.code === "unsupported-document")) throw error;
          formOperationsByStream.set(key, null);
        }
      }
      return formOperationsByStream.get(key)!;
    };
    const statesByWay = new Map<
      string,
      Readonly<{ states: ReturnType<typeof computeClipStates>; doIndexes: ReadonlyMap<number, number> }> | null
    >();
    for (const host of hosts) {
      // A form written on a clone (shared with other pages, drawn more than
      // once, or drawn inside another) carries the chain every writer adds
      // to its request.
      const hostInspection: VectorStreamInspection = host.formInstanceIsolation
        ? Object.freeze({ ...host.streamInspection, formInstanceIsolation: host.formInstanceIsolation })
        : host.streamInspection;
      const doIndex = pageDoIndexes.get(host.path[0].doOffset) ?? -1;
      if (doIndex < 0) continue;
      const outerClip = pageStates.clips[doIndex];
      if (outerClip === "opaque") continue;
      let entered = enter(pageStates.matrices[doIndex], outerClip, host.path[0]);
      for (let depth = 1; entered && depth < host.path.length; depth += 1) {
        const parentLevel = host.path[depth - 1];
        const wayKey = host.path
          .slice(0, depth)
          .map((level) => `${level.form.objectNumber}:${level.form.generation}@${level.doOffset}`)
          .join("/");
        let way = statesByWay.get(wayKey);
        if (way === undefined) {
          const parentOperations = formOperationsOf(parentLevel.streamInspection, parentLevel.form);
          if (!parentOperations) {
            way = null;
          } else {
            try {
              const states = computeClipStates(parentOperations, entered.matrix, page.visibleBox, entered.clip);
              const doIndexes = new Map<number, number>();
              parentOperations.forEach((operation, index) => {
                if (operation.operator === "Do") doIndexes.set(operation.start, index);
              });
              way = Object.freeze({ states, doIndexes });
            } catch (error) {
              if (!(error instanceof PdfNativeTransactionError && error.code === "unsupported-document")) throw error;
              way = null;
            }
          }
          statesByWay.set(wayKey, way);
        }
        const childIndex = way?.doIndexes.get(host.path[depth].doOffset);
        entered =
          way && childIndex !== undefined
            ? enter(way.states.matrices[childIndex], way.states.clips[childIndex], host.path[depth])
            : null;
      }
      if (!entered) continue;
      const formMatrix = entered.matrix;
      const initialClip = entered.clip;
      let formShapes: readonly ScannedShape[] = [];
      let formSubpaths: readonly ScannedSubpath[] = [];
      const formOperations = formOperationsOf(hostInspection, host.form);
      if (!formOperations) continue;
      try {
        formShapes = scanShapes(formOperations, formMatrix, page.visibleBox, initialClip);
      } catch (error) {
        if (error instanceof PdfNativeTransactionError && error.code === "unsupported-document") continue;
        throw error;
      }
      // A form whose paths the subpath family cannot read keeps its shapes:
      // one combined reading offers what the shape reading alone offers.
      let formSubpathsRead = mode !== "shapes" && subpathInspectionAvailable;
      if (formSubpathsRead) {
        try {
          formSubpaths = scanSubpaths(formOperations, formMatrix, page.visibleBox, initialClip);
        } catch (error) {
          if (!(error instanceof PdfNativeTransactionError && error.code === "unsupported-document")) throw error;
          formSubpathsRead = false;
        }
      }
      if (mode === "subpaths" && !formSubpathsRead) continue;
      if (mode !== "subpaths") {
        for (const scannedShape of formShapes) {
          await attributeShape(scannedShape, hostInspection, host.form, 0, 1, host.formInstance);
        }
      }
      if (formSubpathsRead) {
        const formSpanKeys = new Set(formShapes.map((shape) => spanKey(host.form, shape.span)));
        excludedExistingShapeCount += formShapes.length;
        for (const scannedSubpath of formSubpaths) {
          if (formSpanKeys.has(spanKey(host.form, scannedSubpath.span))) continue;
          await attributeSubpath(scannedSubpath, hostInspection, host.form, 0, 1, host.formInstance);
        }
      }
    }
  }
  return Object.freeze({
    sourceSha256: page.sourceSha256,
    pageNumber: page.pageNumber,
    shapes: Object.freeze(shapes),
    subpaths: Object.freeze(subpaths),
    excludedExistingShapeCount,
    shapeLosses: Object.freeze(shapeLosses),
    subpathLosses: Object.freeze(subpathLosses),
    subpathError,
  });
}

/** The most recent pages' combined proof readings, by the SHA-256 of the
 * source they were read from and the page: an edit reads its page's proofs
 * about a dozen times (the route, the plan, the writer's proof, the reopened
 * candidate, then all of it again when the adapter applies the command to
 * its own copy of the same bytes), and a schematic sheet of fifty thousand
 * paths took seconds per reading. The key is the bytes' digest, so a copy
 * finds the reading and nothing another source holds can be served. The
 * shape and subpath readings are views of it: the same proofs, in the same
 * order, that a reading of one family alone makes. */
const combinedPageInspections = new Map<
  string,
  Promise<InternalCombinedPageInspection>
>();
/** An edit's source and candidate, and the page read before them. */
const COMBINED_PAGE_INSPECTION_LIMIT = 3;

async function cachedCombinedPage(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<InternalCombinedPageInspection> {
  const page = await inspectMraiNativePageContentGraph(sourceBytes, pageNumber);
  const key = `${page.sourceSha256}:${page.pageNumber}`;
  const known = combinedPageInspections.get(key);
  if (known) {
    combinedPageInspections.delete(key);
    combinedPageInspections.set(key, known);
    return known;
  }
  const pending = inspectCombinedPage(sourceBytes, page, "combined");
  combinedPageInspections.set(key, pending);
  // A failed reading is not remembered: the next caller reads again.
  pending.catch(() => {
    if (combinedPageInspections.get(key) === pending) {
      combinedPageInspections.delete(key);
    }
  });
  while (combinedPageInspections.size > COMBINED_PAGE_INSPECTION_LIMIT) {
    combinedPageInspections.delete(combinedPageInspections.keys().next().value!);
  }
  return pending;
}

async function inspectInternal(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<
  Readonly<{
    sourceSha256: string;
    pageNumber: number;
    shapes: readonly InternalShape[];
    losses: PdfNativeVectorContentInspection["losses"];
  }>
> {
  const combined = await cachedCombinedPage(sourceBytes, pageNumber);
  return Object.freeze({
    sourceSha256: combined.sourceSha256,
    pageNumber: combined.pageNumber,
    shapes: combined.shapes,
    losses: combined.shapeLosses,
  });
}

/** The matrix each published selection was scanned under. */
const publishedSelectionMatrices = new WeakMap<object, Matrix>();
/** The clip each published selection was scanned under (null: none). */
const publishedSelectionClips = new WeakMap<object, readonly ClipRegion[] | null>();

function publicShape(shape: InternalShape): PdfNativeVectorContentShape {
  const {
    streamInspection: _streamInspection,
    matrix,
    paintSpan: _paintSpan,
    paintOperatorToken: _paintOperatorToken,
    ...visible
  } = shape;
  const published = Object.freeze(visible);
  publishedSelectionMatrices.set(published, matrix);
  publishedSelectionClips.set(published, shape.clip);
  return published;
}

function publicSubpath(subpath: InternalSubpath): PdfNativeVectorSubpathProof {
  const {
    streamInspection: _streamInspection,
    matrix,
    paintSpan: _paintSpan,
    ...visible
  } = subpath;
  const published = Object.freeze(visible);
  publishedSelectionMatrices.set(published, matrix);
  publishedSelectionClips.set(published, subpath.clip);
  return published;
}

/**
 * Whether `bounds` (page space) lie inside every clip the published selection
 * is drawn under -- the rule a moved or resized shape answers to (mirror of
 * the replay's _vector_clip_allows). False for an object no scan published.
 */
export function pdfNativeVectorSelectionClipAllows(
  selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof,
  bounds: PdfNativeVectorBounds,
): boolean {
  if (!publishedSelectionClips.has(selection)) return false;
  return clipAllows(publishedSelectionClips.get(selection) ?? null, bounds);
}

/**
 * The matrix a native vector is painted under -- its own user space to the
 * page's -- as the scan that published this selection read it; null for an
 * object no scan published.
 */
export function pdfNativeVectorSelectionMatrix(
  selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof,
): readonly [number, number, number, number, number, number] | null {
  return publishedSelectionMatrices.get(selection) ?? null;
}

export async function inspectMraiNativeVectorContent(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativeVectorContentInspection> {
  const inspection = await inspectInternal(sourceBytes, pageNumber);
  return Object.freeze({
    sourceSha256: inspection.sourceSha256,
    pageNumber: inspection.pageNumber,
    shapes: Object.freeze(inspection.shapes.map(publicShape)),
    losses: inspection.losses,
  });
}

async function inspectSubpathsInternal(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<
  Readonly<{
    sourceSha256: string;
    pageNumber: number;
    subpaths: readonly InternalSubpath[];
    excludedExistingShapeCount: number;
    losses: PdfNativeVectorSubpathInspection["losses"];
  }>
> {
  const combined = await cachedCombinedPage(sourceBytes, pageNumber);
  if (combined.subpathError) throw combined.subpathError;
  return Object.freeze({
    sourceSha256: combined.sourceSha256,
    pageNumber: combined.pageNumber,
    subpaths: combined.subpaths,
    excludedExistingShapeCount: combined.excludedExistingShapeCount,
    losses: combined.subpathLosses,
  });
}

export async function inspectMraiNativeVectorSubpaths(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfNativeVectorSubpathInspection> {
  const inspection = await inspectSubpathsInternal(sourceBytes, pageNumber);
  return Object.freeze({
    sourceSha256: inspection.sourceSha256,
    pageNumber: inspection.pageNumber,
    subpaths: Object.freeze(inspection.subpaths.map(publicSubpath)),
    excludedExistingShapeCount: inspection.excludedExistingShapeCount,
    losses: inspection.losses,
  });
}

/**
 * Inspect all pages from one pinned page-tree parse. Each eligible content
 * stream is authenticated and scanned once, then shared by the strict-shape
 * and generic-subpath proof builders.
 */
export async function inspectMraiNativeVectorDocumentAuthority(
  sourceBytes: Uint8Array,
): Promise<PdfNativeVectorDocumentAuthorityInspection> {
  return (await inspectVectorDocumentAuthorityPages(sourceBytes, 1, null)).authority;
}









export async function inspectMraiNativeVectorDocumentAuthorityPages(
  sourceBytes: Uint8Array,
  firstPage: number,
  budgetMs: number,
): Promise<Readonly<{ authority: PdfNativeVectorDocumentAuthorityInspection; totalPages: number }>> {
  if (!Number.isSafeInteger(firstPage) || firstPage < 1 || !(budgetMs > 0)) {
    invalid("PDF vector document page range is invalid.");
  }
  return inspectVectorDocumentAuthorityPages(sourceBytes, firstPage, budgetMs);
}

/** The drawing a bounded (per-chunk) read takes on one page. */
const MAX_BOUNDED_PAGE_DRAWING_BYTES = 600 * 1024;








type VectorDocumentReadContext = Readonly<{
  fingerprint: string;
  pageGraphs: Awaited<ReturnType<typeof inspectMraiNativeDocumentPageContentGraphs>>;
  referenceCountByStream: ReadonlyMap<string, number>;
}>;
const vectorDocumentReadContexts = new WeakMap<Uint8Array, Promise<VectorDocumentReadContext>>();
function vectorDocumentFingerprint(sourceBytes: Uint8Array): string {
  const head = sourceBytes.subarray(0, 64);
  const tail = sourceBytes.subarray(Math.max(0, sourceBytes.length - 64));
  return `${sourceBytes.length}:${latin1(head)}:${latin1(tail)}`;
}
async function vectorDocumentReadContext(sourceBytes: Uint8Array): Promise<VectorDocumentReadContext> {
  const fingerprint = vectorDocumentFingerprint(sourceBytes);
  const cached = vectorDocumentReadContexts.get(sourceBytes);
  if (cached) {
    const context = await cached.catch(() => null);
    if (context && context.fingerprint === fingerprint) return context;
  }
  const read = (async (): Promise<VectorDocumentReadContext> => {
    const pageGraphs =
      await inspectMraiNativeDocumentPageContentGraphs(sourceBytes, {
        includeExtGStates: true,
      });
    const contentStreams = Array.from(
      new Map(
        pageGraphs.flatMap((page) =>
          page.contentStreams.map((stream) => [
            `${stream.objectNumber}:${stream.generation}`,
            stream,
          ] as const),
        ),
      ).values(),
    );
    const referenceCounts = countMraiNativeIndirectObjectReferencesBatch(
      sourceBytes,
      contentStreams,
    );
    return Object.freeze({
      fingerprint,
      pageGraphs,
      referenceCountByStream: new Map(
        contentStreams.map((stream, index) => [
          `${stream.objectNumber}:${stream.generation}`,
          referenceCounts[index],
        ]),
      ),
    });
  })();
  vectorDocumentReadContexts.set(sourceBytes, read);
  return read;
}

async function inspectVectorDocumentAuthorityPages(
  sourceBytes: Uint8Array,
  firstPage: number,
  budgetMs: number | null,
): Promise<Readonly<{ authority: PdfNativeVectorDocumentAuthorityInspection; totalPages: number }>> {
  const started = Date.now();
  const { pageGraphs, referenceCountByStream } = await vectorDocumentReadContext(sourceBytes);
  if (firstPage > pageGraphs.length) invalid("PDF vector document page range is invalid.");
  const pages: PdfNativeVectorDocumentPageInspection[] = [];
  const authorityPages: PdfNativeVectorDocumentPageAuthority[] = [];
  let sourceSha256: string | null = null;
  for (let index = 0; index < pageGraphs.length; index += 1) {
    const pageGraph = pageGraphs[index];
    if (
      pageGraph.pageNumber !== index + 1 ||
      (sourceSha256 !== null && pageGraph.sourceSha256 !== sourceSha256)
    ) {
      invalid("PDF vector document page authority is inconsistent.");
    }
    sourceSha256 = pageGraph.sourceSha256;
    if (index + 1 < firstPage) continue;
    if (budgetMs !== null && pages.length > 0 && Date.now() - started > budgetMs) break;
    const combined = await inspectCombinedPage(
      sourceBytes,
      pageGraph,
      "combined",
      pageGraph.contentStreams.length === 1
        ? referenceCountByStream.get(
            `${pageGraph.contentStreams[0].objectNumber}:${pageGraph.contentStreams[0].generation}`,
          )
        : undefined,
      budgetMs !== null ? MAX_BOUNDED_PAGE_DRAWING_BYTES : undefined,
    );
    if (combined.subpathError) throw combined.subpathError;
    pages.push(
      Object.freeze({
        pageNumber: combined.pageNumber,
        shapes: Object.freeze(combined.shapes.map(publicShape)),
        subpaths: Object.freeze(combined.subpaths.map(publicSubpath)),
        excludedExistingShapeCount: combined.excludedExistingShapeCount,
        shapeLosses: combined.shapeLosses,
        subpathLosses: combined.subpathLosses,
      }),
    );
    authorityPages.push(
      Object.freeze({
        pageNumber: combined.pageNumber,
        pageGraph,
        contentStream:
          combined.shapes[0]?.streamInspection ??
          combined.subpaths[0]?.streamInspection ??
          null,
      }),
    );
  }
  if (!sourceSha256 || pages.length === 0) {
    invalid("PDF vector document has no visible pages.");
  }
  return Object.freeze({
    authority: Object.freeze({
      document: Object.freeze({
        sourceSha256,
        pageCount: pages.length,
        pages: Object.freeze(pages),
      }),
      pages: Object.freeze(authorityPages),
    }),
    totalPages: pageGraphs.length,
  });
}

export async function inspectMraiNativeVectorDocument(
  sourceBytes: Uint8Array,
): Promise<PdfNativeVectorDocumentInspection> {
  return (await inspectMraiNativeVectorDocumentAuthority(sourceBytes)).document;
}

function formatNumber(value: number): string {
  const roundedValue = rounded(value);
  if (!Number.isFinite(roundedValue)) invalid("PDF vector number is invalid.");
  return roundedValue
    .toFixed(6)
    .replace(/\.(\d*?)0+$/, (_, kept: string) => (kept ? `.${kept}` : "")) // no lookbehind: Safari < 16.4
    .replace(/\.$/, "");
}

function formatSubpathNumber(value: number): string {
  if (!Number.isFinite(value)) invalid("PDF vector number is invalid.");
  const fixed = value
    .toFixed(15)
    .replace(/\.(\d*?)0+$/, (_, kept: string) => (kept ? `.${kept}` : "")) // no lookbehind: Safari < 16.4
    .replace(/\.$/, "");
  return /^-0(?:\.0*)?$/.test(fixed) ? "0" : fixed;
}

function pathReplacement(
  geometryInput: PdfNativeVectorGeometry,
  matrix: Matrix,
): Uint8Array {
  const geometry = exactGeometry(geometryInput);
  const inverseMatrix = inverse(matrix);
  const localPoint = (value: PdfNativeVectorPoint) =>
    point(inverseMatrix, value.x, value.y);
  const xy = (value: PdfNativeVectorPoint) =>
    `${formatNumber(value.x)} ${formatNumber(value.y)}`;
  if (geometry.kind === "line") {
    return ascii(
      `${xy(localPoint(geometry.start))} m\n${xy(localPoint(geometry.end))} l`,
    );
  }
  if (geometry.kind === "rectangle") {
    const corners = geometry.corners.map(localPoint);
    return ascii(
      `${xy(corners[0])} m\n${xy(corners[1])} l\n${xy(corners[2])} l\n${xy(corners[3])} l\nh`,
    );
  }
  const rightPoint = add(geometry.center, geometry.axisX);
  const topPoint = add(geometry.center, geometry.axisY);
  const leftPoint = add(geometry.center, geometry.axisX, -1);
  const bottomPoint = add(geometry.center, geometry.axisY, -1);
  const segments = [
    [
      rightPoint,
      add(rightPoint, geometry.axisY, KAPPA),
      add(topPoint, geometry.axisX, KAPPA),
      topPoint,
    ],
    [
      topPoint,
      add(topPoint, geometry.axisX, -KAPPA),
      add(leftPoint, geometry.axisY, KAPPA),
      leftPoint,
    ],
    [
      leftPoint,
      add(leftPoint, geometry.axisY, -KAPPA),
      add(bottomPoint, geometry.axisX, -KAPPA),
      bottomPoint,
    ],
    [
      bottomPoint,
      add(bottomPoint, geometry.axisX, KAPPA),
      add(rightPoint, geometry.axisY, -KAPPA),
      rightPoint,
    ],
  ] as const;
  const lines = [`${xy(localPoint(rightPoint))} m`];
  for (const [, control1, control2, end] of segments) {
    lines.push(
      `${xy(localPoint(control1))} ${xy(localPoint(control2))} ${xy(localPoint(end))} c`,
    );
  }
  return ascii(lines.join("\n"));
}

/** The shape a reopen reads from path bytes written under `matrix`. The
 * writers round each coordinate in the path's own space, so what lands on
 * the page is these bytes read back through the CTM, not the page-space
 * geometry that was asked for. */
function writtenShapeGeometry(
  bytes: Uint8Array,
  paintOperatorToken: string,
  matrix: Matrix,
  kind: PdfNativeVectorKind,
): PdfNativeVectorGeometry {
  const shapes = scanShapes(
    operations(`${latin1(bytes)}\n${paintOperatorToken}`),
    matrix,
  );
  if (shapes.length !== 1 || shapes[0].kind !== kind) {
    invalid("PDF vector replacement does not read back as the same shape.");
  }
  return shapes[0].geometry;
}

function writtenSubpathCommands(
  bytes: Uint8Array,
  paintOperatorToken: string,
  matrix: Matrix,
): readonly PdfNativeVectorSubpathCommand[] {
  const subpaths = scanSubpaths(
    operations(`${latin1(bytes)}\n${paintOperatorToken}`),
    matrix,
  );
  if (subpaths.length !== 1) {
    invalid("PDF vector replacement does not read back as one subpath.");
  }
  return subpaths[0].commands;
}

function exactAffineTransform(
  transform: PdfNativeVectorAffineTransform,
): Readonly<{ publicValue: PdfNativeVectorAffineTransform; matrix: Matrix }> {
  const values: Matrix = [
    transform.a,
    transform.b,
    transform.c,
    transform.d,
    transform.e,
    transform.f,
  ];
  if (
    values.some(
      (value) =>
        !Number.isFinite(value) || Math.abs(value) > MAX_ABS_COORDINATE,
    )
  ) {
    invalid("PDF vector affine transform is invalid.");
  }
  const determinant = values[0] * values[3] - values[1] * values[2];
  if (!Number.isFinite(determinant) || Math.abs(determinant) <= 1e-12) {
    invalid("PDF vector affine transform is singular.");
  }
  if (sameMatrix(values, IDENTITY)) {
    invalid("PDF vector affine transform is a no-op.");
  }
  return Object.freeze({
    publicValue: Object.freeze({
      a: values[0],
      b: values[1],
      c: values[2],
      d: values[3],
      e: values[4],
      f: values[5],
    }),
    matrix: values,
  });
}

function transformSubpathCommands(
  commands: readonly PdfNativeVectorSubpathCommand[],
  transform: Matrix,
): readonly PdfNativeVectorSubpathCommand[] {
  const mapped = commands.map((command, index) => {
    if (command.operator === "h") {
      return Object.freeze({ operator: "h" as const });
    }
    if ("point" in command) {
      return Object.freeze({
        operator: command.operator,
        point: exactPoint(
          point(transform, command.point.x, command.point.y),
          `subpath command ${index + 1}`,
        ),
      });
    }
    return Object.freeze({
      operator: command.operator,
      control1: exactPoint(
        point(transform, command.control1.x, command.control1.y),
        `subpath command ${index + 1} control 1`,
      ),
      control2: exactPoint(
        point(transform, command.control2.x, command.control2.y),
        `subpath command ${index + 1} control 2`,
      ),
      end: exactPoint(
        point(transform, command.end.x, command.end.y),
        `subpath command ${index + 1} end`,
      ),
    });
  });
  if (
    mapped.length < 2 ||
    mapped.length > MAX_NATIVE_SUBPATH_COMMANDS ||
    mapped[0].operator !== "m"
  ) {
    invalid("PDF vector subpath command sequence is invalid.");
  }
  boundsForSubpath(mapped);
  return Object.freeze(mapped);
}

function subpathReplacement(
  commands: readonly PdfNativeVectorSubpathCommand[],
  pathMatrix: Matrix,
): Uint8Array {
  const inverseMatrix = inverse(pathMatrix);
  const localPoint = (value: PdfNativeVectorPoint) =>
    exactPoint(point(inverseMatrix, value.x, value.y), "subpath local point");
  const xy = (value: PdfNativeVectorPoint) =>
    `${formatSubpathNumber(value.x)} ${formatSubpathNumber(value.y)}`;
  const lines: string[] = [];
  let currentPoint: PdfNativeVectorPoint | null = null;
  let startPoint: PdfNativeVectorPoint | null = null;
  let closed = false;
  let segments = 0;
  for (const [index, command] of commands.entries()) {
    if (command.operator === "m") {
      if ((index !== 0 && segments === 0 && !closed) || (index === 0 && currentPoint)) {
        invalid("PDF vector subpath has a lone move operator.");
      }
      const local = localPoint(command.point);
      lines.push(`${xy(local)} m`);
      currentPoint = command.point;
      startPoint = command.point;
      closed = false;
      segments = 0;
      continue;
    }
    if (!currentPoint || !startPoint || (closed && command.operator === "h")) {
      invalid("PDF vector subpath topology changed.");
    }
    // A segment after a close goes on from the closed subpath's start.
    if (command.operator !== "h") closed = false;
    segments += 1;
    if (command.operator === "l") {
      const local = localPoint(command.point);
      lines.push(`${xy(local)} l`);
      currentPoint = command.point;
      continue;
    }
    if (command.operator === "h") {
      // A close ends its subpath (`m h` a subpath of one point); the next
      // command, if any, opens another.
      segments -= 1;
      lines.push("h");
      currentPoint = startPoint;
      closed = true;
      continue;
    }
    if (!("control1" in command)) {
      invalid("PDF vector subpath curve topology changed.");
    }
    if (
      command.operator === "v" &&
      distance(command.control1, currentPoint) > EPSILON
    ) {
      invalid("PDF vector v implicit control topology changed.");
    }
    if (
      command.operator === "y" &&
      distance(command.control2, command.end) > EPSILON
    ) {
      invalid("PDF vector y implicit control topology changed.");
    }
    const control1 = localPoint(command.control1);
    const control2 = localPoint(command.control2);
    const end = localPoint(command.end);
    if (command.operator === "c") {
      lines.push(`${xy(control1)} ${xy(control2)} ${xy(end)} c`);
    } else if (command.operator === "v") {
      lines.push(`${xy(control2)} ${xy(end)} v`);
    } else {
      lines.push(`${xy(control1)} ${xy(end)} y`);
    }
    currentPoint = command.end;
  }
  if (lines.length !== commands.length) {
    invalid("PDF vector subpath replacement is incomplete.");
  }
  return ascii(lines.join("\n"));
}

function sameSubpathAuthority(
  actual: PdfNativeVectorSubpathProof,
  supplied: PdfNativeVectorSubpathProof,
): boolean {
  try {
    return (
      supplied.schemaVersion === MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION &&
      actual.id === supplied.id &&
      actual.sourceSha256 === supplied.sourceSha256 &&
      actual.pageNumber === supplied.pageNumber &&
      actual.pageSubpathIndex === supplied.pageSubpathIndex &&
      actual.proofDigest === supplied.proofDigest &&
      actual.contentStream.objectNumber ===
        supplied.contentStream.objectNumber &&
      actual.contentStream.generation === supplied.contentStream.generation &&
      actual.operatorSpan.start === supplied.operatorSpan.start &&
      actual.operatorSpan.end === supplied.operatorSpan.end &&
      actual.operatorSpan.expectedSha256 ===
        supplied.operatorSpan.expectedSha256 &&
      subpathFingerprint(actual) === subpathFingerprint(supplied)
    );
  } catch {
    return false;
  }
}

function sameShapeAuthority(
  actual: PdfNativeVectorContentShape,
  supplied: PdfNativeVectorContentShape,
): boolean {
  try {
    return (
      supplied.schemaVersion === MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION &&
      actual.id === supplied.id &&
      actual.sourceSha256 === supplied.sourceSha256 &&
      actual.pageNumber === supplied.pageNumber &&
      actual.pageVectorIndex === supplied.pageVectorIndex &&
      actual.proofDigest === supplied.proofDigest &&
      actual.contentStream.objectNumber ===
        supplied.contentStream.objectNumber &&
      actual.contentStream.generation === supplied.contentStream.generation &&
      actual.operatorSpan.start === supplied.operatorSpan.start &&
      actual.operatorSpan.end === supplied.operatorSpan.end &&
      actual.operatorSpan.expectedSha256 ===
        supplied.operatorSpan.expectedSha256 &&
      shapeFingerprint(actual) === shapeFingerprint(supplied)
    );
  } catch {
    return false;
  }
}

export async function planMraiNativeVectorDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorDeleteSelection;
}): Promise<PdfNativeVectorDeletePlan> {
  const inspection = await cachedCombinedPage(
    input.sourceBytes,
    input.pageNumber,
  );
  if (inspection.subpathError) throw inspection.subpathError;
  const shapes = inspection.shapes;
  const subpaths = inspection.subpaths;
  const shapeSelection = "kind" in input.selection;
  const selected = shapeSelection
    ? shapes.find((shape) => shape.id === input.selection.id)
    : subpaths.find((subpath) => subpath.id === input.selection.id);
  const authorityMatches =
    selected && shapeSelection && "kind" in selected
      ? sameShapeAuthority(selected, input.selection)
      : selected && !shapeSelection && !("kind" in selected)
        ? sameSubpathAuthority(selected, input.selection)
        : false;
  if (!selected || !authorityMatches) {
    invalid("Selected PDF vector delete proof is stale or unavailable.");
  }
  if (!selected.editable || selected.preserveReason !== null) {
    refuse("Selected PDF vector cannot be deleted from a shared stream.");
  }
  if (!SHA256_PATTERN.test(selected.proofDigest)) {
    invalid("Selected PDF vector delete proof digest is invalid.");
  }
  const originalOperatorBytes = Uint8Array.from(
    selected.streamInspection.streamBytes.slice(
      selected.operatorSpan.start,
      selected.operatorSpan.end,
    ),
  );
  if (
    originalOperatorBytes.length === 0 ||
    (await sha256PdfBytes(originalOperatorBytes)) !==
      selected.operatorSpan.expectedSha256
  ) {
    invalid("Selected PDF vector delete span is inconsistent.");
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: selected.operatorSpan,
    replacement: ascii("n"),
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
    selection: shapeSelection
      ? publicShape(selected as InternalShape)
      : publicSubpath(selected as InternalSubpath),
    request,
    originalOperatorBytes,
    beforeShapeFingerprints: Object.freeze(
      shapes.map(shapeFingerprint),
    ),
    beforeSubpathFingerprints: Object.freeze(
      subpaths.map(subpathFingerprint),
    ),
  });
}

export async function prepareMraiNativeVectorDelete(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorDeletePlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (
    plan.schemaVersion !== MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION ||
    plan.selection.sourceSha256 !== plan.request.identity.sourceSha256 ||
    plan.request.replacement.length !== 1 ||
    plan.request.replacement[0] !== 0x6e
  ) {
    invalid("PDF vector delete plan authority is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeVectorDeleteCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorDeletePlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorDeleteVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const inspection = await cachedCombinedPage(
    prepared.candidateBytes,
    plan.selection.pageNumber,
  );
  if (inspection.subpathError) throw inspection.subpathError;
  let proofFamily: "shape" | "subpath";
  let deletedIndex: number;
  if ("kind" in plan.selection) {
    proofFamily = "shape";
    deletedIndex = plan.selection.pageVectorIndex;
  } else {
    proofFamily = "subpath";
    deletedIndex = plan.selection.pageSubpathIndex;
  }
  const expectedShapes = plan.beforeShapeFingerprints.filter(
    (_fingerprint, index) => proofFamily !== "shape" || index !== deletedIndex,
  );
  const expectedSubpaths = plan.beforeSubpathFingerprints.filter(
    (_fingerprint, index) =>
      proofFamily !== "subpath" || index !== deletedIndex,
  );
  if (
    JSON.stringify(inspection.shapes.map(shapeFingerprint)) !==
      JSON.stringify(expectedShapes) ||
    JSON.stringify(inspection.subpaths.map(subpathFingerprint)) !==
      JSON.stringify(expectedSubpaths)
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector deletion changed an unselected object.",
    );
  }
  return Object.freeze({
    verified: true,
    pageNumber: plan.selection.pageNumber,
    deletedProofFamily: proofFamily,
    deletedProofIndex: deletedIndex,
    candidateSha256: prepared.candidateSha256,
  });
}

export async function planMraiNativeVectorGeometryEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selectionId: string;
  readonly replacement: PdfNativeVectorGeometry;
}): Promise<PdfNativeVectorGeometryEditPlan> {
  const inspection = await inspectInternal(input.sourceBytes, input.pageNumber);
  const selected = inspection.shapes.find(
    (shape) => shape.id === input.selectionId,
  );
  if (!selected) invalid("Selected PDF vector proof is stale or unavailable.");
  if (!selected.editable) {
    refuse("Selected PDF vector content stream is shared.");
  }
  if (!SHA256_PATTERN.test(selected.proofDigest)) {
    invalid("Selected PDF vector proof digest is invalid.");
  }
  const replacementBytes = pathReplacement(
    exactGeometry(input.replacement, selected.kind),
    selected.matrix,
  );
  const replacement = writtenShapeGeometry(
    replacementBytes,
    selected.paintOperatorToken,
    selected.matrix,
    selected.kind,
  );
  if (!clipAllows(selected.clip, boundsForGeometry(replacement))) {
    refuse(MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE);
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: selected.operatorSpan,
    replacement: replacementBytes,
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION,
    selection: publicShape(selected),
    replacement,
    request,
    beforeShapeFingerprints: Object.freeze(
      inspection.shapes.map(shapeFingerprint),
    ),
  });
}

export async function prepareMraiNativeVectorGeometryEdit(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorGeometryEditPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (plan.selection.sourceSha256 !== plan.request.identity.sourceSha256) {
    invalid("PDF vector plan source identity is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeVectorGeometryCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorGeometryEditPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorGeometryVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const reopened = await inspectInternal(
    prepared.candidateBytes,
    plan.selection.pageNumber,
  );
  if (reopened.shapes.length !== plan.beforeShapeFingerprints.length) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector shape count changed.",
    );
  }
  const selected = reopened.shapes[plan.selection.pageVectorIndex];
  if (
    !selected ||
    selected.kind !== plan.selection.kind ||
    selected.paintMode !== plan.selection.paintMode ||
    shapeFingerprint(selected) !==
      shapeFingerprint({
        kind: plan.selection.kind,
        paintMode: plan.selection.paintMode,
        geometry: plan.replacement,
      })
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector geometry does not match the replacement.",
    );
  }
  for (let index = 0; index < reopened.shapes.length; index += 1) {
    if (index === plan.selection.pageVectorIndex) continue;
    if (
      shapeFingerprint(reopened.shapes[index]) !==
      plan.beforeShapeFingerprints[index]
    ) {
      throw new PdfNativeTransactionError(
        "verification-failed",
        "Unselected PDF vector geometry changed.",
      );
    }
  }
  return Object.freeze({
    verified: true,
    pageNumber: plan.selection.pageNumber,
    pageVectorIndex: plan.selection.pageVectorIndex,
    candidateSha256: prepared.candidateSha256,
    reopenedShape: publicShape(selected),
  });
}

export async function planMraiNativeVectorSubpathAffineEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorSubpathProof;
  readonly transform: PdfNativeVectorAffineTransform;
}): Promise<PdfNativeVectorSubpathAffineEditPlan> {
  const { inspection, selected } = await selectedSubpathForEdit(input);
  const affine = exactAffineTransform(input.transform);
  return subpathCommandsEditPlan({
    sourceBytes: input.sourceBytes,
    pageNumber: input.pageNumber,
    inspection,
    selected,
    commands: transformSubpathCommands(selected.commands, affine.matrix),
    transform: affine.publicValue,
    unchangedMessage: "PDF vector affine transform does not change the subpath.",
  });
}

/**
 * Rewrite one subpath's operands to the given commands (page user space),
 * keeping every operator in place: a vertex edit that moves one anchor and
 * the control points attached to it. `transform` records the anchor's own
 * translation for the evidence; the commands are the authority.
 */
export async function planMraiNativeVectorSubpathCommandsEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorSubpathProof;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
  readonly transform: PdfNativeVectorAffineTransform;
}): Promise<PdfNativeVectorSubpathAffineEditPlan> {
  const { inspection, selected } = await selectedSubpathForEdit(input);
  if (
    input.commands.length !== selected.commands.length ||
    input.commands.some(
      (command, index) => command.operator !== selected.commands[index].operator,
    )
  ) {
    invalid("PDF vector subpath command edit changes the operator sequence.");
  }
  const affine = exactAffineTransform(input.transform);
  return subpathCommandsEditPlan({
    sourceBytes: input.sourceBytes,
    pageNumber: input.pageNumber,
    inspection,
    selected,
    // The same walk the affine path takes: every point checked finite and in
    // range, the sequence shape checked, the bounds computable.
    commands: transformSubpathCommands(input.commands, IDENTITY),
    transform: affine.publicValue,
    unchangedMessage: "PDF vector command edit does not change the subpath.",
  });
}

async function selectedSubpathForEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorSubpathProof;
}): Promise<
  Readonly<{
    inspection: Awaited<ReturnType<typeof inspectSubpathsInternal>>;
    selected: InternalSubpath;
  }>
> {
  const inspection = await inspectSubpathsInternal(
    input.sourceBytes,
    input.pageNumber,
  );
  const selected = inspection.subpaths.find(
    (subpath) => subpath.id === input.selection.id,
  );
  if (!selected || !sameSubpathAuthority(selected, input.selection)) {
    invalid("Selected PDF vector subpath proof is stale or unavailable.");
  }
  if (!selected.editable) {
    refuse("Selected PDF vector subpath content stream is shared.");
  }
  if (!SHA256_PATTERN.test(selected.proofDigest)) {
    invalid("Selected PDF vector subpath proof digest is invalid.");
  }
  return Object.freeze({ inspection, selected });
}

async function subpathCommandsEditPlan(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly inspection: Awaited<ReturnType<typeof inspectSubpathsInternal>>;
  readonly selected: InternalSubpath;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
  readonly transform: PdfNativeVectorAffineTransform;
  readonly unchangedMessage: string;
}): Promise<PdfNativeVectorSubpathAffineEditPlan> {
  const { selected, inspection } = input;
  const replacement = subpathReplacement(input.commands, selected.matrix);
  const transformedCommands = writtenSubpathCommands(
    replacement,
    selected.paintOperator,
    selected.matrix,
  );
  if (
    JSON.stringify(canonicalSubpathCommands(transformedCommands)) ===
    JSON.stringify(canonicalSubpathCommands(selected.commands))
  ) {
    invalid(input.unchangedMessage);
  }
  if (!clipAllows(selected.clip, boundsForSubpath(transformedCommands))) {
    refuse(MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE);
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: selected.operatorSpan,
    replacement,
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  const existingShapes = await inspectInternal(
    input.sourceBytes,
    input.pageNumber,
  );
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
    selection: publicSubpath(selected),
    transform: input.transform,
    transformedCommands,
    request,
    beforeSubpathFingerprints: Object.freeze(
      inspection.subpaths.map(subpathFingerprint),
    ),
    beforeShapeFingerprints: Object.freeze(
      existingShapes.shapes.map(shapeFingerprint),
    ),
  });
}

export async function prepareMraiNativeVectorSubpathAffineEdit(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorSubpathAffineEditPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (
    plan.schemaVersion !== MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION ||
    plan.selection.sourceSha256 !== plan.request.identity.sourceSha256
  ) {
    invalid("PDF vector subpath plan source identity is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeVectorSubpathCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorSubpathAffineEditPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorSubpathVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const reopened = await inspectSubpathsInternal(
    prepared.candidateBytes,
    plan.selection.pageNumber,
  );
  if (reopened.subpaths.length !== plan.beforeSubpathFingerprints.length) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector subpath count changed.",
    );
  }
  const selected = reopened.subpaths[plan.selection.pageSubpathIndex];
  const expectedFingerprint = subpathFingerprint({
    paintMode: plan.selection.paintMode,
    paintOperator: plan.selection.paintOperator,
    commands: plan.transformedCommands,
    closed: plan.selection.closed,
  });
  if (
    !selected ||
    selected.paintMode !== plan.selection.paintMode ||
    selected.paintOperator !== plan.selection.paintOperator ||
    selected.closed !== plan.selection.closed ||
    selected.commands.map((command) => command.operator).join("/") !==
      plan.selection.commands.map((command) => command.operator).join("/") ||
    subpathFingerprint(selected) !== expectedFingerprint
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector subpath does not match the affine edit.",
    );
  }
  for (let index = 0; index < reopened.subpaths.length; index += 1) {
    if (index === plan.selection.pageSubpathIndex) continue;
    if (
      subpathFingerprint(reopened.subpaths[index]) !==
      plan.beforeSubpathFingerprints[index]
    ) {
      throw new PdfNativeTransactionError(
        "verification-failed",
        "Unselected PDF vector subpath changed.",
      );
    }
  }
  const reopenedShapes = await inspectInternal(
    prepared.candidateBytes,
    plan.selection.pageNumber,
  );
  if (reopenedShapes.shapes.length !== plan.beforeShapeFingerprints.length) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Existing PDF vector shape count changed during subpath editing.",
    );
  }
  for (let index = 0; index < reopenedShapes.shapes.length; index += 1) {
    if (
      shapeFingerprint(reopenedShapes.shapes[index]) !==
      plan.beforeShapeFingerprints[index]
    ) {
      throw new PdfNativeTransactionError(
        "verification-failed",
        "Existing PDF vector shape changed during subpath editing.",
      );
    }
  }
  return Object.freeze({
    verified: true,
    pageNumber: plan.selection.pageNumber,
    pageSubpathIndex: plan.selection.pageSubpathIndex,
    candidateSha256: prepared.candidateSha256,
    reopenedSubpath: publicSubpath(selected),
  });
}












export interface PdfNativeVectorOutlineEditPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
  

  readonly subpathPlan: PdfNativeVectorSubpathAffineEditPlan | null;
  readonly paintOperator: PdfNativeVectorPaintOperator;
  /** The explicit path the object drew, page space. */
  readonly outline: readonly PdfNativeVectorSubpathCommand[];
  readonly transform: PdfNativeVectorAffineTransform;
  /** The explicit path as written and read back, page space. */
  readonly transformedCommands: readonly PdfNativeVectorSubpathCommand[];
  /** The strict shape the written path reads as, or null for a subpath. */
  readonly afterShape: Pick<
    PdfNativeVectorContentShape,
    "kind" | "paintMode" | "geometry"
  > | null;
  readonly request: PdfNativeContentStreamTransactionRequest;
  readonly beforeShapeFingerprints: readonly string[];
  readonly beforeSubpathFingerprints: readonly string[];
}

export interface PdfNativeVectorOutlineVerification {
  readonly verified: true;
  readonly pageNumber: number;
  readonly candidateSha256: string;
  readonly reopened: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
}

function outlinePaintOperator(
  selected: InternalShape | InternalSubpath,
): PdfNativeVectorPaintOperator {
  const token = "paintOperatorToken" in selected
    ? selected.paintOperatorToken
    : selected.paintOperator;
  const operator = paintOperator(token);
  if (!operator) refuse("PDF vector paint operator is unsupported.");
  return operator;
}

/** The explicit path a scanned object draws: its own operator bytes read as
 * one subpath under its own transform (a lone `re` reads as `m l l l h`). */
function outlineOf(
  selected: InternalShape | InternalSubpath,
): readonly PdfNativeVectorSubpathCommand[] {
  if (!("kind" in selected)) return selected.commands;
  const spanBytes = selected.streamInspection.streamBytes.slice(
    selected.operatorSpan.start,
    selected.operatorSpan.end,
  );
  return writtenSubpathCommands(
    spanBytes,
    outlinePaintOperator(selected),
    selected.matrix,
  );
}

async function selectedPathForOutline(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
}): Promise<InternalShape | InternalSubpath> {
  if ("kind" in input.selection) {
    const inspection = await inspectInternal(input.sourceBytes, input.pageNumber);
    const selected = inspection.shapes.find(
      (shape) => shape.id === input.selection.id,
    );
    if (!selected || !sameShapeAuthority(selected, input.selection)) {
      invalid("Selected PDF vector proof is stale or unavailable.");
    }
    if (!selected.editable) {
      refuse("Selected PDF vector content stream is shared.");
    }
    return selected;
  }
  return (
    await selectedSubpathForEdit({
      sourceBytes: input.sourceBytes,
      pageNumber: input.pageNumber,
      selection: input.selection,
    })
  ).selected;
}

/**
 * The explicit path (page space) a selected rectangle, ellipse, line or
 * subpath draws: what its vertices are read from.
 */
export async function readMraiNativeVectorOutline(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
}): Promise<readonly PdfNativeVectorSubpathCommand[]> {
  return outlineOf(await selectedPathForOutline(input));
}




function writtenStrictShape(
  bytes: Uint8Array,
  paintOperatorToken: string,
  matrix: Matrix,
): PdfNativeVectorOutlineEditPlan["afterShape"] {
  const shapes = scanShapes(
    operations(`${latin1(bytes)}\n${paintOperatorToken}`),
    matrix,
  );
  if (shapes.length === 0) return null;
  if (shapes.length !== 1) {
    invalid("PDF vector replacement reads as more than one shape.");
  }
  return Object.freeze({
    kind: shapes[0].kind,
    paintMode: shapes[0].paintMode,
    geometry: shapes[0].geometry,
  });
}

/**
 * Rewrite one path object as `commands` (page space): its explicit path with
 * the same operators (see readMraiNativeVectorOutline), one anchor moved.
 * `transform` records the anchor's own translation for the evidence.
 *
 * `replaceOutline`: the object's path replaced whole by `commands` under the
 * same graphics state, transform and paint operator -- Merge Shapes, whose
 * outline is the merged shapes' (any number of subpaths, a hole drawn the
 * other way round from the outline it is in).
 */
export async function planMraiNativeVectorOutlineEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
  readonly transform: PdfNativeVectorAffineTransform;
  readonly replaceOutline?: boolean;
}): Promise<PdfNativeVectorOutlineEditPlan> {
  const selected = await selectedPathForOutline(input);
  if (!SHA256_PATTERN.test(selected.proofDigest)) {
    invalid("Selected PDF vector proof digest is invalid.");
  }
  const outline = outlineOf(selected);
  if (input.replaceOutline) {
    if (input.commands.length < 3 || input.commands.length > MAX_NATIVE_SUBPATH_COMMANDS) {
      invalid("PDF vector outline replacement is outside the bounded contract.");
    }
  } else if (
    input.commands.length !== outline.length ||
    input.commands.some(
      (command, index) => command.operator !== outline[index].operator,
    )
  ) {
    invalid("PDF vector outline edit changes the operator sequence.");
  }
  // An outline replacement's change is its commands; its affine is none.
  const affine = input.replaceOutline
    ? Object.freeze({ publicValue: Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }), matrix: IDENTITY })
    : exactAffineTransform(input.transform);
  const paint = outlinePaintOperator(selected);
  // The same walk an affine edit takes: every point finite and in range,
  // the sequence a path, the bounds computable.
  const commands = transformSubpathCommands(input.commands, IDENTITY);
  const replacement = subpathReplacement(commands, selected.matrix);
  const afterShape = writtenStrictShape(replacement, paint, selected.matrix);
  if (!("kind" in selected) && !afterShape) {
    const inspection = await inspectSubpathsInternal(
      input.sourceBytes,
      input.pageNumber,
    );
    const subpathPlan = await subpathCommandsEditPlan({
      sourceBytes: input.sourceBytes,
      pageNumber: input.pageNumber,
      inspection,
      selected,
      commands,
      transform: affine.publicValue,
      unchangedMessage: "PDF vector command edit does not change the subpath.",
    });
    return Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
      selection: subpathPlan.selection,
      subpathPlan,
      paintOperator: paint,
      outline,
      transform: subpathPlan.transform,
      transformedCommands: subpathPlan.transformedCommands,
      afterShape: null,
      request: subpathPlan.request,
      beforeShapeFingerprints: subpathPlan.beforeShapeFingerprints,
      beforeSubpathFingerprints: subpathPlan.beforeSubpathFingerprints,
    });
  }
  
  // which a Form written on a clone does not keep.
  if (
    selected.formInstance !== undefined ||
    streamRequestIsolation(selected.streamInspection).formInstanceIsolation
  ) {
    refuse("PDF vector in a Form drawn more than once cannot change its kind.");
  }
  const transformedCommands = writtenSubpathCommands(
    replacement,
    paint,
    selected.matrix,
  );
  if (
    JSON.stringify(canonicalSubpathCommands(transformedCommands)) ===
    JSON.stringify(canonicalSubpathCommands(outline))
  ) {
    invalid("PDF vector command edit does not change the subpath.");
  }
  if (!clipAllows(selected.clip, boundsForSubpath(transformedCommands))) {
    refuse(MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE);
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: selected.operatorSpan,
    replacement,
    untouchedObjects: selected.streamInspection.untouchedObjects,
  });
  const combined = await cachedCombinedPage(input.sourceBytes, input.pageNumber);
  if (combined.subpathError) throw combined.subpathError;
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
    selection: "kind" in selected ? publicShape(selected) : publicSubpath(selected),
    subpathPlan: null,
    paintOperator: paint,
    outline,
    transform: affine.publicValue,
    transformedCommands,
    afterShape,
    request,
    beforeShapeFingerprints: Object.freeze(combined.shapes.map(shapeFingerprint)),
    beforeSubpathFingerprints: Object.freeze(
      combined.subpaths.map(subpathFingerprint),
    ),
  });
}

export async function prepareMraiNativeVectorOutlineEdit(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorOutlineEditPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (plan.subpathPlan) {
    return prepareMraiNativeVectorSubpathAffineEdit(sourceBytes, plan.subpathPlan);
  }
  if (
    plan.schemaVersion !== MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION ||
    plan.selection.sourceSha256 !== plan.request.identity.sourceSha256
  ) {
    invalid("PDF vector outline plan source identity is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

/** `rows` without the one entry equal to `row` (the first such). */
function withoutRow(rows: readonly string[], row: string): readonly string[] | null {
  const index = rows.indexOf(row);
  return index < 0 ? null : [...rows.slice(0, index), ...rows.slice(index + 1)];
}

export async function verifyMraiNativeVectorOutlineCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorOutlineEditPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorOutlineVerification> {
  if (plan.subpathPlan) {
    const verified = await verifyMraiNativeVectorSubpathCandidate(
      sourceBytes,
      plan.subpathPlan,
      prepared,
    );
    return Object.freeze({
      verified: true,
      pageNumber: verified.pageNumber,
      candidateSha256: verified.candidateSha256,
      reopened: verified.reopenedSubpath,
    });
  }
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const failed = (message: string): never => {
    throw new PdfNativeTransactionError("verification-failed", message);
  };
  const pageNumber = plan.selection.pageNumber;
  const reopened = await cachedCombinedPage(prepared.candidateBytes, pageNumber);
  if (reopened.subpathError) throw reopened.subpathError;
  const stream = plan.request.identity;
  const start = plan.request.operatorSpan.start;
  const end = start + plan.request.replacement.length;
  const at = (
    proof: InternalShape | InternalSubpath,
  ) =>
    proof.contentStream.objectNumber === stream.objectNumber &&
    proof.contentStream.generation === stream.generation &&
    proof.operatorSpan.start === start &&
    proof.operatorSpan.end === end &&
    proof.formInstance === undefined;
  const shapes = reopened.shapes.filter(at);
  const subpaths = reopened.subpaths.filter(at);
  if (shapes.length + subpaths.length !== 1) {
    failed("Reopened PDF vector outline edit is not one object.");
  }
  const edited = shapes[0] ?? subpaths[0];
  if (
    outlinePaintOperator(edited) !== plan.paintOperator ||
    JSON.stringify(canonicalSubpathCommands(outlineOf(edited))) !==
      JSON.stringify(canonicalSubpathCommands(plan.transformedCommands))
  ) {
    failed("Reopened PDF vector does not match the outline edit.");
  }
  const editedShapeRow = shapes[0]
    ? shapeFingerprint(shapes[0])
    : null;
  if (
    (plan.afterShape === null) !== (editedShapeRow === null) ||
    (plan.afterShape && editedShapeRow !== shapeFingerprint(plan.afterShape))
  ) {
    failed("Reopened PDF vector kind does not match the outline edit.");
  }
  
  const selectedRow =
    "kind" in plan.selection
      ? shapeFingerprint(plan.selection)
      : subpathFingerprint(plan.selection);
  const beforeShapes =
    "kind" in plan.selection
      ? withoutRow(plan.beforeShapeFingerprints, selectedRow)
      : plan.beforeShapeFingerprints;
  const beforeSubpaths =
    "kind" in plan.selection
      ? plan.beforeSubpathFingerprints
      : withoutRow(plan.beforeSubpathFingerprints, selectedRow);
  const afterShapes = editedShapeRow
    ? withoutRow(reopened.shapes.map(shapeFingerprint), editedShapeRow)
    : reopened.shapes.map(shapeFingerprint);
  const afterSubpaths = editedShapeRow
    ? reopened.subpaths.map(subpathFingerprint)
    : withoutRow(
        reopened.subpaths.map(subpathFingerprint),
        subpathFingerprint(subpaths[0]),
      );
  if (
    !beforeShapes ||
    !beforeSubpaths ||
    !afterShapes ||
    !afterSubpaths ||
    JSON.stringify(beforeShapes) !== JSON.stringify(afterShapes) ||
    JSON.stringify(beforeSubpaths) !== JSON.stringify(afterSubpaths)
  ) {
    failed("An unselected PDF vector changed during the outline edit.");
  }
  return Object.freeze({
    verified: true,
    pageNumber,
    candidateSha256: prepared.candidateSha256,
    reopened: shapes[0] ? publicShape(shapes[0]) : publicSubpath(subpaths[0]),
  });
}

const PDF_WHITESPACE_ONLY = /^[\t\n\f\r ]*$/;
const PDF_WHITESPACE_EDGE = /^[\t\n\f\r ]+|[\t\n\f\r ]+$/g;
/**
 * Operators allowed between two stream-adjacent path objects when they are
 * reordered: pure graphics-state settings that can be replayed verbatim on
 * both sides of the swapped pair without changing any other object.
 */
const REORDER_GAP_OPERATORS = new Set([
  "g", "rg", "k", "G", "RG", "K",
  "cs", "CS", "sc", "SC", "scn", "SCN",
  "w", "d", "J", "j", "M", "gs", "ri", "i",
  // A bare `n` right after a paint is a no-op path end (reportlab emits it
  // after every state change); replaying it on both sides changes nothing.
  "n",
]);

function internalSelectionFor(
  inspection: InternalCombinedPageInspection,
  selection: PdfNativeVectorDeleteSelection,
  label: string,
): InternalShape | InternalSubpath {
  const shapeSelection = "kind" in selection;
  const selected = shapeSelection
    ? inspection.shapes.find((shape) => shape.id === selection.id)
    : inspection.subpaths.find((subpath) => subpath.id === selection.id);
  const authorityMatches =
    selected && shapeSelection && "kind" in selected
      ? sameShapeAuthority(selected, selection)
      : selected && !shapeSelection && !("kind" in selected)
        ? sameSubpathAuthority(selected, selection)
        : false;
  if (!selected || !authorityMatches) {
    invalid(`Selected PDF vector ${label} proof is stale or unavailable.`);
  }
  if (!selected.editable || selected.preserveReason !== null) {
    refuse(`Selected PDF vector cannot be ${label}d in a shared stream.`);
  }
  if (!SHA256_PATTERN.test(selected.proofDigest)) {
    invalid(`Selected PDF vector ${label} proof digest is invalid.`);
  }
  return selected;
}

function objectBlockSpan(
  selected: InternalShape | InternalSubpath,
): Readonly<{ start: number; end: number }> {
  const stream = selected.streamInspection.streamBytes;
  if (
    selected.paintSpan.start < selected.operatorSpan.end ||
    selected.paintSpan.end > stream.length ||
    !PDF_WHITESPACE_ONLY.test(
      latin1(
        stream.slice(selected.operatorSpan.end, selected.paintSpan.start),
      ),
    )
  ) {
    refuse("PDF vector paint operator is not adjacent to its path.");
  }
  return Object.freeze({
    start: selected.operatorSpan.start,
    end: selected.paintSpan.end,
  });
}

function publicSelection(
  selected: InternalShape | InternalSubpath,
): PdfNativeVectorDeleteSelection {
  return "kind" in selected
    ? publicShape(selected)
    : publicSubpath(selected);
}

function selectionFingerprint(
  selected: PdfNativeVectorDeleteSelection,
): string {
  return "kind" in selected
    ? shapeFingerprint(selected)
    : subpathFingerprint(selected);
}

function translatedCopyBytes(
  selected: InternalShape | InternalSubpath,
  offset: Readonly<{ dx: number; dy: number }>,
): Readonly<{ bytes: Uint8Array; fingerprint: string }> {
  const translate: Matrix = [1, 0, 0, 1, offset.dx, offset.dy];
  if ("kind" in selected) {
    const bytes = pathReplacement(
      translatedGeometry(selected.geometry, offset),
      selected.matrix,
    );
    return Object.freeze({
      bytes,
      fingerprint: shapeFingerprint({
        kind: selected.kind,
        paintMode: selected.paintMode,
        geometry: writtenShapeGeometry(
          bytes,
          selected.paintOperatorToken,
          selected.matrix,
          selected.kind,
        ),
      }),
    });
  }
  const bytes = subpathReplacement(
    transformSubpathCommands(selected.commands, translate),
    selected.matrix,
  );
  return Object.freeze({
    bytes,
    fingerprint: subpathFingerprint({
      paintMode: selected.paintMode,
      paintOperator: selected.paintOperator,
      commands: writtenSubpathCommands(
        bytes,
        selected.paintOperator,
        selected.matrix,
      ),
      closed: selected.closed,
    }),
  });
}

function translatedGeometry(
  geometry: PdfNativeVectorGeometry,
  offset: Readonly<{ dx: number; dy: number }>,
): PdfNativeVectorGeometry {
  const move = (value: PdfNativeVectorPoint) =>
    Object.freeze({ x: value.x + offset.dx, y: value.y + offset.dy });
  if (geometry.kind === "line") {
    return exactGeometry({
      kind: "line",
      start: move(geometry.start),
      end: move(geometry.end),
    });
  }
  if (geometry.kind === "rectangle") {
    return exactGeometry({
      kind: "rectangle",
      corners: [
        move(geometry.corners[0]),
        move(geometry.corners[1]),
        move(geometry.corners[2]),
        move(geometry.corners[3]),
      ],
    });
  }
  return exactGeometry({
    kind: "ellipse",
    center: move(geometry.center),
    axisX: geometry.axisX,
    axisY: geometry.axisY,
  });
}

function exactOffset(
  value: Readonly<{ dx: number; dy: number }>,
): Readonly<{ dx: number; dy: number }> {
  if (
    !value ||
    typeof value !== "object" ||
    !Number.isFinite(value.dx) ||
    !Number.isFinite(value.dy) ||
    Math.abs(value.dx) > MAX_ABS_COORDINATE ||
    Math.abs(value.dy) > MAX_ABS_COORDINATE
  ) {
    invalid("PDF vector duplicate offset is invalid.");
  }
  return Object.freeze({ dx: value.dx, dy: value.dy });
}

async function combinedPageFor(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<InternalCombinedPageInspection> {
  const inspection = await cachedCombinedPage(sourceBytes, pageNumber);
  if (inspection.subpathError) throw inspection.subpathError;
  return inspection;
}

export async function planMraiNativeVectorDuplicate(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorDeleteSelection;
  readonly offset: Readonly<{ dx: number; dy: number }>;
}): Promise<PdfNativeVectorDuplicatePlan> {
  const offset = exactOffset(input.offset);
  const inspection = await combinedPageFor(input.sourceBytes, input.pageNumber);
  const selected = internalSelectionFor(inspection, input.selection, "duplicate");
  const block = objectBlockSpan(selected);
  const stream = selected.streamInspection.streamBytes;
  const originalOperatorBytes = Uint8Array.from(stream.slice(block.start, block.end));
  if (
    (await sha256PdfBytes(
      stream.slice(selected.operatorSpan.start, selected.operatorSpan.end),
    )) !== selected.operatorSpan.expectedSha256
  ) {
    invalid("Selected PDF vector duplicate span is inconsistent.");
  }
  const paintBytes = stream.slice(selected.paintSpan.start, selected.paintSpan.end);
  const paintOperator = latin1(paintBytes);
  const copy = translatedCopyBytes(selected, offset);
  if (
    !clipAllows(selected.clip, {
      x: selected.bounds.x + offset.dx,
      y: selected.bounds.y + offset.dy,
      width: selected.bounds.width,
      height: selected.bounds.height,
    })
  ) {
    refuse(MRAI_NATIVE_VECTOR_LEAVES_CLIP_MESSAGE);
  }
  const insertedBytes = concatVectorBytes([
    ascii("\n"),
    copy.bytes,
    ascii("\n"),
    paintBytes,
  ]);
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: Object.freeze({
      start: block.start,
      end: block.end,
      expectedSha256: await sha256PdfBytes(originalOperatorBytes),
    }),
    replacement: concatVectorBytes([originalOperatorBytes, insertedBytes]),
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
    direction: "insert",
    selection: publicSelection(selected),
    offset,
    paintOperator,
    request,
    originalOperatorBytes,
    insertedBytes,
    expectedCopyFingerprint: copy.fingerprint,
    beforeShapeFingerprints: Object.freeze(inspection.shapes.map(shapeFingerprint)),
    beforeSubpathFingerprints: Object.freeze(
      inspection.subpaths.map(subpathFingerprint),
    ),
  });
}

/**
 * Remove a previously inserted duplicate: the copy must still be preceded by
 * exactly the newline the insert wrote, and the removed bytes must be one
 * path object plus its paint operator and nothing else.
 */
export async function planMraiNativeVectorDuplicateRemoval(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorDeleteSelection;
  readonly expectedInsertedBytes: Uint8Array;
}): Promise<PdfNativeVectorDuplicatePlan> {
  const inspection = await combinedPageFor(input.sourceBytes, input.pageNumber);
  const selected = internalSelectionFor(inspection, input.selection, "duplicate");
  const block = objectBlockSpan(selected);
  const stream = selected.streamInspection.streamBytes;
  if (block.start < 1 || stream[block.start - 1] !== 0x0a) {
    invalid("PDF vector duplicate copy is not preceded by its insertion break.");
  }
  const start = block.start - 1;
  const originalOperatorBytes = Uint8Array.from(stream.slice(start, block.end));
  if (
    originalOperatorBytes.length !== input.expectedInsertedBytes.length ||
    originalOperatorBytes.some(
      (value, index) => value !== input.expectedInsertedBytes[index],
    )
  ) {
    invalid("PDF vector duplicate copy bytes changed since insertion.");
  }
  const paintOperator = latin1(
    stream.slice(selected.paintSpan.start, selected.paintSpan.end),
  );
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: Object.freeze({
      start,
      end: block.end,
      expectedSha256: await sha256PdfBytes(originalOperatorBytes),
    }),
    replacement: new Uint8Array(0),
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
    direction: "remove",
    selection: publicSelection(selected),
    offset: Object.freeze({ dx: 0, dy: 0 }),
    paintOperator,
    request,
    originalOperatorBytes,
    insertedBytes: Uint8Array.from(originalOperatorBytes),
    expectedCopyFingerprint: selectionFingerprint(publicSelection(selected)),
    beforeShapeFingerprints: Object.freeze(inspection.shapes.map(shapeFingerprint)),
    beforeSubpathFingerprints: Object.freeze(
      inspection.subpaths.map(subpathFingerprint),
    ),
  });
}

export async function prepareMraiNativeVectorDuplicate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorDuplicatePlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (
    plan.schemaVersion !== MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION ||
    plan.selection.sourceSha256 !== plan.request.identity.sourceSha256 ||
    (plan.direction === "insert"
      ? plan.request.replacement.length !==
        plan.originalOperatorBytes.length + plan.insertedBytes.length
      : plan.request.replacement.length !== 0)
  ) {
    invalid("PDF vector duplicate plan authority is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeVectorDuplicateCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorDuplicatePlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorDuplicateVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const reopened = await combinedPageFor(
    prepared.candidateBytes,
    plan.selection.pageNumber,
  );
  const proofFamily = "kind" in plan.selection ? "shape" : "subpath";
  const index =
    "kind" in plan.selection
      ? plan.selection.pageVectorIndex
      : plan.selection.pageSubpathIndex;
  const expectedShapes = [...plan.beforeShapeFingerprints];
  const expectedSubpaths = [...plan.beforeSubpathFingerprints];
  let copyProofIndex: number;
  if (plan.direction === "insert") {
    copyProofIndex = index + 1;
    (proofFamily === "shape" ? expectedShapes : expectedSubpaths).splice(
      copyProofIndex,
      0,
      plan.expectedCopyFingerprint,
    );
  } else {
    copyProofIndex = index;
    (proofFamily === "shape" ? expectedShapes : expectedSubpaths).splice(
      index,
      1,
    );
  }
  if (
    JSON.stringify(reopened.shapes.map(shapeFingerprint)) !==
      JSON.stringify(expectedShapes) ||
    JSON.stringify(reopened.subpaths.map(subpathFingerprint)) !==
      JSON.stringify(expectedSubpaths)
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector duplicate changed an unselected object.",
    );
  }
  const reopenedCopy =
    plan.direction === "insert"
      ? proofFamily === "shape"
        ? publicShape(reopened.shapes[copyProofIndex])
        : publicSubpath(reopened.subpaths[copyProofIndex])
      : null;
  return Object.freeze({
    verified: true,
    pageNumber: plan.selection.pageNumber,
    proofFamily,
    copyProofIndex,
    candidateSha256: prepared.candidateSha256,
    reopenedCopy,
  });
}

function concatVectorBytes(parts: readonly Uint8Array[]): Uint8Array {
  const length = parts.reduce((total, part) => total + part.length, 0);
  const result = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

function reorderGapBytes(stream: Uint8Array, start: number, end: number): Uint8Array {
  const raw = latin1(stream.slice(start, end));
  if (!PDF_WHITESPACE_ONLY.test(raw)) {
    let parsed: readonly ContentOperation[];
    try {
      parsed = operations(raw);
    } catch {
      refuse("PDF vector neighbours are separated by unsupported content.");
    }
    if (
      parsed.length === 0 ||
      parsed.some(
        (operation) =>
          !REORDER_GAP_OPERATORS.has(operation.operator) ||
          operation.operands.some((operand) => operand.kind === "operator"),
      )
    ) {
      refuse("PDF vector neighbours are separated by unsupported content.");
    }
  }
  return ascii(raw.replace(PDF_WHITESPACE_EDGE, ""));
}

function reorderedPairBytes(
  first: Uint8Array,
  gap: Uint8Array,
  second: Uint8Array,
): Uint8Array {
  // No state between the two objects: swap them outright, so the pair stays
  // free of wrappers and can be reordered again.
  if (gap.length === 0) {
    return concatVectorBytes([second, ascii("\n"), first]);
  }
  return concatVectorBytes([
    ascii("q\n"),
    gap,
    ascii("\n"),
    second,
    ascii("\nQ\n"),
    first,
    ascii("\n"),
    gap,
  ]);
}

/**
 * Widen an object's block to the `q…Q` group that owns it, when that group
 * holds only graphics state plus this one object (what a style edit or an
 * earlier reorder writes). Such a block is self-contained, so reordering it
 * carries its state along instead of stranding it.
 */
function wrappedObjectBlock(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): Readonly<{ start: number; end: number }> {
  // A block wrapped more than once (a style edit, then a move past content
  // that carried state) is widened group by group, to the outermost one
  // that holds only state and the object.
  let current = block;
  for (;;) {
    const wider = wrappedObjectBlockOnce(parsedOperations, current);
    if (wider.start === current.start && wider.end === current.end) return current;
    current = wider;
  }
}

function wrappedObjectBlockOnce(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): Readonly<{ start: number; end: number }> {
  const firstIndex = parsedOperations.findIndex(
    (operation) => operation.start === block.start,
  );
  // The block ends with its paint, or with the `Q` of a group already
  // widened to.
  const paintIndex = parsedOperations.findIndex(
    (operation) =>
      operation.end === block.end &&
      (paintMode(operation.operator) !== null || operation.operator === "Q"),
  );
  if (firstIndex < 0 || paintIndex < firstIndex) return block;
  let index = firstIndex - 1;
  while (
    index >= 0 &&
    REORDER_GAP_OPERATORS.has(parsedOperations[index].operator)
  ) {
    index -= 1;
  }
  const open = parsedOperations[index];
  const close = parsedOperations[paintIndex + 1];
  if (
    !open ||
    open.operator !== "q" ||
    open.operands.length !== 0 ||
    !close ||
    close.operator !== "Q" ||
    close.operands.length !== 0
  ) {
    return block;
  }
  return Object.freeze({ start: open.start, end: close.end });
}

function streamOrderedObjects(
  inspection: InternalCombinedPageInspection,
): readonly (InternalShape | InternalSubpath)[] {
  return [...inspection.shapes, ...inspection.subpaths].sort(
    (left, right) => left.operatorSpan.start - right.operatorSpan.start,
  );
}

function swappedFingerprints(
  inspection: InternalCombinedPageInspection,
  first: InternalShape | InternalSubpath,
  second: InternalShape | InternalSubpath,
): Readonly<{ shapes: readonly string[]; subpaths: readonly string[] }> {
  const shapes = inspection.shapes.map(shapeFingerprint);
  const subpaths = inspection.subpaths.map(subpathFingerprint);
  if ("kind" in first && "kind" in second) {
    const a = first.pageVectorIndex;
    const b = second.pageVectorIndex;
    [shapes[a], shapes[b]] = [shapes[b], shapes[a]];
  } else if (!("kind" in first) && !("kind" in second)) {
    const a = first.pageSubpathIndex;
    const b = second.pageSubpathIndex;
    [subpaths[a], subpaths[b]] = [subpaths[b], subpaths[a]];
  }
  return Object.freeze({
    shapes: Object.freeze(shapes),
    subpaths: Object.freeze(subpaths),
  });
}

/** Whether a gap holds only replayable state (no unsupported content). */
function reorderGapIsReplayable(stream: Uint8Array, start: number, end: number): boolean {
  try {
    reorderGapBytes(stream, start, end);
    return true;
  } catch {
    return false;
  }
}

// The graphics state a path paint reads, by the operators that set it. The
// two colour keys hold the colour space and the colour as a pair: a space
// resets its colour to the initial one, so `cs` alone is a state of its own.
const REORDER_STATE_KEY_BY_OPERATOR: ReadonlyMap<string, string> = new Map([
  ["g", "fill"], ["rg", "fill"], ["k", "fill"], ["cs", "fill"], ["sc", "fill"], ["scn", "fill"],
  ["G", "stroke"], ["RG", "stroke"], ["K", "stroke"], ["CS", "stroke"], ["SC", "stroke"], ["SCN", "stroke"],
  ["w", "w"], ["d", "d"], ["J", "J"], ["j", "j"], ["M", "M"],
  ["gs", "gs"], ["ri", "ri"], ["i", "i"],
]);
const REORDER_STATE_KEY_ORDER = ["fill", "stroke", "w", "d", "J", "j", "M"] as const;
const REORDER_STATE_DEFAULTS: ReadonlyMap<string, string> = new Map([
  ["fill", "0 g"], ["stroke", "0 G"], ["w", "1 w"], ["d", "[] 0 d"], ["J", "0 J"], ["j", "0 j"], ["M", "10 M"],
]);
const REORDER_TEXT_UNIT_FORBIDDEN = new Set(["q", "Q", "cm", "W", "W*", "Do", "sh", "BI", "BT", "d0", "d1"]);
const PATH_CONSTRUCTION_OPERATORS = new Set(["m", "l", "c", "v", "y", "h", "re"]);

interface ReorderStateInForce {
  /** Per key: the operations in force, as written (`cs` then colour). */
  readonly keys: ReadonlyMap<string, readonly string[]>;
  /** The last `Tr` in force (text render mode), -1 when none. */
  readonly renderMode: number;
}

/**
 * The state operators in force at one offset, read through the `q`/`Q`
 * stack: for each key the operation text that set it (a colour as its
 * space then its colour), the way a later paint would read it.
 */
function reorderStateInForceAt(
  parsedOperations: readonly ContentOperation[],
  streamText: string,
  offset: number,
): ReorderStateInForce {
  let current = new Map<string, readonly string[]>();
  let renderMode = -1;
  const stack: { keys: Map<string, readonly string[]>; renderMode: number }[] = [];
  for (const operation of parsedOperations) {
    if (operation.start >= offset) break;
    const text = streamText.slice(operation.start, operation.end);
    if (operation.operator === "q") {
      stack.push({ keys: current, renderMode });
      current = new Map(current);
    } else if (operation.operator === "Q") {
      const saved = stack.pop();
      if (saved) {
        current = saved.keys;
        renderMode = saved.renderMode;
      }
    } else if (operation.operator === "Tr") {
      renderMode = Number(operation.operands[0]?.value ?? -1);
    } else {
      const key = REORDER_STATE_KEY_BY_OPERATOR.get(operation.operator);
      if (!key) continue;
      if (key === "fill" || key === "stroke") {
        // A colour is its space then its components: `cs` starts a space (and
        // resets the colour), `sc`/`scn` set components in that space, and
        // `g`/`rg`/`k` name both at once.
        const previous = current.get(key) ?? [];
        const operator = operation.operator;
        if (operator === "cs" || operator === "CS") {
          current.set(key, [text]);
        } else if (operator === "sc" || operator === "scn" || operator === "SC" || operator === "SCN") {
          const space = previous.length > 0 && /\s(cs|CS)$/.test(previous[0]) ? [previous[0]] : [];
          current.set(key, [...space, text]);
        } else {
          current.set(key, [text]);
        }
      } else {
        current.set(key, [text]);
      }
    }
  }
  return Object.freeze({ keys: current, renderMode });
}

/** The state keys the operations between two offsets set. */
function reorderDisturbedKeys(
  parsedOperations: readonly ContentOperation[],
  start: number,
  end: number,
): ReadonlySet<string> {
  const keys = new Set<string>();
  for (const operation of parsedOperations) {
    if (operation.start < start) continue;
    if (operation.start >= end) break;
    const key = REORDER_STATE_KEY_BY_OPERATOR.get(operation.operator);
    if (key) keys.add(key);
  }
  return keys;
}

/**
 * The one unit of content after (`forward`) or before (`backward`) a block,
 * past the replayable gap: a text object, a balanced group, an XObject or a
 * shading. Null when the stream ends there; a refusal when what follows is
 * content the object cannot pass (a clip, a transform, marked content, a
 * text object that clips, an unbalanced group).
 */
function reorderPassedUnit(
  parsedOperations: readonly ContentOperation[],
  streamText: string,
  block: Readonly<{ start: number; end: number }>,
  forward: boolean,
): PdfNativeVectorPassedUnit | null {
  const unsupported = () => refuse("PDF vector neighbours are separated by unsupported content.");
  const index = forward
    ? parsedOperations.findIndex((operation) => operation.start >= block.end)
    : (() => {
        let last = -1;
        for (let i = 0; i < parsedOperations.length; i += 1) {
          if (parsedOperations[i].end <= block.start) last = i;
          else break;
        }
        return last;
      })();
  let at = index;
  while (at >= 0 && at < parsedOperations.length &&
    REORDER_GAP_OPERATORS.has(parsedOperations[at].operator)) {
    at += forward ? 1 : -1;
  }
  if (at < 0 || at >= parsedOperations.length) return null;
  const edge = parsedOperations[at];
  const unit = (kind: PdfNativeVectorPassedUnitKind, start: number, end: number) =>
    Object.freeze({ kind, start, end });
  if (edge.operator === "Do") return unit("xobject", edge.start, edge.end);
  if (edge.operator === "sh") return unit("shading", edge.start, edge.end);
  // One path object: its construction operators up to the paint that ends
  // it. A path that ends without painting (`n`, a clip) is not passed.
  if (forward && (edge.operator === "m" || edge.operator === "re")) {
    let i = at + 1;
    while (i < parsedOperations.length && PATH_CONSTRUCTION_OPERATORS.has(parsedOperations[i].operator)) i += 1;
    const paint = parsedOperations[i];
    if (!paint || paintMode(paint.operator) === null) unsupported();
    return unit("path", edge.start, paint.end);
  }
  if (!forward && paintMode(edge.operator) !== null) {
    let i = at - 1;
    while (i >= 0 && PATH_CONSTRUCTION_OPERATORS.has(parsedOperations[i].operator)) i -= 1;
    const first = parsedOperations[i + 1];
    if (!first || i + 1 === at) unsupported();
    return unit("path", first.start, edge.end);
  }
  if (forward ? edge.operator === "BT" : edge.operator === "ET") {
    // One text object; marked content inside it must close inside it, and
    // it must not clip (a render mode of 4 or more adds to the clip).
    let depthMarked = 0;
    let renderMode = reorderStateInForceAt(parsedOperations, streamText, forward ? edge.start : 0).renderMode;
    let i = at;
    for (;;) {
      i += forward ? 1 : -1;
      if (i < 0 || i >= parsedOperations.length) unsupported();
      const operation = parsedOperations[i];
      if (forward ? operation.operator === "ET" : operation.operator === "BT") break;
      if (REORDER_TEXT_UNIT_FORBIDDEN.has(operation.operator) || operation.operator === "ET" || operation.operator === "BT") unsupported();
      if (operation.operator === "BMC" || operation.operator === "BDC") depthMarked += forward ? 1 : -1;
      if (operation.operator === "EMC") depthMarked += forward ? -1 : 1;
      if (depthMarked < 0) unsupported();
      if (operation.operator === "Tr") renderMode = Number(operation.operands[0]?.value ?? -1);
    }
    if (depthMarked !== 0) unsupported();
    const start = forward ? edge.start : parsedOperations[i].start;
    const end = forward ? parsedOperations[i].end : edge.end;
    if (!forward) {
      // Read the modes the object sets, in stream order, after the one in
      // force before it.
      renderMode = reorderStateInForceAt(parsedOperations, streamText, start).renderMode;
      for (const operation of parsedOperations) {
        if (operation.start < start || operation.start >= end) continue;
        if (operation.operator === "Tr") renderMode = Number(operation.operands[0]?.value ?? -1);
      }
    }
    if (renderMode >= 4) unsupported();
    return unit("text", start, end);
  }
  if (forward ? edge.operator === "q" : edge.operator === "Q") {
    let depth = 0;
    let depthMarked = 0;
    let i = at;
    for (;;) {
      const operation = parsedOperations[i];
      if (operation.operator === (forward ? "q" : "Q")) depth += 1;
      else if (operation.operator === (forward ? "Q" : "q")) depth -= 1;
      else if (operation.operator === "BMC" || operation.operator === "BDC") depthMarked += forward ? 1 : -1;
      else if (operation.operator === "EMC") depthMarked += forward ? -1 : 1;
      if (depthMarked < 0) unsupported();
      if (depth === 0) break;
      i += forward ? 1 : -1;
      if (i < 0 || i >= parsedOperations.length) unsupported();
    }
    if (depthMarked !== 0) unsupported();
    return forward
      ? unit("group", edge.start, parsedOperations[i].end)
      : unit("group", parsedOperations[i].start, edge.end);
  }
  return unsupported();
}

/** The operations, as written, that keep a path's state past the content
 * it passes: one per key that content sets, the object's own value or the
 * PDF default. Empty when nothing it passes touches what it paints with. */
function reorderCarriedState(
  parsedOperations: readonly ContentOperation[],
  streamText: string,
  blockStart: number,
  disturbed: ReadonlySet<string>,
): string {
  if (disturbed.size === 0) return "";
  if (disturbed.has("ri") || disturbed.has("i")) {
    refuse("PDF vector neighbours are separated by unsupported content.");
  }
  const inForce = reorderStateInForceAt(parsedOperations, streamText, blockStart);
  const lines: string[] = [];
  for (const key of REORDER_STATE_KEY_ORDER) {
    if (!disturbed.has(key)) continue;
    const own = inForce.keys.get(key);
    lines.push(...(own && own.length > 0 ? own : [REORDER_STATE_DEFAULTS.get(key)!]));
  }
  return lines.join("\n");
}

/**
 * The object's block with the carried state in front of it: inside its own
 * `q`…`Q` when the block is one already (the state it names there still
 * wins, as it did), else in a new one. Nothing nests, so a shape moved many
 * times keeps one group.
 */
function wrappedWithState(block: Uint8Array, carried: string): Uint8Array {
  if (!carried) return block;
  const text = latin1(block);
  let operationsOfBlock: readonly ContentOperation[] | null = null;
  try {
    operationsOfBlock = operations(text);
  } catch {
    operationsOfBlock = null;
  }
  const first = operationsOfBlock?.[0];
  const last = operationsOfBlock?.[operationsOfBlock.length - 1];
  if (
    operationsOfBlock &&
    first &&
    last &&
    first.operator === "q" &&
    first.operands.length === 0 &&
    first.start === 0 &&
    last.operator === "Q" &&
    last.operands.length === 0 &&
    last.end === block.length &&
    groupIsBalancedWhole(operationsOfBlock)
  ) {
    const inner = text.slice(first.end, last.start).replace(PDF_WHITESPACE_EDGE, "");
    return concatVectorBytes([ascii(`q\n${carried}\n`), latin1Bytes(inner), ascii("\nQ")]);
  }
  return concatVectorBytes([ascii(`q\n${carried}\n`), block, ascii("\nQ")]);
}

/** Whether the first `q` of these operations is closed only by the last. */
function groupIsBalancedWhole(operationsOfBlock: readonly ContentOperation[]): boolean {
  let depth = 0;
  for (let index = 0; index < operationsOfBlock.length; index += 1) {
    const operator = operationsOfBlock[index].operator;
    if (operator === "q") depth += 1;
    else if (operator === "Q") depth -= 1;
    if (depth === 0 && index < operationsOfBlock.length - 1) return false;
  }
  return depth === 0;
}

/**
 * The object moved past one unit: forward `[A][gap][X]` becomes
 * `[gap][X]\n[q state A Q]`, backward `[X][gap][A]` becomes
 * `[q state A Q]\n[X][\ngap]`. X keeps the state it was drawn in and the
 * content after the pair still sees what X and the gap leave.
 */
function passedUnitBytes(
  stream: Uint8Array,
  block: Readonly<{ start: number; end: number }>,
  gap: Uint8Array,
  unit: PdfNativeVectorPassedUnit,
  carried: string,
  forward: boolean,
): Uint8Array {
  const object = wrappedWithState(stream.slice(block.start, block.end), carried);
  const content = stream.slice(unit.start, unit.end);
  return forward
    ? concatVectorBytes([
        ...(gap.length > 0 ? [gap, ascii("\n")] : []),
        content,
        ascii("\n"),
        object,
      ])
    : concatVectorBytes([
        object,
        ascii("\n"),
        content,
        ...(gap.length > 0 ? [ascii("\n"), gap] : []),
      ]);
}

/** The fingerprints once the selected object has moved past `passedCount`
 * objects of its own family, and where it then stands. */
function movedFingerprints(
  inspection: InternalCombinedPageInspection,
  selected: InternalShape | InternalSubpath,
  passedCount: number,
  forward: boolean,
): Readonly<{ shapes: readonly string[]; subpaths: readonly string[]; index: number }> {
  const shapes = inspection.shapes.map(shapeFingerprint);
  const subpaths = inspection.subpaths.map(subpathFingerprint);
  const family = "kind" in selected ? shapes : subpaths;
  const from = "kind" in selected ? selected.pageVectorIndex : selected.pageSubpathIndex;
  const to = forward ? from + passedCount : from - passedCount;
  const [picked] = family.splice(from, 1);
  family.splice(to, 0, picked);
  return Object.freeze({
    shapes: Object.freeze(shapes),
    subpaths: Object.freeze(subpaths),
    index: to,
  });
}

/** How many objects of the selected object's family a span holds. */
function familyObjectsWithin(
  inspection: InternalCombinedPageInspection,
  selected: InternalShape | InternalSubpath,
  start: number,
  end: number,
): number {
  const family: readonly (InternalShape | InternalSubpath)[] =
    "kind" in selected ? inspection.shapes : inspection.subpaths;
  return family.filter(
    (entry) =>
      entry !== selected &&
      entry.streamInspection === selected.streamInspection &&
      entry.operatorSpan.start >= start &&
      entry.paintSpan.end <= end,
  ).length;
}

// ---------------------------------------------------------------------------
// Layer moves past any content. Office's Bring Forward and Send Backward move
// an object one place in the drawing order, past the next object whatever it
// is, and Bring to Front and Send to Back put it last or first (ISO/IEC
// 29500-1 19.3.1.45: z-order is document order). On a page the next object is
// the next thing that marks it -- a path paint, a text show, an XObject, a
// shading, an inline image -- with the groups and marked-content sections it
// is drawn in. The moved object leaves its place and is written on the far
// side of that object inside its own `q`…`Q`, which re-establishes exactly
// the state it was drawn in: the transforms and clips of the groups it leaves
// (replayed as written), the transform the content it passes adds (undone),
// and the colours, line state and graphics states that content changes
// (named again). Every other byte keeps its order and the state it is drawn
// in, and a tagged page keeps its structure: the moved object's own sections
// go with it, everything else stays where it was.
// ---------------------------------------------------------------------------

export type PdfNativeLayerMoveMode = "forward" | "backward" | "front" | "back";

/** The refusal a layer move raises for content it cannot carry the object
 * past: what the window says as "the next object clips or restyles it". */
const LAYER_MOVE_REFUSAL = "PDF vector neighbours are separated by unsupported content.";
export const MRAI_NATIVE_LAYER_MOVE_OTHER_STREAM =
  "PDF vector cannot pass another of its page's content streams.";

const LAYER_TEXT_SHOW_OPERATORS = new Set(["Tj", "TJ", "'", '"']);

/** Whether an operator puts a mark on the page: what a layer move passes. */
function layerMarks(operator: string): boolean {
  return (
    paintMode(operator) !== null ||
    operator === "Do" ||
    operator === "sh" ||
    operator === "BI" ||
    LAYER_TEXT_SHOW_OPERATORS.has(operator)
  );
}

interface LayerNesting {
  save: number;
  marked: number;
  text: number;
  path: boolean;
}

function layerNest(nesting: LayerNesting, operator: string): void {
  if (operator === "q") nesting.save += 1;
  else if (operator === "Q") nesting.save -= 1;
  else if (operator === "BMC" || operator === "BDC") nesting.marked += 1;
  else if (operator === "EMC") nesting.marked -= 1;
  else if (operator === "BT") nesting.text += 1;
  else if (operator === "ET") nesting.text -= 1;
  else if (PATH_CONSTRUCTION_OPERATORS.has(operator)) nesting.path = true;
  else if (paintMode(operator) !== null || operator === "n") nesting.path = false;
}

/** The index of the first operation at or after one offset. */
function layerIndexAt(parsedOperations: readonly ContentOperation[], offset: number): number {
  const index = parsedOperations.findIndex((operation) => operation.start >= offset);
  return index < 0 ? parsedOperations.length : index;
}

/**
 * Where a layer move takes a block, or null when nothing marks the page on
 * that side of it (the object already stands at that end).
 *
 * After the block (forward, front): past the first mark after it (forward)
 * or the last (front), at the first place outside any text object and path
 * under construction where every group and section opened since the block
 * has closed again -- the object then sits beside what it passed, never
 * inside it. Before the block (backward, back): before the last mark ahead
 * of it (backward) or the first (back), at the latest such place from which
 * nothing open there closes before the block.
 */
function layerMoveDestination(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
  mode: PdfNativeLayerMoveMode,
): number | null {
  if (mode === "forward" || mode === "front") {
    const first = layerIndexAt(parsedOperations, block.end);
    let target = -1;
    for (let index = first; index < parsedOperations.length; index += 1) {
      if (!layerMarks(parsedOperations[index].operator)) continue;
      target = index;
      if (mode === "forward") break;
    }
    if (target < 0) return null;
    const nesting: LayerNesting = { save: 0, marked: 0, text: 0, path: false };
    let lowestSave = 0;
    let lowestMarked = 0;
    for (let index = first; index < parsedOperations.length; index += 1) {
      layerNest(nesting, parsedOperations[index].operator);
      if (nesting.text < 0) refuse("PDF page content stream nesting is unbalanced.");
      lowestSave = Math.min(lowestSave, nesting.save);
      lowestMarked = Math.min(lowestMarked, nesting.marked);
      if (
        index >= target &&
        nesting.text === 0 &&
        !nesting.path &&
        nesting.save === lowestSave &&
        nesting.marked === lowestMarked
      ) {
        return parsedOperations[index].end;
      }
    }
    refuse("PDF page content stream nesting is unbalanced.");
  }
  const limit = layerIndexAt(parsedOperations, block.start);
  let target = -1;
  for (let index = limit - 1; index >= 0; index -= 1) {
    if (!layerMarks(parsedOperations[index].operator)) continue;
    target = index;
    if (mode === "backward") break;
  }
  if (target < 0) return null;
  // Place k stands before operation k; place 0 is the drawing's start.
  const places: { position: number; save: number; marked: number; valid: boolean }[] = [
    { position: parsedOperations[0]?.start ?? 0, save: 0, marked: 0, valid: true },
  ];
  const nesting: LayerNesting = { save: 0, marked: 0, text: 0, path: false };
  for (let index = 0; index < limit; index += 1) {
    layerNest(nesting, parsedOperations[index].operator);
    places.push({
      position: parsedOperations[index].end,
      save: nesting.save,
      marked: nesting.marked,
      valid: nesting.text === 0 && !nesting.path,
    });
  }
  let lowestSave = places[limit].save;
  let lowestMarked = places[limit].marked;
  for (let index = limit; index >= 0; index -= 1) {
    lowestSave = Math.min(lowestSave, places[index].save);
    lowestMarked = Math.min(lowestMarked, places[index].marked);
    if (
      index <= target &&
      places[index].valid &&
      places[index].save === lowestSave &&
      places[index].marked === lowestMarked
    ) {
      return places[index].position;
    }
  }
  refuse("PDF page content stream nesting is unbalanced.");
}

/** One piece of graphics state in force: a transform, a clip (its path and
 * `W n` as written), text drawn as a clip, or one state operator. */
interface LayerStateEntry {
  readonly kind: "cm" | "clip" | "text-clip" | "state";
  /** The state key a "state" entry sets (REORDER_STATE_KEY_BY_OPERATOR). */
  readonly key: string | null;
  readonly operation: ContentOperation | null;
  /** Where it was written; a replay writes `text`. */
  readonly start: number;
  readonly text: string;
}

interface LayerWalk {
  /** Each open `q` group's state entries, the drawing's own frame first. */
  readonly frames: readonly (readonly LayerStateEntry[])[];
  /** The marked-content sections open, outermost first. */
  readonly sections: readonly ContentOperation[];
}

/**
 * The graphics state the operations before `endIndex` leave in force, by the
 * group that applied it, skipping the operations inside `skip` (a block that
 * has left its place). Text state is read only for the render mode, which
 * makes a text object a clip.
 */
function layerWalk(
  parsedOperations: readonly ContentOperation[],
  streamText: string,
  endIndex: number,
  skip: Readonly<{ start: number; end: number }> | null = null,
): LayerWalk {
  const frames: LayerStateEntry[][] = [[]];
  const renderModes: number[] = [];
  const sections: ContentOperation[] = [];
  let renderMode = 0;
  let inText = false;
  let textClips = false;
  let path: ContentOperation[] = [];
  let clip: ContentOperation | null = null;
  for (let index = 0; index < endIndex; index += 1) {
    const operation = parsedOperations[index];
    if (skip && operation.start >= skip.start && operation.end <= skip.end) continue;
    const operator = operation.operator;
    const top = frames[frames.length - 1];
    if (operator === "q") {
      frames.push([]);
      renderModes.push(renderMode);
    } else if (operator === "Q") {
      if (frames.length < 2) refuse("PDF page content stream nesting is unbalanced.");
      frames.pop();
      renderMode = renderModes.pop() ?? 0;
    } else if (operator === "BMC" || operator === "BDC") {
      sections.push(operation);
    } else if (operator === "EMC") {
      if (sections.length === 0) refuse("PDF page content stream nesting is unbalanced.");
      sections.pop();
    } else if (operator === "BT") {
      inText = true;
      textClips = false;
    } else if (operator === "ET") {
      if (textClips) {
        top.push({ kind: "text-clip", key: null, operation, start: operation.start, text: "" });
      }
      inText = false;
      textClips = false;
    } else if (operator === "Tr") {
      renderMode = Number(operation.operands[0]?.value ?? -1);
    } else if (LAYER_TEXT_SHOW_OPERATORS.has(operator)) {
      if (inText && renderMode >= 4) textClips = true;
    } else if (PATH_CONSTRUCTION_OPERATORS.has(operator)) {
      path.push(operation);
    } else if (operator === "W" || operator === "W*") {
      clip = operation;
    } else if (operator === "n" || paintMode(operator) !== null) {
      if (clip && path.length > 0) {
        // A clip ended by a paint is replayed as the same clip ended by `n`.
        const text =
          operator === "n"
            ? streamText.slice(path[0].start, operation.end)
            : `${streamText.slice(path[0].start, clip.end)}\nn`;
        top.push({ kind: "clip", key: null, operation: null, start: path[0].start, text });
      }
      path = [];
      clip = null;
    } else if (operator === "cm") {
      top.push({
        kind: "cm",
        key: null,
        operation,
        start: operation.start,
        text: streamText.slice(operation.start, operation.end),
      });
    } else {
      const key = REORDER_STATE_KEY_BY_OPERATOR.get(operator);
      if (key) {
        top.push({
          kind: "state",
          key,
          operation,
          start: operation.start,
          text: streamText.slice(operation.start, operation.end),
        });
      }
    }
  }
  return Object.freeze({ frames, sections });
}

/** The state keys a paint reads, beyond its transform and clip. */
const LAYER_FILL_KEYS = ["fill"] as const;
const LAYER_STROKE_KEYS = ["stroke", "w", "d", "J", "j", "M"] as const;
const LAYER_RESET_KEY_ORDER = ["fill", "stroke", "w", "d", "J", "j", "M", "ri", "i"] as const;
const LAYER_RESET_DEFAULTS: ReadonlyMap<string, string> = new Map([
  ...REORDER_STATE_DEFAULTS,
  ["ri", "/RelativeColorimetric ri"],
  ["i", "1 i"],
]);
/** The ExtGState key that sets what each operator key sets. */
const LAYER_GRAPHICS_STATE_KEY: ReadonlyMap<string, string> = new Map([
  ["w", "LW"], ["d", "D"], ["J", "LC"], ["j", "LJ"], ["M", "ML"], ["ri", "RI"], ["i", "FL"],
]);

/** The keys a path painted with this operator reads. */
export function layerPathKeys(paintOperatorText: string): ReadonlySet<string> {
  const mode = paintMode(paintOperatorText);
  if (mode === null) refuse("PDF vector paint operator is not addressable.");
  return new Set<string>([
    ...(mode === "stroke" ? [] : LAYER_FILL_KEYS),
    ...(mode === "fill" ? [] : LAYER_STROKE_KEYS),
    "ri",
    "i",
  ]);
}

/**
 * Widen a moved block to everything that exists only for it: the `q…Q`
 * group that holds only its state, and the marked-content section written
 * right around it, as many times as they nest.
 */
function layerObjectBlock(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): Readonly<{ start: number; end: number }> {
  let current = block;
  for (;;) {
    const firstIndex = parsedOperations.findIndex((operation) => operation.start === current.start);
    const lastIndex = parsedOperations.findIndex((operation) => operation.end === current.end);
    if (firstIndex < 0 || lastIndex < firstIndex) return current;
    const after = parsedOperations[lastIndex + 1];
    let index = firstIndex - 1;
    while (index >= 0 && REORDER_GAP_OPERATORS.has(parsedOperations[index].operator)) index -= 1;
    const open = parsedOperations[index];
    if (
      open &&
      open.operator === "q" &&
      open.operands.length === 0 &&
      after &&
      after.operator === "Q" &&
      after.operands.length === 0
    ) {
      current = Object.freeze({ start: open.start, end: after.end });
      continue;
    }
    const before = parsedOperations[firstIndex - 1];
    if (
      before &&
      (before.operator === "BMC" || before.operator === "BDC") &&
      after &&
      after.operator === "EMC" &&
      after.operands.length === 0
    ) {
      current = Object.freeze({ start: before.start, end: after.end });
      continue;
    }
    return current;
  }
}

/** The state keys a block sets itself before its first mark. */
function layerOwnKeys(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): ReadonlySet<string> {
  const keys = new Set<string>();
  for (const operation of parsedOperations) {
    if (operation.start < block.start || operation.end > block.end) continue;
    if (layerMarks(operation.operator)) break;
    const key = REORDER_STATE_KEY_BY_OPERATOR.get(operation.operator);
    if (key) keys.add(key);
  }
  return keys;
}

/** The `gs` operations a block applies itself before its first mark,
 * through the groups it opens. */
function layerOwnGraphicsStates(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): readonly ContentOperation[] {
  const frames: ContentOperation[][] = [[]];
  for (const operation of parsedOperations) {
    if (operation.start < block.start || operation.end > block.end) continue;
    if (layerMarks(operation.operator)) break;
    if (operation.operator === "q") frames.push([]);
    else if (operation.operator === "Q") {
      if (frames.length > 1) frames.pop();
    } else if (operation.operator === "gs") {
      frames[frames.length - 1].push(operation);
    }
  }
  return frames.flat();
}

/** The `EMC` that closes a section. */
function layerSectionClose(
  parsedOperations: readonly ContentOperation[],
  open: ContentOperation,
): ContentOperation {
  let depth = 0;
  for (const operation of parsedOperations) {
    if (operation.start < open.start) continue;
    if (operation.operator === "BMC" || operation.operator === "BDC") depth += 1;
    else if (operation.operator === "EMC") {
      depth -= 1;
      if (depth === 0) return operation;
    }
  }
  refuse("PDF page content stream nesting is unbalanced.");
}

export interface PdfNativeLayerMove {
  readonly spanStart: number;
  readonly spanEnd: number;
  /** What the object passed, in stream offsets. */
  readonly passedStart: number;
  readonly passedEnd: number;
  readonly replacement: Uint8Array;
}

/**
 * One layer move of `block` within one drawing (see the section comment), or
 * null when nothing marks the page on that side of it. `keys` are the state
 * keys the object paints with; `paintStates` and `namedProperties` are the
 * page's ExtGState and `/Properties` names (empty for a Form's own drawing).
 */
export function layerMove(input: {
  readonly parsedOperations: readonly ContentOperation[];
  readonly stream: Uint8Array;
  readonly block: Readonly<{ start: number; end: number }>;
  readonly mode: PdfNativeLayerMoveMode;
  readonly keys: ReadonlySet<string>;
  readonly paintStates: ReadonlyMap<string, PdfNativePaintState | null>;
  readonly namedProperties: ReadonlyMap<string, string | null>;
}): PdfNativeLayerMove | null {
  const { parsedOperations, stream, block, mode } = input;
  const streamText = latin1(stream);
  if (!balancedContentBlock(contentNesting(parsedOperations, 0, stream.length))) {
    refuse("PDF page content stream nesting is unbalanced.");
  }
  if (!balancedContentBlock(contentNesting(parsedOperations, block.start, block.end))) {
    refuse("PDF vector is not painted by a self-contained block.");
  }
  const destination = layerMoveDestination(parsedOperations, block, mode);
  if (destination === null) return null;
  const after = mode === "forward" || mode === "front";
  const atBlock = layerWalk(parsedOperations, streamText, layerIndexAt(parsedOperations, block.start));
  const atDestination = layerWalk(
    parsedOperations,
    streamText,
    layerIndexAt(parsedOperations, destination),
    after ? block : null,
  );
  const depth = atDestination.frames.length - 1;
  const markedDepth = atDestination.sections.length;
  if (
    depth > atBlock.frames.length - 1 ||
    markedDepth > atBlock.sections.length ||
    atDestination.sections.some((section, index) => section !== atBlock.sections[index])
  ) {
    refuse("PDF page content stream nesting is unbalanced.");
  }
  // What the object takes with it: the state of the groups it leaves, and
  // for a backward move the state its own group gained after the place it
  // lands at.
  const carried: LayerStateEntry[] = [
    ...(after ? [] : atBlock.frames[depth].filter((entry) => entry.start >= destination)),
    ...atBlock.frames.slice(depth + 1).flat(),
  ];
  if (carried.some((entry) => entry.kind === "text-clip")) refuse(LAYER_MOVE_REFUSAL);
  // What the content it passes leaves in force where it lands (forward).
  const passedState = after
    ? atDestination.frames[depth].filter((entry) => entry.start >= block.end)
    : [];
  if (passedState.some((entry) => entry.kind === "clip" || entry.kind === "text-clip")) {
    refuse(LAYER_MOVE_REFUSAL);
  }
  // The transform the passed content adds, undone.
  let shift: Matrix = IDENTITY;
  let shifted = false;
  for (const entry of passedState) {
    if (entry.kind !== "cm" || !entry.operation) continue;
    const [a, b, c, d, e, f] = numberOperands(entry.operation, 6);
    shift = multiply(shift, [a, b, c, d, e, f]);
    shifted = true;
  }
  let compensation: readonly string[] | null = null;
  if (shifted) {
    const [a, b, c, d, e, f] = shift;
    const determinant = a * d - b * c;
    if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
      refuse(LAYER_MOVE_REFUSAL);
    }
    const inverseShift = [
      d / determinant,
      -b / determinant,
      -c / determinant,
      a / determinant,
      (c * f - d * e) / determinant,
      (b * e - a * f) / determinant,
    ].map(formatImageMoveCompensationNumber);
    if (inverseShift.join(" ") !== "1 0 0 1 0 0") compensation = inverseShift;
  }
  // The keys the passed content sets that the object paints with and does
  // not set itself, named again at the values it was drawn with.
  const passedKeys = new Set(
    passedState.flatMap((entry) => (entry.kind === "state" && entry.key ? [entry.key] : [])),
  );
  // A key the object's own block sets, or a group it leaves sets again in
  // its wrapper, is in force at the value it was drawn with already.
  const ownKeys = layerOwnKeys(parsedOperations, block);
  const carriedKeys = new Set(
    carried.flatMap((entry) => (entry.kind === "state" && entry.key ? [entry.key] : [])),
  );
  const resetKeys = LAYER_RESET_KEY_ORDER.filter(
    (key) =>
      input.keys.has(key) && passedKeys.has(key) && !ownKeys.has(key) && !carriedKeys.has(key),
  );
  if (resetKeys.some((key) => LAYER_GRAPHICS_STATE_KEY.has(key))) {
    // A graphics state may set the same thing; the value in force is then
    // whichever came last, which an operator named again cannot reproduce.
    const applied = [...atBlock.frames.flat(), ...passedState].filter(
      (entry) => entry.kind === "state" && entry.key === "gs",
    );
    for (const entry of applied) {
      const name = entry.operation?.operands[0];
      const state =
        name?.kind === "name" ? input.paintStates.get(name.value.replace(/^\//, "")) : undefined;
      if (
        !state ||
        resetKeys.some((key) => {
          const graphicsKey = LAYER_GRAPHICS_STATE_KEY.get(key);
          return graphicsKey !== undefined && state.has(graphicsKey);
        })
      ) {
        refuse(LAYER_MOVE_REFUSAL);
      }
    }
  }
  const inForce = reorderStateInForceAt(parsedOperations, streamText, block.start);
  const resetLines = resetKeys.flatMap((key) => {
    const own = inForce.keys.get(key);
    return own && own.length > 0 ? own : [LAYER_RESET_DEFAULTS.get(key)!];
  });
  // A graphics state the passed content applies is passed only where the
  // object then paints as it did, naming the states it was drawn under again
  // when that makes it so.
  let carriedStates: readonly ContentOperation[] = [];
  if (passedState.some((entry) => entry.kind === "state" && entry.key === "gs")) {
    const gsOf = (entries: readonly LayerStateEntry[]) =>
      entries.flatMap((entry) =>
        entry.kind === "state" && entry.key === "gs" && entry.operation ? [entry.operation] : [],
      );
    const own = layerOwnGraphicsStates(parsedOperations, block);
    const drawnIn = paintStateAt(parsedOperations, block.start, input.paintStates);
    const landsIn = paintStateAt(parsedOperations, destination, input.paintStates);
    const target = drawnIn === null ? null : applyPaintStates(drawnIn, own, input.paintStates);
    const ambient = gsOf(atBlock.frames.slice(0, depth + 1).flat());
    const reapplied = gsOf(carried);
    let proven = false;
    for (
      let tail = 0;
      landsIn !== null &&
      tail <= Math.min(MRAI_NATIVE_EMBEDDED_IMAGE_CARRIED_PAINT_STATE_LIMIT, ambient.length);
      tail += 1
    ) {
      const candidate = ambient.slice(ambient.length - tail);
      if (
        samePaintState(
          applyPaintStates(landsIn, [...candidate, ...reapplied, ...own], input.paintStates),
          target,
        )
      ) {
        carriedStates = candidate;
        proven = true;
        break;
      }
    }
    if (!proven) refuse(LAYER_MOVE_REFUSAL);
  }
  // The sections the object leaves: copied around it where a copy says
  // nothing the page does not already say, otherwise taken along -- a tagged
  // leaf that marks only this object.
  const leaving = atBlock.sections.slice(markedDepth);
  const emptied: { start: number; end: number }[] = [];
  for (const section of leaving) {
    if (liftableMarkedContent(section, input.namedProperties)) continue;
    if (!movableMarkedContent(section)) refuse(LAYER_MOVE_REFUSAL);
    const close = layerSectionClose(parsedOperations, section);
    if (
      parsedOperations.some(
        (operation) =>
          operation.start >= section.end &&
          operation.end <= close.start &&
          layerMarks(operation.operator) &&
          !(operation.start >= block.start && operation.end <= block.end),
      ) ||
      (after ? close.end > destination : close.start < block.end)
    ) {
      refuse(LAYER_MOVE_REFUSAL);
    }
    emptied.push(
      { start: section.start, end: section.end },
      { start: close.start, end: close.end },
    );
  }
  emptied.sort((left, right) => left.start - right.start);
  const innerLines = [
    ...resetLines,
    ...carriedStates.map((operation) => streamText.slice(operation.start, operation.end)),
    ...carried.map((entry) => entry.text),
  ];
  const grouped = compensation !== null || innerLines.length > 0;
  const head =
    leaving.map((section) => `${streamText.slice(section.start, section.end)}\n`).join("") +
    (grouped
      ? `q${compensation ? ` ${compensation.join(" ")} cm` : ""}\n${innerLines
          .map((line) => `${line}\n`)
          .join("")}`
      : "");
  const tail = `${grouped ? "\nQ" : ""}${"\nEMC".repeat(leaving.length)}`;
  const spanStart = after
    ? Math.min(block.start, ...emptied.map((range) => range.start))
    : destination;
  const spanEnd = after ? destination : Math.max(block.end, ...emptied.map((range) => range.end));
  const kept: Uint8Array[] = [];
  let at = spanStart;
  for (const range of [...emptied, { start: block.start, end: block.end }].sort(
    (left, right) => left.start - right.start,
  )) {
    if (range.start > at) kept.push(stream.slice(at, range.start));
    at = Math.max(at, range.end);
  }
  if (spanEnd > at) kept.push(stream.slice(at, spanEnd));
  const moved = concatVectorBytes([
    latin1Bytes(head),
    stream.slice(block.start, block.end),
    ascii(tail),
  ]);
  const rest = concatVectorBytes(kept);
  return Object.freeze({
    spanStart,
    spanEnd,
    passedStart: after ? block.end : destination,
    passedEnd: after ? destination : block.start,
    // A place before the object is the end of an operator, which may run
    // straight into the next token: the moved bytes start on a line of
    // their own.
    replacement: after
      ? concatVectorBytes([rest, ascii("\n"), moved])
      : concatVectorBytes([ascii("\n"), moved, ascii("\n"), rest]),
  });
}

/** Whether a stretch of a drawing marks the page at all. */
function layerRangeMarks(
  parsedOperations: readonly ContentOperation[],
  range: Readonly<{ start: number; end: number }>,
): boolean {
  return parsedOperations.some(
    (operation) =>
      operation.start >= range.start &&
      operation.end <= range.end &&
      layerMarks(operation.operator),
  );
}

/** Whether another of the page's streams marks the page on that side of the
 * selected one: a move cannot pass content in another stream. */
async function layerOtherStreamsMark(
  sourceBytes: Uint8Array,
  contentStreams: readonly Readonly<{ objectNumber: number; generation: number }>[],
  index: number,
  after: boolean,
): Promise<boolean> {
  const others = after ? contentStreams.slice(index + 1) : contentStreams.slice(0, index);
  for (const other of others) {
    const inspected = await inspectMraiNativeObject(sourceBytes, other);
    if (!inspected.streamBytes) return true;
    try {
      if (operations(latin1(inspected.streamBytes), true).some((operation) => layerMarks(operation.operator))) {
        return true;
      }
    } catch {
      return true;
    }
  }
  return false;
}

/** Refusals of the one-unit planner that a layer move may still carry the
 * object past: content it cannot read as one unit, an inline image, a path
 * neighbour it would have to swap with. */
const UNIT_REORDER_FALLBACK_MESSAGES = new Set([
  LAYER_MOVE_REFUSAL,
  "PDF vector neighbour lives in a shared stream.",
  "PDF page content stream cannot be reordered.",
]);

export async function planMraiNativeVectorReorder(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorDeleteSelection;
  readonly direction: PdfNativeVectorReorderDirection;
}): Promise<PdfNativeVectorReorderPlan> {
  if (
    input.direction !== "forward" &&
    input.direction !== "backward" &&
    input.direction !== "front" &&
    input.direction !== "back"
  ) {
    invalid("PDF vector reorder direction is invalid.");
  }
  const inspection = await combinedPageFor(input.sourceBytes, input.pageNumber);
  const selected = internalSelectionFor(inspection, input.selection, "reorder");
  // One step past a path neighbour or one unit of other content keeps the
  // form every earlier reorder wrote; anything else -- a tagged paragraph, a
  // group that leaves the object's own group, an empty text object that
  // marks nothing -- is a layer move, and so are the moves to either end.
  if (input.direction === "forward" || input.direction === "backward") {
    let unit: PdfNativeVectorReorderPlan | null;
    try {
      unit = await planUnitVectorReorder(input, inspection, selected);
    } catch (error) {
      if (
        !(error instanceof PdfNativeTransactionError) ||
        !UNIT_REORDER_FALLBACK_MESSAGES.has(error.message)
      ) {
        throw error;
      }
      unit = null;
    }
    if (unit) return unit;
  }
  return planLayerVectorReorder(input, inspection, selected);
}

/** One step written as a pair swap or a pass of one unit, or null when the
 * step is a layer move instead (nothing the unit grammar can pass, or a unit
 * that marks nothing). */
async function planUnitVectorReorder(
  input: Readonly<{
    sourceBytes: Uint8Array;
    pageNumber: number;
    direction: PdfNativeVectorReorderDirection;
  }>,
  inspection: InternalCombinedPageInspection,
  selected: InternalShape | InternalSubpath,
): Promise<PdfNativeVectorReorderPlan | null> {
  const forward = input.direction === "forward";
  const ordered = streamOrderedObjects(inspection);
  const position = ordered.findIndex((entry) => entry === selected);
  const candidate = forward ? ordered[position + 1] : ordered[position - 1];
  let parsedOperations: readonly ContentOperation[];
  try {
    parsedOperations = operations(latin1(selected.streamInspection.streamBytes));
  } catch {
    refuse("PDF page content stream cannot be reordered.");
  }
  const stream = selected.streamInspection.streamBytes;
  const streamText = latin1(stream);
  const selectedBlock = wrappedObjectBlock(parsedOperations, objectBlockSpan(selected));
  const beforeShapeFingerprints = Object.freeze(inspection.shapes.map(shapeFingerprint));
  const beforeSubpathFingerprints = Object.freeze(inspection.subpaths.map(subpathFingerprint));
  const requestFor = async (
    span: Readonly<{ start: number; end: number }>,
    reorderedBytes: Uint8Array,
  ) => {
    const originalOperatorBytes = Uint8Array.from(stream.slice(span.start, span.end));
    const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
      identity: selected.streamInspection.identity,
      expectedObjectSha256: selected.streamInspection.objectSha256,
      expectedStreamSha256: selected.streamInspection.streamSha256,
      operatorSpan: Object.freeze({
        start: span.start,
        end: span.end,
        expectedSha256: await sha256PdfBytes(originalOperatorBytes),
      }),
      replacement: reorderedBytes,
      untouchedObjects: selected.streamInspection.untouchedObjects,
      ...streamRequestIsolation(selected.streamInspection),
    });
    return { request, originalOperatorBytes };
  };
  // The next object in stream order is a path object with only state between
  // the two: swap the pair (the form every earlier reorder wrote).
  const neighbour =
    candidate && candidate.streamInspection === selected.streamInspection
      ? candidate
      : null;
  if (neighbour) {
    const neighbourBlock = wrappedObjectBlock(parsedOperations, objectBlockSpan(neighbour));
    const firstBlock = forward ? selectedBlock : neighbourBlock;
    const secondBlock = forward ? neighbourBlock : selectedBlock;
    if (firstBlock.end > secondBlock.start) {
      invalid("PDF vector neighbours overlap.");
    }
    if (reorderGapIsReplayable(stream, firstBlock.end, secondBlock.start)) {
      if (!neighbour.editable || neighbour.preserveReason !== null) {
        refuse("PDF vector neighbour lives in a shared stream.");
      }
      const first = forward ? selected : neighbour;
      const second = forward ? neighbour : selected;
      const gap = reorderGapBytes(stream, firstBlock.end, secondBlock.start);
      const reorderedBytes = reorderedPairBytes(
        stream.slice(firstBlock.start, firstBlock.end),
        gap,
        stream.slice(secondBlock.start, secondBlock.end),
      );
      const expected = swappedFingerprints(inspection, first, second);
      const { request, originalOperatorBytes } = await requestFor(
        { start: firstBlock.start, end: secondBlock.end },
        reorderedBytes,
      );
      const sameFamily = ("kind" in selected) === ("kind" in neighbour);
      return Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
        direction: input.direction,
        selection: publicSelection(selected),
        neighbour: publicSelection(neighbour),
        passed: null,
        request,
        originalOperatorBytes,
        reorderedBytes,
        beforeShapeFingerprints,
        beforeSubpathFingerprints,
        expectedShapeFingerprints: expected.shapes,
        expectedSubpathFingerprints: expected.subpaths,
        expectedSelectionIndex: sameFamily
          ? "kind" in neighbour
            ? neighbour.pageVectorIndex
            : neighbour.pageSubpathIndex
          : "kind" in selected
            ? selected.pageVectorIndex
            : selected.pageSubpathIndex,
      });
    }
  }
  // Otherwise the object passes one unit of other content: a text object, a
  // group, a picture or a shading, carrying the state that content changes.
  const unit = reorderPassedUnit(parsedOperations, streamText, selectedBlock, forward);
  // Nothing more in the stream, or a unit that marks nothing (an empty text
  // object, a group of state alone): the layer move finds the next object.
  if (!unit || !layerRangeMarks(parsedOperations, unit)) return null;
  const gapStart = forward ? selectedBlock.end : unit.end;
  const gapEnd = forward ? unit.start : selectedBlock.start;
  const gap = reorderGapBytes(stream, gapStart, gapEnd);
  const disturbed = new Set<string>([
    ...reorderDisturbedKeys(parsedOperations, gapStart, gapEnd),
    ...(unit.kind === "text" ? reorderDisturbedKeys(parsedOperations, unit.start, unit.end) : []),
  ]);
  if (disturbed.has("gs")) {
    // A named state sets what no operator can set back (alpha, blend,
    // a soft mask): the object may pass it only where it paints the same.
    // A Form's own names are not the page's, so inside a Form no name is
    // readable and the move is refused (the server reads it the same way).
    const states = streamRequestIsolation(selected.streamInspection).formInstanceIsolation
      ? new Map<string, PdfNativePaintState | null>()
      : await pagePaintStates(input.sourceBytes, input.pageNumber);
    const drawnIn = paintStateAt(parsedOperations, selectedBlock.start, states);
    const landsIn = paintStateAt(parsedOperations, forward ? unit.end : unit.start, states);
    if (!samePaintState(drawnIn, landsIn)) {
      refuse("PDF vector neighbours are separated by unsupported content.");
    }
    disturbed.delete("gs");
  }
  const carried = reorderCarriedState(parsedOperations, streamText, selectedBlock.start, disturbed);
  const reorderedBytes = passedUnitBytes(stream, selectedBlock, gap, unit, carried, forward);
  const span = forward
    ? { start: selectedBlock.start, end: unit.end }
    : { start: unit.start, end: selectedBlock.end };
  const passedCount = familyObjectsWithin(inspection, selected, unit.start, unit.end);
  const expected = movedFingerprints(inspection, selected, passedCount, forward);
  const { request, originalOperatorBytes } = await requestFor(span, reorderedBytes);
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
    direction: input.direction,
    selection: publicSelection(selected),
    neighbour: null,
    passed: unit,
    request,
    originalOperatorBytes,
    reorderedBytes,
    beforeShapeFingerprints,
    beforeSubpathFingerprints,
    expectedShapeFingerprints: expected.shapes,
    expectedSubpathFingerprints: expected.subpaths,
    expectedSelectionIndex: expected.index,
  });
}

/** A layer move of a path object: forward or backward past the next object
 * whatever it is, or to the front or back of its drawing in one write. */
async function planLayerVectorReorder(
  input: Readonly<{
    sourceBytes: Uint8Array;
    pageNumber: number;
    direction: PdfNativeVectorReorderDirection;
  }>,
  inspection: InternalCombinedPageInspection,
  selected: InternalShape | InternalSubpath,
): Promise<PdfNativeVectorReorderPlan> {
  const mode = input.direction;
  const after = mode === "forward" || mode === "front";
  const stream = selected.streamInspection.streamBytes;
  let parsedOperations: readonly ContentOperation[];
  try {
    parsedOperations = operations(latin1(stream), true);
  } catch {
    refuse("PDF page content stream cannot be reordered.");
  }
  const block = layerObjectBlock(parsedOperations, objectBlockSpan(selected));
  // A Form's drawing names its own resources, which the page's say nothing
  // about (the server reads it the same way).
  const graph = await inspectMraiNativePageContentGraph(input.sourceBytes, input.pageNumber);
  const pageStreamIndex = graph.contentStreams.findIndex(
    (entry) =>
      entry.objectNumber === selected.streamInspection.identity.objectNumber &&
      entry.generation === selected.streamInspection.identity.generation,
  );
  const formDrawing = pageStreamIndex < 0;
  const move = layerMove({
    parsedOperations,
    stream,
    block,
    mode,
    keys: layerPathKeys(latin1(stream.slice(selected.paintSpan.start, selected.paintSpan.end))),
    paintStates: formDrawing
      ? new Map<string, PdfNativePaintState | null>()
      : await pagePaintStates(input.sourceBytes, input.pageNumber),
    namedProperties: formDrawing
      ? EMPTY_NAMED_PROPERTIES
      : await pageMarkedContentProperties(input.sourceBytes, input.pageNumber),
  });
  // Another of the page's streams that marks the page on that side is
  // content the object cannot pass: a step that reaches it, or a move to an
  // end that would stop short of it, is refused rather than half made.
  if (
    (!move || mode === "front" || mode === "back") &&
    !formDrawing &&
    (await layerOtherStreamsMark(input.sourceBytes, graph.contentStreams, pageStreamIndex, after))
  ) {
    refuse(MRAI_NATIVE_LAYER_MOVE_OTHER_STREAM);
  }
  if (!move) {
    refuse(
      after
        ? "PDF vector is already the topmost path on its page."
        : "PDF vector is already the bottommost path on its page.",
    );
  }
  const originalOperatorBytes = Uint8Array.from(stream.slice(move.spanStart, move.spanEnd));
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: Object.freeze({
      start: move.spanStart,
      end: move.spanEnd,
      expectedSha256: await sha256PdfBytes(originalOperatorBytes),
    }),
    replacement: move.replacement,
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  const passedCount = familyObjectsWithin(inspection, selected, move.passedStart, move.passedEnd);
  const expected = movedFingerprints(inspection, selected, passedCount, after);
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
    direction: mode,
    selection: publicSelection(selected),
    neighbour: null,
    passed: Object.freeze({ kind: "content" as const, start: move.passedStart, end: move.passedEnd }),
    request,
    originalOperatorBytes,
    reorderedBytes: move.replacement,
    beforeShapeFingerprints: Object.freeze(inspection.shapes.map(shapeFingerprint)),
    beforeSubpathFingerprints: Object.freeze(inspection.subpaths.map(subpathFingerprint)),
    expectedShapeFingerprints: expected.shapes,
    expectedSubpathFingerprints: expected.subpaths,
    expectedSelectionIndex: expected.index,
  });
}

/**
 * Put a reordered pair back: the exact reordered bytes must still occur once
 * in the stream (any later edit to either object makes the history stale).
 */
export async function planMraiNativeVectorReorderRestore(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly selection: PdfNativeVectorDeleteSelection;
  readonly neighbour: PdfNativeVectorDeleteSelection | null;
  readonly reorderedBytes: Uint8Array;
  readonly originalBytes: Uint8Array;
}): Promise<PdfNativeVectorReorderPlan> {
  const inspection = await combinedPageFor(input.sourceBytes, input.pageNumber);
  const selected = internalSelectionFor(inspection, input.selection, "reorder");
  const neighbour = input.neighbour
    ? internalSelectionFor(inspection, input.neighbour, "reorder")
    : null;
  const stream = selected.streamInspection.streamBytes;
  const haystack = latin1(stream);
  const needle = latin1(input.reorderedBytes);
  const start = haystack.indexOf(needle);
  if (start < 0 || haystack.indexOf(needle, start + 1) >= 0) {
    invalid("PDF vector reorder history bytes are not unique in the stream.");
  }
  const end = start + input.reorderedBytes.length;
  const inside = (entry: InternalShape | InternalSubpath) =>
    entry.operatorSpan.start >= start && entry.paintSpan.end <= end;
  if (!inside(selected) || (neighbour && !inside(neighbour))) {
    invalid("PDF vector reorder history no longer covers both objects.");
  }
  // Putting the bytes back moves the object past the same objects of its
  // family the reorder moved it past, the other way: it is drawn last of
  // the span after a forward move and first after a backward one.
  const passedCount = familyObjectsWithin(inspection, selected, start, end);
  const movedForward = familyObjectsWithin(inspection, selected, start, selected.operatorSpan.start) === passedCount;
  const expected = movedFingerprints(inspection, selected, passedCount, !movedForward);
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: selected.streamInspection.identity,
    expectedObjectSha256: selected.streamInspection.objectSha256,
    expectedStreamSha256: selected.streamInspection.streamSha256,
    operatorSpan: Object.freeze({
      start,
      end,
      expectedSha256: await sha256PdfBytes(input.reorderedBytes),
    }),
    replacement: Uint8Array.from(input.originalBytes),
    untouchedObjects: selected.streamInspection.untouchedObjects,
    ...streamRequestIsolation(selected.streamInspection),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
    direction: "restore",
    selection: publicSelection(selected),
    neighbour: neighbour ? publicSelection(neighbour) : null,
    passed: null,
    request,
    originalOperatorBytes: Uint8Array.from(input.reorderedBytes),
    reorderedBytes: Uint8Array.from(input.originalBytes),
    beforeShapeFingerprints: Object.freeze(inspection.shapes.map(shapeFingerprint)),
    beforeSubpathFingerprints: Object.freeze(
      inspection.subpaths.map(subpathFingerprint),
    ),
    expectedShapeFingerprints: expected.shapes,
    expectedSubpathFingerprints: expected.subpaths,
    expectedSelectionIndex: expected.index,
  });
}

export async function prepareMraiNativeVectorReorder(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorReorderPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (
    plan.schemaVersion !== MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION ||
    plan.selection.sourceSha256 !== plan.request.identity.sourceSha256 ||
    plan.request.replacement.length !== plan.reorderedBytes.length
  ) {
    invalid("PDF vector reorder plan authority is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

/** Whether two fingerprints name the same object, numbers compared to well
 * under anything a page draws (a compensating transform reads back within a
 * rounding of the original). */
function reorderFingerprintsAgree(actual: string, expected: string): boolean {
  if (actual === expected) return true;
  const agree = (left: unknown, right: unknown): boolean => {
    if (typeof left === "number" && typeof right === "number") {
      return Math.abs(left - right) <= 1e-5 * Math.max(1, Math.abs(right));
    }
    if (Array.isArray(left) && Array.isArray(right)) {
      return left.length === right.length && left.every((value, index) => agree(value, right[index]));
    }
    if (left && right && typeof left === "object" && typeof right === "object") {
      const leftKeys = Object.keys(left);
      const rightKeys = Object.keys(right);
      return (
        leftKeys.length === rightKeys.length &&
        leftKeys.every(
          (key) =>
            Object.prototype.hasOwnProperty.call(right, key) &&
            agree((left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key]),
        )
      );
    }
    return left === right;
  };
  try {
    return agree(JSON.parse(actual), JSON.parse(expected));
  } catch {
    return false;
  }
}

export async function verifyMraiNativeVectorReorderCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeVectorReorderPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<PdfNativeVectorReorderVerification> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const reopened = await combinedPageFor(
    prepared.candidateBytes,
    plan.selection.pageNumber,
  );
  // Every other object reads back exactly as it was; the moved object is
  // read through the transform a layer move undoes for it, which reads back
  // to within a rounding of where it was.
  const selectionIsShape = "kind" in plan.selection;
  const sameOrder = (
    actual: readonly string[],
    expected: readonly string[],
    selectionFamily: boolean,
  ) =>
    actual.length === expected.length &&
    actual.every(
      (value, index) =>
        value === expected[index] ||
        (selectionFamily &&
          index === plan.expectedSelectionIndex &&
          reorderFingerprintsAgree(value, expected[index])),
    );
  if (
    !sameOrder(reopened.shapes.map(shapeFingerprint), plan.expectedShapeFingerprints, selectionIsShape) ||
    !sameOrder(
      reopened.subpaths.map(subpathFingerprint),
      plan.expectedSubpathFingerprints,
      !selectionIsShape,
    )
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector order does not match the requested swap.",
    );
  }
  const selectionFingerprintValue = selectionFingerprint(plan.selection);
  const familyEntries: readonly (InternalShape | InternalSubpath)[] =
    "kind" in plan.selection ? reopened.shapes : reopened.subpaths;
  const reopenedSelection = familyEntries[plan.expectedSelectionIndex];
  if (
    !reopenedSelection ||
    !reorderFingerprintsAgree(
      selectionFingerprint(publicSelection(reopenedSelection)),
      selectionFingerprintValue,
    )
  ) {
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF vector selection is missing after reordering.",
    );
  }
  return Object.freeze({
    verified: true,
    pageNumber: plan.selection.pageNumber,
    candidateSha256: prepared.candidateSha256,
    reopenedSelection: publicSelection(reopenedSelection),
  });
}

// ---------------------------------------------------------------------------
// Image z-order. A picture is moved to the front or the back of its page by
// rotating one contiguous span of the page content stream around the block
// that paints it, so the file keeps every other byte it had.
// ---------------------------------------------------------------------------


export type PdfNativeEmbeddedImageReorderDirection =
  | "front"
  | "back"
  | "forward"
  | "backward"
  | "restore";



export interface PdfNativeEmbeddedImageReorderPlan {
  readonly schemaVersion: typeof MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION;
  readonly direction: PdfNativeEmbeddedImageReorderDirection;
  /** True for a one-step move (and the restore of one): the span reaches
   * past one unit of content, not to the drawing's end. */
  readonly stepped: boolean;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly imageOrderAfter: number;
  readonly resourceName: string;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly resourceOwner: Readonly<{ objectNumber: number; generation: number }>;
  readonly image: Readonly<{ objectNumber: number; generation: number }>;
  readonly contentStream: Readonly<{ objectNumber: number; generation: number }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  /** Byte length of the moved block inside the un-rotated span. */
  readonly blockLength: number;
  /** True when the moved block leads the un-rotated span (a front move). */
  readonly blockLeads: boolean;
  /** The transform the moved block is wrapped in so it keeps its place when
   * the drawing it passes moves the page with top-level `cm`s; absent when
   * it passes none. */
  readonly compensation?: readonly number[];
  /** The graphics states the moved block names again inside its wrapper so
   * the picture paints exactly as it was drawn, each as the drawing wrote
   * it; absent when the drawing it passes leaves its paint alone. */
  readonly carriedStates?: readonly string[];
  /** The groups the moved block left, and carries the state of in its
   * wrapper; absent when it stays inside every group it was drawn in. */
  readonly lift?: PdfNativeEmbeddedImageMoveLift;
  /** The places a marked-content section that travels with the picture was
   * written in, which the move empties: `BDC` and `EMC` ranges in the
   * un-rotated span, in the order they are written there. */
  readonly emptiedMarkers?: readonly Readonly<{ start: number; end: number }>[];
  /** Bytes at the span's outer edge the rotation leaves where they are: the
   * state a travelling section's `BDC` was written among (a front move), or
   * what stands between the block and the `EMC` that closes it (a back one).
   * Zero for a span that is nothing but the block and what it passes. */
  readonly stationaryBytes?: number;
  readonly request: PdfNativeContentStreamTransactionRequest;
}

const MARKED_CONTENT_OPEN_OPERATORS = new Set(["BDC", "BMC"]);
const PDF_PATH_CONSTRUCTION_OPERATORS = new Set(["m", "l", "c", "v", "y", "h", "re"]);

/** Operators that end an object, so a `q…Q` group holding one is not ours. */
const IMAGE_BLOCK_FOREIGN_OPERATORS = new Set(["Do", "BT", "ET", "sh", "BI", "ID", "EI"]);
/** Ambient state a moved picture would inherit differently at its destination:
 * a transform, a clip, and a graphics state it cannot name again. */
const IMAGE_MOVE_GEOMETRY_OPERATORS = new Set(["cm", "W", "W*", "gs"]);
const IMAGE_MOVE_COLOR_OPERATORS = new Set([
  "g", "rg", "k", "G", "RG", "K",
  "cs", "CS", "sc", "SC", "scn", "SCN",
]);

/**
 * Widen a bare `/Name Do` to the `q…Q` group that owns it. The group carries
 * the picture's own `cm` and clip, so moving it moves the picture whole. When
 * the invocation is not wrapped the bare span stays the block, and the checks
 * below then refuse the move rather than strand what it depends on.
 */
function imageObjectBlock(
  parsedOperations: readonly ContentOperation[],
  paint: Readonly<{ start: number; end: number }>,
): Readonly<{ start: number; end: number }> {
  const bare = Object.freeze({ start: paint.start, end: paint.end });
  const paintIndex = parsedOperations.findIndex(
    (operation) => operation.operator === "Do" && operation.end === paint.end,
  );
  if (paintIndex < 0) refuse("PDF image paint operator is not addressable.");
  const close = parsedOperations[paintIndex + 1];
  if (!close || close.operator !== "Q" || close.operands.length !== 0) {
    return bare;
  }
  let depth = 0;
  for (let index = paintIndex - 1; index >= 0; index -= 1) {
    const operation = parsedOperations[index];
    if (operation.operator === "Q") {
      depth += 1;
      continue;
    }
    if (operation.operator === "q") {
      if (depth > 0) {
        depth -= 1;
        continue;
      }
      return operation.operands.length === 0
        ? widenImageBlockThroughOwnGroups(parsedOperations, index, paintIndex + 1)
        : bare;
    }
    if (depth > 0) continue;
    if (
      IMAGE_BLOCK_FOREIGN_OPERATORS.has(operation.operator) ||
      paintMode(operation.operator) !== null
    ) {
      return bare;
    }
  }
  return bare;
}

/** Operators that only set up what the next paint is drawn with: its
 * transform, clip path, graphics state and colours. */
const IMAGE_BLOCK_SETUP_OPERATORS = new Set([
  "cm", "gs", "ri", "i", "w", "J", "j", "M", "d",
  "g", "rg", "k", "G", "RG", "K", "cs", "CS", "sc", "SC", "scn", "SCN",
  "m", "l", "c", "v", "y", "h", "re", "W", "W*", "n",
]);

/**
 * A picture's `q…Q` group may itself sit in a group that exists only for it:
 * pdfTeX clips an included figure with `q <path> W n … q <cm> /Im Do Q Q`.
 * That outer group is the picture's too -- its clip belongs to the picture --
 * so the block widens to it whenever everything between the two groups only
 * sets up the paint (a transform, a clip, colours) and nothing follows the
 * inner group before the outer one closes.
 */
function widenImageBlockThroughOwnGroups(
  parsedOperations: readonly ContentOperation[],
  openIndex: number,
  closeIndex: number,
): Readonly<{ start: number; end: number }> {
  let open = openIndex;
  let close = closeIndex;
  for (;;) {
    let parent = -1;
    for (let index = open - 1; index >= 0; index -= 1) {
      const operator = parsedOperations[index].operator;
      if (operator === "q" && parsedOperations[index].operands.length === 0) {
        parent = index;
        break;
      }
      if (!IMAGE_BLOCK_SETUP_OPERATORS.has(operator)) break;
    }
    const parentClose = parsedOperations[close + 1];
    if (
      parent < 0 ||
      !parentClose ||
      parentClose.operator !== "Q" ||
      parentClose.operands.length !== 0
    ) {
      break;
    }
    // A clip path is set up and ended, never left open into the picture.
    let pathOpen = false;
    for (let index = parent + 1; index < open; index += 1) {
      const operator = parsedOperations[index].operator;
      if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) pathOpen = true;
      else if (operator === "n") pathOpen = false;
    }
    if (pathOpen) break;
    open = parent;
    close += 1;
  }
  return Object.freeze({
    start: parsedOperations[open].start,
    end: parsedOperations[close].end,
  });
}

/** The whole block a picture's paint is: the `Do` or this writer's wrapper,
 * the `q…Q` group that owns it, every group that exists only for it, and the
 * marked-content tags around them. */
function imagePaintBlock(
  parsedOperations: readonly ContentOperation[],
  paintSpan: Readonly<{ start: number; end: number }>,
  wrapped: boolean,
): Readonly<{ start: number; end: number }> {
  const own = wrapped ? paintSpan : imageObjectBlock(parsedOperations, paintSpan);
  // A wrapper this writer's grammar recognises may still sit in a group that
  // exists only for the picture (a figure's clip); it widens the same way.
  const openIndex = parsedOperations.findIndex(
    (operation) => operation.start === own.start && operation.operator === "q",
  );
  const closeIndex = parsedOperations.findIndex(
    (operation) => operation.end === own.end && operation.operator === "Q",
  );
  return markedContentBlock(
    parsedOperations,
    openIndex >= 0 && closeIndex > openIndex
      ? widenImageBlockThroughOwnGroups(parsedOperations, openIndex, closeIndex)
      : own,
  );
}

/** Operators that put something on the page, or open an object that does. */
function drawsOnPage(operator: string): boolean {
  return paintMode(operator) !== null || IMAGE_BLOCK_FOREIGN_OPERATORS.has(operator);
}

/**
 * Where a move takes the block: for a front move the first place after which
 * nothing is drawn, for a back move the last place before which nothing is.
 * The place sits at the block's own save and marked-content depth, outside
 * any text object and any path under construction, and the drawing between
 * the block and it never closes a group the block is inside -- the picture
 * keeps the groups
 * (a page-wide `q <cm> … Q`, a form's clip) it is drawn in. Null when no such
 * place exists: the picture cannot get there without leaving its groups.
 */
function imageMoveDestination(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
  front: boolean,
): number | null {
  if (front) {
    let depth = 0;
    let marked = 0;
    let text = 0;
    let pathOpen = false;
    let destination = block.end;
    let pending = false;
    for (const operation of parsedOperations) {
      if (operation.start < block.end) continue;
      const operator = operation.operator;
      if (operator === "q") depth += 1;
      else if (operator === "Q") depth -= 1;
      else if (MARKED_CONTENT_OPEN_OPERATORS.has(operator)) marked += 1;
      else if (operator === "EMC") marked -= 1;
      else if (operator === "BT") text += 1;
      else if (operator === "ET") text -= 1;
      else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) pathOpen = true;
      else if (paintMode(operator) !== null || operator === "n") pathOpen = false;
      if (depth < 0 || marked < 0) {
        // The block's own group closes here; anything drawn after it is
        // drawn outside that group, where the picture cannot follow.
        return parsedOperations.some(
          (item) => item.start >= operation.end && drawsOnPage(item.operator),
        )
          ? null
          : pending
            ? null
            : destination;
      }
      if (drawsOnPage(operator)) pending = true;
      if (pending && depth === 0 && marked === 0 && text === 0 && !pathOpen) {
        destination = operation.end;
        pending = false;
      }
    }
    return pending ? null : destination;
  }
  let depth = 0;
  let marked = 0;
  let text = 0;
  let pathOpen = false;
  const boundaries: { position: number; depth: number; marked: number; valid: boolean }[] = [
    { position: parsedOperations[0]?.start ?? 0, depth: 0, marked: 0, valid: true },
  ];
  let firstDraw: number | null = null;
  for (const operation of parsedOperations) {
    if (operation.end > block.start) break;
    const operator = operation.operator;
    if (firstDraw === null && drawsOnPage(operator)) firstDraw = operation.start;
    if (operator === "q") depth += 1;
    else if (operator === "Q") depth -= 1;
    else if (MARKED_CONTENT_OPEN_OPERATORS.has(operator)) marked += 1;
    else if (operator === "EMC") marked -= 1;
    else if (operator === "BT") text += 1;
    else if (operator === "ET") text -= 1;
    else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) pathOpen = true;
    else if (paintMode(operator) !== null || operator === "n") pathOpen = false;
    if (depth < 0 || marked < 0) return null;
    boundaries.push({ position: operation.end, depth, marked, valid: text === 0 && !pathOpen });
  }
  const blockDepth = boundaries[boundaries.length - 1].depth;
  const blockMarked = boundaries[boundaries.length - 1].marked;
  if (firstDraw === null) return block.start;
  // The last place before anything is drawn at the block's depth, from which
  // the drawing up to the block never drops below that depth.
  let destination: number | null = null;
  for (const boundary of boundaries) {
    if (boundary.position > firstDraw) break;
    if (
      boundary.valid &&
      boundary.depth === blockDepth &&
      boundary.marked === blockMarked
    ) {
      destination = boundary.position;
    }
  }
  if (destination === null) return null;
  const rest = boundaries.filter((boundary) => boundary.position >= destination!);
  return rest.every(
    (boundary) => boundary.depth >= blockDepth && boundary.marked >= blockMarked,
  )
    ? destination
    : null;
}

/** How far a move takes a picture out of the groups it is drawn in: the
 * `q…Q` groups and the marked-content sections it leaves behind. */
export interface PdfNativeEmbeddedImageMoveLift {
  readonly save: number;
  readonly marked: number;
}

/**
 * Where a move takes a picture that cannot reach that end of its drawing
 * inside every group it is drawn in -- a figure's picture sits in its own
 * clip groups, and the drawing goes on after they close. Bring to Front puts
 * it after the last thing drawn; Send to Back before the first. The place is
 * outside any text object and path under construction, and inside no group
 * the picture was not already in: every `q` and marked-content section open
 * there is one the picture is drawn in. The groups between the two are the
 * ones the picture leaves, and carries the state of with it (see
 * imageLiftWrapper). The first such place after the last paint for a front
 * move, the last one before the first paint for a back move, so the picture
 * leaves as few groups as it can. Null when there is none.
 */
function imageLiftDestination(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
  front: boolean,
): Readonly<{ destination: number; lift: PdfNativeEmbeddedImageMoveLift }> | null {
  if (front) {
    const first = parsedOperations.findIndex((operation) => operation.start >= block.end);
    if (first < 0) return null;
    let lastDraw = -1;
    for (let index = first; index < parsedOperations.length; index += 1) {
      if (drawsOnPage(parsedOperations[index].operator)) lastDraw = index;
    }
    if (lastDraw < 0) return null;
    let save = 0;
    let marked = 0;
    let text = 0;
    let pathOpen = false;
    let lowestSave = 0;
    let lowestMarked = 0;
    for (let index = first; index < parsedOperations.length; index += 1) {
      const operator = parsedOperations[index].operator;
      if (operator === "q") save += 1;
      else if (operator === "Q") save -= 1;
      else if (MARKED_CONTENT_OPEN_OPERATORS.has(operator)) marked += 1;
      else if (operator === "EMC") marked -= 1;
      else if (operator === "BT") text += 1;
      else if (operator === "ET") text -= 1;
      else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) pathOpen = true;
      else if (paintMode(operator) !== null || operator === "n") pathOpen = false;
      if (text < 0) return null;
      lowestSave = Math.min(lowestSave, save);
      lowestMarked = Math.min(lowestMarked, marked);
      if (
        index >= lastDraw &&
        text === 0 &&
        !pathOpen &&
        save === lowestSave &&
        marked === lowestMarked
      ) {
        return Object.freeze({
          destination: parsedOperations[index].end,
          lift: Object.freeze({ save: 0 - save, marked: 0 - marked }),
        });
      }
    }
    return null;
  }
  const blockIndex = parsedOperations.findIndex((operation) => operation.start >= block.start);
  const before = blockIndex < 0 ? parsedOperations.length : blockIndex;
  const firstDraw = parsedOperations
    .slice(0, before)
    .findIndex((operation) => drawsOnPage(operation.operator));
  if (firstDraw < 0) return null;
  // The state at each place before the block: place 0 is the drawing's start,
  // place i + 1 the end of operation i.
  const places: { position: number; save: number; marked: number; valid: boolean }[] = [
    { position: parsedOperations[0]?.start ?? 0, save: 0, marked: 0, valid: true },
  ];
  let save = 0;
  let marked = 0;
  let text = 0;
  let pathOpen = false;
  for (let index = 0; index < before; index += 1) {
    const operator = parsedOperations[index].operator;
    if (operator === "q") save += 1;
    else if (operator === "Q") save -= 1;
    else if (MARKED_CONTENT_OPEN_OPERATORS.has(operator)) marked += 1;
    else if (operator === "EMC") marked -= 1;
    else if (operator === "BT") text += 1;
    else if (operator === "ET") text -= 1;
    else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) pathOpen = true;
    else if (paintMode(operator) !== null || operator === "n") pathOpen = false;
    places.push({
      position: parsedOperations[index].end,
      save,
      marked,
      valid: text === 0 && !pathOpen,
    });
  }
  // From a place the drawing up to the block may open groups, never close
  // one that place is inside.
  let lowestSave = places[places.length - 1].save;
  let lowestMarked = places[places.length - 1].marked;
  const reaches: boolean[] = new Array(places.length).fill(false);
  for (let index = places.length - 1; index >= 0; index -= 1) {
    lowestSave = Math.min(lowestSave, places[index].save);
    lowestMarked = Math.min(lowestMarked, places[index].marked);
    reaches[index] =
      lowestSave === places[index].save && lowestMarked === places[index].marked;
  }
  for (let index = firstDraw; index >= 0; index -= 1) {
    const place = places[index];
    if (!place.valid || !reaches[index]) continue;
    const atBlock = places[places.length - 1];
    return Object.freeze({
      destination: place.position,
      lift: Object.freeze({
        save: atBlock.save - place.save,
        marked: atBlock.marked - place.marked,
      }),
    });
  }
  return null;
}

/**
 * Where a picture moved out from under the clips it is drawn in lands: just
 * after the outermost `q…Q` group open at the picture that sets a clip
 * closes, out of that group and every group inside it (Chrome's printable
 * area `q x y w h re W* n … Q` holds the whole page; a picture dragged into
 * the margin was cut off at it). The picture leaves as few groups as it can
 * and keeps its place in the drawing as nearly as it can: what those groups
 * draw after it is drawn under it now. Null when no group the picture is in
 * sets a clip, when a clip is set outside every group (none to leave), or
 * when the place is inside a text object or a path under construction.
 */
function imageEscapeDestination(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): Readonly<{ destination: number; lift: PdfNativeEmbeddedImageMoveLift }> | null {
  const groups: boolean[] = [];
  let topClips = false;
  for (const operation of parsedOperations) {
    if (operation.start >= block.start) break;
    const operator = operation.operator;
    if (operator === "q") groups.push(false);
    else if (operator === "Q") {
      if (groups.length === 0) return null;
      groups.pop();
    } else if (operator === "W" || operator === "W*" ||
      (operator === "Tr" && Number(operation.operands[0]?.value) >= 4)) {
      if (groups.length === 0) topClips = true;
      else groups[groups.length - 1] = true;
    }
  }
  const outermost = groups.indexOf(true);
  if (topClips || outermost < 0) return null;
  const save = groups.length - outermost;
  let depth = 0;
  let marked = 0;
  let text = 0;
  let pathOpen = false;
  for (const operation of parsedOperations) {
    if (operation.start < block.end) continue;
    const operator = operation.operator;
    if (operator === "q") depth += 1;
    else if (operator === "Q") depth -= 1;
    else if (MARKED_CONTENT_OPEN_OPERATORS.has(operator)) marked += 1;
    else if (operator === "EMC") marked -= 1;
    else if (operator === "BT") text += 1;
    else if (operator === "ET") text -= 1;
    else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) pathOpen = true;
    else if (paintMode(operator) !== null || operator === "n") pathOpen = false;
    if (depth === -save) {
      if (text !== 0 || pathOpen || marked > 0) return null;
      return Object.freeze({
        destination: operation.end,
        lift: Object.freeze({ save, marked: 0 - marked }),
      });
    }
  }
  return null;
}

/** Operators a lifted picture's groups may set its state with; they are
 * copied, in order, into the wrapper it is drawn in. */
const IMAGE_LIFT_STATE_OPERATORS = new Set([
  "cm", "gs", "ri", "i", "w", "J", "j", "M", "d",
  "g", "rg", "k", "G", "RG", "K", "cs", "CS", "sc", "SC", "scn", "SCN",
]);
/** Operators that draw or mark and set nothing a picture is drawn with. */
const IMAGE_LIFT_INERT_OPERATORS = new Set([
  "Do", "sh", "BI", "MP", "DP", "BX", "EX",
  "Tc", "Tw", "Tz", "TL", "Tf", "Ts",
]);

/** Keys whose value a reader reads in place of the content their section
 * marks; a section that carries one says it again wherever it is copied. */
const MARKED_CONTENT_SPOKEN_KEYS = /\/(?:ActualText|Alt|E)\b/;

/** Whether a marked-content section a picture leaves may be copied around it,
 * which it may when the copy says nothing the page already says: an
 * optional-content membership (`/OC /name BDC`), which decides whether the
 * picture shows at all; a bare `BMC` tag; or a section whose own dictionary
 * names neither a marked-content id nor text read in its place -- Acrobat
 * writes a picture's own group as `/Figure <<>> BDC` on an untagged page.
 * A section with an /MCID is a tagged page's structure leaf, unique on its
 * page: it travels with the picture instead (see movedMarkedContent). */
function liftableMarkedContent(
  operation: ContentOperation,
  /** The page's own `/Properties`, each read as the producer wrote it, for a
   * section that names its dictionary instead of writing it out. */
  namedProperties: ReadonlyMap<string, string | null> = new Map(),
): boolean {
  const [tag, properties] = operation.operands;
  if (operation.operator === "BMC") {
    return operation.operands.length === 1 && tag?.kind === "name";
  }
  if (operation.operator !== "BDC" || operation.operands.length !== 2 || tag?.kind !== "name") {
    return false;
  }
  // A name is read as its value: `/O#43` is /OC, and the page's
  // `/Properties` are keyed by value.
  if (decodePdfSourceName(tag.value) === "OC") return properties.kind === "name";
  // A section that names its dictionary through the page's `/Properties`
  // (what a placed drawing is tagged with) is read there; the copy names the
  // same entry, so it says exactly what the original said. A name the page
  // does not map, or one whose dictionary this reader cannot read, says
  // nothing here about what the section holds.
  const declared =
    properties.kind === "name"
      ? namedProperties.get(decodePdfSourceName(properties.value)) ?? null
      : properties.kind === "value" && properties.value.startsWith("<<")
        ? properties.value
        : null;
  const read = declared === null ? null : pdfSourceNamesRead(declared);
  return (
    read !== null &&
    !/\/MCID\b/.test(read) &&
    !MARKED_CONTENT_SPOKEN_KEYS.test(read)
  );
}

/**
 * The wrapper a picture is drawn in once it leaves groups: the sections it
 * leaves, copied around it, then `q` (with the compensation `cm` when the
 * drawing it passes moves the page), then every state-setting operator of the
 * groups it leaves that is still in effect where it was -- transforms,
 * graphics states, colours, and clips (a path ended by `W n`) -- in the order
 * they ran; the picture's block; `Q`; and the sections' `EMC`s. A front move
 * leaves the innermost `lift.save` groups the drawing before the block has
 * open, a back move every group the drawing opens between its place and the
 * block (and that place's own later state). Refused when that state cannot be
 * carried: a clip that also paints, text drawn as a clip, a section that
 * cannot be copied, or an operator this reader does not know.
 */
function imageLiftWrapper(
  parsedOperations: readonly ContentOperation[],
  drawingText: string,
  from: number,
  block: Readonly<{ start: number; end: number }>,
  front: boolean,
  lift: PdfNativeEmbeddedImageMoveLift,
  compensation: readonly number[] | null,
  /** Graphics states the picture takes with it, written before the groups'
   * own state exactly as the drawing applied them (imageCarriedPaintStates). */
  carriedStates: readonly string[] = EMPTY_CARRIED_PAINT_STATES,
  /** The page's own `/Properties`, for a section that names its dictionary
   * there rather than writing it out (pageMarkedContentProperties). */
  namedProperties: ReadonlyMap<string, string | null> = EMPTY_NAMED_PROPERTIES,
  /** The picture leaves the groups' clips behind (imageEscapeDestination):
   * their transforms and states go with it, the paths they clip by do not. */
  dropClips = false,
): Readonly<{
  prefix: Uint8Array;
  suffix: Uint8Array;
  graphicsStates: ReadonlySet<string>;
  /** The `gs` the wrapper's own group state applies, in the order it writes
   * them, which is what a carried state has to be proven against. */
  appliedStates: readonly ContentOperation[];
  /** The sections the picture leaves, outermost first. */
  carried: readonly ContentOperation[];
  /** Those of them the wrapper does not copy but takes with it: their
   * `BDC`s, whose place in the drawing the move empties. */
  moved: readonly ContentOperation[];
}> {
  const frames: ContentOperation[][] = [[]];
  const sections: ContentOperation[] = [];
  let path: ContentOperation[] = [];
  let clip = false;
  let inText = false;
  for (const operation of parsedOperations) {
    if (operation.start < from) continue;
    if (operation.start >= block.start) break;
    const operator = operation.operator;
    if (operator === "Tr" && Number(operation.operands[0]?.value) >= 4) {
      refuse("PDF image would leave a group whose text clips it.");
    }
    if (inText) {
      if (operator === "ET") inText = false;
      continue;
    }
    if (operator === "BT") inText = true;
    else if (operator === "q") frames.push([]);
    else if (operator === "Q") {
      if (frames.length < 2) refuse("PDF image group nesting is unbalanced.");
      frames.pop();
    } else if (MARKED_CONTENT_OPEN_OPERATORS.has(operator)) sections.push(operation);
    else if (operator === "EMC") {
      if (sections.length === 0) refuse("PDF image group nesting is unbalanced.");
      sections.pop();
    } else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) path.push(operation);
    else if (operator === "W" || operator === "W*") {
      clip = true;
      path.push(operation);
    } else if (operator === "n") {
      if (clip && !dropClips) frames[frames.length - 1].push(...path, operation);
      path = [];
      clip = false;
    } else if (paintMode(operator) !== null) {
      if (clip) refuse("PDF image would leave a clip that also paints.");
      path = [];
    } else if (IMAGE_LIFT_STATE_OPERATORS.has(operator)) {
      frames[frames.length - 1].push(operation);
    } else if (!IMAGE_LIFT_INERT_OPERATORS.has(operator)) {
      refuse("PDF image would leave a group this reader cannot carry.");
    }
  }
  if (inText || path.length > 0) refuse("PDF image group nesting is unbalanced.");
  const carried = front ? frames.slice(frames.length - lift.save) : frames;
  const copied = front ? sections.slice(sections.length - lift.marked) : sections;
  if (
    (front && (lift.save > frames.length - 1 || lift.marked > sections.length)) ||
    (!front && (frames.length - 1 !== lift.save || sections.length !== lift.marked))
  ) {
    refuse("PDF image group nesting is unbalanced.");
  }
  // A section that carries a marked-content id is one page's structure leaf,
  // and the id is that leaf's only name: a second sequence carrying it would
  // leave the tree pointing at two places. Such a section is not copied --
  // it travels with the picture, out of the place it was written in.
  const moved = copied.filter(
    (operation) => !liftableMarkedContent(operation, namedProperties),
  );
  if (!moved.every(movableMarkedContent)) {
    refuse("PDF image would leave a tagged section it cannot carry.");
  }
  const opening = copied
    .map((operation) => `${drawingText.slice(operation.start, operation.end)}\n`)
    .join("");
  const transform = compensation
    ? ` ${compensation.map(formatImageMoveCompensationNumber).join(" ")} cm`
    : "";
  const state = carried
    .flat()
    .map((operation) => `${drawingText.slice(operation.start, operation.end)}\n`)
    .join("");
  return Object.freeze({
    prefix: latin1Bytes(
      `${opening}q${transform}\n${carriedPaintStateText(carriedStates)}${state}`,
    ),
    suffix: ascii(`\nQ${"\nEMC".repeat(copied.length)}`),
    graphicsStates: graphicsStateNames(carried.flat()),
    appliedStates: Object.freeze(
      carried.flat().filter((operation) => operation.operator === "gs"),
    ),
    carried: Object.freeze(copied),
    moved: Object.freeze(moved),
  });
}

/** Whether a section that cannot be copied may travel with the picture: one
 * this reader can read whole, written with its own dictionary, so the move
 * takes exactly the bytes it is written as. A section named through the
 * page's `/Properties` says nothing here about what it holds. */
function movableMarkedContent(operation: ContentOperation): boolean {
  const [tag, properties] = operation.operands;
  return (
    operation.operator === "BDC" &&
    operation.operands.length === 2 &&
    tag?.kind === "name" &&
    properties.kind === "value" &&
    properties.value.startsWith("<<") &&
    properties.value.endsWith(">>")
  );
}

/** A section that travels with the picture: the byte ranges of the `BDC` it
 * is written as and of the `EMC` that closes it, both of which the move
 * takes out of the drawing and puts back around the picture. */
interface MovedMarkedContentMarkers {
  readonly open: Readonly<{ start: number; end: number }>;
  readonly close: Readonly<{ start: number; end: number }>;
}

/**
 * Where the sections a picture takes with it are written, and whether they
 * may be emptied: a section only travels when the picture is all it holds,
 * so that what is left behind draws nothing and what the picture lands in
 * marks the same content it always did. The closing `EMC`s are found after
 * the block, innermost first, in the drawing the move walks.
 *
 * A section that holds more than the picture is refused whichever way it is
 * read: InDesign writes a placed graphic as one `/Figure` leaf over both the
 * drop shadow and the picture, and the leaf's own `/A <</BBox …>>` is drawn
 * around the two of them. Taking the section along would leave the shadow
 * inside no sequence at all; leaving it behind would untag the picture. The
 * third answer -- marking the picture again where it lands -- needs a second
 * id in the leaf's `/K` and a slot for it in the page's `/StructParents`
 * entry, which are objects no picture move writes.
 */
function movedMarkedContentMarkers(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
  carried: readonly ContentOperation[],
  moved: readonly ContentOperation[],
  limit: number,
): readonly MovedMarkedContentMarkers[] {
  if (moved.length === 0) return [];
  // The sections open at the block close after it, innermost first; the ones
  // this move carries are the innermost `carried.length` of them.
  const closers: ContentOperation[] = [];
  let depth = 0;
  for (const operation of parsedOperations) {
    if (operation.start < block.end) continue;
    if (operation.start >= limit) break;
    if (MARKED_CONTENT_OPEN_OPERATORS.has(operation.operator)) depth += 1;
    else if (operation.operator === "EMC") {
      if (depth > 0) depth -= 1;
      else if (closers.push(operation) >= carried.length) break;
    }
  }
  if (closers.length < carried.length) {
    refuse("PDF image group nesting is unbalanced.");
  }
  const markers = moved.map((open) => {
    const close = closers[carried.length - 1 - carried.indexOf(open)];
    return Object.freeze({
      open: Object.freeze({ start: open.start, end: open.end }),
      close: Object.freeze({ start: close.start, end: close.end }),
    });
  });
  // The sections a move carries are nested, so the outermost one holds
  // everything the others do: the picture, and nothing else that draws.
  const outermost = markers[0];
  for (const operation of parsedOperations) {
    if (operation.end <= outermost.open.end || operation.start >= outermost.close.start) {
      continue;
    }
    if (operation.start >= block.start && operation.end <= block.end) continue;
    if (drawsOnPage(operation.operator)) {
      refuse("PDF image would empty a tagged section that holds other drawing.");
    }
  }
  return Object.freeze(markers);
}

/** What the page's graphics states leave in force where one operation
 * stands: every key they set, as the producer wrote it. Two places paint the
 * same when they hold the same keys at the same values, whatever those keys
 * mean -- which is what a picture moved between them needs. */
type PdfNativePaintState = ReadonlyMap<string, string>;

/**
 * The paint state in force where one operation stands: every `gs` the
 * drawing applies on the way, through the `q`/`Q` stack that saves and
 * restores them. Null when a state this reader cannot read is applied, which
 * is what a soft mask or a font in an ExtGState looks like from here.
 */
function paintStateAt(
  parsedOperations: readonly ContentOperation[],
  offset: number,
  states: ReadonlyMap<string, PdfNativePaintState | null>,
): PdfNativePaintState | null {
  let current: PdfNativePaintState = new Map<string, string>();
  const stack: PdfNativePaintState[] = [];
  for (const operation of parsedOperations) {
    if (operation.start >= offset) break;
    if (operation.operator === "q") stack.push(current);
    else if (operation.operator === "Q") current = stack.pop() ?? current;
    else if (operation.operator === "gs" && operation.operands[0]?.kind === "name") {
      // An operand names the state with its leading slash; the page's
      // resources name it without.
      const applied = states.get(operation.operands[0].value.replace(/^\//, ""));
      // A state this reader cannot read says nothing about what it leaves.
      if (applied === undefined || applied === null) return null;
      // An ExtGState sets the keys it names and leaves the rest alone.
      const next = new Map(current);
      for (const [key, value] of applied) next.set(key, value);
      current = next;
    }
  }
  return current;
}

/**
 * The `gs` operations in force where one operation stands, by the group each
 * was applied in: the drawing's own frame first, the one the block is drawn
 * in last. The same walk `paintStateAt` makes, keeping the operations rather
 * than what they leave.
 */
function graphicsStateFramesAt(
  parsedOperations: readonly ContentOperation[],
  offset: number,
): readonly (readonly ContentOperation[])[] {
  const frames: ContentOperation[][] = [[]];
  for (const operation of parsedOperations) {
    if (operation.start >= offset) break;
    if (operation.operator === "q") frames.push([]);
    else if (operation.operator === "Q") {
      if (frames.length > 1) frames.pop();
    } else if (operation.operator === "gs" && operation.operands[0]?.kind === "name") {
      frames[frames.length - 1].push(operation);
    }
  }
  return frames;
}

/**
 * The `gs` operations a picture's own block applies before it paints, in the
 * order they run, through the `q`/`Q` stack the block opens itself. A key one
 * of them sets is the producer's own, written next to the picture: it stands
 * at that value wherever the block is drawn, so it says nothing about what
 * the drawing around the block left in force. They all sit inside the block's
 * own group, so its `Q` puts back whatever the page had.
 */
function imageOwnPaintStates(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): readonly ContentOperation[] {
  const frames: ContentOperation[][] = [[]];
  for (const operation of parsedOperations) {
    if (operation.start < block.start || operation.end > block.end) continue;
    if (IMAGE_BLOCK_FOREIGN_OPERATORS.has(operation.operator)) break;
    if (operation.operator === "q") frames.push([]);
    else if (operation.operator === "Q") {
      if (frames.length > 1) frames.pop();
    } else if (operation.operator === "gs" && operation.operands[0]?.kind === "name") {
      frames[frames.length - 1].push(operation);
    }
  }
  return Object.freeze(frames.flat());
}

/** One state applied on top of another: an ExtGState sets the keys it names
 * and leaves the rest alone. Null when a state this reader cannot read is
 * applied, which says nothing about what it would leave. */
function applyPaintStates(
  base: PdfNativePaintState,
  applied: readonly ContentOperation[],
  states: ReadonlyMap<string, PdfNativePaintState | null>,
): PdfNativePaintState | null {
  let current = new Map(base);
  for (const operation of applied) {
    if (operation.operands[0]?.kind !== "name") return null;
    const state = states.get(operation.operands[0].value.replace(/^\//, ""));
    if (state === undefined || state === null) return null;
    const next = new Map(current);
    for (const [key, value] of state) next.set(key, value);
    current = next;
  }
  return current;
}

/** How many graphics states a moved picture may take with it. A producer
 * names the one it drew the picture under; a page that needs a longer tail
 * of them to be read back is one this reader stops proving. */
export const MRAI_NATIVE_EMBEDDED_IMAGE_CARRIED_PAINT_STATE_LIMIT = 4;

/**
 * The graphics states a moved picture applies again inside its own `q…Q` so
 * it paints exactly as it was drawn. A picture inherits the state the
 * drawing around it left in force, and a producer that sets one state for
 * its pictures and another for its text (an overprinting page sets both)
 * leaves the picture painting differently at the other end of the page.
 * Naming those states again at the new place restores what the picture had,
 * and restores it only for the picture: the `Q` that closes its block puts
 * back whatever the page had.
 *
 * Proven, never assumed: the shortest tail of the states in force where the
 * picture was drawn -- below the groups the move carries, whose own states
 * the wrapper applies after them (`reapplied`) -- for which the state in
 * force at the destination, with the whole wrapper applied to it, is the
 * state the picture was drawn in, key for key. Null when no such tail
 * exists, which leaves the move refused as before.
 */
function imageCarriedPaintStates(
  ambient: readonly ContentOperation[],
  reapplied: readonly ContentOperation[],
  drawnIn: PdfNativePaintState | null,
  landsIn: PdfNativePaintState | null,
  states: ReadonlyMap<string, PdfNativePaintState | null>,
): readonly ContentOperation[] {
  if (drawnIn === null || landsIn === null) return EMPTY_CONTENT_OPERATIONS;
  const limit = Math.min(MRAI_NATIVE_EMBEDDED_IMAGE_CARRIED_PAINT_STATE_LIMIT, ambient.length);
  for (let tail = 1; tail <= limit; tail += 1) {
    const candidate = ambient.slice(ambient.length - tail);
    if (
      samePaintState(
        applyPaintStates(landsIn, [...candidate, ...reapplied], states),
        drawnIn,
      )
    ) {
      return Object.freeze(candidate);
    }
  }
  return EMPTY_CONTENT_OPERATIONS;
}

const EMPTY_CONTENT_OPERATIONS: readonly ContentOperation[] = Object.freeze([]);
const EMPTY_CARRIED_PAINT_STATES: readonly string[] = Object.freeze([]);
const EMPTY_NAMED_PROPERTIES: ReadonlyMap<string, string | null> = new Map();

/** What a wrapper writes for the states a picture takes with it: each one on
 * its own line, exactly as the drawing wrote it. */
function carriedPaintStateText(carried: readonly string[]): string {
  return carried.map((text) => `${text}\n`).join("");
}

/** What the page starts every key at, for the keys a producer sets with a
 * graphics state (ISO 32000-1 8.4.5, 11.6.6): a key a state does not name
 * stands at the value it had, and one no state has named stands here. */
const PAINT_STATE_DEFAULTS: ReadonlyMap<string, string> = new Map([
  ["ca", "1"], ["CA", "1"], ["BM", "/Normal"], ["SMask", "/None"],
  ["op", "false"], ["OP", "false"], ["OPM", "0"], ["AIS", "false"],
  ["SA", "false"], ["TK", "true"], ["LW", "1"], ["LC", "0"], ["LJ", "0"],
  ["ML", "10"], ["FL", "1"], ["RI", "/RelativeColorimetric"], ["SM", "0.1"],
]);

/** One value as the page reads it: a number by what it is worth, a blend
 * mode by the mode it names, anything else by what the producer wrote. */
function paintStateValue(key: string, value: string): string {
  if (key === "BM") return value === "/Compatible" ? "/Normal" : value;
  const number = Number(value);
  return Number.isFinite(number) && /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(value)
    ? String(number)
    : value;
}

/** Whether two places paint the same: every key either one names, and every
 * key neither names, stands at the same value. */
function samePaintState(
  left: PdfNativePaintState | null,
  right: PdfNativePaintState | null,
): boolean {
  if (left === null || right === null) return false;
  for (const key of new Set([...left.keys(), ...right.keys()])) {
    const fallback = PAINT_STATE_DEFAULTS.get(key);
    const first = left.get(key) ?? fallback;
    const second = right.get(key) ?? fallback;
    // A key this reader knows no default for is one it cannot say was
    // unchanged unless both places name it the same.
    if (first === undefined || second === undefined || first !== second) return false;
  }
  return true;
}

/**
 * The page's own `/Properties` names, each read as the dictionary the
 * producer wrote -- or null for one this reader cannot reach. A section that
 * names its dictionary there rather than writing it out is read the same way
 * as one written inline, so it is copied on the same terms.
 */
async function pageMarkedContentProperties(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<ReadonlyMap<string, string | null>> {
  const properties = new Map<string, string | null>();
  try {
    const graph = await inspectMraiNativePageContentGraph(sourceBytes, pageNumber);
    const owner = await inspectMraiNativeObject(sourceBytes, graph.resourceAuthority.owner);
    const entries = parseMraiNativeResourceCategoryEntries(
      owner.objectBytes,
      graph.resourceAuthority.dictionaryKind,
      "Properties",
      "PDF page marked-content properties",
    );
    for (const entry of entries) {
      properties.set(
        entry.resourceName,
        entry.kind === "direct"
          ? entry.dictionary
          : await (async () => {
              const stored = await inspectMraiNativeObject(sourceBytes, entry.target);
              return stored.streamBytes !== null ? null : latin1(stored.objectBytes);
            })(),
      );
    }
  } catch {
    // Reading the page's names is an aid, never a reason to refuse: a
    // category this reader cannot read leaves every name unknown, which is
    // where the walk started.
    return new Map();
  }
  return properties;
}

/** The page's own ExtGState names, each read for the keys it sets -- or null
 * for one this reader cannot read whole, which nothing may be moved past. */
async function pagePaintStates(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<ReadonlyMap<string, PdfNativePaintState | null>> {
  const states = new Map<string, PdfNativePaintState | null>();
  try {
    const graph = await inspectMraiNativePageContentGraph(sourceBytes, pageNumber);
    const owner = await inspectMraiNativeObject(sourceBytes, graph.resourceAuthority.owner);
    const entries = parseMraiNativeExtGStateDictionaryEntries(
      owner.objectBytes,
      graph.resourceAuthority.dictionaryKind,
      "PDF page ExtGState resources",
    );
    for (const entry of entries) {
      const dictionary =
        entry.kind === "direct"
          ? entry.dictionary
          : await (async () => {
              const state = await inspectMraiNativeObject(sourceBytes, entry.target);
              return state.streamBytes !== null ? null : latin1(state.objectBytes);
            })();
      states.set(entry.resourceName, dictionary === null ? null : paintStateKeys(dictionary));
    }
  } catch {
    // Reading the page's states is an aid, never a reason to refuse: an
    // unreadable category leaves every name unknown, which is where the
    // walk started.
    return new Map();
  }
  return states;
}

/** The keys one ExtGState sets, as the producer wrote them; null when this
 * reader cannot say what the whole dictionary sets. */
function paintStateKeys(dictionary: string): PdfNativePaintState | null {
  const open = dictionary.indexOf("<<");
  const close = dictionary.lastIndexOf(">>");
  if (open < 0 || close <= open) return null;
  const body = dictionary.slice(open + 2, close);
  const keys = new Map<string, string>();
  const entry =
    /\/([A-Za-z][A-Za-z0-9]{0,31})\s*((?:\/[^\s/<>\[\]()%]+)|(?:\[[^\[\]]*\])|(?:<<[^<>]*>>)|(?:\d+\s+\d+\s+R)|(?:true|false)|(?:[-+]?(?:\d+(?:\.\d*)?|\.\d+)))/g;
  let read = 0;
  for (const match of body.matchAll(entry)) {
    if (body.slice(read, match.index!).trim()) return null;
    read = match.index! + match[0].length;
    if (keys.has(match[1])) return null;
    if (match[1] !== "Type") {
      keys.set(match[1], paintStateValue(match[1], match[2].replace(/\s+/g, " ").trim()));
    }
  }
  return body.slice(read).trim() ? null : keys;
}

/** The graphics-state resources these operators apply, by name. */
function graphicsStateNames(applied: readonly ContentOperation[]): ReadonlySet<string> {
  return new Set(
    applied
      .filter((operation) => operation.operator === "gs" && operation.operands[0]?.kind === "name")
      .map((operation) => operation.operands[0].value),
  );
}

/** Graphics-state names a picture applies itself before its paint: in its
 * own block, and in the groups it carries out of. A state the drawing it
 * passes sets under one of these names is one the picture sets again after
 * it, so it is not inherited (the keys that state sets are overwritten by the
 * same values; the ones it does not set it never touched). */
function imageOwnGraphicsStates(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
  carried: ReadonlySet<string> = new Set(),
): ReadonlySet<string> {
  const own: ContentOperation[] = [];
  for (const operation of parsedOperations) {
    if (operation.start < block.start || operation.end > block.end) continue;
    if (operation.operator === "Do") break;
    own.push(operation);
  }
  return new Set([...graphicsStateNames(own), ...carried]);
}

/**
 * Absorb the marked-content sequences that wrap a block, so a tagged page
 * (what Word and WPS write) moves the picture with its own `BDC…EMC` tag
 * instead of leaving an empty tag behind and untagging the picture.
 */
function markedContentBlock(
  parsedOperations: readonly ContentOperation[],
  block: Readonly<{ start: number; end: number }>,
): Readonly<{ start: number; end: number }> {
  let start = block.start;
  let end = block.end;
  for (;;) {
    const openIndex = parsedOperations.findIndex(
      (operation) => operation.start === start,
    );
    const closeIndex = parsedOperations.findIndex(
      (operation) => operation.end === end,
    );
    const before = openIndex > 0 ? parsedOperations[openIndex - 1] : undefined;
    const after = closeIndex >= 0 ? parsedOperations[closeIndex + 1] : undefined;
    if (
      !before ||
      !MARKED_CONTENT_OPEN_OPERATORS.has(before.operator) ||
      !after ||
      after.operator !== "EMC" ||
      after.operands.length !== 0
    ) {
      return Object.freeze({ start, end });
    }
    start = before.start;
    end = after.end;
  }
}

interface ContentNesting {
  readonly save: number;
  readonly marked: number;
  readonly text: number;
  readonly path: boolean;
  readonly minimum: number;
}

function contentNesting(
  parsedOperations: readonly ContentOperation[],
  start: number,
  end: number,
): ContentNesting {
  let save = 0;
  let marked = 0;
  let text = 0;
  let path = false;
  let minimum = 0;
  for (const operation of parsedOperations) {
    if (operation.start < start || operation.end > end) continue;
    if (operation.operator === "q") save += 1;
    else if (operation.operator === "Q") save -= 1;
    else if (MARKED_CONTENT_OPEN_OPERATORS.has(operation.operator)) marked += 1;
    else if (operation.operator === "EMC") marked -= 1;
    else if (operation.operator === "BT") text += 1;
    else if (operation.operator === "ET") text -= 1;
    else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operation.operator)) path = true;
    else if (paintMode(operation.operator) !== null || operation.operator === "n") {
      path = false;
    }
    minimum = Math.min(minimum, save, marked, text);
  }
  return { save, marked, text, path, minimum };
}

function balancedContentBlock(nesting: ContentNesting): boolean {
  return (
    nesting.save === 0 &&
    nesting.marked === 0 &&
    nesting.text === 0 &&
    !nesting.path &&
    nesting.minimum === 0
  );
}

/** What a moved block is drawn inside at its new place: nothing, the
 * compensation `q <cm> … Q`, or the wrapper of the groups it left. */
interface ImageMoveWrapper {
  readonly prefix: Uint8Array;
  readonly suffix: Uint8Array;
}

function compensationWrapper(
  compensation: readonly number[] | null,
  carriedStates: readonly string[] = EMPTY_CARRIED_PAINT_STATES,
): ImageMoveWrapper | null {
  const transform = compensation
    ? ` ${compensation.map(formatImageMoveCompensationNumber).join(" ")} cm`
    : "";
  const states = carriedPaintStateText(carriedStates);
  return transform || states
    ? Object.freeze({
        prefix: latin1Bytes(`q${transform}\n${states}`),
        suffix: ascii("\nQ"),
      })
    : null;
}

/**
 * A move's rotation as pieces of the span it rotates: the block and the rest
 * are ranges of that span, and only the separator and the block's wrapper
 * are bytes of the move's own.
 */
function rotatedImageSpanPieces(
  baseLength: number,
  blockLength: number,
  blockLeads: boolean,
  wrapper: ImageMoveWrapper | null = null,
  emptied: readonly Readonly<{ start: number; end: number }>[] = [],
  stationary = 0,
): readonly PdfNativeReplacementPiece[] {
  const prefix = wrapper?.prefix ?? new Uint8Array(0);
  const suffix = wrapper?.suffix ?? new Uint8Array(0);
  const blockStart = blockLeads ? stationary : baseLength - stationary - blockLength;
  const blockEnd = blockStart + blockLength;
  // The bytes the span keeps where they are, in order, with the places the
  // sections that travel with the picture were written in left out.
  const kept = (start: number, end: number): PdfNativeReplacementPiece[] => {
    const pieces: PdfNativeReplacementPiece[] = [];
    let at = start;
    for (const range of emptied) {
      if (range.start < start || range.end > end) continue;
      if (range.start > at) pieces.push({ spanStart: at, spanEnd: range.start });
      at = range.end;
    }
    if (end > at) pieces.push({ spanStart: at, spanEnd: end });
    return pieces;
  };
  const pieces: PdfNativeReplacementPiece[] = blockLeads
    ? [
        ...kept(0, blockStart),
        ...kept(blockEnd, baseLength),
        { literal: concatVectorBytes([ascii("\n"), prefix]) },
        { spanStart: blockStart, spanEnd: blockEnd },
        { literal: suffix },
      ]
    : [
        { literal: prefix },
        { spanStart: blockStart, spanEnd: blockEnd },
        { literal: concatVectorBytes([suffix, ascii("\n")]) },
        ...kept(0, blockStart),
        ...kept(blockEnd, baseLength),
      ];
  return pieces.filter(
    (piece) => ("literal" in piece ? piece.literal.length > 0 : piece.spanEnd > piece.spanStart),
  );
}

function rotatedImageSpanBytes(
  base: Uint8Array,
  blockLength: number,
  blockLeads: boolean,
  wrapper: ImageMoveWrapper | null = null,
  emptied: readonly Readonly<{ start: number; end: number }>[] = [],
  stationary = 0,
): Uint8Array {
  return expandPdfNativeReplacementPieces(
    base,
    rotatedImageSpanPieces(base.length, blockLength, blockLeads, wrapper, emptied, stationary),
  );
}

/** A compensation number, precise enough that the picture moves by no
 * visible amount; the value a record keeps is the one this text reads as. */
function formatImageMoveCompensationNumber(value: number): string {
  if (!Number.isFinite(value)) refuse("PDF image move compensation is not finite.");
  const text = (Math.round(value * 1e9) / 1e9).toFixed(9).replace(/\.?0+$/, "");
  return text === "" || text === "-" || text === "-0" ? "0" : text;
}

/**
 * What the drawing a move passes does to the page's transform, undone for the
 * moved block. A producer that positions everything with top-level `cm`s
 * (Chrome's print writer does) shifts the page between the block and its new
 * place; wrapping the block in the inverse of that shift -- or, moving it
 * back, in the shift itself -- draws it exactly where it was. A clip or a
 * graphics state left open cannot be undone this way and is still refused,
 * as is a fill colour a stencil would inherit. Null when the drawing leaves
 * the transform as it was.
 */
function imageMoveCompensation(
  parsedOperations: readonly ContentOperation[],
  restStart: number,
  restEnd: number,
  stencil: boolean,
  blockLeads: boolean,
  /** The depth, relative to the block, the picture is drawn at once it has
   * left its groups: 0, or minus the groups a front move leaves. */
  level = 0,
  /** Graphics states the picture applies itself (imageOwnGraphicsStates). */
  ownStates: ReadonlySet<string> = new Set(),
  /** True when the drawing the picture passes leaves it painting exactly as
   * it was drawn: every `gs` on the way has been read, and what they leave in
   * force is what the picture already had -- or is made so by the states the
   * picture's own block names (see paintStateAt, imageOwnPaintStates). */
  paintStateKept = false,
): readonly number[] | null {
  let save = 0;
  let shift: Matrix = [1, 0, 0, 1, 0, 0];
  let moved = false;
  for (const operation of parsedOperations) {
    if (operation.start < restStart || operation.end > restEnd) continue;
    if (operation.operator === "q") {
      save += 1;
      continue;
    }
    if (operation.operator === "Q") {
      save -= 1;
      continue;
    }
    if (save !== level) continue;
    if (operation.operator === "cm") {
      const [a, b, c, d, e, f] = numberOperands(operation, 6);
      shift = multiply(shift, [a, b, c, d, e, f]);
      moved = true;
      continue;
    }
    // Text drawn as a clip clips everything after it at this depth.
    if (operation.operator === "Tr" && Number(operation.operands[0]?.value) >= 4) {
      refuse("PDF page clips with text the moved picture would pass.");
    }
    if (
      operation.operator === "gs" &&
      operation.operands[0]?.kind === "name" &&
      (paintStateKept || ownStates.has(operation.operands[0].value))
    ) {
      continue;
    }
    if (operation.operator === "gs") {
      refuse("PDF page sets a graphics state the moved picture would inherit.");
    }
    if (IMAGE_MOVE_GEOMETRY_OPERATORS.has(operation.operator)) {
      refuse("PDF page sets a transform or clip the moved picture would inherit.");
    }
    if (stencil && IMAGE_MOVE_COLOR_OPERATORS.has(operation.operator)) {
      refuse("PDF page sets a fill colour the moved stencil would inherit.");
    }
  }
  if (!moved) return null;
  const [a, b, c, d, e, f] = shift;
  const compensation: Matrix = blockLeads
    ? (() => {
        const determinant = a * d - b * c;
        if (!Number.isFinite(determinant) || Math.abs(determinant) < 1e-12) {
          refuse("PDF page transform the picture would pass cannot be undone.");
        }
        return [
          d / determinant,
          -b / determinant,
          -c / determinant,
          a / determinant,
          (c * f - d * e) / determinant,
          (b * e - a * f) / determinant,
        ];
      })()
    : shift;
  if (compensation.every((value, index) => value === [1, 0, 0, 1, 0, 0][index])) return null;
  return Object.freeze(
    compensation.map((value) => Number(formatImageMoveCompensationNumber(value))),
  );
}

const PDF_REGULAR_BYTE = (value: number | undefined): boolean =>
  value !== undefined &&
  value > 0x20 &&
  value < 0x7f &&
  !"()<>[]{}/%".includes(String.fromCharCode(value));

/**
 * Where a back move writes the block, given the place it found: a back move
 * writes the block first in its span with nothing of its own before it, and
 * the place is the end of an operator, so `1 0 0 1 0 0 cm` and the block's
 * `q` ran together into one `cmq` a reader cannot read -- the page lost
 * everything after it (user audit PDF 0049). The block goes after the
 * whitespace that follows the operator, which still comes before anything
 * drawn; with no whitespace there to write after, the move is refused rather
 * than run two tokens together. (A step lands at the start of the object it
 * passes, after whatever separated that object, and keeps its place.)
 */
function separatedBackPlace(drawing: Uint8Array, place: number, limit: number): number {
  let at = place;
  while (at < limit && PDF_WHITESPACE_ONLY.test(String.fromCharCode(drawing[at]))) at += 1;
  if (at > place || at === 0) return at;
  // What is written there (the block, or the group it is wrapped in) may
  // start with a regular character too.
  if (PDF_REGULAR_BYTE(drawing[at - 1])) {
    refuse("PDF page has no space before its first drawing to put the picture in.");
  }
  return at;
}

/** One front or back move of a picture's block within one drawing: the span
 * it rotates, and what the block is drawn inside at its new place. */
interface ImageMove {
  readonly destination: number;
  readonly spanStart: number;
  readonly spanEnd: number;
  readonly compensation: readonly number[] | null;
  readonly lift: PdfNativeEmbeddedImageMoveLift | null;
  readonly wrapper: ImageMoveWrapper | null;
  /** The graphics states the picture names again at its new place so it
   * paints as it was drawn, each exactly as the drawing wrote it. */
  readonly carriedStates: readonly string[];
  /** The markers of the sections that travel with the picture, in the
   * drawing's own offsets; the span reaches over them so the move can empty
   * the places they were written in. */
  readonly markers: readonly MovedMarkedContentMarkers[];
}

/**
 * Where a front or back move takes a block, and how it is drawn there. Inside
 * every group it is drawn in when it can get there (imageMoveDestination);
 * otherwise out of the groups it has to leave, carrying their state with it
 * (imageLiftDestination, imageLiftWrapper). A front move then undoes what the
 * drawing it passes does to the page's transform at the depth it lands at; a
 * back move is drawn before all of that drawing and carries the state it
 * had whole, so nothing is left to undo.
 */
function imageMove(
  parsedOperations: readonly ContentOperation[],
  drawing: Uint8Array,
  block: Readonly<{ start: number; end: number }>,
  blockLeads: boolean,
  stencil: boolean,
  /** The page's own ExtGState names, read for what they set. */
  paintStates: ReadonlyMap<string, PdfNativePaintState | null> = new Map(),
  /** The page's own `/Properties` names, read for what they declare. */
  namedProperties: ReadonlyMap<string, string | null> = EMPTY_NAMED_PROPERTIES,
  /** A one-step move: where the block lands, past exactly one unit of
   * content at its own depth, so it never leaves a group. */
  stepDestination: number | null = null,
  /** Out from under the clips the picture is drawn in (a front move only):
   * it lands after the groups that set them, which it leaves without them. */
  leaveClips = false,
): ImageMove {
  let destination = leaveClips
    ? null
    : stepDestination ?? imageMoveDestination(parsedOperations, block, blockLeads);
  let lift: PdfNativeEmbeddedImageMoveLift | null = null;
  if (leaveClips) {
    const escaped = blockLeads ? imageEscapeDestination(parsedOperations, block) : null;
    if (escaped === null) refuse(MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED);
    destination = escaped.destination;
    lift = escaped.lift;
  } else if (destination === null) {
    const lifted = imageLiftDestination(parsedOperations, block, blockLeads);
    if (lifted === null) {
      refuse("PDF image cannot reach that end of the page without leaving its group.");
    }
    destination = lifted.destination;
    if (lifted.lift.save > 0 || lifted.lift.marked > 0) lift = lifted.lift;
  } else if (stepDestination !== null) {
    // A step lands beside the next object, which may stand outside the
    // groups and sections the picture is drawn in: it leaves those, carrying
    // their state the way a move to either end does.
    const nesting = blockLeads
      ? contentNesting(parsedOperations, block.end, stepDestination)
      : contentNesting(parsedOperations, stepDestination, block.start);
    const save = blockLeads ? 0 - nesting.save : nesting.save;
    const marked = blockLeads ? 0 - nesting.marked : nesting.marked;
    if (save > 0 || marked > 0) lift = Object.freeze({ save, marked });
  }
  if (!blockLeads && stepDestination === null) {
    destination = separatedBackPlace(drawing, destination, block.start);
  }
  const restStart = blockLeads ? block.end : destination;
  const restEnd = blockLeads ? destination : block.start;
  if (
    restEnd <= restStart ||
    PDF_WHITESPACE_ONLY.test(latin1(drawing.slice(restStart, restEnd)))
  ) {
    refuse(
      blockLeads
        ? MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT
        : MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK,
    );
  }
  const spanStart = blockLeads ? block.start : destination;
  const spanEnd = blockLeads ? destination : block.end;
  // A picture keeps the way it is painted when what the drawing it passes
  // leaves in force is what it already had: a producer that sets its own
  // default state between the things it draws (a deck writes `/GS5 gs` with
  // `ca 1` and `/BM /Normal` before every element) changes nothing. And when
  // the two places are not the same, the picture still paints the same if its
  // own block names a state that sets every key they differ in -- a report
  // that writes `/GS0 gs` inside each picture's group, and overprint on for
  // the text around them, is drawn in its own state at either end.
  const drawnIn = paintStateAt(parsedOperations, block.start, paintStates);
  const landsIn = paintStateAt(parsedOperations, destination, paintStates);
  const ownPaintStates = imageOwnPaintStates(parsedOperations, block);
  const asPainted = (state: PdfNativePaintState | null) =>
    state === null ? null : applyPaintStates(state, ownPaintStates, paintStates);
  const paintStateKept =
    samePaintState(drawnIn, landsIn) ||
    samePaintState(asPainted(drawnIn), asPainted(landsIn));
  const drawingText = latin1(drawing);
  // A state the drawing it passes leaves in a different place is one the
  // picture may take with it, named again inside its own group.
  const stateText = (operations: readonly ContentOperation[]) =>
    Object.freeze(
      operations.map((operation) => drawingText.slice(operation.start, operation.end)),
    );
  const frames = graphicsStateFramesAt(parsedOperations, block.start);
  if (lift === null) {
    const carriedStates = paintStateKept
      ? EMPTY_CARRIED_PAINT_STATES
      : stateText(
          imageCarriedPaintStates(
            frames.flat(),
            EMPTY_CONTENT_OPERATIONS,
            drawnIn,
            landsIn,
            paintStates,
          ),
        );
    const compensation = imageMoveCompensation(
      parsedOperations,
      restStart,
      restEnd,
      stencil,
      blockLeads,
      0,
      imageOwnGraphicsStates(parsedOperations, block),
      paintStateKept || carriedStates.length > 0,
    );
    return Object.freeze({
      destination,
      spanStart,
      spanEnd,
      compensation,
      lift: null,
      wrapper: compensationWrapper(compensation, carriedStates),
      carriedStates,
      markers: [],
    });
  }
  if (!blockLeads) {
    const wrapper = imageLiftWrapper(
      parsedOperations,
      drawingText,
      destination,
      block,
      false,
      lift,
      null,
      EMPTY_CARRIED_PAINT_STATES,
      namedProperties,
    );
    const markers = movedMarkedContentMarkers(
      parsedOperations,
      block,
      wrapper.carried,
      wrapper.moved,
      drawing.length,
    );
    return Object.freeze({
      destination,
      spanStart,
      // A section that travels with the picture is written before the block
      // and closed after it, so a back move's span reaches past the block to
      // the `EMC` it takes with it.
      spanEnd: markers.length > 0 ? Math.max(spanEnd, markers[0].close.end) : spanEnd,
      compensation: null,
      lift,
      wrapper,
      carriedStates: EMPTY_CARRIED_PAINT_STATES,
      markers,
    });
  }
  const carried = imageLiftWrapper(
    parsedOperations,
    drawingText,
    0,
    block,
    true,
    lift,
    null,
    EMPTY_CARRIED_PAINT_STATES,
    namedProperties,
    leaveClips,
  );
  // The groups the picture leaves are re-entered by the wrapper, so only the
  // states applied below them are the picture's to take.
  const carriedStates = paintStateKept
    ? EMPTY_CARRIED_PAINT_STATES
    : stateText(
        imageCarriedPaintStates(
          frames.slice(0, Math.max(0, frames.length - lift.save)).flat(),
          carried.appliedStates,
          drawnIn,
          landsIn,
          paintStates,
        ),
      );
  const compensation = imageMoveCompensation(
    parsedOperations,
    restStart,
    restEnd,
    stencil,
    true,
    -lift.save,
    imageOwnGraphicsStates(parsedOperations, block, carried.graphicsStates),
    paintStateKept || carriedStates.length > 0,
  );
  const markers = movedMarkedContentMarkers(
    parsedOperations,
    block,
    carried.carried,
    carried.moved,
    destination,
  );
  return Object.freeze({
    destination,
    // A section that travels with the picture opens before the block, so a
    // front move's span starts at the `BDC` it takes with it.
    spanStart: markers.length > 0 ? Math.min(spanStart, markers[0].open.start) : spanStart,
    spanEnd,
    compensation,
    lift,
    wrapper:
      compensation || carriedStates.length > 0
        ? imageLiftWrapper(
            parsedOperations,
            drawingText,
            0,
            block,
            true,
            lift,
            compensation,
            carriedStates,
            namedProperties,
            leaveClips,
          )
        : carried,
    carriedStates,
    markers,
  });
}

async function imageReorderRequest(input: {
  readonly stream: PdfNativeEmbeddedImagePageStream;
  readonly spanStart: number;
  readonly spanEnd: number;
  readonly pieces: readonly PdfNativeReplacementPiece[];
}): Promise<PdfNativeContentStreamTransactionRequest> {
  const original = input.stream.streamBytes.slice(input.spanStart, input.spanEnd);
  return Object.freeze({
    identity: input.stream.identity,
    expectedObjectSha256: input.stream.objectSha256,
    expectedStreamSha256: input.stream.streamSha256,
    operatorSpan: Object.freeze({
      start: input.spanStart,
      end: input.spanEnd,
      expectedSha256: await sha256PdfBytes(original),
    }),
    // A move only reorders its span, so the record names those bytes by
    // range instead of carrying them again.
    replacement: expandPdfNativeReplacementPieces(original, input.pieces),
    replacementPieces: input.pieces,
    untouchedObjects: input.stream.untouchedObjects,
  });
}

interface ResolvedImageBlock {
  /** The drawing the picture is moved inside: the page's streams, or the one
   * stream of the Form XObject host that paints it. Its paints are that
   * drawing's own, each keeping its order on the whole page. */
  readonly paints: PdfNativeEmbeddedImagePagePaints;
  readonly parsedOperations: readonly ContentOperation[];
  readonly paint: PdfNativeEmbeddedImagePaintBlock;
  /** The content stream this move rewrites, and nothing else. */
  readonly stream: PdfNativeEmbeddedImagePageStream;
  /** Whether a Form XObject host paints the picture. */
  readonly hosted: boolean;
  /** How many paints the whole page lists. */
  readonly pagePaintCount: number;
  /** The moved block, as an offset into the whole page drawing. */
  readonly block: Readonly<{ start: number; end: number }>;
}

async function resolveImageBlock(
  sourceBytes: Uint8Array,
  pageNumber: number,
  imageOrder: number,
): Promise<ResolvedImageBlock> {
  const page = await inspectMraiNativeEmbeddedImagePagePaints(
    sourceBytes,
    pageNumber,
  );
  const selected = page.paints[imageOrder];
  if (!selected) invalid("Selected PDF image is out of range.");
  // A picture a form paints moves inside that form's stream: moving it out
  // would change the space and the resources it is drawn with.
  const hostIndex = selected.formStreamIndex;
  const paints: PdfNativeEmbeddedImagePagePaints =
    hostIndex === undefined
      ? Object.freeze({
          ...page,
          paints: Object.freeze(page.paints.filter((entry) => entry.formStreamIndex === undefined)),
        })
      : Object.freeze({
          ...page,
          streams: Object.freeze([Object.freeze({ ...page.formStreams[hostIndex], index: 0 })]),
          drawingBytes: page.formStreams[hostIndex].streamBytes,
          paints: Object.freeze(
            page.paints
              .filter((entry) => entry.formStreamIndex === hostIndex)
              .map((entry) => Object.freeze({ ...entry, streamIndex: 0 })),
          ),
        });
  const paint = paints.paints.find((entry) => entry.imageOrder === imageOrder);
  if (!paint) invalid("Selected PDF image is out of range.");
  const stream = paints.streams[paint.streamIndex];
  if (!stream) invalid("Selected PDF image is out of range.");
  // Order on a page is the order of the whole drawing, so the block is read
  // out of the concatenated streams and only then placed back in the one
  // stream this transaction rewrites.
  let parsedOperations: readonly ContentOperation[];
  try {
    parsedOperations = operations(latin1(paints.drawingBytes), true);
  } catch {
    refuse("PDF page content stream cannot be reordered.");
  }
  const whole = contentNesting(parsedOperations, 0, paints.drawingBytes.length);
  if (!balancedContentBlock(whole)) {
    refuse("PDF page content stream nesting is unbalanced.");
  }
  const block = imagePaintBlock(parsedOperations, paint.drawingSpan, paint.wrapped);
  const inner = contentNesting(parsedOperations, block.start, block.end);
  if (!balancedContentBlock(inner)) {
    refuse("PDF image is not painted by a self-contained block.");
  }
  // A move rewrites one content stream, so everything it carries has to live
  // in that stream: a block that reaches into a neighbour is refused rather
  // than split across two objects.
  if (
    block.start < stream.drawingSpan.start ||
    block.end > stream.drawingSpan.end
  ) {
    refuse("PDF image block crosses a page content stream boundary.");
  }
  const owned = paints.paints.filter(
    (entry) =>
      entry.drawingSpan.start >= block.start &&
      entry.drawingSpan.end <= block.end,
  );
  if (owned.length !== 1 || owned[0].imageOrder !== imageOrder) {
    refuse("PDF image shares its block with another picture.");
  }
  return {
    paints,
    parsedOperations,
    paint,
    stream,
    block,
    hosted: hostIndex !== undefined,
    pagePaintCount: page.paints.length,
  };
}

export async function planMraiNativeEmbeddedImageReorder(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly direction: "front" | "back" | "forward" | "backward";
  /** A front move out from under the clips the picture is drawn in, which
   * it lands just after (imageEscapeDestination): a picture moved past a
   * clip the page draws everything under is drawn whole. */
  readonly leaveClips?: boolean;
}): Promise<PdfNativeEmbeddedImageReorderPlan> {
  const leaveClips = input.leaveClips === true;
  if (leaveClips && input.direction !== "front") {
    invalid("PDF image leaves its clips only on a front move.");
  }
  if (
    input.direction !== "front" &&
    input.direction !== "back" &&
    input.direction !== "forward" &&
    input.direction !== "backward"
  ) {
    invalid("PDF image reorder direction is invalid.");
  }
  const resolved = await resolveImageBlock(
    input.sourceBytes,
    input.pageNumber,
    input.imageOrder,
  );
  const drawing = resolved.paints.drawingBytes;
  const { block, stream } = resolved;
  const blockLeads = input.direction === "front" || input.direction === "forward";
  const stepped = input.direction === "forward" || input.direction === "backward";
  // One step moves the picture past exactly one unit of content -- a path
  // object, a text object, a group, another XObject or a shading -- the same
  // unit a path object passes (reorderPassedUnit); nothing to pass is the
  // end the picture already stands at.
  // A unit that marks nothing (an empty text object) is no object to pass,
  // and content the unit grammar cannot read -- a tagged paragraph, a group
  // the picture has to leave its own group for -- is passed the way a layer
  // move passes the next object (layerMoveDestination).
  let steppedUnit: Readonly<{ start: number; end: number }> | null = null;
  if (stepped) {
    let unit: PdfNativeVectorPassedUnit | null = null;
    try {
      unit = reorderPassedUnit(resolved.parsedOperations, latin1(drawing), block, blockLeads);
    } catch (error) {
      if (!(error instanceof PdfNativeTransactionError) || error.message !== LAYER_MOVE_REFUSAL) {
        throw error;
      }
    }
    if (unit && layerRangeMarks(resolved.parsedOperations, unit)) {
      steppedUnit = unit;
    } else {
      const landing = layerMoveDestination(
        resolved.parsedOperations,
        block,
        blockLeads ? "forward" : "backward",
      );
      if (landing !== null) {
        steppedUnit = blockLeads
          ? Object.freeze({ start: block.end, end: landing })
          : Object.freeze({ start: landing, end: block.start });
      }
    }
  }
  if (stepped && !steppedUnit) {
    refuse(
      blockLeads
        ? MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_FRONT
        : MRAI_NATIVE_EMBEDDED_IMAGE_ALREADY_BACK,
    );
  }
  // In front means nothing is drawn after the picture and behind means
  // nothing before it; the picture stays inside every group it can.
  // A picture a form paints is drawn with that form's own states, which the
  // page's names say nothing about.
  const paintStates = resolved.hosted
    ? new Map<string, PdfNativePaintState | null>()
    : await pagePaintStates(input.sourceBytes, resolved.paints.pageNumber);
  const namedProperties = resolved.hosted
    ? EMPTY_NAMED_PROPERTIES
    : await pageMarkedContentProperties(input.sourceBytes, resolved.paints.pageNumber);
  // A form's own box clips what it paints, which no move leaves.
  if (leaveClips && resolved.hosted) refuse(MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED);
  let move: ImageMove;
  try {
    move = imageMove(
      resolved.parsedOperations,
      drawing,
      block,
      blockLeads,
      resolved.paint.stencil,
      paintStates,
      namedProperties,
      steppedUnit ? (blockLeads ? steppedUnit.end : steppedUnit.start) : null,
      leaveClips,
    );
  } catch (error) {
    // A step refused for what the next object leaves in force (a clip, a
    // state the picture cannot name again) is said the way the window says
    // it: content the picture cannot pass.
    if (
      stepped &&
      error instanceof PdfNativeTransactionError &&
      error.code === "unsupported-document" &&
      !pdfNativeEmbeddedImageIsAlreadyArranged(error)
    ) {
      refuse(LAYER_MOVE_REFUSAL);
    }
    throw error;
  }
  const { destination } = move;
  if (
    destination < stream.drawingSpan.start ||
    destination > stream.drawingSpan.end
  ) {
    // The place is in another of the page's streams. A reader draws them as
    // one sequence, so the page is rewritten as that one stream -- the way
    // the text route consolidates a page whose streams split a text object --
    // and the picture moves inside it.
    if (resolved.hosted || resolved.paints.streams.length < 2 || stepped || leaveClips) {
      refuse("PDF image cannot pass another of its page's content streams.");
    }
    return planConsolidatedEmbeddedImageReorder(
      { ...input, direction: blockLeads ? "front" : "back" },
      resolved,
      blockLeads,
    );
  }
  // One step past a unit that paints pictures moves this one past them.
  const passedUnit = steppedUnit ?? (leaveClips ? { start: block.end, end: destination } : null);
  const passedPictures = passedUnit
    ? resolved.paints.paints.filter(
        (entry) =>
          entry.imageOrder !== input.imageOrder &&
          entry.drawingSpan.start >= passedUnit.start &&
          entry.drawingSpan.end <= passedUnit.end,
      ).length
    : 0;
  const { spanStart, spanEnd, compensation, lift } = move;
  const blockLength = block.end - block.start;
  const emptiedMarkers = movedMarkerSpanRanges(move.markers, spanStart);
  const stationaryBytes = blockLeads ? block.start - spanStart : spanEnd - block.end;
  const request = await imageReorderRequest({
    stream,
    spanStart: spanStart - stream.drawingSpan.start,
    spanEnd: spanEnd - stream.drawingSpan.start,
    pieces: rotatedImageSpanPieces(
      spanEnd - spanStart,
      blockLength,
      blockLeads,
      move.wrapper,
      emptiedMarkers,
      stationaryBytes,
    ),
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
    direction: input.direction,
    stepped,
    pageNumber: resolved.paints.pageNumber,
    imageOrder: input.imageOrder,
    // The end of the page drawing is the last paint on the page and its start
    // the first; a form host's ends are its own last and first paint, which
    // keep their places among the page's other paints. One step passes the
    // pictures the unit paints.
    imageOrderAfter: stepped || leaveClips
      ? input.imageOrder + (blockLeads ? passedPictures : -passedPictures)
      : resolved.hosted
        ? blockLeads
          ? resolved.paints.paints[resolved.paints.paints.length - 1].imageOrder
          : resolved.paints.paints[0].imageOrder
        : blockLeads
          ? resolved.pagePaintCount - 1
          : 0,
    resourceName: resolved.paint.resourceName,
    page: resolved.paints.page,
    // A picture a form host paints is named in the form's resources.
    resourceOwner: stream.resourceOwner ?? resolved.paints.resourceOwner,
    image: resolved.paint.image,
    contentStream: stream.contentStream,
    operatorSpan: request.operatorSpan,
    blockLength,
    blockLeads,
    ...(compensation ? { compensation } : {}),
    ...(move.carriedStates.length > 0 ? { carriedStates: move.carriedStates } : {}),
    ...(lift ? { lift } : {}),
    ...(emptiedMarkers.length > 0 ? { emptiedMarkers, stationaryBytes } : {}),
    request,
  });
}

/** The markers a move empties, in the span's own offsets and in writing
 * order: the `BDC`s the sections open with, then the `EMC`s that close them,
 * which a section closed after the picture writes the other way round. */
function movedMarkerSpanRanges(
  markers: readonly MovedMarkedContentMarkers[],
  spanStart: number,
): readonly Readonly<{ start: number; end: number }>[] {
  return Object.freeze(
    markers
      .flatMap((marker) => [marker.open, marker.close])
      .map((range) =>
        Object.freeze({ start: range.start - spanStart, end: range.end - spanStart }),
      )
      .sort((left, right) => left.start - right.start),
  );
}

/**
 * A front or back move whose place is in another of the page's streams: the
 * streams are joined into one new stream the page draws instead (their bytes
 * end to end, which is how a reader reads them), and the picture moves inside
 * it. The page object is rewritten to name the new stream, so it is no
 * longer pinned; the source streams are, by the consolidation plan's hashes.
 */
async function planConsolidatedEmbeddedImageReorder(
  input: Readonly<{ sourceBytes: Uint8Array; imageOrder: number; direction: "front" | "back" }>,
  resolved: ResolvedImageBlock,
  blockLeads: boolean,
): Promise<PdfNativeEmbeddedImageReorderPlan> {
  const { paints, paint } = resolved;
  const pageObject = await inspectMraiNativeObject(input.sourceBytes, paints.page);
  const joined = concatVectorBytes(paints.streams.map((item) => item.streamBytes));
  const plan: PdfNativePageContentConsolidationPlan = Object.freeze({
    pageNumber: paints.pageNumber,
    page: Object.freeze({ ...paints.page, expectedObjectSha256: pageObject.objectSha256 }),
    sourceStreams: Object.freeze(
      paints.streams.map((item) =>
        Object.freeze({ ...item.contentStream, expectedObjectSha256: item.objectSha256 }),
      ),
    ),
    joinedDecodedSha256: await sha256PdfBytes(joined),
    stream: Object.freeze({
      objectNumber: nextMraiNativeIndirectObjectNumber(input.sourceBytes),
      generation: 0 as const,
    }),
  });
  const consolidated = await inspectMraiNativeConsolidatedContentStream(
    input.sourceBytes,
    plan,
  );
  const drawing = consolidated.streamBytes;
  let parsedOperations: readonly ContentOperation[];
  try {
    parsedOperations = operations(latin1(drawing), true);
  } catch {
    refuse("PDF page content stream cannot be reordered.");
  }
  if (!balancedContentBlock(contentNesting(parsedOperations, 0, drawing.length))) {
    refuse("PDF page content stream nesting is unbalanced.");
  }
  // Each stream starts in the joined bytes where the ones before it end.
  const joinedStart = (streamIndex: number) =>
    paints.streams
      .slice(0, streamIndex)
      .reduce((sum, item) => sum + item.streamBytes.length, 0);
  const joinedSpan = (entry: PdfNativeEmbeddedImagePaintBlock) =>
    Object.freeze({
      start: joinedStart(entry.streamIndex) + entry.span.start,
      end: joinedStart(entry.streamIndex) + entry.span.end,
    });
  const block = imagePaintBlock(parsedOperations, joinedSpan(paint), paint.wrapped);
  if (!balancedContentBlock(contentNesting(parsedOperations, block.start, block.end))) {
    refuse("PDF image is not painted by a self-contained block.");
  }
  const owned = paints.paints.filter((entry) => {
    const span = joinedSpan(entry);
    return span.start >= block.start && span.end <= block.end;
  });
  if (owned.length !== 1 || owned[0].imageOrder !== input.imageOrder) {
    refuse("PDF image shares its block with another picture.");
  }
  const move = imageMove(
    parsedOperations,
    drawing,
    block,
    blockLeads,
    paint.stencil,
    await pagePaintStates(input.sourceBytes, paints.pageNumber),
    await pageMarkedContentProperties(input.sourceBytes, paints.pageNumber),
  );
  const { spanStart, spanEnd, compensation, lift } = move;
  const blockLength = block.end - block.start;
  const emptiedMarkers = movedMarkerSpanRanges(move.markers, spanStart);
  const stationaryBytes = blockLeads ? block.start - spanStart : spanEnd - block.end;
  const original = drawing.slice(spanStart, spanEnd);
  const pageKey = `${paints.page.objectNumber}:${paints.page.generation}`;
  const pins = new Map<string, PdfNativeContentStreamInspection["untouchedObjects"][number]>();
  for (const item of paints.streams.flatMap((entry) => entry.untouchedObjects)) {
    const key = `${item.objectNumber}:${item.generation}`;
    if (key !== pageKey) pins.set(key, item);
  }
  for (const item of paints.streams) {
    pins.set(
      `${item.contentStream.objectNumber}:${item.contentStream.generation}`,
      Object.freeze({
        ...item.contentStream,
        expectedObjectSha256: item.objectSha256,
        role: "untouched-object" as const,
      }),
    );
  }
  const request: PdfNativeContentStreamTransactionRequest = Object.freeze({
    identity: consolidated.identity,
    expectedObjectSha256: consolidated.objectSha256,
    expectedStreamSha256: consolidated.streamSha256,
    operatorSpan: Object.freeze({
      start: spanStart,
      end: spanEnd,
      expectedSha256: await sha256PdfBytes(original),
    }),
    replacement: rotatedImageSpanBytes(
      original,
      blockLength,
      blockLeads,
      move.wrapper,
      emptiedMarkers,
      stationaryBytes,
    ),
    replacementPieces: rotatedImageSpanPieces(
      original.length,
      blockLength,
      blockLeads,
      move.wrapper,
      emptiedMarkers,
      stationaryBytes,
    ),
    untouchedObjects: Object.freeze([...pins.values()]),
    pageContentConsolidation: plan,
  });
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
    direction: input.direction,
    stepped: false,
    pageNumber: paints.pageNumber,
    imageOrder: input.imageOrder,
    imageOrderAfter: blockLeads ? resolved.pagePaintCount - 1 : 0,
    resourceName: paint.resourceName,
    page: paints.page,
    resourceOwner: paints.resourceOwner,
    image: paint.image,
    contentStream: Object.freeze({ ...plan.stream }),
    operatorSpan: request.operatorSpan,
    blockLength,
    blockLeads,
    ...(compensation ? { compensation } : {}),
    ...(move.carriedStates.length > 0 ? { carriedStates: move.carriedStates } : {}),
    ...(lift ? { lift } : {}),
    ...(emptiedMarkers.length > 0 ? { emptiedMarkers, stationaryBytes } : {}),
    request,
  });
}
/** The sections a wrapper took with the picture rather than copying, read
 * back out of the wrapper it wrote: its first `marked` lines are the
 * sections the picture leaves, outermost first, and the ones that may not be
 * copied are the ones whose place in the drawing the move emptied. */
function movedSectionOpeningTexts(
  wrapperPrefixText: string,
  marked: number,
): readonly string[] {
  const lines = wrapperPrefixText.split("\n").slice(0, marked);
  if (lines.length !== marked) {
    invalid("PDF image reorder history does not name the sections it moved.");
  }
  return lines.filter((line) => {
    let parsed: readonly ContentOperation[];
    try {
      parsed = operations(line, true);
    } catch {
      invalid("PDF image reorder history does not name the sections it moved.");
    }
    if (parsed.length !== 1 || !MARKED_CONTENT_OPEN_OPERATORS.has(parsed[0].operator)) {
      invalid("PDF image reorder history does not name the sections it moved.");
    }
    return !liftableMarkedContent(parsed[0]);
  });
}

/** The span as it stood before a move that emptied marked-content markers:
 * the bytes the rotation kept, in their old order, with each emptied place
 * written back and the block where it stood. */
function restoredImageSpanPieces(input: {
  readonly keptStart: number;
  readonly blockPiece: Readonly<{ spanStart: number; spanEnd: number }>;
  readonly markers: readonly Readonly<{ start: number; end: number }>[];
  readonly texts: readonly string[];
  readonly stationary: number;
  readonly blockLength: number;
  readonly blockLeads: boolean;
  readonly baseLength: number;
}): readonly PdfNativeReplacementPiece[] {
  const blockStart = input.blockLeads
    ? input.stationary
    : input.baseLength - input.stationary - input.blockLength;
  const pieces: PdfNativeReplacementPiece[] = [];
  let kept = input.keptStart;
  let at = 0;
  let index = 0;
  const upTo = (to: number) => {
    while (index < input.markers.length && input.markers[index].start < to) {
      const marker = input.markers[index];
      if (marker.start > at) {
        pieces.push({ spanStart: kept, spanEnd: kept + (marker.start - at) });
        kept += marker.start - at;
      }
      pieces.push({ literal: latin1Bytes(input.texts[index]) });
      at = marker.end;
      index += 1;
    }
    if (to > at) {
      pieces.push({ spanStart: kept, spanEnd: kept + (to - at) });
      kept += to - at;
    }
    at = to;
  };
  upTo(blockStart);
  pieces.push(input.blockPiece);
  at = blockStart + input.blockLength;
  upTo(input.baseLength);
  return pieces.filter((piece) =>
    "literal" in piece ? piece.literal.length > 0 : piece.spanEnd > piece.spanStart,
  );
}

/** Put a moved picture back, proven by the exact rotation it undoes. */
export async function planMraiNativeEmbeddedImageReorderRestore(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly imageOrderAfter: number;
  /** The forward move's span; the rotation it wrote is one byte longer, and
   * longer again by the wrapper it put around a compensated block. */
  readonly spanStart: number;
  readonly spanEnd: number;
  readonly blockLength: number;
  readonly blockLeads: boolean;
  readonly compensation?: readonly number[] | null;
  /** The graphics states the forward move wrote inside the wrapper. */
  readonly carriedStates?: readonly string[] | null;
  readonly lift?: PdfNativeEmbeddedImageMoveLift | null;
  /** What the forward move recorded of the sections it took with the
   * picture: the places it emptied, and the bytes it left at the span's
   * outer edge. Undoing writes those markers back where they were. */
  readonly emptiedMarkers?: readonly Readonly<{ start: number; end: number }>[] | null;
  readonly stationaryBytes?: number | null;
  /** True when the move being undone was one step past one unit. */
  readonly stepped?: boolean;
}): Promise<PdfNativeEmbeddedImageReorderPlan> {
  const resolved = await resolveImageBlock(
    input.sourceBytes,
    input.pageNumber,
    input.imageOrder,
  );
  const { stream } = resolved;
  const streamBytes = stream.streamBytes;
  const compensation = input.compensation ?? null;
  const lift = input.lift ?? null;
  const carriedStates = input.carriedStates ?? EMPTY_CARRIED_PAINT_STATES;
  const namedProperties = resolved.hosted
    ? EMPTY_NAMED_PROPERTIES
    : await pageMarkedContentProperties(input.sourceBytes, resolved.paints.pageNumber);
  const plain = compensationWrapper(compensation, carriedStates);
  const prefix = plain ? plain.prefix : null;
  // The moved piece is the block, or the block inside its compensation
  // wrapper; a block that left its groups is the whole wrapper it is drawn
  // in now, which is a group that exists only for it.
  const pieceLength = lift
    ? resolved.block.end - resolved.block.start
    : input.blockLength + (prefix ? prefix.length + 2 : 0);
  const suffixLength = lift ? 2 + 4 * lift.marked : prefix ? 2 : 0;
  const prefixLength = lift
    ? pieceLength - input.blockLength - suffixLength
    : prefix
      ? prefix.length
      : 0;
  if (prefixLength < 0) {
    invalid("PDF image reorder history no longer covers this picture.");
  }
  const start = input.spanStart;
  const end =
    input.spanEnd +
    1 +
    (pieceLength - input.blockLength) -
    (input.emptiedMarkers ?? []).reduce((sum, range) => sum + (range.end - range.start), 0);
  if (
    start < 0 ||
    end > streamBytes.length ||
    end - start <= pieceLength + 1 ||
    input.blockLength <= 0
  ) {
    invalid("PDF image reorder history no longer fits this page.");
  }
  const emptied = input.emptiedMarkers ?? [];
  const stationary = input.stationaryBytes ?? 0;
  // The rotation is shorter than the span it replaced by the markers it
  // emptied, which it wrote around the picture instead.
  const emptiedLength = emptied.reduce((sum, range) => sum + (range.end - range.start), 0);
  const rotated = streamBytes.slice(start, end);
  const separator = input.blockLeads
    ? rotated.length - pieceLength - 1
    : pieceLength;
  if (rotated[separator] !== 0x0a) {
    invalid("PDF image reorder history does not match the current bytes.");
  }
  const movedPiece = input.blockLeads
    ? rotated.slice(separator + 1)
    : rotated.slice(0, pieceLength);
  const movedBlock = movedPiece.slice(prefixLength, movedPiece.length - suffixLength);
  const rest = input.blockLeads
    ? rotated.slice(0, separator)
    : rotated.slice(separator + 1);
  // The recorded span is an offset into the stream the move rewrote, so the
  // block it named is compared in that stream's own coordinates. A wrapper
  // is a group that exists only for the picture, so the block reads as it.
  const blockStart = resolved.block.start - stream.drawingSpan.start;
  if (
    blockStart !== start + (input.blockLeads ? separator + 1 : 0) ||
    resolved.block.end - resolved.block.start !== pieceLength
  ) {
    invalid("PDF image reorder history no longer covers this picture.");
  }
  // The markers the move emptied are written back where they were: the
  // sections it took are the ones its wrapper opens that may not be copied,
  // and each is closed by the `EMC` the wrapper writes for it.
  // Every section opens before the picture and closes after it, so the
  // emptied places hold the openings first, in the order the wrapper writes
  // them, and then one `EMC` each.
  const openings = emptied.length
    ? movedSectionOpeningTexts(latin1(movedPiece.slice(0, prefixLength)), lift?.marked ?? 0)
    : [];
  if (openings.length * 2 !== emptied.length) {
    invalid("PDF image reorder history does not name the sections it moved.");
  }
  const emptiedTexts = emptied.map((_range, index) =>
    index < openings.length ? openings[index] : "EMC",
  );
  const restoredLength = rest.length + input.blockLength + emptiedLength;
  const base = emptied.length
    ? expandPdfNativeReplacementPieces(
        rotated,
        restoredImageSpanPieces({
          keptStart: input.blockLeads ? 0 : separator + 1,
          blockPiece: input.blockLeads
            ? { spanStart: separator + 1 + prefixLength, spanEnd: rotated.length - suffixLength }
            : { spanStart: prefixLength, spanEnd: pieceLength - suffixLength },
          markers: emptied,
          texts: emptiedTexts,
          stationary,
          blockLength: input.blockLength,
          blockLeads: input.blockLeads,
          baseLength: restoredLength,
        }),
      )
    : input.blockLeads
      ? concatVectorBytes([movedBlock, rest])
      : concatVectorBytes([rest, movedBlock]);
  // The wrapper a lifted block went out in is the one its original place
  // gives it: for a front move the drawing before the span (the same bytes
  // now), for a back move the drawing the span passed.
  // A front move out from under its clips (leaveClips) wrote the wrapper
  // without the paths they clip by; the bytes say which one it wrote.
  const liftedWrapper = (dropClips: boolean): ImageMoveWrapper => {
    if (input.blockLeads && emptied.length === 0) {
      const drawingStart = stream.drawingSpan.start + start;
      return imageLiftWrapper(
        resolved.parsedOperations,
        latin1(resolved.paints.drawingBytes),
        0,
        { start: drawingStart, end: drawingStart + input.blockLength },
        true,
        lift!,
        compensation,
        carriedStates,
        namedProperties,
        dropClips,
      );
    } else if (input.blockLeads) {
      // The sections the move emptied stood inside the span, so the drawing
      // the wrapper was read from is the one before it and the span's own
      // head, put back together.
      const walkedText =
        latin1(resolved.paints.drawingBytes.slice(0, stream.drawingSpan.start + start)) +
        latin1(base.slice(0, stationary));
      let walked: readonly ContentOperation[];
      try {
        walked = operations(walkedText, true);
      } catch {
        invalid("PDF image reorder history does not undo the current bytes.");
      }
      return imageLiftWrapper(
        walked,
        walkedText,
        0,
        { start: walkedText.length, end: walkedText.length + input.blockLength },
        true,
        lift!,
        compensation,
        carriedStates,
        namedProperties,
        dropClips,
      );
    } else {
      const baseText = latin1(base);
      let baseOperations: readonly ContentOperation[];
      try {
        baseOperations = operations(baseText, true);
      } catch {
        invalid("PDF image reorder history does not undo the current bytes.");
      }
      return imageLiftWrapper(
        baseOperations,
        baseText,
        0,
        { start: rest.length, end: base.length },
        false,
        lift!,
        null,
        EMPTY_CARRIED_PAINT_STATES,
        namedProperties,
        dropClips,
      );
    }
  };
  const rotatedWith = (wrapper: ImageMoveWrapper | null) => latin1(
    rotatedImageSpanBytes(
      base,
      input.blockLength,
      input.blockLeads,
      wrapper,
      emptied,
      stationary,
    ),
  ) === latin1(rotated);
  if (
    !(lift
      ? rotatedWith(liftedWrapper(false)) || (input.blockLeads && rotatedWith(liftedWrapper(true)))
      : rotatedWith(plain))
  ) {
    invalid("PDF image reorder history does not undo the current bytes.");
  }
  // The block and the rest it went past are both still in the span.
  const blockPiece = input.blockLeads
    ? { spanStart: separator + 1 + prefixLength, spanEnd: rotated.length - suffixLength }
    : { spanStart: prefixLength, spanEnd: pieceLength - suffixLength };
  const restPiece = input.blockLeads
    ? { spanStart: 0, spanEnd: separator }
    : { spanStart: separator + 1, spanEnd: rotated.length };
  const request = await imageReorderRequest({
    stream,
    spanStart: start,
    spanEnd: end,
    pieces: emptied.length
      ? restoredImageSpanPieces({
          keptStart: input.blockLeads ? 0 : separator + 1,
          blockPiece,
          markers: emptied,
          texts: emptiedTexts,
          stationary,
          blockLength: input.blockLength,
          blockLeads: input.blockLeads,
          baseLength: base.length,
        })
      : input.blockLeads
        ? [blockPiece, restPiece]
        : [restPiece, blockPiece],
  });
  if (latin1(request.replacement) !== latin1(base)) {
    invalid("PDF image reorder history does not undo the current bytes.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
    direction: "restore",
    stepped: input.stepped ?? false,
    pageNumber: resolved.paints.pageNumber,
    imageOrder: input.imageOrder,
    imageOrderAfter: input.imageOrderAfter,
    resourceName: resolved.paint.resourceName,
    page: resolved.paints.page,
    resourceOwner: stream.resourceOwner ?? resolved.paints.resourceOwner,
    image: resolved.paint.image,
    contentStream: stream.contentStream,
    operatorSpan: request.operatorSpan,
    blockLength: input.blockLength,
    blockLeads: input.blockLeads,
    ...(compensation ? { compensation } : {}),
    ...(carriedStates.length > 0 ? { carriedStates } : {}),
    ...(lift ? { lift } : {}),
    ...(emptied.length > 0 ? { emptiedMarkers: emptied, stationaryBytes: stationary } : {}),
    request,
  });
}

export async function prepareMraiNativeEmbeddedImageReorder(
  sourceBytes: Uint8Array,
  plan: PdfNativeEmbeddedImageReorderPlan,
): Promise<PdfNativePreparedContentStreamTransaction> {
  if (
    plan.schemaVersion !== MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION ||
    plan.blockLength <= 0 ||
    plan.request.operatorSpan.end - plan.request.operatorSpan.start <=
      plan.blockLength
  ) {
    invalid("PDF image reorder plan authority is inconsistent.");
  }
  return prepareMraiNativeContentStreamTransaction(sourceBytes, plan.request);
}

export async function verifyMraiNativeEmbeddedImageReorderCandidate(
  sourceBytes: Uint8Array,
  plan: PdfNativeEmbeddedImageReorderPlan,
  prepared: PdfNativePreparedContentStreamTransaction,
): Promise<
  Readonly<{ verified: true; imageOrder: number; candidateSha256: string }>
> {
  await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
  const before = await inspectMraiNativeEmbeddedImagePagePaints(
    sourceBytes,
    plan.pageNumber,
  );
  const reopened = await inspectMraiNativeEmbeddedImagePagePaints(
    prepared.candidateBytes,
    plan.pageNumber,
  );
  const paint = reopened.paints[plan.imageOrderAfter];
  if (
    reopened.paints.length !== before.paints.length ||
    !paint ||
    paint.resourceName !== plan.resourceName
  ) {
    prepared.candidateBytes.fill(0);
    throw new PdfNativeTransactionError(
      "verification-failed",
      "Reopened PDF picture order does not match the requested move.",
    );
  }
  return Object.freeze({
    verified: true,
    imageOrder: plan.imageOrderAfter,
    candidateSha256: prepared.candidateSha256,
  });
}


/** Text state a group keeps (it is part of the graphics state, set inside a
 * text object or out of one): written again when the group is reopened. */
const CLIP_ESCAPE_TEXT_STATE_OPERATORS = new Set(["Tc", "Tw", "Tz", "TL", "Tf", "Ts", "Tr"]);
/** What draws or marks and leaves no state behind. */
const CLIP_ESCAPE_INERT_OPERATORS = new Set([
  "Do", "sh", "BI", "MP", "DP", "BDC", "BMC", "EMC", "BX", "EX",
]);
/** Text operators that position or show text: no state a group keeps past
 * its text object. */
const CLIP_ESCAPE_TEXT_INERT_OPERATORS = new Set([
  "Td", "TD", "Tm", "T*", "Tj", "TJ", "'", "\"", "d0", "d1",
]);















function imageClipEscapeInPlace(
  parsedOperations: readonly ContentOperation[],
  drawingText: string,
  block: Readonly<{ start: number; end: number }>,
): Readonly<{ prefix: string; suffix: string; closed: number }> | null {
  type Frame = { ops: string[]; carried: string[]; clips: boolean };
  const frames: Frame[] = [{ ops: [], carried: [], clips: false }];
  let path: string[] = [];
  let clip = false;
  let inText = false;
  for (const operation of parsedOperations) {
    if (operation.start >= block.start) break;
    const operator = operation.operator;
    const text = drawingText.slice(operation.start, operation.end);
    const top = frames[frames.length - 1];
    if (CLIP_ESCAPE_TEXT_STATE_OPERATORS.has(operator)) {
      if (operator === "Tr" && Number(operation.operands[0]?.value) >= 4) return null;
      top.ops.push(text);
      continue;
    }
    if (inText) {
      if (operator === "ET") inText = false;
      else if (!CLIP_ESCAPE_TEXT_INERT_OPERATORS.has(operator) && !CLIP_ESCAPE_INERT_OPERATORS.has(operator) &&
        !IMAGE_LIFT_STATE_OPERATORS.has(operator)) return null;
      else if (IMAGE_LIFT_STATE_OPERATORS.has(operator)) top.ops.push(text);
      continue;
    }
    if (operator === "BT") {
      inText = true;
    } else if (operator === "q") {
      frames.push({ ops: [], carried: [], clips: false });
    } else if (operator === "Q") {
      if (frames.length < 2) return null;
      frames.pop();
    } else if (PDF_PATH_CONSTRUCTION_OPERATORS.has(operator)) {
      path.push(text);
    } else if (operator === "W" || operator === "W*") {
      clip = true;
      path.push(text);
    } else if (operator === "n") {
      if (clip) {
        top.ops.push(...path, text);
        top.clips = true;
      }
      path = [];
      clip = false;
    } else if (paintMode(operator) !== null) {
      if (clip) return null;
      path = [];
    } else if (IMAGE_LIFT_STATE_OPERATORS.has(operator)) {
      top.ops.push(text);
      top.carried.push(text);
    } else if (!CLIP_ESCAPE_INERT_OPERATORS.has(operator)) {
      return null;
    }
  }
  if (inText || path.length > 0 || frames[0].clips) return null;
  const outermost = frames.findIndex((frame, index) => index > 0 && frame.clips);
  if (outermost < 0) return null;
  const closed = frames.slice(outermost);
  const prefix = `${"Q\n".repeat(closed.length)}q\n${closed.flatMap((frame) => frame.carried).map((op) => `${op}\n`).join("")}`;
  const suffix = `\nQ\n${closed.map((frame) => `q\n${frame.ops.map((op) => `${op}\n`).join("")}`).join("")}`;
  return Object.freeze({ prefix, suffix, closed: closed.length });
}

/**
 * A picture moved past the clips it is drawn under drawn whole where it is
 * (imageClipEscapeInPlace), as an incremental revision of its content
 * stream: published as its bytes, undone by restoring that object.
 * Refused (as a picture no clip holds back) when it cannot be.
 */
export async function planMraiNativeEmbeddedImageClipEscape(input: {
  readonly sourceBytes: Uint8Array;
  readonly pageNumber: number;
  readonly imageOrder: number;
}): Promise<Readonly<{ candidateBytes: Uint8Array; candidateSha256: string }>> {
  const resolved = await resolveImageBlock(input.sourceBytes, input.pageNumber, input.imageOrder);
  if (resolved.hosted) refuse(MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED);
  const drawingText = latin1(resolved.paints.drawingBytes);
  const escape = imageClipEscapeInPlace(resolved.parsedOperations, drawingText, resolved.block);
  if (!escape) refuse(MRAI_NATIVE_EMBEDDED_IMAGE_NOT_CLIPPED);
  const { stream, block } = resolved;
  const start = block.start - stream.drawingSpan.start;
  const end = block.end - stream.drawingSpan.start;
  const blockBytes = resolved.paints.drawingBytes.slice(block.start, block.end);
  return planMraiNativeContentStreamSpanRevision({
    sourceBytes: input.sourceBytes,
    stream: stream.contentStream,
    start,
    end,
    expectedSpanSha256: await sha256PdfBytes(blockBytes),
    replacement: concatVectorBytes([latin1Bytes(escape.prefix), blockBytes, latin1Bytes(escape.suffix)]),
  });
}
