/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  type PdfSemanticAppearanceCommandKind,
  type PdfSemanticAppearanceCommandTarget,
  type PdfSemanticAppearanceRouteFailureReason,
  type PdfSemanticAppearanceSelectionRoute,
  type PdfSemanticAppearanceUnresolvedRoute,
  type PdfSemanticAppearanceVectorExclusion,
} from "./semantic-appearance-command-routing";
import type { PdfSemanticAppearanceOperatorIds } from "./semantic-appearance-reconciliation";
import {
  isPdfStandardBlendMode,
  type PdfStandardBlendMode,
} from "./blend-modes";
import type {
  PdfNativeVectorCmyk,
  PdfNativeVectorDeviceCmyk,
  PdfNativeVectorObjectStyle,
} from "./engine/native/vector-style-transaction";
import type { PdfSemanticAppearanceSelectionSourceEvidence } from "./semantic-appearance-native-pages";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";
import type {
  PdfVectorContentItem,
  PdfVectorPaintMode,
  PdfVectorSimplePathCommand,
} from "./vector-content-inventory";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
import type { PdfNativePlanners } from "./remote/pdf-native-planners";

export const PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION =
  "mrai.pdf-semantic-appearance-inspection.v3" as const;
export const PDF_SEMANTIC_APPEARANCE_INIT_TIMEOUT_MS = 30_000;



export function pdfSemanticAppearanceLocalEvidenceEnabled(): boolean {
  const flag = process.env.NEXT_PUBLIC_MRAI_PDF_BROWSER_PREPARE;
  if (process.env.NODE_ENV === "test" && flag !== "1") return false;
  return flag !== "0";
}
/** One chunk of the source's selection evidence, well under the engine
 * call's 8 MiB answer bound. */
const PDF_SEMANTIC_EVIDENCE_CHUNK_BYTES = 4 * 1024 * 1024;
export const PDF_SEMANTIC_APPEARANCE_PAGE_TIMEOUT_MS = 5_000;





export const PDF_SEMANTIC_APPEARANCE_PAGE_RETRY_FACTOR = 3;
export const PDF_SEMANTIC_APPEARANCE_PAGE_RETRY_LIMIT = 2;
const MIN_INIT_TIMEOUT_MS = 1_000;
const MAX_INIT_TIMEOUT_MS = 600_000;
/** A page drawing tens of thousands of paths (a wiring diagram) is routed in
 * the worker in time linear in its paths: past the first ten thousand each
 * one extends the page's deadline, so the page is not failed as hung. */
const PAGE_TIMEOUT_FREE_VECTORS = 10_000;
const PAGE_TIMEOUT_MS_PER_VECTOR = 0.25;

export function pdfSemanticAppearancePageTimeoutMs(visibleVectorCount: number, priorTimeouts = 0): number {
  const factor = PDF_SEMANTIC_APPEARANCE_PAGE_RETRY_FACTOR **
    Math.min(Math.max(0, Math.floor(priorTimeouts)), PDF_SEMANTIC_APPEARANCE_PAGE_RETRY_LIMIT);
  return factor * pdfSemanticAppearancePageBaseTimeoutMs(visibleVectorCount);
}

function pdfSemanticAppearancePageBaseTimeoutMs(visibleVectorCount: number): number {
  return (
    PDF_SEMANTIC_APPEARANCE_PAGE_TIMEOUT_MS +
    Math.ceil(
      Math.max(0, visibleVectorCount - PAGE_TIMEOUT_FREE_VECTORS) *
        PAGE_TIMEOUT_MS_PER_VECTOR,
    )
  );
}

const SOURCE_MAX_BYTES = PDF_EDITABLE_MAX_BYTES;
const MAX_SOURCE_GENERATION_LENGTH = 512;
const MAX_TARGETS_PER_PAGE = 16_384;
const MAX_RESULT_ROWS_PER_PAGE = 16_384;
const MAX_OPERATORS_PER_PAGE = 1_000_000;
const MAX_VECTORS_PER_PAGE = 100_000;
// Mirror of the inventory: outlined lettering writes a line of glyphs as one
// path of hundreds of commands, and a longer path failed the whole page.
const MAX_SIMPLE_SUBPATH_COMMANDS = 4096;
const MAX_PROOF_ID_LENGTH = 2_048;
const MAX_ABS_GEOMETRY_VALUE = 10_000_000;
const SHA256_PATTERN = /^[0-9a-f]{64}$/;
const ANNOTATION_REF_PATTERN = /^[1-9][0-9]{0,9}R[0-9]{0,5}$/;

export interface PdfSemanticAppearancePageOperatorEvidence {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly operatorIds: PdfSemanticAppearanceOperatorIds;
  readonly operatorList: Readonly<{
    readonly fnArray: readonly number[];
    /** Only sanitized beginAnnotation tuples survive; every other row is null. */
    readonly argsArray: readonly unknown[];
  }>;
  readonly visibleVectors: readonly PdfVectorContentItem[];
}

export type PdfNativeVectorSelectionProofFamily = "shape" | "subpath";
export type PdfNativeVectorSelectionKind =
  "line" | "rectangle" | "ellipse" | "subpath";

export interface PdfNativeVectorStyleSelectionEvidence {
  readonly schemaVersion: "mrai.native-vector-object-style.v1";
  readonly proofId: string;
  readonly proofDigest: string;
  readonly pageProofIndex: number;
  readonly style: PdfNativeVectorObjectStyle;
  /** The paint's DeviceCMYK ink, where it was set as such (a copy pasted
   * elsewhere says it again, not the naive RGB in `style`). */
  readonly deviceCmyk?: PdfNativeVectorDeviceCmyk;
}

/**
 * Source-bound selection evidence only. The native writer must later re-open
 * the current bytes and independently re-prove every locator and digest.
 */
export interface PdfNativeVectorSelectionRoute {
  readonly selectionOnly: true;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly rendererOperatorIndex: number;
  readonly proofFamily: PdfNativeVectorSelectionProofFamily;
  readonly nativeKind: PdfNativeVectorSelectionKind;
  readonly paintMode: PdfVectorPaintMode;
  readonly proofId: string;
  readonly proofDigest: string;
  readonly contentStreamObjectNumber: number;
  readonly contentStreamGeneration: number;
  readonly operatorStart: number;
  readonly operatorEnd: number;
  /** Which drawing of a Form the page shows more than once the proof reads
   * (part of its id). */
  readonly formInstance?: string;
  readonly style?: PdfNativeVectorStyleSelectionEvidence;
  /** How much a line width of the drawing's own space grows on the page
   * (its paint matrix's scale; absent: 1). The shape tab shows and takes
   * widths in points on paper (user audit PDF 0058). */
  readonly lineWidthScale?: number;
  readonly bounds: Readonly<{
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
}

export interface PdfSemanticAppearanceWorkerInitRequest {
  readonly protocolVersion: typeof PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION;
  readonly kind: "init";
  readonly sessionId: string;
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  readonly sourceBytes: ArrayBuffer;
  


  readonly localEvidence?: boolean;
}

/** The source's owners and native vector pages, read by the engine host
 * after `init` (`null`: the engine could not read them; the session fails). */
export interface PdfSemanticAppearanceWorkerNativeEvidenceRequest {
  readonly protocolVersion: typeof PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION;
  readonly kind: "native-evidence";
  readonly sessionId: string;
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  readonly nativeEvidence: PdfSemanticAppearanceSelectionSourceEvidence | null;
}

export interface PdfSemanticAppearanceWorkerRoutePageRequest {
  readonly protocolVersion: typeof PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION;
  readonly kind: "route-page";
  readonly sessionId: string;
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  readonly requestId: string;
  readonly pageNumber: number;
  readonly targets: readonly PdfSemanticAppearanceCommandTarget[];
  readonly operatorEvidence?: PdfSemanticAppearancePageOperatorEvidence;
}

export interface PdfSemanticAppearanceWorkerDestroyRequest {
  readonly protocolVersion: typeof PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION;
  readonly kind: "destroy";
  readonly sessionId: string;
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
}

export type PdfSemanticAppearanceWorkerRequest =
  | PdfSemanticAppearanceWorkerInitRequest
  | PdfSemanticAppearanceWorkerNativeEvidenceRequest
  | PdfSemanticAppearanceWorkerRoutePageRequest
  | PdfSemanticAppearanceWorkerDestroyRequest;

interface PdfSemanticAppearanceWorkerResponseBase {
  readonly protocolVersion: typeof PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION;
  readonly sessionId: string;
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
}

export interface PdfSemanticAppearanceWorkerReadyResponse extends PdfSemanticAppearanceWorkerResponseBase {
  readonly kind: "ready";
  /** False only for a local-evidence session: the worker keeps the source to
   * read its native vector pages from, and wipes it on destroy. */
  readonly sourceBytesZeroized: boolean;
}

export interface PdfSemanticAppearanceWorkerPageRoutesResponse extends PdfSemanticAppearanceWorkerResponseBase {
  readonly kind: "page-routes";
  readonly requestId: string;
  readonly pageNumber: number;
  readonly routes: readonly PdfSemanticAppearanceSelectionRoute[];
  readonly unresolved: readonly PdfSemanticAppearanceUnresolvedRoute[];
  readonly vectorExclusions: readonly PdfSemanticAppearanceVectorExclusion[];
  readonly nativeVectorRoutes: readonly PdfNativeVectorSelectionRoute[];
}

export interface PdfSemanticAppearanceWorkerErrorResponse extends PdfSemanticAppearanceWorkerResponseBase {
  readonly kind: "error";
  readonly requestId?: string;
  readonly code: "invalid-request" | "inspection-failed" | "routing-failed";
}

export type PdfSemanticAppearanceWorkerResponse =
  | PdfSemanticAppearanceWorkerReadyResponse
  | PdfSemanticAppearanceWorkerPageRoutesResponse
  | PdfSemanticAppearanceWorkerErrorResponse;

export interface PdfSemanticAppearanceInspectionWorkerLike {
  onmessage:
    ((event: MessageEvent<PdfSemanticAppearanceWorkerResponse>) => void) | null;
  onerror: ((event: Event) => void) | null;
  postMessage(
    message: PdfSemanticAppearanceWorkerRequest,
    transfer?: readonly Transferable[],
  ): void;
  terminate(): void;
}

export interface PdfSemanticAppearanceInspectionInput {
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  readonly sourceBytes: Uint8Array;
  /** The planners of the page the bytes are shown on (a reader's own); the
   * engine reads the source's vector evidence with them. Absent: the set
   * bound to no page, refused when two pages hold these bytes. */
  readonly planners?: PdfNativePlanners;
  /** How long the worker may take to accept the file (default
   * PDF_SEMANTIC_APPEARANCE_INIT_TIMEOUT_MS); a retry asks for more. */
  readonly initTimeoutMs?: number;
}

export interface PdfSemanticAppearanceInspectionPageInput {
  readonly pageNumber: number;
  readonly targets: readonly PdfSemanticAppearanceCommandTarget[];
  readonly operatorEvidence?: PdfSemanticAppearancePageOperatorEvidence;
}

export interface PdfSemanticAppearanceInspectionPageResult {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly routes: readonly PdfSemanticAppearanceSelectionRoute[];
  readonly unresolved: readonly PdfSemanticAppearanceUnresolvedRoute[];
  readonly vectorExclusions: readonly PdfSemanticAppearanceVectorExclusion[];
  readonly nativeVectorRoutes: readonly PdfNativeVectorSelectionRoute[];
}

export interface PdfSemanticAppearanceInspectionReadyReceipt {
  readonly sourceSha256: string;
  /** False only for a local-evidence session: the worker keeps the source to
   * read its native vector pages from, and wipes it on destroy. */
  readonly sourceBytesZeroized: boolean;
}

export interface PdfSemanticAppearanceInspectionSession {
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  ready(): Promise<PdfSemanticAppearanceInspectionReadyReceipt>;
  routePage(
    input: PdfSemanticAppearanceInspectionPageInput,
  ): Promise<PdfSemanticAppearanceInspectionPageResult>;
  destroy(): void;
}

export type PdfSemanticAppearanceInspectionErrorCode =
  | "worker-unavailable"
  | "worker-failed"
  | "invalid-source"
  | "invalid-request"
  | "init-timeout"
  | "page-timeout"
  | "invalid-response"
  /** One page's answer could not be taken (or its routing failed): that
   * page has no routes; the session and every other page go on. */
  | "page-refused"
  | "superseded"
  | "destroyed";









export function pdfSemanticAppearanceTargetsOnFilePage<T extends Readonly<{ pageNumber: number }>>(
  targets: readonly T[], shownPage: number, filePage: number): readonly T[] {
  if (filePage === shownPage) return targets;
  return Object.freeze(targets.map((target) => (target.pageNumber === shownPage
    ? Object.freeze({ ...target, pageNumber: filePage }) : target)));
}

export class PdfSemanticAppearanceInspectionError extends Error {
  constructor(readonly code: PdfSemanticAppearanceInspectionErrorCode) {
    super(`PDF semantic appearance inspection failed: ${code}.`);
    this.name = "PdfSemanticAppearanceInspectionError";
  }
}

type WorkerFactory = () => PdfSemanticAppearanceInspectionWorkerLike;

let testWorkerFactory: WorkerFactory | null = null;

export function setPdfSemanticAppearanceWorkerFactoryForTests(
  factory: WorkerFactory | null,
): void {
  testWorkerFactory = factory;
}

function defaultWorkerFactory(): PdfSemanticAppearanceInspectionWorkerLike {
  if (typeof Worker === "undefined") {
    throw new PdfSemanticAppearanceInspectionError("worker-unavailable");
  }
  return new Worker(
    new URL("./semantic-appearance-inspection.worker.ts", import.meta.url),
    { type: "module", name: "mrnote-pdf-semantic-appearance" },
  ) as PdfSemanticAppearanceInspectionWorkerLike;
}

function sessionId(): string {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  return `pdf-semantic-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function exactSha256(value: unknown): string {
  const normalized =
    typeof value === "string" ? value.trim().toLowerCase() : "";
  if (!SHA256_PATTERN.test(normalized)) {
    throw new PdfSemanticAppearanceInspectionError("invalid-source");
  }
  return normalized;
}

function exactSourceGeneration(value: unknown): string {
  const normalized = typeof value === "string" ? value.trim() : "";
  if (
    normalized.length === 0 ||
    normalized.length > MAX_SOURCE_GENERATION_LENGTH
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-source");
  }
  return normalized;
}

function isExactInteger(value: unknown, minimum: number): value is number {
  return Number.isSafeInteger(value) && (value as number) >= minimum;
}

function exactFiniteNumber(value: unknown): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    Math.abs(value) > MAX_ABS_GEOMETRY_VALUE
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  return Object.is(value, -0) ? 0 : value;
}

function exactTuple(value: unknown, length: number): readonly number[] | null {
  if (
    (!Array.isArray(value) &&
      (!ArrayBuffer.isView(value) || value instanceof DataView)) ||
    (value as ArrayLike<unknown>).length !== length
  ) {
    return null;
  }
  try {
    return Object.freeze(
      Array.from(value as ArrayLike<unknown>, exactFiniteNumber),
    );
  } catch {
    return null;
  }
}

function sanitizedAnnotationPayload(value: unknown): readonly unknown[] {
  const raw = Array.isArray(value) ? value : [];
  const rawId = raw[0];
  const id =
    typeof rawId === "string" && ANNOTATION_REF_PATTERN.test(rawId)
      ? rawId
      : null;
  return Object.freeze([
    id,
    exactTuple(raw[1], 4),
    exactTuple(raw[2], 6),
    exactTuple(raw[3], 6),
    typeof raw[4] === "boolean" ? raw[4] : null,
  ]);
}

function exactPageGeometry(
  value: PdfVectorPageGeometry,
): PdfVectorPageGeometry {
  const crop = value?.cropBox;
  if (
    !crop ||
    ![0, 90, 180, 270].includes(value.rotation) ||
    !Number.isFinite(crop.xMin) ||
    !Number.isFinite(crop.yMin) ||
    !Number.isFinite(crop.xMax) ||
    !Number.isFinite(crop.yMax) ||
    crop.xMax <= crop.xMin ||
    crop.yMax <= crop.yMin
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  return Object.freeze({
    cropBox: Object.freeze({
      xMin: exactFiniteNumber(crop.xMin),
      yMin: exactFiniteNumber(crop.yMin),
      xMax: exactFiniteNumber(crop.xMax),
      yMax: exactFiniteNumber(crop.yMax),
    }),
    rotation: value.rotation,
  });
}

function exactPoint(value: unknown): Readonly<{ x: number; y: number }> {
  const point = value as { x?: unknown; y?: unknown } | null;
  if (!point || typeof point !== "object") {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  return Object.freeze({
    x: exactFiniteNumber(point.x),
    y: exactFiniteNumber(point.y),
  });
}

function exactPathCommand(value: unknown): PdfVectorSimplePathCommand {
  const command = value as Record<string, unknown> | null;
  if (!command || typeof command !== "object") {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  if (command.kind === "moveTo" || command.kind === "lineTo") {
    return Object.freeze({
      kind: command.kind,
      point: exactPoint(command.point),
    });
  }
  if (command.kind === "curveTo") {
    return Object.freeze({
      kind: "curveTo",
      control1: exactPoint(command.control1),
      control2: exactPoint(command.control2),
      end: exactPoint(command.end),
    });
  }
  if (command.kind === "quadraticCurveTo") {
    return Object.freeze({
      kind: "quadraticCurveTo",
      control: exactPoint(command.control),
      end: exactPoint(command.end),
    });
  }
  if (command.kind === "closePath") {
    return Object.freeze({ kind: "closePath" });
  }
  throw new PdfSemanticAppearanceInspectionError("invalid-request");
}

function exactVisibleVector(value: PdfVectorContentItem): PdfVectorContentItem {
  if (
    !value ||
    typeof value !== "object" ||
    (value.kind !== "rule" &&
      value.kind !== "rect-like" &&
      value.kind !== "generic-path") ||
    (value.paintMode !== "stroke" &&
      value.paintMode !== "fill" &&
      value.paintMode !== "fill-stroke") ||
    !isExactInteger(value.operatorIndex, 0) ||
    !isExactInteger(value.pathValueCount, 0) ||
    value.pathValueCount > 100_000 ||
    typeof value.rotatedOrSkewed !== "boolean" ||
    typeof value.simplePathEvidence !== "boolean"
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  const x = exactFiniteNumber(value.x);
  const y = exactFiniteNumber(value.y);
  const width = exactFiniteNumber(value.width);
  const height = exactFiniteNumber(value.height);
  const angle = exactFiniteNumber(value.angle);
  if (
    x < 0 ||
    y < 0 ||
    width <= 0 ||
    height <= 0 ||
    x + width > 1.000_001 ||
    y + height > 1.000_001
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  const simpleLineEndpoints = value.simpleLineEndpoints
    ? Object.freeze({
        start: exactPoint(value.simpleLineEndpoints.start),
        end: exactPoint(value.simpleLineEndpoints.end),
      })
    : undefined;
  const simpleSubpath = value.simpleSubpath
    ? (() => {
        if (
          !Array.isArray(value.simpleSubpath.commands) ||
          value.simpleSubpath.commands.length < 2 ||
          value.simpleSubpath.commands.length > MAX_SIMPLE_SUBPATH_COMMANDS ||
          typeof value.simpleSubpath.closed !== "boolean"
        ) {
          throw new PdfSemanticAppearanceInspectionError("invalid-request");
        }
        return Object.freeze({
          commands: Object.freeze(
            value.simpleSubpath.commands.map(exactPathCommand),
          ),
          closed: value.simpleSubpath.closed,
        });
      })()
    : undefined;
  return Object.freeze({
    kind: value.kind,
    paintMode: value.paintMode,
    operatorIndex: value.operatorIndex,
    x,
    y,
    width,
    height,
    angle,
    rotatedOrSkewed: value.rotatedOrSkewed,
    pathValueCount: value.pathValueCount,
    simplePathEvidence: value.simplePathEvidence,
    ...(simpleLineEndpoints ? { simpleLineEndpoints } : {}),
    ...(simpleSubpath ? { simpleSubpath } : {}),
  });
}

/** Strip raw pdf.js arguments down to geometry-only selection evidence. */
export function createPdfSemanticAppearancePageOperatorEvidence(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly operatorIds: PdfSemanticAppearanceOperatorIds;
  readonly operatorList: Readonly<{
    readonly fnArray: ArrayLike<number>;
    readonly argsArray: ArrayLike<unknown>;
  }>;
  readonly visibleVectors: readonly PdfVectorContentItem[];
}): PdfSemanticAppearancePageOperatorEvidence {
  const { beginAnnotation, endAnnotation } = input.operatorIds;
  if (
    !isExactInteger(beginAnnotation, 0) ||
    !isExactInteger(endAnnotation, 0) ||
    beginAnnotation === endAnnotation ||
    !input.operatorList ||
    input.operatorList.fnArray.length !== input.operatorList.argsArray.length ||
    input.operatorList.fnArray.length > MAX_OPERATORS_PER_PAGE ||
    !Array.isArray(input.visibleVectors) ||
    input.visibleVectors.length > MAX_VECTORS_PER_PAGE
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  const fnArray = Object.freeze(
    Array.from(input.operatorList.fnArray, (value) => {
      if (!isExactInteger(value, 0)) {
        throw new PdfSemanticAppearanceInspectionError("invalid-request");
      }
      return value;
    }),
  );
  const argsArray = Object.freeze(
    fnArray.map((fn, index) =>
      fn === beginAnnotation
        ? sanitizedAnnotationPayload(input.operatorList.argsArray[index])
        : null,
    ),
  );
  const operatorIndexes = new Set<number>();
  const visibleVectors = Object.freeze(
    input.visibleVectors.map((vector) => {
      const sanitized = exactVisibleVector(vector);
      if (operatorIndexes.has(sanitized.operatorIndex)) {
        throw new PdfSemanticAppearanceInspectionError("invalid-request");
      }
      operatorIndexes.add(sanitized.operatorIndex);
      return sanitized;
    }),
  );
  return Object.freeze({
    pageGeometry: exactPageGeometry(input.pageGeometry),
    operatorIds: Object.freeze({ beginAnnotation, endAnnotation }),
    operatorList: Object.freeze({ fnArray, argsArray }),
    visibleVectors,
  });
}

function surfaceForKind(
  kind: PdfSemanticAppearanceCommandKind,
): PdfSemanticAppearanceCommandTarget["surface"] {
  if (
    kind === "text" ||
    kind === "checkbox" ||
    kind === "radio" ||
    kind === "dropdown" ||
    kind === "listbox"
  ) {
    return "form-field";
  }
  if (kind === "Link") return "native-link";
  if (kind === "Text") return "native-note";
  if (kind === "Highlight" || kind === "Underline") {
    return "native-text-markup";
  }
  return "native-shape";
}

function routeIdentityKey(route: {
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly semanticKind: PdfSemanticAppearanceCommandKind;
}): string {
  return `${route.pageNumber}:${route.objectNumber}:${route.generation}:${route.semanticKind}`;
}

function isSemanticKind(
  value: unknown,
): value is PdfSemanticAppearanceCommandKind {
  return (
    value === "text" ||
    value === "checkbox" ||
    value === "radio" ||
    value === "dropdown" ||
    value === "listbox" ||
    value === "Link" ||
    value === "Text" ||
    value === "Highlight" ||
    value === "Underline" ||
    value === "Square"
  );
}

function sanitizeTarget(
  target: PdfSemanticAppearanceCommandTarget,
  sourceSha256: string,
  pageNumber: number,
): PdfSemanticAppearanceCommandTarget {
  if (
    !target ||
    typeof target !== "object" ||
    exactSha256(target.sourceSha256) !== sourceSha256 ||
    target.pageNumber !== pageNumber ||
    !isExactInteger(target.objectNumber, 1) ||
    !isExactInteger(target.generation, 0) ||
    !isSemanticKind(target.semanticKind) ||
    target.surface !== surfaceForKind(target.semanticKind)
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-request");
  }
  return Object.freeze({
    sourceSha256,
    pageNumber,
    objectNumber: target.objectNumber,
    generation: target.generation,
    semanticKind: target.semanticKind,
    surface: target.surface,
  });
}

const ROUTE_FAILURE_REASONS = new Set<PdfSemanticAppearanceRouteFailureReason>([
  "duplicate-owner",
  "target-missing",
  "malformed-target",
  "stale-target",
  "target-kind-mismatch",
  "duplicate-target",
  "appearance-evidence-missing",
  "instance-evidence-mismatch",
  "ambiguous-instance",
  "owner-missing",
]);
const INSTANCE_DISPOSITIONS = new Set([
  "semantic-owner-eligible",
  "semantic-owner-preserve-only",
  "preserve-only",
] as const);

function sanitizeRoute(
  route: PdfSemanticAppearanceSelectionRoute,
  sourceSha256: string,
  pageNumber: number,
): PdfSemanticAppearanceSelectionRoute {
  if (
    !route ||
    typeof route !== "object" ||
    route.selectionOnly !== true ||
    route.sourceSha256 !== sourceSha256 ||
    route.pageNumber !== pageNumber ||
    !isExactInteger(route.objectNumber, 1) ||
    !isExactInteger(route.generation, 0) ||
    !isSemanticKind(route.semanticKind) ||
    route.surface !== surfaceForKind(route.semanticKind) ||
    (route.ownerDisposition !== "exact-owner" &&
      route.ownerDisposition !== "shared-appearance-preserve-only") ||
    !(
      (route.evidenceSource === "source-owner-only" &&
        route.instanceDisposition === null) ||
      (route.evidenceSource === "annotation-instance" &&
        route.instanceDisposition !== null &&
        INSTANCE_DISPOSITIONS.has(route.instanceDisposition))
    )
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  return Object.freeze({
    selectionOnly: true,
    sourceSha256,
    pageNumber,
    objectNumber: route.objectNumber,
    generation: route.generation,
    semanticKind: route.semanticKind,
    surface: route.surface,
    ownerDisposition: route.ownerDisposition,
    evidenceSource: route.evidenceSource,
    instanceDisposition: route.instanceDisposition,
  });
}

function exactRouteBounds(value: unknown) {
  const bounds = value as {
    x?: unknown;
    y?: unknown;
    width?: unknown;
    height?: unknown;
  } | null;
  if (!bounds || typeof bounds !== "object") {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const x = exactFiniteNumber(bounds.x);
  const y = exactFiniteNumber(bounds.y);
  const width = exactFiniteNumber(bounds.width);
  const height = exactFiniteNumber(bounds.height);
  if (
    x < 0 ||
    y < 0 ||
    width <= 0 ||
    height <= 0 ||
    x + width > 1.000_001 ||
    y + height > 1.000_001
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  return Object.freeze({ x, y, width, height });
}

function sanitizeVectorExclusion(
  exclusion: PdfSemanticAppearanceVectorExclusion,
  pageNumber: number,
): PdfSemanticAppearanceVectorExclusion {
  if (
    !exclusion ||
    typeof exclusion !== "object" ||
    exclusion.pageNumber !== pageNumber ||
    !isExactInteger(exclusion.rendererOperatorIndex, 0) ||
    exclusion.excludedFromGenericVectorPool !== true ||
    !INSTANCE_DISPOSITIONS.has(exclusion.disposition)
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  return Object.freeze({
    pageNumber,
    rendererOperatorIndex: exclusion.rendererOperatorIndex,
    excludedFromGenericVectorPool: true,
    disposition: exclusion.disposition,
  });
}

function plainDataRecord(
  value: unknown,
): Readonly<Record<string, unknown>> | null {
  try {
    if (!value || typeof value !== "object" || Array.isArray(value))
      return null;
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) return null;
    const descriptors = Object.getOwnPropertyDescriptors(value);
    const result: Record<string, unknown> = {};
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== "string") return null;
      const descriptor = descriptors[key];
      if (!descriptor?.enumerable || !("value" in descriptor)) return null;
      result[key] = descriptor.value;
    }
    return Object.freeze(result);
  } catch {
    return null;
  }
}

function exactNativeVectorStyleColor(
  value: unknown,
): Readonly<{ r: number; g: number; b: number }> | null {
  if (value === null) return null;
  const color = plainDataRecord(value);
  if (!color) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const components = [color.r, color.g, color.b];
  if (
    components.some(
      (component) =>
        typeof component !== "number" ||
        !Number.isFinite(component) ||
        component < 0 ||
        component > 1,
    )
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  return Object.freeze({
    r: color.r as number,
    g: color.g as number,
    b: color.b as number,
  });
}

function sanitizeNativeVectorStyle(
  value: unknown,
  paintMode: PdfNativeVectorSelectionRoute["paintMode"],
): PdfNativeVectorObjectStyle {
  const style = plainDataRecord(value);
  if (!style) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const fillColor = exactNativeVectorStyleColor(style.fillColor);
  const strokeColor = exactNativeVectorStyleColor(style.strokeColor);
  const lineWidth = style.lineWidth;
  const dashPhase = style.dashPhase;
  const miterLimit = style.miterLimit;
  const opacity = style.opacity;
  const blendMode = style.blendMode;
  const dashArray = style.dashArray;
  const strokes = paintMode === "stroke" || paintMode === "fill-stroke";
  const fills = paintMode === "fill" || paintMode === "fill-stroke";
  if (
    (!fills && fillColor !== null) ||
    (!strokes &&
      (strokeColor !== null ||
        lineWidth !== null ||
        dashArray !== null ||
        dashPhase !== null ||
        style.lineCap !== null ||
        style.lineJoin !== null ||
        miterLimit !== null)) ||
    (lineWidth !== null &&
      (typeof lineWidth !== "number" ||
        !Number.isFinite(lineWidth) ||
        lineWidth < 0 ||
        lineWidth > 1_000)) ||
    (dashPhase !== null &&
      (typeof dashPhase !== "number" ||
        !Number.isFinite(dashPhase) ||
        Math.abs(dashPhase) > MAX_ABS_GEOMETRY_VALUE)) ||
    (miterLimit !== null &&
      (typeof miterLimit !== "number" ||
        !Number.isFinite(miterLimit) ||
        miterLimit < 1 ||
        miterLimit > MAX_ABS_GEOMETRY_VALUE)) ||
    (opacity !== null &&
      (typeof opacity !== "number" ||
        !Number.isFinite(opacity) ||
        opacity < 0 ||
        opacity > 1)) ||
    (blendMode !== null && !isPdfStandardBlendMode(blendMode)) ||
    (style.lineCap !== null &&
      style.lineCap !== "butt" &&
      style.lineCap !== "round" &&
      style.lineCap !== "square") ||
    (style.lineJoin !== null &&
      style.lineJoin !== "miter" &&
      style.lineJoin !== "round" &&
      style.lineJoin !== "bevel") ||
    typeof style.hasExternalGraphicsState !== "boolean"
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  let exactDashArray: readonly number[] | null = null;
  if (dashArray !== null) {
    if (
      !Array.isArray(dashArray) ||
      dashArray.length > 64 ||
      dashArray.some(
        (entry) =>
          typeof entry !== "number" ||
          !Number.isFinite(entry) ||
          entry < 0 ||
          entry > MAX_ABS_GEOMETRY_VALUE,
      )
    ) {
      throw new PdfSemanticAppearanceInspectionError("invalid-response");
    }
    exactDashArray = Object.freeze([...dashArray]);
  }
  return Object.freeze({
    fillColor,
    strokeColor,
    lineWidth: lineWidth as number | null,
    dashArray: exactDashArray,
    dashPhase: dashPhase as number | null,
    lineCap: style.lineCap as PdfNativeVectorObjectStyle["lineCap"],
    lineJoin: style.lineJoin as PdfNativeVectorObjectStyle["lineJoin"],
    miterLimit: miterLimit as number | null,
    opacity: opacity as number | null,
    blendMode: blendMode as PdfStandardBlendMode | null,
    hasExternalGraphicsState: style.hasExternalGraphicsState,
    ...(style.fillUnreadable === true ? { fillUnreadable: true } : {}),
  });
}

function sanitizeNativeVectorStyleEvidence(
  value: unknown,
  route: PdfNativeVectorSelectionRoute,
): PdfNativeVectorStyleSelectionEvidence | undefined {
  if (value === undefined) return undefined;
  const style = plainDataRecord(value);
  const proofId = typeof style?.proofId === "string" ? style.proofId : "";
  if (
    !style ||
    typeof style !== "object" ||
    style.schemaVersion !== "mrai.native-vector-object-style.v1" ||
    !SHA256_PATTERN.test(String(style.proofDigest ?? "")) ||
    !isExactInteger(style.pageProofIndex, 0) ||
    proofId.length === 0 ||
    proofId.length > MAX_PROOF_ID_LENGTH ||
    proofId !==
      `pdf-native-vector-object-style:${route.sourceSha256}:p${route.pageNumber}:${route.proofFamily}:${style.pageProofIndex}:${style.proofDigest}`
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  return Object.freeze({
    schemaVersion: "mrai.native-vector-object-style.v1",
    proofId,
    proofDigest: String(style.proofDigest),
    pageProofIndex: style.pageProofIndex,
    style: sanitizeNativeVectorStyle(style.style, route.paintMode),
    ...sanitizeNativeVectorDeviceCmyk(style.deviceCmyk),
  });
}

function sanitizeNativeVectorDeviceCmyk(
  value: unknown,
): { deviceCmyk: PdfNativeVectorDeviceCmyk } | Record<string, never> {
  if (value === undefined) return {};
  const record = plainDataRecord(value);
  const components = (entry: unknown): PdfNativeVectorCmyk | null => {
    if (entry === null) return null;
    if (
      !Array.isArray(entry) ||
      entry.length !== 4 ||
      entry.some((part) => typeof part !== "number" || !(part >= 0 && part <= 1))
    ) {
      throw new PdfSemanticAppearanceInspectionError("invalid-response");
    }
    return Object.freeze([entry[0], entry[1], entry[2], entry[3]] as const);
  };
  if (!record) throw new PdfSemanticAppearanceInspectionError("invalid-response");
  const fill = components(record.fill ?? null);
  const stroke = components(record.stroke ?? null);
  return fill || stroke ? { deviceCmyk: Object.freeze({ fill, stroke }) } : {};
}

/** Mirror of the byte engine's pdfNativeVectorProofId (kept here so this
 * main-thread module does not load the engine). */
const PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN = /^\d{1,9}(?:\.\d{1,9}){0,7}$/;
function pdfNativeVectorProofId(input: {
  readonly family: "shape" | "subpath";
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly objectNumber: number;
  readonly generation: number;
  readonly start: number;
  readonly end: number;
  readonly formInstance?: string;
}): string {
  return `pdf-native-vector${input.family === "subpath" ? "-subpath" : ""}:${input.sourceSha256}:p${input.pageNumber}:${input.objectNumber}:${input.generation}:${input.start}:${input.end}${
    input.formInstance !== undefined ? `:i${input.formInstance}` : ""
  }`;
}

function sanitizeNativeVectorRoute(
  route: PdfNativeVectorSelectionRoute,
  sourceSha256: string,
  pageNumber: number,
): PdfNativeVectorSelectionRoute {
  const styleDescriptor = Object.getOwnPropertyDescriptor(route, "style");
  if (styleDescriptor && !("value" in styleDescriptor)) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const proofId =
    typeof route?.proofId === "string" ? route.proofId.trim() : "";
  const proofFamily = route?.proofFamily;
  const nativeKind = route?.nativeKind;
  if (
    !route ||
    typeof route !== "object" ||
    route.selectionOnly !== true ||
    route.sourceSha256 !== sourceSha256 ||
    route.pageNumber !== pageNumber ||
    !isExactInteger(route.rendererOperatorIndex, 0) ||
    (proofFamily !== "shape" && proofFamily !== "subpath") ||
    (nativeKind !== "line" &&
      nativeKind !== "rectangle" &&
      nativeKind !== "ellipse" &&
      nativeKind !== "subpath") ||
    (proofFamily === "subpath") !== (nativeKind === "subpath") ||
    (route.paintMode !== "stroke" &&
      route.paintMode !== "fill" &&
      route.paintMode !== "fill-stroke") ||
    proofId.length === 0 ||
    proofId.length > MAX_PROOF_ID_LENGTH ||
    route.proofId !== proofId ||
    !SHA256_PATTERN.test(route.proofDigest) ||
    !isExactInteger(route.contentStreamObjectNumber, 1) ||
    !isExactInteger(route.contentStreamGeneration, 0) ||
    !isExactInteger(route.operatorStart, 0) ||
    !isExactInteger(route.operatorEnd, 1) ||
    route.operatorEnd <= route.operatorStart
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const formInstance = route.formInstance;
  if (
    formInstance !== undefined &&
    (typeof formInstance !== "string" ||
      !PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN.test(formInstance))
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const expectedProofId = pdfNativeVectorProofId({
    family: proofFamily,
    sourceSha256,
    pageNumber,
    objectNumber: route.contentStreamObjectNumber,
    generation: route.contentStreamGeneration,
    start: route.operatorStart,
    end: route.operatorEnd,
    formInstance,
  });
  if (proofId !== expectedProofId) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  const sanitized = {
    selectionOnly: true,
    sourceSha256,
    pageNumber,
    rendererOperatorIndex: route.rendererOperatorIndex,
    proofFamily,
    nativeKind,
    paintMode: route.paintMode,
    proofId,
    proofDigest: route.proofDigest,
    contentStreamObjectNumber: route.contentStreamObjectNumber,
    contentStreamGeneration: route.contentStreamGeneration,
    operatorStart: route.operatorStart,
    operatorEnd: route.operatorEnd,
    ...(formInstance !== undefined ? { formInstance } : {}),
    ...(typeof route.lineWidthScale === "number" &&
    Number.isFinite(route.lineWidthScale) &&
    route.lineWidthScale > 0
      ? { lineWidthScale: route.lineWidthScale }
      : {}),
    bounds: exactRouteBounds(route.bounds),
  } as const;
  const style = sanitizeNativeVectorStyleEvidence(
    styleDescriptor?.value,
    sanitized,
  );
  return Object.freeze({
    ...sanitized,
    ...(style ? { style } : {}),
  });
}

function sanitizeUnresolved(
  route: PdfSemanticAppearanceUnresolvedRoute,
  pageNumber: number,
): PdfSemanticAppearanceUnresolvedRoute {
  if (
    !route ||
    typeof route !== "object" ||
    route.pageNumber !== pageNumber ||
    !isExactInteger(route.objectNumber, 1) ||
    !isExactInteger(route.generation, 0) ||
    !isSemanticKind(route.semanticKind) ||
    !ROUTE_FAILURE_REASONS.has(route.reason)
  ) {
    throw new PdfSemanticAppearanceInspectionError("invalid-response");
  }
  return Object.freeze({
    pageNumber,
    objectNumber: route.objectNumber,
    generation: route.generation,
    semanticKind: route.semanticKind,
    reason: route.reason,
  });
}

interface PendingPageRequest {
  readonly pageNumber: number;
  readonly targetKeys: ReadonlySet<string>;
  readonly visibleOperatorIndexes: ReadonlySet<number> | null;
  readonly resolve: (result: PdfSemanticAppearanceInspectionPageResult) => void;
  readonly reject: (error: PdfSemanticAppearanceInspectionError) => void;
  readonly timer: ReturnType<typeof setTimeout>;
}

class InspectionSession implements PdfSemanticAppearanceInspectionSession {
  readonly sourceGeneration: string;
  readonly sourceSha256: string;
  readonly #sessionId = sessionId();
  readonly #planners: PdfNativePlanners | null;
  readonly #worker: PdfSemanticAppearanceInspectionWorkerLike;
  readonly #readyPromise: Promise<PdfSemanticAppearanceInspectionReadyReceipt>;
  #resolveReady!: (
    receipt: PdfSemanticAppearanceInspectionReadyReceipt,
  ) => void;
  #rejectReady!: (error: PdfSemanticAppearanceInspectionError) => void;
  #readyState: "pending" | "ready" | "failed" = "pending";
  #initTimer: ReturnType<typeof setTimeout> | null = null;
  #destroyed = false;
  #requestSequence = 0;
  readonly #pending = new Map<string, PendingPageRequest>();
  readonly #latestRequestByPage = new Map<number, string>();
  /** How often each page's request timed out: the next asks for more time. */
  readonly #pageTimeoutsByPage = new Map<number, number>();

  constructor(input: PdfSemanticAppearanceInspectionInput) {
    this.sourceGeneration = exactSourceGeneration(input.sourceGeneration);
    this.sourceSha256 = exactSha256(input.sourceSha256);
    this.#planners = input.planners ?? null;
    if (
      !(input.sourceBytes instanceof Uint8Array) ||
      input.sourceBytes.byteLength === 0 ||
      input.sourceBytes.byteLength > SOURCE_MAX_BYTES
    ) {
      throw new PdfSemanticAppearanceInspectionError("invalid-source");
    }
    try {
      this.#worker = (testWorkerFactory ?? defaultWorkerFactory)();
    } catch (error) {
      if (error instanceof PdfSemanticAppearanceInspectionError) throw error;
      throw new PdfSemanticAppearanceInspectionError("worker-unavailable");
    }
    this.#readyPromise = new Promise((resolve, reject) => {
      this.#resolveReady = resolve;
      this.#rejectReady = reject;
    });
    this.#worker.onmessage = (event) => this.#receive(event.data);
    this.#worker.onerror = () =>
      this.#fail(new PdfSemanticAppearanceInspectionError("worker-failed"));
    const initTimeoutMs =
      typeof input.initTimeoutMs === "number" && Number.isFinite(input.initTimeoutMs)
        ? Math.min(MAX_INIT_TIMEOUT_MS, Math.max(MIN_INIT_TIMEOUT_MS, Math.floor(input.initTimeoutMs)))
        : PDF_SEMANTIC_APPEARANCE_INIT_TIMEOUT_MS;
    this.#initTimer = globalThis.setTimeout(
      () =>
        this.#fail(new PdfSemanticAppearanceInspectionError("init-timeout")),
      initTimeoutMs,
    );
    const clone = Uint8Array.from(input.sourceBytes);
    const sourceBytes = clone.buffer as ArrayBuffer;
    // The worker reads the owners and native vector pages itself, the page
    
    // off the engine host reads them and hands them over.
    const localEvidence = pdfSemanticAppearanceLocalEvidenceEnabled();
    const engineCopy = localEvidence ? null : Uint8Array.from(input.sourceBytes);
    try {
      this.#worker.postMessage(
        {
          protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
          kind: "init",
          sessionId: this.#sessionId,
          sourceGeneration: this.sourceGeneration,
          sourceSha256: this.sourceSha256,
          sourceBytes,
          ...(localEvidence ? { localEvidence: true } : {}),
        },
        [sourceBytes],
      );
    } catch {
      if (clone.byteLength > 0) clone.fill(0);
      engineCopy?.fill(0);
      this.#fail(new PdfSemanticAppearanceInspectionError("worker-failed"));
      return;
    }
    if (engineCopy) void this.#sendNativeEvidence(engineCopy);
  }

  async #sendNativeEvidence(engineCopy: Uint8Array): Promise<void> {
    let nativeEvidence: PdfSemanticAppearanceSelectionSourceEvidence | null = null;
    try {
      const { inspectPdfSemanticAppearanceSelectionSource } =
        this.#planners ?? (await import("./remote/pdf-native-planners")).createPdfNativePlanners(null);
      // Read in chunks of pages: a long paper's figures came to one answer
      // past the engine call's bound (21 MB for 47 pages), and the whole
      // inspection failed on open. The engine inspects the source once.
      let inventory: PdfSemanticAppearanceSelectionSourceEvidence["inventory"] | null = null;
      const pages: PdfSemanticAppearanceSelectionSourceEvidence["nativePages"][number][] = [];
      for (let next: number | null = 1, calls = 0; next !== null; calls += 1) {
        let part: Awaited<ReturnType<typeof inspectPdfSemanticAppearanceSelectionSource>>;
        try {
          part = await inspectPdfSemanticAppearanceSelectionSource(engineCopy,
            { firstPage: next, maxBytes: PDF_SEMANTIC_EVIDENCE_CHUNK_BYTES });
        } catch (error) {
          // One page whose vectors the engine could not read in time (a
          // datasheet's 15 s diagram page) has no vector objects to select;
          // the pages after it still do. Before the first answer there is no
          // inventory to go on.
          if (!inventory || calls > inventory.pageCount) throw error;
          pages.push(Object.freeze({ pageNumber: next, shapes: Object.freeze([]), subpaths: Object.freeze([]),
            styles: Object.freeze([]) }));
          next = next < inventory.pageCount ? next + 1 : null;
          continue;
        }
        inventory ??= part.inventory;
        pages.push(...part.nativePages);
        const following: number | null = part.nextPage ?? null;
        if (following !== null && (following <= next || calls > inventory.pageCount)) throw new Error("inspection-failed");
        next = following;
      }
      if (inventory) nativeEvidence = Object.freeze({ inventory, nativePages: Object.freeze(pages) });
    } catch {
      nativeEvidence = null;
    } finally {
      engineCopy.fill(0);
    }
    if (this.#destroyed) return;
    try {
      this.#worker.postMessage({
        protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
        kind: "native-evidence",
        sessionId: this.#sessionId,
        sourceGeneration: this.sourceGeneration,
        sourceSha256: this.sourceSha256,
        nativeEvidence,
      });
    } catch {
      this.#fail(new PdfSemanticAppearanceInspectionError("worker-failed"));
    }
  }

  ready(): Promise<PdfSemanticAppearanceInspectionReadyReceipt> {
    return this.#readyPromise;
  }

  async routePage(
    input: PdfSemanticAppearanceInspectionPageInput,
  ): Promise<PdfSemanticAppearanceInspectionPageResult> {
    await this.#readyPromise;
    if (this.#destroyed || this.#readyState !== "ready") {
      throw new PdfSemanticAppearanceInspectionError("destroyed");
    }
    if (
      !isExactInteger(input.pageNumber, 1) ||
      !Array.isArray(input.targets) ||
      input.targets.length > MAX_TARGETS_PER_PAGE
    ) {
      throw new PdfSemanticAppearanceInspectionError("invalid-request");
    }
    const targets = Object.freeze(
      input.targets.map((target) =>
        sanitizeTarget(target, this.sourceSha256, input.pageNumber),
      ),
    );
    const operatorEvidence = input.operatorEvidence
      ? createPdfSemanticAppearancePageOperatorEvidence(input.operatorEvidence)
      : undefined;
    const visibleOperatorIndexes = operatorEvidence
      ? new Set(
          operatorEvidence.visibleVectors.map((vector) => vector.operatorIndex),
        )
      : null;
    const targetKeys = new Set<string>();
    for (const target of targets) {
      const key = routeIdentityKey(target);
      if (targetKeys.has(key)) {
        throw new PdfSemanticAppearanceInspectionError("invalid-request");
      }
      targetKeys.add(key);
    }
    const previousId = this.#latestRequestByPage.get(input.pageNumber);
    if (previousId) {
      const previous = this.#pending.get(previousId);
      if (previous) {
        globalThis.clearTimeout(previous.timer);
        this.#pending.delete(previousId);
        previous.reject(new PdfSemanticAppearanceInspectionError("superseded"));
      }
    }
    this.#requestSequence += 1;
    const requestId = `${this.#sessionId}:${this.#requestSequence}`;
    this.#latestRequestByPage.set(input.pageNumber, requestId);
    return new Promise((resolve, reject) => {
      // A page that takes too long fails this request only — the session
      // and every other page go on — and is given more time when asked
      // again (PDF_SEMANTIC_APPEARANCE_PAGE_RETRY_FACTOR).
      const timer = globalThis.setTimeout(() => {
        const timedOut = this.#pending.get(requestId);
        if (!timedOut) return;
        this.#pending.delete(requestId);
        if (this.#latestRequestByPage.get(input.pageNumber) === requestId) {
          this.#latestRequestByPage.delete(input.pageNumber);
        }
        this.#pageTimeoutsByPage.set(
          input.pageNumber,
          (this.#pageTimeoutsByPage.get(input.pageNumber) ?? 0) + 1,
        );
        timedOut.reject(new PdfSemanticAppearanceInspectionError("page-timeout"));
      }, pdfSemanticAppearancePageTimeoutMs(
        operatorEvidence?.visibleVectors.length ?? 0,
        this.#pageTimeoutsByPage.get(input.pageNumber) ?? 0,
      ));
      this.#pending.set(requestId, {
        pageNumber: input.pageNumber,
        targetKeys,
        visibleOperatorIndexes,
        resolve,
        reject,
        timer,
      });
      try {
        this.#worker.postMessage({
          protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
          kind: "route-page",
          sessionId: this.#sessionId,
          sourceGeneration: this.sourceGeneration,
          sourceSha256: this.sourceSha256,
          requestId,
          pageNumber: input.pageNumber,
          targets,
          ...(operatorEvidence ? { operatorEvidence } : {}),
        });
      } catch {
        this.#fail(new PdfSemanticAppearanceInspectionError("worker-failed"));
      }
    });
  }

  destroy(): void {
    if (this.#destroyed) return;
    try {
      this.#worker.postMessage({
        protocolVersion: PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION,
        kind: "destroy",
        sessionId: this.#sessionId,
        sourceGeneration: this.sourceGeneration,
        sourceSha256: this.sourceSha256,
      });
    } catch {
      // Termination below is the authority boundary.
    }
    this.#fail(new PdfSemanticAppearanceInspectionError("destroyed"));
  }

  #receive(response: PdfSemanticAppearanceWorkerResponse): void {
    if (this.#destroyed) return;
    if (
      !response ||
      typeof response !== "object" ||
      response.protocolVersion !==
        PDF_SEMANTIC_APPEARANCE_INSPECTION_PROTOCOL_VERSION ||
      response.sessionId !== this.#sessionId ||
      response.sourceGeneration !== this.sourceGeneration ||
      response.sourceSha256 !== this.sourceSha256
    ) {
      this.#fail(new PdfSemanticAppearanceInspectionError("invalid-response"));
      return;
    }
    if (response.kind === "ready") {
      if (
        this.#readyState !== "pending" ||
        typeof response.sourceBytesZeroized !== "boolean"
      ) {
        this.#fail(
          new PdfSemanticAppearanceInspectionError("invalid-response"),
        );
        return;
      }
      if (this.#initTimer !== null) {
        globalThis.clearTimeout(this.#initTimer);
        this.#initTimer = null;
      }
      this.#readyState = "ready";
      this.#resolveReady(
        Object.freeze({
          sourceSha256: this.sourceSha256,
          sourceBytesZeroized: true,
        }),
      );
      return;
    }
    if (response.kind === "error") {
      // One page the worker could not route is that page's failure: it took
      
      const failedRequest = (response as { requestId?: unknown; code?: unknown }).requestId;
      if (typeof failedRequest === "string" && (response as { code?: unknown }).code === "routing-failed" &&
        this.#pending.has(failedRequest)) {
        this.#refusePage(failedRequest);
        return;
      }
      this.#fail(new PdfSemanticAppearanceInspectionError("worker-failed"));
      return;
    }
    if (response.kind !== "page-routes" || this.#readyState !== "ready") {
      this.#fail(new PdfSemanticAppearanceInspectionError("invalid-response"));
      return;
    }
    const pending = this.#pending.get(response.requestId);
    if (!pending) return;
    if (
      response.pageNumber !== pending.pageNumber ||
      !Array.isArray(response.routes) ||
      !Array.isArray(response.unresolved) ||
      !Array.isArray(response.vectorExclusions) ||
      !Array.isArray(response.nativeVectorRoutes) ||
      response.routes.length > MAX_RESULT_ROWS_PER_PAGE ||
      response.unresolved.length > MAX_RESULT_ROWS_PER_PAGE ||
      response.vectorExclusions.length > MAX_VECTORS_PER_PAGE ||
      response.nativeVectorRoutes.length > MAX_VECTORS_PER_PAGE
    ) {
      this.#refusePage(response.requestId);
      return;
    }
    let routes: readonly PdfSemanticAppearanceSelectionRoute[];
    let unresolved: readonly PdfSemanticAppearanceUnresolvedRoute[];
    let vectorExclusions: readonly PdfSemanticAppearanceVectorExclusion[];
    let nativeVectorRoutes: readonly PdfNativeVectorSelectionRoute[];
    try {
      routes = Object.freeze(
        response.routes.map((route) =>
          sanitizeRoute(route, this.sourceSha256, pending.pageNumber),
        ),
      );
      unresolved = Object.freeze(
        response.unresolved.map((route) =>
          sanitizeUnresolved(route, pending.pageNumber),
        ),
      );
      vectorExclusions = Object.freeze(
        response.vectorExclusions.map((exclusion) =>
          sanitizeVectorExclusion(exclusion, pending.pageNumber),
        ),
      );
      nativeVectorRoutes = Object.freeze(
        response.nativeVectorRoutes.map((route) =>
          sanitizeNativeVectorRoute(
            route,
            this.sourceSha256,
            pending.pageNumber,
          ),
        ),
      );
      const resultKeys = new Set<string>();
      for (const route of [...routes, ...unresolved]) {
        const key = routeIdentityKey(route);
        if (!pending.targetKeys.has(key) || resultKeys.has(key)) {
          throw new PdfSemanticAppearanceInspectionError("invalid-response");
        }
        resultKeys.add(key);
      }
      if (resultKeys.size !== pending.targetKeys.size) {
        throw new PdfSemanticAppearanceInspectionError("invalid-response");
      }
      const claimedOperatorIndexes = new Set<number>();
      const proofIds = new Set<string>();
      for (const exclusion of vectorExclusions) {
        if (
          !pending.visibleOperatorIndexes?.has(
            exclusion.rendererOperatorIndex,
          ) ||
          claimedOperatorIndexes.has(exclusion.rendererOperatorIndex)
        ) {
          throw new PdfSemanticAppearanceInspectionError("invalid-response");
        }
        claimedOperatorIndexes.add(exclusion.rendererOperatorIndex);
      }
      for (const route of nativeVectorRoutes) {
        if (
          !pending.visibleOperatorIndexes?.has(route.rendererOperatorIndex) ||
          claimedOperatorIndexes.has(route.rendererOperatorIndex) ||
          proofIds.has(route.proofId)
        ) {
          throw new PdfSemanticAppearanceInspectionError("invalid-response");
        }
        claimedOperatorIndexes.add(route.rendererOperatorIndex);
        proofIds.add(route.proofId);
      }
      if (
        pending.visibleOperatorIndexes === null &&
        (vectorExclusions.length !== 0 || nativeVectorRoutes.length !== 0)
      ) {
        throw new PdfSemanticAppearanceInspectionError("invalid-response");
      }
    } catch (error) {
      if (process.env.NODE_ENV !== "production") {
        // Which check refused the page's answer (every one throws the same).
        console.debug("[pdf-semantic] page answer refused", pending.pageNumber,
          (error instanceof Error ? error.stack ?? "" : String(error)).split("\n").slice(0, 4).join(" | "));
      }
      this.#refusePage(response.requestId);
      return;
    }
    globalThis.clearTimeout(pending.timer);
    this.#pending.delete(response.requestId);
    if (
      this.#latestRequestByPage.get(pending.pageNumber) === response.requestId
    ) {
      this.#latestRequestByPage.delete(pending.pageNumber);
    }
    pending.resolve(
      Object.freeze({
        sourceSha256: this.sourceSha256,
        pageNumber: pending.pageNumber,
        routes,
        unresolved,
        vectorExclusions,
        nativeVectorRoutes,
      }),
    );
  }

  /** One page's request refused; the session goes on. */
  #refusePage(requestId: string): void {
    const pending = this.#pending.get(requestId);
    if (!pending) return;
    globalThis.clearTimeout(pending.timer);
    this.#pending.delete(requestId);
    if (this.#latestRequestByPage.get(pending.pageNumber) === requestId) {
      this.#latestRequestByPage.delete(pending.pageNumber);
    }
    pending.reject(new PdfSemanticAppearanceInspectionError("page-refused"));
  }

  #fail(error: PdfSemanticAppearanceInspectionError): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    if (this.#initTimer !== null) {
      globalThis.clearTimeout(this.#initTimer);
      this.#initTimer = null;
    }
    if (this.#readyState === "pending") {
      this.#readyState = "failed";
      this.#rejectReady(error);
    }
    for (const pending of this.#pending.values()) {
      globalThis.clearTimeout(pending.timer);
      pending.reject(error);
    }
    this.#pending.clear();
    this.#latestRequestByPage.clear();
    this.#worker.onmessage = null;
    this.#worker.onerror = null;
    this.#worker.terminate();
  }
}

export function createPdfSemanticAppearanceInspectionSession(
  input: PdfSemanticAppearanceInspectionInput,
): PdfSemanticAppearanceInspectionSession {
  return new InspectionSession(input);
}
