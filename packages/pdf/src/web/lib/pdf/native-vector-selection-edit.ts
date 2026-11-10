/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { displayAffineToPdfUserSpace, pdfNativeVectorDisplayMatrixAffine, PDF_NATIVE_VECTOR_DUPLICATE_NUDGE_POINTS, pdfNativeVectorDisplayShiftOffset, GEOMETRY_EPSILON, MRAI_NATIVE_VECTOR_DELETE_ACTION, MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_DUPLICATE_ACTION, MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_REORDER_ACTION, MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION, PAGE_EPSILON, PDF_NATIVE_VECTOR_DELETE_CONTROL_ID, PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE, PDF_NATIVE_VECTOR_OFF_PAGE_MESSAGE, PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID, PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID, buildPdfNativeVectorDeleteCommand, buildPdfNativeVectorDuplicateCommand, buildPdfNativeVectorGeometryCommand, buildPdfNativeVectorReorderCommand, createPdfNativeVectorGeometryOwnershipManifest, createPdfNativeVectorHistoryEndpoint, createPdfNativeVectorRotationOrigin, exactAffine, exactCommands, exactFinite, exactGeometry, exactInteger, exactPageGeometry, fail, geometryState, historyEndpoint, nativeBoundsMatch, nativeVectorLineThicknessDrag, pdfNativeVectorAcceptedTarget, pdfNativeVectorGroupBounds, pdfNativeVectorReorderRefusal, pdfNativeVectorRouteIsAcceptedTarget, readPdfNativeVectorAwareReplayAuthority, readPdfNativeVectorDeletePayload, readPdfNativeVectorDuplicatePayload, readPdfNativeVectorReorderPayload, readPdfNativeVectorSelectionEditPayload, samePdfNativeVectorHistoryEndpoint } from "./native-vector-commands";
export { MRAI_NATIVE_VECTOR_DELETE_ACTION, PDF_NATIVE_VECTOR_DUPLICATE_NUDGE_POINTS, pdfNativeVectorDisplayShiftOffset, displayAffineToPdfUserSpace, MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_DUPLICATE_ACTION, MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_REORDER_ACTION, MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION, MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION, PDF_NATIVE_VECTOR_DELETE_CONTROL_ID, PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE, PDF_NATIVE_VECTOR_OFF_PAGE_MESSAGE, PDF_NATIVE_VECTOR_GEOMETRY_CONTROL_ID, PDF_NATIVE_VECTOR_GEOMETRY_FEATURE_ID, buildPdfNativeVectorDeleteCommand, buildPdfNativeVectorDuplicateCommand, buildPdfNativeVectorGeometryCommand, buildPdfNativeVectorReorderCommand, createPdfNativeVectorGeometryOwnershipManifest, createPdfNativeVectorHistoryEndpoint, createPdfNativeVectorRotationOrigin, nativeVectorLineThicknessDrag, pdfNativeVectorAcceptedTarget, pdfNativeVectorGroupBounds, pdfNativeVectorReorderRefusal, pdfNativeVectorRouteIsAcceptedTarget, readPdfNativeVectorAwareReplayAuthority, readPdfNativeVectorDeletePayload, readPdfNativeVectorDuplicatePayload, readPdfNativeVectorReorderPayload, readPdfNativeVectorSelectionEditPayload, samePdfNativeVectorHistoryEndpoint };
import {
  MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
} from "./engine/native/vocabulary";
export {
  MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
};
import { PDFDict, PDFDocument, PDFName, PDFRef } from "pdf-lib";
import {
  MRAI_NATIVE_SOURCE_CONSOLIDATION_ACTION,
  MRAI_NATIVE_SOURCE_CONSOLIDATION_KIND,
} from "./engine/native/source-consolidation";

import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";

import { readPdfCommandOwnershipManifest } from "./engine/command-ownership";
import type {
  PdfEngineCommand,
  PdfEngineSourceIdentity,
} from "./engine/contracts";
import {
  inspectMraiNativeVectorContent,
  inspectMraiNativeVectorSubpaths,
  PDF_NATIVE_VECTOR_FORM_INSTANCE_PATTERN,
  pdfNativeVectorProofId,
  MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
  planMraiNativeVectorDelete,
  planMraiNativeVectorDuplicate,
  planMraiNativeVectorDuplicateRemoval,
  planMraiNativeVectorGeometryEdit,
  planMraiNativeVectorReorder,
  planMraiNativeVectorReorderRestore,
  planMraiNativeVectorOutlineEdit,
  planMraiNativeVectorSubpathAffineEdit,
  prepareMraiNativeVectorDuplicate,
  prepareMraiNativeVectorGeometryEdit,
  prepareMraiNativeVectorDelete,
  prepareMraiNativeVectorReorder,
  prepareMraiNativeVectorOutlineEdit,
  prepareMraiNativeVectorSubpathAffineEdit,
  readMraiNativeVectorOutline,
  verifyMraiNativeVectorDuplicateCandidate,
  verifyMraiNativeVectorGeometryCandidate,
  verifyMraiNativeVectorDeleteCandidate,
  verifyMraiNativeVectorOutlineCandidate,
  verifyMraiNativeVectorReorderCandidate,
  verifyMraiNativeVectorSubpathCandidate,
  type PdfNativeVectorAffineTransform,
  type PdfNativeVectorContentShape,
  type PdfNativeVectorDuplicatePlan,
  type PdfNativeVectorGeometry,
  type PdfNativeVectorGeometryEditPlan,
  type PdfNativeVectorDeletePlan,
  type PdfNativeVectorPassedUnit,
  type PdfNativeVectorBounds,
  type PdfNativeVectorReorderPlan,
  type PdfNativeVectorSubpathAffineEditPlan,
  type PdfNativeVectorSubpathCommand,
  type PdfNativeVectorSubpathProof,
} from "./engine/native/vector-content-transaction";
import {
  inspectMraiNativePageContentGraph,
  prepareMraiNativeContentStreamTransaction,
  readMraiNativePdfXrefSectionCount,
  sha256PdfBytes,
  type PdfNativePreparedContentStreamTransaction,
} from "./engine/native/content-stream-transaction";
import type { PdfNativeListTextReplayAuthority } from "./engine/native/list-text-command-lifecycle";
import type { PdfNativeVectorSelectionRoute } from "./semantic-appearance-inspection";
import type { PdfVectorPageGeometry } from "./vector-capability-reconciliation";



export type PdfNativeRectangleGeometry = Extract<
  PdfNativeVectorGeometry,
  { kind: "rectangle" }
>;

export interface PdfNativeVectorNormalizedBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export type PdfNativeVectorWriteSelection = Readonly<
  | {
      proofFamily: "shape";
      sourceSha256: string;
      pageNumber: number;
      pageProofIndex: number;
      nativeKind: "line" | "rectangle" | "ellipse";
      paintMode: PdfNativeVectorContentShape["paintMode"];
      proofId: string;
      proofDigest: string;
      contentStreamObjectNumber: number;
      contentStreamGeneration: number;
      operatorStart: number;
      operatorEnd: number;
      operatorSha256: string;
      bounds: PdfNativeVectorBounds;
      geometry: PdfNativeVectorGeometry;
      /** Which drawing of a Form the page shows more than once (part of the
       * proof id). */
      formInstance?: string;
    }
  | {
      proofFamily: "subpath";
      sourceSha256: string;
      pageNumber: number;
      pageProofIndex: number;
      nativeKind: "subpath";
      paintMode: PdfNativeVectorSubpathProof["paintMode"];
      paintOperator: PdfNativeVectorSubpathProof["paintOperator"];
      proofId: string;
      proofDigest: string;
      contentStreamObjectNumber: number;
      contentStreamGeneration: number;
      operatorStart: number;
      operatorEnd: number;
      operatorSha256: string;
      bounds: PdfNativeVectorBounds;
      commands: readonly PdfNativeVectorSubpathCommand[];
      closed: boolean;
      formInstance?: string;
    }
>;

type PdfNativeVectorHistoryEndpointFor<
  Selection extends PdfNativeVectorWriteSelection,
> = Selection extends PdfNativeVectorWriteSelection
  ? Omit<
      Selection,
      | "sourceSha256"
      | "proofId"
      | "proofDigest"
      | "operatorStart"
      | "operatorEnd"
      | "operatorSha256"
    >
  : never;

export type PdfNativeVectorHistoryEndpoint = Readonly<
  PdfNativeVectorHistoryEndpointFor<PdfNativeVectorWriteSelection>
>;

/**
 * One anchor of the object moved by the payload's translation (WPS "编辑顶点"):
 * a line's start or end, or a subpath's `m`/`l` point or curve end, counted in
 * drawing order. The control points attached to that anchor travel with it;
 * every other operand and every operator stays. The affine is then the
 * anchor's own move, and the server derives the same replacement from it
 * (mirror: _vector_move_selection_anchor in native_surgical_transaction.py).
 */
export interface PdfNativeVectorVertexEdit {
  readonly anchorIndex: number;
}

/** Office's Merge Shapes, as one object's outline replaced whole. */
export type PdfNativeVectorMergeOperation = "union" | "combine" | "fragment" | "intersect" | "subtract";

/**
 * Merge Shapes (PowerPoint's Shape Format > Merge Shapes): the object's path
 * replaced by the merged outline (page space; any number of subpaths, each
 * hole drawn the other way round from the outline it is in) under its own
 * graphics state, transform and paint operator, as PowerPoint gives the
 * result the first selected shape's format. The other merged objects are
 * deleted in the same step (PdfPageWindow).
 */
export interface PdfNativeVectorOutlineReplace {
  readonly operation: PdfNativeVectorMergeOperation;
  readonly commands: readonly PdfNativeVectorSubpathCommand[];
}

export interface PdfNativeVectorSelectionEditPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly affine: PdfNativeVectorAffineTransform;
  readonly rotationHistory?: PdfNativeVectorRotationHistory;
  readonly vertexEdit?: PdfNativeVectorVertexEdit;
  readonly outlineReplace?: PdfNativeVectorOutlineReplace;
}

export interface PdfNativeVectorRotationOrigin {
  readonly schemaVersion: "mrai.pdf-native-vector-rotation-origin.v1";
  readonly sourceLength: number;
  readonly candidateLength: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly affine: PdfNativeVectorAffineTransform;
  readonly before: PdfNativeVectorWriteSelection;
  readonly after: PdfNativeVectorWriteSelection;
}

export interface PdfNativeVectorRotationHistory {
  readonly direction: "undo" | "redo";
  readonly origin: PdfNativeVectorRotationOrigin;
}

export interface PdfNativeVectorGeometryEditEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly proofFamily: PdfNativeVectorWriteSelection["proofFamily"];
  readonly nativeKind: PdfNativeVectorWriteSelection["nativeKind"];
  readonly paintMode: PdfNativeVectorWriteSelection["paintMode"];
  readonly paintOperator: PdfNativeVectorSubpathProof["paintOperator"] | null;
  readonly sourceProofId: string;
  readonly sourceProofDigest: string;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly affine: PdfNativeVectorAffineTransform;
  readonly rotationHistory?: PdfNativeVectorRotationHistory;
  readonly vertexEdit?: PdfNativeVectorVertexEdit;
  /** Merge Shapes: the outline replaced whole (its commands are the
   * replacement the record writes). */
  readonly outlineReplace?: Readonly<{ operation: PdfNativeVectorMergeOperation }>;
  readonly before: PdfNativeVectorWriteSelection;
  readonly after: PdfNativeVectorWriteSelection;
  readonly siblingFingerprintCount: number;
  readonly siblingFingerprintSha256: string;
}

export interface PdfNativeVectorSelectionEditPlan {
  readonly payload: PdfNativeVectorSelectionEditPayload;
  readonly evidence: PdfNativeVectorGeometryEditEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly beforeSelection: PdfNativeVectorWriteSelection;
  readonly afterSelection: PdfNativeVectorWriteSelection;
  readonly beforeEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly afterEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly inverseAffine: PdfNativeVectorAffineTransform;
  readonly request:
    | PdfNativeVectorGeometryEditPlan["request"]
    | PdfNativeVectorSubpathAffineEditPlan["request"];
}

export interface PdfNativeVectorDeletePayload {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly pageGeometry: PdfVectorPageGeometry;
}

export interface PdfNativeVectorDeleteEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly proofFamily: PdfNativeVectorWriteSelection["proofFamily"];
  readonly nativeKind: PdfNativeVectorWriteSelection["nativeKind"];
  readonly paintMode: PdfNativeVectorWriteSelection["paintMode"];
  readonly paintOperator: PdfNativeVectorSubpathProof["paintOperator"] | null;
  readonly sourceProofId: string;
  readonly sourceProofDigest: string;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly before: PdfNativeVectorWriteSelection;
  readonly originalOperatorBytes: readonly number[];
  readonly removedOperatorBytes: readonly [110];
  readonly siblingFingerprintCount: number;
  readonly siblingFingerprintSha256: string;
}

export interface PdfNativeVectorSelectionDeletePlan {
  readonly payload: PdfNativeVectorDeletePayload;
  readonly evidence: PdfNativeVectorDeleteEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly beforeSelection: PdfNativeVectorWriteSelection;
  readonly request: PdfNativeVectorDeletePlan["request"];
}

export interface PdfNativeVectorDeleteCommitRequest {
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}

/**
 * Merge Shapes on two selected source vectors (PowerPoint's Shape Format >
 * Merge Shapes, which the other editors take on two shapes): `routes[0]` is
 * the first selected, whose format the result takes. The host reads both
 * outlines from the file and merges them; every operation but Fragment makes
 * one piece, Fragment one per region. Piece i is written into routes[i] (its
 * outline replaced), the pieces beyond the selection into copies of the
 * first, and the objects beyond the pieces are deleted: one step, one Undo.
 */
export interface PdfNativeVectorMergeCommitRequest {
  readonly routes: readonly PdfNativeVectorSelectionRoute[];
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly operation: PdfNativeVectorMergeOperation;
}

export type PdfNativeVectorFlipAxis = "horizontal" | "vertical";

/** Which display axes a resize carried across the frame's opposite edge. */
export type PdfNativeVectorMirror = Readonly<{ x: boolean; y: boolean }>;

export type PdfNativeVectorGeometryCommitRequest =
  | Readonly<{
      route: PdfNativeVectorSelectionRoute;
      pageGeometry: PdfVectorPageGeometry;
      targetBounds: PdfNativeVectorNormalizedBounds;
      /** The frame the gesture started on (the route's box when absent): the
       * shape's frame is its whole path grown by half its stroke, which the
       * route's box (the part the page shows, without the stroke) is not. */
      fromBounds?: PdfNativeVectorNormalizedBounds;
      /** A handle pulled across the opposite edge: the shape comes out
       * mirrored on that axis into `targetBounds`, as in WPS. */
      mirror?: PdfNativeVectorMirror;
      rotationDegrees?: never;
      flipAxis?: never;
      vertex?: never;
      displayMatrix?: never;
    }>
  | Readonly<{
      route: PdfNativeVectorSelectionRoute;
      pageGeometry: PdfVectorPageGeometry;
      rotationDegrees: number;
      targetBounds?: never;
      flipAxis?: never;
      vertex?: never;
      displayMatrix?: never;
    }>
  | Readonly<{
      route: PdfNativeVectorSelectionRoute;
      pageGeometry: PdfVectorPageGeometry;
      flipAxis: PdfNativeVectorFlipAxis;
      targetBounds?: never;
      rotationDegrees?: never;
      vertex?: never;
      displayMatrix?: never;
    }>
  | Readonly<{
      route: PdfNativeVectorSelectionRoute;
      pageGeometry: PdfVectorPageGeometry;
      /** One anchor dragged from `from` to `to`, both in normalized display
       * space; `from` is checked against the source so a stale handle never
       * moves another point. */
      vertex: PdfNativeVectorVertexDrag;
      targetBounds?: never;
      rotationDegrees?: never;
      flipAxis?: never;
      displayMatrix?: never;
    }>
  | Readonly<{
      route: PdfNativeVectorSelectionRoute;
      pageGeometry: PdfVectorPageGeometry;
      /** Any affine of the page as shown, in display points (a turned frame
       * resized along its own axes): written into the shape's points. */
      displayMatrix: readonly [number, number, number, number, number, number];
      targetBounds?: never;
      rotationDegrees?: never;
      flipAxis?: never;
      vertex?: never;
    }>;

export interface PdfNativeVectorVertexDrag {
  readonly anchorIndex: number;
  readonly from: Readonly<{ x: number; y: number }>;
  readonly to: Readonly<{ x: number; y: number }>;
}

export interface PdfNativeVectorDuplicateCommitRequest {
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  /** Where the copy goes (display points, right and down from the source);
   * WPS's nudge when absent. A paste lands it at the pointer. */
  readonly shiftPoints?: Readonly<{ x: number; y: number }>;
}

/**
 * Where a layer command sends an object: one step past the next object
 * whatever it is (WPS "上移一层" / "下移一层", Office Bring Forward / Send
 * Backward) or all the way (WPS "置于顶层" / "置于底层", Office Bring to Front /
 * Send to Back). Each is one write and one history entry.
 */
export type PdfNativeVectorLayerMove = "forward" | "backward" | "front" | "back";

export interface PdfNativeVectorReorderCommitRequest {
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly direction: PdfNativeVectorLayerMove;
}

/** Why the reorder planner refused a move, read from its refusal: the object
 * is already drawn last (first), the next object clips it or changes how it
 * paints in a way it cannot carry, the next object is in another of the
 * page's content streams, that neighbour is drawn from content several pages
 * share, or the object is not drawn on its page at all. Anything else is not
 * a reorder refusal. */
export type PdfNativeVectorReorderRefusal =
  | "at-front"
  | "at-back"
  | "other-content"
  | "other-stream"
  | "shared-neighbour"
  | "off-page";


/** The refusal the reorder planner raises once an object is drawn last (or
 * first): what tells a "front" / "back" loop it has arrived. */
export function pdfNativeVectorIsAtStackEnd(error: unknown): boolean {
  const refusal = pdfNativeVectorReorderRefusal(error);
  return refusal === "at-front" || refusal === "at-back";
}

/**
 * One transform applied to every object of a multiple selection, expressed
 * in normalized display space (page fractions, /Rotate honoured) so the same
 * transform serves each object: the group frame's move or resize as the box
 * it went from and to, or a turn or flip about the group's shared centre.
 */
export type PdfNativeVectorGroupTransform = Readonly<
  | {
      kind: "bounds";
      fromBounds: PdfNativeVectorNormalizedBounds;
      toBounds: PdfNativeVectorNormalizedBounds;
    }
  | {
      kind: "rotation";
      rotationDegrees: number;
      center: Readonly<{ x: number; y: number }>;
    }
  | {
      kind: "flip";
      axis: PdfNativeVectorFlipAxis;
      center: Readonly<{ x: number; y: number }>;
    }
>;


export interface PdfNativeVectorGroupGeometryCommitRequest {
  readonly routes: readonly PdfNativeVectorSelectionRoute[];
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly transform: PdfNativeVectorGroupTransform;
  /** Each member's own transform, by its place in `routes` (对齐 and 分布
   * move each object by its own amount); `transform` for any not given. */
  readonly memberTransforms?: readonly PdfNativeVectorGroupTransform[];
}

/**
 * The edited object's identity in the revision an accepted write produced,
 * read back from the committed bytes. It is the only thing a viewer may use
 * to select the object again once the new revision's routes arrive: offsets
 * alone can name a different object after a deletion or a source swap.
 */
export type PdfNativeVectorAcceptedTarget = Readonly<
  Pick<
    PdfNativeVectorSelectionRoute,
    | "sourceSha256"
    | "pageNumber"
    | "proofFamily"
    | "nativeKind"
    | "paintMode"
    | "proofId"
    | "proofDigest"
    | "contentStreamObjectNumber"
    | "contentStreamGeneration"
    | "operatorStart"
    | "operatorEnd"
    | "formInstance"
  >
>;

/**
 * What a source-vector commit hands back to the viewer: `false` when nothing
 * was written, `true` when a write landed without an object to select next
 * (a deletion), or the accepted target.
 */
export type PdfNativeVectorCommitOutcome =
  | boolean
  | void
  | Readonly<{
      accepted: true;
      target: PdfNativeVectorAcceptedTarget;
      /** The other objects of a group write, as the accepted revision names
       * them, so the viewer keeps the whole group selected. */
      extraTargets?: readonly PdfNativeVectorAcceptedTarget[];
    }>;



export interface PdfNativeVectorDisplayEditInput {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly targetBounds: PdfNativeVectorNormalizedBounds;
  /** The box the gesture went from; the route's box when absent. Only the
   * change between the two is applied to the shape's own geometry. */
  readonly fromBounds?: PdfNativeVectorNormalizedBounds;
  readonly mirror?: PdfNativeVectorMirror;
}

interface PageProofs {
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly shapes: readonly PdfNativeVectorContentShape[];
  readonly subpaths: readonly PdfNativeVectorSubpathProof[];
}









function exactBounds(
  value: unknown,
  label: string,
): PdfNativeVectorNormalizedBounds {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    fail(`Invalid PDF vector ${label}.`);
  }
  const bounds = value as Record<
    keyof PdfNativeVectorNormalizedBounds,
    unknown
  >;
  const result = Object.freeze({
    x: exactFinite(bounds.x, `${label} x`),
    y: exactFinite(bounds.y, `${label} y`),
    width: exactFinite(bounds.width, `${label} width`),
    height: exactFinite(bounds.height, `${label} height`),
  });
  // A frame may reach past the page (a shape the page cuts, or a stroke
  // half outside it): it only has to be on the page somewhere.
  if (
    result.width <= 0 ||
    result.height <= 0 ||
    result.x >= 1 ||
    result.y >= 1 ||
    result.x + result.width <= 0 ||
    result.y + result.height <= 0
  ) {
    fail(`PDF vector ${label} is outside the page.`);
  }
  return result;
}







function selectionFromShape(
  shape: PdfNativeVectorContentShape,
): PdfNativeVectorWriteSelection {
  return Object.freeze({
    proofFamily: "shape" as const,
    sourceSha256: shape.sourceSha256,
    pageNumber: shape.pageNumber,
    pageProofIndex: shape.pageVectorIndex,
    nativeKind: shape.kind,
    paintMode: shape.paintMode,
    proofId: shape.id,
    proofDigest: shape.proofDigest,
    contentStreamObjectNumber: shape.contentStream.objectNumber,
    contentStreamGeneration: shape.contentStream.generation,
    operatorStart: shape.operatorSpan.start,
    operatorEnd: shape.operatorSpan.end,
    operatorSha256: shape.operatorSpan.expectedSha256,
    bounds: Object.freeze({ ...shape.bounds }),
    geometry: exactGeometry(shape.geometry, shape.kind),
    ...(shape.formInstance !== undefined ? { formInstance: shape.formInstance } : {}),
  });
}

function selectionFromSubpath(
  proof: PdfNativeVectorSubpathProof,
): PdfNativeVectorWriteSelection {
  return Object.freeze({
    proofFamily: "subpath" as const,
    sourceSha256: proof.sourceSha256,
    pageNumber: proof.pageNumber,
    pageProofIndex: proof.pageSubpathIndex,
    nativeKind: "subpath" as const,
    paintMode: proof.paintMode,
    paintOperator: proof.paintOperator,
    proofId: proof.id,
    proofDigest: proof.proofDigest,
    contentStreamObjectNumber: proof.contentStream.objectNumber,
    contentStreamGeneration: proof.contentStream.generation,
    operatorStart: proof.operatorSpan.start,
    operatorEnd: proof.operatorSpan.end,
    operatorSha256: proof.operatorSpan.expectedSha256,
    bounds: Object.freeze({ ...proof.bounds }),
    commands: exactCommands(proof.commands),
    closed: proof.closed,
    ...(proof.formInstance !== undefined ? { formInstance: proof.formInstance } : {}),
  });
}

export function createPdfNativeVectorWriteSelection(
  proof: PdfNativeVectorContentShape | PdfNativeVectorSubpathProof,
): PdfNativeVectorWriteSelection {
  return "kind" in proof ? selectionFromShape(proof) : selectionFromSubpath(proof);
}


function selectionMatches(
  left: PdfNativeVectorWriteSelection,
  right: PdfNativeVectorWriteSelection,
): boolean {
  return (
    left.proofFamily === right.proofFamily &&
    left.sourceSha256 === right.sourceSha256 &&
    left.pageNumber === right.pageNumber &&
    left.pageProofIndex === right.pageProofIndex &&
    left.nativeKind === right.nativeKind &&
    left.paintMode === right.paintMode &&
    left.proofId === right.proofId &&
    left.proofDigest === right.proofDigest &&
    left.contentStreamObjectNumber === right.contentStreamObjectNumber &&
    left.contentStreamGeneration === right.contentStreamGeneration &&
    left.operatorStart === right.operatorStart &&
    left.operatorEnd === right.operatorEnd &&
    left.operatorSha256 === right.operatorSha256 &&
    left.formInstance === right.formInstance &&
    JSON.stringify(left.bounds) === JSON.stringify(right.bounds) &&
    geometryState(left) === geometryState(right)
  );
}


async function readPageProofs(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PageProofs> {
  const [shapes, subpaths] = await Promise.all([
    inspectMraiNativeVectorContent(sourceBytes, pageNumber),
    inspectMraiNativeVectorSubpaths(sourceBytes, pageNumber),
  ]);
  if (
    shapes.sourceSha256 !== subpaths.sourceSha256 ||
    shapes.pageNumber !== subpaths.pageNumber
  ) {
    fail("PDF vector proof families disagree on source authority.");
  }
  return Object.freeze({
    sourceSha256: shapes.sourceSha256,
    pageNumber,
    shapes: shapes.shapes,
    subpaths: subpaths.subpaths,
  });
}

function findExactSelection(
  proofs: PageProofs,
  supplied: PdfNativeVectorWriteSelection,
): PdfNativeVectorWriteSelection {
  const candidates =
    supplied.proofFamily === "shape"
      ? proofs.shapes
          .filter((proof) => proof.id === supplied.proofId)
          .map(selectionFromShape)
      : proofs.subpaths
          .filter((proof) => proof.id === supplied.proofId)
          .map(selectionFromSubpath);
  if (candidates.length !== 1 || !selectionMatches(candidates[0], supplied)) {
    fail("Selected PDF vector proof is stale or unavailable.");
  }
  const proof =
    supplied.proofFamily === "shape"
      ? proofs.shapes[supplied.pageProofIndex]
      : proofs.subpaths[supplied.pageProofIndex];
  if (
    !proof ||
    proof.id !== supplied.proofId ||
    !proof.editable ||
    proof.preserveReason !== null
  ) {
    fail("Selected PDF vector is not independently writable.");
  }
  return candidates[0];
}



function endpointMatches(
  selection: PdfNativeVectorWriteSelection,
  endpoint: PdfNativeVectorHistoryEndpoint,
): boolean {
  return (
    selection.proofFamily === endpoint.proofFamily &&
    selection.pageNumber === endpoint.pageNumber &&
    selection.pageProofIndex === endpoint.pageProofIndex &&
    selection.nativeKind === endpoint.nativeKind &&
    selection.paintMode === endpoint.paintMode &&
    selection.contentStreamObjectNumber ===
      endpoint.contentStreamObjectNumber &&
    selection.contentStreamGeneration === endpoint.contentStreamGeneration &&
    selection.formInstance === endpoint.formInstance &&
    nativeBoundsMatch(selection.bounds, endpoint.bounds) &&
    geometryState(selection) ===
      geometryState(endpoint as PdfNativeVectorWriteSelection)
  );
}


export async function resolvePdfNativeVectorHistorySelection(input: {
  readonly sourceBytes: Uint8Array;
  readonly endpoint: PdfNativeVectorHistoryEndpoint;
}): Promise<PdfNativeVectorWriteSelection> {
  const proofs = await readPageProofs(
    input.sourceBytes,
    input.endpoint.pageNumber,
  );
  const candidates = [
    ...proofs.shapes.map(selectionFromShape),
    ...proofs.subpaths.map(selectionFromSubpath),
  ].filter((selection) => endpointMatches(selection, input.endpoint));
  if (candidates.length !== 1) {
    fail("PDF vector history endpoint is stale or ambiguous.");
  }
  return candidates[0];
}

type SourcePageGeometry = Readonly<{
  cropBox: Readonly<{ xMin: number; yMin: number; xMax: number; yMax: number }>;
  rotation: number;
}>;

/**
 * Every page's CropBox and rotation of recently read sources, by content
 * hash. One vector edit re-proves its page geometry several times (route,
 * plan, command, save), and each read parsed the whole file with pdf-lib --
 * seconds per read on a large report. The bytes a hash names never change,
 * so the answer is kept for the last few sources.
 */
const SOURCE_PAGE_GEOMETRY_LIMIT = 4;
const sourcePageGeometries = new Map<string, Promise<readonly SourcePageGeometry[]>>();

export async function readPdfNativeSourcePageGeometries(
  sourceBytes: Uint8Array,
): Promise<readonly SourcePageGeometry[]> {
  const key = await sha256PdfBytes(sourceBytes);
  const cached = sourcePageGeometries.get(key);
  if (cached) {
    sourcePageGeometries.delete(key);
    sourcePageGeometries.set(key, cached);
    return cached;
  }
  const loading = (async () => {
    const document = await PDFDocument.load(Uint8Array.from(sourceBytes), {
      updateMetadata: false,
    });
    return Object.freeze(
      document.getPages().map((page) => {
        const crop = page.getCropBox();
        return Object.freeze({
          cropBox: Object.freeze({
            xMin: crop.x,
            yMin: crop.y,
            xMax: crop.x + crop.width,
            yMax: crop.y + crop.height,
          }),
          rotation: ((page.getRotation().angle % 360) + 360) % 360,
        });
      }),
    );
  })();
  sourcePageGeometries.set(key, loading);
  loading.catch(() => {
    if (sourcePageGeometries.get(key) === loading) sourcePageGeometries.delete(key);
  });
  while (sourcePageGeometries.size > SOURCE_PAGE_GEOMETRY_LIMIT) {
    sourcePageGeometries.delete(sourcePageGeometries.keys().next().value!);
  }
  return loading;
}

async function readActualPageGeometry(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfVectorPageGeometry> {
  let pages: readonly SourcePageGeometry[];
  try {
    pages = await readPdfNativeSourcePageGeometries(sourceBytes);
  } catch {
    fail("PDF vector page authority could not be reopened.");
  }
  const page = pages[pageNumber - 1];
  if (!page) fail("PDF vector page is out of range.");
  return exactPageGeometry({
    cropBox: { ...page.cropBox },
    rotation: page.rotation as PdfVectorPageGeometry["rotation"],
  });
}

function samePageGeometry(
  left: PdfVectorPageGeometry,
  right: PdfVectorPageGeometry,
): boolean {
  return (
    left.rotation === right.rotation &&
    (Object.keys(left.cropBox) as Array<keyof typeof left.cropBox>).every(
      (key) => Math.abs(left.cropBox[key] - right.cropBox[key]) <= PAGE_EPSILON,
    )
  );
}

function proofBounds(selection: PdfNativeVectorWriteSelection) {
  return selection.bounds;
}

/**
 * How far past its CropBox a shape may still reach and be drawn whole on the
 * page. Two residues land here and neither is a clip.
 *
 * PAGE_EPSILON is a FRACTION of the page, the unit every other page-boundary
 * guard here works in (see affineFromDisplayBounds). The display layer hands
 * this writer a box rounded to six decimals of the page, so a shape dragged
 * flush against the page edge lands up to one such fraction past the CropBox
 * in user space. Reading the same number as an absolute point would make this
 * guard ~600x tighter than the one that already accepted the edit, and refuse
 * a full-bleed shape over the rounding residue alone.
 *
 * The other residue is the source's own: a box and a path are both written as
 * decimal reals, often to a different number of places (a report whose page is
 * `1190.55` wide carries a background drawn to `1190.552`), so a shape meant
 * to bleed to the page edge overshoots it by a unit or two of the last place
 * either number was written with. A hundredth of a point is below that, and
 * below a device pixel at any resolution a page is rasterized at, so a shape
 * within it is a bleed rather than something the page cuts. Relative alone
 * cannot express this: the residue comes from how the numbers were written,
 * not from how big the page is, and the guard would be tightest exactly where
 * the page is smallest.
 */
const CROP_BLEED_TOLERANCE = 0.01;

function cropTolerance(
  geometry: PdfVectorPageGeometry,
): Readonly<{ x: number; y: number }> {
  const crop = geometry.cropBox;
  return Object.freeze({
    x: Math.max(PAGE_EPSILON * (crop.xMax - crop.xMin), CROP_BLEED_TOLERANCE),
    y: Math.max(PAGE_EPSILON * (crop.yMax - crop.yMin), CROP_BLEED_TOLERANCE),
  });
}

/** The same tolerance as a fraction of the page, in the display's own axes. */
function displayCropTolerance(
  geometry: PdfVectorPageGeometry,
): Readonly<{ x: number; y: number }> {
  const crop = geometry.cropBox;
  const width = crop.xMax - crop.xMin;
  const height = crop.yMax - crop.yMin;
  const upright = geometry.rotation % 180 === 0;
  const tolerance = cropTolerance(geometry);
  return Object.freeze({
    x: (upright ? tolerance.x : tolerance.y) / (upright ? width : height),
    y: (upright ? tolerance.y : tolerance.x) / (upright ? height : width),
  });
}

/** The page draws the whole shape. */
function boundsInsideCropBox(
  selection: PdfNativeVectorWriteSelection,
  geometry: PdfVectorPageGeometry,
): boolean {
  const bounds = proofBounds(selection);
  const crop = geometry.cropBox;
  const tolerance = cropTolerance(geometry);
  return (
    bounds.x >= crop.xMin - tolerance.x &&
    bounds.y >= crop.yMin - tolerance.y &&
    bounds.x + bounds.width <= crop.xMax + tolerance.x &&
    bounds.y + bounds.height <= crop.yMax + tolerance.y
  );
}

/**
 * Some of the shape falls inside the page's own box -- it is drawn where a
 * reader can see it and a pointer can reach it. An imposed sheet (four 612pt
 * slides side by side on one 2448pt page, each PDF page a CropBox window onto
 * it) draws shapes that straddle a window's edge, and shapes that miss the
 * window altogether: only the first kind is this page's to edit.
 */
function boundsMeetCropBox(
  selection: PdfNativeVectorWriteSelection,
  geometry: PdfVectorPageGeometry,
): boolean {
  const bounds = proofBounds(selection);
  const crop = geometry.cropBox;
  const tolerance = cropTolerance(geometry);
  return (
    bounds.x < crop.xMax - tolerance.x &&
    bounds.y < crop.yMax - tolerance.y &&
    bounds.x + bounds.width > crop.xMin + tolerance.x &&
    bounds.y + bounds.height > crop.yMin + tolerance.y
  );
}

/** A vertex edit moves one anchor by the affine's own translation. */
function translationOnly(affine: PdfNativeVectorAffineTransform): boolean {
  return (
    Math.abs(affine.a - 1) <= GEOMETRY_EPSILON &&
    Math.abs(affine.d - 1) <= GEOMETRY_EPSILON &&
    Math.abs(affine.b) <= GEOMETRY_EPSILON &&
    Math.abs(affine.c) <= GEOMETRY_EPSILON
  );
}

/** An outline replacement moves nothing: its commands are the path. */
function identityAffine(affine: PdfNativeVectorAffineTransform): boolean {
  return translationOnly(affine) &&
    Math.abs(affine.e) <= GEOMETRY_EPSILON && Math.abs(affine.f) <= GEOMETRY_EPSILON;
}

/**
 * The page's own shape: some of it is drawn inside the CropBox. A shape the
 * page cuts is edited as the whole shape it is (WPS moves, resizes, turns and
 * mirrors it, and lets an edit take part of any shape past the page edge,
 * where the page simply does not draw it); one wholly outside is not this
 * page's (an imposed sheet's other windows) and stays preserve-only. Both
 * sides hold every edit to the same rule: the shape meets the CropBox before
 * and after (native_surgical_transaction.py).
 */
function projectedRawBounds(
  selection: PdfNativeVectorWriteSelection,
  geometry: PdfVectorPageGeometry,
) {
  if (!boundsOnPage(selection, geometry)) {
    fail(PDF_NATIVE_VECTOR_OFF_PAGE_MESSAGE);
  }
  return rawProjectedBounds(selection, geometry);
}

/** Whether the page draws the selection, whole or in part: what every
 * vector edit -- a resize written as a style `cm` among them -- leaves true. */
export function pdfNativeVectorSelectionOnPage(
  selection: PdfNativeVectorWriteSelection,
  geometry: PdfVectorPageGeometry,
): boolean {
  return boundsOnPage(selection, geometry);
}

/**
 * One shape of a page as the AI names it (pdf.object.drawn_shape, planned on
 * the engine host by shape-edit-plan.ts): its write selection, the route a
 * click on it gives (built from its own proof), and its box and anchors on
 * the shown page (fractions from the top-left; a straight line's box may
 * have no width or no height).
 */
export interface PdfNativeVectorPageShape {
  readonly selection: PdfNativeVectorWriteSelection;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly bounds: PdfNativeVectorNormalizedBounds;
  /** A line's ends, a path's vertices, a rectangle's corners or an
   * ellipse's quadrant points (pdfNativeVectorShapeAnchors). */
  readonly anchors: readonly Readonly<{ x: number; y: number }>[] | null;
}

/**
 * The page's shapes the writers take -- every proof that is editable, not
 * preserved and drawn on the page -- in one stable order (by content stream,
 * then where in it the path is drawn), with the page geometry the display
 * planners check. The window's own routes also leave out what its renderer
 * did not reconcile and what a semantic owner claims; the writers' checks
 * stay the authority for both.
 */
export async function readPdfNativeVectorPageShapes(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<Readonly<{ pageGeometry: PdfVectorPageGeometry; shapes: readonly PdfNativeVectorPageShape[] }>> {
  const pageGeometry = await readActualPageGeometry(sourceBytes, pageNumber);
  const proofs = await readPageProofs(sourceBytes, pageNumber);
  const writable = [
    ...proofs.shapes
      .filter((proof) => proof.editable && proof.preserveReason === null)
      .map(selectionFromShape),
    ...proofs.subpaths
      .filter((proof) => proof.editable && proof.preserveReason === null)
      .map(selectionFromSubpath),
  ].filter((selection) => boundsOnPage(selection, pageGeometry));
  writable.sort(
    (left, right) =>
      left.contentStreamObjectNumber - right.contentStreamObjectNumber ||
      left.contentStreamGeneration - right.contentStreamGeneration ||
      (left.formInstance ?? "").localeCompare(right.formInstance ?? "") ||
      left.operatorStart - right.operatorStart,
  );
  const crop = pageGeometry.cropBox;
  const upright = pageGeometry.rotation % 180 === 0;
  const displayWidth = upright ? crop.xMax - crop.xMin : crop.yMax - crop.yMin;
  const displayHeight = upright ? crop.yMax - crop.yMin : crop.xMax - crop.xMin;
  const shapes = writable.map((selection, index): PdfNativeVectorPageShape => {
    const bounds = Object.freeze(rawProjectedBounds(selection, pageGeometry));
    const anchors = pdfNativeVectorShapeAnchors(selection);
    return Object.freeze({
      selection,
      route: Object.freeze({
        selectionOnly: true as const,
        sourceSha256: selection.sourceSha256,
        pageNumber,
        rendererOperatorIndex: index,
        proofFamily: selection.proofFamily,
        nativeKind: selection.nativeKind,
        paintMode: selection.paintMode,
        proofId: selection.proofId,
        proofDigest: selection.proofDigest,
        contentStreamObjectNumber: selection.contentStreamObjectNumber,
        contentStreamGeneration: selection.contentStreamGeneration,
        operatorStart: selection.operatorStart,
        operatorEnd: selection.operatorEnd,
        ...(selection.formInstance !== undefined ? { formInstance: selection.formInstance } : {}),
        bounds,
      }),
      bounds,
      anchors: anchors.length > 0
        ? Object.freeze(
            anchors.map((point) => {
              const shown = displayPointForPdfPoint(point, pageGeometry);
              return Object.freeze({ x: shown.x / displayWidth, y: shown.y / displayHeight });
            }),
          )
        : null,
    });
  });
  return Object.freeze({ pageGeometry, shapes: Object.freeze(shapes) });
}

/** The page draws the shape, whole or in part (a rule lying on the page's
 * own edge is drawn whole without crossing into it). */
function boundsOnPage(
  selection: PdfNativeVectorWriteSelection,
  geometry: PdfVectorPageGeometry,
): boolean {
  return (
    boundsInsideCropBox(selection, geometry) ||
    boundsMeetCropBox(selection, geometry)
  );
}

/** The same verdict for a box in display fractions, with the same bleed
 * tolerance carried there: an edit may take a shape past the page edge,
 * never wholly off it (there would be no frame to take hold of again). */
function displayBoundsMeetPage(
  next: Readonly<{ x: number; y: number; width: number; height: number }>,
  geometry: PdfVectorPageGeometry,
): boolean {
  const tolerance = displayCropTolerance(geometry);
  const whole =
    next.x >= -tolerance.x &&
    next.y >= -tolerance.y &&
    next.x + next.width <= 1 + tolerance.x &&
    next.y + next.height <= 1 + tolerance.y;
  return (
    whole ||
    (next.x < 1 - tolerance.x &&
      next.y < 1 - tolerance.y &&
      next.x + next.width > tolerance.x &&
      next.y + next.height > tolerance.y)
  );
}

function rawProjectedBounds(
  selection: PdfNativeVectorWriteSelection,
  geometry: PdfVectorPageGeometry,
) {
  const bounds = proofBounds(selection);
  const crop = geometry.cropBox;
  const x2 = bounds.x + bounds.width;
  const y2 = bounds.y + bounds.height;
  const width = crop.xMax - crop.xMin;
  const height = crop.yMax - crop.yMin;
  if (geometry.rotation === 0) {
    return {
      x: (bounds.x - crop.xMin) / width,
      y: (crop.yMax - y2) / height,
      width: bounds.width / width,
      height: bounds.height / height,
    };
  }
  if (geometry.rotation === 90) {
    return {
      x: (bounds.y - crop.yMin) / height,
      y: (bounds.x - crop.xMin) / width,
      width: bounds.height / height,
      height: bounds.width / width,
    };
  }
  if (geometry.rotation === 180) {
    return {
      x: (crop.xMax - x2) / width,
      y: (bounds.y - crop.yMin) / height,
      width: bounds.width / width,
      height: bounds.height / height,
    };
  }
  return {
    x: (crop.yMax - y2) / height,
    y: (crop.xMax - x2) / width,
    width: bounds.height / height,
    height: bounds.width / width,
  };
}

function axisTransform(
  sourceMin: number,
  sourceExtent: number,
  deltaMin: number,
  deltaMax: number,
  /** The resize crossed the opposite edge: the axis comes out reversed,
   * its start landing on `nextMax` and its end on `nextMin`. */
  mirrored = false,
): Readonly<{ scale: number; translate: number }> {
  if (sourceExtent <= GEOMETRY_EPSILON) {
    if (Math.abs(deltaMin - deltaMax) > GEOMETRY_EPSILON) {
      fail("A zero-thickness PDF vector axis can only be translated.");
    }
    // Reversing an axis with no extent changes nothing.
    return Object.freeze({ scale: 1, translate: (deltaMin + deltaMax) / 2 });
  }
  const nextMin = sourceMin + deltaMin;
  const nextMax = sourceMin + sourceExtent + deltaMax;
  const nextExtent = nextMax - nextMin;
  if (nextExtent <= GEOMETRY_EPSILON) {
    fail("PDF vector scale would collapse its geometry.");
  }
  const scale = ((mirrored ? -1 : 1) * nextExtent) / sourceExtent;
  return Object.freeze({
    scale,
    translate: (mirrored ? nextMax : nextMin) - scale * sourceMin,
  });
}



export function invertPdfNativeVectorAffine(
  input: PdfNativeVectorAffineTransform,
): PdfNativeVectorAffineTransform {
  const affine = exactAffine(input);
  const determinant = affine.a * affine.d - affine.b * affine.c;
  return exactAffine({
    a: affine.d / determinant,
    b: -affine.b / determinant,
    c: -affine.c / determinant,
    d: affine.a / determinant,
    e: (affine.c * affine.f - affine.d * affine.e) / determinant,
    f: (affine.b * affine.e - affine.a * affine.f) / determinant,
  });
}

function affineFromDisplayBounds(input: {
  readonly selection: PdfNativeVectorWriteSelection;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly routeBounds: PdfNativeVectorNormalizedBounds;
  readonly targetBounds: PdfNativeVectorNormalizedBounds;
  readonly mirror?: PdfNativeVectorMirror;
}) {
  // Nothing of it drawn on this page: this page's frame is not its frame and
  // no gesture here is about it. A shape the page cuts is the page's whole:
  // its frame is its whole path (the part past the edge too), so a drag on
  // any edge it shows says exactly where that edge goes.
  const actual = projectedRawBounds(input.selection, input.pageGeometry);
  const route = exactBounds(input.routeBounds, "route bounds");
  const target = exactBounds(input.targetBounds, "target bounds");
  const leftDelta = target.x - route.x;
  const rightDelta = target.x + target.width - (route.x + route.width);
  const topDelta = target.y - route.y;
  const bottomDelta = target.y + target.height - (route.y + route.height);
  const x = axisTransform(actual.x, actual.width, leftDelta, rightDelta, input.mirror?.x === true);
  const y = axisTransform(actual.y, actual.height, topDelta, bottomDelta, input.mirror?.y === true);
  const next = {
    x: actual.x + leftDelta,
    y: actual.y + topDelta,
    width: actual.width + rightDelta - leftDelta,
    height: actual.height + bottomDelta - topDelta,
  };
  // Past the page edge is the page's to cut, as WPS lets it; wholly off it
  // there would be no frame to take hold of again.
  if (!displayBoundsMeetPage(next, input.pageGeometry)) {
    fail(PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE);
  }
  return displayAffineToPdfUserSpace({
    pageGeometry: input.pageGeometry,
    scaleX: x.scale,
    scaleY: y.scale,
    translateX: x.translate,
    translateY: y.translate,
  });
}

function transformPoint(
  affine: PdfNativeVectorAffineTransform,
  point: Readonly<{ x: number; y: number }>,
) {
  return Object.freeze({
    x: affine.a * point.x + affine.c * point.y + affine.e,
    y: affine.b * point.x + affine.d * point.y + affine.f,
  });
}

function transformedGeometry(
  geometry: PdfNativeVectorGeometry,
  affine: PdfNativeVectorAffineTransform,
): PdfNativeVectorGeometry {
  if (geometry.kind === "line") {
    return exactGeometry({
      kind: geometry.kind,
      start: transformPoint(affine, geometry.start),
      end: transformPoint(affine, geometry.end),
    });
  }
  if (geometry.kind === "rectangle") {
    return exactGeometry({
      kind: geometry.kind,
      corners: geometry.corners.map((point) => transformPoint(affine, point)),
    });
  }
  const transformedCenter = transformPoint(affine, geometry.center);
  const transformedAxisXEnd = transformPoint(affine, {
    x: geometry.center.x + geometry.axisX.x,
    y: geometry.center.y + geometry.axisX.y,
  });
  const transformedAxisYEnd = transformPoint(affine, {
    x: geometry.center.x + geometry.axisY.x,
    y: geometry.center.y + geometry.axisY.y,
  });
  return exactGeometry({
    kind: geometry.kind,
    center: transformedCenter,
    axisX: {
      x: transformedAxisXEnd.x - transformedCenter.x,
      y: transformedAxisXEnd.y - transformedCenter.y,
    },
    axisY: {
      x: transformedAxisYEnd.x - transformedCenter.x,
      y: transformedAxisYEnd.y - transformedCenter.y,
    },
  });
}

function pageFingerprintRows(
  proofs: PageProofs,
  excluded: PdfNativeVectorWriteSelection,
  /** The Forms the edit wrote on clones (its formInstanceIsolation): each
   * clone's proofs are read as its original's, the stream the page drew them
   * from before the edit. */
  cloned: readonly Readonly<{
    clone: Readonly<{ objectNumber: number; generation: number }>;
    original: Readonly<{ objectNumber: number; generation: number }>;
  }>[] = [],
   
  unindexed = false,
): readonly string[] {
  const stream = (selection: PdfNativeVectorWriteSelection) =>
    cloned.find(
      (item) =>
        selection.contentStreamObjectNumber === item.clone.objectNumber &&
        selection.contentStreamGeneration === item.clone.generation,
    )?.original ?? {
      objectNumber: selection.contentStreamObjectNumber,
      generation: selection.contentStreamGeneration,
    };
  return Object.freeze(
    [
      ...proofs.shapes.map(selectionFromShape),
      ...proofs.subpaths.map(selectionFromSubpath),
    ]
      .filter(
        (selection) =>
          !(
            selection.proofFamily === excluded.proofFamily &&
            selection.pageProofIndex === excluded.pageProofIndex &&
            selection.contentStreamObjectNumber ===
              excluded.contentStreamObjectNumber
          ),
      )
      .map((selection) => {
        const { objectNumber, generation } = stream(selection);
        const index = unindexed ? "" : `${selection.pageProofIndex}:`;
        return `${selection.proofFamily}:${index}${objectNumber}:${generation}:${geometryState(selection)}`;
      }),
  );
}

function proofFromSelection(
  selection: PdfNativeVectorWriteSelection,
): PdfNativeVectorContentShape | PdfNativeVectorSubpathProof {
  const common = {
    id: selection.proofId,
    sourceSha256: selection.sourceSha256,
    pageNumber: selection.pageNumber,
    paintMode: selection.paintMode,
    bounds: selection.bounds,
    contentStream: Object.freeze({
      objectNumber: selection.contentStreamObjectNumber,
      generation: selection.contentStreamGeneration,
    }),
    operatorSpan: Object.freeze({
      start: selection.operatorStart,
      end: selection.operatorEnd,
      expectedSha256: selection.operatorSha256,
    }),
    proofDigest: selection.proofDigest,
    editable: true,
    preserveReason: null,
    ...(selection.formInstance !== undefined
      ? { formInstance: selection.formInstance }
      : {}),
  } as const;
  return selection.proofFamily === "shape"
    ? Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_GEOMETRY_SCHEMA_VERSION,
        ...common,
        pageVectorIndex: selection.pageProofIndex,
        kind: selection.nativeKind,
        geometry: selection.geometry,
      })
    : Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_SUBPATH_SCHEMA_VERSION,
        ...common,
        pageSubpathIndex: selection.pageProofIndex,
        paintOperator: selection.paintOperator,
        commands: selection.commands,
        closed: selection.closed,
      });
}

function deletePlanSiblingRows(
  plan: PdfNativeVectorDeletePlan,
): readonly string[] {
  const shapeRows = plan.beforeShapeFingerprints
    .filter(
      (_fingerprint, index) =>
        !("kind" in plan.selection) ||
        index !== plan.selection.pageVectorIndex,
    )
    .map((fingerprint) => `shape:${fingerprint}`);
  const subpathRows = plan.beforeSubpathFingerprints
    .filter(
      (_fingerprint, index) =>
        "kind" in plan.selection ||
        index !== plan.selection.pageSubpathIndex,
    )
    .map((fingerprint) => `subpath:${fingerprint}`);
  return Object.freeze([...shapeRows, ...subpathRows]);
}

async function siblingDigest(rows: readonly string[]): Promise<string> {
  return sha256PdfBytes(new TextEncoder().encode(JSON.stringify(rows)));
}




async function verifyRotationHistorySource(
  sourceBytes: Uint8Array,
  payload: PdfNativeVectorSelectionEditPayload,
): Promise<void> {
  const history = payload.rotationHistory;
  if (!history) return;
  const { origin, direction } = history;
  if (origin.candidateLength > sourceBytes.length ||
    !samePageGeometry(payload.pageGeometry, origin.pageGeometry)) {
    fail("PDF vector rotation history source is unavailable.");
  }
  const beforeBytes = sourceBytes.subarray(0, origin.sourceLength);
  const afterBytes = sourceBytes.subarray(0, origin.candidateLength);
  const [beforeSha, afterSha] = await Promise.all([
    sha256PdfBytes(beforeBytes), sha256PdfBytes(afterBytes),
  ]);
  if (beforeSha !== origin.before.sourceSha256 || afterSha !== origin.after.sourceSha256 ||
    readMraiNativePdfXrefSectionCount(afterBytes) !== readMraiNativePdfXrefSectionCount(beforeBytes) + 1) {
    fail("PDF vector rotation history prefix changed.");
  }
  const expected = direction === "undo" ? origin.after : origin.before;
  const affine = direction === "undo" ? invertPdfNativeVectorAffine(origin.affine) : origin.affine;
  if (!samePdfNativeVectorHistoryEndpoint(historyEndpoint(payload.selection), historyEndpoint(expected)) ||
    (Object.keys(affine) as (keyof typeof affine)[]).some(
      key => Math.abs(affine[key] - payload.affine[key]) > GEOMETRY_EPSILON,
    )) {
    fail("PDF vector rotation history is not the recorded operation.");
  }
}



/**
 * How far (page points, per axis) a subpath's last segment may end from the
 * subpath's start and still close on it: one vertex there, not two.
 */
export const PDF_NATIVE_VECTOR_CLOSING_VERTEX_TOLERANCE = 0.001;

function closingEnd(
  command: PdfNativeVectorSubpathCommand,
): Readonly<{ x: number; y: number }> | null {
  if (command.operator === "h") return null;
  return "point" in command ? command.point : command.end;
}

/**
 * The vertices of a path in drawing order, each as the commands that draw
 * it: a `m`/`l` point or a curve end (never a control point, never a close).
 * A subpath whose last segment comes back to where its move started (an
 * ellipse's fourth curve, a polygon written with its closing line) has one
 * vertex there, as PowerPoint's freeform does: the move's, which then names
 * that last segment too. Mirror of _vector_vertex_commands in
 * native_surgical_transaction.py.
 */
export function pdfNativeVectorVertexCommands(
  commands: readonly PdfNativeVectorSubpathCommand[],
): readonly (readonly number[])[] {
  const vertices: number[][] = [];
  let open: { vertex: number; start: Readonly<{ x: number; y: number }> } | null = null;
  let segments = 0;
  let last: number | null = null;
  const finish = () => {
    if (open && segments >= 2 && last !== null) {
      const end = closingEnd(commands[last]);
      if (
        end &&
        Math.abs(end.x - open.start.x) <= PDF_NATIVE_VECTOR_CLOSING_VERTEX_TOLERANCE &&
        Math.abs(end.y - open.start.y) <= PDF_NATIVE_VECTOR_CLOSING_VERTEX_TOLERANCE
      ) {
        vertices.pop();
        vertices[open.vertex].push(last);
      }
    }
    open = null;
    segments = 0;
    last = null;
  };
  commands.forEach((command, index) => {
    if (command.operator === "h") {
      finish();
      return;
    }
    if (command.operator === "m") {
      finish();
      vertices.push([index]);
      open = { vertex: vertices.length - 1, start: command.point };
      return;
    }
    // A segment straight after a close opens a subpath with no move of its
    // own: it has no start vertex to close on.
    vertices.push([index]);
    segments += 1;
    last = index;
  });
  finish();
  return Object.freeze(vertices.map((indexes) => Object.freeze(indexes)));
}

/**
 * The points a vertex edit can take hold of, in drawing order and page user
 * space (see pdfNativeVectorVertexCommands).
 */
export function pdfNativeVectorOutlineAnchors(
  commands: readonly PdfNativeVectorSubpathCommand[],
): readonly Readonly<{ x: number; y: number }>[] {
  return pdfNativeVectorVertexCommands(commands).map(
    (indexes) => closingEnd(commands[indexes[0]])!,
  );
}

/**
 * The points a vertex edit can take hold of: a line's start and end; the
 * vertices of a subpath (pdfNativeVectorOutlineAnchors). A rectangle's and an
 * ellipse's are read from the path they draw (readPdfNativeVectorAnchors);
 * here they have none.
 */
export function pdfNativeVectorAnchors(
  selection: PdfNativeVectorWriteSelection,
): readonly Readonly<{ x: number; y: number }>[] | null {
  if (selection.proofFamily === "shape") {
    return selection.geometry.kind === "line"
      ? [selection.geometry.start, selection.geometry.end]
      : null;
  }
  return pdfNativeVectorOutlineAnchors(selection.commands);
}

/**
 * Every kind's points as its proof names them: a rectangle's corners and an
 * ellipse's quadrant points (right, top, left, bottom, the order the four
 * curves it draws end in) are the vertices of the path each draws, which
 * Edit Points turns it into; the writer reads them exactly from that path
 * (readPdfNativeVectorAnchors) and holds a handle to them.
 */
export function pdfNativeVectorShapeAnchors(
  selection: PdfNativeVectorWriteSelection,
): readonly Readonly<{ x: number; y: number }>[] {
  const direct = pdfNativeVectorAnchors(selection);
  if (direct) return direct;
  if (selection.proofFamily !== "shape") return [];
  const geometry = selection.geometry;
  if (geometry.kind === "rectangle") return geometry.corners;
  if (geometry.kind !== "ellipse") return [];
  const at = (axis: Readonly<{ x: number; y: number }>, sign: number) =>
    Object.freeze({
      x: geometry.center.x + axis.x * sign,
      y: geometry.center.y + axis.y * sign,
    });
  return [
    at(geometry.axisX, 1),
    at(geometry.axisY, 1),
    at(geometry.axisX, -1),
    at(geometry.axisY, -1),
  ];
}

/** The explicit path a write selection draws, read from its proof. */
async function readSelectionOutline(
  sourceBytes: Uint8Array,
  proofs: PageProofs,
  selection: PdfNativeVectorWriteSelection,
): Promise<readonly PdfNativeVectorSubpathCommand[]> {
  if (selection.proofFamily === "subpath") return selection.commands;
  const proof = proofs.shapes[selection.pageProofIndex];
  if (!proof || proof.id !== selection.proofId) {
    fail("PDF vector proof index changed.");
  }
  return readMraiNativeVectorOutline({
    sourceBytes,
    pageNumber: selection.pageNumber,
    selection: proof,
  });
}

/**
 * The points a vertex edit can take hold of, every kind included: a
 * rectangle's four corners and an ellipse's four quadrant points come from
 * the path each draws (Office's Edit Points on a preset shape).
 */
async function readPdfNativeVectorAnchors(
  sourceBytes: Uint8Array,
  proofs: PageProofs,
  selection: PdfNativeVectorWriteSelection,
): Promise<readonly Readonly<{ x: number; y: number }>[]> {
  const direct = pdfNativeVectorAnchors(selection);
  if (direct) return direct;
  return pdfNativeVectorOutlineAnchors(
    await readSelectionOutline(sourceBytes, proofs, selection),
  );
}

function shiftPoint(
  point: Readonly<{ x: number; y: number }>,
  delta: Readonly<{ x: number; y: number }>,
) {
  return Object.freeze({ x: point.x + delta.x, y: point.y + delta.y });
}

/**
 * Move one vertex (see pdfNativeVectorVertexCommands) by `delta`. A curve
 * ending at the vertex keeps its shape at that end (its second control point
 * travels), as does the curve leaving it (its first control point travels),
 * so a smooth point stays smooth; a vertex a subpath closes on moves the
 * move and the segment that comes back to it together. The implicit control
 * points of `v` and `y` are then restated from the points they are defined
 * by, so the writer's topology checks hold. Mirror of
 * _vector_move_subpath_anchor in native_surgical_transaction.py.
 */
export function movePdfNativeVectorAnchor(
  commands: readonly PdfNativeVectorSubpathCommand[],
  anchorIndex: number,
  delta: Readonly<{ x: number; y: number }>,
): readonly PdfNativeVectorSubpathCommand[] {
  const targets = pdfNativeVectorVertexCommands(commands)[anchorIndex];
  if (targets === undefined) fail("PDF vector vertex is not part of the path.");
  const moved = commands.map((command, index): PdfNativeVectorSubpathCommand => {
    if (command.operator === "h") return command;
    if ("point" in command) {
      return targets.includes(index)
        ? Object.freeze({ operator: command.operator, point: shiftPoint(command.point, delta) })
        : command;
    }
    const ends = targets.includes(index);
    const leaves = targets.includes(index - 1);
    if (!ends && !leaves) return command;
    return Object.freeze({
      operator: command.operator,
      control1: leaves ? shiftPoint(command.control1, delta) : command.control1,
      control2: ends ? shiftPoint(command.control2, delta) : command.control2,
      end: ends ? shiftPoint(command.end, delta) : command.end,
    });
  });
  // Restate the implicit control points: `v` starts its curve at the current
  // point, `y` ends it at its end point. Mirrors the pen the engine's writer
  // tracks (subpathReplacement): a close returns the pen to the subpath start.
  let currentPoint: Readonly<{ x: number; y: number }> | null = null;
  let startPoint: Readonly<{ x: number; y: number }> | null = null;
  return Object.freeze(
    moved.map((command): PdfNativeVectorSubpathCommand => {
      if (command.operator === "h") {
        currentPoint = startPoint;
        return command;
      }
      if ("point" in command) {
        if (command.operator === "m") startPoint = command.point;
        currentPoint = command.point;
        return command;
      }
      const restated =
        command.operator === "v" && currentPoint
          ? Object.freeze({ ...command, control1: currentPoint })
          : command.operator === "y"
            ? Object.freeze({ ...command, control2: command.end })
            : command;
      currentPoint = command.end;
      return restated;
    }),
  );
}

function movedLineGeometry(
  geometry: PdfNativeVectorGeometry,
  vertexEdit: PdfNativeVectorVertexEdit,
  affine: PdfNativeVectorAffineTransform,
): PdfNativeVectorGeometry {
  if (geometry.kind !== "line") {
    fail("PDF vector rectangles and ellipses have no editable vertices.");
  }
  if (vertexEdit.anchorIndex !== 0 && vertexEdit.anchorIndex !== 1) {
    fail("PDF vector vertex is not part of the path.");
  }
  const delta = { x: affine.e, y: affine.f };
  return exactGeometry({
    kind: "line",
    start: vertexEdit.anchorIndex === 0 ? shiftPoint(geometry.start, delta) : geometry.start,
    end: vertexEdit.anchorIndex === 1 ? shiftPoint(geometry.end, delta) : geometry.end,
  });
}

/** How far (page fractions) a handle may sit from the source anchor it names. */
const VERTEX_HANDLE_TOLERANCE = 0.004;

/**
 * Plan a vertex drag from the viewer: the anchor the handle names is located
 * in the source, checked against where the handle was drawn, and moved to
 * where the pointer let go.
 */
export async function planPdfNativeVectorDisplayVertexEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly vertex: PdfNativeVectorVertexDrag;
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  const anchors = await readPdfNativeVectorAnchors(
    input.sourceBytes,
    await readPageProofs(input.sourceBytes, input.route.pageNumber),
    resolved.selection,
  );
  const anchorIndex = exactInteger(input.vertex.anchorIndex, "vertex anchor index");
  const anchor = anchors[anchorIndex];
  if (!anchor) fail("PDF vector vertex is not part of the path.");
  const geometry = resolved.pageGeometry;
  const crop = geometry.cropBox;
  const pageWidth = crop.xMax - crop.xMin;
  const pageHeight = crop.yMax - crop.yMin;
  const displayWidth = geometry.rotation % 180 === 0 ? pageWidth : pageHeight;
  const displayHeight = geometry.rotation % 180 === 0 ? pageHeight : pageWidth;
  const shown = displayPointForPdfPoint(anchor, geometry);
  const from = exactDisplayPoint(input.vertex.from, "vertex handle");
  if (
    Math.abs(shown.x / displayWidth - from.x) > VERTEX_HANDLE_TOLERANCE ||
    Math.abs(shown.y / displayHeight - from.y) > VERTEX_HANDLE_TOLERANCE
  ) {
    fail("PDF vector vertex handle is stale.");
  }
  const to = exactDisplayPoint(input.vertex.to, "vertex target");
  const target = pdfPointForDisplayPoint(
    { x: to.x * displayWidth, y: to.y * displayHeight },
    geometry,
  );
  const affine = Object.freeze({
    a: 1,
    b: 0,
    c: 0,
    d: 1,
    e: target.x - anchor.x,
    f: target.y - anchor.y,
  });
  if (Math.abs(affine.e) <= GEOMETRY_EPSILON && Math.abs(affine.f) <= GEOMETRY_EPSILON) {
    fail("PDF vector vertex is unchanged.");
  }
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: geometry,
      affine,
      vertexEdit: Object.freeze({ anchorIndex }),
    }),
  });
}

export async function planPdfNativeVectorSelectionDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorDeletePayload;
}): Promise<PdfNativeVectorSelectionDeletePlan> {
  return planSelectionDelete(input);
}

/**
 * The page without one vector, for drawing only (the drag backdrop): its drag
 * shows the page without it rather than a crop that carries the text painted
 * over it. Nothing here is written.
 */
export async function planPdfNativeVectorBackdropSelectionDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorDeletePayload;
}): Promise<PdfNativeVectorSelectionDeletePlan> {
  return planSelectionDelete(input);
}

async function planSelectionDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorDeletePayload;
}): Promise<PdfNativeVectorSelectionDeletePlan> {
  const payload = readPdfNativeVectorDeletePayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (sourceSha256 !== payload.selection.sourceSha256) {
    fail("PDF vector delete selection source changed.");
  }
  const actualGeometry = await readActualPageGeometry(
    input.sourceBytes,
    payload.selection.pageNumber,
  );
  if (!samePageGeometry(actualGeometry, payload.pageGeometry)) {
    fail("PDF vector delete CropBox or rotation changed.");
  }
  const beforeSelection = payload.selection;
  // A shape the page cuts is the page's to delete, as WPS deletes it; one
  // drawn wholly outside the page is not (an imposed sheet's other windows).
  projectedRawBounds(beforeSelection, actualGeometry);
  const deletePlan = await planMraiNativeVectorDelete({
    sourceBytes: input.sourceBytes,
    pageNumber: beforeSelection.pageNumber,
    selection: proofFromSelection(beforeSelection),
  });
  const prepared = await prepareMraiNativeVectorDelete(
    input.sourceBytes,
    deletePlan,
  );
  await verifyMraiNativeVectorDeleteCandidate(
    input.sourceBytes,
    deletePlan,
    prepared,
  );
  const beforeSiblings = deletePlanSiblingRows(deletePlan);
  const pageInspection = await inspectMraiNativePageContentGraph(
    input.sourceBytes,
    beforeSelection.pageNumber,
  );
  const evidence: PdfNativeVectorDeleteEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
    sourceSha256,
    pageNumber: beforeSelection.pageNumber,
    page: Object.freeze({ ...pageInspection.page }),
    pageGeometry: actualGeometry,
    proofFamily: beforeSelection.proofFamily,
    nativeKind: beforeSelection.nativeKind,
    paintMode: beforeSelection.paintMode,
    paintOperator:
      beforeSelection.proofFamily === "subpath"
        ? beforeSelection.paintOperator
        : null,
    sourceProofId: beforeSelection.proofId,
    sourceProofDigest: beforeSelection.proofDigest,
    contentStream: Object.freeze({
      objectNumber: beforeSelection.contentStreamObjectNumber,
      generation: beforeSelection.contentStreamGeneration,
    }),
    operatorSpan: Object.freeze({
      start: beforeSelection.operatorStart,
      end: beforeSelection.operatorEnd,
      expectedSha256: beforeSelection.operatorSha256,
    }),
    before: beforeSelection,
    originalOperatorBytes: Object.freeze([...deletePlan.originalOperatorBytes]),
    removedOperatorBytes: Object.freeze([110] as const),
    siblingFingerprintCount: beforeSiblings.length,
    siblingFingerprintSha256: await siblingDigest(beforeSiblings),
  });
  return Object.freeze({
    payload,
    evidence,
    prepared,
    beforeSelection,
    request: deletePlan.request,
  });
}

export async function planPdfNativeVectorDisplayDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<PdfNativeVectorSelectionDeletePlan> {
  return planSelectionDelete({
    sourceBytes: input.sourceBytes,
    payload: await routeDeletePayload(input),
  });
}

/** The route's page without that vector, for drawing only (see
 * planPdfNativeVectorBackdropSelectionDelete). */
export async function planPdfNativeVectorBackdropDelete(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<PdfNativeVectorSelectionDeletePlan> {
  return planSelectionDelete({
    sourceBytes: input.sourceBytes,
    payload: await routeDeletePayload(input),
  });
}

/**
 * The page with the route's vector replaced, where it is painted, by a fill
 * of the whole page in black and in white -- for drawing only. What is
 * painted after the vector lies over both fills alike and what was painted
 * before it is hidden by them, so the two renders give the content over the
 * vector as one layer (its coverage and colour) across the whole page: a
 * dragged shape then slides under the text that is drawn after it, as the
 * page paints it once moved, instead of tinting it. The fill keeps the
 * vector's own graphics state (its opacity included); the reader solves
 * that from the renders. Nothing here is written.
 */
export async function planPdfNativeVectorBackdropFills(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<{ readonly black: Uint8Array; readonly white: Uint8Array }> {
  const payload = await routeDeletePayload(input);
  const deletePlan = await planMraiNativeVectorDelete({
    sourceBytes: input.sourceBytes,
    pageNumber: payload.selection.pageNumber,
    selection: proofFromSelection(payload.selection),
  });
  const fill = async (gray: 0 | 1) =>
    (
      await prepareMraiNativeContentStreamTransaction(input.sourceBytes, {
        ...deletePlan.request,
        replacement: new TextEncoder().encode(
          `q ${gray} g -100000 -100000 200000 200000 re f Q`,
        ),
      })
    ).candidateBytes;
  return Object.freeze({ black: await fill(0), white: await fill(1) });
}

async function routeDeletePayload(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<PdfNativeVectorDeletePayload> {
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (
    input.route.selectionOnly !== true ||
    input.route.sourceSha256 !== sourceSha256
  ) {
    fail("PDF vector delete route is stale.");
  }
  const proofSelections =
    input.route.proofFamily === "shape"
      ? (
          await inspectMraiNativeVectorContent(
            input.sourceBytes,
            input.route.pageNumber,
          )
        ).shapes.map(selectionFromShape)
      : (
          await inspectMraiNativeVectorSubpaths(
            input.sourceBytes,
            input.route.pageNumber,
          )
        ).subpaths.map(selectionFromSubpath);
  const candidates = proofSelections.filter(
    (selection) =>
      selection.proofFamily === input.route.proofFamily &&
      selection.proofId === input.route.proofId &&
      selection.nativeKind === input.route.nativeKind &&
      selection.paintMode === input.route.paintMode &&
      selection.proofDigest === input.route.proofDigest &&
      selection.contentStreamObjectNumber ===
        input.route.contentStreamObjectNumber &&
      selection.contentStreamGeneration ===
        input.route.contentStreamGeneration &&
      selection.operatorStart === input.route.operatorStart &&
      selection.operatorEnd === input.route.operatorEnd,
  );
  if (candidates.length !== 1) {
    fail("PDF vector delete route is ambiguous.");
  }
  return Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
    selection: candidates[0],
    pageGeometry: exactPageGeometry(input.pageGeometry),
  });
}

export async function planPdfNativeVectorSelectionEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorSelectionEditPayload;
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const payload = readPdfNativeVectorSelectionEditPayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (sourceSha256 !== payload.selection.sourceSha256) {
    fail("PDF vector selection source changed.");
  }
  const actualGeometry = await readActualPageGeometry(
    input.sourceBytes,
    payload.selection.pageNumber,
  );
  if (!samePageGeometry(actualGeometry, payload.pageGeometry)) {
    fail("PDF vector CropBox or rotation changed.");
  }
  const beforeProofs = await readPageProofs(
    input.sourceBytes,
    payload.selection.pageNumber,
  );
  const beforeSelection = findExactSelection(beforeProofs, payload.selection);
  // The page's shape is the one it draws some of, and any edit leaves it so:
  // what an edit takes past the page edge the page cuts (WPS), and a shape
  // the page already cut is edited as the whole shape it is. The server holds
  // the record to the same rule (native_surgical_transaction.py).
  projectedRawBounds(beforeSelection, actualGeometry);
  await verifyRotationHistorySource(input.sourceBytes, payload);
  const vertexEdit = payload.vertexEdit;
  const outlineReplace = payload.outlineReplace;
  if (outlineReplace && (vertexEdit || payload.rotationHistory || !identityAffine(payload.affine))) {
    fail("PDF vector outline replacement carries no other edit.");
  }
  if (vertexEdit) {
    if (payload.rotationHistory) {
      fail("PDF vector vertex edits carry no rotation history.");
    }
    if (!translationOnly(payload.affine)) {
      fail("PDF vector vertex edit affine must be a translation.");
    }
  }
  let prepared: PdfNativePreparedContentStreamTransaction;
  let afterSelection: PdfNativeVectorWriteSelection;
  let request: PdfNativeVectorSelectionEditPlan["request"];
  if (
    outlineReplace ||
    (vertexEdit &&
      !(beforeSelection.proofFamily === "shape" && beforeSelection.nativeKind === "line"))
  ) {
    // Edit Points: the object is written as the explicit path it draws with
    // one vertex moved -- a rectangle's `re` as its `m l l l h`, an
    
    // claims the result.
    const proof =
      beforeSelection.proofFamily === "shape"
        ? beforeProofs.shapes[beforeSelection.pageProofIndex]
        : beforeProofs.subpaths[beforeSelection.pageProofIndex];
    if (!proof || proof.id !== beforeSelection.proofId) {
      fail("PDF vector proof index changed.");
    }
    const outline = await readSelectionOutline(
      input.sourceBytes,
      beforeProofs,
      beforeSelection,
    );
    const plan = await planMraiNativeVectorOutlineEdit({
      sourceBytes: input.sourceBytes,
      pageNumber: beforeSelection.pageNumber,
      selection: proof,
      commands: outlineReplace
        ? outlineReplace.commands
        : movePdfNativeVectorAnchor(outline, vertexEdit!.anchorIndex, {
          x: payload.affine.e,
          y: payload.affine.f,
        }),
      transform: payload.affine,
      ...(outlineReplace ? { replaceOutline: true } : {}),
    });
    prepared = await prepareMraiNativeVectorOutlineEdit(input.sourceBytes, plan);
    const verified = await verifyMraiNativeVectorOutlineCandidate(
      input.sourceBytes,
      plan,
      prepared,
    );
    afterSelection = createPdfNativeVectorWriteSelection(verified.reopened);
    request = plan.request;
  } else if (beforeSelection.proofFamily === "shape") {
    const plan = await planMraiNativeVectorGeometryEdit({
      sourceBytes: input.sourceBytes,
      pageNumber: beforeSelection.pageNumber,
      selectionId: beforeSelection.proofId,
      replacement: vertexEdit
        ? movedLineGeometry(beforeSelection.geometry, vertexEdit, payload.affine)
        : transformedGeometry(beforeSelection.geometry, payload.affine),
    });
    prepared = await prepareMraiNativeVectorGeometryEdit(
      input.sourceBytes,
      plan,
    );
    const verified = await verifyMraiNativeVectorGeometryCandidate(
      input.sourceBytes,
      plan,
      prepared,
    );
    afterSelection = selectionFromShape(verified.reopenedShape);
    request = plan.request;
  } else {
    const currentProof = beforeProofs.subpaths[beforeSelection.pageProofIndex];
    if (!currentProof || currentProof.id !== beforeSelection.proofId) {
      fail("PDF vector subpath proof index changed.");
    }
    const plan = await planMraiNativeVectorSubpathAffineEdit({
      sourceBytes: input.sourceBytes,
      pageNumber: beforeSelection.pageNumber,
      selection: currentProof,
      transform: payload.affine,
    });
    prepared = await prepareMraiNativeVectorSubpathAffineEdit(
      input.sourceBytes,
      plan,
    );
    const verified = await verifyMraiNativeVectorSubpathCandidate(
      input.sourceBytes,
      plan,
      prepared,
    );
    afterSelection = selectionFromSubpath(verified.reopenedSubpath);
    request = plan.request;
  }
  if (!boundsOnPage(afterSelection, actualGeometry)) {
    prepared.candidateBytes.fill(0);
    fail(PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE);
  }
  const afterProofs = await readPageProofs(
    prepared.candidateBytes,
    payload.selection.pageNumber,
  );
  
  
  // after it: the others are then compared in drawing order, unindexed.
  const laneChanged = beforeSelection.proofFamily !== afterSelection.proofFamily;
  const beforeSiblings = pageFingerprintRows(
    beforeProofs,
    beforeSelection,
    [],
    laneChanged,
  );
  const afterSiblings = pageFingerprintRows(
    afterProofs,
    afterSelection,
    prepared.formInstanceIsolation?.clones ?? [],
    laneChanged,
  );
  if (JSON.stringify(beforeSiblings) !== JSON.stringify(afterSiblings)) {
    prepared.candidateBytes.fill(0);
    fail("An unselected PDF vector changed during geometry editing.");
  }
  const pageInspection = await inspectMraiNativePageContentGraph(
    input.sourceBytes,
    beforeSelection.pageNumber,
  );
  const digest = await siblingDigest(beforeSiblings);
  const evidence: PdfNativeVectorGeometryEditEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
    sourceSha256,
    pageNumber: beforeSelection.pageNumber,
    page: Object.freeze({ ...pageInspection.page }),
    pageGeometry: actualGeometry,
    proofFamily: beforeSelection.proofFamily,
    nativeKind: beforeSelection.nativeKind,
    paintMode: beforeSelection.paintMode,
    paintOperator:
      beforeSelection.proofFamily === "subpath"
        ? beforeSelection.paintOperator
        : null,
    sourceProofId: beforeSelection.proofId,
    sourceProofDigest: beforeSelection.proofDigest,
    contentStream: Object.freeze({
      objectNumber: beforeSelection.contentStreamObjectNumber,
      generation: beforeSelection.contentStreamGeneration,
    }),
    operatorSpan: Object.freeze({
      start: beforeSelection.operatorStart,
      end: beforeSelection.operatorEnd,
      expectedSha256: beforeSelection.operatorSha256,
    }),
    affine: payload.affine,
    ...(payload.rotationHistory ? { rotationHistory: payload.rotationHistory } : {}),
    ...(vertexEdit ? { vertexEdit } : {}),
    ...(outlineReplace ? { outlineReplace: Object.freeze({ operation: outlineReplace.operation }) } : {}),
    before: beforeSelection,
    after: afterSelection,
    siblingFingerprintCount: beforeSiblings.length,
    siblingFingerprintSha256: digest,
  });
  return Object.freeze({
    payload,
    evidence,
    prepared,
    beforeSelection,
    afterSelection,
    beforeEndpoint: historyEndpoint(beforeSelection),
    afterEndpoint: historyEndpoint(afterSelection),
    // Merge Shapes is undone by writing its record's bytes back, not by an
    // inverse transform (it has none).
    inverseAffine: outlineReplace ? payload.affine : invertPdfNativeVectorAffine(payload.affine),
    request,
  });
}

async function resolvePdfNativeVectorDisplayRoute(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}) {
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (
    input.route.selectionOnly !== true ||
    input.route.sourceSha256 !== sourceSha256
  ) {
    fail("PDF vector selection route is stale.");
  }
  const proofs = await readPageProofs(
    input.sourceBytes,
    input.route.pageNumber,
  );
  const candidate = [
    ...proofs.shapes.map(selectionFromShape),
    ...proofs.subpaths.map(selectionFromSubpath),
  ].filter(
    (selection) =>
      selection.proofFamily === input.route.proofFamily &&
      selection.proofId === input.route.proofId &&
      selection.nativeKind === input.route.nativeKind &&
      selection.paintMode === input.route.paintMode &&
      selection.proofDigest === input.route.proofDigest &&
      selection.contentStreamObjectNumber ===
        input.route.contentStreamObjectNumber &&
      selection.contentStreamGeneration ===
        input.route.contentStreamGeneration &&
      selection.operatorStart === input.route.operatorStart &&
      selection.operatorEnd === input.route.operatorEnd,
  );
  if (candidate.length !== 1) fail("PDF vector selection route is ambiguous.");
  const pageGeometry = exactPageGeometry(input.pageGeometry);
  const actualPageGeometry = await readActualPageGeometry(
    input.sourceBytes,
    input.route.pageNumber,
  );
  if (!samePageGeometry(pageGeometry, actualPageGeometry)) {
    fail("PDF vector route page geometry is stale.");
  }
  return Object.freeze({ selection: candidate[0], pageGeometry });
}

/**
 * Re-prove one selection route against the given bytes and return the write
 * selection it names. Callers that must address several objects across
 * consecutive writes turn these into history endpoints, which survive edits
 * to the other objects on the page.
 */
export async function resolvePdfNativeVectorRouteSelection(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<PdfNativeVectorWriteSelection> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  return resolved.selection;
}

/**
 * Resolve several routes on one page from a single inspection. A batch edit
 * addresses every selected object up front, and re-proving the page once per
 * route is what makes that step slow enough to miss its command window.
 */
export async function resolvePdfNativeVectorRouteSelections(input: {
  readonly sourceBytes: Uint8Array;
  readonly routes: readonly PdfNativeVectorSelectionRoute[];
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<readonly PdfNativeVectorWriteSelection[]> {
  if (input.routes.length === 0) return Object.freeze([]);
  const pageNumber = input.routes[0].pageNumber;
  if (input.routes.some((route) => route.pageNumber !== pageNumber)) {
    fail("PDF vector batch spans more than one page.");
  }
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (
    input.routes.some(
      (route) =>
        route.selectionOnly !== true || route.sourceSha256 !== sourceSha256,
    )
  ) {
    fail("PDF vector selection route is stale.");
  }
  const actualGeometry = await readActualPageGeometry(
    input.sourceBytes,
    pageNumber,
  );
  if (!samePageGeometry(exactPageGeometry(input.pageGeometry), actualGeometry)) {
    fail("PDF vector route page geometry is stale.");
  }
  const proofs = await readPageProofs(input.sourceBytes, pageNumber);
  const selections = [
    ...proofs.shapes.map(selectionFromShape),
    ...proofs.subpaths.map(selectionFromSubpath),
  ];
  return Object.freeze(
    input.routes.map((route) => {
      const candidates = selections.filter(
        (selection) =>
          selection.proofFamily === route.proofFamily &&
          selection.proofId === route.proofId &&
          selection.nativeKind === route.nativeKind &&
          selection.paintMode === route.paintMode &&
          selection.proofDigest === route.proofDigest &&
          selection.contentStreamObjectNumber ===
            route.contentStreamObjectNumber &&
          selection.contentStreamGeneration ===
            route.contentStreamGeneration &&
          selection.operatorStart === route.operatorStart &&
          selection.operatorEnd === route.operatorEnd,
      );
      if (candidates.length !== 1) {
        fail("PDF vector selection route is ambiguous.");
      }
      return candidates[0];
    }),
  );
}

export async function planPdfNativeVectorDisplayEdit(
  input: PdfNativeVectorDisplayEditInput,
): Promise<PdfNativeVectorSelectionEditPlan> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  const affine = affineFromDisplayBounds({
    selection: resolved.selection,
    pageGeometry: resolved.pageGeometry,
    routeBounds: input.fromBounds ?? input.route.bounds,
    targetBounds: input.targetBounds,
    mirror: input.mirror,
  });
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: resolved.pageGeometry,
      affine,
    }),
  });
}

function displayPointForPdfPoint(
  point: Readonly<{ x: number; y: number }>,
  geometry: PdfVectorPageGeometry,
) {
  const crop = geometry.cropBox;
  if (geometry.rotation === 0) {
    return { x: point.x - crop.xMin, y: crop.yMax - point.y };
  }
  if (geometry.rotation === 90) {
    return { x: point.y - crop.yMin, y: point.x - crop.xMin };
  }
  if (geometry.rotation === 180) {
    return { x: crop.xMax - point.x, y: point.y - crop.yMin };
  }
  return { x: crop.yMax - point.y, y: crop.xMax - point.x };
}

function pdfPointForDisplayPoint(
  point: Readonly<{ x: number; y: number }>,
  geometry: PdfVectorPageGeometry,
) {
  const crop = geometry.cropBox;
  if (geometry.rotation === 0) {
    return { x: crop.xMin + point.x, y: crop.yMax - point.y };
  }
  if (geometry.rotation === 90) {
    return { x: crop.xMin + point.y, y: crop.yMin + point.x };
  }
  if (geometry.rotation === 180) {
    return { x: crop.xMax - point.x, y: crop.yMin + point.y };
  }
  return { x: crop.xMax - point.y, y: crop.yMax - point.x };
}

function displayRotationAffine(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly rotationDegrees: number;
  /** Pivot in normalized display space; the selection's own centre when
   * absent. A group turns every member about the group's centre. */
  readonly center?: Readonly<{ x: number; y: number }>;
}): PdfNativeVectorAffineTransform {
  const degrees = exactFinite(input.rotationDegrees, "display rotation");
  const normalized = ((degrees % 360) + 360) % 360;
  if (normalized <= 0.000001 || 360 - normalized <= 0.000001) {
    fail("PDF vector rotation is unchanged.");
  }
  const geometry = exactPageGeometry(input.pageGeometry);
  // Renderer bounds are rounded and may include stroke/hit-target padding.
  // The native source geometry alone owns the rotation pivot and boundary.
  const bounds = projectedRawBounds(input.selection, geometry);
  const crop = geometry.cropBox;
  const pageWidth = crop.xMax - crop.xMin;
  const pageHeight = crop.yMax - crop.yMin;
  const displayWidth = geometry.rotation % 180 === 0 ? pageWidth : pageHeight;
  const displayHeight = geometry.rotation % 180 === 0 ? pageHeight : pageWidth;
  const pivot = input.center
    ? exactDisplayPoint(input.center, "rotation centre", CENTRE_REACH)
    : { x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 };
  const center = { x: pivot.x * displayWidth, y: pivot.y * displayHeight };
  const radians = (normalized * Math.PI) / 180;
  const cosine = Math.cos(radians);
  const sine = Math.sin(radians);
  const rotateDisplayPoint = (point: Readonly<{ x: number; y: number }>) => {
    const dx = point.x - center.x;
    const dy = point.y - center.y;
    return {
      x: center.x + cosine * dx - sine * dy,
      y: center.y + sine * dx + cosine * dy,
    };
  };
  const corners = [
    [bounds.x, bounds.y],
    [bounds.x + bounds.width, bounds.y],
    [bounds.x + bounds.width, bounds.y + bounds.height],
    [bounds.x, bounds.y + bounds.height],
  ].map(([x, y]) =>
    rotateDisplayPoint({ x: x * displayWidth, y: y * displayHeight }),
  );
  // A turn may take part of the shape past the page edge (WPS turns a
  // page-size background 30 degrees and the page cuts its corners); it may
  // not take all of it off.
  const turnedLeft = Math.min(...corners.map((point) => point.x));
  const turnedTop = Math.min(...corners.map((point) => point.y));
  if (
    !displayBoundsMeetPage(
      {
        x: turnedLeft / displayWidth,
        y: turnedTop / displayHeight,
        width: (Math.max(...corners.map((point) => point.x)) - turnedLeft) / displayWidth,
        height: (Math.max(...corners.map((point) => point.y)) - turnedTop) / displayHeight,
      },
      geometry,
    )
  ) {
    fail(PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE);
  }
  const map = (point: Readonly<{ x: number; y: number }>) => {
    const display = displayPointForPdfPoint(point, geometry);
    return pdfPointForDisplayPoint(
      rotateDisplayPoint(display),
      geometry,
    );
  };
  const origin = map({ x: 0, y: 0 });
  const xAxis = map({ x: 1, y: 0 });
  const yAxis = map({ x: 0, y: 1 });
  return exactAffine({
    a: xAxis.x - origin.x,
    b: xAxis.y - origin.y,
    c: yAxis.x - origin.x,
    d: yAxis.y - origin.y,
    e: origin.x,
    f: origin.y,
  });
}

/**
 * Mirror the selection across the vertical (horizontal flip) or horizontal
 * (vertical flip) axis through the centre of its own source bounds. The
 * reflection is expressed in display space so page /Rotate is honoured, then
 * conjugated back into PDF user space like the rotation affine.
 */
function displayFlipAffine(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly axis: PdfNativeVectorFlipAxis;
  /** Mirror axis position in normalized display space; the selection's own
   * centre when absent. A group mirrors every member about the group's. */
  readonly center?: Readonly<{ x: number; y: number }>;
}): PdfNativeVectorAffineTransform {
  if (input.axis !== "horizontal" && input.axis !== "vertical") {
    fail("PDF vector flip axis is invalid.");
  }
  const geometry = exactPageGeometry(input.pageGeometry);
  const bounds = projectedRawBounds(input.selection, geometry);
  const crop = geometry.cropBox;
  const pageWidth = crop.xMax - crop.xMin;
  const pageHeight = crop.yMax - crop.yMin;
  const displayWidth = geometry.rotation % 180 === 0 ? pageWidth : pageHeight;
  const displayHeight = geometry.rotation % 180 === 0 ? pageHeight : pageWidth;
  const pivot = input.center
    ? exactDisplayPoint(input.center, "flip centre", CENTRE_REACH)
    : { x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 };
  const center = { x: pivot.x * displayWidth, y: pivot.y * displayHeight };
  const flipDisplayPoint = (point: Readonly<{ x: number; y: number }>) =>
    input.axis === "horizontal"
      ? { x: 2 * center.x - point.x, y: point.y }
      : { x: point.x, y: 2 * center.y - point.y };
  const map = (point: Readonly<{ x: number; y: number }>) =>
    pdfPointForDisplayPoint(
      flipDisplayPoint(displayPointForPdfPoint(point, geometry)),
      geometry,
    );
  const origin = map({ x: 0, y: 0 });
  const xAxis = map({ x: 1, y: 0 });
  const yAxis = map({ x: 0, y: 1 });
  return exactAffine({
    a: xAxis.x - origin.x,
    b: xAxis.y - origin.y,
    c: yAxis.x - origin.x,
    d: yAxis.y - origin.y,
    e: origin.x,
    f: origin.y,
  });
}

function exactDisplayPoint(
  value: Readonly<{ x: number; y: number }>,
  label: string,
  /** How far past the page the point may lie (page fractions): a pointer is
   * on the page, the centre of a shape the page cuts may not be. */
  reach = 0,
): Readonly<{ x: number; y: number }> {
  const x = exactFinite(value.x, `${label} x`);
  const y = exactFinite(value.y, `${label} y`);
  const low = -reach - PAGE_EPSILON;
  const high = 1 + reach + PAGE_EPSILON;
  if (x < low || y < low || x > high || y > high) {
    fail(`PDF vector ${label} is outside the page.`);
  }
  return Object.freeze({ x, y });
}

/** How far past the page a turn's or flip's centre may lie (page widths). */
const CENTRE_REACH = 10;

/**
 * The affine one member of a group takes for the group's transform. A move or
 * resize of the group frame scales and translates every member by the frame's
 * own change, so members keep their places relative to each other; each
 * member may reach past the page edge, never wholly off it. A turn or flip is
 * the single-object one about the group's shared centre instead of the
 * member's own.
 */
function groupTransformAffine(input: {
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly transform: PdfNativeVectorGroupTransform;
}): PdfNativeVectorAffineTransform {
  const transform = input.transform;
  if (transform.kind === "rotation") {
    return displayRotationAffine({
      pageGeometry: input.pageGeometry,
      selection: input.selection,
      rotationDegrees: transform.rotationDegrees,
      center: transform.center,
    });
  }
  if (transform.kind === "flip") {
    return displayFlipAffine({
      pageGeometry: input.pageGeometry,
      selection: input.selection,
      axis: transform.axis,
      center: transform.center,
    });
  }
  const from = exactBounds(transform.fromBounds, "group bounds");
  const to = exactBounds(transform.toBounds, "group target bounds");
  if (from.width <= GEOMETRY_EPSILON || from.height <= GEOMETRY_EPSILON) {
    fail("PDF vector group bounds are degenerate.");
  }
  const scaleX = to.width / from.width;
  const scaleY = to.height / from.height;
  if (scaleX <= GEOMETRY_EPSILON || scaleY <= GEOMETRY_EPSILON) {
    fail("PDF vector group scale would collapse its geometry.");
  }
  const translateX = to.x - scaleX * from.x;
  const translateY = to.y - scaleY * from.y;
  const actual = projectedRawBounds(input.selection, input.pageGeometry);
  const next = {
    x: scaleX * actual.x + translateX,
    y: scaleY * actual.y + translateY,
    width: scaleX * actual.width,
    height: scaleY * actual.height,
  };
  if (!displayBoundsMeetPage(next, input.pageGeometry)) {
    fail(PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE);
  }
  return displayAffineToPdfUserSpace({
    pageGeometry: input.pageGeometry,
    scaleX,
    scaleY,
    translateX,
    translateY,
  });
}

/**
 * Plan one member's share of a group transform from a write selection already
 * proved against `sourceBytes` (a history endpoint resolved after the previous
 * member's write landed).
 */
export async function planPdfNativeVectorSelectionGroupTransform(input: {
  readonly sourceBytes: Uint8Array;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly transform: PdfNativeVectorGroupTransform;
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const pageGeometry = exactPageGeometry(input.pageGeometry);
  const actualPageGeometry = await readActualPageGeometry(
    input.sourceBytes,
    input.selection.pageNumber,
  );
  if (!samePageGeometry(pageGeometry, actualPageGeometry)) {
    fail("PDF vector route page geometry is stale.");
  }
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: input.selection,
      pageGeometry,
      affine: groupTransformAffine({
        pageGeometry,
        selection: input.selection,
        transform: input.transform,
      }),
    }),
  });
}

/** Plan one member's share of a group transform from its selection route. */
export async function planPdfNativeVectorDisplayGroupTransform(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly transform: PdfNativeVectorGroupTransform;
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: resolved.pageGeometry,
      affine: groupTransformAffine({
        pageGeometry: resolved.pageGeometry,
        selection: resolved.selection,
        transform: input.transform,
      }),
    }),
  });
}

export async function planPdfNativeVectorDisplayFlip(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly flipAxis: PdfNativeVectorFlipAxis;
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: resolved.pageGeometry,
      affine: displayFlipAffine({
        pageGeometry: resolved.pageGeometry,
        selection: resolved.selection,
        axis: input.flipAxis,
      }),
    }),
  });
}

export async function planPdfNativeVectorDisplayRotation(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly rotationDegrees: number;
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: resolved.pageGeometry,
      affine: displayRotationAffine({
        pageGeometry: resolved.pageGeometry,
        selection: resolved.selection,
        rotationDegrees: input.rotationDegrees,
      }),
    }),
  });
}






/** A turned frame resized along its own axes: one affine of the page as
 * shown (display points), baked into the shape's points as a turn is. A
 * rectangle stays a rectangle: the scale is along its own sides. */
export async function planPdfNativeVectorDisplayAffine(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly matrix: readonly [number, number, number, number, number, number];
}): Promise<PdfNativeVectorSelectionEditPlan> {
  const resolved = await resolvePdfNativeVectorDisplayRoute(input);
  return planPdfNativeVectorSelectionEdit({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: resolved.pageGeometry,
      affine: pdfNativeVectorDisplayMatrixAffine({
        pageGeometry: resolved.pageGeometry,
        matrix: input.matrix,
      }),
    }),
  });
}

// ---------------------------------------------------------------------------
// Duplicate and z-order (reorder) bridges. Both ride the geometry control and
// feature; each has its own action, schema and record kind so the server can
// validate the exact byte transform independently.
// ---------------------------------------------------------------------------

export type PdfNativeVectorDuplicateDirection = "insert" | "remove";
/** One step past the next object, or all the way to that end of the
 * drawing in one write (see the engine's layerMove). */
export type PdfNativeVectorReorderDirection = "forward" | "backward" | "front" | "back";

export interface PdfNativeVectorDuplicatePayload {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly direction: PdfNativeVectorDuplicateDirection;
  readonly offset: Readonly<{ dx: number; dy: number }>;
  /** Remove only: the exact bytes the insert appended. */
  readonly insertedBytes?: readonly number[];
}

export interface PdfNativeVectorDuplicateEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly direction: PdfNativeVectorDuplicateDirection;
  readonly proofFamily: PdfNativeVectorWriteSelection["proofFamily"];
  readonly nativeKind: PdfNativeVectorWriteSelection["nativeKind"];
  readonly paintMode: PdfNativeVectorWriteSelection["paintMode"];
  readonly paintOperator: string;
  readonly sourceProofId: string;
  readonly sourceProofDigest: string;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly offset: Readonly<{ dx: number; dy: number }>;
  readonly source: PdfNativeVectorWriteSelection;
  readonly copySelection: PdfNativeVectorWriteSelection | null;
  readonly insertedBytes: readonly number[];
  readonly siblingFingerprintCount: number;
  readonly siblingFingerprintSha256: string;
}

export interface PdfNativeVectorSelectionDuplicatePlan {
  readonly payload: PdfNativeVectorDuplicatePayload;
  readonly evidence: PdfNativeVectorDuplicateEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly beforeSelection: PdfNativeVectorWriteSelection;
  readonly beforeEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly copySelection: PdfNativeVectorWriteSelection | null;
  readonly copyEndpoint: PdfNativeVectorHistoryEndpoint | null;
  readonly insertedBytes: Uint8Array;
  readonly request: PdfNativeVectorDuplicatePlan["request"];
}

export interface PdfNativeVectorReorderPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly direction: PdfNativeVectorReorderDirection | "restore";
  /** Restore only: the neighbour the pair was swapped with. */
  readonly neighbour?: PdfNativeVectorWriteSelection;
  /** Restore only: the pair bytes currently in the stream. */
  readonly reorderedBytes?: readonly number[];
  /** Restore only: the pair bytes to put back. */
  readonly originalBytes?: readonly number[];
}

export interface PdfNativeVectorReorderEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly direction: PdfNativeVectorReorderDirection | "restore";
  readonly selection: PdfNativeVectorWriteSelection;
  /** The path object swapped with; null when the object passed other
   * content, named by `passed`. */
  readonly neighbour: PdfNativeVectorWriteSelection | null;
  /** The text object, group, XObject or shading the object was moved past
   * (stream offsets in the source), null for a pair swap or a restore. */
  readonly passed: PdfNativeVectorPassedUnit | null;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly operatorSpan: Readonly<{
    start: number;
    end: number;
    expectedSha256: string;
  }>;
  readonly originalBytes: readonly number[];
  readonly reorderedBytes: readonly number[];
  readonly siblingFingerprintCount: number;
  readonly siblingFingerprintSha256: string;
}

export interface PdfNativeVectorSelectionReorderPlan {
  readonly payload: PdfNativeVectorReorderPayload;
  readonly evidence: PdfNativeVectorReorderEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly beforeSelection: PdfNativeVectorWriteSelection;
  readonly afterSelection: PdfNativeVectorWriteSelection;
  readonly neighbourSelection: PdfNativeVectorWriteSelection | null;
  readonly beforeEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly afterEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly afterNeighbourEndpoint: PdfNativeVectorHistoryEndpoint | null;
  readonly originalBytes: Uint8Array;
  readonly reorderedBytes: Uint8Array;
  readonly request: PdfNativeVectorReorderPlan["request"];
  /** The object to select after the move, when it is not the one the plan
   * moved (a grouped shape whose group a neighbour swapped with). */
  readonly selectEndpoint?: PdfNativeVectorHistoryEndpoint;
}





async function resolvedRouteSelection(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
}) {
  return resolvePdfNativeVectorDisplayRoute(input);
}

/** A display-space shift as the PDF user-space offset a duplicate takes. */
function duplicateOffsetForPage(
  geometry: PdfVectorPageGeometry,
  shift: Readonly<{ x: number; y: number }>,
): Readonly<{ dx: number; dy: number }> {
  return pdfNativeVectorDisplayShiftOffset(geometry, shift);
}

export async function planPdfNativeVectorSelectionDuplicate(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorDuplicatePayload;
}): Promise<PdfNativeVectorSelectionDuplicatePlan> {
  const payload = readPdfNativeVectorDuplicatePayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (sourceSha256 !== payload.selection.sourceSha256) {
    fail("PDF vector duplicate selection source changed.");
  }
  const actualGeometry = await readActualPageGeometry(
    input.sourceBytes,
    payload.selection.pageNumber,
  );
  if (!samePageGeometry(actualGeometry, payload.pageGeometry)) {
    fail("PDF vector duplicate CropBox or rotation changed.");
  }
  const beforeSelection = payload.selection;
  projectedRawBounds(beforeSelection, actualGeometry);
  const lowLevel =
    payload.direction === "insert"
      ? await planMraiNativeVectorDuplicate({
          sourceBytes: input.sourceBytes,
          pageNumber: beforeSelection.pageNumber,
          selection: proofFromSelection(beforeSelection),
          offset: payload.offset,
        })
      : await planMraiNativeVectorDuplicateRemoval({
          sourceBytes: input.sourceBytes,
          pageNumber: beforeSelection.pageNumber,
          selection: proofFromSelection(beforeSelection),
          expectedInsertedBytes: Uint8Array.from(payload.insertedBytes ?? []),
        });
  const prepared = await prepareMraiNativeVectorDuplicate(
    input.sourceBytes,
    lowLevel,
  );
  const verified = await verifyMraiNativeVectorDuplicateCandidate(
    input.sourceBytes,
    lowLevel,
    prepared,
  );
  const copySelection = verified.reopenedCopy
    ? "kind" in verified.reopenedCopy
      ? selectionFromShape(verified.reopenedCopy)
      : selectionFromSubpath(verified.reopenedCopy)
    : null;
  // The copy may reach past the page edge (a paste at the pointer, WPS),
  // never wholly off it.
  if (copySelection && !boundsOnPage(copySelection, actualGeometry)) {
    prepared.candidateBytes.fill(0);
    fail(PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE);
  }
  const beforeSiblings = duplicatePlanSiblingRows(lowLevel);
  const pageInspection = await inspectMraiNativePageContentGraph(
    input.sourceBytes,
    beforeSelection.pageNumber,
  );
  const evidence: PdfNativeVectorDuplicateEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
    sourceSha256,
    pageNumber: beforeSelection.pageNumber,
    page: Object.freeze({ ...pageInspection.page }),
    pageGeometry: actualGeometry,
    direction: payload.direction,
    proofFamily: beforeSelection.proofFamily,
    nativeKind: beforeSelection.nativeKind,
    paintMode: beforeSelection.paintMode,
    paintOperator: lowLevel.paintOperator,
    sourceProofId: beforeSelection.proofId,
    sourceProofDigest: beforeSelection.proofDigest,
    contentStream: Object.freeze({
      objectNumber: beforeSelection.contentStreamObjectNumber,
      generation: beforeSelection.contentStreamGeneration,
    }),
    operatorSpan: Object.freeze({ ...lowLevel.request.operatorSpan }),
    offset: payload.offset,
    source: beforeSelection,
    copySelection,
    insertedBytes: Object.freeze([...lowLevel.insertedBytes]),
    siblingFingerprintCount: beforeSiblings.length,
    siblingFingerprintSha256: await siblingDigest(beforeSiblings),
  });
  return Object.freeze({
    payload,
    evidence,
    prepared,
    beforeSelection,
    beforeEndpoint: historyEndpoint(beforeSelection),
    copySelection,
    copyEndpoint: copySelection ? historyEndpoint(copySelection) : null,
    insertedBytes: lowLevel.insertedBytes,
    request: lowLevel.request,
  });
}

function duplicatePlanSiblingRows(
  plan: PdfNativeVectorDuplicatePlan,
): readonly string[] {
  const shapeRows = plan.beforeShapeFingerprints
    .filter(
      (_fingerprint, index) =>
        !("kind" in plan.selection) ||
        index !== plan.selection.pageVectorIndex,
    )
    .map((fingerprint) => `shape:${fingerprint}`);
  const subpathRows = plan.beforeSubpathFingerprints
    .filter(
      (_fingerprint, index) =>
        "kind" in plan.selection ||
        index !== plan.selection.pageSubpathIndex,
    )
    .map((fingerprint) => `subpath:${fingerprint}`);
  return Object.freeze([...shapeRows, ...subpathRows]);
}

export async function planPdfNativeVectorDisplayDuplicate(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly nudgePoints?: number;
  /** Where the copy goes instead (display points, right and down): a paste
   * at the pointer (WPS). */
  readonly shiftPoints?: Readonly<{ x: number; y: number }>;
}): Promise<PdfNativeVectorSelectionDuplicatePlan> {
  const resolved = await resolvedRouteSelection(input);
  return planPdfNativeVectorSelectionDuplicate({
    sourceBytes: input.sourceBytes,
    payload: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
      selection: resolved.selection,
      pageGeometry: resolved.pageGeometry,
      direction: "insert",
      offset: duplicateOffsetForPage(resolved.pageGeometry, pdfNativeVectorDuplicateShift(input)),
    }),
  });
}

/** The display shift (points, right and down) a duplicate is made at: the
 * one asked for, else WPS's nudge on both axes. */
function pdfNativeVectorDuplicateShift(input: {
  readonly nudgePoints?: number;
  readonly shiftPoints?: Readonly<{ x: number; y: number }>;
}): Readonly<{ x: number; y: number }> {
  if (input.shiftPoints) {
    return Object.freeze({
      x: exactFinite(input.shiftPoints.x, "duplicate shift x"),
      y: exactFinite(input.shiftPoints.y, "duplicate shift y"),
    });
  }
  const nudge = input.nudgePoints ?? PDF_NATIVE_VECTOR_DUPLICATE_NUDGE_POINTS;
  return Object.freeze({ x: nudge, y: nudge });
}

function reorderPlanSiblingRows(
  plan: PdfNativeVectorReorderPlan,
): readonly string[] {
  const excludedShape = new Set<number>();
  const excludedSubpath = new Set<number>();
  for (const entry of [plan.selection, plan.neighbour]) {
    if (!entry) continue;
    if ("kind" in entry) excludedShape.add(entry.pageVectorIndex);
    else excludedSubpath.add(entry.pageSubpathIndex);
  }
  return Object.freeze([
    ...plan.beforeShapeFingerprints
      .filter((_fingerprint, index) => !excludedShape.has(index))
      .map((fingerprint) => `shape:${fingerprint}`),
    ...plan.beforeSubpathFingerprints
      .filter((_fingerprint, index) => !excludedSubpath.has(index))
      .map((fingerprint) => `subpath:${fingerprint}`),
  ]);
}

export async function planPdfNativeVectorSelectionReorder(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorReorderPayload;
}): Promise<PdfNativeVectorSelectionReorderPlan> {
  const payload = readPdfNativeVectorReorderPayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (sourceSha256 !== payload.selection.sourceSha256) {
    fail("PDF vector reorder selection source changed.");
  }
  const actualGeometry = await readActualPageGeometry(
    input.sourceBytes,
    payload.selection.pageNumber,
  );
  if (!samePageGeometry(actualGeometry, payload.pageGeometry)) {
    fail("PDF vector reorder CropBox or rotation changed.");
  }
  const beforeSelection = payload.selection;
  // A reorder moves bytes, never geometry: a shape the page cuts keeps its
  // place and its clip, so it may change layers as long as part of it shows
  // (the server reaches the same verdict).
  projectedRawBounds(beforeSelection, actualGeometry);
  const lowLevel =
    payload.direction === "restore"
      ? await planMraiNativeVectorReorderRestore({
          sourceBytes: input.sourceBytes,
          pageNumber: beforeSelection.pageNumber,
          selection: proofFromSelection(beforeSelection),
          neighbour: payload.neighbour ? proofFromSelection(payload.neighbour) : null,
          reorderedBytes: Uint8Array.from(payload.reorderedBytes!),
          originalBytes: Uint8Array.from(payload.originalBytes!),
        })
      : await planMraiNativeVectorReorder({
          sourceBytes: input.sourceBytes,
          pageNumber: beforeSelection.pageNumber,
          selection: proofFromSelection(beforeSelection),
          direction: payload.direction,
        });
  const prepared = await prepareMraiNativeVectorReorder(
    input.sourceBytes,
    lowLevel,
  );
  const verified = await verifyMraiNativeVectorReorderCandidate(
    input.sourceBytes,
    lowLevel,
    prepared,
  );
  const neighbourSelection = lowLevel.neighbour
    ? "kind" in lowLevel.neighbour
      ? selectionFromShape(lowLevel.neighbour)
      : selectionFromSubpath(lowLevel.neighbour)
    : null;
  const afterSelection =
    "kind" in verified.reopenedSelection
      ? selectionFromShape(verified.reopenedSelection)
      : selectionFromSubpath(verified.reopenedSelection);
  const afterNeighbourEndpoint = neighbourSelection
    ? historyEndpoint(
        Object.freeze({
          ...neighbourSelection,
          pageProofIndex:
            neighbourSelection.proofFamily === beforeSelection.proofFamily
              ? beforeSelection.pageProofIndex
              : neighbourSelection.pageProofIndex,
        }),
      )
    : null;
  const beforeSiblings = reorderPlanSiblingRows(lowLevel);
  const pageInspection = await inspectMraiNativePageContentGraph(
    input.sourceBytes,
    beforeSelection.pageNumber,
  );
  const evidence: PdfNativeVectorReorderEvidence = Object.freeze({
    schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
    sourceSha256,
    pageNumber: beforeSelection.pageNumber,
    page: Object.freeze({ ...pageInspection.page }),
    pageGeometry: actualGeometry,
    direction: payload.direction,
    selection: beforeSelection,
    neighbour: neighbourSelection,
    passed: lowLevel.passed,
    contentStream: Object.freeze({
      objectNumber: beforeSelection.contentStreamObjectNumber,
      generation: beforeSelection.contentStreamGeneration,
    }),
    operatorSpan: Object.freeze({ ...lowLevel.request.operatorSpan }),
    // Evidence always names the pair layout as `originalBytes` and the
    // wrapped layout as `reorderedBytes`; a restore reads the wrapped bytes
    // and writes the pair back, so its plan fields are swapped here.
    originalBytes: Object.freeze([
      ...(payload.direction === "restore"
        ? lowLevel.reorderedBytes
        : lowLevel.originalOperatorBytes),
    ]),
    reorderedBytes: Object.freeze([
      ...(payload.direction === "restore"
        ? lowLevel.originalOperatorBytes
        : lowLevel.reorderedBytes),
    ]),
    siblingFingerprintCount: beforeSiblings.length,
    siblingFingerprintSha256: await siblingDigest(beforeSiblings),
  });
  return Object.freeze({
    payload,
    evidence,
    prepared,
    beforeSelection,
    afterSelection,
    neighbourSelection,
    beforeEndpoint: historyEndpoint(beforeSelection),
    afterEndpoint: historyEndpoint(afterSelection),
    afterNeighbourEndpoint,
    originalBytes: lowLevel.originalOperatorBytes,
    reorderedBytes: lowLevel.reorderedBytes,
    request: lowLevel.request,
  });
}

export async function planPdfNativeVectorDisplayReorder(input: {
  readonly sourceBytes: Uint8Array;
  readonly route: PdfNativeVectorSelectionRoute;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly direction: PdfNativeVectorReorderDirection;
}): Promise<PdfNativeVectorSelectionReorderPlan> {
  const resolved = await resolvedRouteSelection(input);
  try {
    return await planPdfNativeVectorSelectionReorder({
      sourceBytes: input.sourceBytes,
      payload: Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
        selection: resolved.selection,
        pageGeometry: resolved.pageGeometry,
        direction: input.direction,
      }),
    });
  } catch (error) {
    const grouped = await planGroupedVectorReorder(input, resolved.selection, resolved.pageGeometry, error);
    if (grouped) return grouped;
    throw error;
  }
}












async function planGroupedVectorReorder(
  input: Readonly<{ sourceBytes: Uint8Array; direction: PdfNativeVectorReorderDirection }>,
  selection: PdfNativeVectorWriteSelection,
  pageGeometry: PdfVectorPageGeometry,
  error: unknown,
): Promise<(PdfNativeVectorSelectionReorderPlan & { readonly selectEndpoint: PdfNativeVectorHistoryEndpoint }) | null> {
  if (input.direction !== "forward" && input.direction !== "backward") return null;
  if (!(error instanceof Error) || (error.message !== "PDF vector is already the topmost path on its page." &&
    error.message !== "PDF vector is already the bottommost path on its page.")) return null;
  const sameShape = (shape: PdfNativeVectorContentShape, of: Readonly<{ contentStreamObjectNumber: number;
    contentStreamGeneration: number; operatorStart: number; operatorEnd: number; formInstance?: string }>) =>
    shape.contentStream.objectNumber === of.contentStreamObjectNumber &&
    shape.contentStream.generation === of.contentStreamGeneration &&
    shape.operatorSpan.start === of.operatorStart && shape.operatorSpan.end === of.operatorEnd &&
    (shape.formInstance ?? undefined) === (of.formInstance ?? undefined);
  const shapes = (await inspectMraiNativeVectorContent(input.sourceBytes, selection.pageNumber)).shapes;
  const index = shapes.findIndex((shape) => sameShape(shape, selection));
  if (index < 0) return null;
  const ours = shapes[index];
  const neighbour = shapes[input.direction === "backward" ? index - 1 : index + 1];
  if (!neighbour || (neighbour.contentStream.objectNumber === ours.contentStream.objectNumber &&
    neighbour.contentStream.generation === ours.contentStream.generation)) return null;
  let plan: PdfNativeVectorSelectionReorderPlan;
  try {
    plan = await planPdfNativeVectorSelectionReorder({
      sourceBytes: input.sourceBytes,
      payload: Object.freeze({
        schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
        selection: selectionFromShape(neighbour),
        pageGeometry,
        direction: input.direction === "backward" ? "forward" : "backward",
      }),
    });
  } catch {
    return null;
  }
  // The swap passes exactly one Form invocation, and it is the group's own:
  // the name it invokes is the page's resource for the group's stream.
  const latin1 = (bytes: Uint8Array) => { let text = ""; for (const byte of bytes) text += String.fromCharCode(byte); return text; };
  const invocation = /\/([^\s/\[\]<>(){}%]+)\s+Do(?=[\s\[<(/%]|$)/gu;
  const original = latin1(plan.originalBytes), reordered = latin1(plan.reorderedBytes);
  const invoked = [...original.matchAll(invocation)];
  const reject = () => { plan.prepared.candidateBytes.fill(0); return null; };
  if (invoked.length !== 1) return reject();
  try {
    const document = await PDFDocument.load(Uint8Array.from(input.sourceBytes), { updateMetadata: false, ignoreEncryption: false });
    const page = document.getPage(selection.pageNumber - 1);
    const xobjects = page.node.Resources()?.lookupMaybe(PDFName.of("XObject"), PDFDict);
    const ref = xobjects?.get(PDFName.of(invoked[0][1]));
    if (!(ref instanceof PDFRef) || ref.objectNumber !== ours.contentStream.objectNumber ||
      ref.generationNumber !== ours.contentStream.generation) return reject();
  } catch {
    return reject();
  }
  // And the move puts the group on the side it was asked to go.
  const before = original.search(invocation), now = reordered.search(new RegExp(invocation.source, "u"));
  if (before < 0 || now < 0 || (input.direction === "backward" ? now >= before : now <= before)) return reject();
  const after = (await inspectMraiNativeVectorContent(plan.prepared.candidateBytes, selection.pageNumber)).shapes;
  const oursAfter = after.find((shape) => sameShape(shape, selection));
  if (!oursAfter) return reject();
  return Object.freeze({ ...plan, selectEndpoint: historyEndpoint(selectionFromShape(oursAfter)) });
}



