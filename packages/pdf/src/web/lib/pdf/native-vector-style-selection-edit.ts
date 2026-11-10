/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION, PDF_NATIVE_VECTOR_STYLE_CONTROL_ID, PDF_NATIVE_VECTOR_STYLE_FEATURE_ID, buildPdfNativeVectorStyleCommand, changes, createPdfNativeVectorStyleOwnershipManifest, invalid, pageGeometry, readPdfNativeVectorStyleSelectionEditPayload, route, samePdfNativeVectorObjectStyle } from "./native-vector-style-commands";
export { MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION, PDF_NATIVE_VECTOR_STYLE_CONTROL_ID, PDF_NATIVE_VECTOR_STYLE_FEATURE_ID, buildPdfNativeVectorStyleCommand, createPdfNativeVectorStyleOwnershipManifest, readPdfNativeVectorStyleSelectionEditPayload, samePdfNativeVectorObjectStyle };
import {
  MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION,
} from "./engine/native/vocabulary";
export {
  MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION,
};
import type { PdfNativeVectorContentShape } from "./engine/native/vector-content-transaction";

import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import {
  isPdfStandardBlendMode,
  type PdfStandardBlendMode,
} from "./blend-modes";
import type {
  PdfEngineCommand,
  PdfEngineSourceIdentity,
} from "./engine/contracts";
import { readPdfCommandOwnershipManifest } from "./engine/command-ownership";
import {
  inspectMraiNativePageContentGraph,
  sha256PdfBytes,
  type PdfNativePreparedContentStreamTransaction,
} from "./engine/native/content-stream-transaction";
import { inspectMraiNativeVectorDocumentAuthority } from "./engine/native/vector-content-transaction";
import {
  MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
  inspectMraiNativeVectorObjectStyleDocument,
  planMraiNativeVectorObjectStyleEdit,
  prepareMraiNativeVectorObjectStyleEdit,
  verifyMraiNativeVectorObjectStyleCandidate,
  type PdfNativeVectorObjectStyle,
  type PdfNativeVectorObjectStyleChanges,
  type PdfNativeVectorObjectStyleProof,
  type PdfNativeVectorRgbColor,
} from "./engine/native/vector-style-transaction";
import {
  createPdfNativeVectorHistoryEndpoint,
  createPdfNativeVectorWriteSelection,
  pdfNativeVectorSelectionOnPage,
  readPdfNativeSourcePageGeometries,
  resolvePdfNativeVectorHistorySelection,
  type PdfNativeVectorHistoryEndpoint,
  type PdfNativeVectorWriteSelection,
} from "./native-vector-selection-edit";
import { PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE } from "./native-vector-commands";
import type {
  PdfNativeVectorSelectionRoute,
  PdfNativeVectorStyleSelectionEvidence,
} from "./semantic-appearance-inspection";
import {
  projectPdfNativeVectorBounds,
  type PdfVectorPageGeometry,
} from "./vector-capability-reconciliation";



export interface PdfNativeVectorStyleSelectionEditPayload {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION;
  readonly route: PdfNativeVectorSelectionRoute & {
    readonly style: PdfNativeVectorStyleSelectionEvidence;
  };
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly changes: PdfNativeVectorObjectStyleChanges;
}

export interface PdfNativeVectorStyleSelectionEditEvidence {
  readonly schemaVersion: typeof MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION;
  readonly sourceSha256: string;
  readonly pageNumber: number;
  readonly page: Readonly<{ objectNumber: number; generation: number }>;
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly proofFamily: PdfNativeVectorSelectionRoute["proofFamily"];
  readonly pageProofIndex: number;
  readonly nativeKind: PdfNativeVectorSelectionRoute["nativeKind"];
  readonly paintMode: PdfNativeVectorSelectionRoute["paintMode"];
  readonly sourceProofId: string;
  readonly sourceProofDigest: string;
  readonly styleProofId: string;
  readonly styleProofDigest: string;
  readonly selection: PdfNativeVectorWriteSelection;
  readonly contentStream: Readonly<{
    objectNumber: number;
    generation: number;
  }>;
  readonly pathSpan: PdfNativeVectorObjectStyleProof["pathSpan"];
  readonly paintSpan: PdfNativeVectorObjectStyleProof["paintSpan"];
  readonly before: PdfNativeVectorObjectStyle;
  readonly after: PdfNativeVectorObjectStyle;
  readonly changes: PdfNativeVectorObjectStyleChanges;
  readonly siblingFingerprintCount: number;
  readonly siblingFingerprintSha256: string;
}

export interface PdfNativeVectorStyleSelectionEditPlan {
  readonly payload: PdfNativeVectorStyleSelectionEditPayload;
  readonly evidence: PdfNativeVectorStyleSelectionEditEvidence;
  readonly prepared: PdfNativePreparedContentStreamTransaction;
  readonly beforeProof: PdfNativeVectorObjectStyleProof;
  readonly afterProof: PdfNativeVectorObjectStyleProof;
  readonly beforeEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly afterEndpoint: PdfNativeVectorHistoryEndpoint;
  readonly inverseChanges: PdfNativeVectorObjectStyleChanges;
}

export interface PdfNativeVectorStyleCommitRequest {
  readonly route: PdfNativeVectorStyleSelectionEditPayload["route"];
  readonly pageGeometry: PdfVectorPageGeometry;
  readonly changes: PdfNativeVectorObjectStyleChanges;
}












/** Crop box and /Rotate of one page, read from the pinned source bytes. */
export async function readPdfNativeVectorPageGeometry(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfVectorPageGeometry> {
  return actualPageGeometry(sourceBytes, pageNumber);
}

async function actualPageGeometry(
  sourceBytes: Uint8Array,
  pageNumber: number,
): Promise<PdfVectorPageGeometry> {
  const page = (await readPdfNativeSourcePageGeometries(sourceBytes))[pageNumber - 1];
  if (!page) invalid("PDF vector style page is unavailable.");
  return pageGeometry({
    cropBox: { ...page.cropBox },
    rotation: page.rotation,
  });
}

function samePageGeometry(
  left: PdfVectorPageGeometry,
  right: PdfVectorPageGeometry,
): boolean {
  return (
    left.rotation === right.rotation &&
    (Object.keys(left.cropBox) as Array<keyof typeof left.cropBox>).every(
      (key) => Math.abs(left.cropBox[key] - right.cropBox[key]) <= 0.000_001,
    )
  );
}

function sameStyle(
  left: PdfNativeVectorObjectStyle,
  right: PdfNativeVectorObjectStyle,
): boolean {
  
  // carried by the engine's proof and not by every route that reaches here
  // (a route sanitised before it, a route written before the field): a
  // field absent reads as false, so that the same style is the same style.
  // Compared as written, every style edit of a vector was refused as stale
  // from the field's first day.
  const canonical = ({ fillUnreadable, ...rest }: PdfNativeVectorObjectStyle) =>
    JSON.stringify({ ...rest, fillUnreadable: fillUnreadable === true });
  return canonical(left) === canonical(right);
}


export async function resolvePdfNativeVectorStyleHistoryRoute(input: {
  readonly sourceBytes: Uint8Array;
  readonly endpoint: PdfNativeVectorHistoryEndpoint;
  readonly pageGeometry: PdfVectorPageGeometry;
}): Promise<PdfNativeVectorStyleSelectionEditPayload["route"]> {
  const selection = await resolvePdfNativeVectorHistorySelection({
    sourceBytes: input.sourceBytes,
    endpoint: input.endpoint,
  });
  const authority = await inspectMraiNativeVectorDocumentAuthority(
    input.sourceBytes,
  );
  const styles = await inspectMraiNativeVectorObjectStyleDocument(
    input.sourceBytes,
    authority,
  );
  const proof = styles.pages[selection.pageNumber - 1]?.proofs.find(
    (candidate) =>
      candidate.proofFamily === selection.proofFamily &&
      candidate.pageProofIndex === selection.pageProofIndex &&
      candidate.selectionId === selection.proofId &&
      candidate.selectionProofDigest === selection.proofDigest,
  );
  const bounds = projectPdfNativeVectorBounds(
    selection.bounds,
    input.pageGeometry,
  );
  if (!proof || !bounds)
    invalid("PDF vector style history is stale or ambiguous.");
  return Object.freeze({
    selectionOnly: true,
    sourceSha256: selection.sourceSha256,
    pageNumber: selection.pageNumber,
    rendererOperatorIndex: 0,
    proofFamily: selection.proofFamily,
    nativeKind: selection.nativeKind,
    paintMode: selection.paintMode,
    proofId: selection.proofId,
    proofDigest: selection.proofDigest,
    contentStreamObjectNumber: selection.contentStreamObjectNumber,
    contentStreamGeneration: selection.contentStreamGeneration,
    operatorStart: selection.operatorStart,
    operatorEnd: selection.operatorEnd,
    style: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_OBJECT_STYLE_SCHEMA_VERSION,
      proofId: proof.id,
      proofDigest: proof.proofDigest,
      pageProofIndex: proof.pageProofIndex,
      style: proof.style,
    }),
    bounds,
  });
}

function inverseChanges(
  before: PdfNativeVectorObjectStyle,
  requested: PdfNativeVectorObjectStyleChanges,
  beforePaintMode?: PdfNativeVectorContentShape["paintMode"],
): PdfNativeVectorObjectStyleChanges {
  // Across a repaint a colour channel the source never defined (null) is
  // restored by the paint mode alone: the hidden channel simply stops being
  // painted, so no colour operator is written back for it.
  return Object.freeze({
    ...(requested.paintMode !== undefined && beforePaintMode
      ? { paintMode: beforePaintMode }
      : {}),
    ...(requested.fillColor && before.fillColor
      ? { fillColor: before.fillColor }
      : {}),
    ...(requested.strokeColor && before.strokeColor
      ? { strokeColor: before.strokeColor }
      : {}),
    ...(requested.lineWidth !== undefined
      ? { lineWidth: before.lineWidth! }
      : {}),
    ...(requested.dashArray !== undefined
      ? { dashArray: before.dashArray! }
      : {}),
    ...(requested.dashPhase !== undefined
      ? { dashPhase: before.dashPhase! }
      : {}),
    ...(requested.lineCap !== undefined ? { lineCap: before.lineCap! } : {}),
    ...(requested.lineJoin !== undefined ? { lineJoin: before.lineJoin! } : {}),
    ...(requested.miterLimit !== undefined
      ? { miterLimit: before.miterLimit! }
      : {}),
    ...(requested.opacity !== undefined ? { opacity: before.opacity! } : {}),
    ...(requested.blendMode !== undefined
      ? { blendMode: before.blendMode! }
      : {}),
    ...(requested.translate
      ? {
          translate: Object.freeze({
            x: -requested.translate.x,
            y: -requested.translate.y,
          }),
        }
      : {}),
    ...(requested.transform
      ? { transform: invertedAffine(requested.transform) }
      : {}),
  });
}

/** The page transform that undoes `affine` (PDF convention). */
function invertedAffine(
  affine: NonNullable<PdfNativeVectorObjectStyleChanges["transform"]>,
): NonNullable<PdfNativeVectorObjectStyleChanges["transform"]> {
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

export async function planPdfNativeVectorStyleSelectionEdit(input: {
  readonly sourceBytes: Uint8Array;
  readonly payload: PdfNativeVectorStyleSelectionEditPayload;
}): Promise<PdfNativeVectorStyleSelectionEditPlan> {
  const payload = readPdfNativeVectorStyleSelectionEditPayload(input.payload);
  const sourceSha256 = await sha256PdfBytes(input.sourceBytes);
  if (sourceSha256 !== payload.route.sourceSha256) {
    invalid("PDF vector style source changed.");
  }
  const actualGeometry = await actualPageGeometry(
    input.sourceBytes,
    payload.route.pageNumber,
  );
  if (!samePageGeometry(actualGeometry, payload.pageGeometry)) {
    invalid("PDF vector style CropBox or rotation changed.");
  }
  const lowLevel = await planMraiNativeVectorObjectStyleEdit({
    sourceBytes: input.sourceBytes,
    pageNumber: payload.route.pageNumber,
    proofId: payload.route.style.proofId,
    changes: payload.changes,
  });
  const proof = lowLevel.proof;
  const selection = lowLevel.selection;
  const expectedNativeKind = "kind" in selection ? selection.kind : "subpath";
  if (
    proof.proofDigest !== payload.route.style.proofDigest ||
    proof.pageProofIndex !== payload.route.style.pageProofIndex ||
    !sameStyle(proof.style, payload.route.style.style) ||
    proof.selectionId !== payload.route.proofId ||
    proof.selectionProofDigest !== payload.route.proofDigest ||
    proof.proofFamily !== payload.route.proofFamily ||
    expectedNativeKind !== payload.route.nativeKind ||
    proof.paintMode !== payload.route.paintMode ||
    proof.contentStream.objectNumber !==
      payload.route.contentStreamObjectNumber ||
    proof.contentStream.generation !== payload.route.contentStreamGeneration ||
    proof.pathSpan.start !== payload.route.operatorStart ||
    proof.pathSpan.end !== payload.route.operatorEnd
  ) {
    invalid("PDF vector style route is stale or ambiguous.");
  }
  const prepared = await prepareMraiNativeVectorObjectStyleEdit(
    input.sourceBytes,
    lowLevel,
  );
  const verified = await verifyMraiNativeVectorObjectStyleCandidate(
    input.sourceBytes,
    lowLevel,
    prepared,
  );
  const page = await inspectMraiNativePageContentGraph(
    input.sourceBytes,
    payload.route.pageNumber,
  );
  const siblingFingerprintSha256 = await sha256PdfBytes(
    new TextEncoder().encode(
      JSON.stringify(lowLevel.beforeSiblingFingerprints),
    ),
  );
  const beforeSelection = createPdfNativeVectorWriteSelection(
    lowLevel.selection,
  );
  const afterSelection = createPdfNativeVectorWriteSelection(
    verified.reopenedSelection,
  );
  // A resize may take part of the shape past the page edge (the page cuts
  // it, WPS), never all of it (the replay holds the record to the same rule).
  if (
    payload.changes.transform &&
    !pdfNativeVectorSelectionOnPage(afterSelection, actualGeometry)
  ) {
    prepared.candidateBytes.fill(0);
    invalid(PDF_NATIVE_VECTOR_LEAVES_PAGE_MESSAGE);
  }
  return Object.freeze({
    payload,
    evidence: Object.freeze({
      schemaVersion: MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION,
      sourceSha256,
      pageNumber: payload.route.pageNumber,
      page: page.page,
      pageGeometry: actualGeometry,
      proofFamily: proof.proofFamily,
      pageProofIndex: proof.pageProofIndex,
      nativeKind: payload.route.nativeKind,
      paintMode: proof.paintMode,
      sourceProofId: proof.selectionId,
      sourceProofDigest: proof.selectionProofDigest,
      styleProofId: proof.id,
      styleProofDigest: proof.proofDigest,
      selection: beforeSelection,
      contentStream: proof.contentStream,
      pathSpan: proof.pathSpan,
      paintSpan: proof.paintSpan,
      before: lowLevel.beforeStyle,
      after: lowLevel.expectedStyle,
      // A translation or a resize is recorded as written: the object's own
      // `cm` operands.
      changes: lowLevel.writtenChanges,
      siblingFingerprintCount: lowLevel.beforeSiblingFingerprints.length,
      siblingFingerprintSha256,
    }),
    prepared,
    beforeProof: proof,
    afterProof: verified.reopenedProof,
    beforeEndpoint: createPdfNativeVectorHistoryEndpoint(beforeSelection),
    afterEndpoint: createPdfNativeVectorHistoryEndpoint(afterSelection),
    inverseChanges: inverseChanges(
      lowLevel.beforeStyle,
      payload.changes,
      proof.paintMode,
    ),
  });
}


