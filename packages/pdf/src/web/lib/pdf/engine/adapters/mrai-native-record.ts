/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The shape of a native PDF record (what the journal stores and the server
 * replays) built from a planned candidate: the content-stream fields every
 * native record carries, and the binders the editor's working session and
 * author recovery use to bind a locally held plan to its durable base. Pure
 * JSON shaping -- no PDF reading or writing -- so the editor keeps it while
 * the planners run on the server.
 */

import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import type { PdfEngineCommand, PdfEngineSourceIdentity } from "../contracts";
import type { PdfSharedStateRecord } from "../shared-state-bridge";
import type {
  PdfNativeEmbeddedImageRemovalEvidence,
  PdfNativeEmbeddedImageTransformEvidence,
  PdfNativeCatalogVersionRaise,
  PdfNativePdfaIdentificationRemoval,
  PdfNativePreparedContentStreamTransaction,
  PdfNativePreparedCatalogVersionRaise,
  PdfNativePreparedPdfaIdentificationRemoval,
} from "../native/content-stream-transaction";
import type {
  PdfNativeListTextFidelityEvidence,
  PdfNativeTextWindowHistoryEvidence,
} from "../native/list-text-fidelity";
import {
  MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
  MRAI_NATIVE_LIST_TEXT_ACTION,
  MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION,
  PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA,
} from "../native/vocabulary";

/** The durable base a local plan was made against, as the editor states it:
 * its identity and the revision depth (xref sections) of its bytes. */
export type PdfNativeRecordCheckpoint = Readonly<{
  source: PdfEngineSourceIdentity;
  xrefSections: number;
}>;

function recordCheckpointClaim(
  checkpoint: PdfNativeRecordCheckpoint,
  source: PdfEngineSourceIdentity,
  prepared: PdfNativePreparedContentStreamTransaction,
): Readonly<{ schemaVersion: typeof PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA; xrefSections: number }> | null {
  if (
    checkpoint.source.resourceId !== source.resourceId ||
    checkpoint.source.resourceVersionId !== source.resourceVersionId ||
    checkpoint.source.sourceSha256 !== source.sourceSha256 ||
    !Number.isSafeInteger(prepared.sourceXrefSections) ||
    checkpoint.xrefSections !== prepared.sourceXrefSections
  ) return null;
  return Object.freeze({ schemaVersion: PDF_NATIVE_SOURCE_CHECKPOINT_SCHEMA, xrefSections: checkpoint.xrefSections });
}






export function pdfNativeStructActualTextRecordField(
  request: Pick<PdfNativePreparedContentStreamTransaction["request"], "structActualText">,
): { structActualText?: JsonValue[] } {
  return request.structActualText
    ? { structActualText: request.structActualText.map((revision): JsonValue => ({
        objectNumber: revision.objectNumber,
        generation: revision.generation,
        expectedObjectSha256: revision.expectedObjectSha256,
        page: { objectNumber: revision.page.objectNumber, generation: revision.page.generation },
        mcids: [...revision.mcids],
        previous: Array.from(revision.previous),
        value: Array.from(revision.value),
        text: revision.text,
      })) }
    : {};
}







export function pdfNativePdfaIdentificationRecordField(
  prepared: Readonly<{
    pdfaIdentificationRemoval?: PdfNativePdfaIdentificationRemoval;
    catalogVersionRaise?: PdfNativeCatalogVersionRaise;
  }>,
): { pdfaIdentificationRemoval?: JsonValue; catalogVersionRaise?: JsonValue } {
  const removal = prepared.pdfaIdentificationRemoval;
  // The catalog stating the version the write needs (user audit PDF 0020)
  // rides the same records, beside the packet.
  const raised = prepared.catalogVersionRaise;
  return {
    ...(removal
      ? { pdfaIdentificationRemoval: {
          objectNumber: removal.objectNumber,
          generation: removal.generation,
          expectedObjectSha256: removal.expectedObjectSha256,
        } }
      : {}),
    ...(raised
      ? { catalogVersionRaise: {
          objectNumber: raised.objectNumber,
          generation: raised.generation,
          expectedObjectSha256: raised.expectedObjectSha256,
          version: raised.version,
        } }
      : {}),
  };
}

/** The packet's and the raised catalog's places in the revision, as a
 * receipt's revised objects name them (nothing when the write writes
 * neither). */
export function pdfNativePdfaIdentificationRevisedObjects(
  prepared: Readonly<{
    pdfaIdentificationRemoval?: PdfNativePreparedPdfaIdentificationRemoval;
    catalogVersionRaise?: PdfNativePreparedCatalogVersionRaise;
  }>,
): Readonly<{ objectNumber: number; generation: number; offset: number; sha256: string }>[] {
  return [prepared.pdfaIdentificationRemoval, prepared.catalogVersionRaise].flatMap((item) =>
    item ? [{ objectNumber: item.objectNumber, generation: item.generation, offset: item.offset, sha256: item.sha256 }] : [],
  );
}

export function stableRecord(
  command: PdfEngineCommand,
  prepared: PdfNativePreparedContentStreamTransaction,
  listFidelity?: PdfNativeListTextFidelityEvidence,
  embeddedImageRemoval?: PdfNativeEmbeddedImageRemovalEvidence,
  embeddedImageTransform?: PdfNativeEmbeddedImageTransformEvidence,
): PdfSharedStateRecord {
  const request = prepared.request;
  return {
    id: command.transactionId,
    schemaVersion:
      listFidelity?.schemaVersion ??
      embeddedImageRemoval?.schemaVersion ??
      embeddedImageTransform?.schemaVersion ??
      MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    transactionSchemaVersion: MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    kind: listFidelity
      ? "text-replacement"
      : embeddedImageRemoval
        ? "embedded-image-removal"
        : embeddedImageTransform
          ? "embedded-image-transform"
          : "content-stream-replacement",
    featureId: command.featureId,
    controlId: command.controlId,
    action: command.action,
    transactionId: command.transactionId,
    baseSessionRevision: command.sessionRevision,
    selectionRevision: command.selectionRevision,
    source: {
      resourceId: command.resourceId,
      resourceVersionId: command.resourceVersionId,
      sha256: request.identity.sourceSha256,
    },
    target: {
      objectNumber: request.identity.objectNumber,
      generation: request.identity.generation,
      expectedObjectSha256: request.expectedObjectSha256,
      expectedStreamSha256: request.expectedStreamSha256,
    },
    operatorSpan: {
      start: request.operatorSpan.start,
      end: request.operatorSpan.end,
      expectedSha256: request.operatorSpan.expectedSha256,
    },
    // A picture move names the bytes of its span by range (the server puts
    // them back together from the span it pins).
    ...(request.replacementPieces
      ? {
          replacement: null,
          replacementPieces: request.replacementPieces.map(
            (piece): JsonValue =>
              "literal" in piece
                ? { literal: Array.from(piece.literal) }
                : { spanStart: piece.spanStart, spanEnd: piece.spanEnd },
          ),
        }
      : { replacement: Array.from(request.replacement) }),
    ...(request.duplicateSpans
      ? { duplicateSpans: request.duplicateSpans.map((span): JsonValue => ({ start: span.start, end: span.end,
          ...(span.sourceStart !== undefined
            ? { sourceStart: span.sourceStart, sourceEnd: span.sourceEnd!,
                replacementStart: span.replacementStart!, replacementEnd: span.replacementEnd! }
            : {}) })) }
      : {}),
    ...(request.textRules
      ? { textRules: request.textRules.map((rule): JsonValue => ({ style: rule.style, start: rule.start, end: rule.end,
          removed: Array.from(rule.removed), added: Array.from(rule.added) })) }
      : {}),
    ...(request.boxMark ? { boxMark: { start: request.boxMark.start, end: request.boxMark.end } } : {}),
    ...pdfNativeStructActualTextRecordField(request),
    ...pdfNativePdfaIdentificationRecordField(prepared),
    replacementSha256: prepared.replacementSha256,
    untouchedObjects: (request.untouchedObjects ?? []).map((item) => ({
      objectNumber: item.objectNumber,
      generation: item.generation,
      expectedObjectSha256: item.expectedObjectSha256,
      role: item.role,
    })),
    ...(request.formInstanceIsolation
      ? {
          formInstanceIsolation: {
            page: { ...request.formInstanceIsolation.page },
            chain: request.formInstanceIsolation.chain.map(
              ({ instance, ...item }) => ({
                ...item,
                ...(instance
                  ? {
                      instance: {
                        contentStream: { ...instance.contentStream },
                        operandStart: instance.operandStart,
                        resourceName: instance.resourceName,
                      },
                    }
                  : {}),
              }),
            ),
          },
        }
      : {}),
    ...(request.pageExtGStateMutation
      ? {
          pageExtGStateMutation: {
            pageNumber: request.pageExtGStateMutation.pageNumber,
            page: { ...request.pageExtGStateMutation.page },
            resourceOwner: { ...request.pageExtGStateMutation.resourceOwner },
            category:
              request.pageExtGStateMutation.category.kind === "indirect"
                ? {
                    ...request.pageExtGStateMutation.category,
                    target: {
                      ...request.pageExtGStateMutation.category.target,
                    },
                  }
                : { ...request.pageExtGStateMutation.category },
            resourceName: request.pageExtGStateMutation.resourceName,
            fillOpacity: request.pageExtGStateMutation.fillOpacity,
            strokeOpacity: request.pageExtGStateMutation.strokeOpacity,
            blendMode: request.pageExtGStateMutation.blendMode,
            reuseExisting: request.pageExtGStateMutation.reuseExisting,
          },
        }
      : {}),
    ...(request.pageContentConsolidation
      ? {
          pageContentConsolidation: {
            pageNumber: request.pageContentConsolidation.pageNumber,
            page: { ...request.pageContentConsolidation.page },
            sourceStreams: request.pageContentConsolidation.sourceStreams.map(
              (item) => ({ ...item }),
            ),
            joinedDecodedSha256:
              request.pageContentConsolidation.joinedDecodedSha256,
            stream: { ...request.pageContentConsolidation.stream },
          },
        }
      : {}),
    derivedCandidate: {
      sha256: prepared.candidateSha256,
      byteLength: prepared.candidateBytes.length,
      sourceLength: prepared.sourceLength,
      sourceXrefOffset: prepared.sourceXrefOffset,
      changedObjectOffset: prepared.changedObjectOffset,
      incrementalXrefOffset: prepared.incrementalXrefOffset,
      changedObjectSha256: prepared.changedObjectSha256,
      changedStreamSha256: prepared.changedStreamSha256,
      // A consolidation writes the new stream and the page; the server
      // replays both and compares this receipt object by object.
      ...(prepared.pageContentConsolidation
        ? {
            revisedObjects: [
              prepared.pageContentConsolidation.stream,
              prepared.pageContentConsolidation.page,
              ...(prepared.structActualText ?? []),
              ...pdfNativePdfaIdentificationRevisedObjects(prepared),
            ]
              .sort(
                (left, right) =>
                  left.objectNumber - right.objectNumber ||
                  left.generation - right.generation,
              )
              .map((item) => ({ ...item })),
          }
        : {}),
    },
    ...(listFidelity
      ? {
          // The planner has already authenticated this JSON evidence. Keep
          // its whole schema, as the vector transactions do, so new proof
          // fields survive save and server replay without a second field list.
          textFidelity: JSON.parse(JSON.stringify(listFidelity)) as JsonValue,
        }
      : {}),
    ...(embeddedImageRemoval
      ? {
          embeddedImageRemoval: {
            schemaVersion: embeddedImageRemoval.schemaVersion,
            pageNumber: embeddedImageRemoval.pageNumber,
            imageOrder: embeddedImageRemoval.imageOrder,
            resourceName: embeddedImageRemoval.resourceName,
            page: { ...embeddedImageRemoval.page },
            resourceOwner: { ...embeddedImageRemoval.resourceOwner },
            image: { ...embeddedImageRemoval.image },
            contentStream: { ...embeddedImageRemoval.contentStream },
            operatorSpan: { ...embeddedImageRemoval.operatorSpan },
            drawOperatorBytes: [...embeddedImageRemoval.drawOperatorBytes],
            removedOperatorBytes: [
              ...embeddedImageRemoval.removedOperatorBytes,
            ],
            ...(embeddedImageRemoval.resourceUnmap
              ? {
                  resourceUnmap: {
                    resourceName:
                      embeddedImageRemoval.resourceUnmap.resourceName,
                    direction: embeddedImageRemoval.resourceUnmap.direction,
                    resourceOwner: {
                      ...embeddedImageRemoval.resourceUnmap.resourceOwner,
                    },
                    expectedOwnerSha256:
                      embeddedImageRemoval.resourceUnmap.expectedOwnerSha256,
                    revisedOwnerSha256:
                      embeddedImageRemoval.resourceUnmap.revisedOwnerSha256,
                    entryOffset: embeddedImageRemoval.resourceUnmap.entryOffset,
                    entryBytes: [
                      ...embeddedImageRemoval.resourceUnmap.entryBytes,
                    ],
                  },
                }
              : {}),
          },
        }
      : {}),
    ...(embeddedImageTransform
      ? {
          embeddedImageTransform: {
            schemaVersion: embeddedImageTransform.schemaVersion,
            pageNumber: embeddedImageTransform.pageNumber,
            imageOrder: embeddedImageTransform.imageOrder,
            resourceName: embeddedImageTransform.resourceName,
            page: { ...embeddedImageTransform.page },
            resourceOwner: { ...embeddedImageTransform.resourceOwner },
            image: { ...embeddedImageTransform.image },
            contentStream: { ...embeddedImageTransform.contentStream },
            originalOperatorBytes: [
              ...embeddedImageTransform.originalOperatorBytes,
            ],
            transformedOperatorBytes: [
              ...embeddedImageTransform.transformedOperatorBytes,
            ],
            matrix: [...embeddedImageTransform.matrix],
            clip: embeddedImageTransform.clip
              ? { ...embeddedImageTransform.clip }
              : null,
            ...(embeddedImageTransform.graphicsStateName
              ? { graphicsStateName: embeddedImageTransform.graphicsStateName }
              : {}),
            ...(embeddedImageTransform.opacity === undefined ||
            embeddedImageTransform.opacity === null
              ? {}
              : { opacity: embeddedImageTransform.opacity }),
          },
        }
      : {}),
  };
}

/** Bind a locally compiled text operation only when its exact base has a
 * real durable resource identity. This does not persist or acknowledge it. */
export function bindMraiNativePreparedTextRecord(input: {
  source: PdfEngineSourceIdentity;
  transactionId: string;
  workingRevision: number;
  selectionRevision: number;
  prepared: PdfNativePreparedContentStreamTransaction;
  fidelity: PdfNativeListTextFidelityEvidence;
  sourceCheckpoint?: PdfNativeRecordCheckpoint;
}): PdfSharedStateRecord {
  if (!input.source.resourceId.trim() || !input.source.resourceVersionId.trim() ||
    input.source.sourceSha256 !== input.prepared.request.identity.sourceSha256 ||
    input.fidelity.selection?.sourceSha256 !== input.source.sourceSha256 ||
    !input.transactionId.trim() || !Number.isSafeInteger(input.workingRevision) || input.workingRevision < 0)
    throw new Error("The local PDF operation does not match the durable source.");
  const checkpoint = input.sourceCheckpoint ? recordCheckpointClaim(input.sourceCheckpoint, input.source, input.prepared) : null;
  if (input.sourceCheckpoint && !checkpoint) throw new Error("The local PDF checkpoint does not match its durable base.");
  const record = stableRecord({
    ...input.source, transactionId: input.transactionId, sessionRevision: input.workingRevision,
    selectionRevision: input.selectionRevision, featureId: "PDF-EDT-TEXT-01",
    controlId: "PDFCTL.RIBBON.EDIT.EDIT_CONTENT", action: MRAI_NATIVE_LIST_TEXT_ACTION,
    output: "current-pdf", payload: {},
  }, input.prepared, input.fidelity);
  return checkpoint ? { ...record, sourceCheckpoint: checkpoint } : record;
}

export function stableTextWindowHistoryRecord(
  command: PdfEngineCommand,
  prepared: PdfNativePreparedContentStreamTransaction,
  history: PdfNativeTextWindowHistoryEvidence,
): PdfSharedStateRecord {
  return {...stableRecord(command, prepared), schemaVersion: MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION,
    kind: "text-window-history", textWindowHistory: {...history}};
}

export function bindMraiNativePreparedTextWindowHistoryRecord(input:
  Omit<Parameters<typeof bindMraiNativePreparedTextRecord>[0], "fidelity"> & {history: PdfNativeTextWindowHistoryEvidence},
): PdfSharedStateRecord {
  if (!input.source.resourceId.trim() || !input.source.resourceVersionId.trim() ||
      input.source.sourceSha256 !== input.prepared.request.identity.sourceSha256 ||
      !input.transactionId.trim() || input.transactionId === input.history.originTransactionId || input.selectionRevision !== 0 ||
      !Number.isSafeInteger(input.workingRevision) || input.workingRevision < 0)
    throw new Error("The PDF history operation does not match the durable source.");
  const checkpoint = input.sourceCheckpoint ? recordCheckpointClaim(input.sourceCheckpoint, input.source, input.prepared) : null;
  if (input.sourceCheckpoint && !checkpoint) throw new Error("The PDF history checkpoint does not match its durable base.");
  const record = stableTextWindowHistoryRecord({...input.source,transactionId:input.transactionId,
    sessionRevision:input.workingRevision,selectionRevision:0,featureId:"PDF-EDT-TEXT-01",
    controlId:"PDFCTL.RIBBON.EDIT.EDIT_CONTENT",action:MRAI_NATIVE_LIST_TEXT_ACTION,output:"current-pdf",payload:{}},input.prepared,input.history);
  return checkpoint ? {...record,sourceCheckpoint:checkpoint} : record;
}

