/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  inspectMraiNativeEmbeddedImageReplacement,
  listMraiNativeEmbeddedImagePaintMatrices,
  planMraiNativeStructElementAlt,
  readMraiNativeEmbeddedImageMarkedContent,
  readMraiNativeEmbeddedImageObject,
  listMraiNativeEmbeddedImagePaintsOfObject,
  planMraiNativePictureOrigin,
  readMraiNativePictureOrigin,
  planMraiNativeEmbeddedImageRemoval,
  planMraiNativeEmbeddedImageRemovalHistory,
  planMraiNativeEmbeddedImageReplacement,
  planMraiNativeEmbeddedImageReplacementHistory,
  planMraiNativeEmbeddedImageTransform,
  planMraiNativeEmbeddedImageTransformHistory,
  planMraiNativePageSplit,
  planMraiNativeObjectRestore,
  planMraiNativeObjectRestoreFromEvidence,
  prepareMraiNativeContentStreamTransaction,
  readMraiNativeManagedOutline,
  readMraiNativePdfXrefSectionCount,
  squashMraiNativeRevisions,
  verifyMraiNativeContentStreamCandidate,
} from "./engine/native/content-stream-transaction";
import {
  planMraiNativeListTextHistoryReplacement,
  planMraiNativeListTextPropertyReplacement,
  planMraiNativeListTextSelectionReplacement,
  proveMraiNativeListTextSelection,
  proveMraiNativeListTextSelectionForWriter,
  warmMraiNativeParagraphFormatPlans,
} from "./engine/native/list-text-fidelity";
import { planMraiNativeByteRestoreHistory } from "./engine/native/byte-restore-history";
import { planMraiNativeTextMove } from "./engine/native/text-move";
import { planMraiNativeTextGroupMove } from "./engine/native/text-group-move";
import { findMraiNativeTextDecoration, readMraiNativeTextRuleBesideRun } from "./engine/native/text-decoration";
import {
  planMraiNativeInvisibleTextRuleChanges,
  proveMraiNativeInvisibleTextRules,
  readMraiNativeInvisibleTextRulesStanding,
} from "./engine/native/invisible-text-rule";
import { planPdfNativeContentAppend } from "./engine/native/content-append-transaction";
import { fetchVerifiedPdfComplexScriptFont, fetchVerifiedPdfUnicodeFallbackFont } from "./unicode-font-asset";
import {
  planMraiNativeParagraphProperties,
  prepareMraiNativeParagraphProperties,
  rebindMraiNativeParagraphProperties,
  verifyMraiNativeParagraphProperties,
} from "./engine/native/paragraph-properties-transaction";
import {
  planPdfNativeVectorBackdropDelete,
  planPdfNativeVectorBackdropFills,
  planPdfNativeVectorBackdropSelectionDelete,
  planPdfNativeVectorDisplayAffine,
  planPdfNativeVectorDisplayDelete,
  planPdfNativeVectorDisplayDuplicate,
  planPdfNativeVectorDisplayEdit,
  planPdfNativeVectorDisplayFlip,
  planPdfNativeVectorDisplayGroupTransform,
  planPdfNativeVectorDisplayReorder,
  planPdfNativeVectorDisplayRotation,
  planPdfNativeVectorDisplayVertexEdit,
  planPdfNativeVectorSelectionDelete,
  planPdfNativeVectorSelectionDuplicate,
  planPdfNativeVectorSelectionEdit,
  planPdfNativeVectorSelectionGroupTransform,
  planPdfNativeVectorSelectionReorder,
  resolvePdfNativeVectorHistorySelection,
  readPdfNativeVectorPageShapes,
  resolvePdfNativeVectorRouteSelections,
} from "./native-vector-selection-edit";
import {
  planPdfNativeVectorStyleSelectionEdit,
  readPdfNativeVectorPageGeometry,
  resolvePdfNativeVectorStyleHistoryRoute,
} from "./native-vector-style-selection-edit";
import { planPdfNativeTextDecoration } from "./native-text-decoration";
import { planPdfNativeTextObjectDelete } from "./native-text-object-delete";
import { planMraiNativeEmbeddedImageClipEscape, planPdfNativeEmbeddedImageReorder } from "./native-embedded-image-reorder";
import { listMraiNativePageContentUnits, planMraiNativeContentUnitReorder } from "./engine/native/vector-content-transaction";
import {
  planPdfNativeEmbeddedFontPropertyEditFromPrograms,
  planPdfNativeFontExtensionTextEdit,
} from "./native-font-extension-text-edit";
import { preparePdfNativeSourceCheckpoint } from "./engine/native/source-checkpoint";
import { bindPdfSourceTextBox, verifyPdfSourceTextBox } from "./source-text-box-proof";
import { locatePdfFallbackText } from "./source-text-fallback";
import { inspectPdfNativeFontReadingRepair, planPdfNativeFontReadingRepair } from "./native-font-reading-repair";
import { inspectPdfSemanticAppearanceSelectionSource } from "./semantic-appearance-native-evidence";

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- a table of heterogeneous entry points
type Entry = (...args: any[]) => unknown;

/** The planners this Worker answers, by the name the editor calls them by. */
export const PDF_NATIVE_BROWSER_PLANNERS: Readonly<Record<string, Entry>> = Object.freeze({
  inspectMraiNativeEmbeddedImageReplacement,
  listMraiNativeEmbeddedImagePaintMatrices,
  planMraiNativeStructElementAlt,
  readMraiNativeEmbeddedImageMarkedContent,
  readMraiNativeEmbeddedImageObject,
  listMraiNativeEmbeddedImagePaintsOfObject,
  planMraiNativePictureOrigin,
  readMraiNativePictureOrigin,
  planMraiNativeEmbeddedImageRemoval,
  planMraiNativeEmbeddedImageRemovalHistory,
  planMraiNativeEmbeddedImageReplacement,
  planMraiNativeEmbeddedImageReplacementHistory,
  planMraiNativeEmbeddedImageTransform,
  planMraiNativeEmbeddedImageTransformHistory,
  planMraiNativePageSplit,
  planMraiNativeObjectRestore,
  planMraiNativeObjectRestoreFromEvidence,
  prepareMraiNativeContentStreamTransaction,
  readMraiNativeManagedOutline,
  readMraiNativePdfXrefSectionCount,
  squashMraiNativeRevisions,
  verifyMraiNativeContentStreamCandidate,
  planMraiNativeListTextHistoryReplacement,
  planMraiNativeListTextPropertyReplacement,
  planMraiNativeListTextSelectionReplacement,
  proveMraiNativeListTextSelection,
  proveMraiNativeListTextSelectionForWriter,
  planMraiNativeByteRestoreHistory,
  planMraiNativeTextMove,
  planMraiNativeTextGroupMove,
  // A head read ahead (its document inspection) for the box plans to come.
  warmMraiNativeParagraphFormatPlans,
  findMraiNativeTextDecoration,
  readMraiNativeTextRuleBesideRun,
  // A scanned line's rules-only proof: where a rule beside its invisible
  // words goes and whether it would be seen.
  proveMraiNativeInvisibleTextRules,
  // The rules standing beside it now, and the rule changes a format asks.
  readMraiNativeInvisibleTextRulesStanding,
  planMraiNativeInvisibleTextRuleChanges,
  // The faces it draws in, fetched here by their pinned digests (a loader
  // the editor holds cannot be posted to this Worker).
  planPdfNativeContentAppend: (input: Parameters<typeof planPdfNativeContentAppend>[0]) => planPdfNativeContentAppend({
    ...input,
    loadFont: input.loadFont ?? (async () => {
      const loaded = await fetchVerifiedPdfUnicodeFallbackFont();
      if (!loaded) throw new Error("The bundled Unicode font is unavailable.");
      return new Uint8Array(loaded);
    }),
    loadComplexFont: input.loadComplexFont ?? (async (sourceId: string) => {
      const loaded = await fetchVerifiedPdfComplexScriptFont(sourceId);
      return loaded ? new Uint8Array(loaded) : null;
    }),
  }),
  planMraiNativeParagraphProperties,
  prepareMraiNativeParagraphProperties,
  rebindMraiNativeParagraphProperties,
  verifyMraiNativeParagraphProperties,
  planPdfNativeVectorDisplayAffine,
  planPdfNativeVectorDisplayDelete,
  planPdfNativeVectorDisplayDuplicate,
  planPdfNativeVectorDisplayEdit,
  planPdfNativeVectorDisplayFlip,
  planPdfNativeVectorDisplayGroupTransform,
  planPdfNativeVectorDisplayReorder,
  planPdfNativeVectorDisplayRotation,
  planPdfNativeVectorDisplayVertexEdit,
  planPdfNativeVectorSelectionDelete,
  planPdfNativeVectorSelectionDuplicate,
  planPdfNativeVectorSelectionEdit,
  planPdfNativeVectorSelectionGroupTransform,
  planPdfNativeVectorSelectionReorder,
  resolvePdfNativeVectorHistorySelection,
  resolvePdfNativeVectorRouteSelections,
  readPdfNativeVectorPageShapes,
  planPdfNativeVectorStyleSelectionEdit,
  readPdfNativeVectorPageGeometry,
  resolvePdfNativeVectorStyleHistoryRoute,
  planPdfNativeTextDecoration,
  planPdfNativeTextObjectDelete,
  planPdfNativeEmbeddedImageReorder,
  planMraiNativeEmbeddedImageClipEscape,
  listMraiNativePageContentUnits,
  planMraiNativeContentUnitReorder,
  planPdfNativeFontExtensionTextEdit,
  // ...with the served font list the editor hands over (its manifest field).
  planPdfNativeEmbeddedFontPropertyEditFromPrograms,
  preparePdfNativeSourceCheckpoint,
  planPdfNativeVectorBackdropDelete,
  planPdfNativeVectorBackdropFills,
  planPdfNativeVectorBackdropSelectionDelete,
  bindPdfSourceTextBox,
  verifyPdfSourceTextBox,
  locatePdfFallbackText,
  planPdfNativeFontReadingRepair,
  inspectPdfNativeFontReadingRepair,
  inspectPdfSemanticAppearanceSelectionSource,
});

type ThrownError = Readonly<{ name: string; message: string; code?: string; resourceIssue?: unknown }>;

export interface PdfNativePlannerWorkerRequest {
  readonly requestId: string;
  readonly fn: string;
  readonly args: readonly unknown[];
}

export type PdfNativePlannerWorkerResponse =
  | Readonly<{ requestId: string; status: "answered"; value: unknown }>
  | Readonly<{ requestId: string; status: "thrown"; error: ThrownError }>
  | Readonly<{ requestId: string; status: "failed"; reason: string }>;

function thrown(error: unknown): ThrownError {
  const record = (error && typeof error === "object" ? error : {}) as Record<string, unknown>;
  const issue = record.resourceIssue;
  return Object.freeze({
    name: error instanceof Error ? error.name : "Error",
    message: error instanceof Error ? error.message : String(error),
    ...(typeof record.code === "string" ? { code: record.code } : {}),
    ...(issue !== undefined ? { resourceIssue: JSON.parse(JSON.stringify(issue)) as unknown } : {}),
  });
}

export async function answerPdfNativePlannerRequest(request: PdfNativePlannerWorkerRequest): Promise<PdfNativePlannerWorkerResponse> {
  const planner = PDF_NATIVE_BROWSER_PLANNERS[request.fn];
  if (!planner) return { requestId: request.requestId, status: "failed", reason: `unknown-planner:${request.fn}` };
  let value: unknown;
  try {
    value = await planner(...request.args);
  } catch (error) {
    return { requestId: request.requestId, status: "thrown", error: thrown(error) };
  }
  return { requestId: request.requestId, status: "answered", value };
}

/** Whether a message is a planner request this module answers (the PDF
 * text Worker, native-text.worker.ts, hands it over). */
export function isPdfNativePlannerWorkerRequest(request: unknown): request is PdfNativePlannerWorkerRequest {
  const value = request as Partial<PdfNativePlannerWorkerRequest> | null;
  return Boolean(value && typeof value.requestId === "string" && typeof value.fn === "string" && Array.isArray(value.args));
}
