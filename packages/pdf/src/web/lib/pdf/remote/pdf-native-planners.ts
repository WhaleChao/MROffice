/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type * as ContentStream from "../engine/native/content-stream-transaction";
import type * as ListText from "../engine/native/list-text-fidelity";
import type * as ByteRestore from "../engine/native/byte-restore-history";
import type * as TextMove from "../engine/native/text-move";
import type * as TextGroupMove from "../engine/native/text-group-move";
import type * as TextDecorationRead from "../engine/native/text-decoration";
import type * as InvisibleTextRule from "../engine/native/invisible-text-rule";
import type * as ContentAppend from "../engine/native/content-append-transaction";
import type * as ParagraphProperties from "../engine/native/paragraph-properties-transaction";
import type * as SourceCheckpoint from "../engine/native/source-checkpoint";
import type * as VectorEdit from "../native-vector-selection-edit";
import type * as VectorStyle from "../native-vector-style-selection-edit";
import type * as TextDecoration from "../native-text-decoration";
import type * as TextObjectDelete from "../native-text-object-delete";
import type * as ImageReorder from "../native-embedded-image-reorder";
import type * as ContentUnits from "../engine/native/vector-content-transaction";
import type * as FontExtension from "../native-font-extension-text-edit";
import type * as AuthorRecovery from "../native-author-recovery";
import type * as SourceTextBoxProof from "../source-text-box-proof";
import type * as SourceTextFallback from "../source-text-fallback";
import type * as FontReadingRepair from "../native-font-reading-repair";
import type * as TaggedText from "../native-tagged-text-roundtrip";
import type * as SemanticNative from "../semantic-appearance-native-evidence";
import { decodePdfNativeWire, encodePdfNativeWire } from "@mrnote/core/web/lib/engine-host/pdf-native-wire.ts";
import { PdfNativeTransactionError, sha256PdfBytes, type PdfNativeTransactionErrorCode } from "../engine/native/transaction-error";
import { readPdfNativeResourceIssue } from "../engine/native/resource-issue";
import { callPdfNativeEngine, createPdfNativeOffload, readablePdfEngineError } from "../pdf-native-remote";
import { callEngine } from "@mrnote/core/web/lib/engine-host/client.ts";
import { pdfNativeWorkingHead, pdfNativeWorkingHeadPages } from "./pdf-native-sources";
import {
  callPdfNativePlannerInBrowser,
  pdfNativeBrowserPlannerEnabled,
  PdfNativeBrowserPlannerThrew,
  PdfNativeBrowserPlannerUnavailable,
} from "../native-planners-browser";
import { markPdfTextLatency } from "../text-edit-latency";

// eslint-disable-next-line @typescript-eslint/no-explicit-any -- any planner signature
type Planner = (...args: any[]) => unknown;
type Remote<F extends Planner> = (...args: Parameters<F>) => Promise<Awaited<ReturnType<F>>>;

const TRANSACTION_CODES = new Set<PdfNativeTransactionErrorCode>([
  "invalid-request", "source-mismatch", "precondition-mismatch", "unsupported-document", "verification-failed",
]);
/** Shorter byte arrays are never a PDF source. */
const MIN_SOURCE_BYTES = 64;

const digests = new WeakMap<Uint8Array, Promise<string>>();






export function notePdfNativeBytesDigest(bytes: Uint8Array, sha256: string): void {
  if (/^[0-9a-f]{64}$/i.test(sha256)) digests.set(bytes, Promise.resolve(sha256.toLowerCase()));
}
function digestOf(bytes: Uint8Array): Promise<string> {
  let digest = digests.get(bytes);
  if (!digest) {
    digest = sha256PdfBytes(bytes);
    digests.set(bytes, digest);
  }
  return digest;
}

function byteArguments(value: unknown, out: Uint8Array[], depth = 0): void {
  if (depth > 4 || !value || typeof value !== "object") return;
  if (ArrayBuffer.isView(value)) {
    if (value.constructor?.name === "Uint8Array" && value.byteLength >= MIN_SOURCE_BYTES) {
      out.push(value instanceof Uint8Array ? value : new Uint8Array(value.buffer, value.byteOffset, value.byteLength));
    }
    return;
  }
  if (Array.isArray(value)) {
    for (const item of value) byteArguments(item, out, depth + 1);
    return;
  }
  if (value instanceof Map || value instanceof Set) return;
  for (const item of Object.values(value)) byteArguments(item, out, depth + 1);
}

/** The working head among a call's arguments: the first byte array whose
 * digest a loaded head registered -- under `pageId`, the page the caller's
 * planner set is bound to. A call bound to no page (`null`) is answered
 * only for bytes a single page holds: bytes two pages registered (twin
 * copies of one PDF open in one tab) are refused, the call names neither. */
export async function workingHead(args: readonly unknown[], pageId: string | null) {
  const candidates: Uint8Array[] = [];
  for (const arg of args) byteArguments(arg, candidates);
  for (const bytes of candidates) {
    const sha256 = await digestOf(bytes);
    const ref = pdfNativeWorkingHead(sha256, pageId);
    if (ref) return { sha256, bytes, ref };
    if (!pageId && pdfNativeWorkingHeadPages(sha256).length > 1) {
      throw new PdfNativeTransactionError("source-mismatch",
        "This PDF is open on more than one page here; the call does not say which.");
    }
  }
  return null;
}

export function rethrow(error: { name: string; message: string; code?: string; resourceIssue?: unknown }): never {
  if (error.code && TRANSACTION_CODES.has(error.code as PdfNativeTransactionErrorCode)) {
    throw new PdfNativeTransactionError(error.code as PdfNativeTransactionErrorCode, error.message,
      readPdfNativeResourceIssue(error.resourceIssue) ?? undefined);
  }
  const rebuilt = new Error(error.message);
  rebuilt.name = error.name;
  if (error.code) Object.assign(rebuilt, { code: error.code });
  throw rebuilt;
}

export async function callPdfNativePlanner(
  /** The page whose working head the call plans on (its planner set's). */
  pageId: string | null,
  fn: string,
  args: readonly unknown[],
  signal?: AbortSignal,
  /** What this browser's Worker is given instead of `args`: the editor's own
   * bytes a host fills in for itself (the bundled font). */
  browserArgs?: readonly unknown[],
): Promise<unknown> {
  const head = await workingHead(args, pageId);
  if (!head) {
    throw new PdfNativeTransactionError("source-mismatch", "The PDF source is not open for editing here.");
  }
  // This browser's own Worker first (browser first, phase 5): the same
  // planner on the same bytes, no server round trip. The engine host
  // answers only what the Worker cannot (no Worker, a crash, a timeout,
  // a planner it does not hold); a planner's own answer is final.
  if (pdfNativeBrowserPlannerEnabled(fn)) {
    const startedAt = performance.now();
    try {
      const value = await callPdfNativePlannerInBrowser(fn, browserArgs ?? args, signal);
      markPdfTextLatency("browser.planner", { fn, ms: Math.round(performance.now() - startedAt) });
      return value;
    } catch (error) {
      if (error instanceof PdfNativeBrowserPlannerThrew) rethrow(error.thrown);
      if (!(error instanceof PdfNativeBrowserPlannerUnavailable)) throw error;
      markPdfTextLatency("browser.planner:fallback", { fn, reason: error.message.slice(0, 120),
        ms: Math.round(performance.now() - startedAt) });
    }
  }
  const source = { sha256: head.sha256, bytes: head.bytes };
  // Large byte arrays (a candidate carrying an installed font, a picture)
  // travel beside the JSON, not in it.
  const offloaded = createPdfNativeOffload();
  const answer = await callPdfNativeEngine("call", {
    source: { ...head.ref, sourceSha256: head.sha256 },
    fn,
    args: (await encodePdfNativeWire([...args], source, undefined, offloaded.offload)) as unknown[],
  }, { signal, workingBytes: () => head.bytes, offloaded });
  if (!answer.ok) rethrow(answer.error);
  return decodePdfNativeWire(answer.value, (base) => (base === head.sha256 ? head.bytes : null));
}










async function holdPdfNativeHeadOnServer(pageId: string | null, bytes: Uint8Array): Promise<void> {
  const head = await workingHead([bytes], pageId);
  if (!head) {
    throw new PdfNativeTransactionError("source-mismatch", "The PDF source is not open for editing here.");
  }
  const source = { sha256: head.sha256, bytes: head.bytes };
  const offloaded = createPdfNativeOffload();
  const answer = await callPdfNativeEngine("call", {
    source: { ...head.ref, sourceSha256: head.sha256 },
    fn: "readMraiNativePdfXrefSectionCount",
    args: (await encodePdfNativeWire([bytes], source, undefined, offloaded.offload)) as unknown[],
  }, { workingBytes: () => head.bytes, offloaded });
  if (!answer.ok) rethrow(answer.error);
}

/** A picture's inspection reads no PDF: `pdf-native/pure`. */
const inspectMraiNativeEmbeddedImageReplacement: Remote<typeof ContentStream.inspectMraiNativeEmbeddedImageReplacement> =
  async (...args) => {
    // A picture of 64 KiB or more is this call's upload, not part of its JSON.
    const offloaded = createPdfNativeOffload();
    const encoded = (await encodePdfNativeWire([...args], null, undefined, offloaded.offload)) as unknown[];
    const uploads = [...offloaded.inputs].map(([sha256, bytes]) => ({ sha256, bytes, role: "input" as const }));
    const answer = await callEngine("pdf-native", "pure", {
      fn: "inspectMraiNativeEmbeddedImageReplacement",
      args: encoded,
    }, uploads.length ? { uploads } : {}).catch((error: unknown) => { throw readablePdfEngineError(error); });
    if (!answer.ok) rethrow(answer.error);
    return decodePdfNativeWire(answer.value, () => null);
  };
function remote<F extends Planner>(pageId: string | null, fn: string): Remote<F> {
  return ((...args: Parameters<F>) => callPdfNativePlanner(pageId, fn, args)) as Remote<F>;
}

/** The editor's font loaders stay here: this browser's Worker fetches the
 * faces itself (native-planners.worker.ts), the engine host reads its own. */
const contentAppend = (pageId: string | null): Remote<typeof ContentAppend.planPdfNativeContentAppend> => (input) => {
  const { loadFont: _loadFont, loadComplexFont: _loadComplexFont, ...posted } = input;
  void _loadFont;
  void _loadComplexFont;
  return callPdfNativePlanner(pageId, "planPdfNativeContentAppend", [input], undefined, [posted]) as
    ReturnType<Remote<typeof ContentAppend.planPdfNativeContentAppend>>;
};





const fontExtensionTextEdit = (pageId: string | null): Remote<typeof FontExtension.planPdfNativeFontExtensionTextEdit> =>
  (input) => {
    // `undefined` does not travel: the engine fills in its font.
    return callPdfNativePlanner(pageId, "planPdfNativeFontExtensionTextEdit", [{ ...input, fontBytes: undefined }],
      undefined,
      [{ ...input, fontBytes: input.fontBytes instanceof Uint8Array ? input.fontBytes : new Uint8Array(input.fontBytes) }],
    ) as ReturnType<Remote<typeof FontExtension.planPdfNativeFontExtensionTextEdit>>;
  };

/** Every planner, each call finding its head under `pageId`. */
function bindPdfNativePlanners(pageId: string | null) {
  return {
    /** The page this set plans for (`null`: none, see `workingHead`). */
    pageId,
    /** See `holdPdfNativeHeadOnServer`. */
    holdPdfNativeHeadOnServer: (bytes: Uint8Array) => holdPdfNativeHeadOnServer(pageId, bytes),
    inspectMraiNativeEmbeddedImageReplacement,
    listMraiNativeEmbeddedImagePaintMatrices:
      remote<typeof ContentStream.listMraiNativeEmbeddedImagePaintMatrices>(pageId, "listMraiNativeEmbeddedImagePaintMatrices"),
    readMraiNativeEmbeddedImageMarkedContent:
      remote<typeof ContentStream.readMraiNativeEmbeddedImageMarkedContent>(pageId, "readMraiNativeEmbeddedImageMarkedContent"),
    readMraiNativeEmbeddedImageObject:
      remote<typeof ContentStream.readMraiNativeEmbeddedImageObject>(pageId, "readMraiNativeEmbeddedImageObject"),
    listMraiNativeEmbeddedImagePaintsOfObject:
      remote<typeof ContentStream.listMraiNativeEmbeddedImagePaintsOfObject>(pageId, "listMraiNativeEmbeddedImagePaintsOfObject"),
    planMraiNativePictureOrigin:
      remote<typeof ContentStream.planMraiNativePictureOrigin>(pageId, "planMraiNativePictureOrigin"),
    readMraiNativePictureOrigin:
      remote<typeof ContentStream.readMraiNativePictureOrigin>(pageId, "readMraiNativePictureOrigin"),
    planMraiNativeStructElementAlt:
      remote<typeof ContentStream.planMraiNativeStructElementAlt>(pageId, "planMraiNativeStructElementAlt"),
    planMraiNativeEmbeddedImageRemoval:
      remote<typeof ContentStream.planMraiNativeEmbeddedImageRemoval>(pageId, "planMraiNativeEmbeddedImageRemoval"),
    planMraiNativeEmbeddedImageRemovalHistory:
      remote<typeof ContentStream.planMraiNativeEmbeddedImageRemovalHistory>(pageId, "planMraiNativeEmbeddedImageRemovalHistory"),
    planMraiNativeEmbeddedImageReplacement:
      remote<typeof ContentStream.planMraiNativeEmbeddedImageReplacement>(pageId, "planMraiNativeEmbeddedImageReplacement"),
    planMraiNativeEmbeddedImageReplacementHistory:
      remote<typeof ContentStream.planMraiNativeEmbeddedImageReplacementHistory>(pageId, "planMraiNativeEmbeddedImageReplacementHistory"),
    planMraiNativeEmbeddedImageTransform:
      remote<typeof ContentStream.planMraiNativeEmbeddedImageTransform>(pageId, "planMraiNativeEmbeddedImageTransform"),
    planMraiNativeEmbeddedImageTransformHistory:
      remote<typeof ContentStream.planMraiNativeEmbeddedImageTransformHistory>(pageId, "planMraiNativeEmbeddedImageTransformHistory"),
    planMraiNativePageSplit: remote<typeof ContentStream.planMraiNativePageSplit>(pageId, "planMraiNativePageSplit"),
    planMraiNativePageBackground:
      remote<typeof ContentStream.planMraiNativePageBackground>(pageId, "planMraiNativePageBackground"),
    prepareMraiNativeContentStreamTransaction:
      remote<typeof ContentStream.prepareMraiNativeContentStreamTransaction>(pageId, "prepareMraiNativeContentStreamTransaction"),
    readMraiNativePdfXrefSectionCount:
      remote<typeof ContentStream.readMraiNativePdfXrefSectionCount>(pageId, "readMraiNativePdfXrefSectionCount"),
    squashMraiNativeRevisions:
      remote<typeof ContentStream.squashMraiNativeRevisions>(pageId, "squashMraiNativeRevisions"),
    readMraiNativeManagedOutline:
      remote<typeof ContentStream.readMraiNativeManagedOutline>(pageId, "readMraiNativeManagedOutline"),
    planMraiNativeObjectRestore:
      remote<typeof ContentStream.planMraiNativeObjectRestore>(pageId, "planMraiNativeObjectRestore"),
    planMraiNativeObjectRestoreFromEvidence:
      remote<typeof ContentStream.planMraiNativeObjectRestoreFromEvidence>(pageId, "planMraiNativeObjectRestoreFromEvidence"),
    verifyMraiNativeContentStreamCandidate:
      remote<typeof ContentStream.verifyMraiNativeContentStreamCandidate>(pageId, "verifyMraiNativeContentStreamCandidate"),
    planMraiNativeListTextHistoryReplacement:
      remote<typeof ListText.planMraiNativeListTextHistoryReplacement>(pageId, "planMraiNativeListTextHistoryReplacement"),
    planMraiNativeListTextPropertyReplacement:
      remote<typeof ListText.planMraiNativeListTextPropertyReplacement>(pageId, "planMraiNativeListTextPropertyReplacement"),
    planMraiNativeListTextSelectionReplacement:
      remote<typeof ListText.planMraiNativeListTextSelectionReplacement>(pageId, "planMraiNativeListTextSelectionReplacement"),
    proveMraiNativeListTextSelection:
      remote<typeof ListText.proveMraiNativeListTextSelection>(pageId, "proveMraiNativeListTextSelection"),
    proveMraiNativeListTextSelectionForWriter:
      remote<typeof ListText.proveMraiNativeListTextSelectionForWriter>(pageId, "proveMraiNativeListTextSelectionForWriter"),
    planMraiNativeByteRestoreHistory:
      remote<typeof ByteRestore.planMraiNativeByteRestoreHistory>(pageId, "planMraiNativeByteRestoreHistory"),
    planMraiNativeTextMove: remote<typeof TextMove.planMraiNativeTextMove>(pageId, "planMraiNativeTextMove"),
    planMraiNativeTextGroupMove:
      remote<typeof TextGroupMove.planMraiNativeTextGroupMove>(pageId, "planMraiNativeTextGroupMove"),
    findMraiNativeTextDecoration:
      remote<typeof TextDecorationRead.findMraiNativeTextDecoration>(pageId, "findMraiNativeTextDecoration"),
    readMraiNativeTextRuleBesideRun:
      remote<typeof TextDecorationRead.readMraiNativeTextRuleBesideRun>(pageId, "readMraiNativeTextRuleBesideRun"),
    proveMraiNativeInvisibleTextRules:
      remote<typeof InvisibleTextRule.proveMraiNativeInvisibleTextRules>(pageId, "proveMraiNativeInvisibleTextRules"),
    readMraiNativeInvisibleTextRulesStanding:
      remote<typeof InvisibleTextRule.readMraiNativeInvisibleTextRulesStanding>(pageId, "readMraiNativeInvisibleTextRulesStanding"),
    planMraiNativeInvisibleTextRuleChanges:
      remote<typeof InvisibleTextRule.planMraiNativeInvisibleTextRuleChanges>(pageId, "planMraiNativeInvisibleTextRuleChanges"),
    planPdfNativeContentAppend: contentAppend(pageId),
    planMraiNativeParagraphProperties:
      remote<typeof ParagraphProperties.planMraiNativeParagraphProperties>(pageId, "planMraiNativeParagraphProperties"),
    prepareMraiNativeParagraphProperties:
      remote<typeof ParagraphProperties.prepareMraiNativeParagraphProperties>(pageId, "prepareMraiNativeParagraphProperties"),
    rebindMraiNativeParagraphProperties:
      remote<typeof ParagraphProperties.rebindMraiNativeParagraphProperties>(pageId, "rebindMraiNativeParagraphProperties"),
    verifyMraiNativeParagraphProperties:
      remote<typeof ParagraphProperties.verifyMraiNativeParagraphProperties>(pageId, "verifyMraiNativeParagraphProperties"),
    planPdfNativeVectorBackdropDelete:
      remote<typeof VectorEdit.planPdfNativeVectorBackdropDelete>(pageId, "planPdfNativeVectorBackdropDelete"),
    planPdfNativeVectorBackdropFills:
      remote<typeof VectorEdit.planPdfNativeVectorBackdropFills>(pageId, "planPdfNativeVectorBackdropFills"),
    planPdfNativeVectorBackdropSelectionDelete:
      remote<typeof VectorEdit.planPdfNativeVectorBackdropSelectionDelete>(pageId, "planPdfNativeVectorBackdropSelectionDelete"),
    planPdfNativeVectorDisplayDelete:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayDelete>(pageId, "planPdfNativeVectorDisplayDelete"),
    planPdfNativeVectorDisplayDuplicate:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayDuplicate>(pageId, "planPdfNativeVectorDisplayDuplicate"),
    planPdfNativeVectorDisplayEdit:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayEdit>(pageId, "planPdfNativeVectorDisplayEdit"),
    planPdfNativeVectorDisplayAffine:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayAffine>(pageId, "planPdfNativeVectorDisplayAffine"),
    planPdfNativeVectorDisplayFlip:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayFlip>(pageId, "planPdfNativeVectorDisplayFlip"),
    planPdfNativeVectorDisplayGroupTransform:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayGroupTransform>(pageId, "planPdfNativeVectorDisplayGroupTransform"),
    planPdfNativeVectorDisplayReorder:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayReorder>(pageId, "planPdfNativeVectorDisplayReorder"),
    planPdfNativeVectorDisplayRotation:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayRotation>(pageId, "planPdfNativeVectorDisplayRotation"),
    planPdfNativeVectorDisplayVertexEdit:
      remote<typeof VectorEdit.planPdfNativeVectorDisplayVertexEdit>(pageId, "planPdfNativeVectorDisplayVertexEdit"),
    planPdfNativeVectorSelectionDelete:
      remote<typeof VectorEdit.planPdfNativeVectorSelectionDelete>(pageId, "planPdfNativeVectorSelectionDelete"),
    planPdfNativeVectorSelectionDuplicate:
      remote<typeof VectorEdit.planPdfNativeVectorSelectionDuplicate>(pageId, "planPdfNativeVectorSelectionDuplicate"),
    planPdfNativeVectorSelectionEdit:
      remote<typeof VectorEdit.planPdfNativeVectorSelectionEdit>(pageId, "planPdfNativeVectorSelectionEdit"),
    planPdfNativeVectorSelectionGroupTransform:
      remote<typeof VectorEdit.planPdfNativeVectorSelectionGroupTransform>(pageId, "planPdfNativeVectorSelectionGroupTransform"),
    planPdfNativeVectorSelectionReorder:
      remote<typeof VectorEdit.planPdfNativeVectorSelectionReorder>(pageId, "planPdfNativeVectorSelectionReorder"),
    resolvePdfNativeVectorHistorySelection:
      remote<typeof VectorEdit.resolvePdfNativeVectorHistorySelection>(pageId, "resolvePdfNativeVectorHistorySelection"),
    resolvePdfNativeVectorRouteSelections:
      remote<typeof VectorEdit.resolvePdfNativeVectorRouteSelections>(pageId, "resolvePdfNativeVectorRouteSelections"),
    readPdfNativeVectorPageShapes:
      remote<typeof VectorEdit.readPdfNativeVectorPageShapes>(pageId, "readPdfNativeVectorPageShapes"),
    planPdfNativeVectorStyleSelectionEdit:
      remote<typeof VectorStyle.planPdfNativeVectorStyleSelectionEdit>(pageId, "planPdfNativeVectorStyleSelectionEdit"),
    readPdfNativeVectorPageGeometry:
      remote<typeof VectorStyle.readPdfNativeVectorPageGeometry>(pageId, "readPdfNativeVectorPageGeometry"),
    resolvePdfNativeVectorStyleHistoryRoute:
      remote<typeof VectorStyle.resolvePdfNativeVectorStyleHistoryRoute>(pageId, "resolvePdfNativeVectorStyleHistoryRoute"),
    planPdfNativeTextDecoration:
      remote<typeof TextDecoration.planPdfNativeTextDecoration>(pageId, "planPdfNativeTextDecoration"),
    planPdfNativeTextObjectDelete:
      remote<typeof TextObjectDelete.planPdfNativeTextObjectDelete>(pageId, "planPdfNativeTextObjectDelete"),
    planPdfNativeEmbeddedImageReorder:
      remote<typeof ImageReorder.planPdfNativeEmbeddedImageReorder>(pageId, "planPdfNativeEmbeddedImageReorder"),
    planMraiNativeEmbeddedImageClipEscape:
      remote<typeof ImageReorder.planMraiNativeEmbeddedImageClipEscape>(pageId, "planMraiNativeEmbeddedImageClipEscape"),
    /** The page content's objects as a layer step passes them, and the layer
     * move of a text object, shading, Form or group (the Selection pane). */
    listMraiNativePageContentUnits:
      remote<typeof ContentUnits.listMraiNativePageContentUnits>(pageId, "listMraiNativePageContentUnits"),
    planMraiNativeContentUnitReorder:
      remote<typeof ContentUnits.planMraiNativeContentUnitReorder>(pageId, "planMraiNativeContentUnitReorder"),
    planPdfNativeFontExtensionTextEdit: fontExtensionTextEdit(pageId),
    /** A part set in a family the PDF does not hold: the slices the editor
     * fetched and verified travel with the call (the engine fetches nothing). */
    planPdfNativeEmbeddedFontPropertyEdit:
      remote<typeof FontExtension.planPdfNativeEmbeddedFontPropertyEditFromPrograms>(pageId, "planPdfNativeEmbeddedFontPropertyEditFromPrograms"),
    preparePdfNativeSourceCheckpoint:
      remote<typeof SourceCheckpoint.preparePdfNativeSourceCheckpoint>(pageId, "preparePdfNativeSourceCheckpoint"),
    preparePdfNativeAuthorTextRecovery:
      remote<typeof AuthorRecovery.preparePdfNativeAuthorTextRecovery>(pageId, "preparePdfNativeAuthorTextRecovery"),
    bindPdfSourceTextBox: remote<typeof SourceTextBoxProof.bindPdfSourceTextBox>(pageId, "bindPdfSourceTextBox"),
    verifyPdfSourceTextBox:
      remote<typeof SourceTextBoxProof.verifyPdfSourceTextBox>(pageId, "verifyPdfSourceTextBox"),
     
    locatePdfFallbackText: remote<typeof SourceTextFallback.locatePdfFallbackText>(pageId, "locatePdfFallbackText"),
    /** The page's fonts whose ToUnicode calls drawn glyphs spaces, named by
     * their own programs: one revision appended, or null (native-font-reading-repair). */
    planPdfNativeFontReadingRepair:
      remote<typeof FontReadingRepair.planPdfNativeFontReadingRepair>(pageId, "planPdfNativeFontReadingRepair"),
    /** The pages that draw with such a font: asked once for a document. */
    inspectPdfNativeFontReadingRepair:
      remote<typeof FontReadingRepair.inspectPdfNativeFontReadingRepair>(pageId, "inspectPdfNativeFontReadingRepair"),
    inspectPdfTaggedTextTargets:
      remote<typeof TaggedText.inspectPdfTaggedTextTargets>(pageId, "inspectPdfTaggedTextTargets"),
    createPdfTaggedTextOperation:
      remote<typeof TaggedText.createPdfTaggedTextOperation>(pageId, "createPdfTaggedTextOperation"),
    verifyPdfTaggedTextOutput:
      remote<typeof TaggedText.verifyPdfTaggedTextOutput>(pageId, "verifyPdfTaggedTextOutput"),
    inspectPdfSemanticAppearanceSelectionSource:
      remote<typeof SemanticNative.inspectPdfSemanticAppearanceSelectionSource>(pageId, "inspectPdfSemanticAppearanceSelectionSource"),
  };
}

/** The planners one page's window (and every module it hands them to)
 * calls, by the planner's own name and argument shape. */
export type PdfNativePlanners = ReturnType<typeof bindPdfNativePlanners>;

const boundSets = new Map<string | null, PdfNativePlanners>();
const BOUND_SET_LIMIT = 64;

/**
 * The planners bound to one page: each call finds its working head among
 * its byte arguments under that page's registration only
 * (`pdfNativeWorkingHead(sha256, pageId)`), never another page's that holds
 * the very same bytes (one PDF uploaded twice, both open in one tab). The
 * same page is always handed the same set.
 *
 * `null` binds no page: such a call is answered only for bytes a single
 * page holds, and refused (`source-mismatch`) for twins. A reader host
 * passes `PDF_NATIVE_READER_PAGE`.
 */
export function createPdfNativePlanners(pageId: string | null): PdfNativePlanners {
  const key = pageId || null;
  let set = boundSets.get(key);
  if (set) return set;
  set = bindPdfNativePlanners(key);
  boundSets.set(key, set);
  while (boundSets.size > BOUND_SET_LIMIT) boundSets.delete(boundSets.keys().next().value as string | null);
  return set;
}
