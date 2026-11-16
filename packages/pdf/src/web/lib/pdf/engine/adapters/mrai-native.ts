/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  MRAI_NATIVE_CONTENT_STREAM_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION,
  MRAI_NATIVE_LIST_TEXT_ACTION,
  MRAI_NATIVE_OUTLINE_ACTION,
} from "../native/vocabulary";
export {
  MRAI_NATIVE_CONTENT_STREAM_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION,
  MRAI_NATIVE_LIST_TEXT_ACTION,
  MRAI_NATIVE_OUTLINE_ACTION,
};
import { pdfNativeSourceCheckpointClaim, type PdfNativeSourceCheckpoint } from "../native/source-checkpoint";
import { commitMraiNativePlannedRecord } from "./mrai-native-commit";
import {
  bindMraiNativePreparedTextRecord,
  bindMraiNativePreparedTextWindowHistoryRecord,
  pdfNativePdfaIdentificationRecordField,
  pdfNativePdfaIdentificationRevisedObjects,
  stableRecord,
  stableTextWindowHistoryRecord,
} from "./mrai-native-record";
export { bindMraiNativePreparedTextRecord, bindMraiNativePreparedTextWindowHistoryRecord };
import {
  refusedPdfEngineResult,
  type PdfEngineAdapter,
  type PdfEngineAdapterContext,
  type PdfEngineCommand,
  type PdfEngineConflictResult,
  type PdfEngineResult,
  type PdfEngineSourceIdentity,
} from "../contracts";
import type {
  PdfSharedStateBridge,
  PdfSharedStateRecord,
} from "../shared-state-bridge";
import {
  MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION,
  MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION,
  MRAI_NATIVE_OUTLINE_SCHEMA_VERSION,
  MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES,
  PdfNativeTransactionError,
  planMraiNativeEmbeddedImageRemoval,
  planMraiNativeEmbeddedImageRemovalHistory,
  planMraiNativeEmbeddedImageTransform,
  planMraiNativeEmbeddedImageTransformHistory,
  planMraiNativeEmbeddedImageReplacement,
  planMraiNativeEmbeddedImageReplacementHistory,
  planMraiNativeOutlineMutation,
  planMraiNativePageBackground,
  planMraiNativePageSplit,
  prepareMraiNativeContentStreamTransaction,
  readMraiNativePdfXrefSectionCount,
  sha256PdfBytes,
  verifyMraiNativeContentStreamCandidate,
  verifyMraiNativeEmbeddedImageReplacementCandidate,
  type PdfNativeContentStreamTransactionRequest,
  type PdfNativeEmbeddedImageRemovalEvidence,
  type PdfNativeEmbeddedImageReplacementEvidence,
  type PdfNativeEmbeddedImageTransformEvidence,
  type PdfNativeEmbeddedImageTransformIntent,
  type PdfNativeEmbeddedImageTransformMatrix,
  type PdfNativeObjectPrecondition,
  type PdfNativePreparedEmbeddedImageReplacement,
  type PdfNativePreparedOutlineMutation,
  type PdfNativePreparedPageBackground,
  type PdfNativePreparedPageSplit,
} from "../native/content-stream-transaction";
import {
  MRAI_NATIVE_PAGE_BACKGROUND_ACTION,
  MRAI_NATIVE_PAGE_BACKGROUND_KIND,
  MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION,
  PDF_NATIVE_PAGE_BACKGROUND_CONTROL_ID,
  PDF_NATIVE_PAGE_BACKGROUND_FEATURE_ID,
  MRAI_NATIVE_PAGE_SPLIT_ACTION,
  MRAI_NATIVE_PAGE_SPLIT_KIND,
  MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION,
  PDF_NATIVE_PAGE_SPLIT_CONTROL_ID,
  PDF_NATIVE_PAGE_SPLIT_FEATURE_ID,
} from "../native/vocabulary";
import {
  readPdfNativePageSplitPayload,
  type PdfNativePageSplitPayload,
} from "../native/page-split-commands";
import {
  readPdfNativePageBackgroundPayload,
  type PdfNativePageBackgroundPayload,
} from "../native/page-background-commands";
import {
  validateManagedOutlineTree,
  type ManagedOutlineNode,
} from "@mrnote/core/web/lib/pdf/outline.ts";
import {
  MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION,
  MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION,
  MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION,
  MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION,
  MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION,
  MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION,
  PDF_NATIVE_MAX_FIXED_FRAME_ITEMS,
  planMraiNativeListTextHistoryReplacement,
  planMraiNativeTextWindowHistory,
  MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION,
  planMraiNativeListTextPropertyReplacement,
  planMraiNativeListTextParagraphFormat,
  pdfNativeParagraphPlanClearOfOtherText,
  planMraiNativeListTextSelectionReplacement,
  type PdfNativeListTextFidelityEvidence,
  type PdfNativeListTextSelectionEvidence,
  type PdfNativeTextPropertyDraft,
  type PdfNativeTextPropertyProof,
} from "../native/list-text-fidelity";
import type { JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import {
  MRAI_NATIVE_VECTOR_DELETE_ACTION,
  MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION,
  MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
  planPdfNativeVectorSelectionDelete,
  planPdfNativeVectorSelectionEdit,
  readPdfNativeVectorDeletePayload,
  readPdfNativeVectorDuplicatePayload,
  readPdfNativeVectorReorderPayload,
  readPdfNativeVectorSelectionEditPayload,
  type PdfNativeVectorDeletePayload,
  type PdfNativeVectorSelectionDeletePlan,
  type PdfNativeVectorSelectionEditPayload,
  type PdfNativeVectorSelectionEditPlan,
  MRAI_NATIVE_VECTOR_DUPLICATE_ACTION,
  MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
  MRAI_NATIVE_VECTOR_REORDER_ACTION,
  MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
  planPdfNativeVectorSelectionDuplicate,
  planPdfNativeVectorSelectionReorder,
  type PdfNativeVectorDuplicatePayload,
  type PdfNativeVectorReorderPayload,
  type PdfNativeVectorSelectionDuplicatePlan,
  type PdfNativeVectorSelectionReorderPlan,
} from "../../native-vector-selection-edit";
import {
  MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION,
  MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION,
  planPdfNativeVectorStyleSelectionEdit,
  readPdfNativeVectorStyleSelectionEditPayload,
  type PdfNativeVectorStyleSelectionEditPayload,
  type PdfNativeVectorStyleSelectionEditPlan,
} from "../../native-vector-style-selection-edit";
import {
  MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
  pdfNativeEmbeddedFontLibraryFromPrograms,
  planPdfNativeEmbeddedFontParagraphFormatEdit,
  planPdfNativeEmbeddedFontPropertyEdit,
  planPdfNativeFontExtensionTextEdit,
  planPdfNativeFontExtensionTextHistory,
  type PdfNativeFontExtensionTextHistoryPlan,
  type PdfNativeFontExtensionTextPlan,
} from "../../native-font-extension-text-edit";
import { isPdfNativeEmbeddedFontKey } from "../../native-embedded-font-key";

import {
  MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION,
  MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
  planMraiNativeParagraphProperties,
  readMraiNativeParagraphPropertyChanges,
  verifyMraiNativeParagraphProperties,
  type PdfNativeParagraphPropertyChanges,
  type PdfNativeParagraphPropertyPlan,
} from "../native/paragraph-properties-transaction";
import { readPdfCompiledTextPreview, type PdfCompiledTextPreview } from "../../native-text-compiled-preview";
import {
  MRAI_NATIVE_TEXT_MOVE_KIND,
  MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
  planMraiNativeTextMove,
  type PdfNativeTextMovePlan,
} from "../native/text-move";
import {
  MRAI_NATIVE_TEXT_GROUP_MOVE_KIND,
  MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
  planMraiNativeTextGroupMove,
  planMraiNativeTextGroupMoveRestore,
  readMraiNativeTextGroupTransform,
  type PdfNativeTextGroupMoveEvidence,
  type PdfNativeTextGroupEdit,
} from "../native/text-group-move";
import type { PdfNativePreparedContentStreamTransaction } from "../native/content-stream-transaction";
import {
  MRAI_NATIVE_CONTENT_APPEND_ACTION,
  MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
  planPdfNativeContentAppend,
  readPdfNativeContentAppendPayload,
  type PdfNativeContentAppendPayload,
  type PdfNativeContentAppendPlanResult,
} from "../native/content-append-transaction";
import { fetchVerifiedPdfComplexScriptFont } from "../../unicode-font-asset";

export {
  MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION,
  MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION,
} from "../../native-font-extension-text-edit";

import {
  MRAI_NATIVE_TEXT_DECORATION_ACTION,
  MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
  planPdfNativeTextDecoration,
  readPdfNativeTextDecorationPayload,
  type PdfNativeTextDecorationPayload,
  type PdfNativeTextDecorationPlanResult,
} from "../../native-text-decoration";
import {
  MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
  MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
  planPdfNativeTextObjectDelete,
  readPdfNativeTextObjectDeletePayload,
  type PdfNativeTextObjectDeletePayload,
  type PdfNativeTextObjectDeletePlanResult,
} from "../../native-text-object-delete";
import {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
  planPdfNativeEmbeddedImageReorder,
  readPdfNativeEmbeddedImageReorderPayload,
  type PdfNativeEmbeddedImageReorderPayload,
  type PdfNativeEmbeddedImageReorderPlanResult,
} from "../../native-embedded-image-reorder";

export {
  MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
  MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
  MRAI_NATIVE_TEXT_DECORATION_ACTION,
};

const SHA256_PATTERN = /^[a-f\d]{64}$/i;

export interface MraiNativePdfAdapterOptions {
  readonly compiledTextPreview?: PdfCompiledTextPreview | null;
  readonly sourceCheckpoint?: PdfNativeSourceCheckpoint;
  /** Exact saved forward paragraph record; the server resolves its id from the journal. */
  readonly textWindowHistoryAuthority?: Readonly<{
    originRecord: PdfSharedStateRecord;
    direction: "undo" | "redo";
  }>;
  readonly paragraphHistoryOrigin?: PdfSharedStateRecord;
  readonly loadPinnedSource: (
    source: PdfEngineSourceIdentity,
  ) => Uint8Array | Promise<Uint8Array>;
  readonly sharedStateBridge: PdfSharedStateBridge;
  /** Trusted bundled font bytes used only for a freshly proven new-glyph edit. */
  readonly loadNativeFontExtensionFont?: () =>
    | Uint8Array
    | ArrayBuffer
    | Promise<Uint8Array | ArrayBuffer>;
  /** Exact accepted origin row and direction for one font-history replay. */
  readonly fontExtensionHistoryAuthority?: Readonly<{
    originRecord: PdfSharedStateRecord;
    direction: "undo" | "redo";
  }>;
  readonly collection?: string;
  /**
   * The exact bytes this command will save, before they are sent: a host
   * may paint them while the server replays and acknowledges the write.
   * Display only -- the save still stands or falls on the server's receipt.
   */
  readonly onCandidatePrepared?: (
    candidate: Readonly<{ bytes: Uint8Array; sha256: string }>,
  ) => void;
  /** Exact receipt/journal authority for a previously generated native PDF. */
  readonly allowReceiptBoundPreviousRevisions?: boolean;
  /** Exact opposite-side v4 authority for one compensating history command. */
  readonly listTextHistoryAuthority?: Readonly<{
    positionedOriginTransactionId?: string;
    expectedSourceSha256: string;
    selectionRevision: number;
    priorFidelity: PdfNativeListTextFidelityEvidence;
    expectedCurrentCodeBytes: readonly number[];
    desiredCodeBytes: readonly number[];
    expectedCurrentOperatorBytes?: readonly number[];
    desiredOperatorBytes?: readonly number[];
  }>;
  /** Fresh v6 proof and requested property values for one Apply command. */
  /** The exact restore of a group move (its undo, text-group-move.ts): the
   * move it puts back and the bytes that stood before it and after it. */
  readonly textGroupMoveRestoreAuthority?: Readonly<{
    originTransactionId: string;
    origin: PdfNativeTextGroupMoveEvidence;
    originalBytes: Uint8Array;
    rewrittenBytes: Uint8Array;
  }>;
  readonly listTextPropertyAuthority?: Readonly<{
    proof: PdfNativeTextPropertyProof;
    properties: PdfNativeTextPropertyDraft;
    /** The rules drawn beside the run, rewritten along it in the same write. */
    keepTextRules?: readonly ("underline" | "strikethrough")[];
    /** The record of the edit that last turned this run (planner input). */
    previousTurn?: Parameters<typeof planMraiNativeListTextPropertyReplacement>[0]["previousTurn"];
    /** The served font slices a draft that sets text in a family the PDF
     * does not hold needs (native-embedded-font-key.ts), as the editor
     * fetched and verified them; checked against the lock again here. */
    embeddedFontPrograms?: readonly Readonly<{
      sha256: string;
      bytes: Uint8Array;
      /** A font from the reader's computer the page keeps (page-fonts.ts). */
      pageFont?: import("../../page-fonts").PdfPageFont;
    }>[];
    /** The served font list those slices come from, narrowed to the draft's
     * families (native-embedded-font-programs.ts): a planner that cannot ask
     * for the list (the browser's plan Worker) plans from this one. */
    embeddedFontManifest?: import("@mrnote/core/web/lib/office-fonts/embeddable-faces.ts").EmbeddableFontManifest | null;
  }>;
  /** The frame the reader resized a paragraph to (WPS: a text block dragged
   * narrower or wider re-wraps its text): the right edge, an x of the page's
   * own space. A plain text edit with it is set again by the paragraph
   * rewrite in that frame (planMraiNativeListTextSelectionReplacement). */
  readonly listTextParagraphFrame?: Readonly<{ right?: number; lineStep?: number; sizeScale?: number }>;
  /** Where the editor's rows start in a text edit's new text
   * (lib/pdf/native-text-row-starts): display only, no authority. */
  readonly listTextRowStarts?: readonly number[];
  /** A group move's paragraph set again in a frame reaching this right edge
   * (the page's user space) after it: the landing is planned with it. */
  readonly textGroupMoveFrameRight?: number;
  




  readonly listTextParagraphFormat?: Readonly<{
    format: import("../native/list-text-fidelity").PdfNativeParagraphFormat;
    fontOptions?: readonly import("../native/list-text-fidelity").PdfNativeTextFontOption[];
    embeddedFontPrograms?: readonly Readonly<{
      sha256: string;
      bytes: Uint8Array;
      pageFont?: import("../../page-fonts").PdfPageFont;
    }>[];
    embeddedFontManifest?: import("@mrnote/core/web/lib/office-fonts/embeddable-faces.ts").EmbeddableFontManifest | null;
    /** The paragraph set again would draw over other text of the page: the
     * write is refused instead of made with a notice (a caller with nobody
     * at the page to read one -- the server's AI commands). */
    refuseOverOtherText?: boolean;
  }>;
  /** Exact current source evidence for one receipt-bound image remove/restore. */
  readonly embeddedImageHistoryAuthority?: Readonly<{
    expectedSourceSha256: string;
    priorEvidence: PdfNativeEmbeddedImageRemovalEvidence;
    expectedCurrentOperatorBytes: readonly number[];
    desiredOperatorBytes: readonly number[];
  }>;
  /** Fresh receipt-bound authority for one replacement undo/redo. */
  readonly embeddedImageReplacementHistoryAuthority?: Readonly<{
    expectedSourceSha256: string;
    priorEvidence: PdfNativeEmbeddedImageReplacementEvidence;
    expectedCurrent: "original" | "replacement";
    desired: "original" | "replacement";
  }>;
  /** Exact current source evidence for one receipt-bound transform/restore. */
  readonly embeddedImageTransformHistoryAuthority?: Readonly<{
    expectedSourceSha256: string;
    priorEvidence: PdfNativeEmbeddedImageTransformEvidence;
    expectedCurrentOperatorBytes: readonly number[];
    desiredOperatorBytes: readonly number[];
  }>;
}

interface ParsedEmbeddedImageTransformPayload {
  readonly pageNumber: number;
  readonly imageOrder: number;
  readonly intent: PdfNativeEmbeddedImageTransformIntent | null;
}

interface ParsedNativePayload {
  readonly objectNumber: number;
  readonly generation: number;
  readonly request: Omit<PdfNativeContentStreamTransactionRequest, "identity">;
}

interface ParsedListPayload {
  readonly selection: PdfNativeListTextSelectionEvidence;
  readonly replacementText: string;
  readonly paragraphChanges?: PdfNativeParagraphPropertyChanges;
  /** A native move of the selected text (text-move.ts), in PDF user space. */
  readonly move?: Readonly<{ x: number; y: number }>;
  /** A native group move of a whole box, or the restore of one (text-group-move.ts). */
  readonly group?: Readonly<
    | { transform: PdfNativeTextGroupEdit; restore?: undefined }
    | { restore: Readonly<{ originTransactionId: string }>; transform?: undefined }
  >;
}

interface ParsedFontExtensionHistoryPayload {
  readonly originTransactionId: string;
  readonly direction: "undo" | "redo";
}

interface ParsedEmbeddedImagePayload {
  readonly pageNumber: number;
  readonly imageOrder: number;
}

interface ParsedEmbeddedImageReplacementPayload extends ParsedEmbeddedImagePayload {
  readonly replacementBytes: Uint8Array | null;
  readonly replacementMimeType: string | null;
}

interface ParsedOutlinePayload {
  readonly operation:
    | "add"
    | "rename"
    | "update-destination"
    | "move"
    | "delete"
    | "rename-source"
    | "delete-source";
  readonly managedTree: readonly ManagedOutlineNode[];
  readonly expectedManagedIds: readonly string[];
  /** The tree names the upload's pages on a file holding the page's
   * structure: for each file page, the upload's page it is. */
  readonly pageMap?: readonly (number | null)[];
  /** rename-source / delete-source: the file's own bookmark by its place in
   * the outline, and a rename's new title. */
  readonly sourceTarget?: Readonly<{ path: readonly number[]; title: string | null; expectedTitle?: string }>;
}

type ParsedVectorPayload = PdfNativeVectorSelectionEditPayload;
type ParsedVectorDeletePayload = PdfNativeVectorDeletePayload;
type ParsedVectorDuplicatePayload = PdfNativeVectorDuplicatePayload;
type ParsedVectorReorderPayload = PdfNativeVectorReorderPayload;
type ParsedVectorStylePayload = PdfNativeVectorStyleSelectionEditPayload;
type ParsedContentAppendPayload = PdfNativeContentAppendPayload;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function integer(value: unknown, label: string, allowZero = false): number {
  if (!Number.isSafeInteger(value) || Number(value) < (allowZero ? 0 : 1)) {
    throw new Error(`Invalid ${label}.`);
  }
  return Number(value);
}

function generation(value: unknown): number {
  const parsed = integer(value, "PDF object generation", true);
  if (parsed > 65_535) throw new Error("Invalid PDF object generation.");
  return parsed;
}

function sha256(value: unknown, label: string): string {
  if (typeof value !== "string" || !SHA256_PATTERN.test(value)) {
    throw new Error(`Invalid ${label} SHA-256.`);
  }
  return value.toLowerCase();
}

function replacementBytes(value: unknown): Uint8Array {
  if (
    !Array.isArray(value) ||
    value.some(
      (item) =>
        !Number.isInteger(item) || Number(item) < 0 || Number(item) > 255,
    )
  ) {
    throw new Error("Native replacement must be a JSON-safe byte array.");
  }
  return Uint8Array.from(value as number[]);
}

function untouchedObjects(value: unknown): PdfNativeObjectPrecondition[] {
  if (value === undefined) return [];
  if (!Array.isArray(value) || value.length > 128) {
    throw new Error("Invalid native untouched-object preconditions.");
  }
  return value.map((item) => {
    if (!isRecord(item)) {
      throw new Error("Invalid native untouched-object precondition.");
    }
    const role = item.role;
    if (role !== "resource" && role !== "untouched-object") {
      throw new Error("Invalid native untouched-object role.");
    }
    return {
      objectNumber: integer(item.objectNumber, "PDF object number"),
      generation: generation(item.generation),
      expectedObjectSha256: sha256(
        item.expectedObjectSha256,
        "untouched object",
      ),
      role,
    };
  });
}

function boundedText(value: unknown, label: string): string {
  if (typeof value !== "string" || value.length === 0 || value.length > 2_000) {
    throw new Error(`Invalid ${label}.`);
  }
  return value;
}

function finite(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`Invalid ${label}.`);
  }
  return value;
}

function parseSelectionItem(
  value: unknown,
): PdfNativeListTextSelectionEvidence["item"] {
  if (
    !isRecord(value) ||
    !Array.isArray(value.transform) ||
    value.transform.length !== 6 ||
    typeof value.hasEOL !== "boolean"
  ) {
    throw new Error("Invalid native selection item.");
  }
  return Object.freeze({
    itemIndex: integer(value.itemIndex, "selection item index", true),
    str: boundedText(value.str, "selected native text"),
    fontName: boundedText(value.fontName, "selection font name"),
    transform: Object.freeze(
      value.transform.map((entry, index) =>
        finite(entry, `selection transform ${index}`),
      ),
    ) as PdfNativeListTextSelectionEvidence["item"]["transform"],
    width: finite(value.width, "selection width"),
    height: finite(value.height, "selection height"),
    hasEOL: value.hasEOL,
  });
}

function parseListPayload(command: PdfEngineCommand): ParsedListPayload {
  const payload = command.payload;
  const schemaIsValid =
    isRecord(payload) &&
    (command.action === MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION
      ? payload.schemaVersion === MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA
      : command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION
      ? payload.schemaVersion === MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION
      : payload.schemaVersion === MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
        payload.schemaVersion === MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION ||
        payload.schemaVersion === MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION ||
        payload.schemaVersion === MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION ||
        payload.schemaVersion === MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION ||
        payload.schemaVersion === MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
        payload.schemaVersion ===
          MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
        payload.schemaVersion ===
          MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
        payload.schemaVersion ===
          MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION);
  if (
    !isRecord(payload) ||
    !schemaIsValid ||
    !isRecord(payload.selection)
  ) {
    throw new Error("Invalid MRAI native text payload.");
  }
  const item = parseSelectionItem(payload.selection.item);
  const items =
    payload.selection.items === undefined
      ? undefined
      : Array.isArray(payload.selection.items) &&
          payload.selection.items.length >= 2 &&
          payload.selection.items.length <= PDF_NATIVE_MAX_FIXED_FRAME_ITEMS
        ? payload.selection.items.map(parseSelectionItem)
        : (() => {
            throw new Error("Invalid fixed-frame native selection items.");
          })();
  const selectionRevision = integer(
    payload.selection.selectionRevision,
    "selection revision",
  );
  if (selectionRevision !== command.selectionRevision) {
    throw new Error("Native text selection revision is stale.");
  }
  // A move keeps the text: it names the translation instead of new words.
  const move = payload.schemaVersion === MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION
    ? (() => {
        if (!isRecord(payload.move) || payload.replacementText !== undefined ||
          command.action !== MRAI_NATIVE_LIST_TEXT_ACTION) {
          throw new Error("Invalid native text move payload.");
        }
        return Object.freeze({
          x: finite(payload.move.x, "native text move"),
          y: finite(payload.move.y, "native text move"),
        });
      })()
    : undefined;
  // A group move keeps the text too: it names the map of the page, or the
  // move it puts back exactly (text-group-move.ts).
  const group = payload.schemaVersion === MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION
    ? (() => {
        if (payload.replacementText !== undefined || command.action !== MRAI_NATIVE_LIST_TEXT_ACTION ||
          isRecord(payload.transform) === isRecord(payload.restore) ||
          (payload.bold !== undefined && !isRecord(payload.transform))) {
          throw new Error("Invalid native group move payload.");
        }
        if (isRecord(payload.transform)) {
          // The weight set with the map rides beside it (a bold alone keeps
          // the map the identity).
          return Object.freeze({ transform: readMraiNativeTextGroupTransform(payload.transform, payload.bold) });
        }
        const restore = payload.restore as Record<string, unknown>;
        if (typeof restore.originTransactionId !== "string" || !restore.originTransactionId.trim() ||
          restore.originTransactionId.length > 200 || Object.keys(restore).length !== 1) {
          throw new Error("Invalid native group move restore payload.");
        }
        return Object.freeze({ restore: Object.freeze({ originTransactionId: restore.originTransactionId }) });
      })()
    : undefined;
  return Object.freeze({
    selection: Object.freeze({
      sourceSha256: sha256(payload.selection.sourceSha256, "selection source"),
      selectionRevision,
      pageNumber: integer(payload.selection.pageNumber, "selection page"),
      startOffset: integer(
        payload.selection.startOffset,
        "selection start offset",
        true,
      ),
      endOffset: integer(
        payload.selection.endOffset,
        "selection end offset",
        true,
      ),
      item,
      ...(items ? { items: Object.freeze(items) } : {}),
    }),
    replacementText: boundedText(
      move || group ? item.str : payload.replacementText,
      "replacement native text",
    ),
    ...(move ? { move } : {}),
    ...(group ? { group } : {}),
    ...(command.action === MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION
      ? { paragraphChanges: readMraiNativeParagraphPropertyChanges(payload.changes) }
      : {}),
  });
}

function parseTextWindowHistoryPayload(command: PdfEngineCommand): ParsedFontExtensionHistoryPayload {
  const payload = command.payload;
  if (!isRecord(payload) || payload.schemaVersion !== MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION ||
      typeof payload.originTransactionId !== "string" || !payload.originTransactionId ||
      (payload.direction !== "undo" && payload.direction !== "redo") || command.selectionRevision !== 0 ||
      command.featureId !== "PDF-EDT-TEXT-01" ||
      !["PDFCTL.RIBBON.HOME.EDIT_CONTENT", "PDFCTL.RIBBON.EDIT.EDIT_CONTENT"].includes(command.controlId)) {
    throw new Error("Invalid native text-window history payload.");
  }
  return {originTransactionId: payload.originTransactionId, direction: payload.direction};
}

function parseFontExtensionHistoryPayload(
  command: PdfEngineCommand,
): ParsedFontExtensionHistoryPayload {
  const payload = command.payload;
  if (
    !isRecord(payload) ||
    payload.schemaVersion !==
      MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION ||
    typeof payload.originTransactionId !== "string" ||
    !payload.originTransactionId ||
    (payload.direction !== "undo" && payload.direction !== "redo") ||
    command.selectionRevision !== 0
  ) {
    throw new Error("Invalid native font extension history payload.");
  }
  return Object.freeze({
    originTransactionId: payload.originTransactionId,
    direction: payload.direction,
  });
}

function parsePayload(command: PdfEngineCommand): ParsedNativePayload {
  const payload = command.payload;
  if (
    !isRecord(payload) ||
    payload.schemaVersion !== MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION ||
    !isRecord(payload.target) ||
    !isRecord(payload.operatorSpan)
  ) {
    throw new Error("Invalid MRAI native content-stream payload.");
  }
  return Object.freeze({
    objectNumber: integer(payload.target.objectNumber, "PDF object number"),
    generation: generation(payload.target.generation),
    request: Object.freeze({
      expectedObjectSha256: sha256(
        payload.target.expectedObjectSha256,
        "target object",
      ),
      expectedStreamSha256: sha256(
        payload.target.expectedStreamSha256,
        "target stream",
      ),
      operatorSpan: Object.freeze({
        start: integer(payload.operatorSpan.start, "operator span start", true),
        end: integer(payload.operatorSpan.end, "operator span end", true),
        expectedSha256: sha256(
          payload.operatorSpan.expectedSha256,
          "operator span",
        ),
      }),
      replacement: replacementBytes(payload.replacement),
      ...(Array.isArray(payload.duplicateSpans)
        ? {
            duplicateSpans: Object.freeze(payload.duplicateSpans.map((span) => {
              if (!isRecord(span)) throw new Error("Invalid MRAI native duplicate span.");
              const partial = span.sourceStart !== undefined || span.sourceEnd !== undefined ||
                span.replacementStart !== undefined || span.replacementEnd !== undefined;
              return Object.freeze({
                start: integer(span.start, "duplicate span start", true),
                end: integer(span.end, "duplicate span end", true),
                // A fake-bold overprint of one text object of a run group
                // copies that object's runs and takes their part of the
                // replacement.
                ...(partial
                  ? {
                      sourceStart: integer(span.sourceStart, "duplicate span source start", true),
                      sourceEnd: integer(span.sourceEnd, "duplicate span source end", true),
                      replacementStart: integer(span.replacementStart, "duplicate span replacement start", true),
                      replacementEnd: integer(span.replacementEnd, "duplicate span replacement end", true),
                    }
                  : {}),
              });
            })),
          }
        : {}),
      ...(Array.isArray(payload.textRules)
        ? {
            textRules: Object.freeze(payload.textRules.map((rule) => {
              if (!isRecord(rule) || (rule.style !== "underline" && rule.style !== "strikethrough")) {
                throw new Error("Invalid MRAI native text rule span.");
              }
              return Object.freeze({
                style: rule.style,
                start: integer(rule.start, "text rule start", true),
                end: integer(rule.end, "text rule end", true),
                removed: replacementBytes(rule.removed),
                added: replacementBytes(rule.added),
              });
            })),
          }
        : {}),
      untouchedObjects: Object.freeze(
        untouchedObjects(payload.untouchedObjects),
      ),
    }),
  });
}

function parseEmbeddedImagePayload(
  command: PdfEngineCommand,
): ParsedEmbeddedImagePayload {
  const payload = command.payload;
  if (
    !isRecord(payload) ||
    payload.schemaVersion !==
      MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION ||
    !isRecord(payload.selection)
  ) {
    throw new Error("Invalid MRAI native embedded-image payload.");
  }
  return Object.freeze({
    pageNumber: integer(payload.selection.pageNumber, "embedded image page"),
    imageOrder: integer(
      payload.selection.imageOrder,
      "embedded image order",
      true,
    ),
  });
}

function parseEmbeddedImageReplacementPayload(
  command: PdfEngineCommand,
): ParsedEmbeddedImageReplacementPayload {
  const payload = command.payload;
  if (
    !isRecord(payload) ||
    payload.schemaVersion !==
      MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION ||
    !isRecord(payload.selection)
  ) {
    throw new Error("Invalid MRAI native embedded-image replacement payload.");
  }
  const base = Object.freeze({
    pageNumber: integer(payload.selection.pageNumber, "embedded image page"),
    imageOrder: integer(
      payload.selection.imageOrder,
      "embedded image order",
      true,
    ),
  });
  if (payload.historySchemaVersion !== undefined) {
    if (
      payload.historySchemaVersion !==
        MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION ||
      payload.replacementDataBase64 !== undefined ||
      payload.replacementMimeType !== undefined
    ) {
      throw new Error("Invalid embedded-image replacement history payload.");
    }
    return Object.freeze({
      ...base,
      replacementBytes: null,
      replacementMimeType: null,
    });
  }
  if (
    typeof payload.replacementDataBase64 !== "string" ||
    payload.replacementDataBase64.length < 4 ||
    payload.replacementDataBase64.length >
      Math.ceil(MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES / 3) * 4 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(payload.replacementDataBase64) ||
    payload.replacementDataBase64.length % 4 !== 0 ||
    (payload.replacementMimeType !== "image/jpeg" &&
      payload.replacementMimeType !== "image/png")
  ) {
    throw new Error("Invalid embedded-image replacement bytes.");
  }
  let binary: string;
  try {
    binary = atob(payload.replacementDataBase64);
  } catch {
    throw new Error("Invalid embedded-image replacement bytes.");
  }
  const replacementBytesValue = Uint8Array.from(binary, (character) =>
    character.charCodeAt(0),
  );
  if (
    replacementBytesValue.length < 1 ||
    replacementBytesValue.length > MAX_MRAI_NATIVE_EMBEDDED_IMAGE_BYTES
  ) {
    throw new Error("Invalid embedded-image replacement bytes.");
  }
  return Object.freeze({
    ...base,
    replacementBytes: replacementBytesValue,
    replacementMimeType: payload.replacementMimeType,
  });
}

function parseEmbeddedImageTransformPayload(
  command: PdfEngineCommand,
): ParsedEmbeddedImageTransformPayload {
  const payload = command.payload;
  if (
    !isRecord(payload) ||
    payload.schemaVersion !==
      MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION ||
    !isRecord(payload.selection)
  ) {
    throw new Error("Invalid MRAI native embedded-image transform payload.");
  }
  const base = Object.freeze({
    pageNumber: integer(payload.selection.pageNumber, "embedded image page"),
    imageOrder: integer(
      payload.selection.imageOrder,
      "embedded image order",
      true,
    ),
  });
  if (payload.historySchemaVersion !== undefined) {
    if (
      payload.historySchemaVersion !==
        MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION ||
      payload.intent !== undefined
    ) {
      throw new Error("Invalid embedded-image transform history payload.");
    }
    return Object.freeze({ ...base, intent: null });
  }
  const intent = payload.intent;
  if (!isRecord(intent)) {
    throw new Error("Invalid embedded-image transform intent.");
  }
  const ctm = intent.ctm;
  if (
    !Array.isArray(ctm) ||
    ctm.length !== 6 ||
    ctm.some((item) => typeof item !== "number" || !Number.isFinite(item))
  ) {
    throw new Error("Invalid embedded-image transform ctm.");
  }
  const readFrame = (value: unknown, label: string) => {
    if (
      !isRecord(value) ||
      ["x", "y", "width", "height"].some(
        (key) => typeof value[key] !== "number" || !Number.isFinite(value[key] as number),
      )
    ) {
      throw new Error(`Invalid embedded-image transform ${label}.`);
    }
    return Object.freeze({
      x: value.x as number,
      y: value.y as number,
      width: value.width as number,
      height: value.height as number,
    });
  };
  const frame = readFrame(intent.frame, "frame");
  const crop = intent.crop === null || intent.crop === undefined ? null : readFrame(intent.crop, "crop");
  const quarterTurns = intent.quarterTurns;
  if (![0, 1, 2, 3].includes(quarterTurns as number)) {
    throw new Error("Invalid embedded-image transform turns.");
  }
  const opacity = intent.opacity;
  if (
    opacity !== undefined &&
    (typeof opacity !== "number" ||
      !Number.isFinite(opacity) ||
      opacity < 0 ||
      opacity > 1)
  ) {
    throw new Error("Invalid embedded-image transform opacity.");
  }
  // The reader frame's gesture as one user-space affine (a free-angle turn
  // included); the planner refuses it beside turns or flips.
  const pageTransform = intent.pageTransform;
  if (
    pageTransform !== undefined &&
    pageTransform !== null &&
    (!Array.isArray(pageTransform) ||
      pageTransform.length !== 6 ||
      pageTransform.some(
        (item) => typeof item !== "number" || !Number.isFinite(item),
      ))
  ) {
    throw new Error("Invalid embedded-image transform page transform.");
  }
  // Word's Crop > Fill / Fit / Reset Crop (the planner refuses it beside any
  // other change).
  const cropMode = intent.cropMode;
  if (cropMode !== undefined && cropMode !== null && cropMode !== "fill" && cropMode !== "fit" && cropMode !== "reset") {
    throw new Error("Invalid embedded-image transform crop mode.");
  }
  return Object.freeze({
    ...base,
    intent: Object.freeze({
      ctm: Object.freeze([...(ctm as number[])]) as unknown as PdfNativeEmbeddedImageTransformMatrix,
      frame,
      quarterTurns: quarterTurns as 0 | 1 | 2 | 3,
      flipX: intent.flipX === true,
      flipY: intent.flipY === true,
      crop,
      ...(opacity === undefined ? {} : { opacity: opacity as number }),
      ...(typeof cropMode === "string" ? { cropMode: cropMode as "fill" | "fit" | "reset" } : {}),
      ...(Array.isArray(pageTransform)
        ? {
            pageTransform: Object.freeze([
              ...(pageTransform as number[]),
            ]) as unknown as PdfNativeEmbeddedImageTransformMatrix,
          }
        : {}),
    }),
  });
}

function parseOutlinePayload(command: PdfEngineCommand): ParsedOutlinePayload {
  const payload = command.payload;
  const operations = new Set([
    "add",
    "rename",
    "update-destination",
    "move",
    "delete",
    "rename-source",
    "delete-source",
  ]);
  if (
    !isRecord(payload) ||
    payload.schemaVersion !== MRAI_NATIVE_OUTLINE_SCHEMA_VERSION ||
    typeof payload.operation !== "string" ||
    !operations.has(payload.operation) ||
    !Array.isArray(payload.expectedManagedIds) ||
    // `every` with a predicate narrows the checked array to string ids.
    !payload.expectedManagedIds.every(
      (id): id is string =>
        typeof id === "string" && /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(id),
    ) ||
    new Set(payload.expectedManagedIds).size !== payload.expectedManagedIds.length
  ) {
    throw new Error("Invalid MRAI native outline payload.");
  }
  const sourceOperation = payload.operation === "rename-source" || payload.operation === "delete-source";
  const target = payload.sourceTarget;
  if (
    sourceOperation !== (target !== undefined) ||
    (target !== undefined && (
      !isRecord(target) ||
      !Array.isArray(target.path) ||
      target.path.length < 1 ||
      target.path.length > 13 ||
      !target.path.every((index) => Number.isSafeInteger(index) && (index as number) >= 0) ||
      (payload.operation === "rename-source"
        ? typeof target.title !== "string" || !target.title.trim() || target.title.length > 300
        : target.title !== null) ||
      (target.expectedTitle !== undefined &&
        (typeof target.expectedTitle !== "string" || target.expectedTitle.length > 300))
    ))
  ) {
    throw new Error("Invalid MRAI native outline payload.");
  }
  const pageMap = payload.pageMap;
  if (
    pageMap !== undefined &&
    (!Array.isArray(pageMap) ||
      pageMap.length < 1 ||
      pageMap.length > 100_000 ||
      !pageMap.every((page) => page === null || (Number.isSafeInteger(page) && (page as number) >= 1)))
  ) {
    throw new Error("Invalid MRAI native outline payload.");
  }
  return Object.freeze({
    operation: payload.operation as ParsedOutlinePayload["operation"],
    managedTree: validateManagedOutlineTree(payload.managedTree),
    expectedManagedIds: Object.freeze([...payload.expectedManagedIds]),
    ...(target !== undefined && isRecord(target)
      ? { sourceTarget: Object.freeze({
          path: Object.freeze([...(target.path as number[])]),
          title: (target.title as string | null),
          ...(typeof target.expectedTitle === "string" ? { expectedTitle: target.expectedTitle } : {}),
        }) }
      : {}),
    ...(Array.isArray(pageMap) ? { pageMap: Object.freeze([...(pageMap as (number | null)[])]) } : {}),
  });
}

function sourceMatches(
  left: PdfEngineSourceIdentity,
  right: PdfEngineSourceIdentity,
): boolean {
  return (
    left.resourceId === right.resourceId &&
    left.resourceVersionId === right.resourceVersionId &&
    left.sourceSha256 === right.sourceSha256
  );
}

function selectionsEqual(
  left: PdfNativeListTextSelectionEvidence,
  right: PdfNativeListTextSelectionEvidence,
): boolean {
  const itemsEqual =
    left.items === undefined
      ? right.items === undefined
      : right.items !== undefined &&
        left.items.length === right.items.length &&
        left.items.every(
          (item, index) =>
            JSON.stringify(item) === JSON.stringify(right.items![index]),
        );
  return (
    itemsEqual &&
    left.sourceSha256 === right.sourceSha256 &&
    left.selectionRevision === right.selectionRevision &&
    left.pageNumber === right.pageNumber &&
    left.startOffset === right.startOffset &&
    left.endOffset === right.endOffset &&
    JSON.stringify(left.item) === JSON.stringify(right.item)
  );
}

function conflict(
  command: PdfEngineCommand,
  context: Readonly<{ source: PdfEngineSourceIdentity }>,
  reason: PdfEngineConflictResult["reason"],
  latestSourceSha256 = context.source.sourceSha256,
): PdfEngineConflictResult {
  return Object.freeze({
    status: "conflict",
    transactionId: command.transactionId,
    reason,
    latestResourceVersionId: context.source.resourceVersionId,
    latestSourceSha256,
    latestSessionRevision: command.sessionRevision,
  });
}

function mapNativeError(
  error: PdfNativeTransactionError,
  command: PdfEngineCommand,
  context: Readonly<{ source: PdfEngineSourceIdentity }>,
): PdfEngineResult {
  if (error.code === "source-mismatch") {
    return conflict(command, context, "source-mismatch");
  }
  if (error.code === "precondition-mismatch") {
    return conflict(command, context, "object-precondition-mismatch");
  }
  if (error.code === "verification-failed") {
    return refusedPdfEngineResult(
      "verification-failed",
      error.message,
      command.transactionId,
    );
  }
  return refusedPdfEngineResult(
    error.code === "unsupported-document"
      ? "unsupported-document"
      : "invalid-command",
    error.message,
    command.transactionId,
  );
}


function stableVectorRecord(
  command: PdfEngineCommand,
  plan: PdfNativeVectorSelectionEditPlan,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION,
    kind: "vector-geometry-edit",
    vectorGeometry: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableVectorDeleteRecord(
  command: PdfEngineCommand,
  plan: PdfNativeVectorSelectionDeletePlan,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION,
    kind: "vector-delete",
    vectorDeletion: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableVectorDuplicateRecord(
  command: PdfEngineCommand,
  plan: PdfNativeVectorSelectionDuplicatePlan,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION,
    kind: "vector-duplicate",
    vectorDuplicate: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableVectorReorderRecord(
  command: PdfEngineCommand,
  plan: PdfNativeVectorSelectionReorderPlan,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION,
    kind: "vector-reorder",
    vectorReorder: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableEmbeddedImageReorderRecord(
  command: PdfEngineCommand,
  plan: PdfNativeEmbeddedImageReorderPlanResult,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION,
    kind: "embedded-image-reorder",
    embeddedImageReorder: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableTextDecorationRecord(
  command: PdfEngineCommand,
  plan: PdfNativeTextDecorationPlanResult,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION,
    kind: "text-decoration",
    textDecoration: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

/** A native move of one item's text (text-move.ts): the text edit command's
 * record, with the move evidence as its only text authority. */
export function stableMraiNativeTextMoveRecord(
  command: PdfEngineCommand,
  plan: PdfNativeTextMovePlan,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION,
    kind: MRAI_NATIVE_TEXT_MOVE_KIND,
    textMove: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

/** A native group move or resize of a whole box (text-group-move.ts), or its
 * exact restore: the text edit command's record, with the group evidence as
 * its only text authority. */
export function stableMraiNativeTextGroupMoveRecord(
  command: PdfEngineCommand,
  plan: Readonly<{
    prepared: PdfNativePreparedContentStreamTransaction;
    evidence: PdfNativeTextGroupMoveEvidence;
  }>,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION,
    kind: MRAI_NATIVE_TEXT_GROUP_MOVE_KIND,
    textGroupMove: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableTextObjectDeleteRecord(
  command: PdfEngineCommand,
  plan: PdfNativeTextObjectDeletePlanResult,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION,
    kind: "text-object-delete",
    textObjectDelete: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableVectorStyleRecord(
  command: PdfEngineCommand,
  plan: PdfNativeVectorStyleSelectionEditPlan,
): PdfSharedStateRecord {
  return {
    ...stableRecord(command, plan.prepared),
    schemaVersion: MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION,
    kind: "vector-style-edit",
    vectorStyle: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableContentAppendRecord(
  command: PdfEngineCommand,
  plan: PdfNativeContentAppendPlanResult,
): PdfSharedStateRecord {
  const prepared = plan.prepared;
  const receipt = prepared.contentAppend;
  if (!receipt) {
    throw new Error("Content append receipt is missing.");
  }
  const revisedObjects = [
    {
      objectNumber: prepared.request.identity.objectNumber,
      generation: prepared.request.identity.generation,
      offset: prepared.changedObjectOffset,
      sha256: prepared.changedObjectSha256,
    },
    ...receipt.objects,
    ...(receipt.resources ? [receipt.resources] : []),
    ...(receipt.page ? [receipt.page] : []),
    ...pdfNativePdfaIdentificationRevisedObjects(prepared),
  ]
    .sort(
      (left, right) =>
        left.objectNumber - right.objectNumber ||
        left.generation - right.generation,
    )
    .map((item) => ({ ...item }));
  const base = stableRecord(command, prepared);
  return {
    ...base,
    schemaVersion: MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION,
    kind: "content-append",
    derivedCandidate: {
      ...(base.derivedCandidate as Record<string, JsonValue>),
      revisedObjects,
    },
    contentAppend: JSON.parse(JSON.stringify(plan.evidence)) as JsonValue,
  };
}

function stableEmbeddedImageReplacementRecord(
  command: PdfEngineCommand,
  prepared: PdfNativePreparedEmbeddedImageReplacement,
): PdfSharedStateRecord {
  const evidence = prepared.evidence;
  return {
    id: command.transactionId,
    schemaVersion: MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION,
    transactionSchemaVersion: MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    kind: "embedded-image-replacement",
    featureId: command.featureId,
    controlId: command.controlId,
    action: command.action,
    transactionId: command.transactionId,
    baseSessionRevision: command.sessionRevision,
    selectionRevision: command.selectionRevision,
    source: {
      resourceId: command.resourceId,
      resourceVersionId: command.resourceVersionId,
      sha256: command.sourceSha256,
    },
    target: null,
    operatorSpan: null,
    replacement: null,
    replacementSha256: null,
    untouchedObjects: [],
    ...pdfNativePdfaIdentificationRecordField(prepared),
    derivedCandidate: {
      sha256: prepared.candidateSha256,
      byteLength: prepared.candidateBytes.length,
      sourceLength: prepared.sourceLength,
      sourceXrefOffset: prepared.sourceXrefOffset,
      incrementalXrefOffset: prepared.incrementalXrefOffset,
      revisedObjects: prepared.revisedObjects.map((item) => ({ ...item })),
    },
    embeddedImageReplacement: {
      schemaVersion: evidence.schemaVersion,
      pageNumber: evidence.pageNumber,
      imageOrder: evidence.imageOrder,
      originalResourceName: evidence.originalResourceName,
      replacementResourceName: evidence.replacementResourceName,
      page: { ...evidence.page },
      resourceOwner: { ...evidence.resourceOwner },
      originalImage: { ...evidence.originalImage },
      replacementImage: { ...evidence.replacementImage },
      ...(evidence.replacementSoftMask
        ? { replacementSoftMask: { ...evidence.replacementSoftMask } }
        : {}),
      contentStream: { ...evidence.contentStream },
      originalOperatorBytes: [...evidence.originalOperatorBytes],
      replacementOperatorBytes: [...evidence.replacementOperatorBytes],
      replacement: {
        ...evidence.replacement,
        decodeParms: evidence.replacement.decodeParms
          ? { ...evidence.replacement.decodeParms }
          : null,
      },
    },
  };
}

function stableOutlineRecord(
  command: PdfEngineCommand,
  prepared: PdfNativePreparedOutlineMutation,
): PdfSharedStateRecord {
  const evidence = prepared.evidence;
  const copyNode = (
    node: (typeof evidence.managedTree)[number],
  ): { [key: string]: JsonValue } => ({
    id: node.id,
    title: node.title,
    destination: {
      pageNumber: node.destination.pageNumber,
      page: { ...node.destination.page },
      left: node.destination.left,
      top: node.destination.top,
      zoom: node.destination.zoom,
    },
    children: node.children.map(copyNode),
  });
  return {
    id: command.transactionId,
    schemaVersion: MRAI_NATIVE_OUTLINE_SCHEMA_VERSION,
    transactionSchemaVersion: MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    kind: "outline-mutation",
    featureId: command.featureId,
    controlId: command.controlId,
    action: command.action,
    transactionId: command.transactionId,
    baseSessionRevision: command.sessionRevision,
    selectionRevision: command.selectionRevision,
    source: {
      resourceId: command.resourceId,
      resourceVersionId: command.resourceVersionId,
      sha256: command.sourceSha256,
    },
    target: null,
    operatorSpan: null,
    replacement: null,
    replacementSha256: null,
    untouchedObjects: [],
    derivedCandidate: {
      sha256: prepared.candidateSha256,
      byteLength: prepared.candidateBytes.length,
      sourceLength: prepared.sourceLength,
      sourceXrefOffset: prepared.sourceXrefOffset,
      incrementalXrefOffset: prepared.incrementalXrefOffset,
      revisedObjects: prepared.revisedObjects.map((item) => ({ ...item })),
    },
    outlineMutation: {
      schemaVersion: evidence.schemaVersion,
      operation: evidence.operation,
      expectedCatalog: { ...evidence.expectedCatalog },
      expectedOutlineRoot: evidence.expectedOutlineRoot
        ? { ...evidence.expectedOutlineRoot }
        : null,
      managedTree: evidence.managedTree.map(copyNode),
      expectedManagedIds: [...evidence.expectedManagedIds],
      ...(evidence.sourceTarget ? { sourceTarget: { ...evidence.sourceTarget } } : {}),
      ...(evidence.pageMap ? { pageMap: [...evidence.pageMap] } : {}),
    },
  };
}

/** WPS 分割页面 (and the merge that undoes it): the record the server replays
 * (`_replay_native_page_split_transaction`). */
function stablePageSplitRecord(
  command: PdfEngineCommand,
  prepared: PdfNativePreparedPageSplit,
): PdfSharedStateRecord {
  const evidence = prepared.evidence;
  return {
    id: command.transactionId,
    schemaVersion: MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION,
    transactionSchemaVersion: MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    kind: MRAI_NATIVE_PAGE_SPLIT_KIND,
    featureId: command.featureId,
    controlId: command.controlId,
    action: command.action,
    transactionId: command.transactionId,
    baseSessionRevision: command.sessionRevision,
    selectionRevision: command.selectionRevision,
    source: {
      resourceId: command.resourceId,
      resourceVersionId: command.resourceVersionId,
      sha256: command.sourceSha256,
    },
    target: null,
    operatorSpan: null,
    replacement: null,
    replacementSha256: null,
    untouchedObjects: [],
    derivedCandidate: {
      sha256: prepared.candidateSha256,
      byteLength: prepared.candidateBytes.length,
      sourceLength: prepared.sourceLength,
      sourceXrefOffset: prepared.sourceXrefOffset,
      incrementalXrefOffset: prepared.incrementalXrefOffset,
      revisedObjects: prepared.revisedObjects.map((item) => ({ ...item })),
    },
    pageSplit: JSON.parse(JSON.stringify(evidence)) as JsonValue,
  };
}


/** WPS 背景 (and the restore that undoes it): the record the server replays
 * (`_replay_native_page_background_transaction`). */
export function stablePageBackgroundRecord(
  command: PdfEngineCommand,
  prepared: PdfNativePreparedPageBackground,
): PdfSharedStateRecord {
  return {
    id: command.transactionId,
    schemaVersion: MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION,
    transactionSchemaVersion: MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION,
    kind: MRAI_NATIVE_PAGE_BACKGROUND_KIND,
    featureId: command.featureId,
    controlId: command.controlId,
    action: command.action,
    transactionId: command.transactionId,
    baseSessionRevision: command.sessionRevision,
    selectionRevision: command.selectionRevision,
    source: {
      resourceId: command.resourceId,
      resourceVersionId: command.resourceVersionId,
      sha256: command.sourceSha256,
    },
    target: null,
    operatorSpan: null,
    replacement: null,
    replacementSha256: null,
    untouchedObjects: [],
    ...pdfNativePdfaIdentificationRecordField(prepared),
    derivedCandidate: {
      sha256: prepared.candidateSha256,
      byteLength: prepared.candidateBytes.length,
      sourceLength: prepared.sourceLength,
      sourceXrefOffset: prepared.sourceXrefOffset,
      incrementalXrefOffset: prepared.incrementalXrefOffset,
      revisedObjects: prepared.revisedObjects.map((item) => ({ ...item })),
    },
    pageBackground: JSON.parse(JSON.stringify(prepared.evidence)) as JsonValue,
  };
}

/** A planned native command: its record, its candidate bytes and how they
 * are released. Planning is the algorithm (it runs on the server's engine
 * host, `engine-host/engines/pdf-native.ts`); writing the record through the
 * shared-state bridge and checking the receipt is the editor's
 * (`mrai-native-commit.ts`). */
export interface MraiNativePlannedCommand {
  readonly status: "planned";
  readonly record: PdfSharedStateRecord;
  /** The bytes a host may paint while the write is replayed; null when the
   * candidate is not shown before its receipt (outline, picture replace). */
  readonly displayCandidate: Readonly<{ bytes: Uint8Array; sha256: string; sourceLength: number }> | null;
  /** The whole candidate (source + its appended revision). */
  readonly candidate: Readonly<{ bytes: Uint8Array; sha256: string; sourceLength: number }> | null;
  /** The candidate the durable receipt must name (before consolidation). */
  readonly expectedCandidateSha256: string;
  /** Every native writer but the raw content-stream replace saves through
   * the durable saveNow seam and needs an exact lossless receipt. */
  readonly durable: boolean;
  /** Zero every buffer the plan holds. */
  dispose(): void;
}

/** The adapter's options a plan reads (no bridge, no source loader, no
 * display hook): JSON-able authorities, plus the bundled font loader. */
export type MraiNativePlanOptions = Omit<
  MraiNativePdfAdapterOptions,
  "loadPinnedSource" | "sharedStateBridge" | "onCandidatePrepared" | "collection"
>;

/**
 * Plan one native command against the pinned source bytes `loadSource`
 * answers: the refusal or conflict the adapter returns, or the planned
 * record and candidate. Pure in its inputs; it persists nothing.
 */
export async function planMraiNativeCommand(
  command: PdfEngineCommand,
  context: Readonly<{ source: PdfEngineSourceIdentity }>,
  options: MraiNativePlanOptions,
  loadSource: () => Uint8Array | Promise<Uint8Array>,
): Promise<PdfEngineResult | MraiNativePlannedCommand> {
      const isTextWindowHistory = command.action === MRAI_NATIVE_LIST_TEXT_ACTION &&
        isRecord(command.payload) && command.payload.schemaVersion === MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION;
      let parsed:
        | ParsedNativePayload
        | ParsedListPayload
        | ParsedFontExtensionHistoryPayload
        | ParsedEmbeddedImagePayload
        | ParsedEmbeddedImageReplacementPayload
        | ParsedEmbeddedImageTransformPayload
        | PdfNativeEmbeddedImageReorderPayload
        | PdfNativeTextObjectDeletePayload
        | PdfNativeTextDecorationPayload
        | ParsedOutlinePayload
        | PdfNativePageSplitPayload
        | PdfNativePageBackgroundPayload
        | ParsedVectorPayload
        | ParsedVectorDeletePayload
        | ParsedVectorDuplicatePayload
        | ParsedVectorReorderPayload
        | ParsedVectorStylePayload
        | ParsedContentAppendPayload;
      try {
        parsed = isTextWindowHistory ? parseTextWindowHistoryPayload(command) :
          command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION
            ? parseFontExtensionHistoryPayload(command)
            : command.action === MRAI_NATIVE_LIST_TEXT_ACTION ||
                command.action === MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION ||
                command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION
              ? parseListPayload(command)
              : command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION
                ? parseEmbeddedImagePayload(command)
                : command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION
                  ? parseEmbeddedImageReplacementPayload(command)
                  : command.action === MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION
                  ? parseEmbeddedImageTransformPayload(command)
                  : command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION
                  ? readPdfNativeEmbeddedImageReorderPayload(command.payload)
                  : command.action === MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION
                  ? readPdfNativeTextObjectDeletePayload(command.payload)
                  : command.action === MRAI_NATIVE_TEXT_DECORATION_ACTION
                  ? readPdfNativeTextDecorationPayload(command.payload)
                  : command.action === MRAI_NATIVE_OUTLINE_ACTION
                    ? parseOutlinePayload(command)
                  : command.action === MRAI_NATIVE_PAGE_SPLIT_ACTION
                    ? readPdfNativePageSplitPayload(command.payload)
                  : command.action === MRAI_NATIVE_PAGE_BACKGROUND_ACTION
                    ? readPdfNativePageBackgroundPayload(command.payload)
                    : command.action ===
                        MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION
                      ? readPdfNativeVectorSelectionEditPayload(command.payload)
                      : command.action === MRAI_NATIVE_VECTOR_DELETE_ACTION
                        ? readPdfNativeVectorDeletePayload(command.payload)
                        : command.action === MRAI_NATIVE_VECTOR_DUPLICATE_ACTION
                          ? readPdfNativeVectorDuplicatePayload(command.payload)
                        : command.action === MRAI_NATIVE_VECTOR_REORDER_ACTION
                          ? readPdfNativeVectorReorderPayload(command.payload)
                        : command.action ===
                            MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION
                          ? readPdfNativeVectorStyleSelectionEditPayload(
                              command.payload,
                            )
                          : command.action === MRAI_NATIVE_CONTENT_APPEND_ACTION
                            ? readPdfNativeContentAppendPayload(command.payload)
                            : parsePayload(command);
      } catch (error) {
        return refusedPdfEngineResult(
          "invalid-command",
          error instanceof Error ? error.message : "Invalid native payload.",
          command.transactionId,
        );
      }

      let sourceBytes: Uint8Array;
      try {
        const loaded = await loadSource();
        sourceBytes = Uint8Array.from(loaded);
      } catch (error) {
        return refusedPdfEngineResult(
          "adapter-rejected",
          error instanceof Error ? error.message : "PDF source load failed.",
          command.transactionId,
        );
      }
      let fontExtensionPlan: PdfNativeFontExtensionTextPlan | undefined;
      let fontExtensionHistoryPlan:
        | PdfNativeFontExtensionTextHistoryPlan
        | undefined;
      let fontExtensionFontBytes: Uint8Array | undefined;
      const clearFontExtensionBytes = () => {
        if (
          command.action !== MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION &&
          command.action !== MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION
        ) {
          return;
        }
        fontExtensionPlan?.prepared.candidateBytes.fill(0);
        fontExtensionHistoryPlan?.prepared.candidateBytes.fill(0);
        fontExtensionFontBytes?.fill(0);
        sourceBytes.fill(0);
      };
      const loadedSha256 = await sha256PdfBytes(sourceBytes);
      if (loadedSha256 !== context.source.sourceSha256) {
        const result = conflict(
          command,
          context,
          "source-mismatch",
          loadedSha256,
        );
        clearFontExtensionBytes();
        return result;
      }
      let checkpointClaim: ReturnType<typeof pdfNativeSourceCheckpointClaim> = null;
      try {
        if (
          options.sourceCheckpoint && ![
            MRAI_NATIVE_LIST_TEXT_ACTION, MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION,
            MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION, MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION,
            MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION, MRAI_NATIVE_VECTOR_DELETE_ACTION,
            MRAI_NATIVE_VECTOR_DUPLICATE_ACTION, MRAI_NATIVE_VECTOR_REORDER_ACTION,
            MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION, MRAI_NATIVE_CONTENT_APPEND_ACTION,
            MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION, MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION,
            MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION,
            MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION,
            MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION,
            MRAI_NATIVE_TEXT_DECORATION_ACTION,
            MRAI_NATIVE_OUTLINE_ACTION,
            MRAI_NATIVE_PAGE_SPLIT_ACTION,
            MRAI_NATIVE_PAGE_BACKGROUND_ACTION,
          ].some((action) => action === command.action)
        ) {
          clearFontExtensionBytes();
          return refusedPdfEngineResult("adapter-rejected", "The source checkpoint does not authorize this writer.", command.transactionId);
        }
        const xrefSections = readMraiNativePdfXrefSectionCount(sourceBytes);
        checkpointClaim = pdfNativeSourceCheckpointClaim(options.sourceCheckpoint, context.source, xrefSections);
        if (options.sourceCheckpoint && !checkpointClaim) {
          clearFontExtensionBytes();
          return refusedPdfEngineResult("adapter-rejected", "The source checkpoint is stale or invalid.", command.transactionId);
        }
        if (
          xrefSections > 1 && !checkpointClaim &&
          options.allowReceiptBoundPreviousRevisions !== true
        ) {
          const result = refusedPdfEngineResult(
            "unsupported-document",
            "Pre-existing PDF revisions require an exact native receipt and journal authority.",
            command.transactionId,
          );
          clearFontExtensionBytes();
          return result;
        }
      } catch (error) {
        if (error instanceof PdfNativeTransactionError) {
          const result = mapNativeError(error, command, context);
          clearFontExtensionBytes();
          return result;
        }
        const result = refusedPdfEngineResult(
          "adapter-rejected",
          error instanceof Error
            ? error.message
            : "PDF revision-chain verification failed.",
          command.transactionId,
        );
        clearFontExtensionBytes();
        return result;
      }

      let prepared: Awaited<
        ReturnType<typeof prepareMraiNativeContentStreamTransaction>
      > | undefined;
      let textWindowPlan: Awaited<ReturnType<typeof planMraiNativeTextWindowHistory>> | undefined;
      let listFidelity: PdfNativeListTextFidelityEvidence | undefined;
      let reusedCandidateBytes: Uint8Array | undefined;
      let paragraphPlan: PdfNativeParagraphPropertyPlan | undefined;
      let embeddedImageTransform:
        | PdfNativeEmbeddedImageTransformEvidence
        | undefined;
      let embeddedImageRemoval:
        PdfNativeEmbeddedImageRemovalEvidence | undefined;
      let embeddedImageReplacement:
        PdfNativePreparedEmbeddedImageReplacement | undefined;
      let outlinePrepared: PdfNativePreparedOutlineMutation | undefined;
      let pageSplitPrepared: PdfNativePreparedPageSplit | undefined;
      let pageBackgroundPrepared: PdfNativePreparedPageBackground | undefined;
      let vectorPlan: PdfNativeVectorSelectionEditPlan | undefined;
      let vectorDeletePlan: PdfNativeVectorSelectionDeletePlan | undefined;
      let vectorDuplicatePlan: PdfNativeVectorSelectionDuplicatePlan | undefined;
      let vectorReorderPlan: PdfNativeVectorSelectionReorderPlan | undefined;
      let embeddedImageReorderPlan:
        | PdfNativeEmbeddedImageReorderPlanResult
        | undefined;
      let textObjectDeletePlan:
        | PdfNativeTextObjectDeletePlanResult
        | undefined;
      let textDecorationPlan: PdfNativeTextDecorationPlanResult | undefined;
      let textMovePlan: PdfNativeTextMovePlan | undefined;
      let textGroupMovePlan:
        | Readonly<{ prepared: PdfNativePreparedContentStreamTransaction; evidence: PdfNativeTextGroupMoveEvidence }>
        | undefined;
      let vectorStylePlan:
        | PdfNativeVectorStyleSelectionEditPlan
        | undefined;
      let contentAppendPlan: PdfNativeContentAppendPlanResult | undefined;
      try {
        if (command.action === MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION) {
          const payload = parsed as ParsedListPayload;
          if (!payload.paragraphChanges || command.featureId !== "PDF-EDT-TEXT-01" ||
              command.controlId !== "PDFCTL.RIBBON.EDIT.EDIT_CONTENT") {
            throw new Error("Paragraph property command authority is invalid.");
          }
          paragraphPlan = await planMraiNativeParagraphProperties({
            sourceBytes,
            selection: payload.selection,
            changes: payload.paragraphChanges,
            replacementText: payload.replacementText,
            ...(options.listTextHistoryAuthority ? { textHistory: options.listTextHistoryAuthority } : {}),
          });
        }
        if (command.action === MRAI_NATIVE_OUTLINE_ACTION) {
          const outlinePayload = parsed as ParsedOutlinePayload;
          outlinePrepared = await planMraiNativeOutlineMutation({
            sourceBytes,
            request: {
              expectedSourceSha256: context.source.sourceSha256,
              operation: outlinePayload.operation,
              expectedManagedIds: outlinePayload.expectedManagedIds,
              managedTree: outlinePayload.managedTree,
              ...(outlinePayload.sourceTarget ? { sourceTarget: outlinePayload.sourceTarget } : {}),
              ...(outlinePayload.pageMap ? { pageMap: outlinePayload.pageMap } : {}),
            },
          });
        }
        if (command.action === MRAI_NATIVE_PAGE_SPLIT_ACTION) {
          const pageSplitPayload = parsed as PdfNativePageSplitPayload;
          if (
            command.featureId !== PDF_NATIVE_PAGE_SPLIT_FEATURE_ID ||
            command.controlId !== PDF_NATIVE_PAGE_SPLIT_CONTROL_ID ||
            pageSplitPayload.sourceSha256 !== context.source.sourceSha256
          ) {
            throw new Error("Page split command authority is invalid.");
          }
          pageSplitPrepared = await planMraiNativePageSplit({
            sourceBytes,
            request: pageSplitPayload.request,
          });
        }
        if (command.action === MRAI_NATIVE_PAGE_BACKGROUND_ACTION) {
          const backgroundPayload = parsed as PdfNativePageBackgroundPayload;
          if (
            command.featureId !== PDF_NATIVE_PAGE_BACKGROUND_FEATURE_ID ||
            command.controlId !== PDF_NATIVE_PAGE_BACKGROUND_CONTROL_ID ||
            backgroundPayload.sourceSha256 !== context.source.sourceSha256
          ) {
            throw new Error("Page background command authority is invalid.");
          }
          pageBackgroundPrepared = await planMraiNativePageBackground({
            sourceBytes,
            request: backgroundPayload.request,
          });
        }
        if (command.action === MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION) {
          vectorPlan = await planPdfNativeVectorSelectionEdit({
            sourceBytes,
            payload: parsed as ParsedVectorPayload,
          });
        }
        if (command.action === MRAI_NATIVE_VECTOR_DELETE_ACTION) {
          vectorDeletePlan = await planPdfNativeVectorSelectionDelete({
            sourceBytes,
            payload: parsed as ParsedVectorDeletePayload,
          });
        }
        if (command.action === MRAI_NATIVE_VECTOR_DUPLICATE_ACTION) {
          vectorDuplicatePlan = await planPdfNativeVectorSelectionDuplicate({
            sourceBytes,
            payload: parsed as ParsedVectorDuplicatePayload,
          });
        }
        if (command.action === MRAI_NATIVE_VECTOR_REORDER_ACTION) {
          vectorReorderPlan = await planPdfNativeVectorSelectionReorder({
            sourceBytes,
            payload: parsed as ParsedVectorReorderPayload,
          });
        }
        if (command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION) {
          embeddedImageReorderPlan = await planPdfNativeEmbeddedImageReorder({
            sourceBytes,
            payload: parsed as PdfNativeEmbeddedImageReorderPayload,
          });
        }
        if (command.action === MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION) {
          textObjectDeletePlan = await planPdfNativeTextObjectDelete({
            sourceBytes,
            payload: parsed as PdfNativeTextObjectDeletePayload,
          });
        }
        if (command.action === MRAI_NATIVE_TEXT_DECORATION_ACTION) {
          textDecorationPlan = await planPdfNativeTextDecoration({
            sourceBytes,
            payload: parsed as PdfNativeTextDecorationPayload,
          });
        }
        if (
          command.action === MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION
        ) {
          vectorStylePlan = await planPdfNativeVectorStyleSelectionEdit({
            sourceBytes,
            payload: parsed as ParsedVectorStylePayload,
          });
        }
        if (command.action === MRAI_NATIVE_CONTENT_APPEND_ACTION) {
          contentAppendPlan = await planPdfNativeContentAppend({
            sourceBytes,
            payload: parsed as ParsedContentAppendPayload,
            ...(options.loadNativeFontExtensionFont
              ? { loadFont: options.loadNativeFontExtensionFont }
              : {}),
            // A shaped line's packaged faces, verified by their pinned digests.
            loadComplexFont: async (sourceId: string) => {
              const buffer = await fetchVerifiedPdfComplexScriptFont(sourceId);
              return buffer ? new Uint8Array(buffer) : null;
            },
          });
        }
        if (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION) {
          const payload = parsed as ParsedFontExtensionHistoryPayload;
          const authority = options.fontExtensionHistoryAuthority;
          if (
            !authority ||
            authority.direction !== payload.direction ||
            authority.originRecord.id !== payload.originTransactionId
          ) {
            throw new Error(
              "Native font extension history authority is unavailable.",
            );
          }
          fontExtensionHistoryPlan =
            await planPdfNativeFontExtensionTextHistory({
              sourceBytes,
              source: context.source,
              originRecord: authority.originRecord,
              direction: authority.direction,
              transactionId: command.transactionId,
              baseSessionRevision: command.sessionRevision,
            });
        }
        if (isTextWindowHistory) {
          const payload = parsed as ParsedFontExtensionHistoryPayload;
          const authority = options.textWindowHistoryAuthority;
          if (!authority || authority.originRecord.id !== payload.originTransactionId || authority.direction !== payload.direction)
            throw new Error("Native text-window history authority is unavailable.");
          textWindowPlan = await planMraiNativeTextWindowHistory({sourceBytes,
            expectedSourceSha256: context.source.sourceSha256, resourceId: context.source.resourceId,
            originRecord: authority.originRecord, direction: authority.direction});
        }
        const payloadOfList = !isTextWindowHistory && (
          command.action === MRAI_NATIVE_LIST_TEXT_ACTION ||
          command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION)
            ? (parsed as ParsedListPayload)
            : null;
        // A move plans its own proof and rewrite; it carries no text edit.
        if (payloadOfList?.move) {
          if (options.listTextHistoryAuthority || options.listTextPropertyAuthority ||
            options.compiledTextPreview) {
            throw new Error("A native text move carries no other text authority.");
          }
          textMovePlan = await planMraiNativeTextMove({
            sourceBytes,
            selection: payloadOfList.selection,
            move: payloadOfList.move,
          });
          if (!selectionsEqual(textMovePlan.evidence.sourceFidelity.selection!, payloadOfList.selection)) {
            textMovePlan.prepared.candidateBytes.fill(0);
            throw new PdfNativeTransactionError(
              "precondition-mismatch",
              "Native text move does not match its fresh byte authority.",
            );
          }
        }
        // A group move plans its own proof and rewrite, or its restore.
        if (payloadOfList?.group) {
          if (options.listTextHistoryAuthority || options.listTextPropertyAuthority ||
            options.compiledTextPreview) {
            throw new Error("A native group move carries no other text authority.");
          }
          const group = payloadOfList.group;
          if (group.transform) {
            const planned = await planMraiNativeTextGroupMove({
              sourceBytes,
              selection: payloadOfList.selection,
              transform: group.transform,
              ...(typeof options.textGroupMoveFrameRight === "number"
                ? { frameRight: options.textGroupMoveFrameRight } : {}),
            });
            textGroupMovePlan = planned;
            if (!selectionsEqual(planned.proofSelection, payloadOfList.selection)) {
              planned.prepared.candidateBytes.fill(0);
              throw new PdfNativeTransactionError(
                "precondition-mismatch",
                "Native group move does not match its fresh byte authority.",
              );
            }
          } else {
            const authority = options.textGroupMoveRestoreAuthority;
            if (!authority || authority.originTransactionId !== group.restore!.originTransactionId ||
              authority.originTransactionId === command.transactionId) {
              throw new Error("Native group move restore authority is unavailable.");
            }
            textGroupMovePlan = await planMraiNativeTextGroupMoveRestore({
              sourceBytes,
              origin: authority.origin,
              originTransactionId: authority.originTransactionId,
              originalBytes: authority.originalBytes,
              rewrittenBytes: authority.rewrittenBytes,
            });
          }
        }
        const listPayload = payloadOfList?.move || payloadOfList?.group ? null : payloadOfList;
        const embeddedAuthority = options.listTextPropertyAuthority;
        const embeddedDraft = embeddedAuthority !== undefined && (
          isPdfNativeEmbeddedFontKey(embeddedAuthority.properties.fontResourceName) ||
          (embeddedAuthority.properties.segments ?? []).some((segment) => isPdfNativeEmbeddedFontKey(segment.fontResourceName)));
        const paragraphFormat = options.listTextParagraphFormat;
        // The editor's rows, where it sent them: whole offsets, in order.
        const rowStarts = Array.isArray(options.listTextRowStarts) && options.listTextRowStarts.length > 0 &&
          options.listTextRowStarts.length <= 128 &&
          options.listTextRowStarts.every((start, index, all) => Number.isSafeInteger(start) && start > (all[index - 1] ?? 0))
          ? options.listTextRowStarts : null;
        // A format naming no family the PDF lacks rides the font-extension
        // action for the text's own sake: characters no font of the
        // paragraph has, set in the substitute face with the formats.
        const formatEmbedsFonts = paragraphFormat !== undefined && (
          (paragraphFormat.embeddedFontPrograms?.length ?? 0) > 0 ||
          [paragraphFormat.format.fontResourceName, ...(paragraphFormat.format.parts ?? []).map((part) => part.fontResourceName)]
            .some((name) => name !== undefined && isPdfNativeEmbeddedFontKey(name)));
        if (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION && paragraphFormat && listPayload &&
          !formatEmbedsFonts) {
          if (!options.loadNativeFontExtensionFont) throw new Error("Native font extension authority is unavailable.");
          const loadedFont = await options.loadNativeFontExtensionFont();
          fontExtensionFontBytes = loadedFont instanceof Uint8Array ? Uint8Array.from(loadedFont) : new Uint8Array(loadedFont.slice(0));
          fontExtensionPlan = await planPdfNativeFontExtensionTextEdit({
            sourceBytes,
            source: {
              resourceId: context.source.resourceId,
              resourceVersionId: context.source.resourceVersionId,
            },
            selection: listPayload.selection,
            replacementText: listPayload.replacementText,
            fontBytes: fontExtensionFontBytes,
            transactionId: command.transactionId,
            baseSessionRevision: command.sessionRevision,
            paragraphFormat: { format: paragraphFormat.format,
              ...(paragraphFormat.fontOptions ? { fontOptions: paragraphFormat.fontOptions } : {}) },
            ...(options.listTextParagraphFrame ? { paragraphFrame: options.listTextParagraphFrame } : {}),
            ...(rowStarts ? { rowStarts } : {}),
          });
        } else if (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION && paragraphFormat && listPayload) {
          fontExtensionPlan = await planPdfNativeEmbeddedFontParagraphFormatEdit({
            sourceBytes,
            source: {
              resourceId: context.source.resourceId,
              resourceVersionId: context.source.resourceVersionId,
            },
            selection: listPayload.selection,
            replacementText: listPayload.replacementText,
            format: paragraphFormat.format,
            ...(paragraphFormat.fontOptions ? { fontOptions: paragraphFormat.fontOptions } : {}),
            library: await pdfNativeEmbeddedFontLibraryFromPrograms(paragraphFormat.embeddedFontPrograms ?? [],
              paragraphFormat.embeddedFontManifest ?? null),
            transactionId: command.transactionId,
            baseSessionRevision: command.sessionRevision,
            ...(paragraphFormat.refuseOverOtherText ? { refuseOverOtherText: true } : {}),
          });
        } else if (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION && embeddedDraft && listPayload) {
          // A property edit setting text in a family the PDF does not hold:
          // the faces are installed from the slices the editor handed over,
          // each checked against the lock by its digest.
          fontExtensionPlan = await planPdfNativeEmbeddedFontPropertyEdit({
            sourceBytes,
            source: {
              resourceId: context.source.resourceId,
              resourceVersionId: context.source.resourceVersionId,
            },
            selection: listPayload.selection,
            replacementText: listPayload.replacementText,
            proof: embeddedAuthority.proof,
            properties: embeddedAuthority.properties,
            ...(embeddedAuthority.keepTextRules?.length ? { keepTextRules: embeddedAuthority.keepTextRules } : {}),
            ...(embeddedAuthority.previousTurn ? { previousTurn: embeddedAuthority.previousTurn } : {}),
            library: await pdfNativeEmbeddedFontLibraryFromPrograms(embeddedAuthority.embeddedFontPrograms ?? [],
              embeddedAuthority.embeddedFontManifest ?? null),
            transactionId: command.transactionId,
            baseSessionRevision: command.sessionRevision,
          });
        } else if (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION) {
          if (!options.loadNativeFontExtensionFont || !listPayload) {
            throw new Error("Native font extension authority is unavailable.");
          }
          const loadedFont = await options.loadNativeFontExtensionFont();
          fontExtensionFontBytes =
            loadedFont instanceof Uint8Array
              ? Uint8Array.from(loadedFont)
              : new Uint8Array(loadedFont.slice(0));
          fontExtensionPlan = await planPdfNativeFontExtensionTextEdit({
            sourceBytes,
            source: {
              resourceId: context.source.resourceId,
              resourceVersionId: context.source.resourceVersionId,
            },
            selection: listPayload.selection,
            replacementText: listPayload.replacementText,
            fontBytes: fontExtensionFontBytes,
            transactionId: command.transactionId,
            baseSessionRevision: command.sessionRevision,
            ...(options.listTextParagraphFrame ? { paragraphFrame: options.listTextParagraphFrame } : {}),
            ...(rowStarts ? { rowStarts } : {}),
          });
        }
        const historyAuthority = options.listTextHistoryAuthority;
        const history =
          command.action === MRAI_NATIVE_LIST_TEXT_ACTION &&
          listPayload &&
          historyAuthority
            ? await planMraiNativeListTextHistoryReplacement({
                sourceBytes,
                expectedSourceSha256: historyAuthority.expectedSourceSha256,
                selectionRevision: historyAuthority.selectionRevision,
                priorFidelity: historyAuthority.priorFidelity,
                positionedOriginTransactionId: historyAuthority.positionedOriginTransactionId,
                expectedCurrentCodeBytes:
                  historyAuthority.expectedCurrentCodeBytes,
                desiredCodeBytes: historyAuthority.desiredCodeBytes,
                expectedCurrentOperatorBytes:
                  historyAuthority.expectedCurrentOperatorBytes,
                desiredOperatorBytes: historyAuthority.desiredOperatorBytes,
              })
            : null;
        const propertyAuthority = options.listTextPropertyAuthority;
        const property =
          command.action === MRAI_NATIVE_LIST_TEXT_ACTION &&
          listPayload &&
          propertyAuthority &&
          !history
            ? await planMraiNativeListTextPropertyReplacement({
                sourceBytes,
                selection: listPayload.selection,
                replacementText: listPayload.replacementText,
                proof: propertyAuthority.proof,
                properties: propertyAuthority.properties,
                ...(propertyAuthority.keepTextRules?.length ? { keepTextRules: propertyAuthority.keepTextRules } : {}),
                ...(propertyAuthority.previousTurn ? { previousTurn: propertyAuthority.previousTurn } : {}),
              })
            : null;
        // A property edit asked in the reader's own spelling of its run (one
        // space for each run of the run's white-space glyphs, a space at each
        // wide kern) writes the run's own codes back, so the text it records
        // is the source's; the reader still reads the item the plan proved
        // for the run after the edit.
        const propertyReaderSpelling =
          property &&
          JSON.stringify(property.evidence.replacementCodeBytes) ===
            JSON.stringify(property.evidence.originalCodeBytes)
            ? property.evidence.propertyMutation?.replacementSelection.item.str
            : undefined;
        if (
          (history || property) &&
          (!selectionsEqual(
            history?.selection ?? property!.evidence.selection!,
            listPayload!.selection,
          ) ||
            (listPayload!.replacementText !==
              (history?.plan.evidence.replacementDisplayText ??
                property!.evidence.replacementDisplayText) &&
              listPayload!.replacementText !== propertyReaderSpelling))
        ) {
          throw new PdfNativeTransactionError(
            "precondition-mismatch",
            "Native text history command does not match its fresh byte authority.",
          );
        }
        // A resized paragraph frame rides a plain text edit only, and its plan
        // is the paragraph set again in that frame (no compiled preview of
        // the text alone stands for it).
        const paragraphFrame = options.listTextParagraphFrame;
        if (paragraphFrame && (history || property ||
          (paragraphFrame.right === undefined && paragraphFrame.lineStep === undefined &&
            paragraphFrame.sizeScale === undefined) ||
          (paragraphFrame.sizeScale !== undefined && !Number.isFinite(paragraphFrame.sizeScale)) ||
          (paragraphFrame.right !== undefined && !Number.isFinite(paragraphFrame.right)) ||
          (paragraphFrame.lineStep !== undefined && !Number.isFinite(paragraphFrame.lineStep)))) {
          throw new Error("A resized paragraph frame rides a plain text edit only.");
        }
        const cachedText = command.action === MRAI_NATIVE_LIST_TEXT_ACTION && listPayload && !history && !property &&
          !paragraphFrame && !paragraphFormat
          ? readPdfCompiledTextPreview(options.compiledTextPreview,listPayload.selection,listPayload.replacementText) : null;
        reusedCandidateBytes = cachedText?.prepared.candidateBytes;
        const request =
          command.action === MRAI_NATIVE_LIST_TEXT_ACTION && listPayload
          ? (history?.plan ??
            property ??
            cachedText?.plan ??
            (paragraphFormat && !paragraphFrame
              ? pdfNativeParagraphPlanClearOfOtherText(await planMraiNativeListTextParagraphFormat({
                sourceBytes,
                selection: listPayload.selection,
                replacementText: listPayload.replacementText,
                format: paragraphFormat.format,
                ...(paragraphFormat.fontOptions ? { fontOptions: paragraphFormat.fontOptions } : {}),
              }), paragraphFormat.refuseOverOtherText === true)
              : null) ??
            (await planMraiNativeListTextSelectionReplacement({
              sourceBytes,
              selection: listPayload.selection,
              replacementText: listPayload.replacementText,
              ...(paragraphFrame?.right !== undefined ? { paragraphFrameRight: paragraphFrame.right } : {}),
              ...(paragraphFrame?.lineStep !== undefined ? { paragraphLineStep: paragraphFrame.lineStep } : {}),
              ...(paragraphFrame?.sizeScale !== undefined ? { paragraphSizeScale: paragraphFrame.sizeScale } : {}),
              ...(rowStarts && !paragraphFrame ? { rowStarts } : {}),
            })))
          : null;
        const imagePayload =
          command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION
            ? (parsed as ParsedEmbeddedImagePayload)
            : null;
        const imageHistoryAuthority = options.embeddedImageHistoryAuthority;
        const imagePlan = imagePayload
          ? imageHistoryAuthority
            ? await planMraiNativeEmbeddedImageRemovalHistory({
                sourceBytes,
                expectedSourceSha256:
                  imageHistoryAuthority.expectedSourceSha256,
                priorEvidence: imageHistoryAuthority.priorEvidence,
                expectedCurrentOperatorBytes:
                  imageHistoryAuthority.expectedCurrentOperatorBytes,
                desiredOperatorBytes:
                  imageHistoryAuthority.desiredOperatorBytes,
              })
            : await planMraiNativeEmbeddedImageRemoval({
                sourceBytes,
                selection: imagePayload,
              })
          : null;
        if (
          imagePlan &&
          imagePayload &&
          (imagePlan.evidence.pageNumber !== imagePayload.pageNumber ||
            imagePlan.evidence.imageOrder !== imagePayload.imageOrder)
        ) {
          throw new PdfNativeTransactionError(
            "precondition-mismatch",
            "Embedded image command does not match its fresh byte authority.",
          );
        }
        listFidelity = request?.evidence;
        embeddedImageRemoval = imagePlan?.evidence;
        const imageTransformPayload =
          command.action === MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION
            ? (parsed as ParsedEmbeddedImageTransformPayload)
            : null;
        const imageTransformHistory = options.embeddedImageTransformHistoryAuthority;
        const imageTransformPlan = imageTransformPayload
          ? imageTransformHistory
            ? await planMraiNativeEmbeddedImageTransformHistory({
                sourceBytes,
                expectedSourceSha256: imageTransformHistory.expectedSourceSha256,
                priorEvidence: imageTransformHistory.priorEvidence,
                expectedCurrentOperatorBytes:
                  imageTransformHistory.expectedCurrentOperatorBytes,
                desiredOperatorBytes: imageTransformHistory.desiredOperatorBytes,
              })
            : imageTransformPayload.intent
              ? await planMraiNativeEmbeddedImageTransform({
                  sourceBytes,
                  selection: imageTransformPayload,
                  intent: imageTransformPayload.intent,
                })
              : (() => {
                  throw new PdfNativeTransactionError(
                    "invalid-request",
                    "Embedded image transform authority is missing.",
                  );
                })()
          : null;
        if (
          imageTransformPlan &&
          imageTransformPayload &&
          (imageTransformPlan.evidence.pageNumber !== imageTransformPayload.pageNumber ||
            imageTransformPlan.evidence.imageOrder !== imageTransformPayload.imageOrder)
        ) {
          throw new PdfNativeTransactionError(
            "precondition-mismatch",
            "Embedded image transform does not match its fresh byte authority.",
          );
        }
        embeddedImageTransform = imageTransformPlan?.evidence;
        const imageReplacementPayload =
          command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION
            ? (parsed as ParsedEmbeddedImageReplacementPayload)
            : null;
        const imageReplacementHistory =
          options.embeddedImageReplacementHistoryAuthority;
        if (imageReplacementPayload) {
          embeddedImageReplacement = imageReplacementHistory
            ? await planMraiNativeEmbeddedImageReplacementHistory({
                sourceBytes,
                expectedSourceSha256:
                  imageReplacementHistory.expectedSourceSha256,
                priorEvidence: imageReplacementHistory.priorEvidence,
                expectedCurrent: imageReplacementHistory.expectedCurrent,
                desired: imageReplacementHistory.desired,
              })
            : imageReplacementPayload.replacementBytes &&
                imageReplacementPayload.replacementMimeType
              ? await planMraiNativeEmbeddedImageReplacement({
                  sourceBytes,
                  selection: imageReplacementPayload,
                  replacementBytes: imageReplacementPayload.replacementBytes,
                  replacementMimeType:
                    imageReplacementPayload.replacementMimeType,
                })
              : (() => {
                  throw new PdfNativeTransactionError(
                    "invalid-request",
                    "Embedded image replacement bytes are missing.",
                  );
                })();
          if (
            embeddedImageReplacement.evidence.pageNumber !==
              imageReplacementPayload.pageNumber ||
            embeddedImageReplacement.evidence.imageOrder !==
              imageReplacementPayload.imageOrder
          ) {
            throw new PdfNativeTransactionError(
              "precondition-mismatch",
              "Embedded image replacement does not match its fresh authority.",
            );
          }
          await verifyMraiNativeEmbeddedImageReplacementCandidate(
            sourceBytes,
            embeddedImageReplacement,
          );
        }
        if (!outlinePrepared && !pageSplitPrepared && !pageBackgroundPrepared && !embeddedImageReplacement) {
          if (fontExtensionHistoryPlan) {
            // Its exact two-object candidate was already reopened by the
            // dedicated planner; it deliberately does not enter the generic
            // content-stream request surface.
          } else if (fontExtensionPlan) {
            prepared = fontExtensionPlan.prepared;
          } else if (paragraphPlan) {
            prepared = await prepareMraiNativeContentStreamTransaction(sourceBytes, paragraphPlan.request);
            await verifyMraiNativeParagraphProperties(sourceBytes, paragraphPlan, prepared);
          } else if (
            vectorPlan || vectorDeletePlan || vectorStylePlan ||
            vectorDuplicatePlan || vectorReorderPlan ||
            embeddedImageReorderPlan || textObjectDeletePlan ||
            textDecorationPlan || textMovePlan || textGroupMovePlan
          ) {
            prepared = (vectorPlan ?? vectorDeletePlan ?? vectorStylePlan ??
              vectorDuplicatePlan ?? vectorReorderPlan ??
              embeddedImageReorderPlan ?? textObjectDeletePlan ??
              textDecorationPlan ?? textMovePlan ?? textGroupMovePlan)!
              .prepared;
          } else if (contentAppendPlan) {
            prepared = contentAppendPlan.prepared;
          } else {
            prepared = cachedText?.prepared ?? await prepareMraiNativeContentStreamTransaction(
              sourceBytes,
              textWindowPlan?.request ?? request?.request ??
                imagePlan?.request ?? imageTransformPlan?.request ?? {
                  identity: {
                    sourceSha256: context.source.sourceSha256,
                    objectNumber: (parsed as ParsedNativePayload).objectNumber,
                    generation: (parsed as ParsedNativePayload).generation,
                  },
                  ...(parsed as ParsedNativePayload).request,
                },
            );
            await verifyMraiNativeContentStreamCandidate(sourceBytes, prepared);
          }
        }
      } catch (error) {
        reusedCandidateBytes?.fill(0);
        clearFontExtensionBytes();
        if (error instanceof PdfNativeTransactionError) {
          return mapNativeError(error, command, context);
        }
        return refusedPdfEngineResult(
          "adapter-rejected",
          error instanceof Error
            ? error.message
            : "Native PDF preparation failed.",
          command.transactionId,
        );
      }

          const record: PdfSharedStateRecord = { ...(textWindowPlan ? stableTextWindowHistoryRecord(command, prepared!, textWindowPlan.history) : paragraphPlan ? {
            ...stableRecord(command, prepared!),
            schemaVersion: MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
            kind: "paragraph-properties",
            paragraphProperties: JSON.parse(JSON.stringify({
              schemaVersion: MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA,
              sourceFidelity: paragraphPlan.sourceProof.evidence,
              changes: paragraphPlan.changes,
              ...(paragraphPlan.textPlan ? {
                textReplacement: paragraphPlan.textPlan.evidence,
                textReplacementBytes: Array.from(paragraphPlan.textPlan.request.replacement),
              } : {}),
              ...(paragraphPlan.textHistory && options.paragraphHistoryOrigin ? {
                historyOriginTransactionId: options.paragraphHistoryOrigin.id,
              } : {}),
            })) as JsonValue,
          } : outlinePrepared
            ? stableOutlineRecord(command, outlinePrepared)
            : pageSplitPrepared
            ? stablePageSplitRecord(command, pageSplitPrepared)
            : pageBackgroundPrepared
            ? stablePageBackgroundRecord(command, pageBackgroundPrepared)
            : fontExtensionHistoryPlan
              ? fontExtensionHistoryPlan.record
              : fontExtensionPlan
                ? fontExtensionPlan.record
                : embeddedImageReplacement
                  ? stableEmbeddedImageReplacementRecord(
                      command,
                      embeddedImageReplacement,
                    )
                  : vectorPlan
                    ? stableVectorRecord(command, vectorPlan)
                    : vectorDeletePlan
                      ? stableVectorDeleteRecord(command, vectorDeletePlan)
                      : vectorDuplicatePlan
                        ? stableVectorDuplicateRecord(command, vectorDuplicatePlan)
                      : vectorReorderPlan
                        ? stableVectorReorderRecord(command, vectorReorderPlan)
                      : embeddedImageReorderPlan
                        ? stableEmbeddedImageReorderRecord(
                            command,
                            embeddedImageReorderPlan,
                          )
                      : textObjectDeletePlan
                        ? stableTextObjectDeleteRecord(
                            command,
                            textObjectDeletePlan,
                          )
                      : textDecorationPlan
                        ? stableTextDecorationRecord(command, textDecorationPlan)
                      : textMovePlan
                        ? stableMraiNativeTextMoveRecord(command, textMovePlan)
                      : textGroupMovePlan
                        ? stableMraiNativeTextGroupMoveRecord(command, textGroupMovePlan)
                      : vectorStylePlan
                        ? stableVectorStyleRecord(command, vectorStylePlan)
                        : contentAppendPlan
                          ? stableContentAppendRecord(command, contentAppendPlan)
                          : stableRecord(
                              command,
                              prepared!,
                              listFidelity,
                              embeddedImageRemoval,
                              embeddedImageTransform,
                            )), ...(checkpointClaim ? {sourceCheckpoint:checkpointClaim} : {}) };
      const built =
        outlinePrepared ?? pageSplitPrepared ?? pageBackgroundPrepared ?? embeddedImageReplacement ??
        fontExtensionHistoryPlan?.prepared ?? prepared;
      const candidate = built
        ? Object.freeze({ bytes: built.candidateBytes, sha256: built.candidateSha256, sourceLength: built.sourceLength })
        : null;
      const displayCandidate =
        prepared && !outlinePrepared && !pageSplitPrepared && !pageBackgroundPrepared && !embeddedImageReplacement
          ? Object.freeze({ bytes: prepared.candidateBytes, sha256: prepared.candidateSha256, sourceLength: prepared.sourceLength })
          : null;
      const expectedCandidateSha256 =
        outlinePrepared?.candidateSha256 ??
        pageSplitPrepared?.candidateSha256 ??
        pageBackgroundPrepared?.candidateSha256 ??
        embeddedImageReplacement?.candidateSha256 ??
        fontExtensionHistoryPlan?.prepared.candidateSha256 ??
        prepared!.candidateSha256;
      return Object.freeze({
        status: "planned" as const,
        record,
        displayCandidate,
        candidate,
        expectedCandidateSha256,
        durable: command.action !== MRAI_NATIVE_CONTENT_STREAM_ACTION,
        dispose() {
          reusedCandidateBytes?.fill(0);
          vectorPlan?.prepared.candidateBytes.fill(0);
          vectorDeletePlan?.prepared.candidateBytes.fill(0);
          vectorDuplicatePlan?.prepared.candidateBytes.fill(0);
          vectorReorderPlan?.prepared.candidateBytes.fill(0);
          embeddedImageReorderPlan?.prepared.candidateBytes.fill(0);
          textObjectDeletePlan?.prepared.candidateBytes.fill(0);
          textDecorationPlan?.prepared.candidateBytes.fill(0);
          textMovePlan?.prepared.candidateBytes.fill(0);
          textGroupMovePlan?.prepared.candidateBytes.fill(0);
          vectorStylePlan?.prepared.candidateBytes.fill(0);
          contentAppendPlan?.prepared.candidateBytes.fill(0);
          clearFontExtensionBytes();
          if (
            vectorPlan || vectorDeletePlan || vectorStylePlan || contentAppendPlan ||
            vectorDuplicatePlan || vectorReorderPlan ||
            embeddedImageReorderPlan || textObjectDeletePlan || textDecorationPlan ||
            textMovePlan || textGroupMovePlan
          ) {
            sourceBytes.fill(0);
          }
        },
      });
}

export function createMraiNativePdfAdapter(
  options: MraiNativePdfAdapterOptions,
): PdfEngineAdapter {
  let openedSource: PdfEngineSourceIdentity | null = null;

  return Object.freeze({
    authority: "mrai-native" as const,

    availability: () =>
      globalThis.crypto?.subtle
        ? Object.freeze({ available: true })
        : Object.freeze({
            available: false,
            reason: "MRAI native PDF SHA-256 support is unavailable.",
          }),

    canHandle: (command: PdfEngineCommand) =>
      isRecord(command.payload) &&
      ((command.action === MRAI_NATIVE_CONTENT_STREAM_ACTION &&
        command.payload.schemaVersion ===
          MRAI_NATIVE_CONTENT_STREAM_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION &&
          command.payload.schemaVersion === MRAI_NATIVE_PARAGRAPH_PROPERTIES_SCHEMA) ||
        (command.action === MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_VECTOR_SELECTION_EDIT_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_VECTOR_DELETE_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_VECTOR_DELETE_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_VECTOR_DUPLICATE_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_VECTOR_DUPLICATE_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_VECTOR_REORDER_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_VECTOR_REORDER_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_CONTENT_APPEND_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_CONTENT_APPEND_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_LIST_TEXT_ACTION &&
          (command.payload.schemaVersion ===
            MRAI_NATIVE_LIST_TEXT_SCHEMA_VERSION ||
            command.payload.schemaVersion === MRAI_NATIVE_TEXT_MOVE_SCHEMA_VERSION ||
            // A group move is planned by this adapter (parseListPayload); it
            // was missing here, so every group move was refused as "cannot
            // handle this command" before it reached the planner.
            command.payload.schemaVersion === MRAI_NATIVE_TEXT_GROUP_MOVE_SCHEMA_VERSION ||
            command.payload.schemaVersion === MRAI_NATIVE_TEXT_WINDOW_HISTORY_SCHEMA_VERSION ||
            command.payload.schemaVersion ===
              MRAI_NATIVE_MULTI_RUN_TEXT_SCHEMA_VERSION ||
            command.payload.schemaVersion ===
              MRAI_NATIVE_TEXT_PROPERTIES_SCHEMA_VERSION ||
            command.payload.schemaVersion ===
              MRAI_NATIVE_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
            command.payload.schemaVersion ===
              MRAI_NATIVE_MULTICOLUMN_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
            command.payload.schemaVersion ===
              MRAI_NATIVE_ORTHOGONAL_PARAGRAPH_REFLOW_SCHEMA_VERSION ||
            command.payload.schemaVersion ===
              MRAI_NATIVE_AFFINE_PARAGRAPH_REFLOW_SCHEMA_VERSION)) ||
        (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_FONT_EXTENSION_TEXT_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_TEXT_OBJECT_DELETE_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_TEXT_DECORATION_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_TEXT_DECORATION_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_OUTLINE_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_OUTLINE_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_PAGE_SPLIT_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_PAGE_SPLIT_SCHEMA_VERSION) ||
        (command.action === MRAI_NATIVE_PAGE_BACKGROUND_ACTION &&
          command.payload.schemaVersion ===
            MRAI_NATIVE_PAGE_BACKGROUND_SCHEMA_VERSION)),

    open(source: PdfEngineSourceIdentity) {
      openedSource = Object.freeze({ ...source });
    },

    async execute(command: PdfEngineCommand, context: PdfEngineAdapterContext) {
      const writableBeforeLoad = context.readWritable();
      if (!writableBeforeLoad.writable) {
        return refusedPdfEngineResult(
          "not-writable",
          writableBeforeLoad.reason ?? "The PDF page is read-only.",
          command.transactionId,
        );
      }
      if (!openedSource || !sourceMatches(openedSource, context.source)) {
        return refusedPdfEngineResult(
          "adapter-rejected",
          "MRAI native adapter is not pinned to this PDF source.",
          command.transactionId,
        );
      }
      if (!sourceMatches(command, context.source)) {
        return conflict(command, context, "source-mismatch");
      }
      if (
        command.action !== MRAI_NATIVE_CONTENT_STREAM_ACTION &&
        command.action !== MRAI_NATIVE_PARAGRAPH_PROPERTIES_ACTION &&
        command.action !== MRAI_NATIVE_LIST_TEXT_ACTION &&
        command.action !== MRAI_NATIVE_FONT_EXTENSION_TEXT_ACTION &&
        command.action !== MRAI_NATIVE_FONT_EXTENSION_TEXT_HISTORY_ACTION &&
        command.action !== MRAI_NATIVE_EMBEDDED_IMAGE_REMOVE_ACTION &&
        command.action !== MRAI_NATIVE_EMBEDDED_IMAGE_REPLACE_ACTION &&
        command.action !== MRAI_NATIVE_EMBEDDED_IMAGE_TRANSFORM_ACTION &&
        command.action !== MRAI_NATIVE_EMBEDDED_IMAGE_REORDER_ACTION &&
        command.action !== MRAI_NATIVE_TEXT_OBJECT_DELETE_ACTION &&
        command.action !== MRAI_NATIVE_TEXT_DECORATION_ACTION &&
        command.action !== MRAI_NATIVE_OUTLINE_ACTION &&
        command.action !== MRAI_NATIVE_VECTOR_SELECTION_EDIT_ACTION &&
        command.action !== MRAI_NATIVE_VECTOR_DELETE_ACTION &&
        command.action !== MRAI_NATIVE_VECTOR_DUPLICATE_ACTION &&
        command.action !== MRAI_NATIVE_VECTOR_REORDER_ACTION &&
        command.action !== MRAI_NATIVE_VECTOR_STYLE_SELECTION_EDIT_ACTION &&
        command.action !== MRAI_NATIVE_CONTENT_APPEND_ACTION
      ) {
        return refusedPdfEngineResult(
          "invalid-command",
          "MRAI native adapter does not own this command.",
          command.transactionId,
        );
      }

      const outcome = await planMraiNativeCommand(command, context, options, () => options.loadPinnedSource(context.source));
      if (outcome.status !== "planned") return outcome;
      try {
        const writableBeforePersistence = context.readWritable();
        if (!writableBeforePersistence.writable) {
          return refusedPdfEngineResult(
            "not-writable",
            writableBeforePersistence.reason ??
              "The PDF page became read-only before mutation.",
            command.transactionId,
          );
        }
        if (options.onCandidatePrepared && outcome.displayCandidate) {
          try {
            options.onCandidatePrepared(
              Object.freeze({
                bytes: Uint8Array.from(outcome.displayCandidate.bytes),
                sha256: outcome.displayCandidate.sha256,
              }),
            );
          } catch {
            // A display hint never stops the write.
          }
        }
        return await commitMraiNativePlannedRecord({
          command,
          context,
          record: outcome.record,
          expectedCandidateSha256: outcome.expectedCandidateSha256,
          durable: outcome.durable,
          sharedStateBridge: options.sharedStateBridge,
          collection: options.collection,
        });
      } finally {
        outcome.dispose();
      }
    },

    reload(source: PdfEngineSourceIdentity) {
      openedSource = Object.freeze({ ...source });
    },

    dispose() {
      openedSource = null;
    },
  });
}
