/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFInvalidObject,
  PDFName,
  PDFNull,
  type PDFObject,
  PDFRawStream,
  PDFRef,
  PDFStream,
} from "pdf-lib";

import {
  opaqueObjectDictionary,
  signatureMutationImpact,
} from "./signature-detect";
import { loadPdfForInspection } from "./encrypted-guard";
import { readPdfNativeReceiptReplayPolicy, type PdfNativeReceiptReplayPolicy } from "./native-receipt-journal";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";
import { inspectPdfRiskProfileOffThread, warmPdfSourceRiskWorker } from "./source-risk-offthread";

export type PdfSavePolicyHost = "reader" | "native-editor";
export type PdfSavePolicyRevision = string | number;
export type PdfSavePolicyDisposition =
  | "native"
  | "read-only"
  | "flatten-with-consent"
  | "blocked";

export type PdfSavePolicyAction =
  | "save-current"
  | "download-original"
  | "download-current"
  | "download-native"
  | "download-annotated"
  | "download-fillable"
  | "download-editable-annotations"
  | "download-archival"
  | "download-properties"
  | "print"
  | "replace-source"
  | "pdf.page.import-native"
  | "pdf.page.replace-native"
  | "pdf.page.insert-blank"
  | "pdf.page.delete"
  | "pdf.page.reorder"
  | "pdf.page.organize"
  | "pdf.page.rotate"
  | "pdf.page.crop"
  | "pdf.page.split"
  | "pdf.page.background.commit"
  | "pdf.page.resize-media-box"
  | "pdf.page.labels.commit"
  | "pdf.decoration.watermark.commit"
  | "pdf.decoration.page-number.commit"
  | "pdf.decoration.header-footer.commit"
  | "pdf.text.free-text.commit"
  | "pdf.text.edit-content.commit"
  | "pdf.text.delete-object"
  | "pdf.text.decoration.commit"
  | "pdf.text.paragraph.justify"
  | "pdf.tagged-text.commit"
  | "pdf.image.insert.commit"
  | "pdf.image.overlay.commit"
  | "pdf.insert.table.commit"
  | "pdf.insert.chart.commit"
  | "pdf.image.remove-source"
  | "pdf.image.replace-source"
  | "pdf.image.transform-source"
  | "pdf.image.arrange-source"
  | "pdf.outline.commit"
  | "pdf.link.commit"
  | "pdf.link.edit.commit"
  | "pdf.vector.geometry.commit"
  | "pdf.vector.style.commit"
  | "pdf.vector.duplicate.commit"
  | "pdf.vector.reorder.commit"
  | "pdf.content.append.commit"
  | "pdf.content.append-image.commit"
  | "pdf.content.append-path.commit"
  | "pdf.signature.visual.commit"
  | "pdf.comment.highlight.commit"
  | "pdf.comment.area-highlight.commit"
  | "pdf.comment.underline.commit"
  | "pdf.comment.strikeout.commit"
  | "pdf.comment.squiggly.commit"
  | "pdf.comment.proofing.commit"
  | "pdf.comment.note.commit"
  | "pdf.comment.batch-delete"
  | "pdf.comment.ink.commit"
  | "pdf.comment.shape.commit"
  | "pdf.comment.stamp.commit"
  | "pdf.comment.file-attachment.commit"
  | "pdf.comment.measure.commit"
  | "pdf.edit.undo"
  | "pdf.edit.redo"
  | "pdf.page.insert-file"
  | "pdf.page.extract"
  | "pdf.file.split"
  | "pdf.file.merge"
  | "pdf.file.compress-lossless"
  | "pdf.file.compress-lossy"
  | "pdf.file.deskew-page"
  | "pdf.file.image-only"
  | "pdf.form.create-detected-fields"
  | "pdf.form.edit-existing-field"
  | "pdf.form.delete-existing-field"
  | "pdf.form.reset-values"
  | "pdf.form.import-data"
  | "pdf.page.apply-background"
  | "pdf.page.remove-background"
  | "pdf.image.apply-overlay-edits"
  | "pdf.redaction.apply";

export type PdfSavePolicyFeatureId =
  | "PDF-STD-A11Y-02"
  | "PDF-FIL-DOC-02"
  | "PDF-FIL-DOC-03"
  | "PDF-FIL-DOC-06"
  | "PDF-FIL-DOC-10"
  | "PDF-INS-PAGE-02"
  | "PDF-INS-DECOR-01"
  | "PDF-INS-DECOR-02"
  | "PDF-INS-DECOR-03"
  | "PDF-INS-CONTENT-01"
  | "PDF-INS-CONTENT-02"
  | "PDF-INS-CONTENT-03"
  | "PDF-INS-CONTENT-04"
  | "PDF-EDT-TEXT-01"
  | "PDF-EDT-IMAGE-01"
  | "PDF-EDT-OBJECT-03"
  | "PDF-EDT-OBJECT-02"
  | "PDF-EDT-OBJECT-01"
  | "PDF-PAG-STRUCT-01"
  | "PDF-PAG-STRUCT-02"
  | "PDF-PAG-STRUCT-03"
  | "PDF-PAG-STRUCT-04"
  | "PDF-PAG-STRUCT-05"
  | "PDF-PAG-STRUCT-07"
  | "PDF-PAG-STRUCT-08"
  | "PDF-PAG-STRUCT-09"
  | "PDF-PAG-STRUCT-11"
  | "PDF-PAG-STRUCT-12"
  | "PDF-TOO-FILE-01"
  | "PDF-TOO-FILE-02"
  | "PDF-TOO-FILE-03"
  | "PDF-TOO-FILE-04"
  | "PDF-TOO-CONV-06"
  | "PDF-INS-DECOR-04"
  | "PDF-EDT-IMAGE-02"
  | "PDF-HOM-NAV-05"
  | "PDF-SEC-REDACT-02"
  | "PDF-FRM-AUTH-02"
  | "PDF-FRM-AUTH-03"
  | "PDF-CMT-MARK-01"
  | "PDF-CMT-MARK-02"
  | "PDF-CMT-MARK-03"
  | "PDF-CMT-MARK-04"
  | "PDF-CMT-MARK-05"
  | "PDF-CMT-MARK-06"
  | "PDF-CMT-NOTE-01"
  | "PDF-CMT-NOTE-03"
  | "PDF-CMT-DRAW-01"
  | "PDF-CMT-DRAW-02"
  | "PDF-CMT-DRAW-04"
  | "PDF-CMT-DRAW-05"
  | "PDF-CMT-DRAW-06";

export type PdfSourceRiskState = "none" | "present" | "unknown";

export interface PdfSourceRiskProfile {
  readonly encryption: PdfSourceRiskState;
  readonly certificateSignatures: PdfSourceRiskState;
  /** A certification (DocMDP) signature: the author's own rule about what may
   * still change, which no edit of ours is part of. An ordinary approval
   * signature is `certificateSignatures` alone: it signs the revision it was
   * made over, which an append-only native edit leaves byte for byte. */
  readonly certifiedSignatures: PdfSourceRiskState;
  readonly xfa: PdfSourceRiskState;
  /** Beside an XFA form only: whether it is drawn from its template when opened
   * (/NeedsRendering). A static one draws its pages and carries its AcroForm. */
  readonly xfaDynamic?: PdfSourceRiskState;
  readonly optionalContent: PdfSourceRiskState;
  readonly attachments: PdfSourceRiskState;
  readonly taggedStructure: PdfSourceRiskState;
  readonly pdfaClaim: PdfSourceRiskState;
  readonly inspectionFailure: PdfSourceRiskState;
}

/** A profile as a caller may write it: a caller that does not distinguish a
 * certification signature from an approval one says nothing, and is read as
 * the stricter of the two. */
export type PdfSourceRiskProfileInput = Omit<
  PdfSourceRiskProfile,
  "certifiedSignatures"
> &
  Partial<Pick<PdfSourceRiskProfile, "certifiedSignatures">>;

export interface PdfSourceIdentity {
  readonly dataItemId: string;
  /** Exact URL or caller-owned key for the bytes that were inspected. */
  readonly byteGenerationKey: string;
  readonly sha256: string;
  readonly documentRevision: PdfSavePolicyRevision;
}

export interface PdfSourceRiskSnapshot {
  readonly identity: PdfSourceIdentity;
  readonly profile: PdfSourceRiskProfile;
  readonly profileDigest: string;
}

export interface PdfPagePackageSourceSnapshot {
  readonly source: PdfSourceRiskSnapshot;
  readonly byteLength: number;
  readonly pageCount: number;
  readonly structureRisk: PdfSourceRiskState;
}

export interface PdfPagePackagePolicyContext {
  readonly sources: readonly PdfPagePackageSourceSnapshot[];
  readonly planDigest: string;
  readonly outputCount: number;
}

/**
 * Immutable preflight for the one-page deskew raster writer. Object references
 * deliberately bind the permit to the exact pdf.js document and rendered
 * canvas; their private identities are never serialized or accepted from a
 * caller as strings.
 */
export interface PdfDeskewRasterPolicyContext {
  readonly dataItemId: string;
  readonly sourceUrl: string;
  readonly previewByteGenerationKey: string;
  readonly previewSha256: string;
  readonly documentRevision: PdfSavePolicyRevision;
  readonly editRevision: PdfSavePolicyRevision;
  readonly pageNumber: number;
  readonly pdfJsDocument: object;
  readonly resolvedRotation: 0 | 90 | 180 | 270;
  readonly canvas: object;
  readonly canvasGeneration: PdfSavePolicyRevision;
  readonly canvasWidth: number;
  readonly canvasHeight: number;
  /** Deterministic output raster, independent of viewer zoom and device DPR. */
  readonly sourceUserUnit: number;
  readonly outputDpi: number;
  readonly outputPixelWidth: number;
  readonly outputPixelHeight: number;
  readonly outputMediaX: number;
  readonly outputMediaY: number;
  readonly outputMediaWidth: number;
  readonly outputMediaHeight: number;
  readonly outputCropX: number;
  readonly outputCropY: number;
  readonly outputCropWidth: number;
  readonly outputCropHeight: number;
  readonly angle: number;
  readonly confidence: number;
  readonly sourceByteLength: number;
  readonly planDigest: string;
}

interface NormalizedPdfDeskewRasterPolicyContext extends PdfDeskewRasterPolicyContext {
  readonly pdfJsDocumentIdentity: number;
  readonly canvasIdentity: number;
}

export interface PdfPageBackgroundPolicyImage {
  readonly byteLength: number;
  readonly sha256: string;
  readonly mimeType: "image/png" | "image/jpeg";
  readonly width: number;
  readonly height: number;
  readonly pixelCount: number;
}

export interface PdfPageBackgroundPolicyContext {
  readonly dataItemId: string;
  readonly sourceUrl: string;
  readonly previewByteGenerationKey: string;
  readonly previewSha256: string;
  readonly documentRevision: PdfSavePolicyRevision;
  readonly editRevision: PdfSavePolicyRevision;
  readonly kind: "apply" | "remove";
  readonly sourceByteLength: number;
  readonly pageCount: number;
  readonly pages: readonly number[];
  readonly markerRefs: readonly string[];
  readonly replaceOrRemoveCount: number;
  readonly color: string | null;
  readonly opacity: number;
  readonly imageFit: "fit" | "stretch";
  readonly image: PdfPageBackgroundPolicyImage | null;
  readonly planDigest: string;
}

export interface PdfEmbeddedImageOverlayPolicyContext {
  readonly dataItemId: string;
  readonly sourceUrl: string;
  readonly previewByteGenerationKey: string;
  readonly previewSha256: string;
  readonly documentRevision: PdfSavePolicyRevision;
  readonly editRevision: PdfSavePolicyRevision;
  readonly sourceByteLength: number;
  readonly pageCount: number;
  readonly targetPages: readonly number[];
  readonly editCount: number;
  readonly replacementCount: number;
  readonly replacementByteLength: number;
  readonly geometriesDigest: string;
  readonly planDigest: string;
}

export interface PdfLossyCompressionPolicyContext {
  readonly previewDataItemId: string;
  readonly previewByteGenerationKey: string;
  readonly previewSha256: string;
  readonly documentRevision: PdfSavePolicyRevision;
  readonly editRevision: PdfSavePolicyRevision;
  readonly sourceKind: "data_item" | "resource_version";
  readonly workspaceId: string;
  readonly datasetId: string | null;
  readonly dataItemId: string | null;
  readonly resourceId: string | null;
  readonly resourceVersionId: string | null;
  readonly versionKind: string | null;
  readonly expectedPageId: string | null;
  readonly storageCustodyDigest: string;
  readonly sourceSize: number;
  readonly preset: "screen" | "ebook" | "print";
  readonly targetDpi: 72 | 150 | 300;
  readonly jpegQuality: 60 | 75 | 85;
  readonly pageCount: number;
  readonly transformedObjectCount: number;
  readonly nativePlanDigest: string;
  readonly evidenceDigest: string;
}

export interface PdfOutputLossEntry {
  readonly code: string;
  readonly label: string;
  readonly detail?: string;
}

export interface PdfOutputLossReport {
  readonly entries: readonly PdfOutputLossEntry[];
  readonly codes: readonly string[];
  readonly digest: string;
}

export const PDF_LOSS_CURRENT_EDITOR_EDITS_EXCLUDED: PdfOutputLossEntry =
  Object.freeze({
    code: "PDF_LOSS_CURRENT_EDITOR_EDITS_EXCLUDED",
    label:
      "Current editor changes absent from the inspected PDF bytes are not included in this output.",
  });

export const PDF_LOSS_VISUAL_MASK_RECOVERABLE_CONTENT: PdfOutputLossEntry =
  Object.freeze({
    code: "PDF_LOSS_VISUAL_MASK_RECOVERABLE_CONTENT",
    label:
      "Visual covers change appearance only; underlying PDF content remains recoverable.",
    detail:
      "This is not a redaction. Use Redact, then Apply, to permanently remove marked content.",
  });

export interface PdfSavePolicyEditSnapshot {
  readonly editRevision: PdfSavePolicyRevision;
  readonly hasEdits: boolean;
  readonly nativeDurableReceipt: boolean;
  readonly lossReport: PdfOutputLossReport;
}

export interface PdfSavePolicyRequest {
  readonly host: PdfSavePolicyHost;
  readonly action: PdfSavePolicyAction;
  readonly requestedDisposition: Exclude<PdfSavePolicyDisposition, "blocked">;
  readonly documentReady: boolean;
  readonly editable: boolean;
  readonly source: PdfSourceRiskSnapshot | null;
  readonly replacementSource?: PdfSourceRiskSnapshot | null;
  readonly pagePackage?: PdfPagePackagePolicyContext | null;
  readonly deskewRaster?: PdfDeskewRasterPolicyContext | null;
  readonly pageBackground?: PdfPageBackgroundPolicyContext | null;
  readonly embeddedImageOverlay?: PdfEmbeddedImageOverlayPolicyContext | null;
  readonly lossyCompression?: PdfLossyCompressionPolicyContext | null;
  readonly edit: PdfSavePolicyEditSnapshot;
  readonly historyReplay?: PdfNativeReceiptReplayPolicy | null;
  /**
   * Save of edits held while AutoSave was off ("save-current" only): the
   * writes Save sends, each judged by the rule it was judged by when it was
   * made (as a history replay is). A comment held on a PDF whose layers keep
   * it from being rewritten is still a comment its Save may write. Absent or
   * empty, Save is judged as a rewrite of the file.
   */
  readonly heldWrites?: readonly PdfSavePolicyAction[] | null;
}

const decisionBrand: unique symbol = Symbol("PdfSavePolicyDecision");

export type PdfSavePolicyDecision = Readonly<{
  [decisionBrand]: true;
}>;

export interface PdfSavePolicyDecisionView {
  readonly host: PdfSavePolicyHost;
  readonly featureId: PdfSavePolicyFeatureId;
  readonly action: PdfSavePolicyAction;
  readonly enabled: boolean;
  readonly disposition: PdfSavePolicyDisposition;
  readonly reasonKey: string | null;
  readonly requiresConsent: boolean;
  readonly lossReport: PdfOutputLossReport;
  readonly documentRevision: PdfSavePolicyRevision;
  readonly sourceRevision: string;
  readonly editRevision: PdfSavePolicyRevision;
}

interface NormalizedPdfSavePolicyRequest {
  readonly host: PdfSavePolicyHost;
  readonly action: PdfSavePolicyAction;
  readonly featureId: PdfSavePolicyFeatureId;
  readonly requestedDisposition: Exclude<PdfSavePolicyDisposition, "blocked">;
  readonly documentReady: boolean;
  readonly editable: boolean;
  readonly source: PdfSourceRiskSnapshot | null;
  readonly replacementSource: PdfSourceRiskSnapshot | null;
  readonly pagePackage: PdfPagePackagePolicyContext | null;
  readonly deskewRaster: NormalizedPdfDeskewRasterPolicyContext | null;
  readonly pageBackground: PdfPageBackgroundPolicyContext | null;
  readonly embeddedImageOverlay: PdfEmbeddedImageOverlayPolicyContext | null;
  readonly lossyCompression: PdfLossyCompressionPolicyContext | null;
  readonly edit: PdfSavePolicyEditSnapshot;
  readonly historyReplay: ReturnType<typeof readPdfNativeReceiptReplayPolicy>;
  readonly historyReplayRequested: boolean;
  readonly heldWrites: readonly PdfSavePolicyAction[] | null;
  readonly requestKey: string;
}

interface IssuedPdfSavePolicyDecision {
  readonly requestKey: string;
  readonly view: PdfSavePolicyDecisionView;
}

const MAX_SOURCE_BYTES = PDF_EDITABLE_MAX_BYTES;
const MAX_STRING_LENGTH = 512;
const MAX_DATA_ITEM_ID_LENGTH = 160;
const MAX_BYTE_GENERATION_KEY_LENGTH = 4096;
const MAX_LOSS_ENTRIES = 64;
const MAX_LOSS_CODE_LENGTH = 128;
const MAX_LOSS_LABEL_LENGTH = 240;
const MAX_LOSS_DETAIL_LENGTH = 512;
const MAX_RISK_CACHE_ENTRIES = 24;
const MAX_RISK_GRAPH_OBJECTS = 100_000;
const MAX_RISK_GRAPH_EDGES = 5_000_000;
const MAX_PAGE_PACKAGE_SOURCES = 16;
const MAX_PAGE_PACKAGE_AGGREGATE_BYTES = 256 * 1024 * 1024;
const MAX_PAGE_PACKAGE_PAGES = 10_000;
const MAX_PAGE_PACKAGE_OUTPUTS = 200;
const MAX_PAGE_PACKAGE_PLAN_PARTS = 512;
const MAX_PAGE_PACKAGE_PLAN_PART_LENGTH = 4096;
const MAX_DESKEW_CANVAS_DIMENSION = 16_384;
// Raw RGBA canvas memory, not the PDF file: it stays 64 MiB (16,777,216 px,
// the iOS Safari canvas area cap) when the editable PDF total is raised.
const MAX_DESKEW_RASTER_BYTES = 64 * 1024 * 1024;
const DESKEW_OUTPUT_SAFETY_BYTES = 1024 * 1024;
const SHA256_RE = /^[a-f0-9]{64}$/;

/** Fixed print-quality raster target; UI zoom and device DPR never affect it. */
export const PDF_DESKEW_OUTPUT_DPI = 300;

/**
 * The raster pixels a deskew viewport dimension needs. pdf.js scales by
 * 300 / 72 = 4.1666..., so a US Letter page's 792 pt came out
 * 3300.0000000000005 px and rounded up to 3301, never the 3300 the physical
 * crop gives: every Letter page refused with 「本页纠偏失败」 (user audit
 * PDF 0032). Float noise is not a pixel.
 */
export function pdfDeskewRasterPixels(viewportSize: number): number {
  return Math.ceil(viewportSize - 1e-6);
}

export function isPdfDeskewRasterCanvasWithinBudget(
  width: unknown,
  height: unknown,
): boolean {
  if (
    typeof width !== "number" ||
    !Number.isSafeInteger(width) ||
    width < 1 ||
    width > MAX_DESKEW_CANVAS_DIMENSION ||
    typeof height !== "number" ||
    !Number.isSafeInteger(height) ||
    height < 1 ||
    height > MAX_DESKEW_CANVAS_DIMENSION
  ) {
    return false;
  }
  const pixels = width * height;
  const bytes = pixels * 4;
  return (
    Number.isSafeInteger(pixels) &&
    Number.isSafeInteger(bytes) &&
    bytes <= MAX_DESKEW_RASTER_BYTES
  );
}

const RISK_KEYS = [
  "encryption",
  "certificateSignatures",
  "certifiedSignatures",
  "xfa",
  "optionalContent",
  "attachments",
  "taggedStructure",
  "pdfaClaim",
  "inspectionFailure",
] as const satisfies readonly (keyof PdfSourceRiskProfile)[];

const ACTION_FEATURES: Readonly<
  Record<PdfSavePolicyAction, PdfSavePolicyFeatureId>
> = Object.freeze({
  "save-current": "PDF-FIL-DOC-02",
  "download-original": "PDF-FIL-DOC-03",
  "download-current": "PDF-FIL-DOC-03",
  "download-native": "PDF-FIL-DOC-03",
  "download-annotated": "PDF-FIL-DOC-03",
  "download-fillable": "PDF-FIL-DOC-03",
  "download-editable-annotations": "PDF-FIL-DOC-03",
  "download-archival": "PDF-FIL-DOC-03",
  "download-properties": "PDF-FIL-DOC-03",
  print: "PDF-FIL-DOC-06",
  "replace-source": "PDF-FIL-DOC-10",
  "pdf.page.import-native": "PDF-INS-PAGE-02",
  "pdf.page.replace-native": "PDF-PAG-STRUCT-04",
  "pdf.page.insert-blank": "PDF-PAG-STRUCT-01",
  "pdf.page.delete": "PDF-PAG-STRUCT-02",
  "pdf.page.reorder": "PDF-PAG-STRUCT-05",
  "pdf.page.organize": "PDF-PAG-STRUCT-05",
  "pdf.page.rotate": "PDF-PAG-STRUCT-07",
  "pdf.page.crop": "PDF-PAG-STRUCT-08",
  "pdf.page.split": "PDF-PAG-STRUCT-09",
  "pdf.page.background.commit": "PDF-INS-DECOR-04",
  "pdf.page.resize-media-box": "PDF-PAG-STRUCT-11",
  "pdf.page.labels.commit": "PDF-PAG-STRUCT-12",
  "pdf.decoration.watermark.commit": "PDF-INS-DECOR-01",
  "pdf.decoration.page-number.commit": "PDF-INS-DECOR-02",
  "pdf.decoration.header-footer.commit": "PDF-INS-DECOR-03",
  "pdf.text.free-text.commit": "PDF-INS-CONTENT-01",
  "pdf.text.edit-content.commit": "PDF-EDT-TEXT-01",
  "pdf.tagged-text.commit": "PDF-STD-A11Y-02",
  "pdf.image.insert.commit": "PDF-INS-CONTENT-02",
  "pdf.image.overlay.commit": "PDF-EDT-IMAGE-01",
  // 表格: written as a /FreeText (managed_table.py), the text family's.
  "pdf.insert.table.commit": "PDF-INS-CONTENT-01",
  // 图表: written as a /FreeText too (the chart engine's picture).
  "pdf.insert.chart.commit": "PDF-INS-CONTENT-01",
  "pdf.image.remove-source": "PDF-EDT-IMAGE-02",
  "pdf.image.replace-source": "PDF-EDT-IMAGE-02",
  "pdf.text.delete-object": "PDF-EDT-TEXT-01",
  "pdf.text.decoration.commit": "PDF-EDT-TEXT-01",
  "pdf.text.paragraph.justify": "PDF-EDT-TEXT-01",
  "pdf.image.transform-source": "PDF-EDT-IMAGE-02",
  "pdf.image.arrange-source": "PDF-EDT-IMAGE-02",
  "pdf.outline.commit": "PDF-HOM-NAV-05",
  "pdf.link.commit": "PDF-INS-CONTENT-03",
  "pdf.link.edit.commit": "PDF-EDT-OBJECT-02",
  "pdf.vector.geometry.commit": "PDF-EDT-OBJECT-01",
  "pdf.vector.style.commit": "PDF-EDT-OBJECT-01",
  "pdf.vector.duplicate.commit": "PDF-EDT-OBJECT-01",
  "pdf.vector.reorder.commit": "PDF-EDT-OBJECT-01",
  "pdf.content.append.commit": "PDF-INS-CONTENT-01",
  "pdf.content.append-image.commit": "PDF-INS-CONTENT-02",
  "pdf.content.append-path.commit": "PDF-CMT-DRAW-02",
  "pdf.signature.visual.commit": "PDF-INS-CONTENT-04",
  "pdf.comment.highlight.commit": "PDF-CMT-MARK-01",
  "pdf.comment.area-highlight.commit": "PDF-CMT-MARK-02",
  "pdf.comment.underline.commit": "PDF-CMT-MARK-03",
  "pdf.comment.strikeout.commit": "PDF-CMT-MARK-04",
  "pdf.comment.squiggly.commit": "PDF-CMT-MARK-05",
  "pdf.comment.proofing.commit": "PDF-CMT-MARK-06",
  "pdf.comment.note.commit": "PDF-CMT-NOTE-01",
  "pdf.comment.batch-delete": "PDF-CMT-NOTE-03",
  "pdf.comment.ink.commit": "PDF-CMT-DRAW-01",
  "pdf.comment.shape.commit": "PDF-CMT-DRAW-02",
  "pdf.comment.stamp.commit": "PDF-CMT-DRAW-04",
  "pdf.comment.file-attachment.commit": "PDF-CMT-DRAW-05",
  "pdf.comment.measure.commit": "PDF-CMT-DRAW-06",
  "pdf.edit.undo": "PDF-EDT-OBJECT-03",
  "pdf.edit.redo": "PDF-EDT-OBJECT-03",
  "pdf.page.insert-file": "PDF-INS-PAGE-02",
  "pdf.page.extract": "PDF-PAG-STRUCT-03",
  "pdf.file.split": "PDF-TOO-FILE-01",
  "pdf.file.merge": "PDF-TOO-FILE-02",
  "pdf.file.compress-lossless": "PDF-TOO-FILE-03",
  "pdf.file.compress-lossy": "PDF-TOO-FILE-03",
  "pdf.file.deskew-page": "PDF-TOO-FILE-04",
  "pdf.file.image-only": "PDF-TOO-CONV-06",
  "pdf.page.apply-background": "PDF-INS-DECOR-04",
  "pdf.page.remove-background": "PDF-INS-DECOR-04",
  "pdf.image.apply-overlay-edits": "PDF-EDT-IMAGE-02",
  "pdf.redaction.apply": "PDF-SEC-REDACT-02",
  "pdf.form.create-detected-fields": "PDF-FRM-AUTH-02",
  "pdf.form.edit-existing-field": "PDF-FRM-AUTH-02",
  "pdf.form.delete-existing-field": "PDF-FRM-AUTH-02",
  "pdf.form.reset-values": "PDF-FRM-AUTH-03",
  "pdf.form.import-data": "PDF-FRM-AUTH-03",
});

const ACTION_DISPOSITIONS = Object.freeze({
  "save-current": ["native"] as const,
  "download-original": ["read-only"] as const,
  "download-current": ["read-only", "flatten-with-consent"] as const,
  "download-native": ["native"] as const,
  "download-annotated": ["read-only", "flatten-with-consent"] as const,
  "download-fillable": ["read-only", "flatten-with-consent"] as const,
  "download-editable-annotations": [
    "read-only",
    "flatten-with-consent",
  ] as const,
  "download-archival": ["flatten-with-consent"] as const,
  "download-properties": ["flatten-with-consent"] as const,
  print: ["read-only", "flatten-with-consent"] as const,
  "replace-source": ["native"] as const,
  "pdf.page.import-native": ["native"] as const,
  "pdf.page.replace-native": ["native"] as const,
  "pdf.page.insert-blank": ["native"] as const,
  "pdf.page.delete": ["native"] as const,
  "pdf.page.reorder": ["native"] as const,
  "pdf.page.organize": ["native"] as const,
  "pdf.page.rotate": ["native"] as const,
  "pdf.page.crop": ["native"] as const,
  "pdf.page.split": ["native"] as const,
  "pdf.page.background.commit": ["native"] as const,
  "pdf.page.resize-media-box": ["native"] as const,
  "pdf.page.labels.commit": ["native"] as const,
  "pdf.decoration.watermark.commit": ["native"] as const,
  "pdf.decoration.page-number.commit": ["native"] as const,
  "pdf.decoration.header-footer.commit": ["native"] as const,
  "pdf.text.free-text.commit": ["native"] as const,
  "pdf.text.edit-content.commit": ["native"] as const,
  "pdf.tagged-text.commit": ["native"] as const,
  "pdf.image.insert.commit": ["native"] as const,
  "pdf.image.overlay.commit": ["native"] as const,
  "pdf.insert.table.commit": ["native"] as const,
  "pdf.insert.chart.commit": ["native"] as const,
  "pdf.image.remove-source": ["native"] as const,
  "pdf.image.replace-source": ["native"] as const,
  "pdf.text.delete-object": ["native"] as const,
  "pdf.text.decoration.commit": ["native"] as const,
  "pdf.text.paragraph.justify": ["native"] as const,
  "pdf.image.transform-source": ["native"] as const,
  "pdf.image.arrange-source": ["native"] as const,
  "pdf.outline.commit": ["native"] as const,
  "pdf.link.commit": ["native"] as const,
  "pdf.link.edit.commit": ["native"] as const,
  "pdf.vector.geometry.commit": ["native"] as const,
  "pdf.vector.style.commit": ["native"] as const,
  "pdf.vector.duplicate.commit": ["native"] as const,
  "pdf.vector.reorder.commit": ["native"] as const,
  "pdf.content.append.commit": ["native"] as const,
  "pdf.content.append-image.commit": ["native"] as const,
  "pdf.content.append-path.commit": ["native"] as const,
  "pdf.signature.visual.commit": ["native"] as const,
  "pdf.comment.highlight.commit": ["native"] as const,
  "pdf.comment.area-highlight.commit": ["native"] as const,
  "pdf.comment.underline.commit": ["native"] as const,
  "pdf.comment.strikeout.commit": ["native"] as const,
  "pdf.comment.squiggly.commit": ["native"] as const,
  "pdf.comment.proofing.commit": ["native"] as const,
  "pdf.comment.note.commit": ["native"] as const,
  "pdf.comment.batch-delete": ["native"] as const,
  "pdf.comment.ink.commit": ["native"] as const,
  "pdf.comment.shape.commit": ["native"] as const,
  "pdf.comment.stamp.commit": ["native"] as const,
  "pdf.comment.file-attachment.commit": ["native"] as const,
  "pdf.comment.measure.commit": ["native"] as const,
  "pdf.edit.undo": ["native"] as const,
  "pdf.edit.redo": ["native"] as const,
  "pdf.page.insert-file": ["flatten-with-consent"] as const,
  "pdf.page.extract": ["flatten-with-consent"] as const,
  "pdf.file.split": ["flatten-with-consent"] as const,
  "pdf.file.merge": ["flatten-with-consent"] as const,
  "pdf.file.compress-lossless": ["flatten-with-consent"] as const,
  "pdf.file.compress-lossy": ["flatten-with-consent"] as const,
  "pdf.file.deskew-page": ["flatten-with-consent"] as const,
  "pdf.file.image-only": ["flatten-with-consent"] as const,
  "pdf.page.apply-background": ["flatten-with-consent"] as const,
  "pdf.page.remove-background": ["flatten-with-consent"] as const,
  "pdf.image.apply-overlay-edits": ["flatten-with-consent"] as const,
  "pdf.redaction.apply": ["flatten-with-consent"] as const,
  "pdf.form.create-detected-fields": [
    "native",
    "flatten-with-consent",
  ] as const,
  "pdf.form.edit-existing-field": [
    "native",
    "flatten-with-consent",
  ] as const,
  "pdf.form.delete-existing-field": [
    "native",
    "flatten-with-consent",
  ] as const,
  "pdf.form.reset-values": ["native"] as const,
  "pdf.form.import-data": ["native"] as const,
}) satisfies Readonly<
  Record<
    PdfSavePolicyAction,
    readonly Exclude<PdfSavePolicyDisposition, "blocked">[]
  >
>;

const issuedDecisions = new WeakMap<
  PdfSavePolicyDecision,
  IssuedPdfSavePolicyDecision
>();
const riskProfileCache = new Map<
  string,
  Promise<Readonly<PdfSourceRiskProfile>>
>();
const deskewOpaqueIdentities = new WeakMap<object, number>();
let nextDeskewOpaqueIdentity = 1;

function exactString(
  value: unknown,
  label: string,
  maxLength = MAX_STRING_LENGTH,
): string {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > maxLength ||
    value !== value.trim() ||
    value.includes("\0")
  ) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function exactRevision(value: unknown, label: string): PdfSavePolicyRevision {
  if (typeof value === "string") return exactString(value, label);
  if (typeof value === "number" && Number.isSafeInteger(value) && value >= 0) {
    return value;
  }
  throw new Error(`Invalid ${label}`);
}

function riskState(value: unknown, label: string): PdfSourceRiskState {
  if (value === "none" || value === "present" || value === "unknown") {
    return value;
  }
  throw new Error(`Invalid ${label}`);
}

function freezeRiskProfile(
  value: PdfSourceRiskProfileInput,
): PdfSourceRiskProfile {
  return Object.freeze({
    encryption: riskState(value.encryption, "PDF encryption risk"),
    certificateSignatures: riskState(
      value.certificateSignatures,
      "PDF certificate-signature risk",
    ),
    certifiedSignatures: riskState(
      // Not knowing whether a signature certifies is the same as its doing so.
      value.certifiedSignatures ?? value.certificateSignatures,
      "PDF certification-signature risk",
    ),
    xfa: riskState(value.xfa, "PDF XFA risk"),
    ...(value.xfaDynamic === undefined
      ? {}
      : { xfaDynamic: riskState(value.xfaDynamic, "PDF XFA kind") }),
    optionalContent: riskState(
      value.optionalContent,
      "PDF optional-content risk",
    ),
    attachments: riskState(value.attachments, "PDF attachment risk"),
    taggedStructure: riskState(
      value.taggedStructure,
      "PDF tagged-structure risk",
    ),
    pdfaClaim: riskState(value.pdfaClaim, "PDF/A claim risk"),
    inspectionFailure: riskState(
      value.inspectionFailure,
      "PDF inspection-failure risk",
    ),
  });
}

export function createPdfSourceRiskProfile(
  value: PdfSourceRiskProfileInput,
): PdfSourceRiskProfile {
  return freezeRiskProfile(value);
}

export function createUnknownPdfSourceRiskProfile(): PdfSourceRiskProfile {
  return freezeRiskProfile({
    encryption: "unknown",
    certificateSignatures: "unknown",
    certifiedSignatures: "unknown",
    xfa: "unknown",
    optionalContent: "unknown",
    attachments: "unknown",
    taggedStructure: "unknown",
    pdfaClaim: "unknown",
    inspectionFailure: "present",
  });
}

export function createSafePdfSourceRiskProfile(): PdfSourceRiskProfile {
  return freezeRiskProfile({
    encryption: "none",
    certificateSignatures: "none",
    certifiedSignatures: "none",
    xfa: "none",
    optionalContent: "none",
    attachments: "none",
    taggedStructure: "none",
    pdfaClaim: "none",
    inspectionFailure: "none",
  });
}

export function pdfSourceRiskProfileDigest(
  value: PdfSourceRiskProfile,
): string {
  const profile = freezeRiskProfile(value);
  return `pdf-risk-v1:${RISK_KEYS.map((key) => profile[key][0]).join("")}`;
}





const createdSourceRiskSnapshots = new WeakSet<PdfSourceRiskSnapshot>();

export function createPdfSourceRiskSnapshot(options: {
  identity: PdfSourceIdentity;
  profile: PdfSourceRiskProfile;
}): PdfSourceRiskSnapshot {
  const dataItemId = exactString(
    options.identity.dataItemId,
    "PDF data item ID",
    MAX_DATA_ITEM_ID_LENGTH,
  );
  const byteGenerationKey = exactString(
    options.identity.byteGenerationKey,
    "PDF byte-generation key",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const sha256 = exactString(
    options.identity.sha256,
    "PDF source SHA-256",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(sha256)) throw new Error("Invalid PDF source SHA-256");
  const profile = freezeRiskProfile(options.profile);
  const snapshot: PdfSourceRiskSnapshot = Object.freeze({
    identity: Object.freeze({
      dataItemId,
      byteGenerationKey,
      sha256,
      documentRevision: exactRevision(
        options.identity.documentRevision,
        "PDF document revision",
      ),
    }),
    profile,
    profileDigest: pdfSourceRiskProfileDigest(profile),
  });
  createdSourceRiskSnapshots.add(snapshot);
  return snapshot;
}

function exactPositiveSafeInteger(value: unknown, label: string): number {
  if (!Number.isSafeInteger(value) || Number(value) < 1) {
    throw new Error(`Invalid ${label}`);
  }
  return Number(value);
}

function exactDeskewCoordinate(
  value: unknown,
  label: string,
  positive = false,
): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    Math.abs(value) > 1_000_000 ||
    (positive && value <= 0)
  ) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function exactDeskewOpaqueIdentity(value: unknown, label: string): number {
  if (
    value === null ||
    (typeof value !== "object" && typeof value !== "function")
  ) {
    throw new Error(`Invalid ${label}`);
  }
  const object = value as object;
  const existing = deskewOpaqueIdentities.get(object);
  if (existing) return existing;
  if (!Number.isSafeInteger(nextDeskewOpaqueIdentity)) {
    throw new Error("PDF deskew identity budget exhausted");
  }
  const identity = nextDeskewOpaqueIdentity;
  nextDeskewOpaqueIdentity += 1;
  deskewOpaqueIdentities.set(object, identity);
  return identity;
}

function normalizePdfDeskewRasterPolicyContext(
  value: PdfDeskewRasterPolicyContext,
): NormalizedPdfDeskewRasterPolicyContext {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid PDF deskew raster policy context");
  }
  const dataItemId = exactString(
    value.dataItemId,
    "PDF deskew data item ID",
    MAX_DATA_ITEM_ID_LENGTH,
  );
  const sourceUrl = exactString(
    value.sourceUrl,
    "PDF deskew source URL",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewByteGenerationKey = exactString(
    value.previewByteGenerationKey,
    "PDF deskew preview byte generation",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewSha256 = exactString(
    value.previewSha256,
    "PDF deskew preview SHA-256",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(previewSha256)) {
    throw new Error("Invalid PDF deskew preview SHA-256");
  }
  const documentRevision = exactRevision(
    value.documentRevision,
    "PDF deskew document revision",
  );
  const editRevision = exactRevision(
    value.editRevision,
    "PDF deskew edit revision",
  );
  const pageNumber = exactPositiveSafeInteger(
    value.pageNumber,
    "PDF deskew page number",
  );
  const pdfJsDocumentIdentity = exactDeskewOpaqueIdentity(
    value.pdfJsDocument,
    "PDF deskew pdf.js document",
  );
  if (
    value.resolvedRotation !== 0 &&
    value.resolvedRotation !== 90 &&
    value.resolvedRotation !== 180 &&
    value.resolvedRotation !== 270
  ) {
    throw new Error("Invalid PDF deskew resolved rotation");
  }
  const canvasIdentity = exactDeskewOpaqueIdentity(
    value.canvas,
    "PDF deskew canvas",
  );
  const canvasGeneration = exactRevision(
    value.canvasGeneration,
    "PDF deskew canvas generation",
  );
  const canvasWidth = exactPositiveSafeInteger(
    value.canvasWidth,
    "PDF deskew canvas width",
  );
  const canvasHeight = exactPositiveSafeInteger(
    value.canvasHeight,
    "PDF deskew canvas height",
  );
  if (!isPdfDeskewRasterCanvasWithinBudget(canvasWidth, canvasHeight)) {
    throw new Error("PDF deskew canvas budget exceeded");
  }
  const canvasBytes = canvasWidth * canvasHeight * 4;
  const sourceUserUnit = exactDeskewCoordinate(
    value.sourceUserUnit,
    "PDF deskew source UserUnit",
    true,
  );
  if (value.outputDpi !== PDF_DESKEW_OUTPUT_DPI) {
    throw new Error("Invalid PDF deskew output DPI");
  }
  const outputPixelWidth = exactPositiveSafeInteger(
    value.outputPixelWidth,
    "PDF deskew output pixel width",
  );
  const outputPixelHeight = exactPositiveSafeInteger(
    value.outputPixelHeight,
    "PDF deskew output pixel height",
  );
  if (
    !isPdfDeskewRasterCanvasWithinBudget(outputPixelWidth, outputPixelHeight)
  ) {
    throw new Error("PDF deskew output canvas budget exceeded");
  }
  const outputMediaX = exactDeskewCoordinate(
    value.outputMediaX,
    "PDF deskew output media x",
  );
  const outputMediaY = exactDeskewCoordinate(
    value.outputMediaY,
    "PDF deskew output media y",
  );
  const outputMediaWidth = exactDeskewCoordinate(
    value.outputMediaWidth,
    "PDF deskew output media width",
    true,
  );
  const outputMediaHeight = exactDeskewCoordinate(
    value.outputMediaHeight,
    "PDF deskew output media height",
    true,
  );
  const outputCropX = exactDeskewCoordinate(
    value.outputCropX,
    "PDF deskew output crop x",
  );
  const outputCropY = exactDeskewCoordinate(
    value.outputCropY,
    "PDF deskew output crop y",
  );
  const outputCropWidth = exactDeskewCoordinate(
    value.outputCropWidth,
    "PDF deskew output crop width",
    true,
  );
  const outputCropHeight = exactDeskewCoordinate(
    value.outputCropHeight,
    "PDF deskew output crop height",
    true,
  );
  if (
    outputCropX < outputMediaX ||
    outputCropY < outputMediaY ||
    outputCropX + outputCropWidth > outputMediaX + outputMediaWidth ||
    outputCropY + outputCropHeight > outputMediaY + outputMediaHeight
  ) {
    throw new Error("PDF deskew output crop is outside the media box");
  }
  if (
    Math.ceil((outputCropWidth * PDF_DESKEW_OUTPUT_DPI) / 72) !==
      outputPixelWidth ||
    Math.ceil((outputCropHeight * PDF_DESKEW_OUTPUT_DPI) / 72) !==
      outputPixelHeight
  ) {
    throw new Error("PDF deskew output pixels do not match physical crop");
  }
  if (
    typeof value.angle !== "number" ||
    !Number.isFinite(value.angle) ||
    Math.abs(value.angle) < 0.3 ||
    Math.abs(value.angle) > 45
  ) {
    throw new Error("Invalid PDF deskew angle");
  }
  if (
    typeof value.confidence !== "number" ||
    !Number.isFinite(value.confidence) ||
    value.confidence < 0.15 ||
    value.confidence > 1
  ) {
    throw new Error("Invalid PDF deskew confidence");
  }
  const sourceByteLength = exactPositiveSafeInteger(
    value.sourceByteLength,
    "PDF deskew source byte length",
  );
  if (sourceByteLength > MAX_SOURCE_BYTES) {
    throw new Error("PDF deskew source byte budget exceeded");
  }
  if (canvasBytes + DESKEW_OUTPUT_SAFETY_BYTES > MAX_DESKEW_RASTER_BYTES) {
    throw new Error("PDF deskew raster byte budget exceeded");
  }
  const outputCanvasBytes = outputPixelWidth * outputPixelHeight * 4;
  if (
    outputCanvasBytes + DESKEW_OUTPUT_SAFETY_BYTES >
    MAX_DESKEW_RASTER_BYTES
  ) {
    throw new Error("PDF deskew output raster byte budget exceeded");
  }
  const planDigest = exactString(
    value.planDigest,
    "PDF deskew plan digest",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(planDigest)) {
    throw new Error("Invalid PDF deskew plan digest");
  }
  return Object.freeze({
    dataItemId,
    sourceUrl,
    previewByteGenerationKey,
    previewSha256,
    documentRevision,
    editRevision,
    pageNumber,
    pdfJsDocument: value.pdfJsDocument,
    pdfJsDocumentIdentity,
    resolvedRotation: value.resolvedRotation,
    canvas: value.canvas,
    canvasIdentity,
    canvasGeneration,
    canvasWidth,
    canvasHeight,
    sourceUserUnit,
    outputDpi: PDF_DESKEW_OUTPUT_DPI,
    outputPixelWidth,
    outputPixelHeight,
    outputMediaX,
    outputMediaY,
    outputMediaWidth,
    outputMediaHeight,
    outputCropX,
    outputCropY,
    outputCropWidth,
    outputCropHeight,
    angle: value.angle,
    confidence: value.confidence,
    sourceByteLength,
    planDigest,
  });
}

export function createPdfDeskewRasterPolicyContext(
  value: PdfDeskewRasterPolicyContext,
): PdfDeskewRasterPolicyContext {
  const normalized = normalizePdfDeskewRasterPolicyContext(value);
  return Object.freeze({
    dataItemId: normalized.dataItemId,
    sourceUrl: normalized.sourceUrl,
    previewByteGenerationKey: normalized.previewByteGenerationKey,
    previewSha256: normalized.previewSha256,
    documentRevision: normalized.documentRevision,
    editRevision: normalized.editRevision,
    pageNumber: normalized.pageNumber,
    pdfJsDocument: normalized.pdfJsDocument,
    resolvedRotation: normalized.resolvedRotation,
    canvas: normalized.canvas,
    canvasGeneration: normalized.canvasGeneration,
    canvasWidth: normalized.canvasWidth,
    canvasHeight: normalized.canvasHeight,
    sourceUserUnit: normalized.sourceUserUnit,
    outputDpi: normalized.outputDpi,
    outputPixelWidth: normalized.outputPixelWidth,
    outputPixelHeight: normalized.outputPixelHeight,
    outputMediaX: normalized.outputMediaX,
    outputMediaY: normalized.outputMediaY,
    outputMediaWidth: normalized.outputMediaWidth,
    outputMediaHeight: normalized.outputMediaHeight,
    outputCropX: normalized.outputCropX,
    outputCropY: normalized.outputCropY,
    outputCropWidth: normalized.outputCropWidth,
    outputCropHeight: normalized.outputCropHeight,
    angle: normalized.angle,
    confidence: normalized.confidence,
    sourceByteLength: normalized.sourceByteLength,
    planDigest: normalized.planDigest,
  });
}

function normalizePdfPageBackgroundPolicyContext(
  value: PdfPageBackgroundPolicyContext,
): PdfPageBackgroundPolicyContext {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid PDF background policy context");
  }
  const dataItemId = exactString(
    value.dataItemId,
    "PDF background data item ID",
    MAX_DATA_ITEM_ID_LENGTH,
  );
  const sourceUrl = exactString(
    value.sourceUrl,
    "PDF background source URL",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewByteGenerationKey = exactString(
    value.previewByteGenerationKey,
    "PDF background byte generation",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewSha256 = exactString(
    value.previewSha256,
    "PDF background source SHA-256",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(previewSha256)) {
    throw new Error("Invalid PDF background source SHA-256");
  }
  const documentRevision = exactRevision(
    value.documentRevision,
    "PDF background document revision",
  );
  const editRevision = exactRevision(
    value.editRevision,
    "PDF background edit revision",
  );
  if (value.kind !== "apply" && value.kind !== "remove") {
    throw new Error("Invalid PDF background operation");
  }
  const sourceByteLength = exactPositiveSafeInteger(
    value.sourceByteLength,
    "PDF background source byte length",
  );
  if (sourceByteLength > MAX_SOURCE_BYTES) {
    throw new Error("PDF background source exceeds byte limit");
  }
  const pageCount = exactPositiveSafeInteger(
    value.pageCount,
    "PDF background page count",
  );
  if (pageCount > 10_000 || !Array.isArray(value.pages)) {
    throw new Error("Invalid PDF background page plan");
  }
  const pages = value.pages.map((page, index) => {
    const normalized = exactPositiveSafeInteger(
      page,
      `PDF background page ${index}`,
    );
    if (normalized > pageCount) {
      throw new Error("PDF background target exceeds page count");
    }
    return normalized;
  });
  if (
    pages.length < 1 ||
    pages.length > 10_000 ||
    pages.some((page, index) => index > 0 && page <= pages[index - 1])
  ) {
    throw new Error("PDF background pages are not sorted unique");
  }
  if (!Array.isArray(value.markerRefs) || value.markerRefs.length > 10_000) {
    throw new Error("Invalid PDF background marker set");
  }
  const markerRefs = value.markerRefs.map((marker, index) =>
    exactString(marker, `PDF background marker ${index}`, 160),
  );
  if (
    markerRefs.some(
      (marker, index) => index > 0 && marker <= markerRefs[index - 1],
    )
  ) {
    throw new Error("PDF background markers are not sorted unique");
  }
  if (
    !Number.isSafeInteger(value.replaceOrRemoveCount) ||
    value.replaceOrRemoveCount < 0 ||
    value.replaceOrRemoveCount !== markerRefs.length
  ) {
    throw new Error("Invalid PDF background replacement count");
  }
  const color =
    value.color === null
      ? null
      : exactString(value.color, "PDF background colour", 7).toLowerCase();
  if (color !== null && !/^#[0-9a-f]{6}$/.test(color)) {
    throw new Error("Invalid PDF background colour");
  }
  if (
    typeof value.opacity !== "number" ||
    !Number.isFinite(value.opacity) ||
    value.opacity < 0 ||
    value.opacity > 1
  ) {
    throw new Error("Invalid PDF background opacity");
  }
  if (value.imageFit !== "fit" && value.imageFit !== "stretch") {
    throw new Error("Invalid PDF background image fit");
  }
  let image: PdfPageBackgroundPolicyImage | null = null;
  if (value.image !== null) {
    const byteLength = exactPositiveSafeInteger(
      value.image.byteLength,
      "PDF background image byte length",
    );
    if (byteLength > 16 * 1024 * 1024) {
      throw new Error("PDF background image exceeds byte limit");
    }
    const sha256 = exactString(
      value.image.sha256,
      "PDF background image SHA-256",
      64,
    ).toLowerCase();
    if (!SHA256_RE.test(sha256)) {
      throw new Error("Invalid PDF background image SHA-256");
    }
    if (
      value.image.mimeType !== "image/png" &&
      value.image.mimeType !== "image/jpeg"
    ) {
      throw new Error("Invalid PDF background image MIME type");
    }
    const width = exactPositiveSafeInteger(
      value.image.width,
      "PDF background image width",
    );
    const height = exactPositiveSafeInteger(
      value.image.height,
      "PDF background image height",
    );
    const pixelCount = exactPositiveSafeInteger(
      value.image.pixelCount,
      "PDF background image pixel count",
    );
    if (width * height !== pixelCount || pixelCount > 40_000_000) {
      throw new Error("Invalid PDF background image dimensions");
    }
    image = Object.freeze({
      byteLength,
      sha256,
      mimeType: value.image.mimeType,
      width,
      height,
      pixelCount,
    });
  }
  if (value.kind === "apply" && color === null && image === null) {
    throw new Error("PDF background apply plan is empty");
  }
  if (
    value.kind === "remove" &&
    (color !== null || image !== null || value.replaceOrRemoveCount < 1)
  ) {
    throw new Error("Invalid PDF background remove plan");
  }
  const planDigest = exactString(
    value.planDigest,
    "PDF background plan digest",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(planDigest)) {
    throw new Error("Invalid PDF background plan digest");
  }
  return Object.freeze({
    dataItemId,
    sourceUrl,
    previewByteGenerationKey,
    previewSha256,
    documentRevision,
    editRevision,
    kind: value.kind,
    sourceByteLength,
    pageCount,
    pages: Object.freeze(pages),
    markerRefs: Object.freeze(markerRefs),
    replaceOrRemoveCount: value.replaceOrRemoveCount,
    color,
    opacity: value.opacity,
    imageFit: value.imageFit,
    image,
    planDigest,
  });
}

export function createPdfPageBackgroundPolicyContext(
  value: PdfPageBackgroundPolicyContext,
): PdfPageBackgroundPolicyContext {
  return normalizePdfPageBackgroundPolicyContext(value);
}

function normalizePdfEmbeddedImageOverlayPolicyContext(
  value: PdfEmbeddedImageOverlayPolicyContext,
): PdfEmbeddedImageOverlayPolicyContext {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid embedded-image overlay policy context");
  }
  const dataItemId = exactString(
    value.dataItemId,
    "embedded-image data item ID",
    MAX_DATA_ITEM_ID_LENGTH,
  );
  const sourceUrl = exactString(
    value.sourceUrl,
    "embedded-image source URL",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewByteGenerationKey = exactString(
    value.previewByteGenerationKey,
    "embedded-image byte generation",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewSha256 = exactString(
    value.previewSha256,
    "embedded-image source SHA-256",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(previewSha256)) {
    throw new Error("Invalid embedded-image source SHA-256");
  }
  const documentRevision = exactRevision(
    value.documentRevision,
    "embedded-image document revision",
  );
  const editRevision = exactRevision(
    value.editRevision,
    "embedded-image edit revision",
  );
  const sourceByteLength = exactPositiveSafeInteger(
    value.sourceByteLength,
    "embedded-image source byte length",
  );
  if (sourceByteLength > MAX_SOURCE_BYTES) {
    throw new Error("Embedded-image source exceeds byte limit");
  }
  const pageCount = exactPositiveSafeInteger(
    value.pageCount,
    "embedded-image page count",
  );
  if (pageCount > 10_000 || !Array.isArray(value.targetPages)) {
    throw new Error("Invalid embedded-image page plan");
  }
  const targetPages = value.targetPages.map((page, index) => {
    const normalized = exactPositiveSafeInteger(
      page,
      `embedded-image target page ${index}`,
    );
    if (normalized > pageCount) {
      throw new Error("Embedded-image target exceeds page count");
    }
    return normalized;
  });
  if (
    targetPages.length < 1 ||
    targetPages.length > 200 ||
    targetPages.some(
      (page, index) => index > 0 && page <= targetPages[index - 1],
    )
  ) {
    throw new Error("Embedded-image target pages are not sorted unique");
  }
  const editCount = exactPositiveSafeInteger(
    value.editCount,
    "embedded-image edit count",
  );
  if (editCount > 200 || editCount < targetPages.length) {
    throw new Error("Invalid embedded-image edit count");
  }
  if (
    !Number.isSafeInteger(value.replacementCount) ||
    value.replacementCount < 0 ||
    value.replacementCount > editCount
  ) {
    throw new Error("Invalid embedded-image replacement count");
  }
  if (
    !Number.isSafeInteger(value.replacementByteLength) ||
    value.replacementByteLength < 0 ||
    value.replacementByteLength > MAX_SOURCE_BYTES ||
    (value.replacementCount === 0) !== (value.replacementByteLength === 0)
  ) {
    throw new Error("Invalid embedded-image replacement byte length");
  }
  const geometriesDigest = exactString(
    value.geometriesDigest,
    "embedded-image geometry digest",
    64,
  ).toLowerCase();
  const planDigest = exactString(
    value.planDigest,
    "embedded-image plan digest",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(geometriesDigest) || !SHA256_RE.test(planDigest)) {
    throw new Error("Invalid embedded-image plan digest");
  }
  return Object.freeze({
    dataItemId,
    sourceUrl,
    previewByteGenerationKey,
    previewSha256,
    documentRevision,
    editRevision,
    sourceByteLength,
    pageCount,
    targetPages: Object.freeze(targetPages),
    editCount,
    replacementCount: value.replacementCount,
    replacementByteLength: value.replacementByteLength,
    geometriesDigest,
    planDigest,
  });
}

export function createPdfEmbeddedImageOverlayPolicyContext(
  value: PdfEmbeddedImageOverlayPolicyContext,
): PdfEmbeddedImageOverlayPolicyContext {
  return normalizePdfEmbeddedImageOverlayPolicyContext(value);
}

const PDF_LOSSY_PRESET_CONTRACT = Object.freeze({
  screen: Object.freeze({ targetDpi: 72, jpegQuality: 60 }),
  ebook: Object.freeze({ targetDpi: 150, jpegQuality: 75 }),
  print: Object.freeze({ targetDpi: 300, jpegQuality: 85 }),
});

function normalizePdfLossyCompressionPolicyContext(
  value: PdfLossyCompressionPolicyContext,
): PdfLossyCompressionPolicyContext {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid lossy-compression policy context");
  }
  const previewDataItemId = exactString(
    value.previewDataItemId,
    "lossy-compression preview data item ID",
    MAX_DATA_ITEM_ID_LENGTH,
  );
  const previewByteGenerationKey = exactString(
    value.previewByteGenerationKey,
    "lossy-compression preview byte generation",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
  const previewSha256 = exactString(
    value.previewSha256,
    "lossy-compression preview SHA-256",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(previewSha256)) {
    throw new Error("Invalid lossy-compression preview SHA-256");
  }
  const documentRevision = exactRevision(
    value.documentRevision,
    "lossy-compression document revision",
  );
  const editRevision = exactRevision(
    value.editRevision,
    "lossy-compression edit revision",
  );
  if (
    value.sourceKind !== "data_item" &&
    value.sourceKind !== "resource_version"
  ) {
    throw new Error("Invalid lossy-compression source kind");
  }
  const workspaceId = exactString(
    value.workspaceId,
    "lossy-compression workspace ID",
    64,
  );
  const optionalId = (input: unknown, label: string): string | null =>
    input === null ? null : exactString(input, label, 64);
  const datasetId = optionalId(value.datasetId, "lossy-compression dataset ID");
  const dataItemId = optionalId(
    value.dataItemId,
    "lossy-compression data item ID",
  );
  const resourceId = optionalId(
    value.resourceId,
    "lossy-compression resource ID",
  );
  const resourceVersionId = optionalId(
    value.resourceVersionId,
    "lossy-compression resource version ID",
  );
  const versionKind = optionalId(
    value.versionKind,
    "lossy-compression resource version kind",
  );
  const expectedPageId = optionalId(
    value.expectedPageId,
    "lossy-compression expected page ID",
  );
  if (
    (value.sourceKind === "data_item" &&
      (datasetId === null ||
        dataItemId !== previewDataItemId ||
        resourceId !== null ||
        resourceVersionId !== null ||
        versionKind !== null ||
        expectedPageId !== null)) ||
    (value.sourceKind === "resource_version" &&
      (datasetId !== null ||
        dataItemId !== null ||
        resourceId === null ||
        resourceVersionId === null ||
        versionKind === null))
  ) {
    throw new Error("Invalid lossy-compression source lineage");
  }
  const storageCustodyDigest = exactString(
    value.storageCustodyDigest,
    "lossy-compression storage custody digest",
    64,
  ).toLowerCase();
  const nativePlanDigest = exactString(
    value.nativePlanDigest,
    "lossy-compression native plan digest",
    64,
  ).toLowerCase();
  const evidenceDigest = exactString(
    value.evidenceDigest,
    "lossy-compression evidence digest",
    64,
  ).toLowerCase();
  if (
    !SHA256_RE.test(storageCustodyDigest) ||
    !SHA256_RE.test(nativePlanDigest) ||
    !SHA256_RE.test(evidenceDigest)
  ) {
    throw new Error("Invalid lossy-compression evidence digest");
  }
  const sourceSize = exactPositiveSafeInteger(
    value.sourceSize,
    "lossy-compression source size",
  );
  if (sourceSize > MAX_SOURCE_BYTES) {
    throw new Error("Lossy-compression source exceeds byte limit");
  }
  if (
    value.preset !== "screen" &&
    value.preset !== "ebook" &&
    value.preset !== "print"
  ) {
    throw new Error("Invalid lossy-compression preset");
  }
  const preset = value.preset;
  const presetContract = PDF_LOSSY_PRESET_CONTRACT[preset];
  if (
    value.targetDpi !== presetContract.targetDpi ||
    value.jpegQuality !== presetContract.jpegQuality
  ) {
    throw new Error("Lossy-compression preset contract mismatch");
  }
  const pageCount = exactPositiveSafeInteger(
    value.pageCount,
    "lossy-compression page count",
  );
  if (pageCount > 1_000) {
    throw new Error("Lossy-compression page count exceeds limit");
  }
  const transformedObjectCount = exactPositiveSafeInteger(
    value.transformedObjectCount,
    "lossy-compression transformed object count",
  );
  if (transformedObjectCount > 128) {
    throw new Error("Lossy-compression image count exceeds limit");
  }
  return Object.freeze({
    previewDataItemId,
    previewByteGenerationKey,
    previewSha256,
    documentRevision,
    editRevision,
    sourceKind: value.sourceKind,
    workspaceId,
    datasetId,
    dataItemId,
    resourceId,
    resourceVersionId,
    versionKind,
    expectedPageId,
    storageCustodyDigest,
    sourceSize,
    preset,
    targetDpi: presetContract.targetDpi,
    jpegQuality: presetContract.jpegQuality,
    pageCount,
    transformedObjectCount,
    nativePlanDigest,
    evidenceDigest,
  });
}

export function createPdfLossyCompressionPolicyContext(
  value: PdfLossyCompressionPolicyContext,
): PdfLossyCompressionPolicyContext {
  return normalizePdfLossyCompressionPolicyContext(value);
}

function normalizePagePackageSource(
  value: PdfPagePackageSourceSnapshot,
  index: number,
): PdfPagePackageSourceSnapshot {
  if (!value || typeof value !== "object") {
    throw new Error(`Invalid PDF page-package source ${index}`);
  }
  const source = normalizeSnapshot(value.source);
  if (!source) throw new Error(`Missing PDF page-package source ${index}`);
  const byteLength = exactPositiveSafeInteger(
    value.byteLength,
    `PDF page-package byte length ${index}`,
  );
  if (byteLength > MAX_SOURCE_BYTES) {
    throw new Error(`PDF page-package source ${index} exceeds byte limit`);
  }
  const pageCount = exactPositiveSafeInteger(
    value.pageCount,
    `PDF page-package page count ${index}`,
  );
  const structureRisk = riskState(
    value.structureRisk,
    `PDF page-package structure risk ${index}`,
  );
  return Object.freeze({ source, byteLength, pageCount, structureRisk });
}

export function createPdfPagePackagePolicyContext(options: {
  readonly sources: readonly PdfPagePackageSourceSnapshot[];
  readonly planDigest: string;
  readonly outputCount: number;
}): PdfPagePackagePolicyContext {
  if (
    !Array.isArray(options.sources) ||
    options.sources.length < 1 ||
    options.sources.length > MAX_PAGE_PACKAGE_SOURCES
  ) {
    throw new Error("Invalid PDF page-package source count");
  }
  const sources: PdfPagePackageSourceSnapshot[] = [];
  let aggregateBytes = 0;
  let totalPages = 0;
  for (let index = 0; index < options.sources.length; index += 1) {
    if (!(index in options.sources)) {
      throw new Error("Sparse PDF page-package sources");
    }
    const source = normalizePagePackageSource(options.sources[index], index);
    aggregateBytes += source.byteLength;
    totalPages += source.pageCount;
    if (
      aggregateBytes > MAX_PAGE_PACKAGE_AGGREGATE_BYTES ||
      totalPages > MAX_PAGE_PACKAGE_PAGES
    ) {
      throw new Error("PDF page-package source budget exceeded");
    }
    sources.push(source);
  }
  const planDigest = exactString(
    options.planDigest,
    "PDF page-package plan digest",
    64,
  ).toLowerCase();
  if (!SHA256_RE.test(planDigest)) {
    throw new Error("Invalid PDF page-package plan digest");
  }
  const outputCount = exactPositiveSafeInteger(
    options.outputCount,
    "PDF page-package output count",
  );
  if (outputCount > MAX_PAGE_PACKAGE_OUTPUTS) {
    throw new Error("PDF page-package output limit exceeded");
  }
  return Object.freeze({
    sources: Object.freeze(sources),
    planDigest,
    outputCount,
  });
}

export async function createPdfPagePackagePlanDigest(
  parts: readonly string[],
): Promise<string> {
  if (
    !Array.isArray(parts) ||
    parts.length < 1 ||
    parts.length > MAX_PAGE_PACKAGE_PLAN_PARTS
  ) {
    throw new Error("Invalid PDF page-package plan");
  }
  const normalized: string[] = [];
  for (let index = 0; index < parts.length; index += 1) {
    if (!(index in parts)) throw new Error("Sparse PDF page-package plan");
    normalized.push(
      exactString(
        parts[index],
        `PDF page-package plan part ${index}`,
        MAX_PAGE_PACKAGE_PLAN_PART_LENGTH,
      ),
    );
  }
  const encoded = new TextEncoder().encode(JSON.stringify(normalized));
  if (encoded.byteLength > 128 * 1024) {
    encoded.fill(0);
    throw new Error("PDF page-package plan exceeds byte limit");
  }
  try {
    return await sha256Hex(encoded);
  } finally {
    encoded.fill(0);
  }
}

/**
 * Bind a non-page-package PDF output plan to the same bounded, canonical
 * digest primitive used by page-package actions. The alias keeps the native
 * action vocabulary out of the cryptographic contract.
 */
export async function createPdfOutputPlanDigest(
  parts: readonly string[],
): Promise<string> {
  return createPdfPagePackagePlanDigest(parts);
}

/** Hash a caller-owned, already bounded input without retaining its bytes. */
export async function createPdfBoundedInputDigest(
  value: Uint8Array,
  maxBytes: number,
): Promise<string> {
  if (
    !(value instanceof Uint8Array) ||
    !Number.isSafeInteger(maxBytes) ||
    maxBytes < 1 ||
    value.byteLength < 1 ||
    value.byteLength > maxBytes
  ) {
    throw new Error("Invalid bounded PDF policy input");
  }
  return sha256Hex(value);
}

function normalizeLossEntry(
  value: PdfOutputLossEntry,
  index: number,
): PdfOutputLossEntry {
  if (!value || typeof value !== "object") {
    throw new Error(`Invalid PDF loss entry ${index}`);
  }
  const code = exactString(
    value.code,
    `PDF loss code ${index}`,
    MAX_LOSS_CODE_LENGTH,
  );
  if (!/^[A-Z0-9][A-Z0-9_.-]*$/.test(code)) {
    throw new Error(`Invalid PDF loss code ${index}`);
  }
  const label = exactString(
    value.label,
    `PDF loss label ${index}`,
    MAX_LOSS_LABEL_LENGTH,
  );
  const detail =
    value.detail === undefined
      ? undefined
      : exactString(
          value.detail,
          `PDF loss detail ${index}`,
          MAX_LOSS_DETAIL_LENGTH,
        );
  return Object.freeze({ code, label, ...(detail ? { detail } : {}) });
}

export function createPdfOutputLossReport(
  values: readonly PdfOutputLossEntry[],
): PdfOutputLossReport {
  if (!Array.isArray(values) || values.length > MAX_LOSS_ENTRIES) {
    throw new Error("Invalid PDF output loss report");
  }
  const entries: PdfOutputLossEntry[] = [];
  const seen = new Set<string>();
  for (let index = 0; index < values.length; index += 1) {
    if (!(index in values)) throw new Error("Sparse PDF output loss report");
    const entry = normalizeLossEntry(values[index], index);
    if (seen.has(entry.code)) throw new Error("Duplicate PDF output loss code");
    seen.add(entry.code);
    entries.push(entry);
  }
  const codes = [...seen].sort();
  return Object.freeze({
    entries: Object.freeze(entries),
    codes: Object.freeze(codes),
    digest: `pdf-loss-v1:${codes.join(",")}`,
  });
}

const OUTPUT_LOSSES: Readonly<
  Partial<Record<PdfSavePolicyAction, readonly PdfOutputLossEntry[]>>
> = Object.freeze({
  "download-current": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FLATTEN_REWRITE",
      label:
        "The edited PDF is fully rewritten instead of incrementally saved.",
      detail: "The immutable original remains available separately.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
  ]),
  "download-annotated": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FLATTEN_REWRITE",
      label: "Annotations and edits are rendered into rewritten PDF content.",
      detail: "The immutable original remains available separately.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INTERACTIVE_ANNOTATIONS",
      label: "Rendered annotations may no longer be independently editable.",
    }),
  ]),
  "download-fillable": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FLATTEN_REWRITE",
      label: "Page edits are rendered into rewritten PDF content.",
      detail: "AcroForm fields remain fillable when supported.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INTERACTIVE_ANNOTATIONS",
      label: "Rendered annotations may no longer be independently editable.",
    }),
  ]),
  "download-editable-annotations": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FLATTEN_REWRITE",
      label: "Page edits are rewritten into a new PDF package.",
      detail: "Supported annotations remain editable; unknown objects do not.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
  ]),
  "download-archival": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FLATTEN_REWRITE",
      label: "The document is fully rewritten for archival output.",
    }),
    Object.freeze({
      code: "PDF_LOSS_NON_PDFA_FEATURES",
      label:
        "Features incompatible with the selected PDF/A profile are removed.",
    }),
  ]),
  "download-properties": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FLATTEN_REWRITE",
      label: "Saving properties fully rewrites the PDF package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
  ]),
  "pdf.redaction.apply": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_REDACTION_TARGET_CONTENT",
      label:
        "Content covered by the selected redaction marks is permanently removed from the new PDF.",
      detail: "The immutable original and workspace marks remain available.",
    }),
    Object.freeze({
      code: "PDF_LOSS_REDACTED_PAGE_STRUCTURE",
      label:
        "Each redacted page becomes a static image, so its searchable text, forms and interactive structure are removed.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  print: Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_PRINT_SNAPSHOT",
      label: "Interactive content is rendered into a static print snapshot.",
    }),
  ]),
  "replace-source": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_SOURCE_REPLACED",
      label: "Current PDF edit layers do not move to the replacement source.",
      detail: "The previous source remains available in version history.",
    }),
  ]),
  "pdf.page.insert-file": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_PAGE_COPY_FULL_REWRITE",
      label: "Pages are copied into a newly written PDF package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_DOCUMENT_LEVEL_STRUCTURES",
      label:
        "Document-level structures outside copied pages are not carried forward.",
    }),
    Object.freeze({
      code: "PDF_LOSS_CROSS_PAGE_REFERENCES",
      label: "Cross-page references are not preserved by page copying.",
    }),
    Object.freeze({
      code: "PDF_LOSS_ADDITIONAL_SOURCE_CONTEXT",
      label:
        "Document-level context from the inserted source is not carried forward.",
    }),
    PDF_LOSS_CURRENT_EDITOR_EDITS_EXCLUDED,
  ]),
  "pdf.page.extract": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_PAGE_COPY_FULL_REWRITE",
      label: "Selected pages are copied into a newly written PDF package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_DOCUMENT_LEVEL_STRUCTURES",
      label:
        "Document-level structures outside copied pages are not carried forward.",
    }),
    Object.freeze({
      code: "PDF_LOSS_CROSS_PAGE_REFERENCES",
      label: "Cross-page references are not preserved by page copying.",
    }),
    Object.freeze({
      code: "PDF_LOSS_UNSELECTED_PAGES",
      label: "Pages outside the selected range are omitted.",
    }),
    PDF_LOSS_CURRENT_EDITOR_EDITS_EXCLUDED,
  ]),
  "pdf.file.split": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_PAGE_COPY_FULL_REWRITE",
      label: "Each split group is copied into a newly written PDF package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_DOCUMENT_LEVEL_STRUCTURES",
      label:
        "Document-level structures outside copied pages are not carried forward.",
    }),
    Object.freeze({
      code: "PDF_LOSS_CROSS_PAGE_REFERENCES",
      label: "Cross-page references are not preserved by page copying.",
    }),
    Object.freeze({
      code: "PDF_LOSS_UNSELECTED_PAGES",
      label:
        "Each split output omits pages assigned to the other output packages.",
    }),
    PDF_LOSS_CURRENT_EDITOR_EDITS_EXCLUDED,
  ]),
  "pdf.file.merge": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_PAGE_COPY_FULL_REWRITE",
      label: "Source pages are copied into a newly written PDF package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_DOCUMENT_LEVEL_STRUCTURES",
      label:
        "Document-level structures outside copied pages are not carried forward.",
    }),
    Object.freeze({
      code: "PDF_LOSS_CROSS_PAGE_REFERENCES",
      label: "Cross-page references are not preserved by page copying.",
    }),
    Object.freeze({
      code: "PDF_LOSS_ADDITIONAL_SOURCE_CONTEXT",
      label:
        "Document-level context from additional source files is not carried forward.",
    }),
    PDF_LOSS_CURRENT_EDITOR_EDITS_EXCLUDED,
  ]),
  "pdf.file.compress-lossless": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label:
        "The PDF is structurally rewritten into a new object-stream package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
    Object.freeze({
      code: "PDF_LOSS_UNREACHABLE_OBJECTS",
      label: "Objects unreachable from the document catalog are omitted.",
    }),
  ]),
  "pdf.file.compress-lossy": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_IMAGE_RESAMPLING",
      label:
        "Eligible raster images are resampled to the selected screen, ebook, or print target.",
    }),
    Object.freeze({
      code: "PDF_LOSS_JPEG_RECOMPRESSION",
      label:
        "Eligible raster images are recompressed as JPEG and may lose fine visual detail.",
    }),
    Object.freeze({
      code: "PDF_LOSS_IMAGE_COLOR_FIDELITY",
      label:
        "JPEG recompression can change pixel-level colour and grayscale fidelity.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten into a new derivative package.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.file.deskew-page": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_DESKEW_TARGET_PAGE_TEXT",
      label:
        "Any selectable text in the target page content becomes raster pixels.",
    }),
    Object.freeze({
      code: "PDF_LOSS_DESKEW_TARGET_PAGE_STRUCTURE",
      label:
        "Logical structure and reading order associated with the target page are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_DESKEW_FIXED_300_DPI",
      label:
        "The target page is regenerated at a fixed 300 DPI; finer source detail may be downsampled.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.file.image-only": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_IMAGE_ONLY_SEARCHABLE_TEXT",
      label:
        "Selectable and searchable page text becomes raster pixels in the new PDF.",
      detail: "The immutable original remains available separately.",
    }),
    Object.freeze({
      code: "PDF_LOSS_IMAGE_ONLY_INTERACTIVE_CONTENT",
      label:
        "Forms, links, comments and other interactive page objects are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_IMAGE_ONLY_ACCESSIBILITY",
      label:
        "Tags, reading order and other accessibility structure are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_IMAGE_ONLY_SIGNATURES",
      label:
        "Digital signatures and certification do not carry into the image-only derivative.",
    }),
    Object.freeze({
      code: "PDF_LOSS_IMAGE_ONLY_FIXED_150_DPI",
      label:
        "Every page is rendered at a fixed 150 DPI; finer source detail may be downsampled.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.form.create-detected-fields": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FORM_FIELD_GRAPH_CREATE",
      label:
        "The accepted detected fields are added as a new editable AcroForm field graph.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten with new form widget appearances.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.form.edit-existing-field": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FORM_FIELD_PROPERTY_EDIT",
      label:
        "The selected AcroForm field's supported native properties and widget geometry are updated in the derivative PDF.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label:
        "The PDF is fully rewritten while the edited field appearance is refreshed or explicitly delegated to the reader.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.form.delete-existing-field": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FORM_FIELD_DELETE",
      label:
        "The selected AcroForm field and all of its native widgets are removed from the derivative PDF.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label:
        "The PDF is fully rewritten while the selected field is removed and page annotation references are reconciled.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.form.reset-values": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FORM_VALUE_RESET",
      label: "Every supported AcroForm field value is reset in the output.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten with refreshed field appearances.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.form.import-data": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_FORM_VALUE_IMPORT",
      label:
        "Matching AcroForm field values are replaced by the selected CSV data.",
      detail: "Unknown field names and unsupported values are skipped.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten with refreshed field appearances.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.page.apply-background": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_BACKGROUND_MANAGED_CONTENT",
      label:
        "A managed page-background stream is added or replaced in the derivative PDF.",
    }),
    Object.freeze({
      code: "PDF_LOSS_BACKGROUND_NATIVE_PORTABILITY",
      label:
        "The background is ordinary PDF marked content, not a portable native background setting in every PDF editor.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten to add the managed background.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.page.remove-background": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_BACKGROUND_MANAGED_CONTENT_REMOVED",
      label:
        "The exact selected MRAI-managed page-background streams are removed from the derivative PDF.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten to remove the managed background.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
  "pdf.image.apply-overlay-edits": Object.freeze([
    Object.freeze({
      code: "PDF_LOSS_EMBEDDED_IMAGE_OVERLAY_SEMANTICS",
      label:
        "Image edits are written as visual masks and overlays; the original image object remains recoverable beneath the mask.",
      detail:
        "This derivative is an approximate visual cover, not a native PDF XObject replacement.",
    }),
    Object.freeze({
      code: "PDF_LOSS_STRUCTURAL_REWRITE",
      label: "The PDF is fully rewritten to add the image-edit overlays.",
    }),
    Object.freeze({
      code: "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
      label: "Original object numbers and byte offsets are not preserved.",
    }),
    Object.freeze({
      code: "PDF_LOSS_INCREMENTAL_HISTORY",
      label: "Previous incremental-update revisions are collapsed.",
    }),
  ]),
});

export function createPdfOutputLossReportForAction(
  action: PdfSavePolicyAction,
  additionalEntries: readonly PdfOutputLossEntry[] = [],
): PdfOutputLossReport {
  return createPdfOutputLossReport([
    ...(OUTPUT_LOSSES[action] ?? []),
    ...additionalEntries,
  ]);
}

export function createPdfSavePolicyEditSnapshot(options: {
  editRevision: PdfSavePolicyRevision;
  hasEdits: boolean;
  nativeDurableReceipt?: boolean;
  lossReport?: PdfOutputLossReport;
}): PdfSavePolicyEditSnapshot {
  const lossReport = options.lossReport ?? createPdfOutputLossReport([]);
  return Object.freeze({
    editRevision: exactRevision(options.editRevision, "PDF edit revision"),
    hasEdits: options.hasEdits === true,
    nativeDurableReceipt: options.nativeDurableReceipt === true,
    lossReport: createPdfOutputLossReport(lossReport.entries),
  });
}

function normalizeSnapshot(
  value: PdfSourceRiskSnapshot | null | undefined,
): PdfSourceRiskSnapshot | null {
  if (value == null) return null;
  if (createdSourceRiskSnapshots.has(value)) return value;
  return createPdfSourceRiskSnapshot(value);
}

function sourceRevision(
  source: PdfSourceRiskSnapshot | null,
  replacementSource: PdfSourceRiskSnapshot | null,
  pagePackage: PdfPagePackagePolicyContext | null,
  deskewRaster: NormalizedPdfDeskewRasterPolicyContext | null,
  pageBackground: PdfPageBackgroundPolicyContext | null,
  embeddedImageOverlay: PdfEmbeddedImageOverlayPolicyContext | null,
  lossyCompression: PdfLossyCompressionPolicyContext | null,
): string {
  if (!source) return "pdf-source-v1:missing";
  const primary = `${source.identity.dataItemId}:${source.identity.sha256}:${source.profileDigest}`;
  const replacement = replacementSource
    ? `>${replacementSource.identity.dataItemId}:${replacementSource.identity.sha256}:${replacementSource.profileDigest}`
    : "";
  const packageRevision = pagePackage
    ? `>${pagePackage.planDigest}:${pagePackage.sources
        .map((entry) => entry.source.identity.sha256)
        .join(".")}`
    : "";
  const deskewRevision = deskewRaster
    ? `>deskew:${deskewRaster.planDigest}:${deskewRaster.pdfJsDocumentIdentity}:${deskewRaster.canvasIdentity}`
    : "";
  const backgroundRevision = pageBackground
    ? `>background:${pageBackground.kind}:${pageBackground.planDigest}`
    : "";
  const imageOverlayRevision = embeddedImageOverlay
    ? `>image-overlay:${embeddedImageOverlay.planDigest}:${embeddedImageOverlay.geometriesDigest}`
    : "";
  const lossyCompressionRevision = lossyCompression
    ? `>lossy-compress:${lossyCompression.preset}:${lossyCompression.nativePlanDigest}:${lossyCompression.evidenceDigest}:${lossyCompression.storageCustodyDigest}`
    : "";
  return exactString(
    `pdf-source-v1:${primary}${replacement}${packageRevision}${deskewRevision}${backgroundRevision}${imageOverlayRevision}${lossyCompressionRevision}`,
    "PDF policy source revision",
    MAX_BYTE_GENERATION_KEY_LENGTH,
  );
}

function normalizeRequest(
  value: PdfSavePolicyRequest,
): NormalizedPdfSavePolicyRequest {
  if (!value || typeof value !== "object") {
    throw new Error("Invalid PDF save-policy request");
  }
  const host = value.host;
  if (host !== "reader" && host !== "native-editor") {
    throw new Error("Invalid PDF save-policy host");
  }
  const action = value.action;
  const featureId = ACTION_FEATURES[action];
  if (!featureId) throw new Error("Invalid PDF save-policy action");
  const requestedDisposition = value.requestedDisposition;
  if (
    !(
      ACTION_DISPOSITIONS[action] as readonly PdfSavePolicyDisposition[]
    ).includes(requestedDisposition)
  ) {
    throw new Error("Invalid PDF save-policy action disposition");
  }
  const source = normalizeSnapshot(value.source);
  const historyReplay = readPdfNativeReceiptReplayPolicy(value.historyReplay);
  const historyReplayRequested = value.historyReplay != null;
  const replacementSource = normalizeSnapshot(value.replacementSource);
  const pagePackage =
    value.pagePackage == null
      ? null
      : createPdfPagePackagePolicyContext(value.pagePackage);
  const deskewRaster =
    value.deskewRaster == null
      ? null
      : normalizePdfDeskewRasterPolicyContext(value.deskewRaster);
  const pageBackground =
    value.pageBackground == null
      ? null
      : normalizePdfPageBackgroundPolicyContext(value.pageBackground);
  const embeddedImageOverlay =
    value.embeddedImageOverlay == null
      ? null
      : normalizePdfEmbeddedImageOverlayPolicyContext(
          value.embeddedImageOverlay,
        );
  const lossyCompression =
    value.lossyCompression == null
      ? null
      : normalizePdfLossyCompressionPolicyContext(value.lossyCompression);
  const edit = createPdfSavePolicyEditSnapshot(value.edit);
  const heldWrites = action === "save-current" && Array.isArray(value.heldWrites) &&
    value.heldWrites.length > 0 &&
    value.heldWrites.every((held) => typeof held === "string" && held !== "save-current" &&
      Object.prototype.hasOwnProperty.call(ACTION_FEATURES, held))
    ? Object.freeze([...new Set(value.heldWrites)].sort())
    : null;
  const requestKey = JSON.stringify({
    host,
    action,
    heldWrites,
    requestedDisposition,
    documentReady: value.documentReady === true,
    editable: value.editable === true,
    source: source
      ? {
          dataItemId: source.identity.dataItemId,
          byteGenerationKey: source.identity.byteGenerationKey,
          sha256: source.identity.sha256,
          documentRevision: source.identity.documentRevision,
          profileDigest: source.profileDigest,
        }
      : null,
    replacementSource: replacementSource
      ? {
          dataItemId: replacementSource.identity.dataItemId,
          byteGenerationKey: replacementSource.identity.byteGenerationKey,
          sha256: replacementSource.identity.sha256,
          documentRevision: replacementSource.identity.documentRevision,
          profileDigest: replacementSource.profileDigest,
        }
      : null,
    pagePackage: pagePackage
      ? {
          sources: pagePackage.sources.map((entry) => ({
            identity: {
              dataItemId: entry.source.identity.dataItemId,
              byteGenerationKey: entry.source.identity.byteGenerationKey,
              sha256: entry.source.identity.sha256,
              documentRevision: entry.source.identity.documentRevision,
            },
            profileDigest: entry.source.profileDigest,
            byteLength: entry.byteLength,
            pageCount: entry.pageCount,
            structureRisk: entry.structureRisk,
          })),
          planDigest: pagePackage.planDigest,
          outputCount: pagePackage.outputCount,
        }
      : null,
    deskewRaster: deskewRaster
      ? {
          dataItemId: deskewRaster.dataItemId,
          sourceUrl: deskewRaster.sourceUrl,
          previewByteGenerationKey: deskewRaster.previewByteGenerationKey,
          previewSha256: deskewRaster.previewSha256,
          documentRevision: deskewRaster.documentRevision,
          editRevision: deskewRaster.editRevision,
          pageNumber: deskewRaster.pageNumber,
          pdfJsDocumentIdentity: deskewRaster.pdfJsDocumentIdentity,
          resolvedRotation: deskewRaster.resolvedRotation,
          canvasIdentity: deskewRaster.canvasIdentity,
          canvasGeneration: deskewRaster.canvasGeneration,
          canvasWidth: deskewRaster.canvasWidth,
          canvasHeight: deskewRaster.canvasHeight,
          sourceUserUnit: deskewRaster.sourceUserUnit,
          outputDpi: deskewRaster.outputDpi,
          outputPixelWidth: deskewRaster.outputPixelWidth,
          outputPixelHeight: deskewRaster.outputPixelHeight,
          outputMediaX: deskewRaster.outputMediaX,
          outputMediaY: deskewRaster.outputMediaY,
          outputMediaWidth: deskewRaster.outputMediaWidth,
          outputMediaHeight: deskewRaster.outputMediaHeight,
          outputCropX: deskewRaster.outputCropX,
          outputCropY: deskewRaster.outputCropY,
          outputCropWidth: deskewRaster.outputCropWidth,
          outputCropHeight: deskewRaster.outputCropHeight,
          angle: deskewRaster.angle,
          confidence: deskewRaster.confidence,
          sourceByteLength: deskewRaster.sourceByteLength,
          planDigest: deskewRaster.planDigest,
        }
      : null,
    pageBackground: pageBackground
      ? {
          dataItemId: pageBackground.dataItemId,
          sourceUrl: pageBackground.sourceUrl,
          previewByteGenerationKey: pageBackground.previewByteGenerationKey,
          previewSha256: pageBackground.previewSha256,
          documentRevision: pageBackground.documentRevision,
          editRevision: pageBackground.editRevision,
          kind: pageBackground.kind,
          sourceByteLength: pageBackground.sourceByteLength,
          pageCount: pageBackground.pageCount,
          pages: pageBackground.pages,
          markerRefs: pageBackground.markerRefs,
          replaceOrRemoveCount: pageBackground.replaceOrRemoveCount,
          color: pageBackground.color,
          opacity: pageBackground.opacity,
          imageFit: pageBackground.imageFit,
          image: pageBackground.image,
          planDigest: pageBackground.planDigest,
        }
      : null,
    embeddedImageOverlay: embeddedImageOverlay
      ? {
          dataItemId: embeddedImageOverlay.dataItemId,
          sourceUrl: embeddedImageOverlay.sourceUrl,
          previewByteGenerationKey:
            embeddedImageOverlay.previewByteGenerationKey,
          previewSha256: embeddedImageOverlay.previewSha256,
          documentRevision: embeddedImageOverlay.documentRevision,
          editRevision: embeddedImageOverlay.editRevision,
          sourceByteLength: embeddedImageOverlay.sourceByteLength,
          pageCount: embeddedImageOverlay.pageCount,
          targetPages: embeddedImageOverlay.targetPages,
          editCount: embeddedImageOverlay.editCount,
          replacementCount: embeddedImageOverlay.replacementCount,
          replacementByteLength: embeddedImageOverlay.replacementByteLength,
          geometriesDigest: embeddedImageOverlay.geometriesDigest,
          planDigest: embeddedImageOverlay.planDigest,
        }
      : null,
    lossyCompression: lossyCompression
      ? {
          previewDataItemId: lossyCompression.previewDataItemId,
          previewByteGenerationKey: lossyCompression.previewByteGenerationKey,
          previewSha256: lossyCompression.previewSha256,
          documentRevision: lossyCompression.documentRevision,
          editRevision: lossyCompression.editRevision,
          sourceKind: lossyCompression.sourceKind,
          workspaceId: lossyCompression.workspaceId,
          datasetId: lossyCompression.datasetId,
          dataItemId: lossyCompression.dataItemId,
          resourceId: lossyCompression.resourceId,
          resourceVersionId: lossyCompression.resourceVersionId,
          versionKind: lossyCompression.versionKind,
          expectedPageId: lossyCompression.expectedPageId,
          storageCustodyDigest: lossyCompression.storageCustodyDigest,
          sourceSize: lossyCompression.sourceSize,
          preset: lossyCompression.preset,
          targetDpi: lossyCompression.targetDpi,
          jpegQuality: lossyCompression.jpegQuality,
          pageCount: lossyCompression.pageCount,
          transformedObjectCount: lossyCompression.transformedObjectCount,
          nativePlanDigest: lossyCompression.nativePlanDigest,
          evidenceDigest: lossyCompression.evidenceDigest,
        }
      : null,
    historyReplay,
    historyReplayRequested,
    edit: {
      editRevision: edit.editRevision,
      hasEdits: edit.hasEdits,
      nativeDurableReceipt: edit.nativeDurableReceipt,
      lossDigest: edit.lossReport.digest,
    },
  });
  return Object.freeze({
    host,
    action,
    featureId,
    requestedDisposition,
    documentReady: value.documentReady === true,
    editable: value.editable === true,
    source,
    historyReplay,
    historyReplayRequested,
    heldWrites,
    replacementSource,
    pagePackage,
    deskewRaster,
    pageBackground,
    embeddedImageOverlay,
    lossyCompression,
    edit,
    requestKey,
  });
}

function profileAllowsMutation(profile: PdfSourceRiskProfile): boolean {
  return RISK_KEYS.every((key) => profile[key] === "none");
}

/** A file taken in by Replace Source is opened as it is, as an imported one
 * is: its tags, layers, attachments or signatures are not rewritten by the
 * swap. It only has to be readable: not encrypted (a file the inspection
 * could not open at all reads as encryption "unknown"). A part of its
 * catalog the inspection could not read is background, not a refusal. */
function replacementSourceUsable(profile: PdfSourceRiskProfile): boolean {
  return profile.encryption === "none";
}

/** Annotation and outline writers: they add or change their own objects
 * and leave the page tree, Tagged-PDF structure, layers and attachments as
 * they stand, so those package siblings keep none of them read-only, as WPS
 * adds a link or a comment to any such PDF. They are judged as the native
 * writers are: an approval signature stands over the revision it signed,
 * a certification forbids. A whole-package rewrite (flattened download,
 * compress, merge) is not one of them. */
const PDF_ANNOTATION_WRITER_ACTIONS: ReadonlySet<PdfSavePolicyAction> = new Set<PdfSavePolicyAction>([
  "pdf.text.free-text.commit",
  "pdf.image.insert.commit",
  "pdf.link.commit",
  "pdf.link.edit.commit",
  "pdf.outline.commit",
  "pdf.signature.visual.commit",
  "pdf.comment.highlight.commit",
  "pdf.comment.area-highlight.commit",
  "pdf.comment.underline.commit",
  "pdf.comment.strikeout.commit",
  "pdf.comment.squiggly.commit",
  "pdf.comment.proofing.commit",
  "pdf.comment.note.commit",
  "pdf.comment.batch-delete",
  "pdf.comment.ink.commit",
  "pdf.comment.shape.commit",
  "pdf.comment.stamp.commit",
  "pdf.comment.file-attachment.commit",
  "pdf.comment.measure.commit",
]);










export const PDF_PAGE_REVISION_ACTIONS: ReadonlySet<PdfSavePolicyAction> = new Set<PdfSavePolicyAction>([
  "pdf.page.insert-blank",
  "pdf.page.delete",
  "pdf.page.reorder",
  "pdf.page.organize",
  "pdf.page.rotate",
  "pdf.page.crop",
  "pdf.page.split",
  "pdf.page.resize-media-box",
  "pdf.page.labels.commit",
  "pdf.decoration.watermark.commit",
  "pdf.decoration.page-number.commit",
  "pdf.decoration.header-footer.commit",
  // A page import or replacement is a page revision like a blank page: the
  // page writer appends or swaps pages and keeps the catalog (a tagged PDF
  // takes the new pages untagged, as WPS inserts them). Judged by the
  // whole-package rule, 导入文件 was refused on every tagged, layered or
  // attachment-carrying PDF as 「未核验的来源风险」 while 插入空白页 on the
  
  "pdf.page.import-native",
  "pdf.page.replace-native",
  // 提取当前页 / 拆分并下载 write a new file from copied pages and leave
  // the current one as it is; their loss report already names what a copy
  // drops (document-level structures). Judged by the whole-package rule,
  // a tagged PDF was refused as 「未能确认安全的结构」 (page-tab audit,
  
  "pdf.page.extract",
  // 纯图片 PDF rasterizes every page into a new file and leaves the current
  // one as it is; its consent names what the raster drops. Judged by the
  // whole-package rule it failed on every tagged file with only 「无法创建
  // 图片型 PDF」 to show (user audit PDF 0140, W3C table; stress-v2 here).
  "pdf.file.image-only",
]);








const PDF_PAGE_COPY_FILE_ACTIONS: ReadonlySet<PdfSavePolicyAction> = new Set<PdfSavePolicyAction>([
  "pdf.file.split",
  "pdf.file.merge",
]);
/** Compression rewrites the file whole but keeps its catalog: tags, layers
 * and attachments are carried over as they are (pruned only of objects
 * nothing reaches). A signature, approval or certification, would not
 * survive the rewrite: those still refuse. */
const PDF_STRUCTURE_KEEPING_REWRITE_ACTIONS: ReadonlySet<PdfSavePolicyAction> = new Set<PdfSavePolicyAction>([
  "pdf.file.compress-lossless",
  "pdf.file.compress-lossy",
]);

/** What a page copy of these sources leaves behind, named for the consent:
 * the tags and document structure, layers, attachments, a PDF/A claim. */
export function pdfSourceStructureLossEntries(
  profiles: readonly PdfSourceRiskProfile[],
): PdfOutputLossEntry[] {
  const has = (key: "taggedStructure" | "optionalContent" | "attachments" | "pdfaClaim") =>
    profiles.some((profile) => profile[key] !== "none");
  return [
    ...(has("taggedStructure") ? [{
      code: "PDF_LOSS_TAGGED_STRUCTURE",
      label: "Tags and the document structure (the reading order screen readers use) are not kept in the new file.",
    }] : []),
    ...(has("optionalContent") ? [{
      code: "PDF_LOSS_OPTIONAL_CONTENT_LAYERS",
      label: "Layer settings are not kept: content on a hidden layer can show in the new file.",
    }] : []),
    ...(has("attachments") ? [{
      code: "PDF_LOSS_EMBEDDED_ATTACHMENTS",
      label: "Files attached to the PDF are not kept in the new file.",
    }] : []),
    ...(has("pdfaClaim") ? [{
      code: "PDF_LOSS_PDFA_CLAIM",
      label: "The new file does not claim PDF/A conformance.",
    }] : []),
  ];
}

function profileAllowsActionMutation(
  action: PdfSavePolicyAction,
  profile: PdfSourceRiskProfile,
): boolean {
  if (PDF_PAGE_COPY_FILE_ACTIONS.has(action)) {
    return profile.encryption === "none" && profile.certifiedSignatures === "none";
  }
  if (PDF_STRUCTURE_KEEPING_REWRITE_ACTIONS.has(action)) {
    return profile.encryption === "none" && profile.certifiedSignatures === "none" &&
      profile.certificateSignatures === "none";
  }
  // Replace Source writes nothing into the current file: the page takes a
  // different file whole, so what the current one carries (tags, layers,
  // attachments) cannot be lost or broken by it. Only the replacement is
  // judged (replacementSourceUsable). A tagged or layered PDF kept its
  // Replace button disabled under 「正在替换 PDF」 (user audit PDF 0003).
  if (action === "replace-source") return true;
  if (
    !PDF_ANNOTATION_WRITER_ACTIONS.has(action) &&
    !PDF_PAGE_REVISION_ACTIONS.has(action) &&
    action !== "pdf.vector.geometry.commit" &&
    action !== "pdf.vector.style.commit" &&
    action !== "pdf.vector.duplicate.commit" &&
    action !== "pdf.vector.reorder.commit" &&
    action !== "pdf.content.append.commit" &&
    action !== "pdf.content.append-image.commit" &&
    action !== "pdf.content.append-path.commit" &&
    // A background is a page-local revision too: new streams in front of
    // each page's drawing and the pages' own dictionaries.
    action !== "pdf.page.background.commit" &&
    action !== "pdf.text.edit-content.commit" &&
    action !== "pdf.text.delete-object" &&
    action !== "pdf.text.decoration.commit" &&
    action !== "pdf.text.paragraph.justify" &&
    action !== "pdf.tagged-text.commit" &&
    action !== "pdf.image.replace-source" &&
    action !== "pdf.image.remove-source" &&
    action !== "pdf.image.transform-source" &&
    action !== "pdf.image.arrange-source"
  ) {
    return profileAllowsMutation(profile);
  }
  // Native source text/vector edits and native embedded-image replace/remove
  // own authenticated page-local content: they rewrite only the selected image
  // XObject (or its draw operator) plus the owning content stream, leaving the
  // page tree, Tagged-PDF structure, OCG membership and attachments as exact
  // source-prefix bytes. Each still requires the exact selection/source proof;
  // this is not authority for the generic FreeText or package-rewrite writers.
  // Package siblings (attachments, OCG and Tagged-PDF structure) remain exact
  // source-prefix bytes. Of the profile only encryption and a certification
  
  // XFA form's page content is written as any page's; a PDF/A claim is dropped
  
  // 「真正声明了 PDF/A 的文件与WPS一致」), as WPS and Acrobat save an edited
  // PDF/A file; an inspection that could not read some other part of the
  // catalog is background. The server's write invariants still refuse a
  // revision that declares objects it should not.
  return (
    profile.encryption === "none" &&
    // An approval signature signs the revision it was made over, and these
    // writers only append a new one: those bytes and that signature stand as
    // they were. A certification signature carries a rule against the edit.
    profile.certifiedSignatures === "none"
  );
}

function isPagePackageAction(
  action: PdfSavePolicyAction,
): action is
  | "pdf.page.insert-file"
  | "pdf.page.extract"
  | "pdf.file.split"
  | "pdf.file.merge" {
  return (
    action === "pdf.page.insert-file" ||
    action === "pdf.page.extract" ||
    action === "pdf.file.split" ||
    action === "pdf.file.merge"
  );
}

function sameSourceSnapshot(
  left: PdfSourceRiskSnapshot,
  right: PdfSourceRiskSnapshot,
): boolean {
  return (
    left.identity.dataItemId === right.identity.dataItemId &&
    left.identity.byteGenerationKey === right.identity.byteGenerationKey &&
    left.identity.sha256 === right.identity.sha256 &&
    left.identity.documentRevision === right.identity.documentRevision &&
    left.profileDigest === right.profileDigest
  );
}

function blockedReason(request: NormalizedPdfSavePolicyRequest): string | null {
  if (!request.documentReady) return "pdf.savePolicy.documentLoading";
  if (!request.source) return "pdf.savePolicy.sourceIdentityRequired";

  const disposition = request.requestedDisposition;
  if (disposition === "read-only") {
    if (request.edit.hasEdits) return "pdf.savePolicy.editsRequireRewrite";
    return null;
  }

  // Downloading an already-confirmed version does not mutate its layers,
  // tags, attachments or signatures. The host must deliver that exact version
  // and revalidate clean state; a pending rewrite keeps the mutation policy.
  const immutableNativeDownload = request.action === "download-native" &&
    !request.edit.hasEdits && request.edit.nativeDurableReceipt;
  if (!request.editable && request.host === "native-editor" && !immutableNativeDownload) {
    return "pdf.savePolicy.readOnly";
  }
  const history = request.historyReplay;
  const historyMatches = history && request.host === "native-editor" &&
    request.action === `pdf.edit.${history.direction}` &&
    history.receipt.outputSha256 === request.source.identity.sha256;
  if (request.historyReplayRequested && !historyMatches) {
    return "pdf.savePolicy.sourceIdentityRequired";
  }
  // A history replay is judged by the rule each write it replays was judged
  // by when it ran forward; without a journal lease, by the strict rule.
  const judgedActions: readonly PdfSavePolicyAction[] = historyMatches
    ? history.sourceActions
    : request.heldWrites ?? [request.action];
  const profile = request.source.profile;
  if (!immutableNativeDownload && (judgedActions.length === 0 ||
    !judgedActions.every((action) => profileAllowsActionMutation(action, profile)))) {
    return "pdf.savePolicy.riskySourceReadOnly";
  }
  if (
    request.replacementSource &&
    !replacementSourceUsable(request.replacementSource.profile)
  ) {
    return "pdf.savePolicy.riskyReplacementReadOnly";
  }
  if (isPagePackageAction(request.action)) {
    const pagePackage = request.pagePackage;
    if (!pagePackage) return "pdf.savePolicy.pagePackageContextRequired";
    if (!sameSourceSnapshot(request.source, pagePackage.sources[0].source)) {
      return "pdf.savePolicy.pagePackagePrimarySourceMismatch";
    }
    // Each source judged as the action is: a tagged, layered or
    // attachment-carrying file is copied like any other (its loss report
    // names what the copy drops); only encryption and certification refuse
    
    // 文件带有签名、加密或未能确认安全的结构」).
    if (
      pagePackage.sources.some(
        (entry) => !profileAllowsActionMutation(request.action, entry.source.profile),
      )
    ) {
      return "pdf.savePolicy.riskyAdditionalSourceReadOnly";
    }
    // Links, annotations, forms and document actions a page copy leaves
    // behind are named in the consent's loss report, not a refusal: WPS
    // extracts and splits such files, and the owner keeps no gate WPS lacks
    
    
  } else if (request.pagePackage) {
    return "pdf.savePolicy.pagePackageActionMismatch";
  }

  if (request.action === "pdf.file.deskew-page") {
    const deskew = request.deskewRaster;
    if (!deskew) return "pdf.savePolicy.deskewContextRequired";
    if (
      deskew.dataItemId !== request.source.identity.dataItemId ||
      deskew.previewByteGenerationKey !==
        request.source.identity.byteGenerationKey ||
      deskew.previewSha256 !== request.source.identity.sha256 ||
      deskew.documentRevision !== request.source.identity.documentRevision ||
      deskew.editRevision !== request.edit.editRevision
    ) {
      return "pdf.savePolicy.deskewIdentityMismatch";
    }
  } else if (request.deskewRaster) {
    return "pdf.savePolicy.deskewActionMismatch";
  }

  const isBackgroundAction =
    request.action === "pdf.page.apply-background" ||
    request.action === "pdf.page.remove-background";
  if (isBackgroundAction) {
    const background = request.pageBackground;
    if (!background) return "pdf.savePolicy.backgroundContextRequired";
    const expectedKind =
      request.action === "pdf.page.apply-background" ? "apply" : "remove";
    if (
      background.kind !== expectedKind ||
      background.dataItemId !== request.source.identity.dataItemId ||
      background.sourceUrl !== request.source.identity.byteGenerationKey ||
      background.previewByteGenerationKey !==
        request.source.identity.byteGenerationKey ||
      background.previewSha256 !== request.source.identity.sha256 ||
      background.documentRevision !==
        request.source.identity.documentRevision ||
      background.editRevision !== request.edit.editRevision
    ) {
      return "pdf.savePolicy.backgroundIdentityMismatch";
    }
  } else if (request.pageBackground) {
    return "pdf.savePolicy.backgroundActionMismatch";
  }

  const isEmbeddedImageOverlayAction =
    request.action === "pdf.image.apply-overlay-edits";
  if (isEmbeddedImageOverlayAction) {
    const overlay = request.embeddedImageOverlay;
    if (!overlay) return "pdf.savePolicy.embeddedImageContextRequired";
    if (
      overlay.dataItemId !== request.source.identity.dataItemId ||
      overlay.sourceUrl !== request.source.identity.byteGenerationKey ||
      overlay.previewByteGenerationKey !==
        request.source.identity.byteGenerationKey ||
      overlay.previewSha256 !== request.source.identity.sha256 ||
      overlay.documentRevision !== request.source.identity.documentRevision ||
      overlay.editRevision !== request.edit.editRevision
    ) {
      return "pdf.savePolicy.embeddedImageIdentityMismatch";
    }
  } else if (request.embeddedImageOverlay) {
    return "pdf.savePolicy.embeddedImageActionMismatch";
  }

  if (request.action === "pdf.file.compress-lossy") {
    const compression = request.lossyCompression;
    if (!request.editable) return "pdf.savePolicy.readOnly";
    if (!compression) {
      return "pdf.savePolicy.lossyCompressionContextRequired";
    }
    if (
      compression.previewDataItemId !== request.source.identity.dataItemId ||
      compression.previewByteGenerationKey !==
        request.source.identity.byteGenerationKey ||
      compression.previewSha256 !== request.source.identity.sha256 ||
      compression.documentRevision !==
        request.source.identity.documentRevision ||
      compression.editRevision !== request.edit.editRevision
    ) {
      return "pdf.savePolicy.lossyCompressionIdentityMismatch";
    }
  } else if (request.lossyCompression) {
    return "pdf.savePolicy.lossyCompressionActionMismatch";
  }

  if (
    request.action === "download-native" &&
    !request.edit.nativeDurableReceipt
  ) {
    return "pdf.savePolicy.nativeDurableReceiptRequired";
  }

  if (
    request.action === "replace-source" &&
    !request.replacementSource &&
    !request.edit.hasEdits
  ) {
    // The exact command may open the native file picker. It has no output side
    // effect; the later confirmation must resolve a second decision bound to
    // the selected source SHA, both profiles, revisions, and visible losses.
    return null;
  }

  const requiresConsent =
    disposition === "flatten-with-consent" ||
    request.action === "replace-source";
  if (requiresConsent && request.edit.lossReport.codes.length === 0) {
    return "pdf.savePolicy.lossManifestRequired";
  }
  if (
    disposition === "flatten-with-consent" &&
    !request.edit.hasEdits &&
    request.action !== "download-archival" &&
    request.action !== "download-properties" &&
    !isPagePackageAction(request.action) &&
    !isBackgroundAction &&
    !isEmbeddedImageOverlayAction
  ) {
    return "pdf.savePolicy.noRewriteRequired";
  }
  if (request.action === "replace-source" && !request.replacementSource) {
    return "pdf.savePolicy.replacementIdentityRequired";
  }
  return null;
}

function decisionView(
  request: NormalizedPdfSavePolicyRequest,
): PdfSavePolicyDecisionView {
  const reasonKey = blockedReason(request);
  const enabled = reasonKey === null;
  const disposition = enabled ? request.requestedDisposition : "blocked";
  const documentRevision = request.source?.identity.documentRevision ?? 0;
  return Object.freeze({
    host: request.host,
    featureId: request.featureId,
    action: request.action,
    enabled,
    disposition,
    reasonKey,
    requiresConsent:
      enabled &&
      (disposition === "flatten-with-consent" ||
        (request.action === "replace-source" &&
          request.replacementSource !== null)),
    lossReport: request.edit.lossReport,
    documentRevision,
    sourceRevision: sourceRevision(
      request.source,
      request.replacementSource,
      request.pagePackage,
      request.deskewRaster,
      request.pageBackground,
      request.embeddedImageOverlay,
      request.lossyCompression,
    ),
    editRevision: request.edit.editRevision,
  });
}

export function resolvePdfSavePolicy(
  value: PdfSavePolicyRequest,
): PdfSavePolicyDecision {
  const request = normalizeRequest(value);
  const decision = Object.freeze(Object.create(null)) as PdfSavePolicyDecision;
  issuedDecisions.set(
    decision,
    Object.freeze({
      requestKey: request.requestKey,
      view: decisionView(request),
    }),
  );
  return decision;
}

export function readPdfSavePolicyDecision(
  value: unknown,
): PdfSavePolicyDecisionView | null {
  if (value === null || typeof value !== "object") return null;
  return issuedDecisions.get(value as PdfSavePolicyDecision)?.view ?? null;
}

export function verifyPdfSavePolicyDecision(
  decision: unknown,
  expected: PdfSavePolicyRequest,
): boolean {
  if (decision === null || typeof decision !== "object") return false;
  const issued = issuedDecisions.get(decision as PdfSavePolicyDecision);
  if (!issued) return false;
  try {
    return issued.requestKey === normalizeRequest(expected).requestKey;
  } catch {
    return false;
  }
}

function state(
  operation: () => boolean,
  markFailure: () => void,
): PdfSourceRiskState {
  try {
    return operation() ? "present" : "none";
  } catch {
    markFailure();
    return "unknown";
  }
}

function hasEmbeddedFileRisk(doc: PDFDocument): boolean {
  const indirectObjects = doc.context.enumerateIndirectObjects();
  if (indirectObjects.length > MAX_RISK_GRAPH_OBJECTS) {
    throw new Error("PDF attachment inspection object budget exceeded");
  }
  const objectsByRef = new Map(
    indirectObjects.map(([ref, object]) => [ref.toString(), object]),
  );
  const stack: PDFObject[] = [
    doc.catalog,
    ...doc.getPages().map((page) => page.node),
    ...indirectObjects.map(([, object]) => object),
  ];
  const seenObjects = new Set<PDFObject>();
  const seenRefs = new Set<string>();
  let traversedEdges = 0;

  while (stack.length > 0) {
    traversedEdges += 1;
    if (traversedEdges > MAX_RISK_GRAPH_EDGES) {
      throw new Error("PDF attachment inspection edge budget exceeded");
    }
    const current = stack.pop();
    if (!current) continue;
    if (current instanceof PDFRef) {
      const key = current.toString();
      if (seenRefs.has(key)) continue;
      seenRefs.add(key);
      const resolved = objectsByRef.get(key);
      if (!resolved) throw new Error("Dangling PDF attachment reference");
      stack.push(resolved);
      continue;
    }
    if (seenObjects.has(current)) continue;
    seenObjects.add(current);
    if (current instanceof PDFInvalidObject) {
      // Read as the signature inspection reads it: an object with no bytes
      // holds no file, an undecoded cross-reference stream is its dictionary
      // (walked below like any other) over table rows. Any other opaque
      // object throws there: it may be the file specification itself.
      const unread = opaqueObjectDictionary(doc, current);
      if (unread) stack.push(unread);
      continue;
    }
    if (current instanceof PDFStream) {
      if (
        String(current.dict.get(PDFName.of("Type")) ?? "") === "/EmbeddedFile"
      ) {
        return true;
      }
      stack.push(current.dict);
      continue;
    }
    if (current instanceof PDFDict) {
      // Associated Files may live directly on Catalog, Page, annotation, or
      // another reachable dictionary without a /Names/EmbeddedFiles tree.
      if (current.get(PDFName.of("AF"))) return true;
      if (
        String(current.get(PDFName.of("Subtype")) ?? "") === "/FileAttachment"
      ) {
        return true;
      }
      const embeddedFileDictionary = current.get(PDFName.of("EF"));
      if (
        embeddedFileDictionary &&
        String(current.get(PDFName.of("Type")) ?? "") === "/Filespec"
      ) {
        return true;
      }
      stack.push(...current.values());
      continue;
    }
    if (current instanceof PDFArray) stack.push(...current.asArray());
  }
  return false;
}

function hasPdfaClaim(doc: PDFDocument): boolean {
  const metadataRef = doc.catalog.get(PDFName.of("Metadata"));
  if (!metadataRef) return false;
  const metadata = doc.context.lookup(metadataRef);
  if (!(metadata instanceof PDFRawStream)) return false;
  let bytes: Uint8Array;
  try {
    bytes = decodePDFRawStream(metadata).decode();
  } catch {
    bytes = metadata.getContents();
  }
  const text = new TextDecoder().decode(bytes);
  // A claim is a part (1-4) stated under the identification schema: in an
  // element or attribute of a prefix bound to it, or unprefixed under it as
  // the default namespace. Anything else claims nothing -- the namespace
  // named in an extension schema's list (NIST's PDF/UA packets), a prefix
  // declared and never used (PLOS, PDFlib), a part with no schema bound --
  
  const prefixes = Array.from(
    text.matchAll(
      /xmlns:([A-Za-z_][\w.-]*)\s*=\s*["']http:\/\/www\.aiim\.org\/pdfa\/ns\/id\/["']/gi,
    ),
    (match) => match[1],
  );
  const defaultIdentificationNamespace =
    /xmlns\s*=\s*["']http:\/\/www\.aiim\.org\/pdfa\/ns\/id\/["']/i.test(text);
  if (
    defaultIdentificationNamespace &&
    /<part\b[^>]*>\s*[1-4]\s*<\/part>/i.test(text)
  ) {
    return true;
  }
  return prefixes.some((prefix) => {
    const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return (
      new RegExp(`<${escaped}:part\\b[^>]*>\\s*([1-4])\\s*</${escaped}:part>`, "i").test(text) ||
      new RegExp(`\\b${escaped}:part\\s*=\\s*["']([1-4])["']`, "i").test(text)
    );
  });
}

/** The inspection itself, for the worker (source-risk.worker.ts): `ownedBytes` are its to keep. */
export function inspectPdfRiskProfileBytes(
  ownedBytes: Uint8Array,
): Promise<Readonly<PdfSourceRiskProfile>> {
  return inspectPdfRiskProfile(ownedBytes);
}

/** ``xfaDynamic``, present only when the file carries an XFA form: /NeedsRendering
 * in the catalog says the viewer draws the form from its template. */
function xfaKindAxis(
  doc: PDFDocument,
  markFailure: () => void,
): { xfaDynamic?: PdfSourceRiskState } {
  const acroForm = doc.catalog.lookupMaybe(PDFName.of("AcroForm"), PDFDict);
  if (!acroForm?.get(PDFName.of("XFA"))) return {};
  return {
    xfaDynamic: state(
      () => String(doc.catalog.get(PDFName.of("NeedsRendering"))) === "true",
      markFailure,
    ),
  };
}

async function inspectPdfRiskProfile(
  ownedBytes: Uint8Array,
): Promise<Readonly<PdfSourceRiskProfile>> {
  try {
    if (ownedBytes.byteLength > MAX_SOURCE_BYTES) {
      return createUnknownPdfSourceRiskProfile();
    }
    // (Parsed in one go in the worker this runs in: lib/pdf/pdf-lib-pacing.)
    const doc = await loadPdfForInspection(ownedBytes);
    if (!doc) {
      return freezeRiskProfile({
        encryption: "present",
        certificateSignatures: "unknown",
        certifiedSignatures: "unknown",
        xfa: "unknown",
        optionalContent: "unknown",
        attachments: "unknown",
        taggedStructure: "unknown",
        pdfaClaim: "unknown",
        inspectionFailure: "none",
      });
    }
    let failed = false;
    const markFailure = () => {
      failed = true;
    };
    const profile = freezeRiskProfile({
      encryption: "none",
      certificateSignatures: state(
        () => signatureMutationImpact(doc).willInvalidate,
        markFailure,
      ),
      certifiedSignatures: state(
        () => signatureMutationImpact(doc).breaksCertification,
        markFailure,
      ),
      xfa: state(() => {
        const acroForm = doc.catalog.lookupMaybe(
          PDFName.of("AcroForm"),
          PDFDict,
        );
        return Boolean(acroForm?.get(PDFName.of("XFA")));
      }, markFailure),
      ...xfaKindAxis(doc, markFailure),
      optionalContent: state(
        () => Boolean(doc.catalog.get(PDFName.of("OCProperties"))),
        markFailure,
      ),
      attachments: state(() => {
        const names = doc.catalog.lookupMaybe(PDFName.of("Names"), PDFDict);
        return (
          Boolean(names?.get(PDFName.of("EmbeddedFiles"))) ||
          hasEmbeddedFileRisk(doc)
        );
      }, markFailure),
      taggedStructure: state(() => {
        const markInfo = doc.catalog.lookupMaybe(
          PDFName.of("MarkInfo"),
          PDFDict,
        );
        return (
          Boolean(doc.catalog.get(PDFName.of("StructTreeRoot"))) ||
          String(markInfo?.get(PDFName.of("Marked")) ?? "") === "true"
        );
      }, markFailure),
      pdfaClaim: state(() => hasPdfaClaim(doc), markFailure),
      inspectionFailure: "none",
    });
    return failed
      ? freezeRiskProfile({ ...profile, inspectionFailure: "present" })
      : profile;
  } catch {
    return createUnknownPdfSourceRiskProfile();
  } finally {
    ownedBytes.fill(0);
  }
}

async function sha256Hex(
  bytes: Uint8Array,
  authenticatedSha256?: string | null,
): Promise<string> {
  const trusted =
    typeof authenticatedSha256 === "string" &&
    /^[0-9a-f]{64}$/u.test(authenticatedSha256.trim().toLowerCase())
      ? authenticatedSha256.trim().toLowerCase()
      : null;
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    if (trusted) return trusted;
    throw new Error("SHA-256 is unavailable");
  }
  const own = new Uint8Array(bytes.byteLength);
  own.set(bytes);
  try {
    const digest = await subtle.digest("SHA-256", own);
    const computed = Array.from(new Uint8Array(digest), (byte) =>
      byte.toString(16).padStart(2, "0"),
    ).join("");
    if (trusted && computed !== trusted) {
      throw new Error("Authenticated PDF source digest mismatch");
    }
    return computed;
  } finally {
    own.fill(0);
  }
}

/**
 * A store of profiles shared across readers (lib/pdf/source-risk-store.ts):
 * asked by digest before a file is inspected, told what had to be inspected.
 * None installed (a worker, a test): every file is inspected here.
 */
export interface PdfSourceRiskProfileStore {
  read(sha256: string): Promise<PdfSourceRiskProfile | null>;
  write(sha256: string, profile: PdfSourceRiskProfile): void;
}

let sourceRiskProfileStore: PdfSourceRiskProfileStore | null = null;

export function setPdfSourceRiskProfileStore(store: PdfSourceRiskProfileStore | null): void {
  sourceRiskProfileStore = store;
}

function inspectOwnedBytes(owned: Uint8Array): Promise<Readonly<PdfSourceRiskProfile>> {
  // Off the main thread where a worker is to be had (source-risk-offthread):
  // a long file's parse and graph walk held the reader's thread for seconds.
  return inspectPdfRiskProfileOffThread(owned).then(
    (profile) => profile ?? inspectPdfRiskProfile(owned),
    () => inspectPdfRiskProfile(owned),
  );
}

function cacheRiskProfile(
  sha256: string,
  bytes: Uint8Array,
): Promise<Readonly<PdfSourceRiskProfile>> {
  const cached = riskProfileCache.get(sha256);
  if (cached) {
    riskProfileCache.delete(sha256);
    riskProfileCache.set(sha256, cached);
    return cached;
  }
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  // The shared store first (one clean answer per file for every reader,
  
  const store = sourceRiskProfileStore;
  // (The worker is started beside the store's answer, not after it: a file
  // the store does not hold is inspected the moment that is known.)
  if (store) warmPdfSourceRiskWorker();
  const inspection = store
    ? store.read(sha256).catch(() => null).then((stored) => {
        if (stored) return stored;
        return inspectOwnedBytes(owned).then((profile) => {
          try { store.write(sha256, profile); } catch { /* the store is a cache */ }
          return profile;
        });
      })
    : inspectOwnedBytes(owned);
  riskProfileCache.set(sha256, inspection);
  while (riskProfileCache.size > MAX_RISK_CACHE_ENTRIES) {
    const oldest = riskProfileCache.keys().next().value as string | undefined;
    if (!oldest) break;
    riskProfileCache.delete(oldest);
  }
  return inspection;
}

/**
 * Start the inspection of `bytes` now, kept by their own digest: the reader
 * that inspects the same bytes later (`inspectPdfSourceRiskSnapshot`) finds
 * it under way or done (lib/pdf/source-risk-ahead). The digest is computed
 * here over these very bytes, never taken from anywhere else.
 */
export async function inspectPdfSourceRiskProfileAhead(bytes: Uint8Array): Promise<void> {
  // Copied before the first await: the caller may zero or reuse its bytes
  // as soon as this returns, and what is inspected must be what was digested.
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  try {
    const sha256 = await sha256Hex(owned);
    // (`cacheRiskProfile` takes its own copy before it returns.)
    void cacheRiskProfile(sha256, owned).catch(() => undefined);
  } finally {
    owned.fill(0);
  }
}

const carriedRiskProfiles = new WeakSet<Promise<Readonly<PdfSourceRiskProfile>>>();
const SOURCE_SHA256_PATTERN = /^[0-9a-f]{64}$/u;

/**
 * Carry the risk profile this client already inspected for one source over
 * to bytes the server derived from it without changing any object: a native
 * source consolidation writes every object once, under its own number, with
 * the bytes the newest revision holds, so every axis of the profile
 * (encryption, signatures, XFA, catalog features, inspection outcome) reads
 * the same objects. Only a profile already inspected here for `fromSha256`
 * is carried, only when it inspected cleanly and `admit` accepts it, and
 * only under `toSha256`: a snapshot still binds the digest of the bytes it
 * is given, so no other bytes can reach the carried profile.
 *
 * "carried" when the profile now answers for `toSha256`, "present" when
 * `toSha256` had its own entry already (nothing replaced), null when there
 * was nothing to carry (the caller inspects the new bytes in full).
 */
export async function carryPdfSourceRiskProfile(input: {
  readonly fromSha256: string;
  readonly toSha256: string;
  readonly admit?: (profile: Readonly<PdfSourceRiskProfile>) => boolean;
}): Promise<"carried" | "present" | null> {
  const from = typeof input.fromSha256 === "string" ? input.fromSha256.toLowerCase() : "";
  const to = typeof input.toSha256 === "string" ? input.toSha256.toLowerCase() : "";
  if (!SOURCE_SHA256_PATTERN.test(from) || !SOURCE_SHA256_PATTERN.test(to) || from === to) {
    return null;
  }
  const inspected = riskProfileCache.get(from);
  if (!inspected) return null;
  let profile: Readonly<PdfSourceRiskProfile>;
  try {
    profile = await inspected;
  } catch {
    return null;
  }
  if (profile.inspectionFailure !== "none" || (input.admit && !input.admit(profile))) {
    return null;
  }
  if (riskProfileCache.has(to)) return "present";
  const carried = Promise.resolve(profile);
  carriedRiskProfiles.add(carried);
  riskProfileCache.set(to, carried);
  while (riskProfileCache.size > MAX_RISK_CACHE_ENTRIES) {
    const oldest = riskProfileCache.keys().next().value as string | undefined;
    if (!oldest) break;
    riskProfileCache.delete(oldest);
  }
  return "carried";
}

/** Drop a profile carryPdfSourceRiskProfile carried to `sha256` (never one
 * inspected from bytes), so the next snapshot inspects in full. */
export function forgetCarriedPdfSourceRiskProfile(sha256: string): void {
  const key = typeof sha256 === "string" ? sha256.toLowerCase() : "";
  const entry = riskProfileCache.get(key);
  if (entry && carriedRiskProfiles.has(entry)) riskProfileCache.delete(key);
}

export async function inspectPdfSourceRiskSnapshot(options: {
  readonly dataItemId: string;
  readonly byteGenerationKey: string;
  readonly bytes: Uint8Array | ArrayBuffer;
  readonly documentRevision: PdfSavePolicyRevision;
  /** SHA from the authenticated DataItem metadata for WebCrypto fallback and verification. */
  readonly authenticatedSha256?: string | null;
  





  readonly computedSha256?: string | null;
}): Promise<PdfSourceRiskSnapshot> {
  const view =
    options.bytes instanceof Uint8Array
      ? options.bytes
      : new Uint8Array(options.bytes);
  const computed = typeof options.computedSha256 === "string" &&
    /^[0-9a-f]{64}$/u.test(options.computedSha256.trim().toLowerCase())
    ? options.computedSha256.trim().toLowerCase() : null;
  const sha256 = computed ?? await sha256Hex(view, options.authenticatedSha256);
  const profile = await cacheRiskProfile(sha256, view);
  return createPdfSourceRiskSnapshot({
    identity: {
      dataItemId: options.dataItemId,
      byteGenerationKey: options.byteGenerationKey,
      sha256,
      documentRevision: options.documentRevision,
    },
    profile,
  });
}

async function inspectPdfPagePackageStructure(
  bytes: Uint8Array,
): Promise<{ pageCount: number; structureRisk: PdfSourceRiskState }> {
  const owned = new Uint8Array(bytes.byteLength);
  owned.set(bytes);
  try {
    const doc = await loadPdfForInspection(owned);
    if (!doc) {
      return { pageCount: 1, structureRisk: "unknown" };
    }
    const pages = doc.getPages();
    if (pages.length < 1 || pages.length > MAX_PAGE_PACKAGE_PAGES) {
      throw new Error("Invalid PDF page-package page count");
    }
    // An entry only counts when it holds something: many producers write an
    // empty /Annots [] on every page (or an empty /AA or /Names), and such a
    // PDF has no annotation, action or name tree to lose by page copying.
    const present = (value: PDFObject | undefined): boolean => {
      const resolved =
        value instanceof PDFRef ? doc.context.lookup(value) : value;
      if (resolved instanceof PDFArray) return resolved.size() > 0;
      if (resolved instanceof PDFDict) return resolved.keys().length > 0;
      return resolved !== undefined && resolved !== PDFNull;
    };
    const catalogRisk =
      present(doc.catalog.get(PDFName.of("AcroForm"))) ||
      present(doc.catalog.get(PDFName.of("OpenAction"))) ||
      present(doc.catalog.get(PDFName.of("AA"))) ||
      present(doc.catalog.get(PDFName.of("Dests"))) ||
      present(doc.catalog.get(PDFName.of("Names")));
    const pageRisk = pages.some(
      (page) =>
        present(page.node.get(PDFName.of("Annots"))) ||
        present(page.node.get(PDFName.of("AA"))),
    );
    return {
      pageCount: pages.length,
      structureRisk: catalogRisk || pageRisk ? "present" : "none",
    };
  } finally {
    owned.fill(0);
  }
}

export async function inspectPdfPagePackageSource(options: {
  readonly dataItemId: string;
  readonly byteGenerationKey: string;
  readonly bytes: Uint8Array | ArrayBuffer;
  readonly documentRevision: PdfSavePolicyRevision;
}): Promise<PdfPagePackageSourceSnapshot> {
  const view =
    options.bytes instanceof Uint8Array
      ? options.bytes
      : new Uint8Array(options.bytes);
  if (view.byteLength < 1 || view.byteLength > MAX_SOURCE_BYTES) {
    throw new Error("PDF page-package source exceeds byte limit");
  }
  const [source, structure] = await Promise.all([
    inspectPdfSourceRiskSnapshot(options),
    inspectPdfPagePackageStructure(view),
  ]);
  return normalizePagePackageSource(
    {
      source,
      byteLength: view.byteLength,
      pageCount: structure.pageCount,
      structureRisk: structure.structureRisk,
    },
    0,
  );
}

export async function validatePdfPagePackageOutputs(
  outputs: readonly {
    readonly bytes: Uint8Array;
    readonly expectedPageCount: number;
  }[],
): Promise<void> {
  if (
    !Array.isArray(outputs) ||
    outputs.length < 1 ||
    outputs.length > MAX_PAGE_PACKAGE_OUTPUTS
  ) {
    throw new Error("Invalid PDF page-package output count");
  }
  let aggregateBytes = 0;
  let aggregatePages = 0;
  for (let index = 0; index < outputs.length; index += 1) {
    if (!(index in outputs)) throw new Error("Sparse PDF page-package outputs");
    const output = outputs[index];
    if (!(output.bytes instanceof Uint8Array) || output.bytes.byteLength < 1) {
      throw new Error(`Invalid PDF page-package output ${index}`);
    }
    aggregateBytes += output.bytes.byteLength;
    if (aggregateBytes > MAX_PAGE_PACKAGE_AGGREGATE_BYTES) {
      throw new Error("PDF page-package output budget exceeded");
    }
    const expectedPageCount = exactPositiveSafeInteger(
      output.expectedPageCount,
      `PDF page-package expected page count ${index}`,
    );
    aggregatePages += expectedPageCount;
    if (aggregatePages > MAX_PAGE_PACKAGE_PAGES) {
      throw new Error("PDF page-package output page budget exceeded");
    }
    const owned = new Uint8Array(output.bytes.byteLength);
    owned.set(output.bytes);
    try {
      const reopened = await loadPdfForInspection(owned);
      if (
        !reopened ||
        reopened.getPageCount() !== expectedPageCount
      ) {
        throw new Error(`PDF page-package output ${index} failed validation`);
      }
    } finally {
      owned.fill(0);
    }
  }
}
