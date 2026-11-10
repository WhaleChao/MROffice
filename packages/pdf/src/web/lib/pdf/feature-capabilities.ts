/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PDF capability catalog (PDF.WP.P0.01).
 *
 * This is a deliberately conservative inventory, not a claim that every row
 * is implemented.  The 155 canonical IDs and their audited baseline status
 * are closed vocabulary.  A row is executable only when it is explicitly
 * runtime-bindable, the host supplies a live handler, every required flag is
 * enabled, and the PDF capability adapter receives an explicit disposition.
 * Unknown, catalog-only, unbound, or malformed rows stay inert.
 */

export const PDF_CAPABILITY_CATALOG_SCHEMA_VERSION =
  "mrai.pdf-capabilities.v1" as const;

export const PDF_CANONICAL_BASELINE = [
  ["PDF-FIL-DOC-01", "PARTIAL"],
  ["PDF-FIL-DOC-02", "PARTIAL"],
  ["PDF-FIL-DOC-03", "PARTIAL"],
  ["PDF-FIL-DOC-04", "PARTIAL"],
  ["PDF-FIL-DOC-05", "PARTIAL"],
  ["PDF-FIL-DOC-06", "PARTIAL"],
  ["PDF-FIL-DOC-07", "PARTIAL"],
  ["PDF-FIL-DOC-08", "PARTIAL"],
  ["PDF-FIL-DOC-09", "PARTIAL"],
  ["PDF-FIL-DOC-10", "PARTIAL"],
  ["PDF-FIL-DOC-11", "PARTIAL"],
  ["PDF-FIL-DOC-12", "PARTIAL"],
  ["PDF-HOM-NAV-01", "PARTIAL"],
  ["PDF-HOM-NAV-02", "PARTIAL"],
  ["PDF-HOM-NAV-03", "PARTIAL"],
  ["PDF-HOM-NAV-04", "PARTIAL"],
  ["PDF-HOM-NAV-05", "PARTIAL"],
  ["PDF-HOM-NAV-06", "PARTIAL"],
  ["PDF-HOM-NAV-07", "PARTIAL"],
  ["PDF-HOM-NAV-08", "PARTIAL"],
  ["PDF-VIE-ZOOM-01", "PARTIAL"],
  ["PDF-VIE-ZOOM-02", "PARTIAL"],
  ["PDF-VIE-ZOOM-03", "PARTIAL"],
  ["PDF-VIE-LAYOUT-01", "PARTIAL"],
  ["PDF-VIE-LAYOUT-02", "PARTIAL"],
  ["PDF-VIE-LAYOUT-03", "PARTIAL"],
  ["PDF-VIE-LAYOUT-04", "PARTIAL"],
  ["PDF-VIE-MODE-01", "PARTIAL"],
  ["PDF-VIE-MODE-02", "PARTIAL"],
  ["PDF-VIE-MODE-03", "PARTIAL"],
  ["PDF-VIE-MODE-04", "PARTIAL"],
  ["PDF-VIE-MODE-05", "PARTIAL"],
  ["PDF-INS-PAGE-01", "PARTIAL"],
  ["PDF-INS-PAGE-02", "PARTIAL"],
  ["PDF-INS-DECOR-01", "PARTIAL"],
  ["PDF-INS-DECOR-02", "PARTIAL"],
  ["PDF-INS-DECOR-03", "PARTIAL"],
  ["PDF-INS-DECOR-04", "PARTIAL"],
  ["PDF-INS-CONTENT-01", "PARTIAL"],
  ["PDF-INS-CONTENT-02", "PARTIAL"],
  ["PDF-INS-CONTENT-03", "PARTIAL"],
  ["PDF-INS-CONTENT-04", "PARTIAL"],
  ["PDF-INS-CONTENT-05", "PARTIAL"],
  ["PDF-EDT-TEXT-01", "PARTIAL"],
  ["PDF-EDT-TEXT-02", "PARTIAL"],
  ["PDF-EDT-TEXT-03", "PARTIAL"],
  ["PDF-EDT-TEXT-04", "PARTIAL"],
  ["PDF-EDT-TEXT-05", "PARTIAL"],
  ["PDF-EDT-IMAGE-01", "PARTIAL"],
  ["PDF-EDT-IMAGE-02", "PARTIAL"],
  ["PDF-EDT-IMAGE-03", "PARTIAL"],
  ["PDF-EDT-OBJECT-01", "PARTIAL"],
  ["PDF-EDT-OBJECT-02", "PARTIAL"],
  ["PDF-EDT-OBJECT-03", "PARTIAL"],
  ["PDF-PAG-STRUCT-01", "PARTIAL"],
  ["PDF-PAG-STRUCT-02", "PARTIAL"],
  ["PDF-PAG-STRUCT-03", "PARTIAL"],
  ["PDF-PAG-STRUCT-04", "PARTIAL"],
  ["PDF-PAG-STRUCT-05", "PARTIAL"],
  ["PDF-PAG-STRUCT-06", "PARTIAL"],
  ["PDF-PAG-STRUCT-07", "PARTIAL"],
  ["PDF-PAG-STRUCT-08", "PARTIAL"],
  ["PDF-PAG-STRUCT-09", "PARTIAL"],
  ["PDF-PAG-STRUCT-10", "PARTIAL"],
  ["PDF-PAG-STRUCT-11", "PARTIAL"],
  ["PDF-PAG-STRUCT-12", "PARTIAL"],
  ["PDF-CMT-MARK-01", "PARTIAL"],
  ["PDF-CMT-MARK-02", "PARTIAL"],
  ["PDF-CMT-MARK-03", "PARTIAL"],
  ["PDF-CMT-MARK-04", "PARTIAL"],
  ["PDF-CMT-MARK-05", "PARTIAL"],
  ["PDF-CMT-MARK-06", "PARTIAL"],
  ["PDF-CMT-NOTE-01", "PARTIAL"],
  ["PDF-CMT-NOTE-02", "PARTIAL"],
  ["PDF-CMT-NOTE-03", "PARTIAL"],
  ["PDF-CMT-DRAW-01", "PARTIAL"],
  ["PDF-CMT-DRAW-02", "PARTIAL"],
  ["PDF-CMT-DRAW-03", "PARTIAL"],
  ["PDF-CMT-DRAW-04", "PARTIAL"],
  ["PDF-CMT-DRAW-05", "PARTIAL"],
  ["PDF-CMT-DRAW-06", "PARTIAL"],
  ["PDF-CMT-DRAW-07", "PARTIAL"],
  ["PDF-FRM-FILL-01", "PARTIAL"],
  ["PDF-FRM-FILL-02", "PARTIAL"],
  ["PDF-FRM-FILL-03", "PARTIAL"],
  ["PDF-FRM-FILL-04", "PARTIAL"],
  ["PDF-FRM-AUTH-01", "PARTIAL"],
  ["PDF-FRM-AUTH-02", "PARTIAL"],
  ["PDF-FRM-AUTH-03", "PARTIAL"],
  ["PDF-FRM-AUTH-04", "PARTIAL"],
  ["PDF-SGN-ELEC-01", "PARTIAL"],
  ["PDF-SGN-ELEC-02", "PARTIAL"],
  ["PDF-SGN-ELEC-03", "PARTIAL"],
  ["PDF-SGN-DIGI-01", "PARTIAL"],
  ["PDF-SGN-DIGI-02", "MISSING"],
  ["PDF-SGN-DIGI-03", "PARTIAL"],
  ["PDF-SEC-PASS-01", "PARTIAL"],
  ["PDF-SEC-PASS-02", "PARTIAL"],
  ["PDF-SEC-PASS-03", "PARTIAL"],
  ["PDF-SEC-PASS-04", "PARTIAL"],
  ["PDF-SEC-SAFE-01", "PARTIAL"],
  ["PDF-SEC-SAFE-02", "PARTIAL"],
  ["PDF-SEC-REDACT-01", "PARTIAL"],
  ["PDF-SEC-REDACT-02", "PARTIAL"],
  ["PDF-SEC-REDACT-03", "PARTIAL"],
  ["PDF-SEC-REDACT-04", "PARTIAL"],
  ["PDF-TOO-OCR-01", "PARTIAL"],
  ["PDF-TOO-OCR-02", "PARTIAL"],
  ["PDF-TOO-OCR-03", "PARTIAL"],
  ["PDF-TOO-CONV-01", "PARTIAL"],
  ["PDF-TOO-CONV-02", "PARTIAL"],
  ["PDF-TOO-CONV-03", "PARTIAL"],
  ["PDF-TOO-CONV-04", "PARTIAL"],
  ["PDF-TOO-CONV-05", "PARTIAL"],
  ["PDF-TOO-CONV-06", "PARTIAL"],
  ["PDF-TOO-CONV-07", "PARTIAL"],
  ["PDF-TOO-CONV-08", "PARTIAL"],
  ["PDF-TOO-CONV-09", "PARTIAL"],
  ["PDF-TOO-CONV-10", "PARTIAL"],
  ["PDF-TOO-CONV-11", "PARTIAL"],
  ["PDF-TOO-CONV-12", "PARTIAL"],
  ["PDF-TOO-FILE-01", "PARTIAL"],
  ["PDF-TOO-FILE-02", "PARTIAL"],
  ["PDF-TOO-FILE-03", "PARTIAL"],
  ["PDF-TOO-FILE-04", "PARTIAL"],
  ["PDF-TOO-AI-01", "PARTIAL"],
  ["PDF-TOO-AI-02", "PARTIAL"],
  ["PDF-TOO-AI-03", "PARTIAL"],
  ["PDF-TOO-AI-04", "PARTIAL"],
  ["PDF-TOO-AI-05", "PARTIAL"],
  ["PDF-STD-PDFA-01", "PARTIAL"],
  ["PDF-STD-PDFA-02", "PARTIAL"],
  ["PDF-STD-A11Y-01", "PARTIAL"],
  ["PDF-STD-A11Y-02", "PARTIAL"],
  ["PDF-STD-INTEROP-01", "PARTIAL"],
  ["PDF-STD-INTEROP-02", "PARTIAL"],
  ["PDF-STD-INTEROP-03", "PARTIAL"],
  ["PDF-STD-INTEROP-04", "PARTIAL"],
  ["PDF-STD-INTEROP-05", "PARTIAL"],
  ["PDF-STD-INTEROP-06", "PARTIAL"],
  ["PDF-STD-INTEROP-07", "PARTIAL"],
  ["PDF-STD-INTEROP-08", "PARTIAL"],
  ["PDF-STD-INTEROP-09", "PARTIAL"],
  ["PDF-STD-INTEROP-10", "PARTIAL"],
  ["PDF-ARC-UI-01", "PARTIAL"],
  ["PDF-ARC-UI-02", "PARTIAL"],
  ["PDF-ARC-UI-03", "PARTIAL"],
  ["PDF-ARC-UI-04", "PARTIAL"],
  ["PDF-ARC-UI-05", "PARTIAL"],
  ["PDF-ARC-SAVE-01", "PARTIAL"],
  ["PDF-ARC-SAVE-02", "PARTIAL"],
  ["PDF-ARC-SAVE-03", "PARTIAL"],
  ["PDF-ARC-SAVE-04", "PARTIAL"],
  ["PDF-ARC-SAVE-05", "PARTIAL"],
  ["PDF-ARC-SAVE-06", "PARTIAL"],
] as const;

export type PdfCanonicalId = (typeof PDF_CANONICAL_BASELINE)[number][0];
export type PdfCanonicalStatus = (typeof PDF_CANONICAL_BASELINE)[number][1];

export const PDF_P0_01_CANONICAL_IDS = Object.freeze([
  "PDF-ARC-UI-01",
  "PDF-ARC-UI-02",
  "PDF-ARC-UI-04",
  "PDF-ARC-SAVE-03",
  "PDF-ARC-SAVE-06",
] as const satisfies readonly PdfCanonicalId[]);

export const PDF_P0_01_CONTROL_IDS = Object.freeze([
  "PDFCTL.RIBBON.HOME.TAB",
  "PDFCTL.RIBBON.INSERT.TAB",
  "PDFCTL.RIBBON.EDIT.TAB",
  "PDFCTL.RIBBON.PAGE.TAB",
  "PDFCTL.RIBBON.COMMENT.TAB",
  "PDFCTL.RIBBON.FEATURED.TAB",
  "PDFCTL.MENU.APPLICATION.ROOT",
  "PDFCTL.MENU.APPLICATION.ABOUT",
  "PDFCTL.MENU.APPLICATION.ADD_EXTENSION",
  "PDFCTL.MENU.APPLICATION.SERVICES",
  "PDFCTL.DIALOG.IMAGE_EXPORT.MEMBERSHIP",
  "PDFCTL.DIALOG.OCR.MEMBERSHIP",
  "PDFCTL.RIBBON.EDIT.EDIT_CONTENT",
  "PDFCTL.RIBBON.HOME.EDIT_CONTENT",
  "PDFCTL.CONTEXT.PICTURE.DELETE_SOURCE",
  "PDFCTL.CONTEXT.PICTURE.REPLACE_SOURCE",
  "PDFCTL.CONTEXT.PICTURE.TRANSFORM_SOURCE",
  "PDFCTL.CONTEXT.PICTURE.ARRANGE_SOURCE",
  "PDFCTL.CONTEXT.TEXT.DELETE_OBJECT",
  "PDFCTL.CONTEXT.TEXT.DECORATION",
  "PDFCTL.PAGE.VECTOR.GEOMETRY",
  "PDFCTL.PAGE.VECTOR.STYLE",
  "PDFCTL.PAGE.CONTENT.APPEND_TEXT",
  "PDFCTL.PAGE.CONTENT.APPEND_IMAGE",
  "PDFCTL.PAGE.CONTENT.APPEND_PATH",
  "PDFCTL.PAGE.STRUCTURE.SPLIT_PAGE",
  "PDFCTL.SURFACE.BACKGROUND",
  "PDFCTL.OUTLINE.MANAGED_LIFECYCLE",
  "PDFCTL.DIALOG.SECURITY.SAFE_COPY",
] as const);

export const PDF_P0_01_UNKNOWN_CONTROL_IDS = Object.freeze([
  "PDFCTL.MENU.APPLICATION.ADD_EXTENSION",
  "PDFCTL.MENU.APPLICATION.SERVICES",
] as const);

export const PDF_P0_01_DEBT_IDS = Object.freeze([
  "PDFDEBT.MENU.APPLICATION.SERVICES.MEMBERS",
] as const);

export const PDF_FEATURE_FLAG_REGISTRY = Object.freeze([
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_RENDER_LIFECYCLE_V2",
    defaultEnabled: true,
    owner: "PDF Platform",
    reviewBy: "2026-09-30",
    metric: "pdf.render.cancel_release_failure_rate",
    rollback: "Set NEXT_PUBLIC_PDF_RENDER_LIFECYCLE_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_OVERLAY_COORDINATES_V2",
    defaultEnabled: true,
    owner: "PDF Editing",
    reviewBy: "2026-09-30",
    metric: "pdf.overlay.coordinate_regression_rate",
    rollback: "Set NEXT_PUBLIC_PDF_OVERLAY_COORDINATES_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_PAGE_WINDOW_MUTATION_GUARD_V2",
    defaultEnabled: true,
    owner: "PDF Security",
    reviewBy: "2026-09-30",
    metric: "pdf.encrypted_mutation_block_rate",
    rollback: "Set NEXT_PUBLIC_PDF_PAGE_WINDOW_MUTATION_GUARD_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_REDACTION_REACHABILITY_V2",
    defaultEnabled: true,
    owner: "PDF Security",
    reviewBy: "2026-09-30",
    metric: "pdf.redaction_reachability_failure_rate",
    rollback: "Set NEXT_PUBLIC_PDF_REDACTION_REACHABILITY_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_ACROFORM_MUTATION_GUARD_V2",
    defaultEnabled: true,
    owner: "PDF Forms",
    reviewBy: "2026-09-30",
    metric: "pdf.acroform_guard_block_rate",
    rollback: "Set NEXT_PUBLIC_PDF_ACROFORM_MUTATION_GUARD_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_TEXT_EDIT_STABILITY_V2",
    defaultEnabled: true,
    owner: "PDF Editing",
    reviewBy: "2026-09-30",
    metric: "pdf.text_edit_stability_rollback_rate",
    rollback: "Set NEXT_PUBLIC_PDF_TEXT_EDIT_STABILITY_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_FREETEXT_RD_V2",
    defaultEnabled: true,
    owner: "PDF Interop",
    reviewBy: "2026-09-30",
    metric: "pdf.freetext_rd_roundtrip_failure_rate",
    rollback: "Set NEXT_PUBLIC_PDF_FREETEXT_RD_V2=0.",
  }),
  Object.freeze({
    id: "NEXT_PUBLIC_PDF_SAVE_POLICY_V1",
    defaultEnabled: true,
    owner: "PDF Platform",
    reviewBy: "2026-09-30",
    metric: "pdf.save_policy_block_and_failure_rate",
    rollback: "Set NEXT_PUBLIC_PDF_SAVE_POLICY_V1=0.",
  }),
] as const);

export type PdfFeatureFlagId = (typeof PDF_FEATURE_FLAG_REGISTRY)[number]["id"];
export type PdfCapabilityDomain =
  | "file"
  | "home"
  | "view"
  | "insert"
  | "edit"
  | "page"
  | "comment"
  | "forms"
  | "signatures"
  | "security"
  | "tools"
  | "standards"
  | "architecture";
export type PdfCapabilityHost = "reader" | "native-editor";
export type PdfCapabilityAvailabilityPredicate =
  | "catalog-only"
  | "document-ready"
  | "document-writable";
export type PdfOutputDisposition =
  | "native"
  | "flatten-with-consent"
  | "read-only"
  | "blocked";

export interface PdfCapabilityEntryPoint {
  readonly controlId: string;
  readonly surface:
    | "ribbon-tab"
    | "ribbon-control"
    | "application-menu"
    | "dialog-membership";
  readonly evidence: "OBSERVED" | "UNKNOWN";
  readonly availability: "not-implemented" | "inert-unknown";
  readonly reasonKey: string;
}

export interface PdfFeatureCapability {
  readonly id: PdfCanonicalId;
  readonly reasonKey: string;
  readonly baselineStatus: PdfCanonicalStatus;
  readonly domain: PdfCapabilityDomain;
  readonly entryPoints: readonly PdfCapabilityEntryPoint[];
  readonly contexts: readonly PdfCapabilityHost[];
  readonly availabilityPredicate: PdfCapabilityAvailabilityPredicate;
  readonly command: Readonly<{
    disposition: "runtime-bindable" | "catalog-only";
    key: PdfCanonicalId | null;
    handlerRef: "host-binding-required" | null;
    undoPolicy: "existing-handler" | "view-only" | "unknown";
  }>;
  readonly nativeObjectKinds: readonly string[];
  readonly import: Readonly<{
    disposition: "unknown" | "not-applicable";
    projectionRefs: readonly string[];
    preserveUnknownObjects: "unknown" | true;
    lossCodes: readonly string[];
  }>;
  readonly nativeSave: Readonly<{
    disposition: "unknown" | "resolver-required" | "not-applicable";
    serializerRefs: readonly string[];
    preservePolicy: string | null;
    allowedDispositions: readonly PdfOutputDisposition[];
    lossCodes: readonly string[];
  }>;
  readonly flatten: Readonly<{
    disposition: "unknown" | "consent-required" | "blocked" | "not-applicable";
    lossCodes: readonly string[];
  }>;
  readonly featureFlags: readonly PdfFeatureFlagId[];
  readonly evidenceTests: readonly string[];
  readonly acceptanceContractIds: readonly string[];
}

export interface PdfFeatureCapabilityCatalog {
  readonly schemaVersion: typeof PDF_CAPABILITY_CATALOG_SCHEMA_VERSION;
  readonly features: readonly PdfFeatureCapability[];
  readonly byId: Readonly<Record<PdfCanonicalId, PdfFeatureCapability>>;
}

type CapabilityOverride = Partial<
  Pick<
    PdfFeatureCapability,
    | "entryPoints"
    | "contexts"
    | "availabilityPredicate"
    | "command"
    | "nativeObjectKinds"
    | "import"
    | "nativeSave"
    | "flatten"
    | "featureFlags"
    | "evidenceTests"
    | "acceptanceContractIds"
  >
>;

const runtimeCommand = (
  id: PdfCanonicalId,
  allowedDispositions: readonly PdfOutputDisposition[],
  undoPolicy: "existing-handler" | "view-only",
): CapabilityOverride => ({
  availabilityPredicate:
    id === "PDF-FIL-DOC-02" || id === "PDF-FIL-DOC-10"
      ? "document-writable"
      : "document-ready",
  command: {
    disposition: "runtime-bindable",
    key: id,
    handlerRef: "host-binding-required",
    undoPolicy,
  },
  nativeObjectKinds: ["document"],
  nativeSave: {
    disposition: "resolver-required",
    serializerRefs: [],
    preservePolicy: null,
    allowedDispositions,
    lossCodes: [],
  },
  flatten: allowedDispositions.includes("flatten-with-consent")
    ? {
        disposition: "consent-required",
        lossCodes: ["PDF_LOSS_FLATTEN_REWRITE"],
      }
    : { disposition: "blocked", lossCodes: [] },
  evidenceTests: [
    "PDFCAPTEST.COMMAND.FAIL_CLOSED",
    "PDFCAPTEST.DISPOSITION.SAME_REASON",
  ],
  acceptanceContractIds: ["PDF.WP.P0.01"],
});

const nativeCrossTabCommand = (
  id: PdfCanonicalId,
  nativeObjectKinds: readonly string[],
): CapabilityOverride => ({
  ...runtimeCommand(id, ["native", "blocked"], "existing-handler"),
  availabilityPredicate: "document-writable",
  nativeObjectKinds,
  evidenceTests: [
    "PDFCAPTEST.CROSS_TAB.EXACT_ACTION_IDENTITY",
    "PDFCAPTEST.CROSS_TAB.RECEIPT_BOUND_MUTATION",
    "PDFCAPTEST.CROSS_TAB.STALE_GENERATION_BLOCKED",
  ],
  acceptanceContractIds: [
    "PDF.WP.P0.01",
    "PDF-ARC-UI-02",
    "PDF-ARC-UI-04",
    "PDF-ARC-SAVE-01",
    "PDF-ARC-SAVE-03",
    "PDF-ARC-SAVE-04",
  ],
});

function entry(
  controlId: string,
  surface: PdfCapabilityEntryPoint["surface"],
  evidence: PdfCapabilityEntryPoint["evidence"] = "OBSERVED",
): PdfCapabilityEntryPoint {
  const unknown = evidence === "UNKNOWN";
  return {
    controlId,
    surface,
    evidence,
    availability: unknown ? "inert-unknown" : "not-implemented",
    reasonKey: unknown
      ? "pdf.capability.unknownEvidence"
      : "pdf.capability.notImplemented",
  };
}

const ALL_PDF_FLAGS = PDF_FEATURE_FLAG_REGISTRY.map((flag) => flag.id);

const CAPABILITY_OVERRIDES: Partial<
  Record<PdfCanonicalId, CapabilityOverride>
> = {
  "PDF-FIL-DOC-07": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "print-queue",
      "source-bound-pdf",
      "page-range",
      "temporary-print-document",
      "system-print-window",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfFileRibbonTab", "PdfBatchPrintDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["buildPdfBatchPrintDocument"],
      preservePolicy: "view-only-temporary-print-derivative",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.BATCH_PRINT.EXACT_SOURCE_IDENTITY",
      "PDFCAPTEST.BATCH_PRINT.STRICT_PAGE_RANGES",
      "PDFCAPTEST.BATCH_PRINT.QUEUE_ORDER",
      "PDFCAPTEST.BATCH_PRINT.REOPEN_GEOMETRY",
      "PDFCAPTEST.BATCH_PRINT.ONE_SYSTEM_PRINT",
      "PDFCAPTEST.BATCH_PRINT.FAILURE_RETAINS_QUEUE",
      "PDFCAPTEST.BATCH_PRINT.LOCAL_FILES_ZERO_UPLOAD",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-FIL-DOC-07"],
  },
  "PDF-TOO-CONV-04": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "tagged-reading-order",
      "text-layer",
      "accepted-ocr-text",
      "page-range",
      "page-separator",
      "page-label",
      "utf8-txt-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["collectCurrentDocumentText", "PdfExtractTextDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["formatPdfExtractedText", "downloadTextFile"],
      preservePolicy: "source-read-only-txt-export",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.TXT.TAGGED_READING_ORDER",
      "PDFCAPTEST.TXT.PAGE_RANGE_AND_SEPARATOR",
      "PDFCAPTEST.TXT.PAGE_LABEL_AND_UTF8_BOM",
      "PDFCAPTEST.TXT.STALE_SOURCE_BLOCKED",
      "PDFCAPTEST.TXT.MASK_REDACTION_GATE",
    ],
  },
  "PDF-TOO-AI-04": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "browser-display-capture",
      "media-recorder-session",
      "local-video-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfScreenRecordingDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["startPdfScreenRecording"],
      preservePolicy: "original-pdf-unchanged-local-screen-recording",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.SCREEN_RECORDING.PREFLIGHT_ZERO_PERMISSION",
      "PDFCAPTEST.SCREEN_RECORDING.SYSTEM_PICKER",
      "PDFCAPTEST.SCREEN_RECORDING.STOP_SINGLE_OUTPUT",
      "PDFCAPTEST.SCREEN_RECORDING.SYSTEM_END",
      "PDFCAPTEST.SCREEN_RECORDING.DISCARD_NO_OUTPUT",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-AI-04"],
  },
  "PDF-TOO-CONV-01": {
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-writable",
    nativeObjectKinds: [
      "source-bound-pdf-projection",
      "editable-document-page",
      "docx-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfOfficeProjectionDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["pageExportFormatForPdfProjection"],
      preservePolicy: "original-pdf-unchanged-office-projection",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.OFFICE_PROJECTION.PREFLIGHT_ZERO_CREATE",
      "PDFCAPTEST.OFFICE_PROJECTION.EXACT_SOURCE_AND_SAVED_STATE",
      "PDFCAPTEST.OFFICE_PROJECTION.DOCX_DOWNLOAD",
      "PDFCAPTEST.OFFICE_PROJECTION.DUPLICATE_DISPATCH_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-CONV-01"],
  },
  "PDF-TOO-CONV-02": {
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-writable",
    nativeObjectKinds: [
      "source-bound-pdf-projection",
      "editable-spreadsheet-page",
      "xlsx-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfOfficeProjectionDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["pageExportFormatForPdfProjection"],
      preservePolicy: "original-pdf-unchanged-office-projection",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.OFFICE_PROJECTION.PREFLIGHT_ZERO_CREATE",
      "PDFCAPTEST.OFFICE_PROJECTION.EXACT_SOURCE_AND_SAVED_STATE",
      "PDFCAPTEST.OFFICE_PROJECTION.XLSX_TEXT_SHEETS",
      "PDFCAPTEST.OFFICE_PROJECTION.DUPLICATE_DISPATCH_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-CONV-02"],
  },
  "PDF-TOO-CONV-03": {
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-writable",
    nativeObjectKinds: [
      "source-bound-pdf-projection",
      "editable-presentation-page",
      "pptx-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfOfficeProjectionDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["pageExportFormatForPdfProjection"],
      preservePolicy: "original-pdf-unchanged-office-projection",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.OFFICE_PROJECTION.PREFLIGHT_ZERO_CREATE",
      "PDFCAPTEST.OFFICE_PROJECTION.EXACT_SOURCE_AND_SAVED_STATE",
      "PDFCAPTEST.OFFICE_PROJECTION.PPTX_TEXT_GEOMETRY",
      "PDFCAPTEST.OFFICE_PROJECTION.DUPLICATE_DISPATCH_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-CONV-03"],
  },
  "PDF-TOO-CONV-05": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "source-bound-vector-geometry",
      "dxf-line",
      "dxf-lightweight-polyline",
      "page-layer",
      "metric-cad-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfCadExportDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["verifyPdfCadDxfOutput"],
      preservePolicy: "source-read-only-vector-dxf-extraction",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.CAD.REVIEW_ZERO_EGRESS",
      "PDFCAPTEST.CAD.EXACT_SOURCE_RECEIPT",
      "PDFCAPTEST.CAD.METRIC_PAGE_LAYERS",
      "PDFCAPTEST.CAD.VECTOR_ONLY_WITNESS",
      "PDFCAPTEST.CAD.STALE_SOURCE_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-CONV-05"],
  },
  "PDF-TOO-CONV-06": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "source-bound-page-raster",
      "image-xobject",
      "image-only-pdf",
      "output-loss-report",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfImageOnlyExportDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [
        "createPdfImageOnlyPageInventory",
        "verifyPdfImageOnlyOutput",
      ],
      preservePolicy: "source-read-only-image-only-derivative",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [
        "PDF_LOSS_IMAGE_ONLY_SEARCHABLE_TEXT",
        "PDF_LOSS_IMAGE_ONLY_INTERACTIVITY",
        "PDF_LOSS_IMAGE_ONLY_ACCESSIBILITY",
        "PDF_LOSS_IMAGE_ONLY_SIGNATURES",
      ],
    },
    flatten: {
      disposition: "consent-required",
      lossCodes: [
        "PDF_LOSS_IMAGE_ONLY_SEARCHABLE_TEXT",
        "PDF_LOSS_IMAGE_ONLY_INTERACTIVITY",
        "PDF_LOSS_IMAGE_ONLY_ACCESSIBILITY",
        "PDF_LOSS_IMAGE_ONLY_SIGNATURES",
      ],
    },
    evidenceTests: [
      "PDFCAPTEST.IMAGE_ONLY.OPTIONS_ZERO_RENDER",
      "PDFCAPTEST.IMAGE_ONLY.EXACT_DPI_AND_ENCODING",
      "PDFCAPTEST.IMAGE_ONLY.LOSS_CONSENT",
      "PDFCAPTEST.IMAGE_ONLY.SEMANTIC_REOPEN",
      "PDFCAPTEST.IMAGE_ONLY.STALE_SOURCE_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-CONV-06"],
  },
  "PDF-ARC-UI-01": {
    entryPoints: [
      entry("PDFCTL.RIBBON.HOME.TAB", "ribbon-tab"),
      entry("PDFCTL.RIBBON.INSERT.TAB", "ribbon-tab"),
      entry("PDFCTL.RIBBON.EDIT.TAB", "ribbon-tab"),
      entry("PDFCTL.RIBBON.PAGE.TAB", "ribbon-tab"),
      entry("PDFCTL.RIBBON.COMMENT.TAB", "ribbon-tab"),
      entry("PDFCTL.RIBBON.FEATURED.TAB", "ribbon-tab"),
    ],
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "office-ribbon-chassis",
      "canonical-tab",
      "command-group",
      "overflow-control",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfOfficeRibbon", "OfficeRibbonChassis"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: "view-only-ribbon-chassis",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.RIBBON.SHARED_CHASSIS",
      "PDFCAPTEST.RIBBON.CANONICAL_TAB_IDS",
      "PDFCAPTEST.RIBBON.ROVING_KEYBOARD_AND_OVERFLOW",
    ],
  },
  "PDF-TOO-CONV-09": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "tagged-reading-order",
      "text-layer",
      "accepted-ocr-text",
      "page-range",
      "txt-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["collectCurrentDocumentText", "PdfExtractTextDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["downloadTextFile"],
      preservePolicy: "source-read-only-text-export",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.EXTRACT_TEXT.LOCAL_REVIEW",
      "PDFCAPTEST.EXTRACT_TEXT.PAGE_RANGE",
      "PDFCAPTEST.EXTRACT_TEXT.STALE_SOURCE_BLOCKED",
      "PDFCAPTEST.EXTRACT_TEXT.MASK_REDACTION_GATE",
    ],
  },
  "PDF-TOO-CONV-10": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "image-xobject",
      "jpeg-encoded-stream",
      "jpeg2000-encoded-stream",
      "soft-mask-reference",
      "image-extraction-manifest",
      "zip-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["inspectPdfImageResources", "PdfExtractImagesDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["buildPdfImageExtractionZip"],
      preservePolicy: "source-read-only-image-extraction",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.EXTRACT_IMAGES.EXACT_ENCODED_STREAM",
      "PDFCAPTEST.EXTRACT_IMAGES.SHARED_RESOURCE_DEDUPE",
      "PDFCAPTEST.EXTRACT_IMAGES.UNSUPPORTED_MANIFEST",
      "PDFCAPTEST.EXTRACT_IMAGES.STALE_SOURCE_BLOCKED",
    ],
  },
  "PDF-TOO-CONV-11": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "source-bound-page-raster",
      "selected-region",
      "ocr-region-egress-policy",
      "read-only-text-result",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfSnapshotTextPopover"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: "source-read-only-region-ocr",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.SCREENSHOT_OCR.EXPLICIT_EGRESS_START",
      "PDFCAPTEST.SCREENSHOT_OCR.EXACT_SOURCE_PAGE_REGION",
      "PDFCAPTEST.SCREENSHOT_OCR.NO_ANNOTATION_PERSISTENCE",
      "PDFCAPTEST.SCREENSHOT_OCR.READ_ONLY_RESULT",
      "PDFCAPTEST.SCREENSHOT_OCR.STALE_SOURCE_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-CONV-11"],
  },
  "PDF-TOO-AI-01": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "tagged-reading-order",
      "text-layer",
      "accepted-ocr-text",
      "source-anchor",
      "translated-text",
      "txt-output",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["collectCurrentDocumentText", "PdfTranslationPreflight"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: ["downloadTextFile"],
      preservePolicy: "source-read-only-translated-text",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.TRANSLATION.PREFLIGHT_ZERO_EGRESS",
      "PDFCAPTEST.TRANSLATION.EXACT_SOURCE_AND_OCR",
      "PDFCAPTEST.TRANSLATION.STREAM_ABORT_ON_DRIFT",
      "PDFCAPTEST.TRANSLATION.READ_ONLY_RESULT",
      "PDFCAPTEST.TRANSLATION.COPY_AND_TXT_OUTPUT",
    ],
  },
  "PDF-HOM-NAV-05": {
    ...nativeCrossTabCommand("PDF-HOM-NAV-05", [
      "catalog",
      "outline-root",
      "outline-item",
      "destination",
    ]),
    entryPoints: [entry("PDFCTL.OUTLINE.MANAGED_LIFECYCLE", "ribbon-control")],
  },
  "PDF-INS-CONTENT-01": {
    ...nativeCrossTabCommand("PDF-INS-CONTENT-01", [
      "page",
      "annotation",
      "free-text",
      "appearance-stream",
    ]),
    entryPoints: [entry("PDFCTL.PAGE.CONTENT.APPEND_TEXT", "ribbon-control")],
  },
  "PDF-INS-CONTENT-02": {
    ...nativeCrossTabCommand("PDF-INS-CONTENT-02", [
      "page",
      "annotation",
      "stamp",
      "stamp-image-appearance",
      "appearance-stream",
    ]),
    entryPoints: [entry("PDFCTL.PAGE.CONTENT.APPEND_IMAGE", "ribbon-control")],
  },
  "PDF-INS-CONTENT-03": nativeCrossTabCommand("PDF-INS-CONTENT-03", [
    "page",
    "annotation",
    "link",
    "action",
    "destination",
  ]),
  "PDF-INS-CONTENT-04": nativeCrossTabCommand("PDF-INS-CONTENT-04", [
    "page",
    "annotation",
    "stamp",
    "visual-signature-appearance",
    "appearance-stream",
  ]),
  "PDF-FIL-DOC-02": runtimeCommand(
    "PDF-FIL-DOC-02",
    ["native", "blocked"],
    "existing-handler",
  ),
  "PDF-FIL-DOC-03": runtimeCommand(
    "PDF-FIL-DOC-03",
    ["native", "flatten-with-consent", "read-only", "blocked"],
    "view-only",
  ),
  "PDF-FIL-DOC-06": runtimeCommand(
    "PDF-FIL-DOC-06",
    ["flatten-with-consent", "read-only", "blocked"],
    "view-only",
  ),
  "PDF-FIL-DOC-10": runtimeCommand(
    "PDF-FIL-DOC-10",
    ["native", "blocked"],
    "existing-handler",
  ),
  "PDF-INS-PAGE-02": {
    ...runtimeCommand(
      "PDF-INS-PAGE-02",
      ["native", "flatten-with-consent", "blocked"],
      "existing-handler",
    ),
    availabilityPredicate: "document-writable",
    nativeObjectKinds: ["document", "page-tree", "page", "resource"],
    evidenceTests: [
      "PDFCAPTEST.CROSS_TAB.EXACT_ACTION_IDENTITY",
      "PDFCAPTEST.PAGE_IMPORT.RECEIPT_BOUND_MUTATION",
      "PDFCAPTEST.PAGE_IMPORT.REOPEN_VERIFIED",
    ],
    acceptanceContractIds: [
      "PDF.WP.P0.01",
      "PDF-ARC-UI-02",
      "PDF-ARC-SAVE-01",
      "PDF-ARC-SAVE-04",
    ],
  },
  "PDF-INS-DECOR-01": nativeCrossTabCommand("PDF-INS-DECOR-01", [
    "page",
    "annotation",
    "stamp",
    "stamp-image-appearance",
    "appearance-stream",
  ]),
  "PDF-INS-DECOR-02": nativeCrossTabCommand("PDF-INS-DECOR-02", [
    "page",
    "annotation",
    "stamp",
    "stamp-image-appearance",
    "appearance-stream",
  ]),
  "PDF-INS-DECOR-03": nativeCrossTabCommand("PDF-INS-DECOR-03", [
    "page",
    "annotation",
    "stamp",
    "stamp-image-appearance",
    "appearance-stream",
  ]),
  // WPS 背景: written into the PDF as its own native revision
  // (planMraiNativePageBackground), undone by the restore it records.
  "PDF-INS-DECOR-04": {
    ...runtimeCommand(
      "PDF-INS-DECOR-04",
      ["native", "flatten-with-consent", "blocked"],
      "existing-handler",
    ),
    entryPoints: [entry("PDFCTL.SURFACE.BACKGROUND", "ribbon-control")],
    nativeObjectKinds: ["page", "content-stream", "resource-dictionary"],
    flatten: {
      disposition: "consent-required",
      lossCodes: [
        "PDF_LOSS_BACKGROUND_MANAGED_CONTENT",
        "PDF_LOSS_BACKGROUND_NATIVE_PORTABILITY",
        "PDF_LOSS_STRUCTURAL_REWRITE",
        "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
        "PDF_LOSS_INCREMENTAL_HISTORY",
      ],
    },
    evidenceTests: [
      "PDFCAPTEST.BACKGROUND.EXACT_AUTHORITY",
      "PDFCAPTEST.BACKGROUND.ALL_OR_ZERO",
      "PDFCAPTEST.BACKGROUND.REOPEN_VERIFIED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.03.BACKGROUND_OUTPUT_POLICY_01"],
  },
  "PDF-EDT-IMAGE-02": {
    ...runtimeCommand(
      "PDF-EDT-IMAGE-02",
      ["native", "flatten-with-consent", "blocked"],
      "existing-handler",
    ),
    entryPoints: [
      entry("PDFCTL.CONTEXT.PICTURE.DELETE_SOURCE", "ribbon-control"),
      entry("PDFCTL.CONTEXT.PICTURE.REPLACE_SOURCE", "ribbon-control"),
      entry("PDFCTL.CONTEXT.PICTURE.TRANSFORM_SOURCE", "ribbon-control"),
      entry("PDFCTL.CONTEXT.PICTURE.ARRANGE_SOURCE", "ribbon-control"),
    ],
    nativeObjectKinds: ["page", "content-stream", "image", "shape"],
    flatten: {
      disposition: "consent-required",
      lossCodes: [
        "PDF_LOSS_EMBEDDED_IMAGE_OVERLAY_SEMANTICS",
        "PDF_LOSS_STRUCTURAL_REWRITE",
        "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
        "PDF_LOSS_INCREMENTAL_HISTORY",
      ],
    },
    evidenceTests: [
      "PDFCAPTEST.EMBEDDED_IMAGE.EXACT_AUTHORITY",
      "PDFCAPTEST.EMBEDDED_IMAGE.ALL_OR_ZERO",
      "PDFCAPTEST.EMBEDDED_IMAGE.REOPEN_VERIFIED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.03.EMBEDDED_IMAGE_OVERLAY_POLICY_01"],
  },
  "PDF-EDT-IMAGE-01": nativeCrossTabCommand("PDF-EDT-IMAGE-01", [
    "page",
    "annotation",
    "stamp",
    "stamp-image-appearance",
    "appearance-stream",
  ]),
  "PDF-EDT-OBJECT-03": nativeCrossTabCommand("PDF-EDT-OBJECT-03", [
    "edit-transaction",
    "document",
  ]),
  "PDF-CMT-MARK-01": nativeCrossTabCommand("PDF-CMT-MARK-01", [
    "annotation",
    "text-markup",
  ]),
  "PDF-CMT-MARK-02": nativeCrossTabCommand("PDF-CMT-MARK-02", [
    "annotation",
    "shape",
    "appearance-stream",
  ]),
  "PDF-CMT-MARK-03": nativeCrossTabCommand("PDF-CMT-MARK-03", [
    "annotation",
    "text-markup",
  ]),
  "PDF-CMT-MARK-04": nativeCrossTabCommand("PDF-CMT-MARK-04", [
    "annotation",
    "text-markup",
  ]),
  "PDF-CMT-MARK-05": nativeCrossTabCommand("PDF-CMT-MARK-05", [
    "annotation",
    "text-markup",
  ]),
  "PDF-CMT-MARK-06": nativeCrossTabCommand("PDF-CMT-MARK-06", [
    "annotation",
    "caret",
    "text-markup",
    "appearance-stream",
  ]),
  "PDF-CMT-NOTE-01": nativeCrossTabCommand("PDF-CMT-NOTE-01", [
    "annotation",
    "note",
  ]),
  "PDF-CMT-NOTE-03": nativeCrossTabCommand("PDF-CMT-NOTE-03", [
    "annotation",
    "note",
    "text-markup",
    "ink",
    "shape",
  ]),
  "PDF-CMT-DRAW-01": nativeCrossTabCommand("PDF-CMT-DRAW-01", [
    "annotation",
    "ink",
    "appearance-stream",
  ]),
  "PDF-CMT-DRAW-02": {
    ...nativeCrossTabCommand("PDF-CMT-DRAW-02", [
      "annotation",
      "shape",
      "appearance-stream",
    ]),
    entryPoints: [entry("PDFCTL.PAGE.CONTENT.APPEND_PATH", "ribbon-control")],
  },
  "PDF-CMT-DRAW-04": nativeCrossTabCommand("PDF-CMT-DRAW-04", [
    "annotation",
    "stamp",
    "stamp-image-appearance",
    "appearance-stream",
  ]),
  "PDF-CMT-DRAW-05": nativeCrossTabCommand("PDF-CMT-DRAW-05", [
    "annotation",
    "file-specification",
    "embedded-file",
  ]),
  "PDF-CMT-DRAW-06": nativeCrossTabCommand("PDF-CMT-DRAW-06", [
    "annotation",
    "measurement",
    "appearance-stream",
  ]),
  "PDF-PAG-STRUCT-01": nativeCrossTabCommand("PDF-PAG-STRUCT-01", [
    "document",
    "page-tree",
    "page",
    "resource",
  ]),
  "PDF-EDT-OBJECT-02": nativeCrossTabCommand("PDF-EDT-OBJECT-02", [
    "page",
    "annotation",
    "link",
    "action",
    "destination",
  ]),
  "PDF-PAG-STRUCT-02": nativeCrossTabCommand("PDF-PAG-STRUCT-02", [
    "document",
    "page-tree",
    "page",
  ]),
  "PDF-PAG-STRUCT-03": runtimeCommand(
    "PDF-PAG-STRUCT-03",
    ["flatten-with-consent", "blocked"],
    "view-only",
  ),
  "PDF-PAG-STRUCT-04": {
    ...nativeCrossTabCommand("PDF-PAG-STRUCT-04", [
      "document",
      "page-tree",
      "page",
      "resource",
      "content-stream",
    ]),
    evidenceTests: [
      "PDFCAPTEST.PAGE_REPLACE.STABLE_PAGE_IDENTITY",
      "PDFCAPTEST.PAGE_REPLACE.RECEIPT_BOUND_MUTATION",
      "PDFCAPTEST.PAGE_REPLACE.UNDO_REOPEN_VERIFIED",
    ],
    acceptanceContractIds: [
      "PDF.WP.P0.01",
      "PDF-ARC-UI-02",
      "PDF-ARC-SAVE-01",
      "PDF-ARC-SAVE-04",
    ],
  },
  "PDF-PAG-STRUCT-05": nativeCrossTabCommand("PDF-PAG-STRUCT-05", [
    "document",
    "page-tree",
    "page",
  ]),
  "PDF-PAG-STRUCT-07": nativeCrossTabCommand("PDF-PAG-STRUCT-07", [
    "page",
    "page-dictionary",
    "annotation",
    "appearance-stream",
  ]),
  "PDF-PAG-STRUCT-08": nativeCrossTabCommand("PDF-PAG-STRUCT-08", [
    "page",
    "page-dictionary",
    "page-box",
  ]),
  // WPS 分割页面: the native page split writer (planMraiNativePageSplit).
  "PDF-PAG-STRUCT-09": {
    ...nativeCrossTabCommand("PDF-PAG-STRUCT-09", [
      "document",
      "page-tree",
      "page",
      "page-box",
    ]),
    entryPoints: [entry("PDFCTL.PAGE.STRUCTURE.SPLIT_PAGE", "ribbon-control")],
  },
  "PDF-PAG-STRUCT-11": {
    ...nativeCrossTabCommand("PDF-PAG-STRUCT-11", [
      "page",
      "page-dictionary",
      "media-box",
      "page-box",
    ]),
    evidenceTests: [
      "PDFCAPTEST.PAGE_SIZE.RECEIPT_BOUND_MUTATION",
      "PDFCAPTEST.PAGE_SIZE.UNDO_REOPEN_VERIFIED",
      "PDFCAPTEST.PAGE_SIZE.UNSAFE_BOUNDARIES_REFUSED",
    ],
  },
  "PDF-PAG-STRUCT-12": nativeCrossTabCommand("PDF-PAG-STRUCT-12", [
    "document",
    "page-labels",
    "number-tree",
  ]),
  "PDF-TOO-FILE-01": runtimeCommand(
    "PDF-TOO-FILE-01",
    ["flatten-with-consent", "blocked"],
    "view-only",
  ),
  "PDF-TOO-FILE-02": runtimeCommand(
    "PDF-TOO-FILE-02",
    ["flatten-with-consent", "blocked"],
    "view-only",
  ),
  "PDF-TOO-FILE-03": {
    ...runtimeCommand(
      "PDF-TOO-FILE-03",
      ["flatten-with-consent", "blocked"],
      "view-only",
    ),
    nativeObjectKinds: ["document", "image-xobject"],
    flatten: {
      disposition: "consent-required",
      lossCodes: [
        "PDF_LOSS_STRUCTURAL_REWRITE",
        "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
        "PDF_LOSS_INCREMENTAL_HISTORY",
        "PDF_LOSS_IMAGE_RESAMPLING",
        "PDF_LOSS_JPEG_RECOMPRESSION",
        "PDF_LOSS_IMAGE_COLOR_FIDELITY",
      ],
    },
    evidenceTests: [
      "PDFCAPTEST.COMPRESS.EXACT_ACTION_AUTHORITY",
      "PDFCAPTEST.COMPRESS.LOSSY_EVIDENCE_PLAN",
      "PDFCAPTEST.COMPRESS.LOSSY_RESULT_PROOF",
    ],
    acceptanceContractIds: [
      "PDF.WP.P0.03.LOSSLESS_COMPRESSION_POLICY_01",
      "PDF.WP.P0.03.LOSSY_COMPRESSION_PRESET_POLICY_01",
    ],
  },
  "PDF-TOO-FILE-04": {
    ...runtimeCommand(
      "PDF-TOO-FILE-04",
      ["flatten-with-consent", "blocked"],
      "view-only",
    ),
    nativeObjectKinds: ["page", "page-raster"],
    flatten: {
      disposition: "consent-required",
      lossCodes: [
        "PDF_LOSS_DESKEW_TARGET_PAGE_TEXT",
        "PDF_LOSS_DESKEW_TARGET_PAGE_STRUCTURE",
        "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
        "PDF_LOSS_INCREMENTAL_HISTORY",
      ],
    },
    evidenceTests: [
      "PDFCAPTEST.DESKEW.EXACT_AUTHORITY",
      "PDFCAPTEST.DESKEW.RASTER_CONSENT",
      "PDFCAPTEST.DESKEW.STALE_INPUT_BLOCKED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.03.DESKEW_RASTER_POLICY_01"],
  },
  "PDF-TOO-OCR-03": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "page",
      "page-raster",
      "ocr-egress-policy",
      "ocr-review-annotation",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: [],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: null,
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.OCR.EGRESS_PREFLIGHT_EXACT_POLICY",
      "PDFCAPTEST.OCR.EGRESS_CANCEL_ZERO_UPLOAD",
      "PDFCAPTEST.OCR.EGRESS_PER_PAGE_REVALIDATION",
      "PDFCAPTEST.OCR.RESULT_MANUAL_ACCEPTANCE",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-TOO-OCR-03"],
  },
  "PDF-FRM-AUTH-02": {
    ...runtimeCommand(
      "PDF-FRM-AUTH-02",
      ["native", "flatten-with-consent", "blocked"],
      "existing-handler",
    ),
    availabilityPredicate: "document-writable",
    nativeObjectKinds: [
      "acroform",
      "form-field",
      "widget",
      "appearance-stream",
    ],
    evidenceTests: [
      "PDFCAPTEST.FORM_AUTHORING.EXACT_ACTION_IDENTITY",
      "PDFCAPTEST.FORM_AUTHORING.RECEIPT_BOUND_MUTATION",
      "PDFCAPTEST.FORM_AUTHORING.REOPEN_VERIFIED",
    ],
    acceptanceContractIds: [
      "PDF.WP.P0.01",
      "PDF-ARC-SAVE-01",
      "PDF-ARC-SAVE-03",
      "PDF-ARC-SAVE-04",
    ],
  },
  "PDF-FRM-AUTH-03": {
    ...runtimeCommand(
      "PDF-FRM-AUTH-03",
      ["native", "blocked"],
      "existing-handler",
    ),
    availabilityPredicate: "document-writable",
    nativeObjectKinds: [
      "acroform",
      "form-field",
      "widget",
      "appearance-stream",
    ],
    nativeSave: {
      disposition: "resolver-required",
      serializerRefs: [
        "PDFViewer#flushFormValues",
        "PdfPageWindow#handleNativeFormFieldChange",
        "pdf_roundtrip.native_edit_primitives.update_form_field",
      ],
      preservePolicy: "native-acroform-field-value-v1",
      allowedDispositions: ["native", "blocked"],
      lossCodes: [],
    },
    evidenceTests: [
      "PDFCAPTEST.FORM_DATA.TYPED_PREFLIGHT",
      "PDFCAPTEST.FORM_DATA.NATIVE_GESTURE_HISTORY",
      "PDFCAPTEST.FORM_DATA.REOPEN_VERIFIED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF.WP.P1.03"],
  },
  "PDF-ARC-UI-02": {
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "command-registry",
      "capability-decision",
      "host-binding",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: [],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: null,
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.COMMAND.EXACT_ACTION_IDENTITY",
      "PDFCAPTEST.COMMAND.HOST_BINDING_REQUIRED",
      "PDFCAPTEST.COMMAND.STALE_GENERATION_BLOCKED",
      "PDFCAPTEST.DISPOSITION.SAME_REASON",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-ARC-UI-02"],
  },
  "PDF-ARC-UI-03": {
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: ["dialog-shell", "focus-contract", "modal-state"],
    import: {
      disposition: "not-applicable",
      projectionRefs: [],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: null,
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.DIALOG.SHARED_SHELL",
      "PDFCAPTEST.DIALOG.CANCEL_INITIAL_FOCUS",
      "PDFCAPTEST.DIALOG.ESCAPE_AND_RETURN_FOCUS",
      "PDFCAPTEST.DIALOG.BUSY_FAIL_CLOSED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-ARC-UI-03"],
  },
  "PDF-ARC-UI-04": {
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-ready",
    entryPoints: [
      entry("PDFCTL.MENU.APPLICATION.ROOT", "application-menu"),
      entry("PDFCTL.MENU.APPLICATION.ABOUT", "application-menu"),
      entry(
        "PDFCTL.MENU.APPLICATION.ADD_EXTENSION",
        "application-menu",
        "UNKNOWN",
      ),
      entry("PDFCTL.MENU.APPLICATION.SERVICES", "application-menu", "UNKNOWN"),
    ],
    nativeObjectKinds: ["none"],
    import: {
      disposition: "not-applicable",
      projectionRefs: [],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: null,
      allowedDispositions: ["read-only"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.CATALOG.BIDIRECTIONAL",
      "PDFCAPTEST.CATALOG.MALFORMED_ROWS_REJECTED",
      "PDFCAPTEST.CATALOG.UNKNOWN_CONTROLS_INERT",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-ARC-UI-04"],
  },
  // Feedback is application-owned (the audit's own rule): the notebook
  // sidebar's Report a problem, sent to the user's technical-team
  // conversation; the PDF ribbon has no help or feedback of its own.
  "PDF-ARC-UI-05": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "manual-feedback-payload",
      "focus-contract",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["NotebookFeedbackPanel"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: "view-only-manual-handoff",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.FEEDBACK.USER_TEXT_ONLY",
      "PDFCAPTEST.FEEDBACK.NON_CONTENT_SUPPORT_DETAILS",
      "PDFCAPTEST.FEEDBACK.RETURN_FOCUS",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-ARC-UI-05"],
  },
  "PDF-ARC-SAVE-03": {
    nativeObjectKinds: ["document"],
    nativeSave: {
      disposition: "resolver-required",
      serializerRefs: [],
      preservePolicy: null,
      allowedDispositions: [
        "native",
        "flatten-with-consent",
        "read-only",
        "blocked",
      ],
      lossCodes: [],
    },
    flatten: {
      disposition: "consent-required",
      lossCodes: ["PDF_LOSS_FLATTEN_REWRITE"],
    },
  },
  "PDF-ARC-SAVE-06": {
    entryPoints: [
      entry("PDFCTL.DIALOG.IMAGE_EXPORT.MEMBERSHIP", "dialog-membership"),
      entry("PDFCTL.DIALOG.OCR.MEMBERSHIP", "dialog-membership"),
    ],
    nativeObjectKinds: ["entitlement-state"],
    featureFlags: ALL_PDF_FLAGS,
  },
  "PDF-VIE-ZOOM-03": {
    featureFlags: ["NEXT_PUBLIC_PDF_RENDER_LIFECYCLE_V2"],
  },
  "PDF-VIE-MODE-05": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "tagged-reading-order",
      "pdf-text-layer",
      "accepted-ocr-text",
      "speech-controller",
      "system-voice",
      "reading-settings",
      "pdf-document-language",
      "device-read-aloud-preferences",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: [
        "PdfSpeechController",
        "PdfReadAloudSettingsDialog",
        "read-aloud-preferences",
      ],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "not-applicable",
      serializerRefs: [],
      preservePolicy: "local-view-state-only",
      allowedDispositions: ["read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.READ_ALOUD.TAGGED_OR_TEXT_OR_ACCEPTED_OCR",
      "PDFCAPTEST.READ_ALOUD.PAGE_CONTINUATION",
      "PDFCAPTEST.READ_ALOUD.LOCAL_SYSTEM_VOICE",
      "PDFCAPTEST.READ_ALOUD.LANGUAGE_VOICE_RATE",
      "PDFCAPTEST.READ_ALOUD.DEVICE_PREFERENCES",
      "PDFCAPTEST.READ_ALOUD.EXPLICIT_DOCUMENT_LANGUAGE",
      "PDFCAPTEST.READ_ALOUD.MALFORMED_PREFERENCE_FALLBACK",
      "PDFCAPTEST.READ_ALOUD.VOICE_LANGUAGE_MATCH",
      "PDFCAPTEST.READ_ALOUD.SETTINGS_RESTART_ONE_SESSION",
      "PDFCAPTEST.READ_ALOUD.PREVIOUS_NEXT_SEGMENT",
      "PDFCAPTEST.READ_ALOUD.EXACT_VISUAL_SEGMENT",
      "PDFCAPTEST.READ_ALOUD.WORD_BOUNDARY_HIGHLIGHT",
      "PDFCAPTEST.READ_ALOUD.LOCAL_KEYBOARD_CONTROLS",
      "PDFCAPTEST.READ_ALOUD.UNSUPPORTED_EXPLAINED",
      "PDFCAPTEST.READ_ALOUD.ZERO_UPLOAD",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-VIE-MODE-05"],
  },
  "PDF-EDT-OBJECT-01": {
    ...nativeCrossTabCommand("PDF-EDT-OBJECT-01", [
      "page",
      "content-stream",
      "vector-shape",
      "vector-subpath",
    ]),
    featureFlags: ["NEXT_PUBLIC_PDF_OVERLAY_COORDINATES_V2"],
    entryPoints: [
      entry("PDFCTL.PAGE.VECTOR.GEOMETRY", "ribbon-control"),
      entry("PDFCTL.PAGE.VECTOR.STYLE", "ribbon-control"),
    ],
    evidenceTests: [
      "PDFCAPTEST.VECTOR_GEOMETRY.SELECTION_IS_NOT_AUTHORITY",
      "PDFCAPTEST.VECTOR_GEOMETRY.FRESH_PROOF_AND_RECEIPT",
      "PDFCAPTEST.VECTOR_GEOMETRY.SIBLING_FIDELITY",
      "PDFCAPTEST.VECTOR_GEOMETRY.FAIL_CLOSED",
      "PDFCAPTEST.VECTOR_STYLE.OBJECT_LOCAL_RECEIPT",
      "PDFCAPTEST.VECTOR_STYLE.SIBLING_FIDELITY",
    ],
  },
  "PDF-SEC-PASS-01": {
    featureFlags: ["NEXT_PUBLIC_PDF_PAGE_WINDOW_MUTATION_GUARD_V2"],
  },
  "PDF-SEC-PASS-02": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "document",
      "encryption-dictionary",
      "standard-security-handler",
      "encrypted-copy",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfDocumentEncryptDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "resolver-required",
      serializerRefs: ["PdfDocumentEncryptDialog", "pdf_encrypt.py"],
      preservePolicy: "verified-encrypted-copy-original-unchanged",
      allowedDispositions: ["native", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.ENCRYPT.OPEN_PASSWORD_VALIDATED",
      "PDFCAPTEST.ENCRYPT.SOURCE_SHA_BOUND",
      "PDFCAPTEST.ENCRYPT.REOPEN_WITH_USER_PASSWORD",
      "PDFCAPTEST.ENCRYPT.ORIGINAL_UNCHANGED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-SEC-PASS-02"],
  },
  "PDF-SEC-PASS-03": {
    contexts: ["reader"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "document",
      "encryption-dictionary",
      "owner-password",
      "permission-bits",
      "encrypted-copy",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfDocumentEncryptDialog"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "resolver-required",
      serializerRefs: ["PdfDocumentEncryptDialog", "pdf_encrypt.py"],
      preservePolicy: "verified-permission-copy-original-unchanged",
      allowedDispositions: ["native", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.ENCRYPT.OWNER_PASSWORD_VALIDATED",
      "PDFCAPTEST.ENCRYPT.PERMISSION_MATRIX_EXACT",
      "PDFCAPTEST.ENCRYPT.PERMISSION_SELECT_ALL_TRI_STATE",
      "PDFCAPTEST.ENCRYPT.REOPEN_WITH_OWNER_PASSWORD",
      "PDFCAPTEST.ENCRYPT.SECRET_NEVER_PERSISTED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-SEC-PASS-03"],
  },
  "PDF-SEC-SAFE-01": {
    ...runtimeCommand("PDF-SEC-SAFE-01", ["read-only", "blocked"], "view-only"),
    entryPoints: [
      entry("PDFCTL.DIALOG.SECURITY.SAFE_COPY", "dialog-membership"),
    ],
    nativeObjectKinds: [
      "document",
      "catalog",
      "action",
      "annotation",
      "acroform",
      "form-field",
    ],
    evidenceTests: [
      "PDFCAPTEST.SECURITY.DANGEROUS_ACTION_LOCATIONS",
      "PDFCAPTEST.SECURITY.SAFE_COPY_ORIGINAL_UNCHANGED",
      "PDFCAPTEST.SECURITY.SAFE_COPY_REOPEN_VERIFIED",
    ],
    acceptanceContractIds: ["PDF.WP.P0.01", "PDF-SEC-SAFE-01"],
  },
  "PDF-SEC-REDACT-02": {
    ...runtimeCommand(
      "PDF-SEC-REDACT-02",
      ["flatten-with-consent", "blocked"],
      "view-only",
    ),
    nativeObjectKinds: [
      "document",
      "page",
      "content-stream",
      "resource-dictionary",
      "image-xobject",
      "form-field",
    ],
    flatten: {
      disposition: "consent-required",
      lossCodes: [
        "PDF_LOSS_REDACTION_TARGET_CONTENT",
        "PDF_LOSS_REDACTED_PAGE_STRUCTURE",
        "PDF_LOSS_SOURCE_OBJECT_IDENTITY",
        "PDF_LOSS_INCREMENTAL_HISTORY",
      ],
    },
    featureFlags: ["NEXT_PUBLIC_PDF_REDACTION_REACHABILITY_V2"],
    evidenceTests: [
      "PDFCAPTEST.REDACTION.EXACT_APPLY_AUTHORITY",
      "PDFCAPTEST.REDACTION.PENDING_OUTPUT_BLOCKED",
      "PDFCAPTEST.REDACTION.FAILURE_ZERO_OUTPUT",
    ],
    acceptanceContractIds: [
      "PDF.WP.P0.05.REDACTION_EXPLICIT_APPLY_FAIL_CLOSED_01",
    ],
  },
  "PDF-FRM-AUTH-04": {
    featureFlags: ["NEXT_PUBLIC_PDF_ACROFORM_MUTATION_GUARD_V2"],
  },
  "PDF-EDT-TEXT-01": {
    ...nativeCrossTabCommand("PDF-EDT-TEXT-01", [
      "page",
      "content-stream",
      "font-resource",
      "to-unicode-cmap",
    ]),
    entryPoints: [
      entry("PDFCTL.RIBBON.EDIT.EDIT_CONTENT", "ribbon-control"),
      entry("PDFCTL.RIBBON.HOME.EDIT_CONTENT", "ribbon-control"),
      entry("PDFCTL.CONTEXT.TEXT.DELETE_OBJECT", "ribbon-control"),
      entry("PDFCTL.CONTEXT.TEXT.DECORATION", "ribbon-control"),
    ],
    featureFlags: ["NEXT_PUBLIC_PDF_TEXT_EDIT_STABILITY_V2"],
  },
  "PDF-STD-INTEROP-03": {
    featureFlags: ["NEXT_PUBLIC_PDF_FREETEXT_RD_V2"],
  },
  "PDF-STD-A11Y-02": {
    ...runtimeCommand("PDF-STD-A11Y-02", ["native", "read-only", "blocked"], "existing-handler"),
    contexts: ["reader", "native-editor"],
    availabilityPredicate: "document-ready",
    nativeObjectKinds: [
      "tagged-structure",
      "reading-order",
      "alternative-text",
      "page-bound-finding",
    ],
    import: {
      disposition: "not-applicable",
      projectionRefs: ["PdfPropertiesDialog", "pdfTaggedReadingOrderText"],
      preserveUnknownObjects: true,
      lossCodes: [],
    },
    nativeSave: {
      disposition: "resolver-required",
      serializerRefs: ["createPdfTaggedTextOperation", "PdfTaggedTextPatch"],
      preservePolicy: "source-bound-semantic-properties-only",
      allowedDispositions: ["native", "read-only", "blocked"],
      lossCodes: [],
    },
    flatten: { disposition: "not-applicable", lossCodes: [] },
    evidenceTests: [
      "PDFCAPTEST.ACCESSIBILITY.PAGE_BOUND_FINDINGS",
      "PDFCAPTEST.ACCESSIBILITY.SELECT_TO_NAVIGATE",
      "PDFCAPTEST.ACCESSIBILITY.FAILURE_STAYS_UNKNOWN",
    ],
  },
};

function domainFor(id: PdfCanonicalId): PdfCapabilityDomain {
  if (id.startsWith("PDF-FIL-")) return "file";
  if (id.startsWith("PDF-HOM-")) return "home";
  if (id.startsWith("PDF-VIE-")) return "view";
  if (id.startsWith("PDF-INS-")) return "insert";
  if (id.startsWith("PDF-EDT-")) return "edit";
  if (id.startsWith("PDF-PAG-")) return "page";
  if (id.startsWith("PDF-CMT-")) return "comment";
  if (id.startsWith("PDF-FRM-")) return "forms";
  if (id.startsWith("PDF-SGN-")) return "signatures";
  if (id.startsWith("PDF-SEC-")) return "security";
  if (id.startsWith("PDF-TOO-")) return "tools";
  if (id.startsWith("PDF-STD-")) return "standards";
  return "architecture";
}

function freezeCapability(
  id: PdfCanonicalId,
  baselineStatus: PdfCanonicalStatus,
): PdfFeatureCapability {
  const override = CAPABILITY_OVERRIDES[id] ?? {};
  const entryPoints = Object.freeze(
    (override.entryPoints ?? []).map((item) => Object.freeze({ ...item })),
  );
  const contexts = Object.freeze([
    ...(override.contexts ?? ["reader", "native-editor"]),
  ]);
  const command = Object.freeze(
    override.command ?? {
      disposition: "catalog-only" as const,
      key: null,
      handlerRef: null,
      undoPolicy: "unknown" as const,
    },
  );
  const importContract = Object.freeze(
    override.import ?? {
      disposition: "unknown" as const,
      projectionRefs: Object.freeze([] as string[]),
      preserveUnknownObjects: "unknown" as const,
      lossCodes: Object.freeze([] as string[]),
    },
  );
  const nativeSave = Object.freeze(
    override.nativeSave ?? {
      disposition: "unknown" as const,
      serializerRefs: Object.freeze([] as string[]),
      preservePolicy: null,
      allowedDispositions: Object.freeze(["blocked"] as PdfOutputDisposition[]),
      lossCodes: Object.freeze([] as string[]),
    },
  );
  const flatten = Object.freeze(
    override.flatten ?? {
      disposition: "unknown" as const,
      lossCodes: Object.freeze([] as string[]),
    },
  );
  return Object.freeze({
    id,
    reasonKey:
      command.disposition === "runtime-bindable"
        ? "pdf.capability.runtimeDecisionRequired"
        : "pdf.capability.catalogOnly",
    baselineStatus,
    domain: domainFor(id),
    entryPoints,
    contexts,
    availabilityPredicate: override.availabilityPredicate ?? "catalog-only",
    command,
    nativeObjectKinds: Object.freeze([
      ...(override.nativeObjectKinds ?? ["unknown"]),
    ]),
    import: Object.freeze({
      ...importContract,
      projectionRefs: Object.freeze([...importContract.projectionRefs]),
      lossCodes: Object.freeze([...importContract.lossCodes]),
    }),
    nativeSave: Object.freeze({
      ...nativeSave,
      serializerRefs: Object.freeze([...nativeSave.serializerRefs]),
      allowedDispositions: Object.freeze([...nativeSave.allowedDispositions]),
      lossCodes: Object.freeze([...nativeSave.lossCodes]),
    }),
    flatten: Object.freeze({
      ...flatten,
      lossCodes: Object.freeze([...flatten.lossCodes]),
    }),
    featureFlags: Object.freeze([...(override.featureFlags ?? [])]),
    evidenceTests: Object.freeze([
      ...(override.evidenceTests ?? ["PDFCAPTEST.CATALOG.BIDIRECTIONAL"]),
    ]),
    acceptanceContractIds: Object.freeze([
      "PDF.WP.P0.01",
      ...(override.acceptanceContractIds ?? []).filter(
        (contractId) => contractId !== "PDF.WP.P0.01",
      ),
    ]),
  });
}

const PDF_CANONICAL_ID_SET = new Set<string>(
  PDF_CANONICAL_BASELINE.map(([id]) => id),
);

export function readPdfFeatureCapabilityCatalog(source: {
  version: string;
  features: readonly PdfFeatureCapability[];
}): PdfFeatureCapabilityCatalog {
  if (source.version !== PDF_CAPABILITY_CATALOG_SCHEMA_VERSION) {
    throw new Error("Unsupported PDF capability catalog version");
  }
  if (source.features.length !== PDF_CANONICAL_BASELINE.length) {
    throw new Error(
      `PDF capability catalog must contain exactly ${PDF_CANONICAL_BASELINE.length} canonical rows`,
    );
  }

  const byId = {} as Record<PdfCanonicalId, PdfFeatureCapability>;
  const controls = new Set<string>();
  for (const feature of source.features) {
    if (!PDF_CANONICAL_ID_SET.has(feature.id)) {
      throw new Error(`Unknown PDF canonical ID: ${feature.id}`);
    }
    if (byId[feature.id]) {
      throw new Error(`Duplicate PDF canonical ID: ${feature.id}`);
    }
    if (feature.domain !== domainFor(feature.id)) {
      throw new Error(`PDF capability domain drift: ${feature.id}`);
    }
    if (
      feature.command.disposition === "runtime-bindable" &&
      (feature.command.key !== feature.id ||
        feature.command.handlerRef !== "host-binding-required" ||
        feature.command.undoPolicy === "unknown")
    ) {
      throw new Error(`Incomplete PDF command contract: ${feature.id}`);
    }
    if (
      feature.command.disposition === "catalog-only" &&
      (feature.command.key !== null || feature.command.handlerRef !== null)
    ) {
      throw new Error(`Inert PDF capability has a command: ${feature.id}`);
    }
    if (
      feature.flatten.disposition === "consent-required" &&
      feature.flatten.lossCodes.length === 0
    ) {
      throw new Error(`PDF flatten contract lacks a loss code: ${feature.id}`);
    }
    if (
      feature.nativeSave.disposition === "resolver-required" &&
      feature.nativeSave.allowedDispositions.length === 0
    ) {
      throw new Error(`PDF save contract lacks dispositions: ${feature.id}`);
    }
    if (
      feature.evidenceTests.length === 0 ||
      feature.acceptanceContractIds.length === 0
    ) {
      throw new Error(
        `PDF capability lacks acceptance evidence: ${feature.id}`,
      );
    }
    for (const control of feature.entryPoints) {
      if (controls.has(control.controlId)) {
        throw new Error(`Duplicate PDF control ID: ${control.controlId}`);
      }
      if (
        control.evidence === "UNKNOWN" &&
        control.availability !== "inert-unknown"
      ) {
        throw new Error(
          `Unknown PDF control is not inert: ${control.controlId}`,
        );
      }
      controls.add(control.controlId);
    }
    byId[feature.id] = feature;
  }

  for (const [id] of PDF_CANONICAL_BASELINE) {
    if (!byId[id]) throw new Error(`Missing PDF canonical ID: ${id}`);
  }
  const expectedControls = new Set<string>(PDF_P0_01_CONTROL_IDS);
  if (
    controls.size !== expectedControls.size ||
    [...controls].some((controlId) => !expectedControls.has(controlId))
  ) {
    throw new Error("PDF.WP.P0.01 control coverage drift");
  }

  return Object.freeze({
    schemaVersion: PDF_CAPABILITY_CATALOG_SCHEMA_VERSION,
    features: Object.freeze([...source.features]),
    byId: Object.freeze({ ...byId }),
  });
}

export const PDF_FEATURE_CAPABILITY_CATALOG = readPdfFeatureCapabilityCatalog({
  version: PDF_CAPABILITY_CATALOG_SCHEMA_VERSION,
  features: PDF_CANONICAL_BASELINE.map(([id, status]) =>
    freezeCapability(id, status),
  ),
});

export const PDF_FEATURE_CAPABILITY_REGISTRY =
  PDF_FEATURE_CAPABILITY_CATALOG.features;

export function findPdfFeatureCapability(
  id: string,
): PdfFeatureCapability | null {
  return PDF_FEATURE_CAPABILITY_CATALOG.byId[id as PdfCanonicalId] ?? null;
}
