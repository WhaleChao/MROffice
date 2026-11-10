/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  documentSignatureStatus,
  judgeSignature,
  type SignatureStatus,
} from "./signature-validation";
import type { PdfSourceRiskSnapshot } from "./save-policy";
import { PDF_EDITABLE_MAX_BYTES } from "@mrnote/core/web/lib/document-size-limits.ts";

const SHA256_RE = /^[0-9a-f]{64}$/;
const UNIT_ID_RE = /^sig-[0-9a-f]{64}$/;
const VALIDATION_TIME_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{6}Z$/;
const REASON_RE = /^[A-Z][A-Z0-9_]{0,63}$/;
const MAX_SOURCE_BYTES = PDF_EDITABLE_MAX_BYTES;
const MAX_SIGNATURES = 32;

const REASONS = new Set([
  "CRYPTO_DISAGREEMENT",
  "CRYPTO_SET_DISAGREEMENT",
  "DUAL_CRYPTO_INVALID",
  "EMPTY_TRUST_ROOTS",
  "IMPLEMENTATION_DUPLICATE",
  "INVENTORY_DISAGREEMENT",
  "NO_SIGNATURES",
  "UNSUPPORTED_PDF_STRUCTURE",
  "WITNESS_SET_INVALID",
  "WITNESS_UNAVAILABLE",
]);

export type PdfSignatureIntegrity = "VALID" | "INVALID" | "UNKNOWN";

export interface PdfSignatureProof {
  readonly schemaVersion: "mrai.pdf.signature-proof.product.v1";
  readonly dataItemId: string;
  readonly source: Readonly<{ sha256: string; size: number }>;
  readonly validationTime: string;
  readonly signatures: readonly Readonly<{
    unitId: string;
    integrity: PdfSignatureIntegrity;
    coversCurrentFile: boolean;
  }>[];
  readonly aggregate: "INDETERMINATE" | "TOTAL_FAILED";
  readonly reason: string;
  readonly certificateTrust: "UNKNOWN";
  readonly timestamp: "UNKNOWN";
  readonly revocation: "UNKNOWN";
  readonly fieldMdp: "UNKNOWN";
  readonly dss: "UNKNOWN";
}

export interface PdfSignatureProofPresentation {
  readonly integrity: "intact" | "mismatch" | "unknown";
  readonly judgement: SignatureStatus;
  readonly headline: string;
  readonly detail: string;
  readonly unknownAxes: string;
  readonly loadingHeadline: string;
  readonly loadingDetail: string;
  /** Short reader-chrome label (Office's status-bar Signatures button). */
  readonly indicatorLabel: string;
  readonly loadingIndicatorLabel: string;
  /** Title of the panel the indicator opens (Office's Signatures pane). */
  readonly panelTitle: string;
  readonly detailsLabel: string;
  readonly verificationTimeLabel: string;
  readonly verificationTime: string | null;
  readonly reasonLabel: string;
  readonly downloadOriginalLabel: string;
  readonly downloadOriginalUnavailableLabel: string;
  readonly signatures: readonly PdfSignatureProofSignaturePresentation[];
  readonly axes: readonly PdfSignatureProofAxisPresentation[];
}

export interface PdfSignatureProofSignaturePresentation {
  readonly label: string;
  readonly integrity: "intact" | "mismatch" | "unknown";
  readonly integrityLabel: string;
  readonly coverage: "current" | "earlier";
  readonly coverageLabel: string;
}

export interface PdfSignatureProofAxisPresentation {
  readonly key:
    | "certificateTrust"
    | "timestamp"
    | "revocation"
    | "fieldMdp"
    | "dss";
  readonly label: string;
  readonly value: "unknown";
  readonly valueLabel: string;
}

export class PdfSignatureProofContractError extends Error {}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasExactKeys(
  value: Record<string, unknown>,
  expected: readonly string[],
): boolean {
  const keys = Object.keys(value).sort();
  const wanted = [...expected].sort();
  return (
    keys.length === wanted.length &&
    keys.every((key, index) => key === wanted[index])
  );
}

function rejectContract(): never {
  throw new PdfSignatureProofContractError(
    "Invalid PDF signature-proof response",
  );
}

function parseSignature(
  value: unknown,
): PdfSignatureProof["signatures"][number] {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ["unitId", "integrity", "coversCurrentFile"]) ||
    typeof value.unitId !== "string" ||
    !UNIT_ID_RE.test(value.unitId) ||
    (value.integrity !== "VALID" &&
      value.integrity !== "INVALID" &&
      value.integrity !== "UNKNOWN") ||
    typeof value.coversCurrentFile !== "boolean"
  ) {
    return rejectContract();
  }
  return Object.freeze({
    unitId: value.unitId,
    integrity: value.integrity,
    coversCurrentFile: value.coversCurrentFile,
  });
}

export function parsePdfSignatureProof(value: unknown): PdfSignatureProof {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, [
      "schemaVersion",
      "dataItemId",
      "source",
      "validationTime",
      "signatures",
      "aggregate",
      "reason",
      "certificateTrust",
      "timestamp",
      "revocation",
      "fieldMdp",
      "dss",
    ]) ||
    value.schemaVersion !== "mrai.pdf.signature-proof.product.v1" ||
    typeof value.dataItemId !== "string" ||
    value.dataItemId.length < 1 ||
    value.dataItemId.length > 160 ||
    value.dataItemId !== value.dataItemId.trim() ||
    !isRecord(value.source) ||
    !hasExactKeys(value.source, ["sha256", "size"]) ||
    typeof value.source.sha256 !== "string" ||
    !SHA256_RE.test(value.source.sha256) ||
    !Number.isSafeInteger(value.source.size) ||
    Number(value.source.size) < 1 ||
    Number(value.source.size) >= MAX_SOURCE_BYTES ||
    typeof value.validationTime !== "string" ||
    !VALIDATION_TIME_RE.test(value.validationTime) ||
    Number.isNaN(Date.parse(value.validationTime)) ||
    !Array.isArray(value.signatures) ||
    value.signatures.length > MAX_SIGNATURES ||
    (value.aggregate !== "INDETERMINATE" &&
      value.aggregate !== "TOTAL_FAILED") ||
    typeof value.reason !== "string" ||
    !REASON_RE.test(value.reason) ||
    !REASONS.has(value.reason) ||
    value.certificateTrust !== "UNKNOWN" ||
    value.timestamp !== "UNKNOWN" ||
    value.revocation !== "UNKNOWN" ||
    value.fieldMdp !== "UNKNOWN" ||
    value.dss !== "UNKNOWN"
  ) {
    return rejectContract();
  }

  const signatures = Object.freeze(value.signatures.map(parseSignature));
  if (
    new Set(signatures.map((signature) => signature.unitId)).size !==
    signatures.length
  ) {
    return rejectContract();
  }
  const allInvalid =
    signatures.length > 0 &&
    signatures.every((signature) => signature.integrity === "INVALID");
  if ((value.aggregate === "TOTAL_FAILED") !== allInvalid) {
    return rejectContract();
  }
  return Object.freeze({
    schemaVersion: value.schemaVersion,
    dataItemId: value.dataItemId,
    source: Object.freeze({
      sha256: value.source.sha256,
      size: Number(value.source.size),
    }),
    validationTime: value.validationTime,
    signatures,
    aggregate: value.aggregate,
    reason: value.reason,
    certificateTrust: value.certificateTrust,
    timestamp: value.timestamp,
    revocation: value.revocation,
    fieldMdp: value.fieldMdp,
    dss: value.dss,
  });
}

export function bindPdfSignatureProofToSource(
  value: unknown,
  source: PdfSourceRiskSnapshot,
): PdfSignatureProof {
  const proof = parsePdfSignatureProof(value);
  if (
    source.profile.certificateSignatures === "none" ||
    proof.dataItemId !== source.identity.dataItemId ||
    proof.source.sha256 !== source.identity.sha256 ||
    (source.profile.certificateSignatures === "present" &&
      proof.signatures.length === 0)
  ) {
    return rejectContract();
  }
  return proof;
}

/** The translator shape the presentation needs (next-intl's `t` fits). */
export type PdfSignatureProofTranslator = (
  key: string,
  values?: Record<string, string | number>,
) => string;

const PROOF_REASON_KEYS: Readonly<Record<string, string>> = {
  DUAL_CRYPTO_INVALID: "pdf.signatureProof.reasonDualCryptoInvalid",
  CRYPTO_DISAGREEMENT: "pdf.signatureProof.reasonCryptoDisagreement",
  CRYPTO_SET_DISAGREEMENT: "pdf.signatureProof.reasonCryptoDisagreement",
  EMPTY_TRUST_ROOTS: "pdf.signatureProof.reasonEmptyTrustRoots",
  IMPLEMENTATION_DUPLICATE: "pdf.signatureProof.reasonImplementationDuplicate",
  INVENTORY_DISAGREEMENT: "pdf.signatureProof.reasonInventoryDisagreement",
  NO_SIGNATURES: "pdf.signatureProof.reasonNoSignatures",
  UNSUPPORTED_PDF_STRUCTURE: "pdf.signatureProof.reasonUnsupportedStructure",
  WITNESS_SET_INVALID: "pdf.signatureProof.reasonWitnessSetInvalid",
  WITNESS_UNAVAILABLE: "pdf.signatureProof.reasonWitnessUnavailable",
};

function presentProofReason(
  reason: string | undefined,
  t: PdfSignatureProofTranslator,
): string {
  if (!reason) return t("pdf.signatureProof.reasonUnavailable");
  return t(
    Object.hasOwn(PROOF_REASON_KEYS, reason)
      ? PROOF_REASON_KEYS[reason]
      : "pdf.signatureProof.reasonInconclusive",
  );
}

function presentSignatureRows(
  proof: PdfSignatureProof | null,
  t: PdfSignatureProofTranslator,
): readonly PdfSignatureProofSignaturePresentation[] {
  return Object.freeze(
    (proof?.signatures ?? []).map((signature, index) => {
      const integrity =
        signature.integrity === "VALID"
          ? "intact"
          : signature.integrity === "INVALID"
            ? "mismatch"
            : "unknown";
      return Object.freeze({
        label: t("pdf.signatureProof.signatureLabel", { index: index + 1 }),
        integrity,
        integrityLabel:
          integrity === "intact"
            ? t("pdf.signatureProof.integrityIntact")
            : integrity === "mismatch"
              ? t("pdf.signatureProof.integrityMismatch")
              : t("pdf.signatureProof.integrityUnknown"),
        coverage: signature.coversCurrentFile ? "current" : "earlier",
        coverageLabel: signature.coversCurrentFile
          ? t("pdf.signatureProof.coverageCurrent")
          : t("pdf.signatureProof.coverageEarlier"),
      });
    }),
  );
}

function presentUnknownAxes(
  t: PdfSignatureProofTranslator,
): readonly PdfSignatureProofAxisPresentation[] {
  const labels = [
    ["certificateTrust", t("pdf.signatureProof.axisCertificateTrust")],
    ["timestamp", t("pdf.signatureProof.axisTimestamp")],
    ["revocation", t("pdf.signatureProof.axisRevocation")],
    ["fieldMdp", "FieldMDP"],
    ["dss", "DSS"],
  ] as const;
  const valueLabel = t("pdf.signatureProof.axisUnknown");
  return Object.freeze(
    labels.map(([key, label]) =>
      Object.freeze({ key, label, value: "unknown" as const, valueLabel }),
    ),
  );
}

export function presentPdfSignatureProof(
  proof: PdfSignatureProof | null,
  t: PdfSignatureProofTranslator,
): PdfSignatureProofPresentation {
  const verdicts = (proof?.signatures ?? []).map((signature) =>
    judgeSignature({
      digestIntact:
        signature.integrity === "VALID"
          ? true
          : signature.integrity === "INVALID"
            ? false
            : undefined,
    }),
  );
  const judgement = documentSignatureStatus(verdicts);
  const hasMismatch = proof?.signatures.some(
    (signature) => signature.integrity === "INVALID",
  );
  const allIntact =
    Boolean(proof?.signatures.length) &&
    proof?.signatures.every((signature) => signature.integrity === "VALID");
  const hasLaterRevision = proof?.signatures.some(
    (signature) => !signature.coversCurrentFile,
  );
  const unknownAxes = t("pdf.signatureProof.unknownAxes");
  const shared = {
    loadingHeadline: t("pdf.signatureProof.loadingHeadline"),
    loadingDetail: t("pdf.signatureProof.loadingDetail"),
    loadingIndicatorLabel: t("pdf.signatureProof.loadingIndicator"),
    panelTitle: t("pdf.signatureProof.panelTitle"),
    detailsLabel: proof
      ? t("pdf.signatureProof.detailsCount", {
          count: proof.signatures.length,
        })
      : t("pdf.signatureProof.details"),
    verificationTimeLabel: t("pdf.signatureProof.verificationTime"),
    verificationTime: proof?.validationTime ?? null,
    reasonLabel: presentProofReason(proof?.reason, t),
    downloadOriginalLabel: t("pdf.signatureProof.downloadOriginal"),
    downloadOriginalUnavailableLabel: t(
      "pdf.signatureProof.downloadOriginalUnavailable",
    ),
    signatures: presentSignatureRows(proof, t),
    axes: presentUnknownAxes(t),
  } as const;

  if (hasMismatch) {
    return Object.freeze({
      ...shared,
      integrity: "mismatch",
      judgement,
      indicatorLabel: t("pdf.signatureProof.mismatchIndicator"),
      headline: t("pdf.signatureProof.mismatchHeadline"),
      detail: t("pdf.signatureProof.mismatchDetail"),
      unknownAxes,
    });
  }
  if (allIntact) {
    return Object.freeze({
      ...shared,
      integrity: "intact",
      judgement,
      // Intact bytes do not establish signer trust, so never say "valid".
      indicatorLabel: t("pdf.signatureProof.intactIndicator"),
      headline: t("pdf.signatureProof.intactHeadline"),
      detail: hasLaterRevision
        ? t("pdf.signatureProof.intactEarlierRevision")
        : t("pdf.signatureProof.intactDetail"),
      unknownAxes,
    });
  }
  return Object.freeze({
    ...shared,
    integrity: "unknown",
    judgement,
    indicatorLabel: t("pdf.signatureProof.unknownIndicator"),
    headline: t("pdf.signatureProof.unknownHeadline"),
    detail: t("pdf.signatureProof.unknownDetail"),
    unknownAxes,
  });
}
