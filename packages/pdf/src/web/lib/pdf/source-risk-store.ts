/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet, apiPut } from "@mrnote/core/web/lib/api.ts";
import {
  createPdfSourceRiskProfile,
  setPdfSourceRiskProfileStore,
  type PdfSourceRiskProfile,
  type PdfSourceRiskState,
} from "./save-policy";

/** Bumped when the inspection rules change: older rows are then not found. */
export const PDF_SOURCE_RISK_INSPECTOR_VERSION = 2;

const AXES = [
  "encryption",
  "certificateSignatures",
  "certifiedSignatures",
  "xfa",
  "optionalContent",
  "attachments",
  "taggedStructure",
  "pdfaClaim",
  "inspectionFailure",
] as const;

const SHA256 = /^[0-9a-f]{64}$/u;

function isState(value: unknown): value is PdfSourceRiskState {
  return value === "none" || value === "present" || value === "unknown";
}

/** A profile every reader may trust: every axis answered, no failure. */
export function pdfSourceRiskProfileIsClean(profile: PdfSourceRiskProfile): boolean {
  return AXES.every((axis) => profile[axis] === "none" || profile[axis] === "present") &&
    profile.inspectionFailure === "none";
}

/** The server's row as a profile, or null when it is not a clean profile. */
export function parseStoredPdfSourceRiskProfile(value: unknown): PdfSourceRiskProfile | null {
  if (!value || typeof value !== "object") return null;
  const raw = (value as { profile?: unknown }).profile;
  if (!raw || typeof raw !== "object") return null;
  const fields = raw as Record<string, unknown>;
  for (const axis of AXES) if (!isState(fields[axis])) return null;
  const profile = createPdfSourceRiskProfile({
    encryption: fields.encryption as PdfSourceRiskState,
    certificateSignatures: fields.certificateSignatures as PdfSourceRiskState,
    certifiedSignatures: fields.certifiedSignatures as PdfSourceRiskState,
    xfa: fields.xfa as PdfSourceRiskState,
    ...(isState(fields.xfaDynamic) ? { xfaDynamic: fields.xfaDynamic } : {}),
    optionalContent: fields.optionalContent as PdfSourceRiskState,
    attachments: fields.attachments as PdfSourceRiskState,
    taggedStructure: fields.taggedStructure as PdfSourceRiskState,
    pdfaClaim: fields.pdfaClaim as PdfSourceRiskState,
    inspectionFailure: fields.inspectionFailure as PdfSourceRiskState,
  });
  return pdfSourceRiskProfileIsClean(profile) ? profile : null;
}

export async function readStoredPdfSourceRiskProfile(
  sha256: string,
): Promise<PdfSourceRiskProfile | null> {
  if (!SHA256.test(sha256)) return null;
  try {
    const row = await apiGet<unknown>(
      `/api/v1/pdf/source-risk/${sha256}?inspector=${PDF_SOURCE_RISK_INSPECTOR_VERSION}`,
      { cache: "no-store" },
      { suppressUnauthorizedHandling: true },
    );
    return parseStoredPdfSourceRiskProfile(row);
  } catch {
    return null;
  }
}

export function storePdfSourceRiskProfile(
  sha256: string,
  profile: PdfSourceRiskProfile,
): void {
  if (!SHA256.test(sha256) || !pdfSourceRiskProfileIsClean(profile)) return;
  void apiPut<unknown>(`/api/v1/pdf/source-risk/${sha256}`, {
    inspector_version: PDF_SOURCE_RISK_INSPECTOR_VERSION,
    profile: { ...profile },
  }).catch(() => undefined);
}

/** Wire the store into the inspection cache (once, from the reader). */
export function installPdfSourceRiskStore(): void {
  setPdfSourceRiskProfileStore({
    read: readStoredPdfSourceRiskProfile,
    write: storePdfSourceRiskProfile,
  });
}
