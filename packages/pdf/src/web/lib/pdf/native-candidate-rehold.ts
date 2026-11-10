/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A native PDF save refused because the server does not hold the bytes its
 * bytes-publish record names (503 `pdf_native_bytes_unavailable`: an API
 * reload or release dropped the engine host's candidates). Nothing was
 * written; the bytes are handed over again and the save is sent again.
 */

export const PDF_NATIVE_BYTES_UNAVAILABLE = "pdf_native_bytes_unavailable";







export function isPdfNativeBytesUnavailable(input: Readonly<{
  thrown?: unknown;
  lastSaveFailureCode?: string | null;
}>): boolean {
  const thrown = input.thrown as { code?: unknown } | null | undefined;
  if (thrown && typeof thrown === "object" && thrown.code === PDF_NATIVE_BYTES_UNAVAILABLE) return true;
  return input.thrown === undefined && input.lastSaveFailureCode === PDF_NATIVE_BYTES_UNAVAILABLE;
}

export type PdfNativeCandidateHold = (input: Readonly<{
  engineSource: Readonly<{ pageId: string; resourceId: string; resourceVersionId: string }>;
  baseSha256: string;
  baseLength: number;
  candidateSha256: string;
  candidateBytes: Uint8Array;
  again?: boolean;
}>) => Promise<void>;

/** Hands the bytes the refused record names to the server again, from the
 * working copy's own copy of them (`again`: whatever the browser noted as
 * held is not). False when the record names no bytes this window has, or the
 * hand-over failed; the caller's retry then asks again. */
export async function reholdPdfNativeCandidate(input: Readonly<{
  pageId: string;
  collection: unknown;
  copyPendingOutput: (recordId: string) => Uint8Array | null;
  hold: PdfNativeCandidateHold;
}>): Promise<boolean> {
  const head = Array.isArray(input.collection) ? input.collection.at(-1) : null;
  const record = head && typeof head === "object" && !Array.isArray(head)
    ? (head as Record<string, unknown>) : null;
  if (!record || typeof record.id !== "string") return false;
  const derived = record.derivedCandidate as { sha256?: unknown; sourceLength?: unknown } | undefined;
  const source = record.source as { resourceId?: unknown; resourceVersionId?: unknown; sha256?: unknown } | undefined;
  if (typeof derived?.sha256 !== "string" || typeof derived.sourceLength !== "number" ||
    typeof source?.resourceId !== "string" || typeof source.resourceVersionId !== "string" ||
    typeof source.sha256 !== "string") return false;
  const candidate = input.copyPendingOutput(record.id);
  if (!candidate) return false;
  try {
    await input.hold({
      engineSource: { pageId: input.pageId, resourceId: source.resourceId, resourceVersionId: source.resourceVersionId },
      baseSha256: source.sha256, baseLength: derived.sourceLength,
      candidateSha256: derived.sha256, candidateBytes: candidate, again: true,
    });
    return true;
  } catch {
    return false;
  } finally {
    candidate.fill(0);
  }
}
