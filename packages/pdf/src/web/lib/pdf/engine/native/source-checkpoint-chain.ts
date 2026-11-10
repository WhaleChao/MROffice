/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { sha256PdfBytes } from "./transaction-error";

/** Portable receipt binding. Full transaction validation stays on the server. */
export async function pdfNativeCheckpointChainSha256(collection: readonly unknown[]): Promise<string> {
  const record = (value: unknown): Record<string, unknown> =>
    value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
  const rows = collection.map(value => {
    const item = record(value), source = record(item.source), derived = record(item.derivedCandidate), checkpoint = record(item.sourceCheckpoint);
    return [...["id", "transactionId", "action", "kind", "schemaVersion"].map(key => item[key] ?? null),
      source.resourceId ?? null, source.resourceVersionId ?? null, source.sha256 ?? null,
      derived.sha256 ?? null, checkpoint.schemaVersion ?? null, checkpoint.xrefSections ?? null];
  });
  return sha256PdfBytes(new TextEncoder().encode(JSON.stringify(rows)));
}
