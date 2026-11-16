/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { inspectPdfSourceRiskProfileAhead } from "./save-policy";
import { installPdfSourceRiskStore } from "./source-risk-store";

/** Larger files are inspected by their reader alone (a copy of the bytes is
 * held for the inspection). */
export const PDF_SOURCE_RISK_AHEAD_MAX_BYTES = 32 * 1024 * 1024;

export function inspectPdfSourceRiskAhead(bytes: Uint8Array): void {
  if (typeof window === "undefined" || bytes.byteLength === 0 || bytes.byteLength > PDF_SOURCE_RISK_AHEAD_MAX_BYTES) return;
  try {
    // The shared store of profiles, as the reader installs it: what is
    // inspected here is looked for there first and handed back to it.
    installPdfSourceRiskStore();
    void inspectPdfSourceRiskProfileAhead(bytes).catch(() => undefined);
  } catch {
    // The reader inspects.
  }
}
