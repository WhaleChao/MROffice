/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeTextDecorationEvidence, PdfNativeTextDecorationHistory } from "./native-text-decoration";

export function fail(message: string): never {
  throw new Error(message);
}

/** What an Undo needs to take the rule back out. */
export function pdfNativeTextDecorationHistory(
  evidence: PdfNativeTextDecorationEvidence,
): PdfNativeTextDecorationHistory {
  if (evidence.direction !== "add") {
    fail("Native PDF text rule cannot be undone twice.");
  }
  return Object.freeze({
    offset: evidence.ruleOffset,
    ruleBytes: evidence.ruleBytes,
  });
}
