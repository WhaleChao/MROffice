/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Why the server cannot make a verified PDF/A copy, in the reader's words.
 *
 * The API reports these as English sentences (verified_pdfa_export.py); the
 * dialog must not show them raw to a reader in another language. Each maps to
 * one of a few reasons the reader can act on; anything unknown reads as the
 * generic "cannot export PDF/A for this PDF right now".
 */
export type PdfaUnavailableReason =
  | "toolsUnavailable"
  | "annotations"
  | "forms"
  | "dangerousActions"
  | "signed"
  | "encrypted"
  | "limit"
  | "sourceChanged"
  | "structure"
  | "busy"
  | "notVerified"
  | "unknown";

const RULES: ReadonlyArray<readonly [RegExp, PdfaUnavailableReason]> = [
  [/annotations require/i, "annotations"],
  [/interactive forms|xfa/i, "forms"],
  [/executable pdf actions|safe copy/i, "dangerousActions"],
  [/^signed pdfs/i, "signed"],
  [/^encrypted pdfs/i, "encrypted"],
  [/exceeds|scratch space/i, "limit"],
  [/changed (?:before|during)/i, "sourceChanged"],
  [/another pdf\/a export is running/i, "busy"],
  [/structure could not be authenticated|pdfium/i, "structure"],
  [
    /not installed|not available|is required|did not report|isolation|could not start|isolated pdf worker|definition|icc/i,
    "toolsUnavailable",
  ],
  [/conversion failed|verapdf|validation|could not be verified/i, "notVerified"],
];

export function classifyPdfaUnavailableReason(
  reason: string | null | undefined,
): PdfaUnavailableReason | null {
  const text = reason?.trim();
  if (!text) return null;
  for (const [pattern, kind] of RULES) {
    if (pattern.test(text)) return kind;
  }
  return "unknown";
}
