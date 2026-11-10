/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The output-loss review (the "what this new PDF changes" list) in the
 * reader's language.
 *
 * The save policy states each loss as an exact English sentence
 * (lib/pdf/save-policy.ts, OUTPUT_LOSSES) and several sentences share one
 * code, so each sentence has its own `console-notebooks` key,
 * `pdf.outputLoss.<FNV-1a of its UTF-8 bytes>`. A sentence the catalogue does
 * not know stays as written.
 */

export function pdfOutputLossKey(text: string): string {
  let hash = 0x811c9dc5;
  for (const byte of new TextEncoder().encode(text)) {
    hash ^= byte;
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return `pdf.outputLoss.${hash.toString(16).padStart(8, "0")}`;
}

/** A catalogue reader: next-intl's `t` (it has `has`). */
export interface PdfOutputLossTranslator {
  (key: string): string;
  has(key: string): boolean;
}

/** The loss sentence in the reader's language, else as written. */
export function localizePdfOutputLossText(
  text: string,
  t: PdfOutputLossTranslator,
): string {
  const key = pdfOutputLossKey(text);
  return t.has(key) ? t(key) : text;
}
