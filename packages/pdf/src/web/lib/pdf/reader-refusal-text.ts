/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isDoNotTrackEnabled, sendErrorBeacon } from "@mrnote/core/web/lib/errors/report-beacon.ts";

/** Scripts a window's language is written in that no English sentence is. */
const NON_LATIN_UI = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}]/u;
const ASCII_SENTENCE = /^[\x20-\x7E]+$/;
const REPEAT_AFTER_MS = 10 * 60 * 1000;
const lastReported = new Map<string, number>();

function reportInternalRefusal(message: string, now: number): void {
  try {
    if (process.env.NODE_ENV !== "production") {
      console.debug("[pdf-refusal] shown in the reader's language: " + message.slice(0, 300));
    }
    if (typeof window === "undefined" || isDoNotTrackEnabled()) return;
    const key = message.slice(0, 200);
    const last = lastReported.get(key);
    if (last !== undefined && now - last < REPEAT_AFTER_MS) return;
    lastReported.set(key, now);
    if (lastReported.size > 200) lastReported.clear();
    sendErrorBeacon({
      name: "PdfRefusalShownGeneric",
      message: key,
      path: "/pdf",
      source: "window",
      window_type: "pdf",
    });
  } catch {
    // Diagnostics never stand in the reader's way.
  }
}

/** `message` as the reader is shown it: an engine sentence in English
 * becomes `pdf.editRefusedGeneric` in a window not written in Latin
 * letters; anything else is kept. */
export function pdfReaderRefusalText(
  message: string,
  t: (key: string) => string,
  now = Date.now(),
): string {
  if (!message) return message;
  const generic = t("pdf.editRefusedGeneric");
  if (!NON_LATIN_UI.test(generic)) return message;
  if (!ASCII_SENTENCE.test(message) || (message.match(/[A-Za-z]{2,}/g)?.length ?? 0) < 3) return message;
  reportInternalRefusal(message, now);
  return generic;
}










export function pdfReaderRefusalReasonText(
  message: string,
  t: (key: string) => string,
  said: string = pdfReaderRefusalText(message, t),
): string {
  return said === t("pdf.editRefusedGeneric") ? t("pdf.viewer.textUsesFontPdfContent") : said;
}

/** Test seam. */
export function resetPdfReaderRefusalReportsForTests(): void {
  lastReported.clear();
}
