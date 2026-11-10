/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Words of a complex script (Arabic, Hebrew, Thai, Devanagari, Bengali)
 * written as page text: the API shapes each line once, as the exporters do
 * (`/api/v1/pdf-text-shape`, app.services.pdf_text_shaping), and the content
 * append draws exactly that (PdfNativeContentAppendTextContent.shapedLines).
 */
import type { PdfNativeShapedTextPiece } from "./engine/native/content-append-transaction";

/** The scripts the API packages a face of (document_packaged_fonts). */
const COMPLEX_SCRIPT = /[֐-׿؀-ۿݐ-ݿࢠ-ࣿיִ-﷿ﹰ-﻿ऀ-ॿঀ-৿฀-๿]/u;

/** Whether `text` has words the bundled face cannot set as they read. */
export function pdfTextNeedsShaping(text: string): boolean {
  return COMPLEX_SCRIPT.test(text);
}



const STRONG_RTL = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/u;
const STRONG_LTR = /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF\u0900-\u0DFF\u0E00-\u0E7F\u3040-\u30FF\u3400-\u9FFF\uAC00-\uD7AF]/u;

/** The paragraph direction of `text` by its first strong letter (Unicode's
 * bidi rules P2 and P3, as Word and the browser take it for a paragraph
 * whose direction is not stated): an Arabic or Hebrew first letter reads
 * right to left, anything else left to right. */
export function pdfTextReadsRightToLeft(text: string): boolean {
  for (const character of text) {
    if (STRONG_RTL.test(character)) return true;
    if (STRONG_LTR.test(character)) return false;
  }
  return false;
}

export async function requestPdfTextShaping(
  lines: readonly string[],
  post: (path: string, body: unknown) => Promise<unknown>,
): Promise<readonly (readonly PdfNativeShapedTextPiece[])[] | null> {
  let answer: unknown;
  try {
    answer = await post("/api/v1/pdf-text-shape", { lines, rtl: pdfTextReadsRightToLeft(lines.join("\n")) });
  } catch {
    return null;
  }
  const shaped = (answer as { lines?: unknown } | null)?.lines;
  if (!Array.isArray(shaped) || shaped.length !== lines.length) return null;
  if (shaped.some((line) => !Array.isArray(line) || line.some((piece) =>
    !piece || typeof piece !== "object" || (piece as { kind?: unknown }).kind === "unshaped"))) return null;
  return shaped as readonly (readonly PdfNativeShapedTextPiece[])[];
}
