/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfNativePartStart = "line-start" | "word-start" | "inside-word";

export function pdfNativePartStartOnItsLine(text: string, at: number,
  /** Where the part ends: the blanks it opens with are no start of its
   * words (a selection taken from the end of the line above begins with the
   * blank the line broke at, and read as begun inside that line's last
   * word). */
  end: number = at): PdfNativePartStart {
  const clamp = (value: number) => Math.max(0, Math.min(text.length, Math.trunc(value)));
  let index = clamp(at);
  for (const stop = clamp(end); index < stop && /\s/u.test(text[index]);) index += 1;
  const lineStart = text.lastIndexOf("\n", index - 1) + 1;
  if (!/\S/u.test(text.slice(lineStart, index))) return "line-start";
  return /\S/u.test(text.slice(index - 1, index)) ? "inside-word" : "word-start";
}
