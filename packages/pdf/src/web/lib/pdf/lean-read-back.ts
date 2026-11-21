/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfTextHasRightToLeftRun } from "./engine/native/reader-order";







export function pdfNativeFormatWithLeanReadByReader<Format extends object>(format: Format | null | undefined):
  (Format & Readonly<{ leanReadBack: "reader" }>) | null {
  if (!format || typeof (format as { slanted?: unknown }).slanted !== "boolean") return null;
  return { ...format, leanReadBack: "reader" as const };
}

/**
 * Whether a page reads back as the same boxes: as many, each with the words
 * it had. (In whatever order: two boxes side by side on one baseline may be
 * listed the other way round once one of them leans, which changes no box.)
 */
export function pdfNativePageBoxesReadTheSame(before: readonly string[], after: readonly string[],
  /** "words": a box that gained or lost a blank at either END of its text is
   * the same box (see pdfNativeWrittenPageReadsBack). */
  by: "text" | "words" = "text"): boolean {
  if (before.length !== after.length) return false;
  const read = (texts: readonly string[]) => (by === "words" ? texts.map((text) => text.trim()) : [...texts]).sort();
  const left = read(before), right = read(after);
  return left.every((text, index) => text === right[index]);
}











export function pdfNativeFormatMaySetWordsAgain(format: unknown, boxText: string | null | undefined): boolean {
  return Boolean(format) && typeof format === "object" && (format as { syntheticBold?: unknown }).syntheticBold === false &&
    typeof boxText === "string" && pdfTextHasRightToLeftRun(boxText);
}




export function pdfNativeFormatWithWordsReadByReader<Format extends object>(format: Format | null | undefined,
  boxText: string | null | undefined): (Format & Readonly<{ wordsReadBack: "reader" }>) | null {
  return format && pdfNativeFormatMaySetWordsAgain(format, boxText) ? { ...format, wordsReadBack: "reader" as const } : null;
}

















export async function pdfNativeWrittenPageReadsBack(
  read: (bytes: Uint8Array) => Promise<readonly string[]>,
  before: Uint8Array,
  after: Uint8Array,
  by: "text" | "words" = "text",
): Promise<boolean | null> {
  try {
    const [was, now] = await Promise.all([read(before), read(after)]);
    return was.length ? pdfNativePageBoxesReadTheSame(was, now, by) : null;
  } catch {
    return null;
  }
}

/** (The name a lean's read-back asks it by.) */
export const pdfNativeLeanReadsBack = pdfNativeWrittenPageReadsBack;
