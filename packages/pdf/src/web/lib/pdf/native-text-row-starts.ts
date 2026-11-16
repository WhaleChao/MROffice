/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfNativeTextInputLayout } from "./native-text-input-layout";
















const rowStartsByEdit = new Map<string, readonly number[]>();
const KEPT_EDITS = 24;
const editKey = (originalText: string, replacementText: string) => `${originalText}\u0000${replacementText}`;

/** The offset, in the layout's text, each row after the first starts at. */
export function pdfNativeTextInputRowStarts(layout: PdfNativeTextInputLayout): readonly number[] {
  let shift = 0;
  return Object.freeze(layout.breaks.map((line) => {
    // The break shows as one character and stands for its own text (the
    // blank the row ended in, or nothing).
    const start = line.offset + shift + line.text.length;
    shift += line.text.length - 1;
    return start;
  }));
}

/** Keep (or, with null, forget) the rows of an edit of `originalText` to
 * `replacementText`. */
export function notePdfNativeTextRowStarts(originalText: string, replacementText: string,
  rowStarts: readonly number[] | null): void {
  const key = editKey(originalText, replacementText);
  rowStartsByEdit.delete(key);
  if (!rowStarts?.length || originalText === replacementText ||
    rowStarts.some((start, index) => !Number.isSafeInteger(start) || start <= (rowStarts[index - 1] ?? 0) ||
      start >= replacementText.length)) return;
  rowStartsByEdit.set(key, Object.freeze([...rowStarts]));
  while (rowStartsByEdit.size > KEPT_EDITS) rowStartsByEdit.delete(rowStartsByEdit.keys().next().value!);
}

/** The rows the editor last showed that edit in, if it did. */
export function readPdfNativeTextRowStarts(originalText: string, replacementText: string): readonly number[] | null {
  return rowStartsByEdit.get(editKey(originalText, replacementText)) ?? null;
}
