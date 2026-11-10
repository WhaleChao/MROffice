/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > Text hands a new box to the native editor once its first words
 * are page text (PdfPageWindow beginCreatedTextHandoff). The editor opens on
 * the run that draws them -- alone, or, when they were written against older
 * words on the same baseline, the line they joined -- and the words typed
 * into the stand-in meanwhile go in where the written ones stand.
 */

export type InsertTextHandoffPlan = Readonly<{
  /** The stretch of the editor's text to replace (empty: an insertion). */
  replaceStart: number;
  replaceEnd: number;
  insert: string;
  /** The stand-in's caret, carried over into the editor's text. */
  caretStart: number;
  caretEnd: number;
}>;

/**
 * Where `written` (the words staged) stands in `fieldValue` (the editor's
 * text): at its end, else its start, else its last occurrence; null when it
 * is not there at all.
 */
export function insertTextHandoffAnchor(fieldValue: string, written: string): number | null {
  if (!written) return null;
  if (fieldValue.endsWith(written)) return fieldValue.length - written.length;
  if (fieldValue.startsWith(written)) return 0;
  const at = fieldValue.lastIndexOf(written);
  return at >= 0 ? at : null;
}

/**
 * The edit that puts `typed` (the stand-in's text, which began as `written`)
 * in place of `written` in the editor, touching nothing else of the line.
 * Typed on after the written words, only the rest is inserted after them;
 * otherwise the written words are replaced. Null when the written words are
 * not in the editor's text (the caller does not guess where they went).
 */
export function planInsertTextHandoff(input: Readonly<{
  fieldValue: string;
  written: string;
  typed: string;
  caretStart: number;
  caretEnd: number;
}>): InsertTextHandoffPlan | null {
  const { fieldValue, written } = input;
  const at = insertTextHandoffAnchor(fieldValue, written);
  if (at === null) return null;
  const typed = input.typed.startsWith(written) ? input.typed : input.typed.trimEnd();
  const clamp = (offset: number) => Math.max(0, Math.min(typed.length, offset));
  const caret = (offset: number) => at + clamp(offset);
  if (typed.startsWith(written)) {
    return Object.freeze({
      replaceStart: at + written.length,
      replaceEnd: at + written.length,
      insert: typed.slice(written.length),
      caretStart: caret(input.caretStart),
      caretEnd: caret(input.caretEnd),
    });
  }
  return Object.freeze({
    replaceStart: at,
    replaceEnd: at + written.length,
    insert: typed,
    caretStart: caret(input.caretStart),
    caretEnd: caret(input.caretEnd),
  });
}
