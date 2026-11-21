/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** How long after a pause's write the line may take to open again by itself
 * (the revision read in, the line found on it and proved). */
export const PDF_CREATED_TEXT_PAUSE_REOPEN_MS = 15_000;
/** An editor closed with nothing written and opened again on the same line
 * at once (taken over on the revision an edit just staged) is one editor. */
export const PDF_CREATED_TEXT_CLOSE_GRACE_MS = 1_500;

/** A line as its editor opened on it: its words and where it starts (page
 * shares). */
export type PdfCreatedTextLine = Readonly<{ text?: string | null; x?: number; y?: number }> | null | undefined;

export type PdfCreatedTextTyping = {
  /** The edit of the working copy that wrote the box's first words: the
   * step it opens takes every later write of the typing. */
  readonly anchor: string;
  /** The words of the box's line as the typing last left it (no white
   * space: a reader of the page spaces them its own way). */
  words: string;
  /** Where that line starts, once an editor has opened on it. */
  origin: Readonly<{ x: number; y: number }> | null;
  /** The line's editor has closed: the typing goes on only if an editor is
   * open on the line again by then (ms). Null while one is open. */
  resumeBy: number | null;
};



const SAME_LINE_SHARE = 0.012;

const wordsOf = (text: string | null | undefined): string => (text ?? "").replace(/\s+/gu, "").normalize("NFC");

const originOf = (line: PdfCreatedTextLine): Readonly<{ x: number; y: number }> | null =>
  line && typeof line.x === "number" && typeof line.y === "number" && Number.isFinite(line.x) && Number.isFinite(line.y)
    ? Object.freeze({ x: line.x, y: line.y }) : null;

/** The typing a new box's first words open (`anchor` the edit that wrote
 * them); null when nothing was written. */
export function beginPdfCreatedTextTyping(anchor: string | null | undefined, written: string): PdfCreatedTextTyping | null {
  const words = wordsOf(written);
  return anchor && words ? { anchor, words, origin: null, resumeBy: null } : null;
}

/** The line editor opened on the first words: alone they are its line, or
 * they joined older words on the same baseline -- that line is the box's
 * then, as the editor reads it. */
export function pdfCreatedTextTypingOpened(typing: PdfCreatedTextTyping, line: PdfCreatedTextLine): void {
  const words = wordsOf(line?.text);
  if (words.includes(typing.words)) typing.words = words;
  typing.origin = originOf(line);
}

/** Whether an editor open on `line` is typing in the box: the line holds
 * the words the typing last left, where it left them. */
export function pdfCreatedTextTypingTakes(typing: PdfCreatedTextTyping | null | undefined, line: PdfCreatedTextLine): boolean {
  if (!typing || !line || !typing.words || wordsOf(line.text) !== typing.words) return false;
  const origin = originOf(line);
  return !typing.origin || !origin ||
    (Math.abs(origin.x - typing.origin.x) <= SAME_LINE_SHARE && Math.abs(origin.y - typing.origin.y) <= SAME_LINE_SHARE);
}

/** The editor of the box's line wrote `text` (it opened on `line`): the
 * line is read with those words from here on. */
export function pdfCreatedTextTypingWrote(typing: PdfCreatedTextTyping, text: string, line?: PdfCreatedTextLine): void {
  typing.words = wordsOf(text);
  typing.origin = originOf(line) ?? typing.origin;
}

/** The line's editor closed: after a pause's write (`graceMs` the time the
 * line has to open again by itself), or with nothing written. */
export function pdfCreatedTextTypingClosed(typing: PdfCreatedTextTyping, now: number, graceMs: number): void {
  if (typing.resumeBy === null) typing.resumeBy = now + graceMs;
}

/** An editor is open on a line (or commits one): false when the typing was
 * over by then -- its editor closed and none opened in time. */
export function pdfCreatedTextTypingResumes(typing: PdfCreatedTextTyping, now: number): boolean {
  if (typing.resumeBy === null) return true;
  if (now > typing.resumeBy) return false;
  typing.resumeBy = null;
  return true;
}













export function pdfCreatedTextUndoTakesBox(typing: PdfCreatedTextTyping | null | undefined, line: PdfCreatedTextLine,
  own: Readonly<{ steps: number; composing: boolean }>): boolean {
  return pdfCreatedTextTypingTakes(typing, line) && !own.composing && own.steps <= 1;
}
