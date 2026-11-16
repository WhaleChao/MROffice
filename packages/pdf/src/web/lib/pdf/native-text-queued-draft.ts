/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A text box clicked while an earlier edit is still saving opens at once on
 * the reader's own erasure (typing, caret and draft all live), but it cannot
 * be proved or written until the saved revision is the page's source. When
 * that revision lands the box is opened again on it -- the same line, found
 * by its text and position (native-text-continuation.ts) -- and what was
 * typed meanwhile is carried into the new box here: its text and caret, and
 * whether the reader had already finished it (clicked away), so it is
 * written exactly as if it had been typed there.
 *
 * Only a carry whose original text is the reopened box's text is applied:
 * a line the save changed is never overwritten with a draft of another text.
 * One change is followed: a line whose font was read right meanwhile
 * (native-font-reading-repair.ts) holds letters where the draft's original
 * had the blanks its font named them, and the draft's own change is put
 * onto the line as it reads now (rereadPdfNativeTextQueuedDraft). And a
 * paragraph's box is its line, though it holds the rows a line break apart
 * where the line reads them a blank apart or not apart at all: the draft's
 * change is put onto the rows (rowedPdfNativeTextQueuedDraft).
 *
 * A carry whose automatic reopen does not come (the saved revision's line
 * was not found, Edit mode was held) is never dropped: the reader is told
 * once that the draft is kept, and it lands in that line (same text, same
 * page) when the reader opens it again -- then without finishing on its own,
 * so the reader sees it before it is written.
 *
 * Keys typed ahead for a line that was not open at all (it was being written
 * under the reader's caret, and the reader moved on before it opened again)
 * are carried too, as `typedAhead`: there is no draft of a box to compare,
 * so the line is known by `typedAhead.line` and the keys go in at the
 * reopened box's caret, as they do when the line opens in time.
 */
import { insertPdfNativeTextTypeAhead } from "./native-text-continuation";

export type PdfNativeTextQueuedDraft = Readonly<{
  /** The line's text when the queued box opened. */
  originalText: string;
  /** The draft typed meanwhile. */
  value: string;
  selectionStart: number;
  selectionEnd: number;
  /** The reader clicked away: the reopened box finishes on its own. */
  finish: boolean;
  /** Epoch ms by which the automatic reopen should have taken the carry. */
  until: number;
  /** The page the line is on (its inline input's page host); a late carry
   * lands only on a box of this page. */
  pageNumber?: number;
  






  typedAhead?: Readonly<{ text: string; caretMove: "start" | "end" | null; line: (value: string) => boolean }>;
}>;

/** Whether the carry holds words of the reader's: a draft that is not the
 * line as it read, or keys typed ahead. Until its line takes it they are in
 * no file, no saved row and no open line, and the window says so (a line
 * given back as it was, to be opened again, holds none). */
export function pdfNativeTextQueuedDraftHoldsWords(carry: PdfNativeTextQueuedDraft): boolean {
  return carry.typedAhead ? carry.typedAhead.text.length > 0 : carry.value !== carry.originalText;
}

/** How long a late carry stays waiting for the reader to reopen its line. */
export const QUEUED_DRAFT_KEEP_MS = 30 * 60_000;

const INLINE_INPUT = "data-pdf-native-text-inline-input";

/**
 * The carry of a line that reads `current` now, where `current` is the
 * carry's original with characters in place of blanks: the same line, its
 * font read right since the draft was typed (a glyph the font's ToUnicode
 * called a space is named by its program). The draft's change -- what it put
 * between the text before its caret's stretch and the text after it -- is
 * kept, and the text around it is the line's as it reads now, so no letter
 * the page draws goes back to the blank the draft knew. Null where `current`
 * is any other text.
 */
export function rereadPdfNativeTextQueuedDraft(current: string, carry: PdfNativeTextQueuedDraft): PdfNativeTextQueuedDraft | null {
  const original = carry.originalText;
  if (current === original || current.length < original.length) return null;
  // Where each unit of the original stands in the line as it reads now.
  const at: number[] = [];
  let read = 0, named = 0;
  for (let index = 0; index < original.length; index += 1) {
    at.push(read);
    if (read >= current.length) return null;
    if (original[index] === current[read]) { read += 1; continue; }
    const point = current.codePointAt(read)!;
    const character = String.fromCodePoint(point);
    if (!/\s/u.test(original[index]) || /^[\s\p{C}]$/u.test(character)) return null;
    read += character.length;
    named += 1;
  }
  if (read !== current.length || named === 0) return null;
  at.push(current.length);
  return carriedOnto(current, carry, at);
}













export function rowedPdfNativeTextQueuedDraft(current: string, carry: PdfNativeTextQueuedDraft): PdfNativeTextQueuedDraft | null {
  const original = carry.originalText;
  if (current === original || !current.includes("\n")) return null;
  // Where each unit of the original stands in the rows (a blank a break
  // stands for, at the break).
  const at: number[] = [];
  let read = 0, rows = 0;
  for (let index = 0; index < original.length;) {
    if (read < current.length && original[index] === current[read]) { at.push(read); index += 1; read += 1; continue; }
    if (current[read] !== "\n") return null;
    // The break stands for the blanks the line holds here, but for the ones
    // the next row begins with.
    let blanks = 0, kept = 0;
    while (original[index + blanks] === " ") blanks += 1;
    while (current[read + 1 + kept] === " ") kept += 1;
    if (kept > blanks) return null;
    for (let owned = blanks - kept; owned > 0; owned -= 1) { at.push(read); index += 1; }
    read += 1;
    rows += 1;
  }
  if (read !== current.length || rows === 0) return null;
  at.push(current.length);
  return carriedOnto(current, carry, at);
}

/** The draft's change put onto `current`, where `at[i]` is the place in it
 * of the original's unit `i` (and `at[original.length]` its end). */
function carriedOnto(current: string, carry: PdfNativeTextQueuedDraft, at: readonly number[]): PdfNativeTextQueuedDraft {
  const original = carry.originalText;
  // The stretch the draft changed: it ends at the caret where the text
  // after the caret is the original's end (typed or deleted in one place),
  // and starts where the draft stops reading as the original.
  const value = carry.value;
  const caret = Math.max(0, Math.min(value.length, Math.max(carry.selectionStart, carry.selectionEnd)));
  let after = value.length - caret;
  if (after > original.length || value.slice(caret) !== original.slice(original.length - after)) {
    after = 0;
    while (after < value.length && after < original.length &&
      value[value.length - 1 - after] === original[original.length - 1 - after]) after += 1;
  }
  let before = 0;
  const most = Math.min(value.length, original.length) - after;
  while (before < most && value[before] === original[before]) before += 1;
  // An end of the stretch among the blanks one break stands for: the whole
  // break is the draft's, its blanks written out as the draft has them.
  while (before > 0 && at[before] === at[before - 1]) before -= 1;
  while (after > 0 && at[original.length - after] === at[original.length - after - 1]) after -= 1;
  const head = current.slice(0, at[before]);
  const tail = current.slice(at[original.length - after]);
  const next = head + value.slice(before, value.length - after) + tail;
  const moved = (offset: number) => {
    const clamped = Math.max(0, Math.min(value.length, offset));
    if (clamped <= before) return at[clamped];
    if (clamped >= value.length - after) return next.length - (current.length - at[original.length - (value.length - clamped)]);
    return head.length + (clamped - before);
  };
  return Object.freeze({ ...carry, originalText: current, value: next,
    selectionStart: moved(carry.selectionStart), selectionEnd: moved(carry.selectionEnd) });
}

/** Puts the carried draft into `input` through the element's own value setter
 * and an input event, so the editor's state follows as if it were typed.
 * Returns false (and changes nothing) when the box is another text (a line
 * whose font was read right meanwhile takes the draft read the same way; a
 * paragraph's box, its rows a line break apart, takes it on its rows). */
export function applyPdfNativeTextQueuedDraft(input: HTMLTextAreaElement, carry: PdfNativeTextQueuedDraft): boolean {
  if (carry.typedAhead) {
    if (!carry.typedAhead.line(input.value)) return false;
    insertPdfNativeTextTypeAhead(input, carry.typedAhead.text, carry.typedAhead.caretMove);
    return true;
  }
  if (input.value !== carry.originalText) {
    const reread = rereadPdfNativeTextQueuedDraft(input.value, carry) ?? rowedPdfNativeTextQueuedDraft(input.value, carry);
    return reread !== null && applyPdfNativeTextQueuedDraft(input, reread);
  }
  if (carry.value !== carry.originalText) {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")?.set;
    if (setter) setter.call(input, carry.value); else input.value = carry.value;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }
  const clamp = (offset: number) => Math.max(0, Math.min(carry.value.length, offset));
  input.setSelectionRange(clamp(carry.selectionStart), clamp(carry.selectionEnd));
  return true;
}

/** Waits (one frame at a time) for the reopened inline input -- focused,
 * writable, still holding the line's original text, and not `skip` (the
 * queued box's own input) -- then applies the carry and calls
 * `onApplied(true, late)`. Past `carry.until` (no automatic reopen) it calls
 * `onLate()` once and keeps waiting, for a box of the same page, until
 * `QUEUED_DRAFT_KEEP_MS` more have passed; `onApplied(false, true)` then. The
 * returned function cancels the wait. */
export function deliverPdfNativeTextQueuedDraft(
  carry: PdfNativeTextQueuedDraft,
  onApplied: (applied: boolean, late: boolean) => void,
  root: Pick<Document, "activeElement"> = document,
  skip: Element | null = null,
  onLate?: () => void,
): () => void {
  let frame = 0;
  let timer: ReturnType<typeof setTimeout> | 0 = 0;
  let done = false;
  let late = false;
  const keepUntil = carry.until + QUEUED_DRAFT_KEEP_MS;
  const stop = () => {
    done = true;
    if (frame) cancelAnimationFrame(frame);
    if (timer) clearTimeout(timer);
    frame = 0;
    timer = 0;
  };
  const onPage = (input: HTMLTextAreaElement) => carry.pageNumber === undefined ||
    input.closest("[data-page-number]")?.getAttribute("data-page-number") === String(carry.pageNumber);
  const attempt = () => {
    frame = 0;
    timer = 0;
    if (done) return;
    const now = Date.now();
    if (!late && now >= carry.until) {
      late = true;
      try { onLate?.(); } catch { /* the notice is best effort; the draft stays */ }
    }
    if (now >= keepUntil) { stop(); onApplied(false, true); return; }
    const active = root.activeElement;
    // Only the reopened box takes the carry: the queued box being retired
    // (or any other box) still holds another value and is waited past.
    if (active instanceof HTMLTextAreaElement && active.isConnected && active !== skip &&
      active.getAttribute(INLINE_INPUT) === "true" && !active.readOnly &&
      (!late || onPage(active)) &&
      applyPdfNativeTextQueuedDraft(active, carry)) {
      stop();
      onApplied(true, late);
      return;
    }
    // Every frame while the automatic reopen is due; a late carry waits for a
    // reader's click, so a few checks a second are enough.
    if (late) timer = setTimeout(attempt, 250);
    else frame = requestAnimationFrame(attempt);
  };
  frame = requestAnimationFrame(attempt);
  return stop;
}
