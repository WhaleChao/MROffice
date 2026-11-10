/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface TextRange {
  start: number;
  end: number;
}

export interface RemoteCaretHold {
  blockId: string;
  direction: "forward" | "backward" | "none";
  /** The box's words before the delivery, and the caret on them. */
  before: string;
  beforeRange: TextRange;
  /** The delivered words, and the caret carried over them. */
  after: string;
  afterRange: TextRange;
}

/** Long enough for the renders a delivery and the member's own last key
 * are drawn in; short enough that a later rewrite is not taken for them. */
export const REMOTE_CARET_HOLD_MS = 2_000;

let held: (RemoteCaretHold & { lastValue: string; until: number }) | null = null;

/** Hold `hold` (the box's value is `hold.after` now, with its caret). */
export function holdRemoteCaret(hold: RemoteCaretHold, now = Date.now()): void {
  held = { ...hold, beforeRange: { ...hold.beforeRange }, afterRange: { ...hold.afterRange },
    lastValue: hold.after, until: now + REMOTE_CARET_HOLD_MS };
}

export function releaseRemoteCaret(): void {
  held = null;
}

/**
 * After a commit of the component drawing `textarea` (the words of
 * `blockId`): the held caret put back when that commit rewrote the value to
 * the words before the delivery or the delivered ones. The range put back,
 * or null (nothing to do, or the hold ended).
 */
export function keepRemoteCaret(
  textarea: HTMLTextAreaElement | null,
  blockId: string,
  options: { composing?: boolean; now?: number } = {},
): TextRange | null {
  const hold = held;
  if (!hold || !textarea || hold.blockId !== blockId) return null;
  if (options.composing || textarea.ownerDocument.activeElement !== textarea ||
      (options.now ?? Date.now()) > hold.until) {
    held = null;
    return null;
  }
  const value = textarea.value;
  if (value !== hold.before && value !== hold.after) {
    // Typed meanwhile (or another delivery): the caret is the member's.
    held = null;
    return null;
  }
  const range = value === hold.after ? hold.afterRange : hold.beforeRange;
  const rewritten = value !== hold.lastValue;
  hold.lastValue = value;
  const start = textarea.selectionStart ?? value.length;
  const end = textarea.selectionEnd ?? value.length;
  if (start === range.start && end === range.end) return null;
  if (!rewritten) {
    // The value is as last seen and the caret elsewhere: the member moved it.
    held = null;
    return null;
  }
  textarea.setSelectionRange(range.start, range.end, hold.direction);
  return { ...range };
}
