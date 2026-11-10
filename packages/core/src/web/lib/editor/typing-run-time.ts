/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { EditorView } from "@tiptap/pm/view";

/**
 * When the reader made the edit a view is applying now: the time of the input
 * event (a key, a character, a composition's end, a paste) it comes from, or
 * null when it comes from none (a command, a remote change, a later pass).
 *
 * Undo takes back a typing run in one press, as Word does. Both histories
 * tell runs apart by time: prosemirror-history by the gap between two
 * transactions' times (newGroupDelay), Yjs's undo manager by the gap between
 * the ends of two of its transactions (captureTimeout), 500 ms each. On a
 * long document one keystroke's own dispatch takes longer than that, and the
 * keys typed behind it (queued while it ran) each became an Undo step of
 * their own: one Cmd+Z took back one character on a 3-million-character page
 * while a short page's took the word (ab27 sac29/sac30). Measured from the
 * input events instead, the gap is the reader's own.
 */
export const TYPING_RUN_GAP_MS = 500;

interface InputMark {
  /** The event's time, on the Date.now() clock. */
  at: number;
  /** The wall time it was noted. */
  noted: number;
  /** How long after that it still dates an edit (see below). */
  lifetime: number;
}

const marks = new WeakMap<EditorView, InputMark>();
/** A key handled by a keymap edits in the task the key ran in. */
const KEY_MARK_LIFETIME_MS = 250;








const TEXT_MARK_LIFETIME_MS = 3_000;

function eventTime(event: Event): number {
  const now = Date.now();
  const origin = typeof performance !== "undefined" ? performance.timeOrigin : NaN;
  const stamp = event.timeStamp;
  if (!Number.isFinite(stamp) || stamp <= 0) return now;
  // Relative to the time origin (every current browser); an engine that
  // stamps events on the epoch clock instead (older WebKit, jsdom) as is.
  const at = stamp > 1e12 ? stamp : Number.isFinite(origin) ? origin + stamp : now;
  // A clock that does not agree (a timeStamp on another origin, a skewed
  // timeOrigin): the event is taken as now.
  return at <= now && now - at < 60_000 ? at : now;
}

/** The view is handling `event`: edits it makes in this task are dated by it. */
export function noteEditorInput(view: EditorView, event: Event): void {
  const key = event.type === "keydown";
  let at = eventTime(event);
  // A key's keydown comes before its beforeinput, whose text mark replaces
  // it; a key that only moves (no text) leaves its short mark. The text a
  // key puts in is that key's: its beforeinput runs in the key's own task
  // (the key's mark is still here), but the browser stamps it when it gets
  // to it. Keys typed behind a slow one wait in the queue carrying the time
  // they were typed, their beforeinput the time they were handled: on the
  // causal 2.8-million-character page that was a second after the first
  // key, and one Cmd+Z took back all but the first character (ab27 sac28,
  
  const before = marks.get(view);
  if (!key && before?.lifetime === KEY_MARK_LIFETIME_MS && Date.now() - before.noted <= KEY_MARK_LIFETIME_MS) {
    at = Math.min(at, before.at);
  }
  const mark: InputMark = {
    at, noted: Date.now(), lifetime: key ? KEY_MARK_LIFETIME_MS : TEXT_MARK_LIFETIME_MS,
  };
  marks.set(view, mark);
  if (key) {
    setTimeout(() => {
      if (marks.get(view) === mark) marks.delete(view);
    }, 0);
  }
}

/** A pointer press (a click on the page or a toolbar): no later edit is dated by the input before it. */
export function endEditorInput(view: EditorView): void {
  marks.delete(view);
}

export function editorInputTime(view: EditorView | null | undefined): number | null {
  if (!view) return null;
  const mark = marks.get(view);
  if (!mark) return null;
  if (Date.now() - mark.noted > mark.lifetime) {
    marks.delete(view);
    return null;
  }
  // A text mark dates one edit: the next is its own input's, or none's.
  if (mark.lifetime === TEXT_MARK_LIFETIME_MS) marks.delete(view);
  return mark.at;
}
