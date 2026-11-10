/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A typing session: a key went down, or text came in, so recently that the
 * next one is on its way. An automatic save in the middle of one is what made
 * typing in a long document stutter: Docs serialises the whole document and
 * reads its page layout (hundreds of milliseconds on a hundred pages), and
 * when that lands in the gap between two keys the next key waits behind it.
 *
 * So, as for the pen (lib/ink-writing-session), the editor's automatic
 * serialisation and save wait until typing pauses ({@link TYPING_IDLE_MS}),
 * then go once; a long burst saves anyway after {@link TYPING_MAX_HOLD_MS}.
 * Nothing waits that must not: Save, a hidden tab, a closing window and every
 * reader that needs the latest content (the AI, export, Word Count) take it at
 * once, as before.
 *
 * Only the browser's own (trusted) key and text events count, anywhere in the
 * window.
 */

import { createInkWritingSaveGate, type InkWritingSaveGate } from "./ink-writing-session";

/** No key or text for this long: the typing paused. */
export const TYPING_IDLE_MS = 800;
/** A save held this long goes even while the typing goes on. */
export const TYPING_MAX_HOLD_MS = 12_000;

let lastTypingAt: number | null = null;
let listening = false;

function listen(): void {
  if (listening || typeof window === "undefined") return;
  listening = true;
  const note = (event: Event) => {
    if (event.isTrusted) lastTypingAt = performance.now();
  };
  for (const type of ["keydown", "beforeinput", "compositionupdate"]) {
    window.addEventListener(type, note, { capture: true, passive: true });
  }
}

/** How long ago the last key or text came; Infinity before any. */
export function msSinceTyping(): number {
  listen();
  return lastTypingAt === null ? Infinity : performance.now() - lastTypingAt;
}

/** The gate an editor puts in front of its timer-driven serialisation or save. */
export function createTypingSaveGate(): InkWritingSaveGate {
  listen();
  return createInkWritingSaveGate({
    setTimeout: (run, ms) => globalThis.setTimeout(run, ms),
    clearTimeout: (handle) => globalThis.clearTimeout(handle as ReturnType<typeof setTimeout>),
    pen: () => ({ penDown: false, lastStrokeEndedAt: lastTypingAt }),
    idleMs: TYPING_IDLE_MS,
    maxHoldMs: TYPING_MAX_HOLD_MS,
    // Keys answer a save at once; its longest hold ends at the next look.
    latePauseMs: 0,
  });
}
