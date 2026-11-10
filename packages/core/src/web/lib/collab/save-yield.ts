/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A save's own turn to let the page breathe between its steps.
 *
 * A save of a causal page runs as a chain of steps (capture, persist,
 * prepare, digest, adopt), each now proportional to what changed. Chained
 * without a break they still ran as one task, and the pen's next samples
 * waited behind all of them together. Between steps the save gives the event
 * loop one task (the browser's own `scheduler.yield` where it has one), and
 * while a stroke is on the glass it waits for the lift, a moment at most, so
 * it never takes the pen's frames; it never waits for the pen in the air
 * between strokes, which would hold a long session's save forever.
 *
 * Every step that follows a yield works on what the save froze before its
 * first await (the request's bytes, body and page): a yield is no different
 * from the digest the chain already awaited there.
 */
import { inkPointerArbiter } from "../ink-pointer-policy";

/** The longest a save waits for a stroke to lift before it goes on. */
export const SAVE_STROKE_WAIT_MS = 600;

function nextTask(): Promise<void> {
  const scheduler = (globalThis as typeof globalThis & { scheduler?: { yield?: () => Promise<void> } }).scheduler;
  if (typeof scheduler?.yield === "function") return scheduler.yield();
  if (typeof MessageChannel === "function") {
    return new Promise<void>((resolve) => {
      const channel = new MessageChannel();
      channel.port1.onmessage = () => { channel.port1.close(); resolve(); };
      channel.port2.postMessage(null);
    });
  }
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

function clock(): number {
  return typeof performance === "undefined" ? Date.now() : performance.now();
}

/** Give the page a task, and the running stroke its lift (at most a moment). */
export async function yieldBetweenSaveSteps(): Promise<void> {
  await nextTask();
  const deadline = clock() + SAVE_STROKE_WAIT_MS;
  while (inkPointerArbiter.strokeRunning() && clock() < deadline) {
    await new Promise<void>((resolve) => setTimeout(resolve, 16));
  }
}
