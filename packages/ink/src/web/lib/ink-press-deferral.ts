/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pencilScribblePossible } from "@mrnote/core/web/lib/ink-scribble-guard.ts";


















export interface InkPressHandlers {
  readonly move: (event: PointerEvent) => void;
  readonly up: (event: PointerEvent) => void;
  readonly cancel: (event: PointerEvent) => void;
}

/** A point of the pen, in client pixels. */
export interface InkPressPoint {
  readonly clientX: number;
  readonly clientY: number;
}

export interface DeferredInkPress {
  /** The setup has not run yet; it will not, and the kept events are dropped. */
  readonly cancel: () => void;
}

const REPLAYED: ReadonlyArray<"pointermove" | "pointerup" | "pointercancel"> = ["pointermove", "pointerup", "pointercancel"];

/**
 * Runs `setUp` now, or (on iPadOS WebKit) in the press's next task with the
 * pointer's events replayed; returns the deferred press to cancel, or null
 * when `setUp` already ran.
 */
/**
 * A press this far (client px) from its first sample, and this near its
 * last, did not go down there: the samples say where.
 */
export const INK_PRESS_STALE_PX = 24;












export function inkPendingStrokeSamples<S extends { clientX: number; clientY: number; t?: number }>(
  press: { clientX: number; clientY: number },
  samples: readonly S[],
): { samples: S[]; begin: S | null } {
  const timed = samples.length > 1 && samples.every((sample) => typeof sample.t === "number" && sample.t > 0);
  const ordered = timed
    ? samples.map((sample, index) => ({ sample, index })).sort((a, b) => a.sample.t! - b.sample.t! || a.index - b.index).map(({ sample }) => sample)
    : [...samples];
  const first = ordered[0];
  const last = ordered[ordered.length - 1];
  // Stale: far from where the samples begin, and where they end -- the pen's
  // point when the page got to the press. A press far from a coarse first
  // sample alone is only a fast start.
  const stale =
    first !== undefined &&
    last !== undefined &&
    first !== last &&
    Math.hypot(first.clientX - press.clientX, first.clientY - press.clientY) > INK_PRESS_STALE_PX &&
    Math.hypot(last.clientX - press.clientX, last.clientY - press.clientY) <= INK_PRESS_STALE_PX;
  return stale ? { samples: ordered.slice(1), begin: first } : { samples: ordered, begin: null };
}

export function runInkPressSetup(
  pointerId: number,
  setUp: (downAt: InkPressPoint | null) => InkPressHandlers | null,
  defer: boolean = pencilScribblePossible(),
  press?: InkPressPoint,
): DeferredInkPress | null {
  if (!defer || typeof window === "undefined") {
    setUp(null);
    return null;
  }
  const kept: PointerEvent[] = [];
  let ended = false;
  const keep = (event: Event) => {
    // Not `instanceof PointerEvent`: a test DOM has no PointerEvent.
    if (ended || !("pointerId" in event) || (event as PointerEvent).pointerId !== pointerId) return;
    kept.push(event as PointerEvent);
    if (event.type !== "pointermove") ended = true;
  };
  for (const type of REPLAYED) window.addEventListener(type, keep, true);
  const release = () => {
    for (const type of REPLAYED) window.removeEventListener(type, keep, true);
  };
  const timer = setTimeout(() => {
    release();
    // The moves in the order the pen made them, and where it went down when
    // the press itself came too late to say (inkPendingStrokeSamples).
    const moves = kept
      .filter((event) => event.type === "pointermove")
      .map((event) => {
        const first = typeof event.getCoalescedEvents === "function" ? event.getCoalescedEvents()[0] : undefined;
        const at = first ?? event;
        return { clientX: at.clientX, clientY: at.clientY, t: event.timeStamp, event };
      });
    const ordered = press ? inkPendingStrokeSamples(press, moves) : null;
    const handlers = setUp(ordered?.begin ? { clientX: ordered.begin.clientX, clientY: ordered.begin.clientY } : null);
    if (!handlers) return;
    const replay = ordered
      ? [...(ordered.begin ? [ordered.begin] : []), ...ordered.samples].map((move) => move.event)
          .concat(kept.filter((event) => event.type !== "pointermove"))
      : kept;
    for (const event of replay) {
      if (event.type === "pointermove") handlers.move(event);
      else if (event.type === "pointerup") handlers.up(event);
      else handlers.cancel(event);
    }
  }, 0);
  return {
    cancel: () => {
      clearTimeout(timer);
      release();
    },
  };
}
