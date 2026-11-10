/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A wheel that steps through items one at a time (PowerPoint's Normal view:
 * a wheel over the slide goes to the next or previous slide).
 *
 * A mouse wheel steps once per notch. A trackpad sends a stream of small
 * deltas and then its momentum: the stream steps once, when it has moved
 * far enough, and nothing more until the wheel has been quiet for
 * WHEEL_GESTURE_IDLE_MS (lib/wheel-gesture), so one swipe is one slide and
 * its momentum does not run on through the deck.
 */
import { WHEEL_GESTURE_IDLE_MS } from "./wheel-gesture";

/** Trackpad travel (px) a gesture needs before it steps. */
export const WHEEL_STEP_TRAVEL_PX = 48;
/** Pixels a line-mode delta stands for. */
const LINE_PX = 16;

export interface WheelStepper {
  /** -1 (previous), 1 (next) or 0 for this wheel event. */
  step(event: Pick<WheelEvent, "deltaX" | "deltaY" | "deltaMode"> & { wheelDeltaY?: number }, now?: number): -1 | 0 | 1;
  /**
   * This wheel event was spent elsewhere (it scrolled a zoomed-in slide): the
   * trackpad swipe it belongs to, momentum included, does not step once the
   * scroll has reached its end; the next swipe does. A mouse wheel's next
   * notch at the end still steps, as PowerPoint's does.
   */
  hold(event: Pick<WheelEvent, "deltaY" | "deltaMode"> & { wheelDeltaY?: number }, now?: number): void;
  reset(): void;
}

/** Whether the event is one notch of a mouse wheel: the legacy wheelDelta of
 * a notch is a multiple of 120 in Chromium and WebKit, a trackpad's is not;
 * a line or page delta is a wheel's. */
export function isWheelNotch(event: Pick<WheelEvent, "deltaY" | "deltaMode"> & { wheelDeltaY?: number }): boolean {
  if (event.deltaMode !== 0) return true;
  const legacy = event.wheelDeltaY;
  return typeof legacy === "number" && legacy !== 0 && legacy % 120 === 0;
}

export function createWheelStepper(): WheelStepper {
  let quietAt = Number.NEGATIVE_INFINITY;
  let travel = 0;
  let stepped = false;
  let notches = false;
  return {
    step(event, now = performance.now()) {
      // Sideways scrolling is not a step.
      if (Math.abs(event.deltaX) > Math.abs(event.deltaY) || event.deltaY === 0) return 0;
      const direction: -1 | 1 = event.deltaY > 0 ? 1 : -1;
      const fresh = now >= quietAt;
      quietAt = now + WHEEL_GESTURE_IDLE_MS;
      // A gesture is a wheel's or a trackpad's by its first tick: a
      // trackpad delta that happens to look like a notch (40 px, 120 in
      // wheelDelta) inside a swipe or its momentum is not one.
      if (fresh) {
        notches = isWheelNotch(event);
        travel = 0;
        stepped = false;
      }
      if (notches && isWheelNotch(event)) return direction;
      if (stepped) return 0;
      const px = event.deltaMode === 1 ? event.deltaY * LINE_PX : event.deltaY;
      // A turn the other way starts the travel again.
      travel = Math.sign(travel) === Math.sign(px) ? travel + px : px;
      if (Math.abs(travel) < WHEEL_STEP_TRAVEL_PX) return 0;
      stepped = true;
      return direction;
    },
    hold(event, now = performance.now()) {
      if (now >= quietAt) notches = isWheelNotch(event);
      quietAt = now + WHEEL_GESTURE_IDLE_MS;
      travel = 0;
      stepped = true;
    },
    reset() {
      quietAt = Number.NEGATIVE_INFINITY;
      travel = 0;
      stepped = false;
      notches = false;
    },
  };
}
