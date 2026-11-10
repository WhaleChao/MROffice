/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A wheel quiet this long has ended its gesture (ms). */
export const WHEEL_GESTURE_IDLE_MS = 140;

export interface WheelGestureLatch<T> {
  /** The gesture's decision: `decide()` on its first tick, kept after. */
  hold(event: WheelEvent, decide: () => T): T;
  /** Forget the gesture (the next tick starts a new one). */
  reset(): void;
}

export function createWheelGestureLatch<T>(): WheelGestureLatch<T> {
  let decided = false;
  let value: T;
  let until = Number.NEGATIVE_INFINITY;
  return {
    hold(_event, decide) {
      const now = performance.now();
      if (!decided || now >= until) {
        value = decide();
        decided = true;
      }
      until = now + WHEEL_GESTURE_IDLE_MS;
      return value;
    },
    reset() {
      decided = false;
      until = Number.NEGATIVE_INFINITY;
    },
  };
}
