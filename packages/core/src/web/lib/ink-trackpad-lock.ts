/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
let target: HTMLElement | null = null;
let claimed = false;
let releaseTimer: ReturnType<typeof setTimeout> | null = null;
/** A lock no frame takes over by then is let go (the mode did not open). */
const UNCLAIMED_RELEASE_MS = 1500;

/** The element the pointer is locked to from the switch (made once). */
export function inkTrackpadLockTarget(): HTMLElement | null {
  if (typeof document === "undefined" || !document.body) return null;
  if (!target || !target.isConnected) {
    target = document.createElement("div");
    target.setAttribute("data-testid", "ink-trackpad-lock");
    target.setAttribute("aria-hidden", "true");
    target.style.cssText =
      "position:fixed;left:0;top:0;width:1px;height:1px;opacity:0;pointer-events:none;";
    document.body.appendChild(target);
  }
  return target;
}

/** Lock the pointer now, inside the click that turns the mode on. */
export function lockInkTrackpadPointer(): void {
  const element = inkTrackpadLockTarget();
  if (!element || typeof element.requestPointerLock !== "function") return;
  claimed = false;
  if (releaseTimer !== null) clearTimeout(releaseTimer);
  releaseTimer = setTimeout(() => {
    releaseTimer = null;
    if (!claimed && document.pointerLockElement === element) document.exitPointerLock?.();
  }, UNCLAIMED_RELEASE_MS);
  try {
    const request = element.requestPointerLock() as unknown;
    if (request instanceof Promise) request.catch(() => {});
  } catch {
    // No lock: the frame asks again, and falls back to pressing and drawing.
  }
}

/** The pointer is locked for the mode: to the frame or to the switch's target. */
export function inkTrackpadLockHeld(frame: Element | null): boolean {
  if (typeof document === "undefined") return false;
  const held = document.pointerLockElement;
  return held !== null && (held === frame || held === target);
}

/** The mode's frame has taken the lock over; it is no longer let go by itself. */
export function claimInkTrackpadLock(): void {
  claimed = true;
}
