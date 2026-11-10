"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

import { noteDoubleTapAnswered } from "./double-tap";
import { yieldTouchHold } from "./touch-hold";

/*
 * Touch gestures shared by the Office editors (Microsoft's touch conventions:
 * Office Touch Guide, Excel for iPad / Android touch guide).
 *
 * Every piece here is input-agnostic about WHAT it drives; each editor keeps
 * its own adapter (the sheet's: spreadsheet/useSpreadsheetTouchGestures):
 *
 * - touch input mode: whether touch affordances (selection grips, a 44px fill
 *   grip) are shown. On for a device whose primary pointer is coarse, and on
 *   any device from the moment a finger presses until a mouse presses again,
 *   as Office on a touch laptop shows its grips after a tap and drops them on
 *   a click.
 * - long press: a finger held still for TOUCH_LONG_PRESS_MS (moved no more
 *   than TOUCH_LONG_PRESS_SLOP_PX) becomes the editor's drag; a finger that
 *   moves first is left to the browser, which scrolls.
 * - holding a gesture: once a drag is the editor's, the page must not scroll
 *   under it (`touch-action: none` on the surface, a non-passive touchmove
 *   that cancels, pointer capture on the surface) and the global press-and-
 *   hold menu (touch-hold) must not also fire.
 * - flick and double tap detectors, and the 44px touch target size.
 */

/** Apple's HIG and Microsoft's touch guidance: 44 × 44 CSS px per target. */
export const TOUCH_TARGET_MIN_PX = 44;
export const TOUCH_LONG_PRESS_MS = 300;
export const TOUCH_LONG_PRESS_SLOP_PX = 8;
export const TOUCH_DOUBLE_TAP_MS = 300;
export const TOUCH_DOUBLE_TAP_SLOP_PX = 24;
/** Release speed (CSS px per ms) above which a grip drag is a flick. */
export const TOUCH_FLICK_MIN_SPEED = 0.6;
/** Only the last stretch of the drag counts toward the release speed. */
const FLICK_SAMPLE_WINDOW_MS = 100;

export const COARSE_POINTER_MEDIA_QUERY = "(pointer: coarse)";

/** A finger, or a pen on a device that cannot hover. */
export function isTouchGesturePointer(event: { pointerType?: string }): boolean {
  if (event.pointerType === "touch") return true;
  if (event.pointerType !== "pen") return false;
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") return false;
  return window.matchMedia("(hover: none)").matches;
}

// ---------------------------------------------------------------------------
// Touch input mode
// ---------------------------------------------------------------------------

let lastPressWasTouch = false;
let installed = false;
const listeners = new Set<() => void>();

function coarseDevice(): boolean {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") return false;
  return window.matchMedia(COARSE_POINTER_MEDIA_QUERY).matches;
}

function notify() {
  for (const listener of listeners) listener();
}

function install() {
  if (installed || typeof window === "undefined") return;
  installed = true;
  window.addEventListener(
    "pointerdown",
    (event) => {
      const touch = isTouchGesturePointer(event);
      // a mouse press ends touch mode; a pen that can hover leaves it as it is
      if (!touch && event.pointerType !== "mouse") return;
      if (touch === lastPressWasTouch) return;
      lastPressWasTouch = touch;
      notify();
    },
    { capture: true, passive: true },
  );
  if (typeof window.matchMedia === "function") {
    window.matchMedia(COARSE_POINTER_MEDIA_QUERY).addEventListener?.("change", notify);
  }
}

/** Whether touch affordances are on right now. */
export function touchInputModeActive(): boolean {
  return lastPressWasTouch || coarseDevice();
}

function subscribeTouchInputMode(onChange: () => void): () => void {
  install();
  listeners.add(onChange);
  return () => listeners.delete(onChange);
}

/**
 * Live touch input mode for render decisions. The server and the first paint
 * assume a mouse, so markup never differs between server and client.
 */
export function useTouchInputMode(): boolean {
  return useSyncExternalStore(subscribeTouchInputMode, touchInputModeActive, () => false);
}

/** Test seam: forget the last press (jsdom keeps module state across tests). */
export function resetTouchInputModeForTests(): void {
  if (!lastPressWasTouch) return;
  lastPressWasTouch = false;
  notify();
}

// ---------------------------------------------------------------------------
// Long press
// ---------------------------------------------------------------------------

export type TouchLongPressAbort = "moved" | "released" | "cancelled" | "second-pointer";

export interface TouchLongPressOptions {
  readonly delayMs?: number;
  readonly slopPx?: number;
  /** The finger has been held still: the gesture is the editor's from here. */
  readonly onLongPress: () => void;
  /** The press ended before it became a long press ("released" is a tap). */
  readonly onAbort?: (reason: TouchLongPressAbort, event: PointerEvent | null) => void;
}

export interface TouchLongPressWatch {
  cancel(): void;
}

/**
 * Watch one pressed pointer for a long press. Window listeners only, all
 * passive: until the timer fires nothing is taken from the browser, so a
 * finger that moves first scrolls as it always did.
 */
export function watchTouchLongPress(
  press: { readonly pointerId: number; readonly clientX: number; readonly clientY: number },
  options: TouchLongPressOptions,
): TouchLongPressWatch {
  const slop = options.slopPx ?? TOUCH_LONG_PRESS_SLOP_PX;
  let done = false;
  const passive = { capture: true, passive: true } as const;
  const stop = () => {
    if (done) return;
    done = true;
    clearTimeout(timer);
    window.removeEventListener("pointermove", onMove, passive);
    window.removeEventListener("pointerup", onUp, passive);
    window.removeEventListener("pointercancel", onCancel, passive);
    window.removeEventListener("pointerdown", onDown, passive);
  };
  const abort = (reason: TouchLongPressAbort, event: PointerEvent | null) => {
    if (done) return;
    stop();
    options.onAbort?.(reason, event);
  };
  function onMove(event: PointerEvent) {
    if (event.pointerId !== press.pointerId) return;
    if (Math.hypot(event.clientX - press.clientX, event.clientY - press.clientY) > slop) {
      abort("moved", event);
    }
  }
  function onUp(event: PointerEvent) {
    if (event.pointerId === press.pointerId) abort("released", event);
  }
  function onCancel(event: PointerEvent) {
    if (event.pointerId === press.pointerId) abort("cancelled", event);
  }
  function onDown(event: PointerEvent) {
    if (event.pointerId !== press.pointerId) abort("second-pointer", event);
  }
  const timer = setTimeout(() => {
    if (done) return;
    stop();
    options.onLongPress();
  }, options.delayMs ?? TOUCH_LONG_PRESS_MS);
  window.addEventListener("pointermove", onMove, passive);
  window.addEventListener("pointerup", onUp, passive);
  window.addEventListener("pointercancel", onCancel, passive);
  window.addEventListener("pointerdown", onDown, passive);
  return { cancel: stop };
}

// ---------------------------------------------------------------------------
// Holding a gesture
// ---------------------------------------------------------------------------

/**
 * Make a running touch gesture the surface's own until the returned release
 * is called (or the pointer lifts): the surface gets `touch-action: none` and
 * the pointer's capture, every touchmove is cancelled so the page cannot
 * start a scroll or a zoom underneath, and the global press-and-hold menu
 * stands down for this pointer (an editor that wants a menu for a hold that
 * never moved opens it itself, on release).
 */
export function holdTouchGesture(surface: HTMLElement | null, pointerId: number): () => void {
  if (typeof window === "undefined") return () => {};
  yieldTouchHold(pointerId);
  const previousTouchAction = surface?.style.touchAction ?? "";
  if (surface) surface.style.touchAction = "none";
  try {
    surface?.setPointerCapture?.(pointerId);
  } catch {
    // the pointer already lifted (or this DOM has no capture)
  }
  const block = (event: TouchEvent) => {
    if (event.cancelable) event.preventDefault();
  };
  const end = (event: PointerEvent) => {
    if (event.pointerId === pointerId) release();
  };
  let released = false;
  function release() {
    if (released) return;
    released = true;
    window.removeEventListener("touchmove", block, true);
    window.removeEventListener("pointerup", end, true);
    window.removeEventListener("pointercancel", end, true);
    if (surface) surface.style.touchAction = previousTouchAction;
    try {
      if (surface?.hasPointerCapture?.(pointerId)) surface.releasePointerCapture(pointerId);
    } catch {
      // already released
    }
  }
  window.addEventListener("touchmove", block, { capture: true, passive: false });
  window.addEventListener("pointerup", end, true);
  window.addEventListener("pointercancel", end, true);
  return release;
}

/**
 * Swallow the click a finished drag's release would send (a touch release
 * on the element it pressed clicks it). One click, within `ms`.
 */
export function swallowNextClick(ms = 450): void {
  if (typeof window === "undefined") return;
  const until = Date.now() + ms;
  const swallow = (event: MouseEvent) => {
    window.removeEventListener("click", swallow, true);
    if (Date.now() > until) return;
    event.preventDefault();
    event.stopPropagation();
  };
  window.addEventListener("click", swallow, true);
  setTimeout(() => window.removeEventListener("click", swallow, true), ms);
}

/** A short vibration where the platform offers one (Android); else nothing. */
export function touchHaptic(): void {
  try {
    (navigator as Navigator & { vibrate?: (ms: number) => boolean }).vibrate?.(8);
  } catch {
    // not allowed here
  }
}

// ---------------------------------------------------------------------------
// Flick and double tap
// ---------------------------------------------------------------------------

export type TouchFlickDirection = "up" | "down" | "left" | "right";

export interface PointerVelocityTracker {
  add(clientX: number, clientY: number, time?: number): void;
  /** CSS px per ms over the last stretch of the drag. */
  velocity(): { vx: number; vy: number };
}

export function createPointerVelocityTracker(windowMs = FLICK_SAMPLE_WINDOW_MS): PointerVelocityTracker {
  const samples: { x: number; y: number; t: number }[] = [];
  const now = () => (typeof performance !== "undefined" ? performance.now() : Date.now());
  return {
    add(clientX, clientY, time = now()) {
      samples.push({ x: clientX, y: clientY, t: time });
      while (samples.length > 2 && time - samples[0]!.t > windowMs) samples.shift();
    },
    velocity() {
      if (samples.length < 2) return { vx: 0, vy: 0 };
      const first = samples[0]!;
      const last = samples[samples.length - 1]!;
      const dt = Math.max(1, last.t - first.t);
      return { vx: (last.x - first.x) / dt, vy: (last.y - first.y) / dt };
    },
  };
}

/** The screen direction of a flick, or null when the release was not one. */
export function touchFlickDirection(
  velocity: { vx: number; vy: number },
  minSpeed = TOUCH_FLICK_MIN_SPEED,
): TouchFlickDirection | null {
  const { vx, vy } = velocity;
  if (Math.hypot(vx, vy) < minSpeed) return null;
  if (Math.abs(vx) > Math.abs(vy)) return vx > 0 ? "right" : "left";
  return vy > 0 ? "down" : "up";
}

export interface DoubleTapDetector {
  /** Report a tap; true when it completes a double tap on the same target. */
  tap(clientX: number, clientY: number, key: string, time?: number): boolean;
  reset(): void;
}

export function createDoubleTapDetector(
  intervalMs = TOUCH_DOUBLE_TAP_MS,
  slopPx = TOUCH_DOUBLE_TAP_SLOP_PX,
): DoubleTapDetector {
  let last: { x: number; y: number; t: number; key: string } | null = null;
  const now = () => (typeof performance !== "undefined" ? performance.now() : Date.now());
  return {
    tap(clientX, clientY, key, time = now()) {
      const previous = last;
      if (
        previous &&
        previous.key === key &&
        time - previous.t <= intervalMs &&
        Math.hypot(clientX - previous.x, clientY - previous.y) <= slopPx
      ) {
        last = null;
        noteDoubleTapAnswered();
        return true;
      }
      last = { x: clientX, y: clientY, t: time, key };
      return false;
    },
    reset() {
      last = null;
    },
  };
}
