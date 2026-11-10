"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Double tap — a touch screen's double-click (touch audit P-10).
 *
 * An iPad does not dependably turn two taps into a `dblclick` (it depends on
 * whether the page lets it zoom on a double tap), so every "double-click to
 * open" entry point was a gamble under a finger. A surface asks, from the
 * tap's own `click`: is this the second tap of a pair — within
 * DOUBLE_TAP_MS of the last tap and DOUBLE_TAP_SLOP_PX of where it landed,
 * on the same target key? The first tap of a pair answers false (and is an
 * ordinary click); the second answers true once, so a third tap starts a
 * new pair. Only a finger's taps count (`isCompatMouseFromTouch`): a mouse
 * keeps the browser's own `dblclick`.
 */

import { isCompatMouseFromTouch } from "./pointer-capability";

export const DOUBLE_TAP_MS = 350;
export const DOUBLE_TAP_SLOP_PX = 12;

let last: { at: number; x: number; y: number; key: unknown } | null = null;
/** When a surface last answered a double tap itself (see the bridge). */
let answeredAt = Number.NEGATIVE_INFINITY;

function now(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}

/**
 * True when this click is a finger's second tap of a double tap on `key`
 * (default: the event's target). Call it from the tap's click handler.
 */
export function isDoubleTap(event: MouseEvent, key: unknown = event.target): boolean {
  if (!isCompatMouseFromTouch(event)) return false;
  const at = now();
  const previous = last;
  if (
    previous &&
    previous.key === key &&
    at - previous.at <= DOUBLE_TAP_MS &&
    Math.hypot(event.clientX - previous.x, event.clientY - previous.y) <= DOUBLE_TAP_SLOP_PX
  ) {
    last = null;
    noteDoubleTapAnswered();
    return true;
  }
  last = { at, x: event.clientX, y: event.clientY, key };
  return false;
}

/**
 * From a tap's click on an element that answers a double-click: on the second
 * tap of a pair, that element's own double-click, so its onDoubleClick runs
 * for a finger as for a mouse. True when it sent one.
 */
export function doubleClickOnDoubleTap(event: {
  nativeEvent: MouseEvent;
  currentTarget: EventTarget;
  clientX: number;
  clientY: number;
}): boolean {
  if (!isDoubleTap(event.nativeEvent, event.currentTarget)) return false;
  event.currentTarget.dispatchEvent(new MouseEvent("dblclick", {
    bubbles: true,
    cancelable: true,
    clientX: event.clientX,
    clientY: event.clientY,
  }));
  return true;
}

/**
 * A surface recognised a double tap and answered it itself (isDoubleTap,
 * createDoubleTapDetector, or its own pointer logic): the product-wide
 * bridge (lib/double-tap-bridge.ts) does not also send it a double-click.
 */
export function noteDoubleTapAnswered(): void {
  answeredAt = now();
}

/** When a surface last answered a double tap itself. */
export function lastDoubleTapAnsweredAt(): number {
  return answeredAt;
}
