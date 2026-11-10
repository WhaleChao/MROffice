"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

/**
 * Hover capability — the one place the product decides "is this a touch
 * gesture?" for UI that otherwise opens on hover.
 *
 * Two inputs, never the screen width (an iPad reaches the full product):
 *
 * - The DEVICE: `(hover: none)` — the primary pointer cannot hover (iPad,
 *   touch-only laptops). CSS uses the same query to show hover-revealed
 *   actions persistently (`@media (hover: none)` blocks in globals.css).
 * - The EVENT: `pointerType` of the actual gesture. A finger is always a
 *   touch; a pen counts as touch only on a device that cannot hover (an
 *   Apple Pencil on an iPad), so a pen tablet on a desktop keeps the mouse
 *   path. A mouse — including an iPad's trackpad — always keeps the
 *   desktop path, so desktop behaviour is unchanged by construction.
 *
 * Desktop: hover opens / previews and a click on the trigger does its own
 * thing. Touch: a tap on the trigger opens (a second tap or a tap outside
 * closes), and the trigger's own destination lives inside what it opens.
 */
export const NO_HOVER_MEDIA_QUERY = "(hover: none)";

/** True when the device's primary pointer can hover (SSR / no matchMedia: true). */
export function deviceCanHover(): boolean {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return true;
  }
  return !window.matchMedia(NO_HOVER_MEDIA_QUERY).matches;
}

/** True when a gesture of this pointer type should take the touch path. */
export function isTouchLikePointerType(
  pointerType: string | null | undefined,
): boolean {
  if (pointerType === "touch") return true;
  if (pointerType === "pen") return !deviceCanHover();
  return false;
}

/** `isTouchLikePointerType` for a (React or DOM) pointer event. */
export function isTouchLikePointer(event: { pointerType?: string }): boolean {
  return isTouchLikePointerType(event.pointerType);
}

/*
 * The pointer type of the most recent pointer that entered something or
 * pressed. A tap fires pointerover/pointerdown (pointerType "touch") BEFORE
 * the browser's compatibility mouseenter/mouseover/click, so a mouse-event
 * handler (`onMouseEnter`) can ask whether the mouse event it is handling is
 * really a finger. A real mouse always fires pointerover("mouse") first, so
 * for a mouse the answer is always false. One passive capture listener pair
 * on the window, installed when this module first loads in a browser.
 */
let lastPointerType = "";
/** When a finger (or a non-hovering pen) last pressed or lifted. */
let lastTouchLikeAt = Number.NEGATIVE_INFINITY;
function clock(): number {
  return typeof performance !== "undefined" ? performance.now() : Date.now();
}
if (typeof window !== "undefined") {
  const remember = (event: PointerEvent) => {
    lastPointerType = event.pointerType;
    if (event.type !== "pointerover" && isTouchLikePointerType(event.pointerType)) {
      lastTouchLikeAt = clock();
    }
  };
  window.addEventListener("pointerover", remember, { capture: true, passive: true });
  window.addEventListener("pointerdown", remember, { capture: true, passive: true });
  window.addEventListener("pointerup", remember, { capture: true, passive: true });
}

/**
 * How long after a finger lifts its compatibility mouse events can still
 * arrive (they follow the touchend at once; a busy main thread delays them).
 */
const COMPAT_MOUSE_WINDOW_MS = 1000;

/**
 * True when this mouse event is the browser's compatibility event for a tap
 * (mousedown / mouseup / mousemove / click after a touch), not a real mouse.
 *
 * A compatibility `mousedown` carries no pointerType, so a handler written
 * for the mouse (a drag-to-select, a "pointer is selecting" flag) cannot tell
 * it from a press of the mouse button and starts a gesture whose pointerup
 * has already gone by. Chrome says so outright (`sourceCapabilities`); other
 * browsers are answered by the pointer that last pressed or lifted: a real
 * mouse always sends its own pointerdown first, so a mouse is never mistaken.
 * Events that are pointer events themselves (pointerdown …) are not
 * compatibility events; a `click` that carries a pointerType (Chrome) is a
 * tap's when that pointerType is touch-like.
 */
export function isCompatMouseFromTouch(event: Event): boolean {
  if (typeof MouseEvent === "undefined" || !(event instanceof MouseEvent)) {
    return false;
  }
  if (event.type.startsWith("pointer")) return false;
  const pointerType = (event as { pointerType?: unknown }).pointerType;
  if (typeof pointerType === "string" && pointerType !== "") {
    return isTouchLikePointerType(pointerType);
  }
  const capabilities = (
    event as { sourceCapabilities?: { firesTouchEvents?: boolean } | null }
  ).sourceCapabilities;
  if (capabilities && typeof capabilities.firesTouchEvents === "boolean") {
    return capabilities.firesTouchEvents;
  }
  return (
    latestPointerIsTouchLike() &&
    clock() - lastTouchLikeAt < COMPAT_MOUSE_WINDOW_MS
  );
}

/** A coarse primary pointer (a finger): `(pointer: coarse)`. */
export const COARSE_POINTER_MEDIA_QUERY = "(pointer: coarse)";

/** True on a device whose primary pointer is coarse (SSR / no matchMedia: false). */
export function devicePointerIsCoarse(): boolean {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return false;
  }
  return window.matchMedia(COARSE_POINTER_MEDIA_QUERY).matches;
}

/**
 * True on a device a finger can reach at all, whatever its primary pointer:
 * a touch laptop's primary pointer is its trackpad (`pointer: fine`, it can
 * hover), yet a finger on its screen still pans the page unless the target
 * says otherwise. SSR / no matchMedia: false.
 */
export function deviceHasTouch(): boolean {
  if (typeof navigator !== "undefined" && (navigator.maxTouchPoints ?? 0) > 0) return true;
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return false;
  }
  return window.matchMedia("(any-pointer: coarse)").matches;
}

/**
 * True while the latest pointer was a finger (or a pen on a device that
 * cannot hover): a mouse event handled now is the compatibility event of a
 * tap, not a hover.
 */
export function latestPointerIsTouchLike(): boolean {
  return isTouchLikePointerType(lastPointerType);
}

/*
 * Tap-to-reveal — the touch stand-in for "hover a row to see its actions".
 *
 * Mark the hovered element with `data-touch-reveal` (spread
 * `TOUCH_REVEAL_PROPS`). A tap inside it gives it `data-touch-revealed`,
 * and a tap anywhere else takes that away, so its CSS reveals the actions
 * next to the existing hover rule:
 *
 *   .row:hover .actions,
 *   .row[data-touch-revealed] .actions { opacity: 1; pointer-events: auto; }
 *
 * The attribute is set on the tap's click, after the browser has chosen
 * what the tap hit, so the tap that reveals the actions never also presses
 * one of them. A mouse never sets it, so desktop rendering is untouched.
 */
export const TOUCH_REVEAL_ATTRIBUTE = "data-touch-reveal";
export const TOUCH_REVEALED_ATTRIBUTE = "data-touch-revealed";
export const TOUCH_REVEAL_PROPS = { [TOUCH_REVEAL_ATTRIBUTE]: "" } as const;

/**
 * For a control inside a tap-to-reveal host whose hover shows something the
 * control's click would otherwise skip past (a citation's preview): true when
 * this press is a finger's on a host not yet revealed, so its click should
 * only reveal. Read it at pointerdown — by the click the host is revealed.
 * A mouse, a hovering pen and the keyboard are never held back.
 */
export function touchPressOnlyReveals(event: {
  pointerType?: string;
  currentTarget: EventTarget | null;
}): boolean {
  if (!isTouchLikePointer(event)) return false;
  const element =
    typeof Element !== "undefined" && event.currentTarget instanceof Element
      ? event.currentTarget
      : null;
  const host = element?.closest(`[${TOUCH_REVEAL_ATTRIBUTE}]`);
  return Boolean(host && !host.hasAttribute(TOUCH_REVEALED_ATTRIBUTE));
}

/** How long a tap's click may take before the tap reveals without one. */
const TAP_REVEAL_CLICK_WAIT_MS = 120;
/** Travel (px) under which a touch press is a tap, not a scroll. */
const TAP_REVEAL_SLOP_PX = 10;

function revealAt(target: Element | null) {
  for (const revealed of document.querySelectorAll(`[${TOUCH_REVEALED_ATTRIBUTE}]`)) {
    if (!target || !revealed.contains(target)) {
      revealed.removeAttribute(TOUCH_REVEALED_ATTRIBUTE);
    }
  }
  target
    ?.closest(`[${TOUCH_REVEAL_ATTRIBUTE}]`)
    ?.setAttribute(TOUCH_REVEALED_ATTRIBUTE, "true");
}

if (typeof document !== "undefined") {
  /*
   * Safari on an iPad sends a tap's `click` (and its `mouseover`, so its
   * `:hover`) only to an element that listens for clicks itself or sits in
   * one that does: a tap on a message bubble's text, a row's label or a
   * card's picture sends none, so the reveal never happened there. A tap
   * whose click does not follow at once reveals from its pointerup instead.
   */
  let tapDown: { pointerId: number; x: number; y: number } | null = null;
  let tapTimer: ReturnType<typeof setTimeout> | null = null;
  const clearTapTimer = () => {
    if (tapTimer !== null) clearTimeout(tapTimer);
    tapTimer = null;
  };
  document.addEventListener(
    "pointerdown",
    (event) => {
      clearTapTimer();
      tapDown =
        isTouchLikePointer(event) && event.isPrimary
          ? { pointerId: event.pointerId, x: event.clientX, y: event.clientY }
          : null;
    },
    { capture: true, passive: true },
  );
  document.addEventListener(
    "pointerup",
    (event) => {
      const down = tapDown;
      tapDown = null;
      if (!down || down.pointerId !== event.pointerId) return;
      if (Math.hypot(event.clientX - down.x, event.clientY - down.y) > TAP_REVEAL_SLOP_PX) {
        return;
      }
      const target = event.target instanceof Element ? event.target : null;
      clearTapTimer();
      tapTimer = setTimeout(() => {
        tapTimer = null;
        revealAt(target);
      }, TAP_REVEAL_CLICK_WAIT_MS);
    },
    { capture: true, passive: true },
  );
  document.addEventListener(
    "pointercancel",
    () => {
      tapDown = null;
    },
    { capture: true, passive: true },
  );
  document.addEventListener(
    "click",
    (event) => {
      if (!latestPointerIsTouchLike()) return;
      // the tap's own click: it reveals, after the browser chose what it hit
      clearTapTimer();
      revealAt(event.target instanceof Element ? event.target : null);
    },
    { capture: true, passive: true },
  );
}

function subscribeHoverCapability(onChange: () => void): () => void {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return () => {};
  }
  const query = window.matchMedia(NO_HOVER_MEDIA_QUERY);
  query.addEventListener?.("change", onChange);
  return () => query.removeEventListener?.("change", onChange);
}

/**
 * Live `deviceCanHover()` for render decisions (e.g. whether to show an
 * explicit "open" affordance). Server render and first paint assume a
 * hovering device, so markup never differs between server and client.
 */
export function useDeviceCanHover(): boolean {
  return useSyncExternalStore(
    subscribeHoverCapability,
    deviceCanHover,
    () => true,
  );
}

function subscribeCoarsePointer(onChange: () => void): () => void {
  if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
    return () => {};
  }
  const query = window.matchMedia(COARSE_POINTER_MEDIA_QUERY);
  query.addEventListener?.("change", onChange);
  return () => query.removeEventListener?.("change", onChange);
}

/** Live `devicePointerIsCoarse()` for render decisions. Server render: false. */
export function useDevicePointerIsCoarse(): boolean {
  return useSyncExternalStore(subscribeCoarsePointer, devicePointerIsCoarse, () => false);
}
