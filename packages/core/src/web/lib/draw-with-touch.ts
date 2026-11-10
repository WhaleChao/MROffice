"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";








































export const DRAW_WITH_TOUCH_KEY = "mrnote.drawWithTouch";
/** On, as Word / PowerPoint for iPad and Android ship it ("Draw with Touch"). */
export const DRAW_WITH_TOUCH_DEFAULT = true;
/**
 * Where the pen was last seen on this page, kept across reloads of the tab.
 * While a pen is "present" (seen within {@link PEN_PRESENCE_MS}) and the
 * reader has not chosen Draw with Touch themselves, a finger pans and only
 * the pen inks, as a tablet with its pencil behaves; with no pen about, the
 * finger inks. The reader's own switch (on or off) always wins.
 */
export const PEN_SEEN_AT_KEY = "mrnote.penSeenAt";
export const PEN_PRESENCE_MS = 30 * 60 * 1000;

const listeners = new Set<() => void>();
// When storage throws (private mode), the page's own choice.
let memory: boolean | null = null;

/** The reader's own choice: true/false when they flipped the switch, null for automatic. */
export function loadDrawWithTouchChoice(): boolean | null {
  try {
    const stored = globalThis.localStorage?.getItem(DRAW_WITH_TOUCH_KEY);
    if (stored === "1") return true;
    if (stored === "0") return false;
  } catch {
    /* storage unavailable: automatic */
  }
  return null;
}

let penSeenAtMemory = Number.NEGATIVE_INFINITY;

function penSeenAt(): number {
  try {
    const stored = globalThis.sessionStorage?.getItem(PEN_SEEN_AT_KEY);
    if (stored) {
      const at = Number(stored);
      if (Number.isFinite(at)) return Math.max(at, penSeenAtMemory);
    }
  } catch {
    /* storage unavailable: this page's memory */
  }
  return penSeenAtMemory;
}

/** True while a pen has touched or hovered this page within {@link PEN_PRESENCE_MS}. */
export function penPresent(now: number = Date.now()): boolean {
  return now - penSeenAt() < PEN_PRESENCE_MS;
}

/** A pen was seen (any pointer event of type "pen"): finger drawing yields to it. */
export function notePenSeen(now: number = Date.now()): void {
  const wasPresent = penPresent(now);
  penSeenAtMemory = Math.max(penSeenAtMemory, now);
  try {
    globalThis.sessionStorage?.setItem(PEN_SEEN_AT_KEY, String(now));
  } catch {
    /* storage unavailable: this page's memory */
  }
  if (!wasPresent) for (const listener of listeners) listener();
}

/** Forget the pen (tests, or a reader who put it away). */
export function forgetPen(): void {
  penSeenAtMemory = Number.NEGATIVE_INFINITY;
  try {
    globalThis.sessionStorage?.removeItem(PEN_SEEN_AT_KEY);
  } catch {
    /* storage unavailable */
  }
  for (const listener of listeners) listener();
}

/**
 * Whether a finger inks right now: the reader's switch when set, otherwise
 * automatic (a finger inks until a pen shows up, then pans).
 */
export function loadDrawWithTouch(): boolean {
  const choice = loadDrawWithTouchChoice();
  if (choice !== null) return choice;
  return !penPresent();
}

export function saveDrawWithTouch(on: boolean): void {
  try {
    globalThis.localStorage?.setItem(DRAW_WITH_TOUCH_KEY, on ? "1" : "0");
  } catch {
    /* storage unavailable: the choice still holds for this page */
    memory = on;
  }
  for (const listener of listeners) listener();
}

function current(): boolean {
  return memory ?? loadDrawWithTouch();
}

let penWatcherInstalled = false;
/**
 * Watch the page for a pen once (capture phase, passive): the first pen
 * event switches fingers to panning on every ink surface at once.
 */
export function installPenPresenceWatcher(): void {
  if (penWatcherInstalled || typeof window === "undefined") return;
  penWatcherInstalled = true;
  const onPointer = (event: PointerEvent) => {
    if (event.pointerType === "pen") notePenSeen();
  };
  window.addEventListener("pointerdown", onPointer, { capture: true, passive: true });
  window.addEventListener("pointermove", onPointer, { capture: true, passive: true });
}

/** Follows the switch: this page's own changes and another tab's (storage). */
export function subscribeDrawWithTouch(listener: () => void): () => void {
  installPenPresenceWatcher();
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === DRAW_WITH_TOUCH_KEY) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}

/** Live switch state for render decisions. Server render: the default. */
export function useDrawWithTouch(): boolean {
  return useSyncExternalStore(subscribeDrawWithTouch, current, () => DRAW_WITH_TOUCH_DEFAULT);
}

/**
 * Whether an ink surface takes this press. Nothing inks unless an ink tool
 * (pen, highlighter, eraser) is picked (`toolArmed`); then a pen or a mouse
 * always does, a finger only while Draw with Touch is on.
 */
export function inkAcceptsPointer(
  event: { pointerType?: string },
  toolArmed: boolean,
  drawWithTouch: boolean = current(),
): boolean {
  if (event.pointerType === "pen") notePenSeen();
  if (!toolArmed) return false;
  return event.pointerType !== "touch" || drawWithTouch;
}





























export function installPenInkTouchGuard(
  node: HTMLElement,
  inking: (target: Element) => boolean = () => true,
): () => void {
  let penDownAt = Number.NEGATIVE_INFINITY;
  const onPointerDown = (event: PointerEvent) => {
    if (event.pointerType === "pen") penDownAt = event.timeStamp;
  };
  const onTouchStart = (event: TouchEvent) => {
    if (!(event.target instanceof Element) || !inking(event.target)) return;
    const stylus = Array.from(event.changedTouches).some(
      (touch) => (touch as Touch & { touchType?: string }).touchType === "stylus",
    );
    // The pen's own touchstart follows its pointerdown within a frame.
    if (stylus || Math.abs(event.timeStamp - penDownAt) < 100) notePenSeen();
  };
  const options: AddEventListenerOptions = { capture: true, passive: true };
  node.addEventListener("pointerdown", onPointerDown, options);
  node.addEventListener("touchstart", onTouchStart, options);
  return () => {
    node.removeEventListener("pointerdown", onPointerDown, options);
    node.removeEventListener("touchstart", onTouchStart, options);
  };
}

/**
 * Whether this device has a touch screen at all: the Draw tab shows Draw
 * with Touch only where a finger can draw, and Draw with Trackpad only where
 * there is no touch screen (Office for Mac has the one, Office for iPad and
 * Windows tablets the other); a desktop never shows both.
 */
export function deviceHasTouch(): boolean {
  return typeof navigator !== "undefined" && (navigator.maxTouchPoints ?? 0) > 0;
}

const noSubscription = () => () => {};

/** `deviceHasTouch` for render; the server and the first paint say no. */
export function useDeviceHasTouch(): boolean {
  return useSyncExternalStore(noSubscription, deviceHasTouch, () => false);
}
