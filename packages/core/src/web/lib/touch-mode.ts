"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

import { TOUCH_MODE_ATTRIBUTE, TOUCH_MODE_KEY } from "./touch-mode-init";

// The spacing rules travel with the switch: every surface that can turn it
// on (the editors' rails live on the dev galleries too) has them.
import "../styles/touch-mode.css";

/**
 * "Touch mode" for the editors, on this device — Office's Touch/Mouse Mode
 * ("Touch/Mouse Mode changes the spacing of the buttons on the ribbon …
 * for the touch screen, the buttons are spaced farther apart so they're
 * easier to tap", Microsoft Support, "Turn Touch Mode on or off"). It is the
 * reader's own switch; a device that can be touched starts with it on
 * (`deviceHasTouch`), and the reader turns it off or on as in Office, where
 * the Quick Access Toolbar's hand button does.
 *
 * On, the workspace carries `data-touch-mode` on <html> and
 * styles/touch-mode.css spaces the shared chrome (the ribbon, the dialogs'
 * steppers, the window controls) — the same faces, farther apart; nothing
 * is hidden, nothing is rearranged — and every finger size in the site
 * (styles/touch-sizing.css and the touch blocks of each stylesheet) keys on
 * it. It is the one switch for sizing: off, an iPad gets the mouse's sizes;
 * what a finger needs because it cannot hover stays on the device's media
 * queries. Saved per device, like "Reduce motion" (lib/reduce-motion), and
 * applied before the first paint (lib/touch-mode-init).
 */

export { TOUCH_MODE_ATTRIBUTE, TOUCH_MODE_KEY };

const listeners = new Set<() => void>();






export function deviceHasTouch(): boolean {
  if (typeof navigator !== "undefined" && (navigator.maxTouchPoints ?? 0) > 0) return true;
  try {
    return globalThis.matchMedia?.("(any-pointer: coarse)").matches ?? false;
  } catch {
    return false;
  }
}

export function loadTouchMode(): boolean {
  let saved: string | null | undefined;
  try {
    saved = globalThis.localStorage?.getItem(TOUCH_MODE_KEY);
  } catch {
    saved = null;
  }
  if (saved === "1") return true;
  if (saved === "0") return false;
  return deviceHasTouch();
}

export function applyTouchMode(on: boolean): void {
  if (typeof document === "undefined") return;
  if (on) document.documentElement.setAttribute(TOUCH_MODE_ATTRIBUTE, "");
  else document.documentElement.removeAttribute(TOUCH_MODE_ATTRIBUTE);
}

export function saveTouchMode(on: boolean): void {
  try {
    globalThis.localStorage?.setItem(TOUCH_MODE_KEY, on ? "1" : "0");
  } catch {
    /* storage unavailable: the choice still holds for this page */
  }
  applyTouchMode(on);
  for (const listener of listeners) listener();
}

/** Follows the switch: this page's own changes and another tab's (storage). */
export function subscribeTouchMode(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === TOUCH_MODE_KEY) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}

/**
 * A window's resize handles in touch mode: react-rnd's edges are 10px and its
 * corners 20px; a finger gets 24px edges and 44px corners, centred on the
 * same border. Passed as `resizeHandleStyles`, merged over the defaults.
 */
export const TOUCH_WINDOW_RESIZE_HANDLE_STYLES = {
  top: { height: 24, top: -12 },
  bottom: { height: 24, bottom: -12 },
  left: { width: 24, left: -12 },
  right: { width: 24, right: -12 },
  topLeft: { width: 44, height: 44, top: -22, left: -22 },
  topRight: { width: 44, height: 44, top: -22, right: -22 },
  bottomLeft: { width: 44, height: 44, bottom: -22, left: -22 },
  bottomRight: { width: 44, height: 44, bottom: -22, right: -22 },
} as const;

/** Live touch-mode state for render decisions. Server render: off. */
export function useTouchMode(): boolean {
  return useSyncExternalStore(subscribeTouchMode, loadTouchMode, () => false);
}
