/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * "Reduce motion" for the app, on this device. On, the workspace carries
 * `data-reduce-motion` on <html> and styles/reduce-motion.css stills its
 * animations and transitions, whatever the operating system says; off, the
 * system setting alone decides, as before.
 */

export const REDUCE_MOTION_KEY = "mrnote.reduceMotion";
export const REDUCE_MOTION_ATTRIBUTE = "data-reduce-motion";

/** Whether motion is reduced right now: the app's switch or the system's. */
export function reducedMotionNow(): boolean {
  if (typeof document === "undefined") return false;
  if (document.documentElement.hasAttribute(REDUCE_MOTION_ATTRIBUTE)) return true;
  return (
    typeof window.matchMedia === "function" &&
    Boolean(window.matchMedia("(prefers-reduced-motion: reduce)")?.matches)
  );
}

export function loadReduceMotion(): boolean {
  try {
    return globalThis.localStorage?.getItem(REDUCE_MOTION_KEY) === "1";
  } catch {
    return false;
  }
}

export function applyReduceMotion(on: boolean): void {
  if (typeof document === "undefined") return;
  if (on) document.documentElement.setAttribute(REDUCE_MOTION_ATTRIBUTE, "");
  else document.documentElement.removeAttribute(REDUCE_MOTION_ATTRIBUTE);
}

export function saveReduceMotion(on: boolean): void {
  try {
    globalThis.localStorage?.setItem(REDUCE_MOTION_KEY, on ? "1" : "0");
  } catch {
    /* storage unavailable: the choice still holds for this page */
  }
  applyReduceMotion(on);
}
