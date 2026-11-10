/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pen's gestures, switched in the pen menu's Settings (a note-taking
 * app's Pen Gestures): one set for every editor's pens, remembered on this
 * device, read by the editors at the moment a gesture would start.
 */

const KEY = "mrnote.inkPenGestures";

export interface InkPenGestures {
  /** Draw a line or shape and hold: it snaps to a straight line or a shape. */
  readonly holdShape: boolean;
  /** Scribble over ink with the pen to erase what it covers. */
  readonly scribbleErase: boolean;
  /** A scribble erases shapes and highlighter strokes too (else only writing). */
  readonly scribbleErasesShapes: boolean;
  /** Circle content with the pen to select it, as the lasso does. */
  readonly circleLasso: boolean;
}







export const INK_PEN_GESTURE_DEFAULTS: InkPenGestures = {
  holdShape: true,
  scribbleErase: false,
  scribbleErasesShapes: true,
  circleLasso: false,
};

let cached: InkPenGestures | null = null;
const listeners = new Set<() => void>();

export function inkPenGestures(): InkPenGestures {
  if (cached) return cached;
  let stored: Partial<Record<keyof InkPenGestures, unknown>> = {};
  try {
    const raw = typeof window === "undefined" ? null : window.localStorage.getItem(KEY);
    const parsed: unknown = raw ? JSON.parse(raw) : null;
    if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) stored = parsed as typeof stored;
  } catch {
    // Unreadable storage: the defaults.
  }
  const pick = (name: keyof InkPenGestures) =>
    typeof stored[name] === "boolean" ? (stored[name] as boolean) : INK_PEN_GESTURE_DEFAULTS[name];
  cached = {
    holdShape: pick("holdShape"),
    scribbleErase: pick("scribbleErase"),
    scribbleErasesShapes: pick("scribbleErasesShapes"),
    circleLasso: pick("circleLasso"),
  };
  return cached;
}

export function setInkPenGesture(name: keyof InkPenGestures, on: boolean): void {
  cached = { ...inkPenGestures(), [name]: on };
  try {
    window.localStorage.setItem(KEY, JSON.stringify(cached));
  } catch {
    // A private window keeps the choice for this page only.
  }
  for (const listener of listeners) listener();
}

/** For useSyncExternalStore: every palette and menu shows one state. */
export function subscribeInkPenGestures(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key !== KEY) return;
    cached = null;
    listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}
