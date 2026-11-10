"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useSyncExternalStore } from "react";

/**
 * 键盘输入触发编辑 -- the tick box on WPS's 文字编辑 tab (on by default).
 *
 * WPS selects a text block with a single click (its frame, the 文字编辑 tab,
 * every format on the whole block, no caret) and puts the caret in with a
 * double click. With the box ticked, a key typed on a selected block starts
 * editing it, the key going after the block's text; unticked, a selected
 * block takes no typing until it is double-clicked (or Enter / F2).
 *
 * WPS keeps the choice in the app, not in the document, so it is kept per
 * device here, as Draw with Touch is (lib/draw-with-touch.ts).
 */

export const PDF_TEXT_KEYBOARD_INPUT_KEY = "mrnote.pdfTextKeyboardInput";
/** Ticked, as WPS ships it. */
export const PDF_TEXT_KEYBOARD_INPUT_DEFAULT = true;

const listeners = new Set<() => void>();
// When storage throws (private mode), the page's own choice.
let memory: boolean | null = null;

export function loadPdfTextKeyboardInput(): boolean {
  try {
    const stored = globalThis.localStorage?.getItem(PDF_TEXT_KEYBOARD_INPUT_KEY);
    if (stored === "1") return true;
    if (stored === "0") return false;
  } catch {
    /* storage unavailable: the default */
  }
  return PDF_TEXT_KEYBOARD_INPUT_DEFAULT;
}

export function savePdfTextKeyboardInput(on: boolean): void {
  try {
    globalThis.localStorage?.setItem(PDF_TEXT_KEYBOARD_INPUT_KEY, on ? "1" : "0");
    memory = null;
  } catch {
    /* storage unavailable: the choice still holds for this page */
    memory = on;
  }
  for (const listener of listeners) listener();
}

/** The choice in force now (for event handlers outside React). */
export function readPdfTextKeyboardInput(): boolean {
  return memory ?? loadPdfTextKeyboardInput();
}

/** Follows the box: this page's own changes and another tab's (storage). */
export function subscribePdfTextKeyboardInput(listener: () => void): () => void {
  listeners.add(listener);
  const onStorage = (event: StorageEvent) => {
    if (event.key === null || event.key === PDF_TEXT_KEYBOARD_INPUT_KEY) listener();
  };
  if (typeof window !== "undefined") window.addEventListener("storage", onStorage);
  return () => {
    listeners.delete(listener);
    if (typeof window !== "undefined") window.removeEventListener("storage", onStorage);
  };
}

/** Live state of the box for render decisions. Server render: the default. */
export function usePdfTextKeyboardInput(): boolean {
  return useSyncExternalStore(
    subscribePdfTextKeyboardInput,
    readPdfTextKeyboardInput,
    () => PDF_TEXT_KEYBOARD_INPUT_DEFAULT,
  );
}
