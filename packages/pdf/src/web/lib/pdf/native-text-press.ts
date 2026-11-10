/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The last press on a source text box, as the reader saw it (WPS click
 * model: one press selects the box as a block, the second of a double click
 * puts the caret in).
 *
 * The editor for a box opens a moment after its first press: a proof is
 * started, a provisional session begun, React renders. The second press of
 * a quick double click lands before that, on the box itself, where the
 * reader can only join the activation already under way -- and the box
 * opened selected, so the keys that followed (Shift+Arrow) moved the block
 * instead of selecting text. The reader notes each press here with its click
 * count; the editor that opens for that box reads it, and one that is
 * already up hears it, and puts the caret in where the second press landed.
 * Intent only: it grants nothing the activation has not proved.
 */

export type PdfNativeTextPress = Readonly<{
  target: HTMLElement;
  /** The browser's click count for the press (2: a double click's second). */
  count: number;
  clientX: number;
  clientY: number;
  /** performance.now() when it was noted. */
  time: number;
}>;

/** How long a double click still speaks for the box it was made on: the
 * editor opens within this in all but a stalled page. */
export const PDF_NATIVE_TEXT_DOUBLE_PRESS_MS = 2_000;

let last: PdfNativeTextPress | null = null;
const listeners = new Set<(press: PdfNativeTextPress) => void>();

export function notePdfNativeTextPress(
  target: HTMLElement,
  count: number,
  clientX: number,
  clientY: number,
): void {
  const press = Object.freeze({
    target, count: Math.max(1, Math.floor(count) || 1), clientX, clientY, time: performance.now(),
  });
  last = press;
  if (press.count >= 2) for (const listener of [...listeners]) listener(press);
}

/** The double click just made on `target` that no editor has taken yet. */
export function readPdfNativeTextDoublePress(target: Element | null | undefined): PdfNativeTextPress | null {
  return target && last && last.target === target && last.count >= 2 &&
    performance.now() - last.time <= PDF_NATIVE_TEXT_DOUBLE_PRESS_MS ? last : null;
}

/** The editor took the double click: it does not speak again. */
export function takePdfNativeTextDoublePress(press: PdfNativeTextPress): void {
  if (last === press) last = null;
}

/** Hears every double click noted from now on. */
export function subscribePdfNativeTextDoublePress(
  listener: (press: PdfNativeTextPress) => void,
): () => void {
  listeners.add(listener);
  return () => { listeners.delete(listener); };
}
