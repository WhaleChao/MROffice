/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PdfNativeTextPressKept = Readonly<{
  /** The press is over or superseded: nothing more is sent for it. */
  cancel: () => void;
}>;

export function keepPdfNativeTextClickOfRemovedPress(view: Window & typeof globalThis, press: Readonly<{
  target: Element;
  /** The press's `detail`, which the click carries (its count of clicks). */
  detail: number;
  /** Whether a click at this element belongs where the press was. */
  isSamePlace: (under: Element) => boolean;
}>): PdfNativeTextPressKept {
  let live = true;
  let stopDroppingClick: (() => void) | null = null;
  const released = (event: MouseEvent) => {
    view.removeEventListener("mouseup", released, true);
    if (!live || event.button !== 0 || press.target.isConnected) return;
    const { clientX, clientY, screenX, screenY, shiftKey, metaKey, ctrlKey, altKey } = event;
    const under = view.document.elementFromPoint?.(clientX, clientY);
    if (!under || !press.isSamePlace(under)) return;
    under.dispatchEvent(new view.MouseEvent("click", {
      bubbles: true, cancelable: true, composed: true, button: 0, detail: press.detail,
      clientX, clientY, screenX, screenY, shiftKey, metaKey, ctrlKey, altKey,
    }));
    // The engine's own, if it sends one, comes in this same task.
    const drop = (click: Event) => {
      if (!click.isTrusted) return;
      click.preventDefault();
      click.stopImmediatePropagation();
    };
    view.addEventListener("click", drop, true);
    const timer = view.setTimeout(() => stopDroppingClick?.(), 0);
    stopDroppingClick = () => {
      stopDroppingClick = null;
      view.clearTimeout(timer);
      view.removeEventListener("click", drop, true);
    };
  };
  view.addEventListener("mouseup", released, true);
  return {
    cancel: () => {
      live = false;
      view.removeEventListener("mouseup", released, true);
      stopDroppingClick?.();
    },
  };
}
