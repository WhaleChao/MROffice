/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect } from "react";

/** Whether `element` takes typed text (an input, a text area, an editable region). */
function takesText(element: Element | null): element is HTMLElement {
  if (!(element instanceof HTMLElement)) return false;
  if (element.isContentEditable) return true;
  if (element instanceof HTMLTextAreaElement) return true;
  if (element instanceof HTMLInputElement) {
    return !["button", "checkbox", "radio", "range", "color", "file", "submit", "reset", "image"].includes(element.type);
  }
  return false;
}

/** Let go of the text focus and the DOM selection inside `scope`. */
export function releaseTextFocusIn(scope: HTMLElement): void {
  const doc = scope.ownerDocument;
  const active = doc.activeElement;
  if (takesText(active) && scope.contains(active)) active.blur();
  const selection = doc.getSelection();
  if (selection && selection.rangeCount > 0 && selection.anchorNode && scope.contains(selection.anchorNode)) {
    selection.removeAllRanges();
  }
}

/**
 * Keep `scope` free of a text focus and a DOM selection while `armed`: let
 * them go when the tool is armed, and again whenever a field inside takes
 * the focus while it stays armed.
 */
export function useInkArmedFocusRelease(armed: boolean, scope: () => HTMLElement | null | undefined): void {
  useEffect(() => {
    if (!armed) return;
    const element = scope();
    if (!element) return;
    releaseTextFocusIn(element);
    const onFocusIn = (event: FocusEvent) => {
      if (takesText(event.target as Element | null)) releaseTextFocusIn(element);
    };
    element.addEventListener("focusin", onFocusIn);
    return () => element.removeEventListener("focusin", onFocusIn);
    // The scope is read when the tool is armed.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [armed]);
}
