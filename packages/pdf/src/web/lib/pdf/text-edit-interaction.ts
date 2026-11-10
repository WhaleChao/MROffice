/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_TEXT_EDIT_PENDING_OUTLINE =
  "1px dashed var(--console-accent, #0d9488)";

/** Let the input/selection feedback paint before expensive PDF preparation.
 * rAF alone is insufficient: its promise continuation runs before that paint.
 * The following task works on Safari too; a deadline covers background tabs. */
export function afterPdfTextInteractionPaint(): Promise<void> {
  return new Promise((resolve) => {
    let settled = false;
    let frame = 0;
    let task: ReturnType<typeof setTimeout> | undefined;
    const finish = () => {
      if (settled) return;
      settled = true;
      clearTimeout(deadline);
      if (task !== undefined) clearTimeout(task);
      if (frame && typeof cancelAnimationFrame === "function")
        cancelAnimationFrame(frame);
      resolve();
    };
    const deadline = setTimeout(finish, 64);
    if (typeof requestAnimationFrame === "function") {
      frame = requestAnimationFrame(() => {
        task = setTimeout(finish, 0);
      });
    } else task = setTimeout(finish, 0);
  });
}

/** Immediate interaction retirement, not a successful-save acknowledgement. */
export function retirePdfTextEditingChrome(
  frame: HTMLElement | null,
  input: HTMLTextAreaElement | null,
): () => void {
  const outline = frame?.style.outline ?? "";
  const handles =
    frame?.style.getPropertyValue("--pdf-text-handle-visibility") ?? "";
  const caret = input?.style.caretColor ?? "";
  const pointer = input?.style.pointerEvents ?? "";
  if (frame) {
    frame.style.outline = PDF_TEXT_EDIT_PENDING_OUTLINE;
    frame.style.setProperty("--pdf-text-handle-visibility", "hidden");
    frame.setAttribute("data-pdf-native-text-finishing", "true");
  }
  if (input) {
    input.style.caretColor = "transparent";
    input.style.pointerEvents = "none";
    input.setSelectionRange(input.selectionEnd, input.selectionEnd);
    input.blur();
  }
  return () => {
    if (frame) {
      frame.style.outline = outline;
      if (handles)
        frame.style.setProperty("--pdf-text-handle-visibility", handles);
      else frame.style.removeProperty("--pdf-text-handle-visibility");
      frame.removeAttribute("data-pdf-native-text-finishing");
    }
    if (input) {
      input.style.caretColor = caret;
      input.style.pointerEvents = pointer;
    }
  };
}
