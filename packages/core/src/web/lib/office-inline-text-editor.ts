/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office edits a shape's text IN the shape: the words are typed where they
 * are drawn, in the shape's own font and colour, with no box of their own
 * (Word, Excel and PowerPoint alike). This is the ONE style the three
 * editors' in-shape text editors wear: the deck's shape textarea, the
 * worksheet's shape and SmartArt member textareas, and Docs' group / SmartArt
 * member textarea. Each host lays the label's own metrics (size, weight,
 * colour, WordArt paint, alignment, inset) over it.
 */
import type { CSSProperties } from "react";

export const OFFICE_INLINE_TEXT_EDITOR_STYLE: CSSProperties = {
  position: "absolute",
  inset: 0,
  width: "100%",
  height: "100%",
  // Every box metric a host's own form styles could set is stated here, so
  // none reaches the editor: the deck's `.presentation-window__slide
  // textarea` rule (a 118 px minimum height, a tinted rounded box) made a
  // shape's centred editor taller than a small shape and its words sat at
  
  minWidth: 0,
  minHeight: 0,
  maxWidth: "none",
  maxHeight: "none",
  boxSizing: "border-box",
  margin: 0,
  padding: 0,
  border: "none",
  borderRadius: 0,
  outline: "none",
  boxShadow: "none",
  background: "transparent",
  appearance: "none",
  resize: "none",
  overflow: "hidden",
  font: "inherit",
  color: "inherit",
  textAlign: "center",
  textIndent: 0,
  textTransform: "none",
  lineHeight: "inherit",
  whiteSpace: "pre-wrap",
  overflowWrap: "anywhere",
};

/** The shared editor style with the host's label metrics laid over it. */
export function officeInlineTextEditorStyle(label: CSSProperties = {}): CSSProperties {
  return { ...OFFICE_INLINE_TEXT_EDITOR_STYLE, ...label };
}

/**
 * A label anchored at the shape's middle is typed at the middle too. A
 * textarea cannot centre its lines, so the editor is as tall as its words
 * (`fitOfficeInlineTextEditor` keeps that true while typing) and hangs from
 * the vertical centre of the label's inset. `insetTop` / `insetBottom` are
 * the label's own top and bottom insets (fractions of the shape).
 */
export function officeInlineTextEditorCenteredStyle(
  insetTop: number,
  insetBottom: number,
): CSSProperties {
  const middle = insetTop + (1 - insetTop - insetBottom) / 2;
  return {
    top: `${middle * 100}%`,
    bottom: "auto",
    height: "auto",
    maxHeight: `${(1 - insetTop - insetBottom) * 100}%`,
    transform: "translateY(-50%)",
  };
}

/** Size a centred editor to its words (call on open and on every input). */
export function fitOfficeInlineTextEditor(element: HTMLTextAreaElement): void {
  element.style.height = "auto";
  element.style.height = `${element.scrollHeight}px`;
}

/** The centred layout on a DOM element, as `applyOfficeInlineTextEditorStyle`. */
export function applyOfficeInlineTextEditorCentered(
  element: HTMLElement,
  insetTop: number,
  insetBottom: number,
): void {
  const centered = officeInlineTextEditorCenteredStyle(insetTop, insetBottom);
  element.style.top = String(centered.top);
  element.style.bottom = String(centered.bottom);
  element.style.height = String(centered.height);
  element.style.maxHeight = String(centered.maxHeight);
  element.style.transform = String(centered.transform);
}

/** The same style on a DOM element (Docs' node views build the DOM by hand). */
export function applyOfficeInlineTextEditorStyle(
  element: HTMLElement,
  label: Record<string, string> = {},
): void {
  // The very style the React hosts wear, so the two can never drift.
  Object.assign(
    element.style,
    Object.fromEntries(
      Object.entries(OFFICE_INLINE_TEXT_EDITOR_STYLE).map(([key, value]) => [key, String(value)]),
    ),
  );
  for (const [key, value] of Object.entries(label)) element.style.setProperty(key, value);
}

/*
 * No text box is drawn around an in-shape edit: Office marks text-edit mode
 * by drawing the object's own selection outline dashed (the shared
 * `data-text-editing` rule, lib/office-text-edit-mode.ts).
 */
type CaretDocument = Document & {
  caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
  caretRangeFromPoint?: (x: number, y: number) => Range | null;
};

/**
 * Where a double-click lands in a label, as an offset into the label's text:
 * the caret then opens at the pointer, as clicking into a text frame does.
 * The label is hit-tested as drawn (its pointer events are lifted for the
 * query), marker glyphs (`aria-hidden`) are not text, and each new line box
 * counts one line break. Null when the point is not on the label's words;
 * the host then opens the caret at the end.
 */
export function officeInlineTextCaretFromPoint(
  label: HTMLElement | null,
  x: number,
  y: number,
): number | null {
  if (!label || !Number.isFinite(x) || !Number.isFinite(y)) return null;
  const doc = label.ownerDocument as CaretDocument;
  const previous = label.style.pointerEvents;
  label.style.pointerEvents = "auto";
  let node: Node | null = null;
  let offset = 0;
  try {
    if (typeof doc.caretPositionFromPoint === "function") {
      const position = doc.caretPositionFromPoint(x, y);
      if (position) {
        node = position.offsetNode;
        offset = position.offset;
      }
    } else if (typeof doc.caretRangeFromPoint === "function") {
      const range = doc.caretRangeFromPoint(x, y);
      if (range) {
        node = range.startContainer;
        offset = range.startOffset;
      }
    }
  } finally {
    label.style.pointerEvents = previous;
  }
  if (!node || node.nodeType !== 3 || !label.contains(node)) return null;
  const view = doc.defaultView;
  const lineOf = (text: Node): Element => {
    let element = text.parentElement;
    while (element && element !== label) {
      const display = view?.getComputedStyle(element).display ?? "";
      if (display === "block" || display === "flex" || display === "grid" || display === "list-item") {
        return element;
      }
      element = element.parentElement;
    }
    return label;
  };
  const walker = doc.createTreeWalker(label, 4 /* NodeFilter.SHOW_TEXT */);
  let total = 0;
  let lastLine: Element | null = null;
  for (let text = walker.nextNode(); text; text = walker.nextNode()) {
    if (text.parentElement?.closest('[aria-hidden="true"]')) continue;
    const line = lineOf(text);
    if (lastLine !== null && line !== lastLine) total += 1;
    lastLine = line;
    const length = text.textContent?.length ?? 0;
    if (text === node) return total + Math.max(0, Math.min(offset, length));
    total += length;
  }
  return null;
}

/**
 * Open a caret, not a selection: at the pointer when the host found one, at
 * the end of the words otherwise (the deck's editor opens the same way).
 */
export function placeOfficeInlineTextCaret(
  element: HTMLTextAreaElement,
  caret?: number | null,
): void {
  const length = element.value.length;
  const at = typeof caret === "number" ? Math.max(0, Math.min(caret, length)) : length;
  element.setSelectionRange(at, at);
}

/**
 * Tab inside an in-shape editor types a tab character, over the selection if
 * there is one, as Excel and PowerPoint do in a shape's or a text box's
 * words. Done here and consumed: the browser's own Tab walks to the next
 * control, and the blur that follows ends the edit, so the key that should
 * have spaced the words closed them instead. Shift+Tab stays in the words and
 * does nothing; Ctrl / Cmd / Alt + Tab are not this key. Returns false for
 * any other key, which the caller then handles.
 */
export function typeOfficeInlineTextTab(
  event: {
    key: string;
    shiftKey: boolean;
    ctrlKey: boolean;
    metaKey: boolean;
    altKey: boolean;
    preventDefault: () => void;
  },
  field: HTMLTextAreaElement,
): boolean {
  if (event.key !== "Tab" || event.ctrlKey || event.metaKey || event.altKey) return false;
  event.preventDefault();
  if (event.shiftKey || field.readOnly || field.disabled) return true;
  const start = field.selectionStart ?? field.value.length;
  const end = field.selectionEnd ?? start;
  if (field.maxLength >= 0 && field.value.length - (end - start) + 1 > field.maxLength) return true;
  field.setRangeText("\t", start, end, "end");
  // The host hears it as typing: a controlled field takes the new words, an
  // uncontrolled one fits its height to them.
  field.dispatchEvent(new Event("input", { bubbles: true }));
  return true;
}












export function moveOfficeInlineTextCaretHomeEnd(
  event: {
    key: string;
    shiftKey: boolean;
    ctrlKey: boolean;
    metaKey: boolean;
    altKey: boolean;
    preventDefault: () => void;
  },
  field: HTMLTextAreaElement | HTMLInputElement,
): boolean {
  if ((event.key !== "Home" && event.key !== "End") || event.altKey) return false;
  const value = field.value;
  const toEnd = event.key === "End";
  const whole = event.ctrlKey || event.metaKey;
  const start = field.selectionStart ?? value.length;
  const end = field.selectionEnd ?? start;
  // The caret the move starts from: the selection's focus, which the field
  // keeps as its direction.
  const focus = field.selectionDirection === "backward" ? start : end;
  let to: number;
  if (whole) {
    to = toEnd ? value.length : 0;
  } else if (toEnd) {
    const next = value.indexOf("\n", focus);
    to = next < 0 ? value.length : next;
  } else {
    to = value.lastIndexOf("\n", focus - 1) + 1;
  }
  event.preventDefault();
  try {
    if (event.shiftKey) {
      const anchor = field.selectionDirection === "backward" ? end : start;
      field.setSelectionRange(Math.min(anchor, to), Math.max(anchor, to), to < anchor ? "backward" : "forward");
    } else {
      field.setSelectionRange(to, to);
    }
  } catch {
    // A field without a selection API keeps its caret.
  }
  return true;
}
