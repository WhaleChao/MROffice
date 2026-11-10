/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What the keyboard keeps in view on a zoomed-in slide canvas (PowerPoint
 * scrolls its slide pane to it): the caret moving through words being typed,
 * else the object(s) an arrow key nudges or Tab selects.
 */

/** Room kept between what is revealed and the canvas's edge (CSS px). */
export const PRESENTATION_REVEAL_MARGIN_PX = 24;

export interface PresentationRevealBox {
  /** From the canvas's visible content corner, in CSS px. */
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

/**
 * How far to scroll (CSS px, signed) so `box` is in a `viewWidth` x
 * `viewHeight` view with PRESENTATION_REVEAL_MARGIN_PX to spare -- the least
 * scroll that does it. A box larger than the view is brought in by its
 * start (its top-left) when it is wholly out of view; partly in view, it
 * is left where it is.
 */
export function presentationRevealScroll(
  box: PresentationRevealBox,
  viewWidth: number,
  viewHeight: number,
  margin = PRESENTATION_REVEAL_MARGIN_PX,
): { dx: number; dy: number } {
  const axis = (start: number, size: number, view: number): number => {
    if (!Number.isFinite(start) || !Number.isFinite(size) || !(view > 0)) return 0;
    const room = Math.min(margin, Math.max(0, (view - size) / 2));
    const end = start + size;
    if (size + room * 2 > view) {
      // larger than the view: only when it is out of view, its start comes in
      return end <= 0 || start >= view ? start - room : 0;
    }
    if (start < room) return start - room;
    if (end > view - room) return end - (view - room);
    return 0;
  };
  return {
    dx: Math.round(axis(box.left, box.width, viewWidth)),
    dy: Math.round(axis(box.top, box.height, viewHeight)),
  };
}

function caretRect(slideFrame: HTMLElement): DOMRect | null {
  const active = slideFrame.ownerDocument.activeElement;
  if (!(active instanceof HTMLElement) || !slideFrame.contains(active)) return null;
  if (active.isContentEditable) {
    const selection = slideFrame.ownerDocument.getSelection();
    if (selection && selection.rangeCount > 0 && selection.focusNode && slideFrame.contains(selection.focusNode)) {
      const range = slideFrame.ownerDocument.createRange();
      try {
        range.setStart(selection.focusNode, selection.focusOffset);
        range.collapse(true);
      } catch {
        return active.getBoundingClientRect();
      }
      const rects = range.getClientRects();
      const rect = rects.length > 0 ? rects[rects.length - 1] : range.getBoundingClientRect();
      if (rect && (rect.height > 0 || rect.width > 0)) return rect;
      const host = selection.focusNode instanceof Element ? selection.focusNode : selection.focusNode.parentElement;
      return host?.getBoundingClientRect() ?? null;
    }
    return active.getBoundingClientRect();
  }
  if (active instanceof HTMLTextAreaElement) return textareaCaretRect(active);
  if (active instanceof HTMLInputElement) return active.getBoundingClientRect();
  return null;
}

/** The text-box properties a field's caret line depends on. */
const MIRRORED_STYLE = [
  "boxSizing", "width", "paddingTop", "paddingRight", "paddingBottom", "paddingLeft",
  "borderTopWidth", "borderRightWidth", "borderBottomWidth", "borderLeftWidth",
  "fontFamily", "fontSize", "fontWeight", "fontStyle", "fontVariant", "fontStretch",
  "letterSpacing", "wordSpacing", "lineHeight", "textAlign", "textIndent", "textTransform",
  "whiteSpace", "wordBreak", "overflowWrap", "tabSize", "direction", "writingMode",
] as const;

/**
 * The caret's line in a text field, on screen: a field draws no caret box, so
 * a twin of it (same box, same type, laid out where it stands, inside the
 * same scaled slide) is given the words up to the caret and a mark there --
 * the mark is where the caret's line is, even in a field taller than the
 * window (the field itself would only say "the whole box").
 */
export function textareaCaretRect(field: HTMLTextAreaElement): DOMRect {
  const parent = field.offsetParent instanceof HTMLElement ? field.offsetParent : field.parentElement;
  const view = field.ownerDocument.defaultView;
  if (!parent || !view) return field.getBoundingClientRect();
  const computed = view.getComputedStyle(field);
  const twin = field.ownerDocument.createElement("div");
  for (const key of MIRRORED_STYLE) {
    (twin.style as unknown as Record<string, string>)[key] = computed[key as keyof CSSStyleDeclaration] as string;
  }
  twin.style.position = "absolute";
  twin.style.visibility = "hidden";
  twin.style.pointerEvents = "none";
  twin.style.overflow = "hidden";
  twin.style.left = `${field.offsetLeft}px`;
  twin.style.top = `${field.offsetTop - field.scrollTop}px`;
  twin.style.height = "auto";
  if (computed.whiteSpace === "normal" || computed.whiteSpace === "") twin.style.whiteSpace = "pre-wrap";
  const caret = field.selectionDirection === "backward" ? field.selectionStart : field.selectionEnd;
  twin.textContent = field.value.slice(0, caret ?? 0);
  const mark = field.ownerDocument.createElement("span");
  // a zero-width character carries the line's height
  mark.textContent = "\u200b";
  twin.append(mark);
  parent.append(twin);
  try {
    const rect = mark.getBoundingClientRect();
    return rect.height > 0 ? rect : field.getBoundingClientRect();
  } finally {
    twin.remove();
  }
}

/** The screen box the keyboard last moved on the slide, or null. */
export function presentationKeyboardRevealRect(slideFrame: HTMLElement): DOMRect | null {
  const caret = caretRect(slideFrame);
  if (caret) return caret;
  const selected = slideFrame.querySelectorAll<HTMLElement>(
    '[role="group"][data-block-id][data-selected="true"], [role="group"][data-block-id][data-multiselected="true"]',
  );
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const element of Array.from(selected)) {
    const rect = element.getBoundingClientRect();
    if (!(rect.width > 0 || rect.height > 0)) continue;
    left = Math.min(left, rect.left);
    top = Math.min(top, rect.top);
    right = Math.max(right, rect.right);
    bottom = Math.max(bottom, rect.bottom);
  }
  if (!Number.isFinite(left)) return null;
  return new DOMRect(left, top, right - left, bottom - top);
}
