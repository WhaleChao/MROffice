/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a text selection sits on screen, in viewport pixels, for the Mini
 * Toolbar over it (office-ribbon/OfficeFloatingMiniToolbar).
 *
 * A contentEditable host has a DOM Range to measure. A text input or a text
 * area does not: its selection is two offsets into its value, so the text is
 * laid out once more in a hidden mirror with the control's own font, padding
 * and wrapping, and the selected part of the mirror is measured instead.
 */

export interface OfficeTextSelectionRect {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

/** The style properties that decide where a control's characters fall. */
const MIRRORED_PROPERTIES = [
  "boxSizing",
  "width",
  "borderTopWidth",
  "borderRightWidth",
  "borderBottomWidth",
  "borderLeftWidth",
  "paddingTop",
  "paddingRight",
  "paddingBottom",
  "paddingLeft",
  "fontStyle",
  "fontVariant",
  "fontWeight",
  "fontStretch",
  "fontSize",
  "fontFamily",
  "lineHeight",
  "letterSpacing",
  "wordSpacing",
  "textIndent",
  "textTransform",
  "textAlign",
  "tabSize",
  "direction",
] as const;

function unionOfRects(rects: readonly DOMRect[]): OfficeTextSelectionRect | null {
  const visible = rects.filter((rect) => rect.width > 0 || rect.height > 0);
  if (visible.length === 0) return null;
  const left = Math.min(...visible.map((rect) => rect.left));
  const top = Math.min(...visible.map((rect) => rect.top));
  const right = Math.max(...visible.map((rect) => rect.right));
  const bottom = Math.max(...visible.map((rect) => rect.bottom));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * A text control's text laid out once more where its characters can be
 * measured: the control's own font, padding and wrapping in a hidden box.
 * Kept in the page until `remove()`.
 */
export interface OfficeTextControlMirror {
  /**
   * The line boxes of the characters `start`..`end`, in viewport pixels and
   * not clipped to the control.
   */
  rangeRects(start: number, end: number): OfficeTextSelectionRect[];
  /**
   * The offset of the character boundary nearest a viewport point: on the
   * line the point's y is on (the first or the last when it is above or
   * below them), at the point's x.
   */
  offsetAt(x: number, y: number): number;
  remove(): void;
}

export function officeTextControlMirror(
  control: HTMLInputElement | HTMLTextAreaElement,
): OfficeTextControlMirror | null {
  const doc = control.ownerDocument;
  const view = doc.defaultView;
  if (!view) return null;
  const computed = view.getComputedStyle(control);
  const box = control.getBoundingClientRect();
  const mirror = doc.createElement("div");
  const style = mirror.style;
  for (const property of MIRRORED_PROPERTIES) {
    style[property] = computed[property];
  }
  const multiline = control instanceof HTMLTextAreaElement;
  // The mirror is laid out in the control's own CSS pixels. A control drawn
  // under a transform (a slide scaled to fit its window) shows those pixels
  // scaled, which its box and its layout size tell apart.
  const scaleX = control.offsetWidth > 0 ? box.width / control.offsetWidth : 1;
  const scaleY = control.offsetHeight > 0 ? box.height / control.offsetHeight : 1;
  style.position = "fixed";
  style.left = `${box.left}px`;
  style.top = `${box.top}px`;
  style.visibility = "hidden";
  style.pointerEvents = "none";
  // All of the text, unscrolled: the control's scroll is taken off below.
  style.overflow = "visible";
  style.whiteSpace = multiline ? "pre-wrap" : "pre";
  style.overflowWrap = multiline ? "break-word" : "normal";
  style.height = "auto";
  const value = control.value;
  const text = doc.createTextNode(value);
  // A closing line break opens a line of its own only with something on it.
  mirror.append(text, doc.createTextNode(" "));
  doc.body.append(mirror);
  const origin = mirror.getBoundingClientRect();
  const range = doc.createRange();
  /** The drawn boxes of `start`..`end`, in the mirror's own pixels. */
  const localRects = (start: number, end: number) => {
    range.setStart(text, Math.max(0, Math.min(start, value.length)));
    range.setEnd(text, Math.max(0, Math.min(end, value.length)));
    const rects = typeof range.getClientRects === "function" ? [...range.getClientRects()] : [];
    return rects
      .filter((rect) => rect.width > 0 || rect.height > 0)
      .map((rect) => ({
        left: rect.left - origin.left,
        top: rect.top - origin.top,
        width: rect.width,
        height: rect.height,
      }));
  };
  /** The box of the one character at `index` (a line break's is empty, at its line's end). */
  const charBox = (index: number) => localRects(index, index + 1)[0] ?? null;
  /** The smallest index in `from`..`to` that `passes`, or `to`. */
  const firstIndex = (from: number, to: number, passes: (index: number) => boolean) => {
    let low = from;
    let high = to;
    while (low < high) {
      const middle = (low + high) >> 1;
      if (passes(middle)) high = middle;
      else low = middle + 1;
    }
    return low;
  };
  return {
    rangeRects(start, end) {
      if (start >= end) return [];
      return localRects(start, end).map((line) => ({
        left: box.left + (line.left - control.scrollLeft) * scaleX,
        top: box.top + (line.top - control.scrollTop) * scaleY,
        width: line.width * scaleX,
        height: line.height * scaleY,
      }));
    },
    offsetAt(x, y) {
      const count = value.length;
      if (count === 0) return 0;
      const localX = (x - box.left) / scaleX + control.scrollLeft;
      const localY = (y - box.top) / scaleY + control.scrollTop;
      let lineStart = 0;
      let lineEnd = count;
      if (multiline) {
        const lastBox = charBox(count - 1);
        // Under everything, after a closing line break: its own empty line.
        if (value.endsWith("\n") && lastBox && localY > lastBox.top + lastBox.height) {
          return count;
        }
        // A character of the pointer's line: the first one not above it.
        const onLine = Math.min(
          count - 1,
          firstIndex(0, count, (index) => {
            const at = charBox(index);
            return at !== null && at.top + at.height > localY;
          }),
        );
        const lineTop = charBox(onLine)?.top ?? 0;
        lineStart = firstIndex(0, onLine, (index) => {
          const at = charBox(index);
          return at !== null && at.top >= lineTop - 1;
        });
        lineEnd = firstIndex(onLine, count, (index) => {
          const at = charBox(index);
          return at !== null && at.top > lineTop + 1;
        });
      }
      const offset = firstIndex(lineStart, lineEnd, (index) => {
        const at = charBox(index);
        return at !== null && at.left + at.width / 2 > localX;
      });
      // Past the line's words: before the break that ends it.
      return offset === lineEnd && lineEnd > lineStart && value[lineEnd - 1] === "\n"
        ? lineEnd - 1
        : offset;
    },
    remove() {
      mirror.remove();
    },
  };
}

/**
 * The line boxes of the characters `start`..`end` of a text input or text
 * area, in viewport pixels and not clipped to the control. Empty when the
 * range holds nothing drawn (or the control is not laid out).
 */
export function officeTextControlRangeRects(
  control: HTMLInputElement | HTMLTextAreaElement,
  start: number,
  end: number,
): OfficeTextSelectionRect[] {
  if (start >= end) return [];
  const mirror = officeTextControlMirror(control);
  if (!mirror) return [];
  const lines = mirror.rangeRects(start, end);
  mirror.remove();
  return lines;
}

/**
 * The selected characters of a text input or text area, or null when the
 * selection is collapsed (or the control is not laid out).
 */
export function officeTextControlSelectionRect(
  control: HTMLInputElement | HTMLTextAreaElement,
): OfficeTextSelectionRect | null {
  const start = control.selectionStart;
  const end = control.selectionEnd;
  if (start === null || end === null || start === end) return null;
  const lines = officeTextControlRangeRects(control, start, end);
  if (lines.length === 0) return null;
  const box = control.getBoundingClientRect();
  const rectLeft = Math.min(...lines.map((line) => line.left));
  const rectTop = Math.min(...lines.map((line) => line.top));
  const rectRight = Math.max(...lines.map((line) => line.left + line.width));
  const rectBottom = Math.max(...lines.map((line) => line.top + line.height));
  const left = Math.max(box.left, rectLeft);
  const right = Math.min(box.right, rectRight);
  const top = Math.max(box.top, rectTop);
  const bottom = Math.min(box.bottom, rectBottom);
  return {
    left,
    top,
    width: Math.max(0, right - left),
    height: Math.max(0, bottom - top),
  };
}

/**
 * The document selection inside a contentEditable host, or null when it is
 * collapsed or lies elsewhere.
 */
export function officeEditableSelectionRect(host: HTMLElement): OfficeTextSelectionRect | null {
  const selection = host.ownerDocument.defaultView?.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return null;
  const range = selection.getRangeAt(0);
  if (!host.contains(range.commonAncestorContainer)) return null;
  // A document without layout (a server or test DOM) may not measure Ranges.
  const lines = typeof range.getClientRects === "function" ? [...range.getClientRects()] : [];
  return (
    unionOfRects(lines) ??
    (typeof range.getBoundingClientRect === "function"
      ? unionOfRects([range.getBoundingClientRect()])
      : null)
  );
}

/** The selection of either kind of text host. */
export function officeTextSelectionRect(host: HTMLElement): OfficeTextSelectionRect | null {
  if (host instanceof HTMLInputElement || host instanceof HTMLTextAreaElement) {
    return officeTextControlSelectionRect(host);
  }
  return officeEditableSelectionRect(host);
}
