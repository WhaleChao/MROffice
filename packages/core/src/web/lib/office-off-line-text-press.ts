/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A press OFF the lines of a text box (its inset above the first line, the
 * empty space under the last one) starts on the nearest line at the
 * pointer's x, and a drag from there selects, as in Word, Excel and
 * PowerPoint.
 *
 * The browser does otherwise on macOS: such a press sends the insertion
 * point to the very start or end of the text and holds it there for as long
 * as the pointer stays off the lines, so a drag along the box selects
 * nothing (a text field taller than its line does the same above and below
 * that line). This takes that one press over, for a text area, a text field
 * and a contentEditable host alike: the insertion point goes onto the nearest line
 * under the pointer, and the selection follows the pointer until the button
 * comes up. A press on a line of text is left to the browser.
 *
 * A text control's characters are found in the mini toolbar's mirror of its
 * text (office-text-selection-rect): the browser's own point-to-character
 * call misreads the lines of a text area drawn under a transform in WebKit.
 * A contentEditable host's are asked of the browser.
 *
 * ProseMirror documents (Docs' own drag selection, dragSelectionTargets),
 * CodeMirror and MathLive make their own selections and are left alone, as
 * is any host that opts out with `data-native-text-press`.
 */

import {
  officeTextControlMirror,
  type OfficeTextControlMirror,
} from "./office-text-selection-rect";

type CaretDocument = Document & {
  caretPositionFromPoint?: (
    x: number,
    y: number,
  ) => { offsetNode: Node; offset: number } | null;
  caretRangeFromPoint?: (x: number, y: number) => Range | null;
};

interface TextLines {
  /** Top of the first line and bottom of the last, in viewport pixels. */
  readonly top: number;
  readonly bottom: number;
  /** A y inside the first line, and one inside the last. */
  readonly firstLineY: number;
  readonly lastLineY: number;
}

interface CaretPoint {
  readonly node: Node;
  readonly offset: number;
}

type TextControl = HTMLTextAreaElement | HTMLInputElement;

/** The kinds of text field whose selection can be read and set. */
const TEXT_FIELD_TYPES: ReadonlySet<string> = new Set([
  "text",
  "search",
  "url",
  "tel",
  "password",
]);

function isTextControl(element: Element): element is TextControl {
  return (
    element instanceof HTMLTextAreaElement ||
    (element instanceof HTMLInputElement && TEXT_FIELD_TYPES.has(element.type))
  );
}

/** Hosts that make their own selections. */
const OWN_SELECTION_HOSTS =
  ".ProseMirror, .cm-editor, math-field, [data-native-text-press]";

/**
 * The text host a press landed in: a text area, a text field, or a
 * contentEditable root.
 */
export function officeTextPressHost(target: EventTarget | null): HTMLElement | null {
  if (!(target instanceof Element)) return null;
  if (isTextControl(target)) {
    return target.disabled || target.closest(OWN_SELECTION_HOSTS) ? null : target;
  }
  const editable = target.closest("[contenteditable]");
  if (!(editable instanceof HTMLElement) || !editable.isContentEditable) return null;
  let host = editable;
  while (host.parentElement?.isContentEditable) host = host.parentElement;
  // A host of its own inside a document (a shape's words) is not the
  // document's: only the host itself decides.
  if (host.matches(OWN_SELECTION_HOSTS) || host.closest("math-field, .cm-editor")) {
    return null;
  }
  if (host.closest("[data-native-text-press]")) return null;
  return host;
}

function linesOfTextArea(
  control: HTMLTextAreaElement,
  mirror: OfficeTextControlMirror,
): TextLines | null {
  const { value } = control;
  if (!value) return null;
  const rects = mirror.rangeRects(0, value.length);
  const first = rects[0];
  const last = rects[rects.length - 1];
  if (!first || !last) return null;
  // A closing line break opens one more (empty) line under the last drawn.
  const openLine = value.endsWith("\n") ? last.height : 0;
  const lastTop = last.top + openLine;
  return {
    top: first.top,
    bottom: lastTop + last.height,
    firstLineY: first.top + first.height / 2,
    lastLineY: lastTop + last.height / 2,
  };
}

/** A text field's one line: centred in its content box. */
function lineOfTextField(control: HTMLInputElement): TextLines | null {
  if (!control.value || !control.offsetHeight) return null;
  const view = control.ownerDocument.defaultView;
  if (!view) return null;
  const computed = view.getComputedStyle(control);
  const box = control.getBoundingClientRect();
  const scale = box.height / control.offsetHeight;
  const px = (value: string) => parseFloat(value) || 0;
  const top = box.top + (px(computed.borderTopWidth) + px(computed.paddingTop)) * scale;
  const bottom =
    box.bottom - (px(computed.borderBottomWidth) + px(computed.paddingBottom)) * scale;
  const fontSize = px(computed.fontSize);
  const lineHeight = (px(computed.lineHeight) || fontSize * 1.2) * scale;
  const middle = (top + bottom) / 2;
  return {
    top: middle - lineHeight / 2,
    bottom: middle + lineHeight / 2,
    firstLineY: middle,
    lastLineY: middle,
  };
}

function linesOfEditable(host: HTMLElement): TextLines | null {
  const doc = host.ownerDocument;
  // The lines are drawn by text and by line breaks: an empty line being
  // typed in holds a <br> alone (a deck's empty paragraph, a Sheets shape's
  // empty line). Read off the text alone, a press on an empty last line fell
  // "under the last line" and was sent to the end of the line above, so
  
  const walker = doc.createTreeWalker(
    host,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
    {
      acceptNode: (node) =>
        node.nodeType === Node.TEXT_NODE || node.nodeName === "BR"
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_SKIP,
    },
  );
  const range = doc.createRange();
  const lineRects = (node: Node): DOMRect[] => {
    if (node.nodeName === "BR") range.selectNode(node);
    else range.selectNodeContents(node);
    return typeof range.getClientRects === "function"
      ? [...range.getClientRects()].filter((rect) => rect.height > 0)
      : [];
  };
  let first: DOMRect | null = null;
  let last: DOMRect | null = null;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const rects = lineRects(node);
    if (rects.length === 0) continue;
    first ??= rects[0] ?? null;
    last = rects[rects.length - 1] ?? last;
  }
  if (!first || !last) return null;
  return {
    top: first.top,
    bottom: last.bottom,
    firstLineY: first.top + first.height / 2,
    lastLineY: last.top + last.height / 2,
  };
}

/** The caret the browser finds under a point of a contentEditable host. */
function editableCaretAt(host: HTMLElement, x: number, y: number): CaretPoint | null {
  const doc = host.ownerDocument as CaretDocument;
  let caret: CaretPoint | null = null;
  if (typeof doc.caretPositionFromPoint === "function") {
    const position = doc.caretPositionFromPoint(x, y);
    caret = position ? { node: position.offsetNode, offset: position.offset } : null;
  } else if (typeof doc.caretRangeFromPoint === "function") {
    const range = doc.caretRangeFromPoint(x, y);
    caret = range ? { node: range.startContainer, offset: range.startOffset } : null;
  }
  return caret && host.contains(caret.node) ? caret : null;
}

/** One text host's lines, and the caret under a point of it. */
interface PressedText {
  readonly lines: TextLines;
  caretAt(x: number, y: number): CaretPoint | null;
  release(): void;
}

function pressedText(host: HTMLElement): PressedText | null {
  if (isTextControl(host)) {
    const mirror = officeTextControlMirror(host);
    if (!mirror) return null;
    const lines =
      host instanceof HTMLTextAreaElement
        ? linesOfTextArea(host, mirror)
        : lineOfTextField(host);
    if (!lines) {
      mirror.remove();
      return null;
    }
    return {
      lines,
      caretAt: (x, y) => ({ node: host, offset: mirror.offsetAt(x, y) }),
      release: () => mirror.remove(),
    };
  }
  const lines = linesOfEditable(host);
  if (!lines) return null;
  return {
    lines,
    caretAt: (x, y) => editableCaretAt(host, x, y),
    release: () => {},
  };
}

/** The caret under the pointer, on the nearest line of `host`. */
function caretOnNearestLine(
  host: HTMLElement,
  text: PressedText,
  x: number,
  y: number,
): CaretPoint | null {
  const { lines } = text;
  const box = host.getBoundingClientRect();
  const lineY = y < lines.top ? lines.firstLineY : y > lines.bottom ? lines.lastLineY : y;
  return text.caretAt(
    Math.min(Math.max(x, box.left + 1), box.right - 1),
    Math.min(Math.max(lineY, box.top + 1), box.bottom - 1),
  );
}

function select(host: HTMLElement, anchor: CaretPoint, head: CaretPoint): void {
  if (isTextControl(host)) {
    const backward = head.offset < anchor.offset;
    host.setSelectionRange(
      Math.min(anchor.offset, head.offset),
      Math.max(anchor.offset, head.offset),
      backward ? "backward" : "forward",
    );
    return;
  }
  const selection = host.ownerDocument.defaultView?.getSelection();
  if (!selection) return;
  if (!host.contains(anchor.node) || !host.contains(head.node)) return;
  selection.setBaseAndExtent(anchor.node, anchor.offset, head.node, head.offset);
}

/** Whether the pointer is over the host's scrollbar rather than its text. */
function onScrollbar(host: HTMLElement, x: number, y: number): boolean {
  if (!host.offsetWidth || !host.offsetHeight) return false;
  const box = host.getBoundingClientRect();
  const scaleX = box.width / host.offsetWidth;
  const scaleY = box.height / host.offsetHeight;
  const right = box.left + (host.clientLeft + host.clientWidth) * scaleX;
  const bottom = box.top + (host.clientTop + host.clientHeight) * scaleY;
  return x > right || y > bottom;
}

/**
 * Take over a press off the lines of the text host under it. Returns whether
 * it did (the event's default is then prevented, and the selection follows
 * the pointer until the button comes up).
 */
export function officeOffLineTextPress(event: MouseEvent): boolean {
  if (
    event.button !== 0 ||
    event.detail > 1 ||
    event.shiftKey ||
    event.metaKey ||
    event.ctrlKey ||
    event.altKey ||
    event.defaultPrevented
  ) {
    return false;
  }
  const host = officeTextPressHost(event.target);
  if (!host) return false;
  const x = event.clientX;
  const y = event.clientY;
  if (onScrollbar(host, x, y)) return false;
  const text = pressedText(host);
  if (!text) return false;
  const onLines = y >= text.lines.top && y <= text.lines.bottom;
  const anchor = onLines ? null : caretOnNearestLine(host, text, x, y);
  if (!anchor) {
    text.release();
    return false;
  }

  event.preventDefault();
  if (host.ownerDocument.activeElement !== host) host.focus({ preventScroll: true });
  select(host, anchor, anchor);

  const view = host.ownerDocument.defaultView ?? window;
  const onMove = (move: MouseEvent) => {
    if (!(move.buttons & 1)) {
      finish();
      return;
    }
    const head = caretOnNearestLine(host, text, move.clientX, move.clientY);
    if (head) select(host, anchor, head);
    move.preventDefault();
  };
  const finish = () => {
    view.removeEventListener("mousemove", onMove, true);
    view.removeEventListener("mouseup", finish, true);
    view.removeEventListener("blur", finish);
    text.release();
  };
  view.addEventListener("mousemove", onMove, true);
  view.addEventListener("mouseup", finish, true);
  view.addEventListener("blur", finish);
  return true;
}

/**
 * Put the caret of a text host at the point nearest to a press that landed
 * outside the host but inside the object it is the words of -- a table cell
 * whose words fill only its first line (PowerPoint: a click anywhere in a
 * cell puts the caret in that cell, on its nearest line). Focuses the host;
 * an empty host takes the caret at its start. Returns whether it did.
 */
export function officePressIntoTextHost(host: HTMLElement, x: number, y: number): boolean {
  const editable = host.isContentEditable ||
    ["", "true", "plaintext-only"].includes(host.getAttribute("contenteditable") ?? "false");
  if (isTextControl(host) ? host.disabled || host.readOnly : !editable) return false;
  if (host.ownerDocument.activeElement !== host) host.focus({ preventScroll: true });
  const text = pressedText(host);
  let place: () => void;
  try {
    const caret = text ? caretOnNearestLine(host, text, x, y) : null;
    place = caret
      ? () => select(host, caret, caret)
      // No lines yet (an empty cell): the caret at its start.
      : () => {
          if (isTextControl(host)) host.setSelectionRange(0, 0);
          else host.ownerDocument.defaultView?.getSelection()?.collapse(host, 0);
        };
  } finally {
    text?.release();
  }
  place();
  // An editor that puts back the range it kept when it is focused does so
  // after the focus (a microtask): the press's caret is placed after it.
  queueMicrotask(() => {
    if (host.ownerDocument.activeElement === host) place();
  });
  return true;
}
