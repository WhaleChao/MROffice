"use client";


























/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Less than this between the visual and layout viewports' bottoms is a
 * hardware keyboard's shortcut bar or a browser toolbar, not a keyboard. */
export const SOFT_KEYBOARD_MIN_INSET_PX = 120;
/** Room kept between the caret and the keyboard (the hide-keyboard bar). */
export const SOFT_KEYBOARD_CARET_MARGIN_PX = 56;

export interface CaretBox {
  top: number;
  bottom: number;
}

export interface VisibleBand {
  /** Client-pixel top and bottom of the visual viewport. */
  top: number;
  bottom: number;
}

const TEXT_INPUT_TYPES = new Set([
  "",
  "text",
  "search",
  "email",
  "url",
  "tel",
  "password",
  "number",
  "date",
  "datetime-local",
  "month",
  "time",
  "week",
]);

/** A field the on-screen keyboard types into. */
export function isKeyboardEditable(element: Element | null): element is HTMLElement {
  if (!element || !(element instanceof HTMLElement)) return false;
  if (element instanceof HTMLTextAreaElement) return !element.readOnly && !element.disabled;
  if (element instanceof HTMLInputElement) {
    return (
      TEXT_INPUT_TYPES.has(element.type) && !element.readOnly && !element.disabled
    );
  }
  return element.isContentEditable;
}

/**
 * How much of the layout viewport the keyboard covers (CSS px), or 0 while it
 * is down, while the page is pinch-zoomed (the visual viewport is then
 * smaller by choice) or where there is no visual viewport.
 */
export function softKeyboardInset(win: Window = window): number {
  const viewport = win.visualViewport;
  if (!viewport || viewport.scale > 1.01) return 0;
  const layoutHeight = win.document.documentElement.clientHeight || win.innerHeight;
  const hidden = layoutHeight - (viewport.offsetTop + viewport.height);
  return hidden >= SOFT_KEYBOARD_MIN_INSET_PX ? Math.round(hidden) : 0;
}

/**
 * How many pixels at the bottom of `element`'s box the on-screen keyboard
 * hides, plus `margin` for the bar on its top edge (0 while it is down). A
 * view that scrolls its own content (the spreadsheet grid) treats this as a
 * cover on its bottom edge, as it does a task pane floating over it, so the
 * cell being typed in stays above the keyboard.
 */
export function softKeyboardCover(
  element: Element,
  margin: number = SOFT_KEYBOARD_CARET_MARGIN_PX,
  win: Window = window,
): number {
  if (softKeyboardInset(win) === 0) return 0;
  const box = element.getBoundingClientRect();
  return Math.max(0, Math.round(box.bottom - (visibleBand(win).bottom - margin)));
}

/**
 * Does this keydown send a chat composer's message? "Enter sends, Shift+Enter
 * starts a new line" left an on-screen keyboard (an iPad's, Android's) no way
 * to write a second line: it has no Shift+Return, and its Return is Enter.
 * While one is up, Return is the line break and the composer's Send button
 * sends, as chat apps do on a tablet; a hardware keyboard keeps Enter to send.
 */
export function enterSendsMessage(
  event: {
    key: string;
    shiftKey: boolean;
    keyCode?: number;
    isComposing?: boolean;
    nativeEvent?: { isComposing?: boolean; keyCode?: number };
  },
  win: Window = window,
): boolean {
  return (
    event.key === "Enter" &&
    !event.shiftKey &&
    !isImeKeydown(event) &&
    softKeyboardInset(win) === 0
  );
}







export function isImeKeydown(event: {
  keyCode?: number;
  isComposing?: boolean;
  nativeEvent?: { isComposing?: boolean; keyCode?: number };
}): boolean {
  return (
    event.isComposing === true ||
    event.nativeEvent?.isComposing === true ||
    event.keyCode === 229 ||
    event.nativeEvent?.keyCode === 229
  );
}

/** The visual viewport in client pixels (what is not under the keyboard). */
export function visibleBand(win: Window = window): VisibleBand {
  const viewport = win.visualViewport;
  if (!viewport) return { top: 0, bottom: win.innerHeight };
  return { top: viewport.offsetTop, bottom: viewport.offsetTop + viewport.height };
}

/** The insertion point's box (client px) in the focused field, or null. */
export function caretBox(doc: Document = document): CaretBox | null {
  const active = doc.activeElement;
  if (!isKeyboardEditable(active)) return null;
  if (active instanceof HTMLInputElement || active instanceof HTMLTextAreaElement) {
    const rect = active.getBoundingClientRect();
    // A tall text area: its caret line is unknown; keep its first lines up.
    const bottom = Math.min(rect.bottom, rect.top + 48);
    return { top: rect.top, bottom };
  }
  const selection = doc.getSelection();
  if (!selection || selection.rangeCount === 0 || !selection.focusNode) return null;
  if (!active.contains(selection.focusNode)) return null;
  const range = doc.createRange();
  try {
    range.setStart(selection.focusNode, selection.focusOffset);
  } catch {
    return null;
  }
  range.collapse(true);
  const rects = range.getClientRects();
  const rect = rects.length > 0 ? rects[0] : range.getBoundingClientRect();
  if (rect && rect.height > 0) {
    return { top: rect.top, bottom: rect.bottom };
  }
  // An empty line has no caret box: the block it sits in has one.
  const node = selection.focusNode;
  const element =
    node.nodeType === Node.ELEMENT_NODE
      ? ((node as Element).childNodes[selection.focusOffset] as Element | undefined) ??
        (node as Element)
      : node.parentElement;
  const box = element instanceof Element ? element.getBoundingClientRect() : null;
  return box ? { top: box.top, bottom: Math.min(box.bottom, box.top + 32) } : null;
}

function scrollsVertically(element: Element): boolean {
  const style = getComputedStyle(element);
  return (
    (style.overflowY === "auto" || style.overflowY === "scroll") &&
    element.scrollHeight > element.clientHeight + 1
  );
}

/**
 * Scroll the caret's scrolling ancestors, inner first (the editor's own page
 * column before anything around its window) and never <body>, so the caret
 * sits inside `band`, `margin` clear of its bottom. An ancestor that cannot
 * scroll (a maximised window's scroll-locked canvas wears overflow hidden)
 * is passed over. Returns the pixels scrolled.
 */
export function revealCaretInBand(
  caret: CaretBox,
  band: VisibleBand,
  from: Element,
  margin: number = SOFT_KEYBOARD_CARET_MARGIN_PX,
): number {
  let { top, bottom } = caret;
  let moved = 0;
  for (
    let element: Element | null = from;
    element && element !== element.ownerDocument.body;
    element = element.parentElement
  ) {
    if (!scrollsVertically(element)) continue;
    const rect = element.getBoundingClientRect();
    const limitBottom = Math.min(band.bottom - margin, rect.bottom - 4);
    const limitTop = Math.max(band.top + 8, rect.top + 4);
    let delta = 0;
    if (bottom > limitBottom) delta = bottom - limitBottom;
    else if (top < limitTop) delta = top - limitTop;
    if (delta !== 0) {
      const before = element.scrollTop;
      const max = element.scrollHeight - element.clientHeight;
      element.scrollTop = Math.max(0, Math.min(max, before + delta));
      const step = element.scrollTop - before;
      top -= step;
      bottom -= step;
      moved += step;
    }
    if (bottom <= band.bottom - margin && top >= band.top) break;
  }
  return moved;
}

export interface SoftKeyboardState {
  /** The keyboard covers this much of the page (0: down). */
  inset: number;
  /** The visual viewport's bottom, in client px (where the keyboard starts). */
  keyboardTop: number;
}

/**
 * Follow the keyboard for the fields inside `root`: report it up/down and
 * keep the caret above it as it opens, as it resizes and as typing moves
 * the caret. Returns the uninstaller.
 */
export function watchSoftKeyboard(
  root: HTMLElement,
  onChange: (state: SoftKeyboardState) => void,
): () => void {
  const win = root.ownerDocument.defaultView;
  const viewport = win?.visualViewport;
  if (!win || !viewport) return () => {};
  const doc = root.ownerDocument;
  let frame = 0;
  let last: SoftKeyboardState = { inset: 0, keyboardTop: 0 };
  let lastInset = 0;

  const focusedHere = () => {
    const active = doc.activeElement;
    return isKeyboardEditable(active) && root.contains(active) ? active : null;
  };

  const update = (reveal: boolean) => {
    frame = 0;
    const field = focusedHere();
    const inset = field ? softKeyboardInset(win) : 0;
    const band = visibleBand(win);
    const next = { inset, keyboardTop: Math.round(band.bottom) };
    if (next.inset !== last.inset || (inset > 0 && next.keyboardTop !== last.keyboardTop)) {
      last = next;
      // Reported to the caller only. It used to be a custom property on
      // <html> as well, which nothing read and which restyled the whole page
      // on every step of the keyboard rising (and on each viewport move
      // while it was up).
      onChange(next);
    }
    // the keyboard rose (or grew), or the caret moved while it is up
    const rose = inset > lastInset;
    lastInset = inset;
    if (!field || inset === 0 || !(reveal || rose)) return;
    const caret = caretBox(doc);
    if (!caret) return;
    const selection = doc.getSelection();
    const from =
      field instanceof HTMLInputElement || field instanceof HTMLTextAreaElement
        ? field
        : (selection?.focusNode?.parentElement ?? field);
    revealCaretInBand(caret, band, from);
  };
  const schedule = (reveal: boolean) => () => {
    if (frame) win.cancelAnimationFrame(frame);
    frame = win.requestAnimationFrame(() => update(reveal));
  };
  const onViewport = schedule(false);
  const onCaret = schedule(true);
  const onSelection = () => {
    if (lastInset > 0 && focusedHere()) onCaret();
  };

  viewport.addEventListener("resize", onViewport);
  viewport.addEventListener("scroll", onViewport);
  doc.addEventListener("selectionchange", onSelection);
  root.addEventListener("focusin", onViewport);
  root.addEventListener("focusout", onViewport);
  return () => {
    if (frame) win.cancelAnimationFrame(frame);
    viewport.removeEventListener("resize", onViewport);
    viewport.removeEventListener("scroll", onViewport);
    doc.removeEventListener("selectionchange", onSelection);
    root.removeEventListener("focusin", onViewport);
    root.removeEventListener("focusout", onViewport);
  };
}

/** Put the keyboard away: the field loses the focus (its selection stays). */
export function hideSoftKeyboard(doc: Document = document): void {
  const active = doc.activeElement;
  if (active instanceof HTMLElement && isKeyboardEditable(active)) active.blur();
}













export function holdIphoneFocusZoom(doc: Document = document): () => void {
  const nav = doc.defaultView?.navigator;
  if (!nav || !/\b(iPhone|iPod)\b/.test(nav.userAgent)) return () => {};
  // Counted: the workspace's layout can hold it twice for a moment (its
  // notebook and page branches swap), and the first to let go must not take
  // it from the other. Watched: the page's head is the router's, and a page's
  // address (…/pages/<id>, which settles on the notebook's) brought the
  // viewport back without the limit.
  if (focusZoomHolds === 0) {
    const apply = () => {
      const meta = doc.querySelector<HTMLMetaElement>('meta[name="viewport"]');
      if (!meta) return;
      const content = meta.getAttribute("content") ?? "";
      if (/maximum-scale\s*=/i.test(content)) return;
      focusZoomHeld.set(meta, content);
      meta.setAttribute(
        "content",
        content ? `${content}, maximum-scale=1` : "width=device-width, initial-scale=1, maximum-scale=1",
      );
    };
    apply();
    const View = doc.defaultView as (Window & typeof globalThis) | null;
    if (View?.MutationObserver && doc.head) {
      focusZoomWatch = new View.MutationObserver(apply);
      focusZoomWatch.observe(doc.head, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["content"],
      });
    }
  }
  focusZoomHolds += 1;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    focusZoomHolds = Math.max(0, focusZoomHolds - 1);
    if (focusZoomHolds > 0) return;
    focusZoomWatch?.disconnect();
    focusZoomWatch = null;
    for (const [meta, was] of focusZoomHeld) {
      if (meta.isConnected && /maximum-scale=1$/.test(meta.getAttribute("content") ?? "")) {
        meta.setAttribute("content", was);
      }
    }
    focusZoomHeld.clear();
  };
}

let focusZoomHolds = 0;
let focusZoomWatch: MutationObserver | null = null;
const focusZoomHeld = new Map<HTMLMetaElement, string>();
