/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { clipToVisibleEditor } from "../components/console/editor/native-document/visibleEditorClip";
import { OFFICE_TEXT_PAINT_INK, OFFICE_TEXT_PAINT_MARK } from "./office-text-paint";

const BLINK = "office-painted-caret-blink";
const CARET_WIDTH_PX = 2;
/** On a host while its painted words are selected: the browser's band is off. */
const SELECTION_ATTRIBUTE = "data-office-painted-selection";
/** One selection colour over every WordArt style. */
const SELECTION_COLOR = "rgba(38, 120, 235, 0.26)";

let installs = 0;
let caret: HTMLDivElement | null = null;
/** The host whose own caret is hidden, its caret-color to restore, and the
 * colour the browser drew that caret in. */
let hidden: { host: HTMLElement; value: string; priority: string; ink: string } | null = null;
let frame = 0;
let last = "";
let detach: (() => void) | null = null;
let band: HTMLDivElement | null = null;
let bandHost: HTMLElement | null = null;
/** Schedules a redraw while installed (a held size changed with no event). */
let requestRedraw: (() => void) | null = null;

/**
 * On the editing host (or an ancestor of the caret inside it) while a font
 * size is held at a bare caret: the held size over the size of the text the
 * caret stands beside.
 */
export const OFFICE_CARET_FONT_SCALE_ATTRIBUTE = "data-office-caret-font-scale";
/** The caret's share above the baseline: a face's ascent over its ascent and descent. */
const CARET_ASCENT_SHARE = 0.8;

export interface CaretBox {
  left: number;
  top: number;
  height: number;
}

/**
 * The scale a held size draws the caret at: `held` over `current`, both in
 * the surface's own unit. Null when either is unknown or they draw the same
 * caret (within 1%), so the browser's own caret stays.
 */
export function officeCaretFontScale(held: unknown, current: unknown): number | null {
  if (typeof held !== "number" || typeof current !== "number") return null;
  if (!(held > 0) || !(current > 0) || !Number.isFinite(held) || !Number.isFinite(current)) return null;
  const scale = held / current;
  return Math.abs(scale - 1) < 0.01 ? null : scale;
}

/**
 * A caret beside text, redrawn for text `scale` times as large typed there:
 * that much taller, on the same baseline, as Word draws the insertion point
 * of a held size.
 */
export function officeScaledCaretBox(box: CaretBox, scale: number): CaretBox {
  if (!(scale > 0) || !Number.isFinite(scale) || scale === 1) return box;
  const baseline = box.top + box.height * CARET_ASCENT_SHARE;
  const height = box.height * scale;
  return { left: box.left, top: baseline - height * CARET_ASCENT_SHARE, height };
}

/**
 * States (or, with null, withdraws) the held size's scale on `element`,
 * the editing host or an element round the caret inside it.
 */
export function setOfficeCaretFontScale(element: Element | null | undefined, scale: number | null): void {
  if (!element) return;
  const value = scale === null || !(scale > 0) || !Number.isFinite(scale) ? null : String(Math.round(scale * 1000) / 1000);
  if (element.getAttribute(OFFICE_CARET_FONT_SCALE_ATTRIBUTE) === value) return;
  if (value === null) element.removeAttribute(OFFICE_CARET_FONT_SCALE_ATTRIBUTE);
  else element.setAttribute(OFFICE_CARET_FONT_SCALE_ATTRIBUTE, value);
  requestRedraw?.();
}

/** The held size's scale round the caret at `at` inside `host`, or null. */
function caretFontScaleAt(host: HTMLElement, at: Element): number | null {
  const holder = at.closest(`[${OFFICE_CARET_FONT_SCALE_ATTRIBUTE}]`);
  if (!holder || !(holder === host || holder.contains(host) || host.contains(holder))) return null;
  const scale = Number(holder.getAttribute(OFFICE_CARET_FONT_SCALE_ATTRIBUTE));
  return scale > 0 && Number.isFinite(scale) && scale !== 1 ? scale : null;
}

/** The colour the browser drew the caret in: the caret colour, else the text's. */
function typedCaretColor(at: Element): string {
  const style = getComputedStyle(at);
  const caret = style.caretColor;
  if (caret && caret !== "auto" && cssAlpha(caret) > 0.02) return caret;
  return cssAlpha(style.color) > 0.02 ? style.color : officeCaretColorFor(at);
}

/** A box with a height: a laid-out character, element or line break. */
function shownRect(rects: Iterable<DOMRect>): DOMRect | null {
  for (const rect of rects) if (rect.height > 0) return rect;
  return null;
}

/** The caret's height beside `rect`: the line's, at most the font's, centred on it. */
function besideBox(left: number, rect: DOMRect, element: Element): CaretBox {
  const fontPx = Number.parseFloat(getComputedStyle(element).fontSize) || rect.height;
  const height = Math.min(rect.height, fontPx * 1.2);
  return { left, top: rect.top + (rect.height - height) / 2, height };
}

/**
 * Where the caret of an EMPTY line draws: where its block's alignment puts
 * the first character (centred, right, after the indent), not its left edge.
 */
function emptyLineBox(element: Element): CaretBox | null {
  let block: Element | null = element;
  while (block && /^inline/.test(getComputedStyle(block).display) && block.parentElement) block = block.parentElement;
  if (!block) return null;
  const rect = block.getBoundingClientRect();
  if (!(rect.height > 0)) return null;
  const style = getComputedStyle(block);
  const px = (value: string) => Number.parseFloat(value) || 0;
  const start = rect.left + px(style.paddingLeft) + px(style.borderLeftWidth);
  const end = rect.right - px(style.paddingRight) - px(style.borderRightWidth);
  const rtl = style.direction === "rtl";
  const align = style.textAlign;
  const indent = px(style.textIndent);
  const left =
    align === "center" || align === "-webkit-center"
      ? (start + end) / 2
      : align === "right" || (align === "end" && !rtl) || ((align === "start" || align === "justify") && rtl)
        ? end - (rtl ? indent : 0)
        : start + (rtl ? 0 : indent);
  const fontPx = px(style.fontSize) || rect.height;
  const top = rect.top + px(style.paddingTop) + px(style.borderTopWidth);
  const height = Math.min(rect.height, fontPx * 1.2);
  return { left, top, height };
}









export function officeCaretBox(range: Range): CaretBox | null {
  const own = shownRect(range.getClientRects());
  if (own) return { left: own.left, top: own.top, height: own.height };
  const node = range.startContainer;
  const offset = range.startOffset;
  if (node.nodeType === Node.TEXT_NODE) {
    const length = (node as Text).length;
    const probe = node.ownerDocument!.createRange();
    if (offset < length) {
      probe.setStart(node, offset);
      probe.setEnd(node, offset + 1);
      const rect = shownRect(probe.getClientRects());
      if (rect) return { left: rect.left, top: rect.top, height: rect.height };
    }
    if (offset > 0) {
      probe.setStart(node, offset - 1);
      probe.setEnd(node, offset);
      const rect = shownRect(probe.getClientRects());
      if (rect) return { left: rect.right, top: rect.top, height: rect.height };
    }
  }
  const element = node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  if (!element) return null;
  if (node.nodeType === Node.ELEMENT_NODE) {
    // Between two child nodes: after the one before it (an equation's
    // right edge), before the one after it, or a line break's own box
    // (where the browser draws the caret of its line).
    const before = node.childNodes[offset - 1] ?? null;
    const after = node.childNodes[offset] ?? null;
    const boxOf = (child: ChildNode): DOMRect | null => {
      if (child.nodeType === Node.ELEMENT_NODE) return shownRect([(child as Element).getBoundingClientRect()]);
      if (child.nodeType !== Node.TEXT_NODE || !(child as Text).length) return null;
      const probe = child.ownerDocument!.createRange();
      probe.selectNodeContents(child);
      const rects = Array.from(probe.getClientRects()).filter((rect) => rect.height > 0);
      return rects.length ? rects[child === before ? rects.length - 1 : 0] : null;
    };
    if (after && after.nodeName === "BR") {
      const rect = boxOf(after);
      if (rect) return besideBox(rect.left, rect, element);
    }
    if (before && before.nodeName !== "BR") {
      const rect = boxOf(before);
      if (rect) return besideBox(rect.right, rect, element);
    }
    if (after) {
      const rect = boxOf(after);
      if (rect) return besideBox(rect.left, rect, element);
    }
  }
  return emptyLineBox(element);
}

/** Dark on a light page, light on a dark one: never the words' colours. */
export function officeCaretColorFor(element: Element | null): string {
  for (let node = element; node; node = node.parentElement) {
    const match = /rgba?\(([^)]+)\)/.exec(getComputedStyle(node).backgroundColor);
    if (!match) continue;
    const [r, g, b, a = 1] = match[1].split(/[ ,/]+/).filter(Boolean).map(Number);
    if (!(a > 0.5)) continue;
    const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
    return luminance < 0.45 ? "#ffffff" : "#000000";
  }
  return "#000000";
}

function cssAlpha(value: string): number {
  if (!value || value === "transparent") return 0;
  const match = /rgba?\(([^)]+)\)/.exec(value);
  if (!match) return 1;
  const parts = match[1].split(/[ ,/]+/).filter(Boolean);
  if (parts.length < 4) return 1;
  const alpha = parts[3]!;
  return alpha.endsWith("%") ? Number.parseFloat(alpha) / 100 : Number(alpha);
}










export function officeShownTextColor(element: Element): string {
  const view = element.ownerDocument.defaultView;
  if (!view) return "#000000";
  const own = view.getComputedStyle(element);
  if (cssAlpha(own.color) > 0.02) return own.color;
  // The paint states the colour it is seen in (officeTextPaintCss).
  const ink = own.getPropertyValue(OFFICE_TEXT_PAINT_INK).trim();
  if (ink && cssAlpha(ink) > 0.02) return ink;
  for (let node: Element | null = element; node; node = node.parentElement) {
    const style = view.getComputedStyle(node);
    if (style.backgroundImage && style.backgroundImage !== "none") {
      for (const match of style.backgroundImage.matchAll(/rgba?\([^)]*\)|#[0-9a-f]{3,8}\b/gi)) {
        if (cssAlpha(match[0]) > 0.02) return match[0];
      }
    }
    const stroke = style.getPropertyValue("-webkit-text-stroke-color");
    if (Number.parseFloat(style.getPropertyValue("-webkit-text-stroke-width")) > 0 && cssAlpha(stroke) > 0.02) return stroke;
    if (node.getAttribute("data-block-id") !== null) break;
  }
  return officeCaretColorFor(element);
}

/**
 * How the words at `element` look, for something typed over them that
 * should look the same (an equation opened in place in WordArt): their
 * shown colour (officeShownTextColor), outline, surface shadows and the
 * paint's own drop shadows (glow, shadow, a clipped fill's outline ring).
 * A fill clipped to the glyphs is not carried: WebKit does not paint it
 * through an editor's shadow tree, and the equation went invisible.
 */
export function officeShownTextLook(element: Element): {
  color: string;
  WebkitTextStroke?: string;
  paintOrder?: string;
  textShadow?: string;
  filter?: string;
} {
  const view = element.ownerDocument.defaultView;
  const color = officeShownTextColor(element);
  if (!view) return { color };
  const own = view.getComputedStyle(element);
  const look: ReturnType<typeof officeShownTextLook> = { color };
  const width = own.getPropertyValue("-webkit-text-stroke-width");
  if (Number.parseFloat(width) > 0) {
    look.WebkitTextStroke = `${width} ${own.getPropertyValue("-webkit-text-stroke-color")}`;
    if (own.paintOrder && own.paintOrder !== "normal") look.paintOrder = own.paintOrder;
  }
  if (own.textShadow && own.textShadow !== "none") look.textShadow = own.textShadow;
  for (let node: Element | null = element; node; node = node.parentElement) {
    const style = view.getComputedStyle(node);
    if (style.getPropertyValue(OFFICE_TEXT_PAINT_MARK).trim() !== "1") break;
    if (style.filter && style.filter !== "none") {
      look.filter = style.filter;
      break;
    }
    if (node.getAttribute("data-block-id") !== null) break;
  }
  return look;
}

function caretElement(doc: Document): HTMLDivElement {
  if (caret?.isConnected) return caret;
  if (!doc.getElementById(BLINK)) {
    const style = doc.createElement("style");
    style.id = BLINK;
    // The browser's selection over painted words follows the paint (a
    // clipped fill drew a dark block); it is turned off while this module
    // draws the selection, the words keeping their own colours.
    style.textContent =
      `@keyframes ${BLINK}{0%,49%{opacity:1}50%,100%{opacity:0}}` +
      `[${SELECTION_ATTRIBUTE}] ::selection,[${SELECTION_ATTRIBUTE}]::selection{background-color:transparent !important;}`;
    doc.head.appendChild(style);
  }
  caret = doc.createElement("div");
  caret.setAttribute("aria-hidden", "true");
  caret.dataset.officePaintedCaret = "true";
  Object.assign(caret.style, {
    position: "fixed",
    width: `${CARET_WIDTH_PX}px`,
    pointerEvents: "none",
    zIndex: "2147483000",
    display: "none",
    borderRadius: "1px",
  });
  doc.body.appendChild(caret);
  return caret;
}

function restoreHostCaret(): void {
  if (!hidden) return;
  if (hidden.value) hidden.host.style.setProperty("caret-color", hidden.value, hidden.priority);
  else hidden.host.style.removeProperty("caret-color");
  hidden = null;
}

function hideBand(): void {
  bandHost?.removeAttribute(SELECTION_ATTRIBUTE);
  bandHost = null;
  if (band) band.replaceChildren();
}

function hide(): void {
  restoreHostCaret();
  hideBand();
  if (caret) caret.style.display = "none";
  last = "";
}

/** The selection's boxes, one per line: a run's boxes on a line merged. */
export function officeSelectionLineBoxes(range: Range): Array<{ left: number; top: number; width: number; height: number }> {
  const lines: Array<{ left: number; top: number; right: number; bottom: number }> = [];
  for (const rect of Array.from(range.getClientRects())) {
    if (!(rect.width > 0.5 && rect.height > 0)) continue;
    const line = lines.find(
      (entry) => Math.min(entry.bottom, rect.bottom) - Math.max(entry.top, rect.top) > Math.min(entry.bottom - entry.top, rect.height) / 2,
    );
    if (line) {
      line.left = Math.min(line.left, rect.left);
      line.right = Math.max(line.right, rect.right);
      line.top = Math.min(line.top, rect.top);
      line.bottom = Math.max(line.bottom, rect.bottom);
    } else {
      lines.push({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom });
    }
  }
  return lines.map((line) => ({ left: line.left, top: line.top, width: line.right - line.left, height: line.bottom - line.top }));
}

function isPainted(node: Node | null): boolean {
  const element = !node ? null : node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement;
  return Boolean(element && getComputedStyle(element).getPropertyValue(OFFICE_TEXT_PAINT_MARK).trim() === "1");
}

/** A highlighted range in painted words: drawn here, one colour for every style. */
function drawBand(doc: Document, host: HTMLElement, range: Range): void {
  if (!band?.isConnected) {
    band = doc.createElement("div");
    band.setAttribute("aria-hidden", "true");
    band.dataset.officePaintedSelectionBand = "true";
    Object.assign(band.style, { position: "fixed", left: "0", top: "0", pointerEvents: "none", zIndex: "2147482999" });
    doc.body.appendChild(band);
  }
  if (bandHost !== host) {
    bandHost?.removeAttribute(SELECTION_ATTRIBUTE);
    bandHost = host;
    host.setAttribute(SELECTION_ATTRIBUTE, "true");
  }
  // Drawn on the body, so cut to what the host's scroller and window leave
  // visible: a line scrolled under the ribbon or past the window is not
  // painted over them (the Docs drawn caret's rule, visibleEditorClip).
  const boxes = officeSelectionLineBoxes(range).flatMap((box) => {
    const shown = clipToVisibleEditor(host, box);
    return shown ? [shown] : [];
  });
  while (band.children.length > boxes.length) band.lastElementChild!.remove();
  boxes.forEach((box, index) => {
    let piece = band!.children[index] as HTMLDivElement | undefined;
    if (!piece) {
      piece = doc.createElement("div");
      Object.assign(piece.style, { position: "fixed", background: SELECTION_COLOR, borderRadius: "2px" });
      band!.appendChild(piece);
    }
    piece.style.left = `${Math.round(box.left)}px`;
    piece.style.top = `${Math.round(box.top)}px`;
    piece.style.width = `${Math.round(box.width)}px`;
    piece.style.height = `${Math.round(box.height)}px`;
  });
}

function isEditingHost(element: HTMLElement): boolean {
  if (element.isContentEditable) return true;
  const value = element.getAttribute("contenteditable");
  return value !== null && value !== "false";
}

/** The editing host a caret is in: the focused contenteditable that holds it. */
function caretHost(doc: Document): { host: HTMLElement; range: Range; at: Element } | null {
  const active = doc.activeElement;
  if (!(active instanceof HTMLElement) || !isEditingHost(active)) return null;
  const selection = doc.getSelection();
  if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return null;
  const anchor = selection.anchorNode;
  if (!anchor || !active.contains(anchor)) return null;
  const at = anchor.nodeType === Node.ELEMENT_NODE ? (anchor as Element) : anchor.parentElement;
  if (!at) return null;
  return { host: active, range: selection.getRangeAt(0), at };
}

function update(doc: Document): void {
  frame = 0;
  // A highlighted range in painted words of the focused editing host.
  const active = doc.activeElement;
  const selection = doc.getSelection();
  if (
    active instanceof HTMLElement && isEditingHost(active) &&
    selection && selection.rangeCount > 0 && !selection.isCollapsed &&
    active.contains(selection.anchorNode) && active.contains(selection.focusNode) &&
    (isPainted(selection.anchorNode) || isPainted(selection.focusNode))
  ) {
    restoreHostCaret();
    if (caret) caret.style.display = "none";
    last = "";
    drawBand(doc, active, selection.getRangeAt(0));
    return;
  }
  hideBand();
  const found = caretHost(doc);
  const painted = found ? getComputedStyle(found.at).getPropertyValue(OFFICE_TEXT_PAINT_MARK).trim() === "1" : false;
  // A size held at the caret (Grow Font with nothing highlighted).
  const scale = found ? caretFontScaleAt(found.host, found.at) : null;
  if (!found || (!painted && scale === null)) {
    hide();
    return;
  }
  const beside = officeCaretBox(found.range);
  const measured = beside && scale !== null ? officeScaledCaretBox(beside, scale) : beside;
  // On the body: cut to what the host's scroller and window leave visible,
  // never painted over the ribbon or past the window (visibleEditorClip).
  const shown = measured
    ? clipToVisibleEditor(found.host, { left: measured.left, top: measured.top, width: 0, height: measured.height })
    : null;
  const box = shown ? { left: shown.left, top: shown.top, height: shown.height } : null;
  if (!box) {
    hide();
    return;
  }
  // Painted words take one caret colour for every style; plain words keep
  // the colour the browser drew their caret in, read before it is hidden.
  if (hidden?.host !== found.host) {
    restoreHostCaret();
    hidden = {
      host: found.host,
      value: found.host.style.getPropertyValue("caret-color"),
      priority: found.host.style.getPropertyPriority("caret-color"),
      ink: typedCaretColor(found.at),
    };
    found.host.style.setProperty("caret-color", "transparent", "important");
  }
  const color = painted ? officeCaretColorFor(found.host) : hidden.ink;
  const element = caretElement(doc);
  const left = Math.round(box.left - CARET_WIDTH_PX / 2);
  const key = `${left}:${Math.round(box.top)}:${Math.round(box.height)}:${color}`;
  element.style.left = `${left}px`;
  element.style.top = `${Math.round(box.top)}px`;
  element.style.height = `${Math.round(box.height)}px`;
  element.style.background = color;
  element.style.display = "block";
  if (key !== last) {
    // A caret that moved shows at once, then blinks.
    element.style.animation = "none";
    void element.offsetWidth;
    element.style.animation = `${BLINK} 1.06s step-end infinite`;
    last = key;
  }
}

/** Draw the WordArt caret in this document until the returned call. */
export function installOfficePaintedCaret(doc: Document = document): () => void {
  installs += 1;
  if (installs === 1) {
    // With no editing host focused and nothing drawn, update() would only
    // hide what is already hidden: such an event (a grid click's focusin, a
    // pointer move, the timer) costs no frame.
    const schedule = () => {
      if (frame) return;
      const active = doc.activeElement;
      const drawn = hidden !== null || bandHost !== null || (caret !== null && caret.style.display !== "none");
      if (!drawn && !(active instanceof HTMLElement && isEditingHost(active))) return;
      frame = requestAnimationFrame(() => update(doc));
    };
    const events: Array<[EventTarget, string]> = [
      [doc, "selectionchange"],
      [doc, "focusin"],
      [doc, "focusout"],
      [doc, "input"],
      [doc, "keyup"],
      [doc, "compositionupdate"],
      [doc, "compositionend"],
      [doc, "pointerup"],
      [window, "resize"],
    ];
    for (const [target, name] of events) target.addEventListener(name, schedule, true);
    doc.addEventListener("scroll", schedule, { capture: true, passive: true });
    doc.addEventListener("pointermove", schedule, { capture: true, passive: true });
    // A relayout with no event (an autofit, a remote edit) moves the words.
    const timer = window.setInterval(schedule, 300);
    requestRedraw = () => {
      if (!frame) frame = requestAnimationFrame(() => update(doc));
    };
    detach = () => {
      requestRedraw = null;
      window.clearInterval(timer);
      doc.removeEventListener("pointermove", schedule, { capture: true });
      for (const [target, name] of events) target.removeEventListener(name, schedule, true);
      doc.removeEventListener("scroll", schedule, { capture: true });
      if (frame) cancelAnimationFrame(frame);
      frame = 0;
      hide();
      caret?.remove();
      caret = null;
      band?.remove();
      band = null;
    };
  }
  let released = false;
  return () => {
    if (released) return;
    released = true;
    installs -= 1;
    if (installs === 0) {
      detach?.();
      detach = null;
    }
  };
}
