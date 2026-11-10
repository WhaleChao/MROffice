/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's Tabs in a text box made in MRAI (text-box-typography.ts `tabStops`):
 * each tab moves the pen to the next stop past it, measured from the text
 * area's left edge; the words after it are set left of, centred on, right
 * of, or with their first decimal point on that stop; past the last stop,
 * the default interval (Word's 0.5 in where a box states none). A leader
 * fills the tab's gap with dots, hyphens or an underline.
 *
 * The writers set the same (native_edit_primitives `_text_box_tab_layout`,
 * and its runs writer); the editor and the view draw each tab as an
 * inline-block reaching its stop (applyPdfTextBoxTabStops), so the words
 * wrap and stand where the PDF has them.
 */
import type { PdfTextBoxTabStop, PdfTextBoxTypographyFields } from "@mrnote/core/web/lib/pdf/text-box-typography.ts";

/** Word's default interval (0.5 in), as a fraction of a page this many points tall. */
export function pdfTextBoxDefaultTabStop(pageHeightPoints: number): number {
  return 36 / Math.max(1, pageHeightPoints);
}

export interface PdfTextBoxTabLayoutStop {
  /** From the text area's left edge, in the unit of the measure. */
  readonly position: number;
  readonly align: PdfTextBoxTabStop["align"];
  readonly leader: PdfTextBoxTabStop["leader"];
}

export interface PdfTextBoxTabPiece {
  readonly text: string;
  /** From the line's left edge. */
  readonly x: number;
  /** The leader the tab before it fills its gap with, and where the gap starts. */
  readonly leader: PdfTextBoxTabStop["leader"];
  readonly gapStart: number;
}

/** The next stop past `pen` (from the text area's left edge); a bar is no
 * stop of the words (Word's: it only draws its rule). */
function nextStop(pen: number, stops: readonly PdfTextBoxTabLayoutStop[], interval: number): PdfTextBoxTabLayoutStop {
  const stop = stops.find((candidate) => candidate.align !== "bar" && candidate.position > pen + 1e-6);
  if (stop) return stop;
  const step = Math.max(1e-6, interval);
  return { position: (Math.floor(pen / step + 1e-9) + 1) * step, align: "left", leader: "none" };
}

/** A line's pieces between its tabs (mirror of `_text_box_tab_layout`):
 * `start` is the line's left edge from the text area's. */
export function pdfTextBoxTabLayout(line: string, options: {
  readonly measure: (text: string) => number;
  readonly start: number;
  readonly stops: readonly PdfTextBoxTabLayoutStop[];
  readonly interval: number;
}): PdfTextBoxTabPiece[] {
  const segments = line.split("\t");
  const pieces: PdfTextBoxTabPiece[] = [{ text: segments[0], x: 0, leader: "none", gapStart: 0 }];
  let pen = options.measure(segments[0]);
  for (const segment of segments.slice(1)) {
    const stop = nextStop(options.start + pen, options.stops, options.interval);
    const width = options.measure(segment);
    const decimal = segment.indexOf(".");
    const shift = stop.align === "right" ? width : stop.align === "center" ? width / 2
      : stop.align === "decimal" ? (decimal >= 0 ? options.measure(segment.slice(0, decimal)) : width) : 0;
    const x = Math.max(pen, stop.position - options.start - shift);
    pieces.push({ text: segment, x, leader: stop.leader, gapStart: pen });
    pen = x + width;
  }
  return pieces;
}

/** Where a box's bar tabs draw their rules, from the text area's left edge
 * in the unit of its stops (Word's Bar: a vertical line through every line
 * of the paragraph, the words untouched). */
export function pdfTextBoxBarPositions(stops: readonly PdfTextBoxTabLayoutStop[] | readonly PdfTextBoxTabStop[] | undefined): number[] {
  return (stops ?? []).filter((stop) => stop.align === "bar").map((stop) => stop.position);
}

/** A box's stops and interval in pixels of a page `pageHeightPx` tall (null
 * where it states none: its tabs are the browser's). */
export function pdfTextBoxTabStopsPx(
  typography: PdfTextBoxTypographyFields | undefined,
  pageHeightPx: number,
  pageHeightPoints: number,
): Readonly<{ stops: PdfTextBoxTabLayoutStop[]; interval: number; barWidth: number }> | null {
  if (!typography?.tabStops?.length && typography?.defaultTabStop === undefined) return null;
  return {
    stops: (typography.tabStops ?? []).map((stop) => ({ ...stop, position: stop.position * pageHeightPx })),
    interval: (typography.defaultTabStop ?? pdfTextBoxDefaultTabStop(pageHeightPoints)) * pageHeightPx,
    barWidth: PDF_TEXT_BOX_BAR_RULE_POINTS * pageHeightPx / Math.max(1, pageHeightPoints),
  };
}

/** A bar tab's rule, in points (the writers' `_TEXT_BOX_BAR_RULE_POINTS`):
 * Word's hairline. */
export const PDF_TEXT_BOX_BAR_RULE_POINTS = 0.75;

export const PDF_TEXT_BOX_TAB_ATTR = "data-pdf-textbox-tab";

const LEADER_BACKGROUND: Readonly<Record<PdfTextBoxTabStop["leader"], string>> = Object.freeze({
  none: "none",
  dot: "radial-gradient(circle, currentColor 0.06em, transparent 0.07em) 0 85% / 0.3em 0.3em repeat-x",
  hyphen: "linear-gradient(currentColor, currentColor) 0 62% / 100% 0.06em no-repeat",
  underscore: "linear-gradient(currentColor, currentColor) 0 100% / 100% 0.06em no-repeat",
});

/**
 * Each tab of `root` (a box's editable or read-only text) drawn as an
 * inline-block reaching its stop, measured where the browser laid the words
 * before it out, in order: the tabs after it move with it. `leftIndentPx`:
 * the box's left indent (its padding), the stops measured from the text
 * area's edge before it. The tab's character stays in the text: the input,
 * the caret and the saved words are unchanged.
 */
export function applyPdfTextBoxTabStops(root: HTMLElement, layout: Readonly<{
  stops: readonly PdfTextBoxTabLayoutStop[];
  interval: number;
  barWidth?: number;
}>, leftIndentPx = 0): void {
  const document = root.ownerDocument;
  // Every tab its own span (a text node holding one split around it).
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const tabNodes: Text[] = [];
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    if (node.data.includes("\t") && !(node.parentElement?.hasAttribute(PDF_TEXT_BOX_TAB_ATTR) && node.data === "\t")) {
      tabNodes.push(node);
    }
  }
  for (const node of tabNodes) {
    let rest: Text | null = node;
    while (rest) {
      const current: Text = rest;
      const index: number = current.data.indexOf("\t");
      if (index < 0) break;
      const tab: Text = index === 0 ? current : current.splitText(index);
      const after: Text | null = tab.data.length > 1 ? tab.splitText(1) : null;
      const span = document.createElement("span");
      span.setAttribute(PDF_TEXT_BOX_TAB_ATTR, "true");
      tab.replaceWith(span);
      span.appendChild(tab);
      rest = after;
    }
  }
  const spans = [...root.querySelectorAll<HTMLElement>(`span[${PDF_TEXT_BOX_TAB_ATTR}]`)];
  if (spans.length === 0) { applyPdfTextBoxBarRules(root, layout, leftIndentPx); return; }
  const box = root.getBoundingClientRect();
  // Screen pixels per pixel of the page the stops are in (a zoomed host).
  const zoom = root.offsetWidth > 0 ? box.width / root.offsetWidth : 1;
  const padding = Number.parseFloat(root.ownerDocument.defaultView?.getComputedStyle(root).paddingLeft ?? "0") || 0;
  const areaLeft = box.left + (root.clientLeft + padding - leftIndentPx) * zoom;
  for (const span of spans) {
    Object.assign(span.style, { display: "inline-block", whiteSpace: "pre", tabSize: "0", width: "0px",
      background: "none", textIndent: "0" } satisfies Partial<CSSStyleDeclaration>);
  }
  for (const span of spans) {
    const pen = (span.getBoundingClientRect().left - areaLeft) / zoom;
    const stop = nextStop(pen, layout.stops, layout.interval);
    let shift = 0;
    if (stop.align !== "left") {
      // The words after the tab, up to the next tab or the end of its text.
      const range = document.createRange();
      range.setStartAfter(span);
      const next = spans[spans.indexOf(span) + 1];
      if (next && next.parentElement && root.contains(next)) range.setEndBefore(next);
      else range.setEnd(root, root.childNodes.length);
      const text = range.toString();
      const line = text.split("\n")[0];
      if (line) {
        // The following words' width in their own face, measured outside
        // the box (nothing is put into an editable text).
        const style = document.defaultView?.getComputedStyle(span.parentElement ?? root);
        const probe = document.createElement("span");
        probe.style.cssText = "position:fixed;left:-10000px;top:0;visibility:hidden;white-space:pre";
        if (style) {
          probe.style.font = style.font;
          probe.style.letterSpacing = style.letterSpacing;
          probe.style.wordSpacing = style.wordSpacing;
        }
        probe.textContent = stop.align === "decimal" && line.includes(".") ? line.slice(0, line.indexOf(".")) : line;
        document.body.appendChild(probe);
        const width = probe.getBoundingClientRect().width;
        probe.remove();
        shift = stop.align === "center" ? width / 2 : width;
      }
    }
    const width = Math.max(0, stop.position - shift - pen);
    span.style.width = `${width}px`;
    span.style.background = LEADER_BACKGROUND[stop.leader];
  }
  // The rules after the tabs are set: the lines stand where they will.
  applyPdfTextBoxBarRules(root, layout, leftIndentPx);
}

export const PDF_TEXT_BOX_BARS_ATTR = "data-pdf-textbox-bars";

/**
 * A box's bar tabs (Word's Bar) drawn as rules behind its words: each from
 * the top of the first line to the foot of the last in a column (the
 * writers draw the same, through the paragraphs' spacing), at its position
 * from the text area's left edge. Drawn as the root's background images,
 * so nothing is put into an editable text; cleared once the box has none.
 */
export function applyPdfTextBoxBarRules(root: HTMLElement, layout: Readonly<{
  stops: readonly PdfTextBoxTabLayoutStop[];
  barWidth?: number;
}>, leftIndentPx: number): void {
  const bars = pdfTextBoxBarPositions(layout.stops);
  const clear = () => {
    if (!root.hasAttribute(PDF_TEXT_BOX_BARS_ATTR)) return;
    root.removeAttribute(PDF_TEXT_BOX_BARS_ATTR);
    Object.assign(root.style, { backgroundImage: "", backgroundPosition: "", backgroundSize: "",
      backgroundRepeat: "", backgroundOrigin: "" } satisfies Partial<CSSStyleDeclaration>);
  };
  if (bars.length === 0) { clear(); return; }
  const view = root.ownerDocument.defaultView;
  const style = view?.getComputedStyle(root);
  const box = root.getBoundingClientRect();
  const zoom = root.offsetWidth > 0 ? box.width / root.offsetWidth : 1;
  const padding = Number.parseFloat(style?.paddingLeft ?? "0") || 0;
  const paddingRight = Number.parseFloat(style?.paddingRight ?? "0") || 0;
  // In the root's padding box (where its backgrounds are placed), unzoomed.
  const toLocal = (clientX: number, clientY: number) => ({
    x: (clientX - box.left) / zoom - root.clientLeft, y: (clientY - box.top) / zoom - root.clientTop });
  const count = Math.max(1, Number.parseInt(style?.columnCount ?? "1", 10) || 1);
  const gap = count > 1 ? Number.parseFloat(style?.columnGap ?? "0") || 0 : 0;
  const columnWidth = count > 1 ? (root.clientWidth - padding - paddingRight - (count - 1) * gap) / count : 0;
  const range = root.ownerDocument.createRange();
  range.selectNodeContents(root);
  // The lines: the text's own boxes, and its inline elements' (words drawn
  // as generated content have no text node). A block's box is left out: its
  // padding is the paragraph's spacing, above its first line.
  const rects: DOMRect[] = [...range.getClientRects()];
  for (const element of root.querySelectorAll<HTMLElement>("*")) {
    if (view?.getComputedStyle(element).display.startsWith("inline")) rects.push(...element.getClientRects());
  }
  const extents: Array<{ top: number; bottom: number } | undefined> = [];
  for (const rect of rects) {
    if (!(rect.height > 0)) continue;
    const local = toLocal(rect.left, rect.top);
    const column = count > 1
      ? Math.min(count - 1, Math.max(0, Math.floor((local.x - padding + gap / 2) / (columnWidth + gap)))) : 0;
    const bottom = local.y + rect.height / zoom;
    const extent = extents[column];
    extents[column] = extent ? { top: Math.min(extent.top, local.y), bottom: Math.max(extent.bottom, bottom) }
      : { top: local.y, bottom };
  }
  const width = Math.max(1, layout.barWidth ?? 1);
  const rules: string[] = [];
  extents.forEach((extent, column) => {
    if (!extent) return;
    for (const position of bars) {
      const x = padding - leftIndentPx + column * (columnWidth + gap) + position - width / 2;
      rules.push(`${x}px ${extent.top}px / ${width}px ${Math.max(0, extent.bottom - extent.top)}px`);
    }
  });
  if (rules.length === 0) { clear(); return; }
  root.setAttribute(PDF_TEXT_BOX_BARS_ATTR, "true");
  Object.assign(root.style, {
    backgroundImage: rules.map(() => "linear-gradient(currentColor, currentColor)").join(", "),
    backgroundPosition: rules.map((rule) => rule.split(" / ")[0]).join(", "),
    backgroundSize: rules.map((rule) => rule.split(" / ")[1]).join(", "),
    backgroundRepeat: "no-repeat",
    backgroundOrigin: "padding-box",
  } satisfies Partial<CSSStyleDeclaration>);
}
