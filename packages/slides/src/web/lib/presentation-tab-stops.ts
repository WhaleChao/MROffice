/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { createContext } from "react";

import { hangingMarginTabTarget, pxToTwips, wordTabAdvanceAt, type TabStop } from "@mrnote/core/web/lib/document-tabstops.ts";

/**
 * A text box's custom tab stops (`a:pPr/a:tabLst/a:tab`, ECMA-376 §21.1.2.2.13):
 * `pos` in design px from the text area's left edge, `align` from `@algn`
 * (`l` / `ctr` / `r` / `dec`). The page stores a LEFT stop as its bare
 * position (the form every page before the alignment reading holds) and any
 * other as `{ pos, align }`.
 */
export type PresentationTabAlign = "left" | "center" | "right" | "decimal";
export interface PresentationTabStop {
  pos: number;
  align: PresentationTabAlign;
}

/** PowerPoint's default tab interval (`a:bodyPr`'s `defTabSz`, 1 in), twips. */
export const POWERPOINT_DEFAULT_TAB_INTERVAL_TWIPS = 1440;

const MAX_PRESENTATION_TAB_STOPS = 20;

export function normalizePresentationTabStops(value: unknown): PresentationTabStop[] {
  if (!Array.isArray(value)) return [];
  const stops: PresentationTabStop[] = [];
  for (const entry of value.slice(0, MAX_PRESENTATION_TAB_STOPS)) {
    if (typeof entry === "number" && Number.isFinite(entry) && entry > 0) {
      stops.push({ pos: entry, align: "left" });
      continue;
    }
    if (!entry || typeof entry !== "object") continue;
    const record = entry as { pos?: unknown; align?: unknown };
    if (typeof record.pos !== "number" || !Number.isFinite(record.pos) || record.pos <= 0) continue;
    const align =
      record.align === "center" || record.align === "right" || record.align === "decimal" ? record.align : "left";
    stops.push({ pos: record.pos, align });
  }
  return stops.sort((a, b) => a.pos - b.pos);
}

/** The stops as the shared tab layout (`document-tabstops`) takes them. */
export function presentationTabStopsForLayout(stops: readonly PresentationTabStop[]): TabStop[] {
  return stops.map((stop) => ({ pos: pxToTwips(stop.pos), align: stop.align, leader: "none" }));
}

/**
 * Whether the text renderer draws each tab as its own measured span (the
 * box states custom stops): the span's width is set by
 * {@link layoutPresentationTabs} after layout.
 */
export const PresentationMeasuredTabsContext = createContext(false);

/** The attribute a measured tab's span carries. */
export const PRESENTATION_TAB_ATTRIBUTE = "data-presentation-tab";

/**
 * Where a tab on a hanging first line goes: {@link hangingMarginTabTarget}
 * (shared with the text engine, which sets the same tabs for print).
 */
export { hangingMarginTabTarget };

/** The first custom stop (design px from the text area's left) past `penPx`, or null. */
function nextCustomStopPx(penPx: number, stops: readonly PresentationTabStop[]): number | null {
  return stops.find((stop) => stop.pos > penPx + 0.75)?.pos ?? null;
}

/** The rows' alignment PowerPoint lays a tabbed line out under before aligning it whole. */
const LINE_ALIGNED_AFTER_TABS = new Set(["center", "right", "end", "-webkit-center", "-webkit-right"]);

/**
 * Lay out every measured tab span under `host`, left to right, as PowerPoint
 * does: a tab moves to the first stop right of the pen; a left stop starts
 * the next text there, a centre stop centres it on the stop, a right stop
 * ends it there, a decimal stop puts its first "." there; past the last
 * custom stop the pen goes to the next default stop. The rule is the shared
 * one Word's layout uses (`wordTabAdvanceAt`); positions run from the text
 * area's left edge (the host's content box), in design px.
 */
export function layoutPresentationTabs(
  host: HTMLElement,
  stops: readonly PresentationTabStop[],
  /** The default interval (`@defTabSz`), design px; unstated, PowerPoint's inch. */
  defaultTabPx?: number,
): void {
  const spans = Array.from(host.querySelectorAll<HTMLElement>(`[${PRESENTATION_TAB_ATTRIBUTE}]`));
  if (!spans.length) return;
  const hostRect = host.getBoundingClientRect();
  const scale = host.offsetWidth > 0 ? hostRect.width / host.offsetWidth : 1;
  if (!(scale > 0)) return;
  const view = host.ownerDocument.defaultView;
  const style = view?.getComputedStyle(host);
  const columnLeft = hostRect.left + (parseFloat(style?.paddingLeft ?? "0") || 0) * scale;
  // A body in columns (`a:bodyPr@numCol`) sets each column's stops from that
  // column's own left edge, as PowerPoint does.
  const flow = [host, host.parentElement].find((element) => {
    const count = element ? Number.parseInt(view?.getComputedStyle(element).columnCount ?? "", 10) : NaN;
    return Number.isFinite(count) && count > 1;
  });
  const flowStyle = flow ? view?.getComputedStyle(flow) : undefined;
  const columns = flowStyle ? Number.parseInt(flowStyle.columnCount, 10) : 1;
  const columnGapPx = flowStyle ? parseFloat(flowStyle.columnGap) || 0 : 0;
  const columnPitchPx =
    flow && columns > 1
      ? (flow.clientWidth -
          (parseFloat(flowStyle?.paddingLeft ?? "0") || 0) -
          (parseFloat(flowStyle?.paddingRight ?? "0") || 0) -
          columnGapPx * (columns - 1)) /
          columns +
        columnGapPx
      : 0;
  const penOf = (span: HTMLElement): number => {
    const pen = (span.getBoundingClientRect().left - columnLeft) / scale;
    if (!(columnPitchPx > 0)) return pen;
    const column = Math.max(0, Math.min(columns - 1, Math.floor((pen + 0.5) / columnPitchPx)));
    return pen - column * columnPitchPx;
  };
  const layoutStops = presentationTabStopsForLayout(stops);
  const rowOf = (span: HTMLElement): HTMLElement =>
    (span.closest("[data-presentation-paragraph-index], [data-testid='presentation-text-row']") as HTMLElement | null) ?? host;
  // A centred or right-aligned row with tabs is laid out as PowerPoint lays
  // it: tabs set from the line's own start (as if aligned left), then the
  
  // segment of "tc1\tcentred\tparagraph" where that rule puts it, to
  // 0.1 pt). Measured from a centred line, each tab moved the line it was
  // measured on and the words ran together.
  const realigned: Array<{ row: HTMLElement; textAlign: string }> = [];
  for (const row of new Set(spans.map(rowOf))) {
    const align = view?.getComputedStyle(row).textAlign ?? "";
    const rtl = view?.getComputedStyle(row).direction === "rtl";
    if (!rtl && LINE_ALIGNED_AFTER_TABS.has(align)) {
      realigned.push({ row, textAlign: row.style.textAlign });
      row.style.textAlign = "left";
    }
  }
  for (const span of spans) span.style.width = "0px";
  for (let index = 0; index < spans.length; index += 1) {
    const span = spans[index];
    const row = rowOf(span);
    const nextSpan = spans[index + 1];
    const range = host.ownerDocument.createRange();
    range.setStartAfter(span);
    if (nextSpan && row.contains(nextSpan)) range.setEndBefore(nextSpan);
    else range.setEnd(row, row.childNodes.length);
    // The text after the tab on THIS line: every box of it on the first
    // line it lands on (a run, a word apart, each is a box of its own; the
    // first alone set "centre text" on a centre stop by "centre", 20 pt off).
    const boxes = Array.from(range.getClientRects()).filter((rect) => rect.width > 0 || rect.height > 0);
    const firstBox = boxes[0];
    const lineBoxes = firstBox
      ? boxes.filter((rect) => rect.top < firstBox.bottom - 0.5 && rect.bottom > firstBox.top + 0.5)
      : [];
    const widthPx = lineBoxes.length
      ? (Math.max(...lineBoxes.map((rect) => rect.right)) - Math.min(...lineBoxes.map((rect) => rect.left))) / scale
      : 0;
    const text = range.toString();
    const point = text.indexOf(".");
    let decimalPx: number | undefined;
    if (point >= 0) {
      const decimal = host.ownerDocument.createRange();
      decimal.setStart(range.startContainer, range.startOffset);
      // Walk to the "."'s text position.
      const walker = host.ownerDocument.createTreeWalker(row, NodeFilter.SHOW_TEXT);
      let remaining = point;
      let node: Node | null;
      let started = false;
      while ((node = walker.nextNode())) {
        if (!started) {
          if (!range.intersectsNode(node)) continue;
          started = true;
        }
        const length = (node.textContent ?? "").length;
        const offset = node === range.startContainer ? range.startOffset : 0;
        if (remaining < length - offset) {
          decimal.setEnd(node, offset + remaining);
          break;
        }
        remaining -= length - offset;
      }
      const rect = decimal.getClientRects()[0];
      decimalPx = rect ? rect.width / scale : undefined;
    }
    // A hanging first line (negative indent, no bullet): its left margin.
    const hangPx = hangingMarginOf(row, span);
    const advanceFrom = (penPx: number) => {
      const hang = hangingMarginTabTarget(penPx, nextCustomStopPx(penPx, stops), hangPx);
      return hang !== null ? hang - penPx : layoutAdvance(penPx);
    };
    const layoutAdvance = (penPx: number) =>
      wordTabAdvanceAt({
        penPx,
        next: decimalPx === undefined ? { widthPx } : { widthPx, decimalPx },
        stops: layoutStops,
        defaultIntervalTwips:
          typeof defaultTabPx === "number" && defaultTabPx > 0
            ? pxToTwips(defaultTabPx)
            : POWERPOINT_DEFAULT_TAB_INTERVAL_TWIPS,
      });
    const placed = span.getBoundingClientRect();
    span.style.width = `${advanceFrom(penOf(span))}px`;
    // A tab whose stop lies past the line's end goes to the next line, and
    
    
    // sets every tab at the start of a line, an inch in). The browser carried
    // the span down; it is measured again where it now stands.
    const moved = span.getBoundingClientRect();
    if (moved.top > placed.top + 1 && moved.left < placed.left) {
      span.style.width = `${advanceFrom(penOf(span))}px`;
    }
  }
  for (const { row, textAlign } of realigned) row.style.textAlign = textAlign;

  /** A row's left margin in pen space when `span` sits on its hanging first line, else null. */
  function hangingMarginOf(row: Element, span: HTMLElement): number | null {
    if (!(row instanceof HTMLElement) || row === host) return null;
    const rowStyle = view?.getComputedStyle(row);
    const indent = parseFloat(rowStyle?.textIndent ?? "0") || 0;
    if (!(indent < -0.01)) return null;
    // A bullet stands in the hang: it is the marker's, not a stop.
    if (row.querySelector("[data-testid='presentation-paragraph-marker'], [data-presentation-marker]")) return null;
    const start = host.ownerDocument.createRange();
    start.selectNodeContents(row);
    const first = Array.from(start.getClientRects()).find((rect) => rect.height > 0);
    const at = span.getBoundingClientRect();
    if (first && Math.abs(at.top - first.top) > Math.max(1, first.height / 2)) return null;
    const rowBox = row.getBoundingClientRect();
    const padding = parseFloat(rowStyle?.paddingLeft ?? "0") || 0;
    return (rowBox.left - columnLeft) / scale + padding;
  }
}
