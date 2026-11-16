/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { SheetFact } from "../../../../../lib/editor/sheet-facts";
import { PX_PER_PT } from "../../../../../lib/office-units";
import { Extension } from "@tiptap/core";
import type { Node as PmNode } from "@tiptap/pm/model";
import { Plugin, PluginKey } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

import {
  normalizePositionalTab,
  normalizeTabStops,
  TAB_STOP_REACH_PX,
  WORD_DEFAULT_TAB_INTERVAL_TWIPS,
  wordTabAdvanceAt,
  wordTabAdvances,
  wordTabLeader,
  withHangingIndentStop,
  type PositionalTab,
} from "../../../../../lib/document-tabstops";

import { inlineTextWidthPx } from "./bandRichDom";
import { BAND_RUN_MARK_NAME } from "./bandRunMark";
import { measurementRange } from "./measurementRange";
import { SheetMeasureMemo, sheetsOf } from "./sheetMeasureMemo";
import { MEASURE_HORIZON_EVENT, sheetsWithinMeasureHorizon } from "./measureHorizon";
import { editorFontFamilies, facesUsedBy, loadedFaceRanges, textDrawsCodePointIn } from "./ledger/ledgerFontRanges";
import { afterFontsReady, newFontFacesOf } from "../../../../../lib/font-face-news";
import { clearLayoutPending, markLayoutPending, registerLayoutSettler } from "./layoutPending";
import { REBALANCE_META as PAGINATION_MOVE_META } from "./pageRebalanceExtension";

/** The tabs' decorations, and how many edits from outside the document has had
 * (the pagination's own moves not counted). */
interface BandTabState {
  set: DecorationSet;
  edits: number;
}

const bandTabLayoutKey = new PluginKey<BandTabState>("bandTabLayout");

interface TabbedLine {
  /** Document positions of the line's tab characters. */
  tabs: number[];
  /** Word's absolute position tab each of them is, if it is one. */
  positional: (PositionalTab | null)[];
  /** [from, to) document ranges of the text around them (tabs + 1). */
  segments: Array<[number, number]>;
  stops: ReturnType<typeof normalizeTabStops>;
  contentStart: number;
  contentEnd: number;
  /** Document position of the text block itself. */
  blockPos: number;
  /** The paragraph's left and right indents, in px. */
  indentLeftPx: number;
  indentRightPx: number;
  /** Word 2010 compatibility: a tab may run past the right margin and the
   * text after it stays on the line (`source.wordTabOverMargin`). */
  overMargin: boolean;
}

function indentPx(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? value * PX_PER_PT : 0;
}

/** Text blocks found to hold no tab: a block keeps its node wherever the
 * pages move it, and a long book's every block built its text again on
 * every layout to look for one. */
const tablessBlocks = new WeakSet<PmNode>();

const HOLDS_TAB = new SheetFact((node) => node.isText && (node.text ?? "").includes("\t"));

function tabbedLines(doc: PmNode): TabbedLine[] {
  const lines: TabbedLine[] = [];
  HOLDS_TAB.descendants(doc, (node, pos) => {
    if (!node.isTextblock) return true;
    // Code keeps its own tab width: Word's stops are a paragraph's.
    if (node.type.spec.code) return false;
    if (tablessBlocks.has(node)) return false;
    if (!node.textContent.includes("\t")) {
      tablessBlocks.add(node);
      return false;
    }
    const contentStart = pos + 1;
    const blockPos = pos;
    const tabs: number[] = [];
    const positional: (PositionalTab | null)[] = [];
    // Shift+Enter ends a line: the text after a tab runs only to it.
    const breaks: number[] = [];
    node.forEach((child, offset) => {
      if (child.type.name === "hardBreak") breaks.push(contentStart + offset);
      if (!child.isText) return;
      const text = child.text ?? "";
      // Every tab of a positional tab's text node is one (the editor joins
      // two alike ones into one node), as the story reads it back.
      const ptab = normalizePositionalTab(
        child.marks.find((mark) => mark.type.name === BAND_RUN_MARK_NAME)?.attrs.ptab ?? null,
      );
      for (let index = text.indexOf("\t"); index >= 0; index = text.indexOf("\t", index + 1)) {
        tabs.push(contentStart + offset + index);
        positional.push(ptab);
      }
    });
    const contentEnd = contentStart + node.content.size;
    const segments: Array<[number, number]> = [];
    let from = contentStart;
    for (const tab of tabs) {
      segments.push([from, tab]);
      from = tab + 1;
    }
    segments.push([from, contentEnd]);
    // The text a tab aligns is the rest of its own line: a right stop on the
    // first line of "label<Tab>12<Shift+Enter>longer label<Tab>34" measured
    // "12" and the second line's label together, and drew no gap (WUA 142).
    for (const [index, [start, end]] of segments.entries()) {
      const cut = breaks.find((at) => at >= start && at < end);
      if (cut !== undefined) segments[index] = [start, cut];
    }
    lines.push({
      tabs,
      positional,
      segments,
      stops: withHangingIndentStop(normalizeTabStops(node.attrs.tabStops ?? null), node.attrs.indentLeftPt, node.attrs.firstLineIndentPt),
      contentStart,
      contentEnd,
      blockPos,
      indentLeftPx: indentPx(node.attrs.indentLeftPt),
      indentRightPx: indentPx(node.attrs.indentRightPt),
      overMargin: (node.attrs.source as { wordTabOverMargin?: unknown } | null)?.wordTabOverMargin === 1,
    });
    return false;
  });
  return lines;
}

/** Whether any of the line's drawn text boxes reaches past the right edge:
 * tabs the browser lets hang there (pre-wrap white space) where Word wraps. */
/**
 * Whether a line's pen passes the paragraph's right edge. The tolerance a stop
 * is reached with is no part of this: a right stop 0.6 pt past the margin
 * (corpus R282-D16: 9072 twips in a 9060-twip column, mode 14) reads as at the
 * edge, the line is laid out as if it fitted, the browser breaks it by that
 * 0.6 pt, and four lines of a form take two each.
 */
export function tabRunsPastRightEdge(atPx: number, rightEdgePx: number): boolean {
  return atPx > rightEdgePx + 0.01;
}

function drawnPastRightEdge(
  view: EditorView,
  line: TabbedLine,
  origin: number,
  scale: number,
  rightEdgePx: number,
): boolean {
  try {
    const start = view.domAtPos(line.contentStart);
    const end = view.domAtPos(line.contentEnd);
    const range = measurementRange(view.dom.ownerDocument, "bandTabLayout");
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    return Array.from(range.getClientRects()).some(
      (rect) => rect.width > 0 && (rect.right - origin) / scale > rightEdgePx,
    );
  } catch {
    return false;
  }
}

/** How many times a wrapped paragraph's tabs are measured again after a
 * change before the widths are taken as they stand. */
const MAX_SETTLE_PASSES = 4;

/** A computed `tab-size` that is a length, in px: the surface's is the
 * document's Default tab stops, a block's is what the browser draws its tabs
 * with (a paragraph with stops of its own states another). */
function nativeTabIntervalPx(style: CSSStyleDeclaration | null | undefined): number | null {
  if (!style) return null;
  const value = style.tabSize ?? "";
  if (!value.endsWith("px")) return null;
  const px = Number.parseFloat(value);
  return Number.isFinite(px) && px > 0 ? px : null;
}

/** One tabbed line's tabs as measured: its keys (without the line's index)
 * and decorations, positions from the line's block. */
interface LineTabDecoration { node: boolean; from: number; to: number; attrs: Record<string, string> }
interface LineTabs { keys: readonly string[]; decorations: readonly LineTabDecoration[]; wrapped: boolean }
interface SheetTabs { at: string; lines: Map<PmNode, LineTabs> }
interface TabLinePlace { sheet: HTMLElement; page: PmNode; block: PmNode }

// A line's measurement reads its own page only (its block's box, its text's
// rects and the root's metrics, which `at` keys): on a page nothing changed
// in since, under the same page node, it measures what it measured (the page
// ledger's rule, sheetMeasureMemo). A font arriving forgets them all.
const tabMeasureMemos = new WeakMap<HTMLElement, SheetMeasureMemo<SheetTabs>>();

function tabMeasureMemo(root: HTMLElement): SheetMeasureMemo<SheetTabs> | null {
  let memo = tabMeasureMemos.get(root);
  if (!memo && typeof root.ownerDocument.defaultView?.MutationObserver === "function") {
    tabMeasureMemos.set(root, (memo = new SheetMeasureMemo<SheetTabs>(root)));
  }
  return memo ?? null;
}

function tabLinePlace(view: EditorView, root: HTMLElement, line: TabbedLine): TabLinePlace | null {
  const doc = view.state.doc;
  const block = doc.nodeAt(line.blockPos);
  const $at = doc.resolve(line.blockPos);
  if (!block || $at.depth < 1) return null;
  const page = $at.node(1);
  if (page.type.name !== "pageBlock") return null;
  const dom = view.nodeDOM(line.blockPos);
  const sheet = dom instanceof HTMLElement ? dom.closest("article[data-page-sheet]") : null;
  return sheet instanceof HTMLElement && sheet.parentElement === root ? { sheet, page, block } : null;
}

function keptLineTabs(memo: SheetMeasureMemo<SheetTabs>, place: TabLinePlace, at: string): LineTabs | undefined {
  const kept = memo.get(place.sheet, place.page);
  return kept && kept.at === at ? kept.lines.get(place.block) : undefined;
}

function keepLineTabs(memo: SheetMeasureMemo<SheetTabs>, place: TabLinePlace, at: string, tabs: LineTabs): void {
  let kept = memo.get(place.sheet, place.page);
  if (!kept || kept.at !== at) memo.set(place.sheet, place.page, (kept = { at, lines: new Map() }));
  kept.lines.set(place.block, tabs);
}

/**
 * The measurements a paragraph's tabs have had since the layout's inputs last
 * changed, by its node. Tabs sized for one line can make that line wrap, and
 * sized for the lines it wraps into make it fit one again: no widths hold
 * (R350-D07's interlinear glosses, eight left stops, 21 px and 43 px in turn).
 * Stopped by the count of passes, a paragraph stood on whichever of the two the
 * count ran out on, and one opening laid the book out on 159 sheets, the next
 * on 160. Once a paragraph comes back to widths it had, it takes, of those it
 * went round, the narrowest: the one its text stands on the fewest lines with
 * (Word, whose tabs are laid out with the line, sets such a gloss on one).
 */
interface TabRound { times: number; widthPx: number; widths: number[]; shape: string; tabs: LineTabs }
export type TabRounds = WeakMap<PmNode, TabRound[]>;

/** A line's tab widths, and what else its keys say (each tab's place and leader). */
function tabWidthsOf(tabs: LineTabs): { widths: number[]; shape: string } {
  const widths: number[] = [];
  const shape: string[] = [];
  for (const item of tabs.keys) {
    const [index, advance, leader] = item.split(":");
    const px = Number.parseFloat(advance ?? "");
    if (Number.isFinite(px)) {
      widths.push(px);
      shape.push(`${index}:${leader}`);
    } else shape.push(item);
  }
  return { widths, shape: shape.join(",") };
}

export function settledRound(rounds: TabRounds | null, block: PmNode | null | undefined, tabs: LineTabs): LineTabs {
  if (!rounds || !block) return tabs;
  let seen = rounds.get(block);
  if (!seen) rounds.set(block, (seen = []));
  const { widths, shape } = tabWidthsOf(tabs);
  // Within a stop's reach of each other, two readings are one width: the half
  // pixels a justified line's spacing reads a tab at are neither a new width
  // nor one that did not come back.
  const entry = seen.find((candidate) =>
    candidate.shape === shape &&
    candidate.widths.length === widths.length &&
    candidate.widths.every((width, index) => Math.abs(width - widths[index]) <= TAB_STOP_REACH_PX));
  if (entry) entry.times += 1;
  else seen.push({ times: 1, widthPx: widths.reduce((sum, width) => sum + width, 0), widths, shape, tabs });
  // Round: widths it had come back, or the paragraph measured on one line and
  // on two in turn (the count of passes can run out before the widths come
  // back: R350-D07, one gloss one line on one opening and two on the next).
  // (A count of widths is no round: taken by one, the narrowest of a justified
  // paragraph's readings moved a Greek paragraph's first line, R282-D25.)
  let flipped = false;
  if (!entry && seen.length >= 2) {
    flipped = seen.some((candidate) => candidate.tabs.wrapped) && seen.some((candidate) => !candidate.tabs.wrapped);
  }
  if ((!entry && !flipped) || seen.length < 2) return tabs;
  let narrowest: TabRound | null = null;
  for (const candidate of seen) {
    if (!narrowest || candidate.widthPx < narrowest.widthPx - 0.01 ||
      (Math.abs(candidate.widthPx - narrowest.widthPx) <= 0.01 && candidate.tabs.keys.join(",") < narrowest.tabs.keys.join(","))) {
      narrowest = candidate;
    }
  }
  // Settled: nothing more to read again for it.
  return narrowest ? { ...narrowest.tabs, wrapped: false } : tabs;
}

function measure(view: EditorView, rounds: TabRounds | null = null): { set: DecorationSet; key: string; wrapped: boolean; deferred: boolean; firstRead: boolean } | null {
  const root = view.dom as HTMLElement;
  const width = root.offsetWidth;
  if (!width) return null;
  const scale = root.getBoundingClientRect().width / width || 1;
  const rootStyle = root.ownerDocument.defaultView?.getComputedStyle(root);
  const paddingLeft = Number.parseFloat(rootStyle?.paddingLeft ?? "0") || 0;
  const paddingRight = Number.parseFloat(rootStyle?.paddingRight ?? "0") || 0;
  const rootOrigin = root.getBoundingClientRect().left + paddingLeft * scale;
  // The story's text column, margin to margin: where a positional tab set
  // relative to the margins measures from.
  const rootColumnPx = Math.max(0, root.clientWidth - paddingLeft - paddingRight);
  // Word's Default tab stops for this document.
  const defaultIntervalPx = nativeTabIntervalPx(rootStyle);
  const defaultIntervalTwips =
    defaultIntervalPx !== null ? defaultIntervalPx * 15 : WORD_DEFAULT_TAB_INTERVAL_TWIPS;
  const decorations: Decoration[] = [];
  const keys: string[] = [];
  let wrapped = false;
  // Each line's tabs as measured, kept per page (see tabMeasureMemo): a
  // keystroke measured every tabbed line of the document again (a long
  // book's contents rows: a layout of every page, ~1 s under load).
  const memo = tabMeasureMemo(root);
  const memoAt = `${width}|${scale}|${paddingLeft}|${paddingRight}|${defaultIntervalPx}`;
  const measureLine = (line: TabbedLine): LineTabs => {
    const lineKeys: string[] = [];
    const lineDecorations: LineTabDecoration[] = [];
    let lineWrapped = false;
    const done = (): LineTabs => ({ keys: lineKeys, decorations: lineDecorations, wrapped: lineWrapped });
    // The paragraph's own containing block: its box plus the margins its
    // indents render as. On a page that is the sheet's text column; in a
    // table it is the cell. The story's root serves when the block has no
    // element of its own (a measurement before the view drew it).
    let origin = rootOrigin;
    let columnPx = rootColumnPx;
    // The interval the browser draws this block's tabs with, and where the
    // block's own text edge stands from the column's: it counts its tab
    // stops from there.
    let blockIntervalPx: number | null = null;
    let textEdgePx = 0;
    const blockDom = view.nodeDOM(line.blockPos);
    if (blockDom instanceof HTMLElement && blockDom.ownerDocument.defaultView) {
      const rect = blockDom.getBoundingClientRect();
      const style = blockDom.ownerDocument.defaultView.getComputedStyle(blockDom);
      blockIntervalPx = nativeTabIntervalPx(style);
      // A right-to-left paragraph lays its tabs out from the right margin,
      // which the pen below (left to right, Word's LTR rule) does not model;
      // its segment widths read off bidi-reordered client rects moved with
      // every tab width installed, so each keystroke re-sized every tab of
      // every RTL line in the document and paginated those pages again.
      // Such a line keeps the browser's uniform tab until RTL stops are
      // modelled.
      if (style.direction === "rtl") return done();
      const marginLeft = Number.parseFloat(style.marginLeft) || 0;
      const marginRight = Number.parseFloat(style.marginRight) || 0;
      if (rect.width > 0) {
        origin = rect.left - marginLeft * scale;
        columnPx = rect.width / scale + marginLeft + marginRight;
      }
      textEdgePx =
        marginLeft +
        (Number.parseFloat(style.borderLeftWidth) || 0) +
        (Number.parseFloat(style.paddingLeft) || 0);
    }
    const customStops = (line.stops ?? []).some((stop) => stop.align !== "bar");
    if (
      !customStops &&
      !line.positional.some(Boolean) &&
      defaultIntervalPx !== null &&
      blockIntervalPx !== null &&
      Math.abs(blockIntervalPx - defaultIntervalPx) < 0.01
    ) {
      // Word's default stops count from the column's edge. The browser's
      // count from the block's text edge: the same stops when that edge
      // stands on one, and then its own tab is Word's on every line.
      const off = ((textEdgePx % defaultIntervalPx) + defaultIntervalPx) % defaultIntervalPx;
      // ... unless a tab runs past the right edge: the browser lets it hang
      // there, where Word (2013 and later) takes it to the next line, to that
      // line's first stop (corpus R37-D22: 22 default tabs, ten to a line).
      const rightEdgePx = columnPx - line.indentRightPx + TAB_STOP_REACH_PX;
      const hangs = !line.overMargin && drawnPastRightEdge(view, line, origin, scale, rightEdgePx);
      if (!hangs && (off <= TAB_STOP_REACH_PX || defaultIntervalPx - off <= TAB_STOP_REACH_PX)) return done();
    }
    // The text between two tabs as drawn, all of it even where it has
    // wrapped: coordinates of its two ends would measure the gap between two
    // lines, and a tab sized from that keeps the text wrapped for good.
    const segmentWidth = ([from, to]: [number, number]) => {
      if (from >= to) return 0;
      try {
        const start = view.domAtPos(from);
        const end = view.domAtPos(to);
        const range = measurementRange(root.ownerDocument, "bandTabLayout");
        range.setStart(start.node, start.offset);
        range.setEnd(end.node, end.offset);
        const width = inlineTextWidthPx(Array.from(range.getClientRects()), scale);
        if (width !== null) return width;
      } catch {
        // No DOM range for these positions: measure their coordinates.
      }
      return Math.max(0, (view.coordsAtPos(to, -1).left - view.coordsAtPos(from, 1).left) / scale);
    };
    // The text right after a tab up to its first break (a space or a soft
    // hyphen; a picture counts whole): Word breaks a line before a tab, not
    // after it, so a tab and that word move to the next line together.
    const firstWordWidth = ([from, to]: [number, number]): number => {
      let end = to;
      view.state.doc.nodesBetween(from, to, (node, pos) => {
        if (end !== to || !node.isText || !node.text) return end === to;
        const start = Math.max(from, pos);
        const text = node.text.slice(start - pos, to - pos);
        const at = text.search(/[ \u00ad]/);
        if (at >= 0) end = start + at;
        return false;
      });
      return segmentWidth([from, end]);
    };
    // A decimal stop aligns the text's first decimal point on the stop, so
    // the width up to it is measured too (DB: Word parity audit, decimal
    // tab drew as a right tab).
    const hasDecimalStop = (line.stops ?? []).some((stop) => stop.align === "decimal");
    const decimalAt = ([from, to]: [number, number]): number | null => {
      let found: number | null = null;
      view.state.doc.nodesBetween(from, to, (node, pos) => {
        if (found !== null || !node.isText || !node.text) return found === null;
        const start = Math.max(from, pos);
        const offset = node.text.slice(start - pos, to - pos).indexOf(".");
        if (offset >= 0) found = start + offset;
        return false;
      });
      return found;
    };
    const segments = line.segments.map((range) => {
      const widthPx = segmentWidth(range);
      if (!hasDecimalStop) return { widthPx };
      const point = decimalAt(range);
      return point === null ? { widthPx } : { widthPx, decimalPx: segmentWidth([range[0], point]) };
    });
    const first = view.coordsAtPos(line.contentStart, 1);
    const startPx = (first.left - origin) / scale;
    const column = { widthPx: columnPx, indentLeftPx: line.indentLeftPx, indentRightPx: line.indentRightPx };
    // Drawn on more than one line: the last character starts at or below
    // the first one's foot.
    const last = line.contentEnd > line.contentStart ? view.coordsAtPos(line.contentEnd, -1) : first;
    // Word 2010 compatibility (and earlier): a tab whose stop lies past the
    // paragraph's right edge still goes to it, and the text after it stays
    // on that line, running past the margin (Word 16.86, corpus R36-D15 in
    // compatibility mode 12: a left stop at 9500 twips in a 9026-twip
    // column puts the rest of the paragraph on one line). Laid out as one
    // line, and the paragraph does not wrap.
    let overMargin = false;
    let oneLine: number[] | null = null;
    if (line.overMargin) {
      oneLine = wordTabAdvances({
        segments,
        stops: line.stops,
        startPx,
        defaultIntervalTwips,
        positional: line.positional,
        line: column,
        stopsEndAtRightEdge: !line.overMargin,
      });
      const rightEdge = columnPx - line.indentRightPx;
      let at = startPx + segments[0].widthPx;
      for (const [index, advance] of oneLine.entries()) {
        at += advance;
        if (tabRunsPastRightEdge(at, rightEdge)) overMargin = true;
        at += segments[index + 1].widthPx;
      }
    }
    // The middle of the last character's box below the first one's foot.
    // Its top alone is not enough: with Word's line metrics the lines can
    // stand closer than the font's own box is tall (Calibri 11 pt: an 18 px
    // box every 15 px at this zoom), so a second line's box starts above the
    // first one's foot and a wrapped paragraph read as one line (corpus
    // R49-D14 TA12f: the tab never went to the next line).
    const wraps = !overMargin && (last.top + last.bottom) / 2 >= first.bottom - 1;
    if (wraps) lineWrapped = true;
    const advances = wraps
      ? null
      : (overMargin ? oneLine : null) ??
        wordTabAdvances({
          segments,
          stops: line.stops,
          startPx,
          defaultIntervalTwips,
          positional: line.positional,
          line: column,
          stopsEndAtRightEdge: !line.overMargin,
        });
    if (overMargin) {
      lineKeys.push("over-margin");
      lineDecorations.push(
        // An attribute, not an inline style: the paragraph's style attribute
        // is rewritten by the page's own line metrics after this is drawn.
        { node: true, from: 0, to: line.contentEnd + 1 - line.blockPos, attrs: { "data-tab-over-margin": "true" } },
      );
    }
    let pen = startPx + segments[0].widthPx;
    for (const [index, tab] of line.tabs.entries()) {
      // On a wrapped paragraph the pen is where this tab stands now, on its
      // own line: where the text before it ends. The tab's own box is not
      // that place: a tab the browser drew too wide (its uniform width, or
      // a width from a pass before) stands at the start of the next line,
      // and a pen read there took it to the line's first stop again -- the
      // tab after a right stop's text went back to that stop, repeated its
      // leader and pushed the rest onto a second line for good (WUA 171).
      if (wraps) {
        const startsLine =
          tab === line.contentStart || view.state.doc.resolve(tab).nodeBefore?.type.name === "hardBreak";
        pen = (view.coordsAtPos(tab, startsLine ? 1 : -1).left - origin) / scale;
      }
      let advance = advances
        ? (advances[index] ?? 0)
        : wordTabAdvanceAt({
            penPx: pen,
            next: segments[index + 1],
            stops: line.stops,
            defaultIntervalTwips,
            positional: line.positional[index],
            line: column,
            stopsEndAtRightEdge: !line.overMargin,
          });
      if (
        wraps &&
        !line.overMargin &&
        !line.positional[index] &&
        pen > line.indentLeftPx + TAB_STOP_REACH_PX &&
        (pen + advance > columnPx - line.indentRightPx + TAB_STOP_REACH_PX ||
          pen + advance + firstWordWidth(line.segments[index + 1]) >
            columnPx - line.indentRightPx + TAB_STOP_REACH_PX)
      ) {
        // No stop is left on this line, or the word after the tab does not
        // fit after its stop: Word starts the next line with the tab, which
        // reaches that line's first stop (Word 16.86, corpus R49-D14 TA12d:
        // a picture after a stop 0.6 pt short of the margin takes the tab to
        // the next line). A box that wide does not fit what is left here, so
        // it is drawn there.
        advance = wordTabAdvanceAt({
          penPx: line.indentLeftPx,
          next: segments[index + 1],
          stops: line.stops,
          defaultIntervalTwips,
          positional: null,
          line: column,
          stopsEndAtRightEdge: !line.overMargin,
        });
      }
      const leader = wordTabLeader(line.stops, pen, line.positional[index]);
      // Keyed by the tab's place in its line, not its document position: an
      // edit above shifts every later position while the tabs themselves
      // are laid out exactly as before, and the mapped decorations already
      // hold them, so nothing is dispatched.
      lineKeys.push(`${index}:${advance}:${leader}`);
      lineDecorations.push({
        node: false,
        from: tab - line.blockPos,
        to: tab + 1 - line.blockPos,
        attrs: {
          class: "sheet-band-story-tab",
          // The tab character fills its box (its one stop is the box's far
          // edge), so the caret after it is drawn after the gap. Zero wide,
          // the caret stood before the gap it had just typed.
          style: `display:inline-block;width:${advance}px;tab-size:${advance}px;white-space:pre`,
          ...(leader !== "none" ? { "data-leader": leader } : {}),
        },
      });
      pen += advance + segments[index + 1].widthPx;
    }
    return done();
  };
  // A line on a sheet outside the ledger's measure horizon (not laid out
  // yet, or left unread as it stands) that has no kept measurement keeps the
  // tabs it has and is measured when the horizon reaches it: measuring it
  // laid that far page out.
  const horizon = sheetsWithinMeasureHorizon(root);
  const installed = horizon ? bandTabLayoutKey.getState(view.state)?.set ?? DecorationSet.empty : null;
  let deferred = false;
  let firstRead = false;
  for (const [lineIndex, line] of tabbedLines(view.state.doc).entries()) {
    const place = memo ? tabLinePlace(view, root, line) : null;
    let result = place ? keptLineTabs(memo!, place, memoAt) : undefined;
    if (!result && horizon && place && !horizon.has(place.sheet)) {
      deferred = true;
      keys.push(`${lineIndex}:deferred`);
      decorations.push(...installed!.find(line.contentStart, line.contentEnd));
      continue;
    }
    if (!result) {
      const block = place?.block ?? view.state.doc.nodeAt(line.blockPos);
      // A paragraph read for the first time (a page the horizon just reached):
      // its settling has passes of its own to come.
      if (rounds && block && !rounds.has(block)) firstRead = true;
      result = settledRound(rounds, block, measureLine(line));
      if (place) keepLineTabs(memo!, place, memoAt, result);
    }
    for (const key of result.keys) keys.push(`${lineIndex}:${key}`);
    for (const item of result.decorations) {
      const from = line.blockPos + item.from, to = line.blockPos + item.to;
      decorations.push(item.node ? Decoration.node(from, to, item.attrs) : Decoration.inline(from, to, item.attrs));
    }
    if (result.wrapped) wrapped = true;
  }
  return { set: DecorationSet.create(view.state.doc, decorations), key: keys.join("|"), wrapped, deferred, firstRead };
}

export const BandTabLayout = Extension.create({
  name: "bandTabLayout",
  addProseMirrorPlugins() {
    return [
      new Plugin<BandTabState>({
        key: bandTabLayoutKey,
        state: {
          init: () => ({ set: DecorationSet.empty, edits: 0 }),
          apply(tr, state) {
            const next = tr.getMeta(bandTabLayoutKey) as DecorationSet | undefined;
            // The pagination moving blocks between sheets (its own steps) is no
            // edit: it changes no paragraph's text, stops or width.
            const edits = tr.docChanged && tr.getMeta(PAGINATION_MOVE_META) !== true ? state.edits + 1 : state.edits;
            return { set: next ?? state.set.map(tr.mapping, tr.doc), edits };
          },
        },
        props: {
          decorations: (state) => bandTabLayoutKey.getState(state)?.set,
        },
        view(editorView) {
          const win = editorView.dom.ownerDocument.defaultView;
          let frame: number | null = null;
          let destroyed = false;
          let installed = "";
          // Passes since the last change from outside (an edit, a resize, a
          // font): a wrapped paragraph's tabs settle over a few.
          let settlePasses = 0;
          let ownDispatchAt = Number.NEGATIVE_INFINITY;
          const now = () => win?.performance?.now() ?? 0;
          // After a face arrives the tabs are measured from no sizing at all:
          // the next sizes depend on where the text stood, and measured over
          // the sizes taken in the fallback a tabbed line could settle where
          // a document opened with its faces in never does (a CJK index's
          
          // with nothing painted between.
          let fromClearSlate = false;
          // What each paragraph's tabs measured since the inputs last changed.
          let rounds: TabRounds = new WeakMap();
          // The last measure left lines outside the horizon for later.
          let lastDeferred = false;
          // Until the tabs measured after a face arrived have settled, Print
          // waits (layoutPending, printReadiness).
          const run = () => {
            try {
              measureNow();
            } finally {
              if (frame === null) clearLayoutPending(editorView.dom, "band-tabs");
            }
          };
          const measureNow = () => {
            frame = null;
            if (destroyed) return;
            if (fromClearSlate) {
              fromClearSlate = false;
              if (bandTabLayoutKey.getState(editorView.state)?.set !== DecorationSet.empty) {
                editorView.dispatch(
                  editorView.state.tr.setMeta(bandTabLayoutKey, DecorationSet.empty).setMeta("addToHistory", false),
                );
              }
              installed = "";
            }
            let next: ReturnType<typeof measure> = null;
            try {
              next = measure(editorView, rounds);
            } catch {
              // A DOM that cannot be measured (not laid out) keeps the
              // uniform tab width; nothing else depends on this.
              return;
            }
            lastDeferred = Boolean(next?.deferred);
            if (!next || next.key === installed) return;
            installed = next.key;
            ownDispatchAt = now();
            editorView.dispatch(
              editorView.state.tr.setMeta(bandTabLayoutKey, next.set).setMeta("addToHistory", false),
            );
            // A tab just sized on a wrapped paragraph moved what follows it:
            // read the tabs again where they stand now.
            if (next.wrapped) {
              if (next.firstRead) settlePasses = 0;
              settleAgain();
            }
          };
          const settleAgain = () => {
            if (settlePasses >= MAX_SETTLE_PASSES) return;
            settlePasses += 1;
            settle();
          };
          const settle = () => {
            if (frame !== null || !win) return;
            frame = win.requestAnimationFrame(run);
          };
          const schedule = () => {
            settlePasses = 0;
            settle();
          };
          schedule();
          // Print lays the tabs out now, from scratch, with the faces that
          // are in, before it copies the pages (layoutPending.settleLayoutNow):
          // the settling below, run to its end in one go.
          const unregisterSettler = registerLayoutSettler(editorView.dom, "band-tabs", () => {
            if (destroyed) return;
            if (frame !== null) win?.cancelAnimationFrame(frame);
            frame = null;
            if (editorView.dom instanceof HTMLElement) tabMeasureMemos.get(editorView.dom)?.clear();
            fromClearSlate = true;
            settlePasses = 0;
            for (let pass = 0; pass <= MAX_SETTLE_PASSES; pass += 1) {
              measureNow();
              // A wrapped line asked for another look: taken now, not next frame.
              if (frame === null) break;
              win?.cancelAnimationFrame(frame);
              frame = null;
            }
            clearLayoutPending(editorView.dom, "band-tabs");
          });
          const fonts = editorView.dom.ownerDocument.fonts;
          // A face arriving changes text widths with no DOM change: every
          // kept measurement is forgotten.
          const fontsChanged = () => {
            if (editorView.dom instanceof HTMLElement) tabMeasureMemos.get(editorView.dom)?.clear();
            rounds = new WeakMap();
            fromClearSlate = true;
            markLayoutPending(editorView.dom, "band-tabs");
            schedule();
          };
          // Read once for every caller of this task, a microtask later: a
          // read of `fonts.ready` lays the page out in Blink (font-face-news).
          const cancelFontsReady = afterFontsReady(editorView.dom.ownerDocument, fontsChanged);
          // A face that arrives later (an Office font asset loaded for this
          // run) changes the text's width without resizing the story. A
          // subset of some characters changes only the sheets drawing them
          // (the page ledger's rule): a Chinese book fetches subsets all the
          // while it opens, and each measured every tabbed line again (1.2 s
          // of a 271-page book's reopening).
          const facesLoaded = (event: Event) => {
            const root = editorView.dom;
            const memo = root instanceof HTMLElement ? tabMeasureMemos.get(root) : undefined;
            const faces = newFontFacesOf(event) ??
              (event as Event & { fontfaces?: ReadonlyArray<{ family?: string; unicodeRange?: string }> }).fontfaces;
            const used = faces && faces.length > 0 && root instanceof HTMLElement ? facesUsedBy(faces, editorFontFamilies(root)) : null;
            if (used && used.length === 0) return;
            const ranges = loadedFaceRanges(used ?? faces);
            if (!memo || !ranges) { fontsChanged(); return; }
            let touched = false;
            for (const sheet of sheetsOf(root)) {
              if (!textDrawsCodePointIn(sheet.textContent ?? "", ranges)) continue;
              memo.forget(sheet);
              touched = true;
            }
            if (!touched) return;
            markLayoutPending(root, "band-tabs");
            schedule();
          };
          fonts?.addEventListener?.("loadingdone", facesLoaded);
          // The horizon reached sheets a measure left for later.
          const horizonWidened = () => {
            if (lastDeferred) schedule();
          };
          editorView.dom.addEventListener(MEASURE_HORIZON_EVENT, horizonWidened);
          // The story's size changing right after the tabs were sized is that
          // sizing (a line more, a line less): part of settling, not a new
          // layout to start over on, or a paragraph whose tabs never hold
          // would be measured for ever.
          // Only a new width is a new layout: the story growing or shrinking by
          // the sheets the pages took or gave back is the pages settling.
          let observedWidth: number | null = null;
          const resized = (entries: readonly ResizeObserverEntry[]) => {
            const width = entries[entries.length - 1]?.contentRect.width ?? null;
            const widthChanged = width !== null && observedWidth !== null && Math.abs(width - observedWidth) > 0.5;
            if (width !== null) observedWidth = width;
            if (widthChanged) rounds = new WeakMap();
            if (widthChanged && now() - ownDispatchAt >= 250) schedule();
            else settleAgain();
          };
          const observer = win && "ResizeObserver" in win ? new win.ResizeObserver(resized) : null;
          observer?.observe(editorView.dom);
          return {
            update(view, previous) {
              if (view.state.doc.eq(previous.doc)) return;
              // A block moved to another sheet is read again, within the passes
              // its settling has left: a page break moved by these very tabs
              // started the count over, and tabs whose widths make their line
              // wrap, and unwrapped come out wider again, were measured for as
              // long as the pages kept moving with them (R350-D07: interlinear
              // glosses one line, then two, then one, every few frames; 157 to
              // 160 sheets, never settled, a tab burning a core).
              if (bandTabLayoutKey.getState(view.state)?.edits !== bandTabLayoutKey.getState(previous)?.edits) schedule();
              else settleAgain();
            },
            destroy() {
              destroyed = true;
              if (frame !== null) win?.cancelAnimationFrame(frame);
              clearLayoutPending(editorView.dom, "band-tabs");
              unregisterSettler();
              observer?.disconnect();
              cancelFontsReady();
              fonts?.removeEventListener?.("loadingdone", facesLoaded);
              editorView.dom.removeEventListener(MEASURE_HORIZON_EVENT, horizonWidened);
            },
          };
        },
      }),
    ];
  },
});
