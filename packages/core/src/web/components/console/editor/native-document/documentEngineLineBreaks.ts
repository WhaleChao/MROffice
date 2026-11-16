/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Docs lines break where the text engine breaks them (lib/text-engine,
 * PLAN-phase2.md; flag `docs`, `?textEngine=on`).
 *
 * The engine fills each paragraph's lines the way Word does, from the
 * served font files (docs-lines.ts). The browser still draws the text; where
 * its own line breaker would break elsewhere -- its glyph advances a hair
 * wider or narrower than the file's, a justified line Word keeps by
 * narrowing its spaces -- this plugin spreads the difference over that line's
 * letter spacing, so the drawn line is as wide as the computed one and the
 * browser's breaker lands on the engine's break. A line it still breaks
 * elsewhere is corrected again, whatever the widths said, a few rounds at
 * most. Paint and line breaking only: nothing is saved, and with the flag off
 * the plugin does nothing.
 *
 * Only the paragraphs a transaction touched are measured again (typing
 * re-reads one paragraph); a font landing re-reads them all.
 */
import { noteEngineDeclinedBlock, takeEngineDeclinedNews } from "./engineDeclinedBlocks";
import { drawnLinesWidth, engineLinesKeptOn, forgetDrawnLines, noteDrawnLines, retireSeededLines, seededLinesFor, spansAt } from "./engineLinesKept";
import { Extension } from "@tiptap/core";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { Plugin, PluginKey, type Transaction } from "@tiptap/pm/state";
import { Decoration, DecorationSet, type EditorView } from "@tiptap/pm/view";

import {
  auditDocsParagraph,
  DOCS_ENGINE_BASE_SPACING_ATTR,
  DOCS_ENGINE_BASE_WORD_SPACING_ATTR,
  DOCS_ENGINE_LINE_ATTR,
  DOCS_ENGINE_TAB_BREAK_CLASS,
  docsBrowserAgreesQuickly,
  docsBrowserLineStarts,
  docsParagraphBesideFloat,
  type DocsFloatCache,
  docsDrawnWidths,
  docsEngineLineStarts,
  docsLastLineOverhangs,
  engineBreakStrutOff,
  docsLineOpensWithSpace,
  normalizeStarts,
  planDocsLineCorrections,
  readDocsParagraph,
  type DocsLayoutRules,
  type DocsParagraph,
} from "../../../../lib/text-engine/docs-lines";
import type { KinsokuList } from "../../../../lib/text-engine/kinsoku";
import { layoutDocsBlockLines } from "../../../../lib/text-engine/docs-block-layout";
import { textEngineEnabled } from "../../../../lib/text-engine/flags";
import { textEngineOn } from "../../../../lib/text-engine/page-adapter";
import { docsBlockLineRule, noteEngineBlockLines } from "./paged/ledger/ledgerEngineLines";
import { normalizeWordCompatibilityMode } from "./documentOverflowPunctuation";
import { wordSetsIdeographsAsLetters } from "./wordFarEastDefaults";
import { touchedTextblockRanges } from "./touchedTextblocks";
import { isImeCompositionActive, onImeCompositionSettled } from "./imeCompositionGuard";
import { DOCUMENT_ENGINE_LINES_PENDING, LAYOUT_METRICS_FLUSH_EVENT, writeDocumentLayoutStatus } from "./documentLayoutStatus";
import { projectionFontReach, sheetRangesOf } from "./projectionFontNews";
import { recordProjectedHorizon, sheetsWithinMeasureHorizon } from "./paged/measureHorizon";
import { createEngineLineCache, familyNews, type EngineLineCache, type EngineLineKey, type FamilyNews } from "./engineLineCache";
import { PAGE_BLOCK_NODE } from "./paged/paginationProjectionKeys";
import { fastRangeMappingOn, mapDecorationsFast, mapRangesFast, mergeRanges } from "./fastRangeMapping";
import { ENGINE_WORD_EXTRA_VAR } from "./balancedSpaces";
import { textEngineAnswerEpoch, textEngineEpoch } from "../../../../lib/text-engine/measure";
import { msSinceTyping, TYPING_IDLE_MS } from "../../../../lib/typing-session";

export const documentEngineLineBreaksKey = new PluginKey<EngineLinesState>("documentEngineLineBreaks");
const META = "documentEngineLineBreaks";
const BLOCKS_PER_SLICE = 30;
/** Rounds a paragraph is corrected again while the browser still disagrees. */
const MAX_ROUNDS = 4;
/** This projection's coverage record in measureHorizon. */
const ENGINE_LINES_HORIZON_KEY = "engine-lines";
/** How long the page ledger waits for face files the engine still loads. */
const FACE_HOLD_MS = 4000;
/** How long, and how often, paragraphs read while a face was still loading
 * are looked at for the engine's answer (the ledger is not held for this). */
const FACE_WATCH_MS = 60_000;
const FACE_WATCH_STEP_MS = 250;












const TYPING_PAUSE_MS = 2 * TYPING_IDLE_MS;
const CJK = /[぀-ヿ㐀-鿿가-힯＀-￯]/;

/** Whether `transaction` publishes the engine's line corrections only. */
export function isDocumentEngineLineBreaksTransaction(transaction: { getMeta(key: string): unknown }): boolean {
  return transaction.getMeta(META) !== undefined;
}

export interface Span {
  from: number;
  to: number;
  /** Extra letter spacing, px. */
  extra: number;
  /** The run's own letter spacing, px. */
  base: number;
  /** A line break before a tab that opens a line (`DocsLineCorrection.tabBreak`):
   * `from === to`, at the tab, and no spacing. */
  tabBreak?: true;
  /** Extra word spacing (a kept word's narrowed spaces), px. */
  wordExtra: number;
  /** The run's own word spacing, px. */
  wordBase: number;
}

interface EngineLinesState {
  set: DecorationSet;
  dirty: Array<[number, number]> | "all";
  /** Paragraphs the browser still broke elsewhere after their correction:
   * measured again, lines forced, a few rounds at most. */
  retry: Array<[number, number]>;
}

interface EngineLinesMeta {
  spans: Span[];
  /** Blocks measured: their old spans are replaced. */
  ranges: Array<[number, number]>;
  retry: Array<[number, number]>;
  /** Dirty ranges left for later: sheets past the page ledger's measure
   * horizon (a long document laid out progressively). */
  dirty: Array<[number, number]>;
}

function tabBreakElement(): HTMLElement {
  const br = document.createElement("br");
  br.className = DOCS_ENGINE_TAB_BREAK_CLASS;
  return br;
}

/**
 * The tabs of a paragraph the browser breaks that hang at a line's end while
 * what follows them opens the next line. Word takes such a tab to the next
 * line with what follows it (docs-lines: "a tab stays with what follows
 * it"); the browser lets it hang, and a reference list indented by typed
 * tabs set its continuation lines at the margin (corpus R283-D12, Times New
 * Roman installed on the Mac: the engine leaves those paragraphs to the
 * browser). A tab a break of ours already opens a line with keeps it.
 */
export function hangingTabBreaks(view: EditorView, element: HTMLElement): Span[] {
  const doc = element.ownerDocument;
  const texts: Text[] = [];
  const walker = doc.createTreeWalker(element, NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => (node.parentElement?.closest('[contenteditable="false"]') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let node = walker.nextNode(); node; node = walker.nextNode()) texts.push(node as Text);
  const range = doc.createRange();
  const boxOf = (text: Text, index: number): DOMRect | null => {
    range.setStart(text, index);
    range.setEnd(text, index + 1);
    const rects = Array.from(range.getClientRects()).filter((rect) => rect.height > 0);
    return rects[0] ?? null;
  };
  const spans: Span[] = [];
  for (let t = 0; t < texts.length; t += 1) {
    const data = texts[t]!.data;
    for (let index = data.indexOf("\t"); index >= 0; index = data.indexOf("\t", index + 1)) {
      const tab = boxOf(texts[t]!, index);
      if (!tab) continue;
      // Already opened by our break: kept.
      const before = texts[t]!.previousSibling;
      const opened = index === 0 && before instanceof HTMLElement && before.classList.contains(DOCS_ENGINE_TAB_BREAK_CLASS);
      let hangs = opened;
      if (!opened) {
        // The first character after the tab that draws something.
        let found: DOMRect | null = null;
        for (let u = t, at = index + 1; u < texts.length && !found; u += 1, at = 0) {
          const text = texts[u]!.data;
          for (let i = at; i < text.length; i += 1) {
            if (/\s/.test(text[i]!)) continue;
            found = boxOf(texts[u]!, i);
            break;
          }
        }
        hangs = found !== null && found.top > tab.top + tab.height / 2;
      }
      if (!hangs) continue;
      try {
        const at = view.posAtDOM(texts[t]!, index);
        spans.push({ from: at, to: at, extra: 0, base: 0, wordExtra: 0, wordBase: 0, tabBreak: true });
      } catch { /* a node outside the editor's content */ }
    }
  }
  return spans;
}

function decorationsFor(doc: ProseMirrorNode, spans: readonly Span[]): Decoration[] {
  return spans.map((span) => {
    // The break ends a line and owns no line box of its own: it stands in
    // the paragraph's wrappers, which take no line height where the
    // paragraph's leaves own it (documentAutoLineSafety). Taken for a
    
    // of its paragraph: an 8 pt cell under an 11 pt Normal drew 10.5 pt lines
    // for Word's 9.77 (tables/probes/narrow, corpus R417-D03's "ID" column).
    if (span.tabBreak) {
      return Decoration.widget(span.from, tabBreakElement, {
        side: -1, key: DOCS_ENGINE_TAB_BREAK_CLASS, engineLine: true, ignoreSelection: true,
        ...(engineBreakStrutOff() ? {} : { zeroLayoutFormattingMark: true }),
      });
    }
    const total = Number((span.base + span.extra).toFixed(4));
    // The extra alone too: a balanced space's own width on the same text adds it (balancedSpaces.ts).
    const words = span.wordExtra
      ? `;word-spacing:${Number((span.wordBase + span.wordExtra).toFixed(4))}px;${ENGINE_WORD_EXTRA_VAR}:${Number(span.wordExtra.toFixed(4))}px`
      : "";
    return Decoration.inline(
      span.from,
      span.to,
      {
        [DOCS_ENGINE_LINE_ATTR]: String(Number(span.extra.toFixed(4))),
        [DOCS_ENGINE_BASE_SPACING_ATTR]: String(span.base),
        [DOCS_ENGINE_BASE_WORD_SPACING_ATTR]: String(span.wordBase),
        style: `letter-spacing:${total}px${words}`,
      },
      { engineLine: true, extra: span.extra, base: span.base },
    );
  });
}

function rulesFor(view: EditorView, node: ProseMirrorNode): DocsLayoutRules {
  const kinsoku: KinsokuList | null = CJK.test(node.textContent) ? "zh-CN" : null;
  const attrs = view.state.doc.attrs;
  return { compatibilityMode: normalizeWordCompatibilityMode(attrs.wordCompatibilityMode), kinsoku, ideographsAsLetters: kinsoku !== null && wordSetsIdeographsAsLetters(attrs) };
}

/** Pieces' natural drawn widths; a justified paragraph is read left-aligned
 * for the moment (its justification stretches CJK glyph boxes), with the
 * editor's DOM observer paused so the style flip is not taken for an edit. */
function naturalWidths(view: EditorView, element: HTMLElement, paragraph: DocsParagraph): number[] {
  const align = paragraph.align;
  if (align !== "justify" && align !== "distribute" && !align.startsWith("-")) return docsDrawnWidths(paragraph);
  const observer = (view as EditorView & { domObserver?: { stop(): void; start(): void } }).domObserver;
  const before = element.style.getPropertyValue("text-align");
  const beforePriority = element.style.getPropertyPriority("text-align");
  const beforeLast = element.style.getPropertyValue("text-align-last");
  const beforeLastPriority = element.style.getPropertyPriority("text-align-last");
  observer?.stop();
  try {
    element.style.setProperty("text-align", "left", "important");
    element.style.setProperty("text-align-last", "auto", "important");
    return docsDrawnWidths(paragraph);
  } finally {
    if (before) element.style.setProperty("text-align", before, beforePriority);
    else element.style.removeProperty("text-align");
    if (beforeLast) element.style.setProperty("text-align-last", beforeLast, beforeLastPriority);
    else element.style.removeProperty("text-align-last");
    observer?.start();
  }
}

/** Same lines? Starts compared on the drawn sequence. */
function sameStarts(a: readonly number[], b: readonly number[]): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index]);
}

/**
 * The spans one paragraph needs now. `round` > 0: the browser disagreed
 * after the last correction; the lines from the first disagreement are
 * corrected whatever their widths say.
 */
/** Where a line's first piece stands in its block's content. */
function lineContentOffset(view: EditorView, paragraph: DocsParagraph, start: number, pos: number): number | null {
  const piece = paragraph.pieces[start];
  if (!piece?.node || piece.offset === undefined) return null;
  try {
    return view.posAtDOM(piece.node, piece.offset) - (pos + 1);
  } catch {
    return null;
  }
}

/**
 * The text blocks of `doc` among `waiting` (paragraphs a pass found waiting
 * for a face file), with their positions, wherever they stand: read again once
 * the faces move, whatever sheets the page ledger's horizon holds then.
 */
export function blocksWaitingForFaces(doc: ProseMirrorNode, waiting: ReadonlySet<ProseMirrorNode>): Array<{ node: ProseMirrorNode; pos: number }> {
  const found: Array<{ node: ProseMirrorNode; pos: number }> = [];
  if (!waiting.size) return found;
  let left = waiting.size;
  doc.descendants((node, pos) => {
    if (!left) return false;
    if (!node.isTextblock) return true;
    if (waiting.has(node)) {
      found.push({ node, pos });
      left -= 1;
    }
    return false;
  });
  return found;
}

/** Counts for probes (`window.__docsEngineLineCache`). */
interface EngineLineCacheStats { measured: number; reused: number; kept: number; missed?: Record<string, number>; seeded?: number }

/** What a block's lines depend on besides itself (engineLineCache). */
function lineKeyOf(view: EditorView, element: HTMLElement, pos: number): EngineLineKey {
  const parent = view.state.doc.resolve(pos).parent;
  return {
    width: element.getBoundingClientRect().width,
    // Lines are kept only from answers the engine gave (a paragraph still
    // waiting for a face is not kept), and a given answer moves only with
    // the faces declared (textEngineAnswerEpoch); the faces a block's
    // families gained since are its entry's own check (familyNews). Keyed on
    // every face file loading, R286-D01's cold opening read 12,565 blocks
    // again for that alone, of 34,449; with this, 25,620 readings in all
    
    // every face file again.
    epoch: (view.dom.ownerDocument.defaultView as (Window & { __docsEngineLinesAnswerEpoch?: boolean }) | null)?.__docsEngineLinesAnswerEpoch === false
      ? textEngineEpoch() : textEngineAnswerEpoch(),
    docAttrs: view.state.doc.attrs,
    parent: parent.type.name === PAGE_BLOCK_NODE ? PAGE_BLOCK_NODE : parent.attrs,
  };
}

type EngineLinesCache = { lines: EngineLineCache; news: FamilyNews; stats: EngineLineCacheStats | null };

/** What a reading drew, for the layout a row keeps (engineLinesKept). */
interface DrawnReading { read: boolean; width: number | null; declined: boolean }

/** A paragraph's spans now (readBlockLines), what it drew noted for a kept
 * layout, or taken back from the one restored. */
function measureBlock(
  view: EditorView, node: ProseMirrorNode, pos: number, round: number, floats: DocsFloatCache,
  cache?: EngineLinesCache,
): { spans: Span[]; agrees: boolean } | null | undefined {
  if (!engineLinesKeptOn(view.dom.ownerDocument.defaultView)) return readBlockLines(view, node, pos, round, floats, cache, null);
  const drew: DrawnReading = { read: false, width: null, declined: false };
  const result = readBlockLines(view, node, pos, round, floats, cache, drew);
  if (result === undefined || !drew.read) forgetDrawnLines(node);
  else noteDrawnLines(node, pos, drew.width ?? drawnLinesWidth(node), result?.spans ?? [], drew.declined);
  return result;
}

function readBlockLines(
  view: EditorView, node: ProseMirrorNode, pos: number, round: number, floats: DocsFloatCache,
  cache: EngineLinesCache | undefined, drew: DrawnReading | null,
): { spans: Span[]; agrees: boolean } | null | undefined {
  const element = view.nodeDOM(pos);
  if (!(element instanceof HTMLElement) || !element.isConnected) return null;
  // A cell of an auto-layout table sizes its column from its content's
  // widths, which a correction would change: the browser keeps it.
  const table = element.closest("table");
  if (table && element.ownerDocument.defaultView?.getComputedStyle(table).tableLayout !== "fixed") {
    noteEngineDeclinedBlock(node);
    return null;
  }
  // A block measured as it is now (moved to another sheet by the page
  // ledger, unchanged) takes its corrections back without a reading.
  const key = cache && round === 0 ? lineKeyOf(view, element, pos) : null;
  if (drew) {
    drew.read = true;
    drew.width = key?.width ?? null;
    // Restored pages draw what they drew when they were kept.
    // A paragraph with corrections beside a float is read again: a float a
    // layout drew late (a text box placed after its first pass) was not there
    // when its corrections were kept, and the browser wraps beside one.
    const seededKept = key ? seededLinesFor(node, key.width) : null;
    const kept = seededKept && (!seededKept.spans.length || !docsParagraphBesideFloat(element, floats)) ? seededKept : null;
    if (kept) {
      if (cache?.stats) cache.stats.seeded = (cache.stats.seeded ?? 0) + 1;
      drew.declined = kept.declined;
      if (kept.declined) noteEngineDeclinedBlock(node);
      return kept.declined && !kept.spans.length ? null : { spans: spansAt<Span>(kept, pos), agrees: true };
    }
  }
  if (cache && key) {
    const kept = cache.lines.lookup<Span>(node, key, pos, cache.news);
    if (kept) {
      if (cache.stats) cache.stats.reused += 1;
      return { spans: kept, agrees: true };
    }
  }
  if (cache?.stats) {
    cache.stats.measured += 1;
    const reason = !key ? "retry" : cache.lines.missReason(node, key, cache.news);
    const missed = (cache.stats.missed ??= {});
    missed[reason] = (missed[reason] ?? 0) + 1;
  }
  const keep = (spans: Span[]) => {
    if (!cache || !key) return;
    const families = new Set<string>();
    for (const piece of paragraph.pieces) {
      const stack = (piece as { font?: { stack?: string } }).font?.stack;
      if (stack) for (const family of stack.split(",")) families.add(family);
    }
    cache.lines.store(node, key, pos, spans, [...families], cache.news);
    if (cache.stats) cache.stats.kept += 1;
  };
  const paragraph = readDocsParagraph(element);
  // undefined: a face is still loading (the ledger may wait for it).
  if (paragraph.pending) return undefined;
  if (paragraph.unsupported || !paragraph.pieces.length) {
    // Left to the browser: its lines take the Word rules the browser's
    // projections keep (justifiedSpaceShrink), and a tab hanging at a line's
    // end goes to the next line with what follows it (hangingTabBreaks).
    if (paragraph.unsupported) {
      noteEngineDeclinedBlock(node);
      if (drew) drew.declined = true;
      if (element.textContent?.includes("\t")) {
        const spans = hangingTabBreaks(view, element);
        if (spans.length) return { spans, agrees: true };
      }
    }
    return null;
  }
  const rules = rulesFor(view, node);
  const raw = docsEngineLineStarts(paragraph, rules);
  const engine = normalizeStarts(paragraph, raw);
  // The page ledger reads these lines (a list's first line rises to its
  // marker, which the model does not state: the page keeps those).
  if (!element.closest("li")) {
    const rule = docsBlockLineRule(node, element);
    const layout = layoutDocsBlockLines(paragraph, engine, rule, null);
    const offsets = layout.status === "ok" && layout.lines.length > 1 ? layout.lines.map((line) => lineContentOffset(view, paragraph, line.start, pos)) : [];
    noteEngineBlockLines(node, layout, rule, offsets, paragraph.chrome);
  }
  // Most paragraphs already break where the engine does: shown cheaply.
  const current = paragraph.pieces.some((piece) => piece.correction);
  // A line the engine opens with a tab is drawn so by a break before it
  // (planDocsLineCorrections' tabBreak), which no spacing shows: the browser
  // reads that tab's box at the next line's start whether the break is there
  // or the tab hangs at the end of the line before (R283-D12 on: the links
  // stayed at the margin, the lines read as agreeing). Such a paragraph is
  // always planned, so its break is drawn and kept.
  // So is a line the engine opens with a space (mode 14, a cell narrower
  // than a letter): the browser reads no space's box at all.
  // And a paragraph whose last line is one letter wider than the measure:
  // that letter is narrowed so the closing widgets stay on its line.
  const tabOpened = engine.some((start, line) => line > 0 && paragraph.pieces[start]?.kind === "tab") ||
    raw.some((_start, line) => docsLineOpensWithSpace(paragraph, raw, line)) || docsLastLineOverhangs(paragraph, raw);
  if (!current && !tabOpened && docsBrowserAgreesQuickly(element, paragraph, engine)) {
    keep([]);
    return { spans: [], agrees: true };
  }
  const browser = docsBrowserLineStarts(paragraph);
  const agrees = sameStarts(engine, browser);
  // Lines a float narrows are the browser's to wrap (the engine fills whole
  // lines); corrections already there stay out of it too.
  if (!agrees && docsParagraphBesideFloat(element, floats)) return { spans: [], agrees: true };
  const force = new Set<number>();
  if (!agrees && round > 0) {
    let first = 0;
    while (first < engine.length && engine[first] === browser[first]) first += 1;
    // The line before the first disagreement decides where it starts.
    for (let line = Math.max(0, first - 1); line <= Math.min(raw.length - 1, first + 1); line += 1) force.add(line);
  }
  if (agrees && !current && !tabOpened) {
    keep([]);
    return { spans: [], agrees };
  }
  const plan = planDocsLineCorrections(paragraph, raw, naturalWidths(view, element, paragraph), force);
  const spans: Span[] = [];
  for (const line of plan) {
    // One span per text node stretch of the line.
    let at = line.from;
    while (at < line.to) {
      const piece = paragraph.pieces[at];
      if ((piece.kind !== "text" && piece.kind !== "space") || !piece.node || piece.offset === undefined) { at += 1; continue; }
      const textNode = piece.node;
      const startOffset = piece.offset;
      let endOffset = startOffset + piece.text.length;
      let next = at + 1;
      // A stretch also ends where the narrowed spaces (after the line's
      // last tab) begin.
      const spaced = at >= line.spacesFrom;
      while (next < line.to) {
        const later = paragraph.pieces[next];
        if (later.node !== textNode || (later.kind !== "text" && later.kind !== "space") || later.offset !== endOffset) break;
        if ((next >= line.spacesFrom) !== spaced) break;
        endOffset += later.text.length;
        next += 1;
      }
      try {
        const from = view.posAtDOM(textNode, startOffset);
        const to = view.posAtDOM(textNode, endOffset);
        if (to > from) {
          spans.push({
            from, to, extra: line.letterSpacing, base: piece.letterSpacing ?? 0,
            wordExtra: spaced ? line.wordSpacing : 0, wordBase: piece.wordSpacing ?? 0,
          });
        }
      } catch { /* a node outside the editor's content */ }
      at = next;
    }
    const tab = line.tabBreak === null ? undefined : paragraph.pieces[line.tabBreak];
    if (tab?.node && tab.offset !== undefined) {
      try {
        const at = view.posAtDOM(tab.node, tab.offset);
        spans.push({ from: at, to: at, extra: 0, base: 0, wordExtra: 0, wordBase: 0, tabBreak: true });
      } catch { /* a node outside the editor's content */ }
    }
  }
  if (agrees) keep(spans);
  return { spans, agrees };
}

export const DocumentEngineLineBreaks = Extension.create({
  name: "documentEngineLineBreaks",
  addProseMirrorPlugins() {
    // Off (the default): no state, no transaction cost.
    if (!textEngineEnabled("docs")) return [];
    const mapRanges = (ranges: Array<[number, number]>, tr: Transaction) =>
      ranges.map(([from, to]): [number, number] => [tr.mapping.map(from, -1), tr.mapping.map(to, 1)]);
    return [
      new Plugin<EngineLinesState>({
        key: documentEngineLineBreaksKey,
        state: {
          init: () => ({ set: DecorationSet.empty, dirty: "all", retry: [] }),
          apply: (tr: Transaction, previous): EngineLinesState => {
            const meta = tr.getMeta(META) as EngineLinesMeta | undefined;
            if (meta) {
              let set = previous.set;
              for (const [from, to] of meta.ranges) set = set.remove(set.find(from, to));
              const added = decorationsFor(tr.doc, meta.spans);
              return { set: added.length ? set.add(tr.doc, added) : set, dirty: meta.dirty, retry: meta.retry };
            }
            if (!tr.docChanged) return previous;
            if (tr.getMeta("addToHistory") !== false) retireSeededLines();
            // An edited paragraph keeps its mapped corrections until it is
            // measured again: no frame of browser lines in between.
            const changed = touchedTextblockRanges(tr);
            // At the cost of what the transaction touches (fastRangeMapping).
            if (fastRangeMappingOn()) {
              return {
                set: mapDecorationsFast(previous.set, tr),
                dirty: previous.dirty === "all" ? "all" : mapRangesFast(previous.dirty, tr, changed),
                retry: previous.retry.length ? mapRangesFast(previous.retry, tr, [], false) : previous.retry,
              };
            }
            const set = previous.set.map(tr.mapping, tr.doc);
            const dirty = previous.dirty === "all" ? "all" : [...mapRanges(previous.dirty, tr), ...changed];
            return { set, dirty, retry: mapRanges(previous.retry, tr) };
          },
        },
        props: {
          decorations: (state) => documentEngineLineBreaksKey.getState(state)?.set,
        },
        view(view) {
          const win = view.dom.ownerDocument.defaultView;
          if (!win || !textEngineOn("docs")) return {};
          // Probes compare a paragraph's engine lines with the drawn ones.
          (win as unknown as Record<string, unknown>).__mraiDocsLineAudit = (element: HTMLElement) => {
            const pos = view.posAtDOM(element, 0) - 1;
            const node = view.state.doc.nodeAt(pos);
            return node ? auditDocsParagraph(element, rulesFor(view, node), true) : null;
          };
          let frame: number | null = null;
          let idle: number | null = null;
          let waitTimer: number | null = null;
          let full = true;
          /** Sheets a face arrived for (projectionFontNews): read with the dirty ones. */
          const fontSheets = new Set<Element>();
          /** What a pass read a frame at a time took to read (every paragraph,
           * the sheets of an arrived face): given back if it is cut short. */
          let inFlight: (() => void) | null = null;
          const rounds = new Map<number, number>();
          const lineCache = createEngineLineCache();
          const cacheStats = (): EngineLineCacheStats | null => {
            const holder = win as Window & { __docsEngineLineCache?: EngineLineCacheStats };
            return holder.__docsEngineLineCache ??= { measured: 0, reused: 0, kept: 0 };
          };
          /** Since when paragraphs wait for a face file (the page ledger
           * waits with them, at most FACE_HOLD_MS). */
          let facesPendingSince: number | null = null;
          // Paragraphs a pass found waiting for a face file, by node. They are
          // read again once the engine's faces move, wherever they are: the
          // pass that answers runs long after the page ledger laid them out,
          // its horizon on other sheets, and a face file read by the engine
          // alone (its metrics, the browser's face long loaded) brings the
          // ledger no font news to widen it. R284-D30's Malgun Gothic Latin
          // 8 s late: page 7's cell paragraph kept the browser's lines for
          // good ("동/일한" where the engine squeezes "동일한").
          // `window.__docsEnginePendingAnywhereOff = true`: horizon only, as before.
          const waitingForFaces = new Set<ProseMirrorNode>();
          let facesLanded = false;
          const pendingAnywhereOff = () =>
            (win as Window & { __docsEnginePendingAnywhereOff?: boolean }).__docsEnginePendingAnywhereOff === true;
          // The page ledger lays pages out from the drawn lines: while lines
          // are about to be corrected it waits (and asks for them at once,
          // LAYOUT_METRICS_FLUSH_EVENT), so its cuts are made on the engine's
          // lines and never undone by a correction after them.
          const markPending = (pending: boolean) => writeDocumentLayoutStatus(view.dom, DOCUMENT_ENGINE_LINES_PENDING, pending ? "true" : null, view);
          const cancelScheduled = () => {
            if (frame !== null) { win.cancelAnimationFrame(frame); frame = null; }
            // A pass under way, cut short by the ledger's flush: what it took
            // to read is read by the flush instead. Taken and dropped, the
            // paragraphs waiting for a face that had just answered stayed the
            // browser's lines through the pages' release and were corrected
            
            if (inFlight) { inFlight(); inFlight = null; }
            if (idle !== null) {
              const cancel = (win as Window & { cancelIdleCallback?: (handle: number) => void }).cancelIdleCallback;
              if (cancel) cancel.call(win, idle);
              else win.clearTimeout(idle);
              idle = null;
            }
          };
          // The page ledger's flush usually measures first (in the frame it
          // lays pages out in, so its pass needs no second one); a pass of
          // our own runs only when no flush came for two frames (nothing to
          // lay out: a face landing on a settled document, a non-paged view).
          let flushes = 0;
          // An editor not in the page has no lines to read and none to wait
          // for: it says nothing is pending (a document printed from an
          // editor that was never attached waited on this flag for ever) and
          // is looked at again until it is attached, without touching the
          // flag meanwhile (print waits for the page to hold still).
          const waitUntilAttached = () => {
            if (waitTimer !== null) return;
            waitTimer = win.setTimeout(() => {
              waitTimer = null;
              if (view.dom.isConnected) schedule();
              else waitUntilAttached();
            }, 250);
          };
          const schedule = () => {
            if (!view.dom.isConnected) {
              markPending(false);
              waitUntilAttached();
              return;
            }
            markPending(true);
            if (frame !== null) return;
            const seen = flushes;
            frame = win.requestAnimationFrame(() => {
              frame = win.requestAnimationFrame(() => {
                frame = null;
                if (flushes === seen) pass(false);
                else settle(false);
              });
            });
          };
          // After an edit the paragraph is measured once the keystroke has
          // been painted: typing never waits for it (iPad, Safari), unless
          // the page ledger asks for the lines first.
          let typingTimer: number | null = null;
          const typingHolds = () =>
            (win as Window & { __docsEngineLinesPerKey?: boolean }).__docsEngineLinesPerKey !== true && msSinceTyping() < TYPING_PAUSE_MS;
          /** Measure the edits once the typing has paused: nothing pending meanwhile. */
          const afterTyping = () => {
            if (typingTimer !== null) return;
            typingTimer = win.setTimeout(() => {
              typingTimer = null;
              if (typingHolds()) afterTyping();
              else scheduleAfterPaint();
            }, Math.max(16, TYPING_PAUSE_MS - msSinceTyping()));
          };
          const scheduleAfterPaint = () => {
            markPending(true);
            if (idle !== null || frame !== null) return;
            const go = () => { idle = null; schedule(); };
            const request = (win as Window & { requestIdleCallback?: (cb: () => void, options?: { timeout: number }) => number }).requestIdleCallback;
            idle = request ? request.call(win, go, { timeout: 150 }) : win.setTimeout(go, 32);
          };
          /** Blocks to measure now: the dirty ones (or all) on the sheets
           * within the page ledger's measure horizon, then the retries; the
           * dirty ranges on sheets past it are left for later. */
          const collectBlocks = () => {
            const state = documentEngineLineBreaksKey.getState(view.state);
            if (!state) return null;
            const doc = view.state.doc;
            const all = full || state.dirty === "all";
            full = false;
            const horizon = sheetsWithinMeasureHorizon(view.dom);
            recordProjectedHorizon(view.dom, ENGINE_LINES_HORIZON_KEY);
            const blocks: Array<{ node: ProseMirrorNode; pos: number; retry: boolean }> = [];
            const later: Array<[number, number]> = [];
            const seen = new Set<number>();
            const retryAt = new Set(state.retry.map(([from]) => from));
            const collect = (from: number, to: number) =>
              doc.nodesBetween(Math.max(0, from), Math.min(doc.content.size, to), (node, pos) => {
                if (!node.isTextblock) return true;
                if (!seen.has(pos) && node.content.size) {
                  seen.add(pos);
                  blocks.push({ node, pos, retry: retryAt.has(pos) });
                }
                return false;
              });
            const takenSheets = [...fontSheets];
            const fromFonts = takenSheets.length ? sheetRangesOf(view, takenSheets) : [];
            fontSheets.clear();
            const wanted: Array<[number, number]> = all ? [[0, doc.content.size]] : [...(state.dirty === "all" ? [] : state.dirty), ...fromFonts];
            if (all) rounds.clear();
            // A long list read sheet by sheet from its merged, ordered form:
            // every range against every sheet was sheets x ranges a round (a
            // document held as parts lists a part's paragraphs with each part).
            const ordered = fastRangeMappingOn() && wanted.length > 64 ? mergeRanges(wanted) : null;
            let cursor = 0;
            const clip = (ranges: readonly [number, number][], sheetFrom: number, sheetTo: number) => ranges
              .map(([from, to]): [number, number] => [Math.max(from, sheetFrom), Math.min(to, sheetTo)])
              .filter(([from, to]) => to > from || (to === from && from > sheetFrom && from < sheetTo));
            doc.forEach((sheet, offset) => {
              const sheetFrom = offset, sheetTo = offset + sheet.nodeSize;
              let parts: Array<[number, number]>;
              if (ordered) {
                while (cursor < ordered.length && ordered[cursor]![1] < sheetFrom) cursor += 1;
                let end = cursor;
                while (end < ordered.length && ordered[end]![0] <= sheetTo) end += 1;
                parts = clip(ordered.slice(cursor, end), sheetFrom, sheetTo);
              } else parts = clip(wanted, sheetFrom, sheetTo);
              if (!parts.length) return;
              const dom = horizon ? view.nodeDOM(offset) : null;
              if (horizon && dom && !horizon.has(dom as Element)) {
                later.push(...parts);
                return;
              }
              for (const [from, to] of parts) collect(from, to);
            });
            for (const [from, to] of state.retry) collect(from, to);
            if (facesLanded && waitingForFaces.size && !pendingAnywhereOff()) {
              facesLanded = false;
              for (const { node, pos } of blocksWaitingForFaces(doc, waitingForFaces)) {
                if (seen.has(pos)) continue;
                seen.add(pos);
                blocks.push({ node, pos, retry: false });
              }
              waitingForFaces.clear();
            }
            return { doc, blocks, all, later, takenSheets, dirty: state.dirty === "all" || state.dirty.length > 0 || state.retry.length > 0 || fromFonts.length > 0 };
          };
          /** Measure; `sync`: every block now, the retry rounds too (the
           * ledger is waiting). Otherwise a slice per frame. */
          const pass = (sync: boolean) => {
            if (!view.dom.isConnected) {
              markPending(false);
              waitUntilAttached();
              return;
            }
            // An IME owns the DOM: nothing is read (a justified paragraph is
            // read left-aligned for a moment) and nothing dispatched, which the
            // composition guard would drop anyway; the lines stay pending and
            // are measured on the settled wake-up.
            if (isImeCompositionActive(view)) {
              markPending(true);
              return;
            }
            let anyFacesPending = false;
            let readAll = false;
            for (let round = 0; round <= MAX_ROUNDS; round += 1) {
              const collected = collectBlocks();
              if (!collected) return;
              readAll ||= collected.all;
              const { doc, blocks, all, dirty, later, takenSheets } = collected;
              if (!blocks.length) {
                // Nothing within the horizon: dispatch only if that changes
                // the state (a flush that finds the same ranges left does
                // nothing, so the ledger's flush rounds stay cheap).
                const state = documentEngineLineBreaksKey.getState(view.state);
                const unchanged = state && state.retry.length === 0 && state.dirty !== "all" &&
                  state.dirty.length === later.length && state.dirty.every(([from, to], at) => later[at][0] === from && later[at][1] === to);
                if (dirty && !unchanged) view.dispatch(view.state.tr.setMeta(META, { spans: [], ranges: [], retry: [], dirty: later } satisfies EngineLinesMeta).setMeta("addToHistory", false));
                settle(anyFacesPending, readAll);
                return;
              }
              let index = 0;
              let facesPending = false;
              const floats: DocsFloatCache = new Map();
              const cache = { lines: lineCache, news: familyNews(view.dom.ownerDocument), stats: cacheStats() };
              const spans: Span[] = [];
              const measured: Array<[number, number]> = [];
              const retry: Array<[number, number]> = [];
              const finish = () => {
                const tr = view.state.tr.setMeta(META, { spans, ranges: measured, retry, dirty: later } satisfies EngineLinesMeta).setMeta("addToHistory", false);
                // Paragraphs newly left to the browser are read again by the
                // projections that keep Word's rules for the browser's lines.
                if (takeEngineDeclinedNews()) tr.setMeta("justifiedSpaceShrinkRemeasure", true);
                view.dispatch(tr);
                return retry.length > 0;
              };
              const slice = (count: number) => {
                const end = Math.min(blocks.length, index + count);
                for (; index < end; index += 1) {
                  const { node, pos, retry: again } = blocks[index];
                  const done = again ? (rounds.get(pos) ?? 0) : 0;
                  const result = measureBlock(view, node, pos, done, floats, cache);
                  measured.push([pos, pos + node.nodeSize]);
                  if (!result) {
                    if (result === undefined) {
                      facesPending = true;
                      waitingForFaces.add(node);
                    }
                    continue;
                  }
                  spans.push(...result.spans);
                  if (!result.agrees && done < MAX_ROUNDS && result.spans.length) {
                    rounds.set(pos, done + 1);
                    retry.push([pos, pos + node.nodeSize]);
                  } else if (result.agrees) rounds.delete(pos);
                }
              };
              if (sync) {
                slice(blocks.length);
                const more = finish();
                anyFacesPending ||= facesPending;
                if (!more) {
                  settle(anyFacesPending, readAll);
                  return;
                }
                continue;
              }
              const giveBack = () => {
                if (all) full = true;
                for (const sheet of takenSheets) fontSheets.add(sheet);
              };
              const step = () => {
                frame = null;
                inFlight = null;
                if (!view.dom.isConnected) { giveBack(); return; }
                if (view.state.doc !== doc) {
                  // Edited meanwhile: what is left is measured next time.
                  giveBack();
                  if (typingHolds()) afterTyping();
                  else schedule();
                  return;
                }
                slice(BLOCKS_PER_SLICE);
                if (index < blocks.length) {
                  frame = win.requestAnimationFrame(step);
                  inFlight = giveBack;
                  return;
                }
                const more = finish();
                settle(facesPending, all);
                if (more) schedule();
              };
              step();
              return;
            }
            settle(anyFacesPending, readAll);
          };
          /** A pass ended: the ledger may lay pages out unless faces are
           * still coming (held a few seconds at most). */
          // Paragraphs read while a face was loading have no lines yet, and
          // nothing said when they could be read: the last pass of an opening
          // document ran 150 ms before its faces answered, no font event
          // followed, and its three justified paragraphs stayed the browser's
          
          // all pending, 0 corrections until an edit). The engine's epoch moves
          // when a pending answer may have come: it is looked at for a minute.
          let facesWatch: number | null = null;
          let facesEpoch = "";
          let facesWatchUntil = 0;
          const watchFaces = () => {
            facesEpoch = textEngineEpoch();
            if (facesWatch !== null) return;
            facesWatchUntil = performance.now() + FACE_WATCH_MS;
            const look = () => {
              facesWatch = null;
              if (textEngineEpoch() !== facesEpoch) {
                // Read again, every paragraph; still pending, the pass watches on.
                full = true;
                facesLanded = true;
                schedule();
                return;
              }
              if (performance.now() < facesWatchUntil) facesWatch = win.setTimeout(look, FACE_WATCH_STEP_MS);
            };
            facesWatch = win.setTimeout(look, FACE_WATCH_STEP_MS);
          };
          





          const settle = (facesPending: boolean, readAll = false) => {
            const state = documentEngineLineBreaksKey.getState(view.state);
            // Dirty ranges past the measure horizon hold nothing: the ledger
            // asks for them (its flush) when its horizon reaches them.
            const work = frame !== null || idle !== null || Boolean(state && (state.dirty === "all" || state.retry.length));
            // A wait past its hold holds nothing and is ended by any pass (a
            // later wait starts its own hold).
            if (facesPending) facesPendingSince ??= performance.now();
            else if (readAll || (facesPendingSince !== null && performance.now() - facesPendingSince >= FACE_HOLD_MS)) facesPendingSince = null;
            if (facesPending) watchFaces();
            const holding = facesPendingSince !== null && performance.now() - facesPendingSince < FACE_HOLD_MS;
            markPending(work || holding);
            if (holding && !work) win.setTimeout(() => { if (facesPendingSince !== null) settle(true); }, FACE_HOLD_MS);
          };
          const onFlush = () => {
            const state = documentEngineLineBreaksKey.getState(view.state);
            const dirty = Boolean(state && (state.dirty === "all" || state.dirty.length || state.retry.length));
            if (!view.dom.hasAttribute(DOCUMENT_ENGINE_LINES_PENDING) && !dirty) return;
            // Ranges a hand is still typing in wait for the pause (afterTyping).
            if (!view.dom.hasAttribute(DOCUMENT_ENGINE_LINES_PENDING) && typingHolds()) return;
            flushes += 1;
            cancelScheduled();
            pass(true);
          };
          view.dom.addEventListener(LAYOUT_METRICS_FLUSH_EVENT, onFlush);
          // Faces landing (the engine's answers move from pending) and a
          // width change re-read every paragraph.
          const onFonts = (event?: Event) => {
            // (projectionFontNews) With the trial, only the paragraphs of the
            // sheets drawing the arrived face's characters are read again.
            const reach = projectionFontReach(view, "engine-lines", event);
            if (reach === "none") return;
            facesLanded = true;
            if (reach === "all") full = true;
            else for (const sheet of reach.sheets) fontSheets.add(sheet);
            schedule();
          };
          view.dom.ownerDocument.fonts?.addEventListener?.("loadingdone", onFonts);
          let lastWidth = view.dom.getBoundingClientRect().width;
          const observer = "ResizeObserver" in win
            ? new win.ResizeObserver(() => {
              const width = view.dom.getBoundingClientRect().width;
              if (Math.abs(width - lastWidth) < 0.5) return;
              lastWidth = width;
              full = true;
              schedule();
            })
            : null;
          observer?.observe(view.dom);
          const stopImeSettled = onImeCompositionSettled(view, () => {
            if (view.dom.hasAttribute(DOCUMENT_ENGINE_LINES_PENDING)) schedule();
          });
          schedule();
          return {
            update(updated, previousState) {
              // Only an edit: the plugin's own transactions (corrections,
              // ranges left past the horizon) schedule nothing, or the ledger
              // would see lines pending after every flush and never lay
              // pages out.
              if (updated.state.doc !== previousState.doc) {
                if (typingHolds()) afterTyping();
                else scheduleAfterPaint();
              }
            },
            destroy() {
              stopImeSettled();
              observer?.disconnect();
              view.dom.removeEventListener(LAYOUT_METRICS_FLUSH_EVENT, onFlush);
              view.dom.ownerDocument.fonts?.removeEventListener?.("loadingdone", onFonts);
              cancelScheduled();
              if (waitTimer !== null) win.clearTimeout(waitTimer);
              if (facesWatch !== null) win.clearTimeout(facesWatch);
              if (typingTimer !== null) win.clearTimeout(typingTimer);
              markPending(false);
            },
          };
        },
      }),
    ];
  },
});
