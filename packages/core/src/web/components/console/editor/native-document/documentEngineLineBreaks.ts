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
import { touchedTextblockRanges } from "./touchedTextblocks";
import { isImeCompositionActive, onImeCompositionSettled } from "./imeCompositionGuard";
import { DOCUMENT_ENGINE_LINES_PENDING, LAYOUT_METRICS_FLUSH_EVENT, writeDocumentLayoutStatus } from "./documentLayoutStatus";
import { recordProjectedHorizon, sheetsWithinMeasureHorizon } from "./paged/measureHorizon";
import { createEngineLineCache, familyNews, type EngineLineCache, type EngineLineKey, type FamilyNews } from "./engineLineCache";
import { PAGE_BLOCK_NODE } from "./paged/paginationProjectionKeys";
import { textEngineEpoch } from "../../../../lib/text-engine/measure";

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
const CJK = /[぀-ヿ㐀-鿿가-힯＀-￯]/;

/** Whether `transaction` publishes the engine's line corrections only. */
export function isDocumentEngineLineBreaksTransaction(transaction: { getMeta(key: string): unknown }): boolean {
  return transaction.getMeta(META) !== undefined;
}

interface Span {
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

function decorationsFor(doc: ProseMirrorNode, spans: readonly Span[]): Decoration[] {
  return spans.map((span) => {
    if (span.tabBreak) return Decoration.widget(span.from, tabBreakElement, { side: -1, key: DOCS_ENGINE_TAB_BREAK_CLASS, engineLine: true, ignoreSelection: true });
    const total = Number((span.base + span.extra).toFixed(4));
    const words = span.wordExtra ? `;word-spacing:${Number((span.wordBase + span.wordExtra).toFixed(4))}px` : "";
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
  return { compatibilityMode: normalizeWordCompatibilityMode(view.state.doc.attrs.wordCompatibilityMode), kinsoku };
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

/** Counts for probes (`window.__docsEngineLineCache`). */
interface EngineLineCacheStats { measured: number; reused: number; kept: number; missed?: Record<string, number> }

/** What a block's lines depend on besides itself (engineLineCache). */
function lineKeyOf(view: EditorView, element: HTMLElement, pos: number): EngineLineKey {
  const parent = view.state.doc.resolve(pos).parent;
  return {
    width: element.getBoundingClientRect().width,
    epoch: textEngineEpoch(),
    docAttrs: view.state.doc.attrs,
    parent: parent.type.name === PAGE_BLOCK_NODE ? PAGE_BLOCK_NODE : parent.attrs,
  };
}

function measureBlock(
  view: EditorView, node: ProseMirrorNode, pos: number, round: number, floats: DocsFloatCache,
  cache?: { lines: EngineLineCache; news: FamilyNews; stats: EngineLineCacheStats | null },
): { spans: Span[]; agrees: boolean } | null | undefined {
  const element = view.nodeDOM(pos);
  if (!(element instanceof HTMLElement) || !element.isConnected) return null;
  // A cell of an auto-layout table sizes its column from its content's
  // widths, which a correction would change: the browser keeps it.
  const table = element.closest("table");
  if (table && element.ownerDocument.defaultView?.getComputedStyle(table).tableLayout !== "fixed") return null;
  // A block measured as it is now (moved to another sheet by the page
  // ledger, unchanged) takes its corrections back without a reading.
  const key = cache && round === 0 ? lineKeyOf(view, element, pos) : null;
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
  if (paragraph.unsupported || !paragraph.pieces.length) return null;
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
  const tabOpened = engine.some((start, line) => line > 0 && paragraph.pieces[start]?.kind === "tab");
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
            // An edited paragraph keeps its mapped corrections until it is
            // measured again: no frame of browser lines in between.
            const changed = touchedTextblockRanges(tr);
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
          const rounds = new Map<number, number>();
          const lineCache = createEngineLineCache();
          const cacheStats = (): EngineLineCacheStats | null => {
            const holder = win as Window & { __docsEngineLineCache?: EngineLineCacheStats };
            return holder.__docsEngineLineCache ??= { measured: 0, reused: 0, kept: 0 };
          };
          /** Since when paragraphs wait for a face file (the page ledger
           * waits with them, at most FACE_HOLD_MS). */
          let facesPendingSince: number | null = null;
          // The page ledger lays pages out from the drawn lines: while lines
          // are about to be corrected it waits (and asks for them at once,
          // LAYOUT_METRICS_FLUSH_EVENT), so its cuts are made on the engine's
          // lines and never undone by a correction after them.
          const markPending = (pending: boolean) => writeDocumentLayoutStatus(view.dom, DOCUMENT_ENGINE_LINES_PENDING, pending ? "true" : null, view);
          const cancelScheduled = () => {
            if (frame !== null) { win.cancelAnimationFrame(frame); frame = null; }
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
            const wanted: Array<[number, number]> = all ? [[0, doc.content.size]] : state.dirty === "all" ? [] : state.dirty;
            if (all) rounds.clear();
            doc.forEach((sheet, offset) => {
              const sheetFrom = offset, sheetTo = offset + sheet.nodeSize;
              const parts = wanted
                .map(([from, to]): [number, number] => [Math.max(from, sheetFrom), Math.min(to, sheetTo)])
                .filter(([from, to]) => to > from || (to === from && from > sheetFrom && from < sheetTo));
              if (!parts.length) return;
              const dom = horizon ? view.nodeDOM(offset) : null;
              if (horizon && dom && !horizon.has(dom as Element)) {
                later.push(...parts);
                return;
              }
              for (const [from, to] of parts) collect(from, to);
            });
            for (const [from, to] of state.retry) collect(from, to);
            return { doc, blocks, all, later, dirty: state.dirty === "all" || state.dirty.length > 0 || state.retry.length > 0 };
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
            for (let round = 0; round <= MAX_ROUNDS; round += 1) {
              const collected = collectBlocks();
              if (!collected) return;
              const { doc, blocks, all, dirty, later } = collected;
              if (!blocks.length) {
                // Nothing within the horizon: dispatch only if that changes
                // the state (a flush that finds the same ranges left does
                // nothing, so the ledger's flush rounds stay cheap).
                const state = documentEngineLineBreaksKey.getState(view.state);
                const unchanged = state && state.retry.length === 0 && state.dirty !== "all" &&
                  state.dirty.length === later.length && state.dirty.every(([from, to], at) => later[at][0] === from && later[at][1] === to);
                if (dirty && !unchanged) view.dispatch(view.state.tr.setMeta(META, { spans: [], ranges: [], retry: [], dirty: later } satisfies EngineLinesMeta).setMeta("addToHistory", false));
                settle(anyFacesPending);
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
                view.dispatch(view.state.tr.setMeta(META, { spans, ranges: measured, retry, dirty: later } satisfies EngineLinesMeta).setMeta("addToHistory", false));
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
                    if (result === undefined) facesPending = true;
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
                  settle(anyFacesPending);
                  return;
                }
                continue;
              }
              const step = () => {
                frame = null;
                if (!view.dom.isConnected) return;
                if (view.state.doc !== doc) {
                  // Edited meanwhile: what is left is measured next time.
                  if (all) full = true;
                  schedule();
                  return;
                }
                slice(BLOCKS_PER_SLICE);
                if (index < blocks.length) {
                  frame = win.requestAnimationFrame(step);
                  return;
                }
                const more = finish();
                settle(facesPending);
                if (more) schedule();
              };
              step();
              return;
            }
            settle(anyFacesPending);
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
                schedule();
                return;
              }
              if (performance.now() < facesWatchUntil) facesWatch = win.setTimeout(look, FACE_WATCH_STEP_MS);
            };
            facesWatch = win.setTimeout(look, FACE_WATCH_STEP_MS);
          };
          const settle = (facesPending: boolean) => {
            const state = documentEngineLineBreaksKey.getState(view.state);
            // Dirty ranges past the measure horizon hold nothing: the ledger
            // asks for them (its flush) when its horizon reaches them.
            const work = frame !== null || idle !== null || Boolean(state && (state.dirty === "all" || state.retry.length));
            if (facesPending) facesPendingSince ??= performance.now();
            else facesPendingSince = null;
            if (facesPending) watchFaces();
            const holding = facesPendingSince !== null && performance.now() - facesPendingSince < FACE_HOLD_MS;
            markPending(work || holding);
            if (holding && !work) win.setTimeout(() => { if (facesPendingSince !== null) settle(true); }, FACE_HOLD_MS);
          };
          const onFlush = () => {
            const state = documentEngineLineBreaksKey.getState(view.state);
            const dirty = Boolean(state && (state.dirty === "all" || state.dirty.length || state.retry.length));
            if (!view.dom.hasAttribute(DOCUMENT_ENGINE_LINES_PENDING) && !dirty) return;
            flushes += 1;
            cancelScheduled();
            pass(true);
          };
          view.dom.addEventListener(LAYOUT_METRICS_FLUSH_EVENT, onFlush);
          // Faces landing (the engine's answers move from pending) and a
          // width change re-read every paragraph.
          const onFonts = () => { full = true; schedule(); };
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
              if (updated.state.doc !== previousState.doc) scheduleAfterPaint();
            },
            destroy() {
              stopImeSettled();
              observer?.disconnect();
              view.dom.removeEventListener(LAYOUT_METRICS_FLUSH_EVENT, onFlush);
              view.dom.ownerDocument.fonts?.removeEventListener?.("loadingdone", onFonts);
              cancelScheduled();
              if (waitTimer !== null) win.clearTimeout(waitTimer);
              if (facesWatch !== null) win.clearTimeout(facesWatch);
              markPending(false);
            },
          };
        },
      }),
    ];
  },
});
