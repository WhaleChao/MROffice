/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { loadedFontsKey } from "../loadedFontsKey";
import { trailingAutoLeadingCreditPx } from "./pageEdgeLeading";
import { paginationLayoutRect, paginationLayoutScale } from "./paginationGeometry";
import { COLUMN_REGIONS_ATTR, unequalColumnCount } from "./unequalColumnProjection";
import { measurementRange } from "./measurementRange";
import { clearPageEdgeMarks } from "./paginationProjectionKeys";

/**
 * The head fragment a text cut produces, rendered before it exists.
 *
 * A page-bottom row is kept when the fragment ending with it fits: its line
 * boxes, less the Auto leading its last line may hang below the page edge.
 * The next pass judges the materialised fragment by exactly that rule. The
 * row planner locates rows by their painted glyph boxes, which do not give
 * the line box edge: a fixed line height does not center the glyph box
 * (Chromium floors the leading above the text to whole pixels, WebKit
 * truncates the line height), and on a row mixing runs the line box is set
 * by the tallest inline box, not by the glyph the search probed. A row the
 * planner kept by its glyph could make a head that failed the fragment
 * rule; the ledger then moved the text back and alternated between the two
 * layouts until it stalled.
 *
 * So the planner confirms its row with the fragment itself: a copy of the
 * block, cut at the row start, laid out off-screen in the block's own
 * context (the editor root and every ancestor between, cloned without
 * their other children), measured like the ledger measures fragments.
 *
 * Cost. A node added anywhere under the editor makes the browser recheck
 * the editor's `:has()`-styled ancestors, which cost several milliseconds
 * per copy. The copies therefore live in one contained host at the end of
 * the document body, under shallow clones of the editor's own ancestors
 * (their classes and styles give the copy the same inherited context; their
 * transforms and animations are dropped, which only scale the picture). The
 * host is attached for the current task, with one context shell per
 * ancestor chain. A measured head is kept until something it depends on
 * changes: a mutation inside the block, its ancestors' attributes, the
 * editor's width or styles, the loaded fonts, or an external layout change
 * the ledger reports (`invalidateRenderedTextHeads`).
 */

export interface RenderedTextHead {
  /** Border-box height of the head, in layout px. */
  heightPx: number;
  /** Auto leading its last line may hang below the page edge. */
  creditPx: number;
}

export interface RenderedTextHeadProbe {
  /** The head that keeps the block's text before `offset` characters, or
   * null when that head cannot be laid out. */
  measure(offset: number, endOfPrevious?: boolean): RenderedTextHead | null;
  dispose(): void;
}

interface TextSegment {
  node: Text;
  start: number;
}

interface ContextShell {
  /** Where the copy goes: the clone of the block's parent. */
  slot: HTMLElement;
  /** Siblings that keep the copy an interior child of a sheet. */
  leading: HTMLElement | null;
  trailing: HTMLElement | null;
}

interface TaskHost {
  element: HTMLElement;
  /** Clone of the editor root's parent, inside the cloned ancestor chain. */
  context: HTMLElement;
  rootWidth: string;
  shells: Map<HTMLElement, ContextShell>;
}

/** Layout-neutral styles for the cloned ancestors: they lend the copy their
 * inherited context and matched rules, never their size or position. */
const NEUTRAL_ANCESTOR_STYLES: Readonly<Record<string, string>> = {
  display: "block",
  position: "static",
  width: "max-content",
  height: "auto",
  "min-width": "0",
  "min-height": "0",
  "max-width": "none",
  "max-height": "none",
  margin: "0",
  overflow: "visible",
  transform: "none",
  translate: "none",
  rotate: "none",
  scale: "none",
  animation: "none",
  transition: "none",
  contain: "none",
  "content-visibility": "visible",
};

function neutralize(element: HTMLElement, styles: Readonly<Record<string, string>>): void {
  for (const [name, value] of Object.entries(styles)) {
    element.style.setProperty(name, value, "important");
  }
}

/** A clone of an application element keeps what styles it and nothing that
 * identifies it to scripts, tests or assistive technology. */
function anonymize(element: HTMLElement): void {
  for (const name of element.getAttributeNames()) {
    if (name === "id" || name === "data-testid" || name === "role" || name === "tabindex" ||
        name === "draggable" || name.startsWith("aria-")) {
      element.removeAttribute(name);
    }
  }
}

interface CachedHeads {
  context: string;
  byOffset: Map<number, RenderedTextHead | null>;
}

interface RootState {
  task: TaskHost | null;
  generation: number;
  observer: MutationObserver | null;
  heads: WeakMap<HTMLElement, CachedHeads>;
}

const stateByRoot = new WeakMap<HTMLElement, RootState>();

function forgetMutated(state: RootState, root: HTMLElement, records: readonly MutationRecord[]): void {
  for (const record of records) {
    // A projection re-stamping the value already there (a layout pass
    // writing a frame's mode and style every transaction) changed nothing
    // the head was measured from.
    if (
      record.type === "attributes" &&
      record.target instanceof Element &&
      record.target.getAttribute(record.attributeName ?? "") === record.oldValue
    ) {
      continue;
    }
    for (let node: Node | null = record.target; node && node !== root; node = node.parentNode) {
      if (state.heads.has(node as HTMLElement)) state.heads.delete(node as HTMLElement);
    }
  }
}

function rootState(root: HTMLElement): RootState {
  let state = stateByRoot.get(root);
  if (!state) {
    const created: RootState = { task: null, generation: 0, observer: null, heads: new WeakMap() };
    if (typeof MutationObserver === "function") {
      created.observer = new MutationObserver((records) => forgetMutated(created, root, records));
      created.observer.observe(root, { attributes: true, attributeOldValue: true, characterData: true, childList: true, subtree: true });
    }
    stateByRoot.set(root, created);
    state = created;
  }
  if (state.observer) forgetMutated(state, root, state.observer.takeRecords());
  return state;
}

/**
 * Run a sheet-level probe in the contained task host: `fn` receives a
 * layout-neutral clone of the editor root — at the root's width, under the
 * cloned ancestor chain that gives it the root's inherited context — into
 * which it clones the sheets it measures. Nothing under the editor is
 * touched, so the document stays laid out and every read that follows in
 * the pass (fragment heights, the edge-margin probe, the selection
 * ProseMirror writes back) pays no full-document layout. A probe appended
 * beside the editor and removed again dirtied the editor's ancestors each
 * time; on a 157-page import each of the four or five reads that followed
 * then laid the whole document out (about 200 ms apiece, a keystroke's
 * worth of stall). Null when the root has no document body to host it.
 */
export function withRootMeasurementSlot<T>(root: HTMLElement, fn: (slot: HTMLElement) => T): T | null {
  const state = rootState(root);
  const task = taskHost(root, state);
  if (!task) return null;
  const shell = root.cloneNode(false) as HTMLElement;
  anonymize(shell);
  neutralize(shell, { ...NEUTRAL_ANCESTOR_STYLES, width: task.rootWidth, "min-width": task.rootWidth });
  task.context.append(shell);
  try {
    return fn(shell);
  } finally {
    shell.remove();
  }
}

/**
 * Forget every rendered head of the editor at `root`: a layout input outside
 * the document changed (surface geometry, view mode, fonts).
 */
export function invalidateRenderedTextHeads(root: HTMLElement): void {
  const state = stateByRoot.get(root);
  if (state) state.generation += 1;
}

/** The contained host for `root` in the current task. */
function taskHost(root: HTMLElement, state: RootState): TaskHost | null {
  const doc = root.ownerDocument;
  const body = doc.body;
  if (!body || !root.parentElement) return null;
  if (state.task?.element.parentElement === body) return state.task;
  const element = doc.createElement("div");
  element.setAttribute("aria-hidden", "true");
  element.setAttribute("inert", "");
  element.style.cssText = [
    "position:absolute",
    "left:-100000px",
    "top:0",
    "width:0",
    "height:0",
    "margin:0",
    "padding:0",
    "border:0",
    "overflow:visible",
    "visibility:hidden",
    "pointer-events:none",
    // A relayout boundary: what happens inside never lays out the editor.
    "contain:strict",
  ].map((declaration) => `${declaration}!important`).join(";");
  const lineage: HTMLElement[] = [];
  for (let ancestor: HTMLElement | null = root.parentElement; ancestor && ancestor !== body; ancestor = ancestor.parentElement) {
    lineage.unshift(ancestor);
  }
  let context: HTMLElement = element;
  for (const ancestor of lineage) {
    const copy = ancestor.cloneNode(false) as HTMLElement;
    anonymize(copy);
    neutralize(copy, NEUTRAL_ANCESTOR_STYLES);
    context.append(copy);
    context = copy;
  }
  const rootWidth = getComputedStyle(root).width;
  const task: TaskHost = {
    element,
    context,
    rootWidth: rootWidth.endsWith("px") && Number.parseFloat(rootWidth) > 0 ? rootWidth : `${root.clientWidth}px`,
    shells: new Map(),
  };
  body.append(element);
  state.task = task;
  queueMicrotask(() => {
    element.remove();
    if (state.task === task) state.task = null;
  });
  return task;
}

function contextShell(
  root: HTMLElement,
  task: TaskHost,
  ancestors: readonly HTMLElement[],
): ContextShell {
  const key = ancestors.at(-1) ?? root;
  const existing = task.shells.get(key);
  if (existing) return existing;
  const doc = root.ownerDocument;
  const shell = root.cloneNode(false) as HTMLElement;
  anonymize(shell);
  neutralize(shell, { ...NEUTRAL_ANCESTOR_STYLES, width: task.rootWidth, "min-width": task.rootWidth });
  let slot: HTMLElement = shell;
  for (const ancestor of ancestors) {
    const copy = ancestor.cloneNode(false) as HTMLElement;
    anonymize(copy);
    slot.append(copy);
    slot = copy;
  }
  // Keep the block an interior child of its sheet, as the edge rules see
  // most blocks (the same shape the page-edge margin probe uses).
  const interior = slot === shell || slot.matches("article[data-page-sheet]");
  const context: ContextShell = {
    slot,
    leading: interior ? doc.createElement("span") : null,
    trailing: interior ? doc.createElement("span") : null,
  };
  task.context.append(shell);
  task.shells.set(key, context);
  return context;
}

function attributeSignature(element: HTMLElement): string[] {
  return element.getAttributeNames()
    .filter((name) => name !== "id")
    .map((name) => `${name}=${element.getAttribute(name)}`);
}

/**
 * A copy of `block` that ends where its text node `source` (at `path` in it)
 * is cut at `offset`, as a page cut's head ends: what follows dropped, the
 * wrappers it leaves empty and trailing source page breaks too.
 *
 * Only the head is copied (its ancestors split as the cut leaves them): a
 * whole long paragraph copied for each candidate head, its tail then
 * deleted, cost a copy of every run per head measured (R461-D71, code
 * listings of 3,500 runs). `whole` copies it whole, as before
 * (`window.__docsRenderedHeadWholeCloneOn`).
 */
export function copyTextHead(block: HTMLElement, source: Text, offset: number, path: readonly number[], whole = false): HTMLElement | null {
  const doc = block.ownerDocument;
  const range = measurementRange(doc, "renderedTextHead");
  const at = Math.max(0, Math.min(source.length, offset));
  let copy: HTMLElement;
  if (whole) {
    copy = block.cloneNode(true) as HTMLElement;
  } else {
    range.setStart(block, 0);
    range.setEnd(source, at);
    copy = block.cloneNode(false) as HTMLElement;
    copy.append(range.cloneContents());
  }
  copy.removeAttribute("id");
  clearPageEdgeMarks(copy);
  copy.setAttribute("data-page-split-continues", "true");
  const text = nodeAtPath(copy, path);
  if (text?.nodeType !== Node.TEXT_NODE) return null;
  range.setStart(text, Math.min((text as Text).length, at));
  range.setEnd(copy, copy.childNodes.length);
  range.deleteContents();
  pruneEmptyTail(copy, text);
  dropTrailingSourcePageBreaks(copy);
  return copy;
}

function wholeCloneOn(doc: Document): boolean {
  return (doc.defaultView as (Window & { __docsRenderedHeadWholeCloneOn?: boolean }) | null)?.__docsRenderedHeadWholeCloneOn === true;
}

export function childPath(root: Node, node: Node): number[] | null {
  const path: number[] = [];
  let current: Node | null = node;
  while (current && current !== root) {
    const parent: Node | null = current.parentNode;
    if (!parent) return null;
    path.unshift(Array.prototype.indexOf.call(parent.childNodes, current) as number);
    current = parent;
  }
  return current === root ? path : null;
}

function nodeAtPath(root: Node, path: readonly number[]): Node | null {
  let current: Node | null = root;
  for (const index of path) {
    current = current?.childNodes[index] ?? null;
  }
  return current;
}

/** Drop what a cut at a run's first character leaves behind: inline
 * wrappers with nothing in them. ProseMirror renders no such wrapper. */
function pruneEmptyTail(block: HTMLElement, from: Node): void {
  let current: Node | null = from;
  while (current && current !== block) {
    const parent: Node | null = current.parentNode;
    const empty = current.nodeType === Node.TEXT_NODE
      ? (current as Text).length === 0
      : current.nodeType === Node.ELEMENT_NODE &&
        current.childNodes.length === 0 &&
        (current as Element).matches("span, a, strong, em, b, i, u, s, sub, sup, mark, code, small");
    if (!empty) return;
    parent?.removeChild(current);
    current = parent;
  }
}

/** The cut leaves a source page break that starts the cut row with the
 * tail (`planTextHeadOnce`): so does the copy, with the separator image
 * ProseMirror paints before that inline atom. */
function dropTrailingSourcePageBreaks(block: HTMLElement): void {
  for (;;) {
    let last: Node | null = block.lastChild;
    while (last && last.nodeType === Node.ELEMENT_NODE && last.lastChild &&
      !(last as Element).matches('[data-type="source-page-break"]')) {
      last = last.lastChild;
    }
    if (!last || last === block) return;
    const element = last.nodeType === Node.ELEMENT_NODE ? (last as Element) : null;
    const removable = element
      ? element.matches('[data-type="source-page-break"], img.ProseMirror-separator')
      : last.nodeType === Node.TEXT_NODE && (last as Text).length === 0;
    if (!removable) return;
    const parent = last.parentNode;
    parent?.removeChild(last);
    if (parent && parent !== block) pruneEmptyTail(block, parent);
  }
}

/**
 * A probe for `block`'s heads, or null when the block's fragments are not
 * measured as plain border boxes (no editor root, multi-column sheets).
 */
export function createRenderedTextHeadProbe(
  block: HTMLElement,
  segments: readonly TextSegment[],
): RenderedTextHeadProbe | null {
  const doc = block.ownerDocument;
  const root = block.closest<HTMLElement>(".ProseMirror");
  if (!doc || !root?.isConnected || !root.parentElement || root === block) return null;
  const sheet = block.closest<HTMLElement>("article[data-page-sheet]");
  if (sheet) {
    // A page of stacked column regions (continuous sections) places each
    // block at its own track's width; the probe's shell lays the head out at
    // the page's width, where 40 lines of a 192px column fit in 14, so a
    // column cut kept all but one line per pass (b-sect2.docx S5).
    if (sheet.hasAttribute(COLUMN_REGIONS_ATTR)) return null;
    const unequal = unequalColumnCount(sheet);
    if (unequal !== null && unequal > 1) return null;
    const columns = Number.parseInt(getComputedStyle(sheet).columnCount, 10);
    if (Number.isFinite(columns) && columns > 1) return null;
  }
  const ancestors: HTMLElement[] = [];
  for (let element = block.parentElement; element !== root; element = element.parentElement) {
    if (!element) return null;
    ancestors.unshift(element);
  }
  if (segments.length === 0) return null;
  // Where each text node stands, read when a head is cut in it. Read for
  // every node up front, each an index among its parent's children, it cost
  // the square of a paragraph's runs per probe (R461-D71's code listings,
  // 3,500 runs). A node no longer in the block makes no head.
  const paths = new Map<number, number[] | null>();
  const pathOf = (index: number): number[] | null => {
    if (!paths.has(index)) paths.set(index, childPath(block, segments[index].node));
    return paths.get(index) ?? null;
  };
  if (pathOf(0) === null || pathOf(segments.length - 1) === null) return null;

  let placed: ContextShell | null = null;
  let contextKey: string | null | undefined;

  const segmentIndexAt = (offset: number): number => {
    let lo = 0;
    let hi = segments.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (segments[mid].start <= offset) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };

  // `endOfPrevious`: the head ends right after the character before
  // `offset`, in that character's own text node, so what stands between it
  // and the next text (an in-line picture's row) goes with the tail.
  const layOut = (task: TaskHost, offset: number, endOfPrevious: boolean): RenderedTextHead | null => {
    const index = segmentIndexAt(endOfPrevious ? offset - 1 : offset);
    const path = pathOf(index);
    if (!path) return null;
    const source = segments[index].node;
    const copy = copyTextHead(block, source, offset - segments[index].start, path, wholeCloneOn(doc));
    if (!copy) return null;
    const context = contextShell(root, task, ancestors);
    context.slot.replaceChildren(
      ...[context.leading, copy, context.trailing].filter((node): node is HTMLElement => node !== null),
    );
    placed = context;
    // The copy sits in the measurement slot, so it is drawn at the slot's
    // scale: read it from the slot, which a layout pass reads once
    // (withPaginationScaleReuse), not from each new copy (a forced layout
    // and a computed style per candidate cut, a quarter of a typed key's
    // pagination in a long document).
    const scale = paginationLayoutScale(context.slot);
    const heightPx = paginationLayoutRect(copy, scale).height;
    if (!Number.isFinite(heightPx) || !(heightPx > 0)) return null;
    return { heightPx, creditPx: trailingAutoLeadingCreditPx(copy) };
  };

  return {
    measure(offset, endOfPrevious = false) {
      if (endOfPrevious && offset <= 0) return null;
      const key = endOfPrevious ? -offset - 0.5 : offset;
      const state = rootState(root);
      const task = taskHost(root, state);
      if (!task) return null;
      if (contextKey === undefined) {
        const fonts = loadedFontsKey(doc);
        contextKey = fonts === null
          ? null
          : JSON.stringify([
              state.generation,
              fonts,
              task.rootWidth,
              root.className,
              root.style.cssText,
              ancestors.map((ancestor) => [ancestor.tagName, attributeSignature(ancestor)]),
            ]);
      }
      let cached = contextKey === null ? undefined : state.heads.get(block);
      if (cached && cached.context !== contextKey) cached = undefined;
      if (cached?.byOffset.has(key)) return cached.byOffset.get(key)!;
      const head = layOut(task, offset, endOfPrevious);
      if (contextKey !== null) {
        if (!cached) {
          cached = { context: contextKey, byOffset: new Map() };
          state.heads.set(block, cached);
        }
        cached.byOffset.set(key, head);
      }
      return head;
    },
    dispose() {
      const context = placed as ContextShell | null;
      context?.slot.replaceChildren();
      placed = null;
    },
  };
}
