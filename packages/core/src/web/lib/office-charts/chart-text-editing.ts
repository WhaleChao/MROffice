/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A chart text typed in place: the editor's own drawing of a chart title or
 * axis title (`OfficeChartTextLayer`) made editable, as Office edits a title
 * as the text box it is -- the caret where the click landed, Enter a new
 * line, the words' runs keeping their format.
 *
 * The words live in the DOM while they are typed: each run is a span marked
 * with the model run it came from (`data-chart-run`), a line break is a
 * `<br>` (every browser edits those natively; WebKit rewrote a typed `\n`
 * character next to new words as a space). As a browser lays them out, one
 * `<br>` at the very end of the words is no line of its own: a new last line
 * keeps a second one after it. These helpers read the words back as runs,
 * find the characters selected, and put the caret.
 */

/** The words typed, in pieces of the model runs they belong to. */
export interface OfficeChartTextValue {
  text: string;
  pieces: { text: string; run: number }[];
}

/** The attribute of a run's span: its index in the model's runs. */
export const CHART_TEXT_RUN_ATTRIBUTE = "data-chart-run";
/** A `<br>` a browser keeps so a trailing new line shows: no character. */
const SENTINEL_ATTRIBUTE = "data-chart-text-sentinel";

const BLOCK_TAGS = new Set(["DIV", "P", "LI"]);

interface Walked {
  node: Node;
  /** Offset of the node's first character in the text. */
  start: number;
  text: string;
  run: number;
}

/** Whether a `<br>` is the very last thing in the host (no line of its own). */
function trailingBreak(host: HTMLElement, br: HTMLElement): boolean {
  const walker = host.ownerDocument.createTreeWalker(host, NodeFilter.SHOW_ALL);
  walker.currentNode = br;
  for (let next = walker.nextNode(); next; next = walker.nextNode()) {
    if (next.nodeType === Node.TEXT_NODE && (next.textContent ?? "") !== "") return false;
    if (next instanceof HTMLElement && next.tagName === "BR") return false;
  }
  return true;
}

/**
 * The host's characters in order: its text nodes, a `<br>` as a line break
 * (not the sentinel, nor the one that ends the words, which a browser draws
 * as no line), and a block a browser made for a new line as a line break
 * before it.
 */
function walk(host: HTMLElement): Walked[] {
  const out: Walked[] = [];
  let at = 0;
  let lastRun = 0;
  const push = (node: Node, text: string, run: number) => {
    out.push({ node, start: at, text, run });
    at += text.length;
    lastRun = run;
  };
  const visit = (node: Node, run: number) => {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = (node.textContent ?? "").replace(/\r\n?/g, "\n").replace(/ /g, " ");
      if (text) push(node, text, run);
      return;
    }
    if (!(node instanceof HTMLElement)) return;
    if (node.tagName === "BR") {
      if (node.hasAttribute(SENTINEL_ATTRIBUTE) || trailingBreak(host, node)) return;
      push(node, "\n", run);
      return;
    }
    const own = node.getAttribute(CHART_TEXT_RUN_ATTRIBUTE);
    const nodeRun = own !== null && /^\d+$/.test(own) ? Number(own) : run;
    if (node !== host && BLOCK_TAGS.has(node.tagName) && at > 0) push(node, "\n", lastRun);
    for (const child of Array.from(node.childNodes)) visit(child, nodeRun);
  };
  for (const child of Array.from(host.childNodes)) visit(child, lastRun);
  return out;
}

/** The words in the host, as the model will keep them. */
export function readChartTextHost(host: HTMLElement): OfficeChartTextValue {
  const pieces: { text: string; run: number }[] = [];
  for (const entry of walk(host)) {
    const last = pieces[pieces.length - 1];
    if (last && last.run === entry.run) last.text += entry.text;
    else pieces.push({ text: entry.text, run: entry.run });
  }
  return { text: pieces.map((piece) => piece.text).join(""), pieces };
}

/** The character offset in the host of a DOM position (node, offset). */
function offsetOf(host: HTMLElement, entries: Walked[], node: Node, offset: number): number | null {
  if (!host.contains(node)) return null;
  if (node.nodeType === Node.TEXT_NODE) {
    const entry = entries.find((item) => item.node === node);
    if (entry) return entry.start + Math.min(offset, entry.text.length);
  }
  // A position between children: the characters of everything before it.
  const marker = node.childNodes[offset] ?? null;
  let total = 0;
  for (const entry of entries) {
    if (marker && (entry.node === marker || marker.contains(entry.node) || isAfter(entry.node, marker))) break;
    if (!marker && !node.contains(entry.node) && isAfter(entry.node, node)) break;
    total = entry.start + entry.text.length;
  }
  return total;
}

function isAfter(node: Node, reference: Node): boolean {
  return Boolean(reference.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING);
}

/** The characters a range covers in the host, or null when it is elsewhere. */
export function chartTextRangeOffsets(
  host: HTMLElement,
  range: Range,
): { start: number; end: number } | null {
  const entries = walk(host);
  const start = offsetOf(host, entries, range.startContainer, range.startOffset);
  const end = offsetOf(host, entries, range.endContainer, range.endOffset);
  if (start === null || end === null) return null;
  return { start: Math.min(start, end), end: Math.max(start, end) };
}

const lastSelections = new WeakMap<HTMLElement, { start: number; end: number }>();

/** Remember the host's selection (the ribbon takes the focus, Office keeps the words selected). */
export function rememberChartTextSelection(host: HTMLElement): void {
  const selection = host.ownerDocument.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!host.contains(range.startContainer) || !host.contains(range.endContainer)) return;
  const offsets = chartTextRangeOffsets(host, range);
  if (offsets) lastSelections.set(host, offsets);
}

/** The characters selected in the host: the live selection, else the last one it had. */
export function chartTextSelection(host: HTMLElement): { start: number; end: number } | null {
  const selection = host.ownerDocument.getSelection();
  if (selection && selection.rangeCount > 0) {
    const range = selection.getRangeAt(0);
    if (host.contains(range.startContainer) && host.contains(range.endContainer)) {
      const offsets = chartTextRangeOffsets(host, range);
      if (offsets) return offsets;
    }
  }
  return lastSelections.get(host) ?? null;
}

/** The DOM position of a character offset in the host. */
function positionAt(host: HTMLElement, offset: number): { node: Node; offset: number } {
  const entries = walk(host);
  for (const entry of entries) {
    if (offset <= entry.start + entry.text.length && entry.node.nodeType === Node.TEXT_NODE) {
      return { node: entry.node, offset: Math.max(0, offset - entry.start) };
    }
    if (offset <= entry.start && entry.node.parentNode) {
      const parent = entry.node.parentNode;
      return { node: parent, offset: Array.prototype.indexOf.call(parent.childNodes, entry.node) };
    }
  }
  const last = entries[entries.length - 1];
  if (last?.node.nodeType === Node.TEXT_NODE) return { node: last.node, offset: last.text.length };
  return { node: host, offset: host.childNodes.length };
}

/** Select the characters [start, end) of the host. */
export function selectChartText(host: HTMLElement, start: number, end = start): void {
  const selection = host.ownerDocument.getSelection();
  if (!selection) return;
  const from = positionAt(host, start);
  const to = positionAt(host, end);
  const range = host.ownerDocument.createRange();
  range.setStart(from.node, from.offset);
  range.setEnd(to.node, to.offset);
  selection.removeAllRanges();
  selection.addRange(range);
  lastSelections.set(host, { start: Math.min(start, end), end: Math.max(start, end) });
}

/** Every character of the host selected (a title entered by keyboard). */
export function selectAllChartText(host: HTMLElement): void {
  const selection = host.ownerDocument.getSelection();
  if (!selection) return;
  const range = host.ownerDocument.createRange();
  range.selectNodeContents(host);
  selection.removeAllRanges();
  selection.addRange(range);
}

/**
 * A selection begun in the host and extended past it (a drag, Shift+click
 * beyond its words) ends at the host's edge: the text is typed in its own
 * box, and an editable page around it (Docs) must never take the keys.
 */
export function keepChartTextSelectionInside(host: HTMLElement): void {
  const selection = host.ownerDocument.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const { anchorNode, anchorOffset, focusNode } = selection;
  if (!anchorNode || !focusNode) return;
  const anchorIn = host.contains(anchorNode);
  const focusIn = host.contains(focusNode);
  if (anchorIn === focusIn) return;
  if (!anchorIn) {
    // Begun outside, ending in the words: not the text's selection.
    return;
  }
  const after = Boolean(host.compareDocumentPosition(focusNode) & Node.DOCUMENT_POSITION_FOLLOWING);
  selection.setBaseAndExtent(anchorNode, anchorOffset, host, after ? host.childNodes.length : 0);
}

type CaretDocument = Document & {
  caretRangeFromPoint?: (x: number, y: number) => Range | null;
  caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
};

/**
 * The caret where a click landed (Office: the click that starts the edit
 * puts the insertion point there). False when the point is not on the
 * host's words, so the caller can select them all instead.
 */
export function placeChartTextCaret(host: HTMLElement, clientX: number, clientY: number): boolean {
  const doc = host.ownerDocument as CaretDocument;
  const at = (x: number, y: number): { node: Node; offset: number } | null => {
    if (typeof doc.caretPositionFromPoint === "function") {
      const position = doc.caretPositionFromPoint(x, y);
      return position ? { node: position.offsetNode, offset: position.offset } : null;
    }
    if (typeof doc.caretRangeFromPoint === "function") {
      const range = doc.caretRangeFromPoint(x, y);
      return range ? { node: range.startContainer, offset: range.startOffset } : null;
    }
    return null;
  };
  let found = at(clientX, clientY);
  if (!found || !host.contains(found.node)) {
    // A click in the title's box beside its words: the nearest point of the
    // words (Office puts the insertion point at the nearest character).
    const box = host.getBoundingClientRect();
    if (box.width > 0 && box.height > 0) {
      const x = Math.min(box.right - 1, Math.max(box.left + 1, clientX));
      const y = Math.min(box.bottom - 1, Math.max(box.top + 1, clientY));
      found = at(x, y);
    }
  }
  if (!found || !host.contains(found.node)) return false;
  const { node, offset } = found;
  const selection = doc.getSelection();
  if (!selection) return false;
  const range = doc.createRange();
  range.setStart(node, offset);
  range.collapse(true);
  selection.removeAllRanges();
  selection.addRange(range);
  return true;
}

/** The `<br>` that shows an empty last line, at the end of `parent`. */
export function appendChartTextSentinel(parent: HTMLElement): void {
  const sentinel = parent.ownerDocument.createElement("br");
  sentinel.setAttribute(SENTINEL_ATTRIBUTE, "true");
  parent.appendChild(sentinel);
}

/**
 * Typed text over the selection in the host: inside the run the caret is in,
 * so it takes that run's format (Office extends the run typed into), a line
 * break a `<br>`. A new line at the very end keeps a sentinel `<br>` after
 * it, without which a browser does not show the empty last line or put the
 * caret on it.
 */
export function insertChartText(host: HTMLElement, text: string): void {
  const doc = host.ownerDocument;
  const selection = doc.getSelection();
  if (!selection || selection.rangeCount === 0) return;
  const range = selection.getRangeAt(0);
  if (!host.contains(range.startContainer) || !host.contains(range.endContainer)) return;
  range.deleteContents();
  let container: Node = range.startContainer;
  let offset = range.startOffset;
  const isRun = (node: Node | undefined): node is HTMLElement =>
    node instanceof HTMLElement && node.hasAttribute(CHART_TEXT_RUN_ATTRIBUTE);
  if (container === host) {
    // Between runs: into the run before the caret (else the one after).
    const before = host.childNodes[offset - 1];
    const after = host.childNodes[offset];
    if (isRun(before)) {
      container = before;
      offset = before.childNodes.length;
    } else if (isRun(after)) {
      container = after;
      offset = 0;
    }
  }
  if (isRun(container)) {
    const before = container.childNodes[offset - 1];
    const after = container.childNodes[offset];
    if (before?.nodeType === Node.TEXT_NODE) {
      container = before;
      offset = (before.textContent ?? "").length;
    } else if (after?.nodeType === Node.TEXT_NODE) {
      container = after;
      offset = 0;
    }
  }
  const caret = doc.createRange();
  caret.setStart(container, offset);
  caret.collapse(true);
  // The words line by line, a `<br>` between (in the run the caret is in).
  text.split("\n").forEach((line, index) => {
    if (index > 0) {
      const br = doc.createElement("br");
      caret.insertNode(br);
      caret.setStartAfter(br);
      caret.collapse(true);
    }
    if (!line) return;
    const start = caret.startContainer;
    if (start.nodeType === Node.TEXT_NODE) {
      const node = start as Text;
      const at = caret.startOffset;
      node.insertData(at, line);
      caret.setStart(node, at + line.length);
    } else {
      const node = doc.createTextNode(line);
      caret.insertNode(node);
      caret.setStartAfter(node);
    }
    caret.collapse(true);
  });
  // A break that ends the words draws no line by itself: a second one
  // makes the new last line, and the caret goes on it.
  for (const stale of Array.from(host.querySelectorAll(`br[${SENTINEL_ATTRIBUTE}]`))) stale.remove();
  const last = caret.startContainer.childNodes[caret.startOffset - 1] ?? null;
  if (text.endsWith("\n") && last instanceof HTMLElement && last.tagName === "BR" && trailingBreak(host, last)) {
    const sentinel = doc.createElement("br");
    sentinel.setAttribute(SENTINEL_ATTRIBUTE, "true");
    caret.insertNode(sentinel);
    caret.setStartBefore(sentinel);
    caret.collapse(true);
  }
  selection.removeAllRanges();
  selection.addRange(caret);
  rememberChartTextSelection(host);
}

/** The editable chart text a ribbon command acts on, in a host's subtree. */
export function editingChartTextHost(root: ParentNode | null | undefined): HTMLElement | null {
  return root?.querySelector<HTMLElement>('[data-office-chart-title][contenteditable="true"]') ?? null;
}
