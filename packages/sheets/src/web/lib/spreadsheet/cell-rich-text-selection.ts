/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Map a DOM Range (a user text selection inside the rich-text cell editor) to
 * CHARACTER OFFSETS in the cell's concatenated text, so the offsets can drive the
 * pure {@link applyRunFormat}. Walks the container's text nodes in document order
 * and accumulates their lengths — independent of how the runs were split into
 * spans, so re-rendering after a format never shifts the mapping.
 *
 * DOM-only (no React); jsdom-testable via document.createRange + TreeWalker.
 */

/** The cumulative character offset of (node, nodeOffset) within `container`. */
function charOffsetOf(
  container: HTMLElement,
  node: Node,
  nodeOffset: number,
): number {
  // A selection endpoint can land on an ELEMENT (offset = child index) rather
  // than a text node; resolve it to the text position by summing the text of the
  // children before that index.
  if (node.nodeType === Node.ELEMENT_NODE) {
    let elementBase = 0;
    const children = Array.from(node.childNodes).slice(0, nodeOffset);
    for (const child of children) {
      elementBase += (child.textContent ?? "").length;
    }
    // Add the text of every node BEFORE `node` in the container.
    return charOffsetBeforeNode(container, node) + elementBase;
  }
  // Text node: sum every text node strictly before it, then add the local offset.
  let offset = 0;
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  let current = walker.nextNode();
  while (current) {
    if (current === node) {
      return offset + Math.min(nodeOffset, (current.textContent ?? "").length);
    }
    offset += (current.textContent ?? "").length;
    current = walker.nextNode();
  }
  return offset;
}

/** Total text length of every text node that comes before `node` in `container`. */
function charOffsetBeforeNode(container: HTMLElement, node: Node): number {
  let offset = 0;
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  let current = walker.nextNode();
  while (current) {
    if (node.contains(current)) return offset;
    offset += (current.textContent ?? "").length;
    current = walker.nextNode();
  }
  return offset;
}

/**
 * The `[start, end)` character offsets of a Range within `container`, normalized
 * so start ≤ end. Returns null when the range is collapsed (no selection) — a
 * format with no selection is a no-op the caller skips.
 */
export function rangeToCharOffsets(
  container: HTMLElement,
  range: Range,
): [number, number] | null {
  const a = charOffsetOf(container, range.startContainer, range.startOffset);
  const b = charOffsetOf(container, range.endContainer, range.endOffset);
  const start = Math.min(a, b);
  const end = Math.max(a, b);
  if (start === end) return null;
  return [start, end];
}

/**
 * {@link rangeToCharOffsets} for an editable box, where a collapsed range is
 * the caret and still names a position: `[caret, caret]`.
 */
export function rangeToCaretOffsets(
  container: HTMLElement,
  range: Range,
): [number, number] {
  const a = charOffsetOf(container, range.startContainer, range.startOffset);
  const b = charOffsetOf(container, range.endContainer, range.endOffset);
  return [Math.min(a, b), Math.max(a, b)];
}

/**
 * The DOM position of character offset `offset` inside `container`: the text
 * node it falls in (the earlier node at a boundary, so a caret stays inside
 * the run it follows), or the container's end.
 */
function domPositionOf(container: HTMLElement, offset: number): [Node, number] {
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  let remaining = Math.max(0, offset);
  let current = walker.nextNode();
  let last: Node | null = null;
  while (current) {
    const length = (current.textContent ?? "").length;
    if (remaining <= length) return [current, remaining];
    remaining -= length;
    last = current;
    current = walker.nextNode();
  }
  return last
    ? [last, (last.textContent ?? "").length]
    : [container, container.childNodes.length];
}

/** Select the characters `[start, end)` of `container` (a caret when equal). */
export function selectCharOffsets(
  container: HTMLElement,
  start: number,
  end: number,
): void {
  const selection = typeof window !== "undefined" ? window.getSelection() : null;
  if (!selection) return;
  const range = document.createRange();
  const [startNode, startOffset] = domPositionOf(container, Math.min(start, end));
  const [endNode, endOffset] = domPositionOf(container, Math.max(start, end));
  range.setStart(startNode, startOffset);
  range.setEnd(endNode, endOffset);
  selection.removeAllRanges();
  selection.addRange(range);
}
