/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PdfTextLayerSelectionOffsets {
  readonly start: number;
  readonly end: number;
  readonly backward: boolean;
}

function textNodesOf(layer: Element): Text[] {
  const nodes: Text[] = [];
  const walker = layer.ownerDocument.createTreeWalker(layer, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) nodes.push(node as Text);
  return nodes;
}

function offsetOf(layer: Element, nodes: readonly Text[], container: Node, offset: number): number | null {
  if (!layer.contains(container)) return null;
  let total = 0;
  if (container.nodeType === 3) {
    for (const node of nodes) {
      if (node === container) return total + offset;
      total += node.data.length;
    }
    return null;
  }
  // An element boundary: the text before its offset-th child.
  const child = container.childNodes[offset] ?? null;
  for (const node of nodes) {
    if (child && (child === node || child.contains(node))) return total;
    total += node.data.length;
  }
  return child ? null : total;
}

/** The selection in `layer` as offsets into its text, or null when the
 * selection is not (wholly) in it. */
export function readPdfTextLayerSelection(layer: Element): PdfTextLayerSelectionOffsets | null {
  const selection = layer.ownerDocument.defaultView?.getSelection();
  if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return null;
  const range = selection.getRangeAt(0);
  const nodes = textNodesOf(layer);
  const start = offsetOf(layer, nodes, range.startContainer, range.startOffset);
  const end = offsetOf(layer, nodes, range.endContainer, range.endOffset);
  if (start === null || end === null || end <= start) return null;
  const anchor = selection.anchorNode;
  const backward = Boolean(anchor && anchor === range.endContainer && selection.anchorOffset === range.endOffset &&
    !(anchor === range.startContainer && selection.anchorOffset === range.startOffset));
  return { start, end, backward };
}

function pointAt(nodes: readonly Text[], offset: number): { node: Text; offset: number } | null {
  let total = 0;
  for (const node of nodes) {
    const length = node.data.length;
    if (offset <= total + length) return { node, offset: offset - total };
    total += length;
  }
  return null;
}

/** Set the selection again on `layer`'s (new) spans; false when its text
 * no longer reaches those offsets. */
export function restorePdfTextLayerSelection(layer: Element, saved: PdfTextLayerSelectionOffsets): boolean {
  const selection = layer.ownerDocument.defaultView?.getSelection();
  if (!selection) return false;
  const nodes = textNodesOf(layer);
  const start = pointAt(nodes, saved.start);
  const end = pointAt(nodes, saved.end);
  if (!start || !end) return false;
  try {
    if (saved.backward && typeof selection.setBaseAndExtent === "function") {
      selection.setBaseAndExtent(end.node, end.offset, start.node, start.offset);
    } else {
      const range = layer.ownerDocument.createRange();
      range.setStart(start.node, start.offset);
      range.setEnd(end.node, end.offset);
      selection.removeAllRanges();
      selection.addRange(range);
    }
    return true;
  } catch {
    return false;
  }
}

/** Replace `layer`'s spans, keeping the reader's selection in it. */
export function replacePdfTextLayerChildren(layer: Element, nodes: readonly Node[]): void {
  const saved = readPdfTextLayerSelection(layer);
  layer.replaceChildren(...nodes);
  if (saved) restorePdfTextLayerSelection(layer, saved);
}
