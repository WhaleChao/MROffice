/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { wordKeepsJustifiedWord, wordTrailingPunctuationLength } from "../../../../../../lib/text-engine/docs-word-rules";

/**
 * A seam left inside a row: the head's last row ends short of the tail's
 * first word, which it now has room for (pageLedgerExtension markStaleSeams
 * for page seams).
 */
/**
 * Whether the head's last row has room for the tail's first word: past the
 * space the row already ends with (it hangs, and its box is the row's), or a
 * space and the word. In a justified paragraph Word narrows the row's spaces
 * to keep a word that overflows by little (`narrows`: compatibility mode 15,
 * wordKeepsJustifiedWord): R283-D05's "They concluded" stood 3.5 px over its
 * edge on one open, and on another "concluded" opened page 3 and the seam,
 * a space-width short of "room", was never re-cut.
 */
export function headRowTakesTailWord(headDom: HTMLElement, tailDom: HTMLElement, narrows = false): boolean {
  const headText = headDom.textContent ?? "";
  const tailText = tailDom.textContent ?? "";
  // Cut inside a word (a hyphenated break): the joined row is not read here.
  if (!/\s$/.test(headText) && !/^\s/.test(tailText)) return false;
  const style = getComputedStyle(headDom);
  if (style.direction === "rtl" || /^(vertical|sideways)/.test(style.writingMode)) return false;
  const doc = headDom.ownerDocument;
  const box = headDom.getBoundingClientRect();
  if (!(box.width > 0) || !(headDom.offsetWidth > 0)) return false;
  const scale = box.width / headDom.offsetWidth;
  const range = doc.createRange();
  range.selectNodeContents(headDom);
  const rects = Array.from(range.getClientRects()).filter((rect) => rect.width > 0 && rect.height > 0);
  if (rects.length === 0) return false;
  const lastTop = Math.max(...rects.map((rect) => rect.top));
  const lastRight = Math.max(...rects.filter((rect) => Math.abs(rect.top - lastTop) < 2 * scale).map((rect) => rect.right));
  const contentRight = box.right - ((Number.parseFloat(style.paddingRight) || 0) + (Number.parseFloat(style.borderRightWidth) || 0)) * scale;
  const room = contentRight - lastRight;
  if (!(room > 0)) return false;
  // The tail's first word, as drawn.
  const walker = doc.createTreeWalker(tailDom, NodeFilter.SHOW_TEXT);
  let node: Node | null;
  while ((node = walker.nextNode())) {
    const data = (node as Text).data;
    if (/^\s*$/.test(data)) continue;
    const match = /^(\s*)(\S+)/.exec(data);
    if (!match) return false;
    const word = doc.createRange();
    word.setStart(node, match[1].length);
    word.setEnd(node, match[1].length + match[2].length);
    const wordRects = Array.from(word.getClientRects()).filter((rect) => rect.width > 0);
    // A word the tail wraps over two rows of its own is no word a row takes.
    if (wordRects.length !== 1) return false;
    const canvas = doc.createElement("canvas").getContext("2d");
    const parent = (node as Text).parentElement;
    if (!canvas || !parent) return false;
    const tailStyle = getComputedStyle(parent);
    canvas.font = `${tailStyle.fontStyle} ${tailStyle.fontWeight} ${tailStyle.fontSize} ${tailStyle.fontFamily}`;
    const space = canvas.measureText(" ").width * scale;
    const wordPx = wordRects[0].width;
    const row = lastRowSpaces(headDom, lastTop, scale);
    const needed = wordPx + (row.endsWithSpace ? 0 : space);
    if (room >= needed + 0.25 * scale) return true;
    if (!narrows || style.textAlign !== "justify" || !(row.spaces + (row.endsWithSpace ? 0 : 1) > 0)) return false;
    const trailing = wordTrailingPunctuationLength(match[2]);
    const trailingPx = trailing > 0 ? canvas.measureText(match[2].slice(-trailing)).width * scale : 0;
    const fontPx = (Number.parseFloat(tailStyle.fontSize) || 0) * scale;
    return wordKeepsJustifiedWord(needed - room, row.spaces + (row.endsWithSpace ? 0 : 1), fontPx, space, wordPx, trailingPx);
  }
  return false;
}

/** The spaces on the head's last row (the one it ends with among them), and
 * whether it ends with one. Read back from the row's end. */
function lastRowSpaces(headDom: HTMLElement, lastTop: number, scale: number): { spaces: number; endsWithSpace: boolean } {
  const doc = headDom.ownerDocument;
  const texts: Text[] = [];
  const walker = doc.createTreeWalker(headDom, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) texts.push(node as Text);
  const range = doc.createRange();
  let spaces = 0;
  let endsWithSpace = false;
  let seenInk = false;
  for (let t = texts.length - 1; t >= 0; t -= 1) {
    const data = texts[t].data;
    for (let i = data.length - 1; i >= 0; i -= 1) {
      range.setStart(texts[t], i);
      range.setEnd(texts[t], i + 1);
      const box = Array.from(range.getClientRects()).find((rect) => rect.width > 0 || rect.height > 0);
      if (!box) continue;
      if (Math.abs(box.top - lastTop) >= 2 * scale) return { spaces, endsWithSpace };
      if (/\s/.test(data[i])) {
        spaces += 1;
        if (!seenInk && box.width > 0) endsWithSpace = true;
      } else {
        seenInk = true;
      }
    }
  }
  return { spaces, endsWithSpace };
}

