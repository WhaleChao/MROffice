/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The words of an editable text box as the reader sees them laid out: every
 * place the box wrapped a line on its own (a soft wrap at its width) becomes
 * a line break, so a box written into the page as text keeps the lines it
 * showed while it was typed (PDF 0060: a two-line box saved as one long line
 * that ran off the paper).
 *
 * Nothing is removed: the space a line wrapped after stays at the end of its
 * line, so taking the inserted breaks out gives back the box's text exactly.
 * Explicit breaks ("\n" in the box's text) are kept as they are. Returns null
 * when the box is not laid out in horizontal lines that can be measured.
 */
export function pdfEditableVisualLineText(root: HTMLElement): string | null {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return null;
  const writingMode = view.getComputedStyle(root).writingMode;
  if (writingMode && writingMode.startsWith("vertical")) return null;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  let out = "";
  let lineTop: number | null = null;
  let lineBottom = 0;
  let measured = 0;
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    let offset = 0;
    for (const character of Array.from(node.data)) {
      const start = offset;
      offset += character.length;
      if (character === "\n") {
        out += character;
        lineTop = null;
        continue;
      }
      // A space never starts a line: the one a line wrapped after hangs at
      // its end, wherever the browser reports it.
      if (/\s/.test(character)) {
        out += character;
        continue;
      }
      range.setStart(node, start);
      range.setEnd(node, offset);
      const rect = range.getClientRects()[0];
      if (!rect || rect.height <= 0) {
        out += character;
        continue;
      }
      measured += 1;
      if (lineTop === null) {
        lineTop = rect.top;
        lineBottom = rect.bottom;
      } else if (
        // A glyph that starts below the middle of the line in hand is on the
        // next line: the box wrapped there.
        rect.top > lineBottom - Math.min(rect.height, lineBottom - lineTop) * 0.5
      ) {
        out += "\n";
        lineTop = rect.top;
        lineBottom = rect.bottom;
      } else {
        lineTop = Math.min(lineTop, rect.top);
        lineBottom = Math.max(lineBottom, rect.bottom);
      }
      out += character;
    }
  }
  range.detach?.();
  return measured > 0 ? out : null;
}

/**
 * `text` with the soft wraps `root` shows, when `root` holds exactly that
 * text once its inserted breaks are ignored; otherwise `text` unchanged.
 * `normalize` is the clean-up the caller already applied to `text`.
 */
export function pdfSoftWrappedBoxText(
  root: HTMLElement | null | undefined,
  text: string,
  normalize: (value: string) => string,
): string {
  if (!root) return text;
  if (normalize(root.textContent ?? "") !== text) return text;
  const visual = pdfEditableVisualLineText(root);
  if (visual === null) return text;
  const wrapped = normalize(visual);
  // The wrapped text differs from the box's only by the breaks it inserted.
  const inserted = wrapped.length - text.length;
  if (inserted <= 0) return text;
  let i = 0;
  let j = 0;
  while (i < wrapped.length && j < text.length) {
    if (wrapped[i] === text[j]) {
      i += 1;
      j += 1;
    } else if (wrapped[i] === "\n") {
      i += 1;
    } else {
      return text;
    }
  }
  return j === text.length && /^\n*$/.test(wrapped.slice(i)) ? wrapped.replace(/\n+$/, "") : text;
}
