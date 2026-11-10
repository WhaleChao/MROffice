/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The marker a paragraph of PDF text starts with, when it is a list item
 * (lib/pdf/text-list.ts): a bullet glyph or a numbering label, and the blank
 * after it. Kept free of the editors' libraries so the native engine reads a
 * marker the same way the ribbon writes one.
 */

/** The bullet glyphs a paragraph may start with (the presentation library's,
 * and the ones a PDF's own lists are commonly set in). */
export const PDF_TEXT_LIST_BULLET_GLYPHS: ReadonlySet<string> = new Set([
  "•", "◦", "▪", "▫", "●", "○", "■", "□", "◆", "◇", "❑", "✓", "➢", "–", "-", "*", "·",
]);

/** A numbering label at a paragraph's start: 1. 1) a. A) i. IV. ① 一． 一、 (1).
 * A bare number is no label (Word's always has its stop or brackets): a
 * line opening "2027 年" is a year. */
export const PDF_TEXT_LIST_NUMBER_LABEL =
  /^(?:\(\d{1,4}\)|\d{1,4}[.)．、]|[a-zA-Z][.)]|[ivxlcdmIVXLCDM]{1,6}[.)]|[①-⑳㉑-㉟㊱-㊿]|[一二三四五六七八九十百]{1,4}[．、.])(?=[ \t　])/u;

/** A paragraph's own marker and the gap after it, when it starts with one. */
export function pdfTextListMarkerOf(paragraph: string): string | null {
  const glyph = [...paragraph][0];
  if (glyph && PDF_TEXT_LIST_BULLET_GLYPHS.has(glyph) && /^[ \t　]/u.test(paragraph.slice(glyph.length))) {
    return paragraph.slice(0, glyph.length + 1);
  }
  const label = PDF_TEXT_LIST_NUMBER_LABEL.exec(paragraph);
  return label ? paragraph.slice(0, label[0].length + 1) : null;
}
