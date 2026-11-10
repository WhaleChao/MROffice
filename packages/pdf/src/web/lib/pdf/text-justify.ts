/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's Justify for a text box, shared by every writer of a box's glyphs
 * (the managed appearance raster, the flattened export and the API's vector
 * appearance follow the same rule):
 *
 * - every wrapped line is stretched to the box's width EXCEPT the last line of
 *   a paragraph (the line before a hard break, and the box's final line);
 * - a line with spaces between its words takes the extra width in those
 *   spaces; a line without any (CJK text) takes it evenly between its
 *   characters;
 * - trailing whitespace at a wrap is not content and is never stretched.
 */

/** Which wrapped lines end their paragraph. `lines` concatenate, in order,
 * to `text` without its line-break characters. */
export function pdfTextLineParagraphEnds(
  text: string,
  lines: readonly string[],
): boolean[] {
  let cursor = 0;
  return lines.map((line, index) => {
    cursor += line.length;
    let ends = index === lines.length - 1;
    if (text[cursor] === "\r") {
      ends = true;
      cursor += 1;
      if (text[cursor] === "\n") cursor += 1;
    } else if (text[cursor] === "\n") {
      ends = true;
      cursor += 1;
    }
    return ends;
  });
}

export interface PdfJustifiedPiece {
  readonly text: string;
  /** Offset of the piece's left edge from the line's left edge. */
  readonly x: number;
}

/**
 * The pieces of one line laid out to `availableWidth`, or null when the line
 * is drawn as it is (nothing to stretch, or already full).
 */
export function justifyPdfTextLine(
  line: string,
  availableWidth: number,
  measure: (value: string) => number,
): { pieces: PdfJustifiedPiece[]; width: number } | null {
  const content = line.replace(/\s+$/u, "");
  if (!content || !Number.isFinite(availableWidth) || availableWidth <= 0) {
    return null;
  }
  const natural = measure(content);
  const extra = availableWidth - natural;
  if (!Number.isFinite(extra) || extra <= 0.01) return null;
  const tokens = content.match(/\s+|\S+/gu) ?? [];
  const interiorGaps = tokens.filter(
    (token, index) => /^\s/u.test(token) && index > 0 && index < tokens.length - 1,
  ).length;
  const pieces: PdfJustifiedPiece[] = [];
  let x = 0;
  if (interiorGaps > 0) {
    const perGap = extra / interiorGaps;
    tokens.forEach((token, index) => {
      if (/^\s/u.test(token)) {
        x += measure(token) + (index > 0 && index < tokens.length - 1 ? perGap : 0);
        return;
      }
      pieces.push({ text: token, x });
      x += measure(token);
    });
    return { pieces, width: availableWidth };
  }
  const characters = Array.from(content);
  if (characters.length < 2) return null;
  const perGap = extra / (characters.length - 1);
  characters.forEach((character, index) => {
    pieces.push({ text: character, x });
    x += measure(character) + (index < characters.length - 1 ? perGap : 0);
  });
  return { pieces, width: availableWidth };
}
