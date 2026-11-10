/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The line breaks of a content append's words (content-append-wrap.ts): at
 * spaces, between CJK characters, and inside a word only when the word
 * alone is wider than the line; spaces hang past the edge, and a line does
 * not start with the spaces it broke at. Each piece is measured with
 * `measure` -- the planner's glyph advances on the server, the browser's
 * measure of the same face in the editor.
 */

/** CJK ideographs, kana, hangul and their punctuation: a line may break
 * before or after any of them. */
const CJK = /[⺀-鿿가-힯豈-﫿＀-￯　-〿]/u;

export function breakPdfContentAppendLines(
  pieces: readonly Readonly<{ text: string; width: number }>[],
  maxWidth: number,
): string[] {
  const lines: string[] = [];
  let start = 0;
  while (start < pieces.length) {
    let width = 0;
    let end = start;
    // The last place the line may end (exclusive piece index), if any.
    let lastBreak = -1;
    while (end < pieces.length) {
      const piece = pieces[end]!;
      const isSpace = /^\s+$/u.test(piece.text);
      // Spaces hang past the edge; anything else must fit.
      if (!isSpace && width + piece.width > maxWidth && end > start) break;
      width += piece.width;
      end += 1;
      if (isSpace) lastBreak = end;
      else if (CJK.test(piece.text) || (end < pieces.length && CJK.test(pieces[end]!.text))) lastBreak = end;
    }
    const stop = end >= pieces.length || lastBreak <= start ? end : lastBreak;
    lines.push(pieces.slice(start, stop).map((piece) => piece.text).join("").trimEnd());
    start = stop;
    while (start < pieces.length && /^\s+$/u.test(pieces[start]!.text)) start += 1;
  }
  return lines;
}

/** The text broken into the lines `maxWidth` holds, each character measured
 * by `measure`; its own line breaks kept, an empty paragraph kept empty. */
export function wrapPdfContentAppendTextByMeasure(
  text: string,
  maxWidth: number,
  measure: (character: string) => number,
): string {
  if (!(maxWidth > 0)) return text;
  return text.split(/\r?\n/u).map((paragraph) => paragraph.trim()
    ? breakPdfContentAppendLines(Array.from(paragraph, (character) => ({ text: character, width: measure(character) })), maxWidth).join("\n")
    : "").join("\n");
}
