/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's line-filling rules that Docs reproduces, kept apart from the editor
 * plugins so the text engine's Docs line layout and the plugins that bend
 * the browser's lines (justifiedSpaceShrink) decide by the same numbers.
 */

/** Word's kept-word limits for a justified line (justifiedSpaceShrink.ts:
 * measured in Word 16.86, compatibility mode 15). */
export const WORD_JUSTIFY_SPACE_SHARE = 0.25;
export const WORD_JUSTIFY_WORD_SHARE = 0.345;
/** What the space before the kept word adds to the word's limit, of its own
 * width. Word 16 for Mac, mode 15, Arial 10 and 12 pt, 12 spaces, 0.05 pt
 * steps of the right indent (fixtures-1110/keep-limit, 164 paragraphs): "in"
 * stays 3.684 pt over and goes at 3.734 (10 pt), 4.461 / 4.511 (12 pt); "a"
 * 2.912 / 2.962 and 3.495 / 3.545. With the letters' 0.345 that is 0.3716 to
 * 0.3726 of Arial's space (0.278 em). It was read as 0.085 em on Calibri,
 * Cambria and Times New Roman (spaces of 0.226, 0.22 and 0.25 em: 0.084, 0.082
 * and 0.093 em here, inside those ladders' 0.25 pt steps), which left Arial's
 * limit 0.018 em short (corpus R280-D14: "in", "a" and "EIR;" sent down 0.24,
 * 0.08 and 0.01 px over the old limit, where Word keeps them). */
export const WORD_JUSTIFY_SPACE_BEFORE_SHARE = 0.372;

/** What closes a word and Word lets hang past what the word's letters may:
 * measured for , . ; : ) ” ’ ! ? %; straight quotes and the other closing
 * brackets taken as the same class. A hyphen is the word's own, and an opening
 * mark counts with the letters (probes sq/m5, sq/m6). */
const WORD_JUSTIFY_TRAILING = /[,.;:!?%)\]}"'\u2019\u201d]+$/u;

/** How many UTF-16 units of `word`'s end are trailing punctuation. */
export function wordTrailingPunctuationLength(word: string): number {
  return WORD_JUSTIFY_TRAILING.exec(word)?.[0].length ?? 0;
}

/** Word's share of a kept word's trailing punctuation over the limits of its
 * letters: fewer spaces on the line, more of it, up to all of it (Word 16.86,
 * compatibility mode 15, Times New Roman at 11 and 22 pt, 6 to 50 spaces, kept
 * words "the,", "EIR;" and "survey.", 0.25 pt steps of the right indent; Arial,
 * Calibri and Cambria agree within 0.3 pt; corpus R280-D14: "EIR;" stays on
 * one line 9.19 pt over at 10 pt Arial with 13 spaces). */
export const WORD_JUSTIFY_TRAILING_SHARE = 0.245;
export const WORD_JUSTIFY_TRAILING_PER_SPACES = 9;

/** Whether a justified line overflowing by `overflowPx` keeps its next word
 * in Word: `spaces` are the spaces after the line's last tab (counting the
 * one before the kept word), `spacePx` a space's natural width, `wordPx` the
 * kept word's width and `fontPx` the font size (only whether it is 0 is read), 0 for what follows a hyphen,
 * which has no space before it. `trailingPx` is how much of `wordPx` is
 * trailing punctuation, which is paid for apart from the letters and past the
 * spaces' own limit. A line with no space at all (it can end only after a
 * hyphen) has nothing to narrow, and keeps what follows the hyphen only when
 * its trailing punctuation alone hangs over (sq/m7: , ; . ) % in Times New
 * Roman and Arial 11 pt, no punctuation none). What follows a hyphen on a line
 * with spaces gets less of its punctuation back, as if it had 5/3 the spaces
 * (sq/m7: "x-the," and "x-EIR;" after 3 and 12 spaces). */
export function wordKeepsJustifiedWord(
  overflowPx: number,
  spaces: number,
  fontPx: number,
  spacePx: number,
  wordPx: number,
  trailingPx = 0,
): boolean {
  if (!(overflowPx > 0) || !(wordPx > 0)) return false;
  if (!(spaces > 0)) return overflowPx <= trailingPx;
  if (!(spacePx > 0)) return false;
  const bySpaces = spacePx * spaces * WORD_JUSTIFY_SPACE_SHARE;
  const byWord = (wordPx - trailingPx) * WORD_JUSTIFY_WORD_SHARE + (fontPx > 0 ? WORD_JUSTIFY_SPACE_BEFORE_SHARE * spacePx : 0);
  const counted = fontPx > 0 ? spaces : (spaces * 5) / 3;
  const share = Math.min(1, WORD_JUSTIFY_TRAILING_SHARE + WORD_JUSTIFY_TRAILING_PER_SPACES / counted);
  return overflowPx <= Math.min(bySpaces, byWord) + trailingPx * share;
}
