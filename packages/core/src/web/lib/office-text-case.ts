/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Change Case > Sentence case, one reading for Docs and the deck: every letter
 * lower-cased, then the first letter of each sentence upper-cased. A sentence
 * starts at the text's start or after `.`, `!`, `?` (and the full-width
 * 。！？), and the quotes, brackets and spaces that lead into it are skipped,
 * so `"HELLO WORLD."` reads `"Hello world."` as Word and PowerPoint write it
 * (PPT-B144: an opening quote kept the first letter lower-case).
 */

/** What may stand between a sentence's start and its first letter. */
const LEAD = String.raw`\s"'“”‘’«»‹›「」『』〈〉《》【】()（）\[\]{}`;

const SENTENCE_START = new RegExp(
  String.raw`(^[${LEAD}]*|[.!?。！？][${LEAD}]*)(\p{L})`,
  "gu",
);

export function officeSentenceCase(text: string): string {
  return text
    .toLowerCase()
    .replace(SENTENCE_START, (_whole, lead: string, letter: string) => lead + letter.toUpperCase());
}
