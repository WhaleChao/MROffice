/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure readability analysis for the document editor — Microsoft Word's
 * Review ▸ Word Count dialog plus its "Show readability statistics" panel
 * (Flesch Reading Ease + Flesch–Kincaid Grade Level), and a modern reading-time
 * estimate. Dependency-free and DOM-free so the counting + scoring is
 * exhaustively unit-testable; the toolbar wires it over `editor.getText()`.
 *
 * The syllable count is the well-known English heuristic (vowel-group counting
 * with silent-trailing-e correction). It is an estimate — exactly as Word's own
 * readability scores are estimates — not a dictionary lookup.
 */

import { characterCount, splitWordCount, withoutWordSpaces } from "./document-word-count";

export interface ReadabilityStats {
  /** Word tokens (runs of letters/digits). */
  words: number;
  /** CJK characters, which WPS reports separately for Chinese documents. */
  cjkCharacters: number;
  /** Latin (whitespace-delimited) words only, the other half of that split. */
  latinWords: number;
  /** Characters excluding whitespace. */
  characters: number;
  /** Characters including whitespace (trimmed ends). */
  charactersWithSpaces: number;
  /** Sentence count (terminal . ! ? groups; ≥1 when there is any text). */
  sentences: number;
  /** Paragraph count (non-empty block lines; ≥1 when there is any text). */
  paragraphs: number;
  /** Estimated syllables across all words. */
  syllables: number;
  avgWordsPerSentence: number;
  avgCharsPerWord: number;
  avgSyllablesPerWord: number;
  avgSentencesPerParagraph: number;
  /** Flesch Reading Ease, 0–100 (higher = easier), clamped for display. */
  fleschReadingEase: number;
  /** Flesch–Kincaid US grade level, ≥0. */
  fleschKincaidGrade: number;
  /** Estimated reading time in minutes at {@link READING_WPM}. */
  readingTimeMinutes: number;
}

/** Average adult silent-reading speed (words per minute) for the time estimate. */
export const READING_WPM = 200;

/**
 * Estimate the syllables in a single word with the standard English heuristic:
 * count vowel groups, drop a silent trailing "e" (but keep a "-le" syllable),
 * and floor at 1 for any word containing a letter.
 */
export function countSyllables(word: string): number {
  const w = word.toLowerCase().replace(/[^a-z]/g, "");
  if (!w) return 0;
  if (w.length <= 3) return 1;
  // Strip a silent trailing e / es / ed — but NOT a "-le" ending (table, little),
  // which keeps its own syllable.
  const trimmed = w
    .replace(/(?:[^laeiouy]e|ed|[^laeiouy]es)$/, "")
    .replace(/^y/, "");
  const groups = trimmed.match(/[aeiouy]{1,2}/g);
  return groups ? Math.max(1, groups.length) : 1;
}

function round1(n: number): number {
  return Number.isFinite(n) ? Math.round(n * 10) / 10 : 0;
}

/** Analyze plain document text into Word-parity counts + readability scores. */
export function analyzeReadability(text: string): ReadabilityStats {
  // Only the paragraph marks around the text go: Word counts a leading or
  // trailing space as a character with spaces, and an em space is a word
  // (String.prototype.trim took both).
  const trimmed = text.replace(/^[\n\r]+|[\n\r]+$/gu, "");
  const wordTokens = trimmed.match(/[\p{L}\p{N}]+(?:['’][\p{L}\p{N}]+)*/gu) ?? [];
  // Han, kana and Hangul are all `\p{L}`, so a run of them matched as ONE
  // token: this reported 2 words for a 24-character Chinese sentence and 1 for
  // a Japanese one. Word and WPS count each CJK character as a word, which is
  // what `countWords` implements, so the user-facing number comes from there.
  const counted = splitWordCount(trimmed);
  const words = counted.words;
  // Characters as the reader sees them, by the one rule `countWords` and
  // the status bar use: "é" typed as e + U+0301 is one character, as the
  
  const characters = characterCount(withoutWordSpaces(trimmed));
  // The text comes one paragraph per line; the line breaks between them are
  // Word's paragraph marks, which Word does not count as characters (a doc of
  // "a" and "b" has 2 with spaces, not 3).
  const charactersWithSpaces = characterCount(trimmed.replace(/\n/gu, ""));

  // Sentences: runs of terminal punctuation, which closing quotes or brackets
  // may follow (“I came.” ends a sentence), and the full-width 。！？ that
  
  // decimal point (1.5) is not followed by a space and so ends nothing. Floor
  // at 1 when there is any word so a punctuation-free fragment still scores
  // as a single sentence.
  const sentenceMatches =
    trimmed.match(
      /[.!?]+["'”’»›)\]}」』）】》〉]*(?=\s|$|[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}])|[。！？｡]+["'”’»›)\]}」』）】》〉]*/gu,
    )?.length ?? 0;
  const sentences = Math.max(sentenceMatches, words > 0 ? 1 : 0);

  // Paragraphs: the block lines (getText joins blocks with newlines) that
  // hold a word. Word does not count a paragraph of "—", "• " or "– —", nor
  // one whose only text is a bullet, while "," and "2." alone are a word and
  // a paragraph each (Word 16.86, `compute statistics`).
  const paragraphMatches = trimmed
    .split(/\n+/u)
    .filter((line) => splitWordCount(line).words > 0).length;
  const paragraphs = Math.max(paragraphMatches, words > 0 ? 1 : 0);

  let syllables = 0;
  for (const token of wordTokens) syllables += countSyllables(token);

  // Flesch is defined over English-shaped words and syllables, so it keeps the
  // Latin token count as its denominator. Feeding it one unit per Han character
  // would not make it right for Chinese — it would make it differently wrong,
  // and it would move the score of every English document that happens to
  // contain a Chinese name.
  const scoringWords = wordTokens.length;
  const scoringWordsPerSentence = sentences > 0 ? scoringWords / sentences : 0;
  const avgSyllablesPerWord = scoringWords > 0 ? syllables / scoringWords : 0;
  // The averages the dialog SHOWS are of the counts it shows: Words per
  // Sentence is Words / Sentences, so "well-known." (one word, one sentence)
  
  const avgWordsPerSentence = sentences > 0 ? words / sentences : 0;
  const avgCharsPerWord = words > 0 ? characters / words : 0;
  const avgSentencesPerParagraph = paragraphs > 0 ? sentences / paragraphs : 0;

  let fleschReadingEase = 0;
  let fleschKincaidGrade = 0;
  if (scoringWords > 0 && sentences > 0) {
    fleschReadingEase =
      206.835 - 1.015 * scoringWordsPerSentence - 84.6 * avgSyllablesPerWord;
    fleschReadingEase = Math.max(0, Math.min(100, fleschReadingEase));
    fleschKincaidGrade =
      0.39 * scoringWordsPerSentence + 11.8 * avgSyllablesPerWord - 15.59;
    fleschKincaidGrade = Math.max(0, fleschKincaidGrade);
  }

  return {
    words,
    cjkCharacters: counted.cjkCharacters,
    latinWords: counted.latinWords,
    characters,
    charactersWithSpaces,
    sentences,
    paragraphs,
    syllables,
    avgWordsPerSentence: round1(avgWordsPerSentence),
    avgCharsPerWord: round1(avgCharsPerWord),
    avgSyllablesPerWord: round1(avgSyllablesPerWord),
    avgSentencesPerParagraph: round1(avgSentencesPerParagraph),
    fleschReadingEase: round1(fleschReadingEase),
    fleschKincaidGrade: round1(fleschKincaidGrade),
    readingTimeMinutes: words > 0 ? round1(words / READING_WPM) : 0,
  };
}
