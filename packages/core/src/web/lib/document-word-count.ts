/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word Count (审阅 ▸ 字数统计).
 *
 * WPS and Word report the same six numbers for a document or a selection, and
 * they count CJK differently from Latin: every Han/Kana/Hangul character is one
 * "word" on its own, while Latin words are whitespace-delimited runs. A mixed
 * paragraph therefore counts as the sum of both, which is what makes a Chinese
 * document's word count match the character count and an English one's not.
 *
 * Pure: text in, counts out. The host decides what to feed it (whole document
 * or selection, with or without footnotes), so the same counter serves the
 * dialog, the status bar, and tests.
 */

export interface WordCountInput {
  /** Body paragraphs, in order. */
  paragraphs: readonly string[];
  /** Footnote and endnote texts, counted only when `includeNotes`. */
  notes?: readonly string[];
  /** Word's "include footnotes and endnotes" checkbox. */
  includeNotes?: boolean;
}

export interface WordCountResult {
  /** Latin words plus individually-counted CJK characters. */
  words: number;
  /** Every character, including spaces inside paragraphs. */
  charactersWithSpaces: number;
  /** Characters excluding whitespace. */
  charactersWithoutSpaces: number;
  /** Non-empty paragraphs, as Word counts them. */
  paragraphs: number;
  /** CJK characters, which WPS shows separately for Chinese documents. */
  cjkCharacters: number;
  /** Latin (whitespace-delimited) words only. */
  latinWords: number;
}

// Han, Hiragana/Katakana, Hangul, and CJK punctuation that reads as a
// character. The two regexes below are exact complements — one counts what the
// other excludes — so the ranges live here once rather than being written out
// twice and drifting apart.
//
// The BMP ranges deliberately START at U+3001, one past the IDEOGRAPHIC SPACE
// (U+3000): a fullwidth space is whitespace, and counting it as a character
// inflated the word count of every document typed with them.
//
// The supplementary-plane blocks are Han extensions B through H plus the
// compatibility supplement. They are not decoration: rare characters in names,
// Hong Kong and Taiwan character sets, and classical text live there, and
// without them a paragraph of 𠀀𪜀 counted as one Latin WORD and zero Chinese
// characters — the opposite of both numbers WPS reports.
const CJK_RANGES =
  "、-ヿ㐀-䶿一-鿿가-힯豈-﫿＀-￯" +
  "\u{20000}-\u{2a6df}\u{2a700}-\u{2ebef}\u{2f800}-\u{2fa1f}\u{30000}-\u{3134f}";
const CJK_RE = new RegExp(`[${CJK_RANGES}]`, "gu");
// Word splits words at an en dash, an em dash and a bullet as it does at a
// space, and a run of them alone is no word ("a — b" and "a—b" are 2, "—" is
// 0, while "a ‒ b", "a ― b", "a - b" and "a − b" are 3; Word 16.86). They
// stay characters.
const WORD_SEPARATORS = "\u2013\u2014\u2022\ufdd0";
/**
 * What the counted text holds for a footnote or endnote reference mark: Word
 * counts the mark as one character and no word, and it splits the words
 * around it ("a<mark>b" is 2 words of 3 characters; Word 16.86). U+FDD0 is a
 * Unicode noncharacter, so no document text can be mistaken for it.
 */
export const WORD_NOTE_MARK = "\ufdd0";
// What Word counts as a space (Word 16.86, `compute statistics` on one
// paragraph per character, alone and between letters): the space, the tab,
// the no-break space, the four-per-em space and the ideographic space split
// words and are no "characters (no spaces)". Every other Unicode space -- the
// en and em quads and spaces, the three- and six-per-em, figure, punctuation,
// thin, hair, narrow no-break, medium mathematical and Ogham spaces, the line
// separator -- is an ordinary character: "a\u2003b" is one word of three
// characters and an em space alone is a word. The line breaks are the
// paragraph marks the counted text is joined with.
const WORD_SPACES = " \\t\\n\\r\\v\\f\\u00a0\\u2005\\u3000";
const WORD_SPACE_RE = new RegExp(`[${WORD_SPACES}]`, "gu");
// The zero width space and the byte order mark are no character at all:
// "a\u200bb" is one word of two characters, and either alone is nothing.
const UNCOUNTED_RE = /[\u200b\ufeff]/gu;
const LATIN_WORD_RE = new RegExp(`[^${WORD_SPACES}${WORD_SEPARATORS}${CJK_RANGES}]+`, "gu");

/** The text with Word's spaces and its uncounted characters taken out. */
export function withoutWordSpaces(text: string): string {
  return text.replace(WORD_SPACE_RE, "").replace(UNCOUNTED_RE, "");
}

/**
 * Characters as a READER counts them, not as the encoding stores them. "é"
 * written with a combining acute is two code points and one character; a flag
 * is two, a ZWJ family emoji five. Word and WPS both report what the reader
 * sees, so grapheme clusters are used wherever the runtime offers them.
 */
const GRAPHEMES =
  typeof Intl !== "undefined" && "Segmenter" in Intl
    ? new Intl.Segmenter(undefined, { granularity: "grapheme" })
    : null;

/**
 * The characters of a text as a reader sees them: one entry per grapheme
 * cluster ("é" with a combining acute is one, a flag emoji is one), falling
 * back to code points where the runtime has no segmenter.
 */
export function graphemeClusters(text: string): string[] {
  if (!GRAPHEMES) return [...text];
  return Array.from(GRAPHEMES.segment(text), (item) => item.segment);
}

/**
 * Code points that can join their neighbours into one reader-visible
 * character: combining marks, ZWJ, variation selectors, emoji (surrogate
 * pairs) and their modifiers, regional-indicator flags, conjoining jamo.
 * Text without any is counted one code point per character, which is what
 * the grapheme pass would return, without its cost over a whole book.
 */
const CLUSTER_FORMING_RE = /[\p{M}\u200d\ufe00-\ufe0f\u1100-\u11ff\u{10000}-\u{10ffff}]/u;

export function characterCount(text: string): number {
  const counted = text.replace(UNCOUNTED_RE, "");
  if (!GRAPHEMES || !CLUSTER_FORMING_RE.test(counted)) return [...counted].length;
  // Segmenter yields one entry per grapheme cluster; only the tally matters.
  return [...GRAPHEMES.segment(counted)].length;
}


/**
 * Just the word split, without the grapheme pass the character counts need.
 *
 * The status bar recomputes on a debounce over the WHOLE document — ~200ms on
 * an 876-page book — so it must not pay for `Intl.Segmenter` over text whose
 * character counts it is not going to show. Same rule as `countWords`, which
 * is the point: there is one definition of a word here, not one per caller.
 */
export function splitWordCount(text: string): {
  words: number;
  cjkCharacters: number;
  latinWords: number;
} {
  const cjkCharacters = (text.match(CJK_RE) ?? []).length;
  // Every whitespace-delimited run counts, including one made only of
  // punctuation. This looks wrong and is not: it is what Word does.
  //
  // MEASURED, against Microsoft Word for Mac driven over AppleScript
  // (`compute statistics … statistic words` — the engine behind 字数统计), one
  // generated .docx per case:
  //
  //     "a , b" → 3      "a - b" → 3      "a / b" → 3      "-" → 1
  //     "a ... b" → 3    "a -- b" → 3     "a ( b ) c" → 5  "- / ," → 3
  //     "中 - 文" → 3    "中文 --- 中文" → 5
  //
  // and, for runs that are NOT space-separated, the run still counts once:
  //
  //     "a-b" → 1        "abc,def" → 1    "中文,中文" → 5   "word ." → 2
  //
  // A `[\p{L}\p{N}]` guard used to drop the punctuation-only runs, which cost a
  // word on every "01 - 名称 / Name" line — a shape this product's own fixtures
  // are full of — and reported 0 for a selection of "- / ,". The guard was
  // written from an assumption about what a word is; the numbers above are what
  // the application we claim parity with actually returns.
  const latinWords = (text.replace(UNCOUNTED_RE, "").match(LATIN_WORD_RE) ?? []).length;
  return { words: cjkCharacters + latinWords, cjkCharacters, latinWords };
}

/** Lines already counted, by their text (most recent last). */
const statusLineCounts = new Map<string, { words: number; characters: number; joins: boolean }>();
const STATUS_LINE_MEMORY = 50_000;

/**
 * The status bar's two numbers for `text` (words as `splitWordCount`,
 * characters as `characterCount(withoutWordSpaces(text))`), counted a line
 * at a time and kept by line: a long document's text is the same after a
 * key typed but for one paragraph, and the whole-document count (the
 * grapheme pass over a book of math symbols: ~0.8 s in a dev build) was paid
 * again whenever the typing paused. A line break separates words, and Word's
 * spaces (the line break among them) are not characters, so the counts of the
 * lines add up, unless a line's first counted character joins the character
 * before it (a combining mark, a joiner, an emoji or flag half, a conjoining
 * jamo): such a text is counted whole, as before.
 */
export function statusBarCounts(text: string): { words: number; characters: number } {
  const lines = text.split("\n");
  let words = 0;
  let characters = 0;
  let counted = false;
  for (const line of lines) {
    let entry = statusLineCounts.get(line);
    if (entry) {
      statusLineCounts.delete(line);
    } else {
      const stripped = withoutWordSpaces(line);
      entry = {
        words: splitWordCount(line).words,
        characters: characterCount(stripped),
        joins: stripped.length > 0 && CLUSTER_FORMING_RE.test(String.fromCodePoint(stripped.codePointAt(0)!)),
      };
    }
    statusLineCounts.set(line, entry);
    if (entry.joins && counted) return { words: splitWordCount(text).words, characters: characterCount(withoutWordSpaces(text)) };
    if (entry.characters > 0) counted = true;
    words += entry.words;
    characters += entry.characters;
  }
  while (statusLineCounts.size > STATUS_LINE_MEMORY) statusLineCounts.delete(statusLineCounts.keys().next().value!);
  return { words, characters };
}

function countIn(text: string): {
  words: number;
  latinWords: number;
  cjk: number;
  withSpaces: number;
  withoutSpaces: number;
} {
  // A Latin "word" is any run with no whitespace and no CJK in it, so
  // "state-of-the-art" is one word and "中文abc" is one CJK-run plus one word.
  const { words, cjkCharacters: cjk, latinWords } = splitWordCount(text);
  return {
    words,
    latinWords,
    cjk,
    withSpaces: characterCount(text),
    withoutSpaces: characterCount(withoutWordSpaces(text)),
  };
}

/** Count a document or a selection the way WPS reports it. */
export function countWords(input: WordCountInput): WordCountResult {
  const texts = [
    ...input.paragraphs,
    ...(input.includeNotes ? (input.notes ?? []) : []),
  ];
  const result: WordCountResult = {
    words: 0,
    charactersWithSpaces: 0,
    charactersWithoutSpaces: 0,
    paragraphs: 0,
    cjkCharacters: 0,
    latinWords: 0,
  };
  for (const text of texts) {
    const counted = countIn(text ?? "");
    result.words += counted.words;
    result.latinWords += counted.latinWords;
    result.cjkCharacters += counted.cjk;
    result.charactersWithSpaces += counted.withSpaces;
    result.charactersWithoutSpaces += counted.withoutSpaces;
    // Word counts a paragraph only when it holds something.
    if (counted.withoutSpaces > 0) result.paragraphs += 1;
  }
  return result;
}

/** One-line summary for a status bar, in the reader's language. */
export function wordCountSummary(
  result: WordCountResult,
  locale: "en" | "zh",
): string {
  return locale === "zh"
    ? `${result.words} 字 · ${result.charactersWithSpaces} 字符`
    : `${result.words} words · ${result.charactersWithSpaces} characters`;
}
