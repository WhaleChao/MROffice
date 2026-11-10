/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The spelling checker every MRAI editor shares: which word of a text the
 * Spelling dialog stops on next, and what it suggests.
 *
 * Semantics follow Office's proofing options (Microsoft Support, "Ignore
 * uppercase words, numbers, or Internet addresses when checking spelling";
 * "Add or edit words in a spell check dictionary"): words in UPPERCASE, words
 * with numbers and Internet / file addresses are passed over when their
 * options are on, a repeated word ("beep beep") is flagged on its second
 * occurrence, the custom dictionary's words are accepted (and offered as
 * suggestions unless "Suggest from main dictionary only" is on), and Ignore
 * All passes a word over for the rest of the session. The main dictionary is a
 * Hunspell dictionary read by nspell (lib/proofing/spelling-dictionaries).
 */

import type { SpellingOptions } from "./proofing-preferences";
import {
  isLatinScriptWord,
  isScriptWord,
  tokenizeForSpelling,
  type SpellingScript,
  type SpellingToken,
} from "./spelling-tokenize";

/** How a dictionary judges a word, from the user's spelling options. */
export interface SpellingJudgeOptions {
  /** Word's "Enforce accented uppercase in French": false accepts an
   *  unaccented capital where the word takes an accented one ("ECOLE"). */
  enforceAccentedUppercase?: boolean;
}

/** A main dictionary: what Hunspell offers, reduced to the checker's needs. */
export interface SpellingDictionary {
  correct(word: string, options?: SpellingJudgeOptions): boolean;
  suggest(word: string): string[];
  /** The script of the words it judges (absent: Latin). */
  readonly script?: SpellingScript;
}

export interface SpellingCheckContext {
  dictionary: SpellingDictionary;
  options: SpellingOptions;
  /** The custom dictionary (Add to Dictionary). */
  customWords: readonly string[];
  /** Words passed over with Ignore All, in the form `ignoreKey` gives. */
  ignoredWords: ReadonlySet<string>;
  /** Letters the dictionary can judge; words in other scripts are skipped.
   *  Absent = the dictionary's own script (Latin when it states none). */
  checksWord?: (word: string) => boolean;
}

export type SpellingIssueKind = "misspelled" | "repeated";

export interface SpellingIssue {
  kind: SpellingIssueKind;
  /** UTF-16 range of the word in the checked text. */
  start: number;
  end: number;
  word: string;
}

/** A range of the text to check: a selection inside a cell being edited. */
export interface SpellingTextRange {
  start: number;
  end: number;
}

/** The key Ignore All remembers a word by: the same word in any case. */
export function ignoreKey(word: string): string {
  return word.normalize("NFC").replace(/’/g, "'").toLowerCase();
}

function apostrophesStraight(word: string): string {
  return word.replace(/’/g, "'");
}

function isUppercaseWord(word: string): boolean {
  return /\p{Lu}/u.test(word) && !/\p{Ll}/u.test(word);
}

function capitalized(word: string): string {
  const [first = "", ...rest] = Array.from(word);
  return first.toUpperCase() + rest.join("");
}

/**
 * Whether the custom dictionary accepts `word`, with Hunspell's personal
 * dictionary case rules: an entry in lower case also accepts its Capitalized
 * and UPPERCASE forms; an entry with capitals ("MRAI", "McDonald") accepts
 * itself and its UPPERCASE form only.
 */
export function customDictionaryAccepts(
  customWords: readonly string[],
  word: string,
): boolean {
  if (customWords.length === 0) return false;
  const candidate = apostrophesStraight(word.normalize("NFC"));
  const upper = candidate.toUpperCase();
  for (const raw of customWords) {
    const entry = apostrophesStraight(raw);
    if (entry === candidate) return true;
    if (entry.toUpperCase() !== upper) continue;
    if (candidate === upper) return true;
    if (entry === entry.toLowerCase() && candidate === capitalized(entry)) return true;
  }
  return false;
}

function hasDigit(word: string): boolean {
  return /\p{N}/u.test(word);
}

/**
 * Whether one word passes: the main dictionary, the custom dictionary or
 * Ignore All accepts it. A hyphenated word passes when the whole word does
 * ("x-ray") or every part does ("well-known").
 */
export function spellingAccepts(word: string, context: SpellingCheckContext): boolean {
  if (context.ignoredWords.has(ignoreKey(word))) return true;
  if (customDictionaryAccepts(context.customWords, word)) return true;
  const judge: SpellingJudgeOptions = {
    enforceAccentedUppercase: context.options.enforceAccentedUppercaseFrench,
  };
  if (context.dictionary.correct(word, judge)) return true;
  // A word the dictionary only knows with a straight apostrophe.
  if (word.includes("’") && context.dictionary.correct(apostrophesStraight(word), judge)) {
    return true;
  }
  if (word.includes("-")) {
    const parts = word.split("-").filter(Boolean);
    return (
      parts.length > 1 &&
      parts.every((part) => !wordIsChecked(part, context) || spellingAccepts(part, context))
    );
  }
  return false;
}

/** Whether the options let the checker judge a word at all. */
function wordIsChecked(word: string, context: SpellingCheckContext): boolean {
  if (!/\p{L}/u.test(word)) return false; // 2026, 3.14: numbers are not words
  const checksWord =
    context.checksWord ??
    (context.dictionary.script && context.dictionary.script !== "Latin"
      ? (candidate: string) => isScriptWord(candidate, context.dictionary.script!)
      : isLatinScriptWord);
  if (!checksWord(word)) return false;
  if (context.options.ignoreUppercase && isUppercaseWord(word)) return false;
  if (context.options.ignoreWordsWithNumbers && hasDigit(word)) return false;
  return true;
}

function onlySpaceBetween(text: string, from: number, to: number): boolean {
  return /^[^\S\n\r\u2028\u2029]+$/u.test(text.slice(from, to));
}

function repeats(previous: SpellingToken | null, token: SpellingToken, text: string): boolean {
  return (
    previous !== null &&
    onlySpaceBetween(text, previous.end, token.start) &&
    ignoreKey(previous.word) === ignoreKey(token.word) &&
    /\p{L}/u.test(token.word) &&
    !hasDigit(token.word)
  );
}

/**
 * The first issue in `text` at or after `from` (and inside `range`, when
 * one is given): a repeated word or a word no dictionary accepts.
 */
export function findSpellingIssue(
  text: string,
  from: number,
  context: SpellingCheckContext,
  range?: SpellingTextRange,
): SpellingIssue | null {
  if (!text) return null;
  const tokens = tokenizeForSpelling(text, {
    ignoreInternetAndFileAddresses: context.options.ignoreInternetAndFileAddresses,
  });
  const lower = Math.max(from, range?.start ?? 0);
  const upper = range?.end ?? text.length;
  let previous: SpellingToken | null = null;
  for (const token of tokens) {
    const inside = token.start >= lower && token.end <= upper;
    if (inside) {
      if (context.options.flagRepeatedWords && repeats(previous, token, text)) {
        // The pair must lie inside the checked range, not straddle it.
        if (previous && previous.start >= (range?.start ?? 0)) {
          return { kind: "repeated", start: token.start, end: token.end, word: token.word };
        }
      }
      if (wordIsChecked(token.word, context) && !spellingAccepts(token.word, context)) {
        return { kind: "misspelled", start: token.start, end: token.end, word: token.word };
      }
    }
    if (token.start >= upper) break;
    previous = token;
  }
  return null;
}

/** Damerau-Levenshtein (optimal string alignment) distance, case-folded. */
export function spellingEditDistance(left: string, right: string): number {
  const a = Array.from(left.toLowerCase());
  const b = Array.from(right.toLowerCase());
  const rows = a.length + 1;
  const cols = b.length + 1;
  const d: number[][] = Array.from({ length: rows }, (_, i) =>
    Array.from({ length: cols }, (_, j) => (i === 0 ? j : j === 0 ? i : 0)),
  );
  for (let i = 1; i < rows; i += 1) {
    for (let j = 1; j < cols; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
      }
    }
  }
  return d[a.length][b.length];
}

function sameLetters(left: string, right: string): boolean {
  const sort = (word: string) => Array.from(word.toLowerCase()).sort().join("");
  return sort(left) === sort(right);
}

/** Give `replacement` the capitalization pattern of `original`. */
export function matchSpellingCase(original: string, replacement: string): string {
  if (!replacement) return replacement;
  if (Array.from(original).length > 1 && isUppercaseWord(original)) {
    return replacement.toUpperCase();
  }
  const [first = ""] = Array.from(original);
  if (first && first === first.toUpperCase() && first !== first.toLowerCase()) {
    return capitalized(replacement);
  }
  return replacement;
}

/**
 * The replacement Change All writes over one more occurrence of a flagged
 * word: the occurrence's own capitalization, not the flagged one's
 * ("Teh" changed to "The" changes "teh" to "the" and "TEH" to "THE").
 */
export function adaptSpellingReplacement(
  flagged: string,
  replacement: string,
  occurrence: string,
): string {
  if (occurrence === flagged) return replacement;
  let base = replacement;
  if (Array.from(flagged).length > 1 && isUppercaseWord(flagged)) {
    base = replacement.toLowerCase();
  } else {
    const [first = ""] = Array.from(flagged);
    const [replacementFirst = ""] = Array.from(replacement);
    if (
      first !== first.toLowerCase() &&
      replacementFirst !== replacementFirst.toLowerCase()
    ) {
      base = replacementFirst.toLowerCase() + replacement.slice(replacementFirst.length);
    }
  }
  return matchSpellingCase(occurrence, base);
}

export const MAX_SPELLING_SUGGESTIONS = 10;

/**
 * Suggestions for a flagged word, best first: the dictionary's, then the
 * custom dictionary's near words (unless "Suggest from main dictionary
 * only"). Ranked by edit distance with transpositions, a candidate made of
 * the same letters first among equals ("teh" offers "the" before "ten").
 */
export function spellingSuggestions(word: string, context: SpellingCheckContext): string[] {
  const seen = new Set<string>();
  const ranked: Array<{ text: string; distance: number; anagram: boolean; order: number }> = [];
  const push = (candidate: string) => {
    const text = candidate.normalize("NFC");
    if (!text || text === word || seen.has(text)) return;
    seen.add(text);
    ranked.push({
      text,
      distance: spellingEditDistance(word, text),
      anagram: sameLetters(word, text),
      order: ranked.length,
    });
  };
  let main: string[] = [];
  try {
    main = context.dictionary.suggest(word);
  } catch {
    main = [];
  }
  for (const candidate of main) push(candidate);
  if (!context.options.suggestFromMainDictionaryOnly) {
    for (const entry of context.customWords) {
      if (spellingEditDistance(word, entry) <= 2) push(matchSpellingCase(word, entry));
    }
  }
  ranked.sort(
    (left, right) =>
      left.distance - right.distance ||
      Number(right.anagram) - Number(left.anagram) ||
      left.order - right.order,
  );
  return ranked.slice(0, MAX_SPELLING_SUGGESTIONS).map((entry) => entry.text);
}

/**
 * Every occurrence of `word` in `text` (the same word in any case), for
 * Change All. Ranges come in text order.
 */
export function spellingOccurrences(
  text: string,
  word: string,
  options: SpellingOptions,
  range?: SpellingTextRange,
): SpellingToken[] {
  const key = ignoreKey(word);
  const lower = range?.start ?? 0;
  const upper = range?.end ?? text.length;
  return tokenizeForSpelling(text, {
    ignoreInternetAndFileAddresses: options.ignoreInternetAndFileAddresses,
  }).filter(
    (token) => token.start >= lower && token.end <= upper && ignoreKey(token.word) === key,
  );
}










type SpellingScope = string | object | undefined;
const sessionIgnoreAll = new Set<string>();
const scopedIgnoreAll = new Map<string, Set<string>>();
const objectScopedIgnoreAll = new WeakMap<object, Set<string>>();

function ignoreAllListFor(scope: SpellingScope, create: boolean): Set<string> | undefined {
  if (scope === undefined) return sessionIgnoreAll;
  if (typeof scope === "string") {
    let list = scopedIgnoreAll.get(scope);
    if (!list && create) {
      list = new Set<string>();
      scopedIgnoreAll.set(scope, list);
    }
    return list;
  }
  let list = objectScopedIgnoreAll.get(scope);
  if (!list && create) {
    list = new Set<string>();
    objectScopedIgnoreAll.set(scope, list);
  }
  return list;
}

const EMPTY_IGNORED: ReadonlySet<string> = new Set<string>();

export function spellingSessionIgnoredWords(scope?: SpellingScope): ReadonlySet<string> {
  return ignoreAllListFor(scope, false) ?? EMPTY_IGNORED;
}

export function ignoreWordForSession(word: string, scope?: SpellingScope): void {
  ignoreAllListFor(scope, true)!.add(ignoreKey(word));
}

export function stopIgnoringWordForSession(word: string, scope?: SpellingScope): void {
  ignoreAllListFor(scope, false)?.delete(ignoreKey(word));
}

/** Tests only: start a fresh session. */
export function resetSpellingSessionForTests(): void {
  sessionIgnoreAll.clear();
  scopedIgnoreAll.clear();
}
