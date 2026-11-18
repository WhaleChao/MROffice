/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a word may be broken by automatic hyphenation (Word's Layout >
 * Hyphenation > Automatic), for the Docs line engine (docs-lines.ts), which
 * decides which of these points a line ends at by Word's own rules.
 *
 * Liang's pattern method over the TeX patterns of the language. Word's own
 * hyphenator is not published; on the 73 words Windows Word hyphenated in a
 * corpus document (en-US, Calibri, mode 15) these patterns give the point
 * Word broke at for 67 ("ap-proaches", "test-crosses", "genera-tions",
 * "technol-ogy"...); Word's own differ on six ("alt-hough", "var-iance",
 * "availa-ble", "ped-igree" twice, a place name).
 *
 * Only languages with patterns here are hyphenated by the engine; the rest
 * are left to the browser (`hyphenationLanguage` returns null).
 */
import { EN_US_EXCEPTIONS, EN_US_PATTERNS } from "./hyphenation-en-us";

interface Hyphenator {
  /** Pattern letters → the digits of their gaps (length + 1). */
  readonly patterns: ReadonlyMap<string, readonly number[]>;
  readonly exceptions: ReadonlyMap<string, readonly number[]>;
  readonly longest: number;
  /** Fewest letters before and after a point. */
  readonly left: number;
  readonly right: number;
}

function build(patternText: string, exceptionText: string, left: number, right: number): Hyphenator {
  const patterns = new Map<string, number[]>();
  let longest = 0;
  for (const pattern of patternText.split(/\s+/)) {
    if (!pattern) continue;
    let letters = "";
    const digits: number[] = [0];
    for (const char of pattern) {
      if (char >= "0" && char <= "9") digits[digits.length - 1] = Number(char);
      else {
        letters += char;
        digits.push(0);
      }
    }
    patterns.set(letters, digits);
    longest = Math.max(longest, letters.length);
  }
  const exceptions = new Map<string, number[]>();
  for (const word of exceptionText.split(/\s+/)) {
    if (!word) continue;
    const points: number[] = [];
    let letters = 0;
    for (const char of word) {
      if (char === "-") points.push(letters);
      else letters += 1;
    }
    exceptions.set(word.replace(/-/g, ""), points);
  }
  return { patterns, exceptions, longest, left, right };
}

const built = new Map<string, Hyphenator>();

/** The patterns a language tag is hyphenated with, or null (none here). */
export function hyphenationLanguage(tag: string | null | undefined): string | null {
  if (!tag) return null;
  const lower = tag.toLowerCase();
  // American English; the other English variants hyphenate by their own
  // dictionaries in Word (British "pro-ject" against American "proj-ect").
  if (lower === "en-us" || lower === "en") return "en-us";
  return null;
}

function hyphenator(language: string): Hyphenator | null {
  const known = built.get(language);
  if (known) return known;
  if (language !== "en-us") return null;
  const made = build(EN_US_PATTERNS, EN_US_EXCEPTIONS, 2, 3);
  built.set(language, made);
  return made;
}

const cache = new Map<string, readonly number[]>();
const CACHE_LIMIT = 20_000;
const NO_POINTS: readonly number[] = [];

/**
 * The points a word may be broken at, as counts of its UTF-16 code units
 * before each point, ascending. `word` is letters only (the caller cuts
 * words at anything else). Empty when the language has no patterns here or
 * the word is too short.
 */
export function hyphenationPoints(word: string, language: string): readonly number[] {
  const key = `${language}\u0000${word}`;
  const known = cache.get(key);
  if (known) return known;
  const result = computePoints(word, language);
  if (cache.size >= CACHE_LIMIT) cache.clear();
  cache.set(key, result);
  return result;
}

function computePoints(word: string, language: string): readonly number[] {
  const engine = hyphenator(language);
  // Patterns are per letter; a word of other code units (astral letters,
  // combining marks) is left whole.
  if (!engine || word.length < engine.left + engine.right || !/^[a-zA-ZÀ-ɏ]+$/.test(word)) return NO_POINTS;
  const lower = word.toLowerCase();
  const exception = engine.exceptions.get(lower);
  if (exception) return exception.filter((at) => at >= engine.left && at <= word.length - engine.right);
  const text = `.${lower}.`;
  const values = new Array<number>(text.length + 1).fill(0);
  for (let start = 0; start < text.length; start += 1) {
    const stop = Math.min(text.length, start + engine.longest);
    for (let end = start + 1; end <= stop; end += 1) {
      const digits = engine.patterns.get(text.slice(start, end));
      if (!digits) continue;
      for (let k = 0; k < digits.length; k += 1) {
        if (digits[k]! > values[start + k]!) values[start + k] = digits[k]!;
      }
    }
  }
  // values[i] is the gap before text[i]; the word's gap after its n-th
  // letter is values[n + 1] (text starts with ".").
  const points: number[] = [];
  for (let at = engine.left; at <= word.length - engine.right; at += 1) {
    if (values[at + 1]! % 2 === 1) points.push(at);
  }
  return points;
}
