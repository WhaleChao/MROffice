/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeEastAsianFontFamilies } from "../../../../lib/document-word-line-metrics";
import { wordFontIsEastAsian, wordFontIsRecordedLatin } from "./wordFontMetrics";
import { normalizeWordCompatibilityMode } from "./documentOverflowPunctuation";





























const answers = new WeakMap<object, boolean>();

export function wordSetsIdeographsAsLetters(attrs: { readonly [key: string]: unknown }): boolean {
  if (typeof window !== "undefined" && (window as Window & { __docsLatinIdeographsOff?: boolean }).__docsLatinIdeographsOff === true) return false;
  // Asked for every paragraph a pass reads: one answer per document attrs.
  let answer = answers.get(attrs);
  if (answer === undefined) answers.set(attrs, (answer = ideographsAsLetters(attrs)));
  return answer;
}

















export function wordOmitsEastAsianAutoSpace(attrs: { readonly [key: string]: unknown }): boolean {
  if (typeof window !== "undefined" && (window as Window & { __docsLatinDefaultsAutoSpaceOff?: boolean }).__docsLatinDefaultsAutoSpaceOff === true) return false;
  let answer = autoSpaceAnswers.get(attrs);
  if (answer === undefined) autoSpaceAnswers.set(attrs, (answer = defaultsNameNoEastAsianFace(attrs)));
  return answer;
}
const autoSpaceAnswers = new WeakMap<object, boolean>();

function ideographsAsLetters(attrs: { readonly [key: string]: unknown }): boolean {
  const mode = normalizeWordCompatibilityMode(attrs.wordCompatibilityMode);
  if (mode === null || mode >= 14) return false;
  return defaultsNameNoEastAsianFace(attrs);
}

/** Whether the document's defaults name no East Asian face (false where the
 * page cannot tell: no defaults' stack). */
function defaultsNameNoEastAsianFace(attrs: { readonly [key: string]: unknown }): boolean {
  const stack = attrs.defaultFontFamily;
  if (typeof stack !== "string" || !stack.trim()) return false;
  if (typeof attrs.defaultEastAsianFontFamily === "string" && attrs.defaultEastAsianFontFamily.trim()) return false;
  const declared = new Set(
    (normalizeEastAsianFontFamilies(attrs.eastAsianFontFamilies) ?? [])
      .filter((family) => !wordFontIsRecordedLatin(family))
      .map((family) => family.toLowerCase()),
  );
  const tokens = stack.split(",").map((token) => token.trim().replace(/^["']|["']$/g, "").trim()).filter(Boolean);
  return !tokens.some((name) => declared.has(name.toLowerCase()) || wordFontIsEastAsian(name));
}
