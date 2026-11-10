/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Typing aids for the linear-format source box, as Word applies them while
 * an equation is in linear form (UTN #28 §§2.5, 3.5, Appendix B):
 *
 * - letters typed as variables become math italic (𝑥), while a recognized
 *   function name ("sin", "lim" …) stays upright ASCII;
 * - a keyword (`\alpha`, `\sqrt`, `\int` …) followed by a space or another
 *   non-letter is replaced by its character; the terminating space is
 *   consumed, as Math AutoCorrect does.
 */

import { RECOGNIZED_FUNCTIONS, encodeMathCharacter, isItalicByDefault } from "./characters";
import { UNICODE_MATH_KEYWORDS } from "./keywords";
import { matchAllBehind } from "../regexp-behind";

const FUNCTION_TERMINATORS = new Set([" ", "(", "[", "{", "⁡", "_", "^", "┬", "┴", "〖", "|"]);

function italic(character: string): string {
  return isItalicByDefault(character) ? encodeMathCharacter(character, {}) : character;
}

/**
 * Translate freshly typed or pasted text. `before` is the text preceding the
 * insertion (to know whether it continues a keyword or a quoted run).
 */
export function translateLinearInput(inserted: string, before = ""): string {
  const characters = Array.from(inserted);
  // Quotes not after a backslash (tested in code: no lookbehind, Safari < 16.4).
  let quoted = [...matchAllBehind(before, /"/g, { char: /\\/, not: true })].length % 2 === 1;
  let out = "";
  for (let index = 0; index < characters.length; index += 1) {
    const character = characters[index];
    if (character === '"' && !(out.endsWith("\\") || (out === "" && before.endsWith("\\")))) {
      quoted = !quoted;
      out += character;
      continue;
    }
    if (quoted || !isItalicByDefault(character)) {
      out += character;
      continue;
    }
    if (/^[A-Za-z]$/.test(character)) {
      // A keyword's letters stay ASCII until the keyword is complete.
      const context = before + out;
      if (/\\[A-Za-z]*$/.test(context)) {
        out += character;
        continue;
      }
      let end = index;
      while (end < characters.length && /^[A-Za-z]$/.test(characters[end])) end += 1;
      const word = characters.slice(index, end).join("");
      const next = characters[end];
      const startsWord = !/[A-Za-z]$/.test(context);
      if (startsWord && RECOGNIZED_FUNCTIONS.has(word) && next !== undefined && FUNCTION_TERMINATORS.has(next)) {
        out += word;
        index = end - 1;
        continue;
      }
    }
    out += italic(character);
  }
  return out;
}

/** Replace a completed keyword before `cursor`; null when there is none. */
export function completeLinearKeyword(
  text: string,
  cursor: number,
): { text: string; cursor: number } | null {
  const head = text.slice(0, cursor);
  const match = /\\([A-Za-z]+)([^A-Za-z])$/.exec(head);
  if (!match) return null;
  const glyph = Object.prototype.hasOwnProperty.call(UNICODE_MATH_KEYWORDS, match[1])
    ? UNICODE_MATH_KEYWORDS[match[1]]
    : undefined;
  if (glyph === undefined) return null;
  const replacement = italic(glyph) + (match[2] === " " ? "" : match[2]);
  const start = head.length - match[0].length;
  return {
    text: text.slice(0, start) + replacement + text.slice(cursor),
    cursor: start + replacement.length,
  };
}

/**
 * When a space, bracket or script follows italic letters that spell a
 * recognized function name, Word shows the name upright: restore ASCII.
 */
export function uprightFunctionNameBefore(
  text: string,
  cursor: number,
): { text: string; cursor: number } | null {
  const head = Array.from(text.slice(0, cursor));
  const terminator = head.at(-1);
  if (!terminator || !FUNCTION_TERMINATORS.has(terminator)) return null;
  let start = head.length - 1;
  let word = "";
  while (start > 0) {
    const character = head[start - 1];
    const plain = PLAIN_ITALIC.get(character);
    if (!plain) break;
    word = plain + word;
    start -= 1;
  }
  if (!word || !RECOGNIZED_FUNCTIONS.has(word)) return null;
  const previous = head[start - 1];
  if (previous && (PLAIN_ITALIC.has(previous) || /^[A-Za-z]$/.test(previous))) return null;
  const prefix = head.slice(0, start).join("");
  const replaced = prefix + word + terminator;
  return { text: replaced + text.slice(cursor), cursor: replaced.length };
}

const PLAIN_ITALIC = new Map<string, string>();
for (const letter of "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") {
  PLAIN_ITALIC.set(encodeMathCharacter(letter, {}), letter);
}
