/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure Python source-coordinate helpers shared by the standalone Code editor
 * and its lazy Jedi language-service provider.
 */

/**
 * Convert a character offset into a 1-based line and 0-based UTF-16 column.
 * Stale offsets are clamped into the source.
 */
export function offsetToLineCol(
  text: string,
  offset: number,
): { line: number; column: number } {
  const clamped = Math.max(0, Math.min(offset, text.length));
  const before = text.slice(0, clamped);
  const newline = before.lastIndexOf("\n");
  return {
    line: before.split("\n").length,
    column: clamped - (newline + 1),
  };
}

/** Convert a code-point column (Jedi) to a UTF-16 column (CodeMirror). */
export function codePointToUtf16Col(lineText: string, codePointCol: number) {
  if (codePointCol <= 0) return 0;
  let codePoints = 0;
  let utf16Units = 0;
  for (const character of lineText) {
    if (codePoints >= codePointCol) break;
    utf16Units += character.length;
    codePoints += 1;
  }
  return utf16Units;
}

/** Convert a UTF-16 column (CodeMirror) to a code-point column (Jedi). */
export function utf16ColToCodePointCol(lineText: string, utf16Col: number) {
  if (utf16Col <= 0) return 0;
  let codePoints = 0;
  let utf16Units = 0;
  for (const character of lineText) {
    if (utf16Units >= utf16Col) break;
    utf16Units += character.length;
    codePoints += 1;
  }
  return codePoints;
}

const PY_IDENTIFIER = /^[A-Za-z_][A-Za-z0-9_]*$/;

const PY_KEYWORDS = new Set([
  "False",
  "None",
  "True",
  "and",
  "as",
  "assert",
  "async",
  "await",
  "break",
  "class",
  "continue",
  "def",
  "del",
  "elif",
  "else",
  "except",
  "finally",
  "for",
  "from",
  "global",
  "if",
  "import",
  "in",
  "is",
  "lambda",
  "nonlocal",
  "not",
  "or",
  "pass",
  "raise",
  "return",
  "try",
  "while",
  "with",
  "yield",
]);

export function isRenamablePythonIdentifier(name: string): boolean {
  return PY_IDENTIFIER.test(name) && !PY_KEYWORDS.has(name);
}

export function identifierSpanAt(
  text: string,
  offset: number,
): { from: number; to: number } | null {
  const isWord = (character: string) => /[A-Za-z0-9_]/.test(character);
  let start = Math.max(0, Math.min(offset, text.length));
  let end = start;
  while (start > 0 && isWord(text[start - 1])) start -= 1;
  while (end < text.length && isWord(text[end])) end += 1;
  const word = text.slice(start, end);
  return word && PY_IDENTIFIER.test(word) ? { from: start, to: end } : null;
}

export function identifierAt(text: string, offset: number): string | null {
  const span = identifierSpanAt(text, offset);
  return span ? text.slice(span.from, span.to) : null;
}
