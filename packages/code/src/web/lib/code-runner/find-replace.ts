/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure find/replace logic for the code editor. Framework-free so it is fully
 * unit-tested; the editor renders the UI and applies the textarea selection.
 *
 * Supports VSCode-style options: case sensitivity, whole-word matching, regular
 * expressions (with $1 capture-group replacement), and find-in-selection.
 */

export interface Match {
  start: number;
  end: number;
  /** [fullMatch, $1, $2, ...] — only populated for regex matches. */
  captures?: string[];
  /** Named capture groups ($<name>) — only populated for regex matches. */
  namedGroups?: Record<string, string>;
}

export interface FindOptions {
  caseSensitive?: boolean;
  wholeWord?: boolean;
  regex?: boolean;
  /** Restrict the search to a document slice (find-in-selection). */
  selectionRange?: { start: number; end: number };
  /**
   * Preserve Case (VSCode Aa/AB): rewrite each replacement to match the case
   * shape (UPPER / lower / Capitalized) of the text it replaces. See
   * {@link preserveCase}. No-op for mixed/uncased matches.
   */
  preserveCase?: boolean;
}

const UPPER_CASED_CHAR = /\p{Lu}/u;
const LOWER_CASED_CHAR = /\p{Ll}/u;
const TITLE_CASED_CHAR = /\p{Lt}/u;
const CASED_CHAR = /[\p{Lu}\p{Ll}\p{Lt}]/u;

/**
 * Rewrite a replacement string so it adopts the CASE PATTERN of the text it is
 * replacing — VSCode's "Preserve Case" (Aa/AB) toggle. The mapping mirrors
 * VSCode exactly:
 *
 *  - the matched text is all UPPERCASE (and has at least one cased letter) ->
 *    UPPERCASE the replacement (`FOO` + `bar` -> `BAR`);
 *  - the matched text is all lowercase -> lowercase the replacement
 *    (`foo` -> `bar`);
 *  - the matched text is Capitalized (first letter upper, rest lower) ->
 *    capitalize the replacement (`Foo` -> `Bar`);
 *  - otherwise (mixed/camel case, or no cased letters) -> the replacement is
 *    returned verbatim.
 *
 * Only the FIRST cased letter decides "capitalized"; a single-letter all-caps
 * match (`F`) counts as uppercase. Pure + total.
 */
export function preserveCase(matched: string, replacement: string): string {
  if (!replacement) return replacement;

  const hasUpper = UPPER_CASED_CHAR.test(matched);
  const hasLower = LOWER_CASED_CHAR.test(matched);
  const hasTitle = TITLE_CASED_CHAR.test(matched);
  const hasCased = hasUpper || hasLower || hasTitle;
  if (!hasCased) return replacement; // nothing to mimic (digits/symbols only)

  // All uppercase (no lowercase letters present, at least one uppercase).
  if (hasUpper && !hasLower && !hasTitle) return replacement.toUpperCase();
  // All lowercase.
  if (hasLower && !hasUpper && !hasTitle) return replacement.toLowerCase();

  // Mixed: treat as "Capitalized" only when the first CASED letter is upper and
  // every later cased letter is lower (e.g. "Foo", "Foo Bar" word-initial). Any
  // other interior uppercase (camelCase / PascalCase like "FooBar") is left
  // untouched, matching VSCode which only preserves the three simple shapes.
  const letters = Array.from(matched).filter((character) =>
    CASED_CHAR.test(character),
  );
  const firstUpper =
    UPPER_CASED_CHAR.test(letters[0] ?? "") ||
    TITLE_CASED_CHAR.test(letters[0] ?? "");
  const restAllLower = letters
    .slice(1)
    .every((character) => LOWER_CASED_CHAR.test(character));
  if (firstUpper && restAllLower) {
    const [first = "", ...rest] = Array.from(replacement);
    return first.toUpperCase() + rest.join("").toLowerCase();
  }
  return replacement;
}

// Whole-word search must treat non-ASCII identifiers as words too. The prior
// ASCII-only class considered every CJK/Arabic/Cyrillic letter a separator, so
// searching "变量" with Whole Word enabled still matched inside "变量名".
// Combining marks are continuations as well: decomposed "e\u0301" is one word,
// not an ASCII "e" followed by a boundary.
const WORD_CHAR = /[\p{L}\p{N}\p{M}_]/u;
export const MAX_FIND_REGEX_CHARS = 512;
export const MAX_FIND_QUERY_CHARS = 4_096;
export const UNSAFE_FIND_REGEX_MESSAGE =
  "Regular expression uses an unsafe backtracking structure.";
export const OVERSIZED_FIND_REGEX_MESSAGE = "Regular expression is too long.";
export const MAX_EDITOR_FIND_MATCHES = 10_000;
export const MAX_FIND_REPLACEMENT_MATCHES = 100_000;
export const MAX_FIND_REPLACEMENT_RESULT_CHARS = 1_000_000;

export function boundFindQuery(query: string): string {
  return query.slice(0, MAX_FIND_QUERY_CHARS);
}

export class FindReplacementTooLargeError extends Error {
  constructor() {
    super("Replacement result exceeds the Code page limit.");
    this.name = "FindReplacementTooLargeError";
  }
}

export function isFindReplacementTooLargeError(
  error: unknown,
): error is FindReplacementTooLargeError {
  return error instanceof FindReplacementTooLargeError;
}

export class FindReplacementTooManyMatchesError extends Error {
  constructor() {
    super("Replacement has too many matches for one Code operation.");
    this.name = "FindReplacementTooManyMatchesError";
  }
}

export class FindReplacementWorkspaceTooLargeError extends Error {
  constructor() {
    super("Workspace replacement result exceeds the Code operation limit.");
    this.name = "FindReplacementWorkspaceTooLargeError";
  }
}

export type FindReplacementPlanError =
  | FindReplacementTooLargeError
  | FindReplacementTooManyMatchesError
  | FindReplacementWorkspaceTooLargeError;
export type FindReplacementPlanMessageKey =
  | "code.find.replacementTooLarge"
  | "code.find.tooManyReplacements"
  | "code.find.workspaceReplacementTooLarge";

export function isFindReplacementPlanError(
  error: unknown,
): error is FindReplacementPlanError {
  return (
    error instanceof FindReplacementTooLargeError ||
    error instanceof FindReplacementTooManyMatchesError ||
    error instanceof FindReplacementWorkspaceTooLargeError
  );
}

export function findReplacementPlanMessageKey(
  error: FindReplacementPlanError,
): FindReplacementPlanMessageKey {
  if (error instanceof FindReplacementTooManyMatchesError) {
    return "code.find.tooManyReplacements";
  }
  return error instanceof FindReplacementWorkspaceTooLargeError
    ? "code.find.workspaceReplacementTooLarge"
    : "code.find.replacementTooLarge";
}
export type FindRegexErrorMessageKey =
  | "code.find.regexUnsafe"
  | "code.find.regexTooLong"
  | "findReplace.regexError";

export function findRegexErrorMessageKey(
  error: string | null,
): FindRegexErrorMessageKey | null {
  if (error === UNSAFE_FIND_REGEX_MESSAGE) return "code.find.regexUnsafe";
  if (error === OVERSIZED_FIND_REGEX_MESSAGE) return "code.find.regexTooLong";
  return error ? "findReplace.regexError" : null;
}

interface RegexGroupSafety {
  hasVariableQuantifier: boolean;
  hasAlternation: boolean;
}

function quantifierAt(
  pattern: string,
  index: number,
): { end: number; repeats: boolean; variable: boolean } | null {
  const token = pattern[index];
  if (token === "*" || token === "+") {
    return { end: index + 1, repeats: true, variable: true };
  }
  if (token === "?") {
    return { end: index + 1, repeats: false, variable: true };
  }
  if (token !== "{") return null;
  const close = pattern.indexOf("}", index + 1);
  if (close === -1) return null;
  const body = pattern.slice(index + 1, close);
  if (!/^\d+(?:,\d*)?$/u.test(body)) return null;
  const [minimumText, maximumText] = body.split(",");
  const minimum = Number(minimumText);
  const hasComma = body.includes(",");
  const maximum = hasComma
    ? maximumText === ""
      ? Number.POSITIVE_INFINITY
      : Number(maximumText)
    : minimum;
  return {
    end: close + 1,
    repeats: maximum > 1,
    variable: minimum !== maximum,
  };
}

/**
 * Conservative main-thread regex guard. JavaScript RegExp has no timeout, so
 * nested unbounded quantifiers, quantified alternation and backreferences are
 * rejected before authored text reaches the engine. Ordinary captures,
 * bounded repeats and unquantified alternation remain available.
 */
function regexSafetyError(pattern: string): string | null {
  if (pattern.length > MAX_FIND_REGEX_CHARS) {
    return OVERSIZED_FIND_REGEX_MESSAGE;
  }
  const groups: RegexGroupSafety[] = [
    { hasVariableQuantifier: false, hasAlternation: false },
  ];
  let inClass = false;
  for (let index = 0; index < pattern.length; index += 1) {
    const token = pattern[index];
    if (token === "\\") {
      const escaped = pattern[index + 1] ?? "";
      if (/^[1-9]$/u.test(escaped) || escaped === "k") {
        return UNSAFE_FIND_REGEX_MESSAGE;
      }
      index += 1;
      continue;
    }
    if (token === "[") {
      inClass = true;
      continue;
    }
    if (token === "]" && inClass) {
      inClass = false;
      continue;
    }
    if (inClass) continue;
    if (token === "(") {
      groups.push({ hasVariableQuantifier: false, hasAlternation: false });
      continue;
    }
    if (token === "|") {
      groups[groups.length - 1]!.hasAlternation = true;
      continue;
    }
    if (token === ")" && groups.length > 1) {
      const group = groups.pop()!;
      const outer = quantifierAt(pattern, index + 1);
      if (
        outer?.repeats &&
        (group.hasVariableQuantifier || group.hasAlternation)
      ) {
        return UNSAFE_FIND_REGEX_MESSAGE;
      }
      const parent = groups[groups.length - 1]!;
      parent.hasVariableQuantifier ||=
        group.hasVariableQuantifier || outer?.variable === true;
      parent.hasAlternation ||= group.hasAlternation;
      if (outer) index = outer.end - 1;
      continue;
    }
    const quantifier = quantifierAt(pattern, index);
    if (quantifier) {
      if (quantifier.variable) {
        groups[groups.length - 1]!.hasVariableQuantifier = true;
      }
      index = quantifier.end - 1;
    }
  }
  return null;
}

function codePointBefore(text: string, index: number): string {
  if (index <= 0) return "";
  const trailing = text.charCodeAt(index - 1);
  if (trailing >= 0xdc00 && trailing <= 0xdfff && index >= 2) {
    const leading = text.charCodeAt(index - 2);
    if (leading >= 0xd800 && leading <= 0xdbff) {
      return text.slice(index - 2, index);
    }
  }
  return text[index - 1] ?? "";
}

function codePointAfter(text: string, index: number): string {
  const codePoint = text.codePointAt(index);
  return codePoint === undefined ? "" : String.fromCodePoint(codePoint);
}

/** A match is "whole word" when the code points around it are non-word. */
function hasWordBoundaries(scope: string, start: number, end: number): boolean {
  const before = codePointBefore(scope, start);
  const after = codePointAfter(scope, end);
  return !WORD_CHAR.test(before) && !WORD_CHAR.test(after);
}

/**
 * Validate a regex pattern. Returns an error message string when invalid, or
 * null when the pattern compiles. The editor uses this to surface an inline
 * error instead of silently returning no matches.
 */
export function validateRegex(query: string, caseSensitive = false): string | null {
  if (!query) return null;
  const safetyError = regexSafetyError(query);
  if (safetyError) return safetyError;
  try {
    void new RegExp(query, caseSensitive ? "gmu" : "gmiu");
    return null;
  } catch (err) {
    return err instanceof Error ? err.message : "Invalid regular expression";
  }
}

function literalMatches(
  scope: string,
  query: string,
  opts: FindOptions,
  limit: number,
): Match[] {
  const out: Match[] = [];
  if (!opts.caseSensitive) {
    // Lowercasing the whole document can change its UTF-16 length (for example
    // `İ` becomes `i\u0307`). Indices in that folded copy no longer address the
    // original source, so a later ASCII match could be replaced at the wrong
    // offset. An escaped Unicode regex performs simple case folding while its
    // match indices remain anchored to the original string.
    const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&");
    let matcher: RegExp;
    try {
      matcher = new RegExp(escapedQuery, "giu");
    } catch {
      return [];
    }
    let match: RegExpExecArray | null;
    while ((match = matcher.exec(scope)) !== null) {
      const start = match.index;
      const end = start + match[0].length;
      if (!opts.wholeWord || hasWordBoundaries(scope, start, end)) {
        out.push({ start, end });
        if (out.length >= limit) break;
      }
    }
    return out;
  }

  let from = 0;
  for (;;) {
    const idx = scope.indexOf(query, from);
    if (idx === -1) break;
    const end = idx + query.length;
    if (!opts.wholeWord || hasWordBoundaries(scope, idx, end)) {
      out.push({ start: idx, end });
      if (out.length >= limit) break;
    }
    from = end; // non-overlapping (query.length >= 1 here)
  }
  return out;
}

function regexMatches(
  scope: string,
  query: string,
  caseSensitive: boolean,
  limit: number,
): Match[] {
  if (regexSafetyError(query)) return [];
  let re: RegExp;
  try {
    // VSCode-style: multiline so ^/$ anchor to line boundaries within the doc.
    re = new RegExp(query, caseSensitive ? "gmu" : "gmiu");
  } catch {
    return []; // invalid pattern — UI surfaces the error via validateRegex
  }
  const out: Match[] = [];
  let guard = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(scope)) !== null) {
    if (guard++ > scope.length + 1) break; // pathological safety
    if (m[0].length === 0) {
      const codePoint = scope.codePointAt(re.lastIndex);
      re.lastIndex += codePoint !== undefined && codePoint > 0xffff ? 2 : 1;
      continue;
    }
    out.push({
      start: m.index,
      end: m.index + m[0].length,
      captures: m.slice(),
      namedGroups: m.groups ? { ...m.groups } : undefined,
    });
    if (out.length >= limit) break;
  }
  return out;
}

/**
 * Non-overlapping matches of `query` in `text`, left to right.
 *
 * `limit` defaults to every match for existing find/replace callers. Bounded
 * workspace search supplies a finite limit so a huge or highly repetitive file
 * cannot allocate an unbounded match array before the UI applies its own cap.
 */
export function findMatches(
  text: string,
  query: string,
  opts: FindOptions = {},
  limit = Number.POSITIVE_INFINITY,
): Match[] {
  const safeLimit =
    limit === Number.POSITIVE_INFINITY
      ? Number.POSITIVE_INFINITY
      : Number.isFinite(limit)
        ? Math.max(0, Math.trunc(limit))
        : 0;
  if (!query || query.length > MAX_FIND_QUERY_CHARS || safeLimit === 0) return [];
  const range = opts.selectionRange;
  const hasRange = !!range && range.end > range.start;
  const base = hasRange ? range!.start : 0;
  const scope = hasRange ? text.slice(range!.start, range!.end) : text;

  const raw = opts.regex
    ? regexMatches(scope, query, !!opts.caseSensitive, safeLimit)
    : literalMatches(scope, query, opts, safeLimit);

  if (!hasRange) return raw;
  return raw.map((m) => ({ ...m, start: m.start + base, end: m.end + base }));
}

export interface FindMatchPreview {
  /** Bounded matches available for count display and next/previous navigation. */
  matches: Match[];
  /** True when at least one additional match exists beyond the preview window. */
  truncated: boolean;
}

/**
 * A bounded editor-facing match window. The extra sentinel match distinguishes
 * an exact-limit result from a truncated one without building the document's
 * full match array on the main thread.
 */
export function findMatchPreview(
  text: string,
  query: string,
  opts: FindOptions = {},
  limit = MAX_EDITOR_FIND_MATCHES,
): FindMatchPreview {
  const safeLimit = Number.isFinite(limit)
    ? Math.max(0, Math.min(MAX_EDITOR_FIND_MATCHES, Math.trunc(limit)))
    : MAX_EDITOR_FIND_MATCHES;
  const window = findMatches(text, query, opts, safeLimit + 1);
  const truncated = window.length > safeLimit;
  return {
    matches: truncated ? window.slice(0, safeLimit) : window,
    truncated,
  };
}

/** The group a `$` token names, or null for a token kept literally. */
function groupForToken(
  token: string,
  captures: string[],
  namedGroups?: Record<string, string>,
  name?: string,
): string | null {
  if (token === "&") return captures[0] ?? "";
  if (name) return namedGroups?.[name] ?? "";
  const n = Number(token);
  if (n < captures.length) return captures[n] ?? "";
  return null;
}

/** VS Code's \u \l (first character) and \U \L (whole group) case operators. */
function applyCaseOps(value: string, ops: readonly string[]): string {
  let out = value;
  if (ops.includes("U")) out = out.toUpperCase();
  else if (ops.includes("L")) out = out.toLowerCase();
  if (ops.includes("u")) out = out.charAt(0).toUpperCase() + out.slice(1);
  else if (ops.includes("l")) out = out.charAt(0).toLowerCase() + out.slice(1);
  return out;
}

const REPLACEMENT_TOKEN = /\$(\$|&|<([A-Za-z_$][\w$]*)>|\d{1,2})/y;

/**
 * Expand a regex replacement the way VS Code does. `$` tokens: $$ (literal $),
 * $& and $0 (full match), $1..$99 (numbered groups; a 2-digit token falls back
 * to the 1-digit group plus a literal digit, as JS String.replace does, else
 * stays literal), $<name> (named groups). Backslash escapes: \n (newline),
 * \t (tab), \\ (backslash), and the case operators \u \l \U \L, which apply
 * to the next group. Any other backslash stays as written. Without captures
 * (a plain-text search) the replacement is inserted literally.
 */
export function expandReplacement(
  replacement: string,
  captures?: string[],
  namedGroups?: Record<string, string>,
): string {
  if (!captures) return replacement;
  let out = "";
  let caseOps: string[] = [];
  let index = 0;
  while (index < replacement.length) {
    const char = replacement[index]!;
    if (char === "\\" && index + 1 < replacement.length) {
      const next = replacement[index + 1]!;
      if (next === "n" || next === "t" || next === "\\") {
        out += next === "n" ? "\n" : next === "t" ? "\t" : "\\";
        index += 2;
        continue;
      }
      if (next === "u" || next === "l" || next === "U" || next === "L") {
        caseOps.push(next);
        index += 2;
        continue;
      }
    }
    if (char === "$") {
      REPLACEMENT_TOKEN.lastIndex = index;
      const token = REPLACEMENT_TOKEN.exec(replacement);
      if (token) {
        index += token[0].length;
        if (token[1] === "$") {
          out += "$";
          continue;
        }
        let group = groupForToken(token[1]!, captures, namedGroups, token[2]);
        let trailing = "";
        if (group === null && token[1]!.length === 2 && !token[2]) {
          group = groupForToken(token[1]![0]!, captures);
          trailing = token[1]![1]!;
        }
        if (group === null) {
          out += token[0];
          continue;
        }
        out += applyCaseOps(group, caseOps) + trailing;
        caseOps = [];
        continue;
      }
    }
    out += char;
    index += 1;
  }
  return out;
}

/**
 * Index of the match to treat as active given the caret, searching forward and
 * wrapping. Returns -1 when there are no matches.
 */
export function nextMatchIndex(
  matches: Match[],
  caret: number,
  direction: 1 | -1 = 1,
): number {
  if (matches.length === 0) return -1;
  if (direction === 1) {
    const i = matches.findIndex((m) => m.start >= caret);
    return i === -1 ? 0 : i;
  }
  for (let i = matches.length - 1; i >= 0; i -= 1) {
    if (matches[i].end <= caret) return i;
  }
  return matches.length - 1;
}

/** Step the active match index with wrap-around. */
export function stepMatch(count: number, current: number, direction: 1 | -1): number {
  if (count === 0) return -1;
  return (current + direction + count) % count;
}

/**
 * The final replacement text for one match: expand regex tokens ($1/$&/…), then
 * — when `preserveCaseFlag` is set — bend the result to the matched text's case
 * shape. `matchedText` is the literal slice the match covers (the caller has it
 * cheaply from the source). This is the single source of truth the editor and
 * both replace functions share, so case preservation is identical everywhere.
 */
export function computeReplacement(
  replacement: string,
  match: Match,
  matchedText: string,
  preserveCaseFlag = false,
): string {
  const expanded = expandReplacement(replacement, match.captures, match.namedGroups);
  return preserveCaseFlag ? preserveCase(matchedText, expanded) : expanded;
}

/** Replace a single match; returns the new text and the caret after it. */
export function replaceMatch(
  text: string,
  match: Match,
  replacement: string,
  opts: FindOptions = {},
): { text: string; caret: number } {
  const matched = text.slice(match.start, match.end);
  const rep = computeReplacement(replacement, match, matched, opts.preserveCase);
  if (
    text.length - (match.end - match.start) + rep.length >
    MAX_FIND_REPLACEMENT_RESULT_CHARS
  ) {
    throw new FindReplacementTooLargeError();
  }
  return {
    text: text.slice(0, match.start) + rep + text.slice(match.end),
    caret: match.start + rep.length,
  };
}

/** Replace every match; returns the new text and how many were replaced. */
export function replaceAll(
  text: string,
  query: string,
  replacement: string,
  opts: FindOptions = {},
  matchLimit = MAX_FIND_REPLACEMENT_MATCHES,
): { text: string; count: number } {
  const safeMatchLimit = Number.isFinite(matchLimit)
    ? Math.max(
        0,
        Math.min(MAX_FIND_REPLACEMENT_MATCHES, Math.trunc(matchLimit)),
      )
    : MAX_FIND_REPLACEMENT_MATCHES;
  const matches = findMatches(text, query, opts, safeMatchLimit + 1);
  if (matches.length > safeMatchLimit) {
    throw new FindReplacementTooManyMatchesError();
  }
  if (matches.length === 0) return { text, count: 0 };
  const replacements: string[] = [];
  let projectedLength = text.length;
  for (const match of matches) {
    const replacementText = computeReplacement(
      replacement,
      match,
      text.slice(match.start, match.end),
      opts.preserveCase,
    );
    projectedLength += replacementText.length - (match.end - match.start);
    if (projectedLength > MAX_FIND_REPLACEMENT_RESULT_CHARS) {
      throw new FindReplacementTooLargeError();
    }
    replacements.push(replacementText);
  }
  const parts: string[] = [];
  let last = 0;
  for (let index = 0; index < matches.length; index += 1) {
    const match = matches[index];
    parts.push(text.slice(last, match.start), replacements[index]);
    last = match.end;
  }
  parts.push(text.slice(last));
  return { text: parts.join(""), count: matches.length };
}
