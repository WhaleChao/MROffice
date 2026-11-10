/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where each braced block starts and ends, for the languages with no parser.
 *
 * The outline knows where names are DECLARED and, through indentation, roughly
 * how deeply they nest. It does not know where a block ENDS, and that missing
 * fact is what made scope-aware Go-to-Definition unsafe: a statement written
 * after `class Beta` closes is indented exactly like one written inside it, so
 * resolving by indentation jumped into Beta from outside it. See the note in
 * structural-language-provider.ts — this module is the fact that was missing.
 *
 * It answers one question: which `{ … }` regions contain a given offset. To do
 * that without a parser it has to be right about the two places a brace is not
 * a brace — inside a comment, and inside a string — so each language brings its
 * own comment and string syntax, including the awkward ones that broke the
 * naive version of this scan:
 *
 *   Rust    `'a` is a lifetime, not the start of a character literal, and
 *           `r#"…"#` counts its own hashes.
 *   Go      a backquoted string runs across lines and honours no escapes.
 *   C#      `@"…"` writes an embedded quote as `""`.
 *   C++     `R"tag(…)tag"` chooses its own delimiter.
 *
 * REFUSAL IS PART OF THE CONTRACT. `blockExtents` returns null — not a partial
 * answer — for an unlisted language, for a document past the scan limit, and,
 * most importantly, for a document whose braces do not balance. A file being
 * typed into is unbalanced most of the time, and a half-scanned file yields
 * confidently wrong extents; callers must fall back to whatever they did before
 * scopes were available rather than trust a guess. That is also the backstop
 * for the string syntax this scanner does not model (PHP and shell heredocs,
 * which is why neither language is listed): mis-reading a string almost always
 * unbalances the braces, and an unbalanced scan reports nothing at all.
 *
 * Offsets are buffer offsets, matching {@link import("./outline-symbols")}.
 */

import { normalizeLanguage } from "./language-metadata";

/** Characters scanned before giving up; a generated blob is not worth scoping. */
export const BLOCK_SCAN_CHAR_LIMIT = 512 * 1024;

/** One `{ … }` region: the offsets of its opening and closing brace. */
export interface BlockExtent {
  open: number;
  close: number;
}

/**
 * A string literal's syntax.
 *
 * `across` is whether it may contain a newline — a plain C string may not, and
 * treating one as if it could turns a single unterminated quote into the rest
 * of the file. `escape` is absent for raw forms, where a backslash is just a
 * backslash. `doubled` is a sequence that starts like the terminator but is
 * not one (C#'s `""` inside `@"…"`).
 */
interface StringRule {
  open: string;
  close: string;
  across: boolean;
  escape?: string;
  doubled?: string;
}

interface BraceSyntax {
  line: readonly string[];
  block: readonly (readonly [string, string])[];
  /** Longest opener first, so `"""` is tried before `"`. */
  strings: readonly StringRule[];
  /** Rust: `'` opens a character literal only when it closes like one. */
  lifetimeAwareChar?: true;
  /** Rust: `r"…"`, `r#"…"#`, `br##"…"##`. */
  rustRawStrings?: true;
  /** C++: `R"tag(…)tag"`. */
  cppRawStrings?: true;
}

const C_COMMENTS = {
  line: ["//"] as const,
  block: [["/*", "*/"]] as const,
};

/** `"…"` and `'…'` as C writes them: escaped with a backslash, one line only. */
const C_QUOTE: StringRule = { open: '"', close: '"', across: false, escape: "\\" };
const C_CHAR: StringRule = { open: "'", close: "'", across: false, escape: "\\" };

/**
 * The languages scope resolution is trusted for.
 *
 * Every one here has C-style comments and a string syntax this scanner models
 * completely. PHP, shell, Perl and R are deliberately absent: their heredocs
 * (and, for Perl, braces inside regular expressions) are not modelled, and a
 * language is either understood or left alone. Ruby and Lua close blocks with
 * `end` rather than a brace, which is a different scan and not this one.
 */
const SYNTAX: Readonly<Record<string, BraceSyntax>> = {
  c: { ...C_COMMENTS, strings: [C_QUOTE, C_CHAR] },
  cpp: { ...C_COMMENTS, strings: [C_QUOTE, C_CHAR], cppRawStrings: true },
  java: {
    ...C_COMMENTS,
    // A text block runs across lines; `\` still escapes inside one.
    strings: [
      { open: '"""', close: '"""', across: true, escape: "\\" },
      C_QUOTE,
      C_CHAR,
    ],
  },
  csharp: {
    ...C_COMMENTS,
    strings: [
      // Raw (C# 11) first: no escapes at all inside.
      { open: '"""', close: '"""', across: true },
      // Verbatim: runs across lines, and `""` is an embedded quote.
      { open: '@"', close: '"', across: true, doubled: '""' },
      C_QUOTE,
      C_CHAR,
    ],
  },
  swift: {
    ...C_COMMENTS,
    // Swift has no character literal — `'` is not a quote here at all.
    strings: [{ open: '"""', close: '"""', across: true, escape: "\\" }, C_QUOTE],
  },
  go: {
    ...C_COMMENTS,
    // A backquoted string spans lines and interprets nothing, not even `\`.
    strings: [{ open: "`", close: "`", across: true }, C_QUOTE, C_CHAR],
  },
  rust: {
    ...C_COMMENTS,
    // A normal Rust string may contain a newline; `'` is handled separately.
    strings: [{ open: '"', close: '"', across: true, escape: "\\" }],
    lifetimeAwareChar: true,
    rustRawStrings: true,
  },
};

/** Whether {@link blockExtents} can answer for `language` at all. */
export function supportsBlockExtents(language: string): boolean {
  return normalizeLanguage(language) in SYNTAX;
}

/** `'a'`, `'\n'`, `'\u{1F600}'` — but not the `'a` of a Rust lifetime. */
const RUST_CHAR = /^'(?:\\(?:u\{[0-9a-fA-F]{1,6}\}|x[0-9a-fA-F]{2}|.)|[^'\\\n])'/;
/** `r"`, `r#"`, `br##"` … the hashes are counted and must match on the way out. */
const RUST_RAW_OPEN = /^b?r(#*)"/;
/** `R"tag(` — the tag is at most 16 characters and excludes spaces and parens. */
const CPP_RAW_OPEN = /^R"([^()\\ \t\n]{0,16})\(/;

/**
 * Every `{ … }` region in `text`, ordered by opening brace, or null when the
 * document cannot be scanned with confidence — see the module comment. A
 * balanced document with no braces returns an empty array, which is an answer;
 * null is the absence of one.
 */
export function blockExtents(text: string, language: string): BlockExtent[] | null {
  const syntax = SYNTAX[normalizeLanguage(language)];
  if (!syntax || !text || text.length > BLOCK_SCAN_CHAR_LIMIT) return null;

  const extents: BlockExtent[] = [];
  const open: number[] = [];
  let index = 0;

  const startsWith = (token: string) => text.startsWith(token, index);

  while (index < text.length) {
    // Comments first: a brace inside one is prose.
    const lineComment = syntax.line.find(startsWith);
    if (lineComment) {
      const newline = text.indexOf("\n", index);
      index = newline < 0 ? text.length : newline;
      continue;
    }
    const blockComment = syntax.block.find(([start]) => startsWith(start));
    if (blockComment) {
      const end = text.indexOf(blockComment[1], index + blockComment[0].length);
      if (end < 0) return null; // the rest of the file is comment, or is broken
      index = end + blockComment[1].length;
      continue;
    }

    if (syntax.rustRawStrings) {
      const raw = RUST_RAW_OPEN.exec(text.slice(index, index + 24));
      if (raw) {
        const terminator = `"${raw[1]}`;
        const end = text.indexOf(terminator, index + raw[0].length);
        if (end < 0) return null;
        index = end + terminator.length;
        continue;
      }
    }
    if (syntax.cppRawStrings) {
      const raw = CPP_RAW_OPEN.exec(text.slice(index, index + 24));
      if (raw) {
        const terminator = `)${raw[1]}"`;
        const end = text.indexOf(terminator, index + raw[0].length);
        if (end < 0) return null;
        index = end + terminator.length;
        continue;
      }
    }
    if (syntax.lifetimeAwareChar && text[index] === "'") {
      const char = RUST_CHAR.exec(text.slice(index, index + 16));
      // Not a character literal: a lifetime or a label, one tick and no more.
      index += char ? char[0].length : 1;
      continue;
    }

    const rule = syntax.strings.find((candidate) => startsWith(candidate.open));
    if (rule) {
      const consumed = skipString(text, index, rule);
      if (consumed < 0) return null; // never closed: structure past here unknown
      index = consumed;
      continue;
    }

    const char = text[index]!;
    if (char === "{") {
      open.push(index);
      index += 1;
      continue;
    }
    if (char === "}") {
      const start = open.pop();
      // A closer with nothing open means the scan lost its place (or the file
      // is broken); either way the extents are not trustworthy.
      if (start === undefined) return null;
      extents.push({ open: start, close: index });
      index += 1;
      continue;
    }
    index += 1;
  }

  // Something never closed. Common while typing, and fatal to a scope answer.
  if (open.length > 0) return null;
  extents.sort((a, b) => a.open - b.open);
  return extents;
}

/** Offset just past the string starting at `start`, or -1 if it never closes. */
function skipString(text: string, start: number, rule: StringRule): number {
  let index = start + rule.open.length;
  while (index < text.length) {
    if (rule.escape && text.startsWith(rule.escape, index)) {
      index += rule.escape.length + 1;
      continue;
    }
    if (!rule.across && text[index] === "\n") return -1;
    if (text.startsWith(rule.close, index)) {
      // `""` inside a C# verbatim string is one quote, not the end of it.
      if (rule.doubled && text.startsWith(rule.doubled, index)) {
        index += rule.doubled.length;
        continue;
      }
      return index + rule.close.length;
    }
    index += 1;
  }
  return -1;
}

/**
 * The blocks containing `offset`, innermost first.
 *
 * The braces themselves are outside: a caret ON a closing brace is in the
 * enclosing scope, which is where a reader would say it is.
 */
export function enclosingBlocks(
  extents: readonly BlockExtent[],
  offset: number,
): BlockExtent[] {
  return extents
    .filter((extent) => offset > extent.open && offset < extent.close)
    .sort((a, b) => a.close - a.open - (b.close - b.open));
}

/** The innermost block containing `offset`, or null when it is at file scope. */
export function innermostBlock(
  extents: readonly BlockExtent[],
  offset: number,
): BlockExtent | null {
  return enclosingBlocks(extents, offset)[0] ?? null;
}

/** Whether two block references are the same block; null is file scope. */
export function sameBlock(a: BlockExtent | null, b: BlockExtent | null): boolean {
  if (a === null || b === null) return a === b;
  return a.open === b.open && a.close === b.close;
}
