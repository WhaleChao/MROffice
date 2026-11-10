/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Structural diagnostics for languages with no language server and no Lezer
 * grammar (C#, Swift).
 *
 * C, C++, Go, Java, Rust and YAML moved to real CodeMirror grammars, which
 * report where the parse actually broke — strictly better than counting
 * brackets, so they are deliberately NOT handled here any more. What remains is
 * the set that still has neither.
 *
 * A browser cannot host clangd or gopls, and those languages are highlighted
 * here by stream modes that build no syntax tree — so the whole class of
 * "this file does not close" errors went unreported, which is precisely the
 * error a parser flags first and the one a person stares past for minutes.
 *
 * This finds exactly that class, and nothing it cannot be sure of: brackets
 * that never close (or close what was never opened), a string left open at the
 * end of its line, and a block comment left open at the end of the file. It
 * knows enough about each language's string and comment syntax not to be fooled
 * by a brace inside a comment — the mistake that would make it worse than
 * useless. It never claims a file is CORRECT; a balanced file simply reports
 * nothing.
 */

/** Characters scanned before giving up; a generated blob is not worth flagging. */
export const BRACKET_SCAN_CHAR_LIMIT = 512 * 1024;
/** Diagnostics reported per document. */
export const BRACKET_DIAGNOSTIC_LIMIT = 50;

export interface BracketDiagnostic {
  from: number;
  to: number;
  severity: "error";
  /** One of a small closed set, so callers can localize. */
  code: "unclosed_bracket" | "unmatched_bracket" | "unterminated_string" | "unterminated_comment";
  /** The bracket or quote involved, for the message. */
  token: string;
}

interface LanguageSyntax {
  /** Line-comment starters. */
  line: readonly string[];
  /** Block comment delimiters. */
  block: readonly (readonly [string, string])[];
  /** Quote characters that must close on their own line. */
  quotes: readonly string[];
  /**
   * Quote sequences that may span lines (raw/multi-line strings).
   *
   * A third element, when present, is a sequence that is NOT a terminator even
   * though it starts with one — C#'s verbatim strings write an embedded quote
   * as `""`, so the first `"` of that pair must not end the string.
   */
  multiline: readonly (readonly [string, string, string?])[];
  /** Multi-line strings in which a backslash escapes the next character. */
  escapedMultiline?: readonly (readonly [string, string])[];
  /**
   * A quote that only starts a character literal when this matches at it
   * (Rust: `'x'`, `'\n'`); a quote it does not match is plain text (a
   * lifetime such as `'a`), never an unterminated string.
   */
  charLiteral?: RegExp;
  /**
   * The code lives only between these tags (PHP inside HTML): text outside
   * them is markup, where an apostrophe in "Don't" is not a string.
   */
  embedded?: { open: readonly string[]; close: string };
  /** Starts of a line comment that are something else (PHP's `#[` attribute). */
  notLineComment?: readonly string[];
  /** PHP heredoc/nowdoc: `<<<ID` … a line that is `ID`. */
  heredoc?: boolean;
}

const C_FAMILY: LanguageSyntax = {
  line: ["//"],
  block: [["/*", "*/"]],
  quotes: ['"', "'"],
  multiline: [],
};

/**
 * Languages this check runs for. Deliberately excludes everything with a real
 * engine (JS/TS, Python, JSON) or a Lezer tree (CSS, HTML, PHP, SQL), which
 * already report better diagnostics than bracket counting.
 */
const SYNTAX_BY_LANGUAGE: Readonly<Record<string, LanguageSyntax>> = {
  // Both of these span lines and both were missing, so the scanner reported an
  // unterminated string on source that is perfectly valid — a red squiggle on
  // correct code, which is worse than the silence this file is careful about
  // everywhere else.
  csharp: {
    ...C_FAMILY,
    // Raw string literals (C# 11) are listed first so they win over the
    // verbatim rule and over a bare quote.
    multiline: [
      ['"""', '"""'],
      ['@"', '"', '""'],
    ],
  },
  swift: {
    ...C_FAMILY,
    // Raw strings first, so a quote inside `#"…"#` does not end them.
    multiline: [
      ['#"""', '"""#'],
      ['##"', '"##'],
      ['#"', '"#'],
      ['"""', '"""'],
    ],
  },
  // Java's tree grammar is too old for Java 14–17 syntax, so it is checked
  // here; text blocks (Java 15) span lines.
  java: { ...C_FAMILY, multiline: [['"""', '"""']] },
  // Like Java, these tree grammars predate syntax people now write (Rust
  // let-else, PHP readonly promoted properties, C++20 templated lambdas).
  cpp: { ...C_FAMILY, multiline: [['R"(', ')"']] },
  php: {
    embedded: { open: ["<?php", "<?="], close: "?>" },
    line: ["//", "#"],
    notLineComment: ["#["],
    heredoc: true,
    block: [["/*", "*/"]],
    quotes: [],
    multiline: [],
    // PHP strings may span lines.
    escapedMultiline: [['"', '"'], ["'", "'"]],
  },
  rust: {
    line: ["//"],
    block: [["/*", "*/"]],
    quotes: [],
    multiline: [
      ['r###"', '"###'],
      ['r##"', '"##'],
      ['r#"', '"#'],
      ['br"', '"'],
      ['r"', '"'],
    ],
    // Rust strings may span lines.
    escapedMultiline: [['"', '"']],
    charLiteral: /'(?:\\(?:u\{[0-9a-fA-F]{1,6}\}|x[0-9a-fA-F]{2}|.)|[^\\'\n])'/y,
  },
};

export function supportsBracketDiagnostics(language: string): boolean {
  return language in SYNTAX_BY_LANGUAGE;
}

const OPENERS: Readonly<Record<string, string>> = {
  "(": ")",
  "[": "]",
  "{": "}",
};
const CLOSERS = new Set([")", "]", "}"]);

/**
 * Report unbalanced structure in `text`. Returns an empty list for a balanced
 * document, an unsupported language, or input past the scan limit — an empty
 * list never means "checked and correct", only "nothing to report".
 */
export function bracketDiagnostics(
  text: string,
  language: string,
  limits: { scanLimit?: number; maxDiagnostics?: number } = {},
): BracketDiagnostic[] {
  const syntax = SYNTAX_BY_LANGUAGE[language];
  const scanLimit = limits.scanLimit ?? BRACKET_SCAN_CHAR_LIMIT;
  const maxDiagnostics = limits.maxDiagnostics ?? BRACKET_DIAGNOSTIC_LIMIT;
  if (!syntax || !text || text.length > scanLimit) return [];

  const diagnostics: BracketDiagnostic[] = [];
  const stack: Array<{ char: string; at: number }> = [];
  let index = 0;
  // Set when the scan stops early because something never closed. Past that
  // point the file's structure is unknown, so the brackets left on the stack
  // are a consequence of that error rather than errors of their own.
  let aborted = false;

  const startsWith = (token: string) => text.startsWith(token, index);
  let inCode = !syntax.embedded;

  while (index < text.length && diagnostics.length < maxDiagnostics) {
    const char = text[index]!;

    if (syntax.embedded) {
      if (!inCode) {
        const open = syntax.embedded.open.find(startsWith);
        if (open) {
          inCode = true;
          index += open.length;
        } else {
          index += 1;
        }
        continue;
      }
      if (startsWith(syntax.embedded.close)) {
        inCode = false;
        index += syntax.embedded.close.length;
        continue;
      }
    }

    if (syntax.heredoc && startsWith("<<<")) {
      const head = /<<<[ \t]*(["']?)([A-Za-z_][A-Za-z0-9_]*)\1\r?\n/y;
      head.lastIndex = index;
      const opened = head.exec(text);
      if (opened) {
        const end = new RegExp(`^[ \\t]*${opened[2]}\\b`, "m");
        const rest = text.slice(head.lastIndex);
        const found = end.exec(rest);
        if (!found) {
          diagnostics.push({
            from: index,
            to: Math.min(text.length, index + 3),
            severity: "error",
            code: "unterminated_string",
            token: "<<<",
          });
          aborted = true;
          break;
        }
        index = head.lastIndex + found.index + found[0].length;
        continue;
      }
    }

    // Comments first: everything inside one is text, including brackets.
    const lineComment = syntax.notLineComment?.some(startsWith)
      ? undefined
      : syntax.line.find(startsWith);
    if (lineComment) {
      const newline = text.indexOf("\n", index);
      index = newline < 0 ? text.length : newline;
      continue;
    }
    const block = syntax.block.find(([open]) => startsWith(open));
    if (block) {
      const end = text.indexOf(block[1], index + block[0].length);
      if (end < 0) {
        diagnostics.push({
          from: index,
          to: Math.min(text.length, index + block[0].length),
          severity: "error",
          code: "unterminated_comment",
          token: block[0],
        });
        aborted = true;
        break;
      }
      index = end + block[1].length;
      continue;
    }

    // Multi-line strings before single-line quotes, so ``` and """ win over " .
    const multiline = syntax.multiline.find(([open]) => startsWith(open));
    if (multiline) {
      const [open, close, escaped] = multiline;
      let end = text.indexOf(close, index + open.length);
      // Skip terminators that are really an escaped quote inside the string.
      while (end >= 0 && escaped && text.startsWith(escaped, end)) {
        end = text.indexOf(close, end + escaped.length);
      }
      if (end < 0) {
        diagnostics.push({
          from: index,
          to: Math.min(text.length, index + open.length),
          severity: "error",
          code: "unterminated_string",
          token: open,
        });
        aborted = true;
        break;
      }
      index = end + close.length;
      continue;
    }

    const escapedMultiline = syntax.escapedMultiline?.find(([open]) =>
      startsWith(open),
    );
    if (escapedMultiline) {
      const [open, close] = escapedMultiline;
      const start = index;
      index += open.length;
      let closed = false;
      while (index < text.length) {
        if (text[index] === "\\") {
          index += 2;
          continue;
        }
        if (text.startsWith(close, index)) {
          closed = true;
          index += close.length;
          break;
        }
        index += 1;
      }
      if (!closed) {
        diagnostics.push({
          from: start,
          to: Math.min(text.length, start + open.length),
          severity: "error",
          code: "unterminated_string",
          token: open,
        });
        aborted = true;
        break;
      }
      continue;
    }

    if (char === "'" && syntax.charLiteral) {
      syntax.charLiteral.lastIndex = index;
      const literal = syntax.charLiteral.exec(text);
      index += literal ? literal[0].length : 1;
      continue;
    }

    if (syntax.quotes.includes(char)) {
      const start = index;
      index += 1;
      let closed = false;
      while (index < text.length) {
        const inner = text[index]!;
        if (inner === "\\") {
          index += 2;
          continue;
        }
        if (inner === "\n") break; // a plain string may not cross a line
        if (inner === char) {
          closed = true;
          index += 1;
          break;
        }
        index += 1;
      }
      if (!closed) {
        diagnostics.push({
          from: start,
          to: Math.min(text.length, start + 1),
          severity: "error",
          code: "unterminated_string",
          token: char,
        });
      }
      continue;
    }

    if (char in OPENERS) {
      stack.push({ char, at: index });
      index += 1;
      continue;
    }
    if (CLOSERS.has(char)) {
      const open = stack.pop();
      if (!open || OPENERS[open.char] !== char) {
        // Closing something that was never opened (or the wrong thing) is the
        // error; putting the marker on the CLOSER is where the reader is.
        if (open) stack.push(open);
        diagnostics.push({
          from: index,
          to: index + 1,
          severity: "error",
          code: "unmatched_bracket",
          token: char,
        });
      }
      index += 1;
      continue;
    }
    index += 1;
  }

  // Whatever is still open never closed. Report the outermost first: that is
  // usually the real mistake, and the rest are its consequence.
  if (aborted) return diagnostics;
  for (const open of stack) {
    if (diagnostics.length >= maxDiagnostics) break;
    diagnostics.push({
      from: open.at,
      to: open.at + 1,
      severity: "error",
      code: "unclosed_bracket",
      token: open.char,
    });
  }
  return diagnostics;
}
