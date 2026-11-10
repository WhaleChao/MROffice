/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Built-in code-snippet library for the MRNote code editor (VSCode-style
 * snippets). Each supported language has a small set of common boilerplate
 * snippets (loops, functions, try/catch, …) that the user triggers by typing a
 * snippet's `prefix` and accepting it from the autocomplete popup — the body is
 * inserted in place of the typed prefix.
 *
 * Framework-free + unit-tested (no React, no CodeMirror). The autocomplete glue
 * (`autocomplete.ts` for the textarea editor, `codemirror-completions.ts` for
 * CM6) surfaces matching snippets as completion items ALONGSIDE the existing
 * keyword/identifier completions, and the editors reuse {@link expandSnippet}
 * to compute the inserted text + final caret position.
 *
 * Snippet bodies support LSP-style TAB-STOP syntax (a subset of TextMate/VSCode):
 *  - `$0` / `${0}`        — the FINAL caret position (after every other stop).
 *  - `$1`, `$2`, …        — numbered tab-stops, cycled with Tab / Shift-Tab.
 *  - `${1:label}`         — a tab-stop with default text, PRE-SELECTED on arrival.
 *  - a repeated index     — a MIRROR: typing in one updates all instances at once.
 * A body with no tab-stop at all puts the caret at the end of the inserted text.
 *
 * Two consumption paths share this one body format:
 *  - {@link expandSnippet} (the legacy single-`$0` path) still computes a plain
 *    inserted text + final caret offset; it understands `$0`/`${0}` and treats any
 *    OTHER tab-stop token as plain text, so it degrades gracefully where the
 *    interactive engine is not wired.
 *  - {@link toCmSnippetTemplate} converts a body into a CodeMirror 6 snippet
 *    template string for `@codemirror/autocomplete`'s `snippet()`, which drives
 *    the live, interactive tab-stop / placeholder / mirror experience. The pure
 *    {@link parseSnippetTabStops} extracts the ordered tab-stop list + mirrors for
 *    testing without an editor.
 *
 * Multi-line bodies are written with `\n` and NO leading indentation on
 * continuation lines; {@link expandSnippet} re-indents continuation lines to the
 * caret's current column so a snippet inserted on an indented line stays aligned
 * (matching how editors expand snippets).
 */

import { userSnippetsFor } from "./user-snippets";

/** The marker in a snippet body for the final caret position. */
export const SNIPPET_CURSOR = "$0";

export interface Snippet {
  /** What the user types to surface the snippet in autocomplete (case-insensitive). */
  prefix: string;
  /** Display label in the completion popup (usually the prefix). */
  label: string;
  /** Short human description of what the snippet inserts. */
  description: string;
  /**
   * The snippet body. `\n`-separated; continuation lines are written WITHOUT
   * leading indentation (re-indented on insert). May contain a single `$0`
   * marking the final caret position (stripped from the inserted text).
   */
  body: string;
}

/**
 * Snippets keyed by canonical language id. TypeScript reuses the JavaScript set
 * (same syntax for these constructs); see {@link snippetLanguageKey}. Only
 * languages present here have snippets; everything else yields an empty set, so
 * the feature is additive and never interferes with unsupported languages.
 */
const SNIPPETS: Record<string, Snippet[]> = {
  javascript: [
    {
      prefix: "log",
      label: "log",
      description: "console.log",
      body: "console.log($0);",
    },
    {
      prefix: "fn",
      label: "fn",
      description: "Function declaration",
      body: "function name(params) {\n  $0\n}",
    },
    {
      prefix: "afn",
      label: "afn",
      description: "Arrow function",
      body: "const name = (params) => {\n  $0\n};",
    },
    {
      prefix: "asfn",
      label: "asfn",
      description: "Async function",
      body: "async function name(params) {\n  $0\n}",
    },
    {
      prefix: "for",
      label: "for",
      description: "for loop (index)",
      body: "for (let i = 0; i < arr.length; i++) {\n  $0\n}",
    },
    {
      prefix: "forof",
      label: "forof",
      description: "for…of loop",
      body: "for (const item of iterable) {\n  $0\n}",
    },
    {
      prefix: "trycatch",
      label: "trycatch",
      description: "try / catch",
      body: "try {\n  $0\n} catch (err) {\n  console.error(err);\n}",
    },
    {
      prefix: "map",
      label: "map",
      description: "Array.map",
      body: "arr.map((item) => $0)",
    },
    {
      prefix: "filter",
      label: "filter",
      description: "Array.filter",
      body: "arr.filter((item) => $0)",
    },
    {
      prefix: "reduce",
      label: "reduce",
      description: "Array.reduce",
      body: "arr.reduce((acc, item) => $0, initial)",
    },
    {
      prefix: "ifelse",
      label: "ifelse",
      description: "if / else",
      body: "if ($0) {\n} else {\n}",
    },
    {
      prefix: "promise",
      label: "promise",
      description: "new Promise",
      body: "new Promise((resolve, reject) => {\n  $0\n})",
    },
  ],
  python: [
    {
      prefix: "def",
      label: "def",
      description: "Function definition",
      body: "def name(params):\n    $0",
    },
    {
      prefix: "class",
      label: "class",
      description: "Class with __init__",
      body: "class Name:\n    def __init__(self):\n        $0",
    },
    {
      prefix: "for",
      label: "for",
      description: "for loop",
      body: "for item in iterable:\n    $0",
    },
    {
      prefix: "while",
      label: "while",
      description: "while loop",
      body: "while $0:\n    pass",
    },
    {
      prefix: "ifmain",
      label: "ifmain",
      description: 'if __name__ == "__main__"',
      body: 'if __name__ == "__main__":\n    $0',
    },
    {
      prefix: "listcomp",
      label: "listcomp",
      description: "List comprehension",
      body: "[$0 for item in iterable]",
    },
    {
      prefix: "dictcomp",
      label: "dictcomp",
      description: "Dict comprehension",
      body: "{key: value for item in iterable}",
    },
    {
      prefix: "tryexcept",
      label: "tryexcept",
      description: "try / except",
      body: "try:\n    $0\nexcept Exception as err:\n    print(err)",
    },
    {
      prefix: "withopen",
      label: "withopen",
      description: "with open(...) as f",
      body: 'with open(path) as f:\n    $0',
    },
    {
      prefix: "import",
      label: "import",
      description: "from … import …",
      body: "from module import name",
    },
    {
      prefix: "print",
      label: "print",
      description: "print()",
      body: "print($0)",
    },
  ],
};

/**
 * Map a language id to the key used in {@link SNIPPETS}. TypeScript shares the
 * JavaScript snippet set (identical syntax for these constructs); every other id
 * maps to itself. Mirrors `keywordsFor` in autocomplete.ts so snippets and
 * keyword completions agree on language aliasing.
 */
export function snippetLanguageKey(language?: string): string {
  const lang = (language || "").toLowerCase();
  if (lang === "typescript") return "javascript";
  return lang;
}

/**
 * All snippets for a language: the built-in set MERGED with any USER-defined
 * snippets that apply (language-specific ones + language-agnostic ones).
 *
 * Byte-stable / additive: when the user has defined NO snippets,
 * {@link userSnippetsFor} returns `[]` and this returns the built-in array
 * UNCHANGED (same identity), so the completion output is byte-for-byte what it
 * was before this feature. The merge cost is only paid once a user snippet
 * exists.
 *
 * COLLISION RULE: a user snippet whose prefix equals a built-in's prefix
 * OVERRIDES the built-in (the user's intent wins — they explicitly redefined
 * that trigger). User snippets are appended AFTER the built-ins, with any
 * built-in sharing a user prefix removed first, so exactly one snippet exists
 * per prefix and it's the user's.
 */
export function snippetsFor(language?: string): Snippet[] {
  const key = snippetLanguageKey(language);
  const builtIns = SNIPPETS[key] ?? [];
  const userSnippets = userSnippetsFor(key);
  // No user snippets -> return the built-in array verbatim (byte-stable).
  if (userSnippets.length === 0) return builtIns;
  const userPrefixes = new Set(userSnippets.map((s) => s.prefix));
  // Drop any built-in whose prefix the user redefined, then append the user set.
  return [
    ...builtIns.filter((s) => !userPrefixes.has(s.prefix)),
    ...userSnippets,
  ];
}

/**
 * Snippets whose prefix matches the typed word, ranked the same way keyword
 * completions are (case-sensitive prefix first, then shorter, then alpha). A
 * cheap O(#snippets) prefix scan — there are ~10 snippets per language, so this
 * adds negligible cost to the (already debounced) autocomplete pass.
 *
 * The match is case-insensitive on the prefix (so `LOG` still finds `log`), and
 * — like keyword completion — a snippet whose prefix EQUALS the typed word is
 * excluded (there is nothing left to complete). Returns `[]` when nothing
 * matches, so callers can treat "no snippets" uniformly.
 */
export function matchSnippets(prefix: string, language?: string): Snippet[] {
  if (!prefix) return [];
  const lower = prefix.toLowerCase();
  return snippetsFor(language)
    .filter(
      (s) =>
        s.prefix.length > prefix.length &&
        s.prefix.toLowerCase().startsWith(lower),
    )
    .sort((a, b) => {
      const aExact = a.prefix.startsWith(prefix) ? 0 : 1;
      const bExact = b.prefix.startsWith(prefix) ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      if (a.prefix.length !== b.prefix.length)
        return a.prefix.length - b.prefix.length;
      return a.prefix.localeCompare(b.prefix);
    });
}

/** Look up a single snippet by its exact prefix for a language (case-sensitive prefix id). */
export function snippetByPrefix(
  prefix: string,
  language?: string,
): Snippet | undefined {
  return snippetsFor(language).find((s) => s.prefix === prefix);
}

export interface ExpandedSnippet {
  /** The text to insert in place of the typed prefix. */
  text: string;
  /**
   * Caret offset WITHIN {@link text} after insertion: the index of the `$0`
   * marker (stripped), or `text.length` when the body had no `$0`.
   */
  cursorOffset: number;
}

/**
 * Expand a snippet body for insertion at the caret.
 *
 * - `indent` is the leading whitespace of the line the caret sits on; every
 *   continuation line (i.e. each line after the first) is prefixed with it so a
 *   snippet inserted on an indented line stays aligned. The first line is left
 *   as-is because the typed prefix it replaces already sits at the caret column.
 * - The single `$0` marker (if any) is removed and its position recorded as
 *   {@link ExpandedSnippet.cursorOffset}; with no marker the caret lands at the
 *   end of the inserted text.
 *
 * Pure string math — no editor state — so it is unit-tested directly and reused
 * by both the textarea and CodeMirror accept paths.
 */
export function expandSnippet(body: string, indent = ""): ExpandedSnippet {
  // Re-indent continuation lines to the caret column.
  const reindented =
    indent.length === 0
      ? body
      : body
          .split("\n")
          .map((line, i) => (i === 0 ? line : indent + line))
          .join("\n");

  const markerIndex = reindented.indexOf(SNIPPET_CURSOR);
  if (markerIndex === -1) {
    return { text: reindented, cursorOffset: reindented.length };
  }
  const text =
    reindented.slice(0, markerIndex) +
    reindented.slice(markerIndex + SNIPPET_CURSOR.length);
  return { text, cursorOffset: markerIndex };
}

/**
 * Leading-whitespace (indentation) of the line containing `caret` in `value`.
 * Used to align multi-line snippet bodies on insert. Exported so the editors
 * compute it identically.
 */
export function lineIndentAt(value: string, caret: number): string {
  const lineStart = value.lastIndexOf("\n", caret - 1) + 1;
  const line = value.slice(lineStart, caret);
  const match = /^[ \t]*/.exec(line);
  return match ? match[0] : "";
}

// --- Interactive tab-stops / placeholders / mirrors -------------------------
//
// The snippet body grammar (a TextMate/LSP subset) recognizes these tokens:
//   $0   ${0}            final caret position
//   $1   ${1}            numbered tab-stop (a repeat of the same number mirrors)
//   ${1:label}           numbered tab-stop with PRE-SELECTED default text
// Everything else in a body — including literal `{`/`}` and a `$` not followed by
// a digit or `{` — is plain text. We normalize bodies into CodeMirror 6's snippet
// template syntax (which ONLY understands the braced `${…}` form) so the native
// `snippet()` engine drives the interactive experience.

/** One parsed tab-stop occurrence within a snippet body. */
export interface SnippetTabStop {
  /** The stop's order index (`$1` -> 1, `${0}` -> 0). `0` is always last. */
  index: number;
  /** Default/placeholder text (`${1:foo}` -> "foo"); "" when none. */
  placeholder: string;
  /** 0-based char position of the token's START within the ORIGINAL body. */
  from: number;
  /** 0-based char position just past the token's END within the ORIGINAL body. */
  to: number;
}

// A single token matcher for the body grammar, applied with a sticky/global
// regex so we can walk the body left-to-right. Order of alternatives matters:
// the braced forms are tried before the bare `$N` so `${1:x}` isn't mis-read as
// bare `$` + text. `\$` (escaped dollar) is matched first so it's treated as a
// literal, never a tab-stop. Group map:
//   1: escaped dollar "\$"      (literal — not a stop)
//   2: braced index "${1}"/"${0}"           -> numbered, no placeholder
//   3: braced index of "${1:label}"
//   4: placeholder label of "${1:label}"
//   5: bare index "$1"/"$0"                  -> numbered, no placeholder
const SNIPPET_TOKEN_RE =
  /(\\\$)|\$\{(\d+)\}|\$\{(\d+):([^{}]*)\}|\$(\d+)/g;

/**
 * Parse a snippet body into its ordered list of tab-stop OCCURRENCES (in the
 * order they appear in the body — mirrors of the same index appear multiple
 * times). Pure string analysis — no editor state — so the tab-stop/mirror model
 * is unit-tested directly. An escaped `\$` is NOT a tab-stop. The legacy single
 * `$0` is reported as one stop with index 0, so old bodies parse identically.
 */
export function parseSnippetTabStops(body: string): SnippetTabStop[] {
  const stops: SnippetTabStop[] = [];
  SNIPPET_TOKEN_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = SNIPPET_TOKEN_RE.exec(body)) !== null) {
    // Group 1 = an escaped dollar: literal, skip (but it DID consume input).
    if (m[1] !== undefined) continue;
    let index: number;
    let placeholder = "";
    if (m[2] !== undefined) {
      index = Number.parseInt(m[2], 10);
    } else if (m[3] !== undefined) {
      index = Number.parseInt(m[3], 10);
      placeholder = m[4] ?? "";
    } else {
      // m[5]: bare $N
      index = Number.parseInt(m[5] as string, 10);
    }
    stops.push({ index, placeholder, from: m.index, to: m.index + m[0].length });
  }
  return stops;
}

/**
 * Unique tab-stop indices in CYCLE ORDER: ascending positive indices first
 * (1, 2, 3, …) then `0` last (the final caret), matching how Tab cycles through
 * a snippet. A mirrored index appears once here (it's one logical stop). Used by
 * tests + callers that need the field order without an editor.
 */
export function snippetTabStopOrder(body: string): number[] {
  const seen = new Set<number>();
  for (const s of parseSnippetTabStops(body)) seen.add(s.index);
  const positives = [...seen].filter((n) => n > 0).sort((a, b) => a - b);
  return seen.has(0) ? [...positives, 0] : positives;
}

/** True when a body contains ANY interactive tab-stop beyond a lone final `$0`. */
export function hasInteractiveTabStops(body: string): boolean {
  return parseSnippetTabStops(body).some((s) => s.index !== 0);
}

/** True when a body already declares a FINAL stop (`$0` / `${0}`). */
export function hasFinalTabStop(body: string): boolean {
  return parseSnippetTabStops(body).some((s) => s.index === 0);
}

/**
 * Ensure a body ends the caret in the same place the legacy single-`$0` path
 * did. CodeMirror's `snippet()` leaves the caret at the FIRST field when a
 * template has no `${0}`, but the legacy {@link expandSnippet} (and every editor
 * before interactive stops) put the caret at the END of a body with no `$0`. So
 * for a body that declares NO final stop we append `$0` at the end, preserving
 * the prior caret-at-end behaviour exactly; a body that already has `$0`/`${0}`
 * is returned unchanged. Apply this BEFORE {@link toCmSnippetTemplate}.
 */
export function ensureFinalTabStop(body: string): string {
  return hasFinalTabStop(body) ? body : `${body}${SNIPPET_CURSOR}`;
}

/**
 * Convert a snippet body into a CodeMirror 6 snippet TEMPLATE string consumable
 * by `@codemirror/autocomplete`'s `snippet()`.
 *
 * CM6 only recognizes the BRACED field form (`${1}`, `${1:label}`, `${0}`); a
 * bare `$1`/`$0` is inserted literally. So we rewrite every recognized tab-stop
 * token into its braced form and leave ALL other text — including literal
 * `{`/`}` and a `$` that doesn't introduce a stop — byte-for-byte untouched (CM6
 * leaves non-`${`/`#{` braces and dollars literal, so existing bodies such as
 * `if ($0) {\n} else {\n}` round-trip with their braces intact). An escaped
 * `\$` is preserved as-is so it stays a literal dollar.
 *
 * Pure: a body with only `$0` becomes `${0}` (CM places the caret there); a body
 * with no tab-stop is returned unchanged (CM appends the caret at the end).
 */
export function toCmSnippetTemplate(body: string): string {
  let out = "";
  let last = 0;
  SNIPPET_TOKEN_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = SNIPPET_TOKEN_RE.exec(body)) !== null) {
    out += body.slice(last, m.index);
    last = m.index + m[0].length;
    if (m[1] !== undefined) {
      // Escaped dollar: keep literal (CM emits a literal `$`).
      out += m[1];
    } else if (m[2] !== undefined) {
      out += `\${${m[2]}}`;
    } else if (m[3] !== undefined) {
      // Preserve the placeholder text verbatim (it can't contain braces by the
      // grammar, so it's already a safe CM placeholder body).
      out += `\${${m[3]}:${m[4] ?? ""}}`;
    } else {
      // Bare $N -> braced.
      out += `\${${m[5]}}`;
    }
  }
  out += body.slice(last);
  return out;
}
