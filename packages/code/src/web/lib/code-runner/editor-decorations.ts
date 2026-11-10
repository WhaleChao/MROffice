/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the code editor's overlay decorations (VSCode parity):
 *  - bracket matching: given a caret, find the bracket adjacent to it and its
 *    depth-aware matching partner;
 *  - indent guides: the column positions of faint vertical guides per line;
 *  - active-line range: the [start, end) char span of the caret's line.
 *
 * Framework-free so the behaviour is unit-tested without a DOM; CodeEditor maps
 * the returned char/line/column coordinates into absolute-positioned overlay
 * elements in the highlight layer (it never mutates the textarea text).
 *
 * All helpers are total (never throw) and defensive about out-of-range carets,
 * empty input and CRLF line endings.
 */

/** The three bracket pairs we match (quotes are intentionally excluded). */
const OPEN_TO_CLOSE: Record<string, string> = {
  "(": ")",
  "[": "]",
  "{": "}",
};
const CLOSE_TO_OPEN: Record<string, string> = {
  ")": "(",
  "]": "[",
  "}": "{",
};

function isOpen(ch: string): boolean {
  return ch === "(" || ch === "[" || ch === "{";
}
function isClose(ch: string): boolean {
  return ch === ")" || ch === "]" || ch === "}";
}

/** A bracket position: its character index and the glyph at that index. */
export interface BracketPos {
  index: number;
  char: string;
}

/** A matched bracket pair (open always precedes close in the source). */
export interface BracketMatch {
  open: BracketPos;
  close: BracketPos;
}

/**
 * A per-character mask marking indices that live inside a string literal or a
 * comment, so bracket matching can skip them. This is a best-effort scanner for
 * C-family / JS / Python-ish syntax (it is language-agnostic by design): it
 * understands `'…'`, `"…"`, `` `…` `` strings with backslash escapes, `//` and
 * `#` line comments, and `/* … *​/` block comments. It never throws and always
 * returns a mask the exact length of `text`.
 */
export function computeSkipMask(text: string): boolean[] {
  const n = text.length;
  const mask = new Array<boolean>(n).fill(false);
  let i = 0;
  while (i < n) {
    const ch = text[i];
    // Line comments: // or # run to the end of the line.
    if ((ch === "/" && text[i + 1] === "/") || ch === "#") {
      while (i < n && text[i] !== "\n") {
        mask[i] = true;
        i += 1;
      }
      continue;
    }
    // Block comment: /* … */ (unterminated runs to EOF).
    if (ch === "/" && text[i + 1] === "*") {
      mask[i] = true;
      mask[i + 1] = true;
      i += 2;
      while (i < n) {
        if (text[i] === "*" && text[i + 1] === "/") {
          mask[i] = true;
          mask[i + 1] = true;
          i += 2;
          break;
        }
        mask[i] = true;
        i += 1;
      }
      continue;
    }
    // String literals: ' " ` with backslash escapes (unterminated runs to the
    // end of the line for ' and ", to EOF for template `` ` ``).
    if (ch === '"' || ch === "'" || ch === "`") {
      const quote = ch;
      mask[i] = true;
      i += 1;
      while (i < n) {
        const c = text[i];
        if (c === "\\") {
          // Escaped char: mask both the backslash and the next char.
          mask[i] = true;
          if (i + 1 < n) mask[i + 1] = true;
          i += 2;
          continue;
        }
        if ((quote === '"' || quote === "'") && c === "\n") break;
        mask[i] = true;
        i += 1;
        if (c === quote) break;
      }
      continue;
    }
    i += 1;
  }
  return mask;
}

/**
 * Find the bracket the caret is adjacent to (preferring the char immediately
 * BEFORE the caret, matching VSCode), then locate its depth-aware partner.
 *
 * `caretIndex` is the textarea selectionStart (0..text.length). Returns the
 * `{ open, close }` pair (open before close) or `null` when the caret is not
 * next to a bracket, the bracket is unbalanced, or the bracket sits inside a
 * string/comment. Brackets inside strings/comments are skipped on both the seed
 * lookup and the scan, via {@link computeSkipMask}.
 */
export function findMatchingBracket(
  text: string,
  caretIndex: number,
): BracketMatch | null {
  if (!text) return null;
  const n = text.length;
  const caret = Math.max(0, Math.min(caretIndex, n));
  const mask = computeSkipMask(text);

  // Candidate bracket: the char before the caret wins over the char after, so
  // typing right after a `)` highlights that pair (VSCode behaviour).
  let seed = -1;
  const before = caret - 1;
  if (before >= 0 && !mask[before] && (isOpen(text[before]) || isClose(text[before]))) {
    seed = before;
  } else if (caret < n && !mask[caret] && (isOpen(text[caret]) || isClose(text[caret]))) {
    seed = caret;
  }
  if (seed < 0) return null;

  const seedChar = text[seed];
  if (isOpen(seedChar)) {
    const wanted = OPEN_TO_CLOSE[seedChar];
    let depth = 0;
    for (let i = seed; i < n; i += 1) {
      if (mask[i]) continue;
      const c = text[i];
      if (c === seedChar) depth += 1;
      else if (c === wanted) {
        depth -= 1;
        if (depth === 0) {
          return { open: { index: seed, char: seedChar }, close: { index: i, char: c } };
        }
      }
    }
    return null;
  }

  // seedChar is a closer: scan backwards for its opener.
  const wanted = CLOSE_TO_OPEN[seedChar];
  let depth = 0;
  for (let i = seed; i >= 0; i -= 1) {
    if (mask[i]) continue;
    const c = text[i];
    if (c === seedChar) depth += 1;
    else if (c === wanted) {
      depth -= 1;
      if (depth === 0) {
        return { open: { index: i, char: c }, close: { index: seed, char: seedChar } };
      }
    }
  }
  return null;
}

/**
 * A colorized bracket: its character index, the glyph, and the nesting DEPTH the
 * colorizer cycles a CSS class through (depth 0 = outermost). Brackets inside
 * strings/comments are excluded (they carry no structural depth).
 */
export interface BracketDepth {
  index: number;
  char: string;
  depth: number;
}

/**
 * Assign a nesting DEPTH to every structural bracket in `text` (VSCode's
 * "bracket pair colorization"). Walking left-to-right, an opener takes the
 * CURRENT depth and increments it; its matching closer takes the depth it
 * dropped back to — so a pair shares one depth and sibling pairs at the same
 * level share a color while nested pairs step to the next. Brackets inside
 * strings/comments are skipped via {@link computeSkipMask} (so a `)` in a string
 * never shifts the color of real code).
 *
 * Unbalanced input is handled defensively: a stray closer with no open pair is
 * still reported (clamped to depth 0) rather than dropped, and depth never goes
 * negative. Different bracket FLAVOURS share one depth counter (matching VSCode,
 * which colors by nesting level, not by bracket kind). Pure and total.
 */
export function computeBracketDepths(text: string): BracketDepth[] {
  if (!text) return [];
  const mask = computeSkipMask(text);
  const result: BracketDepth[] = [];
  // Stack of the bracket FLAVOUR opened at each level, so a closer only pops a
  // matching opener; a mismatched closer is treated as depth-0 noise.
  const stack: string[] = [];
  for (let i = 0; i < text.length; i += 1) {
    if (mask[i]) continue;
    const ch = text[i];
    if (isOpen(ch)) {
      const depth = stack.length;
      stack.push(ch);
      result.push({ index: i, char: ch, depth });
    } else if (isClose(ch)) {
      const wantedOpener = CLOSE_TO_OPEN[ch];
      if (stack.length > 0 && stack[stack.length - 1] === wantedOpener) {
        stack.pop();
        result.push({ index: i, char: ch, depth: stack.length });
      } else {
        // Unbalanced / mismatched closer — color it at the base level rather
        // than disturbing the open stack.
        result.push({ index: i, char: ch, depth: 0 });
      }
    }
  }
  return result;
}

/**
 * Number of distinct colors the bracket colorizer cycles through; a bracket at
 * depth `d` uses class `cm-bracket-depth-${d % BRACKET_DEPTH_COLORS}`. Matches
 * the 6-color cycle the editor theme defines.
 */
export const BRACKET_DEPTH_COLORS = 6;

/** The CSS class for a bracket at `depth` (cycled through the color palette). */
export function bracketDepthClass(depth: number): string {
  const safe = ((Math.trunc(depth) % BRACKET_DEPTH_COLORS) + BRACKET_DEPTH_COLORS) % BRACKET_DEPTH_COLORS;
  return `cm-bracket-depth-${safe}`;
}

/** A single indent guide on a line: a 0-based column where a faint rule draws. */
export interface IndentGuide {
  /** 0-based line index. */
  line: number;
  /** 0-based column the guide sits at (a multiple of `tabSize`). */
  column: number;
}

/**
 * A single bracket-pair guide cell: a 0-based (line, column) where a thin
 * vertical connector draws for one matched multi-line bracket pair. Same shape
 * as {@link IndentGuide} so the CM6 layer renders both with the identical
 * per-line background-gradient mechanism (no source-byte mutation).
 */
export interface BracketGuide {
  /** 0-based line index a connector rule is drawn on. */
  line: number;
  /** 0-based column (char cell from the line start) the connector sits at. */
  column: number;
}

/**
 * Compute the thin vertical connectors VSCode draws between a matched bracket
 * pair (its "bracket pair guides"). For every BALANCED pair whose opener and
 * closer sit on DIFFERENT lines, a 1-cell-wide rule is drawn at the OPENING
 * bracket's column on each line strictly between the opener line and the closer
 * line, plus on the opener and closer lines themselves — visually tying the
 * block's `{ … }` / `( … )` / `[ … ]` together down the gutter.
 *
 * Built on the SAME structural scan as {@link computeBracketDepths}: a depth
 * stack of `{ char, index }` openers, popped by a matching closer; brackets
 * inside strings/comments are skipped via {@link computeSkipMask} (so a `)` in a
 * string never sprouts a guide). Single-line pairs contribute NOTHING (there is
 * no vertical span to connect), and an unbalanced / mismatched closer is ignored
 * (it pops nothing) — so a doc with no multi-line brackets yields `[]` and the
 * editor is byte-for-byte the pre-feature surface.
 *
 * Columns are measured in char cells from each line's start, matching the
 * monospace grid the editor renders on (tabs count as one cell here, same
 * convention as {@link indexToLineColumn}). The returned cells are sorted by
 * (line, column) ascending and de-duplicated, so two pairs that would draw a
 * rule at the same (line, column) collapse to one. Pure and total.
 */
export function computeBracketPairGuides(text: string): BracketGuide[] {
  if (!text) return [];
  const mask = computeSkipMask(text);

  // Line/column of every char index, computed in one pass so we don't re-scan
  // the whole buffer per bracket. `lineStart[i]` = offset of the start of the
  // line that index `i` lives on; `lineOf[i]` = that line's 0-based number.
  const n = text.length;
  const lineOf = new Array<number>(n + 1);
  const lineStart = new Array<number>(n + 1);
  let curLine = 0;
  let curStart = 0;
  for (let i = 0; i <= n; i += 1) {
    lineOf[i] = curLine;
    lineStart[i] = curStart;
    if (i < n && text[i] === "\n") {
      curLine += 1;
      curStart = i + 1;
    }
  }

  // Pairs collected as { openLine, closeLine, column } from a depth stack.
  const stack: { char: string; line: number; column: number }[] = [];
  const pairs: { openLine: number; closeLine: number; column: number }[] = [];
  for (let i = 0; i < n; i += 1) {
    if (mask[i]) continue;
    const ch = text[i];
    if (isOpen(ch)) {
      stack.push({ char: ch, line: lineOf[i], column: i - lineStart[i] });
    } else if (isClose(ch)) {
      const wantedOpener = CLOSE_TO_OPEN[ch];
      if (stack.length > 0 && stack[stack.length - 1].char === wantedOpener) {
        const opener = stack.pop() as { char: string; line: number; column: number };
        const closeLine = lineOf[i];
        if (closeLine > opener.line) {
          pairs.push({ openLine: opener.line, closeLine, column: opener.column });
        }
      }
      // A mismatched / unbalanced closer pops nothing — it cannot anchor a guide.
    }
  }

  // Expand each multi-line pair into one connector cell per spanned line and
  // de-duplicate identical (line, column) cells (sibling/nested pairs can align).
  const seen = new Set<string>();
  const guides: BracketGuide[] = [];
  for (const pair of pairs) {
    for (let line = pair.openLine; line <= pair.closeLine; line += 1) {
      const key = `${line}:${pair.column}`;
      if (seen.has(key)) continue;
      seen.add(key);
      guides.push({ line, column: pair.column });
    }
  }
  guides.sort((a, b) => a.line - b.line || a.column - b.column);
  return guides;
}

/**
 * Compute indent-guide columns for every line. A guide is drawn at each
 * indentation level strictly *inside* a line's leading whitespace: a line
 * indented by `k` levels gets guides at columns `0, tabSize, …, (k-1)*tabSize`.
 *
 * Tabs count as one `tabSize`-wide step; spaces are counted as written and
 * floored to whole levels (a half-level of trailing spaces yields no extra
 * guide). Blank / whitespace-only lines INHERIT the indentation of the nearest
 * following non-blank line, so guides run continuously through blank gaps the
 * way VSCode draws them. CRLF is handled (the trailing `\r` is ignored).
 */
export function computeIndentGuides(text: string, tabSize = 2): IndentGuide[] {
  const step = Math.max(1, Math.floor(tabSize) || 1);
  const rawLines = text.split("\n");
  const lines = rawLines.map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));

  // Indent width (in columns) of each line's leading whitespace, or null for a
  // blank / whitespace-only line (resolved by look-ahead below).
  const indentOf = lines.map((line) => {
    let width = 0;
    let i = 0;
    for (; i < line.length; i += 1) {
      const c = line[i];
      if (c === " ") width += 1;
      else if (c === "\t") width += step;
      else break;
    }
    // Whitespace-only (or empty) line: no own indentation level.
    if (i >= line.length) return null;
    return width;
  });

  // Resolve blank lines to the indent of the next non-blank line (inherit).
  const resolved = indentOf.slice();
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    if (resolved[i] === null) {
      resolved[i] = i + 1 < lines.length ? resolved[i + 1] : 0;
    }
  }

  const guides: IndentGuide[] = [];
  for (let line = 0; line < lines.length; line += 1) {
    const width = resolved[line] ?? 0;
    const levels = Math.floor(width / step);
    for (let level = 0; level < levels; level += 1) {
      guides.push({ line, column: level * step });
    }
  }
  return guides;
}

/** A half-open [start, end) char range for the active line. */
export interface LineRange {
  start: number;
  end: number;
}

/**
 * The char range of the line the caret sits on, as `[start, end)` where `start`
 * is the index just after the previous newline and `end` is the index of this
 * line's newline (or `text.length` for the last line). The trailing `\n` is
 * NOT included; CRLF callers get the `\r` inside the range, which is fine for a
 * full-width line-background overlay. Total/defensive about out-of-range carets.
 */
export function activeLineRange(text: string, caretIndex: number): LineRange {
  const n = text.length;
  const caret = Math.max(0, Math.min(caretIndex, n));
  const prevNl = text.lastIndexOf("\n", caret - 1);
  const start = prevNl === -1 ? 0 : prevNl + 1;
  const nextNl = text.indexOf("\n", caret);
  const end = nextNl === -1 ? n : nextNl;
  return { start, end };
}

/**
 * The 0-based line index the caret sits on. Convenience for the overlay, which
 * positions the active-line highlight by line rather than char offset.
 */
export function activeLineNumber(text: string, caretIndex: number): number {
  const caret = Math.max(0, Math.min(caretIndex, text.length));
  // Count newlines before the caret.
  let line = 0;
  for (let i = 0; i < caret; i += 1) {
    if (text[i] === "\n") line += 1;
  }
  return line;
}

/**
 * The 0-based (line, column) of a char index. Columns count chars from the line
 * start (tabs count as one column — the overlay uses the same monospace cell as
 * the textarea, so a literal `\t` advances one cell here too). Used to place the
 * bracket-match boxes. Defensive about out-of-range indices.
 */
export interface LineColumn {
  line: number;
  column: number;
}

export function indexToLineColumn(text: string, index: number): LineColumn {
  const idx = Math.max(0, Math.min(index, text.length));
  const prevNl = text.lastIndexOf("\n", idx - 1);
  const start = prevNl === -1 ? 0 : prevNl + 1;
  let line = 0;
  for (let i = 0; i < idx; i += 1) if (text[i] === "\n") line += 1;
  return { line, column: idx - start };
}
