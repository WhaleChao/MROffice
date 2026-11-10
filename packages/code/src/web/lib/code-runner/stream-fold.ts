/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure fold-range computation for the CodeMirror 6 `foldService` facet.
 *
 * The stream-mode languages (go / rust / java / c / cpp / csharp / ruby /
 * swift / r / perl / lua / yaml / toml / bash) are highlighted by
 * `@codemirror/legacy-modes` stream parsers, which produce NO Lezer syntax
 * tree. CM6's default folding asks the syntax tree for fold ranges via
 * `foldNodeProp`, so those languages render a fold gutter that never actually
 * folds. CM6 consults the {@link https://codemirror.net/6/docs/ref/#language.foldService
 * `foldService`} facet FIRST (before the tree), and a foldService callback only
 * needs the document text + a line's offset bounds — no tree — so it is the
 * right hook to give the stream-mode languages real, tree-free folding.
 *
 * A `foldService` callback has the shape
 *   `(state, lineStart, lineEnd) => { from, to } | null`
 * where `lineStart`/`lineEnd` are the START/END document offsets of the line CM6
 * is testing for foldability, and the returned `{ from, to }` is the half-open
 * document range to HIDE when that line's fold marker is collapsed (the header
 * line itself stays visible; folding hides from `from` to `to`). Returning
 * `null` means "this line is not a fold header".
 *
 * Everything here is framework-free string/offset math (it takes the doc text +
 * an offset, not an EditorState), so it is fully unit-testable in node and the
 * thin CM6 wiring (`streamFoldExtension` in editor-cm-extensions.ts) just adapts
 * an `EditorState` to these functions.
 *
 * Two complementary strategies, tried in order by the wiring:
 *  1. EXPLICIT region markers — `#region` / `#endregion` (and the
 *     `// #region` / `/* #region *\/` / `<!-- #region -->` comment variants),
 *     matching VSCode's marker folding. See {@link computeRegionFoldRange}.
 *  2. INDENTATION — a line whose following non-blank line is more deeply
 *     indented opens a region running to the end of that more-indented body.
 *     See {@link computeIndentFoldRange}.
 *
 * All helpers are total (never throw) and defensive about empty input, blank /
 * whitespace-only lines, and CRLF endings (a trailing `\r` is treated as line
 * content, never as indentation).
 */

/** A half-open document range `[from, to)` to hide when a fold is collapsed. */
export interface FoldOffsetRange {
  from: number;
  to: number;
}

/**
 * The visual indent width (in columns) of a line's leading whitespace. Tabs
 * advance to the next `step`-wide tab stop is intentionally NOT modelled here:
 * for fold-header DETECTION we only ever compare two lines' indents for
 * strict greater/less-or-equal, so a simple "tab = `step` columns, space = 1
 * column" measure is sufficient and matches {@link
 * (computeFoldRegions) fold-regions.ts}. Returns `null` for a blank /
 * whitespace-only line (it has no indentation level of its own).
 */
function indentWidth(line: string, step: number): number | null {
  let width = 0;
  let i = 0;
  for (; i < line.length; i += 1) {
    const c = line[i];
    if (c === " ") width += 1;
    else if (c === "\t") width += step;
    else break;
  }
  if (i >= line.length) return null; // blank / whitespace-only
  return width;
}

/** Strip a single trailing carriage return so CRLF lines measure like LF. */
function stripCr(line: string): string {
  return line.endsWith("\r") ? line.slice(0, -1) : line;
}

/**
 * Split `text` into lines (LF, CRLF-safe) plus the START offset of each line in
 * the original document. The offsets let a per-line fold computation map a
 * 0-based line index back to a document offset (what the CM6 foldService needs).
 *
 * The returned `starts[i]` is the offset of the first character of line `i`;
 * `starts[lines.length]` is NOT included (callers derive an end offset from the
 * next line's start, or the document length for the last line).
 */
/**
 * The most recent split, kept because CM6 asks the fold services about MANY
 * lines against the SAME document: every visible line, every time the fold
 * gutter renders. Without this, one keystroke in a 21,000-line file re-split
 * half a megabyte dozens of times — profiling put ~50% of keystroke CPU in this
 * function alone.
 *
 * One entry, keyed by the exact text: a new document (any edit) replaces it, so
 * this holds at most one line array — strictly less memory churn than
 * allocating a fresh one per query, and it never outlives the current document.
 *
 * The returned arrays are SHARED. Callers here only read them; anything that
 * needs to mutate must copy first.
 */
let cachedSplitText: string | null = null;
let cachedSplit: { lines: string[]; starts: number[] } | null = null;

export function splitLinesWithOffsets(text: string): {
  lines: string[];
  starts: number[];
} {
  if (text === cachedSplitText && cachedSplit) return cachedSplit;
  const lines: string[] = [];
  const starts: number[] = [];
  let start = 0;
  for (let i = 0; i <= text.length; i += 1) {
    if (i === text.length || text[i] === "\n") {
      lines.push(text.slice(start, i));
      starts.push(start);
      start = i + 1;
    }
  }
  const split = { lines, starts };
  cachedSplitText = text;
  cachedSplit = split;
  return split;
}

/**
 * Resolve the 0-based line index whose START offset is exactly `lineStart`.
 * CM6 always hands the foldService a real line-start offset, so this is an exact
 * lookup; it returns `-1` for an offset that is not a line start (defensive — a
 * caller passing an interior offset gets a clean "no fold" rather than a wrong
 * line). Linear scan is fine: the wiring calls it once per fold query, and the
 * starts array is already built for the same text.
 */
function lineIndexAtStart(starts: number[], lineStart: number): number {
  for (let i = 0; i < starts.length; i += 1) {
    if (starts[i] === lineStart) return i;
    if (starts[i] > lineStart) return -1;
  }
  return -1;
}

/**
 * INDENTATION-based fold range for the line starting at offset `lineStart`.
 *
 * The line is a fold header when the next NON-BLANK line is strictly more
 * indented than it. The folded range then runs from the END of the header line
 * (so the header text stays visible) to the END of the last line of the
 * more-indented body — trailing blank lines just before the dedent are NOT
 * swallowed (they belong to the surrounding scope, matching VSCode and
 * {@link (computeFoldRegions) fold-regions.ts}).
 *
 * @param text      The whole document.
 * @param lineStart Document offset of the first char of the candidate header line.
 * @param tabSize   Columns per tab stop for the indent measure (>= 1; floored).
 * @returns The `[from, to)` range to hide, or `null` when the line is not a
 *          header (blank, last line, or its body is not more indented).
 */
export function computeIndentFoldRange(
  text: string,
  lineStart: number,
  tabSize = 2,
): FoldOffsetRange | null {
  if (!text) return null;
  const step = Math.max(1, Math.floor(tabSize) || 1);
  const { lines, starts } = splitLinesWithOffsets(text);
  const h = lineIndexAtStart(starts, lineStart);
  if (h < 0) return null;
  const n = lines.length;

  const headerIndent = indentWidth(stripCr(lines[h]), step);
  if (headerIndent === null) return null; // blank header never folds

  // The next non-blank line decides whether this is a header.
  let j = h + 1;
  while (j < n && indentWidth(stripCr(lines[j]), step) === null) j += 1;
  if (j >= n) return null;
  const nextIndent = indentWidth(stripCr(lines[j]), step);
  if (nextIndent === null || nextIndent <= headerIndent) return null;

  // Walk the body: the last NON-BLANK line whose indent stays > header level.
  let end = j;
  let k = j;
  while (k < n) {
    const ind = indentWidth(stripCr(lines[k]), step);
    if (ind === null) {
      k += 1;
      continue; // blank: tentatively inside, but never extends `end`
    }
    if (ind <= headerIndent) break; // dedent: body ended before this line
    end = k;
    k += 1;
  }
  if (end <= h) return null;

  // Hide from the END of the header line to the END of the last body line.
  const from = starts[h] + lines[h].length;
  const endLineStart = starts[end];
  const to = endLineStart + lines[end].length;
  if (to <= from) return null;
  return { from, to };
}

// ---------------------------------------------------------------------------
// Explicit #region / #endregion marker folding (VSCode "marker" folding).
// ---------------------------------------------------------------------------

/**
 * Whether a line is a `#region` START marker, accepting the bare form and the
 * common comment-prefixed variants so it works across the stream-mode AND
 * tree-backed languages (and HTML-ish markup):
 *   `#region [label]`            (C# / PowerShell / many)
 *   `# region [label]`           (Python / Ruby / Perl / R / bash / yaml /
 *                                 toml — `#` line comment + space, VSCode's
 *                                 default Python region marker)
 *   `// #region [label]`         (C / C++ / Java / Go / Rust / Swift / JS /
 *                                 TS)
 *   `// region [label]`          (IntelliJ-style, no hash)
 *   `/* #region [label] *\/`     (C-block-comment)
 *   `-- #region [label]`         (SQL / Lua / Haskell)
 *   `<!-- #region [label] -->`   (HTML / XML)
 * Leading whitespace is ignored and the match is case-insensitive (VSCode
 * accepts `#region`/`#Region`). The optional label after the keyword is allowed
 * but not required.
 */
export function isRegionStart(line: string): boolean {
  return REGION_START_RE.test(stripCr(line));
}

/** Whether a line is the matching `#endregion` / `#end region` END marker. */
export function isRegionEnd(line: string): boolean {
  return REGION_END_RE.test(stripCr(line));
}

// A comment-lead is an optional `//`, `--`, `;`, `#`, block-comment opener, or
// HTML comment opener, all optional, then optional whitespace. We then require
// the `#region` / `region` keyword. Kept as two anchored regexes (start vs end)
// so the scan is a single cheap test per line.
//
// Group breakdown of the lead `^\s*(?:\/\/|--|;|#|\/\*|<!--)?\s*`:
//   - `\/\/`  line comment (C-family / JS / TS),
//   - `--`    line comment (SQL / Lua),
//   - `;`     line comment (assembly / ini),
//   - `#`     line comment (Python / Ruby / Perl / R / bash / yaml / toml) —
//             so `# region` (hash + space + keyword, the idiomatic Python marker)
//             matches as comment-lead `#` + whitespace + `region`,
//   - `\/\*`  block comment open (C-family),
//   - `<!--`  HTML/XML comment open,
//   - all optional, so the bare `#region` also matches.
const REGION_LEAD = String.raw`^\s*(?:\/\/|--|;|#|\/\*|<!--)?\s*`;
// `#?region` accepts `#region` (bare, no space — the `#` here, NOT consumed by
// the lead, attaches directly to `region`) and the hash-less `region`
// (IntelliJ). A word boundary after stops `regional` from matching. Anything
// after (a label, the closing `*/` / `-->`) is ignored. Note `# region` (hash +
// space) is handled by the lead `#` above, not this optional hash.
const REGION_START_RE = new RegExp(`${REGION_LEAD}#?region\\b`, "i");
// End marker: `#endregion`, `#end region`, `# endregion`, or `endregion`
// (hash-less), same optional comment lead.
const REGION_END_RE = new RegExp(`${REGION_LEAD}#?end\\s?region\\b`, "i");

/**
 * EXPLICIT-marker fold range for the line starting at offset `lineStart`.
 *
 * When that line is a `#region` start marker, returns the range from the END of
 * the start-marker line to the END of the line holding its MATCHING
 * `#endregion` (nesting-aware: an inner region's start/end pair is balanced so
 * the outer region pairs with the correct closer). The end-marker line is
 * INCLUDED in the hidden range (collapsing a region hides through its
 * `#endregion`, like VSCode). Returns `null` when the line is not a region
 * start, or when it has no matching `#endregion` (an unbalanced marker is not
 * foldable rather than folding to end-of-file).
 *
 * @param text      The whole document.
 * @param lineStart Document offset of the first char of the candidate line.
 * @returns The `[from, to)` range to hide, or `null`.
 */
export function computeRegionFoldRange(
  text: string,
  lineStart: number,
): FoldOffsetRange | null {
  if (!text) return null;
  const { lines, starts } = splitLinesWithOffsets(text);
  const h = lineIndexAtStart(starts, lineStart);
  if (h < 0) return null;
  if (!isRegionStart(lines[h])) return null;

  // Scan forward for the matching #endregion, tracking nesting depth so an
  // inner region's own start/end pair doesn't close the outer one.
  let depth = 1;
  for (let k = h + 1; k < lines.length; k += 1) {
    const line = lines[k];
    // Check END before START: a single line is realistically only one or the
    // other, but checking end first keeps a same-line pathological case from
    // double-counting.
    if (isRegionEnd(line)) {
      depth -= 1;
      if (depth === 0) {
        const from = starts[h] + lines[h].length;
        const to = starts[k] + lines[k].length;
        if (to <= from) return null;
        return { from, to };
      }
    } else if (isRegionStart(line)) {
      depth += 1;
    }
  }
  return null; // unbalanced — no matching #endregion
}
