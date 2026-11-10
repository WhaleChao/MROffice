/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure line-level editor commands (VSCode parity): toggle line comment, move
 * lines, duplicate lines. Each takes the current value + selection and returns
 * the next value + selection (or null to fall through). Framework-free + unit
 * tested; CodeEditor wires them to key chords.
 */

export interface CommandInput {
  value: string;
  selectionStart: number;
  selectionEnd: number;
  language?: string;
}

export interface CommandResult {
  value: string;
  selectionStart: number;
  selectionEnd: number;
}

const HASH = new Set(["python", "ruby", "r", "perl", "bash", "shell", "lua", "yaml", "toml"]);
const DASH = new Set(["sql"]);
const PERCENT = new Set(["latex", "tex"]);

/** The single-line comment prefix for a language (defaults to //). */
export function lineCommentPrefix(language?: string): string {
  const lang = (language || "").toLowerCase();
  if (HASH.has(lang)) return "#";
  if (DASH.has(lang)) return "--";
  if (PERCENT.has(lang)) return "%";
  return "//";
}

interface LineSpan {
  /** Index of the first char of the first selected line. */
  blockStart: number;
  /** Index just past the last char of the last selected line. */
  blockEnd: number;
  lines: string[];
}

function selectedLineSpan(value: string, selStart: number, selEnd: number): LineSpan {
  const blockStart = value.lastIndexOf("\n", selStart - 1) + 1;
  let blockEnd = value.indexOf("\n", selEnd);
  // When the selection ends exactly at a line start, don't include the next line.
  if (selEnd > selStart && value[selEnd - 1] === "\n") {
    blockEnd = selEnd - 1;
  }
  if (blockEnd === -1) blockEnd = value.length;
  return { blockStart, blockEnd, lines: value.slice(blockStart, blockEnd).split("\n") };
}

/** Toggle `//`/`#`/`--` line comments over the selected lines. */
export function toggleLineComment(input: CommandInput): CommandResult {
  const { value, selectionStart, selectionEnd, language } = input;
  const prefix = lineCommentPrefix(language);
  const { blockStart, blockEnd, lines } = selectedLineSpan(value, selectionStart, selectionEnd);
  const nonEmpty = lines.filter((l) => l.trim() !== "");
  const allCommented =
    nonEmpty.length > 0 &&
    nonEmpty.every((l) => l.trimStart().startsWith(prefix));

  const newLines = lines.map((line) => {
    if (line.trim() === "") return line;
    const indentLen = line.length - line.trimStart().length;
    const indent = line.slice(0, indentLen);
    const rest = line.slice(indentLen);
    if (allCommented) {
      // Remove the prefix and a single following space if present.
      let stripped = rest.slice(prefix.length);
      if (stripped.startsWith(" ")) stripped = stripped.slice(1);
      return indent + stripped;
    }
    return indent + prefix + " " + rest;
  });

  const newBlock = newLines.join("\n");
  const head = value.slice(0, blockStart);
  const tail = value.slice(blockEnd);
  return {
    value: head + newBlock + tail,
    selectionStart: blockStart,
    selectionEnd: blockStart + newBlock.length,
  };
}

/** Move the selected line(s) up or down by one line. */
export function moveLines(input: CommandInput, direction: -1 | 1): CommandResult | null {
  const { value, selectionStart, selectionEnd } = input;
  const lines = value.split("\n");
  const startLine = value.slice(0, selectionStart).split("\n").length - 1;
  const endChar = selectionEnd > selectionStart && value[selectionEnd - 1] === "\n"
    ? selectionEnd - 1
    : selectionEnd;
  const endLine = value.slice(0, endChar).split("\n").length - 1;

  if (direction === -1 && startLine === 0) return null;
  if (direction === 1 && endLine === lines.length - 1) return null;

  const block = lines.slice(startLine, endLine + 1);
  if (direction === -1) {
    const above = lines[startLine - 1];
    lines.splice(startLine - 1, block.length + 1, ...block, above);
  } else {
    const below = lines[endLine + 1];
    lines.splice(startLine, block.length + 1, below, ...block);
  }
  const newValue = lines.join("\n");
  // Recompute the selection to follow the moved block.
  const newStartLine = startLine + direction;
  const newStart =
    lines.slice(0, newStartLine).reduce((n, l) => n + l.length + 1, 0);
  const newEnd = newStart + block.join("\n").length;
  return { value: newValue, selectionStart: newStart, selectionEnd: newEnd };
}

/** Duplicate the selected line(s) up or down. */
export function duplicateLines(input: CommandInput, direction: -1 | 1): CommandResult {
  const { value, selectionStart, selectionEnd } = input;
  const { blockStart, blockEnd, lines } = selectedLineSpan(value, selectionStart, selectionEnd);
  const block = lines.join("\n");
  const head = value.slice(0, blockStart);
  const tail = value.slice(blockEnd);
  if (direction === 1) {
    // Insert a copy below; move the selection onto the copy.
    const inserted = block + "\n" + block;
    const copyStart = blockStart + block.length + 1;
    return {
      value: head + inserted + tail,
      selectionStart: copyStart,
      selectionEnd: copyStart + block.length,
    };
  }
  // Insert a copy above; keep the selection on the original (now shifted).
  const inserted = block + "\n" + block;
  return {
    value: head + inserted + tail,
    selectionStart: blockStart,
    selectionEnd: blockStart + block.length,
  };
}

// ---------------------------------------------------------------------------
// Pure text transforms (VSCode parity): case conversion + line sorting.
//
// These are framework-free string→string functions so they can be unit tested
// in isolation and reused per-selection by CodeMirrorEditor (which folds them
// over each selection range / the selected line block via `changeByRange`).
// ---------------------------------------------------------------------------

/** Unicode identifier continuations used for caret-word editing commands. */
const WORD_CHAR = /[\p{ID_Continue}$\u200c\u200d]/u;
const WORD_RUN = /[\p{ID_Continue}$\u200c\u200d]+/gu;

/**
 * The half-open `[from, to)` range of the "word" surrounding `pos` in `value`
 * (Unicode identifier continuations plus `$`). Returns a zero-width range at `pos` when the
 * caret is not on/adjacent to a word char, so callers can no-op gracefully. The
 * offset is clamped into the string, so an out-of-bounds `pos` is safe.
 */
export function wordRangeAt(value: string, pos: number): { from: number; to: number } {
  const clamped = Math.max(0, Math.min(pos, value.length));
  WORD_RUN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = WORD_RUN.exec(value))) {
    const to = match.index + match[0].length;
    if (clamped >= match.index && clamped <= to) {
      return { from: match.index, to };
    }
    if (match.index > clamped) break;
  }
  return { from: clamped, to: clamped };
}

/** The case transforms exposed as editor commands. */
export type CaseTransform = "upper" | "lower" | "title";

/**
 * Title-case `text` the way VS Code's "Transform to Title Case" does: lower-case
 * it, then upper-case each letter that starts a word — at a line start, after
 * anything that is not a letter, digit or apostrophe, or after an apostrophe
 * that does not follow a letter. So "it's" stays "It's" (not "It'S"),
 * "'quoted'" reads "'Quoted'", "foo_bar" reads "Foo_Bar" and "3d" stays "3d".
 * The typographic apostrophe counts as one too ("it’s" → "It’s"). Pure + total.
 */
const TITLE_CASE_BOUNDARY = /(^|[^\p{L}\p{N}'’]|((^|\P{L})['’]))\p{L}/gmu;

function toTitleCase(text: string): string {
  return text
    .toLowerCase()
    .replace(TITLE_CASE_BOUNDARY, (boundary) => boundary.toUpperCase());
}

/**
 * Apply a {@link CaseTransform} to `text`. Pure + total; an empty string maps to
 * an empty string. `upper`/`lower` are locale-agnostic `toUpper/LowerCase`;
 * `title` upper-cases each word's first char (see {@link toTitleCase}).
 */
export function transformCase(text: string, transform: CaseTransform): string {
  switch (transform) {
    case "upper":
      return text.toUpperCase();
    case "lower":
      return text.toLowerCase();
    case "title":
      return toTitleCase(text);
    default:
      return text;
  }
}

/**
 * Sort the lines of `text` (split on "\n") in ascending or descending order and
 * re-join with "\n". STABLE: equal lines keep their original relative order
 * (the comparator falls back to the original index on a tie). By default the
 * comparison is case-insensitive locale order (VSCode "Sort Lines" default);
 * pass `caseSensitive` to use a code-point comparison instead.
 *
 * Pure + total: a single-line buffer (no "\n") is returned unchanged, and the
 * line count / trailing newline structure is preserved (sorting reorders the
 * lines but never adds or drops one).
 */
export function sortLines(
  text: string,
  direction: 1 | -1,
  caseSensitive = false,
): string {
  const lines = text.split("\n");
  // Decorate with the original index so equal keys keep input order (stable).
  const indexed = lines.map((line, index) => ({ line, index }));
  indexed.sort((a, b) => {
    let cmp: number;
    if (caseSensitive) {
      cmp = a.line < b.line ? -1 : a.line > b.line ? 1 : 0;
    } else {
      // Case-insensitive locale order. Lines that differ only in case compare
      // equal (sensitivity: "accent" ignores case), so they fall through to the
      // stability fallback below and keep their original input order — a
      // case-fold sort that is also stable.
      cmp = a.line.localeCompare(b.line, undefined, { sensitivity: "accent" });
    }
    if (cmp === 0) return a.index - b.index; // stable
    return direction === 1 ? cmp : -cmp;
  });
  return indexed.map((entry) => entry.line).join("\n");
}

/**
 * Strip trailing whitespace (spaces / tabs) from the END of every line in
 * `text` (VSCode "Files: Trim Trailing Whitespace" / `files.trimTrailingWhitespace`).
 * Lines are split on "\n"; each line's trailing run of `[ \t\f\v]` is removed and
 * the lines are re-joined with "\n". The line COUNT and the document's line
 * endings are preserved — only end-of-line whitespace is dropped, so a buffer
 * with no trailing whitespace (and no editor that ever runs this) is returned
 * byte-for-byte unchanged.
 *
 * Notes:
 *  - A trailing blank line (final "\n" then "") becomes an empty last line, which
 *    re-joins identically, so a file's final newline is untouched.
 *  - Internal blank lines collapse their whitespace to "" but the line itself
 *    remains (an empty string between two "\n"), so the structure is preserved.
 *  - Pure + total: empty input maps to empty output and it never throws.
 *
 * Carriage returns are NOT stripped: a "\r" before "\n" (CRLF, when the buffer
 * is loaded with literal "\r\n") is content, not trailing whitespace, so it is
 * preserved. (The editor normalizes EOL elsewhere; this transform is EOL-agnostic.)
 */
export function trimTrailingWhitespace(text: string): string {
  // Strip a maximal trailing run of space / tab / form-feed / vertical-tab from
  // the end of each line. Splitting on "\n" leaves a "\r" (from a literal CRLF
  // buffer) as the LAST char of the segment; the optional `\r?$` keeps that
  // carriage return while still removing whitespace that precedes it, so CRLF
  // line endings survive and only end-of-line whitespace is dropped.
  return text
    .split("\n")
    .map((line) => line.replace(/[ \t\f\v]+(\r?)$/, "$1"))
    .join("\n");
}

/** A document line-ending style. */
export type EolStyle = "LF" | "CRLF";

/**
 * Normalize EVERY line ending in `text` to `eol` ("LF" → "\n", "CRLF" → "\r\n")
 * (VSCode "Change End of Line Sequence"). All existing `\r\n` / lone `\r` / `\n`
 * are first collapsed to a single `\n`, then re-expanded to the target, so the
 * result is uniform regardless of the input's mix (and converting to the style a
 * file already uses is a byte-for-byte no-op). A lone `\r` (old-Mac line break)
 * is treated as a line ending too, so it is upgraded rather than left dangling.
 *
 * Pure + total: empty input maps to empty output and it never throws. The line
 * COUNT is preserved (only the separators change); a trailing newline stays a
 * trailing newline in the new style.
 */
export function convertEol(text: string, eol: EolStyle): string {
  // Collapse any CRLF or lone CR to LF first, so a mixed buffer normalizes
  // cleanly, then expand to the requested ending.
  const lf = text.replace(/\r\n?/g, "\n");
  return eol === "CRLF" ? lf.replace(/\n/g, "\r\n") : lf;
}

/** An indentation style: hard tabs, or N spaces. */
export interface IndentStyle {
  /** True for hard tabs, false for spaces. */
  useTabs: boolean;
  /** Spaces-per-indent (spaces mode) / the tab WIDTH used to convert (tabs mode). */
  width: number;
}

/**
 * Re-indent `text` by converting each line's LEADING whitespace run from one
 * indentation style to another (VSCode "Convert Indentation to Tabs/Spaces").
 * Only the leading run of spaces/tabs on every line is touched; the rest of the
 * line (including interior whitespace and any trailing `\r` from a CRLF buffer)
 * is preserved verbatim, so converting code never disturbs alignment past the
 * indent or the document's line endings.
 *
 * The leading run is measured as a VISUAL COLUMN using `from.width` as the tab
 * stop: a tab advances to the next multiple of `from.width`, a space advances by
 * one. That column is then re-emitted in the target style — `to.useTabs` emits
 * `floor(col / to.width)` tabs plus `col % to.width` spaces (so a partial indent
 * that doesn't fill a whole tab keeps its remainder as spaces, matching VSCode);
 * `!to.useTabs` emits `col` spaces.
 *
 * Pure + total: a clamp keeps widths ≥ 1, empty input maps to empty output, a
 * line with no leading whitespace is returned unchanged, and converting to the
 * style a buffer already uses is a no-op.
 */
export function convertIndentation(
  text: string,
  from: IndentStyle,
  to: IndentStyle,
): string {
  const fromWidth = Math.max(1, Math.floor(from.width));
  const toWidth = Math.max(1, Math.floor(to.width));
  return text
    .split("\n")
    .map((line) => {
      // Split the leading whitespace run from the rest of the line (a lone "\r"
      // at the end of a CRLF segment is NOT leading whitespace, so it stays in
      // `rest` untouched).
      let i = 0;
      let col = 0;
      while (i < line.length) {
        const ch = line.charCodeAt(i);
        if (ch === 32 /* space */) {
          col += 1;
        } else if (ch === 9 /* tab */) {
          // Advance to the next tab stop measured at the SOURCE tab width.
          col += fromWidth - (col % fromWidth);
        } else {
          break;
        }
        i += 1;
      }
      if (i === 0) return line; // no indentation → unchanged
      const rest = line.slice(i);
      let indent: string;
      if (to.useTabs) {
        const tabs = Math.floor(col / toWidth);
        const spaces = col % toWidth;
        indent = "\t".repeat(tabs) + " ".repeat(spaces);
      } else {
        indent = " ".repeat(col);
      }
      return indent + rest;
    })
    .join("\n");
}

/**
 * Join the lines spanned by the selection (or the current line with the one
 * BELOW it when the selection is empty) into a single line (VSCode "Join Lines",
 * Ctrl+J). Each newline inside the span is replaced by a single space, and the
 * leading whitespace of every joined-in line is collapsed so the merge reads as
 * one clean line — i.e. the boundary between two lines becomes exactly one space
 * (no space is inserted when one side is already empty, and no trailing space is
 * left when the next line was blank).
 *
 * Returns the new buffer plus the selection placed on the joined region (anchor
 * at the original block start, head just past the joined text) so the caller can
 * re-select it. Returns `null` when there is nothing to join — the caret is on
 * the LAST line with an empty selection — so the caller can no-op.
 *
 * Pure offset math + total: offsets are clamped into the buffer, and the line
 * structure outside the joined span (and the document's other line endings) is
 * preserved verbatim. CRLF (`\r\n`) joins identically: the `\r?\n` boundary is
 * consumed as the separator, so no stray carriage return survives in the seam.
 */
export function joinLines(
  value: string,
  selectionStart: number,
  selectionEnd: number,
): CommandResult | null {
  const start = Math.max(0, Math.min(selectionStart, value.length));
  const end = Math.max(start, Math.min(selectionEnd, value.length));
  // Widen to whole lines: from the start of the first selected line to the end
  // of the last selected line. (`lastIndexOf(x, -1)` searches from index 0 in
  // JS, so guard `start === 0` explicitly — the block then begins at offset 0.)
  const blockStart = start === 0 ? 0 : value.lastIndexOf("\n", start - 1) + 1;
  let lastLineSearchFrom = end;
  // When the selection ends exactly at a line start, the "last line" is the one
  // that ended at that boundary, not the next (empty) line.
  if (end > start && value[end - 1] === "\n") {
    lastLineSearchFrom = end - 1;
  }
  let blockEnd = value.indexOf("\n", lastLineSearchFrom);
  if (blockEnd === -1) {
    // Selection is on the last line. With an empty selection there is no line
    // below to join — nothing to do.
    if (start === end) return null;
    blockEnd = value.length;
  } else if (start === end) {
    // Empty selection: join the current line with the one below it, so extend
    // the block past the next newline to include that following line.
    const next = value.indexOf("\n", blockEnd + 1);
    blockEnd = next === -1 ? value.length : next;
  }

  const block = value.slice(blockStart, blockEnd);
  // Split on CRLF or LF; trim each piece so the seam collapses to one space and
  // re-join the non-empty pieces with a single space. Empty pieces (blank lines)
  // contribute nothing, so blank lines collapse cleanly without doubling spaces.
  const pieces = block.split(/\r?\n/);
  if (pieces.length < 2) return null; // single line → nothing to join
  const first = pieces[0];
  const restTrimmed = pieces
    .slice(1)
    .map((p) => p.replace(/^[ \t]+/, ""));
  let joined = first;
  for (const piece of restTrimmed) {
    if (piece.length === 0) continue; // blank line: no separator, no content
    // One space between the running text and the next piece, unless the running
    // text is empty (avoid a leading space).
    joined += (joined.length > 0 ? " " : "") + piece;
  }

  return {
    value: value.slice(0, blockStart) + joined + value.slice(blockEnd),
    selectionStart: blockStart,
    selectionEnd: blockStart + joined.length,
  };
}

/**
 * Splice a formatted replacement for the half-open range `[from, to)` back into
 * `value`, returning the new buffer plus the range the replacement now occupies
 * (so the caller can re-select exactly the reformatted text). Returns `null`
 * when the replacement is byte-identical to the original slice — i.e. nothing
 * changed, so the caller should no-op (no transaction, no selection churn).
 *
 * Pure offset math, separated from the async formatter so it can be unit tested:
 * the new selection end is `from + replacement.length`, and everything before
 * `from` / after `to` is preserved verbatim.
 */
export function spliceFormattedRange(
  value: string,
  from: number,
  to: number,
  replacement: string,
): { value: string; selectionStart: number; selectionEnd: number } | null {
  const lo = Math.max(0, Math.min(from, value.length));
  const hi = Math.max(lo, Math.min(to, value.length));
  const original = value.slice(lo, hi);
  if (replacement === original) return null;
  return {
    value: value.slice(0, lo) + replacement + value.slice(hi),
    selectionStart: lo,
    selectionEnd: lo + replacement.length,
  };
}

export interface CommandKeyInput extends CommandInput {
  key: string;
  metaKey: boolean;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
}

/** Map a key chord to a line command, or null if none applies. */
export function handleEditorCommandKey(input: CommandKeyInput): CommandResult | null {
  const mod = input.metaKey || input.ctrlKey;
  if (mod && !input.altKey && input.key === "/") {
    return toggleLineComment(input);
  }
  if (input.altKey && !mod) {
    if (input.key === "ArrowUp") {
      return input.shiftKey ? duplicateLines(input, -1) : moveLines(input, -1);
    }
    if (input.key === "ArrowDown") {
      return input.shiftKey ? duplicateLines(input, 1) : moveLines(input, 1);
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Fold-by-level (VSCode "Fold Level N", Ctrl+K Ctrl+1..7).
//
// CM6 surfaces the foldable RANGE for each line (`foldable`), but not its
// nesting DEPTH. VSCode's "Fold Level N" folds every region whose nesting depth
// is N. We compute that depth purely from containment of the `[from, to)` ranges
// so the editor command stays a thin shell over tested logic.
// ---------------------------------------------------------------------------

/** A half-open foldable document range, as returned by CM6's `foldable`. */
export interface FoldableRange {
  from: number;
  to: number;
}

/** A foldable range tagged with its 1-based nesting level. */
export interface LeveledFoldRange extends FoldableRange {
  /** Outermost regions are level 1; a region nested inside it is level 2; … */
  level: number;
}

/**
 * Assign each foldable range its 1-based nesting LEVEL by strict containment: a
 * range R is one level deeper than the SMALLEST range that strictly contains R's
 * start. Outermost ranges (contained by nothing) are level 1. Pure + total:
 *  - duplicate / identical ranges are de-duplicated (a `from→to` key) so the same
 *    foldable region reported by several of its lines counts once,
 *  - degenerate / inverted ranges (`to <= from`) are dropped,
 *  - the result is ordered by `from` ascending (ties: larger range — the outer —
 *    first) so callers iterate top-to-bottom.
 *
 * "Strictly contains" uses the start position only (`outer.from < r.from`), which
 * is how editor fold regions nest — a child block always opens after its parent's
 * header line — and avoids miscounting two sibling regions that merely abut.
 */
export function assignFoldLevels(ranges: readonly FoldableRange[]): LeveledFoldRange[] {
  // De-dupe by the exact span and drop degenerate ranges.
  const byKey = new Map<string, FoldableRange>();
  for (const r of ranges) {
    if (!Number.isFinite(r.from) || !Number.isFinite(r.to)) continue;
    if (r.to <= r.from) continue;
    byKey.set(`${r.from}:${r.to}`, { from: r.from, to: r.to });
  }
  const unique = Array.from(byKey.values());
  const leveled = unique.map((r) => {
    let level = 1;
    for (const other of unique) {
      if (other === r) continue;
      // `other` strictly contains `r`'s start and is a wider span.
      if (other.from < r.from && other.to >= r.to) level += 1;
    }
    return { from: r.from, to: r.to, level };
  });
  leveled.sort((a, b) => a.from - b.from || b.to - a.to);
  return leveled;
}

/**
 * The foldable ranges at exactly 1-based `level` (see {@link assignFoldLevels}).
 * `level < 1` yields `[]`. Used by the editor to fold every region at a chosen
 * indentation/nesting level (Ctrl+K Ctrl+N).
 */
export function foldRangesAtLevel(
  ranges: readonly FoldableRange[],
  level: number,
): LeveledFoldRange[] {
  if (level < 1) return [];
  return assignFoldLevels(ranges).filter((r) => r.level === level);
}

// ---------------------------------------------------------------------------
// Change All Occurrences (VSCode Ctrl+F2): select every occurrence of the word
// under the caret (no prior selection needed) as a multi-cursor selection.
// ---------------------------------------------------------------------------

/**
 * Every WHOLE-WORD, case-sensitive occurrence of `word` in `value`, returned as
 * half-open `[from, to)` ranges in document order. A match is "whole word" when
 * neither the char before `from` nor the char at `to` is a word char
 * (`[A-Za-z0-9_$]`), matching the identifier class used elsewhere — so searching
 * `foo` selects `foo` but not the `foo` inside `food`/`barfoo`. Pure + total: an
 * empty `word`, or a `word` containing a non-word char, yields `[]` (there is no
 * single "word" to change-all on), and it never throws.
 */
export function wordOccurrences(value: string, word: string): { from: number; to: number }[] {
  if (!word) return [];
  // The target must itself be a single word token; bail otherwise so we never
  // do a substring sweep for arbitrary text.
  for (let i = 0; i < word.length; i += 1) {
    if (!WORD_CHAR.test(word[i])) return [];
  }
  const out: { from: number; to: number }[] = [];
  let index = value.indexOf(word);
  while (index !== -1) {
    const before = index > 0 ? value[index - 1] : "";
    const afterPos = index + word.length;
    const after = afterPos < value.length ? value[afterPos] : "";
    const boundedLeft = before === "" || !WORD_CHAR.test(before);
    const boundedRight = after === "" || !WORD_CHAR.test(after);
    if (boundedLeft && boundedRight) out.push({ from: index, to: afterPos });
    index = value.indexOf(word, index + word.length);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Final newline on save (VSCode files.insertFinalNewline / trimFinalNewlines).
// ---------------------------------------------------------------------------

/** How the buffer's end-of-file newline run is treated on save. */
export type FinalNewlineMode = "insert" | "trim";

/**
 * Apply VSCode's final-newline save behaviour to `text`:
 *  - `"insert"` (files.insertFinalNewline): ensure the file ends with EXACTLY one
 *    trailing newline — append "\n" when missing, and collapse a run of trailing
 *    blank lines down to a single final newline.
 *  - `"trim"` (files.trimFinalNewlines): remove ALL trailing newlines so the file
 *    ends on its last non-empty line's content (no final newline at all).
 *
 * The EOL flavour is preserved: a CRLF buffer keeps "\r\n" (the inserted newline
 * matches the buffer's prevailing ending). Pure + total:
 *  - empty input maps to empty output for BOTH modes (nothing to anchor a newline
 *    to / nothing to trim) so a brand-new empty buffer is untouched,
 *  - a buffer already in the target shape is returned byte-for-byte unchanged,
 *  - it never throws.
 */
export function applyFinalNewline(text: string, mode: FinalNewlineMode): string {
  if (text === "") return "";
  // Detect the prevailing EOL so an inserted newline matches it: CRLF only when
  // the buffer uses "\r\n" and has no lone-LF lines (a mixed buffer defaults to
  // LF, the editor's normalized ending).
  const hasCrlf = text.includes("\r\n");
  const hasLoneLf = /(^|[^\r])\n/.test(text);
  const eol = hasCrlf && !hasLoneLf ? "\r\n" : "\n";
  // Strip every trailing newline (LF or CRLF) to find the content end.
  const trimmed = text.replace(/(\r?\n)+$/, "");
  if (mode === "trim") return trimmed;
  // insert: content + exactly one EOL. (If trimming removed everything — the
  // buffer was only blank lines — `trimmed` is "" and we still emit one EOL so
  // the file ends with a single newline, matching VSCode.)
  return trimmed + eol;
}

// ---------------------------------------------------------------------------
// Expand / Shrink selection (VSCode Shift+Alt+Right / Shift+Alt+Left).
//
// CM6 ships `selectParentSyntax` (expand) but has NO inverse. We drive both from
// a small SELECTION-HISTORY stack: expanding pushes the pre-expand range, so
// shrinking can pop back to it. These pure helpers own the stack arithmetic so
// the editor command is a thin shell. Ranges are `[anchor, head)` numeric pairs.
// ---------------------------------------------------------------------------

/** A selection range remembered on the expand/shrink stack. */
export interface SelectionRange {
  anchor: number;
  head: number;
}

/** Two ranges cover the same span (order-independent on anchor/head). */
function sameRange(a: SelectionRange, b: SelectionRange): boolean {
  return (
    (a.anchor === b.anchor && a.head === b.head) ||
    (a.anchor === b.head && a.head === b.anchor)
  );
}

/**
 * Push `current` onto the expand-history `stack` ahead of an expand step.
 * Returns the new stack. De-duplicates the top so repeated expands from an
 * unchanged selection don't stack duplicates, and guards against unbounded
 * growth with a fixed `limit` (oldest entries are dropped). Pure (no mutation of
 * the input array).
 */
export function pushExpandHistory(
  stack: readonly SelectionRange[],
  current: SelectionRange,
  limit = 50,
): SelectionRange[] {
  const top = stack[stack.length - 1];
  if (top && sameRange(top, current)) return stack.slice();
  const next = [...stack, { anchor: current.anchor, head: current.head }];
  if (next.length > limit) next.splice(0, next.length - limit);
  return next;
}

/**
 * Pop the expand-history `stack` to SHRINK the selection. The top of the stack is
 * the range we should shrink back to ONLY when it is strictly smaller than the
 * `current` selection (so a stale entry — e.g. after the user clicked elsewhere —
 * is discarded rather than jumping the caret). Returns the range to restore plus
 * the remaining stack, or `null` when there is nothing smaller to shrink to.
 * Pure (no mutation of the input array).
 */
export function popShrinkHistory(
  stack: readonly SelectionRange[],
  current: SelectionRange,
): { range: SelectionRange; stack: SelectionRange[] } | null {
  const next = stack.slice();
  const curLen = Math.abs(current.head - current.anchor);
  while (next.length > 0) {
    const candidate = next.pop() as SelectionRange;
    // Skip an entry identical to the current selection (nothing would change).
    if (sameRange(candidate, current)) continue;
    const candLen = Math.abs(candidate.head - candidate.anchor);
    // Only shrink: the remembered range must be smaller than the current one.
    if (candLen < curLen) {
      return { range: candidate, stack: next };
    }
  }
  return null;
}
