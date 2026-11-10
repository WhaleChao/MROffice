/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure status-bar derivations for the CodeMirror 6 code editor (VSCode parity).
 *
 * The editor's bottom status bar shows three things, all derived here from plain
 * strings + the caret offset (no CM6 / React state), so the logic is unit-tested
 * directly and the component only renders the result:
 *
 *  - {@link lineColAt} — the caret's 1-based line + column from a 0-based document
 *    offset (the "Ln X, Col Y" indicator). Column is counted in UTF-16 code units
 *    from the line start, +1, exactly like VSCode reports it.
 *  - {@link detectIndent} — whether the document indents with tabs or spaces and
 *    the indent WIDTH, derived from the first indented line (the "Spaces: 2" /
 *    "Tab Size: 2" indicator). Falls back to the editor's configured width when
 *    nothing in the buffer is indented yet.
 *  - {@link detectEol} — the document's line ending, "CRLF" when any `\r\n` is
 *    present else "LF" (the EOL indicator).
 *
 * Framework-free + TOTAL (never throws): a caret past the end clamps into the
 * document, an empty buffer reports Ln 1, Col 1, etc.
 */

/** A 1-based line/column position, as shown in the status bar. */
export interface LineCol {
  /** 1-based line number. */
  line: number;
  /** 1-based column (UTF-16 code units from the line start, +1). */
  col: number;
}

/**
 * The 1-based {line, col} of a 0-based document `offset` within `text`.
 *
 * Line = (count of `\n` before the offset) + 1; column = (offset - lineStart) + 1,
 * where lineStart is the index just after the preceding `\n`. A `\r` in a CRLF
 * pair counts as a normal character on its line (column includes it), matching
 * how an editor that stores CRLF reports the caret. The offset is clamped into
 * `[0, text.length]` so a stale/out-of-range caret can never produce NaN or a
 * negative column.
 */
export function lineColAt(text: string, offset: number): LineCol {
  const pos = Math.max(0, Math.min(offset, text.length));
  let line = 1;
  let lineStart = 0;
  // Walk newlines up to the caret; this is O(offset) but the caret listener only
  // re-derives on an actual selection change, so it never runs in a hot loop.
  for (let i = 0; i < pos; i += 1) {
    if (text.charCodeAt(i) === 10 /* \n */) {
      line += 1;
      lineStart = i + 1;
    }
  }
  return { line, col: pos - lineStart + 1 };
}

/** Indentation style of a document: tabs vs spaces, plus the indent width. */
export interface IndentInfo {
  /** True when the document indents with hard tabs, false for spaces. */
  useTabs: boolean;
  /**
   * Indent WIDTH: for spaces, the number of leading spaces on the first
   * space-indented line; for tabs, the configured tab display width. Always a
   * positive integer.
   */
  width: number;
}

/**
 * Detect a document's indentation from its content, mirroring VSCode's
 * "Spaces: N" / "Tab Size: N" status indicator.
 *
 * Scans lines top-to-bottom for the FIRST one that begins with a tab or a space:
 *  - leading tab  -> tabs, width = `configuredTabWidth`
 *  - leading space(s) -> spaces, width = the count of those leading spaces
 *    (clamped to a sane 1..8 so a stray deep-indented first line can't report
 *    "Spaces: 37").
 * When NO line is indented (empty buffer, or every line flush-left), falls back
 * to spaces at `configuredTabWidth` — the editor's own setting — so the indicator
 * always shows the active default rather than guessing.
 *
 * Pure + total: reads only the string and the configured width.
 */
export function detectIndent(text: string, configuredTabWidth = 2): IndentInfo {
  const tabWidth = configuredTabWidth > 0 ? configuredTabWidth : 2;
  const lines = text.split("\n");
  for (const raw of lines) {
    // Ignore a trailing \r (CRLF) — it is not indentation.
    const line = raw.endsWith("\r") ? raw.slice(0, -1) : raw;
    if (line.length === 0) continue;
    const first = line.charCodeAt(0);
    if (first === 9 /* \t */) {
      return { useTabs: true, width: tabWidth };
    }
    if (first === 32 /* space */) {
      let n = 0;
      while (n < line.length && line.charCodeAt(n) === 32) n += 1;
      // Clamp into a sane indent-width range so an oddly-indented first line
      // (e.g. a wrapped comment) doesn't report an absurd width.
      const width = Math.max(1, Math.min(8, n));
      return { useTabs: false, width };
    }
    // A non-whitespace-leading line is not indented; keep scanning.
  }
  return { useTabs: false, width: tabWidth };
}

/** A document's line-ending style, as shown in the status bar. */
export type Eol = "LF" | "CRLF";

/**
 * The document's end-of-line style: "CRLF" when ANY `\r\n` pair appears, else
 * "LF". (A buffer with no newline at all reports "LF" — the editor's default and
 * what a single-line buffer will use the moment a newline is typed.) Pure +
 * total.
 */
export function detectEol(text: string): Eol {
  return text.includes("\r\n") ? "CRLF" : "LF";
}

/**
 * Resolve the status-bar EOL without rescanning the source when the host already
 * owns authoritative file metadata. Standalone/local-file editors preserve EOL
 * separately because CodeMirror normalizes its document to LF; using that value
 * first keeps an O(source length) scan off every controlled edit. Embedded
 * editors that do not provide metadata retain the historical content-derived
 * fallback.
 */
export function resolveEol(text: string, controlled?: Eol): Eol {
  return controlled ?? detectEol(text);
}

/**
 * The human label for the indent indicator, e.g. "Spaces: 2" / "Tab Size: 4".
 * The two label strings are passed in (localized by the caller) with a `{width}`
 * placeholder already substituted, so this stays framework-free.
 */
export function indentLabel(
  indent: IndentInfo,
  spacesLabel: string,
  tabsLabel: string,
): string {
  return indent.useTabs ? tabsLabel : spacesLabel;
}

/** A selection range as a pair of document offsets (order-independent). */
export interface OffsetRange {
  from: number;
  to: number;
}

/**
 * The status-bar selection summary VSCode shows next to Ln/Col:
 *  - `selectedChars` — total characters across all non-empty ranges (the
 *    "(N selected)" suffix). 0 when every caret is collapsed.
 *  - `cursorCount` — number of carets/ranges (the "N selections" indicator,
 *    shown only when > 1).
 *
 * `selectedChars` counts each range's `|to - from|`, summed — exactly what
 * VSCode reports for a multi-range selection. Pure + total: ignores nothing,
 * clamps no offsets (the caller passes CM6's own ranges), and an empty list
 * reports zeros.
 */
export interface SelectionSummary {
  selectedChars: number;
  cursorCount: number;
}

/**
 * Derive the {@link SelectionSummary} from the editor's selection ranges. Each
 * range contributes `|to - from|` characters; the count is the number of
 * ranges. A single collapsed caret yields `{ selectedChars: 0, cursorCount: 1 }`.
 */
export function selectionSummary(ranges: readonly OffsetRange[]): SelectionSummary {
  let selectedChars = 0;
  for (const r of ranges) {
    selectedChars += Math.abs(r.to - r.from);
  }
  return { selectedChars, cursorCount: ranges.length };
}

/**
 * Document offsets to place a caret at the END of every line touched by the
 * given selection ranges — the "Add Cursors to Line Ends" command (VSCode
 * Shift+Alt+I). For each range, every line index from the line containing
 * `from` through the line containing `to` gets one caret at that line's end
 * offset (just before its trailing `\n`, or the document end for the last line).
 *
 * The result is the SORTED, DEDUPLICATED set of those end offsets, so two
 * ranges that overlap lines don't produce stacked carets and the caller can
 * hand them straight to `EditorSelection.create`. Pure + total: offsets are
 * clamped into the document, and an empty/whitespace doc yields a single end
 * offset.
 *
 * NOTE: when a range's `to` sits exactly at the START of a line (a downward
 * line selection whose tail wraps to col 1), that trailing line is NOT
 * included — matching VSCode, which doesn't drop a cursor on the empty line
 * below a full-line selection. The exception is a fully collapsed/zero-line
 * range, which still contributes its own line end.
 */
export function lineEndCursorPositions(
  text: string,
  ranges: readonly OffsetRange[],
): number[] {
  const len = text.length;
  // Precompute the end offset of each line: index i -> offset of the char
  // before line i's terminating "\n" (or `len` for the final line).
  const lineEnds: number[] = [];
  let cursor = 0;
  for (;;) {
    const nl = text.indexOf("\n", cursor);
    if (nl === -1) {
      lineEnds.push(len);
      break;
    }
    lineEnds.push(nl);
    cursor = nl + 1;
  }
  // Map an offset to its 0-based line index.
  const lineIndexAt = (offset: number): number => {
    const pos = Math.max(0, Math.min(offset, len));
    let line = 0;
    for (let i = 0; i < pos; i += 1) {
      if (text.charCodeAt(i) === 10) line += 1;
    }
    return line;
  };

  const ends = new Set<number>();
  for (const r of ranges) {
    const lo = Math.min(r.from, r.to);
    const hi = Math.max(r.from, r.to);
    const firstLine = lineIndexAt(lo);
    let lastLine = lineIndexAt(hi);
    // Drop the trailing line when the selection ends exactly at its start AND
    // spans more than one line (VSCode: a full-line downward selection doesn't
    // seed a cursor on the next, untouched line).
    if (lastLine > firstLine && hi === lineStartOffset(text, lastLine)) {
      lastLine -= 1;
    }
    for (let li = firstLine; li <= lastLine; li += 1) {
      ends.add(lineEnds[li] ?? len);
    }
  }
  return [...ends].sort((a, b) => a - b);
}

/** The document offset of the START of 0-based `lineIndex`. */
function lineStartOffset(text: string, lineIndex: number): number {
  if (lineIndex <= 0) return 0;
  let seen = 0;
  for (let i = 0; i < text.length; i += 1) {
    if (text.charCodeAt(i) === 10) {
      seen += 1;
      if (seen === lineIndex) return i + 1;
    }
  }
  return text.length;
}
