/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for STICKY SCROLL (VSCode's signature feature): as the user scrolls
 * down inside a function / class / block, the enclosing "header" lines (e.g.
 * `function foo() {`, `class Bar {`, `if (...) {`) pin to the TOP of the editor
 * viewport so the surrounding context stays visible.
 *
 * This module is framework-free + fully unit-tested. Given the document, the
 * line currently at the viewport top, and a depth cap, {@link stickyHeaderLines}
 * returns the ordered stack of 0-based header line indices that should stick.
 * `CodeMirrorEditor` renders the returned lines as a small absolute overlay
 * (each clickable to scroll to its header) behind a persisted, default-OFF
 * toggle — so the editor is byte-for-byte unchanged until the user opts in.
 *
 * The "header" detection reuses the editor's existing indentation-based
 * {@link computeFoldRegions}: a header is the `startLine` of a foldable region
 * whose BODY (the lines after the header, up to `endLine`) contains the viewport
 * top line. Stacking the enclosing regions outermost-first yields the same
 * nesting context VSCode shows, with no dependency on a live language tree (so it
 * works uniformly across every language the editor supports, including the
 * legacy stream-parser ones with no lezer tree).
 */

import { computeFoldRegions, type FoldRegion } from "./fold-regions";

/** Default maximum number of stacked sticky headers (VSCode caps at 5 too). */
export const DEFAULT_STICKY_MAX_DEPTH = 5;

/**
 * A line that should stick to the top of the viewport. `line` is the 0-based
 * header line index; `text` is that line's source (with any trailing `\r`
 * stripped) so the overlay can render it without re-splitting the document.
 */
export interface StickyHeader {
  line: number;
  text: string;
}

/**
 * The stack of enclosing header LINE INDICES (0-based) for the line currently at
 * the viewport top, outermost first. A region encloses `topLine` when its header
 * is strictly above the viewport top and its body still covers it
 * (`startLine < topLine <= endLine`): the header has scrolled off (or is about
 * to) while its block is still on screen, which is exactly when its context
 * should pin. The header line that is itself the viewport top is NOT returned —
 * it is already visible, so pinning it would just duplicate the first real line.
 *
 * Regions are returned outermost-first (ascending `startLine`) and capped to
 * `maxDepth` by keeping the INNERMOST `maxDepth` headers (the closest context is
 * the most useful), still rendered outermost-first.
 *
 * Total + defensive: empty/blank input, an out-of-range `topLine`, and CRLF line
 * endings are all handled without throwing.
 *
 * @param text     The full document source.
 * @param topLine  0-based index of the line at the top of the viewport.
 * @param tabSize  Columns per tab stop for the indentation heuristic (matches
 *                 the editor's tab size; defaults to 2).
 * @param maxDepth Maximum stacked headers (defaults to {@link DEFAULT_STICKY_MAX_DEPTH}).
 * @returns Ordered header line indices, outermost first.
 */
export function stickyHeaderLineNumbers(
  text: string,
  topLine: number,
  tabSize = 2,
  maxDepth = DEFAULT_STICKY_MAX_DEPTH,
): number[] {
  if (!text) return [];
  const cap = Math.max(0, Math.floor(maxDepth));
  if (cap === 0) return [];
  if (!Number.isFinite(topLine) || topLine < 0) return [];

  const regions = computeFoldRegions(text, tabSize);
  // Every region whose BODY covers the viewport top line: the header is above
  // the top (already scrolled away or flush with it) and the block extends down
  // through it. computeFoldRegions returns regions ascending by startLine, so
  // the enclosing set is naturally outermost-first.
  const enclosing: FoldRegion[] = [];
  for (const region of regions) {
    if (region.startLine < topLine && topLine <= region.endLine) {
      enclosing.push(region);
    }
  }
  if (enclosing.length === 0) return [];

  const lines = enclosing.map((r) => r.startLine);
  // Cap by keeping the INNERMOST `cap` headers (the deepest context), but render
  // them outermost-first — the slice from the tail preserves that order.
  return lines.length > cap ? lines.slice(lines.length - cap) : lines;
}

/**
 * Like {@link stickyHeaderLineNumbers} but resolves each header line index to its
 * source text (trailing `\r` stripped), ready for the overlay to render. Headers
 * stay outermost-first.
 */
export function stickyHeaderLines(
  text: string,
  topLine: number,
  tabSize = 2,
  maxDepth = DEFAULT_STICKY_MAX_DEPTH,
): StickyHeader[] {
  const numbers = stickyHeaderLineNumbers(text, topLine, tabSize, maxDepth);
  if (numbers.length === 0) return [];
  const rawLines = text.split("\n");
  return numbers.map((line) => {
    const raw = rawLines[line] ?? "";
    return { line, text: raw.endsWith("\r") ? raw.slice(0, -1) : raw };
  });
}
