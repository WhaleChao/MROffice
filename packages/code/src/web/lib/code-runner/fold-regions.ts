/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for code folding (VSCode parity), detected by INDENTATION.
 *
 * `computeFoldRegions(text, tabSize)` returns the foldable regions of a source
 * document: each region is anchored on a "header" line whose following lines are
 * more deeply indented, and runs to the last line of that more-indented body
 * (trailing blank lines are NOT swallowed into the region — they belong to the
 * surrounding scope, matching how VSCode folds). Regions are depth-aware and may
 * nest; the result is ordered by `startLine` ascending.
 *
 * Folding is a VIEW concern: CodeEditor uses these regions to draw fold chevrons
 * in the gutter and to grey/collapse the body in the overlay — it never mutates
 * the underlying textarea text. So this module is framework-free and only deals
 * in 0-based line indices.
 *
 * All helpers are total (never throw) and defensive about empty input, blank /
 * whitespace-only lines, and CRLF line endings (the trailing `\r` is ignored).
 */

/**
 * A foldable region. `startLine` is the 0-based header line that gets the fold
 * chevron; `endLine` is the 0-based last line included in the region. Collapsing
 * the region hides lines `startLine + 1 .. endLine` (inclusive). A region always
 * spans at least two lines (`endLine > startLine`).
 */
export interface FoldRegion {
  startLine: number;
  endLine: number;
}

/**
 * Indent width (in columns) of a line's leading whitespace. Tabs count as one
 * `step`-wide stop; spaces count one column each. Returns `null` for a blank /
 * whitespace-only line (it carries no indentation level of its own and is
 * skipped when deciding region boundaries).
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
  // Whitespace-only (or empty) line: no own indentation level.
  if (i >= line.length) return null;
  return width;
}

/**
 * Compute the foldable regions of `text` by indentation.
 *
 * Algorithm: split into lines (CRLF-safe), compute each non-blank line's indent
 * width. A line `h` opens a region when the next NON-BLANK line is strictly more
 * indented than `h`. The region body is every subsequent line up to (but not
 * including) the first non-blank line whose indent is `<=` the header's indent;
 * trailing blank lines just before that boundary are trimmed off the region so a
 * blank separator between sibling blocks is not folded away.
 *
 * Nested deeper blocks naturally produce their own (inner) regions, so the
 * output can contain overlapping ranges where one fully contains another.
 *
 * @param text    Source document.
 * @param tabSize Columns per tab stop (>= 1; floored). Defaults to 2 to match
 *                the editor's `tab-size`.
 * @returns Regions ordered by `startLine` ascending (ties broken by `endLine`).
 */
export function computeFoldRegions(text: string, tabSize = 2): FoldRegion[] {
  if (!text) return [];
  const step = Math.max(1, Math.floor(tabSize) || 1);
  const rawLines = text.split("\n");
  const lines = rawLines.map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  const n = lines.length;

  // Per-line indent (null for blank / whitespace-only lines).
  const indents = lines.map((l) => indentWidth(l, step));

  const regions: FoldRegion[] = [];

  for (let h = 0; h < n; h += 1) {
    const headerIndent = indents[h];
    // Blank lines never open a region.
    if (headerIndent === null) continue;

    // Find the next non-blank line after the header.
    let j = h + 1;
    while (j < n && indents[j] === null) j += 1;
    // No following content, or it's not more indented -> not a header.
    if (j >= n) continue;
    const nextIndent = indents[j];
    if (nextIndent === null || nextIndent <= headerIndent) continue;

    // Walk the body: every line until indentation returns to <= header level.
    // `end` tracks the last NON-BLANK body line so trailing blanks are excluded.
    let end = j;
    let k = j;
    while (k < n) {
      const ind = indents[k];
      if (ind === null) {
        k += 1;
        continue; // blank line: tentatively inside, but don't extend `end`
      }
      if (ind <= headerIndent) break; // dedent: body ended before this line
      end = k;
      k += 1;
    }

    if (end > h) regions.push({ startLine: h, endLine: end });
  }

  // Stable order: by start, then by end (outer regions before inner when they
  // happen to share a start line, which they cannot here, but keep it total).
  regions.sort((a, b) => a.startLine - b.startLine || a.endLine - b.endLine);
  return regions;
}

/**
 * The set of line indices hidden when the given regions are collapsed. A line is
 * hidden if it lies strictly inside any collapsed region (`start < line <= end`)
 * — the header line itself stays visible (it shows the `…` affordance). Total
 * and defensive: unknown / out-of-range region indices contribute nothing
 * harmful, and overlapping/nested collapsed regions union cleanly.
 *
 * @param regions        All foldable regions (typically from computeFoldRegions).
 * @param collapsedStarts Set of header line indices that are currently collapsed.
 * @returns A set of 0-based hidden line indices.
 */
export function hiddenLines(
  regions: FoldRegion[],
  collapsedStarts: ReadonlySet<number>,
): Set<number> {
  const hidden = new Set<number>();
  for (const r of regions) {
    if (!collapsedStarts.has(r.startLine)) continue;
    for (let line = r.startLine + 1; line <= r.endLine; line += 1) {
      hidden.add(line);
    }
  }
  return hidden;
}

/**
 * Whether `line` is a fold header (the start of some region). Convenience for
 * the gutter, which draws a chevron only on header lines.
 */
export function isFoldHeader(regions: FoldRegion[], line: number): boolean {
  for (const r of regions) if (r.startLine === line) return true;
  return false;
}
