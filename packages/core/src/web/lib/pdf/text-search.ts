/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * In-document text search for the PDF viewer find bar.
 *
 * The viewer extracts a plain-text string per page (via pdf.js
 * `getTextContent()` joined with spaces). These pure helpers take those page
 * strings plus a query and return an ordered, flat list of matches across the
 * whole document so the find bar can show "n of N", step next/prev with
 * wraparound, and tell the viewer which page to scroll to / which character
 * range to highlight.
 *
 * Everything here is pure (no DOM, no pdf.js) so it can be unit-tested without
 * jsdom or a real document — the viewer layers the rect-overlay highlighting on
 * top using these character offsets.
 */

/** A single occurrence of the query inside one page's text. */
export interface PdfTextMatch {
  /** Zero-based index into the array of page texts passed to findMatches. */
  pageIndex: number;
  /** Inclusive character offset of the match start within that page's text. */
  start: number;
  /** Exclusive character offset of the match end within that page's text. */
  end: number;
}

/** Search options. */
export interface PdfTextSearchOptions {
  /** When true, "Cat" does not match "cat". Defaults to false. */
  caseSensitive?: boolean;
  /**
   * When true, only WHOLE-WORD occurrences match: each hit must be bounded by a
   * non-word character (or the start/end of the text) on both sides, so "cat"
   * matches "a cat sat" but not "category" or "scatter". Defaults to false.
   * Composes with {@link caseSensitive}.
   */
  wholeWord?: boolean;
  /**
   * Stop after this many matches. Matching stops allocating as soon as the cap
   * is reached, so a one-letter query on a huge document costs at most `limit`
   * match objects. Absent, NaN or +Infinity means no cap; 0 or less yields
   * `[]`.
   * Use {@link findMatchWindow} to also learn whether more matches exist.
   */
  limit?: number;
}

/**
 * Default match window shared by the find bar and the document-wide results
 * panel. One cap for both keeps panel rows 1:1 with find-bar indexes.
 */
export const PDF_FIND_MAX_MATCHES = 5000;

/**
 * A bounded run of matches plus an explicit flag for the ones left out, so a
 * caller can say "5000+" instead of silently dropping hits.
 */
export interface PdfTextMatchWindow {
  /** The first matches in document order, at most the window size. */
  matches: PdfTextMatch[];
  /**
   * True when at least one more match exists past `matches`. The scan stops
   * there, so the full total is known (`matches.length`) only when false.
   */
  truncated: boolean;
}

function isHighSurrogate(unit: number): boolean {
  return unit >= 0xd800 && unit <= 0xdbff;
}

function isLowSurrogate(unit: number): boolean {
  return unit >= 0xdc00 && unit <= 0xdfff;
}

/**
 * The whole code point that ends just before `index` (both halves of a
 * surrogate pair), or undefined at the start of the text.
 */
function codePointBefore(text: string, index: number): number | undefined {
  if (index <= 0 || index > text.length) return undefined;
  if (
    index >= 2 &&
    isLowSurrogate(text.charCodeAt(index - 1)) &&
    isHighSurrogate(text.charCodeAt(index - 2))
  ) {
    return text.codePointAt(index - 2);
  }
  return text.charCodeAt(index - 1);
}

/** The whole code point that starts at `index`, or undefined at the end. */
function codePointAfter(text: string, index: number): number | undefined {
  if (index < 0 || index >= text.length) return undefined;
  return text.codePointAt(index);
}

/** Whether `index` falls BETWEEN the two halves of one surrogate pair. */
export function splitsSurrogatePair(text: string, index: number): boolean {
  return (
    index > 0 &&
    index < text.length &&
    isHighSurrogate(text.charCodeAt(index - 1)) &&
    isLowSurrogate(text.charCodeAt(index))
  );
}

/**
 * A "word character" for the whole-word boundary test: ASCII letters/digits/`_`
 * PLUS any Unicode letter or number (so accented Latin, CJK — including astral
 * ideographs such as U+20BB7 — count as word characters and a match flanked by
 * them is correctly rejected as non-whole). We classify one whole code point
 * rather than building a regex from the (untrusted) needle, so a query with
 * regex metacharacters can't change the matching semantics or cost.
 */
function isWordCodePoint(codePoint: number | undefined): boolean {
  if (codePoint === undefined) return false;
  // Combining marks (\p{M}) belong to the letter they sit on: a decomposed
  // "CAFE\u0301" is one word, so whole-word "CAFE" must not end inside it
  
  return /[\p{L}\p{M}\p{N}_]/u.test(String.fromCodePoint(codePoint));
}

/**
 * Whether the half-open match range [start, end) within `text` is bounded by a
 * non-word character (or a text edge) on BOTH sides — i.e. it is a whole word.
 * The whole code point just before `start` and the one at `end` must each be
 * a non-word character (or absent); a boundary that cuts a surrogate pair in
 * half lies inside a character and is never whole. Exported for direct unit
 * testing.
 */
export function isWholeWordMatch(
  text: string,
  start: number,
  end: number,
): boolean {
  if (splitsSurrogatePair(text, start) || splitsSurrogatePair(text, end)) {
    return false;
  }
  return (
    !isWordCodePoint(codePointBefore(text, start)) &&
    !isWordCodePoint(codePointAfter(text, end))
  );
}

/** Resolve {@link PdfTextSearchOptions.limit}: Infinity means "no cap". */
function resolveMatchLimit(limit: unknown): number {
  if (typeof limit !== "number" || Number.isNaN(limit)) {
    return Number.POSITIVE_INFINITY;
  }
  if (!Number.isFinite(limit)) {
    return limit > 0 ? Number.POSITIVE_INFINITY : 0;
  }
  return Math.max(0, Math.trunc(limit));
}

/**
 * The one matching loop behind {@link findMatches} and {@link countMatches}:
 * visit every non-overlapping occurrence in document order and stop as soon
 * as `visit` returns false. It allocates nothing per match itself, and pages
 * after the stop are never lower-cased or scanned.
 */
function forEachMatch(
  pageTexts: ReadonlyArray<string>,
  query: string,
  options: PdfTextSearchOptions,
  visit: (pageIndex: number, start: number, end: number) => boolean,
): void {
  const needle = typeof query === "string" ? query : "";
  if (needle.length === 0) return;

  const caseSensitive = options.caseSensitive === true;
  const wholeWord = options.wholeWord === true;
  const compareNeedle = caseSensitive ? needle : foldCase(needle).folded;
  const needleLength = compareNeedle.length;

  for (let pageIndex = 0; pageIndex < pageTexts.length; pageIndex += 1) {
    const raw = pageTexts[pageIndex];
    const text = typeof raw === "string" ? raw : "";
    if (text.length === 0) continue;
    // Lower-casing can change the string's length (U+0130 becomes "i" plus a
    // combining dot, one code unit becoming two), so an offset found in the
    // folded text is mapped back to the original before anyone slices it.
    // Using the folded offset directly replaced "TARGET" in "İ TARGET END"
    
    const folded = caseSensitive ? null : foldCase(text);
    const haystack = folded ? folded.folded : text;

    let from = 0;
    let idx = haystack.indexOf(compareNeedle, from);
    while (idx !== -1) {
      const foldedEnd = idx + needleLength;
      const start = folded ? folded.originalOffset(idx) : idx;
      const end = folded ? folded.originalOffset(foldedEnd) : foldedEnd;
      // Boundary test against the ORIGINAL-cased text (lower-casing never changes
      // word/non-word class for the ranges we test). When the candidate fails the
      // whole-word check we advance by ONE so a valid whole word that starts just
      // after this position isn't skipped; a true match still advances past its
      // end so overlapping occurrences aren't double-counted. A hit that starts
      // or ends inside one original character's folded expansion is no hit.
      if (
        start !== null &&
        end !== null &&
        end > start &&
        (!wholeWord || isWholeWordMatch(text, start, end))
      ) {
        if (!visit(pageIndex, start, end)) return;
        from = foldedEnd;
      } else {
        from = idx + 1;
      }
      idx = haystack.indexOf(compareNeedle, from);
    }
  }
}











function foldCodePoint(codePoint: number): string {
  const char = String.fromCodePoint(codePoint);
  if (codePoint === 0x0131) return char;
  const upper = char.toUpperCase();
  const single =
    upper.length === 1 ||
    (upper.length === 2 && (upper.codePointAt(0) ?? 0) > 0xffff);
  return single ? upper.toLowerCase() : char.toLowerCase();
}

/**
 * Case-fold `text` one code point at a time, remembering where each original
 * code point landed, so an offset in the folded string maps back to the
 * original. `originalOffset` answers null for an offset inside one code
 * point's expansion (such a match does not correspond to whole characters).
 */
function foldCase(text: string): {
  folded: string;
  originalOffset: (foldedIndex: number) => number | null;
} {
  let folded = "";
  // foldedStarts[k] = folded offset where original code unit k begins
  // (undefined for the low half of a surrogate pair); one extra entry marks
  // the end of the string.
  const originalByFolded = new Map<number, number>();
  let index = 0;
  while (index < text.length) {
    const codePoint = text.codePointAt(index) ?? text.charCodeAt(index);
    const width = codePoint > 0xffff ? 2 : 1;
    const piece = foldCodePoint(codePoint);
    originalByFolded.set(folded.length, index);
    folded += piece;
    index += width;
  }
  originalByFolded.set(folded.length, text.length);
  return {
    folded,
    originalOffset: (foldedIndex) => originalByFolded.get(foldedIndex) ?? null,
  };
}

/**
 * Find every occurrence of `query` across the ordered `pageTexts`, returned in
 * document order (page 0 first, then by ascending start offset within a page).
 *
 * Overlapping matches are intentionally NOT produced: after a hit we advance
 * past its end, mirroring how a reader steps through distinct occurrences (so
 * searching "aa" in "aaaa" yields two matches, not three).
 *
 * An empty / whitespace-only query yields no matches. Non-string page entries
 * are treated as empty. `options.limit` caps the result DURING matching (see
 * {@link PdfTextSearchOptions.limit}); without it every match is returned.
 */
export function findMatches(
  pageTexts: ReadonlyArray<string>,
  query: string,
  options: PdfTextSearchOptions = {},
): PdfTextMatch[] {
  const limit = resolveMatchLimit(options.limit);
  const matches: PdfTextMatch[] = [];
  if (limit === 0) return matches;
  forEachMatch(pageTexts, query, options, (pageIndex, start, end) => {
    matches.push({ pageIndex, start, end });
    return matches.length < limit;
  });
  return matches;
}

/**
 * The first `limit` matches (default {@link PDF_FIND_MAX_MATCHES}) plus
 * whether any more exist. One scan with a one-match sentinel: it never builds
 * more than `limit + 1` match objects however many hits the document has. A
 * non-finite `limit` falls back to the default so the window is always
 * bounded; `options.limit` is ignored in favour of the explicit argument.
 *
 * This is the scan the viewer should run ONCE per query: its `matches` drive
 * the find bar and, through `buildDocSearchResults` (doc-search.ts), the
 * results panel, and `truncated` tells both to show that more exist.
 */
export function findMatchWindow(
  pageTexts: ReadonlyArray<string>,
  query: string,
  options: PdfTextSearchOptions = {},
  limit: number = PDF_FIND_MAX_MATCHES,
): PdfTextMatchWindow {
  const size = Number.isFinite(limit)
    ? Math.max(0, Math.trunc(limit))
    : PDF_FIND_MAX_MATCHES;
  const matches = findMatches(pageTexts, query, { ...options, limit: size + 1 });
  const truncated = matches.length > size;
  if (truncated) matches.length = size;
  return { matches, truncated };
}

/**
 * Count every match without building match objects — the exact total
 * {@link findMatches} would return with no cap, at the cost of one scan.
 */
export function countMatches(
  pageTexts: ReadonlyArray<string>,
  query: string,
  options: Pick<PdfTextSearchOptions, "caseSensitive" | "wholeWord"> = {},
): number {
  let count = 0;
  forEachMatch(pageTexts, query, options, () => {
    count += 1;
    return true;
  });
  return count;
}

/**
 * The number of distinct pages that contain at least one match, in document
 * order. Useful for "matches on N pages" style summaries (not used by the
 * default find bar but handy and cheap to expose).
 */
export function matchedPageIndexes(
  matches: ReadonlyArray<PdfTextMatch>,
): number[] {
  const seen = new Set<number>();
  const ordered: number[] = [];
  for (const match of matches) {
    if (seen.has(match.pageIndex)) continue;
    seen.add(match.pageIndex);
    ordered.push(match.pageIndex);
  }
  return ordered;
}

/**
 * Step to the next match index with wraparound.
 *
 * `current` is the currently-active match index (or -1 / null when nothing is
 * active yet, in which case the first match is selected). Returns -1 when there
 * are no matches at all.
 */
export function nextMatchIndex(
  total: number,
  current: number | null,
): number {
  if (!Number.isFinite(total) || total <= 0) return -1;
  if (current == null || current < 0 || current >= total) return 0;
  return (current + 1) % total;
}

/**
 * Step to the previous match index with wraparound.
 *
 * Mirror of {@link nextMatchIndex}: from -1 / null it selects the LAST match so
 * pressing "previous" first jumps to the end of the document. Returns -1 when
 * there are no matches.
 */
export function prevMatchIndex(
  total: number,
  current: number | null,
): number {
  if (!Number.isFinite(total) || total <= 0) return -1;
  if (current == null || current < 0 || current >= total) return total - 1;
  return (current - 1 + total) % total;
}

/**
 * Join per-item text fragments (the `str` values from pdf.js getTextContent)
 * into a single page string while recording the character offset at which each
 * fragment begins. Fragments are joined with the `separator` (default a single
 * space, mirroring how the viewer builds its searchable page text).
 *
 * The returned `text` is exactly what {@link findMatches} should search, and
 * `itemStarts[i]` is the offset of `items[i]` within that text — letting the
 * viewer map a match's character range back to the text-layer span(s) that
 * cover it for highlighting.
 */
export function buildPageText(
  items: ReadonlyArray<string>,
  separator: string = " ",
): { text: string; itemStarts: number[] } {
  const itemStarts: number[] = [];
  let text = "";
  for (let i = 0; i < items.length; i += 1) {
    if (i > 0) text += separator;
    itemStarts.push(text.length);
    const fragment = typeof items[i] === "string" ? items[i] : "";
    text += fragment;
  }
  return { text, itemStarts };
}

/**
 * Given the per-item start offsets from {@link buildPageText} and a match's
 * character range, return the indexes of the items (text-layer spans) that the
 * match overlaps. Used by the viewer to compute the highlight overlay rect for
 * the current match.
 *
 * A match overlaps item `i` when the item's [start, nextStart) span intersects
 * the [matchStart, matchEnd) range. The last item is treated as extending to
 * the end of the text.
 */
export function itemsForRange(
  itemStarts: ReadonlyArray<number>,
  matchStart: number,
  matchEnd: number,
  textLength: number,
): number[] {
  const result: number[] = [];
  if (matchEnd <= matchStart) return result;
  for (let i = 0; i < itemStarts.length; i += 1) {
    const itemStart = itemStarts[i];
    const itemEnd =
      i + 1 < itemStarts.length ? itemStarts[i + 1] : Math.max(textLength, itemStart);
    // Half-open intersection test.
    if (itemStart < matchEnd && itemEnd > matchStart) {
      result.push(i);
    }
  }
  return result;
}

/**
 * Human-facing "n of N" position for a match index. `current` is zero-based;
 * the returned `position` is one-based for display. When there are no matches
 * (or the index is out of range) `position` is 0.
 */
export function matchCountLabel(
  total: number,
  current: number | null,
): { position: number; total: number } {
  const safeTotal = Number.isFinite(total) && total > 0 ? Math.trunc(total) : 0;
  if (
    safeTotal === 0 ||
    current == null ||
    current < 0 ||
    current >= safeTotal
  ) {
    return { position: 0, total: safeTotal };
  }
  return { position: current + 1, total: safeTotal };
}
