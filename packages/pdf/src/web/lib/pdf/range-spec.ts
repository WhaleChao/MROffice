/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure page-range geometry for FILE-LEVEL PDF split.
 *
 * This module is the single source of truth for turning a human-typed
 * page-range spec (Acrobat-style, e.g. `"1-3, 5, 8-"`) into concrete,
 * validated, clamped 1-based page-number groups, and for chopping a document
 * into fixed-size chunks ("split every N pages"). It has ZERO pdf-lib / DOM /
 * React dependency so it can be unit-tested in isolation and reused by the
 * pdf-lib writer in `lib/pdf/merge-split.ts` and by the PDFViewer UI.
 *
 * Design notes:
 *  - Page numbers are 1-based everywhere a HUMAN sees them; the pdf-lib writer
 *    converts to 0-based indices via {@link rangeGroupToZeroBased}.
 *  - An OPEN range (`"8-"` or `"-3"`) clamps to the document's first / last
 *    page, matching Acrobat's "to end" / "from start" behaviour.
 *  - Out-of-range and reversed endpoints are clamped, not thrown: a spec like
 *    `"0-999"` against a 5-page doc yields `[1,2,3,4,5]` so a fat-fingered
 *    range never errors, it just selects what exists.
 *  - Duplicate / overlapping pages WITHIN one group are de-duped and sorted, so
 *    `"3,1-2,2"` → `[1,2,3]` and a single output file never repeats a page.
 *  - Whitespace is ignored; both ASCII hyphen `-` and Unicode en/em dashes
 *    (`–` `—`) are accepted as range separators (users paste these).
 */

/** One validated, clamped, 1-based group of page numbers (one output file). */
export interface RangeGroup {
  /** Sorted, de-duped, in-range 1-based page numbers. Never empty. */
  pages: number[];
  /** Stable 1-based index of this group (its position in the parse), for labels. */
  index: number;
  /** A human label like "1-3" or "5" derived from the contiguous runs. */
  label: string;
}

/** Defensive upper bound on a single document's page count we will address. */
export const MAX_SPLIT_PAGES = 10_000;
/** Defensive upper bound on how many comma-separated tokens we will parse. */
export const MAX_RANGE_TOKENS = 5_000;
/** Length cap on a raw spec string accepted from untrusted input. */
export const MAX_SPEC_LENGTH = 20_000;

/** Normalize Unicode dashes to ASCII hyphen and strip all whitespace. */
function normalizeSpec(spec: string): string {
  return spec
    .replace(/[‒–—―−]/g, "-") // figure/en/em/horiz-bar/minus
    .replace(/\s+/g, "");
}

/** Clamp a value into [1, totalPages] (returns null when totalPages < 1). */
function clampPage(value: number, totalPages: number): number | null {
  if (totalPages < 1) return null;
  if (!Number.isFinite(value)) return null;
  const n = Math.trunc(value);
  if (n < 1) return 1;
  if (n > totalPages) return totalPages;
  return n;
}

/** Sort, de-dupe, drop non-positive — the canonical form of a page list. */
function canonicalize(pages: number[]): number[] {
  const seen = new Set<number>();
  for (const p of pages) {
    if (Number.isInteger(p) && p >= 1) seen.add(p);
  }
  return [...seen].sort((a, b) => a - b);
}

/**
 * Derive a compact "1-3,5,8-9" style label from a sorted, de-duped page list by
 * collapsing contiguous runs. Empty input → "" (callers never pass empty).
 */
export function pagesToLabel(pages: number[]): string {
  const sorted = canonicalize(pages);
  if (sorted.length === 0) return "";
  const runs: string[] = [];
  let runStart = sorted[0];
  let prev = sorted[0];
  for (let i = 1; i <= sorted.length; i += 1) {
    const cur = sorted[i];
    if (cur === prev + 1) {
      prev = cur;
      continue;
    }
    runs.push(runStart === prev ? `${runStart}` : `${runStart}-${prev}`);
    if (cur === undefined) break;
    runStart = cur;
    prev = cur;
  }
  return runs.join(",");
}

/**
 * The clamped, inclusive `[start, end]` bounds a SINGLE comma-token denotes,
 * or null when it denotes nothing — {@link parseRangeToken} without the
 * expansion, so callers can size, merge and bound ranges before allocating a
 * page number per page. A non-finite `totalPages` denotes no document and
 * resolves nothing, so no range can escape the page count unclamped.
 */
function rangeTokenBounds(
  token: string,
  totalPages: number,
): [number, number] | null {
  if (!Number.isFinite(totalPages) || totalPages < 1) return null;
  const trimmed = token.trim();
  if (trimmed === "") return null;

  // Plain single page "5".
  if (/^\d+$/.test(trimmed)) {
    const p = clampPage(Number(trimmed), totalPages);
    return p === null ? null : [p, p];
  }

  // Range forms: "1-3", "8-" (open end), "-3" (open start). Exactly one hyphen.
  const dashMatch = /^(\d*)-(\d*)$/.exec(trimmed);
  if (!dashMatch) return null; // garbage like "1-2-3" or "a-b"
  const [, startRaw, endRaw] = dashMatch;
  if (startRaw === "" && endRaw === "") return null; // bare "-" denotes nothing

  const startNum = startRaw === "" ? 1 : Number(startRaw);
  const endNum = endRaw === "" ? totalPages : Number(endRaw);
  let start = clampPage(startNum, totalPages);
  let end = clampPage(endNum, totalPages);
  if (start === null || end === null) return null;
  if (start > end) [start, end] = [end, start]; // normalize reversed
  return [start, end];
}

/**
 * The canonical (integer, >= 1) pages of token bounds as an integer interval,
 * or null when none survive {@link canonicalize}. Expansion steps by 1 from
 * `start`, so a non-integer `start` (only reachable through a non-integer
 * `totalPages`) yields no integer page at all.
 */
function canonicalBounds(
  bounds: readonly [number, number],
): [number, number] | null {
  const [start, end] = bounds;
  if (!Number.isInteger(start) || start < 1) return null;
  const last = Math.floor(end);
  return last >= start ? [start, last] : null;
}

/** Tokenize a raw spec exactly as every parser here does (caps included). */
function specTokens(spec: string): string[] {
  const normalized = normalizeSpec(spec.slice(0, MAX_SPEC_LENGTH));
  return normalized.split(",").slice(0, MAX_RANGE_TOKENS);
}

/** Every page from `start` up to `end` in steps of one, ascending. */
function expandInterval(start: number, end: number): number[] {
  const pages: number[] = [];
  for (let p = start; p <= end; p += 1) pages.push(p);
  return pages;
}

/**
 * Parse a SINGLE comma-token (e.g. `"8-"`, `"1-3"`, `"5"`, `"-3"`) into the
 * clamped, inclusive 1-based page list it denotes against a document of
 * `totalPages` pages. Returns `[]` for a token that resolves to nothing
 * (empty, all out of range with no clamp target, or pure garbage). Reversed
 * endpoints (`"5-2"`) are normalized to ascending. Open ends clamp to the
 * document bounds.
 */
export function parseRangeToken(token: string, totalPages: number): number[] {
  const bounds = rangeTokenBounds(token, totalPages);
  return bounds ? expandInterval(bounds[0], bounds[1]) : [];
}

/**
 * Parse a full Acrobat-style spec (`"1-3, 5, 8-"`) into the FLAT, sorted,
 * de-duped 1-based page list it selects against `totalPages`. Tokens that
 * resolve to nothing are skipped. Returns `[]` for an empty / all-garbage spec.
 * Use this when the split should produce ONE file from the union of all ranges.
 *
 * Tokens are reduced to intervals and MERGED before any page is expanded, so
 * repeated or overlapping ranges ("1-10000" two hundred times) cost one sort
 * of at most {@link MAX_RANGE_TOKENS} intervals plus one pass over the
 * selected pages — never tokens x pages.
 */
export function parseRangeSpec(spec: string, totalPages: number): number[] {
  if (typeof spec !== "string" || totalPages < 1) return [];
  const intervals: Array<[number, number]> = [];
  for (const token of specTokens(spec)) {
    const bounds = rangeTokenBounds(token, totalPages);
    const interval = bounds ? canonicalBounds(bounds) : null;
    if (interval) intervals.push(interval);
  }
  intervals.sort((a, b) => a[0] - b[0]);
  const pages: number[] = [];
  let next = 1; // smallest page not yet emitted
  for (const [start, end] of intervals) {
    for (let p = Math.max(start, next); p <= end; p += 1) pages.push(p);
    if (end + 1 > next) next = end + 1;
  }
  return pages;
}

/** Why a typed page range cannot be taken literally. */
export interface RangeSpecIssue {
  kind: "syntax" | "out-of-range";
  /** The offending comma-token as the user typed it (whitespace removed). */
  token: string;
}

/**
 * Strict reading of a typed page range for commands that act on the pages
 * the user NAMED (delete, rotate, resize, crop, ...): unlike
 * {@link parseRangeSpec}, which clamps "0" to page 1 for split's forgiving
 * "select what exists", a page number outside 1..totalPages or an
 * unreadable token is reported instead of silently standing in for another
 * page. Open ranges ("8-", "-3") and reversed ranges ("5-2") keep their
 * documented meaning; empty tokens (a trailing comma) are ignored. `pages`
 * is empty whenever `issue` is set.
 */
export function validateRangeSpec(
  spec: string,
  totalPages: number,
): { pages: number[]; issue: RangeSpecIssue | null } {
  if (typeof spec !== "string") {
    return { pages: [], issue: { kind: "syntax", token: "" } };
  }
  const total =
    Number.isFinite(totalPages) && totalPages >= 1 ? Math.trunc(totalPages) : 0;
  for (const token of specTokens(spec)) {
    if (token === "") continue;
    const single = /^\d+$/.test(token);
    const dash = single ? null : /^(\d*)-(\d*)$/.exec(token);
    if (!single && (!dash || (dash[1] === "" && dash[2] === ""))) {
      return { pages: [], issue: { kind: "syntax", token } };
    }
    const numbers = single
      ? [Number(token)]
      : [dash?.[1] ?? "", dash?.[2] ?? ""]
          .filter((part) => part !== "")
          .map(Number);
    if (numbers.some((value) => value < 1 || value > total)) {
      return { pages: [], issue: { kind: "out-of-range", token } };
    }
  }
  return { pages: parseRangeSpec(spec, total), issue: null };
}

/**
 * How many groups {@link parseRangeGroups} would return, computed from the
 * token bounds WITHOUT expanding any page list. Callers that cap the number
 * of output files check this first, so an over-limit spec ("1-" five
 * thousand times) is refused before tokens x pages numbers are allocated.
 */
export function countRangeGroups(spec: string, totalPages: number): number {
  if (typeof spec !== "string" || totalPages < 1) return 0;
  let count = 0;
  for (const token of specTokens(spec)) {
    const bounds = rangeTokenBounds(token, totalPages);
    if (bounds && canonicalBounds(bounds)) count += 1;
  }
  return count;
}

/**
 * Parse a spec into ONE {@link RangeGroup} PER comma-token, so each token
 * becomes a separate output file (Acrobat's "Split by ranges → one file per
 * range"). Tokens that resolve to no pages are dropped (no empty file). Within
 * a group, pages are sorted + de-duped. Returns `[]` when nothing resolves.
 * A token is one contiguous interval, so each group expands straight from its
 * bounds; use {@link countRangeGroups} to refuse too many groups before this
 * allocates them.
 */
export function parseRangeGroups(spec: string, totalPages: number): RangeGroup[] {
  if (typeof spec !== "string" || totalPages < 1) return [];
  const groups: RangeGroup[] = [];
  for (const token of specTokens(spec)) {
    const bounds = rangeTokenBounds(token, totalPages);
    const interval = bounds ? canonicalBounds(bounds) : null;
    if (!interval) continue;
    const [start, end] = interval;
    groups.push({
      pages: expandInterval(start, end),
      index: groups.length + 1,
      label: start === end ? `${start}` : `${start}-${end}`,
    });
  }
  return groups;
}

/**
 * Split a document of `totalPages` into fixed-size groups of `chunkSize` pages
 * each ("split every N pages"). The final group may be shorter. `chunkSize` is
 * clamped to `[1, totalPages]` so 0 / negative / NaN never produces an empty or
 * runaway result. Returns `[]` only when `totalPages < 1`.
 *
 * Example: totalPages=5, chunkSize=2 → [[1,2],[3,4],[5]].
 */
export function splitEveryN(totalPages: number, chunkSize: number): RangeGroup[] {
  if (!Number.isFinite(totalPages) || totalPages < 1) return [];
  const total = Math.min(Math.trunc(totalPages), MAX_SPLIT_PAGES);
  let size = Math.trunc(Number(chunkSize));
  if (!Number.isFinite(size) || size < 1) size = 1;
  if (size > total) size = total;
  const groups: RangeGroup[] = [];
  for (let start = 1; start <= total; start += size) {
    const end = Math.min(start + size - 1, total);
    const pages: number[] = [];
    for (let p = start; p <= end; p += 1) pages.push(p);
    groups.push({ pages, index: groups.length + 1, label: pagesToLabel(pages) });
  }
  return groups;
}

/**
 * Build split groups from the START pages of a document's TOP-LEVEL bookmarks
 * (Acrobat "Split → Top-level bookmarks" / one file per outline section).
 *
 * Each top-level bookmark begins a section that runs UNTIL the next top-level
 * bookmark's start page (exclusive); the last bookmark's section runs to the end
 * of the document. `startPages` are 1-based page numbers (in any order, possibly
 * duplicated or out of range) — they are clamped to `[1, totalPages]`, sorted,
 * and de-duped first, so two bookmarks on the SAME page collapse into one
 * section and a stale/garbage target never escapes the document.
 *
 * If the first bookmark does not start on page 1, the pages BEFORE it
 * (`1 .. firstStart-1`, the front matter) become their own leading group so no
 * page is dropped from the split — mirroring Acrobat, which emits the cover /
 * preface as the first file. When `startPages` resolves to nothing usable, the
 * whole document is returned as a single group (so a bookmark-less PDF still
 * splits to exactly itself rather than to nothing).
 *
 * Pure page math (no pdf-lib): feed the result to {@link parseRangeGroups}'s
 * sibling {@link RangeGroup} consumers (e.g. `splitPdf`). Returns `[]` only when
 * `totalPages < 1`.
 */
export function bookmarkSplitGroups(
  startPages: ReadonlyArray<number>,
  totalPages: number,
): RangeGroup[] {
  if (!Number.isFinite(totalPages) || totalPages < 1) return [];
  const total = Math.min(Math.trunc(totalPages), MAX_SPLIT_PAGES);

  // Clamp every start onto the document, then sort + de-dupe.
  const clamped: number[] = [];
  for (const raw of startPages) {
    const p = clampPage(raw, total);
    if (p !== null) clamped.push(p);
  }
  const starts = canonicalize(clamped);

  // No usable bookmark → the whole document is one group.
  if (starts.length === 0) {
    const pages: number[] = [];
    for (let p = 1; p <= total; p += 1) pages.push(p);
    return [{ pages, index: 1, label: pagesToLabel(pages) }];
  }

  // Section boundaries: each start, plus a leading boundary at page 1 for any
  // front matter before the first bookmark.
  const boundaries = starts[0] > 1 ? [1, ...starts] : [...starts];

  const groups: RangeGroup[] = [];
  for (let i = 0; i < boundaries.length; i += 1) {
    const start = boundaries[i];
    const end = i + 1 < boundaries.length ? boundaries[i + 1] - 1 : total;
    if (end < start) continue; // adjacent boundaries (shouldn't happen post-dedupe)
    const pages: number[] = [];
    for (let p = start; p <= end; p += 1) pages.push(p);
    if (pages.length === 0) continue;
    groups.push({ pages, index: groups.length + 1, label: pagesToLabel(pages) });
  }
  return groups;
}

/** Convert a group's 1-based page numbers to 0-based indices for pdf-lib. */
export function rangeGroupToZeroBased(group: RangeGroup): number[] {
  return group.pages.map((p) => p - 1);
}
