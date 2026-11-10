/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the PDF viewer's document-wide search results panel and for
 * sanitizing link-annotation destinations.
 *
 * The viewer already has a per-page find bar (`lib/pdf/text-search.ts`) that
 * steps through matches in document order. This module reuses that
 * `findMatches` primitive to build the data the *results panel* needs: a flat,
 * ordered list of every hit across the whole document, each carrying the page it
 * is on plus a short, sanitized CONTEXT SNIPPET around the match so the panel
 * can show "p.3  …the quick brown fox jumps…" rows the reader can click to jump
 * to.
 *
 * It also holds `sanitizeLinkUrl`, the http/https-only guard the viewer uses
 * before rendering a PDF link annotation's external URL as a real anchor — every
 * other scheme (javascript:, data:, file:, mailto:, …) is rejected so a hostile
 * PDF can't smuggle a dangerous href through an "open link" click.
 *
 * Everything here is pure (no DOM, no pdf.js) so it can be unit-tested without
 * jsdom or a real document. The viewer layers navigation / rendering on top.
 */

import {
  PDF_FIND_MAX_MATCHES,
  countMatches,
  findMatches,
  splitsSurrogatePair,
  type PdfTextMatch,
} from "./text-search";

/**
 * One row in the document-wide search results list: a match plus the precomputed
 * context snippet shown in the panel. Extends the raw character range
 * (U+0000-U+001F), DEL (U+007F), and the C1 range (U+0080-U+009F) that `\s`
 * path the find bar uses (it maps `pageIndex` + `start`/`end` to a page + rect).
 */
export interface PdfDocSearchResult {
  /** Zero-based index into the array of page texts passed to collectMatches. */
  pageIndex: number;
  /** Inclusive character offset of the match start within that page's text. */
  start: number;
  /** Exclusive character offset of the match end within that page's text. */
  end: number;
  /**
   * Short, sanitized context window around the match (whitespace-collapsed,
   * length-capped, rendered as a TEXT node downstream). Leading/trailing
   * ellipses mark where the page text was truncated.
   */
  snippet: string;
}

/**
 * Optional per-page ORIGINAL text for snippet building. When the caller
 * searches NORMALIZED page text (see lib/pdf/find-normalize.ts) the match
 * offsets live in folded space; supplying the raw text plus a range mapper
 * lets the panel quote the page as written (ligatures, accents, original
 * whitespace) instead of the folded form.
 */
export interface PdfDocSearchSnippetSource {
  /** The original page text the snippet should quote. */
  text: string;
  /** Maps a match range in the SEARCHED text into `text` offsets. */
  mapRange: (start: number, end: number) => { start: number; end: number };
}

/** Options for {@link collectMatches}. */
export interface PdfDocSearchOptions {
  /** When true, "Cat" does not match "cat". Defaults to false. */
  caseSensitive?: boolean;
  /** When true, only whole-word occurrences match (see find bar). Defaults to false. */
  wholeWord?: boolean;
  /**
   * Characters of context to include BEFORE the match in the snippet. Negative /
   * non-finite values are treated as 0. Defaults to {@link SNIPPET_BEFORE}.
   */
  before?: number;
  /**
   * Characters of context to include AFTER the match (beyond the matched text
   * itself). Negative / non-finite values are treated as 0. Defaults to
   * {@link SNIPPET_AFTER}.
   */
  after?: number;
  /**
   * Hard cap on the total results returned, protecting the panel from a
   * pathological query ("e") on a huge document producing tens of thousands of
   * rows. The cap is applied WHILE matching, so no match past it is ever
   * built. Excess matches are dropped (document order is preserved, so the
   * kept ones are the first N); a caller that must tell the reader more exist
   * uses `findMatchWindow` + {@link buildDocSearchResults} instead. 0 /
   * negative disables the cap; absent / non-finite uses
   * {@link SNIPPET_MAX_RESULTS}.
   */
  maxResults?: number;
  /**
   * Per-page original text + offset mapping used ONLY for snippets (indexed
   * by `pageIndex`, aligned with `pageTexts`). Pages with a missing / null
   * entry fall back to quoting the searched text. Match offsets in the
   * results are ALWAYS in searched-text space regardless of this option.
   */
  snippetSources?: ReadonlyArray<PdfDocSearchSnippetSource | null | undefined>;
}

/** Default characters of leading context in a snippet. */
export const SNIPPET_BEFORE = 32;
/** Default characters of trailing context (beyond the match) in a snippet. */
export const SNIPPET_AFTER = 48;
/**
 * Default cap on the number of results collected for the panel — the same
 * window as the find bar (`PDF_FIND_MAX_MATCHES`) so rows stay 1:1 with it.
 */
export const SNIPPET_MAX_RESULTS = PDF_FIND_MAX_MATCHES;

/**
 * Whitespace + ASCII/Unicode control characters collapsed in a snippet. `\s`
 * covers tabs/newlines; the explicit ranges strip the C0 controls
 * (U+0000-U+001F), DEL (U+007F), and the C1 range (U+0080-U+009F) that `\s`
 * misses so a hostile text stream can't smuggle control chars into a row.
 */
const SNIPPET_NOISE = /[\s\u0000-\u001f\u007f-\u009f]+/g;

/**
 * Resolve an OPTION window/cap size: a finite number is truncated and clamped to
 * >= 0 (so an explicit negative means "none"); anything else (undefined /
 * non-number / NaN / Infinity) falls back to the supplied default, so an absent
 * option uses the default.
 */
function toCount(value: unknown, fallback: number): number {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.max(0, Math.trunc(value));
  }
  return fallback;
}

/**
 * Clamp a DIRECT window argument to a non-negative integer. Unlike `toCount`,
 * there is no default fallback here — `buildSnippet`'s parameter defaults
 * already supply the default for an omitted argument, so any explicitly-passed
 * non-finite / negative value is treated as 0 ("no context").
 */
function clampWindow(value: number): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? Math.trunc(value)
    : 0;
}

/**
 * Build the sanitized context snippet for a single match within `text`.
 *
 *   - takes `before` chars before `start` and the matched text plus `after`
 *     chars after `end`,
 *   - collapses runaway whitespace + control chars into single spaces (so a
 *     garbage text stream can't blow out the row or smuggle control chars),
 *   - prefixes "…" when context was clipped at the start and suffixes "…" when
 *     clipped at the end.
 *
 * The result is rendered as a React TEXT node downstream — this is normalization
 * (layout safety), NOT HTML escaping. Exported for direct unit testing. An
 * explicitly-passed non-finite / negative `before`/`after` is treated as 0.
 */
export function buildSnippet(
  text: string,
  start: number,
  end: number,
  before: number = SNIPPET_BEFORE,
  after: number = SNIPPET_AFTER,
): string {
  const source = typeof text === "string" ? text : "";
  const len = source.length;
  // Clamp the match range into the text so a malformed offset can't slice past
  // the ends (Array#slice tolerates out-of-range but we want correct ellipses).
  const safeStart = Math.max(0, Math.min(start, len));
  const safeEnd = Math.max(safeStart, Math.min(end, len));
  const beforeCount = clampWindow(before);
  const afterCount = clampWindow(after);

  let windowStart = Math.max(0, safeStart - beforeCount);
  let windowEnd = Math.min(len, safeEnd + afterCount);
  // Never cut an astral character in half at a context edge: widen by the one
  // missing unit so the snippet stays a well-formed string.
  if (splitsSurrogatePair(source, windowStart)) windowStart -= 1;
  if (splitsSurrogatePair(source, windowEnd)) windowEnd += 1;

  const slice = source.slice(windowStart, windowEnd);
  // Collapse + trim AFTER slicing so the ellipsis decisions reflect the raw
  // window edges, not post-collapse length.
  const collapsed = slice.replace(SNIPPET_NOISE, " ").trim();
  const prefix = windowStart > 0 ? "…" : "";
  const suffix = windowEnd < len ? "…" : "";
  if (collapsed.length === 0) return `${prefix}${suffix}`;
  return `${prefix}${collapsed}${suffix}`;
}

/** Snippet options for {@link buildDocSearchResults}. */
export type PdfDocSearchSnippetOptions = Pick<
  PdfDocSearchOptions,
  "before" | "after" | "snippetSources"
>;

/**
 * Build panel rows for matches the caller ALREADY found — no second scan.
 * Rows are 1:1 with `matches` (same order, same offsets), so the viewer can
 * run `findMatchWindow` once and feed the same window to the find bar and to
 * the results panel; a row's index is its find index. `matches` should come
 * from a bounded scan: this builds one snippet per entry it is given.
 */
export function buildDocSearchResults(
  pageTexts: ReadonlyArray<string>,
  matches: ReadonlyArray<PdfTextMatch>,
  options: PdfDocSearchSnippetOptions = {},
): PdfDocSearchResult[] {
  const before = toCount(options.before, SNIPPET_BEFORE);
  const after = toCount(options.after, SNIPPET_AFTER);
  const results: PdfDocSearchResult[] = [];
  for (const match of matches) {
    const raw = pageTexts[match.pageIndex];
    const text = typeof raw === "string" ? raw : "";
    // Quote the ORIGINAL page text when the caller searched a normalized
    // form: map the match range back and slice the source. A malformed
    // source entry (non-string text / missing mapper) falls back to the
    // searched text so a bad page can never sink the whole result list.
    const source = options.snippetSources?.[match.pageIndex];
    let snippet: string;
    if (
      source &&
      typeof source.text === "string" &&
      typeof source.mapRange === "function"
    ) {
      const mapped = source.mapRange(match.start, match.end);
      snippet = buildSnippet(source.text, mapped.start, mapped.end, before, after);
    } else {
      snippet = buildSnippet(text, match.start, match.end, before, after);
    }
    results.push({
      pageIndex: match.pageIndex,
      start: match.start,
      end: match.end,
      snippet,
    });
  }
  return results;
}

/**
 * Collect the occurrences of `query` across the ordered `pageTexts`, in
 * document order (page 0 first, then ascending start offset within a page),
 * each annotated with a sanitized context snippet for the results panel.
 *
 * Reuses `findMatches` from `text-search` so the panel and the find bar share
 * one definition of "what is a match" (case sensitivity, non-overlapping
 * stepping, empty-query handling). An empty / whitespace-only query yields no
 * results. Non-string page entries are treated as empty.
 *
 * The `maxResults` cap (see options) bounds the scan itself, so a 1-char
 * query on a giant document builds at most that many matches. This helper
 * does not say whether matches were left out; callers that show results to a
 * reader use `findMatchWindow` + {@link buildDocSearchResults}, which do.
 */
export function collectMatches(
  pageTexts: ReadonlyArray<string>,
  query: string,
  options: PdfDocSearchOptions = {},
): PdfDocSearchResult[] {
  const cap = toCount(options.maxResults, SNIPPET_MAX_RESULTS);
  const matches: PdfTextMatch[] = findMatches(pageTexts, query, {
    caseSensitive: options.caseSensitive === true,
    wholeWord: options.wholeWord === true,
    ...(cap > 0 ? { limit: cap } : {}),
  });
  if (matches.length === 0) return [];
  return buildDocSearchResults(pageTexts, matches, options);
}

/**
 * Total hit count across the document for the given query (cap-independent).
 * Counts in one scan without building match objects.
 */
export function totalMatchCount(
  pageTexts: ReadonlyArray<string>,
  query: string,
  options: { caseSensitive?: boolean; wholeWord?: boolean } = {},
): number {
  return countMatches(pageTexts, query, {
    caseSensitive: options.caseSensitive === true,
    wholeWord: options.wholeWord === true,
  });
}

// ---- Link-annotation URL sanitization -------------------------------------

/** The ONLY URL schemes a PDF link annotation may open externally. */
const SAFE_LINK_SCHEMES = new Set(["http:", "https:"]);

/**
 * Any whitespace / control char inside a candidate URL. The WHATWG URL parser
 * silently strips tabs/newlines (and some leading/trailing C0 controls) which
 * can be used to disguise a scheme (e.g. "java\tscript:"); rejecting them up
 * front closes that hole. Covers the SAME ranges as {@link SNIPPET_NOISE}:
 * whitespace (`\s`), the C0 controls (U+0000-U+001F), DEL (U+007F), and the C1
 * range (U+0080-U+009F) so a control char in either band can't slip through.
 */
const URL_NOISE = /[\s\u0000-\u001f\u007f-\u009f]/;

/**
 * Validate a PDF link annotation's external URL, returning a safe href string
 * ONLY for `http:` / `https:` URLs and `null` for everything else.
 *
 * pdf.js surfaces link annotations' `url` field verbatim from the document; a
 * hostile PDF can put `javascript:…`, `data:…`, `file://…`, `vbscript:…`, etc.
 * there. We parse with the WHATWG `URL` parser (which normalizes the scheme to
 * lowercase and rejects malformed input) and allow only the two web schemes, so
 * the viewer can render the result directly as an anchor `href` with
 * `target="_blank" rel="noopener noreferrer"` without risk of script/file
 * execution. Protocol-relative (`//host`) and relative URLs have no scheme and
 * are rejected.
 *
 * Returns the parser-normalized href (so callers get a single canonical form),
 * or `null` when the input is not a string, is empty, fails to parse, or uses a
 * non-web scheme.
 */
export function sanitizeLinkUrl(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (trimmed.length === 0) return null;
  // Reject embedded control chars / whitespace before the parser can strip them
  // (which would let "java\nscript:" slip through as "javascript:").
  if (URL_NOISE.test(trimmed)) return null;
  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    return null;
  }
  if (!SAFE_LINK_SCHEMES.has(parsed.protocol)) return null;
  return parsed.href;
}
