/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * pdf.js-grade text normalization for the PDF viewer's find pipeline.
 *
 * The naive pipeline (join item strings with spaces, search with `indexOf`)
 * misses matches whenever the page's text stream differs from what a person
 * types: Unicode ligatures (`ﬁle` vs "file"), diacritics (`café` vs "cafe"),
 * words hyphenated across a line break (`mathe-\nmatics`), and words split
 * across adjacent text items with no space character between them.
 *
 * This module fixes that with the approach used by Mozilla pdf.js's find
 * controller (`web/pdf_find_controller.js`, Apache-2.0 — normalization
 * character classes and page-text assembly rules adapted with attribution;
 * the offset bookkeeping here uses explicit per-character position arrays
 * instead of pdf.js's compressed diff array, trading memory for simplicity):
 *
 *   1. PAGE TEXT ASSEMBLY — item strings are concatenated with NO separator
 *      unless the item ends a line (`hasEOL` → "\n") or the geometry between
 *      adjacent items implies whitespace (a gap along the text's own baseline
 *      → " ", a jump across it → "\n", whatever the rotation). This is what
 *      lets a word split across two items match.
 *   2. FOLDING — the assembled raw text is normalized: an explicit character
 *      table (curly quotes → straight, U+2010 → "-", vulgar fractions),
 *      NFKD decomposition (expands ligatures, superscripts, fullwidth forms),
 *      combining-mark stripping (café → cafe), end-of-line hyphen folding
 *      between letters (mathe-\nmatics → mathematics), soft-hyphen removal,
 *      whitespace-run collapsing (CJK line breaks join with NO space).
 *   3. OFFSET MAPPING — every normalized character records the raw-text range
 *      of its source character, so a match found in the normalized text maps
 *      back to raw offsets → item indexes → text-layer spans for highlighting.
 *
 * Everything here is pure (no DOM, no pdf.js) so it unit-tests without jsdom.
 * The query must be folded with {@link normalizeSearchQuery} before searching
 * the folded page text so both sides agree ("ﬁ" in the query matches "fi" on
 * the page and vice versa).
 */

/** The subset of a pdf.js `TextItem` the searchable-text builder consumes. */
export interface PdfSearchTextItem {
  /** The item's text run (pdf.js `str`). */
  str: string;
  /** True when pdf.js marks an end-of-line after this item. */
  hasEOL?: boolean;
  /** pdf.js text matrix `[a, b, c, d, e, f]`; `e`/`f` are x/y in page space. */
  transform?: ReadonlyArray<number>;
  /** Advance length of the run along its own baseline, in page units. */
  width?: number;
  /** Font height in text space. */
  height?: number;
}

/** Searchable text for one page plus the offset maps back to raw geometry. */
export interface PdfPageSearchText {
  /** Folded, searchable text — search it with a {@link normalizeSearchQuery}'d needle. */
  text: string;
  /** Raw page text: item strings joined with the inferred separators. */
  rawText: string;
  /** Raw-text offset where each item's string begins (1:1 with the input items). */
  itemStarts: number[];
  /** For folded char `i`, the raw offset where its source character starts. */
  normStarts: number[];
  /** For folded char `i`, the raw offset just past its source character. */
  normEnds: number[];
}

/**
 * Explicit single-character folds applied BEFORE NFKD. Adapted from pdf.js's
 * `CHARACTERS_TO_NORMALIZE` (web/pdf_find_controller.js, Apache-2.0): curly
 * quotes → straight quotes, typographic hyphens → "-", vulgar fractions →
 * ASCII, plus the fraction slash NFKD itself would emit for other fractions.
 */
const NORMALIZE_CHAR_MAP: Readonly<Record<string, string>> = {
  "‐": "-", // hyphen
  "‑": "-", // non-breaking hyphen
  "‘": "'", // left single quotation mark
  "’": "'", // right single quotation mark
  "‚": "'", // single low-9 quotation mark
  "‛": "'", // single high-reversed-9 quotation mark
  "“": '"', // left double quotation mark
  "”": '"', // right double quotation mark
  "„": '"', // double low-9 quotation mark
  "‟": '"', // double high-reversed-9 quotation mark
  "¼": "1/4", // vulgar fraction one quarter
  "½": "1/2", // vulgar fraction one half
  "¾": "3/4", // vulgar fraction three quarters
  "⁄": "/", // fraction slash (NFKD output for other vulgar fractions)
};

/**
 * A character that joins across an UNSPACED line break: ideographs plus the
 * Japanese kana ranges (same class pdf.js uses for its CJK-newline rule) —
 * CJK prose has no inter-word spaces, so a "\n" between two such characters
 * is purely visual and must vanish rather than become a space.
 */
function isCjkAt(value: string, index: number): boolean {
  const cp = codePointAtBoundary(value, index);
  if (cp === undefined) return false;
  return /[\p{Ideographic}぀-ヿ]/u.test(String.fromCodePoint(cp));
}

/** Whether the character at (or containing) `index` is a Unicode letter. */
function isLetterAt(value: string, index: number): boolean {
  const cp = codePointAtBoundary(value, index);
  if (cp === undefined) return false;
  return /\p{L}/u.test(String.fromCodePoint(cp));
}

/**
 * Code point at `index`, stepping back one unit when `index` lands on a
 * trailing surrogate so "last character" checks see the whole astral char.
 */
function codePointAtBoundary(value: string, index: number): number | undefined {
  if (index < 0 || index >= value.length) return undefined;
  const unit = value.charCodeAt(index);
  if (unit >= 0xdc00 && unit <= 0xdfff && index > 0) {
    return value.codePointAt(index - 1);
  }
  return value.codePointAt(index);
}

/** Result of folding a raw string (page text or query) for search. */
export interface FoldedText {
  /** The folded text. */
  text: string;
  /** Per folded char: raw offset where its source character starts. */
  normStarts: number[];
  /** Per folded char: raw offset just past its source character. */
  normEnds: number[];
}

export interface FoldTextOptions {
  readonly keepDiacritics?: boolean;
}

/**
 * Fold a raw string into its searchable form, recording for every emitted
 * character the raw range of its source character. The folds (in order per
 * source character):
 *
 *   - soft hyphens (U+00AD) vanish; when one splits a word across a line
 *     break the "\n" vanishes with it,
 *   - the {@link NORMALIZE_CHAR_MAP} table (quotes / hyphens / fractions),
 *   - an end-of-line hyphen BETWEEN LETTERS folds away with its "\n"
 *     ("mathe-\nmatics" → "mathematics"),
 *   - a whitespace run collapses to ONE space — except a run containing a
 *     line break between two CJK characters, which vanishes entirely; leading
 *     and trailing runs also vanish,
 *   - everything else is NFKD-decomposed (ligatures → letters, superscripts /
 *     fullwidth → ASCII) with combining marks stripped (café → cafe). A
 *     source character that is ITSELF a combining mark disappears, extending
 *     the previous emitted character's raw range so a match ending there
 *     still covers the mark in raw space.
 *
 * Case is preserved — the caller's case-sensitivity handling is unchanged.
 *
 * `keepDiacritics` (区分重音, as WPS's default find and Word's Match
 * diacritics): marks stay on their letters, so Å finds Å and never a plain
 * A (user audit PDF 0038). Text and query are folded the same way.
 */
export function foldTextForSearch(raw: string, options: FoldTextOptions = {}): FoldedText {
  const source = typeof raw === "string" ? raw : "";
  let text = "";
  const normStarts: number[] = [];
  const normEnds: number[] = [];
  // How many folded chars the most recent source character emitted, so a
  // following combining mark can extend exactly that character's raw range.
  let lastEmitCount = 0;

  const emit = (chars: string, start: number, end: number): void => {
    for (let k = 0; k < chars.length; k += 1) {
      text += chars[k];
      normStarts.push(start);
      normEnds.push(end);
    }
    lastEmitCount = chars.length;
  };

  let i = 0;
  while (i < source.length) {
    const cp = source.codePointAt(i) as number;
    const ch = String.fromCodePoint(cp);
    const chLen = ch.length;

    // Soft hyphen: invisible; drop it (and a following line break when the
    // soft hyphen splits a word, mirroring the visible-hyphen fold below).
    if (ch === "\u00ad") {
      if (
        source[i + 1] === "\n" &&
        isLetterAt(text, text.length - 1) &&
        isLetterAt(source, i + 2)
      ) {
        i += 2;
      } else {
        i += 1;
      }
      continue;
    }

    const mapped = NORMALIZE_CHAR_MAP[ch];

    // End-of-line hyphen between letters folds away with its line break.
    if (
      (mapped ?? ch) === "-" &&
      source[i + chLen] === "\n" &&
      isLetterAt(text, text.length - 1) &&
      isLetterAt(source, i + chLen + 1)
    ) {
      i += chLen + 1;
      continue;
    }

    if (/\s/u.test(ch)) {
      // Absorb the whole whitespace run, then decide what (if anything) it
      // becomes. Whitespace code points are all BMP, so unit-wise is safe.
      let j = i + chLen;
      let sawLineBreak = ch === "\n" || ch === "\r";
      while (j < source.length && /\s/u.test(source[j])) {
        if (source[j] === "\n" || source[j] === "\r") sawLineBreak = true;
        j += 1;
      }
      const hasPrev = text.length > 0;
      const hasNext = j < source.length;
      const joinsCjkLineBreak =
        sawLineBreak &&
        isCjkAt(text, text.length - 1) &&
        isCjkAt(source, j);
      if (hasPrev && hasNext && !joinsCjkLineBreak) {
        emit(" ", i, j);
      } else {
        // The run vanished — a combining mark after it must not extend a
        // character on the other side of the (former) whitespace.
        lastEmitCount = 0;
      }
      i = j;
      continue;
    }

    const base = mapped ?? ch;
    const decomposed = base.normalize("NFKD");
    if (options.keepDiacritics && /^\p{M}+$/u.test(decomposed)) {
      // A mark of its own stays with the letter before it.
      emit(decomposed, i, i + chLen);
      i += chLen;
      continue;
    }
    const stripped = options.keepDiacritics
      ? decomposed
      : decomposed.replace(/\p{M}+/gu, "");
    if (stripped.length === 0) {
      // Pure combining mark: strip it, extending the previous character's raw
      // range so a raw-space highlight still covers the mark's code units.
      for (let k = 1; k <= lastEmitCount; k += 1) {
        normEnds[normEnds.length - k] = i + chLen;
      }
      i += chLen;
      continue;
    }
    emit(stripped, i, i + chLen);
    i += chLen;
  }

  return { text, normStarts, normEnds };
}

/**
 * Fold a user-typed query the same way page text is folded, so both sides of
 * the search agree on ligatures, diacritics, quotes, and whitespace. Returns
 * "" for non-string / effectively-empty input (the match helpers treat an
 * empty needle as "no matches").
 */
export function normalizeSearchQuery(
  query: string,
  options: FoldTextOptions = {},
): string {
  return foldTextForSearch(typeof query === "string" ? query : "", options).text;
}

/** A page's searchable text folded with its marks kept (区分重音). */
export function pageSearchTextKeepingDiacritics(
  entry: PdfPageSearchText,
): PdfPageSearchText {
  const folded = foldTextForSearch(entry.rawText, { keepDiacritics: true });
  return {
    text: folded.text,
    rawText: entry.rawText,
    itemStarts: entry.itemStarts,
    normStarts: folded.normStarts,
    normEnds: folded.normEnds,
  };
}

/**
 * Unit vector along an item's BASELINE (its advance direction) in page space.
 * pdf.js text transforms are `[a, b, c, d, e, f]` where `(a, b)` is the scaled
 * advance axis and `(c, d)` the scaled glyph-up axis, so rotation, mirroring
 * and skew all live in the matrix. Falls back to the up axis turned a quarter
 * clockwise when the advance axis is degenerate (e.g. `Tz 0`), and to null
 * when both are.
 */
function baselineDirection(
  transform: ReadonlyArray<number>,
): { x: number; y: number } | null {
  const a = transform[0] ?? 0;
  const b = transform[1] ?? 0;
  const advance = Math.hypot(a, b);
  if (advance > 0) return { x: a / advance, y: b / advance };
  const c = transform[2] ?? 0;
  const d = transform[3] ?? 0;
  const up = Math.hypot(c, d);
  if (up > 0) return { x: d / up, y: -c / up };
  return null;
}

/**
 * Minimum cosine between two items' baselines for them to count as the SAME
 * text direction (about 26 degrees). Items of one line share their matrix
 * orientation; a label turned against its neighbour is a separate line even
 * when its origin happens to sit where the neighbour's advance ends.
 */
const SAME_BASELINE_MIN_COS = 0.9;

/**
 * Whitespace the geometry-based separator inference implies between two
 * adjacent items (only consulted when the earlier item does NOT carry
 * `hasEOL`). The origin step between the items is projected onto the earlier
 * item's own baseline (`along`) and its perpendicular (`across`), so 0/90/
 * 180/270-degree and arbitrarily rotated text all use one rule:
 *
 *   - no usable geometry on either side → " " (the legacy join, safe),
 *   - baselines pointing in clearly different directions → "\n" (a turned
 *     label or watermark is its own line),
 *   - an `across` offset of more than half the font size → "\n" (new line),
 *   - an `along` gap past the earlier item's advance wider than ~a third of
 *     the font size → " ",
 *   - a large BACKWARD `along` jump on the same baseline → " " (new column /
 *     region; small negatives are kerning overlap and join),
 *   - otherwise the items abut → "" (this is what heals split words).
 *
 * `width` is pdf.js's advance length along the item's baseline, which is
 * rotation-independent, so it compares directly with `along`.
 */
function separatorBetween(
  prev: PdfSearchTextItem,
  next: PdfSearchTextItem,
): string {
  const pt = prev.transform;
  const nt = next.transform;
  if (!pt || pt.length < 6 || !nt || nt.length < 6) return " ";
  const fontSize =
    Math.hypot(pt[2] ?? 0, pt[3] ?? 0) ||
    Math.hypot(pt[0] ?? 0, pt[1] ?? 0) ||
    (typeof prev.height === "number" && prev.height > 0 ? prev.height : 0) ||
    1;
  const prevBaseline = baselineDirection(pt);
  const nextBaseline = baselineDirection(nt);
  if (
    prevBaseline &&
    nextBaseline &&
    prevBaseline.x * nextBaseline.x + prevBaseline.y * nextBaseline.y <
      SAME_BASELINE_MIN_COS
  ) {
    return "\n";
  }
  // A fully degenerate matrix carries no direction: read it upright, which is
  // the page-axis behaviour this function had before it projected.
  const baseline = prevBaseline ?? { x: 1, y: 0 };
  const dx = (nt[4] ?? 0) - (pt[4] ?? 0);
  const dy = (nt[5] ?? 0) - (pt[5] ?? 0);
  const along = dx * baseline.x + dy * baseline.y;
  const across = -dx * baseline.y + dy * baseline.x;
  if (Math.abs(across) > fontSize * 0.5) return "\n";
  const prevWidth =
    typeof prev.width === "number" && Number.isFinite(prev.width) && prev.width > 0
      ? prev.width
      : 0;
  const gap = along - prevWidth;
  if (gap > fontSize * 0.3) return " ";
  if (gap < -Math.max(prevWidth, fontSize)) return " ";
  return "";
}

/**
 * Assemble one page's searchable text from its text items (pdf.js
 * `getTextContent` items, or plain strings for the OCR-runs fallback) and
 * fold it for search.
 *
 * Items are concatenated with NO separator unless `hasEOL` (→ "\n") or the
 * inter-item geometry implies whitespace (see {@link separatorBetween}).
 * Plain-string items have no geometry and join with a single space — the
 * legacy behavior the OCR path relies on. `itemStarts[i]` is the raw-text
 * offset of item `i`, 1:1 with the text-layer spans pdf.js renders, so a
 * match's raw range maps to spans via `itemsForRange`.
 */
export function buildSearchablePageText(
  items: ReadonlyArray<PdfSearchTextItem | string>,
): PdfPageSearchText {
  const itemStarts: number[] = [];
  let rawText = "";
  // The last item that contributed text — separators are only inferred
  // between CONTENT, so empty items (marked content, artifacts) are inert
  // apart from a `hasEOL` they may carry.
  let prev: PdfSearchTextItem | null = null;
  let pendingEol = false;

  for (const rawItem of items) {
    const item: PdfSearchTextItem =
      typeof rawItem === "string" ? { str: rawItem } : rawItem;
    const str = typeof item.str === "string" ? item.str : "";
    if (str.length > 0 && prev) {
      rawText +=
        prev.hasEOL === true || pendingEol
          ? "\n"
          : separatorBetween(prev, item);
    }
    itemStarts.push(rawText.length);
    rawText += str;
    if (str.length > 0) {
      prev = item;
      pendingEol = false;
    } else if (item.hasEOL === true) {
      pendingEol = true;
    }
  }

  const folded = foldTextForSearch(rawText);
  return {
    text: folded.text,
    rawText,
    itemStarts,
    normStarts: folded.normStarts,
    normEnds: folded.normEnds,
  };
}

/**
 * Coerce one raw pdf.js text-content item (untyped over the wire) into the
 * shape {@link buildSearchablePageText} consumes. Anything without a string
 * `str` becomes an empty item, matching how the viewer already treats
 * malformed entries (they still occupy a span slot for offset alignment).
 */
export function searchTextItemFromPdfItem(item: unknown): PdfSearchTextItem {
  if (
    !item ||
    typeof item !== "object" ||
    typeof (item as { str?: unknown }).str !== "string"
  ) {
    return { str: "" };
  }
  const it = item as {
    str: string;
    hasEOL?: unknown;
    transform?: unknown;
    width?: unknown;
    height?: unknown;
  };
  const transform =
    Array.isArray(it.transform) &&
    it.transform.length >= 6 &&
    it.transform
      .slice(0, 6)
      .every((v) => typeof v === "number" && Number.isFinite(v))
      ? (it.transform as number[])
      : undefined;
  return {
    str: it.str,
    hasEOL: it.hasEOL === true,
    transform,
    width:
      typeof it.width === "number" && Number.isFinite(it.width)
        ? it.width
        : undefined,
    height:
      typeof it.height === "number" && Number.isFinite(it.height)
        ? it.height
        : undefined,
  };
}

/**
 * Map a match range in FOLDED text back to the raw-text range it came from.
 * The result is what `itemsForRange` (raw `itemStarts`) and the text-layer
 * span sub-ranges consume. A match ending inside a multi-char expansion
 * (e.g. only the "f" of a folded "ﬁ") maps to the WHOLE source character —
 * the same whole-glyph highlight pdf.js produces. Out-of-range inputs are
 * clamped; an empty range maps to an empty raw range at the equivalent spot.
 */
export function rawRangeForNormalizedRange(
  entry: Pick<
    PdfPageSearchText,
    "text" | "rawText" | "normStarts" | "normEnds"
  >,
  start: number,
  end: number,
): { start: number; end: number } {
  const len = entry.text.length;
  const s = Math.max(0, Math.min(Math.trunc(start), len));
  const e = Math.max(s, Math.min(Math.trunc(end), len));
  if (len === 0 || e === s) {
    const at = s < len ? (entry.normStarts[s] ?? 0) : entry.rawText.length;
    return { start: at, end: at };
  }
  const rawStart = entry.normStarts[s] ?? 0;
  const rawEnd = entry.normEnds[e - 1] ?? entry.rawText.length;
  return { start: rawStart, end: Math.max(rawStart, rawEnd) };
}
