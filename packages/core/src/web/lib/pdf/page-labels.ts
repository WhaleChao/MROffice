/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF PAGE LABELS — the document's LOGICAL page numbering
 * (Acrobat "Page Thumbnails ▸ Number Pages", WPS 页码 / 起始页码). This is the
 * `/PageLabels` number tree in the PDF catalog: it tells a viewer to DISPLAY a
 * page as e.g. "iii" or "A-2" instead of its physical 1-based index, without
 * painting anything on the page (distinct from the Bates / page-number STAMP in
 * `stamps.ts`, which draws ink onto the content).
 *
 * The model is a list of {@link PageLabelRange}s. Each range starts at a 0-based
 * page index and applies a numbering STYLE (decimal, upper/lower roman,
 * upper/lower alphabetic, or none), an optional textual PREFIX, and a START
 * value, to every page from its start index up to (but not including) the next
 * range's start. Example: front matter as lowercase roman then the body as
 * decimal →
 *
 *   [ { start: 0, style: "lower-roman" },
 *     { start: 4, style: "decimal", firstValue: 1 } ]
 *
 * yields i, ii, iii, iv, 1, 2, 3, …
 *
 * Everything here is DOM-free / pdf.js-free / pdf-lib-free so it unit-tests in
 * isolation. The viewer maps the {@link buildPageLabelNums} descriptor onto a
 * pdf-lib `/PageLabels` dict; the numeral formatting is reused for any UI preview
 * of the resulting labels. Mirrors the defensive, total style of `stamps.ts`.
 */

/** The numbering styles a page-label range can use (PDF /S values + "none"). */
export type PageLabelStyle =
  | "decimal" // /D   1, 2, 3
  | "upper-roman" // /R   I, II, III
  | "lower-roman" // /r   i, ii, iii
  | "upper-alpha" // /A   A, B, …, Z, AA
  | "lower-alpha" // /a   a, b, …, z, aa
  | "none"; // (no /S) — prefix only, no number

export const PAGE_LABEL_STYLES: readonly PageLabelStyle[] = [
  "decimal",
  "upper-roman",
  "lower-roman",
  "upper-alpha",
  "lower-alpha",
  "none",
] as const;

/** The single-letter PDF `/S` style code for a label style (null for "none"). */
export function pdfStyleCode(style: PageLabelStyle): string | null {
  switch (style) {
    case "decimal":
      return "D";
    case "upper-roman":
      return "R";
    case "lower-roman":
      return "r";
    case "upper-alpha":
      return "A";
    case "lower-alpha":
      return "a";
    case "none":
      return null;
  }
}

/** Narrow an arbitrary value to a page-label style (defaults to "decimal"). */
export function asPageLabelStyle(value: unknown): PageLabelStyle {
  return typeof value === "string" &&
    (PAGE_LABEL_STYLES as readonly string[]).includes(value)
    ? (value as PageLabelStyle)
    : "decimal";
}

/** Max characters kept for a label prefix (a runaway prefix is truncated). */
export const PAGE_LABEL_PREFIX_MAX = 64;

/** Largest page count we will build labels for (defensive upper bound). */
export const PAGE_LABEL_MAX_PAGES = 10_000;

/** Largest persisted `/St` value accepted by the native page authority. */
export const PAGE_LABEL_FIRST_VALUE_MAX = 1_000_000;

/** Largest page value any label range can produce (`/St` max + page offset). */
const ROMAN_VALUE_MAX = PAGE_LABEL_FIRST_VALUE_MAX + PAGE_LABEL_MAX_PAGES;

/**
 * One labeling range. `start` is the 0-based index of the FIRST page the range
 * governs. `firstValue` is the numeric value assigned to that first page
 * (default 1; PDF's `/St`). `prefix` is prepended to every label in the range
 * (PDF's `/P`). The range extends until the next range's `start` (or the end of
 * the document).
 */
export interface PageLabelRange {
  /** 0-based page index where this range begins. */
  start: number;
  style: PageLabelStyle;
  /** Optional textual prefix prepended to each label. */
  prefix?: string;
  /** Numeric value of the first page in the range (>= 1). PDF `/St`. */
  firstValue?: number;
}

function clampPrefix(value: unknown): string {
  if (typeof value !== "string") return "";
  // Strip ASCII control chars + DEL (a label is a single inline token) and cap.
  const cleaned = value.replace(/[\x00-\x1f\x7f]/g, "").trim();
  return cleaned.slice(0, PAGE_LABEL_PREFIX_MAX);
}

function clampFirstValue(value: unknown): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isInteger(n) || n < 1) return 1;
  // PDF page values are positive integers; cap so a silly value can't blow up
  // roman/alpha formatting.
  return Math.min(n, PAGE_LABEL_FIRST_VALUE_MAX);
}

/**
 * Convert a positive integer to an UPPERCASE Roman numeral. Values of 4000 and
 * up repeat "M" once per thousand (4000 = MMMM, 4001 = MMMMI) — exactly what
 * pdf.js `toRomanNumerals` renders for a written `/S /R` (or `/r`) label, so
 * the preview matches the saved PDF and distinct pages keep distinct labels.
 * <1 ⇒ "". A value no label range can reach (past
 * {@link PAGE_LABEL_FIRST_VALUE_MAX} + {@link PAGE_LABEL_MAX_PAGES}) also
 * yields "", which bounds the run of M's for direct callers.
 */
export function toRoman(value: number): string {
  let n = Math.trunc(value);
  if (!Number.isFinite(n) || n < 1) return "";
  if (n > ROMAN_VALUE_MAX) return "";
  const thousands = Math.floor(n / 1000);
  n -= thousands * 1000;
  const table: ReadonlyArray<readonly [number, string]> = [
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let out = "M".repeat(thousands);
  for (const [v, sym] of table) {
    while (n >= v) {
      out += sym;
      n -= v;
    }
  }
  return out;
}

/**
 * Convert a positive integer to a PDF alphabetic label (A, B, …, Z, AA, BB,
 * …, ZZ, AAA, BBB, …), uppercase. PDF `/A` repeats the selected letter for
 * each completed alphabet; it is deliberately NOT spreadsheet-style
 * bijective base-26 (`27 = AA`, but `28 = BB`, not `AB`). <1 ⇒ "".
 */
export function toAlpha(value: number): string {
  const n = Math.trunc(value);
  if (!Number.isFinite(n) || n < 1) return "";
  const letter = String.fromCharCode(65 + ((n - 1) % 26));
  return letter.repeat(Math.floor((n - 1) / 26) + 1);
}

const PAGE_LABEL_RANGE_KEYS = new Set([
  "start",
  "style",
  "prefix",
  "firstValue",
]);
const PAGE_LABEL_STYLE_SET = new Set<string>(PAGE_LABEL_STYLES);
const PAGE_LABEL_CONTROL_CHARS = /[\x00-\x1f\x7f]/;

/** Strict native prefix check: UTF-16 cap, no controls, valid scalar values. */
export function isValidPageLabelPrefix(value: string): boolean {
  if (
    value.length > PAGE_LABEL_PREFIX_MAX ||
    PAGE_LABEL_CONTROL_CHARS.test(value)
  ) {
    return false;
  }
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      index += 1;
    } else if (codeUnit >= 0xdc00 && codeUnit <= 0xdfff) {
      return false;
    }
  }
  return true;
}

/**
 * Strictly parse the native `pageLabelRanges` persistence value.
 *
 * This is intentionally separate from {@link sanitizePageLabelRanges}. The
 * latter is a forgiving helper for the standalone export dialog; native page
 * content is durable package authority and must never be repaired, defaulted,
 * reordered, de-duplicated, or truncated on its way into `content_json`.
 * Empty is valid and means "remove `/PageLabels`". Non-empty arrays must be in
 * canonical page order and start at physical page index 0.
 */
export function parsePageLabelRangesStrict(
  value: unknown,
  pageCount: number = PAGE_LABEL_MAX_PAGES,
): PageLabelRange[] | null {
  if (!Array.isArray(value)) return null;
  if (
    !Number.isInteger(pageCount) ||
    pageCount < 0 ||
    pageCount > PAGE_LABEL_MAX_PAGES ||
    value.length > pageCount
  ) {
    return null;
  }
  if (value.length === 0) return [];
  if (pageCount === 0) return null;

  const parsed: PageLabelRange[] = [];
  let previousStart = -1;
  for (const raw of value) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
    const record = raw as Record<string, unknown>;
    if (Object.keys(record).some((key) => !PAGE_LABEL_RANGE_KEYS.has(key))) {
      return null;
    }
    if (
      typeof record.start !== "number" ||
      !Number.isInteger(record.start) ||
      record.start < 0 ||
      record.start >= pageCount ||
      record.start <= previousStart
    ) {
      return null;
    }
    if (
      typeof record.style !== "string" ||
      !PAGE_LABEL_STYLE_SET.has(record.style)
    ) {
      return null;
    }
    if (
      record.prefix !== undefined &&
      (typeof record.prefix !== "string" ||
        !isValidPageLabelPrefix(record.prefix))
    ) {
      return null;
    }
    if (
      record.firstValue !== undefined &&
      (typeof record.firstValue !== "number" ||
        !Number.isInteger(record.firstValue) ||
        record.firstValue < 1 ||
        record.firstValue > PAGE_LABEL_FIRST_VALUE_MAX)
    ) {
      return null;
    }

    const range: PageLabelRange = {
      start: record.start,
      style: record.style as PageLabelStyle,
    };
    if (record.prefix !== undefined) range.prefix = record.prefix as string;
    if (record.firstValue !== undefined) {
      range.firstValue = record.firstValue as number;
    }
    parsed.push(range);
    previousStart = range.start;
  }
  return parsed[0]?.start === 0 ? parsed : null;
}

/**
 * Native `null` means the current package still owns an unprojected label tree.
 * Merely opening its empty editor draft must not authorize deletion; an
 * explicit Replace-all/Add/edit gesture establishes replacement intent first.
 * Array authorities (including explicit-empty) are already editor-owned.
 */
export function nativePageLabelReplacementReady(
  authority: readonly PageLabelRange[] | null,
  replacementStarted: boolean,
): boolean {
  return authority !== null || replacementStarted;
}

/** Exact semantic equality for canonical native range arrays. */
export function pageLabelRangesEqual(
  left: readonly PageLabelRange[],
  right: readonly PageLabelRange[],
): boolean {
  return (
    left.length === right.length &&
    left.every((range, index) => {
      const candidate = right[index];
      return (
        candidate !== undefined &&
        range.start === candidate.start &&
        range.style === candidate.style &&
        range.prefix === candidate.prefix &&
        range.firstValue === candidate.firstValue
      );
    })
  );
}

export type NativePageLabelTarget = readonly PageLabelRange[] | null;

/**
 * Compare a desired native authority with the strict content authority. Null
 * means the JSON key is absent and the immutable source owns `/PageLabels`;
 * an empty array remains the distinct explicit-delete authority.
 */
export function nativePageLabelTargetMatches(
  target: NativePageLabelTarget,
  authority:
    | { state: "absent" }
    | { state: "complete"; ranges: readonly PageLabelRange[] }
    | { state: "malformed" },
): boolean {
  if (target === null) return authority.state === "absent";
  return (
    authority.state === "complete" &&
    pageLabelRangesEqual(target, authority.ranges)
  );
}

/** Build the exact durable content target used by Apply, Undo, and Redo. */
export function applyNativePageLabelTargetToContent<
  T extends Record<string, unknown>,
>(content: T, target: NativePageLabelTarget): T {
  const next: Record<string, unknown> = { ...content };
  if (target === null) {
    delete next.pageLabelRanges;
  } else {
    next.pageLabelRanges = target.map((range) => ({ ...range }));
  }
  return next as T;
}

/**
 * A genuine load may retain only the exact repair currently leased to the
 * repair-owned reload and still authenticated against the loaded commit.
 */
export function retainNativePageLabelRepairOnLoad(options: {
  repairReloadOwned: boolean;
  exactCommitMatches: boolean;
}): boolean {
  return options.repairReloadOwned && options.exactCommitMatches;
}

/** Same-value Apply is safe only after native package materialization proof. */
export function canShortCircuitNativePageLabelApply(
  current: {
    ranges: readonly PageLabelRange[];
    materialized: boolean;
  },
  next: readonly PageLabelRange[],
): boolean {
  return current.materialized && pageLabelRangesEqual(current.ranges, next);
}

export type NativePageLabelApplyPlan = "save" | "noop" | "repair" | "blocked";

/**
 * Decide the native Apply lifecycle without collapsing persisted-but-unwritten
 * authority into a semantic no-op. A same-value unmaterialized authority may
 * only enter the bounded patchback-repair seam when this mounted editor still
 * holds its authenticated before/after repair record.
 */
export function planNativePageLabelApply(
  current: {
        ranges: readonly PageLabelRange[];
        materialized: boolean;
  } | null,
  next: readonly PageLabelRange[],
  hasBoundedRepair: boolean,
): NativePageLabelApplyPlan {
  if (!current) return "save";
  const sameValue = pageLabelRangesEqual(current.ranges, next);
  if (current.materialized) return sameValue ? "noop" : "save";
  if (!sameValue) return "save";
  return hasBoundedRepair ? "repair" : "blocked";
}

/**
 * Accept the one server-authenticated patchback skip that proves the requested
 * label value already exists in the source package. The caller separately
 * authenticates exact submitted/saved content equality. Generic unsupported
 * skips are deliberately excluded.
 */
export function isVerifiedPageLabelPatchbackNoopReceipt(
  receipt: {
        status: unknown;
        adapter: unknown;
        outputSha256: unknown;
        patchedAnchorCount: unknown;
        lossEntryCount: unknown;
        reason: unknown;
  } | null,
): boolean {
  return Boolean(
    receipt?.status === "skipped" &&
      typeof receipt.adapter === "string" &&
      receipt.adapter.toLowerCase() === "pdf" &&
      receipt.reason === "no_pdf_patches_applied" &&
      receipt.lossEntryCount === 0 &&
      receipt.patchedAnchorCount === 0 &&
      typeof receipt.outputSha256 === "string" &&
      /^[0-9a-f]{64}$/.test(receipt.outputSha256),
  );
}

export function buildNativePageLabelPatchbackRepairRequest<
  TContent extends Record<string, unknown>,
  TMarker,
>(repair: {
  content: TContent;
  previousContent: TContent;
  contentVersion: number;
  incompleteMarker: TMarker | null;
}): {
  content_json: TContent;
  base_content_version: number;
  editable_page_patchback_repair: {
    adapter: "pdf";
    previous_content_json: TContent;
    expected_incomplete_marker: TMarker;
  };
} | null {
  if (
    !Number.isSafeInteger(repair.contentVersion) ||
    repair.contentVersion < 1 ||
    !repair.incompleteMarker
  ) {
    return null;
  }
  return {
    content_json: repair.content,
    base_content_version: repair.contentVersion,
    editable_page_patchback_repair: {
      adapter: "pdf",
      previous_content_json: repair.previousContent,
      expected_incomplete_marker: repair.incompleteMarker,
    },
  };
}

/**
 * Authenticate the receipt returned by an explicit marker-bound page-label
 * repair request. `exactContentCommit` is supplied by the caller's durable
 * page/content identity check; this helper owns the remaining PDF receipt
 * boundary, including the one verified native no-op allowed by the backend.
 */
export function isAcceptedNativePageLabelRepairReceipt(options: {
  exactContentCommit: boolean;
  markerResourceId: string;
  disallowedResourceVersionIds: readonly string[];
  receipt: {
        status: unknown;
        adapter: unknown;
        resourceId: unknown;
        resourceVersionId: unknown;
        outputSha256: unknown;
        patchedAnchorCount: unknown;
        lossEntryCount: unknown;
        reason: unknown;
  } | null;
}): boolean {
  if (!options.exactContentCommit || !options.receipt) return false;
  if (options.receipt.resourceId !== options.markerResourceId) return false;
  if (isVerifiedPageLabelPatchbackNoopReceipt(options.receipt)) return true;
  const { receipt } = options;
  return Boolean(
    receipt.status === "succeeded" &&
      receipt.adapter === "pdf" &&
      receipt.resourceId === options.markerResourceId &&
      typeof receipt.resourceVersionId === "string" &&
      receipt.resourceVersionId.length > 0 &&
    !options.disallowedResourceVersionIds.includes(receipt.resourceVersionId) &&
      typeof receipt.outputSha256 === "string" &&
      /^[0-9a-f]{64}$/.test(receipt.outputSha256) &&
      receipt.lossEntryCount === 0 &&
      Number.isSafeInteger(receipt.patchedAnchorCount) &&
      Number(receipt.patchedAnchorCount) > 0,
  );
}

/** Format a numeric page VALUE in the given style (no prefix). */
export function formatLabelNumber(
  value: number,
  style: PageLabelStyle,
): string {
  switch (style) {
    case "decimal":
      return Number.isFinite(value) && value >= 1
        ? String(Math.trunc(value))
        : "";
    case "upper-roman":
      return toRoman(value);
    case "lower-roman":
      return toRoman(value).toLowerCase();
    case "upper-alpha":
      return toAlpha(value);
    case "lower-alpha":
      return toAlpha(value).toLowerCase();
    case "none":
      return "";
  }
}

/**
 * Validate + normalize raw ranges into a clean, de-duplicated, sorted set:
 *
 *   - drop entries whose `start` is not a non-negative integer,
 *   - clamp `start` below `pageCount`,
 *   - sort by `start`,
 *   - collapse duplicate starts (keep the LAST — later config wins),
 *   - ensure a range exists at index 0 (synthesize a decimal `start:0` range
 *     when the first range begins after page 0, so every page is covered).
 *
 * Returns at most `pageCount` ranges. An empty / non-array input with a positive
 * `pageCount` yields a single default decimal range.
 */
export function sanitizePageLabelRanges(
  ranges: unknown,
  pageCount: number,
): PageLabelRange[] {
  const count =
    Number.isInteger(pageCount) && pageCount > 0
      ? Math.min(pageCount, PAGE_LABEL_MAX_PAGES)
      : 0;
  if (count === 0) return [];
  const list = Array.isArray(ranges) ? ranges : [];
  const byStart = new Map<number, PageLabelRange>();
  for (const raw of list) {
    if (!raw || typeof raw !== "object") continue;
    const record = raw as Record<string, unknown>;
    const startRaw = record.start;
    if (typeof startRaw !== "number" || !Number.isInteger(startRaw)) continue;
    if (startRaw < 0 || startRaw >= count) continue;
    byStart.set(startRaw, {
      start: startRaw,
      style: asPageLabelStyle(record.style),
      prefix: clampPrefix(record.prefix) || undefined,
      firstValue: clampFirstValue(record.firstValue),
    });
  }
  // Always cover page 0.
  if (!byStart.has(0)) {
    byStart.set(0, { start: 0, style: "decimal", firstValue: 1 });
  }
  return [...byStart.values()].sort((a, b) => a.start - b.start);
}

/**
 * Compute the displayed label for EVERY page (index 0..pageCount-1) under a set
 * of ranges. Each page's label = the governing range's prefix + the numeral for
 * its position within the range (firstValue + offset). A "none"-style range
 * yields prefix-only (or "" when it has no prefix). Always returns exactly
 * `min(pageCount, MAX)` strings; an out-of-range / empty config falls back to
 * plain decimal "1".."N".
 */
export function computePageLabels(
  ranges: unknown,
  pageCount: number,
): string[] {
  const clean = sanitizePageLabelRanges(ranges, pageCount);
  const count =
    Number.isInteger(pageCount) && pageCount > 0
      ? Math.min(pageCount, PAGE_LABEL_MAX_PAGES)
      : 0;
  return computeCanonicalPageLabels(clean, count);
}

function computeCanonicalPageLabels(
  clean: readonly PageLabelRange[],
  count: number,
): string[] {
  const out: string[] = [];
  if (count === 0) return out;
  if (clean.length === 0) {
    return Array.from({ length: count }, (_, index) => String(index + 1));
  }
  // Find the governing range for each page by walking the sorted ranges.
  let rangeIdx = 0;
  for (let page = 0; page < count; page += 1) {
    // Advance to the last range whose start <= page.
    while (rangeIdx + 1 < clean.length && clean[rangeIdx + 1].start <= page) {
      rangeIdx += 1;
    }
    const range = clean[rangeIdx];
    const offset = page - range.start;
    const value = (range.firstValue ?? 1) + offset;
    const numeral = formatLabelNumber(value, range.style);
    out.push(`${range.prefix ?? ""}${numeral}`);
  }
  return out;
}

/**
 * Compute preview labels from strict native authority without applying the
 * standalone dialog's forgiving prefix trim/default/reorder behavior.
 */
export function computeNativePageLabels(
  ranges: unknown,
  pageCount: number,
): string[] | null {
  const canonical = parsePageLabelRangesStrict(ranges, pageCount);
  return canonical ? computeCanonicalPageLabels(canonical, pageCount) : null;
}

// --------------------------------------------------------------------------
// DOCUMENT-provided labels (pdf.js `getPageLabels()`) — the READ side of the
// `/PageLabels` tree, powering the toolbar page indicator ("iii (14/296)") and
// label-aware page-input navigation. Distinct from the range EDITOR above,
// which builds labels for export.
// --------------------------------------------------------------------------

/** Max characters SHOWN for a document-provided label in a tight spot (see
 *  {@link displayPageLabel}). The label itself is kept whole. */
export const PAGE_LABEL_DISPLAY_MAX = 48;




export const PAGE_LABEL_MAX_LENGTH = 1024;

/** A label shortened for display only; never compare or navigate by it. */
export function displayPageLabel(label: string): string {
  return label.length > PAGE_LABEL_DISPLAY_MAX
    ? `${label.slice(0, PAGE_LABEL_DISPLAY_MAX - 1)}…`
    : label;
}

/**
 * Sanitize the array from pdf.js `getPageLabels()` into per-page display
 * labels, or `null` when the document has no USEFUL labels:
 *
 *   - not an array, or length ≠ `pageCount` → null (a partial label set can't
 *     be trusted; mirrors pdf.js `PDFViewer#setPageLabels`),
 *   - every label identical to its physical 1-based number → null (trivial
 *     labels add nothing to the indicator; also mirrors pdf.js),
 *   - each label is control-stripped, trimmed, and length-capped; a non-string
 *     entry becomes "" (so indexing stays aligned to pages).
 */
export function sanitizeDocPageLabels(
  labels: unknown,
  pageCount: number,
): string[] | null {
  if (!Array.isArray(labels)) return null;
  if (!Number.isInteger(pageCount) || pageCount < 1) return null;
  if (pageCount > PAGE_LABEL_MAX_PAGES) return null;
  if (labels.length !== pageCount) return null;
  const out = labels.map((raw) =>
    typeof raw === "string"
      ? raw
          .replace(/[\x00-\x1f\x7f]/g, "")
          .trim()
          .slice(0, PAGE_LABEL_MAX_LENGTH)
      : "",
  );
  if (out.every((label, index) => label === String(index + 1))) return null;
  return out;
}

/**
 * Resolve the toolbar page-input text to a 1-based PHYSICAL page number:
 * a matching document label wins (exact first, then case-insensitive; first
 * occurrence on duplicates — matching Acrobat's label-aware page box), else
 * the input is parsed as a physical number (truncated; must be ≥ 1 — the
 * caller clamps overshoot to the last page like it always has). Returns null
 * when the input resolves to nothing (caller restores the previous value).
 */
export function resolvePageInput(
  input: string,
  labels: readonly string[] | null,
  pageCount: number,
): number | null {
  if (typeof input !== "string") return null;
  const text = input.trim();
  if (text.length === 0) return null;
  const count =
    Number.isInteger(pageCount) && pageCount > 0 ? pageCount : Infinity;
  if (labels && labels.length > 0) {
    const exact = labels.indexOf(text);
    if (exact >= 0 && exact < count) return exact + 1;
    const lower = text.toLowerCase();
    const loose = labels.findIndex(
      (label) => label.length > 0 && label.toLowerCase() === lower,
    );
    if (loose >= 0 && loose < count) return loose + 1;
  }
  const parsed = Number(text);
  if (!Number.isFinite(parsed) || parsed < 1) return null;
  return Math.trunc(parsed);
}

/** A plain descriptor for one `/PageLabels /Nums` entry (the writer maps it). */
export interface PageLabelNum {
  /** 0-based page index key. */
  index: number;
  /** PDF `/S` style code, or null for a prefix-only ("none") range. */
  styleCode: string | null;
  /** PDF `/P` prefix (omitted by the writer when ""). */
  prefix: string;
  /** PDF `/St` start value (omitted by the writer when 1). */
  firstValue: number;
}

/**
 * Build the `/PageLabels` number-tree `/Nums` descriptor from a set of ranges —
 * one entry per (sanitized) range, in page order. The viewer turns each entry
 * into the dict `<< /S /D /P (prefix) /St n >>` keyed by `index`. Pure: no
 * pdf-lib here, just the typed plan. Returns [] when there are no pages.
 */
export function buildPageLabelNums(
  ranges: unknown,
  pageCount: number,
): PageLabelNum[] {
  const clean = sanitizePageLabelRanges(ranges, pageCount);
  return clean.map((r) => ({
    index: r.start,
    styleCode: pdfStyleCode(r.style),
    prefix: r.prefix ?? "",
    firstValue: r.firstValue ?? 1,
  }));
}
