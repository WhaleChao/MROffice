/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * OCR acceptance (WPS PDF ▸ 文字识别).
 *
 * Recognising a scan is the easy half. The half that decides whether the
 * feature is trustworthy is what happens to a word the engine is unsure about:
 * dropped into the invisible text layer anyway, it makes search return the
 * wrong page and copy produce nonsense that looks authoritative. So this module
 * decides what may enter the text layer, what must be shown to a human first,
 * and when a page's recognition is too weak to claim at all.
 *
 * Pure: recognised words in, an acceptance decision out.
 */

export interface OcrWord {
  text: string;
  /** 0…1 from the engine. */
  confidence: number;
  /** Page-relative box, used to catch a word placed where nothing was drawn. */
  box?: { x: number; y: number; width: number; height: number };
}

export interface OcrPageResult {
  /** 1-based. */
  page: number;
  words: readonly OcrWord[];
  /** True when the page already carried real text (a digital, non-scanned page). */
  hadTextLayer?: boolean;
}

export interface OcrAcceptanceOptions {
  /** Below this a word is held back for review. Default 0.75. */
  reviewBelow?: number;
  /** Below this a word is discarded outright. Default 0.4. */
  discardBelow?: number;
  /**
   * A page whose accepted words fall below this share of its words is reported
   * as unreliable rather than silently half-recognised. Default 0.6.
   */
  minimumPageCoverage?: number;
}

export type OcrPageVerdict =
  "accepted" | "needs-review" | "unreliable" | "skipped-has-text";

export interface OcrPageAcceptance {
  page: number;
  verdict: OcrPageVerdict;
  /** Words that may go into the invisible text layer. */
  accepted: OcrWord[];
  /** Words a human should confirm before they become searchable. */
  review: OcrWord[];
  /** Words dropped as noise. */
  discarded: OcrWord[];
  /** Accepted share of the page's words, 0…1. */
  coverage: number;
}

export const DEFAULT_REVIEW_BELOW = 0.75;
export const DEFAULT_DISCARD_BELOW = 0.4;
export const DEFAULT_MINIMUM_COVERAGE = 0.6;

function isPlausibleBox(word: OcrWord): boolean {
  if (!word.box) return true;
  const { x, y, width, height } = word.box;
  return (
    [x, y, width, height].every((value) => Number.isFinite(value)) &&
    width > 0 &&
    height > 0 &&
    x >= -0.001 &&
    y >= -0.001 &&
    x + width <= 1.001 &&
    y + height <= 1.001
  );
}

/** Decide what a page's recognition may contribute to the document. */
export function acceptOcrPage(
  result: OcrPageResult,
  options: OcrAcceptanceOptions = {},
): OcrPageAcceptance {
  const reviewBelow = options.reviewBelow ?? DEFAULT_REVIEW_BELOW;
  const discardBelow = options.discardBelow ?? DEFAULT_DISCARD_BELOW;
  const minimumCoverage =
    options.minimumPageCoverage ?? DEFAULT_MINIMUM_COVERAGE;

  // A page that already has real text must not get a second, guessed layer:
  // search would find both and copy would duplicate.
  if (result.hadTextLayer) {
    return {
      page: result.page,
      verdict: "skipped-has-text",
      accepted: [],
      review: [],
      discarded: [],
      coverage: 0,
    };
  }

  const accepted: OcrWord[] = [];
  const review: OcrWord[] = [];
  const discarded: OcrWord[] = [];
  // Two kinds of discard, and only one of them says anything about how well the
  // page was read: a word the engine read POORLY is evidence of bad
  // recognition, while an empty fragment or a word boxed off the page is
  // structural junk the engine emitted regardless of the scan's quality.
  let poorlyRead = 0;

  for (const word of result.words) {
    const text = word.text.trim();
    const confidence = Number(word.confidence);
    // The field is declared 0…1, and a number outside it is not a confidence
    // this module can read. An engine reporting percentages sent 95, 12 and 3;
    // every one of them cleared the review threshold, so a page of garbage came
    // back "accepted" with the garbage in its text layer. Refusing the word is
    // the safe direction: the page loses coverage and says it was read badly.
    if (
      !text ||
      !Number.isFinite(confidence) ||
      confidence < 0 ||
      confidence > 1 ||
      !isPlausibleBox(word)
    ) {
      discarded.push(word);
      continue;
    }
    if (confidence < discardBelow) {
      discarded.push(word);
      poorlyRead += 1;
    } else if (confidence < reviewBelow) review.push(word);
    else accepted.push(word);
  }

  // Coverage is measured against words the engine actually READ — accepted,
  // held for review, or read too poorly to use. Structural junk is excluded so
  // it cannot drag a cleanly recognised page under the floor.
  const countable = accepted.length + review.length + poorlyRead;
  const coverage = countable === 0 ? 0 : accepted.length / countable;
  const total = countable;
  const verdict: OcrPageVerdict =
    total === 0 || coverage < minimumCoverage
      ? "unreliable"
      : review.length > 0
        ? "needs-review"
        : "accepted";

  return { page: result.page, verdict, accepted, review, discarded, coverage };
}

/**
 * The text layer a page contributes once its verdict is honoured — which means
 * nothing at all from a page the verdict rejected.
 *
 * It used to return the accepted words whatever the verdict said, so a page
 * ruled `unreliable` still handed over its best-read words while `ocrSummary`
 * told the user that page was "recognised too poorly to add". Two functions in
 * this module disagreed about the same page, and the one the text layer is
 * built from was the one ignoring the ruling.
 */
export function searchableText(acceptance: OcrPageAcceptance): string {
  if (
    acceptance.verdict === "unreliable" ||
    acceptance.verdict === "skipped-has-text"
  ) {
    return "";
  }
  return acceptance.accepted.map((word) => word.text.trim()).join(" ");
}

/** One line for the panel: what the user should do about this run. */
export function ocrSummary(
  pages: readonly OcrPageAcceptance[],
  locale: "en" | "zh",
): string {
  const review = pages.filter((page) => page.verdict === "needs-review").length;
  const unreliable = pages.filter(
    (page) => page.verdict === "unreliable",
  ).length;
  if (locale === "zh") {
    if (unreliable > 0) return `${unreliable} 页识别质量不足,未加入文字层。`;
    if (review > 0) return `${review} 页有低置信度文字,待确认。`;
    return "识别结果已全部加入文字层。";
  }
  if (unreliable > 0) {
    return `${unreliable} page(s) recognised too poorly to add.`;
  }
  if (review > 0) return `${review} page(s) have words to confirm.`;
  return "Every recognised word was added to the text layer.";
}

// ---------------------------------------------------------------------------
// Source-bound provider-result review
// ---------------------------------------------------------------------------

export type PdfOcrReviewStatus =
  | "pending_review"
  | "accepted"
  | "rejected";

export interface PdfOcrReviewRun {
  readonly text: string;
  readonly bbox: readonly [number, number, number, number];
}

export interface PdfOcrReviewAnnotation {
  readonly id: string;
  readonly pageIndex: number;
  readonly sourceSha256: string;
  readonly status: PdfOcrReviewStatus;
  readonly runs: readonly PdfOcrReviewRun[];
  readonly payloadSha256: string;
}

export interface PdfOcrManualReviewDecision {
  readonly verdict: "manual-review-required";
  readonly confidenceAvailable: false;
  readonly accepted: readonly [];
  readonly review: readonly PdfOcrReviewRun[];
}

type JsonRecord = Record<string, unknown>;

const PDF_OCR_REVIEW_SCHEMA = 2;
const PDF_OCR_SHA256 = /^[0-9a-f]{64}$/;
const PDF_OCR_MAX_RUNS = 500;
const PDF_OCR_MAX_TEXT = 1_000;

function jsonRecord(value: unknown): JsonRecord | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : null;
}

function exactJsonRecord(
  value: unknown,
  keys: readonly string[],
): JsonRecord | null {
  const record = jsonRecord(value);
  if (!record) return null;
  const actual = Object.keys(record).sort();
  const expected = [...keys].sort();
  return actual.length === expected.length &&
    actual.every((key, index) => key === expected[index])
    ? record
    : null;
}

function parsePdfOcrReviewRuns(value: unknown): readonly PdfOcrReviewRun[] | null {
  if (!Array.isArray(value) || value.length > PDF_OCR_MAX_RUNS) return null;
  const runs: PdfOcrReviewRun[] = [];
  for (const candidate of value) {
    const run = exactJsonRecord(candidate, ["text", "bbox"]);
    const text = run?.text;
    const bbox = run?.bbox;
    if (
      typeof text !== "string" ||
      !text ||
      text.trim() !== text ||
      text.length > PDF_OCR_MAX_TEXT ||
      !Array.isArray(bbox) ||
      bbox.length !== 4 ||
      !bbox.every(
        (coordinate) =>
          typeof coordinate === "number" &&
          Number.isFinite(coordinate) &&
          coordinate >= 0 &&
          coordinate <= 1,
      ) ||
      bbox[2] <= bbox[0] ||
      bbox[3] <= bbox[1]
    ) {
      return null;
    }
    runs.push(
      Object.freeze({
        text,
        bbox: Object.freeze([
          bbox[0],
          bbox[1],
          bbox[2],
          bbox[3],
        ]) as readonly [number, number, number, number],
      }),
    );
  }
  return Object.freeze(runs);
}

function parsePdfOcrReviewPayload(
  value: unknown,
  expectedSourceSha256: string,
): Omit<PdfOcrReviewAnnotation, "id" | "payloadSha256"> | null {
  const payload = exactJsonRecord(value, [
    "schema",
    "source_sha256",
    "page_index",
    "status",
    "runs",
  ]);
  if (
    !payload ||
    payload.schema !== PDF_OCR_REVIEW_SCHEMA ||
    payload.source_sha256 !== expectedSourceSha256 ||
    !PDF_OCR_SHA256.test(expectedSourceSha256) ||
    typeof payload.page_index !== "number" ||
    !Number.isInteger(payload.page_index) ||
    payload.page_index < 0 ||
    payload.page_index > 9_999 ||
    (payload.status !== "pending_review" &&
      payload.status !== "accepted" &&
      payload.status !== "rejected")
  ) {
    return null;
  }
  const runs = parsePdfOcrReviewRuns(payload.runs);
  if (!runs || (payload.status === "rejected" && runs.length > 0)) return null;
  return Object.freeze({
    pageIndex: payload.page_index,
    sourceSha256: expectedSourceSha256,
    status: payload.status,
    runs,
  });
}

/** Strict adapter for the generic annotation GET/PATCH response. */
export function parsePdfOcrReviewAnnotation(
  value: unknown,
  expectedSourceSha256: string,
): PdfOcrReviewAnnotation | null {
  const row = jsonRecord(value);
  if (
    !row ||
    typeof row.id !== "string" ||
    !row.id ||
    row.id.trim() !== row.id ||
    row.id.length > 200 ||
    row.type !== "pdf_ocr_text" ||
    typeof row.payload_sha256 !== "string" ||
    !PDF_OCR_SHA256.test(row.payload_sha256)
  ) {
    return null;
  }
  const payload = parsePdfOcrReviewPayload(
    row.payload_json,
    expectedSourceSha256,
  );
  return payload
    ? Object.freeze({
        id: row.id,
        payloadSha256: row.payload_sha256,
        ...payload,
      })
    : null;
}

/**
 * The source a well-formed schema-2 review row is bound to, whatever source
 * the reader holds: a row for other bytes (an earlier version of an edited
 * page) is recognized as such instead of read as malformed. Null for any row
 * the strict adapter would refuse.
 */
export function pdfOcrReviewAnnotationSourceSha256(
  value: unknown,
): string | null {
  const payload = jsonRecord(jsonRecord(value)?.payload_json);
  const sourceSha256 = payload?.source_sha256;
  return typeof sourceSha256 === "string" &&
    PDF_OCR_SHA256.test(sourceSha256) &&
    parsePdfOcrReviewAnnotation(value, sourceSha256) !== null
    ? sourceSha256
    : null;
}

/** Strict adapter for `/pdf-ocr-page`; unknown response fields fail closed. */
export function parsePdfOcrPageReviewResult(
  value: unknown,
  expected: Readonly<{
    sourceSha256: string;
    pageIndex: number;
  }>,
): PdfOcrReviewAnnotation | null {
  const result = exactJsonRecord(value, [
    "runs",
    "cached",
    "schema",
    "source_sha256",
    "page_index",
    "status",
    "annotation_id",
    "payload_sha256",
  ]);
  if (
    !result ||
    typeof result.cached !== "boolean" ||
    result.schema !== PDF_OCR_REVIEW_SCHEMA ||
    result.page_index !== expected.pageIndex
  ) {
    return null;
  }
  return parsePdfOcrReviewAnnotation(
    {
      id: result.annotation_id,
      type: "pdf_ocr_text",
      payload_sha256: result.payload_sha256,
      payload_json: {
        schema: result.schema,
        source_sha256: result.source_sha256,
        page_index: result.page_index,
        status: result.status,
        runs: result.runs,
      },
    },
    expected.sourceSha256,
  );
}

/**
 * This provider does not expose trustworthy per-run confidence. Every valid
 * run therefore starts in human review; none is silently promoted by a made-up
 * threshold.
 */
export function requireManualReviewForOcrRuns(
  runs: readonly PdfOcrReviewRun[],
): PdfOcrManualReviewDecision {
  return Object.freeze({
    verdict: "manual-review-required",
    confidenceAvailable: false,
    accepted: Object.freeze([]),
    review: Object.freeze([...runs]),
  } satisfies PdfOcrManualReviewDecision);
}

export function buildPdfOcrReviewPayload(
  annotation: PdfOcrReviewAnnotation,
  status: "accepted" | "rejected",
  runs: readonly PdfOcrReviewRun[],
): JsonRecord | null {
  const cleanRuns = parsePdfOcrReviewRuns(status === "rejected" ? [] : runs);
  if (!cleanRuns) return null;
  return {
    schema: PDF_OCR_REVIEW_SCHEMA,
    source_sha256: annotation.sourceSha256,
    page_index: annotation.pageIndex,
    status,
    runs: cleanRuns.map((run) => ({
      text: run.text,
      bbox: [...run.bbox],
    })),
  };
}
