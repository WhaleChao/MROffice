/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * WPS 分割页面 in the editor: the requests it sends the native page split
 * writer (planned on the server's engine host) and the line arithmetic of
 * the on-page split mode. Browser-safe: the planner's types only.
 */
import type {
  PdfNativePageSplitEvidence,
  PdfNativePageSplitGrid,
  PdfNativePageSplitPageRequest,
  PdfNativePageSplitRequest,
} from "./engine/native/content-stream-transaction";
import {
  PDF_NATIVE_PAGE_SPLIT_MAX_LINES,
  PDF_NATIVE_PAGE_SPLIT_MAX_PAGES,
  PDF_NATIVE_PAGE_SPLIT_MAX_PARTS,
} from "./engine/native/vocabulary";
import { validateRangeSpec } from "./range-spec";

export type { PdfNativePageSplitEvidence, PdfNativePageSplitGrid };

/** What the split mode asks the host to write: the pages, each with its own
 * lines, and whether the original pages stay (WPS 保留原页面). */
export interface PdfSplitPagesIntent {
  readonly keepOriginal: boolean;
  readonly pages: readonly PdfNativePageSplitPageRequest[];
}

/** The merge that undoes a split, as the editor asks for it. */
export function pageSplitMergeRequest(
  split: PdfNativePageSplitEvidence,
  expectedSourceSha256: string,
): PdfNativePageSplitRequest {
  return Object.freeze({
    operation: "merge" as const,
    splitId: split.splitId,
    expectedSourceSha256,
    keepOriginal: split.keepOriginal,
    groups: Object.freeze(
      split.groups.map((group) =>
        Object.freeze({
          page: Object.freeze({
            objectNumber: group.page.objectNumber,
            generation: group.page.generation,
          }),
          parts: Object.freeze(
            (split.keepOriginal ? group.parts : group.parts.slice(1)).map((part) =>
              Object.freeze({
                objectNumber: part.objectNumber,
                generation: part.generation,
              }),
            ),
          ),
          restoreBoxes: Object.freeze({ ...group.originalBoxes }),
          restoreMarkers: Object.freeze({ ...group.originalMarkers }),
        }),
      ),
    ),
  });
}

/** The narrowest part a line may leave, as a fraction of the page. */
export const PDF_PAGE_SPLIT_MIN_GAP = 0.02;

/** Lines spread evenly: `count` lines cut the page into count + 1 parts. */
export function evenPageSplitLines(count: number): readonly number[] {
  const lines = Math.max(0, Math.min(PDF_NATIVE_PAGE_SPLIT_MAX_LINES, Math.floor(count)));
  return Object.freeze(
    Array.from({ length: lines }, (_, index) => roundPageSplitFraction((index + 1) / (lines + 1))),
  );
}

/** A line every `spacing` along a side of `length` (same unit): WPS 间隔. */
export function spacedPageSplitLines(length: number, spacing: number): readonly number[] {
  if (!(length > 0) || !(spacing > 0)) return Object.freeze([]);
  const lines: number[] = [];
  for (let at = spacing; at < length - 1e-9; at += spacing) {
    const fraction = roundPageSplitFraction(at / length);
    if (fraction >= PDF_PAGE_SPLIT_MIN_GAP && 1 - fraction >= PDF_PAGE_SPLIT_MIN_GAP) {
      lines.push(fraction);
    }
    if (lines.length >= PDF_NATIVE_PAGE_SPLIT_MAX_LINES) break;
  }
  return Object.freeze(lines);
}

export function roundPageSplitFraction(value: number): number {
  return Math.round(value * 10_000) / 10_000;
}

/** Lines sorted, inside the page and at least the minimum gap apart. */
export function normalizePageSplitLines(lines: readonly number[]): readonly number[] {
  const sorted = lines
    .filter((line) => Number.isFinite(line))
    .map(roundPageSplitFraction)
    .sort((left, right) => left - right);
  const kept: number[] = [];
  for (const line of sorted) {
    const previous = kept.length ? kept[kept.length - 1] : 0;
    if (line - previous < PDF_PAGE_SPLIT_MIN_GAP || 1 - line < PDF_PAGE_SPLIT_MIN_GAP) continue;
    kept.push(line);
    if (kept.length >= PDF_NATIVE_PAGE_SPLIT_MAX_LINES) break;
  }
  return Object.freeze(kept);
}

/** The parts a grid cuts one page into. */
export function pageSplitPartCount(grid: PdfNativePageSplitGrid): number {
  return (grid.rows.length + 1) * (grid.columns.length + 1);
}

export function pageSplitGridIsWritable(grid: PdfNativePageSplitGrid): boolean {
  const parts = pageSplitPartCount(grid);
  return parts >= 2 && parts <= PDF_NATIVE_PAGE_SPLIT_MAX_PARTS;
}

export type PdfPageSplitParity = "all" | "odd" | "even";

/**
 * The pages a WPS 页面范围 names: every page, or a list such as `1,3,5-9`
 * (read strictly, as page commands read a typed range), kept to odd or
 * even pages when asked. Null when the list cannot be read.
 */
export function pageSplitRangePages(
  pageCount: number,
  range: Readonly<{ mode: "all" } | { mode: "custom"; text: string }>,
  parity: PdfPageSplitParity,
): readonly number[] | null {
  let pages: number[];
  if (range.mode === "all") {
    pages = Array.from({ length: pageCount }, (_, index) => index + 1);
  } else {
    const text = range.text.replace(/[，、；;]/g, ",").trim();
    if (!text) return null;
    const read = validateRangeSpec(text, pageCount);
    if (read.issue || read.pages.length === 0) return null;
    pages = read.pages;
  }
  if (parity !== "all") {
    pages = pages.filter((page) => (page % 2 === 1) === (parity === "odd"));
  }
  return Object.freeze(pages);
}

export { PDF_NATIVE_PAGE_SPLIT_MAX_LINES, PDF_NATIVE_PAGE_SPLIT_MAX_PAGES, PDF_NATIVE_PAGE_SPLIT_MAX_PARTS };
