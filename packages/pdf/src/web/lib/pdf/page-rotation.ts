/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { validateRangeSpec } from "./range-spec";
import { normalizeRotation, type PdfRotation } from "@mrnote/core/web/lib/pdf/page-ops.ts";

export type PdfPageRotationScope = "page" | "all" | "odd" | "even" | "range";
export type PdfPageRotationDirection =
  "clockwise" | "counterclockwise" | "half-turn";

export interface PdfPageRotationOperation {
  op: "rotate_pdf_page";
  page: number;
  degrees: PdfRotation;
}

/** Resolve an absolute page rotation, honoring an explicit 0 override. */
export function resolveAbsolutePageRotation(
  page: number,
  pageRotations: Readonly<Record<number, number>> = {},
  sourceRotation: unknown = 0,
): PdfRotation {
  return normalizeRotation(
    Object.prototype.hasOwnProperty.call(pageRotations, page)
      ? pageRotations[page]
      : sourceRotation,
  );
}

/**
 * Resolve the concrete 1-based pages selected by the WPS-style rotation scope.
 * Invalid/blank range text returns an empty list so callers can visibly refuse
 * the operation instead of silently falling back to the current page.
 */
export function rotationTargetPages(
  scope: PdfPageRotationScope,
  currentPage: number,
  totalPages: number,
  rangeSpec = "",
): number[] {
  const total = Number.isInteger(totalPages) && totalPages > 0 ? totalPages : 0;
  if (total === 0) return [];
  if (scope === "all") {
    return Array.from({ length: total }, (_, index) => index + 1);
  }
  if (scope === "odd" || scope === "even") {
    const parity = scope === "odd" ? 1 : 0;
    return Array.from({ length: total }, (_, index) => index + 1).filter(
      (page) => page % 2 === parity,
    );
  }
  if (scope === "range") {
    // Rotate exactly the pages named; page 0 or a page past the end is an
    
    return validateRangeSpec(rangeSpec, total).pages;
  }
  const current = Math.min(
    total,
    Math.max(1, Number.isInteger(currentPage) ? currentPage : 1),
  );
  return [current];
}

/**
 * Build one atomic native-document command batch. Persisted page transforms
 * are absolute `/Rotate` values; when a page has no override, its source PDF
 * rotation is the starting point.
 */
export function buildPageRotationOperations(options: {
  pages: ReadonlyArray<number>;
  direction: PdfPageRotationDirection;
  pageRotations?: Readonly<Record<number, number>>;
  sourceRotations?: Readonly<Record<number, number>>;
}): PdfPageRotationOperation[] {
  const {
    pages,
    direction,
    pageRotations = {},
    sourceRotations = {},
  } = options;
  const delta =
    direction === "counterclockwise"
      ? -90
      : direction === "half-turn"
        ? 180
        : 90;
  const seen = new Set<number>();
  const operations: PdfPageRotationOperation[] = [];

  for (const page of pages) {
    if (!Number.isInteger(page) || page < 1 || seen.has(page)) continue;
    seen.add(page);
    const current = resolveAbsolutePageRotation(
      page,
      pageRotations,
      sourceRotations[page],
    );
    operations.push({
      op: "rotate_pdf_page",
      page,
      degrees: normalizeRotation(current + delta),
    });
  }

  return operations;
}
