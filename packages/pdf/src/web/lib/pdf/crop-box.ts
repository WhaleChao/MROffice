/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for PDF page CROP (Acrobat "Crop Pages" / WPS 去白边 parity).
 *
 * A crop is the visible RECTANGLE OF THE PAGE TO KEEP, captured in page-relative
 * *percentage* coordinates [0..100] with a TOP-LEFT origin (resolution-
 * independent, matching how highlight rects / sticky notes / ink / redaction
 * boxes / shapes are stored). Cropping a page sets its PDF `CropBox` to this
 * region on export: everything outside the rectangle is clipped away (the white
 * margins disappear) without deleting any underlying content — re-opening the
 * file and clearing the crop restores the full page.
 *
 * This module mirrors `ink.ts` / `shapes.ts` (per-page geometry): it holds
 * everything with no DOM / React / pdf-lib dependency so it can be unit-tested in
 * isolation and reused by BOTH the editor UI and the flatten-export planner:
 *
 *   - the persisted crop shape + its sanitizer ({@link sanitizeCropBox}),
 *   - construction from edge MARGINS ({@link cropBoxFromMargins}) — the natural
 *     "trim N% off each edge" gesture — and the inverse ({@link marginsForCropBox}),
 *   - a normalizer ({@link normalizeCropKeepRect}) that clamps the keep-rect onto
 *     the page and enforces a minimum size so a crop can never collapse the page,
 *   - {@link isFullPageCrop} so an effectively-no-op crop is dropped (no CropBox
 *     op emitted) and the persisted state stays minimal.
 *
 * The keep-rect is a {@link CropKeepRect} (top-left origin {x,y,width,height} in
 * percent), so it feeds straight into the same `percentRectToUserSpace` mapping
 * the flatten planner already uses for every other percent input — including
 * rotation handling — and onto a `setCropBox(x, y, w, h)` call in user space.
 *
 * The persisted field is `cropBoxes` on the PDF page content and is additive:
 * older PDFs omit it (no page is cropped). Each entry is keyed by 1-based page.
 *
 * A crop can be applied to MORE than the current page (Acrobat "Crop Pages →
 * Page Range: All / From-To" / WPS 应用到所有页面): {@link cropTargetPages}
 * turns a {@link CropScope} + the current page + page count (+ an optional
 * Acrobat-style range spec) into the concrete 1-based page list, and
 * {@link setCropForPages} bakes one keep-rect across that whole list at once.
 */

import { parseRangeSpec } from "./range-spec";

/** A page-relative percent rect (top-left origin) describing the region to KEEP. */
export interface CropKeepRect {
  /** Left edge as a percent of page width [0..100). */
  x: number;
  /** Top edge as a percent of page height [0..100). */
  y: number;
  /** Width as a percent of page width (0, 100]. */
  width: number;
  /** Height as a percent of page height (0, 100]. */
  height: number;
}

/** A persisted per-page crop: the 1-based page + the keep-rect to bake. */
export interface CropBoxEntry extends CropKeepRect {
  /** 1-based page number this crop applies to. */
  page: number;
}

/** Edge margins (percent) trimmed from each side of the page. */
export interface CropMargins {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * Smallest keep-rect dimension, in percent. A crop can never shrink the page
 * below this on either axis (so a stray drag / crafted payload can't collapse a
 * page to a sliver or to nothing). 10% of the page is still a very aggressive
 * crop while staying safe.
 */
export const CROP_MIN_PERCENT = 10;

/** The full, uncropped page keep-rect (the whole page). */
export const FULL_PAGE_CROP: CropKeepRect = { x: 0, y: 0, width: 100, height: 100 };

/** Tolerance (percent) within which a crop is treated as "the full page". */
const FULL_PAGE_EPSILON = 0.01;

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

/** Clamp a value into [lo, hi]; non-finite -> lo. */
function clampTo(value: number, lo: number, hi: number): number {
  if (!Number.isFinite(value)) return lo;
  if (value < lo) return lo;
  if (value > hi) return hi;
  return value;
}

/**
 * Normalize an arbitrary keep-rect onto the page: clamp the origin into
 * [0,100], clamp the size so the rect stays within the page AND is at least
 * {@link CROP_MIN_PERCENT} on each axis, nudging the origin inward if a clamped
 * minimum-size rect would otherwise overflow the right/bottom edge. Always
 * returns a valid on-page rect (never null) so callers can rely on it.
 */
export function normalizeCropKeepRect(rect: {
  x: unknown;
  y: unknown;
  width: unknown;
  height: unknown;
}): CropKeepRect {
  // Origin first (a sane starting corner), leaving room for the minimum size.
  let x = clampTo(isFiniteNumber(rect.x) ? rect.x : 0, 0, 100 - CROP_MIN_PERCENT);
  let y = clampTo(isFiniteNumber(rect.y) ? rect.y : 0, 0, 100 - CROP_MIN_PERCENT);
  // Size: at least the minimum, at most what fits from the (clamped) origin.
  let width = clampTo(
    isFiniteNumber(rect.width) ? rect.width : 100,
    CROP_MIN_PERCENT,
    100 - x,
  );
  let height = clampTo(
    isFiniteNumber(rect.height) ? rect.height : 100,
    CROP_MIN_PERCENT,
    100 - y,
  );
  // If the origin sat so far right/down that even the minimum size overflows,
  // pull the origin back so the minimum-size rect fits on the page.
  if (x + width > 100) x = 100 - width;
  if (y + height > 100) y = 100 - height;
  width = clampTo(width, CROP_MIN_PERCENT, 100 - x);
  height = clampTo(height, CROP_MIN_PERCENT, 100 - y);
  return { x, y, width, height };
}

/**
 * Build a keep-rect from edge MARGINS (the "trim N% off each side" gesture).
 * Negative / non-finite margins are treated as 0; the result is normalized so
 * over-aggressive margins (left+right >= 100) can't invert or collapse the page.
 */
export function cropBoxFromMargins(margins: {
  left?: unknown;
  top?: unknown;
  right?: unknown;
  bottom?: unknown;
}): CropKeepRect {
  const left = Math.max(0, isFiniteNumber(margins.left) ? margins.left : 0);
  const top = Math.max(0, isFiniteNumber(margins.top) ? margins.top : 0);
  const right = Math.max(0, isFiniteNumber(margins.right) ? margins.right : 0);
  const bottom = Math.max(0, isFiniteNumber(margins.bottom) ? margins.bottom : 0);
  return normalizeCropKeepRect({
    x: left,
    y: top,
    width: 100 - left - right,
    height: 100 - top - bottom,
  });
}

/** The inverse of {@link cropBoxFromMargins}: edge margins for a keep-rect. */
export function marginsForCropBox(rect: CropKeepRect): CropMargins {
  const left = clampTo(rect.x, 0, 100);
  const top = clampTo(rect.y, 0, 100);
  // `100 - (x + width)` carries binary float noise (7.299999999999997 for a
  // 7.3 margin); settle it far below any geometry the page can express so an
  
  const right = settleFloat(clampTo(100 - (rect.x + rect.width), 0, 100));
  const bottom = settleFloat(clampTo(100 - (rect.y + rect.height), 0, 100));
  return { left, top, right, bottom };
}












export function uncroppedPageSize(
  croppedSize: { width: number; height: number },
  appliedCrop: CropKeepRect,
): { width: number; height: number } {
  const kept = normalizeCropKeepRect(appliedCrop);
  return {
    width: settleFloat(croppedSize.width / (kept.width / 100)),
    height: settleFloat(croppedSize.height / (kept.height / 100)),
  };
}

/** Round to 1e-9 percent, removing float noise without moving the geometry. */
function settleFloat(value: number): number {
  return Math.round(value * 1e9) / 1e9;
}

/**
 * True when the keep-rect is (within a tiny epsilon) the whole page — i.e. the
 * crop is a no-op. Used to drop an empty crop from persistence and to skip
 * emitting a CropBox op on export.
 */
export function isFullPageCrop(rect: CropKeepRect): boolean {
  return (
    Math.abs(rect.x) <= FULL_PAGE_EPSILON &&
    Math.abs(rect.y) <= FULL_PAGE_EPSILON &&
    Math.abs(rect.width - 100) <= FULL_PAGE_EPSILON &&
    Math.abs(rect.height - 100) <= FULL_PAGE_EPSILON
  );
}

/**
 * Validate + normalize a persisted crop entry into a usable {@link CropBoxEntry},
 * or return null when it is structurally invalid (no 1-based page) OR effectively
 * the full page (so a no-op crop is never persisted / baked). Mirrors the other
 * sanitizers: out-of-range geometry is clamped (never thrown) onto the page.
 */
export function sanitizeCropBox(
  payload: Record<string, unknown> | null | undefined,
): CropBoxEntry | null {
  if (!payload || typeof payload !== "object") return null;
  const page = Number((payload as Record<string, unknown>).page);
  if (!Number.isInteger(page) || page < 1 || page > 100_000) return null;
  const rect = normalizeCropKeepRect({
    x: (payload as Record<string, unknown>).x,
    y: (payload as Record<string, unknown>).y,
    width: (payload as Record<string, unknown>).width,
    height: (payload as Record<string, unknown>).height,
  });
  if (isFullPageCrop(rect)) return null;
  return { page, ...rect };
}

/**
 * Upsert a crop for `page` into an existing list, replacing any prior entry for
 * the same page. A full-page (no-op) crop REMOVES the page's entry instead (so
 * "reset crop" clears it). Returns a new array; the input is left untouched.
 */
export function setCropForPage(
  entries: ReadonlyArray<CropBoxEntry>,
  page: number,
  rect: CropKeepRect,
): CropBoxEntry[] {
  const others = entries.filter((e) => e.page !== page);
  const normalized = normalizeCropKeepRect(rect);
  if (!Number.isInteger(page) || page < 1 || isFullPageCrop(normalized)) {
    return others;
  }
  return [...others, { page, ...normalized }];
}

/** Read the keep-rect for a 1-based page, or the full page when uncropped. */
export function cropForPage(
  entries: ReadonlyArray<CropBoxEntry>,
  page: number,
): CropKeepRect {
  const found = entries.find((e) => e.page === page);
  return found
    ? { x: found.x, y: found.y, width: found.width, height: found.height }
    : { ...FULL_PAGE_CROP };
}

/**
 * Which pages a crop is applied to (Acrobat "Crop Pages → Page Range"):
 *   - `"page"`  — just the current page (the historical default),
 *   - `"all"`   — every page of the document,
 *   - `"range"` — an Acrobat-style 1-based range spec (`"1-3, 5, 8-"`).
 */
export type CropScope = "page" | "all" | "range";

/**
 * Resolve a {@link CropScope} to the concrete, sorted, de-duped list of 1-based
 * page numbers the crop should bake onto.
 *
 *   - `"all"`   → every page `1..totalPages`.
 *   - `"range"` → the pages an Acrobat-style spec selects (parsed/clamped by the
 *     shared {@link parseRangeSpec}); a blank / all-garbage spec selects nothing.
 *   - `"page"` (or any unknown scope) → just `currentPage`, when it is a valid
 *     in-document page.
 *
 * `totalPages` is clamped to a sane non-negative integer; a `currentPage`
 * outside `[1, totalPages]` is dropped (so a stale page never yields an
 * off-document target). Always returns a fresh array (possibly empty), so the
 * caller can treat "nothing selected" as a no-op rather than crash.
 */
export function cropTargetPages(
  scope: CropScope,
  currentPage: number,
  totalPages: number,
  rangeSpec?: string,
): number[] {
  const total =
    Number.isFinite(totalPages) && totalPages > 0 ? Math.trunc(totalPages) : 0;
  if (total < 1) return [];
  if (scope === "all") {
    const pages: number[] = [];
    for (let p = 1; p <= total; p += 1) pages.push(p);
    return pages;
  }
  if (scope === "range") {
    return parseRangeSpec(typeof rangeSpec === "string" ? rangeSpec : "", total);
  }
  // "page" (default): the current page, if it is on the document.
  const page = Number.isFinite(currentPage) ? Math.trunc(currentPage) : 0;
  if (page < 1 || page > total) return [];
  return [page];
}

/**
 * Upsert one keep-rect across MANY pages at once, replacing any prior entry for
 * each target page. A full-page (no-op) rect REMOVES each target's entry instead
 * (so "reset crop" can clear a whole range). Non-integer / non-positive target
 * pages are ignored. Returns a new array; the input is left untouched. Mirrors
 * {@link setCropForPage} applied across the list, but in a single pass so a big
 * "apply to all" never rebuilds the array once per page.
 */
export function setCropForPages(
  entries: ReadonlyArray<CropBoxEntry>,
  pages: ReadonlyArray<number>,
  rect: CropKeepRect,
): CropBoxEntry[] {
  const targets = new Set<number>();
  for (const p of pages) {
    if (Number.isInteger(p) && p >= 1) targets.add(p);
  }
  if (targets.size === 0) return entries.slice();
  const others = entries.filter((e) => !targets.has(e.page));
  const normalized = normalizeCropKeepRect(rect);
  // A no-op (full-page) rect clears the targets instead of cropping them.
  if (isFullPageCrop(normalized)) return others;
  const applied: CropBoxEntry[] = [];
  for (const page of targets) applied.push({ page, ...normalized });
  return [...others, ...applied];
}
