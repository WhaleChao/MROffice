/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation PRINT / HANDOUT settings (打印 / 讲义) — PowerPoint / WPS parity.
 *
 * These are the deck's print layout: Print (the title-row face) prints what
 * the Print Settings dialog shows, and the server applies the scale. Like
 * PowerPoint's Print Layout the default is full-page slides (1 per page);
 * the Export ▸ PDF handout item keeps its historical 2-per-page grid when
 * the layout is full-page slides, since a handout is always a grid.
 *
 * A pure, framework-free module mirroring `lib/spreadsheet/print-settings.ts`:
 * a small settings model, a normalizer that collapses unknown / stale /
 * hand-edited values to a safe default, and a helper that maps the chosen
 * slides-per-page to the export's `layout` query param. Consumed by the deck
 * PDF export URL (PowerPoint "Print ▸ Handouts: N slides per page") and
 * validated by the FE schema.
 *
 * Key concepts
 * ─────────────
 * slidesPerPage : how many slide thumbnails are tiled onto each printed page.
 *                 1 = one slide per page (the plain "PDF" export, no handout
 *                 grid); 2 / 3 / 4 / 6 / 9 = a handout grid, PowerPoint's
 *                 Handouts choices (the 3-slide page has note lines).
 * printHiddenSlides : PowerPoint's Print Hidden Slides (off by default, as
 *                 Save as PDF leaves hidden slides out).
 * scale         : a print SCALE percentage [10, 400] (100 = unscaled, default).
 *                 Persisted for round-trip; the handout grid already auto-fits
 *                 each cell, so this is an additive future / PPTX-handout hint.
 *
 * Neither field is mandatory; a deck without printSettings exports exactly as
 * before (the handout item uses `handout-2`). That default path is byte-for-byte
 * unchanged.
 */

// ── Types ────────────────────────────────────────────────────────────────────

/** The supported slides-per-page counts (PowerPoint handout layouts). */
export const SLIDES_PER_PAGE_OPTIONS = [1, 2, 3, 4, 6, 9] as const;
export type SlidesPerPage = (typeof SLIDES_PER_PAGE_OPTIONS)[number];
const SLIDES_PER_PAGE_SET: ReadonlySet<number> = new Set(SLIDES_PER_PAGE_OPTIONS);

/** The default slides-per-page: full-page slides, PowerPoint's default print
 * layout. (It was 2 while these settings only fed the handout export; a deck
 * that stored nothing still exports the same handout-2 grid, see
 * {@link handoutExportLayoutForSlidesPerPage}.) */
export const DEFAULT_SLIDES_PER_PAGE: SlidesPerPage = 1;

/** The handout export's grid when the print layout is full-page slides. */
export const DEFAULT_HANDOUT_SLIDES_PER_PAGE: SlidesPerPage = 2;

/** Print SCALE percent bounds (matches the spreadsheet percent scaling). */
export const PRINT_SCALE_MIN = 10;
export const PRINT_SCALE_MAX = 400;
export const DEFAULT_PRINT_SCALE = 100;

export interface PresentationPrintSettings {
  /** How many slides are tiled per printed page (1 / 2 / 3 / 4 / 6 / 9). */
  slidesPerPage?: SlidesPerPage;
  /** Print scale percentage [10, 400] (100 = unscaled). */
  scale?: number;
  /** Print Hidden Slides (stored only when on). */
  printHiddenSlides?: boolean;
}

// ── Normalisation ─────────────────────────────────────────────────────────────

/**
 * Clamp / whitelist a persisted slides-per-page value to one of
 * {@link SLIDES_PER_PAGE_OPTIONS}, defaulting to {@link DEFAULT_SLIDES_PER_PAGE}
 * for a missing / non-finite / non-whitelisted value. Pure + side-effect free.
 */
export function normalizeSlidesPerPage(value: unknown): SlidesPerPage {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_SLIDES_PER_PAGE;
  }
  return SLIDES_PER_PAGE_SET.has(value)
    ? (value as SlidesPerPage)
    : DEFAULT_SLIDES_PER_PAGE;
}

/**
 * Clamp a persisted print scale percentage to [PRINT_SCALE_MIN, PRINT_SCALE_MAX]
 * (rounded), defaulting to {@link DEFAULT_PRINT_SCALE} (100) for a missing /
 * non-finite value. Mirrors the spreadsheet percent-scaling clamp.
 */
export function normalizePrintScale(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_PRINT_SCALE;
  }
  const rounded = Math.round(value);
  if (rounded < PRINT_SCALE_MIN) return PRINT_SCALE_MIN;
  return rounded > PRINT_SCALE_MAX ? PRINT_SCALE_MAX : rounded;
}

/**
 * Normalize a raw persisted print-settings blob into a
 * {@link PresentationPrintSettings}, keeping ONLY the non-default fields so a
 * deck whose settings equal the defaults (2 slides per page, 100 % scale)
 * serializes to `undefined` (no `printSettings` key persisted, byte-clean).
 * Returns `undefined` when the blob is not an object or every field is at its
 * default. Pure + side-effect free.
 */
export function normalizePresentationPrintSettings(
  raw: unknown,
): PresentationPrintSettings | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const r = raw as Record<string, unknown>;
  const out: PresentationPrintSettings = {};

  const slidesPerPage = normalizeSlidesPerPage(r.slidesPerPage);
  if (slidesPerPage !== DEFAULT_SLIDES_PER_PAGE) out.slidesPerPage = slidesPerPage;

  const scale = normalizePrintScale(r.scale);
  if (scale !== DEFAULT_PRINT_SCALE) out.scale = scale;

  if (r.printHiddenSlides === true) out.printHiddenSlides = true;

  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve the deck's print settings into a render-ready descriptor with every
 * field filled (defaults applied). Convenience for the export URL + the dialog
 * controls, which always want a concrete value.
 */
export function resolvePresentationPrintSettings(
  raw: unknown,
): Required<PresentationPrintSettings> {
  const r =
    raw && typeof raw === "object" && !Array.isArray(raw)
      ? (raw as Record<string, unknown>)
      : {};
  return {
    slidesPerPage: normalizeSlidesPerPage(r.slidesPerPage),
    scale: normalizePrintScale(r.scale),
    printHiddenSlides: r.printHiddenSlides === true,
  };
}

/**
 * Map a slides-per-page count to the deck PDF export's `layout` query value.
 * 1 slide per page is the plain one-slide-per-page PDF (NO `layout` param, so
 * `null` is returned); 2 / 3 / 4 / 6 / 9 map to the backend's `handout-N` grids.
 * Any non-whitelisted value defaults to the historical `handout-2`.
 */
export function handoutLayoutForSlidesPerPage(
  slidesPerPage: unknown,
): string | null {
  const n = normalizeSlidesPerPage(slidesPerPage);
  return n === 1 ? null : `handout-${n}`;
}

/**
 * The Export ▸ PDF handout item's layout: the print layout's grid, and the
 * historical handout-2 when the print layout is full-page slides (a handout is
 * always a grid; the plain PDF export is the one-slide-per-page file).
 */
export function handoutExportLayoutForSlidesPerPage(slidesPerPage: unknown): string {
  return (
    handoutLayoutForSlidesPerPage(slidesPerPage) ??
    `handout-${DEFAULT_HANDOUT_SLIDES_PER_PAGE}`
  );
}
