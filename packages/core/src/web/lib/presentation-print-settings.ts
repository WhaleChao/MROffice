/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
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
  /** Print Hidden Slides: absent = PowerPoint's default (on); stated only
   * when explicit (see {@link printHiddenSlidesAnswer}). */
  printHiddenSlides?: boolean;
  /** The imported `p:prnPr` (its attributes), kept for the deck written
   * whole; opaque to the editor. */
  source?: PresentationPrintSettingsSource;
}

export interface PresentationPrintSettingsSource {
  readonly prnPr: Readonly<Record<string, string>>;
}

/** The print settings with every editable field filled (no receipt). */
export type ResolvedPresentationPrintSettings = Required<
  Omit<PresentationPrintSettings, "source">
>;

/** `CT_PrintProperties` attributes a receipt may name (the API checks values). */
const PRN_PR_ATTRIBUTES: ReadonlySet<string> = new Set([
  "prnWhat",
  "clrMode",
  "hiddenSlides",
  "scaleToFitPaper",
  "frameSlides",
]);

function normalizePrintSettingsSource(
  value: unknown,
): PresentationPrintSettingsSource | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const prnPr = (value as Record<string, unknown>).prnPr;
  if (!prnPr || typeof prnPr !== "object" || Array.isArray(prnPr)) return undefined;
  const entries = Object.entries(prnPr as Record<string, unknown>);
  if (
    entries.some(
      ([name, text]) =>
        !PRN_PR_ATTRIBUTES.has(name) || typeof text !== "string" || text.length > 16,
    )
  ) {
    return undefined;
  }
  return { prnPr: Object.fromEntries(entries) as Record<string, string> };
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
 * deck whose settings equal the defaults (full-page slides, 100 % scale, Print
 * Hidden Slides unstated) serializes to `undefined` (no `printSettings` key
 * persisted, byte-clean). An explicit Print Hidden Slides, on or off, is kept.
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

  // Three-way: an explicit on or off is kept, absent is the default (on).
  if (typeof r.printHiddenSlides === "boolean") out.printHiddenSlides = r.printHiddenSlides;

  // The imported p:prnPr, carried as it came.
  const source = normalizePrintSettingsSource(r.source);
  if (source) out.source = source;

  return Object.keys(out).length > 0 ? out : undefined;
}

/**
 * Resolve the deck's print settings into a render-ready descriptor with every
 * field filled (defaults applied). Convenience for the export URL + the dialog
 * controls, which always want a concrete value.
 */
export function resolvePresentationPrintSettings(
  raw: unknown,
): ResolvedPresentationPrintSettings {
  const r =
    raw && typeof raw === "object" && !Array.isArray(raw)
      ? (raw as Record<string, unknown>)
      : {};
  return {
    slidesPerPage: normalizeSlidesPerPage(r.slidesPerPage),
    scale: normalizePrintScale(r.scale),
    // PowerPoint's Print prints hidden slides unless turned off.
    printHiddenSlides: r.printHiddenSlides !== false,
  };
}

/**
 * What the Print Settings dialog's Print Hidden Slides box stores: off is an
 * explicit `false`; on is the default (nothing stored) unless the deck held an
 * explicit value, when it is an explicit `true` -- a deck whose import read
 * Print Hidden Slides off states the user's on as such. (The page's alone:
 * nothing of it is written into the file, as PowerPoint writes nothing of its
 * Print pane.)
 */
export function printHiddenSlidesAnswer(
  checked: boolean,
  stored: boolean | undefined,
): boolean | undefined {
  if (!checked) return false;
  return stored === undefined ? undefined : true;
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
