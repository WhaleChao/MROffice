/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the PDF *read mode* (night / invert / sepia) viewer filter.
 *
 * A long-read comfort feature (Acrobat "Invert colors" / WPS night mode): the
 * reader can flip the rendered PAGE BITMAP to an inverted (dark) or sepia tone
 * without altering the document. This is purely an ephemeral viewer preference
 * (like zoom or the single/continuous view mode) — nothing is persisted and the
 * exported file is never touched.
 *
 * The filter is a CSS `filter` string applied to the page-canvas layer ONLY
 * (`canvas[data-pdf-page]`), so the SVG annotation overlays (highlights, ink,
 * shapes, notes) keep their own colors and stay legible. This module holds the
 * mode model + the filter mapping with no DOM / React dependency so it can be
 * unit-tested in isolation and reused by the viewer.
 *
 * Modes:
 *   - "normal" — no filter (the default; same bytes as today).
 *   - "invert" — `invert(1) hue-rotate(180deg)`: a true dark mode. The bare
 *     `invert(1)` flips brightness AND hue (blues→orange), which makes colour
 *     figures read wrong; the `hue-rotate(180deg)` cancels the hue flip so a
 *     black-on-white page becomes white-on-black while colour content keeps its
 *     hue, matching how Acrobat's "invert" looks.
 *   - "sepia" — a warm, low-contrast tone for reduced glare (`sepia` + a slight
 *     brightness/contrast trim), without inverting.
 */

/** The supported read modes (cycle order: normal → invert → sepia → normal). */
export type PdfReadMode = "normal" | "invert" | "sepia";

/** Ordered list of read modes (the toolbar toggle cycles through these). */
export const PDF_READ_MODES: readonly PdfReadMode[] = [
  "normal",
  "invert",
  "sepia",
];

/** Default read mode on load (no filter — byte-for-byte the current behavior). */
export const PDF_DEFAULT_READ_MODE: PdfReadMode = "normal";

const PDF_READ_MODE_SET = new Set<string>(PDF_READ_MODES);

/** True when a value is one of the supported read modes. */
export function isPdfReadMode(value: unknown): value is PdfReadMode {
  return typeof value === "string" && PDF_READ_MODE_SET.has(value);
}

/**
 * The CSS `filter` value for a read mode, applied to the page-canvas layer only.
 * Returns "none" for the normal mode (and for any unknown value) so the property
 * can be set unconditionally. The strings are inert filter literals (no
 * url()/expression), so they are safe to drop into an inline `style.filter`.
 *
 *   - invert: flip brightness then cancel the hue flip (true dark mode).
 *   - sepia:  warm tone with a small contrast/brightness trim to soften glare.
 */
export function cssFilterForReadMode(mode: unknown): string {
  switch (mode) {
    case "invert":
      return "invert(1) hue-rotate(180deg)";
    case "sepia":
      return "sepia(0.6) brightness(0.96) contrast(0.95)";
    case "normal":
    default:
      return "none";
  }
}

/**
 * The next read mode when the toggle is pressed, cycling
 * normal → invert → sepia → normal. An unknown / current-unmatched value starts
 * the cycle from the first non-normal mode (so a single press always changes
 * something visible).
 */
export function nextReadMode(mode: unknown): PdfReadMode {
  const index = PDF_READ_MODES.indexOf(mode as PdfReadMode);
  if (index < 0) return PDF_READ_MODES[1] ?? PDF_DEFAULT_READ_MODE;
  return PDF_READ_MODES[(index + 1) % PDF_READ_MODES.length];
}

/** True when the mode applies a non-trivial filter (used to flag the toggle active). */
export function isReadModeActive(mode: unknown): boolean {
  return isPdfReadMode(mode) && mode !== "normal";
}
