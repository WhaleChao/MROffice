/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * View ▸ Color/Grayscale (PPT-A109): how the editor shows the slides --
 * Color, Grayscale or Pure Black and White. A view only (PowerPoint keeps
 * the colours in the deck and uses these to check how a slide reads printed
 * or projected without colour); session state, never saved.
 *
 * PowerPoint decides per object how it renders in grayscale (its Grayscale
 * tab: Automatic, Grayscale, Light Grayscale, Inverse, Black with White Fill
 * ...); without those settings the editor previews the whole slide with a
 * filter: Grayscale drops the colour, Pure Black and White thresholds every
 * tone to black or white.
 */
export type PresentationViewColorMode = "color" | "grayscale" | "blackWhite";

export function presentationViewColorFilter(
  mode: PresentationViewColorMode,
): string | undefined {
  switch (mode) {
    case "grayscale":
      return "grayscale(1)";
    case "blackWhite":
      return "grayscale(1) contrast(1000%)";
    default:
      return undefined;
  }
}
