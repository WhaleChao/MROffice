/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Insert Blank Page button's choice as a pure function: where the new page
 * goes (after the current page, or after the page before it) and its size.
 * Shared by the insert-page dialog and the headless PDF editor commands.
 */

import {
  PDF_PAGE_SIZE_PRESETS,
  type PdfPageSizePresetKey,
} from "./page-size";

export type PdfInsertPagePlacement = "before" | "after";
export type PdfInsertPagePreset = Extract<
  PdfPageSizePresetKey,
  "a4" | "letter" | "legal"
>;
export type PdfInsertPageOrientation = "portrait" | "landscape";

export interface PdfInsertPageSelection {
  afterPage: number;
  width: number;
  height: number;
  placement: PdfInsertPagePlacement;
  preset: PdfInsertPagePreset;
  orientation: PdfInsertPageOrientation;
  /** The page it goes beside was named by the command itself (the page
   * grid's 在前面/后面插入空白页 on the page pressed), not read off the
   * current page as the dialog reads it. */
  explicitTarget?: boolean;
}

export const PDF_INSERT_PAGE_SIZES: Record<
  PdfInsertPagePreset,
  { width: number; height: number; label: string; dimensions: string }
> = {
  a4: {
    width: PDF_PAGE_SIZE_PRESETS.a4.width,
    height: PDF_PAGE_SIZE_PRESETS.a4.height,
    label: "A4",
    dimensions: "210 × 297 mm",
  },
  letter: {
    width: PDF_PAGE_SIZE_PRESETS.letter.width,
    height: PDF_PAGE_SIZE_PRESETS.letter.height,
    label: "Letter",
    dimensions: "8.5 × 11 in",
  },
  legal: {
    width: PDF_PAGE_SIZE_PRESETS.legal.width,
    height: PDF_PAGE_SIZE_PRESETS.legal.height,
    label: "Legal",
    dimensions: "8.5 × 14 in",
  },
};

export function resolvePdfInsertPageSelection(options: {
  placement: PdfInsertPagePlacement;
  currentPage: number;
  previousPage?: number | null;
  preset: PdfInsertPagePreset;
  orientation: PdfInsertPageOrientation;
}): PdfInsertPageSelection {
  const currentPage =
    Number.isInteger(options.currentPage) && options.currentPage > 0
      ? options.currentPage
      : 1;
  const previousPage =
    Number.isInteger(options.previousPage) && (options.previousPage ?? -1) >= 0
      ? (options.previousPage as number)
      : 0;
  const size = PDF_INSERT_PAGE_SIZES[options.preset] ?? PDF_INSERT_PAGE_SIZES.letter;
  const landscape = options.orientation === "landscape";
  return {
    afterPage: options.placement === "before" ? previousPage : currentPage,
    width: landscape ? size.height : size.width,
    height: landscape ? size.width : size.height,
    placement: options.placement,
    preset: options.preset,
    orientation: options.orientation,
  };
}
