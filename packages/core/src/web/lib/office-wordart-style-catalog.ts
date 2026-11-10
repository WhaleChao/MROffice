/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import source from "./office-wordart-style-catalog.json";
import {
  normalizeOfficeTextPaint,
  type OfficeTextPaint,
} from "./office-text-paint";

export interface OfficeWordArtRecipe {
  id: string;
  label: { en: string; zh: string };
  paint: OfficeTextPaint;
}

/** Owner screenshot: five columns by three rows. These are explicit paint
 * recipes, not claimed Microsoft style-reference IDs or theme transforms. */
export const OFFICE_WORDART_STYLE_RECIPES: readonly OfficeWordArtRecipe[] =
  source.map((recipe) => ({
    ...recipe,
    paint: normalizeOfficeTextPaint(recipe.paint)!,
  }));

/** The style's name as a `console-notebooks` catalogue key (every editor's
 * WordArt gallery reads the same names). */
export function officeWordArtLabelKey(recipe: OfficeWordArtRecipe): string {
  return `spreadsheet.chartFormatCopy.wordArtGallery.styles.${recipe.id}`;
}
