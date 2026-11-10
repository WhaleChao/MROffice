/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { presentationThemeGridColors } from "./presentation-theme-color-binding";
import {
  presentationWordArtStyleName,
  presentationWordArtStyles,
} from "./presentation-wordart-gallery";
import {
  OFFICE_WORDART_STYLE_RECIPES,
  officeWordArtLabelKey,
  type OfficeWordArtRecipe,
} from "./office-wordart-style-catalog";

/** The twenty under a theme's ten-colour grid (`tx1`, `bg1`, `accent1`...). */
export function officeWordArtGalleryRecipes(
  theme?: Readonly<Record<string, string>> | null,
): readonly OfficeWordArtRecipe[] {
  return presentationWordArtStyles(theme ?? presentationThemeGridColors(null)).map((style) => ({
    id: style.id,
    label: { en: "", zh: "" },
    paint: style.paint,
  }));
}

/**
 * The twenty under the Office theme with a host's own accents (a chart that
 * knows only its deck's six accents). Cached by the accents: a ribbon asks on
 * every render.
 */
const BY_ACCENTS = new Map<string, readonly OfficeWordArtRecipe[]>();
export function officeWordArtGalleryRecipesForAccents(
  accents: readonly string[] | null | undefined,
): readonly OfficeWordArtRecipe[] {
  if (!accents?.length) return OFFICE_WORDART_GALLERY_RECIPES;
  const key = accents.join(",");
  let recipes = BY_ACCENTS.get(key);
  if (!recipes) {
    const theme: Record<string, string> = { ...presentationThemeGridColors(null) };
    accents.slice(0, 6).forEach((color, index) => {
      if (/^#[0-9a-f]{6}$/i.test(color)) theme[`accent${index + 1}`] = color.toLowerCase();
    });
    recipes = officeWordArtGalleryRecipes(theme);
    if (BY_ACCENTS.size > 32) BY_ACCENTS.clear();
    BY_ACCENTS.set(key, recipes);
  }
  return recipes;
}

/** The twenty under the Office theme: a document's and a workbook's gallery. */
export const OFFICE_WORDART_GALLERY_RECIPES: readonly OfficeWordArtRecipe[] =
  officeWordArtGalleryRecipes();

/** A style by id: one of the twenty, or one of the fifteen offered before. */
export function officeWordArtRecipe(
  id: string,
  recipes: readonly OfficeWordArtRecipe[] = OFFICE_WORDART_GALLERY_RECIPES,
): OfficeWordArtRecipe | null {
  return (
    recipes.find((recipe) => recipe.id === id) ??
    OFFICE_WORDART_STYLE_RECIPES.find((recipe) => recipe.id === id) ??
    null
  );
}

/**
 * A tile's name: PowerPoint's for the twenty ("Fill: Black, Text color 1;
 * Shadow"), the catalog's for the fifteen. `t` is the console-notebooks
 * translator.
 */
export function officeWordArtRecipeName(
  recipe: OfficeWordArtRecipe,
  t: (key: string, values?: Record<string, string>) => string,
): string {
  const item = /^ppt-wordart-(\d+)$/.exec(recipe.id);
  return item ? presentationWordArtStyleName(Number(item[1]), t) : t(officeWordArtLabelKey(recipe));
}
