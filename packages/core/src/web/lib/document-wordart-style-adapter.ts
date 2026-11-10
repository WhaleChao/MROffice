/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  DOCUMENT_WORD_ART_DEFAULT_FONT_SIZE_PT,
  DOCUMENT_WORD_ART_DEFAULT_HEIGHT_EMU,
  DOCUMENT_WORD_ART_DEFAULT_WIDTH_EMU,
  documentShapeDefaultAttrs,
  normalizeDocumentShapeAttrs,
  type DocumentShapeAttrs,
} from "./document-shapes";
import { sameOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";
import type { OfficeWordArtRecipe } from "./office-wordart-style-catalog";
import type { OfficeDefaultShapeStyle } from "./office-default-shape-style";

/** Word's Set as Default Shape: the selected shape's formatting as the
 * SHARED default-shape record (fill, outline, shadow, WordArt paint). */
export function documentDefaultShapeStyleFromAttrs(
  attrs: DocumentShapeAttrs,
): OfficeDefaultShapeStyle {
  return {
    fill: attrs.fill,
    outline: attrs.outline,
    outlineWidthPt: attrs.outlineWidthPt,
    outlineDash: attrs.outlineStyle,
    shadow: attrs.effect === "shadow" || Boolean(attrs.shapeShadow),
    ...(attrs.textPaint ? { textPaint: attrs.textPaint } : {}),
  };
}

/** The attrs patch a stored default-shape record puts on a new shape. */
export function documentDefaultShapeStylePatch(
  style: OfficeDefaultShapeStyle | null,
): Partial<DocumentShapeAttrs> {
  if (!style) return {};
  return {
    ...(style.fill !== undefined ? { fill: style.fill } : {}),
    ...(style.outline !== undefined ? { outline: style.outline } : {}),
    ...(style.outlineWidthPt !== undefined ? { outlineWidthPt: style.outlineWidthPt } : {}),
    ...(style.outlineDash !== undefined ? { outlineStyle: style.outlineDash } : {}),
    ...(style.shadow !== undefined ? { effect: style.shadow ? "shadow" : "none" } : {}),
    ...(style.textPaint
      ? {
          textPaint: style.textPaint,
          textFill: style.textPaint.fill,
          textOutline: style.textPaint.outline,
          textOutlineWidthPt: style.textPaint.outlineWidthPt,
        }
      : {}),
  };
}

/**
 * Word's Insert > WordArt: a fill-less, outline-less text box carrying one
 * SHARED WordArt recipe (the same catalog the Shape Format tab and the deck
 * / worksheet insert commands use), bold, at Word's default WordArt size.
 */
export function documentWordArtInsertAttrs(
  recipe: OfficeWordArtRecipe,
  text: string,
): DocumentShapeAttrs {
  return normalizeDocumentShapeAttrs({
    // A plain shape carrying text, not a text box: the native save writer
    // has no text-box (wps:txbx) writer, so a WordArt inserted as a text box
    // failed every save of the document with "isTextBox non-default shape
    
    // and fill-less look are what the writer round-trips.
    ...documentShapeDefaultAttrs("rect"),
    textDirection: "horizontal",
    widthEmu: DOCUMENT_WORD_ART_DEFAULT_WIDTH_EMU,
    heightEmu: DOCUMENT_WORD_ART_DEFAULT_HEIGHT_EMU,
    fill: "none",
    outline: "none",
    text,
    ...documentWordArtStylePatch(recipe),
    textFontFamily: "Arial",
    textFontSizePt: DOCUMENT_WORD_ART_DEFAULT_FONT_SIZE_PT,
    textBold: true,
    textVerticalAlign: "middle",
  });
}

export function documentWordArtStylePatch(recipe: OfficeWordArtRecipe) {
  return {
    textPaint: recipe.paint,
    textFill: recipe.paint.fill,
    textOutline: recipe.paint.outline,
    textOutlineWidthPt: recipe.paint.outlineWidthPt,
    textEffect: "none" as const,
  };
}

export function documentWordArtMatchesStyle(
  attrs: DocumentShapeAttrs,
  recipe: OfficeWordArtRecipe,
) {
  return Boolean(
    attrs.textPaint && sameOfficeTextPaint(attrs.textPaint, recipe.paint),
  );
}

export function officeTextPaintEffect(
  paint: OfficeTextPaint | null | undefined,
) {
  if (paint?.bevel) return "bevel" as const;
  if (paint?.glow) return "glow" as const;
  if (paint?.reflection) return "reflection" as const;
  if (paint?.shadow) return "shadow" as const;
  return "none" as const;
}
