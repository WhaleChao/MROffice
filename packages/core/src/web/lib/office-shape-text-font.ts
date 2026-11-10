/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The FONT of a shape's label, in the one vocabulary the Office hosts share.
 *
 * DrawingML keeps it on the label's runs (`a:rPr`, ISO/IEC 29500-1 §21.1.2.3.9):
 * `@b` and `@i` are booleans, `@u` one of `ST_TextUnderlineType` (`sng`, `dbl`,
 * …), `@strike` one of `ST_TextStrikeType` (`noStrike`, `sngStrike`,
 * `dblStrike`), and the typeface is a child — `a:latin`, `a:ea`, `a:cs` — whose
 * `@typeface` names the font for each script. Selecting a shape (not its text)
 * in Excel and pressing Bold, or picking a font on Home, rewrites every run of
 * the label, which is why a whole-shape font is one value per property here
 * and not a per-run list.
 *
 * The field names are the ones the Docs shape already persists
 * (`textFontFamily`, `textBold` in `lib/document-shapes`), extended with the
 * rest of Excel's Home ▸ Font group. Every field is optional and absent means
 * "what the label inherits", so a shape stored before this module draws and
 * saves exactly as it did.
 */

import type { CSSProperties } from "react";

/** The underline variants Excel's Home tab offers (`a:rPr@u` `sng` / `dbl`). */
export const OFFICE_SHAPE_TEXT_UNDERLINES = ["single", "double"] as const;
export type OfficeShapeTextUnderline =
  (typeof OFFICE_SHAPE_TEXT_UNDERLINES)[number];

/** A family name longer than this is not one any Office font picker offers. */
export const OFFICE_SHAPE_MAX_FONT_FAMILY_LENGTH = 80;

export interface OfficeShapeTextFont {
  /** `a:latin@typeface` (and `a:ea`/`a:cs` where the writer states them). */
  textFontFamily?: string;
  /**
   * `a:rPr@b`. A stated `false` is kept: WordArt draws bold by default, so an
   * author who turned bold OFF on one says something an absent field does not.
   */
  textBold?: boolean;
  /** `a:rPr@i`. */
  textItalic?: boolean;
  /** `a:rPr@u`; absent is no underline. */
  textUnderline?: OfficeShapeTextUnderline;
  /** `a:rPr@strike="sngStrike"`; absent is no strikethrough. */
  textStrike?: boolean;
}

export const OFFICE_SHAPE_TEXT_FONT_KEYS = [
  "textFontFamily",
  "textBold",
  "textItalic",
  "textUnderline",
  "textStrike",
] as const satisfies readonly (keyof OfficeShapeTextFont)[];

/**
 * A family name as the pickers and the writers accept it: control characters
 * and markup delimiters removed, trimmed, bounded. `undefined` for anything
 * that leaves no name, so the caller decides what "no family" means (Docs
 * falls back to its default face, a sheet shape inherits the theme's).
 */
export function normalizeOfficeShapeFontFamily(
  value: unknown,
): string | undefined {
  if (typeof value !== "string") return undefined;
  const normalized = value
    .replace(/[\u0000-\u001f\u007f<>]/g, "")
    .trim()
    .slice(0, OFFICE_SHAPE_MAX_FONT_FAMILY_LENGTH);
  return normalized || undefined;
}

/** Only the fields that say something; junk is dropped field by field. */
export function normalizeOfficeShapeTextFont(
  value: unknown,
): OfficeShapeTextFont {
  const o =
    value && typeof value === "object" && !Array.isArray(value)
      ? (value as Record<string, unknown>)
      : {};
  const family = normalizeOfficeShapeFontFamily(o.textFontFamily);
  return {
    ...(family ? { textFontFamily: family } : {}),
    ...(typeof o.textBold === "boolean" ? { textBold: o.textBold } : {}),
    ...(o.textItalic === true ? { textItalic: true } : {}),
    ...((OFFICE_SHAPE_TEXT_UNDERLINES as readonly unknown[]).includes(
      o.textUnderline,
    )
      ? { textUnderline: o.textUnderline as OfficeShapeTextUnderline }
      : {}),
    ...(o.textStrike === true ? { textStrike: true } : {}),
  };
}

/**
 * The CSS a label is drawn with. `familyStack` lets a host put its own
 * fallback faces after the named one (the sheet uses its cell stack); without
 * it the name is followed by a generic sans-serif. `boldByDefault` is WordArt,
 * which draws bold unless the author said otherwise.
 */
export function officeShapeTextFontCss(
  font: OfficeShapeTextFont,
  options: {
    familyStack?: (family: string) => string;
    boldByDefault?: boolean;
  } = {},
): CSSProperties {
  const family = font.textFontFamily;
  const bold = font.textBold ?? options.boldByDefault === true;
  const decorations = [
    font.textUnderline ? "underline" : "",
    font.textStrike ? "line-through" : "",
  ].filter(Boolean);
  return {
    ...(family
      ? {
          fontFamily: options.familyStack
            ? options.familyStack(family)
            : `"${family.replace(/["\\]/g, "")}", sans-serif`,
        }
      : {}),
    fontWeight: bold ? 700 : 400,
    ...(font.textItalic ? { fontStyle: "italic" } : {}),
    ...(decorations.length
      ? {
          textDecorationLine: decorations.join(" "),
          ...(font.textUnderline === "double"
            ? { textDecorationStyle: "double" as const }
            : {}),
        }
      : {}),
  };
}
