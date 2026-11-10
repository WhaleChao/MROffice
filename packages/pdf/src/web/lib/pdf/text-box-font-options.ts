/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { CJK_SANS_STACK, CJK_SERIF_STACK, FONT_STACK } from "./font-resolve";

/**
 * Curated font-family choices for the selected-box toolbar. Each option's
 * `stack` is one of the stable stacks from font-resolve (so the overlay renders
 * a real face) and its `sourceFontName` is a canonical name the export-side
 * matchers recognise: `standardFontForFamily` (flatten-apply) consults
 * `sourceFontName` BEFORE `fontFamily`, so a pick must overwrite the original
 * PDF font name or the export would silently keep the old face. Labels state
 * the export target honestly: Latin picks bake the named base-14 face; CJK
 * picks bake CJK glyphs with the embedded Noto Sans SC fallback font (the only
 * bundled CJK program) while their Latin glyphs use the closest base-14 face.
 */
export interface TextBoxFontFamilyOption {
  id: "helvetica" | "times" | "courier" | "cjkSans" | "cjkSerif";
  stack: string;
  sourceFontName: string;
  labelKey: string;
}

export const TEXTBOX_FONT_FAMILY_OPTIONS: readonly TextBoxFontFamilyOption[] = [
  {
    id: "helvetica",
    stack: FONT_STACK.sans,
    sourceFontName: "Helvetica",
    labelKey: "pdf.textbox.fontFamily.helvetica",
  },
  {
    id: "times",
    stack: FONT_STACK.serif,
    sourceFontName: "Times New Roman",
    labelKey: "pdf.textbox.fontFamily.times",
  },
  {
    id: "courier",
    stack: FONT_STACK.mono,
    sourceFontName: "Courier New",
    labelKey: "pdf.textbox.fontFamily.courier",
  },
  {
    id: "cjkSans",
    stack: CJK_SANS_STACK,
    sourceFontName: "Microsoft YaHei",
    labelKey: "pdf.textbox.fontFamily.cjkSans",
  },
  {
    id: "cjkSerif",
    stack: CJK_SERIF_STACK,
    sourceFontName: "SimSun",
    labelKey: "pdf.textbox.fontFamily.cjkSerif",
  },
];
