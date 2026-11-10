/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The family, bold and italic a font's base name says (mirror of replay's
 * _native_property_font_identity). */
export function pdfNativeTextFontIdentity(
  baseFont: string,
): Readonly<{ family: string; bold: boolean; italic: boolean }> {
  const styleName = baseFont.replace(/^[A-Z]{6}\+/, "");
  const bold = /(?:^|[-_,])(bold|black|heavy|semibold)/i.test(styleName);
  const italic =
    /(?:^|[-_,])(italic|oblique)/i.test(styleName) ||
    /(?:italic|oblique)$/i.test(styleName);
  const family =
    styleName
      .replace(/[-_,]?(?:bold|black|heavy|semibold)/gi, "")
      .replace(/[-_,]?(?:italic|oblique)/gi, "")
      .replace(/[-_,]+$/g, "") || styleName;
  return Object.freeze({ family, bold, italic });
}










export function pdfNativeFontFamilyKey(baseFontOrFamily: string): string {
  return pdfNativeTextFontIdentity(baseFontOrFamily).family
    .replace(/["'\s_,-]/gu, "")
    .toLowerCase()
    .replace(/(?:psmt|mt|ps)$/u, "")
    .replace(/(?:regular|roman|book|normal|plain)$/u, "")
    .replace(/(?:condensed|cond|cn)$/u, "cn");
}

/** Whether a face's name says bold (italic) as a reader reads it, run on
 * with the family too: "Univers-CondensedBold", "ArialBoldItalicMT" -- which
 * pdfNativeTextFontIdentity, mirroring the replay, reads only after a
 * separator. For what the editor shows (加粗 pressed on a bold heading). */





export function pdfNativeFontNameKey(name: string): string {
  return name.replace(/^[A-Z]{6}\+/u, "").replace(/["'\s_-]/gu, "").toLowerCase();
}

/** How a face is set, as its name says: bold for any heavier weight
 * (bold, black, heavy, semibold, demi), italic for italic, oblique or
 * slanted. The writers read a face so (the paragraph rewrite's face style,
 * the replay's _face_style), and the editor reads a face by the same words:
 * a face it named otherwise ("Univers-Demi" bold to the writer, regular in
 * the ribbon) was pressed bold and set bolder. */
export function pdfNativeFaceNameStyle(name: string): Readonly<{ bold: boolean; italic: boolean }> {
  return Object.freeze({ bold: /bold|black|heavy|semibold|demi/iu.test(name), italic: /italic|oblique|slanted/iu.test(name) });
}

export function pdfNativeFontNameReadsStyle(
  baseFont: string,
): Readonly<{ bold: boolean; italic: boolean }> {
  const own = pdfNativeTextFontIdentity(baseFont);
  const named = pdfNativeFaceNameStyle(baseFont.replace(/^[A-Z]{6}\+/, ""));
  // TeX's math and text italics say so only in their family's code (CMMI10, CMTI10).
  const texItalic = /^(?:cmmi|cmti|cmsl|cmbxti|lmmi|lmmathitalic)/iu.test(baseFont.replace(/^[A-Z]{6}\+/, ""));
  return Object.freeze({ bold: own.bold || named.bold, italic: own.italic || named.italic || texItalic });
}
