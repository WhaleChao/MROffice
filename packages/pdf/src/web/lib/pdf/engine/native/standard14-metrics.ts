/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isStandard14LatinBaseFont, refuse } from "./list-text-model";
import { PDF_SIMPLE_FONT_BASE_ENCODINGS, PDF_STANDARD_ENCODING_GLYPH_NAMES } from "./simple-font-encodings";
import { StandardFontEmbedder, isStandardFont } from "pdf-lib";
import type { PdfNativeStandard14EncodingKind, PdfNativeStandard14WinAnsiGlyphAuthority } from "./list-text-fidelity";

export function readMraiNativeStandard14WinAnsiAuthority(
  baseFont: string,
): readonly PdfNativeStandard14WinAnsiGlyphAuthority[] | null {
  if (!isStandard14LatinBaseFont(baseFont) || !isStandardFont(baseFont)) {
    return null;
  }
  const embedder = StandardFontEmbedder.for(baseFont);
  const codes = new Set<number>();
  const entries = embedder.encoding.supportedCodePoints.map(
    (unicodeCodePoint) => {
      const { code, name: glyphName } =
        embedder.encoding.encodeUnicodeCodePoint(unicodeCodePoint);
      // pdf.js and the WinAnsi Standard 14 contract resolve byte 0x9F to the
      // uppercase Ydieresis glyph. pdf-lib's encoding table exposes the
      // lowercase alias name even though the Unicode character is U+0178.
      const metricGlyphName = code === 0x9f ? "Ydieresis" : glyphName;
      const width = embedder.font.getWidthOfGlyph(metricGlyphName);
      if (
        !Number.isSafeInteger(code) ||
        code < 0 ||
        code > 0xff ||
        codes.has(code) ||
        typeof width !== "number" ||
        !Number.isFinite(width) ||
        width < 0 ||
        width > 10_000
      ) {
        refuse("Standard 14 WinAnsi metrics are missing or ambiguous.");
      }
      codes.add(code);
      return Object.freeze({
        code,
        unicodeCodePoint,
        character: String.fromCodePoint(unicodeCodePoint),
        glyphName: metricGlyphName,
        width,
      });
    },
  );
  if (entries.length !== 218 || codes.size !== 218) {
    refuse("Standard 14 WinAnsi coverage is incomplete.");
  }
  return Object.freeze(entries.sort((left, right) => left.code - right.code));
}

/** Glyph authority of one Standard 14 font under the encoding its dictionary
 * draws with: WinAnsi through pdf-lib's encoder, or the font's built-in
 * StandardEncoding (the default when /Encoding is absent) through the frozen
 * code tables the simple-font path already shares with the server. */
export function readMraiNativeStandard14Authority(
  baseFont: string,
  kind: PdfNativeStandard14EncodingKind,
): readonly PdfNativeStandard14WinAnsiGlyphAuthority[] | null {
  if (kind === "standard14-winansi") {
    return readMraiNativeStandard14WinAnsiAuthority(baseFont);
  }
  if (kind === "standard14-encoded") {
    // That font's codes are whatever its own /Encoding dictionary makes them;
    // a base font name and a kind are not enough to rebuild them here.
    return null;
  }
  if (!isStandard14LatinBaseFont(baseFont) || !isStandardFont(baseFont)) {
    return null;
  }
  const embedder = StandardFontEmbedder.for(baseFont);
  const characters = PDF_SIMPLE_FONT_BASE_ENCODINGS.StandardEncoding;
  const codes = Object.keys(characters).map(Number);
  if (
    codes.length !== 149 ||
    Object.keys(PDF_STANDARD_ENCODING_GLYPH_NAMES).length !== 149 ||
    codes.some((code) => PDF_STANDARD_ENCODING_GLYPH_NAMES[code] === undefined)
  ) {
    refuse("Standard 14 StandardEncoding coverage is incomplete.");
  }
  const entries = codes.map((code) => {
    const character = characters[code];
    const glyphName = PDF_STANDARD_ENCODING_GLYPH_NAMES[code];
    const unicodeCodePoint = character.codePointAt(0);
    const width = embedder.font.getWidthOfGlyph(glyphName);
    if (
      [...character].length !== 1 ||
      unicodeCodePoint === undefined ||
      typeof width !== "number" ||
      !Number.isFinite(width) ||
      width < 0 ||
      width > 10_000
    ) {
      refuse("Standard 14 StandardEncoding metrics are missing or ambiguous.");
    }
    return Object.freeze({ code, unicodeCodePoint, character, glyphName, width });
  });
  return Object.freeze(entries.sort((left, right) => left.code - right.code));
}
