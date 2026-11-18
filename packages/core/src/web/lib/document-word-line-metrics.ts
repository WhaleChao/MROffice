/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { wordWindowsEastAsianFactor } from "./office-fonts/word-windows-lines";

/**
 * Word's single-line height rule, shared by the browser DOCX importer, the
 * editor's auto line-spacing projection and the document attribute contract.
 *
 * For `w:line="240" w:lineRule="auto"` (Word "Single") with no document grid,
 * one line of text advances by
 *
 *   S = fontSize * ( eastAsianFont
 *                      ? 1.3 * (hhea.ascender - hhea.descender) / unitsPerEm
 *                      : (hhea.ascender - hhea.descender + hhea.lineGap) / unitsPerEm )
 *
 * measured from Word 16 (macOS) save-as-PDF output for ten installed faces at
 * 9pt and 12pt — Aptos, Calibri, Times New Roman (the one face whose lineGap is
 * non-zero, 1.1467), Arial Unicode MS (1.74), SimSun (1.2933), Songti SC
 * (1.82), DengXian (1.3533), PingFang SC (1.82), Microsoft YaHei (1.7133) and
 * STHeiti (1.30). Every font fits its branch within 0.5 percent.
 *
 * "East Asian font" is a property of the FONT, not of the text: the Arial
 * Unicode MS control paragraph is pure Latin and still takes the 1.3 factor.
 * The test is the font's OS/2 `ulCodePageRange1` CJK declaration, which Word
 * copies verbatim into `word/fontTable.xml` as `w:sig/@w:csb0` — so the
 * importer can read it out of the document without opening a local font file.
 */

/**
 * OS/2 `ulCodePageRange1` bits 17-21: JIS/Japan, Chinese Simplified (GB2312),
 * Korean Wansung, Chinese Traditional (Big5) and Korean Johab.
 */
export const WORD_EAST_ASIAN_CODE_PAGE_BITS = [17, 18, 19, 20, 21] as const;

/** True when a `w:sig/@w:csb0` value declares any East Asian code page. */
export function wordEastAsianCodePages(codePages: number): boolean {
  if (!Number.isInteger(codePages) || codePages < 0) return false;
  return WORD_EAST_ASIAN_CODE_PAGE_BITS.some(
    (bit) => Math.trunc(codePages / 2 ** bit) % 2 === 1,
  );
}

/**
 * The extra pitch Word gives an East Asian face. It belongs to the FONT's line
 * box; it is never a paragraph line-height multiplier, so `w:line` keeps
 * meaning exactly what the document says.
 */
export const WORD_EAST_ASIAN_SINGLE_LINE_FACTOR = 1.3;

/**
 * The factor where Word was measured to differ from the 1.3 rule: SimSun's
 * box is exactly 1 em and Word sets its line at 1.2933 em (Word 16 macOS,
 * 9 / 10 / 12 pt). At 1.3 every 宋体 line was 0.08 pt taller than Word's,
 * which over a 296-page textbook's text lines came to more than a page.
 */
const WORD_MEASURED_EAST_ASIAN_LINE_FACTORS: readonly (readonly [RegExp, number])[] = [
  [/^(simsun|宋体|nsimsun|新宋体)$/i, 1.2933],
];

/** Word's single-line factor for the East Asian face `family` (1.3 where
 * no face-specific value was measured). */
export function wordEastAsianSingleLineFactor(family: string | null | undefined): number {
  const name = (family ?? "").trim().replace(/^["']|["']$/g, "");
  for (const [pattern, factor] of WORD_MEASURED_EAST_ASIAN_LINE_FACTORS) if (pattern.test(name)) return factor;
  return WORD_EAST_ASIAN_SINGLE_LINE_FACTOR;
}

/**
 * Faces whose single line Windows Word reads from another box than the one
 * recorded (the hhea box of Word for macOS): Windows Word sets a face that
 * does not ask for its typographic metrics (OS/2 fsSelection bit 7) on its
 * OS/2 win box, usWinAscent + usWinDescent. For Yu Mincho and Yu Gothic that
 * box is 2636 / 2048 em where hhea's is 2257 / 2048, and Word's line is 1.3x
 * it: 1.6732 em -- 17.56 / 20.08 pt at 10.5 / 12 pt single, 20.17 pt at
 * 10.5 pt and 1.15 lines, two pitches of a 17.4 pt line grid at 10.5 pt as at
 * 12 pt (docx-look/probes/yu yu1, yu2). The other recorded East Asian faces
 * have one box both ways (MS Mincho 1.3 em, Meiryo 1.95 em measured alike).
 * Japanese documents take Yu Mincho from their theme's minorEastAsia face:
 * read at 1.43 em, corpus R315-D15 set its Japanese body one grid pitch a
 * line where Word sets two, a page short of three. The factor is relative to
 * the recorded hhea box, so a caller multiplies the box it already reads.
 */
const WORD_WINDOWS_EAST_ASIAN_LINE_FACTORS: readonly (readonly [RegExp, number])[] = [
  [/^(yu mincho|yu gothic|游明朝|游ゴシック)( regular)?$/i, (2636 * 1.3) / 2257],
];

/** `window.__docsWinBoxLineOff = true` asks the older reading (Yu Mincho and
 * Yu Gothic on their hhea box) for an A/B of one build. */
function windowsBoxLineOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWinBoxLineOff?: boolean }).__docsWinBoxLineOff === true;
}

/**
 * Word's single-line factor for the East Asian face `family` in a Word
 * document, as Windows Word sets it: `wordEastAsianSingleLineFactor`, save
 * for the faces Windows Word reads from their win box. The Docs editor's own;
 * other editors keep `wordEastAsianSingleLineFactor`.
 */
export function wordDocumentEastAsianSingleLineFactor(family: string | null | undefined): number {
  if (!windowsBoxLineOff()) {
    const name = (family ?? "").trim().replace(/^["']|["']$/g, "");
    for (const [pattern, factor] of WORD_WINDOWS_EAST_ASIAN_LINE_FACTORS) if (pattern.test(name)) return factor;
  }
  // Every other East Asian face whose Windows line is not 1.3x the recorded
  // box: the win box plus 0.15 of it on each side in whole font units, so
  // 1.2969 em for SimSun, KaiTi, MS Mincho (256 units), 1.3008 for Batang,
  // Gulim, MingLiU (1024) -- SimSun 15.56 pt at 12 pt where the Mac Word
  // measurement gave 15.52 (word-windows-lines.ts).
  const windows = wordWindowsEastAsianFactor(family);
  if (windows !== undefined) return windows;
  return wordEastAsianSingleLineFactor(family);
}

/** Upper bound mirrored from the OOXML font table the values are read from. */
const MAX_EAST_ASIAN_FONT_FAMILIES = 256;

/** A declared font name is a display string: no C0 controls and no DEL. */
export function hasControlCharacter(value: string): boolean {
  for (const character of value) {
    const code = character.codePointAt(0)!;
    if (code < 0x20 || code === 0x7f) return true;
  }
  return false;
}

/**
 * Closed validator for the `eastAsianFontFamilies` document attribute: the
 * document's own declared Word font names, deduplicated case-insensitively and
 * sorted. Returns null for anything else, so an unknown shape can never widen
 * the rule.
 */
export function normalizeEastAsianFontFamilies(value: unknown): string[] | null {
  if (!Array.isArray(value) || value.length > MAX_EAST_ASIAN_FONT_FAMILIES) {
    return null;
  }
  const seen = new Set<string>();
  const families: string[] = [];
  for (const entry of value) {
    if (
      typeof entry !== "string" ||
      entry.trim() !== entry ||
      !entry ||
      entry.length > 256 ||
      hasControlCharacter(entry)
    ) {
      return null;
    }
    const key = entry.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    families.push(entry);
  }
  return families.length ? families.sort() : null;
}

/** Which East Asian script a face is made for, by the names Office ships. */
const EAST_ASIAN_FACE_SCRIPTS: readonly [RegExp, "zh" | "ja" | "ko"][] = [
  [/^(ms (p)?(mincho|gothic|ui gothic)|ｍｓ ?(ｐ)?(明朝|ゴシック)|yu (mincho|gothic)|游(明朝|ゴシック)|meiryo|メイリオ|biz ud|hiragino|ヒラギノ|noto (sans|serif) (cjk )?jp)/i, "ja"],
  [/^(malgun gothic|맑은 고딕|batang|바탕|gulim|굴림|dotum|돋움|gungsuh|궁서|apple sd gothic|noto (sans|serif) (cjk )?kr)/i, "ko"],
  [/^(simsun|宋体|nsimsun|新宋体|simhei|黑体|fangsong|仿宋|kaiti|楷体|dengxian|等线|microsoft (ya|jheng)hei|微软雅黑|微軟正黑體|pmingliu|mingliu|新細明體|細明體|dfkai|標楷體|pingfang|苹方|hei(ti)? (sc|tc)|song(ti)? (sc|tc)|stsong|stkaiti|stfangsong|stxihei|华文|noto (sans|serif) (cjk )?(sc|tc|hk))/i, "zh"],
];

export function eastAsianFaceScript(family: string): "zh" | "ja" | "ko" | null {
  const name = family.trim().replace(/^["']|["']$/g, "");
  for (const [pattern, script] of EAST_ASIAN_FACE_SCRIPTS) if (pattern.test(name)) return script;
  return null;
}

/**
 * The document's East Asian faces in the order Word would reach for them
 * for `text`: the faces of the text's own script first (Kana -> Japanese,
 * Hangul -> Korean, Han alone -> Chinese, which is where Word's theme sends
 * Han under an English or Chinese document), the other scripts' faces after,
 * faces of no known script last. {@link normalizeEastAsianFontFamilies}
 * sorts by name, which put MS Gothic before 宋体 for a Chinese label.
 */
export function orderEastAsianFontFamiliesForText(families: readonly string[], text: string): string[] {
  const script = /[\u3040-\u30ff\u31f0-\u31ff]/u.test(text)
    ? "ja"
    : /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]/u.test(text)
      ? "ko"
      : "zh";
  const rank = (family: string) => {
    const own = eastAsianFaceScript(family);
    return own === script ? 0 : own === null ? 2 : 1;
  };
  return [...families].sort((left, right) => rank(left) - rank(right));
}
