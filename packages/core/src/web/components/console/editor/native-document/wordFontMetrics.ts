/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { wordEastAsianSingleLineFactor } from "../../../../lib/document-word-line-metrics";
import substitutionTable from "../../../../lib/office-fonts/substitutes.generated.json";
import wordLatinFamilies from "../../../../lib/office-fonts/word-latin-families.generated.json";
import wordLineRatios from "../../../../lib/office-fonts/word-line-ratios.generated.json";
import table from "./wordFontMetrics.json";

/**
 * Design metrics of the fonts Word lays text out with, read from the font
 * files themselves rather than from whichever face the browser substitutes.
 *
 * The browser often has no such font to measure: Aptos Mono is an Office
 * cloud face, and system faces are not loaded font faces a page can inspect.
 * A metric-compatible substitute paints the same advances but has its own
 * ascent and descent. Widths and line boxes Word derives from its font are
 * nevertheless facts of that font.
 */

interface RecordedFont {
  names: readonly string[];
  /** sha256 of the very file these metrics were read from. */
  sha256: string;
  unitsPerEm: number;
  /** hhea ascender, descender (negative) and line gap, in font units. */
  ascender: number;
  descender: number;
  lineGap: number;
  /** The face declares an East Asian code page (OS/2 ulCodePageRange1). */
  eastAsian?: boolean;
  advances: Readonly<Record<string, number>>;
}

const byName = new Map<string, RecordedFont>();
const bySha = new Map<string, RecordedFont>();
for (const font of table.fonts as readonly RecordedFont[]) {
  for (const name of font.names) byName.set(name.toLowerCase(), font);
  bySha.set(font.sha256.toLowerCase(), font);
}

/** A page paints a bound face under the family name its file hashes to
 * (`MRAI-Font-<sha256>…`, one hash per role file), so the declared name is
 * gone from the computed stack. The recorded file is the very file being
 * painted when one of those hashes is its own. */
/**
 * The recorded Office face a metric-compatible open face stands in for.
 * Word, lacking an open face a document names, sets it in the Office face
 * with the same metrics: R36-D29's Carlito 11pt footer comes out of Word in
 * Calibri, 13.4pt a line, where reading the CJK fallback at the end of the
 * render stack gave 15.9pt and lifted the body a line too high. A face is
 * followed through its own exact substitute (Liberation Serif -> Tinos ->
 * Times New Roman); the recorded Office face named most often wins
 * (Cousine -> Courier New).
 */
const exactOfficeFor = new Map<string, RecordedFont | null>();
function recordedThroughExactSubstitute(name: string): RecordedFont | undefined {
  const key = name.toLowerCase();
  const known = exactOfficeFor.get(key);
  if (known !== undefined) return known ?? undefined;
  const entries = Object.values(substitutionTable as Record<string, { office: string; substitutes: readonly string[]; metric: string; weight?: number }>);
  const exact = entries.filter((entry) => (entry.metric === "exact" || entry.metric === "cloned") && entry.weight === undefined);
  const faces = new Set([key]);
  for (const entry of exact) {
    if (entry.office.toLowerCase() === key) for (const substitute of entry.substitutes) faces.add(substitute.toLowerCase());
  }
  const votes = new Map<RecordedFont, number>();
  for (const entry of exact) {
    if (!entry.substitutes.some((substitute) => faces.has(substitute.toLowerCase()))) continue;
    const font = byName.get(entry.office.toLowerCase());
    if (font) votes.set(font, (votes.get(font) ?? 0) + 1);
  }
  let best: RecordedFont | null = null;
  let bestVotes = 0;
  for (const [font, count] of votes) {
    if (count > bestVotes) {
      best = font;
      bestVotes = count;
    }
  }
  exactOfficeFor.set(key, best);
  return best ?? undefined;
}

const recorded = (family: string) => {
  const name = family.trim().replace(/^["']|["']$/g, "");
  const direct = byName.get(name.toLowerCase());
  if (direct) return direct;
  const alias = /^MRAI-Font-((?:[0-9a-f]{64})(?:-[0-9a-f]{64})*)$/i.exec(name);
  if (!alias) return recordedThroughExactSubstitute(name);
  for (const sha of alias[1].toLowerCase().split("-")) {
    const font = bySha.get(sha);
    if (font) return font;
  }
  return undefined;
};

/**
 * Whether Word takes this face as an East Asian font: a property of the font
 * file (its OS/2 code page range), which Word reads from the face itself
 * when the document's font table does not list it.
 */
export function wordFontIsEastAsian(family: string): boolean {
  return recorded(family)?.eastAsian === true;
}

/**
 * Whether this family names a recorded font file that declares no East Asian
 * code page. Word reads that from the installed face itself: Calibri named in
 * a document's `w:eastAsia` slot with no signed font-table entry still runs
 * its Latin single line (word-stops-floats2.docx: 13.4pt at 11pt, not 1.3x).
 */
export function wordFontIsRecordedLatin(family: string): boolean {
  const font = recorded(family);
  if (font !== undefined) return font.eastAsian !== true;
  // Word's other Latin faces, read from its own files (word-vertical-metrics,
  // scripts/measure-word-vertical-metrics.py): Franklin Gothic Book in the
  // `w:eastAsia` slot runs its Latin line, 12.47pt at 11pt, not 1.3x that
  // (corpus R83-D07, word-stops-font-heights).
  return WORD_LATIN_FAMILIES.has(family.trim().toLowerCase());
}

const WORD_LATIN_FAMILIES: ReadonlySet<string> = new Set(
  (wordLatinFamilies as readonly string[]).map((name) => name.toLowerCase()),
);

/** Recorded East Asian families Word installs as one regular face. */
const REGULAR_ONLY_FAMILIES = new Set(["simsun", "simhei", "kaiti", "fangsong", "ms mincho", "ms pmincho", "ms gothic", "ms ui gothic", "ms pgothic"]);

/**
 * Whether Word has only the regular face of this recorded family (SimSun,
 * SimHei, KaiTi, FangSong, the MS Mincho and Gothic faces). Word sets a bold
 * or italic run in that face and emboldens or slants it, so the run keeps
 * the regular face's metrics.
 */
export function wordFontShipsRegularOnly(family: string): boolean {
  const font = recorded(family);
  return font !== undefined && font.names.some((name) => REGULAR_ONLY_FAMILIES.has(name.toLowerCase()));
}

/** Every family name of a recorded East Asian face, lower-cased. */
export const WORD_EAST_ASIAN_RECORDED_FAMILIES: ReadonlySet<string> = new Set(
  (table.fonts as readonly RecordedFont[])
    .filter((font) => font.eastAsian === true)
    .flatMap((font) => font.names.map((name) => name.toLowerCase())),
);

/** Whether this computed family is a face the page bound from a recorded
 * Word font file, which the page paints and Word measured. */
export function wordFontFileIsRecorded(family: string): boolean {
  return /^MRAI-Font-/i.test(family.trim().replace(/^["']|["']$/g, "")) && recorded(family) !== undefined;
}

/**
 * The width Word gives a decimal numbering label, in CSS px, or null when the
 * font or a glyph is not recorded.
 *
 * Word ends the label, then its tab suffix advances to the next usable tab
 * stop; a label wider than the hanging indent pushes the first line of text
 * to the next default stop ("27." in 9pt Aptos ends past a 9pt hang, so the
 * text starts 36pt from the margin instead of 16.45pt).
 */
export function wordNumberingLabelAdvancePx(
  family: string,
  label: string,
  fontSizePx: number,
): number | null {
  const font = recorded(family);
  if (!font || !(fontSizePx > 0) || !label) return null;
  let units = 0;
  for (const character of label) {
    const advance = font.advances[character];
    if (typeof advance !== "number") return null;
    units += advance;
  }
  return (units * fontSizePx) / font.unitsPerEm;
}

/**
 * The regular face's line metrics per em, or null when the font is not
 * recorded. `box` is ascender minus descender; `lineGap` is kept apart
 * because Word's East Asian single line excludes it.
 */
export function wordFontLineMetrics(
  family: string,
): { box: number; lineGap: number } | null {
  const font = recorded(family);
  if (!font) return null;
  return {
    box: (font.ascender - font.descender) / font.unitsPerEm,
    lineGap: font.lineGap / font.unitsPerEm,
  };
}










export function wordGridFontLineMetrics(family: string): { box: number; lineGap: number } | null {
  const own = wordFileLineMetrics(family);
  if (own) return own;
  // An alias of an Office family ("Times" for Times New Roman) is that family.
  const name = family.trim().replace(/^["']|["']$/g, "").toLowerCase();
  const office = (substitutionTable as Record<string, { office?: string }>)[name]?.office;
  return office && office.toLowerCase() !== name ? wordFileLineMetrics(office) : null;
}

/**
 * The single line of `family` as Word's own file sets it: a recorded font's,
 * else the ratio read from Word's file (`word-line-ratios.generated.json`),
 * or null when Word's file was never read.
 */
export function wordFileLineMetrics(family: string): { box: number; lineGap: number } | null {
  const recordedMetrics = wordFontLineMetrics(family);
  if (recordedMetrics) return recordedMetrics;
  const name = family.trim().replace(/^["']|["']$/g, "").toLowerCase();
  const ratio = (wordLineRatios as Readonly<Record<string, readonly number[]>>)[name];
  return ratio ? { box: ratio[0], lineGap: ratio[1] } : null;
}













/**
 * How far below the baseline Word extends a line for a run set in this face,
 * per em, or null when the font is not recorded: the descender, plus half of
 * an East Asian face's extra 0.3 box. A Latin face's line gap is all above.
 */
export function wordFontLineDescent(family: string, eastAsian: boolean): number | null {
  const font = recorded(family);
  if (!font) return null;
  const descent = -font.descender / font.unitsPerEm;
  if (!eastAsian) return descent;
  const box = (font.ascender - font.descender) / font.unitsPerEm;
  return descent + (box * (wordEastAsianSingleLineFactor(family) - 1)) / 2;
}









const gridVerticals = new Map<string, { ascent: number; descent: number } | null>();
export function wordGridFontVerticals(family: string): { ascent: number; descent: number } | null {
  const name = family.trim().replace(/^["']|["']$/g, "");
  const key = name.toLowerCase();
  const known = gridVerticals.get(key);
  if (known !== undefined) return known;
  const office = (substitutionTable as Record<string, { office?: string }>)[name]?.office;
  const face = recorded(name) ? name : office && recorded(office) ? office : null;
  let verticals: { ascent: number; descent: number } | null = null;
  if (face) {
    const eastAsian = wordFontIsEastAsian(face);
    const ascent = wordFontLineAscent(face, eastAsian);
    const descent = wordFontLineDescent(face, eastAsian);
    if (ascent !== null && descent !== null) verticals = { ascent, descent };
  } else {
    // Word's own file (word-line-ratios: box, line gap, East Asian, ascender).
    const ratios = wordLineRatios as Readonly<Record<string, readonly number[]>>;
    const ratio = ratios[key] ?? (office ? ratios[office.toLowerCase()] : undefined);
    if (ratio && ratio.length >= 4) {
      const [box, lineGap, eastAsianFlag, ascender] = ratio;
      if (eastAsianFlag === 1) {
        const half = (box * (wordEastAsianSingleLineFactor(office ?? name) - 1)) / 2;
        verticals = { ascent: ascender + half, descent: box - ascender + half };
      } else {
        verticals = { ascent: ascender + lineGap, descent: box - ascender };
      }
    }
  }
  gridVerticals.set(key, verticals);
  return verticals;
}

export function wordFontLineAscent(family: string, eastAsian: boolean): number | null {
  const font = recorded(family);
  if (!font) return null;
  const ascent = font.ascender / font.unitsPerEm;
  if (!eastAsian) return ascent + font.lineGap / font.unitsPerEm;
  const box = (font.ascender - font.descender) / font.unitsPerEm;
  return ascent + (box * (wordEastAsianSingleLineFactor(family) - 1)) / 2;
}
