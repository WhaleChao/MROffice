/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { leadingFamily, splitFamilyList } from "./face-index";
import { officeFontMetricMatchesWidths, officeFontSubstitution } from "../office-fonts/substitutes";
import { measureText, resolveFace, textEngineEpoch, type EngineFont } from "./measure";
import type { OfficeFaceVerticals, OfficeMeasureAnswer, OfficeMeasuredGrapheme, OfficeTextFont, OfficeTextMeasurer } from "./office-text-body";

export const PT_PER_PX = 0.75;



















const SYNTHETIC_BOLD_STEP_PT = 0.12;
const SYNTHETIC_BOLD_STEP_FROM_PT = 16;











const SYNTHETIC_BOLD_STEP_FACES: ReadonlySet<string> = new Set([
  "simsun", "宋体", "simhei", "黑体", "kaiti", "楷体", "fangsong", "仿宋",
  "ms gothic", "ｍｓ ゴシック", "ms mincho", "ｍｓ 明朝",
  // MS PGothic, which Windows has in regular only (msgothic.ttc): the step from 16 pt (probe t20cjk, Windows'
  
  // 36 / 44 pt). Meiryo, Microsoft YaHei, Malgun Gothic and Yu Gothic have bold faces on Windows: not emboldened.
  "ms pgothic", "ｍｓ ｐゴシック",
]);
const SYNTHETIC_BOLD_EM_BY_FAMILY: ReadonlyMap<string, number> = new Map([
  ["batang", -0.0086],
  ["batangche", -0.0086],
  ["바탕", -0.0086],
  ["바탕체", -0.0086],
  ["gulim", -0.0086],
  ["gulimche", -0.0086],
  ["굴림", -0.0086],
  ["굴림체", -0.0086],
  ["calibri light", -0.0063],
  ["segoe ui light", 0],
  ["pmingliu", 0],
  ["新細明體", 0],
  
  
  ["dfkai-sb", 0],
  ["標楷體", 0],
  // Families named for a weight, which Windows has with no bold face (windows-family-styles.json), emboldened at
  // their regular advances, no GDI step at any size: each grapheme of the ",Bold" font in Windows PowerPoint's own
  
  
  
  
  ["open sans semibold", 0],
  ["montserrat semibold", 0],
  ["oswald semibold", 0],
  ["poppins semibold", 0],
  ["nunito sans semibold", 0],
  ["dosis semibold", 0],
  ["berlin sans fb demi", 0],
]);

/** The extra advance (pt) PowerPoint for Windows gives a grapheme of `stack`'s face set in a bold it synthesizes. */
export function powerPointSyntheticBoldExtraPt(stack: string, sizePt: number): number {
  const family = leadingFamily(stack).trim().replace(/^["']|["']$/g, "").toLowerCase();
  const em = SYNTHETIC_BOLD_EM_BY_FAMILY.get(family);
  if (em !== undefined) return em * sizePt;
  if (!SYNTHETIC_BOLD_STEP_FACES.has(family)) return 0;
  return sizePt >= SYNTHETIC_BOLD_STEP_FROM_PT ? SYNTHETIC_BOLD_STEP_PT : 0;
}

/**
 * An Office family the page declares behind a local()-only face (Arial's
 * `local("Arial")` for a machine that has it): where the machine does not,
 * the browser passes it over to the served face, and so does the engine
 * (`followLocalOnly`); where it does, the text is the browser's.
 */
function engineFont(font: OfficeTextFont): EngineFont {
  return { stack: font.stack, weight: font.weight, italic: font.italic, sizePx: font.sizePt / PT_PER_PX, followLocalOnly: true, ...(font.kerning ? { kerning: true } : {}) };
}












const OFFICE_WIN_BOX: Readonly<Record<string, { readonly unitsPerEm: number; readonly regular: readonly [number, number]; readonly bold?: readonly [number, number] }>> = {
  arial: { unitsPerEm: 2048, regular: [1854, 434] },
  georgia: { unitsPerEm: 2048, regular: [1878, 449] },
  helvetica: { unitsPerEm: 2048, regular: [1946, 461], bold: [1966, 475] },
  // The Japanese families' stand-in is Noto Sans JP (win 1160/288 of 1000,
  // 0.96 em of single-line ascent where Meiryo UI's is 1.00): Office for
  // Mac's own meiryo.ttc, msgothic.ttc and YuGothR.ttc, by the family the
  // document names (Meiryo UI's descent is not Meiryo's). P3-faces: every box
  // of the five at 100 % (and Meiryo UI's at 90 %) on these, 67 of 68.
  meiryo: { unitsPerEm: 2048, regular: [2171, 901] },
  "meiryo ui": { unitsPerEm: 2048, regular: [2171, 430] },
  "ms gothic": { unitsPerEm: 256, regular: [220, 36] },
  "ms pgothic": { unitsPerEm: 256, regular: [220, 36] },
  "ms ui gothic": { unitsPerEm: 256, regular: [220, 36] },
  "yu gothic": { unitsPerEm: 2048, regular: [2017, 619] },
  // Their Mincho members, and the Chinese and Korean families Office themes
  // name, whose stand-ins (Noto Sans / Serif SC, KR: win 1160/288 of 1000)
  // are another box: Office's own msmincho.ttc, simsun.ttc, SimHei, KaiTi
  // and FangSong (220/36 of 256, MS Gothic's), batang.ttc / gulim.ttc
  // (879/145 of 1024), malgun.ttf (2229/495), msyh.ttc (2167/536) and
  // Deng.ttf (1659/475), every weight alike. Latin words in them, and a line
  
  // SimHei numbers 1.031 em down (Windows' print, 0.432-0.437 em under their
  // boxes' middle), where Noto Sans SC's box put them at 0.96 (look-fit 13th,
  // ea13/center.py).
  "ms mincho": { unitsPerEm: 256, regular: [220, 36] },
  "ms pmincho": { unitsPerEm: 256, regular: [220, 36] },
  simsun: { unitsPerEm: 256, regular: [220, 36] },
  nsimsun: { unitsPerEm: 256, regular: [220, 36] },
  "宋体": { unitsPerEm: 256, regular: [220, 36] },
  "新宋体": { unitsPerEm: 256, regular: [220, 36] },
  simhei: { unitsPerEm: 256, regular: [220, 36] },
  "黑体": { unitsPerEm: 256, regular: [220, 36] },
  kaiti: { unitsPerEm: 256, regular: [220, 36] },
  "楷体": { unitsPerEm: 256, regular: [220, 36] },
  fangsong: { unitsPerEm: 256, regular: [220, 36] },
  "仿宋": { unitsPerEm: 256, regular: [220, 36] },
  batang: { unitsPerEm: 1024, regular: [879, 145] },
  "바탕": { unitsPerEm: 1024, regular: [879, 145] },
  gulim: { unitsPerEm: 1024, regular: [879, 145] },
  "굴림": { unitsPerEm: 1024, regular: [879, 145] },
  "malgun gothic": { unitsPerEm: 2048, regular: [2229, 495] },
  "맑은 고딕": { unitsPerEm: 2048, regular: [2229, 495] },
  "microsoft yahei": { unitsPerEm: 2048, regular: [2167, 536] },
  "微软雅黑": { unitsPerEm: 2048, regular: [2167, 536] },
  dengxian: { unitsPerEm: 2048, regular: [1659, 475] },
  "等线": { unitsPerEm: 2048, regular: [1659, 475] },
  // Office's cloud Source Sans Pro does not set USE_TYPO_METRICS (win 984/273
  // of 1000, every weight); the served Source Sans 3 does, on a typo box
  
  // high on it).
  "source sans pro": { unitsPerEm: 1000, regular: [984, 273] },
  "source sans pro black": { unitsPerEm: 1000, regular: [984, 273] },
  "source sans pro light": { unitsPerEm: 1000, regular: [984, 273] },
  "source sans pro semibold": { unitsPerEm: 1000, regular: [984, 273] },
  "source sans pro extralight": { unitsPerEm: 1000, regular: [984, 273] },
  // PowerPoint for Mac sets Avenir in macOS's Avenir.ttc (win 1000/325 of 1000,
  
  // Avenir Light titles 1.2 pt high on its win box, 0.8 low on its typo box).
  avenir: { unitsPerEm: 1000, regular: [1000, 325] },
  "avenir light": { unitsPerEm: 1000, regular: [1000, 325] },
  "avenir book": { unitsPerEm: 1000, regular: [1000, 325] },
  "avenir roman": { unitsPerEm: 1000, regular: [1000, 325] },
  "avenir medium": { unitsPerEm: 1000, regular: [1000, 325] },
  "avenir heavy": { unitsPerEm: 1000, regular: [1000, 325] },
  "avenir black": { unitsPerEm: 1000, regular: [1000, 366] },
};

/**
 * The Office files' external leading (hhea ascender - descender + line gap
 * beyond the win box), em, where the served stand-in's differs or is not
 * known to match: Office for Mac's arial.ttf 67/2048, times.ttf and
 * Calibri.ttf none (Liberation Serif / Tinos state an 87/2048 hhea gap).
 */
const OFFICE_EXTERNAL_LEADING: Readonly<Record<string, number>> = {
  arial: 67 / 2048,
  "times new roman": 0,
  calibri: 0,
};

/**
 * Whether the face drawn for this family is a stand-in of another design
 * (catalog metric "style" or looser): its own USE_TYPO_METRICS says nothing
 * of the Office file's, so its win box stays the estimate.
 */
function drawnByOtherDesign(stack: string): boolean {
  const metric = officeFontSubstitution(leadingFamily(stack))?.metric;
  return metric !== undefined && !officeFontMetricMatchesWidths(metric);
}

/** The Office file's win box in ems, when the family's stand-in does not carry it. */
export function officeWinBox(stack: string, weight: number): { ascent: number; descent: number } | null {
  const family = leadingFamily(stack);
  const office = officeFontSubstitution(family)?.office;
  const plain = family.trim().replace(/^["']|["']$/g, "").trim();
  // A full-width spelling (ＭＳ Ｐゴシック) is the same family.
  const box = OFFICE_WIN_BOX[plain.toLowerCase()] ?? OFFICE_WIN_BOX[plain.normalize("NFKC").toLowerCase()] ??
    (office ? OFFICE_WIN_BOX[office.toLowerCase()] : undefined);
  if (!box) return null;
  const [ascent, descent] = weight >= 600 && box.bold ? box.bold : box.regular;
  return { ascent: ascent / box.unitsPerEm, descent: descent / box.unitsPerEm };
}

const verticalsCache = new Map<string, OfficeMeasureAnswer<OfficeFaceVerticals>>();

export const pageOfficeTextMeasurer: OfficeTextMeasurer = {
  graphemes(text, font): OfficeMeasureAnswer<readonly OfficeMeasuredGrapheme[]> {
    const measured = measureText(text, engineFont(font));
    if (measured.status !== "ok") return measured;
    const boldExtra = measured.graphemes.some((grapheme) => grapheme.syntheticBold) ? powerPointSyntheticBoldExtraPt(font.stack, font.sizePt) : 0;
    return {
      status: "ok",
      value: measured.graphemes.map((grapheme) => ({
        text: grapheme.text,
        advance: grapheme.advance * PT_PER_PX + (grapheme.syntheticBold ? boldExtra : 0),
      })),
    };
  },
  verticals(font): OfficeMeasureAnswer<OfficeFaceVerticals> {
    const key = `${textEngineEpoch()}\u0000${font.stack}\u0000${font.weight}\u0000${font.italic ? 1 : 0}`;
    const known = verticalsCache.get(key);
    if (known && known.status !== "pending") return known;
    // The primary face: the one that draws a space (CSS's first available font).
    const resolved = resolveFace(splitFamilyList(font.stack), font.weight, font.italic, 0x20, true);
    let answer: OfficeMeasureAnswer<OfficeFaceVerticals>;
    if (resolved === "pending") answer = { status: "pending" };
    else if (typeof resolved === "string") answer = { status: "unsupported", reason: resolved };
    else {
      const { metrics, face } = resolved;
      const unit = face.sizeAdjust / metrics.unitsPerEm;
      const office = officeWinBox(font.stack, font.weight);
      const os2 = metrics.os2;
      const typo = os2 !== null && os2.useTypoMetrics && !drawnByOtherDesign(font.stack);
      const ascent = office ? office.ascent : (os2 ? (typo ? os2.typoAscender : os2.winAscent) : metrics.hhea.ascender) * unit;
      const descent = office ? office.descent : (os2 ? (typo ? -os2.typoDescender : os2.winDescent) : -metrics.hhea.descender) * unit;
      // External leading: the Office file's where known, else the served
      // file's -- its typo line gap where it lays in its typo box, else what
      // its hhea line stands beyond its win box.
      const known = OFFICE_EXTERNAL_LEADING[leadingFamily(font.stack).trim().toLowerCase()];
      const hheaLine = (metrics.hhea.ascender - metrics.hhea.descender + metrics.hhea.lineGap) * unit;
      const lineGap = known !== undefined
        ? known
        : typo && os2
          ? Math.max(0, os2.typoLineGap * unit)
          : os2
            ? Math.max(0, hheaLine - (os2.winAscent + os2.winDescent) * unit)
            : Math.max(0, metrics.hhea.lineGap * unit);
      answer = { status: "ok", value: { ascent, descent, lineGap } };
    }
    if (verticalsCache.size > 2_000) verticalsCache.clear();
    verticalsCache.set(key, answer);
    return answer;
  },
};

/** Tests: forget what the page measurer remembered. */
export function resetOfficeTextMeasurerForTests(): void {
  verticalsCache.clear();
}
