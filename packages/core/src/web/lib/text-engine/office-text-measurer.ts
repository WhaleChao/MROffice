/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { splitFamilyList } from "./face-index";
import { officeFontMetricMatchesWidths, officeFontSubstitution } from "../office-fonts/substitutes";
import { measureText, resolveFace, textEngineEpoch, type EngineFont } from "./measure";
import type { OfficeFaceVerticals, OfficeMeasureAnswer, OfficeMeasuredGrapheme, OfficeTextFont, OfficeTextMeasurer } from "./office-text-body";

export const PT_PER_PX = 0.75;











const SYNTHETIC_BOLD_EXTRA_EM = 0.02;

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
 * Whether the face drawn for this family is a stand-in of another design
 * (catalog metric "style" or looser): its own USE_TYPO_METRICS says nothing
 * of the Office file's, so its win box stays the estimate.
 */
function drawnByOtherDesign(stack: string): boolean {
  const metric = officeFontSubstitution(splitFamilyList(stack)[0] ?? "")?.metric;
  return metric !== undefined && !officeFontMetricMatchesWidths(metric);
}

/** The Office file's win box in ems, when the family's stand-in does not carry it. */
export function officeWinBox(stack: string, weight: number): { ascent: number; descent: number } | null {
  const family = splitFamilyList(stack)[0] ?? "";
  const office = officeFontSubstitution(family)?.office;
  const box = OFFICE_WIN_BOX[family.trim().toLowerCase()] ?? (office ? OFFICE_WIN_BOX[office.toLowerCase()] : undefined);
  if (!box) return null;
  const [ascent, descent] = weight >= 600 && box.bold ? box.bold : box.regular;
  return { ascent: ascent / box.unitsPerEm, descent: descent / box.unitsPerEm };
}

const verticalsCache = new Map<string, OfficeMeasureAnswer<OfficeFaceVerticals>>();

export const pageOfficeTextMeasurer: OfficeTextMeasurer = {
  graphemes(text, font): OfficeMeasureAnswer<readonly OfficeMeasuredGrapheme[]> {
    const measured = measureText(text, engineFont(font));
    if (measured.status !== "ok") return measured;
    return {
      status: "ok",
      value: measured.graphemes.map((grapheme) => ({
        text: grapheme.text,
        advance: grapheme.advance * PT_PER_PX + (grapheme.syntheticBold && !/^\s+$/u.test(grapheme.text) ? SYNTHETIC_BOLD_EXTRA_EM * font.sizePt : 0),
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
      answer = { status: "ok", value: { ascent, descent } };
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
