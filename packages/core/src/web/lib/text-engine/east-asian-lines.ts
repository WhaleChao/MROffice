/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where PowerPoint for Windows sets an East Asian face's single line, for the
 * faces its general rule misses, and how the paragraph mark's Latin face
 * moves the line of ideographs it ends.
 *
 * The rule (`office-text-body.ts` lineVerticals): ideographs keep at least a
 * fifth of the 1.2 em line under the baseline, the first baseline at
 * 1.2 x min(ascent share, 0.8) em. PowerPoint for Windows' print of look-fit
 * truth-probes/lookfit-cjk-faces-probe.pptx (lookfit-cjk-faces-probe-win.pdf,
 * 24 pt ideographs, single spacing; Windows is the reference where the two
 * prints differ) holds it within 0.005 em for MS Gothic / PGothic / Mincho,
 * SimSun, SimHei, KaiTi, FangSong, Batang, Gulim, Yu Gothic / Mincho and
 * DengXian, and sets these lower or higher, the same under Calibri, Arial and
 * Times New Roman (Meiryo the same on the Mac). These are every line's but the
 * one that holds the paragraph mark (`eastAsianMarkLine`).
 *
 * Within 0.005 em was close while the engine stood baselines on whole points;
 * on no grid (as Windows) it is 0.3 pt at 60 pt, so the faces the rule holds
 * within it are given their measured place too: -faces- at 24 pt, and
 * -size- (lookfit-cjk-size-probe-win.pdf, 10-115 pt) where it has the face
 * (its mean over nine sizes; MS Gothic 0.953-0.957, Yu Gothic 0.920-0.924,
 * Meiryo 0.870-0.876).
 */
export const EAST_ASIAN_FIRST_BASELINE_EM: Readonly<Record<string, number>> = {
  // The rule's 0.848 (win 2171/901): 0.873.
  meiryo: 0.873,
  "メイリオ": 0.873,
  // The rule's 0.918 (win 2017/619): 0.922; Yu Mincho (the rule's 0.928) 0.931.
  "yu gothic": 0.922,
  "游ゴシック": 0.922,
  "yu mincho": 0.931,
  "游明朝": 0.931,
  // The rule's 0.96: SimHei, KaiTi 0.957, FangSong 0.959, Batang, Gulim 0.956, DengXian (the rule's 0.933) 0.937.
  simhei: 0.957,
  "黑体": 0.957,
  kaiti: 0.957,
  "楷体": 0.957,
  fangsong: 0.959,
  "仿宋": 0.959,
  batang: 0.956,
  "바탕": 0.956,
  gulim: 0.956,
  "굴림": 0.956,
  dengxian: 0.937,
  "等线": 0.937,
  // Its UI member (win 2171/430): the rule's 0.96, 0.950.
  "meiryo ui": 0.95,
  // The rule's 0.96 (the fifth's floor): 0.942 - 0.949.
  "microsoft yahei": 0.942,
  "微软雅黑": 0.942,
  mingliu: 0.944,
  "細明體": 0.944,
  pmingliu: 0.944,
  "新細明體": 0.944,
  "malgun gothic": 0.946,
  "맑은 고딕": 0.946,
  "microsoft jhenghei": 0.949,
  "微軟正黑體": 0.949,
};
































const BOX_GROUPS: Readonly<Record<string, readonly string[]>> = {
  // win 220 / 36 of 256: msgothic.ttc, msmincho.ttc, simsun.ttc, SimHei, KaiTi, FangSong, Word's HGRSMP /
  // HGRSGU / HGRGE / HGRME.
  "ms-gothic-box": [
    "ms gothic", "ms pgothic", "ms ui gothic", "ms mincho", "ms pmincho", "simsun", "nsimsun",
    "ms ゴシック", "ms pゴシック", "ms 明朝", "ms p明朝", "宋体", "新宋体",
    
    // SimHei labels under a SimHei mark stand 1.031 em down in Windows' print, the box's
    // Latin line (0.431-0.437 em under their boxes' middle).
    "simhei", "黑体", "kaiti", "楷体", "fangsong", "仿宋",
    "hgmarugothicmpro", "hg丸ゴシックm-pro", "hgsoeikakugothicub", "hgpsoeikakugothicub", "hgssoeikakugothicub",
    "hg創英角ゴシックub", "hgp創英角ゴシックub", "hgs創英角ゴシックub", "hggothice", "hgpgothice", "hgsgothice",
    "hgゴシックe", "hgpゴシックe", "hgsゴシックe", "hgminchoe", "hgpminchoe", "hgsminchoe", "hg明朝e", "hgp明朝e", "hgs明朝e",
  ],
  // win 2171 / 901 of 2048 (not Meiryo UI's 2171 / 430).
  meiryo: ["meiryo", "メイリオ"],
  // win 2017 / 619 of 2048.
  "yu gothic": ["yu gothic", "游ゴシック"],
  // win 2167 / 536 of 2048.
  "microsoft yahei": ["microsoft yahei", "微软雅黑"],
  // win 1659 / 475 of 2048.
  dengxian: ["dengxian", "等线"],
  // win 2229 / 495 of 2048.
  "malgun gothic": ["malgun gothic", "맑은 고딕"],
  "microsoft jhenghei": ["microsoft jhenghei", "微軟正黑體"],
  calibri: ["calibri"],
  arial: ["arial"],
  "times new roman": ["times new roman"],
  "microsoft sans serif": ["microsoft sans serif"],
};

/**
 * The measured first baseline (em) of a win-box group's ideographs, for the
 * families its names list (`BOX_GROUPS`) that the table above does not: the
 * 220 / 36 box (MS Gothic, PGothic and Mincho 0.955 at 24 pt, 0.953-0.957
 * over -size-'s nine sizes, HG Maru Gothic M-PRO the same; SimSun 0.957), the
 * rule's 0.96.
 */
export const EAST_ASIAN_GROUP_FIRST_BASELINE_EM: Readonly<Record<string, number>> = {
  "ms-gothic-box": 0.956,
};

/** The Windows-measured first baseline (em) of a family's ideographs, by its name or its win box's group. */
export function eastAsianMeasuredFirstBaselineEm(family: string): number | undefined {
  const plain = family.normalize("NFKC").trim().toLowerCase();
  const named = EAST_ASIAN_FIRST_BASELINE_EM[plain] ?? EAST_ASIAN_FIRST_BASELINE_EM[family.trim().toLowerCase()];
  if (named !== undefined) return named;
  const group = GROUP_OF.get(plain);
  return group === undefined ? undefined : EAST_ASIAN_GROUP_FIRST_BASELINE_EM[group];
}

const GROUP_OF = new Map<string, string>();
for (const [group, names] of Object.entries(BOX_GROUPS)) for (const name of names) GROUP_OF.set(name, group);

/** The win-box group of a family name (any width or case), if it is one measured. */
export function eastAsianBoxGroup(family: string): string | undefined {
  return GROUP_OF.get(family.normalize("NFKC").trim().toLowerCase());
}

/**
 * Em the paragraph mark's line of ideographs moves, by East Asian face group
 * then the mark's Latin face group (first baseline of a one-line 24 pt
 * paragraph, less the same under a Calibri mark).
 */
export const EAST_ASIAN_MARK_SHIFT_EM: Readonly<Record<string, Readonly<Record<string, number>>>> = {
  // (Microsoft Sans Serif: -malgun-mark-'s MS Gothic row, 0.992 / 0.989 over 0.959.)
  "ms-gothic-box": { arial: 0.03, "times new roman": 0.025, meiryo: -0.06, "yu gothic": -0.01, "microsoft sans serif": 0.032 },
  meiryo: { arial: 0.01, "times new roman": 0.01, "ms-gothic-box": 0.025, "yu gothic": -0.005 },
  "yu gothic": { arial: 0.01, "times new roman": 0.01, "ms-gothic-box": 0.025, meiryo: -0.055 },
  "microsoft yahei": { arial: 0.025, "times new roman": 0.025, "ms-gothic-box": 0.035, meiryo: -0.06, "yu gothic": -0.01 },
  dengxian: { arial: 0.005, "times new roman": 0.005, "ms-gothic-box": 0.02, meiryo: -0.065, "yu gothic": -0.015 },
  // Malgun Gothic's Latin marks over Hangul in ko-KR (look-fit truth-probes lookfit-malgun-mark-probe-win.pdf,
  // 18 and 24 pt, compatLnSpc either way): Arial 0.973 / 0.975 em, Times New Roman 0.967 / 0.970, Microsoft
  // Sans Serif 0.980 over the Calibri mark's 0.946 / 0.945 -- -markline-'s 0.02 / 0.02 were measured on Han,
  
  "malgun gothic": { arial: 0.028, "times new roman": 0.023, "microsoft sans serif": 0.035, "ms-gothic-box": 0.03, meiryo: -0.06, "yu gothic": -0.015 },
  "microsoft jhenghei": { arial: 0.025, "times new roman": 0.02, "ms-gothic-box": 0.035, meiryo: -0.065, "yu gothic": -0.015 },
};

/**
 * What the paragraph mark in `markFace` does to its line of ideographs in
 * `face` (family names): lay it as the face's Latin text, move it by a
 * measured em, or nothing (a pair not measured, or a Calibri mark).
 */
export function eastAsianMarkLine(face: string, markFace: string): { latin: true } | { shiftEm: number } | null {
  const plain = (name: string) => name.normalize("NFKC").trim().toLowerCase();
  const group = eastAsianBoxGroup(face);
  const markGroup = eastAsianBoxGroup(markFace);
  if (plain(face) === plain(markFace) || (group !== undefined && group === markGroup)) return { latin: true };
  if (group === undefined || markGroup === undefined) return null;
  const shift = EAST_ASIAN_MARK_SHIFT_EM[group]?.[markGroup];
  return shift === undefined ? null : { shiftEm: shift };
}
