/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The names a Simplified Chinese system gives the Office families MRAI
 * serves (each font's own zh-CN `name` record): what WPS and Word list them
 * as in a Chinese interface (宋体, not SimSun). The family a write names is
 * still the Office one; only the list shows these.
 */

const ZH_HANS: Readonly<Record<string, string>> = Object.freeze({
  SimSun: "宋体",
  SimHei: "黑体",
  KaiTi: "楷体",
  FangSong: "仿宋",
  "Microsoft YaHei": "微软雅黑",
  "Microsoft YaHei Light": "微软雅黑 Light",
  DengXian: "等线",
  "DengXian Light": "等线 Light",
  STZhongsong: "华文中宋",
  STLiti: "华文隶书",
  STXihei: "华文细黑",
  STXingkai: "华文行楷",
});

/** The Office families' names in the interface's language, for the locales
 * whose systems name them otherwise; empty for the rest. */
export function officeFamilyDisplayNames(locale: string | null | undefined): Readonly<Record<string, string>> {
  return /^zh(?:-(?:Hans|CN|SG))?\b/i.test(locale ?? "") && !/^zh-(?:Hant|TW|HK|MO)\b/i.test(locale ?? "")
    ? ZH_HANS
    : {};
}

/** The name a Simplified Chinese system gives an Office family (SimHei → 黑体), or null. */
export function officeFamilyLocalizedName(family: string): string | null {
  return Object.prototype.hasOwnProperty.call(ZH_HANS, family) ? ZH_HANS[family] : null;
}

/** The Office family a displayed name stands for (宋体 → SimSun), else the
 * name itself: a family typed or pasted as the list shows it. */
export function officeFamilyFromDisplayName(name: string): string {
  for (const [family, shown] of Object.entries(ZH_HANS)) if (shown === name) return family;
  return name;
}

const HAN = /\p{Script=Han}/u;
let pinyin: Intl.Collator | null = null;

/** WPS's order for a Chinese interface: families named in Chinese first, in
 * pinyin order (等线, 仿宋, 黑体, 楷体, 宋体, 微软雅黑), then the rest. */
export function compareFontLabelsEastAsianFirst(left: string, right: string): number {
  const leftHan = HAN.test(left), rightHan = HAN.test(right);
  if (leftHan !== rightHan) return leftHan ? -1 : 1;
  if (leftHan) {
    pinyin ??= new Intl.Collator("zh-Hans-CN-u-co-pinyin");
    return pinyin.compare(left, right);
  }
  return left.localeCompare(right);
}
