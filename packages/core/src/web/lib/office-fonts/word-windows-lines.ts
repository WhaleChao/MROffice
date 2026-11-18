/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import table from "./word-windows-lines.generated.json";



























interface WindowsLines {
  /** Per family and "weight/italic": [ascent, descent, line gap], percent of the size. */
  readonly faces: Readonly<Record<string, Readonly<Record<string, readonly [number, number, number]>>>>;
  readonly ratios: Readonly<Record<string, readonly number[]>>;
  readonly eastAsianFactors: Readonly<Record<string, number>>;
}
const lines = table as unknown as WindowsLines;

/** `window.__docsWinMetricsOff = true` keeps the Mac box everywhere, for an A/B of one build. */
export function wordWindowsLinesOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsWinMetricsOff?: boolean }).__docsWinMetricsOff === true;
}

const key = (family: string) => family.trim().replace(/^["']|["']$/g, "").trim().toLowerCase();

/** The regular face's Windows line [box, line gap, 0, ascender] em (word-line-ratios' form), or undefined. */
export function wordWindowsLineRatio(family: string): readonly number[] | undefined {
  return wordWindowsLinesOff() ? undefined : lines.ratios[key(family)];
}

/** Windows Word's single-line factor over the recorded box of an East Asian face, or undefined. */
export function wordWindowsEastAsianFactor(family: string | null | undefined): number | undefined {
  return family && !wordWindowsLinesOff() ? lines.eastAsianFactors[key(family)] : undefined;
}

/** Windows Word's single line of a Latin `family` at this weight and slant, em, where it differs from the Mac one; or undefined. */
export function wordWindowsLineEm(family: string, weight: number, italic: boolean): number | undefined {
  if (wordWindowsLinesOff()) return undefined;
  const own = lines.faces[key(family)]?.[`${weight >= 600 ? 700 : 400}/${italic ? 1 : 0}`];
  return own ? (own[0] + own[1] + own[2]) / 100 : undefined;
}
