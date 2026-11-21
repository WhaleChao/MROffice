/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Windows Calibri's Arabic as Windows Excel prints and wraps it (lookwin
// p220-lw). Windows' Calibri has Arabic letters of its own, and Excel sets a
// Calibri cell's Arabic in them, joined and kerned by Uniscribe: each glyph
// at its advance plus the kerning after it, rounded half up to whole device
// units at the device's ppem. office-arabic-advances.json holds each
// character's joining form glyphs, the two-letter ligatures (lam-alef,
// seen-reh, feh-yeh...), the advances and the kerning classes, in font units;
// the API's writer reads the same file (office_arabic.py,
// office_arabic_advances.json).
import TABLE from "./office-arabic-advances.json";

interface ArabicChar {
  readonly j: string;
  readonly isol?: number;
  readonly init?: number;
  readonly medi?: number;
  readonly fina?: number;
}

interface ArabicWeight {
  readonly adv: Readonly<Record<string, number>>;
  readonly first: Readonly<Record<string, number>>;
  readonly second: Readonly<Record<string, number>>;
  readonly kern: Readonly<Record<string, Readonly<Record<string, number>>>>;
}

interface ArabicFace {
  readonly chars: Readonly<Record<string, ArabicChar>>;
  readonly pairs: Readonly<Record<string, Readonly<Record<string, string>>>>;
  readonly R: ArabicWeight;
  readonly B: ArabicWeight;
}

/** A face's table at one weight. */
export interface OfficeArabicTable {
  readonly chars: ArabicFace["chars"];
  readonly pairs: ArabicFace["pairs"];
  readonly weight: ArabicWeight;
}

const table = TABLE as unknown as { readonly upm: number; readonly faces: Readonly<Record<string, ArabicFace>> };

/** The table of a cell face whose own Arabic Excel prints (Calibri), by weight; undefined for any other face. */
export function officeArabicFace(family: string | null | undefined, bold: boolean): OfficeArabicTable | undefined {
  const face = table.faces[(family ?? "").trim().replace(/^["']|["']$/g, "").trim().toLowerCase()];
  if (!face) return undefined;
  return { chars: face.chars, pairs: face.pairs, weight: bold ? face.B : face.R };
}

const JOINS_LEFT = new Set(["D", "C"]);
const JOINS_RIGHT = new Set(["D", "R", "C"]);

function codeKey(character: string): string {
  return (character.codePointAt(0) ?? 0).toString(16).padStart(4, "0");
}

/**
 * Each base glyph of `text` as Excel sets it: [index of its first character
 * (UTF-16), its advance with the kerning before the next glyph, font units];
 * null units where the table lacks the character. Marks set no advance.
 */
export function officeArabicUnits(text: string, arabic: OfficeArabicTable): Array<[number, number | null]> {
  const letters: Array<[number, ArabicChar | undefined]> = [];
  for (let index = 0; index < text.length; index += 1) {
    const entry = arabic.chars[codeKey(text[index])];
    if (entry?.j === "T") continue;
    letters.push([index, entry]);
  }
  const glyphs: Array<[number, string | null]> = letters.map(([index, entry], order) => {
    if (!entry) return [index, null];
    const before = order > 0 ? letters[order - 1][1] : undefined;
    const after = order + 1 < letters.length ? letters[order + 1][1] : undefined;
    const right = JOINS_RIGHT.has(entry.j) && before !== undefined && JOINS_LEFT.has(before.j);
    const left = JOINS_LEFT.has(entry.j) && after !== undefined && JOINS_RIGHT.has(after.j);
    const form = right && left ? "medi" : right ? "fina" : left ? "init" : "isol";
    const glyph = entry[form] ?? entry.isol;
    return [index, glyph === undefined ? null : String(glyph)];
  });
  // Two joined glyphs Calibri sets as one (lam-alef, seen-reh, feh-yeh...).
  const merged: Array<[number, string | null]> = [];
  for (const [index, glyph] of glyphs) {
    const last = merged[merged.length - 1];
    const both = last && last[1] !== null && glyph !== null ? arabic.pairs[last[1]]?.[glyph] : undefined;
    if (last && both !== undefined) last[1] = String(both);
    else merged.push([index, glyph]);
  }
  const { adv, first, second, kern } = arabic.weight;
  return merged.map(([index, glyph], place) => {
    if (glyph === null) return [index, null];
    let units = adv[glyph] ?? 0;
    const following = place + 1 < merged.length ? merged[place + 1][1] : null;
    if (following !== null && first[glyph] !== undefined && second[following] !== undefined) {
      units += kern[String(first[glyph])]?.[String(second[following])] ?? 0;
    }
    return [index, units];
  });
}

/** `text`'s width in whole device units at `ppem` (each glyph's advance and
 * kerning rounded half up); undefined where the table lacks a character. */
export function officeArabicDots(text: string, arabic: OfficeArabicTable, ppem: number): number | undefined {
  let dots = 0;
  for (const [, units] of officeArabicUnits(text, arabic)) {
    if (units === null) return undefined;
    dots += Math.floor(units * ppem / table.upm + 0.5);
  }
  return dots;
}

/**
 * Each character's share of `text`'s printed width, px at 96 dpi, as the
 * wrap counts it: a word (its letters joined and kerned, an Arabic letter
 * alone being its wider isolated form) on its first character, each space
 * its own; undefined where the face has no table or the table lacks a
 * character. At the screen's ppem in whole pixels, the lines that fit the
 * row (lookwin p235-lw: R463-X31's Calibri 12 sentences wrap where these
 * widths wrap them; the printer's dots read back stood 1-2% narrower), as the
 * API's writer wraps such a cell (notebooks.py complex_advance, arabic_screen).
 */
export function officeArabicWrapAdvances(text: string, family: string, bold: boolean, sizePt: number): number[] | undefined {
  const arabic = officeArabicFace(family, bold);
  if (!arabic || !(sizePt > 0) || /[\ud800-\udfff]/.test(text)) return undefined;
  const ppem = Math.floor(sizePt * 96 / 72 + 0.5);
  const out = new Array<number>(text.length).fill(0);
  for (const word of text.matchAll(/\S+|\s/g)) {
    const at = word.index ?? 0;
    if (word[0] === "\n") continue;
    const pixels = officeArabicDots(word[0], arabic, ppem);
    if (pixels === undefined) return undefined;
    out[at] = pixels;
  }
  return out;
}
