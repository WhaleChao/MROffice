/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// A Calibri cell's Arabic drawn as the API's sheet writer prints it
// (notebooks.py complex_pieces / complex_advance): each word in the packaged
// Noto Sans Arabic at the cell's own size, spread or drawn in to the width
// Windows Calibri sets it at (officeArabicDots: its joined, kerned glyphs in
// whole pixels at the screen's pixels per em, or in whole printer dots on the
// print page). The cell stack's "MRAI Excel arabic 818" face set the letters
// 0.818 of their size instead (office-complex-scales.json): about as wide on
// average, but 18% short and 3-14% off word by word (lookwin p220-lw).
import { ensureOfficeScaledScriptFamily } from "@mrnote/core/web/lib/office-fonts/substitutes.ts";
import { officeComplexScaleFamily } from "@mrnote/core/web/lib/spreadsheet/office-complex-scales.ts";
import { officeArabicDots, officeArabicFace } from "@mrnote/core/web/lib/spreadsheet/office-arabic.ts";
import { awaitSpreadsheetFamily, nativeTypographyGeneration } from "@mrnote/core/web/lib/spreadsheet/native-typography.ts";

/** One piece of a cell's text: as it is, or an Arabic word set at Calibri's width (`width`, px at the text's size). */
export type OfficeArabicStretchPiece =
  | { readonly kind: "plain"; readonly text: string }
  | { readonly kind: "arabic"; readonly text: string; readonly width: number };

/** The letters (and their marks) the API's writer reads as Arabic (notebooks.py _complex_script_of). */
const ARABIC_WORD = /[\u0600-\u08ff\ufb50-\ufdff\ufe70-\ufeff]+/gu;

/** The face an Arabic word is drawn in before it is spread: Noto Sans Arabic at its size, the Arabic letters only. */
export const OFFICE_ARABIC_STRETCH_FAMILY = officeComplexScaleFamily("arabic", 1);

/**
 * `text` in pieces, each Arabic word with the width Windows Calibri sets it
 * at, px, for a cell in `family` (bold or not) at `sizePx`; undefined when
 * the face sets no Arabic of its own (every face but Calibri) or the text has
 * none. `dpi` 96 is the screen's whole pixels; a printer's (600) is its whole
 * dots at its pixels per em, as Excel prints. A word holding a character the
 * table lacks stays plain.
 */
export function officeArabicStretchPieces(text: string, family: string | null | undefined, bold: boolean, sizePx: number,
  dpi = 96): OfficeArabicStretchPiece[] | undefined {
  if (!(sizePx > 0) || !/[\u0600-\u08ff\ufb50-\ufdff\ufe70-\ufeff]/u.test(text)) return undefined;
  const arabic = officeArabicFace(family, bold);
  if (!arabic) return undefined;
  const sizePt = sizePx * 72 / 96;
  const ppem = Math.floor(sizePt * dpi / 72 + 0.5);
  if (ppem <= 0) return undefined;
  const pieces: OfficeArabicStretchPiece[] = [];
  let at = 0;
  let stretched = false;
  const plain = (piece: string) => {
    if (!piece) return;
    const last = pieces[pieces.length - 1];
    if (last?.kind === "plain") pieces[pieces.length - 1] = { kind: "plain", text: last.text + piece };
    else pieces.push({ kind: "plain", text: piece });
  };
  for (const word of text.matchAll(ARABIC_WORD)) {
    const start = word.index ?? 0;
    plain(text.slice(at, start));
    const dots = officeArabicDots(word[0], arabic, ppem);
    if (dots === undefined || dots <= 0) plain(word[0]);
    else {
      pieces.push({ kind: "arabic", text: word[0], width: dots * 96 / dpi });
      stretched = true;
    }
    at = start + word[0].length;
  }
  plain(text.slice(at));
  return stretched ? pieces : undefined;
}

/**
 * The horizontal scale an Arabic word of `width` px is drawn at, its
 * `natural` width in {@link OFFICE_ARABIC_STRETCH_FAMILY} at the cell's size
 * known; undefined when it is not (the face not yet arrived: the cell's stack
 * draws it meanwhile).
 */
export function officeArabicStretchScale(width: number, natural: number | undefined): number | undefined {
  if (natural === undefined || !(natural > 0) || !(width > 0)) return undefined;
  return width / natural;
}

type NaturalMeasure = (word: string, font: string) => number | undefined;
let measureForTests: NaturalMeasure | undefined;
/** Tests: the natural width of a word in a CSS font (jsdom draws no text). */
export function setOfficeArabicNaturalMeasureForTests(measure: NaturalMeasure | undefined): void {
  measureForTests = measure;
  naturals.clear();
}

let context: CanvasRenderingContext2D | null | undefined;
const naturals = new Map<string, number | null>();
let naturalsGeneration = -1;

/**
 * `word`'s width in {@link OFFICE_ARABIC_STRETCH_FAMILY} at `sizePx`, px, as
 * the browser draws it; undefined until the face has arrived (asked for now,
 * and the text drawn again when it lands: awaitSpreadsheetFamily).
 */
export function officeArabicNaturalWidth(word: string, fontStyle: string, fontWeight: string | number, sizePx: number): number | undefined {
  const font = `${fontStyle} ${fontWeight} ${sizePx}px "${OFFICE_ARABIC_STRETCH_FAMILY}"`;
  if (measureForTests) return measureForTests(word, font);
  if (typeof document === "undefined") return undefined;
  const generation = nativeTypographyGeneration();
  if (generation !== naturalsGeneration) {
    naturals.clear();
    naturalsGeneration = generation;
  }
  const key = `${font}\u0000${word}`;
  const known = naturals.get(key);
  if (known !== undefined) return known ?? undefined;
  ensureOfficeScaledScriptFamily(OFFICE_ARABIC_STRETCH_FAMILY, "Noto Sans Arabic", 0x0627, "U+0600-08FF, U+FB50-FDFF, U+FE70-FEFF", "100.00%");
  if (context === undefined) context = document.createElement("canvas").getContext("2d");
  if (!context) return undefined;
  // The face draws the word when two different fallbacks behind it give the
  // same width (FontFaceSet.check says yes for a family not yet declared).
  const widths = ["serif", "monospace"].map((fallback) => {
    context!.font = `${font}, ${fallback}`;
    return context!.measureText(word).width;
  });
  const natural = Math.abs(widths[0] - widths[1]) < 0.01 && widths[0] > 0 ? widths[0] : null;
  if (natural === null) {
    awaitSpreadsheetFamily(OFFICE_ARABIC_STRETCH_FAMILY);
    void document.fonts?.load(font, word).catch(() => undefined);
  }
  if (naturals.size > 20_000) naturals.clear();
  naturals.set(key, natural);
  return natural ?? undefined;
}
