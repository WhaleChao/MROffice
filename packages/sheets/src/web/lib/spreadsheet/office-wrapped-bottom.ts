/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where Windows Excel prints a WRAPPED cell's lines (lookwin p121b, p122,
 * p122b and p122c-lw), the same rules as the API's office_rows
 * (office_printed_descent_dots, office_bottom_rise_dots,
 * office_kept_wrapped_lines; office-row-rules.cases.json holds the cases both
 * answer alike):
 *
 * - a bottom-aligned line's baseline stands (D + 1) px of the printed row
 *   (1.1232 of its dot pitch) above the row's lower edge less 3.8 dots, D the
 *   row's deepest screen descent; a wrapped line D px more less its printer
 *   descent and a dot: small wrapped text prints higher than the same text
 *   unwrapped (Arial 8 in a Calibri 11 sheet 41 dots against 30);
 * - a row too low for all the lines keeps the first ones, down to the line
 *   that reaches its height, and aligns those as the cell says (two Calibri
 *   11 lines in a 22-40 px row still end on the bottom);
 * - a top line stands on its printer ascent and leading, a centred line's
 *   printer box in the middle, and all printed text OFFICE_PRINTED_TEXT_DROP_DOTS
 *   lower than the page-scaled layout (p123/p123b-lw; office_printed_top_dots,
 *   office_printed_box_dots).
 */
import OFFICE_ROW_METRICS from "@mrnote/core/web/lib/spreadsheet/office-row-metrics.json";
import OFFICE_PRINTED_GDI from "./office-printed-gdi.json";
import { excelWindowsFaceKey } from "@mrnote/core/web/lib/office-fonts/excel-windows-faces.ts";
import type { OfficeRowFont } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

/** The printer descent, em, where office-row-metrics' descent does not give
 * it (p122/p122c-lw against office_bottom_rise_dots; 맑은 고딕's and 游明朝's a floor, Segoe UI from 9-11 pt). The API's
 * _PRINTED_DESCENT_EM is the same table. */
const PRINTED_DESCENT_EM: Readonly<Record<string, number>> = {
  "游ゴシック": 0.41, "游ゴシック light": 0.41, "游ゴシック medium": 0.41, "游明朝": 0.41, "游明朝 light": 0.41,
  "游明朝 demibold": 0.41, "yu gothic": 0.41, "yu gothic light": 0.41, "yu gothic medium": 0.41, "yu mincho": 0.41,
  "meiryo ui": 0.267, "franklin gothic book": 0.2935, "맑은 고딕": 0.31, "malgun gothic": 0.31,
  "lucida sans unicode": 0.209, "trebuchet ms": 0.247, "segoe ui": 0.32, "courier new": 0.27,
  "微软雅黑": 0.28, "microsoft yahei": 0.28, "メイリオ": 0.472, "meiryo": 0.472,
};

function ownKey(family: string | undefined): string {
  return typeof family === "string" ? family.trim().toLowerCase() : "";
}

/** The printer descent of a face whose hinted small sizes print deeper than
 * its outline's, em, up to the printer pixels per em named (p147-lw; the API's
 * _PRINTED_DESCENT_SMALL_EM, the same table). */
const PRINTED_DESCENT_SMALL_EM: Readonly<Record<string, readonly [number, number]>> = {
  arial: [0.219, 200], "times new roman": [0.239, 150],
};

/** The deepest printer descent of one line's faces, 600-dpi dots at 100 %;
 * undefined for no face or a face the metrics lack. */
export function officePrintedDescentDots(fonts: readonly OfficeRowFont[]): number | undefined {
  let deepest: number | undefined;
  for (const font of fonts) {
    // A script face's line stands on the screen's line (R79-X23; the API alike).
    if (font.script) return undefined;
    const size = font.sizePt;
    if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return undefined;
    const key = excelWindowsFaceKey(font.family, font.record);
    let em = PRINTED_DESCENT_EM[ownKey(font.family)] ?? PRINTED_DESCENT_EM[key];
    const small = PRINTED_DESCENT_SMALL_EM[key];
    if (em === undefined && small !== undefined && (size * 600) / 72 <= small[1]) em = small[0];
    if (em === undefined) {
      const metrics = (OFFICE_ROW_METRICS as Readonly<Record<string, readonly number[]>>)[key];
      const descent = metrics?.[1];
      const units = metrics?.[3];
      if (descent === undefined || !units) return undefined;
      em = descent / units;
    }
    const dots = (em * size * 600) / 72;
    deepest = deepest === undefined ? dots : Math.max(deepest, dots);
  }
  return deepest;
}

/** Where office-row-metrics' (ascent + leading) does not give a top line's
 * printer ascent and external leading, em (lookwin p123/p123b-lw; the API's
 * _PRINTED_TOP_EM). */
const PRINTED_TOP_EM: Readonly<Record<string, number>> = {
  "游ゴシック": 1.1427, "游ゴシック light": 1.1427, "游ゴシック medium": 1.1427, "游明朝": 1.1427, "游明朝 light": 1.1427,
  "游明朝 demibold": 1.1427, "yu gothic": 1.1427, "yu gothic light": 1.1427, "yu gothic medium": 1.1427, "yu mincho": 1.1427,
  "times new roman": 0.933,
  // The MS Gothic and Mincho family print 1.5-2 dots lower than their ascent gives (p123/p123b-lw; R88-X11).
  "ｍｓ ゴシック": 0.881, "ｍｓ ｐゴシック": 0.881, "ｍｓ 明朝": 0.881, "ｍｓ ｐ明朝": 0.881, "ms gothic": 0.881, "ms pgothic": 0.881,
  "ms mincho": 0.881, "ms pmincho": 0.881,
};

/** Windows Excel's printed text stands this many 600-dpi dots lower than its
 * page-scaled layout, at every scale (lookwin p123b-lw; the API's
 * OFFICE_PRINTED_TEXT_DROP_DOTS): a top line 4.3 dots below its printer
 * ascent and leading, a bottom line's rise 4.3 x (1 - the page's scale) less. */
export const OFFICE_PRINTED_TEXT_DROP_DOTS = 4.3;

/** A top-aligned line's printer ascent and external leading, 600-dpi dots at
 * 100 % (the tallest of its faces); undefined for no face or a face the
 * metrics lack. */
export function officePrintedTopDots(fonts: readonly OfficeRowFont[]): number | undefined {
  let tallest: number | undefined;
  for (const font of fonts) {
    // A script face's line stands on the screen's line (R79-X23; the API alike).
    if (font.script) return undefined;
    const size = font.sizePt;
    if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return undefined;
    const key = excelWindowsFaceKey(font.family, font.record);
    let em = PRINTED_TOP_EM[ownKey(font.family)] ?? PRINTED_TOP_EM[key];
    if (em === undefined) {
      const metrics = (OFFICE_ROW_METRICS as Readonly<Record<string, readonly number[]>>)[key];
      const ascent = metrics?.[0];
      const leading = metrics?.[2];
      const units = metrics?.[3];
      if (ascent === undefined || leading === undefined || !units) return undefined;
      em = (ascent + leading) / units;
    }
    const dots = (em * size * 600) / 72;
    tallest = tallest === undefined ? dots : Math.max(tallest, dots);
  }
  return tallest;
}

/** A top-aligned line's baseline below its GDI line, and a centred line's
 * above the middle of its GDI box, 600-dpi dots at any page scale (lookwin
 * p148-lw; the API's OFFICE_PRINTED_GDI_TOP_DOTS / _MIDDLE_RISE_DOTS). */
export const OFFICE_PRINTED_GDI_TOP_DOTS = 3.5;
export const OFFICE_PRINTED_GDI_MIDDLE_RISE_DOTS = 0.75;
/** A merged block's centred lines (one line or a wrapped block, over one row
 * or several) print this many dots higher than an unmerged cell's, at any
 * face and page scale (lookwin p170/p170b-lw; the API's office_rows). */
export const OFFICE_PRINTED_MERGED_MIDDLE_RISE_DOTS = 4;

type PrintedGdiEntry = { u: number; a: number; d: number; l: number; v0?: number; vmax?: number[]; vmin?: number[] };

/**
 * One line's GDI box at the printer, 600-dpi dots on the paper at the page's
 * `scale`: its ascent with external leading and its descent, the tallest and
 * deepest of its faces. Windows Excel's printer sets a face at a whole ppem,
 * round(size x scale x 600 / 72); GDI reads the ascent and descent there from
 * the face's VDMX table (else its rounded win metrics) and adds the rounded
 * external leading (lookwin p148-lw: Arial 9 pt 67 + 2 dots, where the
 * metrics' 70.4 put the line 2 dots low). Undefined for a face the table
 * (office-printed-gdi.json, Latin faces) lacks, a script line or no face.
 * The API's office_printed_gdi_line_dots.
 */
export function officePrintedGdiLineDots(fonts: readonly OfficeRowFont[], scale = 1): { top: number; descent: number } | undefined {
  const table = OFFICE_PRINTED_GDI as unknown as Readonly<Record<string, { regular?: PrintedGdiEntry; bold?: PrintedGdiEntry }>>;
  let top: number | undefined;
  let descent: number | undefined;
  for (const font of fonts) {
    if (font.script) return undefined;
    const size = font.sizePt;
    if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409 || !(scale > 0)) return undefined;
    const styles = table[ownKey(font.family)] ?? table[excelWindowsFaceKey(font.family, font.record)];
    if (!styles) return undefined;
    const entry = font.bold === true && styles.bold ? styles.bold : styles.regular ?? styles.bold;
    if (!entry) return undefined;
    const ppem = Math.max(1, Math.floor((size * scale * 600) / 72 + 0.5));
    let ascent = Math.floor((entry.a * ppem) / entry.u + 0.5);
    let below = Math.floor((entry.d * ppem) / entry.u + 0.5);
    const at = entry.v0 === undefined ? -1 : ppem - entry.v0;
    const vmax = entry.vmax?.[at];
    const vmin = entry.vmin?.[at];
    if (at >= 0 && vmax !== undefined && vmin !== undefined) {
      ascent += vmax;
      below += vmin;
    }
    const line = ascent + Math.floor((entry.l * ppem) / entry.u + 0.5);
    top = top === undefined ? line : Math.max(top, line);
    descent = descent === undefined ? below : Math.max(descent, below);
  }
  return top === undefined || descent === undefined ? undefined : { top, descent };
}

/** One line's printer box, 600-dpi dots at 100 %: its ascent with leading and
 * its descent; a centred line stands it in the middle of its row. */
export function officePrintedBoxDots(fonts: readonly OfficeRowFont[]): { top: number; descent: number } | undefined {
  const top = officePrintedTopDots(fonts);
  const descent = officePrintedDescentDots(fonts);
  return top === undefined || descent === undefined ? undefined : { top, descent };
}

/** A bottom line's rise: this share of the printed row's dot pitch a screen
 * pixel, less OFFICE_BOTTOM_DROP_DOTS, at any page scale (lookwin
 * p121b/p122/p123b-lw; the API's OFFICE_BOTTOM_PX_SHARE). */
export const OFFICE_BOTTOM_PX_SHARE = 1.1232;
export const OFFICE_BOTTOM_DROP_DOTS = 3.8;

/** A bottom-aligned line's baseline above the row's lower edge, 600-dpi dots
 * on the page: `rowDescent` the row's deepest screen descent (px), `pxDots`
 * the printed row's dots a screen px at 100 % (6.25 x its printed row scale),
 * `scale` the page's; a wrapped line (`descentDots`, its printer descent at
 * 100 %) stands higher by the row's descent less that descent and a dot. The
 * API's office_bottom_rise_dots. */
/** Printer dots a screen pixel of a bottom line's rise takes, and the dots
 * added, by the sheet's Normal font (Excel's face key|size): a bottom line
 * stands s (D + 1) x the page's scale + c dots above its row's foot, a wrapped
 * one max(0, D s - its printer descent) x the scale higher still (lookwin
 * p147/p149-lw, 28 Normal fonts; the API's _BOTTOM_PITCH, the same table). A
 * Normal font the table lacks keeps the pitch model. */
const BOTTOM_PITCH: Readonly<Record<string, readonly [number, number]>> = {
  "calibri|11": [6.597, -2.48], "calibri|10": [7.516, -2.63], "arial|10": [6.516, -1.63], "arial|11": [7.0, -2.5],
  "arial|12": [7.516, -2.6], "arial|8": [5.786, -1.57], "aptos narrow|11": [6.8, -2.5], "ｍｓ ｐゴシック|11": [7.0, -2.5],
  "ms pgothic|11": [7.0, -2.5], "游ゴシック|11": [6.1, -1.3], "yu gothic|11": [6.1, -1.3], "等线|11": [7.24, -2.46],
  "dengxian|11": [7.24, -2.46], "times new roman|12": [6.403, -1.52], "verdana|10": [6.238, -1.45],
  "宋体|12": [7.3, -2.7], "simsun|12": [7.3, -2.7], "cambria|11": [7.0, -2.5],
  // p149-lw.
  "calibri|12": [7.0, -2.5], "calibri|9": [7.0, -2.5], "times new roman|10": [7.0, -2.5], "gill sans nova|11": [7.0, -2.5],
  "ｍｓ ゴシック|11": [7.0, -2.5], "ms gothic|11": [7.0, -2.5], "新細明體|12": [6.0, -1.5], "pmingliu|12": [6.0, -1.5],
  "arial|9": [6.0, -1.5], "tahoma|10": [6.239, -1.46], "segoe ui|9": [8.239, -3.46], "aptos|11": [6.754, -2.27],
  "맑은 고딕|11": [6.518, -1.61], "malgun gothic|11": [6.518, -1.61], "微软雅黑|11": [5.814, -1.38],
  "microsoft yahei|11": [5.814, -1.38], "aptos narrow|10": [6.403, -1.52],
};

/** (dots a screen px, dots added) of a bottom line's rise under the sheet's
 * Normal font, or undefined for one the table lacks (the API's office_bottom_pitch). */
export function officeBottomPitch(font: OfficeRowFont | null | undefined): readonly [number, number] | undefined {
  const size = font?.sizePt;
  if (!font || typeof font.family !== "string" || size === undefined || !(size > 0)) return undefined;
  const half = Math.floor(size * 2 + 0.5) / 2;
  for (const key of [font.family.trim().toLowerCase(), excelWindowsFaceKey(font.family, font.record)]) {
    const found = BOTTOM_PITCH[`${key}|${half}`];
    if (found) return found;
  }
  return undefined;
}

/** A scaled page's bottom line stands its Normal font's pitch dots added at
 * the page's scale and this many dots lower for the rest of the way to 0 %
 * (lookwin p172b-lw; the API's OFFICE_BOTTOM_SCALED_DROP_DOTS). */
export const OFFICE_BOTTOM_SCALED_DROP_DOTS = 3.4;

export function officeBottomRiseDots(rowDescent: number, pxDots: number, scale = 1, descentDots?: number,
  pitch?: readonly [number, number], gdiDescentDots?: number): number {
  if (pitch) {
    // Its dots added scale with the page, the rest of the way to
    // OFFICE_BOTTOM_SCALED_DROP_DOTS below (lookwin p172b-lw, as the API).
    let measured = (rowDescent + 1) * pitch[0] * scale + pitch[1] * scale - OFFICE_BOTTOM_SCALED_DROP_DOTS * (1 - scale);
    // A wrapped line's GDI descent at 100 % (officePrintedGdiLineDots): the
    // rise over the unwrapped line is the row's descent in the pitch's dots at
    // the page's scale, rounded, less that descent at the scale, floored
    // (lookwin p171-lw; the API's office_bottom_rise_dots).
    if (gdiDescentDots !== undefined) measured += Math.max(0, Math.round(rowDescent * pitch[0] * scale) - Math.floor(gdiDescentDots * scale + 1e-9));
    else if (descentDots !== undefined) measured += Math.max(0, (rowDescent * pitch[0] - descentDots) * scale);
    return measured;
  }
  let rise = (rowDescent + 1) * OFFICE_BOTTOM_PX_SHARE * pxDots * scale - OFFICE_BOTTOM_DROP_DOTS;
  if (descentDots !== undefined) rise += Math.max(0, (rowDescent * pxDots - descentDots + 1) * scale);
  return rise;
}

/** The dots a row the file marks thickBot raises its bottom-aligned text,
 * at 100 % (lookwin p134-lw; the API's OFFICE_THICK_BOTTOM_RISE_DOTS). */
export const OFFICE_THICK_BOTTOM_RISE_DOTS = 8;

/** How much higher Windows Excel prints a bottom-aligned line in a row the
 * file marks thickBot (`rowThickEdges` & 2), 600-dpi dots on the page: 8 at
 * 100 %, floored at a scale, twice that wrapped (p134-lw: 8 / 16 at 100 %,
 * 5 / 10 at 70 %). The API's office_thick_bottom_rise_dots. */
export function officeThickBottomRiseDots(scale = 1, wrapped = false): number {
  const rise = Math.floor(OFFICE_THICK_BOTTOM_RISE_DOTS * scale + 1e-9);
  return wrapped ? 2 * rise : rise;
}

/** The printer dots Windows Excel opens between repeated print titles whose
 * last row the file marks thickBot and the page's body, at any scale (lookwin
 * p132b/p134-lw: 8 at 100, 90 and 70 %; the API's OFFICE_TITLE_THICK_GAP_DOTS). */
export const OFFICE_TITLE_THICK_GAP_DOTS = 8;

/** The dots a merged block's wrapped bottom line rises less than an unmerged
 * cell's, over one row or several (lookwin p137-lw, p143-lw; the API's OFFICE_MERGED_WRAPPED_LESS_DOTS). */
export const OFFICE_MERGED_WRAPPED_LESS_DOTS = 8;

/** A merged block's wrapped-bottom extra: the unmerged cell's (the
 * descent term of officeBottomRiseDots) less 8 dots at any scale, never below
 * none (p137-lw: Arial 8 in a Calibri 11 sheet 3 dots where unmerged 11). The
 * API's office_merged_wrapped_extra_dots. */
export function officeMergedWrappedExtraDots(extra: number): number {
  return Math.max(0, extra - OFFICE_MERGED_WRAPPED_LESS_DOTS);
}

/** A merged block's bottom line stands its printer descent and this many dots
 * at the page's scale above its foot, less OFFICE_BOTTOM_DROP_DOTS (lookwin
 * p145-lw; the API's OFFICE_MERGED_ROWS_RISE_DOTS). */
export const OFFICE_MERGED_ROWS_RISE_DOTS = 8.6;

/** How far above its foot Windows Excel prints a merged block's bottom line,
 * 600-dpi dots on the page: no lower than its own printer descent and 8.6
 * dots at the page's scale less 3.8 put it, above `floorRise` -- the row's
 * own rise for a one-row block, the Normal font's for a block over several
 * rows, whose rows' faces do not enter (p145-lw: Calibri 24 over two rows
 * 58.5 dots where unmerged 63.5; p137-lw: Segoe UI 9 merged on one row of an
 * Arial 10 sheet 5 dots above the same text unmerged). The API's
 * office_merged_rows_rise_dots. */
export function officeMergedRowsRiseDots(floorRise: number, descentDots: number | undefined, scale = 1): number {
  if (descentDots === undefined) return floorRise;
  return Math.max(floorRise, (descentDots + OFFICE_MERGED_ROWS_RISE_DOTS) * scale - OFFICE_BOTTOM_DROP_DOTS);
}

/** How many of a wrapped cell's lines (their screen heights, px) a row of
 * `rowPx` keeps: the first ones, down to the line that reaches its height. */
export function officeKeptWrappedLines(lineHeights: readonly number[], rowPx: number): number {
  let total = 0;
  let kept = 0;
  for (const height of lineHeights) {
    if (kept && total >= rowPx - 0.01) break;
    total += height;
    kept += 1;
  }
  return kept;
}

/** Printed, how far (px) a line whose baseline the layout stands `pageTopPx`
 * below the paper's top edge moves to sit where Windows Excel prints it: on
 * the next whole printer dot (`dotPx` a dot) down from the page's top (every
 * baseline in Excel's printed PDFs is a whole 600-dpi dot; the API's writer
 * alike). The browser prints a baseline on the nearest whole pixel of its
 * page and a transform moves the painted line after that, so the shift is
 * from that pixel; none when it is under a thousandth of a pixel. */
/** Printed, how far (px) the layout moves a line whose baseline it stands
 * `pageTopPx` below the paper's top so that it stands on a whole pixel there:
 * the pixel officePrintedBaselineShiftPx counts from. Left where the sums put
 * it, a baseline within a hair of half a pixel was printed on the pixel above
 * the one counted from (the browser's own sums run a 64th of a pixel short):
 * a whole pixel, 0.75 pt, high (lookwin yu17 R450-X99: the centred numbers of
 * one row in eight). */
export function officePrintedBaselineSnapPx(pageTopPx: number): number {
  return Number.isFinite(pageTopPx) ? Math.round(pageTopPx) - pageTopPx : 0;
}

export function officePrintedBaselineShiftPx(pageTopPx: number, dotPx: number): number | undefined {
  if (!Number.isFinite(pageTopPx) || !(dotPx > 0)) return undefined;
  const shift = Math.ceil(pageTopPx / dotPx - 1e-6) * dotPx - Math.round(pageTopPx);
  return Math.abs(shift) > 1e-3 ? shift : undefined;
}
