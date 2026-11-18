/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import OFFICE_DIGIT_ADVANCES from "@mrnote/core/web/lib/spreadsheet/office-digit-advances.json";
import OFFICE_DIGIT_WHOLE_EM from "@mrnote/core/web/lib/spreadsheet/office-digit-whole-em.json";
import OFFICE_PRINT_METRICS from "./office-print-metrics.json";
import OFFICE_ROW_METRICS from "@mrnote/core/web/lib/spreadsheet/office-row-metrics.json";
import { officeDigitWidth, officeRowHeight } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { excelWindowsFaceKey, type ExcelFontRecordClass } from "@mrnote/core/web/lib/office-fonts/excel-windows-faces.ts";





























export const OFFICE_PRINT_DPI = 600;
const DOTS_PER_SCREEN_PIXEL = OFFICE_PRINT_DPI / 96;

export interface OfficePrintScale {
  /** Printed width of a screen column pixel, in screen pixels. */
  readonly x: number;
  /** Printed height of a screen row pixel, in screen pixels. */
  readonly y: number;
}

export const OFFICE_PRINT_IDENTITY: OfficePrintScale = Object.freeze({ x: 1, y: 1 });

type MeasuredEntry = readonly [screenRowPx: number, printedRowDots: number, printedDigitDots: number];

/**
 * The Japanese MS faces print alike (one table, the server's
 * _OFFICE_MS_JAPANESE_PRINT_FACES): a face of them the table lacks a size for
 * reads its siblings' (R102-X10's Normal 明朝, drawn in ＭＳ Ｐ明朝 11: its
 * 15 px rows print 85 dots), and GDI's ASCII names read the Japanese ones.
 */
const MS_JAPANESE_PRINT_FACES = ["ｍｓ 明朝", "ｍｓ ｐ明朝", "ｍｓ ゴシック", "ｍｓ ｐゴシック"] as const;
const PRINT_FACE_ALIASES: Readonly<Record<string, string>> = {
  "ms mincho": "ｍｓ 明朝",
  "ms pmincho": "ｍｓ ｐ明朝",
  "ms gothic": "ｍｓ ゴシック",
  "ms pgothic": "ｍｓ ｐゴシック",
};

function measuredSizes(family: string): Readonly<Record<string, unknown>> | undefined {
  const table = OFFICE_PRINT_METRICS as unknown as Readonly<Record<string, unknown>>;
  const at = PRINT_FACE_ALIASES[family] ?? family;
  const own = table[at];
  const sizes = own && typeof own === "object" ? (own as Readonly<Record<string, unknown>>) : undefined;
  if (!(MS_JAPANESE_PRINT_FACES as readonly string[]).includes(at)) return sizes;
  const merged: Record<string, unknown> = {};
  for (const face of MS_JAPANESE_PRINT_FACES) {
    const sibling = table[face];
    if (face !== at && sibling && typeof sibling === "object") Object.assign(merged, sibling);
  }
  if (sizes) Object.assign(merged, sizes);
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function measured(family: string, sizePt: number): MeasuredEntry | undefined {
  const sizes = measuredSizes(family);
  if (!sizes) return undefined;
  const entry = sizes[String(Math.round(sizePt * 2) / 2)];
  return Array.isArray(entry) && entry.length === 3 && entry.every((value) => typeof value === "number" && value > 0)
    ? (entry as unknown as MeasuredEntry)
    : undefined;
}

/**
 * The printed default row of a face the probes did not measure, in printer
 * dots: the face's line box at the whole-dot em (as officeRowHeight builds
 * the screen row) plus 8 dots, what the measured faces carry above their line
 * box (6-10 dots for every face without a measured screen correction).
 */
function estimatedPrintedRowDots(family: string, sizePt: number): number {
  const metrics = OFFICE_ROW_METRICS as Readonly<Record<string, readonly number[]>>;
  const [ascent, descent, leading, unitsPerEm] = metrics[family] ?? metrics.calibri;
  const em = Math.ceil((sizePt * OFFICE_PRINT_DPI) / 72 - 1e-9);
  return (
    Math.ceil((ascent * em) / unitsPerEm) +
    Math.floor((2 * descent * em + unitsPerEm) / (2 * unitsPerEm)) +
    Math.floor((leading * em) / unitsPerEm) +
    8
  );
}

/** The printed maximum digit width of a face the probes did not measure, in
 * printer dots: the digit's advance at the whole-dot em (half the em for the
 * East Asian fixed-pitch faces), as officeDigitWidth counts it on screen. */
function estimatedPrintedDigitDots(family: string, sizePt: number): number {
  const em = Math.floor((sizePt * OFFICE_PRINT_DPI) / 72 + 0.5);
  if ((OFFICE_DIGIT_WHOLE_EM as readonly string[]).includes(family)) return Math.max(1, Math.floor(em / 2 + 0.5));
  const advances = OFFICE_DIGIT_ADVANCES as Readonly<Record<string, number>>;
  return Math.max(1, Math.round((advances[family] ?? advances.calibri) * em));
}

/**
 * How far Windows Excel sets a left- and a right-aligned printed line in from
 * its cell's edges, in printer dots on the paper: half the margin its printer
 * re-wraps in, m = round(digit / 2) + 7 (digit: the face's widest digit at the
 * printer's pixels per em at the page's `scale`), floor(m / 2) + 1 on the
 * left and ceil(m / 2) on the right (lookwin p141-lw: 8 faces at 8-48 pt, 100
 * and 70 %: Calibri 12 17 dots, 48 55, Arial Bold 36 46). The API's
 * draw_cell_text_upright inset_left / inset_right; `margin` is m itself, the
 * room a printed line re-wraps in being the printed column less it (the API's
 * cell_text_lines, lookwin p138-lw).
 */
export function officePrintedTextInsetDots(family: string, sizePt: number, scale = 1): { left: number; right: number; margin: number } | undefined {
  if (!(sizePt > 0) || !(scale > 0)) return undefined;
  const key = excelWindowsFaceKey(family, null);
  const em = (sizePt * scale * OFFICE_PRINT_DPI) / 72;
  const advances = OFFICE_DIGIT_ADVANCES as Readonly<Record<string, number>>;
  const digit = (OFFICE_DIGIT_WHOLE_EM as readonly string[]).includes(key)
    ? Math.max(1, Math.floor(Math.floor(em + 0.5) / 2 + 0.5))
    : Math.max(1, Math.round((advances[key] ?? advances.calibri ?? 0.50684) * em));
  const margin = Math.floor(digit / 2 + 0.5) + 7;
  return { left: Math.floor(margin / 2) + 1, right: margin - Math.floor(margin / 2), margin };
}

/**
 * How Windows Excel's printer scales a workbook's grid: from its Normal font
 * (font 0, the native grid font) and the screen digit width its columns are
 * counted in (officeDigitWidth of that font unless given).
 */
/** East Asian faces a cell may name in Latin letters (office_cjk_latin_face's). */
const LATIN_NAMED_EAST_ASIAN_FACES = new Set(["ms gothic", "ms pgothic", "ms mincho", "ms pmincho", "meiryo", "meiryo ui",
  "yu gothic", "yu mincho", "simsun", "nsimsun", "simhei", "microsoft yahei", "dengxian", "pmingliu", "mingliu",
  "microsoft jhenghei", "malgun gothic", "gulim", "dotum", "batang", "gungsuh"]);

/**
 * A printed Western face's size, page px: Windows Excel prints it at a whole
 * number of printer pixels per em at the page's scale, its advances with it
 * (lookwin bench: Calibri 11 at 100% 92 dots, 11.04 pt; Times New Roman 18
 * at 35% 6.36 pt; the API's draw_cell_text_upright). `dotPx` is the page px
 * of one printer dot. An East Asian face keeps its size: its words keep the
 * unrounded size's widths (R356-X27).
 */
export function officePrintedFontPx(family: string, sizePx: number, dotPx: number): number {
  if (!(sizePx > 0) || !(dotPx > 0)) return sizePx;
  if (officeEastAsianFamily(family)) return sizePx;
  return Math.max(1, Math.floor(sizePx / dotPx + 0.5)) * dotPx;
}

/** An East Asian Office face, by a name in its own script or in Latin letters. */
export function officeEastAsianFamily(family: string): boolean {
  const name = family.trim();
  return [...name].some((character) => (character.codePointAt(0) ?? 0) >= 0x3000) || LATIN_NAMED_EAST_ASIAN_FACES.has(name.toLowerCase());
}

export function officePrintScale(
  font: { readonly family?: string; readonly sizePt?: number; readonly record?: ExcelFontRecordClass | null } | undefined,
  screenDigitWidth?: number,
): OfficePrintScale {
  const size = font?.sizePt;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return OFFICE_PRINT_IDENTITY;
  // The face Windows Excel draws the Normal font in (a face it lacks prints as
  // Arial); a print measured under the name itself comes first (Osaka 12).
  const family = excelWindowsFaceKey(font?.family, font?.record);
  const entry = measured((font?.family ?? "").trim().toLowerCase(), size) ?? measured(family, size);
  const screenRow = entry?.[0] ?? officeRowHeight(font);
  const printedRow = entry?.[1] ?? estimatedPrintedRowDots(family, size);
  const screenDigit =
    typeof screenDigitWidth === "number" && Number.isFinite(screenDigitWidth) && screenDigitWidth > 0
      ? screenDigitWidth
      : officeDigitWidth(font);
  const printedDigit = entry?.[2] ?? estimatedPrintedDigitDots(family, size);
  if (!(screenRow > 0) || !(screenDigit > 0)) return OFFICE_PRINT_IDENTITY;
  return {
    x: printedDigit / (DOTS_PER_SCREEN_PIXEL * screenDigit),
    y: printedRow / (DOTS_PER_SCREEN_PIXEL * screenRow),
  };
}

/**
 * One line of a face as Windows Excel's printer sets it, in 600-dpi dots: the
 * face's printed default row (office-print-metrics.json), a size the table
 * lacks read off the face's measured sizes where it has three or more (a
 * straight line through them: Aptos Narrow 6-12 pt give 202 dots at 18 pt,
 * lookwin p107-lw prints 201), else the estimate. Wrapped lines print this
 * far apart whatever the sheet's Normal font (p107-lw: Calibri 11 lines 121
 * dots apart in an Aptos Narrow 11 sheet). The API's office_printed_row_dots.
 */
export function officePrintedRowDots(
  font: { readonly family?: string; readonly sizePt?: number; readonly record?: ExcelFontRecordClass | null },
): number | undefined {
  const size = font.sizePt;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return undefined;
  const family = excelWindowsFaceKey(font.family, font.record);
  for (const at of [(font.family ?? "").trim().toLowerCase(), family]) {
    if (!at || at.startsWith("_")) continue;
    const entry = measured(at, size);
    if (entry) return entry[1];
    const sizes = measuredSizes(at);
    if (!sizes || typeof sizes !== "object") continue;
    const points: [number, number][] = [];
    for (const [key, value] of Object.entries(sizes as Readonly<Record<string, unknown>>)) {
      if (Array.isArray(value) && value.length === 3 && typeof value[1] === "number") points.push([Number(key), value[1]]);
    }
    if (points.length < 3) continue;
    const meanX = points.reduce((total, [x]) => total + x, 0) / points.length;
    const meanY = points.reduce((total, [, y]) => total + y, 0) / points.length;
    const spread = points.reduce((total, [x]) => total + (x - meanX) ** 2, 0);
    if (!(spread > 0)) continue;
    const slope = points.reduce((total, [x, y]) => total + (x - meanX) * (y - meanY), 0) / spread;
    return Math.max(1, Math.floor(meanY + slope * (size - meanX) + 0.5));
  }
  return estimatedPrintedRowDots(family, size);
}

/** The anchor a drawing prints by: Windows Excel prints an object the file
 * places by two cell markers (`fileMarkers`, an `xdr:twoCellAnchor` of any
 * `editAs`), or that the export writes so (`editAs` equal to its kind), by
 * those markers in the printed grid, as a two-cell anchor (lookwin p75-lw,
 * p78-lw; the API's drawing_rotation_box.marker_print_anchor). */
export function officePrintAnchorKind(
  anchor: { readonly kind?: string; readonly fileMarkers?: unknown; readonly editAs?: unknown } | undefined,
): "twoCell" | "oneCell" | "absolute" {
  const kind = anchor?.kind === "oneCell" || anchor?.kind === "absolute" ? anchor.kind : "twoCell";
  if (kind !== "twoCell" && (anchor?.fileMarkers === true || anchor?.editAs === kind)) return "twoCell";
  return kind;
}

/** A drawing's frame on the printed page, in screen pixels: a two-cell anchor
 * spans its cells (both axes scale), a one-cell anchor sits at its cell but
 * keeps its own extent, an absolute anchor keeps both. */
export function officePrintedDrawingFrame(
  frame: readonly [left: number, top: number, width: number, height: number],
  kind: "twoCell" | "oneCell" | "absolute" | undefined,
  scale: OfficePrintScale,
): readonly [left: number, top: number, width: number, height: number] {
  if (scale.x === 1 && scale.y === 1) return frame;
  const [left, top, width, height] = frame;
  if (kind === "absolute") return frame;
  if (kind === "oneCell") return [left * scale.x, top * scale.y, width, height];
  return [left * scale.x, top * scale.y, width * scale.x, height * scale.y];
}
