/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Add Chart Element > Data Table: the grid of the chart's values drawn under
 * its plot, one column per category and one row per series, each row led by
 * the series' legend key unless the table was added "No Legend Keys" (Office's
 * `c:dTable/c:showKeys`). Docs and the deck draw it from this one reading of a
 * Chart V2 block; each host lays it out in its own frame.
 */
import type { CSSProperties } from "react";

import {
  chartSupportsDataTable,
  type ChartBlockAttrsV2,
} from "../../components/console/editor/extensions/chartBlockShared";
import { presentationChartAreaStyleCss } from "../presentation-chart-area-style";
import { officeChartTextEChartsStyle } from "./chart-text";
import { applyNumberFormat } from "../spreadsheet/number-format";

export type OfficeChartDataTable = Readonly<{
  categories: readonly string[];
  series: readonly Readonly<{ name: string; values: readonly string[] }>[];
  legendKeys: boolean;
  /** Format Data Table ▸ Table Borders. */
  borders?: Readonly<{ horizontal: boolean; vertical: boolean; outline: boolean }>;
  /**
   * Each series' General values with decimals (null elsewhere), which a
   * deck's table rounds to its columns' width
   * (`officeDeckChartDataTableGeneralFitted`).
   */
  general?: readonly (readonly (number | null)[])[];
}>;









function officeChartDataTableIsGeneral(numberFormat?: string): boolean {
  const code = typeof numberFormat === "string" ? numberFormat.trim() : "";
  return !code || code.toLowerCase() === "general";
}

export function officeChartDataTableValue(value: unknown, numberFormat?: string): string {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  const code = typeof numberFormat === "string" ? numberFormat.trim() : "";
  if (!officeChartDataTableIsGeneral(code)) {
    try {
      return applyNumberFormat(value, code);
    } catch {
      // A code the engine refuses prints the number as before.
    }
  }
  return officeChartDataTableGeneral(value);
}







function officeChartDataTableGeneral(value: number, decimals?: number): string {
  if (Number.isInteger(value)) return String(value);
  const magnitude = Math.abs(value);
  if (magnitude >= 1e10 || magnitude < 1e-4) return String(Number(value.toPrecision(6)));
  const whole = magnitude < 1 ? 1 : Math.floor(Math.log10(magnitude)) + 1;
  return String(Number(value.toFixed(Math.min(decimals ?? Infinity, Math.max(0, 10 - whole)))));
}

/** The decimals General writes `value` with (the eleven characters' share). */
function officeChartDataTableGeneralDecimals(value: number): number {
  const text = officeChartDataTableGeneral(value);
  const point = text.indexOf(".");
  return point < 0 || /e/i.test(text) ? 0 : text.length - point - 1;
}









export function officeDeckChartDataTableGeneralFitted(
  look: OfficeDeckChartDataTableLook,
  plotWidthPx: number,
  measure: (text: string, font: string) => number,
): void {
  const { table } = look;
  const count = table.categories.length;
  if (!table.general || !(plotWidthPx > 0) || !count) return;
  const room = Math.max(1, plotWidthPx / count - 2 * POWERPOINT_DATA_TABLE.categoryPadPt * PT_PX);
  let changed = false;
  const series = table.series.map((entry, index) => {
    const general = table.general?.[index];
    if (!general) return entry;
    const values = entry.values.map((text, at) => {
      const value = general[at];
      if (typeof value !== "number") return text;
      let decimals = officeChartDataTableGeneralDecimals(value);
      let fitted = officeChartDataTableGeneral(value, decimals);
      while (decimals > 0 && measure(fitted, look.font) > room) {
        decimals -= 1;
        fitted = officeChartDataTableGeneral(value, decimals);
      }
      if (fitted !== text) changed = true;
      return fitted;
    });
    return { ...entry, values };
  });
  if (changed) look.table = { ...table, series };
}

/** The table a chart shows, or null when it shows none. */
export function officeChartDataTable(
  attrs: ChartBlockAttrsV2,
): OfficeChartDataTable | null {
  if (
    attrs.format?.dataTable !== true ||
    attrs.data.kind !== "category" ||
    !chartSupportsDataTable(attrs.chartType)
  ) {
    return null;
  }
  // The table is the plotted data's: a series or a category the Chart
  // Filters hide is not in it either (PPT-B292).
  const hiddenSeries = new Set(attrs.format?.hiddenSeries ?? []);
  const hiddenCategories = new Set(attrs.format?.hiddenCategories ?? []);
  const shownCategory = (_category: unknown, index: number) => !hiddenCategories.has(index);
  const shown = attrs.data.series.filter((_series, index) => !hiddenSeries.has(index));
  const general = shown.map((series) =>
    series.values
      .filter(shownCategory)
      .map((value) =>
        typeof value === "number" && Number.isFinite(value) && !Number.isInteger(value) && officeChartDataTableIsGeneral(series.numberFormat)
          ? value
          : null,
      ),
  );
  return {
    categories: attrs.data.categories.filter(shownCategory),
    series: shown.map((series) => ({
      name: series.name,
      values: series.values.filter(shownCategory).map((value) => officeChartDataTableValue(value, series.numberFormat)),
    })),
    ...(general.some((values) => values.some((value) => value !== null)) ? { general } : {}),
    legendKeys: attrs.format?.dataTableLegendKeys !== false,
    borders: {
      horizontal: attrs.format?.dataTableHorizontalBorder !== false,
      vertical: attrs.format?.dataTableVerticalBorder !== false,
      outline: attrs.format?.dataTableOutline !== false,
    },
  };
}

/** One table row's height, design px: the 9 px text, its padding and rule. */
export const OFFICE_CHART_DATA_TABLE_ROW_PX = 14;
/** The most of its frame a chart gives its data table. */
export const OFFICE_CHART_DATA_TABLE_MAX_SHARE = 0.7;
/** The lowest a row is drawn once the table is fitted into its share. */
export const OFFICE_CHART_DATA_TABLE_MIN_ROW_PX = 9;
/** The table's text size when the chart states none (design px). */
export const OFFICE_CHART_DATA_TABLE_FONT_PX = 9;

/**
 * The band a data table takes under the plot of a chart `frameHeightPx`
 * tall: its rows at their own height, the plot giving way as PowerPoint's
 * does, up to the share that still leaves a plot to read; a longer table is
 * drawn smaller, row and text alike (PPT-B328). `fitted` is the text size
 * the rows were shrunk to, when they were.
 */
export function officeChartDataTableBand(
  rows: number,
  frameHeightPx: number,
  fontSizePx?: number,
): { heightPx: number; rowPx: number; fitted?: { fontSize: number } } {
  let rowPx = Math.max(
    OFFICE_CHART_DATA_TABLE_ROW_PX,
    typeof fontSizePx === "number" ? Math.round(fontSizePx * 1.4) + 1 : OFFICE_CHART_DATA_TABLE_ROW_PX,
  );
  let fitted: { fontSize: number } | undefined;
  const cap = frameHeightPx > 0 ? frameHeightPx * OFFICE_CHART_DATA_TABLE_MAX_SHARE : Infinity;
  if (rows * rowPx + 1 > cap) {
    rowPx = Math.max(OFFICE_CHART_DATA_TABLE_MIN_ROW_PX, Math.floor((cap - 1) / rows));
    fitted = { fontSize: Math.max(6, Math.floor((rowPx - 1) / 1.4)) };
  }
  return { heightPx: Math.max(0, Math.min(rows * rowPx + 1, cap)), rowPx, ...(fitted ? { fitted } : {}) };
}

/** The four edges of one table cell, each drawn or not. */
export type OfficeChartDataTableCellEdges = Readonly<{
  top: boolean;
  right: boolean;
  bottom: boolean;
  left: boolean;
}>;

/**
 * Which edges of the cell at `row`, `column` of a `rows` x `columns` table
 * (the header row and the series-name column included) Format Data Table ▸
 * Table Borders draws: Horizontal the lines between rows, Vertical the lines
 * between columns, Outline the table's own edge (`c:dTable/c:showHorzBorder`,
 * `c:showVertBorder`, `c:showOutline`). A table that states no borders draws
 * them all.
 */
export function officeChartDataTableCellEdges(
  borders: OfficeChartDataTable["borders"],
  row: number,
  column: number,
  rows: number,
  columns: number,
  /** Excel leaves the corner above the series names open: no lines around
   * it, the outline running round the header cells and the name column
   * instead (C29_ChartDataTable). */
  openCorner = false,
): OfficeChartDataTableCellEdges {
  const horizontal = borders?.horizontal !== false;
  const vertical = borders?.vertical !== false;
  const outline = borders?.outline !== false;
  if (openCorner && row === 0 && column === 0) return { top: false, right: false, bottom: false, left: false };
  if (openCorner && row === 1 && column === 0) {
    return { top: outline, bottom: rows === 2 ? outline : horizontal, left: outline, right: columns === 1 ? outline : vertical };
  }
  if (openCorner && row === 0 && column === 1) {
    return { top: outline, bottom: rows === 1 ? outline : horizontal, left: outline, right: columns === 2 ? outline : vertical };
  }
  return {
    top: row === 0 ? outline : horizontal,
    bottom: row === rows - 1 ? outline : horizontal,
    left: column === 0 ? outline : vertical,
    right: column === columns - 1 ? outline : vertical,
  };
}










export function officeChartDataTableFontPx(dataTableStyle: unknown, chartText: unknown): number | undefined {
  const own = presentationChartAreaStyleCss(dataTableStyle as Parameters<typeof presentationChartAreaStyleCss>[0]).fontSize;
  if (typeof own === "number" && own > 0) return own;
  const chart = officeChartTextEChartsStyle(chartText as Parameters<typeof officeChartTextEChartsStyle>[0]).fontSize;
  return typeof chart === "number" && chart > 0 ? chart : undefined;
}

/** A workbook chart's table text size when the chart states none (px). */
export const OFFICE_SHEET_CHART_DATA_TABLE_FONT_PX = 12;

/**
 * A workbook chart's data table as Excel draws it inside the chart area: its
 * own Format Data Table look over the chart's text, the band it takes under
 * the plot (`officeChartDataTableBand`) at a chart `heightPx` tall, and the
 * text size its rows were fitted to. The sheet's canvas and the render server
 * (for the sheet's print) read it here, so the two lay the table out alike.
 */
export function officeSheetChartDataTableLook(
  attrs: ChartBlockAttrsV2,
  heightPx: number,
): { table: OfficeChartDataTable; band: ReturnType<typeof officeChartDataTableBand>; style: CSSProperties } | null {
  const table = officeChartDataTable(attrs);
  if (!table) return null;
  const text = officeChartTextEChartsStyle(attrs.format?.text);
  // Its own font (Format Data Table > Font, `c:dTable/c:txPr`) with its own
  // colour: the area style's CSS writes no text colour (XW242's bold italic
  // red table drew grey; X-346).
  const ownText = (attrs.format?.dataTableStyle as { text?: { color?: unknown } } | undefined)?.text;
  const own: CSSProperties = {
    ...presentationChartAreaStyleCss(attrs.format?.dataTableStyle),
    ...(typeof ownText?.color === "string" ? { color: ownText.color } : {}),
  };
  const inherited: CSSProperties = {
    ...(typeof text.fontSize === "number" ? { fontSize: text.fontSize } : {}),
    ...(typeof text.color === "string" ? { color: text.color } : {}),
    ...(typeof text.fontFamily === "string" ? { fontFamily: text.fontFamily } : {}),
  };
  const fontSize = officeChartDataTableFontPx(attrs.format?.dataTableStyle, attrs.format?.text);
  const band = officeChartDataTableBand(
    table.series.length + 1,
    heightPx,
    typeof fontSize === "number" ? fontSize : OFFICE_SHEET_CHART_DATA_TABLE_FONT_PX,
  );
  const fitted: CSSProperties = band.fitted ? { fontSize: band.fitted.fontSize, lineHeight: 1.1 } : {};
  return { table, band, style: { ...inherited, ...own, ...fitted } };
}

/** A point in CSS px. */
const PT_PX = 4 / 3;
/** PowerPoint's chart text size when neither the table nor the chart states one (pt). */
const POWERPOINT_CHART_TEXT_PT = 10;

/**
 * How Windows PowerPoint lays out a chart's data table, measured on its own
 * pictures (office-truth/probes/c30dtab, 16 charts, Slide.Export; TRIAGE C40):
 * a row is 1.2 of its text and 3.6 pt, a line of 1.2 em more for each line a
 * wrapped name takes (10 / 12 / 16 pt: 15.6 / 18.0 / 22.8 pt; a two-line 16 pt
 * name 42.0 pt). The series names' column is the legend key's slot (0.975 em,
 * none without keys), the widest name and 10 pt, at most a third of the chart,
 * less 6 pt (12 pt Calibri "Bravo": 44.4 pt drawn; keys off 32.4; a 576 pt
 * chart's capped column 186.6), its names wrapping inside it. The table's
 * left edge stands 16.8 pt in from the chart's at every size (the inner plot
 * begins after the names), its foot 5.4 pt above the chart's. The workbook
 * table keeps Excel's own band (`officeChartDataTableBand`): these numbers are
 * PowerPoint's, not Excel's.
 */
export const POWERPOINT_DATA_TABLE = {
  rowFactor: 1.2,
  rowPadPt: 3.6,
  keySlotEm: 0.975,
  namePadPt: 10,
  nameColumnShare: 1 / 3,
  nameColumnInsetPt: 6,
  leftPt: 16.8,
  footPt: 5.4,
  categoryPadPt: 3,
} as const;

export interface OfficeDeckChartDataTableLook {
  table: OfficeChartDataTable;
  /** The text's CSS font (for measuring) and one line's height (px). */
  font: string;
  linePx: number;
  /** The lines the longest category name takes in the header row (absent: one). */
  headerLines?: number;
  /** The table's text size (px). */
  fontPx: number;
  /** Each row's height (px): the categories' row, then a row per series. */
  rowHeightsPx: number[];
  /** The band: every row together. */
  heightPx: number;
  /** The series names' column, left of the plot (px). */
  nameColumnPx: number;
  /** The lines each series' name takes in its column. */
  nameLines: number[];
  style: CSSProperties;
}

/**
 * A deck chart's data table laid out as Windows PowerPoint lays it out under
 * a chart `widthPx` wide (`POWERPOINT_DATA_TABLE`), names measured with
 * `measure(text, font)` (px). The rows' text is the table's own size, else the
 * chart's, else PowerPoint's 10 pt (`officeChartDataTableFontPx`).
 */
export function officeDeckChartDataTableLook(
  attrs: ChartBlockAttrsV2,
  widthPx: number,
  measure: (text: string, font: string) => number,
  /** The face the chart's text that names none is drawn in (the deck's theme body face). */
  defaultFamily?: string | null,
): OfficeDeckChartDataTableLook | null {
  const table = officeChartDataTable(attrs);
  if (!table) return null;
  const fontPx = officeChartDataTableFontPx(attrs.format?.dataTableStyle, attrs.format?.text) ?? POWERPOINT_CHART_TEXT_PT * PT_PX;
  const text = officeChartTextEChartsStyle(attrs.format?.text);
  const ownText = (attrs.format?.dataTableStyle as { text?: { color?: unknown } } | undefined)?.text;
  const own: CSSProperties = {
    ...presentationChartAreaStyleCss(attrs.format?.dataTableStyle),
    ...(typeof ownText?.color === "string" ? { color: ownText.color } : {}),
  };
  const inherited: CSSProperties = {
    ...(typeof text.color === "string" ? { color: text.color } : {}),
    ...(typeof text.fontFamily === "string" ? { fontFamily: text.fontFamily } : {}),
  };
  // The table's own face, else the chart's, else the deck's body face: the
  
  // measured as Calibri and drawn in the page's face, its names wrapped).
  const family =
    typeof own.fontFamily === "string" && own.fontFamily
      ? own.fontFamily
      : typeof text.fontFamily === "string" && text.fontFamily
        ? text.fontFamily
        : defaultFamily || "Calibri";
  const weight = own.fontWeight === "bold" || own.fontWeight === 700 ? "bold " : "";
  const font = `${weight}${fontPx}px ${family}`;
  const key = table.legendKeys ? POWERPOINT_DATA_TABLE.keySlotEm * fontPx : 0;
  const pad = POWERPOINT_DATA_TABLE.namePadPt * PT_PX;
  const widest = Math.max(0, ...table.series.map((series) => measure(series.name, font)));
  const cap = widthPx > 0 ? widthPx * POWERPOINT_DATA_TABLE.nameColumnShare : Infinity;
  const inset = POWERPOINT_DATA_TABLE.nameColumnInsetPt * PT_PX;
  const nameColumnPx = Math.max(key + 1, Math.min(key + widest + pad, cap) - inset);
  const room = Math.max(1, nameColumnPx + inset - key - pad);
  const nameLines = table.series.map((series) => wrappedLines(series.name, font, room, measure));
  const line = POWERPOINT_DATA_TABLE.rowFactor * fontPx;
  const rowPad = POWERPOINT_DATA_TABLE.rowPadPt * PT_PX;
  const rowHeightsPx = [line + rowPad, ...nameLines.map((lines) => lines * line + rowPad)];
  const heightPx = rowHeightsPx.reduce((sum, height) => sum + height, 0);
  return { table, fontPx, font, linePx: line, rowHeightsPx, heightPx, nameColumnPx, nameLines, style: { ...inherited, ...own, fontFamily: family, fontSize: fontPx } };
}









export function officeDeckChartDataTableHeaderGrowth(
  look: OfficeDeckChartDataTableLook,
  plotWidthPx: number,
  measure: (text: string, font: string) => number,
): number {
  const count = look.table.categories.length;
  if (!(plotWidthPx > 0) || !count) return 0;
  const room = Math.max(1, plotWidthPx / count - 2 * POWERPOINT_DATA_TABLE.categoryPadPt * PT_PX);
  const lines = Math.max(1, ...look.table.categories.map((category) => wrappedLines(String(category), look.font, room, measure)));
  const growth = (lines - 1) * look.linePx;
  if (growth > 0) {
    look.rowHeightsPx[0] += growth;
    look.heightPx += growth;
    look.headerLines = lines;
  }
  return growth;
}

/** The lines `text` takes broken at its spaces within `room` px (a word wider
 * than the room is a line of its own). */
function wrappedLines(text: string, font: string, room: number, measure: (text: string, font: string) => number): number {
  const words = text.split(/\s+/).filter(Boolean);
  if (!words.length) return 1;
  let lines = 1;
  let current = "";
  for (const word of words) {
    const next = current ? `${current} ${word}` : word;
    if (current && measure(next, font) > room) {
      lines += 1;
      current = word;
    } else {
      current = next;
    }
  }
  return lines;
}
