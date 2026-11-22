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

export type OfficeChartDataTable = Readonly<{
  categories: readonly string[];
  series: readonly Readonly<{ name: string; values: readonly string[] }>[];
  legendKeys: boolean;
  /** Format Data Table ▸ Table Borders. */
  borders?: Readonly<{ horizontal: boolean; vertical: boolean; outline: boolean }>;
}>;

/** A value as the table prints it: the number itself, without float noise. */
export function officeChartDataTableValue(value: unknown): string {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  return String(Number.isInteger(value) ? value : Number(value.toFixed(6)));
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
  return {
    categories: attrs.data.categories.filter(shownCategory),
    series: attrs.data.series
      .filter((_series, index) => !hiddenSeries.has(index))
      .map((series) => ({
        name: series.name,
        values: series.values.filter(shownCategory).map(officeChartDataTableValue),
      })),
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
