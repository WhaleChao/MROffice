/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A worksheet chart as its print draws it: the planner input the sheet's own
 * canvas hands the chart engine (`SpreadsheetChartCanvas`), built on the
 * server from the chart's record.
 *
 * The worksheet PDF used to send the chart's bare model as a Docs chart node,
 * so the print lost everything the sheet adds around the model -- an imported
 * chart's built-in look (black text, dark gridlines, its legend), the
 * automatic title, the chart's body face, the axis' linked number format --
 * and a printed chart was not the chart on the screen (X-221 / item 5 of the
 * chart-fixer-5 round: a "0.000%" pie printed 29%, X01 printed light grey
 * gridlines and no legend). The print now asks with the same input.
 */
import type { ChartBlockAttrsV2 } from "@mrnote/core/web/components/console/editor/extensions/chartBlockShared.ts";
import {
  normalizeChartSeriesRefs,
  normalizeChartValueAxisNumberFormat,
  normalizeSpreadsheetChartType,
  parseChartRefRange,
  resolveLinkedValueAxisNumberFormat,
  type SpreadsheetChartNumberFormatCell,
} from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import type { SpreadsheetChartV2Input } from "@mrnote/core/web/lib/spreadsheet/chart-v2-binding.ts";
import { normalizePresentationChartAreaStyle } from "@mrnote/core/web/lib/presentation-chart-area-style.ts";
import { officeChartBorderStyleCss, officeChartColorWithOpacity } from "@mrnote/core/web/lib/office-chart-shape-paint.ts";
import {
  spreadsheetChartModelInput,
  spreadsheetChartStoredModel,
  type SpreadsheetChartLike,
} from "@mrnote/core/web/lib/spreadsheet/chart-v2-model.ts";

/** What the host knows of the workbook beside the chart's record. */
export interface SpreadsheetChartPrintContext {
  /** The sheet the chart sits on (an unqualified reference's sheet). */
  readonly sheetName?: string;
  /**
   * The cells the chart's references name, keyed `sheet!row:col` (sheet name
   * lower-cased, 0-based row and column): their values and number formats,
   * for the axis' Linked to source format and the category cells' format.
   */
  readonly cells?: Readonly<Record<string, SpreadsheetChartNumberFormatCell>>;
}

/** The value axis' number format as the sheet hands it to its canvas. */
function valueAxisFormat(record: SpreadsheetChartLike, context: SpreadsheetChartPrintContext): string | undefined {
  const type = normalizeSpreadsheetChartType(record.chartType ?? record.type);
  const secondary = Array.isArray(record.secondaryAxisSeries) && record.secondaryAxisSeries.length > 0;
  if (!["column", "stacked-column", "line", "area"].includes(type) || secondary) return undefined;
  const stated = normalizeChartValueAxisNumberFormat(record.valueAxisNumberFormat) ?? {
    formatCode: "General",
    sourceLinked: true,
  };
  if (!stated.sourceLinked) return stated.formatCode;
  return resolveLinkedValueAxisNumberFormat(
    record as Parameters<typeof resolveLinkedValueAxisNumberFormat>[0],
    context.sheetName ?? "",
    (sheetName, row, col) => context.cells?.[`${sheetName.toLowerCase()}!${row}:${col}`],
  ).formatCode;
}

/** The format the category cells are written in (a date axis Linked to
 * source), or an X Y chart's X cells (its horizontal value axis). */
function categoryFormat(record: SpreadsheetChartLike, context: SpreadsheetChartPrintContext): string | undefined {
  const refs = normalizeChartSeriesRefs(record.seriesRefs);
  const range = parseChartRefRange(refs?.[0]?.categories ?? refs?.[0]?.xValues);
  if (!range) return undefined;
  const sheet = (range.sheetName ?? context.sheetName ?? "").toLowerCase();
  const format = context.cells?.[`${sheet}!${range.range.r1}:${range.range.c1}`]?.numberFormat;
  return typeof format === "string" && format ? format : undefined;
}

/** The model's own data as the sheet's rows (a header row of series names). */
function rowsOfModelData(model: ChartBlockAttrsV2): Pick<SpreadsheetChartV2Input, "rows" | "headerRow" | "categoryLevels"> | null {
  const data = model.data;
  if (data.kind === "xy") return {};
  if (data.kind !== "category") return null;
  const cell = (value: unknown) => (typeof value === "number" && Number.isFinite(value) ? String(value) : "");
  return {
    headerRow: true,
    rows: [
      ["", ...data.series.map((series) => series.name)],
      ...data.categories.map((category, index) => [category, ...data.series.map((series) => cell(series.values[index]))]),
    ],
    ...(data.categoryLevels?.length ? { categoryLevels: data.categoryLevels } : {}),
  };
}

/**
 * The sheet canvas' input for a chart record whose stored model already holds
 * the data it plots (the print re-reads the cells into it), or null when the
 * record has no model this reading draws (the caller prints it as before).
 */
export function spreadsheetChartPrintInput(
  record: SpreadsheetChartLike,
  context: SpreadsheetChartPrintContext,
  extras: Partial<SpreadsheetChartV2Input>,
): SpreadsheetChartV2Input | null {
  const model = spreadsheetChartStoredModel(record);
  if (!model) return null;
  const rows = rowsOfModelData(model);
  if (!rows) return null;
  const base = spreadsheetChartModelInput({ ...record, chartV2: model });
  const valueAxisNumberFormat = valueAxisFormat(record, context);
  const categoryNumberFormat = categoryFormat(record, context);
  return {
    ...base,
    ...rows,
    // An imported chart's look is the record's own, as the canvas reads it.
    ...(record.nativeLook && typeof record.nativeLook === "object" ? { nativeLook: record.nativeLook as SpreadsheetChartV2Input["nativeLook"] } : {}),
    ...(valueAxisNumberFormat ? { valueAxisNumberFormat } : {}),
    ...(categoryNumberFormat ? { categoryNumberFormat } : {}),
    ...extras,
    storedModel: model,
    storedModelOwnData: true,
  };
}

/** The chart area's frame line as the sheet's canvas draws it. */
export interface SpreadsheetChartFrameLine {
  readonly color: string;
  readonly widthPt: number;
  readonly style: "solid" | "dotted" | "dashed" | "double";
}

/**
 * The line the sheet's canvas draws round a chart (its frame's CSS border in
 * `SpreadsheetChartCanvas`): the chart area's own outline, else -- the chart
 * stating none -- the inserted-chart style's 0.75pt hairline in the style's
 * border colour (#d9d9d9), unless an imported Office 2016 look says No line.
 * The print draws the same line round the picture, which never holds it.
 */
export function spreadsheetChartFrameLine(
  chartAreaStyle: unknown,
  automatic: { readonly drawn: boolean; readonly color?: string },
): SpreadsheetChartFrameLine | null {
  const style = normalizePresentationChartAreaStyle(chartAreaStyle);
  const outline = style?.outline;
  if (outline === "none") return null;
  if (outline === undefined) {
    return automatic.drawn ? { color: automatic.color ?? "#d9d9d9", widthPt: 0.75, style: "solid" } : null;
  }
  const paint = style?.shapePaint;
  return {
    color: officeChartColorWithOpacity(outline, paint?.lineAlpha),
    widthPt: typeof paint?.outlineWidthPt === "number" && paint.outlineWidthPt > 0 ? paint.outlineWidthPt : 0.75,
    style: officeChartBorderStyleCss(paint) ?? "solid",
  };
}

/**
 * The chart frame's CSS line moved onto the frame's edge, as Windows Excel
 * prints a worksheet chart's area line: its path on the frame, half of it
 * outside (lookwin R457-X21; the API's draw_chart_frame, ae768db12c0). The
 * box's `border` (drawn inside a border-box) becomes an `outline` offset in
 * by half its width; every other property is kept. A style stating no line
 * is returned without one.
 */
export function chartFrameLineOnEdge(style: Readonly<Record<string, unknown>>): Record<string, unknown> {
  const out: Record<string, unknown> = { ...style };
  let width: unknown;
  let lineStyle: unknown;
  let color: unknown;
  const shorthand = typeof style.border === "string" ? style.border.trim() : undefined;
  if (shorthand !== undefined && shorthand !== "" && !/^(none|0(px|pt)?)\b/.test(shorthand)) {
    const parts = shorthand.split(/\s+(?![^(]*\))/);
    width = parts.find((part) => /^\d*\.?\d+(px|pt)?$/.test(part));
    lineStyle = parts.find((part) => /^(solid|dotted|dashed|double)$/.test(part));
    color = parts.find((part) => part !== width && part !== lineStyle);
  }
  if (style.borderWidth !== undefined) width = style.borderWidth;
  if (style.borderStyle !== undefined) lineStyle = style.borderStyle;
  if (style.borderColor !== undefined) color = style.borderColor;
  for (const key of Object.keys(out)) if (/^border(?!Radius)/.test(key)) delete out[key];
  const drawn = shorthand === undefined || !/^(none|0(px|pt)?)\b/.test(shorthand) || style.borderStyle !== undefined;
  if (!drawn || lineStyle === undefined || lineStyle === "none" || color === undefined) return out;
  const measured = /^(\d*\.?\d+)(px|pt)?$/.exec(String(width ?? 1));
  const amount = measured ? Number(measured[1]) : 1;
  const unit = measured?.[2] ?? "px";
  if (!(amount > 0)) return out;
  out.outline = `${amount}${unit} ${String(lineStyle)} ${String(color)}`;
  out.outlineOffset = `${-amount / 2}${unit}`;
  return out;
}

/** How far past its frame a chart's printed area line reaches, in CSS px:
 * half the line (`spreadsheetChartFrameLine`), which a page counts as the
 * chart's (R457-X21: a frame ending on a page's line printed the next page
 * for its outer half). 0 when the chart draws no line. */
export function chartFramePrintedReachPx(
  chartAreaStyle: unknown,
  automatic: { readonly drawn: boolean },
): number {
  const line = spreadsheetChartFrameLine(chartAreaStyle, automatic);
  return line ? (line.widthPt / 2) * (96 / 72) : 0;
}

/** {@link chartFramePrintedReachPx} of a chart's record (a sheet chart or an
 * imported chart object): its area style, and whether the automatic hairline
 * is drawn -- not when an imported Office 2016 look says No line, as
 * `SpreadsheetChartCanvas` decides. */
export function chartRecordFramePrintedReachPx(record: unknown): number {
  const fields = (record && typeof record === "object" ? record : {}) as {
    chartFormat?: { chartAreaStyle?: unknown };
    chartStyle?: unknown;
    nativeLook?: { chartArea?: { outline?: unknown } };
  };
  return chartFramePrintedReachPx(fields.chartFormat?.chartAreaStyle, {
    drawn: !(fields.chartStyle === undefined && fields.nativeLook?.chartArea?.outline === "none"),
  });
}
