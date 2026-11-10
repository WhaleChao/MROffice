/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Data > Forecast Sheet: the history read off the open sheet, the forecast
 * table built from it and the new worksheet (table + its chart) that holds the
 * result. The window's Forecast Sheet dialog and the AI's
 * spreadsheet.data.forecast_sheet command both run these, so both read the
 * same cells, refuse the same inputs and write the same sheet.
 */

import { spreadsheetChartBornWithModel } from "@mrnote/core/web/lib/spreadsheet/chart-v2-model.ts";
import {
  drawingChartRangesOf,
  drawingChartSeriesRefsForRange,
  drawingChartSheetPrefix,
} from "@mrnote/core/web/lib/spreadsheet/drawing-chart-series.ts";
import {
  buildForecastSheet,
  type ForecastDuplicateAggregation,
  type ForecastMissingPointPolicy,
  type ForecastSheetError,
} from "./forecast-sheet";
import { DEFAULT_SPREADSHEET_COLUMN_COUNT, DEFAULT_SPREADSHEET_ROW_COUNT } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { MAX_SHEETS } from "./sheet-buttons";

type CellLike = { value?: unknown; display?: unknown };
type Rect = { r1: number; r2: number; c1: number; c2: number };

export interface ForecastTableRequest {
  /** A1 range of the timeline, on the sheet being read. */
  timeline: string;
  /** A1 range of the values, on the sheet being read. */
  values: string;
  periods: number;
  /** 1 detects the season, 0 none, N >= 2 a fixed period. */
  seasonality: number;
  /** 0..1 */
  confidence: number;
  missingPoints: ForecastMissingPointPolicy;
  duplicateAggregation: ForecastDuplicateAggregation;
  /** The headers' language (they become worksheet data). */
  language: "zh" | "en";
  /** Timeline, values, forecast, lower and upper bound in the reader's
   * language (the UI catalogue's words); without them, `language` picks the
   * built-in zh/en set. */
  headers?: readonly [string, string, string, string, string];
}

/** Excel's FORECAST.ETS aggregation argument for duplicate timestamps. */
const AGGREGATION_ARGUMENT: Record<ForecastDuplicateAggregation, number> = {
  average: 1, count: 2, counta: 3, max: 4, median: 5, min: 6, sum: 7,
};

export type ForecastTableError = ForecastSheetError | "timeline-range" | "values-range";

/** The text a cell shows, by 0-based row and column. */
export type ForecastCellText = (row: number, col: number) => string;

/** What a cell shows, read as the window reads it: its display, else its value. */
export function forecastCellMapText(cells: Readonly<Record<string, CellLike | undefined>>): ForecastCellText {
  return (row, col) => {
    const cell = cells[`${row}:${col}`];
    return typeof cell?.display === "string" ? cell.display : String(cell?.value ?? "");
  };
}

function readNumbers(textAt: ForecastCellText, range: Rect): number[] {
  const numbers: number[] = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    for (let col = range.c1; col <= range.c2; col += 1) {
      const text = textAt(row, col).trim();
      const parsed = text === "" ? Number.NaN : Number(text);
      numbers.push(Number.isFinite(parsed) ? parsed : Number.NaN);
    }
  }
  return numbers;
}

function readValues(textAt: ForecastCellText, range: Rect): Array<string | null> {
  const values: Array<string | null> = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    for (let col = range.c1; col <= range.c2; col += 1) {
      const text = textAt(row, col);
      values.push(text.trim() === "" ? null : text);
    }
  }
  return values;
}

/**
 * The forecast sheet's table (header row first), as Excel's Create Forecast
 * Worksheet writes it (Microsoft Support, "Create a forecast in Excel for
 * Windows"): the history (timeline and values) and then the forecast, whose
 * values are FORECAST.ETS formulas over that history and whose bounds are the
 * forecast less and plus FORECAST.ETS.CONFINT. The last historical point is
 * repeated in the forecast columns so the lines meet. `rows` is what the
 * cells hold; `snapshot` the values they show, computed by the workbook's own
 * engine, for the chart born with them.
 */
export function forecastTableRows(
  textAt: ForecastCellText,
  request: ForecastTableRequest,
):
  | { ok: true; rows: string[][]; snapshot: string[][]; forecastCount: number }
  | { ok: false; error: ForecastTableError } {
  const timelineRange = parseA1Range(request.timeline.trim());
  const valuesRange = parseA1Range(request.values.trim());
  if (!timelineRange) return { ok: false, error: "timeline-range" };
  if (!valuesRange) return { ok: false, error: "values-range" };
  const built = buildForecastSheet({
    timeline: readNumbers(textAt, timelineRange),
    values: readValues(textAt, valuesRange),
    periods: request.periods,
    seasonality: request.seasonality,
    missingPoints: request.missingPoints,
    duplicateAggregation: request.duplicateAggregation,
    confidence: request.confidence,
  });
  if (!built.ok) return built;
  const headers = request.headers
    ? [...request.headers]
    : request.language === "zh"
      ? ["时间轴", "值", "预测值", "置信下限", "置信上限"]
      : ["Timeline", "Values", "Forecast", "Lower bound", "Upper bound"];
  // The history as the source shows it, one row per point.
  const history: Array<[string, string]> = [];
  for (let row = timelineRange.r1; row <= timelineRange.r2; row += 1) {
    for (let col = timelineRange.c1; col <= timelineRange.c2; col += 1) {
      history.push([textAt(row, col).trim(), ""]);
    }
  }
  let index = 0;
  for (let row = valuesRange.r1; row <= valuesRange.r2; row += 1) {
    for (let col = valuesRange.c1; col <= valuesRange.c2; col += 1) {
      if (history[index]) history[index]![1] = textAt(row, col).trim();
      index += 1;
    }
  }
  const lastHistoryRow = history.length + 1;
  const timelineRef = `$A$2:$A$${lastHistoryRow}`;
  const valuesRef = `$B$2:$B$${lastHistoryRow}`;
  const tail = [
    request.seasonality,
    request.missingPoints === "zero" ? 0 : 1,
    ...(request.duplicateAggregation === "average" ? [] : [AGGREGATION_ARGUMENT[request.duplicateAggregation]]),
  ].join(",");
  const confidence = Number(request.confidence.toFixed(6));
  const rows: string[][] = [headers];
  history.forEach(([timeline, value], offset) => {
    // The last point starts the forecast lines.
    const joins = offset === history.length - 1 && value !== "";
    rows.push([timeline, value, joins ? value : "", joins ? value : "", joins ? value : ""]);
  });
  built.result.rows.forEach((point, offset) => {
    const sheetRow = lastHistoryRow + offset + 1;
    const forecast = `=FORECAST.ETS(A${sheetRow},${valuesRef},${timelineRef},${tail})`;
    const interval = `FORECAST.ETS.CONFINT(A${sheetRow},${valuesRef},${timelineRef},${confidence},${tail})`;
    rows.push([String(point.timeline), "", forecast, `=C${sheetRow}-${interval}`, `=C${sheetRow}+${interval}`]);
  });
  // What the new sheet's formula cells will show, for the chart's first
  // drawing: the forecast and its bounds as computed above. The grid's
  // evaluator is server-only (it reached the public bundle and blocked every
  
  const firstForecastRow = history.length + 1;
  const snapshot = rows.map((row, r) => {
    if (r === 0) return row;
    const point = r >= firstForecastRow ? built.result.rows[r - firstForecastRow] : undefined;
    return row.map((text, c) => {
      if (!text.startsWith("=")) return text;
      const shown = !point ? undefined : c === 2 ? point.forecast : c === 3 ? point.lower : c === 4 ? point.upper : undefined;
      return typeof shown === "number" && Number.isFinite(shown) ? String(shown) : "";
    });
  });
  return { ok: true, rows, snapshot, forecastCount: built.result.rows.length };
}

/** Excel's sheet-name rules, as the dialog checks them. */
export function forecastSheetNameValid(name: string): boolean {
  const trimmed = name.trim();
  return (
    trimmed.length >= 1 &&
    trimmed.length <= 31 &&
    !/[\\/\?*\[\]:]/.test(trimmed) &&
    !trimmed.startsWith("'") &&
    !trimmed.endsWith("'")
  );
}

export interface ForecastOutputRequest {
  sheetName: string;
  chartType: "line" | "column";
  /** What the cells hold (header row first; formulas start with "="). */
  rows: string[][];
  /** The values those cells show, for the chart's first drawing. */
  snapshot?: string[][];
  /** Where the history came from (the table now holds a copy of it). */
  sourceSheetName: string;
  sourceRanges: [string, string];
}

/** Excel's default chart: 5 x 3 inches. */
const FORECAST_CHART_WIDTH_EMU = 4_572_000;
const FORECAST_CHART_HEIGHT_EMU = 2_743_200;

/** The new worksheet holding the forecast table and its chart, or why not. */
export function forecastOutputSheet(
  sheets: ReadonlyArray<{ id: string; name: string }>,
  { sheetName, chartType, rows, snapshot }: ForecastOutputRequest,
): { ok: true; sheet: Record<string, unknown> } | { ok: false; error: "sheet-limit" | "invalid-name" | "duplicate-name" } {
  if (sheets.length >= MAX_SHEETS) return { ok: false, error: "sheet-limit" };
  const trimmedName = sheetName.trim();
  if (!forecastSheetNameValid(trimmedName)) return { ok: false, error: "invalid-name" };
  const foldedName = trimmedName.toLocaleLowerCase();
  if (sheets.some((candidate) => candidate.name.trim().toLocaleLowerCase() === foldedName)) {
    return { ok: false, error: "duplicate-name" };
  }
  const existingIds = new Set(sheets.map((candidate) => candidate.id));
  let suffix = sheets.length + 1;
  let id = `sheet-${suffix}`;
  while (existingIds.has(id)) {
    suffix += 1;
    id = `sheet-${suffix}`;
  }
  const cells: Record<string, Record<string, unknown>> = {};
  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => {
      const numeric = rowIndex > 0 ? Number(value) : Number.NaN;
      if (value === "") return;
      cells[`${rowIndex}:${columnIndex}`] = {
        value,
        kind: Number.isFinite(numeric) ? "number" : "text",
        // A formula cell names its formula, as a typed one does
        // (writeCellValueInto); without it "=FORECAST.ETS(...)" reads as text.
        ...(value.startsWith("=") ? { formula: value } : {}),
        ...(rowIndex === 0 ? { style: { bold: true, backgroundColor: "#e8f3ec" } } : {}),
      };
    });
  });
  const lastRow = Math.max(1, rows.length);
  const lastColumn = String.fromCharCode(64 + Math.max(1, rows[0]?.length ?? 1));
  const range = `A1:${lastColumn}${lastRow}`;
  const shown = snapshot ?? rows;
  // The chart's series are the table's cells, as Excel's forecast chart's
  // are: it draws what the FORECAST.ETS / CONFINT cells show (the snapshot
  // only paints it before the grid has calculated them), and follows them
  // when the history is edited. Born with constants it showed bounds of its
  // own that disagreed with the table's (live audit S-103).
  const seriesRefs = drawingChartSeriesRefsForRange(
    drawingChartSheetPrefix(trimmedName),
    { r1: 0, c1: 0, r2: lastRow - 1, c2: Math.max(1, (rows[0]?.length ?? 1) - 1) },
    true,
    chartType === "column" ? "bar" : "line",
  );
  return {
    ok: true,
    sheet: {
      id,
      name: trimmedName,
      rowCount: Math.max(DEFAULT_SPREADSHEET_ROW_COUNT, rows.length),
      columnCount: Math.max(DEFAULT_SPREADSHEET_COLUMN_COUNT, 10),
      cells,
      objects: [
        // Born with its model over the forecast table it draws.
        spreadsheetChartBornWithModel(
          {
            id: `${id}-forecast-chart`,
            type: "chart",
            chartVisibility: { state: "known", plotVisOnly: true },
            title: trimmedName,
            caption: trimmedName,
            chartType: chartType === "column" ? "bar" : "line",
            range,
            // The table holds the history too, so the chart reads only it.
            ranges: seriesRefs ? drawingChartRangesOf(seriesRefs) : [range],
            ...(seriesRefs ? { seriesRefs } : {}),
            showLegend: true,
            showAxisLabels: true,
            // A drawing chart plots `data.rows` and reads their first row as
            // the series names (importedChartRenderData): the header row
            // leads, or the first point became "Series 1"'s name.
            data: { headers: shown[0] ?? [], rows: shown },
            // Beside the table, at Excel's default chart size (without a
            // size it was drawn a few cells small and unreadable).
            anchor: {
              kind: "oneCell",
              from: { row: 1, column: (rows[0]?.length ?? 4) + 1 },
              widthEmu: FORECAST_CHART_WIDTH_EMU,
              heightEmu: FORECAST_CHART_HEIGHT_EMU,
            },
          } as never,
          // The header row names the series (Values, Forecast, bounds) and
          // the timeline is the category axis.
          { rows: shown, headerRow: true },
          { sheetIdByName: (name) => (name.toLocaleLowerCase() === foldedName ? id : undefined) },
        ),
      ],
    },
  };
}
