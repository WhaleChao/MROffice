/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { SpreadsheetChartType } from "@mrnote/core/web/lib/spreadsheet/chart.ts";

export interface ChartRecommendation {
  type: SpreadsheetChartType;
  reason: "time-series" | "xy-pairs" | "part-to-whole" | "multi-series" | "distribution" | "general";
}

/** The numeric-column analysis of one selection matrix in series-in-columns
 * orientation. Computed once per matrix and shared by every gate and by the
 * recommendation ranking, so a large selection is profiled a single time. */
export interface ChartMatrixProfile {
  body: readonly (readonly string[])[];
  numericColumns: number[];
  firstCategory: string[];
  nonNegative: boolean;
}

function numeric(value: string): boolean {
  const text = value.trim().replace(/[$,%\s]/g, "").replace(/,/g, "");
  return text !== "" && Number.isFinite(Number(text));
}

function dateLike(value: string): boolean {
  const text = value.trim();
  return (
    /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(text) ||
    /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(text) ||
    (/^[A-Za-z]{3,9}\s+\d{1,2}/.test(text) && !Number.isNaN(Date.parse(text)))
  );
}

export function profileChartMatrix(
  matrix: readonly (readonly string[])[],
): ChartMatrixProfile {
  const width = Math.max(0, ...matrix.map((row) => row.length));
  const body = matrix.length > 1 ? matrix.slice(1) : matrix;
  const numericCounts = Array.from({ length: width }, (_, col) =>
    body.reduce((count, row) => count + (numeric(row[col] ?? "") ? 1 : 0), 0),
  );
  const populated = Math.max(1, body.length);
  const numericColumns = numericCounts
    .map((count, col) => ({ count, col }))
    .filter(({ count }) => count / populated >= 0.6)
    .map(({ col }) => col);
  const firstCategory = body.map((row) => row[0] ?? "").filter(Boolean);
  const nonNegative = numericColumns.every((col) =>
    body.every((row) => {
      const raw = (row[col] ?? "").replace(/[$,%\s,]/g, "");
      return raw === "" || Number(raw) >= 0;
    }),
  );
  return { body, numericColumns, firstCategory, nonNegative };
}

function orientChartMatrix(
  matrix: readonly (readonly string[])[],
  seriesInColumns: boolean,
): readonly (readonly string[])[] {
  if (seriesInColumns) return matrix;
  const width = Math.max(0, ...matrix.map((row) => row.length));
  return Array.from({ length: width }, (_, column) =>
    matrix.map((row) => row[column] ?? ""),
  );
}

/** True only when the current selection can execute the requested native kind. */
export function chartMatrixSupportsType(
  matrix: readonly (readonly string[])[],
  type: SpreadsheetChartType,
  seriesInColumns = true,
  profile?: ChartMatrixProfile,
): boolean {
  const oriented = orientChartMatrix(matrix, seriesInColumns);
  if (oriented.length === 0 || (oriented[0]?.length ?? 0) === 0) return false;
  // A supplied profile describes `matrix` as given; a transposed reading
  // still profiles its own orientation.
  const { body, numericColumns } =
    seriesInColumns && profile ? profile : profileChartMatrix(oriented);
  const width = Math.max(0, ...oriented.map((row) => row.length));
  const numericValues = numericColumns.filter((column) => column > 0);
  const valueColumns = width === 1 ? numericColumns : numericValues;
  const valueSeriesCount = valueColumns.length;
  switch (type) {
    case "scatter": {
      if (numericColumns.includes(0)) return numericColumns.length >= 2;
      // Excel plots a text X column at 1..n: a text first column beside a
      // value column is a scatter too (buildChartSeries numbers the points).
      const xs = body.map((row) => (row[0] ?? "").trim()).filter((value) => value !== "");
      return (
        xs.length > 0 &&
        xs.every((value) => !numeric(value)) &&
        numericColumns.some((column) => column > 0)
      );
    }
    case "bubble":
      return (
        width === 3 &&
        numericColumns.length === 3 &&
        numericColumns.includes(0) &&
        numericColumns.includes(1) &&
        numericColumns.includes(2)
      );
    case "combo":
      return valueSeriesCount >= 2;
    case "surface":
      // Excel plots a surface over at least two series and two categories.
      return valueSeriesCount >= 2 && body.length >= 2;
    case "pie":
    case "doughnut":
      // Excel charts the first value series of a wider selection as the pie
      // (and keeps the others, insertChartSelectionColumns). A negative
      // value is no refusal: Excel draws it as a slice of its size and its
      // label still reads the negative number (chartV2Echarts' ringSource).
      return valueSeriesCount >= 1;
    case "column":
    case "bar":
    case "line":
    case "area":
    case "radar":
      return valueSeriesCount >= 1;
    default:
      return false;
  }
}

/**
 * Why a selection cannot make a `type` chart, for the message the Insert
 * button shows: a combination or a surface needs a second series (Excel's
 * Combo pane: "To create a combo chart, select at least two series of
 * data"), a surface a second category, a bubble chart its X, Y and size
 * columns; anything else is a selection with nothing to plot. A single value
 * column was told it had no numbers to chart (S-079).
 */
export type ChartMatrixShortfall = "no-data" | "two-series" | "two-categories" | "bubble-columns" | "negative-values";
export function chartMatrixShortfall(
  matrix: readonly (readonly string[])[],
  type: SpreadsheetChartType,
  seriesInColumns = true,
  profile?: ChartMatrixProfile,
): ChartMatrixShortfall {
  const oriented = orientChartMatrix(matrix, seriesInColumns);
  if (oriented.length === 0 || (oriented[0]?.length ?? 0) === 0) return "no-data";
  const { body, numericColumns } =
    seriesInColumns && profile ? profile : profileChartMatrix(oriented);
  if (numericColumns.length === 0) return "no-data";
  const width = Math.max(0, ...oriented.map((row) => row.length));
  const valueSeriesCount = (width === 1 ? numericColumns : numericColumns.filter((column) => column > 0)).length;
  if (type === "combo" && valueSeriesCount === 1) return "two-series";
  if (type === "surface") {
    if (valueSeriesCount === 1) return "two-series";
    if (valueSeriesCount >= 2 && body.length < 2) return "two-categories";
  }
  if (type === "bubble") return "bubble-columns";
  // A hierarchy chart plots no negative value: a selection with one is told
  // so, not that it has no numbers (S-079).
  const negativeIn = (column: number) =>
    body.some((row) => {
      const raw = (row[column] ?? "").replace(/[$,%\s,]/g, "");
      return raw !== "" && Number(raw) < 0;
    });
  const valueColumns = width === 1 ? numericColumns : numericColumns.filter((column) => column > 0);
  if ((type === "treemap" || type === "sunburst") && valueColumns.some(negativeIn)) return "negative-values";
  return "no-data";
}

/**
 * Deterministic local counterpart to Excel's Recommended Charts analysis. It
 * reads only the selected display matrix and returns suitable, supported chart
 * families; no recommendation is a decorative placeholder.
 */
export function recommendChartTypes(
  matrix: readonly (readonly string[])[],
  profile?: ChartMatrixProfile,
): readonly ChartRecommendation[] {
  if (matrix.length === 0 || (matrix[0]?.length ?? 0) === 0) {
    return [{ type: "bar", reason: "general" }];
  }
  const { body, numericColumns, firstCategory, nonNegative } =
    profile ?? profileChartMatrix(matrix);
  const timeSeries =
    firstCategory.length > 1 &&
    firstCategory.filter(dateLike).length / firstCategory.length >= 0.6;

  const picks: ChartRecommendation[] = [];
  const add = (type: SpreadsheetChartType, reason: ChartRecommendation["reason"]) => {
    if (!picks.some((entry) => entry.type === type)) picks.push({ type, reason });
  };

  if (timeSeries && numericColumns.length >= 1) {
    add("line", "time-series");
    add("area", "time-series");
  }
  if (numericColumns.length >= 3 && numericColumns[0] === 0) {
    add("bubble", "xy-pairs");
  }
  if (numericColumns.length >= 2 && numericColumns[0] === 0) {
    add("scatter", "xy-pairs");
  }
  if (
    numericColumns.length === 1 &&
    firstCategory.length >= 2 &&
    firstCategory.length <= 8 &&
    nonNegative
  ) {
    add("pie", "part-to-whole");
    add("doughnut", "part-to-whole");
  }
  if (numericColumns.length >= 2) {
    add("combo", "multi-series");
    add("stacked-column", "multi-series");
  }
  if (numericColumns.length === 1 && body.length >= 12) {
    add("histogram", "distribution");
  }
  add("column", "general");
  add("line", "general");
  add("bar", "general");
  return picks.slice(0, 4);
}
