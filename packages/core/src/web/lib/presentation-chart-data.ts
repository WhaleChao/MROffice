/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation chart block data editing (PowerPoint / WPS parity subset).
 *
 * Pure, side-effect-free helpers shared by the presentation editor (the inline
 * data grid + chart-type selector) and the schema/serialize path so the bounds
 * (max rows / cols / cell length), the editable-grid shape, and the row/cell
 * mutations live in exactly one place. Keep this module dependency-free (no
 * React / no DOM) so it can be unit tested in isolation.
 *
 * The chart is rendered read-only by the existing `MiniChart` component; this
 * module only shapes the underlying `{ headers, rows }` data and the allowed
 * `chartType` vocabulary — it never renders. `MiniChart` already renders a much
 * larger family of kinds when called with `extended` (doughnut / scatter /
 * stacked / combo / radar / bubble / histogram / waterfall / box-and-whisker /
 * pareto / treemap / sunburst / stock), so the presentation chart picker now
 * exposes that FULL set (matching the spreadsheet floating-chart picker) and the
 * presentation render passes `extended` so each type renders as itself rather
 * than folding back to bar/line/pie.
 */

/** A flat, fully-rectangular editable grid: a header row + value rows. */
export interface ChartDataGrid {
  headers: string[];
  rows: string[][];
}

export interface PresentationChartSeries {
  name: string;
  values: number[];
}

export function presentationChartHasNativeSource(block: {
  source?: unknown;
}): boolean {
  const source = block.source;
  return Boolean(
    source &&
      typeof source === "object" &&
      !Array.isArray(source) &&
      (typeof (source as Record<string, unknown>).slideId === "string" ||
        typeof (source as Record<string, unknown>).slidePart === "string"),
  );
}

/**
 * The families Office stores as a `cx:chartSpace` part (the server's
 * `office_chart_parts.CHART_EX_FAMILIES`). The shared classic chart editor does
 * not write those parts; the ChartEx writer does, taking a title, a legend
 * (`presentationChartCanEditNativeTitleAndLegend`) and the values of the grid
 * the part's own caches project (`office_chart_edit.edit_extended_chart_data`).
 */
export const OFFICE_2016_CHART_FAMILIES: ReadonlySet<string> = new Set([
  "hierarchy",
  "statistical",
  "waterfall",
  "funnel",
  "map",
]);

/** Inclusive bounds, mirroring the chart-data zod schema in native-page-schemas. */
export const MAX_CHART_COLUMNS = 20;
export const MAX_CHART_ROWS = 200;
export const MAX_CHART_CELL_LENGTH = 2_000;
export const MAX_CHART_HEADER_LENGTH = 500;
/** Complete, reversible subset currently authenticated by the PPTX importer. */
export const MAX_NATIVE_CHART_SERIES = 12;
export const MAX_NATIVE_CHART_CATEGORIES = 50;

export function nativeChartGridFitsEditableProjection(
  grid: ChartDataGrid,
): boolean {
  const current = chartDataToGrid(grid);
  return (
    current.headers.length - 1 <= MAX_NATIVE_CHART_SERIES &&
    current.rows.length <= MAX_NATIVE_CHART_CATEGORIES
  );
}

/**
 * The chart types the editor offers. `MiniChart` (called with `extended`)
 * renders each of these as its own kind — the presentation picker exposes the
 * FULL set, the same family the spreadsheet floating-chart picker offers, so an
 * author picks a familiar PowerPoint / Excel name. `bar` is the implicit
 * default. The trailing group (doughnut … stock) are the extended kinds that
 * historically folded back to bar/line/pie and now render distinctly.
 */
export const CHART_TYPES = [
  "bar",
  "column",
  "line",
  "area",
  "pie",
  "scatter",
  "doughnut",
  "stacked",
  "combo",
  "radar",
  "bubble",
  "histogram",
  "waterfall",
  "box-and-whisker",
  "pareto",
  "treemap",
  "sunburst",
  "stock",
] as const;

export type ChartType = (typeof CHART_TYPES)[number];

/** The chart type used when a block has no explicit (or invalid) value. */
export const DEFAULT_CHART_TYPE: ChartType = "bar";

const CHART_TYPE_SET: ReadonlySet<string> = new Set(CHART_TYPES);

/** Type guard: is `value` exactly one of the offered chart types? */
export function isChartType(value: unknown): value is ChartType {
  return typeof value === "string" && CHART_TYPE_SET.has(value);
}

/**
 * Normalize an arbitrary persisted chart type into one of {@link CHART_TYPES}
 * for the SELECTOR's value. Unknown values (including legacy import names like
 * `barChart`, `lineChart`, `pieChart`) map to the closest offered type so the
 * selector always shows a valid option; everything else falls back to the
 * default. Case-insensitive, trims whitespace. NB: this is only for the editor
 * control — the persisted `chartType` string is preserved verbatim elsewhere so
 * an imported deck's exact label round-trips (MiniChart reads it directly).
 */
export function normalizeChartType(value: unknown): ChartType {
  if (typeof value !== "string") return DEFAULT_CHART_TYPE;
  const candidate = value.trim().toLowerCase();
  if (!candidate) return DEFAULT_CHART_TYPE;
  if (isChartType(candidate)) return candidate;
  // Legacy / import substring matches, ordered to MIRROR MiniChart's own
  // extended `normalizeType` so the selector shows the kind MiniChart actually
  // renders. The order matters (bubble before scatter; box-and-whisker before a
  // bare "box"; doughnut/donut before the legacy pie fold).
  if (candidate.includes("bubble")) return "bubble";
  if (candidate.includes("scatter") || candidate === "xy") return "scatter";
  if (candidate.includes("histogram") || candidate.includes("histo")) return "histogram";
  if (candidate.includes("waterfall") || candidate.includes("bridge")) return "waterfall";
  if (
    candidate.includes("boxplot") ||
    candidate.includes("box plot") ||
    (candidate.includes("box") && candidate.includes("whisker")) ||
    candidate === "box-and-whisker"
  ) {
    return "box-and-whisker";
  }
  if (candidate.includes("pareto")) return "pareto";
  if (candidate.includes("treemap") || candidate.includes("tree map") || candidate.includes("tree-map")) {
    return "treemap";
  }
  if (candidate.includes("sunburst") || candidate.includes("sun burst") || candidate.includes("sun-burst")) {
    return "sunburst";
  }
  if (
    candidate.includes("stock") ||
    candidate.includes("candlestick") ||
    candidate === "ohlc" ||
    candidate === "hlc"
  ) {
    return "stock";
  }
  if (candidate.includes("stack")) return "stacked";
  if (candidate.includes("combo") || candidate.includes("bar+line") || candidate.includes("bar-line")) {
    return "combo";
  }
  if (candidate.includes("radar") || candidate.includes("spider")) return "radar";
  if (candidate.includes("doughnut") || candidate.includes("donut")) return "doughnut";
  if (candidate.includes("pie")) return "pie";
  if (candidate.includes("area")) return "area";
  if (candidate.includes("line")) return "line";
  if (candidate.includes("column")) return "column";
  return DEFAULT_CHART_TYPE;
}

/** Clamp a cell to the value-cell length bound. */
function clampCell(value: unknown): string {
  return String(value ?? "").slice(0, MAX_CHART_CELL_LENGTH);
}

/** Clamp a header to the header length bound. */
function clampHeader(value: unknown): string {
  return String(value ?? "").slice(0, MAX_CHART_HEADER_LENGTH);
}

/**
 * Resolve a chart block's persisted `{ headers?, rows? }` into a rectangular,
 * fully-populated editable grid. Guarantees at least 2 columns and 1 row so the
 * grid is never empty; pads short rows and trims long ones to the column count;
 * applies a default header label per column. Bounded to the max column count.
 *
 * The header labels default to `Category` (col 0) then `Series N` so a brand-new
 * chart's grid reads sensibly. Pure — used both to seed the editor and to test
 * the mutations deterministically.
 */
export function chartDataToGrid(
  data: { headers?: readonly string[]; rows?: readonly (readonly string[])[] } | undefined,
): ChartDataGrid {
  const sourceHeaders = Array.isArray(data?.headers) ? data!.headers : [];
  const sourceRows = Array.isArray(data?.rows) ? data!.rows : [];
  const columnCount = Math.min(
    MAX_CHART_COLUMNS,
    Math.max(
      2,
      sourceHeaders.length,
      ...sourceRows.map((row) => (Array.isArray(row) ? row.length : 0)),
    ),
  );
  const headers = Array.from({ length: columnCount }, (_, index) =>
    clampHeader(sourceHeaders[index] || (index === 0 ? "Category" : `Series ${index}`)),
  );
  const rows = (sourceRows.length ? sourceRows : [[]])
    .slice(0, MAX_CHART_ROWS)
    .map((row) =>
      Array.from({ length: columnCount }, (_, index) =>
        clampCell(Array.isArray(row) ? row[index] : undefined),
      ),
    );
  return { headers, rows };
}

/**
 * Clamp a free-form grid (e.g. the output of a mutation) to the persisted
 * bounds: at most {@link MAX_CHART_COLUMNS} columns, {@link MAX_CHART_ROWS}
 * rows, each cell/header clamped to its length bound. Rows are padded/trimmed to
 * the header column count so the persisted data stays rectangular.
 */
export function clampChartGrid(grid: ChartDataGrid): ChartDataGrid {
  const headers = grid.headers.slice(0, MAX_CHART_COLUMNS).map(clampHeader);
  const columnCount = Math.max(1, headers.length);
  const rows = grid.rows
    .slice(0, MAX_CHART_ROWS)
    .map((row) =>
      Array.from({ length: columnCount }, (_, index) => clampCell(row[index])),
    );
  return { headers, rows };
}

/** Set a single header cell (by index) to `value`, clamped. */
export function setChartHeader(
  grid: ChartDataGrid,
  index: number,
  value: string,
): ChartDataGrid {
  return clampChartGrid({
    ...grid,
    headers: grid.headers.map((header, i) => (i === index ? value : header)),
  });
}

/** Set a single data cell (row, col) to `value`, clamped. */
export function setChartCell(
  grid: ChartDataGrid,
  rowIndex: number,
  colIndex: number,
  value: string,
): ChartDataGrid {
  return clampChartGrid({
    ...grid,
    rows: grid.rows.map((row, r) =>
      r === rowIndex
        ? row.map((cell, c) => (c === colIndex ? value : cell))
        : row,
    ),
  });
}

/** Append an empty data row (one cell per column). No-op at the row cap. */
export function addChartRow(grid: ChartDataGrid): ChartDataGrid {
  if (grid.rows.length >= MAX_CHART_ROWS) return clampChartGrid(grid);
  return clampChartGrid({
    ...grid,
    rows: [...grid.rows, Array.from({ length: grid.headers.length }, () => "")],
  });
}

/**
 * Remove the data row at `index`. Always keeps at least one row so the grid is
 * never empty (removing the last row clears it to a single blank row).
 */
export function removeChartRow(grid: ChartDataGrid, index: number): ChartDataGrid {
  const rows = grid.rows.filter((_, i) => i !== index);
  return clampChartGrid({
    ...grid,
    rows: rows.length ? rows : [Array.from({ length: grid.headers.length }, () => "")],
  });
}

/** Append a new series column (header + an empty cell on every row). No-op at the column cap. */
export function addChartSeries(grid: ChartDataGrid): ChartDataGrid {
  if (grid.headers.length >= MAX_CHART_COLUMNS) return clampChartGrid(grid);
  return clampChartGrid({
    headers: [...grid.headers, `Series ${grid.headers.length}`],
    rows: grid.rows.map((row) => [...row, ""]),
  });
}

export function switchChartRowsAndColumns(
  grid: ChartDataGrid,
): ChartDataGrid | null {
  const current = chartDataToGrid(grid);
  const nextColumnCount = current.rows.length + 1;
  const nextRowCount = current.headers.length - 1;
  if (
    nextColumnCount > MAX_CHART_COLUMNS ||
    nextRowCount < 1 ||
    nextRowCount > MAX_CHART_ROWS
  ) {
    return null;
  }
  return clampChartGrid({
    headers: [current.headers[0] || "Category", ...current.rows.map((row) => row[0] || "")],
    rows: current.headers.slice(1).map((seriesName, seriesIndex) => [
      seriesName,
      ...current.rows.map((row) => row[seriesIndex + 1] || ""),
    ]),
  });
}

export function chartGridToCategoriesAndSeries(grid: ChartDataGrid): {
  categories: string[];
  series: PresentationChartSeries[];
} {
  const current = chartDataToGrid(grid);
  return {
    categories: current.rows.map((row) => row[0] || ""),
    series: current.headers.slice(1).map((name, seriesIndex) => ({
      name,
      values: current.rows.map((row) => {
        const value = Number(row[seriesIndex + 1]);
        return Number.isFinite(value) ? value : 0;
      }),
    })),
  };
}

export function categoriesAndSeriesToChartGrid(
  categories: readonly string[],
  series: readonly PresentationChartSeries[],
): ChartDataGrid {
  return chartDataToGrid({
    headers: ["Category", ...series.map((entry) => entry.name)],
    rows: categories.map((category, categoryIndex) => [
      category,
      ...series.map((entry) => String(entry.values[categoryIndex] ?? 0)),
    ]),
  });
}

/** A table cell's text as a chart value: thousands separators, a unit sign
 * or a trailing percent are dropped; anything else is not a number. */
export function chartValueFromCellText(text: string): number | null {
  const cleaned = text
    .replace(/[\s,，]/g, "")
    .replace(/^[¥$€£]/, "")
    .replace(/[%％]$/, "");
  if (!/^[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?$/i.test(cleaned)) return null;
  const value = Number(cleaned);
  return Number.isFinite(value) ? value : null;
}
