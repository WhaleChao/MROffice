/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "../office-units";
import { MAX_DRAWING_LAYER_X_PX, MAX_DRAWING_LAYER_Y_PX, MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";
import {
  OFFICE_TRENDLINE_ORDER,
  normalizeOfficeTrendline,
  officeTrendlineLabel,
  officeTrendlinePolyline,
  type OfficeTrendline,
  type OfficeTrendlineOptions,
  type OfficeTrendlineType,
} from "../office-charts/trendline";
import { normalizeChartVisibility, type ChartVisibility } from "./chart-visibility";
import {
  normalizePivotChartOptions,
  type SpreadsheetPivotChartOptions,
} from "./pivot-chart";
import { columnToLetters, lettersToColumn } from "./formula-engine";
import {
  CHART_COLOR_PALETTE_IDS,
  CHART_COLOR_PALETTES,
  chartColorPaletteColors,
  isChartColorPaletteId,
  type ChartColorPaletteId,
} from "../office-charts/chart-color-palettes";
import { normalizeNumberFormatCode } from "./number-format";
import { MAX_CHART_POINTS, MAX_CHART_SERIES } from "../office-charts/chart-limits.generated";
import { normalizeSpreadsheetChartFormat, type SpreadsheetChartFormat } from "./chart-format";
import {
  CHART_STYLE_PRESET_FOR_C_STYLE,
  CHART_STYLE_PRESET_IDS,
  CHART_STYLE_PRESETS,
  DEFAULT_CHART_STYLE_PRESET_ID,
  type ChartStylePresetId,
} from "../office-charts/chart-style-presets.generated";
import {
  MAX_CHART_CATEGORY_LEVELS,
  isComboRenderAs,
  normalizeChartBlockAttrsV2Wire,
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ComboRenderAs,
} from "../../components/console/editor/extensions/chartBlockShared";
import {
  normalizePresentationChartSeriesStyles,
  normalizePresentationChartSliceStyles,
  type PresentationChartSeriesStyle,
  type PresentationChartSliceStyle,
} from "../presentation-chart-area-style";

/**
 * Pure helpers for floating CHART objects on a native spreadsheet sheet
 * (Excel / Sheets parity).
 *
 * A chart here is a FLOATING object drawn OVER the grid (like a slicer panel,
 * not an in-cell sparkline): it is bound to a rectangular SOURCE RANGE of cells
 * and re-renders live whenever those cells change. The sheet model stores a
 * small `charts` array (one entry per chart). Each chart carries its kind
 * (bar / line / pie / area), the {r1,c1,r2,c2} source range, an anchor cell it
 * floats near, an optional pixel size + title, and a `seriesInColumns` flag that
 * decides whether labels run down the first COLUMN (series in columns) or across
 * the first ROW (series in rows). Selected values are read by the host from the
 * live grid and passed to `buildChartSeries`.
 *
 * NOT a chart-config edit (see chart-config.ts, which edits the legacy chart
 * BLOCK object's type/range/legend): this is the floating, range-bound chart the
 * "Insert chart" toolbar action creates from the current selection.
 *
 * This module is dependency-free and deterministic so it can be unit-tested in
 * isolation: it carries the chart MODEL, the normalize/persist coercion (a
 * persisted value can come from an imported file, an AI write, or a hand-edited
 * JSON, so every field is clamped/dropped rather than trusted), and the pure
 * data mapping (`buildChartSeries`) that turns a range of display strings into
 * the [label, value] rows MiniChart consumes. The React layer reads the cell
 * values, calls these helpers, and renders the result; it owns no chart math.
 */

/**
 * The chart kinds a floating chart draws.
 *
 * The first four are SINGLE-series (label + one value column). The rest are the
 * extended kinds:
 *  - `scatter`: XY points — the first numeric column is X, the second is Y (no
 *    category axis); degrades to the bar fallback below when fewer than 2 numeric
 *    columns exist.
 *  - `stacked-bar` / `stacked-column`: cumulative stacking of every value column
 *    (synonyms here — both render vertical stacked columns; the names mirror
 *    Excel's "stacked bar" / "stacked column" picker labels).
 *  - `combo`: the first value column draws as bars, the second as a line overlay.
 *  - `doughnut`: a pie with a hollow center (single-series, like pie).
 *  - `radar`: a spider/polygon plot of each value column across the categories.
 *  - `bubble`: scatter (first value column X, second Y) where a THIRD value column
 *    sets each point's bubble radius; degrades to a plain scatter when no third
 *    numeric column exists.
 *  - `histogram`: bins the FIRST numeric value column into N buckets and draws the
 *    per-bucket count as bars (a distribution chart, no category axis).
 *  - `waterfall`: running cumulative bars — each bar starts where the previous one
 *    ended (up/down colored) with an optional final total bar.
 *  - `box-and-whisker`: box plot per series — shows min, Q1, median, Q3, max
 *    (Excel QUARTILE.INC / linear-interpolation), IQR whiskers (1.5 × IQR),
 *    outlier dots beyond the whiskers, and a mean × marker (Excel parity).
 *    NOTE: renders in-app only; xlsx chartEx export (cx:) is not yet supported
 *    by openpyxl — deferred.
 *  - `pareto`: categories sorted descending by value, rendered as bars plus a
 *    secondary cumulative-percentage line on a right 0–100 % axis (Excel parity).
 *    NOTE: renders in-app only; xlsx chartEx export (cx:) is not yet supported
 *    by openpyxl — deferred.
 *  - `treemap`: hierarchical nested rectangles sized proportionally by value
 *    (Excel 2016 treemap chart parity). Accepts flat input where the last numeric
 *    column is the size; multi-level category columns are supported for hierarchy
 *    building. Rendered with a slice-and-dice layout (single-level: squarified).
 *    NOTE: renders in-app only — single-level flat rectangles squarified by value.
 *    Multi-level hierarchy (Excel multi-column category paths) is limited to a
 *    flat grouping of the first category column. xlsx chartEx (cx:) export is not
 *    yet supported by openpyxl — deferred.
 *  - `sunburst`: same hierarchy as treemap but radial — concentric ring segments,
 *    inner ring = top level, angular sweep proportional to value. With single-level
 *    data renders as a doughnut ring. xlsx chartEx (cx:) export is not yet
 *    supported by openpyxl — deferred.
 *  - `stock`: candlestick (OHLC) chart for financial/price data. Expects 4 numeric
 *    series in the order Open, High, Low, Close (plus a category/date column). Each
 *    category gets a candlestick: a high-low wick plus an open-close box, colored
 *    up (close >= open) or down (close < open). Supports the High-Low-Close (3
 *    series) and Open-High-Low-Close (4 series) Excel variants. xlsx export via
 *    openpyxl's classic StockChart is not yet wired — deferred.
 */
export type SpreadsheetChartType =
  | "column"
  | "bar"
  | "line"
  | "pie"
  | "area"
  | "scatter"
  | "stacked-bar"
  | "stacked-column"
  | "combo"
  | "doughnut"
  | "radar"
  | "bubble"
  | "histogram"
  | "waterfall"
  | "box-and-whisker"
  | "pareto"
  | "treemap"
  | "sunburst"
  | "stock"
  | "funnel"
  | "map"
  // A surface chart: the height matrix is the selected range's series (a
  // series per row), as Excel reads it.
  | "surface"
  // "camera" is NOT a chart kind — it is the Excel Camera tool (a live picture of
  // a cell range) reusing the chart object's float/drag/resize/remap/persist
  // infrastructure. It is deliberately ABSENT from SPREADSHEET_CHART_TYPES so the
  // chart-type picker never offers it; the renderer branches on it separately.
  | "camera";

export const SPREADSHEET_CHART_TYPES: readonly SpreadsheetChartType[] = [
  "column",
  "bar",
  "line",
  "pie",
  "area",
  "scatter",
  "stacked-bar",
  "stacked-column",
  "combo",
  "doughnut",
  "radar",
  "bubble",
  "histogram",
  "waterfall",
  "box-and-whisker",
  "pareto",
  "treemap",
  "sunburst",
  "stock",
  "funnel",
  "map",
  "surface",
] as const;

/**
 * Bounded Excel chart-style gallery. The numeric IDs are the native `c:style`
 * values written to XLSX; colours here are the Web preview for the same
 * theme-driven choice, not a replacement for the workbook theme.
 *
 * The worksheet no longer reads this table: its gallery, its canvas and its
 * model read `c:style` through the one preset table
 * ({@link SPREADSHEET_CHART_STYLE_GALLERY}). It stays only for the deck's
 * legacy chart display, which phase 1c of the chart unification removes.
 */
export const SPREADSHEET_CHART_STYLE_PRESETS = Object.freeze([
  {
    id: 2,
    background: "#ffffff",
    plotBackground: "#ffffff",
    text: "#334155",
    gridline: "#e2e8f0",
    border: "#cbd5e1",
    palette: ["#0d9488", "#2563eb", "#f59e0b", "#db2777"],
  },
  {
    id: 4,
    background: "#f8fafc",
    plotBackground: "#ffffff",
    text: "#1e293b",
    gridline: "#cbd5e1",
    border: "#94a3b8",
    palette: ["#2563eb", "#0284c7", "#4f46e5", "#0891b2"],
  },
  {
    id: 8,
    background: "#fff7ed",
    plotBackground: "#fffbeb",
    text: "#7c2d12",
    gridline: "#fed7aa",
    border: "#fdba74",
    palette: ["#ea580c", "#dc2626", "#d97706", "#0f766e"],
  },
  {
    // Excel's default (style 201 over `c:style` 10): text and hairlines in the
    // theme's text colour at 65 % and 15 %, series in the theme's accents.
    id: 10,
    background: "#ffffff",
    plotBackground: "#ffffff",
    text: "#595959",
    gridline: "#d9d9d9",
    border: "#d9d9d9",
    palette: ["#0d9488", "#2563eb", "#f59e0b", "#db2777"],
  },
  {
    id: 11,
    background: "#0f172a",
    plotBackground: "#111827",
    text: "#f8fafc",
    gridline: "#334155",
    border: "#475569",
    palette: ["#22d3ee", "#fbbf24", "#34d399", "#a78bfa"],
  },
  {
    id: 12,
    background: "#f8fafc",
    plotBackground: "#ffffff",
    text: "#18181b",
    gridline: "#d4d4d8",
    border: "#a1a1aa",
    palette: ["#18181b", "#52525b", "#71717a", "#a1a1aa"],
  },
] as const);

export type SpreadsheetChartStyleId =
  (typeof SPREADSHEET_CHART_STYLE_PRESETS)[number]["id"];
export type SpreadsheetChartStylePreset =
  (typeof SPREADSHEET_CHART_STYLE_PRESETS)[number];
export const DEFAULT_SPREADSHEET_CHART_STYLE_ID: SpreadsheetChartStyleId = 10;

/**
 * The worksheet's chart style gallery: the ONE preset table the Docs and deck
 * galleries, the canvas and every writer share
 * (`office-charts/chart-style-presets.generated`, from the server's
 * `office_chart_style_presets`). Each entry's id is the `c:style` a classic
 * part is written with and read back by; its preset is the chart model's
 * `stylePreset`. Text and hairlines draw in Office's default style greys
 * (style 201), as phase 0 of the chart unification states for every host.
 */
export const SPREADSHEET_CHART_STYLE_GALLERY = Object.freeze(
  CHART_STYLE_PRESET_IDS.map((stylePreset) => ({
    id: CHART_STYLE_PRESETS[stylePreset].cStyle,
    stylePreset,
    background: "#ffffff",
    plotBackground: "#ffffff",
    text: "#595959",
    gridline: "#d9d9d9",
    border: "#d9d9d9",
    palette: CHART_STYLE_PRESETS[stylePreset].palette,
  })),
);
export type SpreadsheetChartStyleGalleryEntry = (typeof SPREADSHEET_CHART_STYLE_GALLERY)[number];
/** The `c:style` of Office's default chart style: the table's default preset. */
export const DEFAULT_SPREADSHEET_CHART_STYLE_GALLERY_ID: number =
  CHART_STYLE_PRESETS[DEFAULT_CHART_STYLE_PRESET_ID].cStyle;

/** The chart model's `stylePreset` a sheet chart's `c:style` reads as (the one table). */
export function spreadsheetChartStylePresetId(raw: unknown): ChartStylePresetId {
  const style = normalizeSpreadsheetChartStyle(raw);
  return (style !== undefined ? CHART_STYLE_PRESET_FOR_C_STYLE[String(style)] : undefined) ?? DEFAULT_CHART_STYLE_PRESET_ID;
}

/** A sheet chart's gallery look: its `c:style`'s preset, else Office's default. */
export function spreadsheetChartStyleAppearance(raw: unknown): SpreadsheetChartStyleGalleryEntry {
  const preset = spreadsheetChartStylePresetId(raw);
  return SPREADSHEET_CHART_STYLE_GALLERY.find((entry) => entry.stylePreset === preset)!;
}

/**
 * Excel's Change Colors gallery is independent from the numeric `c:style`
 * gallery.  Persist only a bounded palette identity: colourful palettes rotate
 * the workbook's six accent slots, while monochromatic palettes reuse one
 * accent with Office's luminance steps.  The Web renderer resolves the
 * current workbook theme at paint time; XLSX writers emit the same schemeClr
 * slots so switching the workbook theme continues to recolour the chart.
 *
 * The table is the one the deck and Docs offer too, generated from the
 * server's copy (`office-charts/chart-color-palettes.generated.ts`); each
 * token's `lumMod` / `lumOff` is resolved by the one colour rule, as Excel
 * resolves it, rather than by an RGB mix toward white.
 */
export const SPREADSHEET_CHART_COLOR_PALETTES = Object.freeze(
  CHART_COLOR_PALETTE_IDS.map((id) => ({
    id,
    group: CHART_COLOR_PALETTES[id].group,
    tokens: CHART_COLOR_PALETTES[id].tokens,
  })),
);

export type SpreadsheetChartColorPaletteId = ChartColorPaletteId;
export type SpreadsheetChartColorPalette =
  (typeof SPREADSHEET_CHART_COLOR_PALETTES)[number];

export function normalizeSpreadsheetChartColorPalette(
  raw: unknown,
): SpreadsheetChartColorPaletteId | undefined {
  return isChartColorPaletteId(raw) ? raw : undefined;
}

/**
 * The colours a pick gives a chart that paints `count` series (a pie's
 * points); six, the gallery's row, when the caller does not say. A
 * Monochromatic palette spreads its shades over exactly that many.
 */
export function spreadsheetChartColorPaletteColors(
  raw: unknown,
  themeAccents: readonly string[],
  count?: number,
): readonly string[] | null {
  const id = normalizeSpreadsheetChartColorPalette(raw);
  if (!id || themeAccents.length === 0) return null;
  return chartColorPaletteColors(
    id,
    themeAccents,
    count !== undefined && Number.isInteger(count) && count > 0 ? count : undefined,
  );
}

/** Preserve any valid native style identity, including styles outside our gallery. */
export function normalizeSpreadsheetChartStyle(raw: unknown): number | undefined {
  return typeof raw === "number" &&
    Number.isInteger(raw) &&
    raw >= 1 &&
    raw <= 48
    ? raw
    : undefined;
}

/** Resolve the bounded Web preview without changing an imported native ID. */
export function spreadsheetChartStylePreset(
  raw: unknown,
): SpreadsheetChartStylePreset {
  const normalized = normalizeSpreadsheetChartStyle(raw);
  return (
    SPREADSHEET_CHART_STYLE_PRESETS.find(
      (preset) => preset.id === normalized,
    ) ??
    SPREADSHEET_CHART_STYLE_PRESETS.find(
      (preset) => preset.id === DEFAULT_SPREADSHEET_CHART_STYLE_ID,
    )!
  );
}

/**
 * The set of valid types, for the normalizer's pass-through check. A persisted
 * value that already equals a canonical kind survives unchanged (back-compat for
 * saved charts); anything else is coerced by the alias rules below.
 */
const VALID_CHART_TYPES: ReadonlySet<string> = new Set<string>(
  SPREADSHEET_CHART_TYPES,
);

/** Max charts a sheet may carry (DoS bound on iteration / overlay render). */
export const MAX_CHARTS = 64;
/** Max value-series indices a chart may pin to the secondary axis (DoS bound). */
const MAX_SECONDARY_AXIS_SERIES = 64;
/** Max length of a chart id / title (DoS bound on a hand-edited value). */
export const MAX_CHART_ID_LENGTH = 200;
export const MAX_CHART_TITLE_LENGTH = 2_000;
/** Grid bounds mirror the schema clamps so a normalized range stays in-grid. */
// The last cell of Excel's grid: a chart may read any rows of a sheet (a
// large workbook's reach far past 10,000).
const MAX_RANGE_ROW = MAX_GRID_ROWS - 1;
const MAX_RANGE_COL = MAX_GRID_COLUMNS - 1;
/** Max index of a value series a chart may name (DoS bound). */
const MAX_SERIES_INDEX = 999;
const MIN_CHART_SIZE = 80;
const MAX_CHART_SIZE = 4_000;
const DEFAULT_CHART_WIDTH = 320;
const DEFAULT_CHART_HEIGHT = 200;
/**
 * Interactive drag-RESIZE bounds (the user-facing min the resize grip enforces).
 * Wider than the persist clamp's floor (MIN_CHART_SIZE) so a hand-shrunk box
 * never collapses below a legible chart, and capped at the same MAX so the model
 * clamp stays the outer authority. The persist clamp (`clampSize`) is untouched
 * — a value imported below this min still round-trips, it just can't be REACHED
 * by dragging.
 */
export const MIN_CHART_RESIZE_WIDTH = 160;
export const MIN_CHART_RESIZE_HEIGHT = 100;
export const MAX_CHART_RESIZE = MAX_CHART_SIZE;

/** A 0-based inclusive cell rectangle (the chart's source data range). */
export interface ChartSourceRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * An optional Excel-style TRENDLINE overlaid on the primary numeric series.
 *  - `linear`: an ordinary least-squares line fit over (x, y) — y is the series
 *    value and x is the point index 0..n-1 (or the scatter X when the chart is a
 *    scatter). The overlay is the straight line from the first to the last x.
 *  - `movingAverage`: a trailing average of the last `period` points (Excel's
 *    "Moving Average" trendline). `period` clamps to [2, n]; default 2. Only the
 *    moving-average kind reads `period`; linear ignores it.
 */
export type ChartTrendlineType = OfficeTrendlineType;

/**
 * The six kinds of Office's Trendline menu, shared with Docs and the deck
 * (`lib/office-charts/trendline`, which also owns the fit). The sheet's own
 * spelling keeps `period` optional (a moving average without one is Office's
 * two-period default) and adds a polynomial's `order` the same way.
 */
export interface ChartTrendlineSpec extends OfficeTrendlineOptions {
  type: ChartTrendlineType;
  /** Window for movingAverage (>=2); ignored for every other kind. */
  period?: number;
  /** Order for polynomial (2–6); ignored for every other kind. */
  order?: number;
}

/** The shared contract's spelling of a sheet trendline (defaults filled in). */
export function spreadsheetTrendlineContract(
  spec: ChartTrendlineSpec,
): OfficeTrendline | undefined {
  const { type, period, order, ...options } = spec;
  return normalizeOfficeTrendline(
    type === "polynomial"
      ? { type, order: order ?? OFFICE_TRENDLINE_ORDER.min, ...options }
      : type === "movingAverage"
        ? {
            type,
            period: Math.min(period ?? DEFAULT_TRENDLINE_PERIOD, 255),
            ...options,
          }
        : { type, ...options },
  );
}

/** Default trailing window for a moving-average trendline (Excel's minimum). */
export const DEFAULT_TRENDLINE_PERIOD = 2;
/** Clamp bound on a hand-edited moving-average window (DoS on iteration). */
const MAX_TRENDLINE_PERIOD = 1_000;

/** A plotted trendline point in DATA space (x = index or scatter-X, y = value). */
export interface TrendlinePoint {
  x: number;
  y: number;
}

/**
 * An optional Excel-style ERROR BAR drawn as a vertical whisker at each plotted
 * point/bar:
 *  - `fixed`: a constant +/- `value` in DATA units around every point.
 *  - `percent`: +/- `value`% of each point's own magnitude (|y| * value/100).
 *  - `stddev`: +/- `value` standard deviations of the plotted series (the same
 *    delta on every point; Excel's "Standard Deviation" error amount). When
 *    `value` is 0 the delta is the series' own population standard deviation
 *    (a faithful 1-σ default for a stddev bar with no explicit multiplier).
 * `value` is meaningful for fixed/percent (and as the σ multiplier for stddev);
 * it clamps to >= 0 so a stale/hand-edited negative never inverts the whisker.
 * DEFER: full Excel custom-range error bars (per-point custom +/- values) — the
 * three amount kinds here cover the common case.
 */
/** Excel's error amounts: Fixed Value, Percentage, Standard Deviation(s) and
 * Standard Error (computed from the series; its `value` is ignored). */
export type ChartErrorBarType = "fixed" | "percent" | "stddev" | "stderr";

export interface ChartErrorBarSpec {
  type: ChartErrorBarType;
  /** Error amount (>= 0): the +/- value for fixed/percent, σ multiplier for stddev. */
  value: number;
  /** Format Error Bars ▸ Direction: the one side drawn (absent: Both). */
  include?: "plus" | "minus";
  /** Format Error Bars ▸ End Style ▸ No Cap (absent: Cap). */
  noEndCap?: true;
}

/** The set of valid error-bar amount kinds (normalizer pass-through check). */
const VALID_ERROR_BAR_TYPES: ReadonlySet<string> = new Set<ChartErrorBarType>([
  "fixed",
  "percent",
  "stddev",
  "stderr",
]);
/** Clamp bound on a hand-edited error amount (DoS / sane-pixel guard). */
const MAX_ERROR_BAR_VALUE = 1_000_000_000;
/** Default amount the editor seeds when a user first picks an error-bar kind. */
export const DEFAULT_ERROR_BAR_VALUE = 1;

/**
 * Where a sheet chart's legend can sit — the four `ST_LegendPos` placements
 * Excel's Chart Elements menu offers and this renderer lays out. Excel also
 * writes a `tr` corner legend, which nothing here draws; a chart carrying one
 * keeps it and reports `nativeChartLegendPositionBlocker` instead.
 */
export const SPREADSHEET_CHART_LEGEND_POSITIONS = [
  "right",
  "top",
  "left",
  "bottom",
] as const;
/**
 * The Legend menu's four placements plus Excel's Format Legend ▸ Top Right
 * (`legendPos val="tr"`), which the sheet lays out and writes back but does
 * not offer from the menu, exactly as Excel's own Legend submenu does not.
 */
export type SpreadsheetChartLegendPosition =
  | SpreadsheetChartLegendMenuPosition
  | "topRight";

/** One of the Legend menu's four placements (never Top Right). */
export type SpreadsheetChartLegendMenuPosition =
  (typeof SPREADSHEET_CHART_LEGEND_POSITIONS)[number];

export function isSpreadsheetChartLegendMenuPosition(
  value: unknown,
): value is SpreadsheetChartLegendMenuPosition {
  return SPREADSHEET_CHART_LEGEND_POSITIONS.some(
    (position) => position === value,
  );
}

/**
 * A floating sheet-level chart bound to a source range. `id` identifies it.
 * `type` is the draw kind. `range` is the inclusive source rectangle. `anchor`
 * is the cell the chart floats near (it shifts on insert/delete so the chart
 * follows its data). `width`/`height` are the pixel size of the chart canvas.
 * `title` is an optional header. `seriesInColumns` (default true) decides the
 * orientation: when true the FIRST COLUMN is labels and each later column is a
 * series; when false the FIRST ROW is labels and each later row is a series.
 * `x`/`y` are the panel's pixel offset within the grid stage once the user has
 * DRAGGED it (optional; the renderer lays panels out by index when absent, so a
 * never-moved chart persists with no x/y key — byte-stable). Mirrors the Slicer
 * model: the anchor stays grid-bound (it remaps on insert/delete so the source
 * still tracks), while x/y are free pixel placement independent of the cells.
 */
// A type (not an interface) so a record reads as the loose chart shape the
// shared model helpers take (`SpreadsheetChartLike`, lib/spreadsheet/chart-v2-model).
export type SpreadsheetChart = {
  id: string;
  groupId?: string;
  /** Native DrawingML cNvPr name used by Excel's Navigation pane. */
  name?: string;
  /** Native DrawingML cNvPr hidden state used by Excel's Navigation pane. */
  hidden?: boolean;
  






  outerCategoryColumns?: number;
  /**
   * The range has no line of category labels: its first column (series in
   * columns; its first row, series in rows) is data like the others and the
   * categories are 1, 2, 3... -- Excel's Insert Chart over a selection whose
   * corner cell is filled and whose first column ends in a number (North and
   * South side by side, a Year column beside Sales), written with no
   * `c:cat`. Absent: the first column (row) is the categories.
   */
  noCategories?: true;
  /**
   * The range has no line of series names: its first row (series in columns;
   * its first column, series in rows) is data and each series is named by
   * its place (Series1, 系列1), as Excel inserts a chart over numbers with no
   * heading -- written with no `c:tx`. Absent: the first row names them.
   */
  noSeriesNames?: true;
  









  nameRowsAbove?: number;
  /** XLS-A081 Edit Alt Text: the chart's description (`cNvPr@descr`). */
  alt?: string;
  /** Older Office's alt-text title (`cNvPr@title`), kept as the file had it. */
  altTitle?: string;
  /** Alt Text > Mark as decorative (the frame's Office 2017 `adec:decorative`
   * extension): a decorative chart states no description. */
  decorative?: boolean;
  /** XLS-A081 Assign Macro: the macro a click on the chart runs. */
  macro?: string;
  /** Position in the sheet-wide paint order shared with pictures and shapes
   * (lib/spreadsheet/drawing-stack); absent = above every object, in array order. */
  stackOrder?: number;
  type: SpreadsheetChartType;
  range: ChartSourceRange;
  anchor: { row: number; col: number };
  width: number;
  height: number;
  title?: string;
  /** Native SpreadsheetML `c:style` identity (1..48). */
  chartStyle?: number;
  /** Desktop chart format atoms; data/range remains the worksheet authority. */
  chartFormat?: SpreadsheetChartFormat;
  chartSubtype?: string;
  desktopStyleId?: string;
  sourceSheetId?: string;
  nativeChartFormatBlocker?: string;
  /** Theme-relative Chart Design ▸ Change Colors gallery identity. */
  chartColorPalette?: SpreadsheetChartColorPaletteId;
  /** Explicit classic-chart series paint; manual srgb colours override themes. */
  seriesStyles?: readonly PresentationChartSeriesStyle[];
  /** Exact Spreadsheet-only two-stop left-to-right series gradients. */
  seriesGradientStyles?: readonly SpreadsheetChartSeriesGradientStyle[];
  /** Imported gradient paint stays visible but cannot be flattened by Change Colors. */
  nativeChartColorEditBlocker?: string;
  /** Per-series classic line marker, smooth state, and exact point width. */
  lineSeriesStyles?: readonly SpreadsheetChartLineSeriesStyle[];
  /** Imported line grammar that cannot be edited or rebuilt safely. */
  nativeChartLineBlocker?: string;
  /** Imported line width that cannot be rendered or rebuilt safely. */
  nativeChartLineWidthBlocker?: string;
  /** Explicit pie/doughnut data-point paint indexed by native point order. */
  sliceStyles?: readonly PresentationChartSliceStyle[];
  seriesInColumns?: boolean;
  /**
   * The chart's source lines were all deleted: like Excel's `=SERIES(#REF!…)`
   * the chart keeps its title, size, position and format but plots nothing
   * until Select Data picks a new range. `range` is then only a placeholder.
   */
  sourceRefError?: true;
  x?: number;
  y?: number;
  /**
   * The chart's cell anchor, the one every worksheet drawing has: Excel
   * inserts a chart Move and size with cells (a two-cell anchor), so it moves
   * and sizes with rows and columns inserted, deleted, resized, hidden or
   * sorted under it. When present it is where the chart is drawn; `x`/`y`/
   * `width`/`height` are its frame as last drawn (lib/spreadsheet/chart-anchor).
   */
  drawingAnchor?: SpreadsheetChartDrawingAnchor;
  /**
   * Optional Excel-style trendline overlaid on the primary numeric series
   * (linear OLS or trailing moving average). Persisted only when set, so a chart
   * without a trendline carries no key (byte-stable).
   */
  trendline?: ChartTrendlineSpec;
  /**
   * When true, MiniChart prints the numeric value next to each plotted point/bar
   * (Excel "Data Labels"). Off by default; persisted only when true so existing
   * charts stay byte-identical.
   */
  showDataLabels?: boolean;
  /** Show the chart legend. Absent means Excel's default (shown). */
  showLegend?: boolean;
  /** Supported Excel Web legend placements in this renderer. */
  legendPosition?: SpreadsheetChartLegendPosition;
  /** Imported legend grammar that cannot be projected without guessing. */
  nativeChartLegendPositionBlocker?: string;
  /** Show category/value axes and their labels. Absent means shown. */
  showAxisLabels?: boolean;
  /** Show primary major gridlines. Absent means shown. */
  showGridlines?: boolean;
  /** Show the chart's source values below the plot. Off by default. */
  showDataTable?: boolean;
  /**
   * Optional axis titles (Excel/WPS "Axis Titles"): `axisTitleX` labels the
   * horizontal (category) axis, `axisTitleY` the vertical (value) axis. Each is
   * persisted only when a non-empty string — a chart without axis titles carries
   * no key (byte-stable). Chart types without cartesian axes (pie / doughnut /
   * radar) ignore them on render.
   */
  axisTitleX?: string;
  axisTitleY?: string;
  /** Primary value-axis number format; linked mode resolves current source cells. */
  valueAxisNumberFormat?: SpreadsheetChartValueAxisNumberFormat;
  /** Imported ambiguous/multi-axis authority remains visible but non-authorable. */
  nativeValueAxisNumberFormatBlocker?: string;
  /**
   * Optional Excel-style error bars (vertical whiskers) drawn at each plotted
   * point/bar on the cartesian kinds (bar / line / scatter). Persisted only when
   * a valid spec survives normalization, so a chart without error bars carries no
   * key (byte-stable). Ignored on pie / doughnut / radar (no value whisker).
   */
  errorBars?: ChartErrorBarSpec;
  /** Show the first-to-last series delta bars on a multi-series line chart. */
  upDownBars?: boolean;
  /** Predefined chart-line elements owned by the classic chart family. */
  dropLines?: boolean;
  highLowLines?: boolean;
  seriesLines?: boolean;
  /**
   * Optional combo / dual-axis support: the 0-based VALUE-series indices plotted
   * against a SECONDARY (right-hand) Y scale instead of the primary left axis
   * (Excel's "Secondary Axis"). A series whose index is listed here is scaled to
   * its own right-side axis so a small-magnitude series stays readable beside a
   * large-magnitude one. Persisted only when a non-empty, validated index list
   * survives normalization (deduped + sorted), so a chart without a secondary
   * axis carries no key (byte-stable). Ignored on pie / doughnut / radar.
   */
  secondaryAxisSeries?: number[];
  /**
   * Combo > Custom Combination on a range-bound chart: how each value series
   * is drawn (Excel's per-series Chart Type), in series order. Read only
   * while `chartSubtype` is "custom"; the Secondary Axis check boxes are
   * `secondaryAxisSeries`. An imported chart states the same per series in
   * `seriesRefs` (`renderType`, `axis`) instead.
   */
  comboSeriesTypes?: ComboRenderAs[];
  /**
   * Native classic-chart series custody. Imported charts carry their exact
   * worksheet references here so the Web renderer, scratch writer and
   * original-package patchback all consume the same bounded descriptor. New
   * range-bound charts omit it: the writer derives the same descriptors from
   * `range` + `seriesInColumns` without duplicating them in editor state.
   */
  seriesRefs?: SpreadsheetChartSeriesRef[];
  /** Read-only source appearance; omitted means unknown historical state. */
  chartVisibility?: ChartVisibility;
  /**
   * The ONE chart model every editor shares (`ChartBlockAttrsV2`): the kind,
   * the titles, the format and a cache of the plotted data. A sheet chart's
   * data lives in its cells; `chartV2.data` is the last reading of the bound
   * range and is refreshed on every draw and every write, never trusted by a
   * writer. The flat fields above are DERIVED from it (`chart-v2-model`) for
   * the writers that still read them; a record without it is read in memory
   * from the flat fields and gains it on its next write.
   */
  chartV2?: ChartBlockAttrsV2;
  /** Where the chart's data lives: the counterpart of the Docs workbook binding. */
  binding?: SpreadsheetChartBinding;
  /**
   * A PivotChart is still one ordinary floating chart, but its data authority
   * is a real PivotTable rather than a static range: the table of that name on
   * that sheet, wherever the chart itself sits (Move Chart takes it to any
   * sheet, as in Excel). See `resolvePivotChartPivot`.
   */
  pivotSource?: SpreadsheetPivotChartSource;
  /** A PivotChart's field buttons; absent shows them all (Excel's default). */
  pivotOptions?: SpreadsheetPivotChartOptions;
};

export const SPREADSHEET_LINE_MARKER_SYMBOLS = [
  "none",
  "circle",
  "dash",
  "diamond",
  "dot",
  "plus",
  "square",
  "star",
  "triangle",
  "x",
] as const;

export type SpreadsheetLineMarkerSymbol =
  (typeof SPREADSHEET_LINE_MARKER_SYMBOLS)[number];

const DRAWINGML_EMU_PER_POINT = 12_700;

export function normalizeSpreadsheetLineWidthPt(
  value: unknown,
): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined;
  const emu = Math.round(value * DRAWINGML_EMU_PER_POINT);
  if (emu < 1 || emu > 24 * DRAWINGML_EMU_PER_POINT) return undefined;
  const normalized = emu / DRAWINGML_EMU_PER_POINT;
  return Math.abs(normalized - value) <= 1e-9 ? normalized : undefined;
}

export function spreadsheetLineWidthPtToCssPx(
  value: unknown,
): number | undefined {
  const normalized = normalizeSpreadsheetLineWidthPt(value);
  return normalized === undefined ? undefined : ptToPx(normalized);
}

export interface SpreadsheetChartLineSeriesStyle {
  seriesIndex: number;
  markerSymbol?: SpreadsheetLineMarkerSymbol;
  markerSize?: number;
  smooth?: boolean;
  lineWidthPt?: number;
}

export function normalizeChartLineSeriesStyles(
  value: unknown,
): readonly SpreadsheetChartLineSeriesStyle[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_CHART_SERIES) {
    return undefined;
  }
  const allowed = new Set<string>(SPREADSHEET_LINE_MARKER_SYMBOLS);
  const seen = new Set<number>();
  const styles: SpreadsheetChartLineSeriesStyle[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
      return undefined;
    }
    const record = entry as Record<string, unknown>;
    const seriesIndex = record.seriesIndex;
    if (
      !Number.isInteger(seriesIndex) ||
      (seriesIndex as number) < 0 ||
      (seriesIndex as number) >= MAX_CHART_SERIES ||
      seen.has(seriesIndex as number)
    ) {
      return undefined;
    }
    const markerSymbol =
      typeof record.markerSymbol === "string" &&
      allowed.has(record.markerSymbol)
        ? (record.markerSymbol as SpreadsheetLineMarkerSymbol)
        : undefined;
    const markerSize =
      Number.isInteger(record.markerSize) &&
      (record.markerSize as number) >= 2 &&
      (record.markerSize as number) <= 72
        ? (record.markerSize as number)
        : undefined;
    const smooth =
      typeof record.smooth === "boolean" ? record.smooth : undefined;
    const lineWidthPt = normalizeSpreadsheetLineWidthPt(record.lineWidthPt);
    if (
      (record.markerSymbol !== undefined && markerSymbol === undefined) ||
      (record.markerSize !== undefined && markerSize === undefined) ||
      (record.smooth !== undefined && smooth === undefined) ||
      (record.lineWidthPt !== undefined && lineWidthPt === undefined) ||
      (markerSize !== undefined &&
        (markerSymbol === undefined || markerSymbol === "none")) ||
      (markerSymbol === undefined &&
        smooth === undefined &&
        lineWidthPt === undefined)
    ) {
      return undefined;
    }
    seen.add(seriesIndex as number);
    styles.push({
      seriesIndex: seriesIndex as number,
      ...(markerSymbol === undefined ? {} : { markerSymbol }),
      ...(markerSize === undefined ? {} : { markerSize }),
      ...(smooth === undefined ? {} : { smooth }),
      ...(lineWidthPt === undefined ? {} : { lineWidthPt }),
    });
  }
  return styles.sort((left, right) => left.seriesIndex - right.seriesIndex);
}

export interface SpreadsheetChartSeriesGradientStyle {
  seriesIndex: number;
  from: `#${string}`;
  to: `#${string}`;
}

export function normalizeChartSeriesGradientStyles(
  value: unknown,
): readonly SpreadsheetChartSeriesGradientStyle[] | undefined {
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_CHART_SERIES) {
    return undefined;
  }
  const seen = new Set<number>();
  const styles: SpreadsheetChartSeriesGradientStyle[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
      return undefined;
    }
    const record = entry as Record<string, unknown>;
    const seriesIndex = record.seriesIndex;
    const from =
      typeof record.from === "string" && /^#[0-9a-fA-F]{6}$/.test(record.from)
        ? (`#${record.from.slice(1).toLowerCase()}` as `#${string}`)
        : undefined;
    const to =
      typeof record.to === "string" && /^#[0-9a-fA-F]{6}$/.test(record.to)
        ? (`#${record.to.slice(1).toLowerCase()}` as `#${string}`)
        : undefined;
    if (
      !Number.isInteger(seriesIndex) ||
      (seriesIndex as number) < 0 ||
      (seriesIndex as number) >= MAX_CHART_SERIES ||
      seen.has(seriesIndex as number) ||
      from === undefined ||
      to === undefined ||
      from === to
    ) {
      return undefined;
    }
    seen.add(seriesIndex as number);
    styles.push({ seriesIndex: seriesIndex as number, from, to });
  }
  return styles.sort((left, right) => left.seriesIndex - right.seriesIndex);
}

export interface SpreadsheetChartValueAxisNumberFormat {
  formatCode: string;
  sourceLinked: boolean;
}

export type SpreadsheetChartLinkedFormatResolution =
  | { state: "resolved"; formatCode: string }
  | {
      state: "indeterminate";
      reason: "missing" | "mixed" | "invalid" | "reference" | "unsupported";
      formatCode: "General";
    };

export function normalizeChartValueAxisNumberFormat(
  raw: unknown,
): SpreadsheetChartValueAxisNumberFormat | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  if (typeof record.sourceLinked !== "boolean") return undefined;
  const normalized = normalizeNumberFormatCode(record.formatCode);
  if (normalized === null) return undefined;
  return {
    formatCode: normalized || "General",
    sourceLinked: record.sourceLinked,
  };
}

export interface SpreadsheetPivotChartSource {
  /** The sheet holding the PivotTable. */
  sheetId: string;
  /** The PivotTable's name; absent names that sheet's first PivotTable. */
  pivotName?: string;
}

/** A bounded PivotChart binding, or undefined for anything else. */
export function normalizeSpreadsheetPivotChartSource(
  raw: unknown,
): SpreadsheetPivotChartSource | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const sheetId = typeof record.sheetId === "string" ? record.sheetId.trim() : "";
  if (!sheetId || sheetId.length > MAX_CHART_ID_LENGTH) return undefined;
  const pivotName =
    typeof record.pivotName === "string" ? record.pivotName.trim() : "";
  return pivotName && pivotName.length <= 255
    ? { sheetId, pivotName }
    : { sheetId };
}

/** A single native classic-chart series, using worksheet A1 references. */
export interface SpreadsheetChartSeriesRef {
  title?: string;
  /** Literal native series title (``c:tx/c:v``), not a worksheet reference. */
  titleText?: string;
  categories?: string;
  xValues?: string;
  /** Excel ordinal X=1..N, authenticated from an empty ``c:xVal/c:numLit``. */
  implicitX?: true;
  values: string;
  bubbleSizes?: string;
  renderType?: "column" | "bar" | "line" | "area" | "scatter" | "bubble" | "pie" | "doughnut";
  axis?: "primary" | "secondary";
  /**
   * The references the chart part states as a defined name rather than
   * cells (`Sheet1!Sales`, `[0]!Sales`), by key: the key itself holds the
   * cells the name stands for, which the chart reads; the name is what a
   * save writes back (API `chart_series_names`).
   */
  names?: SpreadsheetChartSeriesRefNames;
}

/** The reference keys of a series that can be read through a defined name. */
export type SpreadsheetChartSeriesRefNames = Partial<
  Record<"title" | "categories" | "values" | "xValues" | "bubbleSizes", string>
>;

/** Series references one chart states: the one limits table's (Office's 255). */
export const MAX_CHART_SERIES_REFS = MAX_CHART_SERIES;

/**
 * Where a chart's data lives, the counterpart of the Docs
 * `documentChart.workbook.binding`: a rectangle of one sheet, plotted by
 * columns (Excel's default for a tall range) or by rows, and -- for a chart
 * another application wrote -- the per-series worksheet references it states.
 * A binding names at least one of the two.
 */
export interface SpreadsheetChartBinding {
  /** The sheet the rectangle is on; absent means the chart's own sheet. */
  sheetId?: string;
  range?: ChartSourceRange;
  /** Absent means "columns". */
  plotBy?: "rows" | "columns";
  seriesRefs?: SpreadsheetChartSeriesRef[];
  /**
   * The chart's data is its own: the values its part caches, not cells the
   * sheet can bind (an XY chart plotted against text X values, a bubble chart
   * of literal points). The model's `data` is then the data, never a cache
   * to refresh; its format and titles edit the part in place.
   */
  literal?: true;
}

export function normalizeSpreadsheetChartBinding(
  raw: unknown,
): SpreadsheetChartBinding | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const out: SpreadsheetChartBinding = {};
  if (typeof record.sheetId === "string" && record.sheetId.trim()) {
    out.sheetId = record.sheetId.trim().slice(0, MAX_CHART_ID_LENGTH);
  }
  const range = record.range === undefined ? null : normalizeChartSourceRange(record.range);
  if (range) out.range = range;
  if (record.plotBy === "rows") out.plotBy = "rows";
  const seriesRefs = normalizeChartSeriesRefs(record.seriesRefs);
  if (seriesRefs) out.seriesRefs = seriesRefs;
  if (!out.range && !out.seriesRefs && record.literal === true) return { literal: true };
  return out.range || out.seriesRefs ? out : undefined;
}

/**
 * A stored `chartV2`, or undefined when it is not a valid V2 record. A gap in
 * a series is `null` in the page (JSON has no NaN) and NaN in memory, as the
 * shared validator reads both.
 */
export function normalizeSpreadsheetChartV2(raw: unknown): ChartBlockAttrsV2 | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  if ((raw as { schemaVersion?: unknown }).schemaVersion !== 2) return undefined;
  const wire = normalizeChartBlockAttrsV2Wire(raw);
  if (!validateChartBlockAttrsV2(wire).valid) return undefined;
  const attrs = structuredClone(wire) as unknown as ChartBlockAttrsV2;
  const data = attrs.data as unknown as Record<string, unknown>;
  if (Array.isArray(data.series)) {
    for (const series of data.series as Record<string, unknown>[]) {
      for (const key of ["values", "xValues", "yValues", "bubbleSizes"]) {
        if (Array.isArray(series[key])) {
          series[key] = (series[key] as unknown[]).map((value) => (value === null ? Number.NaN : value));
        }
      }
    }
  }
  if (Array.isArray(data.values)) {
    data.values = (data.values as unknown[]).map((value) => (value === null ? Number.NaN : value));
  }
  return attrs;
}

/**
 * A chart record in its stored form: the model's gaps (a sheet reading's
 * NaN) are written as null, the JSON form `normalizeSpreadsheetChartV2` reads
 * back as a gap. A record without a gap keeps its identity.
 */
export function spreadsheetChartWithStoredGaps<T extends object>(chart: T): T {
  const attrs = (chart as { chartV2?: unknown }).chartV2 as
    | { data?: Record<string, unknown> }
    | undefined;
  const data = attrs?.data;
  if (!data || typeof data !== "object") return chart;
  const gapFree = (values: unknown): unknown => {
    if (!Array.isArray(values) || !values.some((value) => typeof value === "number" && Number.isNaN(value))) {
      return values;
    }
    return values.map((value) => (typeof value === "number" && Number.isNaN(value) ? null : value));
  };
  let changed = false;
  const series = Array.isArray(data.series)
    ? (data.series as Record<string, unknown>[]).map((entry) => {
        let next = entry;
        for (const key of ["values", "xValues", "yValues", "bubbleSizes"]) {
          const stored = gapFree(entry?.[key]);
          if (stored !== entry?.[key]) {
            next = next === entry ? { ...entry } : next;
            next[key] = stored;
          }
        }
        if (next !== entry) changed = true;
        return next;
      })
    : data.series;
  const values = gapFree(data.values);
  if (values !== data.values) changed = true;
  if (!changed) return chart;
  return {
    ...chart,
    chartV2: {
      ...attrs,
      data: {
        ...data,
        ...(series !== undefined ? { series } : {}),
        ...(values !== undefined ? { values } : {}),
      },
    },
  };
}

export interface MiniChartXYPoint {
  x: number;
  y: number;
  size?: number;
}

/** Authenticated native XY series ready for MiniChart's per-series renderer. */
export interface MiniChartXYSeries {
  name: string;
  points: MiniChartXYPoint[];
}

/** One [label, value] point ready for MiniChart's row consumption. */
export interface ChartSeriesRow {
  label: string;
  value: string;
}

/** A "row:col" -> resolved display string lookup over the live grid cells. */
export type ChartCellLookup = (row: number, col: number) => string | undefined;

/**
 * Coerce a free-form type label to a valid SpreadsheetChartType; a genuinely
 * unknown / legacy value falls back to "bar" (MiniChart's own fallback).
 *
 * Back-compat is load-bearing here: a saved chart whose `type` already equals a
 * canonical kind (including the new extended kinds) passes through UNCHANGED, so
 * an existing bar/line/pie/area chart keeps rendering exactly as before and a
 * newly saved scatter/stacked/combo/doughnut/radar no longer silently collapses
 * to bar. Only non-canonical labels go through the alias rules below.
 *
 * Alias rules (checked most-specific first so e.g. "stacked-column" beats the
 * bare "column" -> bar rule):
 *  - "bubble" -> bubble (a 3-column scatter; checked before "scatter" so a
 *    "bubble scatter" label still resolves to bubble)
 *  - "scatter" / "xy" -> scatter
 *  - "stack" + "column" (or "stacked-column") -> stacked-column;
 *     any other "stack" -> stacked-bar
 *  - "combo" / "bar+line" / "bar-line" -> combo
 *  - "radar" / "spider" -> radar
 *  - "doughnut" / "donut" -> doughnut (a hollow pie; was previously folded into
 *    pie, now its own kind)
 *  - "histogram" / "histo" -> histogram
 *  - "waterfall" / "bridge" -> waterfall
 *  - "box" + "whisker" (or "box-and-whisker" / "boxplot" / "box plot") ->
 *     box-and-whisker
 *  - "pareto" -> pareto
 *  - "treemap" / "tree map" / "tree-map" -> treemap
 *  - "sunburst" / "sun burst" / "sun-burst" -> sunburst
 *  - "stock" / "candlestick" / "ohlc" / "hlc" -> stock
 *  - "pie" -> pie
 *  - "area" -> area, "line" -> line
 *  - "column" -> column (vertical columns)
 *  - "bar" -> bar (horizontal bars)
 *  - else -> column
 */
export function normalizeSpreadsheetChartType(
  type: unknown,
): SpreadsheetChartType {
  const t = typeof type === "string" ? type.trim().toLowerCase() : "";
  // Fast path: an already-canonical value (old or new kind) passes through so a
  // persisted chart round-trips byte-for-byte.
  if (VALID_CHART_TYPES.has(t)) return t as SpreadsheetChartType;
  // The Camera tool rides the chart object but is not in VALID_CHART_TYPES (so it
  // stays out of the chart-type picker); pass it through explicitly so a saved
  // camera round-trips instead of collapsing to "bar".
  if (t === "camera") return "camera";
  // Bubble before scatter: a bubble is a 3-column scatter, so a label that names
  // both ("bubble scatter") should resolve to bubble.
  if (t.includes("bubble")) return "bubble";
  if (t.includes("scatter") || t === "xy") return "scatter";
  if (t.includes("stack")) {
    return t.includes("column") ? "stacked-column" : "stacked-bar";
  }
  if (t.includes("combo") || t.includes("bar+line") || t.includes("bar-line")) {
    return "combo";
  }
  if (t.includes("radar") || t.includes("spider")) return "radar";
  if (t.includes("doughnut") || t.includes("donut")) return "doughnut";
  if (t.includes("histogram") || t.includes("histo")) return "histogram";
  if (t.includes("waterfall") || t.includes("bridge")) return "waterfall";
  // Box-and-whisker: "boxplot" / "box plot" / "box" + "whisker" all resolve here.
  if (
    t.includes("boxplot") ||
    t.includes("box plot") ||
    (t.includes("box") && t.includes("whisker"))
  ) {
    return "box-and-whisker";
  }
  if (t.includes("pareto")) return "pareto";
  // Treemap: "treemap" / "tree map" / "tree-map".
  if (t.includes("treemap") || t.includes("tree map") || t.includes("tree-map")) {
    return "treemap";
  }
  // Sunburst: "sunburst" / "sun burst" / "sun-burst".
  if (
    t.includes("sunburst") ||
    t.includes("sun burst") ||
    t.includes("sun-burst")
  ) {
    return "sunburst";
  }
  // Stock / candlestick / OHLC / HLC.
  if (
    t.includes("stock") ||
    t.includes("candlestick") ||
    t === "ohlc" ||
    t === "hlc"
  ) {
    return "stock";
  }
  if (t.includes("pie")) return "pie";
  if (t.includes("area")) return "area";
  if (t.includes("line")) return "line";
  if (t.includes("column")) return "column";
  if (t.includes("bar")) return "bar";
  // The historic renderer's default geometry was a vertical column. Keep that
  // visual fallback without mislabelling it as Excel's horizontal Bar family.
  return "column";
}

/** Clamp an integer index into [0, max]; non-finite -> 0. */
function clampIndex(value: unknown, max: number): number {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return 0;
  const t = Math.trunc(n);
  if (t < 0) return 0;
  if (t > max) return max;
  return t;
}

/** Clamp a persisted pixel size into [MIN, MAX], falling back to `fallback`. */
function clampSize(raw: unknown, fallback: number): number {
  const n = typeof raw === "number" ? raw : Number(raw);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(MIN_CHART_SIZE, Math.min(MAX_CHART_SIZE, Math.round(n)));
}

/**
 * Clamp a persisted panel offset to a finite, non-negative pixel value, or
 * undefined when not a usable number (mirrors slicer's clampOffset). A chart
 * that was never dragged carries no x/y so the no-position path stays
 * byte-identical with the index-stack default.
 */
function clampChartOffset(raw: unknown, max: number): number | undefined {
  if (typeof raw !== "number" || !Number.isFinite(raw)) return undefined;
  // The grid's own reach, not a fixed 100,000 px (about row 5,000).
  return Math.max(0, Math.min(max, Math.round(raw)));
}

/**
 * Coerce a raw persisted/edited trendline value into a safe ChartTrendlineSpec,
 * or undefined when absent or not a real trendline. A `"none"` type (the UI's
 * "no trendline" sentinel) and any unrecognized type drop to undefined so a
 * chart without a trendline carries NO `trendline` key (byte-stable). The
 * moving-average `period` is clamped to [DEFAULT_TRENDLINE_PERIOD,
 * MAX_TRENDLINE_PERIOD] and stored ONLY for the movingAverage kind (linear never
 * persists a period, so it stays byte-stable too). Accepts the bare string
 * `"linear"`/`"movingAverage"` as a convenience for the editor.
 */
export function normalizeTrendline(
  raw: unknown,
): ChartTrendlineSpec | undefined {
  const typeRaw =
    typeof raw === "string"
      ? raw
      : raw && typeof raw === "object" && !Array.isArray(raw)
        ? (raw as Record<string, unknown>).type
        : undefined;
  const type = typeof typeRaw === "string" ? typeRaw.trim() : "";
  // Office's Format Trendline options ride along when the shared contract
  // accepts them for this kind; one that does not belong is dropped, the
  // trendline stays.
  const withOptions = (base: ChartTrendlineSpec): ChartTrendlineSpec => {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) return base;
    const record = raw as Record<string, unknown>;
    const options: Record<string, unknown> = {};
    for (const key of [
      "name",
      "forward",
      "backward",
      "intercept",
      "displayEquation",
      "displayRSquared",
    ])
      if (record[key] !== undefined && record[key] !== null)
        options[key] = record[key];
    if (!Object.keys(options).length) return base;
    const accepted = spreadsheetTrendlineContract({
      ...base,
      ...(options as OfficeTrendlineOptions),
    });
    if (!accepted) return base;
    const { type: _type, ...rest } = accepted as unknown as Record<string, unknown>;
    delete rest.order;
    delete rest.period;
    return { ...base, ...(rest as OfficeTrendlineOptions) };
  };
  if (
    type === "linear" ||
    type === "exponential" ||
    type === "logarithmic" ||
    type === "power"
  )
    return withOptions({ type });
  if (type === "polynomial") {
    const orderRaw =
      raw && typeof raw === "object" && !Array.isArray(raw)
        ? (raw as Record<string, unknown>).order
        : undefined;
    const n = typeof orderRaw === "number" ? orderRaw : Number(orderRaw);
    return withOptions({
      type: "polynomial",
      order: Number.isFinite(n)
        ? Math.max(
            OFFICE_TRENDLINE_ORDER.min,
            Math.min(OFFICE_TRENDLINE_ORDER.max, Math.trunc(n)),
          )
        : OFFICE_TRENDLINE_ORDER.min,
    });
  }
  if (type === "movingAverage") {
    const out: ChartTrendlineSpec = { type: "movingAverage" };
    const periodRaw =
      raw && typeof raw === "object" && !Array.isArray(raw)
        ? (raw as Record<string, unknown>).period
        : undefined;
    const n = typeof periodRaw === "number" ? periodRaw : Number(periodRaw);
    if (Number.isFinite(n)) {
      out.period = Math.max(
        DEFAULT_TRENDLINE_PERIOD,
        Math.min(MAX_TRENDLINE_PERIOD, Math.trunc(n)),
      );
    }
    return withOptions(out);
  }
  // "none" and anything unrecognized -> no trendline (drop the key entirely).
  return undefined;
}

/**
 * Coerce a raw persisted/edited error-bar value into a safe ChartErrorBarSpec,
 * or undefined when absent or not a real spec. A `"none"` type (the UI's "no
 * error bars" sentinel) and any unrecognized type drop to undefined so a chart
 * without error bars carries NO `errorBars` key (byte-stable). The amount is
 * coerced to a finite number, clamped to [0, MAX_ERROR_BAR_VALUE] (a stale
 * negative can't invert the whisker), and defaults to 0 when missing/non-finite.
 * Accepts the bare string `"fixed"`/`"percent"`/`"stddev"` (defaulting value to
 * 0) as a convenience for the editor.
 */
export function normalizeErrorBars(
  raw: unknown,
): ChartErrorBarSpec | undefined {
  const isObject =
    raw != null && typeof raw === "object" && !Array.isArray(raw);
  const typeRaw =
    typeof raw === "string"
      ? raw
      : isObject
        ? (raw as Record<string, unknown>).type
        : undefined;
  const type = typeof typeRaw === "string" ? typeRaw.trim() : "";
  if (!VALID_ERROR_BAR_TYPES.has(type)) return undefined;
  const valueRaw = isObject
    ? (raw as Record<string, unknown>).value
    : undefined;
  const n = typeof valueRaw === "number" ? valueRaw : Number(valueRaw);
  const value = Number.isFinite(n)
    ? Math.max(0, Math.min(MAX_ERROR_BAR_VALUE, n))
    : 0;
  const record = isObject ? (raw as Record<string, unknown>) : {};
  return {
    type: type as ChartErrorBarType,
    value,
    ...(record.include === "plus" || record.include === "minus" ? { include: record.include } : {}),
    ...(record.noEndCap === true ? { noEndCap: true as const } : {}),
  };
}

/**
 * Coerce a raw persisted/edited secondary-axis series list into a sorted, deduped
 * array of in-range 0-based value-series indices, or undefined when absent /
 * empty / unusable. Non-integer, negative, and non-finite entries are dropped;
 * duplicates collapse; the list is sorted ascending and capped at
 * MAX_SECONDARY_AXIS_SERIES (DoS bound). Returns undefined (NOT an empty array)
 * when nothing valid remains so a chart without a secondary axis carries NO
 * `secondaryAxisSeries` key (byte-stable). Accepts a bare number as a one-element
 * convenience for the editor.
 */
export function normalizeSecondaryAxisSeries(
  raw: unknown,
): number[] | undefined {
  const list = Array.isArray(raw)
    ? raw
    : typeof raw === "number"
      ? [raw]
      : undefined;
  if (!list) return undefined;
  const seen = new Set<number>();
  for (const entry of list) {
    const n = typeof entry === "number" ? entry : Number(entry);
    if (!Number.isFinite(n)) continue;
    const idx = Math.trunc(n);
    if (idx < 0 || idx > MAX_SERIES_INDEX) continue;
    seen.add(idx);
    if (seen.size >= MAX_SECONDARY_AXIS_SERIES) break;
  }
  if (seen.size === 0) return undefined;
  return Array.from(seen).sort((a, b) => a - b);
}

/** A Custom Combination's per-series chart types, or undefined (byte-stable). */
export function normalizeComboSeriesTypes(
  raw: unknown,
): ComboRenderAs[] | undefined {
  if (!Array.isArray(raw) || raw.length < 2 || raw.length > 255) return undefined;
  return raw.every((entry) => isComboRenderAs(entry))
    ? (raw as ComboRenderAs[]).slice()
    : undefined;
}

const CHART_SERIES_RENDER_TYPES = new Set([
  "column",
  "bar",
  "line",
  "area",
  "scatter",
  "bubble",
  // A combination's ring series (a pie and a doughnut drawn over each other).
  "pie",
  "doughnut",
]);

function boundedChartRef(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const ref = value.trim();
  return ref && ref.length <= 500 ? ref : undefined;
}

/** Strict, bounded receipt for the shared native classic-chart series wire. */
export function normalizeChartSeriesRefs(
  raw: unknown,
): SpreadsheetChartSeriesRef[] | undefined {
  if (!Array.isArray(raw) || raw.length === 0 || raw.length > MAX_CHART_SERIES_REFS) {
    return undefined;
  }
  const refs: SpreadsheetChartSeriesRef[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object" || Array.isArray(item)) return undefined;
    const source = item as Record<string, unknown>;
    const values = boundedChartRef(source.values);
    if (!values) return undefined;
    const entry: SpreadsheetChartSeriesRef = { values };
    for (const key of [
      "title",
      "titleText",
      "categories",
      "xValues",
      "bubbleSizes",
    ] as const) {
      const value = boundedChartRef(source[key]);
      if (source[key] !== undefined && !value) return undefined;
      if (value) entry[key] = value;
    }
    if (source.implicitX !== undefined) {
      if (source.implicitX !== true) return undefined;
      entry.implicitX = true;
    }
    if (source.renderType !== undefined) {
      if (
        typeof source.renderType !== "string" ||
        !CHART_SERIES_RENDER_TYPES.has(source.renderType)
      ) {
        return undefined;
      }
      entry.renderType = source.renderType as SpreadsheetChartSeriesRef["renderType"];
    }
    if (source.axis !== undefined) {
      if (source.axis !== "primary" && source.axis !== "secondary") return undefined;
      entry.axis = source.axis;
    }
    // A name is kept only beside the cells it stands for; one this cannot
    // read is let go (the key's cells stay what the chart reads).
    if (source.names && typeof source.names === "object" && !Array.isArray(source.names)) {
      const names: SpreadsheetChartSeriesRefNames = {};
      for (const key of ["title", "categories", "values", "xValues", "bubbleSizes"] as const) {
        const name = boundedChartRef((source.names as Record<string, unknown>)[key]);
        if (name && entry[key]) names[key] = name;
      }
      if (Object.keys(names).length) entry.names = names;
    }
    const effectiveType = entry.renderType;
    if (entry.xValues && entry.implicitX) return undefined;
    if (
      entry.implicitX &&
      (entry.categories !== undefined ||
        entry.bubbleSizes !== undefined ||
        effectiveType === "bubble")
    ) {
      return undefined;
    }
    if (
      (effectiveType === "scatter" || effectiveType === "bubble") &&
      !entry.xValues &&
      !entry.implicitX
    ) {
      return undefined;
    }
    if (effectiveType === "bubble" && !entry.bubbleSizes) return undefined;
    refs.push(entry);
  }
  return refs;
}

/**
 * Pure resize-commit clamp for a drag-RESIZE: round + clamp width/height into
 * the interactive resize bounds [MIN_..RESIZE, MAX_CHART_RESIZE]. Used by the
 * resize-handle drop so a dragged size can never collapse the box below a
 * legible chart or balloon past the model max. Separate from the persist clamp
 * (`clampSize`, floor 80) so import round-trips are unaffected; this is the
 * value the UI WRITES when the user releases the handle.
 */
export function clampChartSize(
  width: number,
  height: number,
): { width: number; height: number } {
  const w = Number.isFinite(width) ? Math.round(width) : MIN_CHART_RESIZE_WIDTH;
  const h = Number.isFinite(height)
    ? Math.round(height)
    : MIN_CHART_RESIZE_HEIGHT;
  return {
    width: Math.max(MIN_CHART_RESIZE_WIDTH, Math.min(MAX_CHART_RESIZE, w)),
    height: Math.max(MIN_CHART_RESIZE_HEIGHT, Math.min(MAX_CHART_RESIZE, h)),
  };
}

/**
 * Normalize a raw source range (a {r1,c1,r2,c2} object) into an in-grid,
 * sorted ChartSourceRange (r1<=r2, c1<=c2). Returns null when the input is not
 * a recognizable rectangle object.
 */
export function normalizeChartSourceRange(input: unknown): ChartSourceRange | null {
  if (!input || typeof input !== "object" || Array.isArray(input)) return null;
  const o = input as Record<string, unknown>;
  if (
    o.r1 === undefined ||
    o.c1 === undefined ||
    o.r2 === undefined ||
    o.c2 === undefined
  ) {
    return null;
  }
  const r1 = clampIndex(o.r1, MAX_RANGE_ROW);
  const r2 = clampIndex(o.r2, MAX_RANGE_ROW);
  const c1 = clampIndex(o.c1, MAX_RANGE_COL);
  const c2 = clampIndex(o.c2, MAX_RANGE_COL);
  return {
    r1: Math.min(r1, r2),
    c1: Math.min(c1, c2),
    r2: Math.max(r1, r2),
    c2: Math.max(c1, c2),
  };
}

/**
 * Coerce a raw persisted entry into a safe SpreadsheetChart, or null when
 * malformed (missing/over-long id, no usable range). The anchor defaults to the
 * range's top-left when absent/invalid; size defaults to the canvas default;
 * title is dropped unless a non-empty string. Extra fields are ignored.
 */
/** A chart's cell anchor, the worksheet drawing anchor's own spelling. */
export interface SpreadsheetChartDrawingAnchorPoint {
  row: number;
  column: number;
  rowOffset?: number;
  columnOffset?: number;
}

export interface SpreadsheetChartDrawingAnchor {
  kind: "twoCell" | "oneCell" | "absolute";
  from?: SpreadsheetChartDrawingAnchorPoint;
  to?: SpreadsheetChartDrawingAnchorPoint;
  xEmu?: number;
  yEmu?: number;
  widthEmu?: number;
  heightEmu?: number;
  origin?: "sheet";
  /** Placement chosen in Properties: written as Excel writes it, a
   * `twoCellAnchor editAs` of this kind (see SpreadsheetObjectAnchor). */
  editAs?: "oneCell" | "absolute";
}

const MAX_CHART_ANCHOR_EMU = 4_000_000_000;

function chartAnchorPoint(raw: unknown): SpreadsheetChartDrawingAnchorPoint | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const p = raw as Record<string, unknown>;
  const index = (value: unknown, max: number) =>
    typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= max ? value : undefined;
  const row = index(p.row, MAX_GRID_ROWS);
  const column = index(p.column ?? p.col, MAX_GRID_COLUMNS);
  if (row === undefined || column === undefined) return undefined;
  const rowOffset = index(p.rowOffset, MAX_CHART_ANCHOR_EMU);
  const columnOffset = index(p.columnOffset ?? p.colOffset, MAX_CHART_ANCHOR_EMU);
  return {
    row,
    column,
    ...(rowOffset !== undefined ? { rowOffset } : {}),
    ...(columnOffset !== undefined ? { columnOffset } : {}),
  };
}

/** A persisted chart cell anchor, or undefined when it is not a complete one. */
export function normalizeChartDrawingAnchor(raw: unknown): SpreadsheetChartDrawingAnchor | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const a = raw as Record<string, unknown>;
  const emu = (value: unknown, min: number) =>
    typeof value === "number" && Number.isFinite(value) && value >= min && value <= MAX_CHART_ANCHOR_EMU
      ? Math.round(value)
      : undefined;
  if (a.kind === "twoCell") {
    const from = chartAnchorPoint(a.from);
    const to = chartAnchorPoint(a.to);
    return from && to ? { kind: "twoCell", from, to } : undefined;
  }
  const widthEmu = emu(a.widthEmu, 1);
  const heightEmu = emu(a.heightEmu, 1);
  if (widthEmu === undefined || heightEmu === undefined) return undefined;
  if (a.kind === "oneCell") {
    const from = chartAnchorPoint(a.from);
    return from ? { kind: "oneCell", from, widthEmu, heightEmu, ...(a.editAs === "oneCell" ? { editAs: "oneCell" as const } : {}) } : undefined;
  }
  if (a.kind === "absolute") {
    const xEmu = emu(a.xEmu, 0);
    const yEmu = emu(a.yEmu, 0);
    return xEmu !== undefined && yEmu !== undefined
      ? { kind: "absolute", xEmu, yEmu, widthEmu, heightEmu, ...(a.origin === "sheet" ? { origin: "sheet" as const } : {}), ...(a.editAs === "absolute" ? { editAs: "absolute" as const } : {}) }
      : undefined;
  }
  return undefined;
}

export function normalizeChart(raw: unknown): SpreadsheetChart | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const id = typeof r.id === "string" ? r.id.trim() : "";
  if (id === "" || id.length > MAX_CHART_ID_LENGTH) return null;
  const range = normalizeChartSourceRange(r.range);
  if (!range) return null;
  const out: SpreadsheetChart = {
    id,
    type: normalizeSpreadsheetChartType(r.type),
    range,
    anchor: { row: range.r1, col: range.c1 },
    width: clampSize(r.width, DEFAULT_CHART_WIDTH),
    height: clampSize(r.height, DEFAULT_CHART_HEIGHT),
  };
  // Anchor: a valid {row,col} wins; otherwise default to the range top-left.
  if (r.anchor && typeof r.anchor === "object" && !Array.isArray(r.anchor)) {
    const a = r.anchor as Record<string, unknown>;
    out.anchor = {
      // Any cell of the grid (it was clamped at row 100,000).
      row: clampIndex(a.row, MAX_RANGE_ROW),
      col: clampIndex(a.col, MAX_RANGE_COL),
    };
  }
  if (typeof r.title === "string") {
    const title = r.title.trim();
    if (title !== "") out.title = title.slice(0, MAX_CHART_TITLE_LENGTH);
  }
  if (typeof r.name === "string") {
    const name = r.name.trim();
    if (name !== "") out.name = name.slice(0, MAX_CHART_TITLE_LENGTH);
  }
  if (r.hidden === true) out.hidden = true;
  if (
    typeof r.outerCategoryColumns === "number" &&
    Number.isInteger(r.outerCategoryColumns) &&
    r.outerCategoryColumns >= 1 &&
    r.outerCategoryColumns < MAX_CHART_CATEGORY_LEVELS &&
    r.outerCategoryColumns <= range.c1
  ) {
    out.outerCategoryColumns = r.outerCategoryColumns;
  }
  if (r.noCategories === true) out.noCategories = true;
  if (r.noSeriesNames === true) out.noSeriesNames = true;
  // Outer category levels sit beside a category column under a row of names.
  if (out.noCategories || out.noSeriesNames) delete out.outerCategoryColumns;
  
  if (
    typeof r.nameRowsAbove === "number" &&
    Number.isInteger(r.nameRowsAbove) &&
    r.nameRowsAbove >= 1 &&
    r.nameRowsAbove <= Math.min(range.r1, MAX_CHART_NAME_ROWS_ABOVE) &&
    r.seriesInColumns !== false &&
    !out.noSeriesNames
  ) {
    out.nameRowsAbove = r.nameRowsAbove;
  }
  // XLS-A081: alt text and the assigned macro, as written.
  if (typeof r.alt === "string" && r.alt.trim()) out.alt = r.alt.slice(0, 2_000);
  if (typeof r.altTitle === "string" && r.altTitle.trim()) out.altTitle = r.altTitle.slice(0, 255);
  if (r.decorative === true) out.decorative = true;
  if (typeof r.macro === "string" && r.macro.trim() && r.macro.length <= 300) out.macro = r.macro.trim();
  // The sheet-wide Arrange position (drawing-stack); a stale or malformed
  // value falls back to the collection default rather than failing the chart.
  if (typeof r.stackOrder === "number" && Number.isInteger(r.stackOrder) && r.stackOrder >= 0) out.stackOrder = r.stackOrder;
  if (r.sourceRefError === true) out.sourceRefError = true;
  if (typeof r.groupId === "string" && r.groupId.trim() && r.groupId.length <= 200) out.groupId = r.groupId.trim();
  const chartStyle = normalizeSpreadsheetChartStyle(r.chartStyle);
  if (chartStyle !== undefined) out.chartStyle = chartStyle;
  if (r.chartFormat !== undefined) {
    const format = normalizeSpreadsheetChartFormat(r.chartFormat);
    if (format) out.chartFormat = format;
    else if (r.chartFormat && typeof r.chartFormat === "object" && !Array.isArray(r.chartFormat)) {
      out.chartFormat = r.chartFormat as SpreadsheetChartFormat;
      out.nativeChartFormatBlocker = "unknown_chart_format";
    }
  }
  for (const key of ["chartSubtype", "desktopStyleId", "sourceSheetId", "nativeChartFormatBlocker"] as const) {
    if (typeof r[key] === "string" && r[key].trim()) out[key] = r[key].trim().slice(0, 200);
  }
  const chartColorPalette = normalizeSpreadsheetChartColorPalette(
    r.chartColorPalette,
  );
  if (chartColorPalette !== undefined) {
    out.chartColorPalette = chartColorPalette;
  }
  const seriesStyles = normalizePresentationChartSeriesStyles(r.seriesStyles);
  if (seriesStyles) out.seriesStyles = seriesStyles;
  const seriesGradientStyles = normalizeChartSeriesGradientStyles(
    r.seriesGradientStyles,
  );
  if (
    seriesGradientStyles &&
    out.type === "column" &&
    !seriesStyles &&
    chartColorPalette === undefined
  ) {
    out.seriesGradientStyles = seriesGradientStyles;
  }
  if (
    typeof r.nativeChartColorEditBlocker === "string" &&
    r.nativeChartColorEditBlocker.trim()
  ) {
    out.nativeChartColorEditBlocker = r.nativeChartColorEditBlocker
      .trim()
      .slice(0, 200);
  }
  // A line series' marker, smoothing and width: a line chart's, or the line
  // series of a combination / secondary-axis chart the model states them on.
  const lineSeriesStyles = normalizeChartLineSeriesStyles(r.lineSeriesStyles);
  if (lineSeriesStyles && (out.type === "line" || out.type === "combo" || normalizeSecondaryAxisSeries(r.secondaryAxisSeries))) {
    out.lineSeriesStyles = lineSeriesStyles;
  }
  if (
    typeof r.nativeChartLineBlocker === "string" &&
    r.nativeChartLineBlocker.trim()
  ) {
    out.nativeChartLineBlocker = r.nativeChartLineBlocker.trim().slice(0, 200);
  }
  if (
    typeof r.nativeChartLineWidthBlocker === "string" &&
    r.nativeChartLineWidthBlocker.trim()
  ) {
    out.nativeChartLineWidthBlocker = r.nativeChartLineWidthBlocker
      .trim()
      .slice(0, 200);
  }
  const sliceStyles = normalizePresentationChartSliceStyles(r.sliceStyles);
  if (sliceStyles && (out.type === "pie" || out.type === "doughnut")) {
    out.sliceStyles = sliceStyles;
  }
  // Default series-in-columns ON (Excel's default for a tall range); only an
  // explicit false flips it. Persisted only when false so the common case adds
  // no key (byte-stable with the implicit default).
  if (r.seriesInColumns === false) out.seriesInColumns = false;
  // Dragged pixel position (mirror slicer): kept only when a finite number, so a
  // never-moved chart adds no x/y key (byte-stable with the index-stack default).
  const x = clampChartOffset(r.x, MAX_DRAWING_LAYER_X_PX);
  if (x !== undefined) out.x = x;
  const y = clampChartOffset(r.y, MAX_DRAWING_LAYER_Y_PX);
  if (y !== undefined) out.y = y;
  // Where the chart sits on the grid and how it follows the cells (Excel's
  // Format Chart Area > Properties); x/y/width/height are its frame as last
  // drawn (chart-anchor).
  const drawingAnchor = normalizeChartDrawingAnchor(r.drawingAnchor);
  if (drawingAnchor) out.drawingAnchor = drawingAnchor;
  // Trendline: persisted only when a real spec survives normalization; a "none"
  // / absent / unrecognized value drops the key so an existing chart stays
  // byte-identical.
  const trendline = normalizeTrendline(r.trendline);
  if (trendline) out.trendline = trendline;
  // Data labels: persisted only when explicitly true (default off), so the
  // common case adds no key (byte-stable with the implicit default).
  if (r.showDataLabels === true) out.showDataLabels = true;
  if (typeof r.showLegend === "boolean") out.showLegend = r.showLegend;
  if (
    r.legendPosition === "topRight" ||
    (SPREADSHEET_CHART_LEGEND_POSITIONS as readonly string[]).includes(
      r.legendPosition as string,
    )
  ) {
    out.legendPosition = r.legendPosition as SpreadsheetChartLegendPosition;
  }
  if (
    typeof r.nativeChartLegendPositionBlocker === "string" &&
    r.nativeChartLegendPositionBlocker.trim()
  ) {
    out.nativeChartLegendPositionBlocker =
      r.nativeChartLegendPositionBlocker.trim().slice(0, 200);
  }
  if (typeof r.showAxisLabels === "boolean") {
    out.showAxisLabels = r.showAxisLabels;
  }
  if (typeof r.showGridlines === "boolean") {
    out.showGridlines = r.showGridlines;
  }
  if (r.showDataTable === true) out.showDataTable = true;
  // Axis titles: persisted only when a non-empty string (same rule as `title`),
  // so a chart without them adds no key (byte-stable).
  if (typeof r.axisTitleX === "string") {
    const axisTitleX = r.axisTitleX.trim();
    if (axisTitleX !== "")
      out.axisTitleX = axisTitleX.slice(0, MAX_CHART_TITLE_LENGTH);
  }
  if (typeof r.axisTitleY === "string") {
    const axisTitleY = r.axisTitleY.trim();
    if (axisTitleY !== "")
      out.axisTitleY = axisTitleY.slice(0, MAX_CHART_TITLE_LENGTH);
  }
  const valueAxisNumberFormat = normalizeChartValueAxisNumberFormat(
    r.valueAxisNumberFormat,
  );
  if (valueAxisNumberFormat) {
    out.valueAxisNumberFormat = valueAxisNumberFormat;
  }
  if (
    typeof r.nativeValueAxisNumberFormatBlocker === "string" &&
    r.nativeValueAxisNumberFormatBlocker.trim()
  ) {
    out.nativeValueAxisNumberFormatBlocker =
      r.nativeValueAxisNumberFormatBlocker.trim().slice(0, 200);
  }
  // Error bars: persisted only when a valid spec survives normalization; a
  // "none" / absent / unrecognized value drops the key so an existing chart
  // stays byte-identical.
  const errorBars = normalizeErrorBars(r.errorBars);
  if (errorBars) out.errorBars = errorBars;

  if (r.upDownBars === true) out.upDownBars = true;
  if (r.dropLines === true) out.dropLines = true;
  if (r.highLowLines === true) out.highLowLines = true;
  if (r.seriesLines === true) out.seriesLines = true;
  // Secondary-axis series: persisted only when a non-empty validated index list
  // survives; an absent / empty / unusable value drops the key (byte-stable).
  const secondaryAxisSeries = normalizeSecondaryAxisSeries(r.secondaryAxisSeries);
  if (secondaryAxisSeries) out.secondaryAxisSeries = secondaryAxisSeries;
  const comboSeriesTypes = normalizeComboSeriesTypes(r.comboSeriesTypes);
  if (comboSeriesTypes) out.comboSeriesTypes = comboSeriesTypes;
  const chartVisibility = normalizeChartVisibility(r.chartVisibility);
  if (chartVisibility) out.chartVisibility = chartVisibility;
  const seriesRefs = normalizeChartSeriesRefs(r.seriesRefs);
  if (seriesRefs) out.seriesRefs = seriesRefs;
  // The shared model and its binding travel together: one without the other
  // is not a reading of this chart, and the flat fields still describe it.
  const chartV2 = normalizeSpreadsheetChartV2(r.chartV2);
  const binding = normalizeSpreadsheetChartBinding(r.binding);
  if (chartV2 && binding) {
    out.chartV2 = chartV2;
    out.binding = binding;
  }
  const pivotSource = normalizeSpreadsheetPivotChartSource(r.pivotSource);
  if (pivotSource) {
    out.pivotSource = pivotSource;
    const pivotOptions = normalizePivotChartOptions(r.pivotOptions);
    if (pivotOptions) out.pivotOptions = pivotOptions;
  }
  return out;
}

/**
 * Normalize a raw `charts` array off a persisted sheet: drop malformed entries
 * and a duplicate id (keep the first), and cap the entry count (DoS bound).
 * Returns a fresh array; empty when nothing valid remains so the no-chart path
 * stays byte-identical.
 */
export function normalizeCharts(raw: unknown): SpreadsheetChart[] {
  if (!Array.isArray(raw)) return [];
  const out: SpreadsheetChart[] = [];
  const usedIds = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_CHARTS) break;
    const chart = normalizeChart(entry);
    if (!chart) continue;
    if (usedIds.has(chart.id)) continue;
    usedIds.add(chart.id);
    out.push(chart);
  }
  return out;
}

/**
 * Build the row grid MiniChart consumes from a chart's source range and the live
 * grid cells. Orientation follows `seriesInColumns`:
 *
 *  - seriesInColumns (default): the FIRST COLUMN of the range supplies labels;
 *    EVERY later column supplies a value series, read top-to-bottom. When the
 *    range is one column wide that single column is both label and value.
 *  - series in ROWS (seriesInColumns === false): the FIRST ROW supplies labels;
 *    EVERY later row supplies a value series, read left-to-right. A one-row range
 *    uses that row for both label and value.
 *
 * The returned rows are [labelString, value1String, value2String, ...]: the
 * label plus one cell per value series. Single-series renderers (bar/line/pie/
 * area, the historical default) read only value column 1 (`row[1]`), so their
 * output is unchanged when the range is the usual two columns wide and is
 * driven purely by the FIRST value series for a wider range — byte-identical
 * with the previous "second column only" behavior. Multi-series renderers
 * (stacked / combo / scatter / radar) read the additional value columns.
 *
 * MiniChart skips a leading non-numeric header row and parses each value.
 * Returns [] when there is nothing chartable.
 */
/** A cell text that reads as a number (a thousands separator, a currency
 * sign or a percent sign allowed), as the chart gates read one. */
function numericText(value: string): boolean {
  const text = value.trim().replace(/[$,%\s]/g, "");
  return text !== "" && Number.isFinite(Number(text));
}

/**
 * Where a range-bound chart's names, categories and values sit in its range,
 * in the range's own terms: with series in columns a "line" is a column and a
 * "point" a row; with series in rows the other way round. `nameLine` is the
 * row (column) of series names and `categoryLine` the column (row) of
 * category labels, each null when the range has none (`noSeriesNames`,
 * `noCategories`, a one-column range); `points` the first and last row
 * (column) a series reads; `seriesLines` the column (row) of each series.
 * The one layout behind the sheet's reading (`buildChartSeries`), the
 * workbook writer's references (API `chart_recipe_binding`) and the source
 * outlines.
 */
export interface NativeChartRangeLayout {
  readonly byColumns: boolean;
  readonly nameLine: number | null;
  readonly categoryLine: number | null;
  readonly points: readonly [number, number];
  readonly seriesLines: readonly number[];
}



export const MAX_CHART_NAME_ROWS_ABOVE = 32;










export function chartNameOverRows(
  rows: readonly [number, number],
  col: number,
  cells: (row: number, col: number) => string | number | undefined,
): string | null {
  const [r1, r2] = rows;
  if (r2 <= r1) return null;
  const parts: string[] = [];
  for (let row = r1; row <= r2; row += 1) {
    let text = "";
    for (let at = col; at >= 0 && text.trim() === ""; at -= 1) text = String(cells(row, at) ?? "");
    if (text.trim() !== "") parts.push(text);
  }
  return parts.join(" ");
}

export function nativeChartRangeLayout(
  chart: Pick<SpreadsheetChart, "range" | "seriesInColumns" | "noCategories" | "noSeriesNames">,
): NativeChartRangeLayout {
  const { r1, c1, r2, c2 } = chart.range;
  const byColumns = chart.seriesInColumns !== false;
  // In the range's own terms: lines run across, points along.
  const [lineFirst, lineLast, pointFirst, pointLast] = byColumns ? [c1, c2, r1, r2] : [r1, r2, c1, c2];
  const categories = chart.noCategories !== true && lineLast > lineFirst;
  const names = chart.noSeriesNames !== true;
  const firstSeries = categories ? lineFirst + 1 : lineFirst;
  return {
    byColumns,
    nameLine: names ? pointFirst : null,
    categoryLine: categories ? lineFirst : null,
    points: [names ? pointFirst + 1 : pointFirst, pointLast],
    seriesLines: Array.from({ length: Math.max(0, lineLast - firstSeries + 1) }, (_, index) => firstSeries + index),
  };
}

/** Whether a range-bound chart states where its labels are (`noCategories`,
 * `noSeriesNames`) rather than leaving its first row and column to be read. */
export function nativeChartStatesLabelLines(chart: Pick<SpreadsheetChart, "noCategories" | "noSeriesNames">): boolean {
  return chart.noCategories === true || chart.noSeriesNames === true;
}

/**
 * The series references a range-bound chart plots, as the workbook writer
 * states them (API `chart_recipe_binding.editor_chart_as_xlsx_object`): per
 * series its name cell, its category labels and its values, each left out
 * when the range has none. `prefix` is the sheet as a reference spells it
 * (none: references on the chart's own sheet). Null for a kind that does not
 * plot categories (XY, a camera) or a range that plots nothing.
 */
export function nativeChartRangeSeriesRefs(
  chart: Pick<SpreadsheetChart, "type" | "range" | "seriesInColumns" | "noCategories" | "noSeriesNames"> &
    Partial<Pick<SpreadsheetChart, "outerCategoryColumns" | "nameRowsAbove">>,
  prefix?: string,
): SpreadsheetChartSeriesRef[] | null {
  if (chart.type === "scatter" || chart.type === "bubble" || chart.type === "camera") return null;
  const layout = nativeChartRangeLayout(chart);
  const [first, last] = layout.points;
  if (last < first || layout.seriesLines.length === 0) return null;
  const sheet = prefix ? `${prefix}!` : "";
  // A cell of the range in its own terms (line, point).
  const at = (line: number, point: number) =>
    layout.byColumns ? `$${columnToLetters(line)}$${point + 1}` : `$${columnToLetters(point)}$${line + 1}`;
  const span = (line: number) => `${sheet}${at(line, first)}${last > first ? `:${at(line, last)}` : ""}`;
  // A tree map's or a sunburst's category area spans its outer text columns,
  
  // N2:O4), as the writer states it.
  const outer = chart.outerCategoryColumns ?? 0;
  const branches =
    (chart.type === "treemap" || chart.type === "sunburst") &&
    layout.byColumns &&
    layout.categoryLine !== null &&
    outer >= 1 &&
    layout.categoryLine - outer >= 0
      ? `${sheet}${at(layout.categoryLine - outer, first)}:${at(layout.categoryLine, last)}`
      : null;
  
  const above = layout.byColumns && layout.nameLine !== null ? Math.min(chart.nameRowsAbove ?? 0, layout.nameLine) : 0;
  return layout.seriesLines.map((line) => ({
    ...(layout.nameLine !== null
      ? {
          title: `${sheet}${
            layout.byColumns
              ? `${above > 0 ? `$${columnToLetters(line)}$${layout.nameLine - above + 1}:` : ""}$${columnToLetters(line)}$${layout.nameLine + 1}`
              : `$${columnToLetters(layout.nameLine)}$${line + 1}`
          }`,
        }
      : {}),
    ...(layout.categoryLine !== null ? { categories: branches ?? span(layout.categoryLine) } : {}),
    values: span(line),
  }));
}

export function buildChartSeries(
  chart: Pick<SpreadsheetChart, "range" | "seriesInColumns"> &
    Partial<Pick<SpreadsheetChart, "type" | "sourceRefError" | "noCategories" | "noSeriesNames" | "nameRowsAbove">>,
  cells: ChartCellLookup,
  // What a VALUE cell reads, when it should differ from what a label reads.
  // Labels want the cell as displayed (a formatted date, a padded code); values
  // must be the number itself: a displayed "1,234.56" or "76,543.21" does not
  // parse as a number, so every row of a thousands-formatted column was dropped
  // from its chart and the chart drew only the rows small enough to need no
  // separator. Omitted, values read the same lookup as labels, as before.
  values: ChartCellLookup = cells,
): string[][] {
  // A #REF! source plots nothing; its placeholder range names no data.
  if (chart.sourceRefError) return [];
  const { range } = chart;
  const seriesInColumns = chart.seriesInColumns !== false;
  const out: string[][] = [];
  const str = (v: string | undefined): string => (v == null ? "" : String(v));
  
  const above = seriesInColumns && chart.noSeriesNames !== true ? Math.min(chart.nameRowsAbove ?? 0, range.r1) : 0;
  const nameOverRows = (row: number, col: number): string | null =>
    above > 0 ? chartNameOverRows([row - above, row], col, cells) : null;
  // XY families have no category column. Keep the first numeric vector both as
  // a stable point label and as the X series so a two-column X/Y selection is
  // executable instead of being misread as category + one value series.
  if (chart.type === "scatter" || chart.type === "bubble") {
    const count = seriesInColumns ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1;
    const xAt = (i: number): string =>
      seriesInColumns ? str(values(range.r1 + i, range.c1)) : str(values(range.r1, range.c1 + i));
    const labelAt = (i: number): string =>
      seriesInColumns ? str(cells(range.r1 + i, range.c1)) : str(cells(range.r1, range.c1 + i));
    const ysAt = (i: number): string[] => {
      const ys: string[] = [];
      if (seriesInColumns) {
        for (let c = range.c1 + 1; c <= range.c2; c += 1) ys.push(str(values(range.r1 + i, c)));
      } else {
        for (let r = range.r1 + 1; r <= range.r2; r += 1) ys.push(str(values(r, range.c1 + i)));
      }
      return ys;
    };
    // Excel plots a text X column at 1, 2, 3… (its axis then counts the
    // points) and keeps the text as each point's label. The first row is
    // the header when a value cell of it is text.
    const header = count > 1 && ysAt(0).some((v) => v.trim() !== "" && !numericText(v));
    const first = header ? 1 : 0;
    let textX = chart.type === "scatter" && count > first;
    let anyX = false;
    for (let i = first; textX && i < count; i += 1) {
      const x = xAt(i).trim();
      if (x === "") continue;
      anyX = true;
      if (numericText(x)) textX = false;
    }
    textX = textX && anyX;
    for (let i = 0; i < count; i += 1) {
      out.push([labelAt(i), textX && i >= first ? String(i - first + 1) : xAt(i), ...ysAt(i)]);
    }
    return out;
  }
  if (nativeChartStatesLabelLines(chart)) {
    // The range says which of its edge lines are labels: one row per point
    // line, header included when it names the series, aligned to the range
    // (row i is the range's row r1 + i with series in columns), each with
    // its category label ("" -- read as 1, 2, 3... -- without one) and one
    // cell per series. Without names its first row is a point, read with
    // `nativeChartHeaderRow`.
    const layout = nativeChartRangeLayout({ ...chart, range });
    const cell = (line: number, point: number, lookup: ChartCellLookup) =>
      str(layout.byColumns ? lookup(point, line) : lookup(line, point));
    const pointFirst = layout.byColumns ? range.r1 : range.c1;
    const pointLast = layout.byColumns ? range.r2 : range.c2;
    for (let point = pointFirst; point <= pointLast; point += 1) {
      const name = point === layout.nameLine;
      out.push([
        name || layout.categoryLine === null ? "" : cell(layout.categoryLine, point, cells),
        ...layout.seriesLines.map((line) =>
          (name && layout.byColumns ? nameOverRows(point, line) : null) ?? cell(line, point, name ? cells : values)),
      ]);
    }
    return out;
  }
  if (seriesInColumns) {
    const oneColumn = range.c1 === range.c2;
    // One-column range: the single column is both label and value. Otherwise the
    // first column is the label and each later column is a value series.
    const firstValueCol = oneColumn ? range.c1 : range.c1 + 1;
    for (let r = range.r1; r <= range.r2; r += 1) {
      const row: string[] = [str(cells(r, range.c1))];
      for (let c = firstValueCol; c <= range.c2; c += 1) {
        row.push((r === range.r1 && !oneColumn ? nameOverRows(r, c) : null) ?? str(values(r, c)));
      }
      out.push(row);
    }
    return out;
  }
  const oneRow = range.r1 === range.r2;
  const firstValueRow = oneRow ? range.r1 : range.r1 + 1;
  for (let c = range.c1; c <= range.c2; c += 1) {
    const row: string[] = [str(cells(range.r1, c))];
    for (let r = firstValueRow; r <= range.r2; r += 1) {
      row.push(str(values(r, c)));
    }
    out.push(row);
  }
  return out;
}

const CHART_A1_RANGE_RE = /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?([0-9]{1,7})(?::\$?([A-Za-z]{1,3})\$?([0-9]{1,7}))?$/;

export interface ParsedChartRefRange {
  sheetName?: string;
  range: ChartSourceRange;
}

export function parseChartRefRange(value: string | undefined): ParsedChartRefRange | null {
  if (!value) return null;
  const match = CHART_A1_RANGE_RE.exec(value.trim());
  if (!match) return null;
  // A quoted name is the sheet's name as it is, its spaces too: R295-X02's
  // "' Global IDP'" found no cells as "Global IDP", and its value axis lost
  // the source's "#,##0" (Linked to source) to General.
  const sheetName = match[1] ? match[1].replace(/''/g, "'") : match[2]?.trim();
  const c1 = lettersToColumn(match[3]);
  const r1 = Number(match[4]) - 1;
  const c2 = match[5] ? lettersToColumn(match[5]) : c1;
  const r2 = match[6] ? Number(match[6]) - 1 : r1;
  if (
    !Number.isInteger(r1) ||
    !Number.isInteger(r2) ||
    r1 < 0 ||
    r2 < r1 ||
    c1 < 0 ||
    c2 < c1 ||
    r2 > MAX_RANGE_ROW ||
    c2 > MAX_RANGE_COL
  ) {
    return null;
  }
  return { ...(sheetName ? { sheetName } : {}), range: { r1, c1, r2, c2 } };
}

/**
 * The areas of a chart reference: one range, or a union of ranges that a
 * chart over a non-contiguous selection states as `(Sheet1!$A$1:$A$2,
 * Sheet1!$A$4:$A$5)` (Excel plots the areas' cells one after the other;
 * R20_chart_column07). Null when any area is not a plain range.
 */
export function parseChartRefAreas(value: string | undefined): ParsedChartRefRange[] | null {
  const parts = chartRefAreaTexts(value);
  if (!parts) return null;
  const areas = parts.map((part) => parseChartRefRange(part));
  return areas.every((area): area is ParsedChartRefRange => area !== null) ? areas : null;
}

/** A chart reference's areas as written: the union's parts, or the one range. */
export function chartRefAreaTexts(value: string | undefined): string[] | null {
  if (!value) return null;
  const text = value.trim();
  if (!(text.startsWith("(") && text.endsWith(")"))) return [text];
  const parts: string[] = [];
  let quoted = false;
  let start = 1;
  for (let index = 1; index < text.length - 1; index += 1) {
    const char = text[index];
    if (char === "'") quoted = !quoted;
    else if (char === "," && !quoted) {
      parts.push(text.slice(start, index));
      start = index + 1;
    }
  }
  parts.push(text.slice(start, text.length - 1));
  return parts.map((part) => part.trim());
}

function chartRefRange(
  value: string | undefined,
  sheetName: string,
): ChartSourceRange | null {
  const parsed = parseChartRefRange(value);
  if (!parsed) return null;
  if (
    parsed.sheetName &&
    parsed.sheetName.toLocaleLowerCase() !== sheetName.toLocaleLowerCase()
  ) {
    return null;
  }
  return parsed.range;
}

export interface SpreadsheetChartNumberFormatCell {
  value?: string;
  numberFormat?: unknown;
}

export type SpreadsheetChartNumberFormatCellLookup = (
  sheetName: string,
  row: number,
  col: number,
) => SpreadsheetChartNumberFormatCell | undefined;

/** A range-bound chart that states where its labels are, read through the
 * references its layout plots (`nativeChartRangeSeriesRefs`) by the source-
 * format readings below; the chart itself otherwise. */
function withStatedLayoutRefs<T extends { type?: unknown; chartType?: unknown; range?: unknown; seriesInColumns?: unknown; seriesRefs?: unknown; noCategories?: unknown; noSeriesNames?: unknown }>(chart: T): T {
  if (chart.seriesRefs !== undefined || (chart.noCategories !== true && chart.noSeriesNames !== true)) return chart;
  const range = normalizeChartSourceRange(chart.range);
  if (!range) return chart;
  const refs = nativeChartRangeSeriesRefs({
    type: normalizeSpreadsheetChartType(chart.type ?? chart.chartType),
    range,
    ...(chart.seriesInColumns === false ? { seriesInColumns: false } : {}),
    ...(chart.noCategories === true ? { noCategories: true as const } : {}),
    ...(chart.noSeriesNames === true ? { noSeriesNames: true as const } : {}),
  });
  return { ...chart, seriesRefs: refs ?? null };
}

export function resolveLinkedValueAxisNumberFormat(
  chart: {
    type?: unknown;
    chartType?: unknown;
    range?: unknown;
    seriesInColumns?: unknown;
    seriesRefs?: unknown;
    secondaryAxisSeries?: unknown;
    noCategories?: unknown;
    noSeriesNames?: unknown;
  },
  currentSheetName: string,
  lookup: SpreadsheetChartNumberFormatCellLookup,
): SpreadsheetChartLinkedFormatResolution {
  chart = withStatedLayoutRefs(chart);
  const type = normalizeSpreadsheetChartType(chart.type ?? chart.chartType);
  if (!["column", "stacked-column", "line", "area"].includes(type)) {
    return { state: "indeterminate", reason: "unsupported", formatCode: "General" };
  }
  if (
    Array.isArray(chart.secondaryAxisSeries) &&
    chart.secondaryAxisSeries.length > 0
  ) {
    return { state: "indeterminate", reason: "unsupported", formatCode: "General" };
  }

  const cells: Array<{ sheetName: string; row: number; col: number }> = [];
  const refs = normalizeChartSeriesRefs(chart.seriesRefs);
  if (chart.seriesRefs !== undefined) {
    if (!refs || refs.some((ref) => ref.axis === "secondary")) {
      return { state: "indeterminate", reason: "reference", formatCode: "General" };
    }
    for (const ref of refs) {
      const parsed = parseChartRefRange(ref.values);
      if (!parsed || (parsed.range.r1 !== parsed.range.r2 && parsed.range.c1 !== parsed.range.c2)) {
        return { state: "indeterminate", reason: "reference", formatCode: "General" };
      }
      const sourceSheet = parsed.sheetName ?? currentSheetName;
      if (parsed.range.c1 === parsed.range.c2) {
        for (let row = parsed.range.r1; row <= parsed.range.r2; row += 1) {
          cells.push({ sheetName: sourceSheet, row, col: parsed.range.c1 });
        }
      } else {
        for (let col = parsed.range.c1; col <= parsed.range.c2; col += 1) {
          cells.push({ sheetName: sourceSheet, row: parsed.range.r1, col });
        }
      }
    }
  } else {
    const range = normalizeChartSourceRange(chart.range);
    if (!range) {
      return { state: "indeterminate", reason: "reference", formatCode: "General" };
    }
    if (chart.seriesInColumns !== false) {
      const firstValueCol = range.c1 === range.c2 ? range.c1 : range.c1 + 1;
      let firstValueRow = range.r1;
      if (
        range.r2 > range.r1 &&
        Array.from({ length: range.c2 - firstValueCol + 1 }, (_, index) =>
          lookup(currentSheetName, range.r1, firstValueCol + index)?.value,
        ).some((value) => value == null || value.trim() === "" || !Number.isFinite(Number(value)))
      ) {
        firstValueRow += 1;
      }
      for (let row = firstValueRow; row <= range.r2; row += 1) {
        for (let col = firstValueCol; col <= range.c2; col += 1) {
          cells.push({ sheetName: currentSheetName, row, col });
        }
      }
    } else {
      if (range.r2 <= range.r1 || range.c2 <= range.c1) {
        return { state: "indeterminate", reason: "reference", formatCode: "General" };
      }
      const headerRowIsLabels = Array.from(
        { length: range.c2 - range.c1 },
        (_, index) => lookup(currentSheetName, range.r1, range.c1 + index + 1)?.value,
      ).every(
        (value) =>
          value != null &&
          value.trim() !== "" &&
          !Number.isFinite(Number(value)),
      );
      const seriesColumnIsLabels = Array.from(
        { length: range.r2 - range.r1 },
        (_, index) => lookup(currentSheetName, range.r1 + index + 1, range.c1)?.value,
      ).every(
        (value) =>
          value != null &&
          value.trim() !== "" &&
          !Number.isFinite(Number(value)),
      );
      if (!headerRowIsLabels || !seriesColumnIsLabels) {
        return { state: "indeterminate", reason: "reference", formatCode: "General" };
      }
      for (let row = range.r1 + 1; row <= range.r2; row += 1) {
        for (let col = range.c1 + 1; col <= range.c2; col += 1) {
          cells.push({ sheetName: currentSheetName, row, col });
        }
      }
    }
  }
  if (cells.length === 0 || cells.length > 200_000) {
    return { state: "indeterminate", reason: "reference", formatCode: "General" };
  }
  const formats = new Set<string>();
  for (const target of cells) {
    const cell = lookup(target.sheetName, target.row, target.col);
    if (!cell || cell.numberFormat === undefined) {
      return { state: "indeterminate", reason: "missing", formatCode: "General" };
    }
    const normalized = normalizeNumberFormatCode(cell.numberFormat);
    if (normalized === null) {
      return { state: "indeterminate", reason: "invalid", formatCode: "General" };
    }
    if (!normalized) {
      return { state: "indeterminate", reason: "missing", formatCode: "General" };
    }
    formats.add(normalized);
    if (formats.size > 1) {
      return { state: "indeterminate", reason: "mixed", formatCode: "General" };
    }
  }
  return { state: "resolved", formatCode: [...formats][0] };
}

/**
 * Each series' source format (Linked to source): the one number format every
 * value cell of the series carries, by series index; undefined where its
 * cells carry none, General, or differing ones (the planner then shows the
 * value in General). What a data label with no format of its own shows, as
 * Excel writes it into the series' `c:numCache/c:formatCode` (CDC p43:
 * "17,422" and "33.1"). Undefined when the references cannot be read.
 */
export function resolveLinkedSeriesNumberFormats(
  chart: { type?: unknown; range?: unknown; seriesInColumns?: unknown; seriesRefs?: unknown; noCategories?: unknown; noSeriesNames?: unknown },
  currentSheetName: string,
  lookup: SpreadsheetChartNumberFormatCellLookup,
): (string | undefined)[] | undefined {
  chart = withStatedLayoutRefs(chart);
  const groups: Array<Array<{ sheetName: string; row: number; col: number }>> = [];
  const refs = normalizeChartSeriesRefs(chart.seriesRefs);
  if (chart.seriesRefs !== undefined) {
    if (!refs) return undefined;
    for (const ref of refs) {
      const parsed = parseChartRefRange(ref.values);
      const cells: Array<{ sheetName: string; row: number; col: number }> = [];
      if (parsed && (parsed.range.r1 === parsed.range.r2 || parsed.range.c1 === parsed.range.c2)) {
        const sheetName = parsed.sheetName ?? currentSheetName;
        for (let row = parsed.range.r1; row <= parsed.range.r2; row += 1) {
          for (let col = parsed.range.c1; col <= parsed.range.c2; col += 1) cells.push({ sheetName, row, col });
        }
      }
      groups.push(cells);
    }
  } else {
    const range = normalizeChartSourceRange(chart.range);
    if (!range) return undefined;
    const numeric = (value: string | undefined) => value != null && value.trim() !== "" && Number.isFinite(Number(value));
    if (chart.seriesInColumns !== false) {
      const firstCol = range.c1 === range.c2 ? range.c1 : range.c1 + 1;
      const header =
        range.r2 > range.r1 &&
        Array.from({ length: range.c2 - firstCol + 1 }, (_, index) => lookup(currentSheetName, range.r1, firstCol + index)?.value).some(
          (value) => !numeric(value),
        );
      for (let col = firstCol; col <= range.c2; col += 1) {
        groups.push(
          Array.from({ length: range.r2 - range.r1 + (header ? 0 : 1) }, (_, index) => ({
            sheetName: currentSheetName,
            row: range.r1 + index + (header ? 1 : 0),
            col,
          })),
        );
      }
    } else {
      if (range.r2 <= range.r1 || range.c2 <= range.c1) return undefined;
      for (let row = range.r1 + 1; row <= range.r2; row += 1) {
        groups.push(
          Array.from({ length: range.c2 - range.c1 }, (_, index) => ({ sheetName: currentSheetName, row, col: range.c1 + index + 1 })),
        );
      }
    }
  }
  if (!groups.length || groups.reduce((total, cells) => total + cells.length, 0) > 200_000) return undefined;
  const formats = groups.map((cells) => {
    let code: string | undefined;
    for (const target of cells) {
      const cell = lookup(target.sheetName, target.row, target.col);
      // An empty cell is a blank point: it states no format of the series'.
      if (!cell || cell.value == null || String(cell.value).trim() === "") continue;
      const normalized = normalizeNumberFormatCode(cell.numberFormat ?? "") ?? "";
      if (code !== undefined && code !== normalized) return undefined;
      code = normalized;
    }
    return code || undefined;
  });
  return formats.some(Boolean) ? formats : undefined;
}

export interface ChartWorkbookLookup {
  display: (sheetName: string, row: number, col: number) => string | number | undefined;
  value: (sheetName: string, row: number, col: number) => string | number | undefined;
}

function valuesForChartRef(
  ref: string | undefined,
  sheetName: string,
  cells: ChartCellLookup,
  workbookCells?: ChartWorkbookLookup["value"],
): string[] | null {
  // A union of areas (a chart over a non-contiguous selection) plots each
  // area's cells in turn (R20_chart_column07).
  const areas = chartRefAreaTexts(ref);
  if (areas && areas.length > 1) {
    const values: string[] = [];
    for (const area of areas) {
      const part = valuesForChartRef(area, sheetName, cells, workbookCells);
      if (!part) return null;
      values.push(...part);
    }
    return values;
  }
  const parsed = workbookCells ? parseChartRefRange(ref) : null;
  const range = workbookCells ? parsed?.range : chartRefRange(ref, sheetName);
  if (!range || (range.r1 !== range.r2 && range.c1 !== range.c2)) return null;
  const lookup = (row: number, col: number) => workbookCells
    ? workbookCells(parsed?.sheetName ?? sheetName, row, col) : cells(row, col);
  const values: string[] = [];
  if (range.c1 === range.c2) {
    for (let row = range.r1; row <= range.r2; row += 1) {
      const value = lookup(row, range.c1);
      if (workbookCells && value === undefined) return null;
      values.push(String(value ?? ""));
    }
  } else {
    for (let col = range.c1; col <= range.c2; col += 1) {
      const value = lookup(range.r1, col);
      if (workbookCells && value === undefined) return null;
      values.push(String(value ?? ""));
    }
  }
  return values;
}

/**
 * A series name read from its Select Data > Series name cells: one cell's
 * text, or several cells' texts joined with a space, as Excel names a series
 * whose name reference spans cells (the API reader's `chartSeriesNameCells`).
 * Null when the reference names no readable cells.
 */
function seriesNameFromChartRef(
  ref: string | undefined,
  sheetName: string,
  cells: ChartCellLookup,
  workbookCells?: ChartWorkbookLookup["display"],
): string | null {
  const values = valuesForChartRef(ref, sheetName, cells, workbookCells);
  if (!values?.length) return null;
  if (values.length === 1) return values[0];
  
  // filled cell left of it, as Excel shows it (probe/n21-x56 p705c).
  const areas = chartRefAreaTexts(ref);
  const parsed = areas && areas.length === 1 ? parseChartRefRange(ref) : null;
  if (parsed && parsed.range.c1 === parsed.range.c2 && parsed.range.r2 > parsed.range.r1) {
    const lookup = (row: number, col: number) => workbookCells
      ? workbookCells(parsed.sheetName ?? sheetName, row, col)
      : cells(row, col);
    const joined = chartNameOverRows([parsed.range.r1, parsed.range.r2], parsed.range.c1, lookup);
    if (joined !== null) return joined;
  }
  return values.filter((value) => value !== "").join(" ");
}

/**
 * XLS-A107: a category area over several columns (or, for series in rows,
 * several rows) -- Excel's multi-level category axis, `c:multiLvlStrRef`: the
 * innermost line's labels, and the outer lines as levels, outermost first,
 * one entry per point (a blank outer cell continues the group above it, as
 * Excel reads it). Null for a one-line area.
 */
export function multiLevelCategoriesForChartRef(
  ref: string | undefined,
  sheetName: string,
  cells: ChartCellLookup,
  workbookCells?: ChartWorkbookLookup["display"],
): { labels: string[]; levels: string[][] } | null {
  const parsed = workbookCells ? parseChartRefRange(ref) : null;
  const range = workbookCells ? parsed?.range : chartRefRange(ref, sheetName);
  if (!range || range.r1 === range.r2 || range.c1 === range.c2) return null;
  const lookup = (row: number, col: number) => workbookCells
    ? workbookCells(parsed?.sheetName ?? sheetName, row, col) : cells(row, col);
  const byColumns = range.r2 - range.r1 >= range.c2 - range.c1;
  const depth = byColumns ? range.c2 - range.c1 + 1 : range.r2 - range.r1 + 1;
  if (depth < 2 || depth > MAX_CHART_CATEGORY_LEVELS) return null;
  const points = byColumns ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1;
  const line = (level: number): string[] | null => {
    const out: string[] = [];
    for (let index = 0; index < points; index += 1) {
      const value = byColumns
        ? lookup(range.r1 + index, range.c1 + level)
        : lookup(range.r1 + level, range.c1 + index);
      if (workbookCells && value === undefined) return null;
      out.push(String(value ?? ""));
    }
    return out;
  };
  const lines: string[][] = [];
  for (let level = 0; level < depth; level += 1) {
    const values = line(level);
    if (!values) return null;
    lines.push(values);
  }
  return { labels: lines[depth - 1]!, levels: lines.slice(0, depth - 1) };
}

/** Whether a range-bound chart's first row of rows (`buildChartSeries`) names
 * its series: false for a range without names; undefined when the reading
 * tells (the first row's value cells). */
export function nativeChartHeaderRow(chart: Pick<SpreadsheetChart, "noSeriesNames" | "noCategories">): boolean | undefined {
  if (chart.noSeriesNames === true) return false;
  return chart.noCategories === true ? true : undefined;
}

/**
 * The outer levels of a range-bound chart's multi-level axis
 * (`outerCategoryColumns`), one entry per row of `range` (header included),
 * outermost first; undefined for a one-level axis.
 */
export function nativeChartOuterCategoryCells(
  chart: Pick<SpreadsheetChart, "range" | "outerCategoryColumns" | "seriesInColumns"> &
    Partial<Pick<SpreadsheetChart, "type" | "noCategories" | "noSeriesNames">>,
  cells: ChartCellLookup,
): string[][] | undefined {
  const count = chart.outerCategoryColumns ?? 0;
  // A tree map's or a sunburst's text columns are its branches with or
  
  // top plots J1:K3 as two levels); a category axis's outer levels sit
  // under a row of names only.
  const hierarchyWithoutNames =
    (chart.type === "treemap" || chart.type === "sunburst") && chart.noCategories !== true;
  if (
    count < 1 ||
    chart.seriesInColumns === false ||
    chart.range.c1 - count < 0 ||
    (nativeChartStatesLabelLines(chart) && !hierarchyWithoutNames)
  )
    return undefined;
  const levels: string[][] = [];
  for (let col = chart.range.c1 - count; col < chart.range.c1; col += 1) {
    const level: string[] = [];
    for (let row = chart.range.r1; row <= chart.range.r2; row += 1) level.push(String(cells(row, col) ?? ""));
    levels.push(level);
  }
  return levels;
}

/**
 * Levels aligned to rows that start with the series-title row (as the
 * chart's live rows do) as the reading wants them: one entry per DATA row.
 * `rows` is the plotted matrix; `levels` are aligned to it row for row.
 */
export function chartCategoryLevelsForRows(
  rows: readonly (readonly string[])[],
  levels: readonly (readonly string[])[] | undefined,
  headerRow?: boolean,
): string[][] | undefined {
  if (!levels?.length || rows.length === 0 || levels.some((level) => level.length !== rows.length)) return undefined;
  const firstValues = (rows[0] ?? []).slice(1);
  const header =
    rows.length > 1 &&
    (headerRow ?? firstValues.some((value) => {
      const text = String(value ?? "").trim();
      return text === "" || !Number.isFinite(Number(text));
    }));
  return levels.map((level) => (header ? level.slice(1) : [...level]));
}

/**
 * XLS-A107: the multi-level axis an imported chart's category area states
 * (several columns), aligned to the rows `buildChartSeriesFromRefs` builds
 * (one entry per data row); undefined for a one-level axis.
 */
export function chartCategoryLevelsFromRefs(
  rawRefs: unknown,
  chartType: SpreadsheetChartType,
  sheetName: string,
  displayCells: ChartCellLookup,
  workbook?: ChartWorkbookLookup,
): string[][] | undefined {
  if (chartType === "scatter" || chartType === "bubble" || chartType === "treemap" || chartType === "sunburst") return undefined;
  const refs = normalizeChartSeriesRefs(rawRefs);
  const multi = refs ? multiLevelCategoriesForChartRef(refs[0]?.categories, sheetName, displayCells, workbook?.display) : null;
  return multi ? multi.levels : undefined;
}

/** Joins a hierarchy chart row's level path inside its one category label. */
export const CHART_HIERARCHY_PATH_SEPARATOR = "";

/**
 * One label per row of a treemap or sunburst category area that spans several
 * level columns, each carrying the row's whole level path. As in Excel, a blank
 * parent cell repeats the parent above it until a level is given anew; a row
 * that leaves a gap between levels fails closed.
 */
function hierarchyLabelsForChartRef(
  ref: string | undefined,
  sheetName: string,
  cells: ChartCellLookup,
  workbookCells?: ChartWorkbookLookup["value"],
): string[] | null {
  const parsed = workbookCells ? parseChartRefRange(ref) : null;
  const range = workbookCells ? parsed?.range : chartRefRange(ref, sheetName);
  if (!range || range.c1 === range.c2 || range.c2 - range.c1 >= 50) return null;
  const lookup = (row: number, col: number) => workbookCells
    ? workbookCells(parsed?.sheetName ?? sheetName, row, col) : cells(row, col);
  const labels: string[] = [];
  let previous: string[] = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    const levels: string[] = [];
    for (let col = range.c1; col <= range.c2; col += 1) {
      const value = lookup(row, col);
      if (workbookCells && value === undefined) return null;
      levels.push(String(value ?? "").trim().split(CHART_HIERARCHY_PATH_SEPARATOR).join(" "));
    }
    const deepest = levels.reduce((last, level, index) => (level ? index : last), -1);
    if (deepest < 0) return null;
    const path: string[] = [];
    let restated = false;
    for (let index = 0; index <= deepest; index += 1) {
      if (levels[index]) {
        path.push(levels[index]);
        restated = true;
      } else if (!restated && index < previous.length) {
        path.push(previous[index]);
      } else {
        return null;
      }
    }
    previous = path;
    labels.push(path.join(CHART_HIERARCHY_PATH_SEPARATOR));
  }
  return labels;
}

function finiteChartVector(values: string[] | null): values is string[] {
  return Boolean(
    values?.length &&
    values.every(
      (value) => value.trim() !== "" && Number.isFinite(Number(value)),
    ),
  );
}

/**
 * Resolve native Scatter/Bubble refs without forcing them into a shared-X row
 * matrix. Every series keeps its own authenticated X/Y/(size) vectors; nothing
 * is aligned by position across series and no adjacent worksheet range is ever
 * inferred. Invalid or mixed graphs fail closed as one unit.
 */
export function buildChartXYSeriesFromRefs(
  rawRefs: unknown,
  chartType: SpreadsheetChartType,
  sheetName: string,
  displayCells: ChartCellLookup,
  valueCells: ChartCellLookup = displayCells,
  workbook?: ChartWorkbookLookup,
): MiniChartXYSeries[] {
  if (chartType !== "scatter" && chartType !== "bubble") return [];
  const refs = normalizeChartSeriesRefs(rawRefs);
  if (!refs) return [];
  if (
    refs.some(
      (ref) =>
        ref.categories !== undefined ||
        (ref.renderType !== undefined && ref.renderType !== chartType) ||
        (chartType === "scatter" && ref.bubbleSizes !== undefined),
    )
  ) {
    return [];
  }
  const allImplicit = refs.every(
    (ref) => ref.implicitX === true && ref.xValues === undefined,
  );
  const allExplicit = refs.every(
    (ref) => ref.implicitX !== true && ref.xValues !== undefined,
  );
  if (
    (!allImplicit && !allExplicit) ||
    (chartType === "bubble" && allImplicit)
  ) {
    return [];
  }

  if (refs.every((ref) => chartRefCellsGone(ref.values))) return [];
  let totalPoints = 0;
  const series: MiniChartXYSeries[] = [];
  for (let seriesIndex = 0; seriesIndex < refs.length; seriesIndex += 1) {
    const ref = refs[seriesIndex];
    if (chartRefCellsGone(ref.values)) {
      // Its cells were deleted: no point, and the other series keep their places.
      series.push({ name: ref.titleText ?? `Series ${seriesIndex + 1}`, points: [] });
      continue;
    }
    const readY = valuesForChartRef(ref.values, sheetName, valueCells, workbook?.value);
    const wholeX = allImplicit
      ? null
      : valuesForChartRef(ref.xValues, sheetName, valueCells, workbook?.value);
    // More X cells than Y cells: the points are the Y cells', each at its X, as
    // Excel draws R12-X06's pH chart (a series of 896 values against 1,442 X cells
    // stops short of the others).
    const readX = wholeX && readY && wholeX.length > readY.length ? wholeX.slice(0, readY.length) : wholeX;
    // A point whose Y cell is empty is no point (Show empty cells as Gaps,
    // the part's own dispBlanksAs): C16_fdo70609 binds five rows and plots
    // the four that hold values. A text X is still placed by its number.
    const plotted = (index: number) => readY?.[index]?.trim() !== "";
    const yValues = readY && readY.some((_, index) => !plotted(index)) ? readY.filter((_, index) => plotted(index)) : readY;
    const statedX =
      readX && readY && readX.length === readY.length && yValues !== readY
        ? readX.filter((_, index) => plotted(index))
        : readX;
    // X values that hold text are plotted at 1, 2, 3, ... as Excel plots
    // them (Microsoft Support, "Present your data in a scatter chart or a
    // line chart"; the API's literal reading restates them the same way):
    // C28_tdf126033 names its two points "Value" and "Value2". A blank or
    // unread vector still leaves the chart in custody.
    const textX =
      statedX !== null &&
      yValues !== null &&
      statedX.length === yValues.length &&
      statedX.some((value) => value.trim() !== "" && !Number.isFinite(Number(value)));
    const xValues = allImplicit || textX
      ? yValues &&
        Array.from({ length: yValues.length }, (_, index) => String(index + 1))
      : statedX;
    const readSizes =
      chartType === "bubble"
        ? valuesForChartRef(ref.bubbleSizes, sheetName, valueCells, workbook?.value)
        : null;
    const sizeValues =
      readSizes && readY && readSizes.length === readY.length && yValues !== readY
        ? readSizes.filter((_, index) => plotted(index))
        : readSizes;
    if (
      !finiteChartVector(xValues) ||
      !finiteChartVector(yValues) ||
      xValues.length !== yValues.length ||
      (chartType === "bubble" &&
        (!finiteChartVector(sizeValues) || sizeValues.length !== xValues.length))
    ) {
      return [];
    }
    const title = seriesNameFromChartRef(ref.title, sheetName, displayCells, workbook?.display);
    if (ref.title !== undefined && title === null) return [];
    const name = title ?? ref.titleText ?? `Series ${seriesIndex + 1}`;
    totalPoints += xValues.length;
    if (totalPoints > MAX_CHART_POINTS) return [];
    series.push({
      name,
      points: xValues.map((x, pointIndex) => ({
        x: Number(x),
        y: Number(yValues[pointIndex]),
        ...(chartType === "bubble"
          ? { size: Number((sizeValues as string[])[pointIndex]) }
          : {}),
      })),
    });
  }
  return series;
}

/**
 * Resolve imported native series refs against the live sheet into MiniChart's
 * bounded row matrix. Categories and titles use the display channel, while
 * plotted values use the raw/computed channel: currency and percent formatting
 * must never turn a valid number into an unplottable "$1,000" or "25%" string.
 * An optional workbook lookup resolves every sheet-qualified vector against
 * its own source. Without it, the original single-sheet boundary is unchanged.
 * No reference is partially projected: unavailable sheets, non-vector ranges or
 * misaligned lengths return [] and leave the imported object in preserve-only
 * custody.
 */
/** A reference whose cells were deleted: Excel stores ``Sheet!#REF!``. */
function chartRefCellsGone(ref: unknown): boolean {
  return typeof ref === "string" && ref.includes("#REF!");
}

export function buildChartSeriesFromRefs(
  rawRefs: unknown,
  chartType: SpreadsheetChartType,
  sheetName: string,
  displayCells: ChartCellLookup,
  valueCells: ChartCellLookup = displayCells,
  workbook?: ChartWorkbookLookup,
  options: { dateCategories?: boolean } = {},
): string[][] {
  const refs = normalizeChartSeriesRefs(rawRefs);
  if (!refs) return [];
  if (chartType === "scatter" || chartType === "bubble") {
    const xySeries = buildChartXYSeriesFromRefs(
      refs,
      chartType,
      sheetName,
      displayCells,
      valueCells,
      workbook,
    );
    if (xySeries.length === 0) return [];
    // MiniChart's bubble matrix has one X/Y/size tuple stream and cannot carry
    // native series identity, legend or style for more than one bubble series.
    // Keep those graphs in preserve-only custody until the renderer owns a
    // per-series XY model rather than flattening and misrepresenting them.
    if (chartType === "bubble" && xySeries.length !== 1) return [];

    if (chartType === "bubble") {
      const entry = xySeries[0];
      const rows: string[][] = [["", "X", entry.name, "Size"]];
      entry.points.forEach((point) => {
        rows.push([
          String(point.x),
          String(point.x),
          String(point.y),
          String(point.size),
        ]);
      });
      return rows;
    }

    const sharedX = xySeries[0].points.map((point) => point.x);
    // MiniChart's scatter matrix has exactly one shared X column. Multiple
    // explicit series are therefore representable only when their authenticated
    // X vectors are numerically identical; differing-X graphs fail closed
    // instead of fabricating alignment or inferring an adjacent range.
    if (
      xySeries.some(
        (entry) =>
          entry.points.length !== sharedX.length ||
          !entry.points.every((point, index) => point.x === sharedX[index]),
      )
    ) {
      return [];
    }
    const rows: string[][] = [["", "X", ...xySeries.map((entry) => entry.name)]];
    for (let index = 0; index < sharedX.length; index += 1) {
      rows.push([
        String(sharedX[index]),
        String(sharedX[index]),
        ...xySeries.map((entry) => String(entry.points[index].y)),
      ]);
    }
    return rows;
  }

  // A series whose cells were deleted (#REF!) has no point; the others still
  // plot. All of them gone leaves nothing to draw.
  const values = refs.map((ref) =>
    chartRefCellsGone(ref.values) ? [] : valuesForChartRef(ref.values, sheetName, valueCells, workbook?.value),
  );
  if (values.some((entry) => !entry) || refs.every((ref) => chartRefCellsGone(ref.values))) return [];
  // Series of different lengths (a gauge's 4-point doughnut over an 8-point
  // pie, C20_tdf111173): Excel plots each series' own points, the chart
  // reaching as far as its longest series, and a shorter series has no
  // point there -- a blank cell, drawn by Show empty cells as.
  let length = Math.max(0, ...(values as string[][]).map((entry) => entry.length));
  const resolved = (values as string[][]).map((entry) =>
    entry.length === length ? entry : [...entry, ...Array.from({ length: length - entry.length }, () => "")],
  );
  if (length === 0) return [];
  // A date axis places each category by its date: the cells' values (Excel
  // serials), not the text their format writes ("10-May" names no year).
  // Cells that are not all dates keep their text, as Excel's axis falls back.
  const dateValues = options.dateCategories
    ? valuesForChartRef(refs[0]?.categories, sheetName, valueCells, workbook?.value)
    : null;
  const datedValues =
    dateValues && dateValues.every((value) => value.trim() === "" || /^\d+(\.\d+)?$/.test(value.trim()))
      ? dateValues
      : null;
  const statedCategories =
    datedValues ??
    (chartType === "treemap" || chartType === "sunburst"
      ? hierarchyLabelsForChartRef(refs[0]?.categories, sheetName, displayCells, workbook?.display)
      : null) ??
    valuesForChartRef(
      refs[0]?.categories,
      sheetName,
      displayCells,
      workbook?.display,
    ) ??
    // XLS-A107: a multi-level category area labels each point with its
    // innermost line; the outer lines are its levels.
    multiLevelCategoriesForChartRef(refs[0]?.categories, sheetName, displayCells, workbook?.display)?.labels ??
    null;
  // A category range shorter than the longest series names the points it
  // reaches; the rest are numbered from their place, as with no categories.
  // Longer than every series, its labels past them have no point: Excel's bands
  // are the points (R24-X02: "SD" and five values; R103-X05's axis stops short of
  // its 118 labels). Refused here, R116-X02's chart of 8 points under 21 labels
  // drew nothing and said its source could not be verified.
  const categoryValues = statedCategories && statedCategories.length > length ? statedCategories.slice(0, length) : statedCategories;
  // On a shared category axis (columns, bars, lines, areas) Excel takes the
  // axis from the first series: its categories set the count, and a later
  // series' points past them get no band (X-303, XW060). A pie or doughnut
  // has no shared axis; each ring keeps its own points (C20).
  const ownPoints = /^(pie|doughnut|ofPie|pie3d|sunburst|treemap|funnel)/i;
  const sharedAxis = !ownPoints.test(String(chartType)) &&
    !refs.some((ref) => ownPoints.test(String((ref as { renderType?: unknown }).renderType ?? "")));
  const firstSeriesCategories = categoryValues?.length ?? 0;
  // A series of the secondary group (its own category axis, shown or
  // deleted) keeps its own points, and the axis widens to the longest of
  // them: Excel draws XW307's 7-point secondary line over 7 bands, the
  // primary series' 5 categories labelling the first five (X-342).
  const ownAxis = (index: number) =>
    refs[0]?.axis !== "secondary" && refs[index]?.axis === "secondary";
  let blankBandsFrom = Infinity;
  if (
    sharedAxis && firstSeriesCategories > 0 && firstSeriesCategories < length &&
    ((values as string[][])[0]?.length ?? 0) <= firstSeriesCategories
  ) {
    const widened = Math.max(
      firstSeriesCategories,
      ...(values as string[][]).map((entry, index) => (ownAxis(index) ? entry.length : 0)),
    );
    const trimmed = resolved.map((entry, index) =>
      ownAxis(index)
        ? entry.slice(0, widened)
        : [
            ...entry.slice(0, firstSeriesCategories),
            ...Array.from({ length: widened - firstSeriesCategories }, () => ""),
          ],
    );
    resolved.splice(0, resolved.length, ...trimmed);
    length = widened;
    blankBandsFrom = firstSeriesCategories;
  }
  const titleValues = refs.map((ref, index) => {
    const title = seriesNameFromChartRef(ref.title, sheetName, displayCells, workbook?.display);
    return title ?? ref.titleText ?? `Series ${index + 1}`;
  });
  const rows: string[][] = [["", ...titleValues]];
  for (let index = 0; index < length; index += 1) {
    rows.push([
      index >= blankBandsFrom ? "" : (categoryValues?.[index] ?? String(index + 1)),
      ...resolved.map((entry) => entry[index]),
    ]);
  }
  return rows;
}

/**
 * The trendline of one series on a CATEGORY axis, in the caller's 0-based index
 * space. Office fits a category axis on the ordinal from 1 — which is what makes
 * a logarithmic or a power fit possible at all (ln 0 is undefined); fed the raw
 * index, those two kinds drew nothing. A forecast is clipped to the half period
 * either side of the categories this small renderer has room for.
 */
export function computeCategoryTrendline(
  values: ReadonlyArray<number>,
  spec: ChartTrendlineSpec | undefined,
): TrendlinePoint[] {
  if (!spec) return [];
  const ordinal = values.map((y, index) => ({ x: index + 1, y }));
  const contract = spreadsheetTrendlineContract(spec);
  const curved =
    contract &&
    contract.type !== "linear" &&
    contract.type !== "movingAverage";
  const fitted = curved
    ? officeTrendlinePolyline(
        contract,
        ordinal,
        // Dense enough for a curve to read as one between two categories.
        Array.from(
          { length: (values.length - 1) * 8 + 1 },
          (_, step) => 1 + step / 8,
        ),
      )
    : computeTrendline(ordinal, spec);
  return fitted
    .map((point) => ({ x: point.x - 1, y: point.y }))
    .filter((point) => point.x >= -0.5 && point.x <= values.length - 0.5);
}

/** The equation / R² label of a sheet trendline, or null when none is shown. */
export function spreadsheetTrendlineLabel(
  points: ReadonlyArray<TrendlinePoint>,
  spec: ChartTrendlineSpec | undefined,
): string | null {
  const contract = spec ? spreadsheetTrendlineContract(spec) : undefined;
  return contract ? officeTrendlineLabel(contract, points) : null;
}

/**
 * Compute the polyline (in DATA space) of a trendline over an ordered set of
 * (x, y) points. Pure + deterministic so it is unit-testable in isolation and
 * carries no rendering — the renderer maps the returned x/y through whatever
 * pixel scale its branch already uses.
 *
 * The caller supplies the points already in data space: for a category chart
 * (bar/line/area) x is the point INDEX (0..n-1) and y is the series value; for a
 * scatter x is the scatter X and y its Y. Non-finite points are dropped first
 * (a blank/non-numeric value never NaNs the fit). Returns [] (no overlay) when
 * fewer than 2 finite points remain or the spec is unusable.
 *
 *  - `linear`: ordinary least-squares fit y = slope·x + intercept over the
 *    finite points, returned as the two endpoints from the smallest to the
 *    largest x (a straight segment spanning the data). A zero-variance x (all
 *    points share one x) has no defined slope -> [] (no overlay), never a
 *    divide-by-zero NaN.
 *  - `movingAverage`: the trailing average of the last `period` y-values at each
 *    position, emitted from the first position that has a full window
 *    (index period-1) onward, each keeping its own x. `period` clamps to
 *    [2, point count]; with <2 points or period>count the window can't form ->
 *    [] for the all-too-short case, else a shortened polyline.
 */
export function computeTrendline(
  points: ReadonlyArray<TrendlinePoint>,
  spec: ChartTrendlineSpec | undefined,
): TrendlinePoint[] {
  if (!spec || !Array.isArray(points)) return [];
  // Keep only finite (x, y) points so a blank/non-numeric cell can't NaN the math.
  const pts = points.filter(
    (p) => p && Number.isFinite(p.x) && Number.isFinite(p.y),
  );
  if (pts.length < 2) return [];

  if (
    spec.type === "exponential" ||
    spec.type === "logarithmic" ||
    spec.type === "power" ||
    spec.type === "polynomial" ||
    (spec.type === "linear" &&
      (spec.forward !== undefined ||
        spec.backward !== undefined ||
        spec.intercept !== undefined))
  ) {
    // The curves, and a line with a forecast or a pinned intercept, are the
    // shared fit's; this module keeps only the two plain kinds it has always
    // drawn, whose exact output its own callers and tests pin.
    const contract = spreadsheetTrendlineContract(spec);
    return contract ? officeTrendlinePolyline(contract, pts) : [];
  }

  if (spec.type === "linear") {
    const n = pts.length;
    let sx = 0;
    let sy = 0;
    let sxx = 0;
    let sxy = 0;
    for (const p of pts) {
      sx += p.x;
      sy += p.y;
      sxx += p.x * p.x;
      sxy += p.x * p.y;
    }
    const denom = n * sxx - sx * sx;
    // Zero x-variance (every x equal) -> slope undefined; no meaningful line.
    if (denom === 0) return [];
    const slope = (n * sxy - sx * sy) / denom;
    const intercept = (sy - slope * sx) / n;
    let minX = pts[0].x;
    let maxX = pts[0].x;
    for (const p of pts) {
      if (p.x < minX) minX = p.x;
      if (p.x > maxX) maxX = p.x;
    }
    return [
      { x: minX, y: slope * minX + intercept },
      { x: maxX, y: slope * maxX + intercept },
    ];
  }

  // movingAverage: trailing average over a window of `period` points.
  const period = Math.max(
    DEFAULT_TRENDLINE_PERIOD,
    Math.min(
      pts.length,
      Math.trunc(
        Number.isFinite(spec.period as number)
          ? (spec.period as number)
          : DEFAULT_TRENDLINE_PERIOD,
      ) || DEFAULT_TRENDLINE_PERIOD,
    ),
  );
  const out: TrendlinePoint[] = [];
  for (let i = period - 1; i < pts.length; i += 1) {
    let sum = 0;
    for (let k = i - period + 1; k <= i; k += 1) sum += pts[k].y;
    out.push({ x: pts[i].x, y: sum / period });
  }
  return out;
}

/**
 * Population standard deviation of a numeric series, ignoring non-finite values.
 * Pure helper for the stddev error-bar amount. Returns 0 for an empty / single
 * finite value (no spread) so a stddev whisker on a degenerate series is a
 * zero-length bar rather than a NaN.
 */
export function seriesStdDev(values: ReadonlyArray<number>): number {
  const finite = values.filter((v) => Number.isFinite(v));
  if (finite.length < 2) return 0;
  const mean = finite.reduce((s, v) => s + v, 0) / finite.length;
  // The SAMPLE standard deviation (n - 1), the one Word, Excel and PowerPoint
  // draw their Standard Deviation and Standard Error bars with (measured on
  // 4, 6, 5, 7: Standard Error ±0.645 = 1.291 / √4, Standard Deviation
  // ±1.291 about the mean; the population value would give 0.559 / 1.118).
  const variance =
    finite.reduce((s, v) => s + (v - mean) * (v - mean), 0) / (finite.length - 1);
  return Math.sqrt(Math.max(0, variance));
}

/**
 * Where an error bar is centred. Office draws Standard Deviation bars about
 * the series' MEAN, the same span at every point (measured in Word: 4, 6, 5, 7
 * with one standard deviation draws 4.21-6.79 at all four points); every other
 * amount is centred on the point's own value.
 */
export function makeErrorBarCenter(
  spec: ChartErrorBarSpec | undefined,
  values: ReadonlyArray<number>,
): (value: number) => number {
  if (spec?.type !== "stddev") return (value) => value;
  const finite = values.filter((v) => Number.isFinite(v));
  if (!finite.length) return (value) => value;
  const mean = finite.reduce((s, v) => s + v, 0) / finite.length;
  return () => mean;
}

/**
 * Build a per-point error-bar DELTA function (the +/- half-length of the whisker
 * in DATA units) for a given error-bar spec over a value series. Pure +
 * deterministic so the renderer just maps the returned delta through its own
 * pixel scale, and it is unit-testable in isolation.
 *
 *  - `fixed`: a constant `value` on every point (independent of the data).
 *  - `percent`: `|y| * value / 100` for each point's own magnitude (so a larger
 *    bar gets a proportionally larger whisker; a zero point gets none).
 *  - `stddev`: `(value || 1) * σ` where σ is the series' population standard
 *    deviation — the SAME delta on every point (Excel's "Standard Deviation"
 *    amount, with `value` as the σ multiplier; a 0/absent multiplier means 1σ).
 *
 * Every branch clamps the result to a finite, non-negative number so a blank /
 * NaN point never emits a NaN whisker. Returns a constant-0 function when the
 * spec is absent so the caller can call it unconditionally (the renderer then
 * draws nothing — a zero-length whisker is skipped).
 */
export function makeErrorBarDelta(
  spec: ChartErrorBarSpec | undefined,
  values: ReadonlyArray<number>,
): (value: number) => number {
  if (!spec) return () => 0;
  if (spec.type === "stderr") {
    // The standard error of the series: its standard deviation over the
    // square root of its point count, the same amount on every point (the
    // shared chart renderer's Standard Error).
    const count = values.filter(Number.isFinite).length;
    const delta = count > 0 ? seriesStdDev(values) / Math.sqrt(count) : 0;
    return () => (Number.isFinite(delta) ? Math.max(0, delta) : 0);
  }
  if (spec.type === "stddev") {
    const sigma = seriesStdDev(values);
    const mult = spec.value > 0 ? spec.value : 1;
    const delta = Math.max(0, mult * sigma);
    return () => (Number.isFinite(delta) ? delta : 0);
  }
  const amount = Math.max(0, spec.value);
  if (spec.type === "percent") {
    return (value: number) => {
      if (!Number.isFinite(value)) return 0;
      const d = (Math.abs(value) * amount) / 100;
      return Number.isFinite(d) ? Math.max(0, d) : 0;
    };
  }
  // fixed: a constant amount on every point.
  return () => amount;
}

export {
  computeHistogramBins,
  DEFAULT_HISTOGRAM_INTERVAL_CLOSED,
  histogramBinLabel,
  type HistogramBin,
  type HistogramIntervalClosed,
} from "../office-charts/histogram-bins";

/** One waterfall step: its start (carry-in), end (carry-out), delta, and sign. */
export interface WaterfallStep {
  label: string;
  /** Running total BEFORE this step (where the bar's base sits). */
  start: number;
  /** Running total AFTER this step (where the bar's top sits). */
  end: number;
  /** The step's signed value (end - start). */
  delta: number;
  /** True for a synthetic final TOTAL bar (drawn from 0 to the grand total). */
  isTotal: boolean;
}

/**
 * Turn an ordered list of (label, delta) points into cumulative WATERFALL steps:
 * each step starts where the previous one ended, so the running total walks up
 * and down. Pure + deterministic (the renderer maps start/end through its pixel
 * scale and colors each bar by delta sign).
 *
 * Non-finite deltas are treated as 0 (a blank cell is a no-op step, not a NaN
 * that poisons every later running total). When `withTotal` is true a final
 * synthetic bar is appended spanning 0..grandTotal (Excel's "Total" column).
 *
 * Returns [] when there are no input points.
 */
export function computeWaterfallSteps(
  points: ReadonlyArray<{ label: string; value: number }>,
  withTotal = false,
): WaterfallStep[] {
  if (!Array.isArray(points) || points.length === 0) return [];
  const out: WaterfallStep[] = [];
  let running = 0;
  for (const p of points) {
    const delta = Number.isFinite(p.value) ? p.value : 0;
    const start = running;
    const end = start + delta;
    running = end;
    out.push({
      label: p?.label ?? "",
      start,
      end,
      delta,
      isTotal: false,
    });
  }
  if (withTotal) {
    out.push({
      label: "Total",
      start: 0,
      end: running,
      delta: running,
      isTotal: true,
    });
  }
  return out;
}

// ─── Box-and-Whisker (Excel QUARTILE.INC / linear interpolation) ──────────────

export {
  computeBoxWhiskerStats,
  quartileExc,
  quartileInc,
  type BoxWhiskerStats,
} from "../office-charts/box-whisker-statistics";

// ─── Pareto (sorted-descending bars + cumulative-% line) ──────────────────────

/** One data point for a Pareto chart: its label, value, and cumulative % (0–100). */
export interface ParetoPoint {
  label: string;
  value: number;
  /** Cumulative percentage 0–100 up to and including this point. */
  cumPct: number;
}

/**
 * Build the Pareto series from an ordered list of (label, value) points: sort by
 * value descending, compute the cumulative percentage after each bar, and return
 * the sorted list. Non-finite values are treated as 0 (a blank cell contributes
 * no count). Returns [] when there are no input points or the total is zero (no
 * cumulative percentage is defined). Pure + deterministic for unit testing.
 *
 * The cumulative % at position k is:
 *   sum(sorted[0..k].value) / totalPositive * 100
 * where totalPositive is the sum of all positive values (negative values and
 * zeros do not contribute to the total, matching Excel Pareto behaviour).
 */
export function computeParetoPoints(
  points: ReadonlyArray<{ label: string; value: number }>,
): ParetoPoint[] {
  if (!Array.isArray(points) || points.length === 0) return [];
  // Coerce non-finite values to 0.
  const coerced = points.map((p) => ({
    label: p?.label ?? "",
    value: Number.isFinite(p.value) ? p.value : 0,
  }));
  // Sort descending by value.
  const sorted = [...coerced].sort((a, b) => b.value - a.value);
  // Total of positive values only (Excel ignores negatives in cumulative %).
  const total = sorted.reduce((s, p) => s + Math.max(0, p.value), 0);
  if (total === 0) {
    // No positive values → return sorted list with 0 cumPct (bars at 0 height).
    return sorted.map((p) => ({ ...p, cumPct: 0 }));
  }
  let running = 0;
  return sorted.map((p) => {
    running += Math.max(0, p.value);
    return { label: p.label, value: p.value, cumPct: (running / total) * 100 };
  });
}

// ─── Treemap (squarified / slice-and-dice layout) ─────────────────────────────

/**
 * One rectangle in a treemap layout, in normalized [0, 1] space. `x`/`y` are the
 * top-left corner and `w`/`h` are the width/height (all 0..1). The renderer maps
 * these through the chart pixel bounds. `label` and `value` are the original data.
 */
export interface TreemapRect {
  label: string;
  value: number;
  /** Top-left x in [0, 1] space. */
  x: number;
  /** Top-left y in [0, 1] space. */
  y: number;
  /** Width in [0, 1] space (proportional to value). */
  w: number;
  /** Height in [0, 1] space. */
  h: number;
  /** 0-based index into the input array (used to pick a palette color). */
  index: number;
}

/**
 * Build a flat treemap using a slice-and-dice layout: items are laid out in the
 * larger remaining axis, alternating horizontal/vertical splits at each level.
 * Pure + deterministic (no SVG geometry) for unit testing.
 *
 * Implementation: simple binary-split / slice-and-dice (the first value takes
 * its proportional share of the current axis, then the rest recurse). Items are
 * sorted descending by value so the largest rectangle appears first (top-left).
 * Non-positive / non-finite values are treated as 0 (they receive 0 area).
 *
 * NOTE: This is a flat single-level treemap. Multi-level category paths from
 * Excel's treemap (where the first N columns form a hierarchy) are limited to
 * grouping by the first column only — full recursive squarification is deferred.
 *
 * Returns [] when there are no positive values.
 */
export function computeTreemapRects(
  points: ReadonlyArray<{ label: string; value: number }>,
): TreemapRect[] {
  if (!Array.isArray(points) || points.length === 0) return [];
  // Coerce values; keep track of original indices for palette mapping.
  const items = points.map((p, i) => ({
    label: p?.label ?? "",
    value: Number.isFinite(p.value) && p.value > 0 ? p.value : 0,
    index: i,
  }));
  const total = items.reduce((s, p) => s + p.value, 0);
  if (total === 0) return [];
  // Sort descending by value (largest first).
  const sorted = [...items].sort((a, b) => b.value - a.value);
  const out: TreemapRect[] = [];

  // Recursive slice-and-dice: split the current rect in the larger axis.
  function layout(
    items2: typeof sorted,
    x: number,
    y: number,
    w: number,
    h: number,
    subtotal: number,
    horizontal: boolean,
  ): void {
    if (items2.length === 0 || subtotal <= 0) return;
    if (items2.length === 1) {
      out.push({ ...items2[0], x, y, w: Math.max(0, w), h: Math.max(0, h) });
      return;
    }
    // Split off the first item in the current axis direction.
    const first = items2[0];
    const rest = items2.slice(1);
    const restTotal = rest.reduce((s, p) => s + p.value, 0);
    if (horizontal) {
      const wFirst = (first.value / subtotal) * w;
      out.push({ ...first, x, y, w: Math.max(0, wFirst), h: Math.max(0, h) });
      layout(rest, x + wFirst, y, Math.max(0, w - wFirst), h, restTotal, false);
    } else {
      const hFirst = (first.value / subtotal) * h;
      out.push({ ...first, x, y, w: Math.max(0, w), h: Math.max(0, hFirst) });
      layout(rest, x, y + hFirst, w, Math.max(0, h - hFirst), restTotal, true);
    }
  }

  layout(sorted, 0, 0, 1, 1, total, true);
  return out;
}

// ─── Sunburst (radial concentric ring segments) ────────────────────────────────

/**
 * One arc segment in a sunburst layout. Angles are in RADIANS. For a flat
 * single-level sunburst this degrades to a doughnut (one ring of segments).
 *
 * `startAngle` and `endAngle` are in radians (0 = 3 o'clock; subtract Math.PI/2
 * to start at 12 o'clock as the renderer normalises). `ring` is the 0-based ring
 * depth (0 = innermost for a multi-level hierarchy; in a flat layout all segments
 * share ring 0). `innerR`/`outerR` are fractional radii in [0, 1] space (1 =
 * the chart's available radius).
 */
export interface SunburstSegment {
  label: string;
  value: number;
  startAngle: number; // radians, 0 = 12 o'clock (after renderer offsets)
  endAngle: number;
  /** 0-based ring depth (0 = outermost for flat data, innermost for deep). */
  ring: number;
  /** Inner radius fraction [0, 1]. */
  innerR: number;
  /** Outer radius fraction [0, 1]. */
  outerR: number;
  /** 0-based index into the input array (for palette). */
  index: number;
}

/**
 * Build a flat single-level sunburst segment list: one ring whose segments span
 * the full circle proportionally to their value. Angles start at -π/2 (12
 * o'clock) and sweep clockwise. Non-positive / non-finite values are treated as
 * 0 (they receive 0 sweep). Returns [] when there are no positive values.
 *
 * For a multi-level hierarchy the innermost ring represents the top-level
 * category. With flat input all segments share `ring = 0` (the outer ring),
 * `innerR = 0.3`, `outerR = 1.0` — matching a doughnut.
 *
 * Pure + deterministic for unit testing.
 */
export function computeSunburstSegments(
  points: ReadonlyArray<{ label: string; value: number }>,
): SunburstSegment[] {
  if (!Array.isArray(points) || points.length === 0) return [];
  const items = points.map((p, i) => ({
    label: p?.label ?? "",
    value: Number.isFinite(p.value) && p.value > 0 ? p.value : 0,
    index: i,
  }));
  const total = items.reduce((s, p) => s + p.value, 0);
  if (total === 0) return [];
  // Flat single-level: inner ring = 0.3 (same hole ratio as doughnut), outer = 1.
  const INNER_R = 0.3;
  const OUTER_R = 1.0;
  const out: SunburstSegment[] = [];
  let angle = -Math.PI / 2; // start at 12 o'clock
  for (const item of items) {
    const sweep = (item.value / total) * Math.PI * 2;
    out.push({
      label: item.label,
      value: item.value,
      startAngle: angle,
      endAngle: angle + sweep,
      ring: 0,
      innerR: INNER_R,
      outerR: OUTER_R,
      index: item.index,
    });
    angle += sweep;
  }
  return out;
}

// ─── Stock / Candlestick (OHLC) ───────────────────────────────────────────────

/**
 * One candlestick in a stock chart. Coordinates are in DATA space; the renderer
 * maps them through its own value scale.
 *
 * Excel supports two variants:
 *  - High-Low-Close (3 series): no open box, just a wick + close tick.
 *  - Open-High-Low-Close (4 series): a box from open to close + the wick.
 *
 * `open` is null in the HLC variant (no fourth series). `up` = true when
 * close >= open (or when only HLC data is present and close >= midpoint(hi, lo)).
 */
export interface CandlestickBar {
  /** Category label (date string or index). */
  label: string;
  /** Open price (null in HLC mode). */
  open: number | null;
  /** High price (top of wick). */
  high: number;
  /** Low price (bottom of wick). */
  low: number;
  /** Close price (bottom or top of the open-close box, or the close tick). */
  close: number;
  /** True when close >= open (up candle: bullish / green); false = down / red. */
  up: boolean;
  /** 0-based index within the category axis. */
  index: number;
}

/**
 * Build the candlestick bar list from a multi-series data view. The series
 * order follows Excel's OHLC convention:
 *
 *  - 4+ series: series[0]=Open, series[1]=High, series[2]=Low, series[3]=Close
 *  - 3 series:  series[0]=High, series[1]=Low, series[2]=Close  (HLC variant)
 *  - < 3 series: returns [] (insufficient data for even an HLC chart).
 *
 * Non-finite values at a given category are dropped (the bar is omitted) so a
 * blank cell in the data never produces a NaN bar. Returns [] when no usable
 * bars exist. Pure + deterministic for unit testing.
 *
 * NOTE: xlsx StockChart export via openpyxl is not yet wired — deferred.
 */
export function computeCandlestickBars(
  categories: ReadonlyArray<string>,
  series: ReadonlyArray<ReadonlyArray<number>>,
): CandlestickBar[] {
  if (!Array.isArray(series) || !Array.isArray(categories)) return [];
  const n = categories.length;
  if (n === 0 || series.length < 3) return [];
  // Determine OHLC vs HLC mode.
  const isOHLC = series.length >= 4;
  const openSeries = isOHLC ? series[0] : null;
  const highSeries = isOHLC ? series[1] : series[0];
  const lowSeries = isOHLC ? series[2] : series[1];
  const closeSeries = isOHLC ? series[3] : series[2];
  const out: CandlestickBar[] = [];
  for (let i = 0; i < n; i += 1) {
    const high = highSeries[i];
    const low = lowSeries[i];
    const close = closeSeries[i];
    if (!Number.isFinite(high) || !Number.isFinite(low) || !Number.isFinite(close)) {
      continue; // skip a row with missing OHLC data
    }
    const open = openSeries ? openSeries[i] : null;
    if (open !== null && !Number.isFinite(open)) continue;
    // Up/down: close >= open in OHLC mode; close >= midpoint in HLC mode.
    const up =
      open !== null
        ? close >= open
        : close >= (high + low) / 2;
    out.push({
      label: categories[i] ?? String(i),
      open,
      high,
      low,
      close,
      up,
      index: i,
    });
  }
  return out;
}
