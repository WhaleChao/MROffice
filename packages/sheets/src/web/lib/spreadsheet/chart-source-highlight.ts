/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's chart data highlight: a selected chart's source data, outlined on
 * the sheet in colour by what each cell is to the chart, and the drags on
 * those outlines that change the chart's data.
 *
 * Microsoft Support, "Add a data series to your chart": "Select the chart.
 * The currently displayed source data is selected on the worksheet, showing
 * sizing handles. On the worksheet, drag the sizing handles to include the
 * new data. The chart is updated automatically." It works for a chart "on the
 * same worksheet as the data you used to create the chart"; a chart whose
 * data is on another sheet takes the Select Data Source dialog instead.
 * "Update the data in an existing chart" names the colours: the row or column
 * of the data series labels in red, the plotted values in blue, the category
 * axis in a third colour (Excel for Windows and Mac draw it purple, as the X
 * values of an XY chart).
 *
 * What is drawn ({@link chartSourceReading}, {@link chartSourceBoxes}):
 * - every series' name cell (red), its categories / X values (purple), its
 *   values (blue) and a bubble chart's sizes (blue), each only where a cell
 *   reference states it (a typed name, `1, 2, 3` X values draw nothing) and
 *   only on the chart's own sheet (a reference to another sheet draws no box
 *   here, a chart reading another sheet draws none);
 * - the chart area (or any element but a series) selected: every series,
 *   the same role's cells of neighbouring series boxed together (two series
 *   in E6:E7 and F6:F7 are one blue box E6:F7, their names one red box);
 *   one series (or one of its points) selected: that series alone;
 * - a range-bound chart reads its range as it plots it (`buildChartSeries`,
 *   the shared reading's header rule): series in columns, the first column
 *   holds the categories (an XY chart's X values) and the first row the
 *   series names when a value cell of it is not a number; series in rows,
 *   the same transposed. A PivotChart, a Camera picture, a #REF! source and
 *   a chart whose data is its own draw nothing.
 *
 * What a drag does ({@link chartSourceDragged}), the rules Excel's outlines
 * follow (resize from a corner handle, move by the edge): the names, the
 * categories and the values stay one layout -- the names follow the value
 * columns, the categories follow the value rows (series in columns; the
 * other way for series in rows):
 * - the values: resized from any corner, moved anywhere, never back over the
 *   names or the categories and never to nothing;
 * - the names: resized or moved along their row only (series added, removed
 *   or read from other columns);
 * - the categories: resized or moved along their column only (points added,
 *   removed or read from other rows).
 * With one series selected the drag edits that series alone (its values stay
 * one row or column). The result is written the way Select Data writes it
 * ({@link ChartSourceEdit}): a chart born here keeps one data range while the
 * layout is one (names right above, categories right beside), else it states
 * its series one by one; an imported chart's series references are rewritten.
 */

import { quotedSheetName } from "@mrnote/core/web/lib/spreadsheet/change-log.ts";
import {
  nativeChartRangeLayout,
  nativeChartStatesLabelLines,
  normalizeChartSeriesRefs,
  normalizeChartSourceRange,
  normalizeSpreadsheetChartBinding,
  normalizeSpreadsheetChartType,
  parseChartRefRange,
  type ChartSourceRange,
  type SpreadsheetChartSeriesRef,
  type SpreadsheetChartType,
} from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { chartRangeToA1, normalizeChartRange } from "@mrnote/core/web/lib/spreadsheet/chart-config.ts";
import { drawingChartSheetOf, drawingChartTakesDataRange } from "@mrnote/core/web/lib/spreadsheet/drawing-chart-series.ts";
import { FORMULA_REFERENCE_COLORS } from "@mrnote/core/web/lib/spreadsheet/formula-reference-highlight.ts";
import { columnToLetters } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

/** What a highlighted cell is to the chart. */
export type ChartSourceRole = "name" | "category" | "values" | "sizes";

/** Excel's colours for the roles (the range finder's own palette). */
export const CHART_SOURCE_ROLE_COLORS: Readonly<Record<ChartSourceRole, string>> = {
  name: FORMULA_REFERENCE_COLORS[1],
  category: FORMULA_REFERENCE_COLORS[2],
  values: FORMULA_REFERENCE_COLORS[0],
  sizes: FORMULA_REFERENCE_COLORS[0],
};

/** One series' cells on the chart's own sheet (absent: not a cell there). */
export interface ChartSourceSeriesCells {
  readonly name?: ChartSourceRange;
  readonly categories?: ChartSourceRange;
  readonly values?: ChartSourceRange;
  readonly sizes?: ChartSourceRange;
}

/** What a series reference carries beside its cells, kept through an edit. */
type SeriesExtras = Pick<SpreadsheetChartSeriesRef, "titleText" | "renderType" | "axis">;

export interface ChartSourceReading {
  readonly kind: "native" | "drawing";
  readonly type: SpreadsheetChartType;
  /** Every series, in the chart's order. */
  readonly series: readonly ChartSourceSeriesCells[];
  /**
   * How the chart's data can be edited by dragging, or null (display only):
   * "range" -- a chart born here bound to one range (its orientation in
   * `seriesInColumns`); "refs" -- a chart that states its series one by one.
   */
  readonly edit: null | {
    readonly via: "range" | "refs";
    /** Series in columns or in rows; undefined: each layout says (series
     * stated one by one need not be one block). */
    readonly seriesInColumns?: boolean;
    /** A range-bound chart's outer category columns (XLS-A107), kept. */
    readonly outerCategoryColumns: number;
    /** A range-bound chart whose range has no category column
     * (`noCategories`): a block of several value columns is still its range. */
    readonly noCategories?: boolean;
    /** The sheet as the series references spell it. */
    readonly prefix: string;
    /** Per series, what its reference states beside its cells. */
    readonly extras: readonly SeriesExtras[];
    /** A combination keeps its number of series. */
    readonly countLocked: boolean;
    /** A chart born here may become one stating its series one by one
     * (Select Data's Legend Entries, XLS-A072) when a drag leaves its range. */
    readonly convertible: boolean;
  };
}

export interface ChartSourceInput {
  readonly kind: "native" | "drawing";
  /** The chart record (a `sheet.charts` entry or a chart drawing object). */
  readonly chart: object;
  /** The chart's own sheet. */
  readonly sheetName: string;
  readonly sheetId: string;
  /** A cell of the chart's own sheet as the chart reads a value. */
  readonly cellValue: (row: number, col: number) => string;
}

const XY_TYPES = new Set<SpreadsheetChartType>(["scatter", "bubble"]);

const transpose = (range: ChartSourceRange): ChartSourceRange => ({
  r1: range.c1,
  c1: range.r1,
  r2: range.c2,
  c2: range.r2,
});

const rect = (r1: number, c1: number, r2: number, c2: number): ChartSourceRange => ({ r1, c1, r2, c2 });

function sameRange(a: ChartSourceRange | undefined, b: ChartSourceRange | undefined): boolean {
  if (!a || !b) return a === b;
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

/** A value cell as the shared reading takes it: a number, or not. */
function finiteValue(text: string): boolean {
  return text.trim() !== "" && Number.isFinite(Number(text));
}

/** A cell an XY chart reads as a number (a separator, currency or percent
 * sign allowed): `buildChartSeries`' own test. */
function numericText(text: string): boolean {
  const stripped = text.trim().replace(/[$,%\s]/g, "");
  return stripped !== "" && Number.isFinite(Number(stripped));
}

/** The series a range-bound chart reads off `range`, in canonical
 * orientation (series in columns), as `buildChartSeries` and the shared
 * reading lay it. */
function rangeSeries(
  range: ChartSourceRange,
  type: SpreadsheetChartType,
  value: (row: number, col: number) => string,
  outer: number,
): ChartSourceSeriesCells[] {
  const { r1, c1, r2, c2 } = range;
  const out: ChartSourceSeriesCells[] = [];
  if (XY_TYPES.has(type)) {
    // X in the first column, each later column a Y series; the first row
    // names them when a Y cell of it is text.
    if (c2 <= c1) return [];
    let header = false;
    if (r2 > r1) {
      for (let col = c1 + 1; col <= c2 && !header; col += 1) {
        const text = value(r1, col);
        header = text.trim() !== "" && !numericText(text);
      }
    }
    const first = header ? r1 + 1 : r1;
    const x = rect(first, c1, r2, c1);
    // A bubble chart plots one series: its Y column and its size column.
    const last = type === "bubble" ? c1 + 1 : c2;
    for (let col = c1 + 1; col <= last; col += 1) {
      out.push({
        ...(header ? { name: rect(r1, col, r1, col) } : {}),
        categories: x,
        values: rect(first, col, r2, col),
        ...(type === "bubble" && c2 >= c1 + 2 ? { sizes: rect(first, c1 + 2, r2, c1 + 2) } : {}),
      });
    }
    return out;
  }
  const oneColumn = c1 === c2;
  const firstValueCol = oneColumn ? c1 : c1 + 1;
  // The first row names the series when a value cell of it is not a number.
  let header = false;
  if (r2 > r1) {
    for (let col = firstValueCol; col <= c2 && !header; col += 1) header = !finiteValue(value(r1, col));
  }
  const first = header ? r1 + 1 : r1;
  // A one-column range plots its column as the values alone.
  const labels = oneColumn ? undefined : rect(first, c1 - outer, r2, c1);
  for (let col = firstValueCol; col <= c2; col += 1) {
    out.push({
      ...(header ? { name: rect(r1, col, r1, col) } : {}),
      ...(labels ? { categories: labels } : {}),
      values: rect(first, col, r2, col),
    });
  }
  return out;
}

/** The series of a range (canonical, series in columns) that states which
 * of its edge lines are labels (`nativeChartRangeLayout`). */
function statedRangeSeries(
  range: ChartSourceRange,
  flags: { noCategories?: true; noSeriesNames?: true },
): ChartSourceSeriesCells[] {
  const layout = nativeChartRangeLayout({ range, ...flags });
  const [first, last] = layout.points;
  if (last < first) return [];
  return layout.seriesLines.map((col) => ({
    ...(layout.nameLine !== null ? { name: rect(layout.nameLine, col, layout.nameLine, col) } : {}),
    ...(layout.categoryLine !== null ? { categories: rect(first, layout.categoryLine, last, layout.categoryLine) } : {}),
    values: rect(first, col, last, col),
  }));
}

function orient(series: ChartSourceSeriesCells, flip: boolean): ChartSourceSeriesCells {
  if (!flip) return series;
  return {
    ...(series.name ? { name: transpose(series.name) } : {}),
    ...(series.categories ? { categories: transpose(series.categories) } : {}),
    ...(series.values ? { values: transpose(series.values) } : {}),
    ...(series.sizes ? { sizes: transpose(series.sizes) } : {}),
  };
}

/** A reference's cells when it is on the chart's own sheet. */
function ownCells(formula: string | undefined, sheetName: string): ChartSourceRange | undefined {
  if (!formula) return undefined;
  const parsed = parseChartRefRange(formula);
  if (!parsed) return undefined;
  if (parsed.sheetName !== undefined && parsed.sheetName.toLowerCase() !== sheetName.toLowerCase()) return undefined;
  return parsed.range;
}

/**
 * What a chart's source data is on its own sheet, or null when the chart
 * shows none there (data on another sheet, a PivotChart, a Camera picture,
 * a #REF! source, data of its own).
 */
export function chartSourceReading(input: ChartSourceInput): ChartSourceReading | null {
  const chart = input.chart as Record<string, unknown>;
  if (chart.pivotSource || chart.sourceRefError === true) return null;
  if (normalizeSpreadsheetChartBinding(chart.binding)?.literal) return null;
  const type = input.kind === "drawing"
    ? normalizeSpreadsheetChartType(chart.chartType)
    : normalizeSpreadsheetChartType(chart.type);
  if (input.kind === "native" && chart.type === "camera") return null;
  const refs = normalizeChartSeriesRefs(chart.seriesRefs);
  if (input.kind === "native" && !refs) {
    const sourceSheetId = typeof chart.sourceSheetId === "string" ? chart.sourceSheetId.trim() : "";
    if (sourceSheetId && sourceSheetId !== input.sheetId) return null;
    const range = normalizeChartSourceRange(chart.range);
    if (!range) return null;
    const seriesInColumns = chart.seriesInColumns !== false;
    const outer = seriesInColumns && typeof chart.outerCategoryColumns === "number" &&
      Number.isInteger(chart.outerCategoryColumns) && chart.outerCategoryColumns > 0 &&
      range.c1 - chart.outerCategoryColumns >= 0
      ? chart.outerCategoryColumns
      : 0;
    const canonical = seriesInColumns ? range : transpose(range);
    const value = seriesInColumns ? input.cellValue : (row: number, col: number) => input.cellValue(col, row);
    const flags = {
      ...(chart.noCategories === true ? { noCategories: true as const } : {}),
      ...(chart.noSeriesNames === true ? { noSeriesNames: true as const } : {}),
    };
    // A range that states where its labels are lays its series out as that
    // (no category column, no row of names); the reading's own guess at a
    // heading does not apply.
    const series = (
      !XY_TYPES.has(type) && nativeChartStatesLabelLines(flags)
        ? statedRangeSeries(canonical, flags)
        : rangeSeries(canonical, type, value, outer)
    ).map((entry) => orient(entry, !seriesInColumns));
    if (!series.length) return null;
    const editable = type !== "bubble";
    return {
      kind: "native",
      type,
      series,
      edit: editable
        ? {
            via: "range",
            seriesInColumns,
            outerCategoryColumns: outer,
            ...(flags.noCategories ? { noCategories: true } : {}),
            prefix: quotedSheetName(input.sheetName),
            extras: series.map(() => ({})),
            countLocked: false,
            // Select Data's own gate for a chart born here (selectDataSeriesAbilities).
            convertible: !XY_TYPES.has(type) && drawingChartTakesDataRange(type),
          }
        : null,
    };
  }
  if (!refs) return null;
  const series: ChartSourceSeriesCells[] = refs.map((ref) => {
    const name = ownCells(ref.title, input.sheetName);
    const categories = ownCells(ref.categories ?? ref.xValues, input.sheetName);
    const values = ownCells(ref.values, input.sheetName);
    const sizes = ownCells(ref.bubbleSizes, input.sheetName);
    return {
      ...(name ? { name } : {}),
      ...(categories ? { categories } : {}),
      ...(values ? { values } : {}),
      ...(sizes ? { sizes } : {}),
    };
  });
  if (!series.some((entry) => entry.name || entry.categories || entry.values || entry.sizes)) return null;
  // Series stated one by one are edited one by one while every reference is
  // a cell range of this sheet the series rewrite takes (Select Data's gate).
  const sheet = drawingChartSheetOf(refs);
  const editable: boolean =
    input.kind === "drawing" &&
    drawingChartTakesDataRange(type) &&
    sheet !== null &&
    sheet.sheetName.toLowerCase() === input.sheetName.toLowerCase() &&
    refs.every((ref) => !ref.names && !ref.implicitX && !ref.bubbleSizes);
  const block = editable ? blockOf(series) : null;
  return {
    kind: input.kind,
    type,
    series,
    edit: editable
      ? {
          via: "refs",
          ...(block ? { seriesInColumns: block.seriesInColumns } : {}),
          outerCategoryColumns: 0,
          prefix: sheet?.prefix ?? quotedSheetName(input.sheetName),
          extras: refs.map((ref) => ({
            ...(ref.titleText !== undefined ? { titleText: ref.titleText } : {}),
            ...(ref.renderType ? { renderType: ref.renderType } : {}),
            ...(ref.axis ? { axis: ref.axis } : {}),
          })),
          countLocked: refs.some((ref) => ref.renderType !== undefined || ref.axis !== undefined),
          convertible: false,
        }
      : null,
  };
}

/** The series a chart's selected element names: `series:S`, `point:S:P`, a
 * pie's `slice:P`; null for every other element (the whole chart). */
export function chartSourceSelectedSeries(area: string | null | undefined): number | null {
  if (!area) return null;
  const match = /^(?:series|point):(\d+)(?::\d+)?$/.exec(area);
  if (match) return Number(match[1]);
  return /^slice:\d+$/.test(area) ? 0 : null;
}

/** One outline: a role's cells, the series it stands for. */
export interface ChartSourceBox {
  readonly key: string;
  readonly role: ChartSourceRole;
  readonly color: string;
  readonly range: ChartSourceRange;
}

const ROLE_ORDER: readonly ChartSourceRole[] = ["category", "name", "values", "sizes"];
const ROLE_FIELD = {
  name: "name",
  category: "categories",
  values: "values",
  sizes: "sizes",
} as const satisfies Record<ChartSourceRole, keyof ChartSourceSeriesCells>;

/** Rectangles of one role, those side by side or one above the other along a
 * whole edge drawn as one, one inside another drawn once. */
function mergedRanges(ranges: readonly ChartSourceRange[]): ChartSourceRange[] {
  let current = ranges.filter(
    (range, index) => ranges.findIndex((other) => sameRange(other, range)) === index,
  );
  for (let pass = 0; pass < 4; pass += 1) {
    const before = current.length;
    current = current.filter(
      (range, index) =>
        !current.some(
          (other, otherIndex) =>
            otherIndex !== index &&
            other.r1 <= range.r1 && other.c1 <= range.c1 && other.r2 >= range.r2 && other.c2 >= range.c2 &&
            // Of two equal rectangles the first stays.
            (!sameRange(other, range) || otherIndex < index),
        ),
    );
    for (const horizontal of [true, false]) {
      const sorted = [...current].sort((a, b) =>
        horizontal ? a.r1 - b.r1 || a.r2 - b.r2 || a.c1 - b.c1 : a.c1 - b.c1 || a.c2 - b.c2 || a.r1 - b.r1,
      );
      const out: ChartSourceRange[] = [];
      for (const range of sorted) {
        const last = out[out.length - 1];
        if (
          last &&
          (horizontal
            ? last.r1 === range.r1 && last.r2 === range.r2 && last.c2 + 1 === range.c1
            : last.c1 === range.c1 && last.c2 === range.c2 && last.r2 + 1 === range.r1)
        ) {
          out[out.length - 1] = horizontal ? { ...last, c2: range.c2 } : { ...last, r2: range.r2 };
        } else out.push(range);
      }
      current = out;
    }
    if (current.length === before) break;
  }
  return current;
}

/** The outlines of the chart's data: every series, or the one `only` names. */
export function chartSourceBoxes(
  series: readonly ChartSourceSeriesCells[],
  only: number | null = null,
): ChartSourceBox[] {
  const shown = only === null ? series : series[only] ? [series[only]!] : series;
  const boxes: ChartSourceBox[] = [];
  for (const role of ROLE_ORDER) {
    const field = ROLE_FIELD[role];
    const ranges = shown.flatMap((entry) => (entry[field] ? [entry[field]!] : []));
    for (const range of mergedRanges(ranges)) {
      boxes.push({
        key: `${role}:${range.r1}:${range.c1}:${range.r2}:${range.c2}`,
        role,
        color: CHART_SOURCE_ROLE_COLORS[role],
        range,
      });
    }
  }
  return boxes;
}

/**
 * The series as one layout, in canonical orientation (series in columns):
 * the names in one row above the values, the categories (the same for every
 * series) beside them on the values' rows, each series one column of the
 * value block, in order. Null when the series are not laid out that way.
 */
interface Block {
  readonly seriesInColumns: boolean;
  readonly nameRow?: number;
  readonly labels?: readonly [number, number];
  readonly rows: readonly [number, number];
  readonly cols: readonly [number, number];
}

function blockIn(series: readonly ChartSourceSeriesCells[], seriesInColumns: boolean): Block | null {
  if (!series.length) return null;
  const cells = series.map((entry) => orient(entry, !seriesInColumns));
  const first = cells[0]!;
  if (!first.values) return null;
  const rows = [first.values.r1, first.values.r2] as const;
  const firstCol = first.values.c1;
  let nameRow: number | undefined;
  const named = Boolean(first.name);
  for (const [index, entry] of cells.entries()) {
    const values = entry.values;
    if (!values || entry.sizes) return null;
    if (values.c1 !== firstCol + index || values.c2 !== values.c1) return null;
    if (values.r1 !== rows[0] || values.r2 !== rows[1]) return null;
    if (Boolean(entry.name) !== named) return null;
    if (entry.name) {
      const { r1, c1, r2, c2 } = entry.name;
      if (r1 !== r2 || c1 !== c2 || c1 !== values.c1 || r1 >= rows[0]) return null;
      if (nameRow !== undefined && nameRow !== r1) return null;
      nameRow = r1;
    }
    if (!sameRange(entry.categories, first.categories)) return null;
  }
  const categories = first.categories;
  if (categories && (categories.r1 !== rows[0] || categories.r2 !== rows[1] || categories.c2 >= firstCol)) return null;
  return {
    seriesInColumns,
    ...(nameRow !== undefined ? { nameRow } : {}),
    ...(categories ? { labels: [categories.c1, categories.c2] as const } : {}),
    rows,
    cols: [firstCol, firstCol + cells.length - 1],
  };
}

function blockOf(series: readonly ChartSourceSeriesCells[], seriesInColumns?: boolean): Block | null {
  if (seriesInColumns !== undefined) return blockIn(series, seriesInColumns);
  return blockIn(series, true) ?? blockIn(series, false);
}

function blockSeries(block: Block): ChartSourceSeriesCells[] {
  const out: ChartSourceSeriesCells[] = [];
  for (let col = block.cols[0]; col <= block.cols[1]; col += 1) {
    out.push(
      orient(
        {
          ...(block.nameRow !== undefined ? { name: rect(block.nameRow, col, block.nameRow, col) } : {}),
          ...(block.labels ? { categories: rect(block.rows[0], block.labels[0], block.rows[1], block.labels[1]) } : {}),
          values: rect(block.rows[0], col, block.rows[1], col),
        },
        !block.seriesInColumns,
      ),
    );
  }
  return out;
}

/** A drag on one outline, in sheet cells. */
export type ChartSourceGesture =
  | {
      readonly kind: "resize";
      readonly role: ChartSourceRole;
      /** The outline as it was. */
      readonly box: ChartSourceRange;
      /** The corner handle pressed, and the cell it is dragged to. */
      readonly corner: "nw" | "ne" | "sw" | "se";
      readonly to: { readonly row: number; readonly col: number };
    }
  | {
      readonly kind: "move";
      readonly role: ChartSourceRole;
      readonly box: ChartSourceRange;
      /** How far the outline is dragged, in rows and columns. */
      readonly rows: number;
      readonly cols: number;
    };

/** A block after a drag on one of its outlines (canonical orientation). */
function blockDragged(block: Block, gesture: ChartSourceGesture, oneSeries: boolean): Block | null {
  const flip = !block.seriesInColumns;
  const floorRow = block.nameRow !== undefined ? block.nameRow + 1 : 0;
  const floorCol = block.labels ? block.labels[1] + 1 : 0;
  let [r1, r2] = block.rows;
  let [c1, c2] = block.cols;
  if (gesture.kind === "move") {
    let dr = flip ? gesture.cols : gesture.rows;
    let dc = flip ? gesture.rows : gesture.cols;
    // The names move along their row, the categories along their column.
    if (gesture.role === "name") dr = 0;
    if (gesture.role === "category") dc = 0;
    if (gesture.role === "sizes") return null;
    // Never back over the names or the categories.
    dr = Math.max(dr, floorRow - r1);
    dc = Math.max(dc, floorCol - c1);
    r1 += dr;
    r2 += dr;
    c1 += dc;
    c2 += dc;
  } else {
    const box = flip ? transpose(gesture.box) : gesture.box;
    const corner = flip
      ? ({ nw: "nw", se: "se", ne: "sw", sw: "ne" } as const)[gesture.corner]
      : gesture.corner;
    const to = flip ? { row: gesture.to.col, col: gesture.to.row } : gesture.to;
    // The corner opposite the one dragged stays.
    const fixedRow = corner[0] === "n" ? box.r2 : box.r1;
    const fixedCol = corner[1] === "w" ? box.c2 : box.c1;
    const rows = [Math.min(fixedRow, to.row), Math.max(fixedRow, to.row)] as const;
    const cols = [Math.min(fixedCol, to.col), Math.max(fixedCol, to.col)] as const;
    if (gesture.role === "values") {
      [r1, r2] = rows;
      // One series' values stay one column.
      if (!oneSeries) [c1, c2] = cols;
    } else if (gesture.role === "name") {
      if (oneSeries) return null;
      [c1, c2] = cols;
    } else if (gesture.role === "category") {
      [r1, r2] = rows;
    } else return null;
    r1 = Math.max(r1, floorRow);
    c1 = Math.max(c1, floorCol);
    if (r1 > r2 || c1 > c2) return null;
  }
  if (r1 < 0 || c1 < 0) return null;
  if (r1 === block.rows[0] && r2 === block.rows[1] && c1 === block.cols[0] && c2 === block.cols[1]) return null;
  return { ...block, rows: [r1, r2], cols: [c1, c2] };
}

/** How a drag is written, the way Select Data writes the same edit. */
export type ChartSourceEdit =
  | {
      /** Select Data's Chart data range, laid as the chart lays it now. */
      readonly kind: "range";
      /** Canonical A1 ("D5:F8"). */
      readonly range: string;
      readonly byRows: boolean;
    }
  | {
      /** Legend Entries edited: every series' references (XLS-A072). */
      readonly kind: "refs";
      readonly refs: SpreadsheetChartSeriesRef[];
      /** Where each series was (null: a series the drag added). */
      readonly seriesOrigins: (number | null)[];
    };

function cellRef(prefix: string, row: number, col: number): string {
  return `${prefix}!$${columnToLetters(col)}$${row + 1}`;
}

function areaRef(prefix: string, range: ChartSourceRange): string {
  const first = cellRef(prefix, range.r1, range.c1);
  return range.r1 === range.r2 && range.c1 === range.c2
    ? first
    : `${first}:$${columnToLetters(range.c2)}$${range.r2 + 1}`;
}

/** The one data range a range-bound chart reads `block` from (names right
 * above, categories right beside with their outer columns), or null. */
function blockRange(block: Block, outer: number, xy: boolean, noCategories = false): ChartSourceRange | null {
  if (block.nameRow !== undefined && block.nameRow !== block.rows[0] - 1) return null;
  if (block.labels) {
    if (noCategories || block.labels[1] !== block.cols[0] - 1 || block.labels[1] - block.labels[0] !== outer) return null;
  } else if (xy || (block.cols[0] !== block.cols[1] && !noCategories)) {
    // Without a label column a range plots one column of values.
    return null;
  }
  const canonical = rect(
    block.nameRow ?? block.rows[0],
    block.labels ? block.labels[1] : block.cols[0],
    block.rows[1],
    block.cols[1],
  );
  return block.seriesInColumns ? canonical : transpose(canonical);
}

/** Whether the outlines shown (every series, or the one `only` names) take
 * a drag: the chart's data is editable and they are one layout. */
export function chartSourceEditable(reading: ChartSourceReading, only: number | null = null): boolean {
  if (!reading.edit) return false;
  const target = only !== null && reading.series[only] ? [reading.series[only]!] : reading.series;
  return blockOf(target, reading.edit.seriesInColumns) !== null;
}

/**
 * The chart's data after a drag on one of its outlines, as Select Data
 * writes it, or null when the drag changes nothing or the chart's data
 * cannot take it (the outline then goes back where it was).
 * `only` is the series shown alone (one selected), null for the whole chart.
 */
export function chartSourceDragged(
  reading: ChartSourceReading,
  gesture: ChartSourceGesture,
  only: number | null = null,
): { edit: ChartSourceEdit; series: ChartSourceSeriesCells[] } | null {
  const edit = reading.edit;
  if (!edit) return null;
  const one = only !== null && reading.series[only] ? only : null;
  const target = one === null ? reading.series : [reading.series[one]!];
  const block = blockOf(target, edit.seriesInColumns);
  if (!block) return null;
  const moved = blockDragged(block, gesture, one !== null);
  if (!moved) return null;
  const laid = blockSeries(moved);
  const series = one === null
    ? laid
    : reading.series.map((entry, index) => (index === one ? laid[0]! : entry));
  const count = reading.series.length;
  if (series.length > 255 || (edit.countLocked && series.length !== count)) return null;
  const xy = XY_TYPES.has(reading.type);
  if (edit.via === "range") {
    const whole = blockOf(series, edit.seriesInColumns);
    const range = whole ? blockRange(whole, edit.outerCategoryColumns, xy, edit.noCategories === true) : null;
    if (range) {
      const a1 = chartRangeToA1(range);
      // A range past the chart's grid is not one Select Data takes.
      if (!sameRange(normalizeChartRange(a1) ?? undefined, range)) return null;
      return { edit: { kind: "range", range: a1, byRows: !edit.seriesInColumns }, series };
    }
    if (!edit.convertible) return null;
  }
  const refs: SpreadsheetChartSeriesRef[] = series.map((entry, index) => {
    const extras = edit.extras[index] ?? {};
    return {
      ...(entry.name ? { title: cellRef(edit.prefix, entry.name.r1, entry.name.c1) } : extras.titleText !== undefined ? { titleText: extras.titleText } : {}),
      ...(entry.categories ? { [xy ? "xValues" : "categories"]: areaRef(edit.prefix, entry.categories) } : {}),
      values: areaRef(edit.prefix, entry.values!),
      ...(extras.renderType ? { renderType: extras.renderType } : {}),
      ...(extras.axis ? { axis: extras.axis } : {}),
    };
  });
  const normalized = normalizeChartSeriesRefs(refs);
  if (!normalized || normalized.length !== refs.length) return null;
  return {
    edit: {
      kind: "refs",
      refs: normalized,
      seriesOrigins: series.map((_, index) => (index < count ? index : null)),
    },
    series,
  };
}
