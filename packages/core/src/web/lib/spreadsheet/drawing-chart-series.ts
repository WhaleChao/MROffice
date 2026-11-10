/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Select Data and Switch Row/Column on an imported (drawing) chart.
 *
 * An imported chart states its data as series references, not a rectangle.
 * Excel's Select Data takes one "Chart data range" and lays series over it:
 * with series in columns the first row names each series and the first
 * column holds the categories (an XY chart's X values), every other column
 * is a series; with series in rows the same, transposed. Switch Row/Column
 * re-lays the same range the other way. The rule here is the one the
 * editor's own charts are written with (API `chart_recipe_binding`), so a
 * range plots the same series whichever chart it is given to.
 *
 * The original-package writer then rewrites the chart's series in place,
 * each keeping its formatting (API `chart_series_rewrite`).
 */

import { quotedSheetName } from "./change-log";
import {
  nativeChartRangeSeriesRefs,
  nativeChartStatesLabelLines,
  parseChartRefRange,
  type ChartSourceRange,
  type SpreadsheetChartSeriesRef,
  type SpreadsheetChartType,
} from "./chart";
import { columnToLetters } from "./formula-engine";

/** The chart types whose series this rewrite lays out (classic, not
 * Office 2016; a bubble chart's sizes have no place in one range). */
const DRAWING_CHART_LAYOUT_TYPES = new Set<SpreadsheetChartType>([
  "column",
  "bar",
  "stacked-column",
  "stacked-bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "radar",
  "scatter",
  "surface",
]);

const XY_TYPES = new Set<SpreadsheetChartType>(["scatter"]);

export interface DrawingChartSource {
  /** The sheet as the chart's own references spell it (`Sheet1`, `'Q1 data'`). */
  prefix: string;
  sheetName: string;
  range: ChartSourceRange;
  seriesInColumns: boolean;
}

export function drawingChartTakesDataRange(type: SpreadsheetChartType): boolean {
  return DRAWING_CHART_LAYOUT_TYPES.has(type);
}

function cell(prefix: string, row: number, col: number): string {
  return `${prefix}!$${columnToLetters(col)}$${row + 1}`;
}

function area(prefix: string, r1: number, c1: number, r2: number, c2: number): string {
  return `${cell(prefix, r1, c1)}:$${columnToLetters(c2)}$${r2 + 1}`;
}

/** The series one range plots for ``type``, or null when it plots none (it
 * needs a row of names and a column of labels besides the values). */
export function drawingChartSeriesRefsForRange(
  prefix: string,
  range: ChartSourceRange,
  seriesInColumns: boolean,
  type: SpreadsheetChartType,
): SpreadsheetChartSeriesRef[] | null {
  const { r1, c1, r2, c2 } = range;
  if (!drawingChartTakesDataRange(type) || r2 <= r1 || c2 <= c1) return null;
  const xy = XY_TYPES.has(type);
  const refs: SpreadsheetChartSeriesRef[] = [];
  if (seriesInColumns) {
    const labels = area(prefix, r1 + 1, c1, r2, c1);
    for (let col = c1 + 1; col <= c2; col += 1) {
      refs.push({
        title: cell(prefix, r1, col),
        ...(xy ? { xValues: labels } : { categories: labels }),
        values: area(prefix, r1 + 1, col, r2, col),
      });
    }
  } else {
    const labels = area(prefix, r1, c1 + 1, r1, c2);
    for (let row = r1 + 1; row <= r2; row += 1) {
      refs.push({
        title: cell(prefix, row, c1),
        ...(xy ? { xValues: labels } : { categories: labels }),
        values: area(prefix, row, c1 + 1, row, c2),
      });
    }
  }
  return refs;
}

/**
 * The series a chart born here lays over its range, laid `seriesInColumns`
 * (its own way, or the other for a switched preview): its Legend Entries
 * as Select Data lists them. A range that states its label lines (no
 * category column, no row of names) lays them as the workbook writer does;
 * any other as `drawingChartSeriesRefsForRange`.
 */
export function nativeChartSeriesRefsForRange(
  prefix: string,
  chart: { range: ChartSourceRange; seriesInColumns?: boolean; noCategories?: true; noSeriesNames?: true },
  seriesInColumns: boolean,
  type: SpreadsheetChartType,
): SpreadsheetChartSeriesRef[] | null {
  if (nativeChartStatesLabelLines(chart) && !XY_TYPES.has(type) && drawingChartTakesDataRange(type)) {
    // Laid the other way, the line that named the series labels the categories.
    const flipped = (chart.seriesInColumns !== false) !== seriesInColumns;
    return nativeChartRangeSeriesRefs(
      {
        type,
        range: chart.range,
        ...(seriesInColumns ? {} : { seriesInColumns: false }),
        ...((flipped ? chart.noSeriesNames : chart.noCategories) ? { noCategories: true as const } : {}),
        ...((flipped ? chart.noCategories : chart.noSeriesNames) ? { noSeriesNames: true as const } : {}),
      },
      prefix,
    );
  }
  return drawingChartSeriesRefsForRange(prefix, chart.range, seriesInColumns, type);
}

/** The formulas a chart part states, in its order, each once (the import's
 * `ranges`): per series its name, labels, values. */
export function drawingChartRangesOf(refs: readonly SpreadsheetChartSeriesRef[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const ref of refs) {
    for (const formula of [ref.title, ref.categories, ref.xValues, ref.values, ref.bubbleSizes]) {
      if (formula && !seen.has(formula)) {
        seen.add(formula);
        out.push(formula);
      }
    }
  }
  return out;
}

function sameRef(left: string | undefined, right: string | undefined): boolean {
  if (left === undefined || right === undefined) return left === right;
  const a = parseChartRefRange(left);
  const b = parseChartRefRange(right);
  return Boolean(
    a &&
      b &&
      (a.sheetName ?? "").toLocaleLowerCase() === (b.sheetName ?? "").toLocaleLowerCase() &&
      a.range.r1 === b.range.r1 &&
      a.range.c1 === b.range.c1 &&
      a.range.r2 === b.range.r2 &&
      a.range.c2 === b.range.c2,
  );
}

function refsEqual(
  left: readonly SpreadsheetChartSeriesRef[],
  right: readonly SpreadsheetChartSeriesRef[],
): boolean {
  return (
    left.length === right.length &&
    left.every((ref, index) => {
      const other = right[index];
      return (
        ref.titleText === undefined &&
        other.titleText === undefined &&
        sameRef(ref.title, other.title) &&
        sameRef(ref.categories, other.categories) &&
        sameRef(ref.xValues, other.xValues) &&
        sameRef(ref.values, other.values) &&
        ref.bubbleSizes === undefined &&
        other.bubbleSizes === undefined
      );
    })
  );
}

/** The sheet a chart's references all name, as they spell it. */
export function drawingChartSheetOf(
  refs: readonly SpreadsheetChartSeriesRef[],
): { prefix: string; sheetName: string } | null {
  let found: { prefix: string; sheetName: string } | null = null;
  for (const ref of refs) {
    for (const formula of [ref.title, ref.categories, ref.xValues, ref.values, ref.bubbleSizes]) {
      if (!formula) continue;
      const parsed = parseChartRefRange(formula);
      const bang = formula.lastIndexOf("!");
      if (!parsed?.sheetName || bang <= 0) return null;
      if (found && found.sheetName.toLocaleLowerCase() !== parsed.sheetName.toLocaleLowerCase()) return null;
      found ??= { prefix: formula.slice(0, bang), sheetName: parsed.sheetName };
    }
  }
  return found;
}

/**
 * The one range (and orientation) whose layout is exactly these series, or
 * null when the chart's data is not one range laid out the usual way (a
 * series from elsewhere, a literal name, a missing row of names) -- the
 * chart Excel calls "too complex" to show in Select Data.
 */
export function drawingChartSourceOf(
  refs: readonly SpreadsheetChartSeriesRef[] | undefined,
  type: SpreadsheetChartType,
): DrawingChartSource | null {
  if (!refs?.length || !drawingChartTakesDataRange(type)) return null;
  const sheet = drawingChartSheetOf(refs);
  if (!sheet) return null;
  let rect: ChartSourceRange | null = null;
  for (const ref of refs) {
    for (const formula of [ref.title, ref.categories, ref.xValues, ref.values]) {
      const parsed = formula ? parseChartRefRange(formula) : null;
      if (!parsed) continue;
      const { r1, c1, r2, c2 } = parsed.range;
      rect = rect
        ? { r1: Math.min(rect.r1, r1), c1: Math.min(rect.c1, c1), r2: Math.max(rect.r2, r2), c2: Math.max(rect.c2, c2) }
        : { r1, c1, r2, c2 };
    }
  }
  if (!rect) return null;
  for (const seriesInColumns of [true, false]) {
    const laid = drawingChartSeriesRefsForRange(sheet.prefix, rect, seriesInColumns, type);
    if (laid && refsEqual(laid, refs)) {
      return { ...sheet, range: rect, seriesInColumns };
    }
  }
  return null;
}

/** Which way a range is laid when the chart says nothing (its series are not
 * one range): Excel puts the longer side on the category axis, so a range
 * with more rows than columns plots its columns as series. */
export function drawingChartDefaultSeriesInColumns(
  refs: readonly SpreadsheetChartSeriesRef[] | undefined,
  range: ChartSourceRange,
): boolean {
  const vectors = (refs ?? [])
    .map((ref) => parseChartRefRange(ref.values)?.range)
    .filter((value): value is ChartSourceRange => Boolean(value));
  if (vectors.length && vectors.every((value) => value.c1 === value.c2 && value.r1 !== value.r2)) return true;
  if (vectors.length && vectors.every((value) => value.r1 === value.r2 && value.c1 !== value.c2)) return false;
  return range.r2 - range.r1 >= range.c2 - range.c1;
}

/** A sheet name as a chart reference spells it. */
export function drawingChartSheetPrefix(sheetName: string): string {
  return quotedSheetName(sheetName);
}
