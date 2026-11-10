/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure projection of a PivotTable result into chart data (Excel "PivotChart"):
 * the pivot's row groups become the category axis, and each output column (a
 * column-field value × an aggregated value field) becomes a series. The chart is
 * a READ-ONLY view bound to the pivot — it is re-derived whenever the pivot
 * recomputes, so it never drifts. Dependency-free → unit-testable apart from the
 * panel + SVG rendering.
 */

import { isPercentShowValuesAs, type PivotConfig, type PivotResult } from "./pivot";
import type { PivotChartReport } from "./pivot-output";
import { normalizeSpreadsheetPivotChartSource, type SpreadsheetChartType, type SpreadsheetPivotChartSource } from "./chart";

/**
 * The chart kinds a PivotChart can be. Microsoft Support ("Overview of
 * PivotTables and PivotCharts"): a PivotChart can be changed to any chart
 * type except an xy (scatter), stock or bubble chart; the Office 2016 kinds
 * (treemap, sunburst, histogram, box and whisker, waterfall, funnel, map)
 * cannot be made from a PivotTable either. Everything else stays.
 */
export const SPREADSHEET_PIVOT_CHART_TYPES = Object.freeze([
  "column",
  "stacked-column",
  "bar",
  "stacked-bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "radar",
  "surface",
  "combo",
] as const satisfies readonly SpreadsheetChartType[]);

/** The kinds Insert PivotChart offers first (its gallery's families). */
export const SPREADSHEET_PIVOT_CHART_INSERT_TYPES = Object.freeze([
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "doughnut",
  "radar",
] as const satisfies readonly SpreadsheetChartType[]);

const SPREADSHEET_PIVOT_CHART_TYPE_SET: ReadonlySet<string> = new Set(
  SPREADSHEET_PIVOT_CHART_TYPES,
);

export type SpreadsheetPivotChartType =
  (typeof SPREADSHEET_PIVOT_CHART_TYPES)[number];

export function isSpreadsheetPivotChartType(
  value: unknown,
): value is SpreadsheetPivotChartType {
  return (
    typeof value === "string" && SPREADSHEET_PIVOT_CHART_TYPE_SET.has(value)
  );
}

export interface PivotChartSeries {
  name: string;
  /** null is a gap: the report shows nothing there. */
  values: (number | null)[];
}

export interface PivotChartData {
  /** One label per pivot row group (the category axis). */
  categories: string[];
  /** The outer axis fields (`pivotChartReportCategoryLevels`), when there are any. */
  categoryLevels?: string[][];
  /** One series per pivot output column. */
  series: PivotChartSeries[];
}

/**
 * Project a {@link PivotResult} into category + series chart data. `valueLabels`
 * are the aggregated value-field display names (config.values, in order) used to
 * label a column when the pivot has no column field, or to qualify a series when
 * there are multiple value fields. A null/missing cell reads as 0 (a chart needs
 * a number). Returns empty arrays for an empty pivot.
 */
export function pivotResultToChartData(
  result: PivotResult,
  valueLabels: readonly string[] = [],
): PivotChartData {
  const categories = result.rowHeaders.map(
    (path) => path.filter((p) => p !== "").join(" / ") || "—",
  );
  const series: PivotChartSeries[] = result.colHeaders.map((col, columnIndex) => {
    const valueLabel =
      valueLabels[col.valueIndex] ?? `Value ${col.valueIndex + 1}`;
    const colLabel = col.columnValues.filter((v) => v !== "").join(" / ");
    // No column field → label by the value field. Column field + multiple value
    // fields → qualify (e.g. "Q1 · Sum of Sales"). Single value field → the
    // column label alone is unambiguous.
    let name: string;
    if (!colLabel) {
      name = valueLabel;
    } else if (valueLabels.length > 1) {
      name = `${colLabel} · ${valueLabel}`;
    } else {
      name = colLabel;
    }
    const values = result.cells.map((row) => {
      const v = row[columnIndex];
      return typeof v === "number" && Number.isFinite(v) ? v : 0;
    });
    return { name, values };
  });
  return { categories, series };
}

/** Whether a pivot result has anything chartable (≥1 category and ≥1 series). */
export function hasPivotChartData(data: PivotChartData): boolean {
  return data.categories.length > 0 && data.series.length > 0;
}

/** Convert the live Pivot projection to the row matrix MiniChart consumes. */
export function pivotChartDataToRows(data: PivotChartData): string[][] {
  return data.categories.map((category, rowIndex) => [
    category,
    ...data.series.map((series) => {
      const value = series.values[rowIndex];
      return value === null || value === undefined ? "" : String(value);
    }),
  ]);
}

/**
 * The row matrix a sheet chart reads from its PivotTable's report: a header
 * row of series names, then one row per category ([label, ...values]). A
 * value the report leaves empty is "", a gap, as a PivotChart shows it
 * (`c:dispBlanksAs val="gap"`).
 */
export function pivotChartReportToRows(report: PivotChartReport): string[][] {
  if (report.series.length === 0 || report.categories.length === 0) return [];
  // With outer axis fields the row label is the innermost level only (the
  // outer ones are `pivotChartReportCategoryLevels`); a collapsed item has
  // no inner label.
  const depth = pivotChartReportAxisDepth(report);
  return [
    ["", ...report.series.map((series) => series.name)],
    ...report.categories.map((category, index) => [
      depth > 1
        ? report.categoryLabels[index]?.length === depth
          ? category
          : ""
        : category,
      ...report.series.map((series) => {
        const value = series.values[index];
        return value === null || value === undefined ? "" : String(value);
      }),
    ]),
  ];
}

/**
 * How many levels a PivotChart's axis shows: its axis fields, less the inner
 * ones every item above them hides (a whole field collapsed shows its outer
 * field alone, as Excel's Collapse Entire Field does).
 */
export function pivotChartReportAxisDepth(report: PivotChartReport): number {
  if (report.rowFields.length <= 1) return report.rowFields.length;
  let depth = 0;
  for (const path of report.categoryPaths) depth = Math.max(depth, path.length);
  return Math.min(report.rowFields.length, Math.max(1, depth));
}

/**
 * A PivotChart's multi-level category axis (the shared model's
 * `categoryLevels`): the axis fields above the innermost, outermost first,
 * one entry per category, an item's label on the first category of its group
 * and "" after it. Null for an axis of one field.
 */
export function pivotChartReportCategoryLevels(
  report: PivotChartReport,
): string[][] | null {
  const depth = pivotChartReportAxisDepth(report);
  if (depth < 2 || report.categories.length === 0) return null;
  const levels: string[][] = Array.from({ length: depth - 1 }, () => []);
  report.categoryPaths.forEach((path, index) => {
    const previous = index > 0 ? report.categoryPaths[index - 1] : null;
    for (let level = 0; level < depth - 1; level += 1) {
      const starts =
        level < path.length &&
        (!previous ||
          previous.length <= level ||
          path.slice(0, level + 1).some((item, i) => item !== previous[i]));
      levels[level].push(starts ? (report.categoryLabels[index]?.[level] ?? "") : "");
    }
  });
  return levels;
}

/** A PivotTable's name: the recipe's own, else the imported table's. */
export function pivotTableName(
  pivot: Pick<PivotConfig, "name"> & { native?: { pivotName?: string } },
): string | undefined {
  const name = pivot.name ?? pivot.native?.pivotName;
  return typeof name === "string" && name.trim() ? name.trim() : undefined;
}

type NamedPivot = Pick<PivotConfig, "name"> & { native?: { pivotName?: string } };
type PivotHost<P> = { id: string; pivot?: P; morePivots?: readonly P[] };
/** The PivotTable type a sheet holds. */
type PivotOf<S> = S extends { pivot?: infer P } ? NonNullable<P> : never;

function hostPivots<P>(sheet: PivotHost<P>): P[] {
  return [
    ...(sheet.pivot !== undefined ? [sheet.pivot] : []),
    ...(sheet.morePivots ?? []),
  ];
}

/**
 * The number format a PivotChart's value axis, data labels and data table
 * are linked to, as Excel links a PivotChart's to its PivotTable: the first
 * value field's own number format (Value Field Settings > Number Format); a
 * value shown as a percentage without one takes Excel's 0.00%; else General.
 * The report's cells are never stored cells, so the sheet chart's source-
 * linked reading of its range found no format and drew 0.2758 for 27.58%.
 */
export function pivotChartLinkedNumberFormat(
  pivot: Pick<PivotConfig, "values"> | undefined,
): string {
  const field = pivot?.values?.[0];
  const own = typeof field?.numberFormat === "string" ? field.numberFormat.trim() : "";
  if (own) return own;
  if (field?.showAs && isPercentShowValuesAs(field.showAs.mode)) return "0.00%";
  return "General";
}

/**
 * The PivotTable a PivotChart is bound to: on the sheet it names, the table
 * of its name (names compare as Excel compares them, case-insensitively), or
 * that sheet's first table for a record written before charts named theirs.
 */
export function resolvePivotChartPivot<S extends PivotHost<NamedPivot>>(
  sheets: readonly S[],
  source: SpreadsheetPivotChartSource | undefined,
): { sheet: S; sheetIndex: number; slot: number; pivot: PivotOf<S> } | null {
  if (!source) return null;
  const sheetIndex = sheets.findIndex((candidate) => candidate.id === source.sheetId);
  if (sheetIndex < 0) return null;
  const sheet = sheets[sheetIndex];
  const pivots = hostPivots(sheet) as PivotOf<S>[];
  if (pivots.length === 0) return null;
  if (!source.pivotName) return { sheet, sheetIndex, slot: 0, pivot: pivots[0] };
  const wanted = source.pivotName.toLowerCase();
  const slot = pivots.findIndex(
    (pivot) => pivotTableName(pivot as NamedPivot)?.toLowerCase() === wanted,
  );
  return slot < 0 ? null : { sheet, sheetIndex, slot, pivot: pivots[slot] };
}

/**
 * The name Excel gives a new PivotTable: "PivotTable" and the next number
 * after every one the workbook already uses.
 */
export function nextPivotTableName(sheets: readonly PivotHost<NamedPivot>[]): string {
  let highest = 0;
  for (const sheet of sheets) {
    for (const pivot of hostPivots(sheet)) {
      const match = /^pivottable(\d{1,6})$/i.exec(pivotTableName(pivot) ?? "");
      if (match) highest = Math.max(highest, Number(match[1]));
    }
  }
  return `PivotTable${highest + 1}`;
}

/** The PivotChart's field buttons (Excel's PivotChart Analyze > Field Buttons). */
export interface SpreadsheetPivotChartOptions {
  /** Hide All clears it; no button shows while it is false. */
  visible?: boolean;
  /** Report Filter field buttons. */
  filter?: boolean;
  /** Legend (Series) field buttons. */
  legend?: boolean;
  /** Axis (Categories) field buttons. */
  axis?: boolean;
  /** Value field buttons. */
  values?: boolean;
  /** Expand/Collapse Entire Field buttons. */
  expandCollapse?: boolean;
}

export type SpreadsheetPivotChartButtonKind = Exclude<
  keyof SpreadsheetPivotChartOptions,
  "visible"
>;

export const SPREADSHEET_PIVOT_CHART_BUTTON_KINDS: readonly SpreadsheetPivotChartButtonKind[] =
  Object.freeze(["filter", "legend", "axis", "values", "expandCollapse"]);

/** Only the booleans; an absent key keeps the default (shown). */
export function normalizePivotChartOptions(raw: unknown): SpreadsheetPivotChartOptions | undefined {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
  const record = raw as Record<string, unknown>;
  const out: SpreadsheetPivotChartOptions = {};
  for (const key of ["visible", ...SPREADSHEET_PIVOT_CHART_BUTTON_KINDS] as const) {
    if (typeof record[key] === "boolean") out[key] = record[key] as boolean;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** Whether a PivotChart shows one kind of field button (Excel's default: all). */
export function pivotChartShowsButtons(
  options: SpreadsheetPivotChartOptions | undefined,
  kind: SpreadsheetPivotChartButtonKind,
): boolean {
  return options?.visible !== false && options?.[kind] !== false;
}

type ChartHost<P> = PivotHost<P> & {
  charts?: readonly { pivotSource?: SpreadsheetPivotChartSource }[];
  objects?: readonly { type?: unknown; pivotSource?: unknown }[];
};

/**
 * Every PivotChart of a workbook re-pointed after its PivotTable moved or was
 * renamed: `rebind` gets each chart's current PivotTable (resolved in
 * `sheets`, the workbook BEFORE the change) and answers its new binding, or
 * undefined to keep it. The result is laid over `nextSheets` (the workbook
 * after the change, same sheet order); a sheet none of whose charts moved
 * keeps its identity.
 */
export function rebindPivotCharts<S extends ChartHost<NamedPivot>>(
  sheets: readonly S[],
  nextSheets: readonly S[],
  rebind: (
    bound: { sheetId: string; slot: number; pivot: PivotOf<S> },
    source: SpreadsheetPivotChartSource,
  ) => SpreadsheetPivotChartSource | undefined,
): S[] {
  const target = (raw: unknown): SpreadsheetPivotChartSource | undefined => {
    if (!raw || typeof raw !== "object") return undefined;
    const source = raw as SpreadsheetPivotChartSource;
    if (typeof source.sheetId !== "string") return undefined;
    const bound = resolvePivotChartPivot(sheets, source);
    if (!bound) return undefined;
    const next = rebind({ sheetId: bound.sheet.id, slot: bound.slot, pivot: bound.pivot }, source);
    return next && (next.sheetId !== source.sheetId || next.pivotName !== source.pivotName)
      ? next
      : undefined;
  };
  return nextSheets.map((sheet) => {
    let touched = false;
    const charts = sheet.charts?.map((chart) => {
      const next = target(chart.pivotSource);
      if (!next) return chart;
      touched = true;
      return { ...chart, pivotSource: next };
    });
    const objects = sheet.objects?.map((object) => {
      if (object.type !== "chart") return object;
      const next = target(object.pivotSource);
      if (!next) return object;
      touched = true;
      return { ...object, pivotSource: next };
    });
    return touched
      ? ({ ...sheet, ...(charts ? { charts } : {}), ...(objects ? { objects } : {}) } as S)
      : sheet;
  });
}

const PIVOT_CHART_REFRESH_LOSSES = [
  "trendline",
  "trendlineStyle",
  "seriesTrendlines",
  "errorBars",
  "errorBarStyle",
  "errorBarsX",
  "errorBarXStyle",
] as const;

/**
 * A chart record without the trendlines and error bars a refresh of its
 * PivotTable takes off (Microsoft Support, "Overview of PivotTables and
 * PivotCharts": a refresh keeps the chart's elements, layout and style, not
 * those). The shared chart model keeps them in its `format` block (Chart
 * Elements ▸ Trendline / Error Bars); older records at its top level. The
 * same record when it carried none.
 */
export function pivotChartWithoutRefreshLosses<T extends object>(record: T): T {
  const source = record as Record<string, unknown>;
  let next = source;
  for (const key of PIVOT_CHART_REFRESH_LOSSES) {
    if (next[key] !== undefined) {
      if (next === source) next = { ...source };
      delete next[key];
    }
  }
  const model = source.chartV2;
  if (model && typeof model === "object" && !Array.isArray(model)) {
    let attrs = model as Record<string, unknown>;
    for (const key of PIVOT_CHART_REFRESH_LOSSES) {
      if (attrs[key] !== undefined) {
        if (attrs === model) attrs = { ...attrs };
        delete attrs[key];
      }
    }
    const format = attrs.format;
    if (format && typeof format === "object" && !Array.isArray(format)) {
      let nextFormat = format as Record<string, unknown>;
      for (const key of PIVOT_CHART_REFRESH_LOSSES) {
        if (nextFormat[key] !== undefined) {
          if (nextFormat === format) nextFormat = { ...nextFormat };
          delete nextFormat[key];
        }
      }
      if (nextFormat !== format) {
        if (attrs === model) attrs = { ...attrs };
        attrs.format = nextFormat;
      }
    }
    if (attrs !== model) {
      if (next === source) next = { ...source };
      next.chartV2 = attrs;
    }
  }
  return next as T;
}

/**
 * The workbook after a refresh of `pivotSheetId`'s PivotTables: every chart
 * bound to one of them (on any sheet) without its trendlines and error bars
 * ({@link pivotChartWithoutRefreshLosses}). The same array when no chart
 * carried any.
 */
export function pivotChartsRefreshedWithoutAnalysis<
  S extends PivotHost<NamedPivot> & { charts?: unknown; objects?: unknown },
>(sheets: readonly S[], pivotSheetId: string): S[] {
  let changed = false;
  const strip = (record: Record<string, unknown>): Record<string, unknown> => {
    const source = normalizeSpreadsheetPivotChartSource(record.pivotSource);
    if (!source || source.sheetId !== pivotSheetId) return record;
    if (!resolvePivotChartPivot(sheets, source)) return record;
    const next = pivotChartWithoutRefreshLosses(record);
    if (next !== record) changed = true;
    return next;
  };
  const stripAll = (list: unknown): unknown => {
    if (!Array.isArray(list)) return list;
    let out: unknown[] | null = null;
    list.forEach((entry, index) => {
      if (!entry || typeof entry !== "object") return;
      const record = entry as Record<string, unknown>;
      if (record.type !== undefined && record.type !== "chart" && !("pivotSource" in record)) return;
      const next = strip(record);
      if (next !== record) {
        out ??= [...list];
        out[index] = next;
      }
    });
    return out ?? list;
  };
  const result = sheets.map((candidate) => {
    const charts = stripAll(candidate.charts);
    const objects = stripAll(candidate.objects);
    if (charts === candidate.charts && objects === candidate.objects) return candidate;
    return {
      ...candidate,
      ...(charts !== candidate.charts ? { charts } : {}),
      ...(objects !== candidate.objects ? { objects } : {}),
    };
  });
  return changed ? result : (sheets as S[]);
}

