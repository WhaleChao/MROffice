/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Chart Design > Select Data: Excel's Select Data Source dialog over a sheet
 * chart (XLS-A072), as pure functions the dialog and its tests share.
 *
 * Excel's dialog (Microsoft Support, "Change the data series in a chart",
 * "Display empty cells, null (#N/A) values, and hidden worksheet data in a
 * chart") states the chart's data three ways at once: one Chart data range
 * (blank when the series are not one range, "too complex"), the Legend
 * Entries (Series) list -- each series a name (a cell or text) and values (a
 * row or column of cells), added, edited, removed and reordered there -- and
 * the Horizontal (Category) Axis Labels every series shares. A range typed in
 * the box replaces every series with the ones it lays; Switch Row/Column lays
 * the same range the other way.
 *
 * How a chart states its series decides how those edits are written:
 * - a chart born in this editor (`sheet.charts`, "native") stores one range
 *   and its orientation. Its Legend Entries are the series that range lays,
 *   edited like any other; while the edited series are still one range the
 *   chart keeps that range, and once they are not (a series removed from
 *   the middle, moved, or read from elsewhere) the chart becomes one that
 *   states its series one by one -- a chart drawing object, as the same
 *   chart is after an XLSX round trip (`nativeChartAsSeriesBoundObject`);
 * - a chart that states its series (`sheet.objects`, "drawing") stores its
 *   series references (`seriesRefs`). The chart's format follows its
 *   series through a removal or a move (`chartFormatFollowingDataEdit`),
 *   as Excel's formatting stays with the series. A chart still bound to
 *   its imported part has that part's series rewritten in place: a removed
 *   or moved series takes its own `c:ser` (everything the model does not
 *   state of it) along, `c:order` renumbered (`slotBound`, XLS-A072).
 */

import { quotedSheetName } from "@mrnote/core/web/lib/spreadsheet/change-log.ts";
import {
  normalizeChartSeriesRefs,
  chartRefAreaTexts,
  parseChartRefAreas,
  parseChartRefRange,
  type ChartSourceRange,
  type SpreadsheetChartSeriesRef,
  type SpreadsheetChartType,
} from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { drawingChartSourceOf, drawingChartTakesDataRange } from "@mrnote/core/web/lib/spreadsheet/drawing-chart-series.ts";
import { columnToLetters } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

/** One Legend Entries (Series) row as the Edit Series boxes show it. */
export interface SelectDataSeriesRow {
  /** Series name: "=Sheet1!$B$1", the name itself, or "" (Series N). */
  readonly name: string;
  /** Series values (an XY chart's Y values): "=Sheet1!$B$2:$B$6". */
  readonly values: string;
  /** An XY chart's Series X values. */
  readonly xValues?: string;
  /** A combination's per-series chart type and axis ride with the row. */
  readonly renderType?: SpreadsheetChartSeriesRef["renderType"];
  readonly axis?: SpreadsheetChartSeriesRef["axis"];
  /** The position the series had when the dialog opened; null (or absent on
   * a row the range relaid) for a series the dialog added. */
  readonly origin?: number | null;
}

/**
 * Why the series list cannot be edited one by one, or null when it can:
 * "derived" -- the chart stores a range, its series are what the range lays;
 * "references" -- its series read through names, literal values or a kind
 * the series rewrite does not take.
 */
export type SelectDataSeriesLock = "derived" | "references" | null;

export interface SelectDataSeriesAbilities {
  readonly lock: SelectDataSeriesLock;
  /** A combination keeps its number of series (each states its own type). */
  readonly countLocked: boolean;
  /** An XY chart: a series has X values, the axis has no labels list. */
  readonly xy: boolean;
  /** The chart's imported part is rewritten in place: a removed or moved
   * series takes its own c:ser along (XLS-A072, `chart_series_rewrite`). */
  readonly slotBound: boolean;
}

export function selectDataSeriesAbilities(input: {
  kind: "native" | "drawing";
  type: SpreadsheetChartType;
  seriesRefs: unknown;
  /** The chart still has the part it was imported with. */
  sourcePart?: boolean;
  /** A native combination (each series states its own type). */
  combination?: boolean;
}): SelectDataSeriesAbilities {
  const xy = input.type === "scatter";
  if (input.kind === "native") {
    // Its series are what its range lays: listed and edited as references
    // over that range. An XY chart reads its range as X / Y columns, and a
    // kind the series writer does not lay keeps the range box only.
    const editable = !xy && drawingChartTakesDataRange(input.type) && input.seriesRefs === undefined;
    return {
      lock: editable ? null : "derived",
      countLocked: !editable || input.combination === true,
      xy,
      slotBound: false,
    };
  }
  const refs = normalizeChartSeriesRefs(input.seriesRefs);
  if (
    !refs ||
    !drawingChartTakesDataRange(input.type) ||
    refs.some((ref) => ref.names || ref.implicitX || ref.bubbleSizes)
  ) {
    return { lock: "references", countLocked: true, xy, slotBound: input.sourcePart === true };
  }
  return {
    lock: null,
    countLocked: refs.some((ref) => ref.renderType !== undefined || ref.axis !== undefined),
    xy,
    slotBound: input.sourcePart === true,
  };
}

/** `seriesOrigins` for `chartFormatFollowingDataEdit`: where each row was. */
export function selectDataSeriesOrigins(rows: readonly SelectDataSeriesRow[]): (number | null)[] {
  return rows.map((row) => (typeof row.origin === "number" ? row.origin : null));
}

/** The dialog's rows for a chart's series references. */
export function selectDataRowsFromRefs(
  refs: readonly SpreadsheetChartSeriesRef[],
): SelectDataSeriesRow[] {
  return refs.map((ref, index) => ({
    name: ref.title ? `=${ref.title}` : (ref.titleText ?? ""),
    values: `=${ref.values}`,
    ...(ref.xValues ? { xValues: `=${ref.xValues}` } : {}),
    ...(ref.renderType ? { renderType: ref.renderType } : {}),
    ...(ref.axis ? { axis: ref.axis } : {}),
    origin: index,
  }));
}

/** The Axis label range the series share ("" when they state none). */
export function selectDataCategoriesOf(
  refs: readonly SpreadsheetChartSeriesRef[],
): string {
  const first = refs.find((ref) => ref.categories)?.categories;
  return first ? `=${first}` : "";
}

/** The Chart data range box for these series: the one range they are laid
 * over and which way, or null (Excel's "too complex" range). */
export function selectDataRangeOfRefs(
  refs: readonly SpreadsheetChartSeriesRef[],
  type: SpreadsheetChartType,
): { range: ChartSourceRange; byRows: boolean } | null {
  const source = drawingChartSourceOf(refs, type);
  return source ? { range: source.range, byRows: !source.seriesInColumns } : null;
}

/**
 * A reference box's text as a chart reference: `=Sheet1!B2:B6`,
 * `Sheet1!$B$2:$B$6` or `B2:B6` (on the chart's sheet, `defaultPrefix`) ->
 * `Sheet1!$B$2:$B$6`. Null when it is not one rectangle of cells of a sheet
 * of the workbook, or -- `shape` -- not one row or column ("vector") or not
 * one cell ("cell").
 */
export function selectDataReference(
  text: string,
  defaultPrefix: string,
  sheetExists: (name: string) => boolean,
  shape: "vector" | "cell",
): string | null {
  const body = text.trim().replace(/^=/, "").trim();
  if (shape === "vector" && body.includes(",")) {
    // Excel's union of areas not next to each other, each a row or column
    // (typed with or without its brackets).
    const areas = chartRefAreaTexts(body.startsWith("(") ? body : `(${body})`);
    const spelled = areas?.map((area) => selectDataReference(area, defaultPrefix, sheetExists, "vector"));
    if (!spelled?.length || spelled.some((area) => area === null)) return null;
    return spelled.length === 1 ? spelled[0] : `(${spelled.join(",")})`;
  }
  const parsed = parseChartRefRange(body);
  if (!parsed) return null;
  const { r1, c1, r2, c2 } = parsed.range;
  if (shape === "cell" ? r1 !== r2 || c1 !== c2 : r1 !== r2 && c1 !== c2) return null;
  if (parsed.sheetName !== undefined && !sheetExists(parsed.sheetName)) return null;
  const prefix = parsed.sheetName !== undefined ? quotedSheetName(parsed.sheetName) : defaultPrefix;
  const first = `$${columnToLetters(c1)}$${r1 + 1}`;
  const last = `$${columnToLetters(c2)}$${r2 + 1}`;
  return `${prefix}!${first}${first === last ? "" : `:${last}`}`;
}

/** The Series name box: a reference to one cell, or the name itself (Excel
 * takes `="Sales"` or plain text as the name); null when a reference does
 * not read. An empty box names no series (Excel's "Series N"). */
export function selectDataSeriesName(
  text: string,
  defaultPrefix: string,
  sheetExists: (name: string) => boolean,
): Pick<SpreadsheetChartSeriesRef, "title" | "titleText"> | null {
  const trimmed = text.trim();
  if (!trimmed) return {};
  if (!trimmed.startsWith("=")) return { titleText: trimmed.slice(0, 255) };
  const quoted = /^=\s*"((?:[^"]|"")*)"\s*$/.exec(trimmed);
  if (quoted) return { titleText: quoted[1].replace(/""/g, '"').slice(0, 255) };
  const title = selectDataReference(trimmed, defaultPrefix, sheetExists, "cell");
  return title ? { title } : null;
}

/**
 * The series references the dialog's rows state, every series taking the
 * shared Axis label range (a category chart's), or the index of the first
 * row that does not read (its name or its values), as `{ invalid }`.
 */
export function selectDataRefsFromRows(
  rows: readonly SelectDataSeriesRow[],
  categories: string,
  options: {
    xy: boolean;
    defaultPrefix: string;
    sheetExists: (name: string) => boolean;
  },
): { refs: SpreadsheetChartSeriesRef[] } | { invalid: number | "categories" } {
  const { xy, defaultPrefix, sheetExists } = options;
  let shared: string | undefined;
  if (!xy && categories.trim()) {
    const parsed = selectDataReference(categories, defaultPrefix, sheetExists, "vector");
    if (!parsed) return { invalid: "categories" };
    shared = parsed;
  }
  const refs: SpreadsheetChartSeriesRef[] = [];
  for (const [index, row] of rows.entries()) {
    const name = selectDataSeriesName(row.name, defaultPrefix, sheetExists);
    const values = selectDataReference(row.values, defaultPrefix, sheetExists, "vector");
    const xValues = xy && row.xValues?.trim()
      ? selectDataReference(row.xValues, defaultPrefix, sheetExists, "vector")
      : undefined;
    if (!name || !values || xValues === null) return { invalid: index };
    refs.push({
      ...name,
      ...(xy ? (xValues ? { xValues } : {}) : shared ? { categories: shared } : {}),
      values,
      ...(row.renderType ? { renderType: row.renderType } : {}),
      ...(row.axis ? { axis: row.axis } : {}),
    });
  }
  const normalized = normalizeChartSeriesRefs(refs);
  return normalized ? { refs: normalized } : { invalid: 0 };
}

/** A picked rectangle as an Edit Series box writes it: `=Sheet1!$B$2:$B$6`. */
export function selectDataPickedReference(prefix: string, range: ChartSourceRange): string {
  const first = `$${columnToLetters(range.c1)}$${range.r1 + 1}`;
  const last = `$${columnToLetters(range.c2)}$${range.r2 + 1}`;
  return `=${prefix}!${first}${first === last ? "" : `:${last}`}`;
}

/** The cells of a reference box, for the dialog's lists (a series' name, the
 * axis labels): `read(sheetName, row, col)` is the displayed value. */
export function selectDataReferenceCells(
  text: string,
  defaultSheetName: string,
  read: (sheetName: string, row: number, col: number) => string | undefined,
  limit = 200,
): string[] | null {
  const typed = text.trim().replace(/^=/, "").trim();
  const areas = parseChartRefAreas(typed.includes(",") && !typed.startsWith("(") ? `(${typed})` : typed);
  if (!areas) return null;
  const out: string[] = [];
  for (const parsed of areas) {
    const { r1, c1, r2, c2 } = parsed.range;
    const sheetName = parsed.sheetName ?? defaultSheetName;
    for (let row = r1; row <= r2; row += 1) {
      for (let col = c1; col <= c2; col += 1) {
        if (out.length >= limit) return out;
        const value = read(sheetName, row, col);
        if (value === undefined) return null;
        out.push(value);
      }
    }
  }
  return out;
}
