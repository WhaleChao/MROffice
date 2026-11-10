/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Every place a worksheet stores a formula or a reference that can NAME
 * another sheet, walked by one function so a sheet rename, a sheet delete and
 * a sheet copy all reach the same set of sources. Missing one is how a rename
 * left a table's calculated-column template on the old name: the existing cells
 * followed, and the next row the table grew was generated from the stale
 * template and came out #REF!.
 *
 * The sources, each read in the shape its own module persists:
 *  - cell formulas (`value` and its parallel `formula` copy),
 *  - defined names (`namedRanges[].ref`),
 *  - data-validation sources (a list rule's `sourceRange`, a custom rule's
 *    `formula`),
 *  - table templates (`dataTables[].calculatedFormulas` / `totalFormulas`),
 *  - conditional formats (a formula rule, and the `formula` thresholds of
 *    colour scales, data bars and icon sets),
 *  - sparkline data sources (`sparklines[key].range`),
 *  - chart series references (`charts[].seriesRefs` and the model's
 *    `charts[].binding.seriesRefs`), and a drawing chart object's
 *    (`objects[]` of type chart: its series references and A1 `range` /
 *    `ranges`).
 *
 * Internal hyperlinks (`#Sheet2!A1`) are deliberately not here: Excel keeps a
 * hyperlink's text as typed when the sheet it names is renamed or deleted.
 *
 * Each rewriter returns the SAME object for a sheet it does not change, so an
 * untouched sheet keeps its identity (no re-render, no spurious save).
 */

import { breakDataTableReferences } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { spreadsheetCellIsText } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import {
  adjustFormulaForSheetMove,
  invalidateSheetInFormula,
  renameSheetInFormula,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

type FormulaRewrite = (formula: string) => string;
type JsonRecord = Record<string, unknown>;

export interface SheetFormulaReferenceOptions {
  /**
   * Leave defined names alone. A workbook name is mirrored into every sheet,
   * so rewriting only one sheet's mirror (a copy's self references) would make
   * the mirrors disagree; renames and deletes rewrite every sheet alike.
   */
  skipNamedRanges?: boolean;
  /** The rewrite a chart's series references take instead of `rewrite`
   * (they are references alone, which Excel stores its own way). */
  seriesRewrite?: FormulaRewrite;
}

function isRecord(value: unknown): value is JsonRecord {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function rewriteString(value: unknown, rewrite: FormulaRewrite): unknown {
  return typeof value === "string" && value !== "" ? rewrite(value) : value;
}

/** Rewrite one keyed record's string values, or return it unchanged. */
function rewriteRecordValues(
  record: unknown,
  rewrite: FormulaRewrite,
): unknown {
  if (!isRecord(record)) return record;
  let changed = false;
  const next: JsonRecord = {};
  for (const [key, value] of Object.entries(record)) {
    const rewritten = rewriteString(value, rewrite);
    if (rewritten !== value) changed = true;
    next[key] = rewritten;
  }
  return changed ? next : record;
}

/** Rewrite named fields of one record, or return it unchanged. */
function rewriteFields(
  record: JsonRecord,
  fields: readonly string[],
  rewrite: FormulaRewrite,
): JsonRecord {
  let next = record;
  for (const field of fields) {
    const value = record[field];
    const rewritten = rewriteString(value, rewrite);
    if (rewritten !== value) {
      if (next === record) next = { ...record };
      next[field] = rewritten;
    }
  }
  return next;
}

function rewriteArray(
  items: unknown,
  rewriteItem: (item: unknown) => unknown,
): unknown {
  if (!Array.isArray(items) || items.length === 0) return items;
  let changed = false;
  const next = items.map((item) => {
    const rewritten = rewriteItem(item);
    if (rewritten !== item) changed = true;
    return rewritten;
  });
  return changed ? next : items;
}

function rewriteCells(cells: unknown, rewrite: FormulaRewrite): unknown {
  if (!isRecord(cells)) return cells;
  let changed = false;
  const next: JsonRecord = {};
  for (const [key, cell] of Object.entries(cells)) {
    let nextCell = cell;
    if (isRecord(cell)) {
      // Only an entered formula is a formula: text that merely mentions a
      // sheet ("see Sheet2!A1") is never rewritten, nor is an entry stored
      // as Text ('=Sheet2!A1), however it starts. A formula's own `kind` is
      // "text" (it is not a number literal), so the text identity is read
      // from the entry's quote prefix or Text format, not from `kind`.
      for (const field of ["value", "formula"] as const) {
        const value = (nextCell as JsonRecord)[field];
        if (typeof value !== "string" || !value.startsWith("=")) continue;
        if (field === "value" && spreadsheetCellIsText(cell as Parameters<typeof spreadsheetCellIsText>[0])) continue;
        const rewritten = rewrite(value);
        if (rewritten === value) continue;
        if (nextCell === cell) nextCell = { ...cell };
        (nextCell as JsonRecord)[field] = rewritten;
      }
    }
    if (nextCell !== cell) changed = true;
    next[key] = nextCell;
  }
  return changed ? next : cells;
}

function rewriteValidation(rule: unknown, rewrite: FormulaRewrite): unknown {
  if (!isRecord(rule)) return rule;
  if (rule.kind === "list") return rewriteFields(rule, ["sourceRange"], rewrite);
  if (rule.kind === "formula") return rewriteFields(rule, ["formula"], rewrite);
  // A number/date/time/length bound typed as a reference (=$H$15) is a
  // formula too, stored without its "=".
  if (rule.kind === "number") return rewriteFields(rule, ["minFormula", "maxFormula"], rewrite);
  return rule;
}

function rewriteDataTable(table: unknown, rewrite: FormulaRewrite): unknown {
  if (!isRecord(table)) return table;
  let next = table;
  for (const field of ["calculatedFormulas", "totalFormulas"] as const) {
    const templates = table[field];
    const rewritten = rewriteRecordValues(templates, rewrite);
    if (rewritten !== templates) {
      if (next === table) next = { ...table };
      next[field] = rewritten;
    }
  }
  return next;
}

function rewriteThreshold(stop: unknown, rewrite: FormulaRewrite): unknown {
  return isRecord(stop) && stop.type === "formula"
    ? rewriteFields(stop, ["formula"], rewrite)
    : stop;
}

function rewriteConditionalFormat(
  rule: unknown,
  rewrite: FormulaRewrite,
): unknown {
  if (!isRecord(rule)) return rule;
  if (rule.kind === "formula") return rewriteFields(rule, ["formula"], rewrite);
  let next = rule;
  const endpoints =
    rule.kind === "colorScale"
      ? (["min", "mid", "max"] as const)
      : rule.kind === "dataBar"
        ? (["min", "max"] as const)
        : [];
  for (const endpoint of endpoints) {
    const stop = rule[endpoint];
    const rewritten = rewriteThreshold(stop, rewrite);
    if (rewritten !== stop) {
      if (next === rule) next = { ...rule };
      next[endpoint] = rewritten;
    }
  }
  if (rule.kind === "iconSet") {
    const stops = rewriteArray(rule.stops, (stop) =>
      rewriteThreshold(stop, rewrite),
    );
    if (stops !== rule.stops) next = { ...rule, stops };
  }
  return next;
}

function rewriteSparklines(sparklines: unknown, rewrite: FormulaRewrite): unknown {
  if (!isRecord(sparklines)) return sparklines;
  let changed = false;
  const next: JsonRecord = {};
  for (const [key, sparkline] of Object.entries(sparklines)) {
    const rewritten = isRecord(sparkline)
      ? rewriteFields(sparkline, ["range"], rewrite)
      : sparkline;
    if (rewritten !== sparkline) changed = true;
    next[key] = rewritten;
  }
  return changed ? next : sparklines;
}

const CHART_SERIES_REFERENCE_FIELDS = [
  "title",
  "categories",
  "xValues",
  "values",
  "bubbleSizes",
] as const;

function rewriteSeriesRefs(refs: unknown, rewrite: FormulaRewrite): unknown {
  return rewriteArray(refs, (ref) =>
    isRecord(ref) ? rewriteFields(ref, CHART_SERIES_REFERENCE_FIELDS, rewrite) : ref,
  );
}

/**
 * A chart's series references: the flat `seriesRefs` and its model's
 * `binding.seriesRefs` (lib/spreadsheet/chart-v2-model), which the flat ones
 * are derived from -- both follow, so neither is left on the old name.
 */
function rewriteChart(chart: unknown, rewrite: FormulaRewrite): unknown {
  if (!isRecord(chart)) return chart;
  const seriesRefs = rewriteSeriesRefs(chart.seriesRefs, rewrite);
  const binding = chart.binding;
  const bindingRefs = isRecord(binding)
    ? rewriteSeriesRefs(binding.seriesRefs, rewrite)
    : undefined;
  const bindingChanged = isRecord(binding) && bindingRefs !== binding.seriesRefs;
  if (seriesRefs === chart.seriesRefs && !bindingChanged) return chart;
  return {
    ...chart,
    ...(seriesRefs !== chart.seriesRefs ? { seriesRefs } : {}),
    ...(bindingChanged ? { binding: { ...binding, seriesRefs: bindingRefs } } : {}),
  };
}

function rewriteChartObject(object: unknown, seriesRewrite: FormulaRewrite): unknown {
  if (!isRecord(object) || object.type !== "chart") return object;
  let next = rewriteChart(object, seriesRewrite) as JsonRecord;
  const range = rewriteString(object.range, seriesRewrite);
  const ranges = Array.isArray(object.ranges)
    ? rewriteArray(object.ranges, (entry) => rewriteString(entry, seriesRewrite))
    : object.ranges;
  if (range !== object.range || ranges !== object.ranges) {
    next = { ...next, range, ranges };
    if (object.ranges === undefined) delete next.ranges;
  }
  return next;
}

/**
 * Apply `rewrite` to every formula source of one sheet. Returns the input when
 * nothing changed. Object spread keeps the sheet's own symbol-keyed runtime
 * hints (the blank-cell-format ranges), which stay valid across a rewrite.
 */
export function rewriteSheetFormulaReferences<T extends object>(
  sheet: T,
  rewrite: FormulaRewrite,
  options: SheetFormulaReferenceOptions = {},
): T {
  const source = sheet as unknown as JsonRecord;
  let next = source;
  const assign = (field: string, value: unknown) => {
    if (value === source[field]) return;
    if (next === source) next = { ...source };
    next[field] = value;
  };
  assign("cells", rewriteCells(source.cells, rewrite));
  if (!options.skipNamedRanges) {
    assign(
      "namedRanges",
      rewriteArray(source.namedRanges, (entry) =>
        isRecord(entry) ? rewriteFields(entry, ["ref"], rewrite) : entry,
      ),
    );
  }
  assign(
    "validations",
    rewriteArray(source.validations, (rule) => rewriteValidation(rule, rewrite)),
  );
  assign(
    "dataTables",
    rewriteArray(source.dataTables, (table) => rewriteDataTable(table, rewrite)),
  );
  assign(
    "conditionalFormats",
    rewriteArray(source.conditionalFormats, (rule) =>
      rewriteConditionalFormat(rule, rewrite),
    ),
  );
  assign("sparklines", rewriteSparklines(source.sparklines, rewrite));
  const seriesRewrite = options.seriesRewrite ?? rewrite;
  assign(
    "charts",
    rewriteArray(source.charts, (chart) => rewriteChart(chart, seriesRewrite)),
  );
  // A chart another application wrote is a drawing object: its series
  // references, and the A1 `range` / `ranges` they are summarised in.
  assign(
    "objects",
    rewriteArray(source.objects, (object) => rewriteChartObject(object, seriesRewrite)),
  );
  return next as unknown as T;
}

/** Apply `rewrite` to one conditional format's formulas (the same object
 * when nothing changed). */
export function rewriteConditionalFormatReferences<R extends object>(
  rule: R,
  rewrite: FormulaRewrite,
): R {
  return rewriteConditionalFormat(rule, rewrite) as R;
}

/** Apply `rewrite` to one data-validation rule's formulas and sources (the
 * same object when nothing changed). */
export function rewriteValidationReferences<R extends object>(
  rule: R,
  rewrite: FormulaRewrite,
): R {
  return rewriteValidation(rule, rewrite) as R;
}

/** Follow a sheet rename across every formula source of `sheet`. */
export function renameSheetReferences<T extends object>(
  sheet: T,
  oldName: string,
  newName: string,
  options?: SheetFormulaReferenceOptions,
): T {
  return rewriteSheetFormulaReferences(
    sheet,
    (formula) => renameSheetInFormula(formula, oldName, newName),
    options,
  );
}

/**
 * Rewrite every 3-D span `sheet` holds for a sheet moved to another tab
 * position (see {@link adjustFormulaForSheetMove}).
 */
export function adjustSheetMoveReferences<T extends object>(
  sheet: T,
  orderBefore: readonly string[],
  orderAfter: readonly string[],
  movedSheet: string,
): T {
  return rewriteSheetFormulaReferences(sheet, (formula) =>
    adjustFormulaForSheetMove(formula, orderBefore, orderAfter, movedSheet),
  );
}

/**
 * Break every reference `sheet` holds to a sheet being deleted (Excel's
 * `#REF!`). `sheetOrder` is the workbook's tab order before the delete.
 */
export function invalidateDeletedSheetReferences<T extends object>(
  sheet: T,
  deletedName: string,
  sheetOrder: readonly string[],
   
  deletedTables: readonly string[] = [],
): T {
  const invalidate = (formula: string) =>
    invalidateSheetInFormula(
      deletedTables.reduce((text, table) => breakDataTableReferences(text, table), formula),
      deletedName,
      sheetOrder,
    );
  const rewritten = rewriteSheetFormulaReferences(sheet, invalidate, {
    // A chart's series reference to the deleted sheet is the bare #REF!
    // Excel writes (`<c:f>#REF!</c:f>`, Excel 16 Mac, evidence/xprobe-misc7;
    
    seriesRewrite: (formula) => {
      const broken = invalidate(formula);
      if (broken === formula) return broken;
      const eq = broken.startsWith("=") ? "=" : "";
      const body = eq ? broken.slice(1) : broken;
      return body.startsWith("#REF!") && !body.includes(",") ? `${eq}#REF!` : broken;
    },
  });
  return withoutBrokenSpills(sheet, rewritten);
}









function withoutBrokenSpills<T extends object>(before: T, after: T): T {
  const beforeCells = (before as { cells?: Record<string, { formula?: unknown; spillRef?: unknown }> }).cells;
  const afterCells = (after as { cells?: Record<string, { formula?: unknown; spillRef?: unknown }> }).cells;
  if (!afterCells || afterCells === beforeCells) return after;
  let cells: Record<string, unknown> | null = null;
  for (const key in afterCells) {
    const cell = afterCells[key];
    const previous = beforeCells?.[key];
    if (!cell || typeof cell !== "object" || cell.spillRef === undefined || !previous) continue;
    if (cell.formula === previous.formula) continue;
    cells ??= { ...afterCells };
    const rest: Record<string, unknown> = { ...(cell as Record<string, unknown>) };
    delete rest.spillRef;
    cells[key] = rest;
  }
  return cells ? ({ ...after, cells } as T) : after;
}
