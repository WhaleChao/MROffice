/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The chart's own numbers, edited in place beside the chart: the fourth quick
 * button under Chart Elements / Styles / Filters opens a small grid of the
 * categories, the series names and their values, so a value can be changed
 * without opening the chart's data sheet.
 *
 * It is the same edit Edit Data makes, not a second model: a chart that keeps
 * a workbook has the typed value written into the workbook cell the chart's
 * source range reads it from, and the chart's data is recalculated from the
 * workbook, exactly as the data sheet's own commit does. A cell that holds a
 * formula is shown and not editable here (its value is the formula's), and a
 * chart without a workbook (an imported classic chart) has its data edited
 * directly. The grid never adds or removes a row or a series: that, and every
 * kind of chart data other than categories by series, stays in the sheet.
 */
import { isCellEditable, type ProtectableSheet } from "../spreadsheet/protection";
import {
  validateChartBlockAttrsV2,
  type ChartBlockAttrsV2,
  type ChartData,
} from "../../components/console/editor/extensions/chartBlockShared";
import { calculateDocumentChartWorkbook } from "../../components/console/editor/extensions/documentChartWorkbookCalculation";
import {
  chartCellFormula,
  normalizeDocumentChartWorkbook,
  type ChartWorkbookCell,
  type DocumentChartWorkbook,
} from "../../components/console/editor/extensions/documentChartWorkbook";

export type ChartQuickDataCell = Readonly<{
  text: string;
  /** False where the value comes from a formula in the chart's sheet. */
  editable: boolean;
}>;

export type ChartQuickData = Readonly<{
  categories: readonly ChartQuickDataCell[];
  series: readonly Readonly<{
    name: ChartQuickDataCell;
    values: readonly ChartQuickDataCell[];
  }>[];
}>;

export type ChartQuickDataEdit =
  | { readonly kind: "category"; readonly index: number; readonly text: string }
  | { readonly kind: "seriesName"; readonly series: number; readonly text: string }
  | {
      readonly kind: "value";
      readonly series: number;
      readonly index: number;
      /** `null` empties the cell: a blank point. */
      readonly value: number | null;
    };

type CategoryLikeData = Extract<ChartData, { kind: "category" | "stock" }>;

function categoryData(attrs: ChartBlockAttrsV2): CategoryLikeData | null {
  return attrs.data.kind === "category" || attrs.data.kind === "stock"
    ? attrs.data
    : null;
}

/** Where the source range keeps a cell: the header row, category column. */
function sourceCell(
  workbook: DocumentChartWorkbook,
  row: number,
  column: number,
): { sheetIndex: number; key: string } | null {
  const binding = workbook.binding;
  const sheetIndex = workbook.content.sheets.findIndex(
    (sheet) => sheet.id === binding.sheetId,
  );
  if (sheetIndex < 0) return null;
  // Plotted by rows, the series run across and the categories down swap.
  const [r, c] =
    binding.plotBy === "rows"
      ? [binding.r1 + column, binding.c1 + row]
      : [binding.r1 + row, binding.c1 + column];
  return { sheetIndex, key: `${r}:${c}` };
}


function editableCell(
  workbook: DocumentChartWorkbook | null,
  row: number,
  column: number,
): boolean {
  if (!workbook) return true;
  const at = sourceCell(workbook, row, column);
  if (!at) return true;
  const sheet = workbook.content.sheets[at.sheetIndex];
  const [r, c] = at.key.split(":").map(Number);
  // A protected sheet lets only an explicitly unlocked cell change, in the
  // full workbook and here alike (PPT-B384).
  if (!isCellEditable(sheet as unknown as ProtectableSheet, r, c)) return false;
  const cell = sheet.cells[at.key];
  return !cell || chartCellFormula(cell) === undefined;
}

function workbookOf(raw: DocumentChartWorkbook | null | undefined): DocumentChartWorkbook | null {
  if (!raw) return null;
  try {
    return normalizeDocumentChartWorkbook(raw);
  } catch {
    return null;
  }
}

/**
 * A value as the grid shows it: to the 15 significant digits a sheet keeps,
 * which clears binary noise (0.1 + 0.2) without rounding a small value to 0
 * (a fixed count of decimals showed 1e-11 as 0, and a 0 typed over it read as
 * no change).
 */
function valueText(value: number | null | undefined): string {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  return String(Number.isInteger(value) ? value : Number(value.toPrecision(15)));
}

/**
 * The grid a chart shows in its quick data flyout, or null where its data is
 * not categories by series (those are edited in the sheet).
 */
export function chartQuickData(
  attrs: ChartBlockAttrsV2,
  rawWorkbook?: DocumentChartWorkbook | null,
): ChartQuickData | null {
  const data = categoryData(attrs);
  if (!data) return null;
  const workbook = workbookOf(rawWorkbook);
  return {
    categories: data.categories.map((category, index) => ({
      text: category,
      editable: editableCell(workbook, index + 1, 0),
    })),
    series: data.series.map((series, seriesIndex) => ({
      name: { text: series.name, editable: editableCell(workbook, 0, seriesIndex + 1) },
      values: data.categories.map((_, index) => ({
        text: valueText(series.values[index]),
        editable: editableCell(workbook, index + 1, seriesIndex + 1),
      })),
    })),
  };
}

/**
 * The chart after one quick edit, with its recalculated workbook where it
 * keeps one; null when the edit cannot be made here (a formula cell, an empty
 * label, a value the chart cannot take).
 */
export function chartWithQuickDataEdit(
  attrs: ChartBlockAttrsV2,
  rawWorkbook: DocumentChartWorkbook | null | undefined,
  edit: ChartQuickDataEdit,
): { attrs: ChartBlockAttrsV2; workbook?: DocumentChartWorkbook } | null {
  const data = categoryData(attrs);
  if (!data) return null;
  if (edit.kind !== "value" && !edit.text.trim()) return null;
  // `null` empties the cell: the chart's blank point (never a 0).
  if (edit.kind === "value" && edit.value !== null && !Number.isFinite(edit.value)) return null;
  const [row, column] =
    edit.kind === "category"
      ? [edit.index + 1, 0]
      : edit.kind === "seriesName"
        ? [0, edit.series + 1]
        : [edit.index + 1, edit.series + 1];
  if (
    (edit.kind !== "seriesName" && (edit.index < 0 || edit.index >= data.categories.length)) ||
    (edit.kind !== "category" && (edit.series < 0 || edit.series >= data.series.length))
  ) {
    return null;
  }
  const workbook = workbookOf(rawWorkbook);
  if (workbook) {
    if (!editableCell(workbook, row, column)) return null;
    const at = sourceCell(workbook, row, column);
    if (!at) return null;
    const next = structuredClone(workbook);
    const sheet = next.content.sheets[at.sheetIndex];
    const { formula: _formula, computed: _computed, ...kept } = sheet.cells[at.key] ?? {};
    if (edit.kind === "value" && edit.value === null) {
      // Clearing a value is Delete on the sheet: the cell's own formatting
      // (a fill, a number format) stays, as the full sheet keeps it; the
      // record goes only when nothing else is on it (PPT-B347).
      const {
        value: _value, kind: _kind, runs: _runs, display: _display, chartValueCache: _cache, phonetic: _phonetic,
        ...formatting
      } = kept as Record<string, unknown>;
      if (Object.keys(formatting).length === 0) delete sheet.cells[at.key];
      else sheet.cells[at.key] = { ...formatting, value: "", kind: "text" } as (typeof sheet.cells)[string];
    } else
      sheet.cells[at.key] =
        edit.kind === "value"
          ? { ...kept, value: String(edit.value), kind: "number" }
          : { ...kept, value: edit.text, kind: "text" };
    try {
      const calculated = calculateDocumentChartWorkbook(attrs, next);
      const nextAttrs = { ...attrs, data: calculated.data };
      return validateChartBlockAttrsV2(nextAttrs).valid
        ? { attrs: nextAttrs, workbook: calculated.workbook }
        : null;
    } catch {
      return null;
    }
  }
  const nextData = structuredClone(data) as CategoryLikeData;
  if (edit.kind === "category") nextData.categories[edit.index] = edit.text;
  else if (edit.kind === "seriesName") nextData.series[edit.series].name = edit.text;
  else nextData.series[edit.series].values[edit.index] = edit.value;
  const nextAttrs = { ...attrs, data: nextData } as ChartBlockAttrsV2;
  return validateChartBlockAttrsV2(nextAttrs).valid ? { attrs: nextAttrs } : null;
}
