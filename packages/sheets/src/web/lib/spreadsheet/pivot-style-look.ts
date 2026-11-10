/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizePivotRange, type PivotConfig, type PivotNativeSavedLayout, type PivotRange } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import type { PivotOutputCell, PivotOutputCellRole, PivotOutputPlan } from "@mrnote/core/web/lib/spreadsheet/pivot-output.ts";
import { isBuiltInPivotStyle, type PivotStyleCellLook, type PivotStyleElement } from "./pivot-style-recipes";

export interface PivotStyleOptions {
  showRowHeaders: boolean;
  showColumnHeaders: boolean;
  bandedRows: boolean;
  bandedColumns: boolean;
}

export function pivotStyleOptionsOf(config: PivotConfig): PivotStyleOptions {
  return {
    showRowHeaders: config.showRowHeaders !== false,
    showColumnHeaders: config.showColumnHeaders !== false,
    bandedRows: config.bandedRows === true,
    bandedColumns: config.bandedColumns === true,
  };
}

const ROW_SUBHEADINGS: readonly PivotStyleElement[] = [
  "firstRowSubheading", "secondRowSubheading", "thirdRowSubheading",
];
const COLUMN_SUBHEADINGS: readonly PivotStyleElement[] = [
  "firstColumnSubheading", "secondColumnSubheading", "thirdColumnSubheading",
];
const SUBTOTAL_ROWS: readonly PivotStyleElement[] = [
  "firstSubtotalRow", "secondSubtotalRow", "thirdSubtotalRow",
];

/** The Row Subheading a row item's level wears: 1, 2, 3, then 2 and 3 in
 * turn for the deeper levels (Excel for Mac, PivotStyleLight2 over six
 * outline row fields: levels 1 and 3 blue, 2 and 4 plain; O06, X-235). */
function rowSubheadingSlot(level: number): number {
  return level <= 2 ? level : 1 + ((level - 1) % 2);
}

/** What the style look reads of a report: a planned report, or the cached
 * one an undrawn PivotTable shows (`savedPivotStyleReport`). */
export type PivotStyleLookPlan = Pick<PivotOutputPlan, "output" | "headerRows" | "firstDataCol" | "cells"> & {
  /** The report's row fields, where the recipe does not name them (an OLAP
   * cube's report). */
  rowFieldCount?: number;
};

export function pivotStyleCellLook(
  plan: PivotStyleLookPlan,
  config: PivotConfig,
  options: PivotStyleOptions,
  row: number,
  col: number,
): PivotStyleCellLook | null {
  const cell = plan.cells.get(`${row}:${col}`);
  if (!cell) return null;
  const output = plan.output;
  if (!output) return null;
  const elements: PivotStyleCellLook["elements"][number][] = [];
  if (cell.role === "page-field" || cell.role === "page-item") {
    elements.push({
      element: cell.role === "page-field" ? "pageFieldLabels" : "pageFieldValues",
      top: true, bottom: true, left: true, right: true,
    });
    return { elements };
  }
  const { r1, c1, r2, c2 } = output;
  if (row < r1 || row > r2 || col < c1 || col > c2) return null;
  const firstBodyRow = r1 + plan.headerRows;
  const firstDataCol = c1 + plan.firstDataCol;
  const tabular = (config.reportLayout ?? "tabular") === "tabular";
  const compact = config.reportLayout === "compact";
  const at = (r: number, c: number) => plan.cells.get(`${r}:${c}`);
  const isGrandTotalRow = (r: number) => at(r, c1)?.role === "grand-total-label";
  // A column's kind shows in an item row's data cell (a grand total or
  // subtotal row carries its own role across every column).
  const columnRoleOf = (c: number) => {
    for (let r = firstBodyRow; r <= r2; r += 1) {
      if (at(r, firstDataCol)?.role !== "value") continue;
      return at(r, c)?.role ?? "value";
    }
    return "value";
  };
  const isGrandTotalCol = (c: number) => c >= firstDataCol && columnRoleOf(c) === "grand-total-value";
  const isSubtotalCol = (c: number) => c >= firstDataCol && columnRoleOf(c) === "subtotal-value";
  // Every label column of a subtotal row carries the role; the level is the
  // column its label stands in (its indent in compact form), not the first
  // column -- else every outline level wore the first level's look (O06:
  // Excel paints levels 1 and 3 as the second Row Subheading, X-235).
  const subtotalLevelOfRow = (r: number): number | null => {
    let level: number | null = null;
    for (let c = c1; c < firstDataCol; c += 1) {
      const candidate = at(r, c);
      if (candidate?.role !== "subtotal-label") continue;
      if (compact) return candidate.indent ?? 0;
      if (level === null || candidate.value != null) level = c - c1;
    }
    return level;
  };
  const edgesRow = (r: number) => ({ top: true, bottom: true, left: col === c1, right: col === c2, r });
  elements.push({ element: "wholeTable", top: row === r1, bottom: row === r2, left: col === c1, right: col === c2 });

  const headerRow = row < firstBodyRow;
  const grandTotalCol = isGrandTotalCol(col);
  const grandTotalRow = isGrandTotalRow(row);
  if (!headerRow) {
    // Banded columns run over the grand total columns too.
    if (options.bandedColumns && col >= firstDataCol) {
      elements.push({
        element: (col - firstDataCol) % 2 === 0 ? "firstColumnStripe" : "secondColumnStripe",
        top: row === firstBodyRow, bottom: row === r2, left: true, right: true,
      });
    }
    // The stripes run on through the grand total row (its own element paints
    // over them where the style says so).
    if (options.bandedRows) {
      const e = edgesRow(row);
      elements.push({
        element: (row - firstBodyRow) % 2 === 0 ? "firstRowStripe" : "secondRowStripe",
        top: e.top, bottom: e.bottom, left: e.left, right: e.right,
      });
    }
  }
  // Row Headers covers the label columns from the top of the report down.
  if (options.showRowHeaders && col < firstDataCol) {
    elements.push({ element: "firstColumn", top: row === r1, bottom: row === r2, left: col === c1, right: col === firstDataCol - 1 });
  }
  // A PivotTable has no "Last Column" option: Excel never paints the
  // lastColumn element of a style on a report (checked against desktop Excel).
  if (headerRow && options.showColumnHeaders) {
    elements.push({ element: "headerRow", top: row === r1, bottom: row === firstBodyRow - 1, left: col === c1, right: col === c2 });
    // The corner above the row labels (the column item rows over the label
    // columns) is the First Header Cell; the caption row itself is not.
    if (col < firstDataCol && row < firstBodyRow - 1) {
      elements.push({
        element: "firstHeaderCell",
        top: row === r1, bottom: row === firstBodyRow - 2, left: col === c1, right: col === firstDataCol - 1,
      });
    }
    // Every header row under the field caption row is a column subheading of
    // its level (the column items, then the value captions under them); a
    // grand total column's caption is not one. An item's region spans the
    // blank cells to its right that it covers.
    const itemLevel = row - r1 - 1;
    if (col >= firstDataCol && itemLevel >= 0 && !grandTotalCol) {
      const element = COLUMN_SUBHEADINGS[Math.min(itemLevel, 2)];
      const covered = (c: number) => {
        const candidate = at(row, c);
        return candidate?.role === "column-header" && candidate.value == null && !isGrandTotalCol(c);
      };
      elements.push({
        element, top: true, bottom: true,
        left: !covered(col) || col === firstDataCol,
        right: !covered(col + 1) || col === c2,
      });
    }
  }
  if (!headerRow && isSubtotalCol(col)) {
    elements.push({ element: "firstSubtotalColumn", top: row === firstBodyRow, bottom: row === r2, left: true, right: true });
  }
  // The level of a body row's item: its label cell's column (its indent in
  // compact form); null for a grand total or blank row.
  const rowLevel = (r: number): number | null => {
    let level: number | null = null;
    for (let c = c1; c < firstDataCol; c += 1) {
      const candidate = at(r, c);
      if (candidate?.role !== "row-header" && candidate?.role !== "subtotal-label") return null;
      if (candidate.value != null) level = compact ? (candidate.indent ?? 0) : c - c1;
    }
    return level;
  };
  if (!headerRow && !grandTotalRow) {
    const level = subtotalLevelOfRow(row);
    if (level !== null) {
      // Outline and compact form give an outer item a row of its own, which
      // its items' rows follow: that whole row is the Row Subheading of its
      // level (checked against Excel). A subtotal row has no deeper row after it.
      const next = rowLevel(row + 1);
      const itemRow = !tabular && next !== null && next > level;
      const e = edgesRow(row);
      elements.push({
        element: itemRow ? ROW_SUBHEADINGS[rowSubheadingSlot(level)] : SUBTOTAL_ROWS[Math.min(level, 2)],
        top: e.top, bottom: e.bottom, left: e.left, right: e.right,
      });
    } else if (cell.role === "row-header" && col < firstDataCol) {
      const level = compact ? (cell.indent ?? 0) : col - c1;
      const rowFieldCount = plan.rowFieldCount ?? config.rows.length;
      if (level < rowFieldCount - 1 && tabular) {
        // Tabular form: the item's label cell and the blank cells under it
        // (its inner items' rows) are one region of the subheading.
        const continues = (r: number) => {
          const candidate = at(r, col);
          return candidate?.role === "row-header" && candidate.value == null;
        };
        elements.push({
          element: ROW_SUBHEADINGS[rowSubheadingSlot(level)],
          top: !continues(row), bottom: !continues(row + 1), left: true, right: true,
        });
      }
    }
  }
  if (grandTotalRow) {
    const e = edgesRow(row);
    elements.push({ element: "totalRow", top: e.top, bottom: e.bottom, left: e.left, right: e.right });
  }
  return { elements };
}

/** An undrawn PivotTable's report as its style paints it: the cells the file
 * cached, laid out by the receipt's saved layout. */
export interface SavedPivotStyleReport {
  config: PivotConfig;
  styleName: string;
  options: PivotStyleOptions;
  plan: PivotStyleLookPlan;
  /** The report and its report filters. */
  area: PivotRange;
}

/**
 * The style an imported PivotTable the grid does not draw wears over the
 * cells the file cached (a sheet's other PivotTables, a preserve-only report
 * on an OLAP cube): the report's own areas -- header rows, row label
 * columns, item / subtotal / grand total rows, subtotal and grand total
 * columns, report filters -- from its `location` and cached `rowItems` /
 * `colItems`, as Excel lays the style over them (X-235). Null without a
 * saved layout that matches the report, or for a style that is not one of
 * Excel's own: Excel draws a style name the workbook does not define as no
 * style (O03's "PivotTableStyle2").
 */
export function savedPivotStyleReport(config: PivotConfig): SavedPivotStyleReport | null {
  const native = config.native;
  const layout = native?.savedLayout;
  if (!native || !layout) return null;
  const styleName = config.styleName ?? native.styleName;
  if (!styleName || !isBuiltInPivotStyle(styleName)) return null;
  const output = normalizePivotRange(native.outputRef);
  if (!output) return null;
  const plan = savedPivotStylePlan(layout, output, config.reportLayout);
  if (!plan) return null;
  let area: PivotRange = output;
  for (const cell of plan.cells.values()) {
    if (cell.role !== "page-field" && cell.role !== "page-item") continue;
    area = {
      r1: Math.min(area.r1, cell.row), c1: Math.min(area.c1, cell.col),
      r2: Math.max(area.r2, cell.row), c2: Math.max(area.c2, cell.col),
    };
  }
  return { config, styleName, options: pivotStyleOptionsOf(config), plan, area };
}

function savedPivotStylePlan(
  layout: PivotNativeSavedLayout,
  output: PivotRange,
  reportLayout: PivotConfig["reportLayout"],
): PivotStyleLookPlan | null {
  const { r1, c1, r2, c2 } = output;
  const headerRows = layout.firstDataRow;
  const firstDataCol = layout.firstDataCol;
  const bodyRows = r2 - r1 + 1 - headerRows;
  const dataCols = c2 - c1 + 1 - firstDataCol;
  if (bodyRows < 0 || dataCols < 0) return null;
  if (
    (layout.rowItems.length && layout.rowItems.length !== bodyRows) ||
    (layout.colItems.length && layout.colItems.length !== dataCols)
  ) {
    return null;
  }
  const compact = reportLayout === "compact" && firstDataCol <= 1;
  const leafLevel = Math.max(0, layout.rowFieldCount - 1);
  const cells = new Map<string, PivotOutputCell>();
  const MARK = "";
  const put = (row: number, col: number, role: PivotOutputCellRole, value: string | null, indent?: number) => {
    cells.set(`${row}:${col}`, { row, col, role, value, ...(indent ? { indent } : {}) });
  };
  // A data column's kind from its cached column item.
  const columnRole = (index: number): "value" | "subtotal-value" | "grand-total-value" => {
    const kind = layout.colItems[index]?.[0] ?? 0;
    return kind === 2 ? "grand-total-value" : kind === 1 ? "subtotal-value" : "value";
  };
  // Header rows: the caption row, then the column items level by level.
  const columnLabels = new Set<string>();
  layout.colItems.forEach(([kind, level, members], index) => {
    const col = c1 + firstDataCol + index;
    const last = kind === 0 ? level + Math.max(1, members) - 1 : level;
    for (let depth = level; depth <= last; depth += 1) {
      columnLabels.add(`${r1 + layout.firstHeaderRow + depth}:${col}`);
    }
  });
  for (let row = r1; row < r1 + headerRows; row += 1) {
    for (let col = c1; col <= c2; col += 1) {
      const key = `${row}:${col}`;
      put(row, col, col < c1 + firstDataCol ? "corner" : "column-header",
        row === r1 || columnLabels.has(key) ? MARK : null);
    }
  }
  // Body rows: each cached row item's kind and level.
  for (let index = 0; index < bodyRows; index += 1) {
    const row = r1 + headerRows + index;
    const [kind, level, members] = layout.rowItems[index] ?? [0, leafLevel, 1];
    const depth = kind === 0 ? level + Math.max(1, members) - 1 : level;
    const leaf = kind === 0 && depth >= leafLevel;
    const labelRole: PivotOutputCellRole =
      kind === 2 ? "grand-total-label" : kind === 3 ? "blank" : leaf ? "row-header" : "subtotal-label";
    for (let offset = 0; offset < firstDataCol; offset += 1) {
      const labelled = kind === 3
        ? false
        : kind === 2
          ? offset === 0
          : compact
            ? offset === 0
            : offset >= level && offset <= depth;
      put(row, c1 + offset, labelRole, labelled ? MARK : null, compact && offset === 0 ? depth : undefined);
    }
    for (let index2 = 0; index2 < dataCols; index2 += 1) {
      const role: PivotOutputCellRole =
        kind === 2 ? "grand-total-value" : kind === 3 ? "blank" : leaf ? columnRole(index2) : "subtotal-value";
      put(row, c1 + firstDataCol + index2, role, null);
    }
  }
  // Report filters above the report (one blank row between), a field's
  // caption beside its selection; several columns of them a column apart.
  const pageCount = layout.pageFieldCount ?? 0;
  const pageRows = Math.max(1, layout.rowPageCount ?? 1);
  const pageCols = Math.max(1, layout.colPageCount ?? 1);
  for (let field = 0; field < pageCount; field += 1) {
    const group = layout.pageOverThenDown ? field % pageCols : Math.floor(field / pageRows);
    const line = layout.pageOverThenDown ? Math.floor(field / pageCols) : field % pageRows;
    const row = r1 - pageRows - 1 + line;
    const col = c1 + group * 3;
    if (row < 0) continue;
    put(row, col, "page-field", MARK);
    put(row, col + 1, "page-item", MARK);
  }
  return { output, headerRows, firstDataCol, cells, rowFieldCount: layout.rowFieldCount };
}
