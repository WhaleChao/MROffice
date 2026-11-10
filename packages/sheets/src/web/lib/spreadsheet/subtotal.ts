/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure planner for Excel / WPS "Data > Subtotal". Given the rows of a contiguous
 * data region (a matrix of DISPLAY strings, header row excluded), a grouping
 * column, and one or more value columns to aggregate, it returns:
 *   - `rows`:   the NEW matrix to write back — the original rows with a SUBTOTAL
 *               summary row inserted after each run of equal grouping-column
 *               values, plus an optional GRAND TOTAL row at the bottom.
 *   - `groups`: the outline ROW groups (relative to the FIRST data row) so the
 *               grid can collapse each group's detail under its summary row —
 *               the same `OutlineGroup` shape the outline-groups model uses.
 *   - `summaryRowOffsets`: the 0-based offsets (within `rows`) of every inserted
 *               summary row, so the caller can style them (bold) if it wants.
 *
 * The SUBTOTAL formulas reference the value column over the group's detail rows
 * using the function number Excel uses (9 = SUM, 1 = AVERAGE, …), so they update
 * live and — because SUBTOTAL ignores nested SUBTOTAL/AGGREGATE cells — the
 * grand total over the whole region does NOT double-count the per-group rows.
 * That mirrors Excel exactly. Everything here is data-only and side-effect free
 * so it unit-tests in isolation; the React layer maps the result onto the grid
 * via its existing paste + outline-group commit paths (undo / autosave covered).
 *
 * Excel's Subtotal command requires the data to be SORTED by the grouping column
 * (otherwise a value that recurs in two non-adjacent runs gets two subtotals).
 * This planner does NOT sort — it groups by *adjacent* equal values exactly as
 * Excel does, so the caller is responsible for sorting first (the UI tells the
 * user, matching Excel's own behaviour of grouping on change-of-value).
 */

import { columnToLetters } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import type { OutlineGroup } from "@mrnote/core/web/lib/spreadsheet/outline-groups.ts";

/**
 * The aggregate functions Excel's Subtotal dialog offers, mapped to the
 * SUBTOTAL `function_num` (the 1-11 family, which COUNTS rows hidden by a
 * filter; the grid's collapse uses the manual-hidden seam so a collapsed group
 * still sums via 101-111 semantics wired elsewhere — here we emit the canonical
 * 1-11 codes Excel writes).
 */
export type SubtotalFunction =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNumbers"
  | "stdDev"
  | "stdDevp"
  | "var"
  | "varp";

/** SUBTOTAL function_num for each function (Excel's 1-11 codes). */
const FUNCTION_NUM: Record<SubtotalFunction, number> = {
  average: 1,
  countNumbers: 2,
  count: 3,
  max: 4,
  min: 5,
  product: 6,
  stdDev: 7,
  stdDevp: 8,
  sum: 9,
  var: 10,
  varp: 11,
};

export interface SubtotalPlanOptions {
  /** The 0-based column to group on (a change of its value ends a group). */
  groupColumn: number;
  /** The 0-based columns to aggregate with SUBTOTAL (must be non-empty). */
  valueColumns: readonly number[];
  /** The aggregate function (default "sum"). */
  function: SubtotalFunction;
  /**
   * The ABSOLUTE sheet row index (0-based) of the FIRST data row, so the emitted
   * SUBTOTAL formulas address the right cells once written back. The header row
   * (if any) sits at `firstDataRow - 1` and is NOT part of `rows`.
   */
  firstDataRow: number;
  /**
   * ABSOLUTE 0-based sheet column of `dataRows[i][0]`: the list's left edge.
   * `groupColumn` / `valueColumns` are offsets from it; the SUBTOTAL formulas
   * address the sheet, so they add it back. Default 0 (a list in column A).
   */
  firstDataColumn?: number;
  /**
   * A label drawn in the grouping column of each summary row, with `{group}`
   * replaced by the group's value (e.g. "{group} Total"). Defaults to just the
   * group value when omitted.
   */
  labelTemplate?: string;
  /** A label for the grand-total row's grouping column (e.g. "Grand Total"). */
  grandTotalLabel?: string;
  /** Whether to append a grand-total row over the whole region (default true). */
  grandTotal?: boolean;
  /**
   * Excel's "Summary below data" (default true). Cleared, each group's summary
   * row goes ABOVE its detail and the grand total above the whole list
   * (XLS-A064).
   */
  summaryBelow?: boolean;
  /**
   * Indices into `dataRows` of rows that are already summary rows (an earlier
   * Subtotal's, kept because "Replace current subtotals" was cleared). They are
   * passed through where they are, never grouped as detail, and end the run
   * they sit in, so the new subtotals nest inside the existing ones as Excel's
   * do (XLS-A064).
   */
  existingSummaryRows?: ReadonlySet<number>;
  /**
   * Excel's "Page break between groups": report where a new page starts after
   * each group (XLS-A064). The offsets land in `pageBreakOffsets`.
   */
  pageBreaks?: boolean;
}

export interface SubtotalPlan {
  /** The rebuilt matrix (original rows + inserted summary rows). */
  rows: string[][];
  /** Outline row groups, 0-based RELATIVE to the first data row. */
  groups: OutlineGroup[];
  /** 0-based offsets within `rows` of every inserted summary/grand-total row. */
  summaryRowOffsets: number[];
  /** How many summary rows (incl. grand total) were inserted. */
  insertedCount: number;
  /**
   * Parallel to `rows`: for each output row, the 0-based index into the ORIGINAL
   * `dataRows` it came from, or null for an inserted summary / grand-total row.
   * Lets the caller map columns OUTSIDE the rebuilt range (cells in the data
   * rows but outside the value/group columns) onto their new positions so they
   * are preserved + shifted, exactly as Excel keeps a row's other columns when
   * it inserts subtotal rows.
   */
  rowSources: (number | null)[];
  /**
   * With `pageBreaks`: the 0-based offsets within `rows` BEFORE which a new
   * page starts (the first row of every group but the first), sorted.
   */
  pageBreakOffsets: number[];
}

/** A1 cell address for a 0-based row/col (absolute, no $). */
function a1(row: number, col: number): string {
  return `${columnToLetters(col)}${row + 1}`;
}

/**
 * Build a SUBTOTAL formula over a contiguous block of a single column.
 * `topRow`/`bottomRow` are ABSOLUTE 0-based sheet rows (inclusive). Always emits
 * a range (`=SUBTOTAL(9,B2:B5)`) even for one row, matching Excel.
 */
function subtotalFormula(
  fnNum: number,
  col: number,
  topRow: number,
  bottomRow: number,
): string {
  return `=SUBTOTAL(${fnNum},${a1(topRow, col)}:${a1(bottomRow, col)})`;
}

/**
 * Plan the Subtotal command over `dataRows` (header excluded). Returns the
 * rebuilt matrix, the outline groups, and the summary-row offsets. When the
 * region is empty, or `valueColumns` is empty, the matrix is returned unchanged
 * with no groups (the caller treats that as a no-op).
 *
 * Grouping is by ADJACENT equal values in `groupColumn` (so the caller must sort
 * first, as Excel requires). Each group's summary row carries a SUBTOTAL formula
 * for every value column over that group's detail rows; the optional grand-total
 * row sums every value column over the WHOLE region (and, thanks to SUBTOTAL's
 * ignore-nested-subtotal rule, does not double-count the per-group summaries).
 */
export function planSubtotals(
  dataRows: readonly (readonly string[])[],
  options: SubtotalPlanOptions,
): SubtotalPlan {
  const {
    groupColumn,
    valueColumns,
    function: fn,
    firstDataRow,
    firstDataColumn = 0,
    labelTemplate,
    grandTotalLabel,
    grandTotal = true,
    summaryBelow = true,
    existingSummaryRows,
    pageBreaks = false,
  } = options;

  const fnNum = FUNCTION_NUM[fn] ?? FUNCTION_NUM.sum;
  // Column count: widen to cover the grouping + value columns even if a row is
  // ragged, so a summary row always has a slot for its formulas + label.
  let width = 0;
  for (const row of dataRows) width = Math.max(width, row.length);
  width = Math.max(width, groupColumn + 1, ...valueColumns.map((c) => c + 1));

  if (dataRows.length === 0 || valueColumns.length === 0) {
    return {
      rows: dataRows.map((r) => [...r]),
      groups: [],
      summaryRowOffsets: [],
      insertedCount: 0,
      rowSources: dataRows.map((_, i) => i),
      pageBreakOffsets: [],
    };
  }

  // Pad a (possibly ragged) detail row to `width`, so column writes never go OOB.
  const pad = (row: readonly string[]): string[] => {
    const out = row.slice(0, width);
    while (out.length < width) out.push("");
    return out;
  };

  // The grouping-column value of a row (empty when out of range). Trimmed only
  // for the EQUALITY check (so trailing spaces don't split a group), but the
  // original value is echoed in the summary label.
  const groupValueOf = (row: readonly string[]): string =>
    (row[groupColumn] ?? "").trim();
  const isExisting = (index: number) => existingSummaryRows?.has(index) === true;

  // 1) Find the runs of equal grouping values. An earlier Subtotal's summary
  //    row ends the run in front of it and belongs to no run.
  const runs: Array<{ start: number; end: number; label: string }> = [];
  let runStart = -1;
  const closeRun = (end: number) => {
    if (runStart < 0) return;
    runs.push({ start: runStart, end, label: dataRows[runStart][groupColumn] ?? "" });
    runStart = -1;
  };
  for (let i = 0; i < dataRows.length; i += 1) {
    if (isExisting(i)) {
      closeRun(i - 1);
      continue;
    }
    if (runStart >= 0 && groupValueOf(dataRows[i]) !== groupValueOf(dataRows[runStart])) {
      closeRun(i - 1);
    }
    if (runStart < 0) runStart = i;
  }
  closeRun(dataRows.length - 1);

  // 2) Lay the output out as entries first: the SUBTOTAL formulas address the
  //    ABSOLUTE rows the detail will occupy, which are only known once every
  //    summary row (above or below its group) has its place.
  type Entry =
    | { kind: "row"; source: number }
    | { kind: "summary"; group: number }
    | { kind: "grand" };
  const entries: Entry[] = [];
  // Where each group's FIRST output row sits (its summary when above), for the
  // page breaks between groups.
  const groupFirstEntry: number[] = [];
  if (grandTotal && !summaryBelow) entries.push({ kind: "grand" });
  let next = 0;
  runs.forEach((run, group) => {
    // Rows before this run that belong to no run are existing summary rows;
    // they stay exactly where they were.
    for (; next < run.start; next += 1) entries.push({ kind: "row", source: next });
    groupFirstEntry.push(entries.length);
    if (!summaryBelow) entries.push({ kind: "summary", group });
    for (let i = run.start; i <= run.end; i += 1) entries.push({ kind: "row", source: i });
    next = run.end + 1;
    if (summaryBelow) entries.push({ kind: "summary", group });
  });
  for (; next < dataRows.length; next += 1) entries.push({ kind: "row", source: next });
  if (grandTotal && summaryBelow) entries.push({ kind: "grand" });

  const offsetOfSource = new Map<number, number>();
  entries.forEach((entry, offset) => {
    if (entry.kind === "row") offsetOfSource.set(entry.source, offset);
  });

  const out: string[][] = [];
  const groups: OutlineGroup[] = [];
  const summaryRowOffsets: number[] = [];
  // Parallel to `out`: the original dataRows index each output row came from, or
  // null for an inserted summary / grand-total row.
  const rowSources: (number | null)[] = [];
  entries.forEach((entry, offset) => {
    if (entry.kind === "row") {
      out.push(pad(dataRows[entry.source]));
      rowSources.push(entry.source);
      return;
    }
    const summary = new Array<string>(width).fill("");
    let top: number;
    let bottom: number;
    if (entry.kind === "summary") {
      const run = runs[entry.group];
      top = offsetOfSource.get(run.start)!;
      bottom = offsetOfSource.get(run.end)!;
      summary[groupColumn] = labelTemplate
        ? labelTemplate.replace("{group}", run.label)
        : run.label;
      // Outline group (relative to the first data row): the body is the detail
      // run; the summary row sits beside it, below or above as the sheet's
      // outline setting says. A single detail row still groups, as in Excel.
      groups.push({ axis: "row", start: top, end: bottom, collapsed: false });
    } else {
      // Grand total over the WHOLE list. SUBTOTAL skips the nested group
      // summaries (and an earlier Subtotal's rows), so this is the true total.
      top = summaryBelow ? 0 : 1;
      bottom = summaryBelow ? entries.length - 2 : entries.length - 1;
      summary[groupColumn] = grandTotalLabel ?? "";
    }
    for (const col of valueColumns) {
      summary[col] = subtotalFormula(fnNum, firstDataColumn + col, firstDataRow + top, firstDataRow + bottom);
    }
    summaryRowOffsets.push(offset);
    out.push(summary);
    rowSources.push(null);
  });

  // Excel's "Page break between groups": a new page for every group after the
  // first; the grand total stays on the last group's page.
  const pageBreakOffsets = pageBreaks ? groupFirstEntry.slice(1) : [];

  return {
    rows: out,
    groups,
    summaryRowOffsets,
    insertedCount: summaryRowOffsets.length,
    rowSources,
    pageBreakOffsets,
  };
}

/** The SUBTOTAL function_num for a function key (exported for the UI label/hint). */
export function subtotalFunctionNum(fn: SubtotalFunction): number {
  return FUNCTION_NUM[fn] ?? FUNCTION_NUM.sum;
}
