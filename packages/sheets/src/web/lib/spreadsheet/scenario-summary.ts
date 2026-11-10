/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Scenario Manager's Summary… report (Excel "Scenario Summary"): a new
 * worksheet that lists every scenario side by side against the current values.
 *
 * Microsoft Support, "Switch between various sets of values by using
 * scenarios": the report lists the scenarios side by side; "The Current Values
 * column represents the values of changing cells when you create the Scenario
 * Summary Report"; "The cells that changed for each scenario are highlighted in
 * gray"; "Excel automatically adds Grouping levels, which expand and collapse
 * the view"; the report does not recalculate afterwards; result cells are
 * optional for this report type. The layout follows Excel's: the title in B2,
 * "Current Values:" and one column per scenario in row 3, a "Changing Cells:"
 * block and a "Result Cells:" block with each cell's absolute address in
 * column C, and the notes beneath.
 *
 * Pure: the host passes the values (current, and each result cell recomputed
 * with the scenario applied) and the localized labels.
 */

import { isFormulaError, parseCellRef, type CellValue } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import type { OutlineGroup } from "@mrnote/core/web/lib/spreadsheet/outline-groups.ts";
import type { Scenario } from "./scenarios";

// A type alias (not an interface) so it assigns to the sheet's cell map.
export type ScenarioSummaryCell = {
  value: string;
  kind?: "text" | "number";
  style?: {
    bold?: boolean;
    italic?: boolean;
    backgroundColor?: string;
    textAlign?: "left" | "center" | "right";
    numberFormat?: string;
  };
};

export interface ScenarioSummaryLabels {
  title: string;
  currentValues: string;
  changingCells: string;
  resultCells: string;
  /** The notes block, one line per row. */
  notes: readonly string[];
}

export interface ScenarioSummaryInput {
  scenarios: readonly Scenario[];
  /** Canonical A1 refs ("C12"), in the order the Result cells box named them. */
  resultRefs: readonly string[];
  /** A cell's current value (a formula's result). */
  currentValue: (ref: string) => CellValue;
  /** A result cell's value with the scenario's changing cells applied. */
  resultValue: (scenario: Scenario, ref: string) => CellValue;
  /** The number format of the cell the row reports (copied onto its values). */
  numberFormat: (ref: string) => string | undefined;
  labels: ScenarioSummaryLabels;
}

export interface ScenarioSummarySheet {
  cells: Record<string, ScenarioSummaryCell>;
  rowCount: number;
  columnCount: number;
  columnWidths: Record<string, number>;
  outlineGroups: OutlineGroup[];
}

/** Result cells one report lists (a guard: a whole column is not a report). */
export const MAX_SCENARIO_SUMMARY_RESULT_CELLS = 256;

/** Excel's header band fill and the gray it marks changed cells with. */
const HEADER_FILL = "#D9D9D9";
const CHANGED_FILL = "#C0C0C0";
const MIN_ROWS = 30;
const MIN_COLUMNS = 10;

/** Column letters for a 0-based index. */
function columnLetters(index: number): string {
  let label = "";
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function absoluteRef(ref: string): string {
  const parsed = parseCellRef(ref);
  return parsed ? `$${columnLetters(parsed.col)}$${parsed.row + 1}` : ref;
}

function canonicalRef(ref: string): string | null {
  const parsed = parseCellRef(ref);
  return parsed ? `${columnLetters(parsed.col)}${parsed.row + 1}` : null;
}

/** A computed value as a stored cell: numbers keep the source's format. */
function valueCell(value: CellValue, numberFormat: string | undefined): ScenarioSummaryCell {
  if (value === null || value === undefined) return { value: "" };
  if (isFormulaError(value)) return { value: value.code };
  if (typeof value === "boolean") return { value: value ? "TRUE" : "FALSE" };
  if (typeof value === "number") {
    return Number.isFinite(value)
      ? {
          value: String(value),
          kind: "number",
          ...(numberFormat ? { style: { numberFormat } } : {}),
        }
      : { value: "#NUM!" };
  }
  return { value: String(value), kind: "text" };
}

/** A stored scenario value (canonical number or text as entered) as a cell. */
function scenarioValueCell(raw: string, numberFormat: string | undefined): ScenarioSummaryCell {
  const text = raw.startsWith("=") ? raw.replace(/^=+/, "") : raw;
  if (text.startsWith("'")) return { value: text.slice(1), kind: "text" };
  const trimmed = text.trim();
  if (trimmed !== "" && /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(trimmed)) {
    const number = Number(trimmed);
    if (Number.isFinite(number)) {
      return {
        value: String(number),
        kind: "number",
        ...(numberFormat ? { style: { numberFormat } } : {}),
      };
    }
  }
  if (/^(TRUE|FALSE)$/i.test(trimmed)) return { value: trimmed.toUpperCase() };
  return { value: text, kind: "text" };
}

export function buildScenarioSummarySheet(input: ScenarioSummaryInput): ScenarioSummarySheet {
  const { scenarios, labels } = input;
  const cells: Record<string, ScenarioSummaryCell> = {};
  const put = (row: number, col: number, cell: ScenarioSummaryCell) => {
    cells[`${row}:${col}`] = cell;
  };
  const withStyle = (
    cell: ScenarioSummaryCell,
    style: NonNullable<ScenarioSummaryCell["style"]>,
  ): ScenarioSummaryCell => ({ ...cell, style: { ...(cell.style ?? {}), ...style } });

  // Changing cells: the union of every scenario's cells, first-seen order.
  const changing: string[] = [];
  const seen = new Set<string>();
  for (const scenario of scenarios) {
    for (const cell of scenario.cells) {
      const ref = canonicalRef(cell.ref);
      if (!ref || seen.has(ref)) continue;
      seen.add(ref);
      changing.push(ref);
    }
  }
  const results: string[] = [];
  const seenResult = new Set<string>();
  for (const raw of input.resultRefs) {
    const ref = canonicalRef(raw);
    if (!ref || seenResult.has(ref)) continue;
    seenResult.add(ref);
    results.push(ref);
  }

  const labelCol = 1; // B
  const refCol = 2; // C
  const currentCol = 3; // D
  const firstScenarioCol = 4; // E
  const lastCol = firstScenarioCol + Math.max(0, scenarios.length - 1);

  // Header band: title (row 2), Current Values + scenario names (row 3).
  let row = 1;
  for (let col = labelCol; col <= lastCol; col += 1) {
    put(row, col, { value: "", style: { backgroundColor: HEADER_FILL } });
    put(row + 1, col, { value: "", style: { backgroundColor: HEADER_FILL } });
  }
  put(row, labelCol, {
    value: labels.title,
    kind: "text",
    style: { bold: true, backgroundColor: HEADER_FILL },
  });
  row += 1;
  put(row, currentCol, {
    value: labels.currentValues,
    kind: "text",
    style: { bold: true, backgroundColor: HEADER_FILL, textAlign: "right" },
  });
  scenarios.forEach((scenario, index) => {
    put(row, firstScenarioCol + index, {
      value: scenario.name,
      kind: "text",
      style: { bold: true, backgroundColor: HEADER_FILL, textAlign: "right" },
    });
  });
  row += 1;
  // The comment each scenario carries ("Created by … on …"), when any does.
  const commentRow = scenarios.some((scenario) => (scenario.comment ?? "").trim() !== "")
    ? row
    : null;
  if (commentRow !== null) {
    scenarios.forEach((scenario, index) => {
      const comment = (scenario.comment ?? "").trim();
      if (comment) {
        put(commentRow, firstScenarioCol + index, {
          value: comment,
          kind: "text",
          style: { italic: true, textAlign: "right" },
        });
      }
    });
    row += 1;
  }

  const outlineGroups: OutlineGroup[] = [];
  put(row, labelCol, { value: labels.changingCells, kind: "text", style: { bold: true } });
  row += 1;
  const changingStart = row;
  for (const ref of changing) {
    const format = input.numberFormat(ref);
    put(row, refCol, { value: absoluteRef(ref), kind: "text" });
    put(row, currentCol, valueCell(input.currentValue(ref), format));
    scenarios.forEach((scenario, index) => {
      const own = scenario.cells.find((cell) => canonicalRef(cell.ref) === ref);
      put(
        row,
        firstScenarioCol + index,
        own
          ? withStyle(scenarioValueCell(own.value, format), { backgroundColor: CHANGED_FILL })
          : valueCell(input.currentValue(ref), format),
      );
    });
    row += 1;
  }
  if (row > changingStart) {
    outlineGroups.push({ axis: "row", start: changingStart, end: row - 1, collapsed: false });
  }

  if (results.length > 0) {
    put(row, labelCol, { value: labels.resultCells, kind: "text", style: { bold: true } });
    row += 1;
    const resultStart = row;
    for (const ref of results) {
      const format = input.numberFormat(ref);
      put(row, refCol, { value: absoluteRef(ref), kind: "text" });
      put(row, currentCol, valueCell(input.currentValue(ref), format));
      scenarios.forEach((scenario, index) => {
        put(row, firstScenarioCol + index, valueCell(input.resultValue(scenario, ref), format));
      });
      row += 1;
    }
    outlineGroups.push({ axis: "row", start: resultStart, end: row - 1, collapsed: false });
  }

  row += 1;
  for (const line of labels.notes) {
    put(row, labelCol, { value: line, kind: "text" });
    row += 1;
  }

  // The scenario columns fold together under one column level.
  if (scenarios.length > 0) {
    outlineGroups.push({ axis: "col", start: currentCol, end: lastCol, collapsed: false });
  }

  const columnWidths: Record<string, number> = { "0": 16, [String(labelCol)]: 24, [String(refCol)]: 72 };
  for (let col = currentCol; col <= lastCol; col += 1) columnWidths[String(col)] = 96;

  return {
    cells,
    rowCount: Math.max(MIN_ROWS, row + 1),
    columnCount: Math.max(MIN_COLUMNS, lastCol + 2),
    columnWidths,
    outlineGroups,
  };
}
