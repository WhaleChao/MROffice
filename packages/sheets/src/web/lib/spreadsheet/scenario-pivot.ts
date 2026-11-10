/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Scenario Manager's Summary… ▸ Scenario PivotTable report (Excel
 * `Scenarios.CreateSummary ReportType:=xlPivotTable`): a PivotTable on a new
 * worksheet whose cache is the scenarios themselves — no worksheet range
 * (OOXML `cacheSource type="scenario"`).
 *
 * Microsoft Support, "Switch between various sets of values by using
 * scenarios": result cells are required for this report type; changing and
 * result cells are named by their references; the report does not
 * recalculate when a scenario changes. Excel's report lays the cache out as:
 *
 * - a page field "<changing cells> by" whose items are the people who made
 *   the scenarios (shown as (All));
 * - a row field "<changing cells>" whose items are the scenario names;
 * - one data field per result cell, named by its reference, holding that
 *   cell's value with the scenario shown.
 *
 * Scenarios over different sets of changing cells get one such pair of
 * fields per set, a scenario being blank in the other sets' fields.
 *
 * Pure: the host computes each result value with the scenario applied (the
 * evaluator the Scenario summary report already uses) and supplies the
 * labels; the recipe is placed on its sheet by the host's PivotTable helpers.
 */

import { parseCellRef, type CellValue } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import {
  createPivotCacheSnapshot,
  pivotRangeToA1,
  type PivotConfig,
  type PivotInferredField,
  type PivotMatrix,
  type PivotNativeScalar,
} from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import {
  formatScenarioChangingCells,
  type Scenario,
} from "./scenarios";

export interface ScenarioPivotInput {
  scenarios: readonly Scenario[];
  /** Canonical A1 refs ("C12"), in the order the Result cells box named them. */
  resultRefs: readonly string[];
  /** A result cell's value with the scenario's changing cells applied. */
  resultValue: (scenario: Scenario, ref: string) => CellValue;
  /** The number format of a result cell (its data field takes it). */
  numberFormat: (ref: string) => string | undefined;
  /** Who the "by" field names for each scenario. */
  creator: (scenario: Scenario) => string;
  labels: {
    /** The page field's caption for a set of changing cells ("$B$2:$B$4 by"). */
    by: (cells: string) => string;
  };
}

export interface ScenarioPivotRecipe {
  /** The PivotTable recipe (no output yet: the host places it). */
  config: PivotConfig;
  /** The cache as a display matrix (header row first). */
  matrix: PivotMatrix;
}

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

function absoluteRef(ref: string): string | null {
  const parsed = parseCellRef(ref);
  return parsed ? `$${columnLetters(parsed.col)}$${parsed.row + 1}` : null;
}

function textScalar(text: string): PivotNativeScalar {
  return text === "" ? { kind: "m", raw: "" } : { kind: "s", raw: text };
}

/**
 * A computed result as a typed cache value. Each result is a Sum data field,
 * which only numbers feed: a text, logical or error result shows blank in the
 * report either way, so it is cached as a missing value (and the workbook can
 * still be written: a summed cache field holds numbers only).
 */
function resultScalar(value: CellValue): PivotNativeScalar {
  if (typeof value === "number" && Number.isFinite(value)) {
    return { kind: "n", raw: String(value) };
  }
  return { kind: "m", raw: "" };
}

function scalarDisplay(value: PivotNativeScalar): string {
  if (value.kind === "m") return "";
  if (value.kind === "b") return value.raw === "1" ? "TRUE" : "FALSE";
  return value.raw;
}

/**
 * The Scenario PivotTable report's recipe, or null when there is nothing to
 * report (no scenarios, or no readable result cell: Excel requires result
 * cells for this report).
 */
export function buildScenarioPivotRecipe(
  input: ScenarioPivotInput,
): ScenarioPivotRecipe | null {
  const { scenarios } = input;
  if (scenarios.length === 0) return null;
  const results: Array<{ ref: string; label: string }> = [];
  const seenResults = new Set<string>();
  for (const raw of input.resultRefs) {
    const label = absoluteRef(raw);
    const parsed = parseCellRef(raw);
    if (!label || !parsed || seenResults.has(label)) continue;
    seenResults.add(label);
    results.push({ ref: `${columnLetters(parsed.col)}${parsed.row + 1}`, label });
  }
  if (results.length === 0) return null;

  // One "by" + row field pair per set of changing cells, first-seen order.
  const sets: string[] = [];
  const setOf = new Map<Scenario, number>();
  for (const scenario of scenarios) {
    const cells = formatScenarioChangingCells(scenario.cells);
    let index = sets.indexOf(cells);
    if (index < 0) {
      sets.push(cells);
      index = sets.length - 1;
    }
    setOf.set(scenario, index);
  }

  // Field names are unique, as a PivotTable cache's are.
  const taken = new Set<string>();
  const unique = (name: string): string => {
    let candidate = name;
    for (let suffix = 2; taken.has(candidate.toLowerCase()); suffix += 1) {
      candidate = `${name} ${suffix}`;
    }
    taken.add(candidate.toLowerCase());
    return candidate;
  };
  const header: string[] = [];
  for (const cells of sets) {
    header.push(unique(input.labels.by(cells)), unique(cells));
  }
  const resultNames = results.map((result) => unique(result.label));
  header.push(...resultNames);

  const scalarMatrix: PivotNativeScalar[][] = [header.map(textScalar)];
  for (const scenario of scenarios) {
    const own = setOf.get(scenario) ?? 0;
    const row: PivotNativeScalar[] = [];
    sets.forEach((_, index) => {
      if (index === own) {
        row.push(textScalar(input.creator(scenario)), textScalar(scenario.name));
      } else {
        row.push(textScalar(""), textScalar(""));
      }
    });
    for (const result of results) {
      row.push(resultScalar(input.resultValue(scenario, result.ref)));
    }
    scalarMatrix.push(row);
  }
  const matrix: PivotMatrix = scalarMatrix.map((row) => row.map(scalarDisplay));
  const width = header.length;
  const fields: PivotInferredField[] = header.map((name, columnIndex) => ({
    name,
    columnIndex,
    type: columnIndex >= sets.length * 2 ? "number" : "text",
  }));
  const cacheSnapshot = createPivotCacheSnapshot(fields, matrix, scalarMatrix);
  if (!cacheSnapshot) return null;

  const config: PivotConfig = {
    // Only the cache's shape: a scenario report reads no worksheet range.
    source: pivotRangeToA1({ r1: 0, c1: 0, r2: matrix.length - 1, c2: width - 1 }),
    cacheSource: "scenario",
    cacheSnapshot,
    rows: sets.map((_, index) => index * 2 + 1),
    columns: [],
    filters: sets.map((_, index) => ({ colIndex: index * 2, selected: [] })),
    values: results.map((result, index) => {
      const numberFormat = input.numberFormat(result.ref);
      return {
        colIndex: sets.length * 2 + index,
        agg: "sum" as const,
        caption: resultNames[index],
        ...(numberFormat && numberFormat !== "General" ? { numberFormat } : {}),
      };
    }),
    grouping: [],
    calculatedFields: [],
    // Excel's classic report: each field under its own caption, and no
    // totals adding alternative scenarios together.
    reportLayout: "tabular",
    grandTotals: "none",
  };
  return { config, matrix };
}
