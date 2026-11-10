/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Formulas ▸ Calculation Options ▸ Automatic Except for Data Tables (xlsx
 * calcMode="autoNoTable"). Excel recalculates every formula on each edit
 * except a What-If Data Table's results, which keep the values of the last
 * calculation until Calculate Now (F9) or Calculate Sheet (Shift+F9); the
 * formulas reading those results read the kept values too (live audit
 * S-143).
 *
 * The calculation reads a held table's results as formulas of the values
 * they last had, so it does not recalculate the table, a cell reading it
 * reads what the table shows, and the grid shows the held values as the
 * calculation's answer for the table's cells.
 */

import { isFormulaError, type CellValue } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { whatIfTableBlocks, whatIfTableResultKeys } from "@mrnote/core/web/lib/spreadsheet/what-if-table-definition.ts";

/** A sheet as the hold reads and rewrites it. */
export interface WhatIfHoldSheet {
  readonly id: string;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly whatIfTables?: unknown;
  readonly cells: Readonly<Record<string, unknown>>;
}

/**
 * One sheet's held results. `values` is null while the sheet waits for the
 * recalculation F9 asked for (its tables calculate live until then).
 * `tables` is the definition list the values were read for: a table added,
 * moved or removed is calculated afresh.
 */
export interface HeldWhatIfSheet {
  readonly tables: unknown;
  readonly values: ReadonlyMap<string, CellValue> | null;
}

/** Each sheet's held results, by sheet id. */
export type HeldWhatIfTables = ReadonlyMap<string, HeldWhatIfSheet>;

/** The result cells ("row:col") of a sheet's What-If Data Tables. */
export function whatIfTableResultCells(sheet: WhatIfHoldSheet): string[] {
  return whatIfTableBlocks(sheet.whatIfTables, sheet).flatMap(whatIfTableResultKeys);
}

/**
 * The values a sheet's table results show, read cell by cell; null when one
 * is not a plain value (an array, a picture), which is then calculated live
 * rather than held in part.
 */
export function readWhatIfTableResults(
  sheet: WhatIfHoldSheet,
  read: (row: number, col: number) => unknown,
): Map<string, CellValue> | null {
  const values = new Map<string, CellValue>();
  for (const key of whatIfTableResultCells(sheet)) {
    const [row, col] = key.split(":").map(Number);
    const value = read(row, col);
    if (
      value !== null &&
      typeof value !== "number" &&
      typeof value !== "string" &&
      typeof value !== "boolean" &&
      !isFormulaError(value)
    ) {
      return null;
    }
    values.set(key, value as CellValue);
  }
  return values;
}

/**
 * A held result as the cell the workbook calculates in its place: a formula
 * of its value alone, so the calculation still answers the cell (the grid
 * shows a table result from the calculation's answer) and a reader of it
 * reads the value.
 */
function heldCell(stored: unknown, value: CellValue): Record<string, unknown> {
  const base = { ...((stored ?? {}) as Record<string, unknown>) };
  delete base.stringValue;
  delete base.runs;
  delete base.kind;
  const formula =
    typeof value === "number"
      ? `=${String(value)}`
      : typeof value === "boolean"
        ? value ? "=TRUE" : "=FALSE"
        : isFormulaError(value)
          ? `=${value.code}`
          : `="${(value ?? "").replace(/"/g, '""')}"`;
  return { ...base, value: formula, formula };
}

/** A calculation binding as the hold reads and rewrites it. */
export interface WhatIfHoldBinding {
  readonly sheet: WhatIfHoldSheet;
  readonly input: { getRawCell(row: number, col: number): string | null };
}

const heldBindings = new WeakMap<object, { held: HeldWhatIfSheet; binding: object }>();

/**
 * The bindings the calculation runs over: each sheet whose tables are held
 * reads its results as formulas of their held values. The grid keeps reading
 * the table's own =TABLE() and shows the calculation's answer for it, which
 * is the held value. The same binding and hold give the same object, so the
 * calculation sees no change until the hold does.
 */
export function bindingsWithHeldWhatIfTables<B extends WhatIfHoldBinding>(
  bindings: readonly B[],
  held: HeldWhatIfTables,
): readonly B[] {
  let changed = false;
  const out = bindings.map((binding) => {
    const entry = held.get(binding.sheet.id);
    const values = entry?.values;
    if (!entry || !values || entry.tables !== binding.sheet.whatIfTables) return binding;
    const known = heldBindings.get(binding);
    if (known?.held === entry) {
      changed = true;
      return known.binding as B;
    }
    const keys = whatIfTableResultCells(binding.sheet);
    if (keys.length === 0 || keys.some((key) => !values.has(key))) return binding;
    const formulas = new Map<string, string>();
    const cells: Record<string, unknown> = { ...binding.sheet.cells };
    for (const key of keys) {
      const cell = heldCell(binding.sheet.cells[key], values.get(key)!);
      formulas.set(key, cell.formula as string);
      cells[key] = cell;
    }
    const read = binding.input.getRawCell;
    const next = {
      ...binding,
      sheet: { ...binding.sheet, cells },
      input: { ...binding.input, getRawCell: (row: number, col: number) => formulas.get(`${row}:${col}`) ?? read(row, col) },
    } as B;
    heldBindings.set(binding, { held: entry, binding: next });
    changed = true;
    return next;
  });
  return changed ? out : bindings;
}
