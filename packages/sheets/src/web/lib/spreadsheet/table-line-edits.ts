/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Insert / delete rows and columns OF A TABLE (XLS-A078): the right-click
// menu's Insert ▸ Table Rows Above / Table Columns to the Left and Delete ▸
// Table Rows / Table Columns.
//
// Microsoft Support, "Resize a table by adding or removing rows and columns":
// a table row or column is inserted or deleted inside the table only — the
// cells beside the table do not move — and the table grows or shrinks with
// it. That is Insert / Delete Cells with a shift over the table's own span
// (references follow the moved cells), with the table's range, its column
// list and its per-column settings adjusted to match. A new row of a
// calculated column takes the column formula; a new column gets the next free
// default name ("Column1", "列1") in its header.
// ---------------------------------------------------------------------------

import type { DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { shiftFormula } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import {
  cellShiftedSheet,
  type CellShiftButtonSheet,
} from "./row-column-buttons";

export interface TableLineEdit {
  readonly axis: "row" | "col";
  readonly op: "insert" | "delete";
  /** The first row / column of the edit (an insert goes above / left of it). */
  readonly at: number;
  readonly count: number;
}

export type TableLineEditRefusal = "outside" | "last-line" | "native" | "shift";

type TableSheet = CellShiftButtonSheet;
const tablesOf = (sheet: unknown): DataTable[] =>
  (sheet as { dataTables?: DataTable[] }).dataTables ?? [];

function shiftOffsets<T>(
  map: Record<number, T> | undefined,
  at: number,
  delta: number,
): Record<number, T> | undefined {
  if (!map) return map;
  const out: Record<number, T> = {};
  for (const [key, value] of Object.entries(map)) {
    const offset = Number(key);
    if (delta < 0 && offset >= at && offset < at - delta) continue;
    out[offset >= at ? offset + delta : offset] = value;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

function withMaps(table: DataTable, at: number, delta: number): DataTable {
  const next: DataTable = { ...table };
  for (const key of ["totals", "calculatedFormulas", "totalFormulas", "totalLabels"] as const) {
    const shifted = shiftOffsets(table[key] as Record<number, never> | undefined, at, delta);
    if (shifted) (next as unknown as Record<string, unknown>)[key] = shifted;
    else delete (next as unknown as Record<string, unknown>)[key];
  }
  return next;
}

/**
 * The sheet with table `tableId` given `edit`, or why not. `defaultColumnName`
 * spells a new column's name for its 1-based number.
 */
export function tableLinesEditedSheet<S extends TableSheet>(
  sheet: S,
  workbook: readonly S[],
  tableId: string,
  edit: TableLineEdit,
  bounds: { maxRows: number; maxCols: number },
  defaultColumnName: (index: number) => string,
): { sheet: S; followOthers: (candidate: S) => S } | TableLineEditRefusal {
  const tables = tablesOf(sheet);
  const table = tables.find((candidate) => candidate.id === tableId);
  if (!table) return "outside";
  if (
    table.nativeResizeBlocker &&
    table.nativeResizeBlocker !== "activeFilter" &&
    table.nativeResizeBlocker !== "sortState"
  ) {
    return "native";
  }
  const { r1, c1, r2, c2 } = table.range;
  const dataTop = r1 + (table.headerRow === false ? 0 : 1);
  const bottom = r2 + (table.totalsRow ? 1 : 0);
  const width = c2 - c1 + 1;
  let range: { r1: number; c1: number; r2: number; c2: number };
  let direction: "down" | "up" | "right" | "left";
  let nextTable: DataTable;
  let count = Math.max(1, Math.floor(edit.count));
  // The new columns of an insert, whose names go into their header cells.
  const added: Array<{ id: number; name: string }> = [];
  if (edit.axis === "row") {
    if (edit.op === "insert") {
      if (edit.at < dataTop || edit.at > r2 + 1) return "outside";
      range = { r1: edit.at, r2: edit.at + count - 1, c1, c2 };
      direction = "down";
      nextTable = { ...table, range: { r1, c1, r2: r2 + count, c2 } };
    } else {
      if (edit.at < dataTop || edit.at > r2) return "outside";
      // A table keeps at least one data row (Excel empties the last one
      // rather than deleting it).
      count = Math.min(count, r2 - edit.at + 1, r2 - dataTop);
      if (count < 1) return "last-line";
      range = { r1: edit.at, r2: edit.at + count - 1, c1, c2 };
      direction = "up";
      nextTable = { ...table, range: { r1, c1, r2: r2 - count, c2 } };
    }
  } else if (edit.op === "insert") {
    if (edit.at < c1 || edit.at > c2 + 1) return "outside";
    range = { r1, r2: bottom, c1: edit.at, c2: edit.at + count - 1 };
    direction = "right";
    const offset = edit.at - c1;
    const names = new Set((table.columns ?? []).map((column) => column.name.toUpperCase()));
    let nextId = Math.max(0, ...(table.columns ?? []).map((column) => column.id)) + 1;
    let number = 1;
    for (let i = 0; i < count; i += 1) {
      while (names.has(defaultColumnName(number).toUpperCase())) number += 1;
      const name = defaultColumnName(number);
      names.add(name.toUpperCase());
      added.push({ id: nextId, name });
      nextId += 1;
    }
    const columns = table.columns ? [...table.columns] : undefined;
    columns?.splice(offset, 0, ...added);
    nextTable = {
      ...withMaps(table, offset, count),
      range: { r1, c1, r2, c2: c2 + count },
      ...(columns ? { columns } : {}),
    };
  } else {
    if (edit.at < c1 || edit.at > c2) return "outside";
    count = Math.min(count, c2 - edit.at + 1, width - 1);
    if (count < 1) return "last-line";
    range = { r1, r2: bottom, c1: edit.at, c2: edit.at + count - 1 };
    direction = "left";
    const offset = edit.at - c1;
    const columns = table.columns ? [...table.columns] : undefined;
    columns?.splice(offset, count);
    nextTable = {
      ...withMaps(table, offset, -count),
      range: { r1, c1, r2, c2: c2 - count },
      ...(columns ? { columns } : {}),
    };
  }
  // The shift runs with this table out of the way (it may not resize a
  // table), then the table comes back at its new size.
  const without = { ...sheet, dataTables: tables.filter((candidate) => candidate.id !== tableId) } as unknown as S;
  const others = workbook.map((candidate) => (candidate.id === sheet.id ? without : candidate));
  const shifted = cellShiftedSheet(without, others, range, direction, bounds);
  if (!shifted) return "shift";
  const cells = { ...shifted.sheet.cells } as Record<string, { value: string; formula?: string }>;
  if (edit.axis === "row" && edit.op === "insert") {
    for (const [key, formula] of Object.entries(table.calculatedFormulas ?? {})) {
      const col = c1 + Number(key);
      for (let row = edit.at; row < edit.at + count; row += 1) {
        const text = shiftFormula(formula, row - dataTop, 0);
        cells[`${row}:${col}`] = { ...(cells[`${row}:${col}`] ?? {}), value: text, formula: text };
      }
    }
  }
  if (added.length > 0 && table.headerRow !== false) {
    added.forEach((column, index) => {
      const key = `${r1}:${edit.at + index}`;
      cells[key] = { ...(cells[key] ?? {}), value: column.name };
    });
  }
  const order = tables.map((candidate) => candidate.id);
  const nextTables = [...tablesOf(shifted.sheet), nextTable].sort(
    (a, b) => order.indexOf(a.id) - order.indexOf(b.id),
  );
  // A slicer on one of this table's columns stays on that column.
  const slicers = (shifted.sheet as { slicers?: Array<{ column: number; source?: { kind?: string; tableId?: string; tableColumnId?: number } }> }).slicers;
  const nextSlicers = slicers?.map((slicer) => {
    if (slicer.source?.kind !== "table" || slicer.source.tableId !== tableId) return slicer;
    const offset = nextTable.columns?.findIndex((column) => column.id === slicer.source?.tableColumnId) ?? -1;
    return offset >= 0 ? { ...slicer, column: nextTable.range.c1 + offset } : slicer;
  });
  const nextSheet = {
    ...shifted.sheet,
    cells,
    dataTables: nextTables,
    ...(nextSlicers ? { slicers: nextSlicers } : {}),
  } as unknown as S;
  return { sheet: nextSheet, followOthers: shifted.followOthers };
}
