/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeDataTables, normalizeDataTableFilterColor, type DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { normalizeConditionFilter, buildRowPredicate, type ConditionFilter } from "@mrnote/core/web/lib/spreadsheet/filter-conditions.ts";
import { filterValueKey, filterValueKeys, normalizeFillColor } from "./sort-filter";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";

export interface FilterVisibilitySource {
  rowCount: number; columnCount: number; autoFilterRef?: string;
  filters?: Record<string, string[]>; colorFilters?: Record<string, string[]>;
  conditionFilters?: Record<string, ConditionFilter>; dataTables?: DataTable[];
}
export interface FilterVisibilityCell {
  display: string; numeric: number | null; fill?: string; font?: string;
  /** A value stored as text: it matches a ticked value by its text alone. */
  text?: boolean;
}
export interface FilterVisibilityConstraint {
  column: number; values?: readonly string[]; condition?: ConditionFilter;
  color?: { kind: "cell" | "font"; values: readonly string[]; worksheet?: true };
}
export interface FilterVisibilityGroup { first: number; last: number; constraints: FilterVisibilityConstraint[] }
export interface PreparedFilterVisibilityGroup extends FilterVisibilityGroup {
  /** Compact columns only: unrelated sheet width never allocates blank slots. */
  columns: Array<{ column: number; cells: FilterVisibilityCell[] }>;
}
export type PreparedFilterVisibility = PreparedFilterVisibilityGroup[];

/** Ordinary worksheet/Table criteria. Specialized date/icon criteria keep the
 * existing conservative host path until their typed authority is migrated. */
export function filterVisibilityGroups(source: FilterVisibilitySource): FilterVisibilityGroup[] | null {
  const tables = normalizeDataTables(source.dataTables, source);
  if (tables.some(table => Object.values(table.columnFilters ?? {}).some(filter => filter.dateGroups || filter.icon || filter.relativeDate))) return null;
  const range = source.autoFilterRef ? parseA1Range(source.autoFilterRef.replace(/\$/g, "")) : null;
  const owns = (col: number) => Number.isInteger(col) && col >= 0 && col < source.columnCount && (!range || (col >= range.c1 && col <= range.c2));
  const constraints: FilterVisibilityConstraint[] = [];
  const conditions = new Map<number, ConditionFilter>();
  for (const [key, raw] of Object.entries(source.conditionFilters ?? {})) {
    const column = Number(key), condition = normalizeConditionFilter(raw);
    if (owns(column) && condition) { conditions.set(column, condition); constraints.push({ column, condition }); }
  }
  for (const [key, values] of Object.entries(source.filters ?? {})) {
    const column = Number(key);
    // A ticked value matches by its filter key (trailing spaces off, case
    
    // as the list offers it and as the synchronous filter matches it.
    if (owns(column) && !conditions.has(column) && Array.isArray(values)) constraints.push({ column, values: values.map(String) });
  }
  for (const [key, values] of Object.entries(source.colorFilters ?? {})) {
    const column = Number(key);
    if (owns(column) && Array.isArray(values)) constraints.push({ column, color: { kind: "cell", values: values.map(value => normalizeFillColor(String(value))), worksheet: true } });
  }
  const groups: FilterVisibilityGroup[] = [];
  if (constraints.length) groups.push({ first: range ? range.r1 + 1 : 1, last: Math.min(source.rowCount - 1, range?.r2 ?? source.rowCount - 1), constraints });
  for (const table of tables) {
    if (table.nativeResizeBlocker || !table.headerRow || !table.columns || !table.columnFilters) continue;
    const byId = new Map(table.columns.map((column, offset) => [String(column.id), table.range.c1 + offset]));
    const constraints: FilterVisibilityConstraint[] = [];
    for (const [id, filter] of Object.entries(table.columnFilters)) {
      const column = byId.get(id); if (column === undefined) continue;
      if (filter.values) constraints.push({ column, values: filter.values });
      else if (filter.condition) constraints.push({ column, condition: filter.condition });
      else if (filter.color) constraints.push({ column, color: { kind: filter.color.kind, values: [filter.color.value] } });
    }
    if (constraints.length) groups.push({ first: table.range.r1 + 1, last: table.range.r2, constraints });
  }
  return groups;
}

export function* prepareFilterVisibilitySteps(groups: readonly FilterVisibilityGroup[], read: (row: number, column: number) => FilterVisibilityCell): Generator<void, PreparedFilterVisibility> {
  const result: PreparedFilterVisibility = []; let work = 0;
  for (const group of groups) {
    const columns: PreparedFilterVisibilityGroup["columns"] = [];
    for (const column of new Set(group.constraints.map(constraint => constraint.column))) {
      const cells: FilterVisibilityCell[] = [];
      for (let row = group.first; row <= group.last; row++) {
        cells.push(read(row, column));
        if (++work % 128 === 0) yield;
      }
      columns.push({ column, cells });
    }
    result.push({ ...group, columns });
  }
  return result;
}

export function* filterVisibilitySteps(groups: PreparedFilterVisibility): Generator<void, number[]> {
  const hidden = new Set<number>(); let work = 0;
  for (const group of groups) {
    const columns = new Map(group.columns.map(column => [column.column, column.cells]));
    const predicates = group.constraints.map(constraint => {
      const cells = columns.get(constraint.column)!;
      const allowed = constraint.values ? filterValueKeys(constraint.values) : undefined;
      // A file's filter values are the texts its author's locale showed
      // ("4,00" in a Hungarian file): a number shown here in another locale
      // ("4.00") is still the value picked (X-146, S27/Hungarian autofilter).
      const allowedNumbers = constraint.values ? new Set(constraint.values.flatMap(localeNumberReadings)) : undefined;
      const colors = constraint.color ? new Set(constraint.color.values) : undefined;
      const predicate = constraint.condition ? buildRowPredicate(constraint.condition, cells.map(cell => cell.display), cells.map(cell => cell.numeric)) : undefined;
      return (offset: number) => {
        const cell = cells[offset];
        // The locale reading is a NUMBER's: a text "5.10" is not the "5.1"
        // ticked, as Excel's value filter compares the text it shows
        
        if (allowed && !allowed.has(filterValueKey(cell.display)) && !(!cell.text && allowedNumbers?.size && localeNumberReadings(cell.display).some(value => allowedNumbers.has(value)))) return false;
        if (predicate && !predicate(cell.display, cell.numeric)) return false;
        if (constraint.color && colors) {
          const raw = constraint.color.kind === "cell" ? cell.fill : cell.font;
          const color = constraint.color.worksheet ? normalizeFillColor(raw) : normalizeDataTableFilterColor({kind: constraint.color.kind, value: raw})?.value ?? null;
          if (color === null || !colors.has(color)) return false;
        }
        return true;
      };
    });
    for (let row = group.first; row <= group.last; row++) {
      if (!predicates.every(predicate => predicate(row - group.first))) hidden.add(row);
      if (++work % 512 === 0) yield;
    }
  }
  return [...hidden];
}

/** The numbers a shown number text reads as in either decimal convention
 * ("4,00" and "4.00" are 4; "1,234" is 1234 or 1.234); none for other text. */
export function localeNumberReadings(text: string): number[] {
  const compact = String(text).replace(/[\s\u00a0\u202f']/g, "");
  if (!/^[+-]?[\d.,]*\d[\d.,]*$/.test(compact)) return [];
  const readings = new Set<number>();
  for (const [decimal, group] of [[".", ","], [",", "."]] as const) {
    const parts = compact.split(decimal);
    if (parts.length > 2) continue;
    const whole = parts[0].split(group).join("");
    if (!/^[+-]?\d*$/.test(whole) || (parts[1] !== undefined && !/^\d*$/.test(parts[1]))) continue;
    const value = Number(`${whole}${parts[1] !== undefined ? `.${parts[1]}` : ""}`);
    if (Number.isFinite(value)) readings.add(value);
  }
  return [...readings];
}

export function exhaustFilterVisibility<T>(steps: Generator<void, T>): T {
  for (;;) { const next = steps.next(); if (next.done) return next.value; }
}
