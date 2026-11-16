/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Tables: Format as Table / Insert > Table's OK and the Table Design tab
 * (Table Style, Banded Rows / Columns, First / Last Column, Filter Button,
 * Total Row, Table Name, Convert to Range) as pure sheet and workbook
 * transactions, plus the table helpers they share with the rest of the
 * window. Moved out of SpreadsheetWindow so the ribbon and the spreadsheet
 * editor commands run the same code; behaviour is unchanged. The window keeps
 * its dialog authority, notices, toasts, confirmation and selection moves.
 */

import { spreadsheetDefaultNames } from "@mrnote/core/web/lib/spreadsheet/default-names.ts";
import {
  hasCellMetadata,
  updateCellStyle,
  updateCellValue,
  writeCellValueInto,
  type CellButtonsSheet,
  type CellMap,
  type SpreadsheetCell,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import type { ConditionalFormatRule } from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import {
  buildDataTableFromRange,
  defaultTableColumnName,
  isAuthorableTableTotalFormula,
  isValidDataTableName,
  MAX_DATA_TABLES,
  normalizeDataTables,
  normalizeTableColumnNames,
  renameDataTableColumnReferences,
  renameDataTableReferences,
  tableCellRole,
  tableColumnNumberFormat,
  tableNativeFilterCustody,
  tableNativeStructureBlocked,
  tableNativeStyle,
  tableStyleCellPaint,
  tableStylePreset,
  tableTotalCellOwnsValue,
  tableTotalCellValue,
  tableWithStyleChoice,
  type DataTable,
  type TableStyleChoice,
  type TableTotalKind,
} from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import type { ValidationRule } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import { spreadsheetCellIsFormula } from "@mrnote/core/web/lib/spreadsheet/display-cache.ts";
import {
  convertRemovedTableReferencesToA1,
  shiftFormula,
  type TableDefinition,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { normalizeMerges, rangeHasMerge, removeMergesInRange, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { cellsAfterUnmerge } from "./merge-buttons";
import { normalizeNamedRanges, type NamedRange } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { normalizeManualHidden, serializeManualHidden, unhideLines } from "@mrnote/core/web/lib/spreadsheet/manual-hidden.ts";
import type { SpreadsheetNativeTableStyle } from "@mrnote/core/web/lib/spreadsheet/native-table-styles.ts";
import { parseA1Range } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";
import { normalizeSlicers } from "./slicer";
import { insertLine, type StructuralSheet } from "@mrnote/core/web/lib/spreadsheet/structural.ts";
import { extendRangeRefsForTableGrowth } from "./table-autocorrect";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** The parts of a worksheet the table buttons read and write. */
export interface TableButtonSheet extends CellButtonsSheet {
  id: string;
  name: string;
  dataTables?: DataTable[];
  namedRanges?: NamedRange[];
  validations?: ValidationRule[];
  conditionalFormats?: ConditionalFormatRule[];
  merges?: MergeSpan[];
  slicers?: unknown;
  splitRow?: number;
  splitCol?: number;
  filters?: Record<string, string[]>;
  colorFilters?: Record<string, string[]>;
  conditionFilters?: Record<string, unknown>;
}

/** A sheet's protection question: is a cell of `range` locked? */
export type LockedCellTest<S> = (sheet: S, range: Rect) => boolean;

/** The cells a table occupies, its Totals Row included. */
export function tableOccupiedRange(table: DataTable): Rect {
  return { ...table.range, r2: table.range.r2 + (table.totalsRow ? 1 : 0) };
}

function sheetTables(sheet: TableButtonSheet): DataTable[] {
  return normalizeDataTables(sheet.dataTables, {
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  });
}

export function coordinateMapTouchesColumnRange(
  raw: unknown,
  c1: number,
  c2: number,
  columnCount: number,
): boolean {
  if (raw === undefined) return false;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return true;
  for (const key of Object.keys(raw)) {
    if (!/^(0|[1-9]\d*)$/.test(key)) return true;
    const column = Number(key);
    if (!Number.isSafeInteger(column) || column < 0 || column >= columnCount) {
      return true;
    }
    if (column >= c1 && column <= c2) return true;
  }
  return false;
}

export function tableRangeResizeBlocked(
  sheet: Pick<TableButtonSheet, "columnCount" | "filters" | "colorFilters" | "conditionFilters">,
  table: DataTable,
  // The caller moves the physical Totals Row itself (Remove Duplicates
  // rewrites it below the rows it keeps), so a Totals Row is no blocker.
  options?: { totalsRowMovable?: boolean },
): boolean {
  return (
    (table.totalsRow === true && !options?.totalsRowMovable) ||
    table.nativeResizeBlocker !== undefined ||
    coordinateMapTouchesColumnRange(
      sheet.filters,
      table.range.c1,
      table.range.c2,
      sheet.columnCount,
    ) ||
    coordinateMapTouchesColumnRange(
      sheet.colorFilters,
      table.range.c1,
      table.range.c2,
      sheet.columnCount,
    ) ||
    coordinateMapTouchesColumnRange(
      sheet.conditionFilters,
      table.range.c1,
      table.range.c2,
      sheet.columnCount,
    )
  );
}

/** Return the exact valid column identities stored in a table's header cells. */
export function tableHeadersForCells(
  table: DataTable,
  cells: CellMap,
): string[] | null {
  const width = table.range.c2 - table.range.c1 + 1;
  if (table.columns?.length === width) {
    return table.columns.map((column) => column.name);
  }
  if (!table.headerRow) {
    return Array.from({ length: width }, (_, index) => `Column${index + 1}`);
  }
  const raw: string[] = [];
  for (let col = table.range.c1; col <= table.range.c2; col += 1) {
    raw.push(String(cells[`${table.range.r1}:${col}`]?.value ?? "").trim());
  }
  const normalized = normalizeTableColumnNames(raw);
  return normalized.every((name, index) => name === raw[index])
    ? normalized
    : null;
}

/** A table column a header edit renamed. */
export interface TableColumnRename {
  readonly tableName: string;
  readonly previousName: string;
  readonly nextName: string;
}

/**
 * The table columns a Replace touched through their header cells: the tables
 * get their column identities from the replaced headers (normalised the way a
 * header edit is) and every renamed column is reported so structured
 * references follow, as they do when the header is edited by hand (live
 * audit B165). The header cells are rewritten to the normalised names.
 */
export function tableHeaderRenamesAfterReplace<S extends TableButtonSheet>(
  sheet: S,
  nextCells: CellMap,
  replacedKeys: readonly string[],
): {
  dataTables: DataTable[] | undefined;
  renames: { tableName: string; previousName: string; nextName: string }[];
} {
  const tables = sheet.dataTables;
  if (!tables || tables.length === 0 || replacedKeys.length === 0) {
    return { dataTables: tables, renames: [] };
  }
  const replaced = new Set(replacedKeys);
  const renames: { tableName: string; previousName: string; nextName: string }[] = [];
  let changed = false;
  const nextTables = tables.map((table) => {
    if (!table.headerRow) return table;
    const { r1, c1, c2 } = table.range;
    let touched = false;
    for (let col = c1; col <= c2 && !touched; col += 1) {
      if (replaced.has(`${r1}:${col}`)) touched = true;
    }
    if (!touched) return table;
    const width = c2 - c1 + 1;
    const previous =
      tableHeadersForCells(table, sheet.cells) ??
      Array.from({ length: width }, (_, index) => `Column${index + 1}`);
    const raw: string[] = [];
    for (let col = c1; col <= c2; col += 1) {
      const key = `${r1}:${col}`;
      raw.push(
        replaced.has(key)
          ? String(nextCells[key]?.value ?? "").trim()
          : previous[col - c1],
      );
    }
    // An emptied header takes the first free "Column1" ("列1" in Chinese),
    // left to right, not its position: Clear Contents on B1 "H2" of T A1:C4
    // reads 列1 in Excel (Mac Excel 16, excel-formats/xprobe-rounds1113 A;
    // xlsx-rounds R379-X23, R380-X18 -- Text to Columns writing an empty
    // piece into a header).
    raw.forEach((name, index) => {
      if (name !== "" || !replaced.has(`${r1}:${c1 + index}`)) return;
      raw[index] = defaultTableColumnName(
        raw.filter((other, at) => at !== index && other !== ""),
        (number) => `${spreadsheetDefaultNames().column}${number}`,
      );
    });
    // Headers not replaced keep their names (S-124's rule, as when typed).
    const normalized = normalizeTableColumnNames(raw, {
      keep: raw.map((_name, index) => !replaced.has(`${r1}:${c1 + index}`)),
    });
    for (let col = c1; col <= c2; col += 1) {
      const key = `${r1}:${col}`;
      if (!replaced.has(key)) continue;
      const cell = nextCells[key];
      // A header cleared (Clear Contents, Delete) is no cell any more: it
      // takes the name its column now has ("Column1"), as Excel writes one
      // into an emptied header. Left blank, the table kept a name no cell
      // showed (sheet model fuzz).
      if (!cell) nextCells[key] = { value: normalized[col - c1], kind: "text" };
      else if (cell.value !== normalized[col - c1]) {
        nextCells[key] = { ...cell, value: normalized[col - c1] };
      }
    }
    normalized.forEach((name, index) => {
      if (name !== previous[index]) {
        renames.push({ tableName: table.name, previousName: previous[index], nextName: name });
      }
    });
    changed = true;
    return {
      ...table,
      columns: normalized.map((name, index) => ({
        id: table.columns?.[index]?.id ?? index + 1,
        name,
      })),
    };
  });
  return { dataTables: changed ? nextTables : tables, renames };
}

/**
 * A table's header cells written by any command that does not name tables
 * itself (Paste, Fill, Clear, a script): the header row IS the columns'
 * names in Excel, so a changed header renames its column and its structured
 * references follow, exactly as typing it does. Keyboard Paste over two
 * headers left the columns and their formulas on the old names and the new
 * ones read #REF! (live audit S-123). `previous` is the sheet before the
 * write; returns the sheet with its tables and references brought along and
 * the renames for the other sheets, or null when no header changed.
 */
export function reconcileTableHeaderWrites<S extends TableButtonSheet>(
  previous: S,
  next: S,
): { sheet: S; renames: TableColumnRename[] } | null {
  const tables = previous.dataTables;
  if (!tables || tables.length === 0 || next.cells === previous.cells || next.dataTables !== tables) return null;
  const changed: string[] = [];
  for (const table of tables) {
    if (!table.headerRow) continue;
    const { r1, c1, c2 } = table.range;
    for (let col = c1; col <= c2; col += 1) {
      const key = `${r1}:${col}`;
      if (String(next.cells[key]?.value ?? "") !== String(previous.cells[key]?.value ?? "")) changed.push(key);
    }
  }
  if (changed.length === 0) return null;
  const cells: CellMap = { ...next.cells };
  const { dataTables, renames } = tableHeaderRenamesAfterReplace(previous, cells, changed);
  if (dataTables === tables && renames.length === 0) return null;
  return {
    sheet: renameTableColumnsInSheet({ ...next, cells, dataTables }, renames),
    renames,
  };
}

/**
 * Several renamed table columns at once. They take effect together: two
 * headers that trade names (A to B, B to A) swap their references instead of
 * the second rename sweeping up the first one's.
 */
export function renameTableColumnsInSheet<S extends TableButtonSheet>(
  candidate: S,
  renames: readonly TableColumnRename[],
): S {
  if (renames.length <= 1) {
    return renames.reduce(
      (current, rename) =>
        renameTableColumnReferencesInSheet(current, rename.tableName, rename.previousName, rename.nextName),
      candidate,
    );
  }
  const stand = (index: number) => `MRAIRENAME${index}X${renames.length}`;
  const parked = renames.reduce(
    (current, rename, index) =>
      renameTableColumnReferencesInSheet(current, rename.tableName, rename.previousName, stand(index)),
    candidate,
  );
  return renames.reduce(
    (current, rename, index) =>
      renameTableColumnReferencesInSheet(current, rename.tableName, stand(index), rename.nextName),
    parked,
  );
}

/**
 * Follow one renamed table column through a sheet (Excel: renaming a column
 * updates every structured reference to it in the workbook). Explicit
 * `Table[Col]` references are renamed wherever they sit; the in-table
 * `[@Col]` / `[Col]` forms only where the renamed table owns the formula, so
 * the calculated-column template and the physical cells it filled stay equal.
 */
export function renameTableColumnReferencesInSheet<S extends TableButtonSheet>(
  candidate: S,
  tableName: string,
  oldColumnName: string,
  newColumnName: string,
): S {
  const tableLower = tableName.toLowerCase();
  return rewriteTableReferencesInSheet(candidate, (formula, owner) =>
    renameDataTableColumnReferences(
      formula,
      tableName,
      oldColumnName,
      newColumnName,
      { shortFormOwnedByTable: owner?.name.toLowerCase() === tableLower },
    ),
  );
}

/**
 * Apply one structured-reference rewrite to every formula a sheet stores:
 * cell formulas, every table's calculated-column and totals-row recipes,
 * defined names and validation sources. `rewrite` receives the table that OWNS
 * the formula — the table whose range (totals row included) contains the cell,
 * or the table a recipe belongs to — because Excel's unqualified `[@Col]` /
 * `[Col]` forms address that table; names and validations own none. A sheet
 * that nothing references comes back unchanged (same object).
 */
export function rewriteTableReferencesInSheet<S extends TableButtonSheet>(
  candidate: S,
  rewrite: (formula: string, owner: DataTable | null) => string,
): S {
  const tables = candidate.dataTables ?? [];
  const ownerAt = (row: number, col: number): DataTable | null =>
    tables.find(
      (table) =>
        row >= table.range.r1 &&
        row <= table.range.r2 + (table.totalsRow ? 1 : 0) &&
        col >= table.range.c1 &&
        col <= table.range.c2,
    ) ?? null;
  // Names and rule sources may be stored without the leading "=".
  const rewriteBare = (source: string, owner: DataTable | null) =>
    source.startsWith("=")
      ? rewrite(source, owner)
      : rewrite(`=${source}`, owner).slice(1);
  let next = candidate;

  let nextCells: CellMap | null = null;
  for (const [key, cell] of Object.entries(candidate.cells)) {
    const hasFormulaValue =
      typeof cell.value === "string" && cell.value.startsWith("=");
    const hasFormula =
      typeof cell.formula === "string" && cell.formula.startsWith("=");
    if (!hasFormulaValue && !hasFormula) continue;
    const separator = key.indexOf(":");
    const owner = ownerAt(
      Number(key.slice(0, separator)),
      Number(key.slice(separator + 1)),
    );
    let updated = cell;
    if (hasFormulaValue) {
      const value = rewrite(cell.value as string, owner);
      if (value !== cell.value) updated = { ...updated, value };
    }
    if (hasFormula) {
      const formula = rewrite(cell.formula as string, owner);
      if (formula !== cell.formula) updated = { ...updated, formula };
    }
    if (updated === cell) continue;
    if (!nextCells) nextCells = { ...candidate.cells };
    nextCells[key] = updated;
  }
  if (nextCells) next = { ...next, cells: nextCells };

  let tablesChanged = false;
  const nextTables = tables.map((table) => {
    if (!table.totalFormulas && !table.calculatedFormulas) return table;
    let formulasChanged = false;
    const renameRecipe = (recipe: Record<number, string> | undefined) =>
      recipe
        ? Object.fromEntries(
            Object.entries(recipe).map(([offset, formula]) => {
              const renamed = rewrite(formula, table);
              formulasChanged ||= renamed !== formula;
              return [offset, renamed];
            }),
          )
        : undefined;
    const totalFormulas = renameRecipe(table.totalFormulas);
    const calculatedFormulas = renameRecipe(table.calculatedFormulas);
    if (!formulasChanged) return table;
    tablesChanged = true;
    return {
      ...table,
      ...(totalFormulas ? { totalFormulas } : {}),
      ...(calculatedFormulas ? { calculatedFormulas } : {}),
    };
  });
  if (tablesChanged) next = { ...next, dataTables: nextTables };

  if (candidate.namedRanges && candidate.namedRanges.length > 0) {
    let namesChanged = false;
    const nextNamed = candidate.namedRanges.map((entry) => {
      const ref = rewriteBare(entry.ref, null);
      if (ref === entry.ref) return entry;
      namesChanged = true;
      return { ...entry, ref };
    });
    if (namesChanged) next = { ...next, namedRanges: nextNamed };
  }

  if (candidate.validations && candidate.validations.length > 0) {
    let rulesChanged = false;
    const nextValidations = candidate.validations.map((rule) => {
      if (rule.kind === "list" && typeof rule.sourceRange === "string") {
        const sourceRange = rewriteBare(rule.sourceRange, null);
        if (sourceRange === rule.sourceRange) return rule;
        rulesChanged = true;
        return { ...rule, sourceRange };
      }
      if (rule.kind === "formula" && typeof rule.formula === "string") {
        const formula = rewriteBare(rule.formula, null);
        if (formula === rule.formula) return rule;
        rulesChanged = true;
        return { ...rule, formula };
      }
      return rule;
    });
    if (rulesChanged) next = { ...next, validations: nextValidations };
  }
  return next;
}

export function convertRemovedTableReferencesInSheet<S extends TableButtonSheet>(
  candidate: S,
  table: TableDefinition,
  tableSheetName: string,
): S | null {
  let nextCells: CellMap | null = null;
  for (const [key, cell] of Object.entries(candidate.cells)) {
    const source =
      cell.formula ??
      (String(cell.value ?? "").trim().startsWith("=")
        ? String(cell.value)
        : null);
    if (!source) continue;
    const separator = key.indexOf(":");
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    const converted = convertRemovedTableReferencesToA1(
      source,
      table,
      { row, col },
      tableSheetName,
      candidate.name,
    );
    if (converted === null) return null;
    if (converted === source) continue;
    if (!nextCells) nextCells = { ...candidate.cells };
    writeCellValueInto(nextCells, row, col, converted);
  }

  let namedRangesChanged = false;
  const namedRanges = candidate.namedRanges?.map((range) => {
    const converted = convertRemovedTableReferencesToA1(
      range.ref,
      table,
      undefined,
      tableSheetName,
    );
    if (converted === null) return null;
    if (converted === range.ref) return range;
    if (range.nativeBlocker) return null;
    namedRangesChanged = true;
    return { ...range, ref: converted };
  });
  if (namedRanges?.some((range) => range === null)) return null;

  let validationsChanged = false;
  const validations = candidate.validations?.map((rule) => {
    const source =
      rule.kind === "formula"
        ? rule.formula
        : rule.kind === "list" && rule.sourceRange
          ? rule.sourceRange
          : null;
    if (!source) return rule;
    const converted = convertRemovedTableReferencesToA1(
      source,
      table,
      { row: rule.range.r1, col: rule.range.c1 },
      tableSheetName,
      candidate.name,
      { thisRowRelative: true },
    );
    if (converted === null) return null;
    if (converted === source) return rule;
    validationsChanged = true;
    return rule.kind === "formula"
      ? { ...rule, formula: converted }
      : { ...rule, sourceRange: converted };
  });
  if (validations?.some((rule) => rule === null)) return null;

  let conditionalFormatsChanged = false;
  const conditionalFormats = candidate.conditionalFormats?.map((rule) => {
    if (rule.kind !== "formula") return rule;
    const converted = convertRemovedTableReferencesToA1(
      rule.formula,
      table,
      { row: rule.range.r1, col: rule.range.c1 },
      tableSheetName,
      candidate.name,
      { thisRowRelative: true },
    );
    if (converted === null) return null;
    if (converted === rule.formula) return rule;
    conditionalFormatsChanged = true;
    return { ...rule, formula: converted };
  });
  if (conditionalFormats?.some((rule) => rule === null)) return null;

  let dataTablesChanged = false;
  const dataTables = candidate.dataTables?.map((candidateTable) => {
    if (
      !candidateTable.totalFormulas &&
      !candidateTable.calculatedFormulas
    ) {
      return candidateTable;
    }
    let formulasChanged = false;
    const formulas: Record<number, string> = {};
    for (const [rawOffset, formula] of Object.entries(
      candidateTable.totalFormulas ?? {},
    )) {
      const offset = Number(rawOffset);
      const converted = convertRemovedTableReferencesToA1(
        formula,
        table,
        {
          row: candidateTable.range.r2 + 1,
          col: candidateTable.range.c1 + offset,
        },
        tableSheetName,
        candidate.name,
      );
      if (converted === null) return null;
      formulas[offset] = converted;
      formulasChanged ||= converted !== formula;
    }
    const calculatedFormulas: Record<number, string> = {};
    for (const [rawOffset, formula] of Object.entries(
      candidateTable.calculatedFormulas ?? {},
    )) {
      const offset = Number(rawOffset);
      const firstDataRow =
        candidateTable.range.r1 + (candidateTable.headerRow ? 1 : 0);
      const converted = convertRemovedTableReferencesToA1(
        formula,
        table,
        { row: firstDataRow, col: candidateTable.range.c1 + offset },
        tableSheetName,
        candidate.name,
      );
      if (converted === null) return null;
      calculatedFormulas[offset] = converted;
      formulasChanged ||= converted !== formula;
    }
    if (!formulasChanged) return candidateTable;
    dataTablesChanged = true;
    return {
      ...candidateTable,
      ...(candidateTable.totalFormulas ? { totalFormulas: formulas } : {}),
      ...(candidateTable.calculatedFormulas
        ? { calculatedFormulas }
        : {}),
    };
  });
  if (dataTables?.some((candidateTable) => candidateTable === null)) {
    return null;
  }

  return nextCells ||
    namedRangesChanged ||
    validationsChanged ||
    conditionalFormatsChanged ||
    dataTablesChanged
    ? {
        ...candidate,
        ...(nextCells ? { cells: nextCells } : {}),
        ...(namedRangesChanged
          ? { namedRanges: namedRanges as NonNullable<S["namedRanges"]> }
          : {}),
        ...(validationsChanged
          ? { validations: validations as NonNullable<S["validations"]> }
          : {}),
        ...(conditionalFormatsChanged
          ? {
              conditionalFormats:
                conditionalFormats as NonNullable<
                  S["conditionalFormats"]
                >,
            }
          : {}),
        ...(dataTablesChanged
          ? { dataTables: dataTables as NonNullable<S["dataTables"]> }
          : {}),
      }
    : candidate;
}

/**
 * The format a table cell shows: the table style underneath, the cell's own
 * (direct) formats on top (live audit B142). Convert to Range writes it into
 * the cell, and a copy of the cell pasted outside any table carries it as its
 * source formatting (live audit B157), as Excel does. Null when the cell is
 * not part of the table or the table style paints nothing there.
 */
export function tableCellEffectiveStyle(
  table: DataTable,
  row: number,
  col: number,
  own: SpreadsheetCell["style"],
  themeAccents: readonly string[],
  nativeStyles?: readonly SpreadsheetNativeTableStyle[],
): NonNullable<SpreadsheetCell["style"]> | null {
  if (!tableCellRole(table, row, col)) return null;
  const { borderEdges: tableBorderEdges, ...patch } = tableStyleCellPaint(
    table.styleId, themeAccents, table, row, col, nativeStyles,
  );
  if (Object.keys(patch).length === 0 && !tableBorderEdges) return null;
  const direct = own ?? {};
  const style: NonNullable<SpreadsheetCell["style"]> = { ...patch, ...direct };
  if (patch.backgroundColor && direct.backgroundColor === undefined) {
    delete style.fillPattern;
    delete style.gradientFill;
  }
  if (tableBorderEdges) {
    style.borderEdges = {
      ...tableBorderEdges,
      ...(direct.borderEdges ?? {}),
    };
  }
  return style;
}

export function materializeDataTableStyle<S extends TableButtonSheet>(
  candidate: S,
  table: DataTable,
  themeAccents: readonly string[],
  nativeStyles?: readonly SpreadsheetNativeTableStyle[],
): S {
  let nextCells: CellMap | null = null;
  const lastRow = table.range.r2 + (table.totalsRow ? 1 : 0);
  for (let row = table.range.r1; row <= lastRow; row += 1) {
    for (let col = table.range.c1; col <= table.range.c2; col += 1) {
      const key = `${row}:${col}`;
      const previous = candidate.cells[key] ?? {
        value: "",
        kind: "text" as const,
      };
      const style = tableCellEffectiveStyle(
        table, row, col, previous.style, themeAccents, nativeStyles,
      );
      if (!style) continue;
      if (!nextCells) nextCells = { ...candidate.cells };
      nextCells[key] = { ...previous, style };
    }
  }
  return nextCells ? { ...candidate, cells: nextCells } : candidate;
}

/**
 * Give one totals-row cell the number format of the column it totals, when a
 * total is set there and the cell carries no number format of its own: Excel
 * formats the total like its column, so "#,##0.00" data sums to "81,427.32"
 * (Microsoft Support, "Total the data in an Excel table"). A format the user
 * set on the totals cell itself is never replaced; a label or blank column is
 * left alone. The format lives on the physical cell, so export writes it.
 */
export function applyTableTotalNumberFormat<S extends TableButtonSheet>(
  sheet: S,
  table: DataTable,
  offset: number,
  totalsRow: number,
): S {
  const kind = table.totals?.[offset] ?? "none";
  if (kind === "none") return sheet;
  const col = table.range.c1 + offset;
  const existing = sheet.cells[`${totalsRow}:${col}`]?.style?.numberFormat;
  if (typeof existing === "string" && existing.trim() !== "") return sheet;
  const format = tableColumnNumberFormat(sheet.cells, table, offset);
  if (!format) return sheet;
  return updateCellStyle(sheet, totalsRow, col, { numberFormat: format });
}

/**
 * Give back one totals-row cell the Totals Row stops occupying (Total Row off,
 * or a resize that moves the row away). Its owned value goes, and so does the
 * number format the Totals Row itself gave the cell — the one
 * applyTableTotalNumberFormat inherits from the column it totals — so the row
 * is left as the Totals Row found it and turning it on again is not refused by
 * the Totals Row's own leftover. Every other format, comment, link or unlock
 * on the cell was put there by the user and stays (and still blocks a later
 * Totals Row from writing over it).
 */
export function releaseTableTotalCell<S extends TableButtonSheet>(
  sheet: S,
  table: DataTable,
  offset: number,
  totalsRow: number,
): S {
  const col = table.range.c1 + offset;
  const key = `${totalsRow}:${col}`;
  const next = updateCellValue(sheet, totalsRow, col, "");
  const cell = next.cells[key];
  const format = cell?.style?.numberFormat;
  if (
    !cell ||
    (table.totals?.[offset] ?? "none") === "none" ||
    typeof format !== "string" ||
    format === "" ||
    format !== tableColumnNumberFormat(sheet.cells, table, offset)
  ) {
    return next;
  }
  const style = { ...cell.style };
  delete style.numberFormat;
  const released: SpreadsheetCell = { ...cell };
  if (Object.keys(style).length > 0) released.style = style;
  else delete released.style;
  const cells = { ...next.cells };
  if (hasCellMetadata(released)) cells[key] = released;
  else delete cells[key];
  return { ...next, cells };
}

/** Why a Total Row function was not set (the hook says why; a command reports not applicable). */
export type TableColumnTotalRefusal = "invalid" | "locked" | "custom_content";

/**
 * Table Design > Total Row, one column's function (None, Average, Count,
 * Count Numbers, Max, Min, Sum, StdDev, Var, or a custom formula): the table's
 * totals and the totals cell rewritten, in the column's number format. The
 * worksheet's total-row dropdown and the AI's spreadsheet.data.table_total both
 * call this. `isLocked` answers whether the totals cell is protected.
 */
export function sheetWithTableColumnTotal<S extends TableButtonSheet>(
  sheet: S,
  tableId: string,
  offset: number,
  kind: TableTotalKind,
  customFormula: string | undefined,
  isLocked: (cell: { row: number; col: number }) => boolean,
): { sheet: S } | { refusal: TableColumnTotalRefusal } {
  const dataTables = sheetTables(sheet);
  const current = dataTables.find((table) => table.id === tableId);
  if (
    !current?.totalsRow ||
    tableNativeStructureBlocked(current) ||
    !Number.isInteger(offset) ||
    offset < 0 ||
    offset > current.range.c2 - current.range.c1
  ) {
    return { refusal: "invalid" };
  }
  const headers = tableHeadersForCells(current, sheet.cells);
  if (!headers) return { refusal: "invalid" };
  const normalizedCustomFormula = kind === "custom" ? customFormula?.trim() ?? "" : "";
  if (kind === "custom" && !isAuthorableTableTotalFormula(normalizedCustomFormula)) {
    return { refusal: "invalid" };
  }
  const totalsRow = current.range.r2 + 1;
  const col = current.range.c1 + offset;
  if (isLocked({ row: totalsRow, col })) return { refusal: "locked" };
  const currentValue = String(sheet.cells[`${totalsRow}:${col}`]?.value ?? "");
  if (!tableTotalCellOwnsValue(current, offset, headers[offset], currentValue)) {
    return { refusal: "custom_content" };
  }
  const totals = { ...(current.totals ?? {}) };
  if (kind === "none") delete totals[offset];
  else totals[offset] = kind;
  const totalFormulas = { ...(current.totalFormulas ?? {}) };
  if (kind === "custom") totalFormulas[offset] = normalizedCustomFormula;
  else delete totalFormulas[offset];
  const totalLabels = { ...(current.totalLabels ?? {}) };
  if (kind !== "none") delete totalLabels[offset];
  const nextTable: DataTable = { ...current };
  if (Object.keys(totals).length > 0) nextTable.totals = totals;
  else delete nextTable.totals;
  if (Object.keys(totalFormulas).length > 0) nextTable.totalFormulas = totalFormulas;
  else delete nextTable.totalFormulas;
  if (Object.keys(totalLabels).length > 0) nextTable.totalLabels = totalLabels;
  else delete nextTable.totalLabels;
  // The chosen total reads in its column's number format; a format the user
  // set on the totals cell itself stays (applyTableTotalNumberFormat).
  // Choosing None gives back the format the previous total inherited.
  const nextSheet = applyTableTotalNumberFormat(
    updateCellValue(
      kind === "none" ? releaseTableTotalCell(sheet, current, offset, totalsRow) : sheet,
      totalsRow,
      col,
      tableTotalCellValue(nextTable, offset, headers[offset]),
    ),
    nextTable,
    offset,
    totalsRow,
  );
  return {
    sheet: {
      ...nextSheet,
      dataTables: dataTables.map((table) => (table.id === tableId ? nextTable : table)),
    },
  };
}

/** Why Create Table refused, as the dialog reports it. */
export type CreateTableRefusal =
  | "limit"
  | "invalidRange"
  | "protected"
  | "rowLimit"
  | "overlap";

export interface CreatedTable<S> {
  sheet: S;
  /** The table's range (one row taller when a header row was generated). */
  range: Rect;
  /** "My table has headers" was off: a header row was inserted above. */
  insertedHeaderRow: boolean;
}

/**
 * Create Table's OK on `range` (Format as Table / Insert > Table). With
 * `hasHeaders` off a real header row "Column1", "Column2", ... is inserted
 * above the data, as Excel does; column names are made non-blank and unique
 * (a formula header gives its displayed text and is frozen to it). The table
 * takes the next free "TableN" name in the workbook and `styleId`.
 */
export function tableCreatedSheet<S extends TableButtonSheet & StructuralSheet>(
  workbook: readonly TableButtonSheet[],
  sheet: S,
  request: { range: Rect; hasHeaders: boolean; styleId: TableStyleChoice },
  context: {
    rangeHasLockedCell: LockedCellTest<S>;
    /** The text a header cell shows (a formula's result). */
    headerDisplay: (cell: SpreadsheetCell | undefined, row: number, col: number) => string;
  },
): CreatedTable<S> | CreateTableRefusal {
  if (sheetTables(sheet).length >= MAX_DATA_TABLES) return "limit";
  const parsed = request.range;
  if (
    parsed.r1 < 0 ||
    parsed.c1 < 0 ||
    parsed.r2 >= MAX_GRID_ROWS ||
    parsed.c2 >= MAX_GRID_COLUMNS
  ) {
    return "invalidRange";
  }
  const expandedSheet: S = {
    ...sheet,
    rowCount: Math.max(sheet.rowCount, parsed.r2 + 1),
    columnCount: Math.max(sheet.columnCount, parsed.c2 + 1),
  };
  let nextSheet = expandedSheet;
  let tableRange = parsed;
  if (request.hasHeaders) {
    if (context.rangeHasLockedCell(sheet, parsed)) return "protected";
  } else {
    if (parsed.r2 + 1 >= MAX_GRID_ROWS) return "rowLimit";
    // Excel generates a real header row when "My table has headers" is off.
    // Reuse the central structural remapper so formulas, names, validations,
    // CF, merges, tables and object anchors all move together. Because that is
    // a whole-row coordinate mutation, authenticate its full affected domain.
    if (
      context.rangeHasLockedCell(sheet, {
        r1: parsed.r1,
        c1: 0,
        r2: Math.max(parsed.r2, sheet.rowCount - 1),
        c2: Math.max(parsed.c2, sheet.columnCount - 1),
      })
    ) {
      return "protected";
    }
    nextSheet = insertLine(expandedSheet, "row", parsed.r1, 1);
    if (nextSheet === expandedSheet) return "rowLimit";
    for (let col = parsed.c1; col <= parsed.c2; col += 1) {
      nextSheet = updateCellValue(nextSheet, parsed.r1, col, `${spreadsheetDefaultNames().column}${col - parsed.c1 + 1}`);
    }
    tableRange = { ...parsed, r2: parsed.r2 + 1 };
  }
  // A table holds no merged cell (Merge is greyed inside one, user audit
  // S-019): Excel's Create Table unmerges the merged cells it covers ("This
  // operation will cause some merged cells to unmerge"), each keeping its
  // value in its upper-left cell. A generated header row would otherwise be
  // written into the covered cells of a merge it ran through (sheet model
  // fuzz), where the loader discards it.
  const mergesBefore = normalizeMerges(nextSheet.merges);
  if (rangeHasMerge(mergesBefore, tableRange)) {
    const kept = removeMergesInRange(mergesBefore, tableRange);
    nextSheet = { ...nextSheet, cells: cellsAfterUnmerge(nextSheet.cells, mergesBefore, kept) as typeof nextSheet.cells };
    if (kept.length > 0) nextSheet.merges = kept;
    else delete nextSheet.merges;
  }

  // Excel table columns are non-blank and case-insensitively unique. Repair
  // only the header cells that need it, inside this same Create transaction,
  // so the persisted names, structured-reference resolver and XLSX writer all
  // consume one stable identity. A header cell that holds a formula gives its
  // DISPLAYED text as the column name and is frozen to that text: the
  // column's identity cannot keep changing under the structured references
  // (live audit B107).
  const rawHeaders: string[] = [];
  const headerIsFormula: boolean[] = [];
  for (let col = tableRange.c1; col <= tableRange.c2; col += 1) {
    const cell = nextSheet.cells[`${tableRange.r1}:${col}`];
    const formula = Boolean(cell && spreadsheetCellIsFormula(cell));
    headerIsFormula.push(formula);
    rawHeaders.push(
      formula
        ? context.headerDisplay(cell, tableRange.r1, col).trim()
        : String(cell?.value ?? "").trim(),
    );
  }
  const tableHeaders = normalizeTableColumnNames(rawHeaders);
  for (let offset = 0; offset < tableHeaders.length; offset += 1) {
    if (tableHeaders[offset] !== rawHeaders[offset] || headerIsFormula[offset]) {
      nextSheet = updateCellValue(nextSheet, tableRange.r1, tableRange.c1 + offset, tableHeaders[offset]);
    }
  }

  const nextTables = sheetTables(nextSheet);
  const overlaps = nextTables.some((candidate) => {
    const existing = candidate.range;
    const occupiedBottom = existing.r2 + (candidate.totalsRow === true ? 1 : 0);
    // A totals row is part of the table's occupied worksheet domain.
    // Authenticate that derived coordinate independently; a malformed or
    // out-of-grid table fails closed instead of authorizing a second table
    // over ambiguous custody.
    if (
      !Number.isSafeInteger(occupiedBottom) ||
      occupiedBottom < existing.r2 ||
      occupiedBottom >= MAX_GRID_ROWS
    ) {
      return true;
    }
    return (
      tableRange.r1 <= occupiedBottom &&
      tableRange.r2 >= existing.r1 &&
      tableRange.c1 <= existing.c2 &&
      tableRange.c2 >= existing.c1
    );
  });
  if (overlaps) return "overlap";

  const workbookTableNames = new Set<string>();
  for (const candidate of workbook) {
    for (const table of sheetTables(candidate)) {
      workbookTableNames.add(table.name.toLowerCase());
    }
  }
  // "Table1", "表1" in Chinese: the interface language's Excel name.
  const tablePrefix = spreadsheetDefaultNames().table;
  let tableOrdinal = 1;
  while (workbookTableNames.has(`${tablePrefix}${tableOrdinal}`.toLowerCase())) tableOrdinal += 1;
  const table = buildDataTableFromRange(tableRange, `${tablePrefix}${tableOrdinal}`, tableHeaders);
  if (!table) return "invalidRange";
  return {
    sheet: {
      ...nextSheet,
      dataTables: [...nextTables, tableWithStyleChoice(table, request.styleId)],
    },
    range: tableRange,
    insertedHeaderRow: !request.hasHeaders,
  };
}

/**
 * Table Design > Table Styles: the table drawn in `styleId`, a built-in or
 * one of the workbook's own styles (only `tableStyleInfo@name` changes; the
 * workbook's style definitions stay as written).
 */
export function tableStyledSheet<S extends TableButtonSheet>(
  sheet: S,
  id: string,
  styleId: TableStyleChoice,
  rangeHasLockedCell: LockedCellTest<S>,
): S | "protected" | null {
  const dataTables = sheetTables(sheet);
  const current = dataTables.find((table) => table.id === id);
  if (!current) return null;
  const chosen = tableWithStyleChoice(current, styleId);
  if (
    chosen.styleId === current.styleId &&
    chosen.nativeStyleName === current.nativeStyleName
  ) {
    return null;
  }
  if (rangeHasLockedCell(sheet, tableOccupiedRange(current))) return "protected";
  const next = dataTables.map((table) => (table.id === id ? chosen : table));
  return { ...sheet, dataTables: next };
}

/** The Table Style Options a check box sets directly on the table. */
export type TableStyleOption = "bandedRows" | "bandedColumns" | "firstColumn" | "lastColumn";

/** Table Design > Table Style Options: Banded Rows / Columns, First / Last Column. */
export function tableStyleOptionSheet<S extends TableButtonSheet>(
  sheet: S,
  id: string,
  option: TableStyleOption,
  nextValue: boolean,
  rangeHasLockedCell: LockedCellTest<S>,
): S | "protected" | null {
  const dataTables = sheetTables(sheet);
  const current = dataTables.find((table) => table.id === id);
  if (!current || Boolean(current[option]) === nextValue) return null;
  if (rangeHasLockedCell(sheet, tableOccupiedRange(current))) return "protected";
  const next = dataTables.map((table) =>
    table.id === id
      ? nextValue
        ? { ...table, [option]: true }
        : (() => {
            const updated = { ...table };
            delete updated[option];
            return updated;
          })()
      : table,
  );
  return { ...sheet, dataTables: next };
}

/** Table Design > Filter Button: shows or hides the table's header arrows. */
export function tableFilterButtonToggledSheet<S extends TableButtonSheet>(
  sheet: S,
  id: string,
  rangeHasLockedCell: LockedCellTest<S>,
): S | "protected" | null {
  const dataTables = sheetTables(sheet);
  const current = dataTables.find((table) => table.id === id);
  if (!current || !current.headerRow || current.nativeResizeBlocker) return null;
  if (rangeHasLockedCell(sheet, current.range)) return "protected";
  const nextTable: DataTable = { ...current };
  if (current.showFilterButton === false) delete nextTable.showFilterButton;
  else nextTable.showFilterButton = false;
  return {
    ...sheet,
    dataTables: dataTables.map((table) => (table.id === id ? nextTable : table)),
  };
}

/**
 * Table Design > Table Name. Tables and defined names share the workbook's
 * name space (Excel refuses a table name a defined name already holds; live
 * audit B111). Renaming updates every reference to the table in the workbook
 * (qualified, bare whole-table, recipes on any sheet, names, rules) in one
 * step. "same" when the name is unchanged; null when refused.
 */
export function tableRenamedSheets<S extends TableButtonSheet>(
  workbook: readonly S[],
  activeSheetIndex: number,
  id: string,
  rawName: string,
): S[] | "same" | null {
  const nextName = rawName.trim();
  const active = workbook[activeSheetIndex];
  if (!active) return null;
  const current = sheetTables(active).find((table) => table.id === id);
  if (!current) return null;
  if (current.name === nextName) return "same";
  if (!isValidDataTableName(nextName)) return null;
  const nameInUse = workbook.some(
    (candidate, sheetIndex) =>
      (candidate.dataTables ?? []).some(
        (table) =>
          !(sheetIndex === activeSheetIndex && table.id === id) &&
          table.name.toLowerCase() === nextName.toLowerCase(),
      ) ||
      normalizeNamedRanges(candidate.namedRanges).some(
        (entry) => entry.name.toLowerCase() === nextName.toLowerCase(),
      ),
  );
  if (nameInUse) return null;
  return workbook.map((candidate, sheetIndex) => {
    const rewritten = rewriteTableReferencesInSheet(candidate, (formula) =>
      renameDataTableReferences(formula, current.name, nextName),
    );
    if (sheetIndex !== activeSheetIndex) return rewritten;
    return {
      ...rewritten,
      dataTables: (rewritten.dataTables ?? []).map((table) =>
        table.id === id ? { ...table, name: nextName } : table,
      ),
    };
  });
}

/** Why Convert to Range refused. */
export type ConvertTableRefusal =
  | "slicer"
  | "native"
  | "customStyle"
  | "protected"
  | "customTotals"
  | "formula";

/**
 * Table Design > Convert to Range. Every structured reference to the table
 * becomes an absolute A1 reference across the workbook, while its rendered
 * style is materialized onto ordinary cells before the table identity
 * disappears. Unsupported native authority or an ambiguous formula fails
 * closed, so conversion never leaves #NAME? references or drops appearance.
 * Null when the table or its header names cannot be read.
 */
export function tableConvertedSheets<S extends TableButtonSheet>(
  workbook: readonly S[],
  activeSheetIndex: number,
  id: string,
  context: {
    rangeHasLockedCell: LockedCellTest<S>;
    themeAccents: readonly string[];
    /** The workbook's own table styles (view options `nativeTableStyles`). */
    nativeStyles?: readonly SpreadsheetNativeTableStyle[];
  },
): S[] | ConvertTableRefusal | null {
  const sheet = workbook[activeSheetIndex];
  if (!sheet) return null;
  const dataTables = sheetTables(sheet);
  const current = dataTables.find((table) => table.id === id);
  if (!current) return null;
  if (
    normalizeSlicers(sheet.slicers, sheet.columnCount).some(
      (slicer) => slicer.source?.kind === "table" && slicer.source.tableId === current.id,
    )
  ) {
    return "slicer";
  }
  if (current.nativeResizeBlocker) return "native";
  const sourceStyle = current.nativeStyleName?.trim();
  const workbookStyle = tableNativeStyle(current, context.nativeStyles);
  // A workbook style the page paints exactly materializes like a built-in
  // (Excel keeps the table's look on the cells it leaves behind); one the
  // page could not read exactly would leave the cells looking different.
  if (
    sourceStyle &&
    sourceStyle.toLowerCase() !== tableStylePreset(current.styleId).excelStyleName.toLowerCase() &&
    (!workbookStyle || workbookStyle.exact === false)
  ) {
    return "customStyle";
  }
  if (context.rangeHasLockedCell(sheet, tableOccupiedRange(current))) return "protected";
  const headers = tableHeadersForCells(current, sheet.cells);
  if (!headers) return null;
  if (current.totalsRow) {
    const totalsRow = current.range.r2 + 1;
    for (let offset = 0; offset < headers.length; offset += 1) {
      const col = current.range.c1 + offset;
      const value = String(sheet.cells[`${totalsRow}:${col}`]?.value ?? "");
      if (!tableTotalCellOwnsValue(current, offset, headers[offset], value)) {
        return "customTotals";
      }
    }
  }
  const tableDefinition: TableDefinition = {
    name: current.name,
    range: current.range,
    hasHeader: current.headerRow,
    hasTotals: current.totalsRow === true,
    columns: headers,
  };
  const convertedSheets: S[] = [];
  for (const candidate of workbook) {
    const converted = convertRemovedTableReferencesInSheet(candidate, tableDefinition, sheet.name);
    if (!converted) return "formula";
    convertedSheets.push(converted);
  }
  let nextSheet = materializeDataTableStyle(
    convertedSheets[activeSheetIndex] ?? sheet,
    current,
    context.themeAccents,
    context.nativeStyles,
  );
  const next = dataTables.filter((table) => table.id !== id);
  nextSheet = { ...nextSheet };
  if (next.length > 0) nextSheet.dataTables = next;
  else delete nextSheet.dataTables;
  convertedSheets[activeSheetIndex] = nextSheet;
  return convertedSheets;
}

/** A cell the Total Row may not take: one with content. A format alone
 * (a blank cell set to two decimals) leaves it free (live audit S-180). */
function totalsRowCellHasContent(cell: SpreadsheetCell): boolean {
  return String(cell.value ?? "") !== "" || Boolean(cell.formula) || Boolean(cell.hyperlink) ||
    Boolean((cell as { picture?: unknown }).picture);
}

/** Why the Total Row check box refused. */
export type TableTotalsRefusal = "invalidTarget" | "protected" | "occupied" | "customTotals";

/**
 * Table Design > Total Row. On: a row under the table with a label in the
 * first column and a SUBTOTAL in the LAST one only, SUM over numbers, COUNT
 * otherwise (Microsoft Support, "Total the data in an Excel table"); every
 * total reads in its column's number format. Off: the row's owned cells are
 * given back. Null when the table cannot be found or resized natively.
 */
export function tableTotalsToggledSheet<S extends TableButtonSheet>(
  sheet: S,
  id: string,
  context: {
    rangeHasLockedCell: LockedCellTest<S>;
    /** The cell's numeric value, or null when it holds no number. */
    numericAt: (cell: SpreadsheetCell, row: number, col: number) => number | null;
    /** The first column's label, "Total" in English. */
    totalsLabel: string;
  },
): S | TableTotalsRefusal | null {
  const dataTables = sheetTables(sheet);
  const current = dataTables.find((table) => table.id === id);
  // A table whose own filter or sort MRAI carries without reading (filter
  // custody) takes a Total Row like any other.
  if (!current || tableNativeStructureBlocked(current)) return null;
  const headers = tableHeadersForCells(current, sheet.cells);
  const totalsRow = current.range.r2 + 1;
  const totalsRange = { r1: totalsRow, r2: totalsRow, c1: current.range.c1, c2: current.range.c2 };
  const conflictsWithTable = dataTables.some((table) => {
    if (table.id === id) return false;
    const otherBottom = table.range.r2 + (table.totalsRow ? 1 : 0);
    return (
      totalsRange.r1 <= otherBottom &&
      totalsRange.r2 >= table.range.r1 &&
      totalsRange.c1 <= table.range.c2 &&
      totalsRange.c2 >= table.range.c1
    );
  });
  if (
    !headers ||
    totalsRow >= MAX_GRID_ROWS ||
    conflictsWithTable ||
    rangeHasMerge(normalizeMerges(sheet.merges), totalsRange)
  ) {
    return "invalidTarget";
  }
  if (context.rangeHasLockedCell(sheet, totalsRange)) return "protected";

  if (!current.totalsRow) {
    for (let col = current.range.c1; col <= current.range.c2; col += 1) {
      // Content takes the row; a format alone does not (a blank cell set to
      // two decimals under the table refused the Total Row, live audit S-180):
      // the total is written into it and keeps the cell's look.
      const cell = sheet.cells[`${totalsRow}:${col}`];
      if (cell && totalsRowCellHasContent(cell)) return "occupied";
    }
    const lastOffset = current.range.c2 - current.range.c1;
    const lastColumnIsNumeric = (() => {
      const firstDataRow = current.range.r1 + (current.headerRow ? 1 : 0);
      for (let row = firstDataRow; row <= current.range.r2; row += 1) {
        const cell = sheet.cells[`${row}:${current.range.c2}`];
        if (cell === undefined || String(cell.value ?? "").trim() === "") continue;
        return context.numericAt(cell, row, current.range.c2) !== null;
      }
      return false;
    })();
    const nextTable: DataTable =
      current.totals || current.totalLabels
        ? { ...current, totalsRow: true }
        : {
            ...current,
            totalsRow: true,
            totalLabels: { 0: context.totalsLabel },
            totals: { [lastOffset]: lastColumnIsNumeric ? "sum" : "count" },
          };
    let nextSheet: S = { ...sheet, rowCount: Math.max(sheet.rowCount, totalsRow + 1) };
    for (let offset = 0; offset < headers.length; offset += 1) {
      const value = tableTotalCellValue(nextTable, offset, headers[offset]);
      if (value) {
        nextSheet = updateCellValue(nextSheet, totalsRow, current.range.c1 + offset, value);
      }
      // The total reads in its column's number format.
      nextSheet = applyTableTotalNumberFormat(nextSheet, nextTable, offset, totalsRow);
    }
    return {
      ...nextSheet,
      dataTables: dataTables.map((table) => (table.id === id ? nextTable : table)),
    };
  }

  for (let offset = 0; offset < headers.length; offset += 1) {
    const col = current.range.c1 + offset;
    const value = String(sheet.cells[`${totalsRow}:${col}`]?.value ?? "");
    if (!tableTotalCellOwnsValue(current, offset, headers[offset], value)) return "customTotals";
  }
  let nextSheet = sheet;
  for (let offset = 0; offset < headers.length; offset += 1) {
    nextSheet = releaseTableTotalCell(nextSheet, current, offset, totalsRow);
  }
  const nextTable = { ...current };
  delete nextTable.totalsRow;
  return {
    ...nextSheet,
    dataTables: dataTables.map((table) => (table.id === id ? nextTable : table)),
  };
}

// ------------------------------------------------------------ Resize Table
// Moved from SpreadsheetWindow (behaviour unchanged): the Resize Table dialog,
// table AutoExpansion, Tab at a table's last cell, the paste's AutoExpansion and
// the spreadsheet editor commands share it.

/**
 * Why a Resize Table cannot be applied. Each is a state the planner cannot
 * rewrite without losing something, and each has its own message:
 *   native-*         a native table part MRAI keeps byte-for-byte (an
 *                    unsupported Total Row recipe, a query, extension data);
 *   header-row       Excel's rule: the headers stay in their row and the new
 *                    range overlaps the old one;
 *   first-column / totals-rows-and-columns  no longer produced: a Total Row
 *                    follows any resize, its totals rebased with the left
 *                    edge (kept for persisted receipts and their messages);
 *   filtered-column  the resize would drop a column that carries a filter;
 *   totals-custom    a Totals Row cell holds content the table does not own;
 *   totals-target    the row the Total Row moves to is not empty;
 *   blocked-cells    merged or protected cells in the old or new range;
 *   overlap-table    the new range runs into another table;
 *   bounds / stale   out of the sheet, or the table changed.
 */
export type DataTableResizeRefusal =
  | "native-totals"
  | "native-query"
  | "native-extension"
  | "header-row"
  | "first-column"
  | "totals-rows-and-columns"
  | "filtered-column"
  | "totals-custom"
  | "totals-target"
  | "blocked-cells"
  | "overlap-table"
  | "bounds"
  | "stale";

export type DataTableResizeOutcome<S = TableButtonSheet> =
  | { ok: true; sheet: S }
  | { ok: false; reason: DataTableResizeRefusal };

/** The refusal a persisted `nativeResizeBlocker` stands for. */
export function dataTableNativeResizeRefusal(
  blocker: DataTable["nativeResizeBlocker"],
): DataTableResizeRefusal | null {
  if (!blocker) return null;
  if (blocker === "totals") return "native-totals";
  if (blocker === "query") return "native-query";
  // "extension". "activeFilter" / "sortState" mark filter custody (the
  // file's own criteria or sort, carried by the writer): Resize Table takes
  // them (planDataTableResizeOutcome asks tableNativeStructureBlocked first);
  // a command that would move rows under such a criterion still refuses.
  return "native-extension";
}

/**
 * Resize Table (Table Design > Resize Table) as one pure sheet transaction:
 * the new range, column identities, a moved or widened Totals Row and the
 * calculated-column formulas of the rows it adds. Returns the same sheet for
 * an unchanged range and null when the resize is refused. planDataTableResize
 * is shared by the Resize Table dialog, table AutoExpansion and Tab at a
 * table's last cell, so all three keep the same fail-closed rules
 * (`nativeResizeBlocker`, protection, merges, overlapping tables, filtered
 * columns, owned Totals Row cells). A table's filters and sort are keyed by
 * column identity, so they follow the resize as Excel's do; a Total Row moves
 * to stay directly under the last row. A column the resize adds under a blank
 * header cell is named "Column1", "Column2", ... (`defaultColumnName` gives
 * the localized text, "列1" in Chinese) with the first free number, and the
 * name is written into the header cell, as Excel does (Microsoft Support,
 * "Resize a table by adding or removing rows and columns").
 */
export function planDataTableResize<S extends TableButtonSheet>(
  sheet: S,
  id: string,
  parsed: { r1: number; c1: number; r2: number; c2: number },
  rangeHasLockedCell: (
    sheet: S,
    range: { r1: number; r2: number; c1: number; c2: number },
  ) => boolean,
  options?: { defaultColumnName?: (index: number) => string },
): S | null {
  const outcome = planDataTableResizeOutcome(
    sheet,
    id,
    parsed,
    rangeHasLockedCell,
    options,
  );
  return outcome.ok ? outcome.sheet : null;
}

/** {@link planDataTableResize}, saying why when it refuses. */
export function planDataTableResizeOutcome<S extends TableButtonSheet>(
  sheet: S,
  id: string,
  parsed: { r1: number; c1: number; r2: number; c2: number },
  rangeHasLockedCell: (
    sheet: S,
    range: { r1: number; r2: number; c1: number; c2: number },
  ) => boolean,
  options?: { defaultColumnName?: (index: number) => string },
): DataTableResizeOutcome<S> {
  const refuse = (reason: DataTableResizeRefusal): DataTableResizeOutcome<S> => ({
    ok: false,
    reason,
  });
  const dataTables = normalizeDataTables(sheet.dataTables, {
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  });
  const mergeSpans = normalizeMerges(sheet.merges);
  const current = dataTables.find((table) => table.id === id);
  if (!current) return refuse("stale");
  // A table the file filters or sorts in ways MRAI carries without reading
  // (filter custody) reshapes like any other: the writer re-points its
  // criteria and sort to the new range.
  const nativeRefusal = tableNativeStructureBlocked(current)
    ? dataTableNativeResizeRefusal(current.nativeResizeBlocker)
    : null;
  if (nativeRefusal) return refuse(nativeRefusal);
  // The sheet's address space, not its stored extent, bounds the table: a
  // blank row below the last stored cell can join it (live audit B159); the
  // extent grows to hold the table below.
  if (parsed.r2 >= MAX_GRID_ROWS || parsed.c2 >= MAX_GRID_COLUMNS) {
    return refuse("bounds");
  }
  if (
    parsed.r1 === current.range.r1 &&
    parsed.c1 === current.range.c1 &&
    parsed.r2 === current.range.r2 &&
    parsed.c2 === current.range.c2
  ) {
    return { ok: true, sheet };
  }
  // Microsoft Support, "Resize a table by adding or removing rows and
  // columns": "Table headers can't move to a different row, and the new
  // range must overlap the original range."
  if (
    parsed.r1 !== current.range.r1 ||
    parsed.c1 > current.range.c2 ||
    parsed.c2 < current.range.c1
  ) {
    return refuse("header-row");
  }
  // A Total Row stays directly under the last row and each total stays with
  // its column, whichever edges move: Excel's Resize Table adds or removes
  // columns on either side, rows with them, the header row staying put
  // (Microsoft Support, "Resize a table by adding or removing rows and
  // columns"). Offsets are rebased by how far the left edge moved.
  const horizontalResize =
    parsed.c1 !== current.range.c1 || parsed.c2 !== current.range.c2;
  const horizontalTotalsResize = Boolean(current.totalsRow && horizontalResize);
  const leftShift = current.range.c1 - parsed.c1;
  const nextTable: DataTable = { ...current, range: parsed };
  if (current.calculatedFormulas && horizontalResize) {
    // An edge moved (Excel lets either, the headers staying in their row):
    // each calculated column stays in its worksheet column, so only its
    // offset from the table's left edge changes; one the resize drops goes.
    // A right edge moved in kept the dropped column's formula, which the
    // loader then discarded (sheet model fuzz).
    const formulas: Record<number, string> = {};
    for (const [rawOffset, formula] of Object.entries(
      current.calculatedFormulas,
    )) {
      const offset = Number(rawOffset) + current.range.c1 - parsed.c1;
      if (offset >= 0 && offset <= parsed.c2 - parsed.c1) {
        formulas[offset] = formula;
      }
    }
    if (Object.keys(formulas).length > 0) nextTable.calculatedFormulas = formulas;
    else delete nextTable.calculatedFormulas;
  }
  // Header cells the resize fills with a default column name (TD: Excel
  // writes "Column1"... into the blank header of a column it adds).
  const headerWrites: Array<{ col: number; name: string }> = [];
  if (horizontalResize) {
    const nextWidth = parsed.c2 - parsed.c1 + 1;
    const currentNames = tableHeadersForCells(current, sheet.cells);
    if (!currentNames) return refuse("stale");
    const existingColumns = current.columns ??
      currentNames.map((name, index) => ({ id: index + 1, name }));
    // Names already taken: every kept column plus the header text of every
    // added column, so a default name never collides with either.
    const takenNames: string[] = [];
    const addedColumns: number[] = [];
    for (let offset = 0; offset < nextWidth; offset += 1) {
      const absoluteColumn = parsed.c1 + offset;
      const currentOffset = absoluteColumn - current.range.c1;
      if (currentOffset >= 0 && currentOffset < existingColumns.length) {
        takenNames.push(existingColumns[currentOffset].name);
        continue;
      }
      addedColumns.push(absoluteColumn);
      const headerText = current.headerRow
        ? String(
            sheet.cells[`${current.range.r1}:${absoluteColumn}`]?.value ?? "",
          ).trim()
        : "";
      if (headerText !== "") takenNames.push(headerText);
    }
    const addedNames = new Map<number, string>();
    for (const absoluteColumn of addedColumns) {
      const headerText = current.headerRow
        ? String(
            sheet.cells[`${current.range.r1}:${absoluteColumn}`]?.value ?? "",
          ).trim()
        : "";
      if (headerText !== "") {
        addedNames.set(absoluteColumn, headerText);
        continue;
      }
      const name = defaultTableColumnName(
        takenNames,
        options?.defaultColumnName,
      );
      takenNames.push(name);
      addedNames.set(absoluteColumn, name);
      if (current.headerRow) {
        headerWrites.push({ col: absoluteColumn, name });
      }
    }
    const keptColumn = (offset: number) => {
      const currentOffset = parsed.c1 + offset - current.range.c1;
      return currentOffset >= 0 && currentOffset < existingColumns.length;
    };
    const rawNames = Array.from({ length: nextWidth }, (_, offset) => {
      const absoluteColumn = parsed.c1 + offset;
      const currentOffset = absoluteColumn - current.range.c1;
      if (keptColumn(offset)) return existingColumns[currentOffset].name;
      return addedNames.get(absoluteColumn) ?? "";
    });
    // The kept columns own their names; an added header that repeats one
    // (in any case), or is too long, is renamed, never a kept column.
    const nextNames = normalizeTableColumnNames(rawNames, {
      keep: rawNames.map((_, offset) => keptColumn(offset)),
    });
    // What the header cell shows is the column's identity: an added header
    // the normalization changed ("Amount" beside "Amount" -> "Amount2") is
    // written back, exactly as a blank one receives its "Column1".
    if (current.headerRow) {
      const pending = new Set(headerWrites.map((write) => write.col));
      nextNames.forEach((name, offset) => {
        if (keptColumn(offset)) return;
        const absoluteColumn = parsed.c1 + offset;
        if (pending.has(absoluteColumn)) return;
        const shown = String(
          sheet.cells[`${current.range.r1}:${absoluteColumn}`]?.value ?? "",
        ).trim();
        if (shown !== name) headerWrites.push({ col: absoluteColumn, name });
      });
    }
    const maxColumnId = Math.max(
      0,
      ...existingColumns.map((column) => column.id),
    );
    let nextColumnId = maxColumnId + 1;
    nextTable.columns = nextNames.map((name, offset) => ({
      id: (() => {
        const currentOffset = parsed.c1 + offset - current.range.c1;
        if (currentOffset >= 0 && currentOffset < existingColumns.length) {
          return existingColumns[currentOffset].id;
        }
        const id = nextColumnId;
        nextColumnId += 1;
        return id;
      })(),
      name,
    }));
    const nextColumnIds = new Set(
      nextTable.columns.map((column) => String(column.id)),
    );
    if (
      Object.keys(current.columnFilters ?? {}).some(
        (columnId) => !nextColumnIds.has(columnId),
      ) ||
      (current.nativeFilterColumnIds ?? []).some(
        (columnId) => !nextColumnIds.has(String(columnId)),
      )
    ) {
      // Removing a filtered table column would silently destroy its native
      // criterion (a modelled one, or one the file keeps that MRAI carries
      // without reading). Keep the resize atomic and fail closed instead.
      return refuse("filtered-column");
    }
  }
  if (horizontalTotalsResize) {
    const nextWidth = parsed.c2 - parsed.c1 + 1;
    const rebased = (rawOffset: string): number | null => {
      const offset = Number(rawOffset) + leftShift;
      return offset >= 0 && offset < nextWidth ? offset : null;
    };
    if (current.totals) {
      const totals: Record<number, TableTotalKind> = {};
      for (const [rawOffset, kind] of Object.entries(current.totals)) {
        const offset = rebased(rawOffset);
        if (offset !== null) totals[offset] = kind;
      }
      if (Object.keys(totals).length > 0) nextTable.totals = totals;
      else delete nextTable.totals;
    }
    if (current.totalLabels) {
      const labels: Record<number, string> = {};
      for (const [rawOffset, label] of Object.entries(
        current.totalLabels,
      )) {
        const offset = rebased(rawOffset);
        if (offset !== null) labels[offset] = label;
      }
      if (Object.keys(labels).length > 0) {
        nextTable.totalLabels = labels;
      } else {
        delete nextTable.totalLabels;
      }
    }
    if (current.totalFormulas) {
      const formulas: Record<number, string> = {};
      for (const [rawOffset, formula] of Object.entries(
        current.totalFormulas,
      )) {
        const offset = rebased(rawOffset);
        if (offset !== null) formulas[offset] = formula;
      }
      if (Object.keys(formulas).length > 0) {
        nextTable.totalFormulas = formulas;
      } else {
        delete nextTable.totalFormulas;
      }
    }
    if (current.calculatedFormulas) {
      const formulas: Record<number, string> = {};
      for (const [rawOffset, formula] of Object.entries(
        current.calculatedFormulas,
      )) {
        const offset = rebased(rawOffset);
        if (offset !== null) formulas[offset] = formula;
      }
      if (Object.keys(formulas).length > 0) {
        nextTable.calculatedFormulas = formulas;
      } else {
        delete nextTable.calculatedFormulas;
      }
    }
  }
  const nextHeaders = tableHeadersForCells(nextTable, sheet.cells);
  if (!nextHeaders) return refuse("stale");
  const nextOccupiedBottom = parsed.r2 + (current.totalsRow ? 1 : 0);
  if (nextOccupiedBottom >= MAX_GRID_ROWS) return refuse("bounds");
  const overlapsAnotherTable = dataTables.some(
    (table) =>
      table.id !== id &&
      parsed.r1 <= table.range.r2 + (table.totalsRow ? 1 : 0) &&
      nextOccupiedBottom >= table.range.r1 &&
      parsed.c1 <= table.range.c2 &&
      parsed.c2 >= table.range.c1,
  );
  if (overlapsAnotherTable) return refuse("overlap-table");
  const currentOccupied = {
    ...current.range,
    r2: current.range.r2 + (current.totalsRow ? 1 : 0),
  };
  const nextOccupied = { ...parsed, r2: nextOccupiedBottom };
  if (
    rangeHasMerge(mergeSpans, nextOccupied) ||
    rangeHasLockedCell(sheet, currentOccupied) ||
    rangeHasLockedCell(sheet, nextOccupied)
  ) {
    return refuse("blocked-cells");
  }

  let nextSheet = sheet;
  for (const { col, name } of headerWrites) {
    nextSheet = updateCellValue(nextSheet, current.range.r1, col, name);
  }
  if (current.totalsRow) {
    const currentHeaders = tableHeadersForCells(current, sheet.cells);
    if (!currentHeaders) return refuse("stale");
    const currentTotalsRow = current.range.r2 + 1;
    const nextTotalsRow = parsed.r2 + 1;
    for (let offset = 0; offset < currentHeaders.length; offset += 1) {
      const col = current.range.c1 + offset;
      const value = String(
        sheet.cells[`${currentTotalsRow}:${col}`]?.value ?? "",
      );
      if (
        !tableTotalCellOwnsValue(
          current,
          offset,
          currentHeaders[offset],
          value,
        )
      ) {
        return refuse("totals-custom");
      }
    }
    for (let offset = 0; offset < nextHeaders.length; offset += 1) {
      const col = parsed.c1 + offset;
      const targetIsOwnedSource =
        nextTotalsRow === currentTotalsRow &&
        col >= current.range.c1 &&
        col <= current.range.c2;
      if (targetIsOwnedSource) continue;
      const target = sheet.cells[`${nextTotalsRow}:${col}`];
      if (target && totalsRowCellHasContent(target)) {
        return refuse("totals-target");
      }
    }
    if (nextTotalsRow !== currentTotalsRow || horizontalTotalsResize) {
      // A row the grown table now covers keeps its column format as a data
      // row; a row left outside the table is given back as it was found.
      const releasedRowLeavesTable =
        currentTotalsRow > parsed.r2 || currentTotalsRow < parsed.r1;
      for (let col = current.range.c1; col <= current.range.c2; col += 1) {
        nextSheet =
          releasedRowLeavesTable || col > parsed.c2 || col < parsed.c1
            ? releaseTableTotalCell(
                nextSheet,
                current,
                col - current.range.c1,
                currentTotalsRow,
              )
            : updateCellValue(nextSheet, currentTotalsRow, col, "");
      }
    }
    nextSheet = {
      ...nextSheet,
      rowCount: Math.max(nextSheet.rowCount, nextTotalsRow + 1),
    };
    // Every owned total is rewritten in its structured-reference form,
    // =SUBTOTAL(109,[Amount]): Excel keeps a structured reference through a
    // resize, so the total follows the column rather than a row span
    // (Microsoft Support, "Using structured references with Excel tables").
    // An imported A1 form is owned too (tableTotalCellOwnsValue) and comes
    // out structured here.
    for (let offset = 0; offset < nextHeaders.length; offset += 1) {
      const col = parsed.c1 + offset;
      nextSheet = updateCellValue(
        nextSheet,
        nextTotalsRow,
        col,
        tableTotalCellValue(nextTable, offset, nextHeaders[offset]),
      );
      nextSheet = applyTableTotalNumberFormat(
        nextSheet,
        nextTable,
        offset,
        nextTotalsRow,
      );
    }
  }
  if (current.calculatedFormulas) {
    // Excel leaves every cell already in a calculated column as it is,
    // exceptions included, and fills only the rows the resize adds; a new
    // row that already holds a value keeps it as one more exception
    // (Microsoft Support, "Use calculated columns in an Excel table").
    // r1 and the Header Row are fixed here, so the first data row is the
    // same before and after; offsets are the remapped ones above. The Total
    // Row has already moved below the new last row, so its old row is
    // filled like an added one.
    const firstDataRow = parsed.r1 + (nextTable.headerRow ? 1 : 0);
    for (const [rawOffset, canonical] of Object.entries(
      nextTable.calculatedFormulas ?? {},
    )) {
      const offset = Number(rawOffset);
      if (offset > parsed.c2 - parsed.c1) continue;
      const col = parsed.c1 + offset;
      for (
        let row = Math.max(firstDataRow, current.range.r2 + 1);
        row <= parsed.r2;
        row += 1
      ) {
        // Read after the Total Row moved: the row it left is a data row now
        // and takes the column's formula like any added row.
        if (String(nextSheet.cells[`${row}:${col}`]?.value ?? "") !== "") {
          continue;
        }
        nextSheet = updateCellValue(
          nextSheet,
          row,
          col,
          shiftFormula(canonical, row - firstDataRow, 0),
        );
      }
    }
  }
  // A total written under the table — =SUM(C4:C10) in C12 — covers the C
  // column's whole data range, so Excel grows it with the table: C4:C11 once
  // row 11 belongs to it. Only a downward growth of the same columns does
  // this, and only to formulas outside the table itself (its calculated
  // columns and Totals Row are written from their own recipes above).
  if (
    parsed.r1 === current.range.r1 &&
    parsed.c1 === current.range.c1 &&
    parsed.c2 === current.range.c2 &&
    parsed.r2 > current.range.r2
  ) {
    const firstDataRow = parsed.r1 + (nextTable.headerRow ? 1 : 0);
    const span = {
      firstColumn: parsed.c1,
      lastColumn: parsed.c2,
      firstDataRow,
      lastDataRow: current.range.r2,
      newLastDataRow: parsed.r2,
    };
    if (firstDataRow <= current.range.r2) {
      for (const [key, cell] of Object.entries(nextSheet.cells)) {
        const formula = String(cell?.value ?? "");
        if (!formula.startsWith("=")) continue;
        const [row, col] = key.split(":").map(Number);
        if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
        const insideTable =
          row >= parsed.r1 &&
          row <= nextOccupiedBottom &&
          col >= parsed.c1 &&
          col <= parsed.c2;
        if (insideTable) continue;
        const grown = extendRangeRefsForTableGrowth(formula, span);
        if (grown !== formula) {
          nextSheet = updateCellValue(nextSheet, row, col, grown);
        }
      }
    }
  }
  // Data validation and conditional formats on a column's data rows grow
  // with the table, as Excel's table expansion carries them to an added row
  // (Tab in the last cell, typing under the table, Resize Table): a rule
  // whose cells end at the last data row and lie within the table's columns
  // reaches the new last data row (audit P-TABLE-TAB-RULE-INHERIT). Its
  // top-left stays, so a relative rule formula reads the same cells.
  if (
    parsed.r1 === current.range.r1 &&
    parsed.c1 === current.range.c1 &&
    parsed.c2 === current.range.c2 &&
    parsed.r2 > current.range.r2
  ) {
    const firstDataRow = parsed.r1 + (nextTable.headerRow ? 1 : 0);
    const lastDataRow = current.range.r2 - (current.totalsRow ? 1 : 0);
    const newLastDataRow = parsed.r2 - (nextTable.totalsRow ? 1 : 0);
    const grows = (range: { r1: number; c1: number; r2: number; c2: number }) =>
      range.r2 === lastDataRow &&
      range.r1 >= firstDataRow &&
      range.c1 >= parsed.c1 &&
      range.c2 <= parsed.c2;
    const grown = <R extends { range: { r1: number; c1: number; r2: number; c2: number }; areas?: Array<{ r1: number; c1: number; r2: number; c2: number }> }>(rule: R): R => {
      const areas = rule.areas?.map((area) => (grows(area) ? { ...area, r2: newLastDataRow } : area));
      if (!grows(rule.range) && !(areas && areas.some((area, index) => area !== rule.areas![index]))) return rule;
      return {
        ...rule,
        range: grows(rule.range) ? { ...rule.range, r2: newLastDataRow } : rule.range,
        ...(areas ? { areas } : {}),
      };
    };
    if (newLastDataRow > lastDataRow && lastDataRow >= firstDataRow) {
      if (nextSheet.validations?.length) {
        nextSheet = { ...nextSheet, validations: nextSheet.validations.map((rule) => grown(rule as never) as typeof rule) };
      }
      if (nextSheet.conditionalFormats?.length) {
        nextSheet = { ...nextSheet, conditionalFormats: nextSheet.conditionalFormats.map((rule) => grown(rule as never) as typeof rule) };
      }
    }
  }
  const next = dataTables.map((table) =>
    table.id === id ? nextTable : table,
  );
  if (!next.some((table) => table.id === id)) return refuse("stale");
  return {
    ok: true,
    sheet: {
      ...nextSheet,
      dataTables: next,
      rowCount: Math.max(nextSheet.rowCount, nextOccupiedBottom + 1),
      columnCount: Math.max(nextSheet.columnCount, parsed.c2 + 1),
    },
  };
}

// --------------------------------------------------------- Header Row toggle

/** Why Table Design > Header Row could not be switched. */
export type TableHeaderRowRefusal =
  | "unreadable"
  | "locked"
  | "no-data-rows"
  | "row-above-taken"
  | "calculated-column"
  | "totals-row";

/**
 * Table Design > Header Row (verified against Excel 16 for Mac): turning the
 * header row off hides it - the header cells go blank, the table shrinks to
 * its data rows and the column names live on in the table, its filters and
 * sort leave with the header buttons; turning it back on grows the table one
 * row up and writes the names there again. Excel inserts cells when that row
 * is taken; here the row above must be free within the table's columns. A
 * calculated column and an owned Total Row move with the data.
 */
export function tableHeaderRowToggledSheet<
  S extends TableButtonSheet & { sortState?: { ref: string }; autoFilterRef?: string },
>(
  sheet: S,
  tables: readonly DataTable[],
  id: string,
  rangeHasLockedCell: LockedCellTest<S>,
): { ok: true; sheet: S } | { ok: false; reason: TableHeaderRowRefusal } | null {
  const current = tables.find((table) => table.id === id);
  if (!current) return null;
  const width = current.range.c2 - current.range.c1 + 1;
  const names = tableHeadersForCells(current, sheet.cells);
  const sortRange = sheet.sortState ? parseA1Range(sheet.sortState.ref.replace(/\$/g, "")) : null;
  const hasTableSort = Boolean(
    sheet.sortState &&
      (!sortRange ||
        (sortRange.r1 <= current.range.r2 &&
          sortRange.r2 >= current.range.r1 &&
          sortRange.c1 <= current.range.c2 &&
          sortRange.c2 >= current.range.c1)),
  );
  // A table whose own filter or sort MRAI carries without reading (filter
  // custody) switches its header like any other.
  if (!names || names.length !== width || tableNativeStructureBlocked(current)) {
    return { ok: false, reason: "unreadable" };
  }
  const occupied = {
    ...current.range,
    r1: Math.max(0, current.range.r1 - (current.headerRow ? 0 : 1)),
    r2: current.range.r2 + (current.totalsRow ? 1 : 0),
  };
  if (rangeHasLockedCell(sheet, occupied)) return { ok: false, reason: "locked" };
  const previousColumns = current.columns;
  const columns = names.map((name, index) => ({ id: previousColumns?.[index]?.id ?? index + 1, name }));
  const turningOff = current.headerRow;
  if (turningOff && current.range.r2 <= current.range.r1) return { ok: false, reason: "no-data-rows" };
  if (!turningOff) {
    const above = current.range.r1 - 1;
    let blocked = above < 0;
    for (let col = current.range.c1; !blocked && col <= current.range.c2; col += 1) {
      const cell = sheet.cells[`${above}:${col}`];
      blocked = String(cell?.value ?? "").trim() !== "" || String(cell?.formula ?? "").trim() !== "";
    }
    if (blocked) return { ok: false, reason: "row-above-taken" };
  }
  const nextTable: DataTable = {
    ...current,
    range: { ...current.range, r1: turningOff ? current.range.r1 + 1 : current.range.r1 - 1 },
    headerRow: !current.headerRow,
    columns,
  };
  let nextSheet = sheet;
  if (!nextTable.headerRow) {
    // The hidden header cells are cleared; the names stay in `columns`.
    for (let col = current.range.c1; col <= current.range.c2; col += 1) {
      nextSheet = updateCellValue(nextSheet, current.range.r1, col, "");
    }
    delete nextTable.columnFilters;
    delete nextTable.sortState;
    if (tableNativeFilterCustody(current)) {
      // "When you turn the header row off, AutoFilter is turned off and any
      // applied filters are removed from the table" (Microsoft Support): the
      // file's own filter and sort go with it (the writer removes them), and
      // the rows a criterion hid show again, as Excel shows every row of the
      // filter range when a filter is removed.
      delete nextTable.nativeResizeBlocker;
      delete nextTable.nativeFilterColumnIds;
      if (current.nativeResizeBlocker === "activeFilter") {
        const hidden = normalizeManualHidden(
          (sheet as { manuallyHiddenRows?: unknown }).manuallyHiddenRows,
          sheet.rowCount,
        );
        const visible = unhideLines(hidden, current.range.r1 + 1, current.range.r2);
        if (visible.size !== hidden.size) {
          const shown = serializeManualHidden(visible);
          const withRows = { ...nextSheet } as S & { manuallyHiddenRows?: number[] };
          if (shown) withRows.manuallyHiddenRows = shown;
          else delete withRows.manuallyHiddenRows;
          nextSheet = withRows;
        }
      }
    }
    // A sheet AutoFilter on another block of the same columns is not the
    // table's: only criteria whose range meets the table are cleared
    // (live audit B143).
    const sheetFilterRange = sheet.autoFilterRef ? parseA1Range(sheet.autoFilterRef.replace(/\$/g, "")) : null;
    const sheetFilterMeetsTable =
      !sheetFilterRange ||
      (sheetFilterRange.r1 <= current.range.r2 &&
        sheetFilterRange.r2 >= current.range.r1 &&
        sheetFilterRange.c1 <= current.range.c2 &&
        sheetFilterRange.c2 >= current.range.c1);
    const withoutTableColumns = <T,>(values: Record<string, T> | undefined): Record<string, T> | undefined => {
      if (!values || !sheetFilterMeetsTable) return values;
      const next = Object.fromEntries(
        Object.entries(values).filter(([rawColumn]) => {
          const column = Number(rawColumn);
          return !Number.isInteger(column) || column < current.range.c1 || column > current.range.c2;
        }),
      );
      return Object.keys(next).length > 0 ? next : undefined;
    };
    const filters = withoutTableColumns(sheet.filters);
    const colorFilters = withoutTableColumns(sheet.colorFilters);
    const conditionFilters = withoutTableColumns(sheet.conditionFilters);
    nextSheet = {
      ...nextSheet,
      ...(filters ? { filters } : {}),
      ...(colorFilters ? { colorFilters } : {}),
      ...(conditionFilters ? { conditionFilters } : {}),
    };
    if (!filters) delete nextSheet.filters;
    if (!colorFilters) delete nextSheet.colorFilters;
    if (!conditionFilters) delete nextSheet.conditionFilters;
    if (hasTableSort) delete nextSheet.sortState;
  }
  if (current.calculatedFormulas) {
    const currentFirstDataRow = current.range.r1 + (current.headerRow ? 1 : 0);
    const nextFirstDataRow = nextTable.range.r1 + (nextTable.headerRow ? 1 : 0);
    const rowDelta = nextFirstDataRow - currentFirstDataRow;
    const calculatedFormulas: Record<number, string> = {};
    for (const [rawOffset, canonical] of Object.entries(current.calculatedFormulas)) {
      const offset = Number(rawOffset);
      if (!Number.isInteger(offset) || offset < 0 || offset >= width) return null;
      const col = current.range.c1 + offset;
      for (let row = currentFirstDataRow; row <= current.range.r2; row += 1) {
        const expected = shiftFormula(canonical, row - currentFirstDataRow, 0);
        const cell = sheet.cells[`${row}:${col}`];
        if ((cell?.formula ?? cell?.value ?? "") !== expected) return { ok: false, reason: "calculated-column" };
      }
      const nextCanonical = shiftFormula(canonical, rowDelta, 0);
      calculatedFormulas[offset] = nextCanonical;
      for (let row = nextFirstDataRow; row <= nextTable.range.r2; row += 1) {
        nextSheet = updateCellValue(nextSheet, row, col, shiftFormula(nextCanonical, row - nextFirstDataRow, 0));
      }
    }
    nextTable.calculatedFormulas = calculatedFormulas;
  }
  if (nextTable.headerRow) {
    for (let offset = 0; offset < names.length; offset += 1) {
      nextSheet = updateCellValue(nextSheet, nextTable.range.r1, nextTable.range.c1 + offset, names[offset]);
    }
  }
  if (current.totalsRow) {
    const totalsRow = current.range.r2 + 1;
    for (let offset = 0; offset < names.length; offset += 1) {
      const col = current.range.c1 + offset;
      const currentValue = String(sheet.cells[`${totalsRow}:${col}`]?.value ?? "");
      if (!tableTotalCellOwnsValue(current, offset, names[offset], currentValue)) {
        return { ok: false, reason: "totals-row" };
      }
      nextSheet = updateCellValue(nextSheet, totalsRow, col, tableTotalCellValue(nextTable, offset, names[offset]));
    }
  }
  return {
    ok: true,
    sheet: { ...nextSheet, dataTables: tables.map((table) => (table.id === id ? nextTable : table)) },
  };
}
