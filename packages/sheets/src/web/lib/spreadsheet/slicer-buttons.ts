/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slicers as pure sheet transactions: Insert Slicer for a table's columns or
 * a PivotTable's fields, a slicer's design (Slicer tab: caption, style,
 * columns, button height, size, header), its removal, and a PivotTable
 * slicer's selection written as the PivotTable's filter. Moved out of
 * SpreadsheetWindow so the window and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps the dialogs, receipts,
 * notices, the slicer selection context and Undo bookkeeping.
 */

import type { DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import {
  normalizePivotConfig,
  normalizePivotRange,
  sheetPivots,
  type PivotConfig,
  type PivotMatrix,
} from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import {
  cascadeSlicerOrigin,
  MAX_SLICER_TITLE_LENGTH,
  MAX_SLICERS,
  normalizeSlicer,
  normalizeSlicers,
  pivotSlicerFieldValues,
  pivotSlicerSourceKey,
  resolvePivotSlicerField,
  slicerSelectionToFilter,
  type Slicer,
} from "./slicer";

/** The parts of a worksheet the slicer transactions read and write. */
export interface SlicerButtonSheet {
  id: string;
  columnCount: number;
  slicers?: Slicer[];
  pivot?: PivotConfig;
  morePivots?: PivotConfig[];
}

/**
 * The table columns Insert Slicer offers: those the table does not already
 * show in a slicer (a legacy column-bound slicer inside the table counts).
 */
export function tableSlicerColumns(
  sheet: SlicerButtonSheet,
  table: DataTable,
): { slicers: Slicer[]; availableColumnIds: number[] } {
  const liveSlicers = normalizeSlicers(sheet.slicers, sheet.columnCount);
  const usedColumnIds = new Set(
    liveSlicers.flatMap((slicer) => {
      if (slicer.source?.kind === "table" && slicer.source.tableId === table.id) {
        return [slicer.source.tableColumnId];
      }
      if (!slicer.source && slicer.column >= table.range.c1 && slicer.column <= table.range.c2) {
        const legacyColumn = table.columns?.[slicer.column - table.range.c1];
        return legacyColumn ? [legacyColumn.id] : [];
      }
      return [];
    }),
  );
  return {
    slicers: liveSlicers,
    availableColumnIds: (table.columns ?? [])
      .filter((column) => !usedColumnIds.has(column.id))
      .map((column) => column.id),
  };
}

/**
 * Insert Slicer on a table: one slicer per picked column the table does not
 * already show in a slicer, placed to the right of the table level with its
 * top row (`firstSpot`) and cascaded, as Excel places them (Microsoft
 * Support, "Use slicers to filter data"). Null when a column is unknown or
 * taken, nothing is picked, or the sheet would pass its slicer limit.
 */
export function tableSlicersAdded(
  sheet: SlicerButtonSheet,
  table: DataTable,
  selectedColumnIds: readonly number[],
  firstSpot: { x: number; y: number },
  identity: string,
): { slicers: Slicer[]; created: Slicer[] } | null {
  if (!table.columns?.length) return null;
  const { slicers: liveSlicers, availableColumnIds } = tableSlicerColumns(sheet, table);
  if (
    selectedColumnIds.length === 0 ||
    new Set(selectedColumnIds).size !== selectedColumnIds.length ||
    selectedColumnIds.length > MAX_SLICERS - liveSlicers.length ||
    selectedColumnIds.some((columnId) => !availableColumnIds.includes(columnId))
  ) {
    return null;
  }
  const selected = new Set(selectedColumnIds);
  const orderedColumns = table.columns.filter(
    (column) => availableColumnIds.includes(column.id) && selected.has(column.id),
  );
  const created: Slicer[] = [];
  for (const [index, column] of orderedColumns.entries()) {
    const spot = cascadeSlicerOrigin([...liveSlicers, ...created], firstSpot);
    const offset = table.columns.findIndex((candidate) => candidate.id === column.id);
    created.push({
      id: `slicer-${table.range.c1 + offset}-${identity}-${index}`,
      column: table.range.c1 + offset,
      source: { kind: "table", tableId: table.id, tableColumnId: column.id },
      x: spot.x,
      y: spot.y,
    });
  }
  if (created.length !== selectedColumnIds.length) return null;
  return { slicers: [...liveSlicers, ...created], created };
}

/**
 * The fields Insert Slicer offers for a PivotTable: every named PivotCache
 * field, and those not yet shown by a slicer of THIS PivotTable (another
 * PivotTable of the sheet may show a field of the same name in its own).
 */
export function pivotSlicerFields(
  sheet: SlicerButtonSheet,
  pivot: PivotConfig,
  liveMatrix: PivotMatrix | null,
): { all: string[]; available: string[]; pivotName: string | undefined } {
  const all: string[] = [];
  const available: string[] = [];
  const matrix = pivot.native?.snapshot?.matrix ?? pivot.cacheSnapshot?.matrix ?? liveMatrix;
  // A later PivotTable of the sheet is named; the first needs no name.
  const pivotName = pivot === sheet.pivot ? undefined : pivot.name ?? pivot.native?.pivotName ?? "";
  const used = new Set(
    normalizeSlicers(sheet.slicers, sheet.columnCount).flatMap((slicer) =>
      slicer.source?.kind === "pivot" ? [pivotSlicerSourceKey(slicer.source)] : [],
    ),
  );
  const seen = new Set<string>();
  for (const raw of matrix?.[0] ?? []) {
    const name = String(raw ?? "").trim();
    const key = name.toLowerCase();
    if (!name || name.length > MAX_SLICER_TITLE_LENGTH || seen.has(key)) continue;
    seen.add(key);
    all.push(name);
    if (!used.has(pivotSlicerSourceKey({ pivotOutput: sheet.id, pivotName, fieldId: name }))) {
      available.push(name);
    }
  }
  return { all, available, pivotName };
}

/**
 * Insert Slicer on a PivotTable: one slicer per picked field, placed to the
 * right of the report level with its top row (`firstSpot`) and cascaded, as
 * a table's slicers are; without a spot the grid places it. The cache column
 * keeps the legacy `column` slot meaningful. Null when a field is unknown or
 * taken, nothing is picked, or the sheet would pass its slicer limit.
 */
export function pivotSlicersAdded(
  sheet: SlicerButtonSheet,
  fields: { all: readonly string[]; available: readonly string[]; pivotName: string | undefined },
  selectedFields: readonly string[],
  identity: string,
  firstSpot?: { x: number; y: number },
): { slicers: Slicer[]; created: Slicer[] } | null {
  const liveSlicers = normalizeSlicers(sheet.slicers, sheet.columnCount);
  const selected = new Set(selectedFields);
  if (
    selected.size === 0 ||
    selected.size !== selectedFields.length ||
    selected.size > MAX_SLICERS - liveSlicers.length ||
    selectedFields.some((field) => !fields.available.includes(field))
  ) {
    return null;
  }
  const lastColumn = Math.max(0, sheet.columnCount - 1);
  const created: Slicer[] = [];
  for (const [index, field] of fields.all.filter((candidate) => selected.has(candidate)).entries()) {
    const spot = firstSpot ? cascadeSlicerOrigin([...liveSlicers, ...created], firstSpot) : null;
    created.push({
      id: `slicer-pivot-${identity}-${index}`,
      column: Math.min(fields.all.indexOf(field), lastColumn),
      source: {
        kind: "pivot",
        pivotOutput: sheet.id,
        fieldId: field,
        ...(fields.pivotName !== undefined ? { pivotName: fields.pivotName } : {}),
      },
      ...(spot ? { x: spot.x, y: spot.y } : {}),
    });
  }
  if (created.length !== selected.size) return null;
  return { slicers: [...liveSlicers, ...created], created };
}

/** A slicer's design change (Slicer tab), normalized; null when it changes
 *  nothing or is not a valid slicer. */
export function slicersWithDesign(
  sheet: SlicerButtonSheet,
  slicerId: string,
  patch: Partial<Slicer>,
): Slicer[] | null {
  const current = normalizeSlicers(sheet.slicers, sheet.columnCount);
  const index = current.findIndex((slicer) => slicer.id === slicerId);
  if (index < 0) return null;
  const nextSlicer = normalizeSlicer(
    { ...current[index], ...patch, id: current[index].id, column: current[index].column },
    sheet.columnCount,
  );
  if (!nextSlicer || JSON.stringify(nextSlicer) === JSON.stringify(current[index])) return null;
  const next = [...current];
  next[index] = nextSlicer;
  return next;
}

/** The sheet without one slicer. The filter it drove is NOT touched (Excel
 *  keeps the filter when a slicer is deleted); the list is dropped when none
 *  remain. Null when the sheet has no such slicer. */
export function sheetWithoutSlicer<S extends SlicerButtonSheet>(sheet: S, slicerId: string): S | null {
  const slicers = normalizeSlicers(sheet.slicers, sheet.columnCount);
  const next = slicers.filter((slicer) => slicer.id !== slicerId);
  if (next.length === slicers.length) return null;
  const nextSheet: S = { ...sheet };
  if (next.length > 0) nextSheet.slicers = next;
  else delete nextSheet.slicers;
  return nextSheet;
}

/** One PivotTable a slicer filters: its host, the PivotTable, field and values. */
export interface PivotSlicerTarget<S> {
  hostSheet: S;
  pivot: PivotConfig;
  colIndex: number;
  values: string[];
  filterKind: "report" | "item";
  filterValues: readonly string[] | undefined;
}

/** What a PivotTable slicer filters: its first PivotTable (which gives the
 *  items and the selection shown) and every PivotTable it is connected to. */
export type PivotSlicerBinding<S> = PivotSlicerTarget<S> & { targets: PivotSlicerTarget<S>[] };

/**
 * The PivotTable fields a slicer is bound to. A PivotTable slicer lists the
 * PivotCache items of its field (Excel's slicer items come from the cache,
 * not the live source). A slicer may filter several PivotTables on one cache
 * (Excel's Slicer > Report Connections): a connected PivotTable that is gone,
 * or no longer has the field, leaves the slicer unusable rather than
 * filtering only some.
 */
export function pivotSlicerBinding<S extends SlicerButtonSheet>(
  sheets: readonly S[],
  slicer: Slicer,
  liveMatrixFor: (host: S, pivot: PivotConfig) => PivotMatrix | null,
): PivotSlicerBinding<S> | null {
  const source = slicer.source;
  if (source?.kind !== "pivot") return null;
  const bindOne = (hostId: string, pivotName: string | undefined): PivotSlicerTarget<S> | null => {
    const hostSheet = sheets.find((candidate) => candidate.id === hostId);
    // The named PivotTable of the host sheet, else its first. A renamed
    // imported PivotTable still answers to the name its file gave it (the
    // slicer cache names it so until the save renames it there).
    const wanted = pivotName?.toLowerCase();
    const pivot = hostSheet
      ? wanted
        ? sheetPivots(hostSheet).find((candidate) =>
            [candidate.name, candidate.native?.pivotName].some((name) => name?.toLowerCase() === wanted),
          )
        : hostSheet.pivot
      : undefined;
    if (!hostSheet || !pivot) return null;
    const matrix = pivot.native?.snapshot?.matrix ?? pivot.cacheSnapshot?.matrix ?? liveMatrixFor(hostSheet, pivot);
    if (!matrix || matrix.length === 0) return null;
    const colIndex = resolvePivotSlicerField(matrix[0], source.fieldId);
    if (colIndex === null) return null;
    const reportFilter = (pivot.filters ?? []).find((entry) => entry.colIndex === colIndex);
    // A slicer lists the field's own values; a filter that names a level of a
    // grouped date field (Years, Quarters) or a manual grouping's group field
    // belongs to that caption.
    const itemFilter = pivot.itemFilters?.find(
      (entry) => entry.colIndex === colIndex && entry.unit === undefined && entry.group === undefined,
    );
    return {
      hostSheet,
      pivot,
      colIndex,
      values: pivotSlicerFieldValues(matrix, colIndex),
      filterKind: reportFilter ? "report" : "item",
      filterValues: reportFilter ? reportFilter.selected : itemFilter?.selected,
    };
  };
  const first = bindOne(source.pivotOutput, source.pivotName);
  if (!first) return null;
  const targets = [first];
  for (const connection of source.connections ?? []) {
    const target = bindOne(connection.pivotOutput, connection.pivotName);
    if (!target || targets.some((known) => known.pivot === target.pivot)) return null;
    targets.push(target);
  }
  return { ...first, targets };
}

/** A slicer selection as the filter it writes: all items selected (or
 *  none) means no filter, as Excel records it. */
export function pivotSlicerVisibleItems(
  binding: Pick<PivotSlicerTarget<unknown>, "values">,
  selection: readonly string[] | null,
): string[] | null {
  return selection === null ? null : slicerSelectionToFilter(selection, binding.values);
}

/**
 * One connected PivotTable with a slicer's filter: the report filter's items
 * when the field is in the Filters area, else the field's manual item filter.
 * Null when nothing changes.
 */
export function pivotWithSlicerFilter(
  target: PivotSlicerTarget<{ columnCount: number }>,
  visible: readonly string[] | null,
): PivotConfig | null {
  const pivot = target.pivot;
  const nextPivot: PivotConfig = { ...pivot };
  if (target.filterKind === "report") {
    nextPivot.filters = (pivot.filters ?? []).map((entry) =>
      entry.colIndex === target.colIndex ? { colIndex: entry.colIndex, selected: visible ? [...visible] : [] } : entry,
    );
  } else {
    const others = (pivot.itemFilters ?? []).filter(
      (entry) => entry.colIndex !== target.colIndex || entry.unit !== undefined || entry.group !== undefined,
    );
    const itemFilters = visible ? [...others, { colIndex: target.colIndex, selected: [...visible] }] : others;
    if (itemFilters.length > 0) nextPivot.itemFilters = itemFilters;
    else delete nextPivot.itemFilters;
  }
  const range = normalizePivotRange(nextPivot.source);
  const maxCol = range ? range.c2 - range.c1 : target.hostSheet.columnCount - 1;
  const normalized = normalizePivotConfig(nextPivot, maxCol);
  return JSON.stringify(normalized) === JSON.stringify(pivot) ? null : normalized;
}
