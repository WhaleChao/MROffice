/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { Editor } from "@tiptap/core";
import { closeHistory } from "@tiptap/pm/history";
import type { Node as ProseMirrorNode } from "@tiptap/pm/model";
import { CellSelection, TableMap } from "@tiptap/pm/tables";

export type RepeatHeaderRowsUnavailableReason =
  | "not_table"
  | "first_row_required"
  | "whole_rows_required"
  | "readonly";

export interface RepeatHeaderRowsCommandState {
  readonly enabled: boolean;
  readonly active: boolean | "mixed";
  readonly selectedRowCount: number;
  readonly unavailableReason: RepeatHeaderRowsUnavailableReason | null;
}

export type RowBreakAcrossPagesUnavailableReason =
  | "not_table"
  | "row_selection_required"
  | "readonly";

export interface RowBreakAcrossPagesCommandState {
  /** Checked means Word's “Allow row to break across pages” option is on. */
  readonly active: boolean | "mixed";
  readonly enabled: boolean;
  readonly selectedRowCount: number;
  readonly unavailableReason: RowBreakAcrossPagesUnavailableReason | null;
}

interface TableSelectionTarget {
  readonly tableNode: ProseMirrorNode;
  readonly tablePos: number;
  readonly selectedRowIndexes: readonly number[];
  readonly selectedRowCount: number;
}

function tableAtSelection(
  editor: Editor,
): { tableNode: ProseMirrorNode; tablePos: number } | null {
  const { doc, selection } = editor.state;
  for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
    const node = selection.$from.node(depth);
    if (node.type.name === "table") {
      return { tableNode: node, tablePos: selection.$from.before(depth) };
    }
  }
  const selectedNode = doc.nodeAt(selection.from);
  return selectedNode?.type.name === "table"
    ? { tableNode: selectedNode, tablePos: selection.from }
    : null;
}

function selectedRows(
  editor: Editor,
  tableNode: ProseMirrorNode,
  tablePos: number,
): { rows: ReadonlySet<number>; wholeRows: boolean } {
  const { selection } = editor.state;
  let map: TableMap;
  try {
    map = TableMap.get(tableNode);
  } catch {
    return { rows: new Set(), wholeRows: false };
  }

  if (selection instanceof CellSelection) {
    const selectedCellPositions = new Set<number>();
    selection.forEachCell((_node, pos) => selectedCellPositions.add(pos));
    const rows = new Set<number>();
    map.map.forEach((offset, index) => {
      if (selectedCellPositions.has(tablePos + 1 + offset)) {
        rows.add(Math.floor(index / map.width));
      }
    });
    return { rows, wholeRows: selection.isRowSelection() };
  }

  let cellPos: number | null = null;
  for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
    const node = selection.$from.node(depth);
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      cellPos = selection.$from.before(depth);
      break;
    }
  }
  if (cellPos === null) return { rows: new Set(), wholeRows: false };
  const cellOffset = cellPos - tablePos - 1;
  const mapIndex = map.map.findIndex((offset) => offset === cellOffset);
  return mapIndex >= 0
    ? { rows: new Set([Math.floor(mapIndex / map.width)]), wholeRows: true }
    : { rows: new Set(), wholeRows: false };
}

function targetAtSelection(
  editor: Editor,
):
  | { target: TableSelectionTarget }
  | { unavailableReason: Exclude<RepeatHeaderRowsUnavailableReason, "readonly"> } {
  const table = tableAtSelection(editor);
  if (!table) return { unavailableReason: "not_table" };
  const selected = selectedRows(editor, table.tableNode, table.tablePos);
  if (!selected.wholeRows || selected.rows.size === 0) {
    return { unavailableReason: "whole_rows_required" };
  }
  const orderedRows = [...selected.rows].sort((left, right) => left - right);
  if (orderedRows[0] !== 0) {
    return { unavailableReason: "first_row_required" };
  }
  for (let index = 0; index < orderedRows.length; index += 1) {
    if (orderedRows[index] !== index) {
      return { unavailableReason: "whole_rows_required" };
    }
  }
  return {
    target: {
      ...table,
      selectedRowIndexes: orderedRows,
      selectedRowCount: orderedRows.length,
    },
  };
}

function rowPropertiesTargetAtSelection(
  editor: Editor,
):
  | { target: TableSelectionTarget }
  | {
      unavailableReason: Exclude<
        RowBreakAcrossPagesUnavailableReason,
        "readonly"
      >;
    } {
  const table = tableAtSelection(editor);
  if (!table) return { unavailableReason: "not_table" };
  const wholeTableSelected =
    editor.state.selection.from === table.tablePos &&
    editor.state.doc.nodeAt(editor.state.selection.from) === table.tableNode;
  const selectedRowIndexes = wholeTableSelected
    ? Array.from({ length: table.tableNode.childCount }, (_, index) => index)
    : [...selectedRows(editor, table.tableNode, table.tablePos).rows].sort(
        (left, right) => left - right,
      );
  if (selectedRowIndexes.length === 0) {
    return { unavailableReason: "row_selection_required" };
  }
  return {
    target: {
      ...table,
      selectedRowIndexes,
      selectedRowCount: selectedRowIndexes.length,
    },
  };
}

/**
 * Resolve the live Word-style Repeat Header Rows command state.
 *
 * A caret in row zero targets the first row. A multi-row operation must be a
 * full-row CellSelection that starts at row zero; Word does not allow a later
 * row to become an independently repeating header.
 */
export function repeatHeaderRowsCommandState(
  editor: Editor,
): RepeatHeaderRowsCommandState {
  if (editor.isDestroyed) {
    return {
      enabled: false,
      active: false,
      selectedRowCount: 0,
      unavailableReason: "not_table",
    };
  }
  const resolved = targetAtSelection(editor);
  if ("unavailableReason" in resolved) {
    return {
      enabled: false,
      active: false,
      selectedRowCount: 0,
      unavailableReason: resolved.unavailableReason,
    };
  }
  const { tableNode, selectedRowCount } = resolved.target;
  let repeatedCount = 0;
  for (let rowIndex = 0; rowIndex < selectedRowCount; rowIndex += 1) {
    if (tableNode.child(rowIndex).attrs.headerRow === true) repeatedCount += 1;
  }
  const active =
    repeatedCount === 0
      ? false
      : repeatedCount === selectedRowCount
        ? true
        : "mixed";
  return {
    enabled: editor.isEditable,
    active,
    selectedRowCount,
    unavailableReason: editor.isEditable ? null : "readonly",
  };
}

/**
 * Toggle the selected leading rows as the table's contiguous repeating header.
 * This changes only the explicit row `headerRow` model bit; it never converts
 * cells to/from the independent `tableHeader` visual/semantic cell type.
 */
export function setRepeatHeaderRows(editor: Editor, enabled: boolean): boolean {
  if (editor.isDestroyed || !editor.isEditable) return false;
  const resolved = targetAtSelection(editor);
  if ("unavailableReason" in resolved) return false;

  const { tableNode, tablePos, selectedRowCount } = resolved.target;
  const transaction = editor.state.tr;
  let changed = false;
  tableNode.forEach((rowNode, rowOffset, rowIndex) => {
    const nextHeaderRow = enabled && rowIndex < selectedRowCount;
    if ((rowNode.attrs.headerRow === true) === nextHeaderRow) return;
    transaction.setNodeMarkup(
      tablePos + 1 + rowOffset,
      undefined,
      { ...rowNode.attrs, headerRow: nextHeaderRow },
      rowNode.marks,
    );
    changed = true;
  });
  if (!changed) return false;
  editor.view.dispatch(closeHistory(transaction).scrollIntoView());
  return true;
}

export function toggleRepeatHeaderRows(editor: Editor): boolean {
  const state = repeatHeaderRowsCommandState(editor);
  return state.enabled && setRepeatHeaderRows(editor, state.active !== true);
}

/** Resolve Word's row-level “Allow row to break across pages” checkbox. */
export function rowBreakAcrossPagesCommandState(
  editor: Editor,
): RowBreakAcrossPagesCommandState {
  if (editor.isDestroyed) {
    return {
      enabled: false,
      active: false,
      selectedRowCount: 0,
      unavailableReason: "not_table",
    };
  }
  const resolved = rowPropertiesTargetAtSelection(editor);
  if ("unavailableReason" in resolved) {
    return {
      enabled: false,
      active: false,
      selectedRowCount: 0,
      unavailableReason: resolved.unavailableReason,
    };
  }
  const { tableNode, selectedRowIndexes } = resolved.target;
  const breakableRows = selectedRowIndexes.filter(
    (rowIndex) => tableNode.child(rowIndex).attrs.cantSplit !== true,
  ).length;
  const active =
    breakableRows === 0
      ? false
      : breakableRows === selectedRowIndexes.length
        ? true
        : "mixed";
  return {
    enabled: editor.isEditable,
    active,
    selectedRowCount: selectedRowIndexes.length,
    unavailableReason: editor.isEditable ? null : "readonly",
  };
}

/**
 * Toggle Word's positive “Allow row to break across pages” option for every
 * row touched by the live table selection. The persisted model is the inverse
 * OOXML property: unchecked maps to `cantSplit=true`; checked maps to false.
 */
export function toggleRowBreakAcrossPages(editor: Editor): boolean {
  if (editor.isDestroyed || !editor.isEditable) return false;
  const resolved = rowPropertiesTargetAtSelection(editor);
  if ("unavailableReason" in resolved) return false;

  const { tableNode, tablePos, selectedRowIndexes } = resolved.target;
  const selectedRowsSet = new Set(selectedRowIndexes);
  const state = rowBreakAcrossPagesCommandState(editor);
  const nextCantSplit = state.active === true;
  const transaction = editor.state.tr;
  let changed = false;
  tableNode.forEach((rowNode, rowOffset, rowIndex) => {
    if (!selectedRowsSet.has(rowIndex)) return;
    if ((rowNode.attrs.cantSplit === true) === nextCantSplit) return;
    transaction.setNodeMarkup(
      tablePos + 1 + rowOffset,
      undefined,
      { ...rowNode.attrs, cantSplit: nextCantSplit },
      rowNode.marks,
    );
    changed = true;
  });
  if (!changed) return false;
  editor.view.dispatch(transaction.scrollIntoView());
  return true;
}
