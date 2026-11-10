/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import * as Y from "yjs";
import { applyJsonDelta, createJsonCollaborationHistory, jsonEquals, yToJson, type JsonValue } from "@mrnote/core/web/lib/collab/json-crdt.ts";
import { cellMapChangesSince } from "@mrnote/core/web/lib/spreadsheet/cell-map-lineage.ts";

/**
 * Whether two cell maps hold the same cells, read at the keys they differ at when
 * one is recorded as following from the other: Undo weighed three whole workbooks,
 * key by key, however small the step.
 */
function cellsEqual(a: Record<string, JsonValue>, b: Record<string, JsonValue>): boolean {
  if (a === b) return true;
  const keys = cellMapChangesSince(a, b) ?? cellMapChangesSince(b, a);
  if (!keys) return jsonEquals(a, b);
  return keys.every((key) => jsonEquals(a[key] ?? null, b[key] ?? null));
}

function sheetsEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  const isSheet = (value: unknown): value is { cells: Record<string, JsonValue> } =>
    Boolean(value) && typeof value === "object" && Boolean((value as { cells?: unknown }).cells) && typeof (value as { cells?: unknown }).cells === "object";
  if (!isSheet(a) || !isSheet(b)) return jsonEquals(a as JsonValue, b as JsonValue);
  const { cells: cellsA, ...restA } = a;
  const { cells: cellsB, ...restB } = b;
  return jsonEquals(restA as unknown as JsonValue, restB as unknown as JsonValue) && cellsEqual(cellsA, cellsB);
}

function statesEqual<T>(a: SheetOwnHistoryState<T>, b: SheetOwnHistoryState<T>): boolean {
  const { sheets: sheetsA, ...restA } = a;
  const { sheets: sheetsB, ...restB } = b;
  return sheetsA.length === sheetsB.length && jsonEquals(restA as JsonValue, restB as JsonValue) &&
    sheetsA.every((sheet, index) => sheetsEqual(sheet, sheetsB[index]));
}

export interface SheetOwnHistoryState<T> {
  sheets: T[];
  workbookTheme?: unknown;
  dataSourceInventory?: unknown;
  workbookStructureProtection?: unknown;
}

/** Ordinary history owns the recorded before→after edit, never a whole saved
 * workbook. Native-copy/query history keep their separate source contracts. */
export function invertSheetOwnHistory<T extends { id: string }>(
  before: SheetOwnHistoryState<T>, after: SheetOwnHistoryState<T>,
  current: SheetOwnHistoryState<T>, coordinateSheetId?: string,
): SheetOwnHistoryState<T> | null {
  for (const state of [before, after, current]) {
    const ids = new Set<string>();
    for (const sheet of state.sheets) {
      if (typeof sheet.id !== "string" || !sheet.id || ids.has(sheet.id)) return null;
      ids.add(sheet.id);
    }
  }
  // A step that changed nothing (Delete on blank columns past the data, a
  // Fill that wrote the values already there) is undone by changing nothing,
  // as Excel's Undo consumes it; it was read as a conflict with later
  // changes, which said current content was kept (sheet model fuzz).
  if (statesEqual(before, after)) return current;
  // Coordinate edits change what row/column addresses mean. Until those
  // operations have a semantic rebase, stop if this sheet advanced at all.
  if (coordinateSheetId && !sheetsEqual(
    current.sheets.find(sheet => sheet.id === coordinateSheetId) ?? null,
    after.sheets.find(sheet => sheet.id === coordinateSheetId) ?? null,
  )) return null;
  // Nothing has changed since the step (the common Undo right after an edit):
  // reversing it restores exactly the state it replaced. The records are the
  // ones the step replaced, so every worksheet and cell map it did not touch
  // keeps its identity, and the cost is the size of the step rather than of
  // the workbook (the rebase below copies the whole workbook through a CRDT
  // three times). `jsonEquals` stops at shared references.
  if (statesEqual(current, after)) {
    return statesEqual(before, current) ? null : before;
  }
  const doc = new Y.Doc();
  const root = doc.getMap<unknown>("sheet-own-history");
  const origin = Symbol("ordinary-sheet-edit");
  const asJson = (value: SheetOwnHistoryState<T>) => value as unknown as Record<string, JsonValue>;
  let history: ReturnType<typeof createJsonCollaborationHistory> | null = null;
  try {
    doc.transact(() => applyJsonDelta(root, asJson(before), {}));
    history = createJsonCollaborationHistory(root, { localOrigin: origin, canWrite: () => true });
    doc.transact(() => applyJsonDelta(root, asJson(after), asJson(before)), origin);
    history.stopCapturing();
    doc.transact(() => applyJsonDelta(root, asJson(current), asJson(after)), "later-peer-state");
    if (!history.undo()) return null;
    const result = yToJson(root) as unknown as SheetOwnHistoryState<T>;
    return jsonEquals(result as unknown as JsonValue, current as unknown as JsonValue) ? null : result;
  } catch {
    // An unsupported or ambiguous snapshot cannot become a history writer.
    return null;
  } finally {
    history?.destroy();
    doc.destroy();
  }
}
