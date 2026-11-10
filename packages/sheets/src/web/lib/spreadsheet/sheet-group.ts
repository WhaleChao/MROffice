/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Grouped worksheets (XLS-A079) — Microsoft Support, "Group worksheets" and
// "Select worksheets".
//
// Several selected sheet tabs form a group: Ctrl+click adds or removes a tab,
// Shift+click selects the run of tabs from the active one, and the tab menu's
// Select All Sheets groups every sheet. While the sheets are grouped, "any
// changes you make on one worksheet are made in all the other worksheets in
// the group in the exact same location" — data, formulas and formatting.
// Clicking a tab outside the group, or Ungroup Sheets, ends the group; with
// ALL sheets grouped, browsing to another sheet ends it as well.
//
// PURE: tab-click rules and the replay of one committed edit of the active
// sheet onto the other grouped sheets.
// ---------------------------------------------------------------------------

import {
  deletedLinesSheet,
  insertedLinesSheet,
} from "./row-column-buttons";
import type { StructuralSheet } from "@mrnote/core/web/lib/spreadsheet/structural.ts";

interface GroupTab {
  readonly id: string;
  readonly hidden?: boolean;
}

/** The grouped sheet ids in workbook order, or null when fewer than two. */
export function normalizeSheetGroup(
  sheets: readonly GroupTab[],
  ids: Iterable<string> | null,
  activeId: string,
): string[] | null {
  if (!ids) return null;
  const wanted = new Set(ids);
  wanted.add(activeId);
  const ordered = sheets
    .filter((sheet) => wanted.has(sheet.id) && !sheet.hidden)
    .map((sheet) => sheet.id);
  return ordered.length >= 2 && ordered.includes(activeId) ? ordered : null;
}

/** Whether a group holds every visible sheet (Select All Sheets). */
export function groupHoldsEverySheet(
  sheets: readonly GroupTab[],
  group: readonly string[] | null,
): boolean {
  if (!group) return false;
  const visible = sheets.filter((sheet) => !sheet.hidden);
  return visible.length >= 2 && visible.every((sheet) => group.includes(sheet.id));
}

/**
 * What a click on a sheet tab does to the group and which sheet shows.
 * - Ctrl/Cmd: an ungrouped tab joins the group and shows; a grouped tab other
 *   than the shown one leaves it.
 * - Shift: the tabs from the shown one to the clicked one form the group; the
 *   shown sheet stays.
 * - Plain: inside a partial group the clicked grouped sheet shows and the
 *   group stays; a tab outside the group, or any tab while ALL sheets are
 *   grouped, ends the group.
 */
export function sheetGroupAfterTabClick(
  sheets: readonly GroupTab[],
  group: readonly string[] | null,
  activeIndex: number,
  clickedIndex: number,
  modifiers: { readonly toggle: boolean; readonly range: boolean },
): { group: string[] | null; activeIndex: number } {
  const active = sheets[activeIndex];
  const clicked = sheets[clickedIndex];
  if (!active || !clicked || clicked.hidden) return { group: group ? [...group] : null, activeIndex };
  const current = group ?? [active.id];
  if (modifiers.range) {
    const [from, to] = activeIndex <= clickedIndex ? [activeIndex, clickedIndex] : [clickedIndex, activeIndex];
    const ids = sheets.slice(from, to + 1).filter((sheet) => !sheet.hidden).map((sheet) => sheet.id);
    return { group: normalizeSheetGroup(sheets, ids, active.id), activeIndex };
  }
  if (modifiers.toggle) {
    if (!current.includes(clicked.id)) {
      return {
        group: normalizeSheetGroup(sheets, [...current, clicked.id], clicked.id),
        activeIndex: clickedIndex,
      };
    }
    if (clicked.id === active.id) return { group: group ? [...group] : null, activeIndex };
    return {
      group: normalizeSheetGroup(sheets, current.filter((id) => id !== clicked.id), active.id),
      activeIndex,
    };
  }
  if (group && group.includes(clicked.id) && !groupHoldsEverySheet(sheets, group)) {
    return { group: [...group], activeIndex: clickedIndex };
  }
  return { group: null, activeIndex: clickedIndex };
}

// ---------------------------------------------------------------------------
// Replaying an edit
// ---------------------------------------------------------------------------

type Cell = Record<string, unknown>;

/** Derived caches a write invalidates; recomputed by the serializer. */
const CACHE_FIELDS = ["display", "chartValueCache"] as const;

function sameJson(a: unknown, b: unknown): boolean {
  return a === b || JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/** The fields one cell edit changed: style keys separately, so a format edit
 *  never carries the edited sheet's other formats or its value along. */
interface CellChange {
  readonly removed: boolean;
  readonly fields: Record<string, unknown>;
  readonly cleared: string[];
  readonly style?: { set: Record<string, unknown>; cleared: string[] };
}

function cellChange(before: Cell | undefined, after: Cell | undefined): CellChange | null {
  if (before === after) return null;
  if (!after) return before ? { removed: true, fields: {}, cleared: [] } : null;
  const fields: Record<string, unknown> = {};
  const cleared: string[] = [];
  let style: CellChange["style"];
  const keys = new Set([...Object.keys(before ?? {}), ...Object.keys(after)]);
  for (const key of keys) {
    if ((CACHE_FIELDS as readonly string[]).includes(key)) continue;
    const b = before?.[key];
    const a = after[key];
    if (sameJson(b, a)) continue;
    if (key === "style" && a && typeof a === "object" && (!b || typeof b === "object")) {
      const set: Record<string, unknown> = {};
      const styleCleared: string[] = [];
      const bs = (b ?? {}) as Record<string, unknown>;
      const as = a as Record<string, unknown>;
      for (const styleKey of new Set([...Object.keys(bs), ...Object.keys(as)])) {
        if (sameJson(bs[styleKey], as[styleKey])) continue;
        if (as[styleKey] === undefined) styleCleared.push(styleKey);
        else set[styleKey] = as[styleKey];
      }
      style = { set, cleared: styleCleared };
      continue;
    }
    if (a === undefined) cleared.push(key);
    else fields[key] = a;
  }
  if (!style && cleared.length === 0 && Object.keys(fields).length === 0) return null;
  return { removed: false, fields, cleared, style };
}

function applyCellChange(target: Cell | undefined, change: CellChange): Cell | undefined {
  if (change.removed) return undefined;
  const next: Cell = { ...(target ?? { value: "" }) };
  for (const [key, value] of Object.entries(change.fields)) next[key] = value;
  for (const key of change.cleared) delete next[key];
  if (change.style) {
    const style = { ...((next.style as Record<string, unknown> | undefined) ?? {}) };
    for (const [key, value] of Object.entries(change.style.set)) style[key] = value;
    for (const key of change.style.cleared) delete style[key];
    if (Object.keys(style).length > 0) next.style = style;
    else delete next.style;
  }
  const valueChanged =
    "value" in change.fields || "formula" in change.fields ||
    change.cleared.includes("value") || change.cleared.includes("formula");
  if (valueChanged) for (const cache of CACHE_FIELDS) delete next[cache];
  if (typeof next.value !== "string") next.value = "";
  const meaningful = Object.keys(next).some((key) => key !== "value") || next.value !== "";
  return meaningful ? next : undefined;
}

function sizeMapChange(
  before: Record<string, unknown> | undefined,
  after: Record<string, unknown> | undefined,
): { set: Record<string, unknown>; cleared: string[] } | null {
  if (before === after) return null;
  const set: Record<string, unknown> = {};
  const cleared: string[] = [];
  for (const key of new Set([...Object.keys(before ?? {}), ...Object.keys(after ?? {})])) {
    if (sameJson(before?.[key], after?.[key])) continue;
    if (after?.[key] === undefined) cleared.push(key);
    else set[key] = after[key];
  }
  return cleared.length === 0 && Object.keys(set).length === 0 ? null : { set, cleared };
}

function linesChange(before: readonly number[] | undefined, after: readonly number[] | undefined) {
  if (before === after) return null;
  const b = new Set(before ?? []);
  const a = new Set(after ?? []);
  const added = [...a].filter((index) => !b.has(index));
  const removed = [...b].filter((index) => !a.has(index));
  return added.length === 0 && removed.length === 0 ? null : { added, removed };
}

type Merge = { row: number; col: number; rowSpan: number; colSpan: number };
const mergeKey = (span: Merge) => `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`;
const overlaps = (a: Merge, b: Merge) =>
  a.row < b.row + b.rowSpan && b.row < a.row + a.rowSpan && a.col < b.col + b.colSpan && b.col < a.col + a.colSpan;

/** A structural line edit (Insert/Delete Sheet Rows/Columns). */
export interface GroupLineEdit {
  readonly axis: "row" | "col";
  readonly op: "insert" | "delete";
  readonly at: number;
  readonly count: number;
}

export type GroupReplayResult<S> =
  | { ok: true; workbook: S[]; changedSheetIds: string[] }
  | { ok: false; reason: "locked" | "three-d" | "insert-refused" | "unsupported" };

/** The fields the replay reads and writes, over any sheet model. */
interface LooseSheet {
  id: string;
  rowCount: number;
  columnCount: number;
  cells: Record<string, Cell>;
  rowHeights?: Record<string, number>;
  columnWidths?: Record<string, number>;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  manuallyHiddenRows?: number[];
  manuallyHiddenCols?: number[];
  rowStyles?: Record<string, unknown>;
  columnStyles?: Record<string, unknown>;
  merges?: Merge[];
}

/**
 * Apply the edit that turned `before` into `after` (the active sheet) to every
 * other sheet in `groupIds`, at the same cells and lines. `workbook` is the
 * workbook already holding `after`. A row/column insert or delete is replayed
 * as the same structural edit (references on the other sheets follow it);
 * other coordinate changes (sort, cell shifts, moves) are not, and refuse.
 * `isLocked(sheet,row,col)` refuses a cell write a protected grouped sheet
 * forbids. Drawings, filters and frozen panes stay with the active sheet.
 */
export function replayEditOnGroup<S extends StructuralSheet & { id: string }>(params: {
  workbook: readonly S[];
  activeIndex: number;
  groupIds: readonly string[];
  before: S;
  after: S;
  coordinate: boolean;
  lineEdit: GroupLineEdit | null;
  isLocked?: (sheet: S, row: number, col: number) => boolean;
}): GroupReplayResult<S> {
  const { activeIndex, groupIds, coordinate, lineEdit, isLocked } = params;
  const before = params.before as unknown as LooseSheet;
  const after = params.after as unknown as LooseSheet;
  let workbook = [...params.workbook];
  const targets = workbook
    .map((sheet, index) => ({ sheet, index }))
    // A chart sheet in the group has no cells or lines to take the edit.
    .filter(({ sheet, index }) =>
      index !== activeIndex && groupIds.includes(sheet.id) &&
      !(sheet as { chartSheet?: unknown }).chartSheet);
  if (targets.length === 0) return { ok: true, workbook, changedSheetIds: [] };
  const changed: string[] = [];

  if (coordinate) {
    if (!lineEdit) return { ok: false, reason: "unsupported" };
    for (const { index } of targets) {
      const target = workbook[index];
      if (lineEdit.op === "insert") {
        const result = insertedLinesSheet(target, workbook, lineEdit.axis, lineEdit.at, lineEdit.count);
        if (result === "three-d") return { ok: false, reason: "three-d" };
        if (typeof result === "string") return { ok: false, reason: "insert-refused" };
        workbook = workbook.map((sheet, i) =>
          i === index ? result.sheet : result.followOthers ? result.followOthers(sheet) : sheet,
        );
      } else {
        const result = deletedLinesSheet(target, workbook, lineEdit.axis, lineEdit.at, lineEdit.count);
        if (result === "three-d") return { ok: false, reason: "three-d" };
        // A delete through a legacy array formula's range on a grouped sheet
        // is refused like an insert that would split one.
        if (typeof result === "string") return { ok: false, reason: "insert-refused" };
        if (!result) continue;
        workbook = workbook.map((sheet, i) =>
          i === index ? result.sheet : result.followOthers ? result.followOthers(sheet) : sheet,
        );
      }
      changed.push(target.id);
    }
    return { ok: true, workbook, changedSheetIds: changed };
  }

  // Content: cells field by field, line sizes / hidden lines / line formats,
  // and merges.
  const cellChanges: Array<[string, CellChange]> = [];
  if (before.cells !== after.cells) {
    for (const key of new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])) {
      const change = cellChange(before.cells[key], after.cells[key]);
      if (change) cellChanges.push([key, change]);
    }
  }
  const mapChanges = (["rowHeights", "columnWidths", "rowStyles", "columnStyles"] as const)
    .map((field) => [field, sizeMapChange(before[field], after[field])] as const)
    .filter(([, change]) => change !== null);
  const hiddenChanges = (["manuallyHiddenRows", "manuallyHiddenCols"] as const)
    .map((field) => [field, linesChange(before[field], after[field])] as const)
    .filter(([, change]) => change !== null);
  const defaultChanges = (["defaultRowHeight", "defaultColumnWidth"] as const).filter(
    (field) => before[field] !== after[field],
  );
  const beforeMerges = new Map((before.merges ?? []).map((span) => [mergeKey(span), span]));
  const afterMerges = new Map((after.merges ?? []).map((span) => [mergeKey(span), span]));
  const addedMerges = [...afterMerges].filter(([key]) => !beforeMerges.has(key)).map(([, span]) => span);
  const removedMerges = [...beforeMerges].filter(([key]) => !afterMerges.has(key)).map(([, span]) => span);
  if (
    cellChanges.length === 0 && mapChanges.length === 0 && hiddenChanges.length === 0 &&
    defaultChanges.length === 0 && addedMerges.length === 0 && removedMerges.length === 0
  ) {
    return { ok: true, workbook, changedSheetIds: [] };
  }

  for (const { index } of targets) {
    const target = workbook[index] as unknown as LooseSheet;
    const next: LooseSheet = { ...target };
    if (cellChanges.length > 0) {
      const cells = { ...target.cells };
      for (const [key, change] of cellChanges) {
        const [row, col] = key.split(":").map(Number);
        if (isLocked?.(workbook[index], row, col)) return { ok: false, reason: "locked" };
        const updated = applyCellChange(cells[key], change);
        if (updated) cells[key] = updated;
        else delete cells[key];
        next.rowCount = Math.max(next.rowCount, row + 1);
        next.columnCount = Math.max(next.columnCount, col + 1);
      }
      next.cells = cells;
    }
    for (const [field, change] of mapChanges) {
      const map: Record<string, unknown> = { ...((next[field] as Record<string, unknown> | undefined) ?? {}) };
      for (const [key, value] of Object.entries(change!.set)) map[key] = value;
      for (const key of change!.cleared) delete map[key];
      if (Object.keys(map).length > 0) next[field] = map as Record<string, number>;
      else delete next[field];
    }
    for (const [field, change] of hiddenChanges) {
      const set = new Set(next[field] ?? []);
      for (const line of change!.added) set.add(line);
      for (const line of change!.removed) set.delete(line);
      if (set.size > 0) next[field] = [...set].sort((a, b) => a - b);
      else delete next[field];
    }
    for (const field of defaultChanges) {
      if (after[field] === undefined) delete next[field];
      else next[field] = after[field];
    }
    if (addedMerges.length > 0 || removedMerges.length > 0) {
      const removedKeys = new Set(removedMerges.map(mergeKey));
      const kept = (next.merges ?? []).filter(
        (span) => !removedKeys.has(mergeKey(span)) && !addedMerges.some((added) => overlaps(added, span)),
      );
      const merges = [...kept, ...addedMerges.map((span) => ({ ...span }))];
      if (merges.length > 0) next.merges = merges;
      else delete next.merges;
    }
    const replayed = next as unknown as S;
    workbook = workbook.map((sheet, i) => (i === index ? replayed : sheet));
    changed.push(target.id);
  }
  return { ok: true, workbook, changedSheetIds: changed };
}

