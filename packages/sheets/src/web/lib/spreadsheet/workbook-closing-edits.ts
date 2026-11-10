/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a window closing on a workbook still has to say, small enough to go out
 * as it closes. A large workbook's first open prepares its rooms for a minute
 * or more, and saves wait for them: a cell typed meanwhile was only in this
 * window (a new 60,000-row sheet, closed at 27 s, lost its edit). The change from
 * the stored body is the ``content_delta`` a save's answer uses, in reverse.
 */

import {
  PAGE_CLOSING_EDITS_MAX_BYTES,
  closingEditsBody,
  waitForClosingEdits as waitForPageClosingEdits,
} from "@mrnote/core/web/lib/page-closing-edits.ts";

export { closingEditsSentRecently, noteClosingEditsSent } from "@mrnote/core/web/lib/page-closing-edits.ts";

/** A request sent as a page closes may carry at most 64 KiB (`keepalive`): the
 * budget every closing page shares (lib/page-closing-edits). */
export const WORKBOOK_CLOSING_EDITS_MAX_BYTES = PAGE_CLOSING_EDITS_MAX_BYTES;

export interface WorkbookClosingSheetDelta {
  fields: Record<string, unknown>;
  removed_fields: string[];
  cells: Record<string, unknown>;
  removed: string[];
}

export interface WorkbookClosingEdits {
  base_content_version: number;
  sheets: Record<string, WorkbookClosingSheetDelta>;
}

type Dict = Record<string, unknown>;

function isDict(value: unknown): value is Dict {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** The same JSON value: a read body and the window's own write members in
 * different orders, so strings of them differ on every cell. */
function same(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Array.isArray(a)) {
    return Array.isArray(b) && a.length === b.length && a.every((item, index) => same(item, b[index]));
  }
  if (!isDict(a) || !isDict(b)) return false;
  const keys = Object.keys(a);
  return keys.length === Object.keys(b).length && keys.every((key) => key in b && same(a[key], b[key]));
}

/**
 * The change from `base` to `next` per sheet, or null when there is none or it
 * is not one this form says (a sheet added or moved, a workbook-wide setting).
 * Which sheet is active is this window's own.
 */
export function workbookClosingEdits(
  base: unknown,
  next: unknown,
  baseContentVersion: number,
): WorkbookClosingEdits | null {
  if (!isDict(base) || !isDict(next)) return null;
  const baseSheets = base.sheets;
  const nextSheets = next.sheets;
  if (!Array.isArray(baseSheets) || !Array.isArray(nextSheets) || baseSheets.length !== nextSheets.length) return null;
  const keys = new Set([...Object.keys(base), ...Object.keys(next)]);
  for (const key of keys) {
    if (key !== "sheets" && key !== "activeSheetId" && !same(base[key], next[key])) return null;
  }
  const sheets: Record<string, WorkbookClosingSheetDelta> = {};
  for (let index = 0; index < baseSheets.length; index += 1) {
    const before: unknown = baseSheets[index];
    const after: unknown = nextSheets[index];
    if (before === after) continue;
    if (!isDict(before) || !isDict(after) || before.id !== after.id) return null;
    const beforeCells = before.cells;
    const afterCells = after.cells;
    if (!isDict(beforeCells) || !isDict(afterCells)) return null;
    const delta: WorkbookClosingSheetDelta = { fields: {}, removed_fields: [], cells: {}, removed: [] };
    for (const field of Object.keys(after)) {
      if (field !== "cells" && !same(before[field], after[field])) delta.fields[field] = after[field];
    }
    for (const field of Object.keys(before)) {
      if (!(field in after)) delta.removed_fields.push(field);
    }
    if (beforeCells !== afterCells) {
      for (const key of Object.keys(afterCells)) {
        if (!same(beforeCells[key], afterCells[key])) delta.cells[key] = afterCells[key];
      }
      for (const key of Object.keys(beforeCells)) {
        if (!(key in afterCells)) delta.removed.push(key);
      }
    }
    if (
      Object.keys(delta.fields).length > 0 || delta.removed_fields.length > 0 ||
      Object.keys(delta.cells).length > 0 || delta.removed.length > 0
    ) {
      sheets[String(index)] = delta;
    }
  }
  if (Object.keys(sheets).length === 0) return null;
  return { base_content_version: baseContentVersion, sheets };
}

/** The request body, when it fits a request sent as a page closes. */
export function workbookClosingEditsBody(edits: WorkbookClosingEdits): string | null {
  return closingEditsBody(edits);
}

// The window that opens next may read the page before those edits are saved. It
// then shows the older page while a save lands behind it, and what is typed there
// is typed on the wrong page. So it opens after that save, by a small read of the
// version; and a recovery copy of the same edits waits for it. A save that does not
// come ends both waits by time: the page opens, and the copy is offered as ever.
// The note and the wait are every closing page's (lib/page-closing-edits).

/**
 * Before reading a page that a window just left with edits on their way: wait
 * until they are saved. True when the page moved past the version they were made
 * on; false when nothing was sent, or nothing came within the wait.
 */
export async function waitForClosingEdits(
  pageId: string,
  readVersion: (pageId: string) => Promise<{ version: number; roomsBehind: boolean } | null>,
  isCancelled: () => boolean,
): Promise<boolean> {
  return waitForPageClosingEdits(pageId, async (id, sent) => {
    const read = await readVersion(id);
    if (read === null) return null;
    return read.version > sent.base && !read.roomsBehind ? "landed" : "pending";
  }, isCancelled);
}

function cellWords(cell: unknown): string {
  if (!isDict(cell)) return "";
  return String(cell.formula ?? cell.value ?? "");
}

/**
 * Whether cells typed since `base` are not in the page the server holds, although
 * the window reads as saved: its view was replaced under it. Only what a person
 * wrote counts (a cell's words or formula), so a cell the server reads back in
 * another form, or another member's edit elsewhere, is not taken for a loss.
 * Nor is another member's later edit of the same cell: a cell that holds
 * neither these words nor the ones they replaced was written by someone after
 * them (two people committing one cell at once), and the last one written
 * stands, as in Excel's co-authoring. Offered as unsaved, it opened Document
 * Recovery over the grid of the member whose words were saved first.
 */
export function typedCellsMissingFromPage(base: unknown, typed: unknown, page: unknown): boolean {
  const edits = workbookClosingEdits(base, typed, 0);
  if (!edits || !isDict(page) || !Array.isArray(page.sheets) || "tiledWorkbook" in page) return false;
  const baseSheets = isDict(base) && Array.isArray(base.sheets) ? base.sheets : [];
  for (const [index, delta] of Object.entries(edits.sheets)) {
    const sheet: unknown = page.sheets[Number(index)];
    if (!isDict(sheet) || !isDict(sheet.cells)) return true;
    const before: unknown = baseSheets[Number(index)];
    const beforeCells = isDict(before) && isDict(before.cells) ? before.cells : {};
    // The page still holds what the cell held before it was typed in.
    const cells: Dict = sheet.cells;
    const unchangedSince = (key: string) => cellWords(cells[key]) === cellWords(beforeCells[key]) &&
      (key in cells) === (key in beforeCells);
    for (const [key, cell] of Object.entries(delta.cells)) {
      if (cellWords(cell) !== cellWords(cells[key]) && unchangedSince(key)) return true;
    }
    for (const key of delta.removed) {
      if (key in cells && unchangedSince(key)) return true;
    }
  }
  return false;
}

/**
 * `edits` laid on `target` by the cell, as the server lays closing edits with
 * merge: the cells and sheet members they name, the bounds the larger of the two.
 * Null when they do not fit it (a sheet it no longer has at that place).
 */
export function applyWorkbookEdits(target: unknown, edits: WorkbookClosingEdits): Record<string, unknown> | null {
  if (!isDict(target) || !Array.isArray(target.sheets)) return null;
  const sheets = [...target.sheets];
  for (const [index, delta] of Object.entries(edits.sheets)) {
    const sheet: unknown = sheets[Number(index)];
    if (!isDict(sheet) || !isDict(sheet.cells)) return null;
    const next: Dict = { ...sheet, ...delta.fields };
    for (const bound of ["rowCount", "columnCount"]) {
      const ours = delta.fields[bound], theirs = sheet[bound];
      if (typeof ours === "number" && typeof theirs === "number") next[bound] = Math.max(ours, theirs);
    }
    for (const field of delta.removed_fields) delete next[field];
    const cells: Dict = { ...sheet.cells };
    for (const key of delta.removed) delete cells[key];
    Object.assign(cells, delta.cells);
    next.cells = cells;
    sheets[Number(index)] = next;
  }
  return { ...target, sheets };
}

export function isWorkbookClosingEdits(value: unknown): value is WorkbookClosingEdits {
  return isDict(value) && typeof value.base_content_version === "number" && isDict(value.sheets);
}
