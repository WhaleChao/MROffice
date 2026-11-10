/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A workbook recovery copy kept beside the page (its base is no version the
 * server confirmed it was made on, so it is never merged on its own): where it
 * differs from the page, cell by cell, and the page with chosen cells taken
 * from it. The member sees each difference and takes or leaves it, as Docs
 * offers its recovered copies; nothing here decides for them.
 */

import { columnName } from "./tiled-workbook";

type Json = unknown;
type Workbook = { sheets?: unknown; [key: string]: unknown };
type Sheet = { id?: unknown; name?: unknown; cells?: unknown; [key: string]: unknown };

export interface KeptWorkbookCopyDifference {
  sheetId: string;
  sheetName: string;
  cellKey: string;
  /** A1 address on its sheet ("B3"). */
  address: string;
  /** What the copy holds there, as shown text ("" when the copy has no cell). */
  copyText: string;
  /** What the page holds there now. */
  pageText: string;
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  Boolean(value) && typeof value === "object" && !Array.isArray(value);

function sheetsOf(workbook: Workbook): Sheet[] {
  return Array.isArray(workbook.sheets) ? (workbook.sheets as unknown[]).filter(isRecord) as Sheet[] : [];
}

function cellsOf(sheet: Sheet | undefined): Record<string, unknown> {
  return sheet && isRecord(sheet.cells) ? sheet.cells : {};
}

/** A cell's data as the member reads it: its formula, else its value. */
function cellData(cell: unknown): { value: Json; formula: Json } {
  return isRecord(cell) ? { value: cell.value ?? null, formula: cell.formula ?? null } : { value: null, formula: null };
}

function cellText(cell: unknown): string {
  const { value, formula } = cellData(cell);
  const shown = typeof formula === "string" && formula ? formula : value;
  return shown === null || shown === undefined ? "" : String(shown);
}

function sameData(left: unknown, right: unknown): boolean {
  const a = cellData(left);
  const b = cellData(right);
  return JSON.stringify(a.value) === JSON.stringify(b.value) && JSON.stringify(a.formula) === JSON.stringify(b.formula);
}

function address(cellKey: string): string {
  const [row, column] = cellKey.split(":").map(Number);
  return Number.isInteger(row) && Number.isInteger(column) ? `${columnName(column)}${row + 1}` : cellKey;
}

/** Every cell whose value or formula differs between the copy and the page,
 * on the sheets both hold, in sheet then row-major order. */
export function keptWorkbookCopyDifferences(copy: Workbook, page: Workbook): KeptWorkbookCopyDifference[] {
  const pageSheets = new Map(sheetsOf(page).map((sheet) => [String(sheet.id), sheet]));
  const out: KeptWorkbookCopyDifference[] = [];
  for (const sheet of sheetsOf(copy)) {
    const current = pageSheets.get(String(sheet.id));
    if (!current) continue;
    const copyCells = cellsOf(sheet);
    const pageCells = cellsOf(current);
    const keys = [...new Set([...Object.keys(copyCells), ...Object.keys(pageCells)])]
      .filter((key) => !sameData(copyCells[key], pageCells[key]))
      .sort((left, right) => {
        const [lr, lc] = left.split(":").map(Number);
        const [rr, rc] = right.split(":").map(Number);
        return lr - rr || lc - rc;
      });
    for (const cellKey of keys) {
      out.push({
        sheetId: String(sheet.id),
        sheetName: typeof current.name === "string" ? current.name : String(sheet.id),
        cellKey,
        address: address(cellKey),
        copyText: cellText(copyCells[cellKey]),
        pageText: cellText(pageCells[cellKey]),
      });
    }
  }
  return out;
}

/** The page with the named cells as the copy holds them (a cell the copy
 * lacks is removed); null when none of them could be taken. */
export function applyKeptWorkbookCopyCells<T extends Workbook>(
  page: T,
  copy: Workbook,
  cells: readonly { sheetId: string; cellKey: string }[],
): T | null {
  const copySheets = new Map(sheetsOf(copy).map((sheet) => [String(sheet.id), sheet]));
  const wanted = new Map<string, string[]>();
  for (const { sheetId, cellKey } of cells) wanted.set(sheetId, [...(wanted.get(sheetId) ?? []), cellKey]);
  let taken = 0;
  const sheets = (Array.isArray(page.sheets) ? page.sheets as unknown[] : []).map((entry) => {
    if (!isRecord(entry)) return entry;
    const keys = wanted.get(String(entry.id));
    const from = copySheets.get(String(entry.id));
    if (!keys || !from) return entry;
    const nextCells = { ...cellsOf(entry as Sheet) };
    const copyCells = cellsOf(from);
    // A cell the copy holds past the page's last row or column (typed into a
    // row the page has not grown to) grows the sheet to hold it, as typing
    // there does; the server refuses a cell outside the declared bounds.
    let rowCount = typeof entry.rowCount === "number" && Number.isFinite(entry.rowCount) ? entry.rowCount : null;
    let columnCount = typeof entry.columnCount === "number" && Number.isFinite(entry.columnCount) ? entry.columnCount : null;
    for (const key of keys) {
      if (Object.hasOwn(copyCells, key)) {
        nextCells[key] = copyCells[key];
        const [row, column] = key.split(":").map(Number);
        if (rowCount !== null && Number.isInteger(row) && row >= rowCount) rowCount = row + 1;
        if (columnCount !== null && Number.isInteger(column) && column >= columnCount) columnCount = column + 1;
      } else delete nextCells[key];
      taken += 1;
    }
    return {
      ...entry,
      cells: nextCells,
      ...(rowCount !== null && rowCount !== entry.rowCount ? { rowCount } : {}),
      ...(columnCount !== null && columnCount !== entry.columnCount ? { columnCount } : {}),
    };
  });
  return taken > 0 ? { ...page, sheets } : null;
}
