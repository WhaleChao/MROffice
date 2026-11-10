/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Each distinct cell style of a spreadsheet page stored once (`cellStyleTable`).
 *
 * SpreadsheetML keeps a workbook's cell formats once, in `cellXfs`, and a cell
 * names its format by index (`<c r="B2" s="5">`, ECMA-376 Part 1 §18.3.1.4 /
 * §18.8.10). A page used to write the whole format out on every cell: a real
 * 62,791-cell formatted workbook became an 18 MB page and opened read-only.
 *
 * This is the browser half of the contract shared with
 * `apps/api/app/services/spreadsheet_cell_styles.py`:
 *
 *   "cellStyleTable": [{ ...style }, ...]                 // workbook level
 *   "cells": { "3:1": { "value": 5, "kind": "number", "styleRef": 0 } }
 *
 * - `styleRef` is an index into the table and stands where the cell's `style`
 *   key was, so expanding gives back exactly what was packed (key order too).
 * - A cell carries `style` or `styleRef`, never both. Inline `style` is valid
 *   forever: old pages and mixed pages read unchanged.
 * - The editor model, collaboration and every reader keep the inline form.
 *   Only a page body sent to the API is packed (`packCellStyleTable`, called
 *   by the notebook SDK), because that is what the request limit and the
 *   page envelope bound; the API stores that form and answers inline.
 *   `expandCellStyleTable` is the one reader of the stored form, applied
 *   where a page body enters the editor (`normalizeSheets`).
 */

export const SPREADSHEET_CONTENT_TYPE = "mrnote.spreadsheet.v1" as const;
export const CELL_STYLE_TABLE_KEY = "cellStyleTable" as const;
export const CELL_STYLE_REF_KEY = "styleRef" as const;
/** Excel's own `cellXfs` ceiling (MS-OI29500: `s` is 0..65,490); styles past
 * it stay inline. Mirrors MAX_CELL_STYLE_TABLE_ENTRIES. */
export const MAX_CELL_STYLE_TABLE_ENTRIES = 65_490;
/** Styled cells a workbook must hold before its styles are packed; a small
 * page is sent exactly as before. Mirrors MIN_PACKED_STYLED_CELLS. */
export const MIN_PACKED_STYLED_CELLS = 64;

type JsonRecord = Record<string, unknown>;

function isRecord(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isStyle(value: unknown): value is JsonRecord {
  return isRecord(value) && Object.keys(value).length > 0;
}

function isRef(value: unknown, size: number): value is number {
  return typeof value === "number" && Number.isInteger(value) && value >= 0 && value < size;
}

/** A style made read-only, nested records included. */
function freezeStyle<T extends JsonRecord>(style: T): T {
  for (const value of Object.values(style)) {
    if (typeof value === "object" && value !== null && !Object.isFrozen(value)) {
      freezeStyle(value as JsonRecord);
    }
  }
  return Object.freeze(style);
}

/**
 * A pool that hands out one shared, frozen object per distinct cell style.
 *
 * SpreadsheetML names a cell's format by index and Excel keeps each format
 * once; the editor model keeps the inline form, but a workbook of 200,000
 * cells has a few dozen distinct styles, and one object per cell was most of
 * its heap. Cells written with the same style now hold the same object.
 * Every writer replaces a cell's `style` with a new object (copy, then
 * change), so a shared style is never changed in place; it is frozen so that
 * a writer that tried would fail rather than restyle every cell sharing it.
 * Two styles are the same when they serialize the same (key order
 * included), as `packCellStyleTable` decides.
 */
export interface CellStyleInterner {
  (style: JsonRecord): JsonRecord;
}

export function createCellStyleInterner(): CellStyleInterner {
  const byObject = new WeakMap<object, JsonRecord>();
  const byText = new Map<string, JsonRecord>();
  return (style) => {
    if (Object.isFrozen(style)) return style;
    const known = byObject.get(style);
    if (known) return known;
    let text: string;
    try {
      text = JSON.stringify(style);
    } catch {
      return style;
    }
    let shared = byText.get(text);
    if (!shared) {
      shared = freezeStyle(JSON.parse(text) as JsonRecord);
      byText.set(text, shared);
    }
    byObject.set(style, shared);
    return shared;
  };
}

/** Whether `content` is a spreadsheet page body holding a style table. */
export function hasCellStyleTable(content: unknown): boolean {
  return isRecord(content) && CELL_STYLE_TABLE_KEY in content;
}

/**
 * The inline form: every `styleRef` replaced by the table's style, in the
 * reference's place. Cells naming one entry share one frozen object (see
 * `createCellStyleInterner`). Copy-on-write; a body without a table is
 * returned as the same object. A reference that does not resolve is dropped
 * (the API refuses such a body, so only a damaged copy can hold one).
 */
export function expandCellStyleTable<T>(content: T): T {
  if (!isRecord(content) || !(CELL_STYLE_TABLE_KEY in content)) return content;
  const rawTable = content[CELL_STYLE_TABLE_KEY];
  const intern = createCellStyleInterner();
  const table: (JsonRecord | null)[] = Array.isArray(rawTable)
    ? rawTable.map((entry) => (isStyle(entry) ? intern(entry) : null))
    : [];
  const expanded: JsonRecord = {};
  for (const [key, value] of Object.entries(content)) {
    if (key !== CELL_STYLE_TABLE_KEY) expanded[key] = value;
  }
  if (Array.isArray(content.sheets)) {
    expanded.sheets = content.sheets.map((sheet) => {
      if (!isRecord(sheet) || !isRecord(sheet.cells)) return sheet;
      let next: JsonRecord | null = null;
      const entries = Object.entries(sheet.cells);
      for (let index = 0; index < entries.length; index += 1) {
        const [key, cell] = entries[index];
        if (!isRecord(cell) || !(CELL_STYLE_REF_KEY in cell)) {
          if (next) next[key] = cell;
          continue;
        }
        if (!next) {
          next = {};
          for (let previous = 0; previous < index; previous += 1) {
            next[entries[previous][0]] = entries[previous][1];
          }
        }
        const ref = cell[CELL_STYLE_REF_KEY];
        const shared = isRef(ref, table.length) && !("style" in cell) ? table[ref] : null;
        const rebuilt: JsonRecord = {};
        for (const [field, value] of Object.entries(cell)) {
          if (field === CELL_STYLE_REF_KEY) {
            if (shared) rebuilt.style = shared;
          } else {
            rebuilt[field] = value;
          }
        }
        next[key] = rebuilt;
      }
      return next ? { ...sheet, cells: next } : sheet;
    });
  }
  return expanded as T;
}

/**
 * The stored form of a spreadsheet page body: each distinct non-empty cell
 * style once in `cellStyleTable`, each styled cell naming it by `styleRef`.
 * Copy-on-write; anything that is not a spreadsheet body, or holds fewer than
 * MIN_PACKED_STYLED_CELLS styled cells, is returned as the same object. Two
 * styles are the same entry when they serialize the same (key order
 * included), so expanding gives back exactly the inline body.
 */
export function packCellStyleTable<T>(content: T): T {
  if (
    !isRecord(content) ||
    content.type !== SPREADSHEET_CONTENT_TYPE ||
    !Array.isArray(content.sheets)
  ) {
    return content;
  }
  const source = (CELL_STYLE_TABLE_KEY in content ? expandCellStyleTable(content) : content) as JsonRecord;
  const sheets = source.sheets as unknown[];
  let styled = 0;
  for (const sheet of sheets) {
    if (!isRecord(sheet) || !isRecord(sheet.cells)) continue;
    for (const cell of Object.values(sheet.cells)) {
      if (isRecord(cell) && isStyle(cell.style)) styled += 1;
    }
    if (styled >= MIN_PACKED_STYLED_CELLS) break;
  }
  if (styled < MIN_PACKED_STYLED_CELLS) return source as T;
  const table: unknown[] = [];
  const refByText = new Map<string, number>();
  const refByObject = new Map<object, number | null>();
  const nextSheets = sheets.map((sheet) => {
    if (!isRecord(sheet) || !isRecord(sheet.cells)) return sheet;
    const nextCells: JsonRecord = {};
    let changed = false;
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const style = isRecord(cell) ? cell.style : undefined;
      if (!isRecord(cell) || !isStyle(style) || CELL_STYLE_REF_KEY in cell) {
        nextCells[key] = cell;
        continue;
      }
      let ref = refByObject.get(style);
      if (ref === undefined) {
        let text: string | null;
        try {
          text = JSON.stringify(style);
        } catch {
          text = null;
        }
        ref = text === null ? null : refByText.get(text) ?? null;
        if (text !== null && ref === null && table.length < MAX_CELL_STYLE_TABLE_ENTRIES) {
          ref = table.length;
          table.push(JSON.parse(text));
          refByText.set(text, ref);
        }
        refByObject.set(style, ref);
      }
      if (ref === null) {
        nextCells[key] = cell;
        continue;
      }
      const packed: JsonRecord = {};
      for (const [field, value] of Object.entries(cell)) {
        if (field === "style") packed[CELL_STYLE_REF_KEY] = ref;
        else packed[field] = value;
      }
      nextCells[key] = packed;
      changed = true;
    }
    return changed ? { ...sheet, cells: nextCells } : sheet;
  });
  if (table.length === 0) return source as T;
  return { ...source, sheets: nextSheets, [CELL_STYLE_TABLE_KEY]: table } as T;
}
