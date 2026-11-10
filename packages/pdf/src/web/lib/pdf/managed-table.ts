/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * 表格 on a PDF page (ONLYOFFICE's and Word's Insert Table): a table made in
 * MRAI, kept as one managed annotation the native writer draws in vector and
 * opens again as a table (apps/api managed_table.py, the same model). The
 * page keeps it as its directive (MRNOTE_NATIVE_TABLE_V1) among the page's
 * annotations, as it keeps a shape or a text box.
 *
 * Lengths are fractions of the page (x, width of its width; y, height, a
 * row's least height and the font size of its height); the columns are their
 * shares of the table's width. A merged cell spans columns and rows; a slot
 * its span covers is null.
 */

export const MANAGED_NATIVE_TABLE_PREFIX = "[[MRNOTE_NATIVE_TABLE_V1|";
export const PDF_TABLE_MAX_COLUMNS = 63;
export const PDF_TABLE_MAX_ROWS = 500;
export const PDF_TABLE_MAX_CELL_TEXT = 4000;
const MAX_MODEL_BYTES = 256 * 1024;
/** The server's cell inner margin and line height (managed_table.py). */
export const PDF_TABLE_CELL_PADDING_EM = 0.45;
export const PDF_TABLE_LINE_HEIGHT = 1.2;

export type PdfTableAlign = "left" | "center" | "right";
export type PdfTableVerticalAlign = "top" | "middle" | "bottom";

export interface PdfTableCell {
  readonly text: string;
  readonly fill?: string;
  readonly color?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly align?: PdfTableAlign;
  readonly verticalAlign?: PdfTableVerticalAlign;
  readonly colSpan?: number;
  readonly rowSpan?: number;
}

export interface PdfTableRow {
  /** Least height, a fraction of the page's height. */
  readonly height: number;
  readonly cells: readonly (PdfTableCell | null)[];
}

export interface PdfTable {
  readonly id: string;
  /** One-based page number. */
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly columns: readonly number[];
  readonly rows: readonly PdfTableRow[];
  readonly fontSize: number;
  readonly font: "sans" | "serif" | "mono";
  /** #RRGGBB, or null for no grid. */
  readonly borderColor: string | null;
  /** Points. */
  readonly borderWidth: number;
  readonly deleted?: false;
}

export type ManagedNativeTableRecord =
  | PdfTable
  | { readonly id: string; readonly page: number; readonly deleted: true };

export interface ManagedNativeTableAnnotationShadow {
  id: string;
  text: string;
  page: number;
  rect: [number, number, number, number];
}

const HEX = /^#[0-9A-Fa-f]{6}$/;

/** A new table of `rows` x `columns` at (x, y): columns of one width, the
 * first row a header (Word's and PowerPoint's default table look: a header
 * in the accent with white bold words, the grid in the accent's dark). */
export function createPdfTable(input: {
  id: string;
  page: number;
  x: number;
  y: number;
  rows: number;
  columns: number;
  /** The page's width over its height. */
  pageAspect: number;
  fontSize: number;
}): PdfTable {
  const columns = Math.max(1, Math.min(PDF_TABLE_MAX_COLUMNS, Math.round(input.columns)));
  const rows = Math.max(1, Math.min(PDF_TABLE_MAX_ROWS, Math.round(input.rows)));
  const rowHeight = input.fontSize * (PDF_TABLE_LINE_HEIGHT + 2 * PDF_TABLE_CELL_PADDING_EM);
  // A column about seven characters of its size wide, within the page.
  const width = Math.min(1 - input.x, Math.max(0.12, columns * input.fontSize * 7 / input.pageAspect));
  const height = Math.min(1 - input.y, rows * rowHeight);
  return {
    id: input.id,
    page: input.page,
    x: input.x,
    y: input.y,
    width,
    height,
    columns: Array.from({ length: columns }, () => 1 / columns),
    rows: Array.from({ length: rows }, (_, index) => ({
      height: rowHeight,
      cells: Array.from({ length: columns }, () =>
        index === 0 ? { text: "", fill: "#4472C4", color: "#FFFFFF", bold: true } : { text: "" }),
    })),
    fontSize: input.fontSize,
    font: "sans",
    borderColor: "#2F5597",
    borderWidth: 0.75,
  };
}

function round(value: number): number {
  return Number(value.toFixed(6));
}

/** The model as the server reads it (managed_table.py `parse_table_model`). */
export function pdfTableModelJson(table: PdfTable): Record<string, unknown> {
  const total = table.columns.reduce((sum, share) => sum + share, 0) || 1;
  const shares = table.columns.map((share) => round(share / total));
  // Rounded shares still sum to one.
  shares[shares.length - 1] = round(1 - shares.slice(0, -1).reduce((sum, share) => sum + share, 0));
  return {
    v: 1,
    cols: shares,
    rows: table.rows.map((row) => ({
      h: round(row.height),
      cells: row.cells.map((cell) => cell === null ? null : {
        ...(cell.text ? { t: cell.text } : {}),
        ...(cell.fill ? { f: cell.fill.toUpperCase() } : {}),
        ...(cell.color && cell.color.toUpperCase() !== "#000000" ? { c: cell.color.toUpperCase() } : {}),
        ...(cell.bold ? { b: true } : {}),
        ...(cell.italic ? { i: true } : {}),
        ...(cell.align && cell.align !== "left" ? { a: cell.align[0] } : {}),
        ...(cell.verticalAlign && cell.verticalAlign !== "top" ? { v: cell.verticalAlign[0] } : {}),
        ...(cell.colSpan && cell.colSpan > 1 ? { cs: cell.colSpan } : {}),
        ...(cell.rowSpan && cell.rowSpan > 1 ? { rs: cell.rowSpan } : {}),
      }),
    })),
    fs: round(table.fontSize),
    font: table.font,
    ...(table.borderColor ? { bc: table.borderColor.toUpperCase() } : {}),
    bw: round(table.borderWidth),
  };
}

function base64Url(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replace(/=+$/, "");
}

function fromBase64Url(token: string): string | null {
  try {
    const binary = atob(token.replaceAll("-", "+").replaceAll("_", "/") + "=".repeat((4 - (token.length % 4)) % 4));
    return new TextDecoder("utf-8", { fatal: true }).decode(Uint8Array.from(binary, (ch) => ch.charCodeAt(0)));
  } catch {
    return null;
  }
}

function textOk(value: unknown, limit: number): value is string {
  return typeof value === "string" && value.length <= limit &&
    !/[\u0000-\u0008\u000B-\u001F\u007F]/.test(value) && !/[\uD800-\uDFFF]/.test(value.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, ""));
}

/** The table's model read back from its JSON, checked as the server checks
 * it; null for anything the window does not write. */
export function readPdfTableModel(
  value: unknown,
): Pick<PdfTable, "columns" | "rows" | "fontSize" | "font" | "borderColor" | "borderWidth"> | null {
  if (!value || typeof value !== "object") return null;
  const model = value as Record<string, unknown>;
  const allowed = new Set(["v", "cols", "rows", "fs", "font", "bc", "bw"]);
  if (model.v !== 1 || Object.keys(model).some((key) => !allowed.has(key))) return null;
  const columns = model.cols;
  const rows = model.rows;
  const number = (item: unknown, low: number, high: number): item is number =>
    typeof item === "number" && Number.isFinite(item) && item >= low && item <= high;
  if (!Array.isArray(columns) || columns.length < 1 || columns.length > PDF_TABLE_MAX_COLUMNS ||
    !columns.every((share) => number(share, 1e-4, 1))) return null;
  if (Math.abs(columns.reduce((sum: number, share: number) => sum + share, 0) - 1) > 1e-3) return null;
  if (!Array.isArray(rows) || rows.length < 1 || rows.length > PDF_TABLE_MAX_ROWS) return null;
  if (!number(model.fs, 1e-4, 0.2)) return null;
  const font = model.font ?? "sans";
  if (font !== "sans" && font !== "serif" && font !== "mono") return null;
  const borderColor = model.bc ?? null;
  if (borderColor !== null && (typeof borderColor !== "string" || !HEX.test(borderColor))) return null;
  const borderWidth = model.bw ?? 0.75;
  if (!number(borderWidth, 0, 12)) return null;
  const width = columns.length;
  const covered = new Set<string>();
  const parsedRows: PdfTableRow[] = [];
  const cellKeys = new Set(["t", "f", "c", "b", "i", "a", "v", "cs", "rs"]);
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] as Record<string, unknown> | null;
    if (!row || typeof row !== "object" || Object.keys(row).sort().join(",") !== "cells,h") return null;
    if (!number(row.h, 0, 1) || !Array.isArray(row.cells) || row.cells.length !== width) return null;
    const cells: (PdfTableCell | null)[] = [];
    for (let columnIndex = 0; columnIndex < width; columnIndex += 1) {
      const slot = row.cells[columnIndex] as Record<string, unknown> | null;
      const key = `${rowIndex}:${columnIndex}`;
      if (slot === null) {
        if (!covered.has(key)) return null;
        cells.push(null);
        continue;
      }
      if (covered.has(key) || !slot || typeof slot !== "object" || Object.keys(slot).some((k) => !cellKeys.has(k))) return null;
      const text = slot.t ?? "";
      if (!textOk(text, PDF_TABLE_MAX_CELL_TEXT)) return null;
      const fill = slot.f;
      const color = slot.c ?? "#000000";
      if ((fill !== undefined && (typeof fill !== "string" || !HEX.test(fill))) ||
        typeof color !== "string" || !HEX.test(color)) return null;
      const align = slot.a ?? "l";
      const vertical = slot.v ?? "t";
      if (!["l", "c", "r"].includes(align as string) || !["t", "m", "b"].includes(vertical as string)) return null;
      const bold = slot.b ?? false;
      const italic = slot.i ?? false;
      if (typeof bold !== "boolean" || typeof italic !== "boolean") return null;
      const colSpan = slot.cs ?? 1;
      const rowSpan = slot.rs ?? 1;
      if (!Number.isInteger(colSpan) || !Number.isInteger(rowSpan) || (colSpan as number) < 1 || (rowSpan as number) < 1 ||
        columnIndex + (colSpan as number) > width || rowIndex + (rowSpan as number) > rows.length) return null;
      for (let dr = 0; dr < (rowSpan as number); dr += 1) {
        for (let dc = 0; dc < (colSpan as number); dc += 1) {
          if (dr || dc) covered.add(`${rowIndex + dr}:${columnIndex + dc}`);
        }
      }
      cells.push({
        text,
        ...(fill ? { fill: (fill as string).toUpperCase() } : {}),
        ...(color.toUpperCase() !== "#000000" ? { color: color.toUpperCase() } : {}),
        ...(bold ? { bold: true } : {}),
        ...(italic ? { italic: true } : {}),
        ...(align !== "l" ? { align: align === "c" ? "center" as const : "right" as const } : {}),
        ...(vertical !== "t" ? { verticalAlign: vertical === "m" ? "middle" as const : "bottom" as const } : {}),
        ...((colSpan as number) > 1 ? { colSpan: colSpan as number } : {}),
        ...((rowSpan as number) > 1 ? { rowSpan: rowSpan as number } : {}),
      });
    }
    parsedRows.push({ height: row.h, cells });
  }
  for (let rowIndex = 0; rowIndex < parsedRows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < width; columnIndex += 1) {
      if (parsedRows[rowIndex].cells[columnIndex] === null && !covered.has(`${rowIndex}:${columnIndex}`)) return null;
    }
  }
  return { columns: [...columns], rows: parsedRows, fontSize: model.fs, font, borderColor: borderColor ? borderColor.toUpperCase() : null, borderWidth };
}

function encodeTableId(id: string): string | null {
  const trimmed = id.trim();
  if (!trimmed) return null;
  const encoded = encodeURIComponent(trimmed).replace(/[!'()*]/g,
    (character) => `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`);
  const token = `table.${encoded}`;
  return token.length <= 384 ? token : null;
}

function decodeTableId(token: string): string | null {
  if (!token.startsWith("table.")) return null;
  try {
    return decodeURIComponent(token.slice("table.".length)) || null;
  } catch {
    return null;
  }
}

function boxToken(value: number): string {
  return String(Number(Math.min(1, Math.max(0, value)).toFixed(6)));
}

/** The table's directive (managed_table.py); null for a table the writer
 * cannot take. */
export function managedNativeTableDirective(
  record: ManagedNativeTableRecord,
  rotation: 0 | 90 | 180 | 270 | "auto" = "auto",
): string | null {
  const nativeId = encodeTableId(record.id);
  if (!nativeId || !Number.isInteger(record.page) || record.page < 1) return null;
  if (record.deleted) return `${MANAGED_NATIVE_TABLE_PREFIX}delete|${nativeId}]]`;
  const json = JSON.stringify(pdfTableModelJson(record));
  if (new TextEncoder().encode(json).length > MAX_MODEL_BYTES) return null;
  if (!readPdfTableModel(JSON.parse(json))) return null;
  const { x, y, width, height } = record;
  if (![x, y, width, height].every(Number.isFinite) || x < 0 || y < 0 || x >= 1 || y >= 1 ||
    width <= 0 || height <= 0 || x + width > 1 + 1e-6 || y + height > 1 + 1e-6) return null;
  return `${MANAGED_NATIVE_TABLE_PREFIX}upsert|${nativeId}|${rotation}|${[x, y, width, height].map(boxToken).join(",")}|${base64Url(json)}]]`;
}

const DIRECTIVE_RE = /^\[\[MRNOTE_NATIVE_TABLE_V1\|(upsert|delete)\|(table\.[A-Za-z0-9._~%-]{1,378})(?:\|(auto|0|90|180|270)\|([0-9.]{1,20}),([0-9.]{1,20}),([0-9.]{1,20}),([0-9.]{1,20})\|([A-Za-z0-9_-]{4,400000}))?\]\]$/;

/** A page annotation row back as the table it is (null for any other). */
export function parseManagedNativeTableShadow(value: unknown): ManagedNativeTableRecord | null {
  if (!value || typeof value !== "object") return null;
  const row = value as Record<string, unknown>;
  const page = Number(row.page);
  if (!Number.isInteger(page) || page < 1 || page > 10_000) return null;
  const text = typeof row.text === "string" ? row.text : "";
  const match = DIRECTIVE_RE.exec(text);
  if (!match) return null;
  const id = decodeTableId(match[2]);
  if (!id) return null;
  if (match[1] === "delete") return match[3] === undefined ? { id, page, deleted: true } : null;
  if (match[3] === undefined) return null;
  const json = fromBase64Url(match[8]);
  if (json === null) return null;
  let model: ReturnType<typeof readPdfTableModel>;
  try {
    model = readPdfTableModel(JSON.parse(json));
  } catch {
    return null;
  }
  if (!model) return null;
  const [x, y, width, height] = [match[4], match[5], match[6], match[7]].map(Number);
  return { id, page, x, y, width, height, ...model };
}

export interface ManagedNativeTablePlan {
  annotations: ManagedNativeTableAnnotationShadow[];
  supported: boolean;
  reason?: "invalid_table" | "duplicate_id";
}

/** Every table of the document as page annotation rows; one the writer
 * cannot take fails the plan closed (no save applies a part of it). */
export function planManagedNativeTableAnnotations(
  records: readonly ManagedNativeTableRecord[],
  rotations: ReadonlyMap<number, 0 | 90 | 180 | 270> = new Map(),
): ManagedNativeTablePlan {
  const ids = new Set<string>();
  const annotations: ManagedNativeTableAnnotationShadow[] = [];
  for (const record of records) {
    if (ids.has(record.id)) return { annotations: [], supported: false, reason: "duplicate_id" };
    ids.add(record.id);
    const text = managedNativeTableDirective(record, rotations.get(record.page) ?? "auto");
    if (!text) return { annotations: [], supported: false, reason: "invalid_table" };
    annotations.push({
      id: `__mrnote_native_table__${record.id}`,
      text,
      page: record.page,
      // The writer sets the /Rect from the table; this placeholder only
      // satisfies the shared annotation transport.
      rect: [24, 24, 44, 44],
    });
  }
  return { annotations, supported: true };
}

/** The table's slot grid: for each row and column, the cell that owns it
 * (a merged cell owns every slot it spans). */
export function pdfTableOwners(table: Pick<PdfTable, "rows" | "columns">): { row: number; column: number }[][] {
  const owners = table.rows.map(() => table.columns.map(() => ({ row: -1, column: -1 })));
  table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, columnIndex) => {
    if (!cell) return;
    for (let dr = 0; dr < (cell.rowSpan ?? 1); dr += 1) {
      for (let dc = 0; dc < (cell.colSpan ?? 1); dc += 1) {
        owners[rowIndex + dr][columnIndex + dc] = { row: rowIndex, column: columnIndex };
      }
    }
  }));
  return owners;
}

function blankLike(cell: PdfTableCell | null | undefined): PdfTableCell {
  // A new row or column takes its neighbour's look, not its words.
  return cell ? { text: "", ...(cell.fill ? { fill: cell.fill } : {}), ...(cell.color ? { color: cell.color } : {}),
    ...(cell.bold ? { bold: true } : {}), ...(cell.italic ? { italic: true } : {}),
    ...(cell.align ? { align: cell.align } : {}), ...(cell.verticalAlign ? { verticalAlign: cell.verticalAlign } : {}) }
    : { text: "" };
}

interface PlacedTableCell { row: number; column: number; rowSpan: number; colSpan: number; cell: PdfTableCell }

function placedCells(table: Pick<PdfTable, "rows">): PlacedTableCell[] {
  return table.rows.flatMap((row, rowIndex) => row.cells.flatMap((cell, column) =>
    cell ? [{ row: rowIndex, column, rowSpan: cell.rowSpan ?? 1, colSpan: cell.colSpan ?? 1, cell }] : []));
}

/** The grid of rows the placed cells make: each at its place, its spans
 * written on it, the slots it covers null. */
function rebuild(placed: readonly PlacedTableCell[], heights: readonly number[], width: number): PdfTableRow[] {
  const rows: (PdfTableCell | null)[][] = heights.map(() => Array.from({ length: width }, () => null));
  for (const entry of placed) {
    const look: PdfTableCell = { ...entry.cell };
    delete (look as { colSpan?: number }).colSpan;
    delete (look as { rowSpan?: number }).rowSpan;
    rows[entry.row][entry.column] = {
      ...look,
      ...(entry.colSpan > 1 ? { colSpan: entry.colSpan } : {}),
      ...(entry.rowSpan > 1 ? { rowSpan: entry.rowSpan } : {}),
    };
  }
  return rows.map((cells, index) => ({ height: heights[index], cells }));
}

/** A row inserted above or below `row` (Word's Insert Above / Below): it
 * takes that row's look; a merged cell across its place grows by it. */
export function insertPdfTableRow(table: PdfTable, row: number, below: boolean): PdfTable {
  if (table.rows.length >= PDF_TABLE_MAX_ROWS) return table;
  const at = below ? row + 1 : row;
  const owners = pdfTableOwners(table);
  const sourceRow = Math.max(0, Math.min(row, table.rows.length - 1));
  const placed = placedCells(table).map((entry) =>
    entry.row >= at ? { ...entry, row: entry.row + 1 }
      : entry.row + entry.rowSpan > at ? { ...entry, rowSpan: entry.rowSpan + 1 } : entry);
  for (let column = 0; column < table.columns.length; column += 1) {
    // A cell that grew across the new row covers its slot there.
    const crossing = placed.some((entry) => entry.row < at && entry.row + entry.rowSpan > at &&
      column >= entry.column && column < entry.column + entry.colSpan);
    if (crossing) continue;
    const owner = owners[sourceRow][column];
    if (owner.column !== column && owner.row === owners[sourceRow][column - 1]?.row && owner.column === owners[sourceRow][column - 1]?.column) {
      // Inside a cell merged across columns: the new row's cell spans as it does.
      const left = placed.find((entry) => entry.row === at && entry.column < column &&
        entry.column + entry.colSpan === column);
      if (left) { left.colSpan += 1; continue; }
    }
    placed.push({ row: at, column, rowSpan: 1, colSpan: 1, cell: blankLike(table.rows[owner.row].cells[owner.column]) });
  }
  const heights = [...table.rows.map((r) => r.height)];
  heights.splice(at, 0, table.rows[sourceRow].height);
  return { ...table, rows: rebuild(placed, heights, table.columns.length), height: table.height + table.rows[sourceRow].height };
}

/** A column inserted left or right of `column`, taking its look; the
 * columns share the table's width with it. */
export function insertPdfTableColumn(table: PdfTable, column: number, right: boolean): PdfTable {
  if (table.columns.length >= PDF_TABLE_MAX_COLUMNS) return table;
  const at = right ? column + 1 : column;
  const owners = pdfTableOwners(table);
  const sourceColumn = Math.max(0, Math.min(column, table.columns.length - 1));
  const placed = placedCells(table).map((entry) =>
    entry.column >= at ? { ...entry, column: entry.column + 1 }
      : entry.column + entry.colSpan > at ? { ...entry, colSpan: entry.colSpan + 1 } : entry);
  for (let row = 0; row < table.rows.length; row += 1) {
    const covered = placed.some((entry) => row >= entry.row && row < entry.row + entry.rowSpan &&
      at >= entry.column && at < entry.column + entry.colSpan);
    if (covered) continue;
    const owner = owners[row][sourceColumn];
    if (owner.row !== row) {
      // Inside a cell merged down: the new column's cell spans down as it does.
      const upper = placed.find((entry) => entry.column === at && entry.row + entry.rowSpan === row &&
        owners[entry.row][sourceColumn].row === owner.row);
      if (upper) { upper.rowSpan += 1; continue; }
    }
    placed.push({ row, column: at, rowSpan: 1, colSpan: 1, cell: blankLike(table.rows[owner.row].cells[owner.column]) });
  }
  const share = table.columns[sourceColumn];
  const total = table.columns.reduce((sum, value) => sum + value, 0) + share;
  const columns = [...table.columns.slice(0, at), share, ...table.columns.slice(at)].map((value) => value / total);
  return { ...table, columns, rows: rebuild(placed, table.rows.map((r) => r.height), columns.length) };
}

/** Row `row` taken out (a merged cell across it shrinks; one starting on
 * it moves down a row); null when it was the table's last. */
export function deletePdfTableRow(table: PdfTable, row: number): PdfTable | null {
  if (table.rows.length <= 1) return null;
  const placed = placedCells(table).flatMap((entry) => {
    if (entry.row === row) return entry.rowSpan > 1 ? [{ ...entry, rowSpan: entry.rowSpan - 1 }] : [];
    if (entry.row < row && entry.row + entry.rowSpan > row) return [{ ...entry, rowSpan: entry.rowSpan - 1 }];
    return [entry.row > row ? { ...entry, row: entry.row - 1 } : entry];
  });
  const heights = table.rows.map((r) => r.height).filter((_, index) => index !== row);
  const removed = table.rows[row].height;
  return { ...table, rows: rebuild(placed, heights, table.columns.length), height: Math.max(removed, table.height - removed) };
}

/** Column `column` taken out (its share goes to the others; a merged cell
 * across it shrinks); null when it was the table's last. */
export function deletePdfTableColumn(table: PdfTable, column: number): PdfTable | null {
  if (table.columns.length <= 1) return null;
  const placed = placedCells(table).flatMap((entry) => {
    if (entry.column === column) return entry.colSpan > 1 ? [{ ...entry, colSpan: entry.colSpan - 1 }] : [];
    if (entry.column < column && entry.column + entry.colSpan > column) return [{ ...entry, colSpan: entry.colSpan - 1 }];
    return [entry.column > column ? { ...entry, column: entry.column - 1 } : entry];
  });
  const columns = table.columns.filter((_, index) => index !== column);
  const total = columns.reduce((sum, value) => sum + value, 0);
  return { ...table, columns: columns.map((value) => value / total), rows: rebuild(placed, table.rows.map((r) => r.height), columns.length) };
}

/** One cell changed (its words or its look). */
export function updatePdfTableCell(
  table: PdfTable, row: number, column: number, change: Partial<PdfTableCell>,
): PdfTable {
  const cell = table.rows[row]?.cells[column];
  if (!cell) return table;
  const rows = table.rows.map((existing, index) => index !== row ? existing : {
    ...existing,
    cells: existing.cells.map((slot, slotIndex) => slotIndex === column ? { ...cell, ...change } : slot),
  });
  return { ...table, rows };
}

/** The table's words, a row a line and a cell a tab (its PDF /Contents). */
export function pdfTableText(table: Pick<PdfTable, "rows">): string {
  return table.rows.map((row) => row.cells.filter((cell): cell is PdfTableCell => cell !== null)
    .map((cell) => cell.text.replaceAll("\n", " ")).join("\t")).join("\n");
}

export interface PdfTableRange {
  readonly row: number;
  readonly column: number;
  readonly endRow: number;
  readonly endColumn: number;
}

/** The range from one cell to another, grown to hold every merged cell it
 * touches whole (Word's and PowerPoint's selection of cells). */
export function pdfTableRange(
  table: Pick<PdfTable, "rows" | "columns">,
  from: { row: number; column: number },
  to: { row: number; column: number } = from,
): PdfTableRange {
  let r0 = Math.max(0, Math.min(from.row, to.row));
  let r1 = Math.min(table.rows.length - 1, Math.max(from.row, to.row));
  let c0 = Math.max(0, Math.min(from.column, to.column));
  let c1 = Math.min(table.columns.length - 1, Math.max(from.column, to.column));
  const placed = placedCells(table);
  for (let changed = true; changed;) {
    changed = false;
    for (const entry of placed) {
      const er1 = entry.row + entry.rowSpan - 1;
      const ec1 = entry.column + entry.colSpan - 1;
      if (entry.row > r1 || er1 < r0 || entry.column > c1 || ec1 < c0) continue;
      if (entry.row < r0) { r0 = entry.row; changed = true; }
      if (er1 > r1) { r1 = er1; changed = true; }
      if (entry.column < c0) { c0 = entry.column; changed = true; }
      if (ec1 > c1) { c1 = ec1; changed = true; }
    }
  }
  return { row: r0, column: c0, endRow: r1, endColumn: c1 };
}

function inRange(entry: { row: number; column: number }, range: PdfTableRange): boolean {
  return entry.row >= range.row && entry.row <= range.endRow && entry.column >= range.column && entry.column <= range.endColumn;
}

/** Every cell of the range changed alike (its look: weight, slant,
 * alignment, fill, ink). */
export function updatePdfTableCells(table: PdfTable, range: PdfTableRange, change: Partial<PdfTableCell>): PdfTable {
  const { text: _ignored, colSpan: _c, rowSpan: _r, ...look } = change;
  void _ignored; void _c; void _r;
  const placed = placedCells(table).map((entry) => inRange(entry, range)
    ? { ...entry, cell: Object.fromEntries(Object.entries({ ...entry.cell, ...look }).filter(([, value]) => value !== undefined)) as unknown as PdfTableCell }
    : entry);
  return { ...table, rows: rebuild(placed, table.rows.map((row) => row.height), table.columns.length) };
}

/** The range merged into one cell (合并单元格): the first cell's look, the
 * words of every cell that has some, each on its own line, in reading order
 * (Word keeps each cell's words as a paragraph). */
export function mergePdfTableCells(table: PdfTable, range: PdfTableRange): PdfTable {
  const placed = placedCells(table);
  const inside = placed.filter((entry) => inRange(entry, range));
  if (inside.length < 2) return table;
  const first = inside.find((entry) => entry.row === range.row && entry.column === range.column) ?? inside[0];
  const text = inside.map((entry) => entry.cell.text).filter((words) => words.trim()).join("\n").slice(0, PDF_TABLE_MAX_CELL_TEXT);
  const merged: PlacedTableCell = {
    row: range.row, column: range.column,
    rowSpan: range.endRow - range.row + 1, colSpan: range.endColumn - range.column + 1,
    cell: { ...first.cell, text },
  };
  return { ...table, rows: rebuild([...placed.filter((entry) => !inRange(entry, range)), merged],
    table.rows.map((row) => row.height), table.columns.length) };
}

/** A merged cell split back into the cells it spans (拆分单元格): its words
 * stay in the first; the others are empty, in its look. */
export function splitPdfTableCell(table: PdfTable, row: number, column: number): PdfTable {
  const placed = placedCells(table);
  const target = placed.find((entry) => entry.row === row && entry.column === column);
  if (!target || (target.rowSpan === 1 && target.colSpan === 1)) return table;
  const pieces: PlacedTableCell[] = [];
  for (let dr = 0; dr < target.rowSpan; dr += 1) {
    for (let dc = 0; dc < target.colSpan; dc += 1) {
      pieces.push({ row: row + dr, column: column + dc, rowSpan: 1, colSpan: 1,
        cell: dr === 0 && dc === 0 ? target.cell : blankLike(target.cell) });
    }
  }
  return { ...table, rows: rebuild([...placed.filter((entry) => entry !== target), ...pieces],
    table.rows.map((r) => r.height), table.columns.length) };
}

/** The table's header and band looks (Word's and PowerPoint's default
 * table style: a header in the accent, white bold words; bands in its
 * lightest tint). */
export const PDF_TABLE_HEADER_FILL = "#4472C4";
export const PDF_TABLE_BAND_FILL = "#D9E2F3";

export function pdfTableHasHeaderRow(table: Pick<PdfTable, "rows">): boolean {
  return table.rows[0]?.cells.every((cell) => cell === null || cell.fill?.toUpperCase() === PDF_TABLE_HEADER_FILL) ?? false;
}

export function pdfTableHasBandedRows(table: Pick<PdfTable, "rows">): boolean {
  return table.rows.some((row) => row.cells.some((cell) => cell?.fill?.toUpperCase() === PDF_TABLE_BAND_FILL));
}

/** 标题行 on or off: the first row in the header look, or plain. */
export function setPdfTableHeaderRow(table: PdfTable, on: boolean): PdfTable {
  const range = { row: 0, column: 0, endRow: 0, endColumn: table.columns.length - 1 };
  return updatePdfTableCells(table, range, on
    ? { fill: PDF_TABLE_HEADER_FILL, color: "#FFFFFF", bold: true }
    : { fill: undefined, color: undefined, bold: undefined });
}

/** 镶边行 on or off: every other body row in the band tint. */
export function setPdfTableBandedRows(table: PdfTable, on: boolean): PdfTable {
  const header = pdfTableHasHeaderRow(table) ? 1 : 0;
  let next = table;
  for (let row = header; row < table.rows.length; row += 1) {
    const banded = on && (row - header) % 2 === 0;
    const range = { row, column: 0, endRow: row, endColumn: table.columns.length - 1 };
    const cells = table.rows[row].cells.filter((cell): cell is PdfTableCell => cell !== null);
    if (banded) next = updatePdfTableCells(next, range, { fill: PDF_TABLE_BAND_FILL });
    else if (cells.some((cell) => cell.fill?.toUpperCase() === PDF_TABLE_BAND_FILL)) {
      next = updatePdfTableCells(next, range, { fill: undefined });
    }
  }
  return next;
}
