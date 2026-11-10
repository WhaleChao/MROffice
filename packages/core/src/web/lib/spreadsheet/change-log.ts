/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The change list behind Review > Show Changes.
 *
 * Excel's Changes pane lists who changed which cell, from what to what, and
 * when, newest first, for every collaborator and session, narrowed on request
 * to a sheet or a range. A bulk edit (a paste, a fill, a clear, a sort) is one
 * card for all of its cells.
 *
 * The workbook knows each edit's before and after at the moment it commits,
 * so this module turns one commit into records (one per worksheet it touched)
 * that the editor uploads to the server's change log, and it narrows, merges
 * and labels the entries the pane shows.
 *
 * What counts follows Excel's rule: a cell's CONTENT (its value or formula).
 * A formatting-only commit, a comment, a chart or a shape is not a change; a
 * formula cell recalculating to a new result is not one either.
 */

import { cellMapChangesSince } from "./cell-map-lineage";
import { columnToLetters } from "./formula-engine";
import { MAX_PARSE_COL, MAX_PARSE_ROW, parseA1Range } from "./print-settings";

/** Cells listed per edit; a bigger paste keeps its count and its range. */
export const CHANGE_LOG_MAX_CELLS = 100;
/** Characters kept of a value; the pane shows a short preview anyway. */
export const CHANGE_LOG_MAX_VALUE_CHARS = 256;

export type SpreadsheetChangeKind =
  | "edit"
  | "undo"
  | "redo"
  | "insert_rows"
  | "delete_rows"
  | "insert_columns"
  | "delete_columns";

export interface SpreadsheetChangeCell {
  r: number;
  c: number;
  before: string;
  after: string;
}

/** One committed edit on one worksheet, as uploaded. */
export interface SpreadsheetChangeRecord {
  clientChangeId: string;
  sheetId: string;
  sheetName: string;
  kind: SpreadsheetChangeKind;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  /** Every cell the edit changed; `cells` lists the first of them. */
  cellCount: number;
  cells: SpreadsheetChangeCell[];
  truncated: boolean;
}

export interface SpreadsheetChangeActor {
  userId: string | null;
  name: string | null;
  avatarKind?: string | null;
  avatarDoodle?: string | null;
  avatarUrlSmall?: string | null;
  accentColor?: string | null;
}

/** One card in the Changes pane. */
export interface SpreadsheetChangeEntry extends SpreadsheetChangeRecord {
  /** Server row id, or the client id while the upload is pending. */
  id: string;
  /** Epoch milliseconds. */
  createdAt: number;
  actor: SpreadsheetChangeActor;
  /** Made in this session and not yet confirmed by the server. */
  pending?: boolean;
}

export type SpreadsheetChangeFilter =
  | { scope: "workbook" }
  | { scope: "sheet"; sheetId: string }
  | {
      scope: "range";
      sheetId: string;
      r1: number;
      c1: number;
      r2: number;
      c2: number;
    };

/** A worksheet insert or delete of whole rows or columns. */
export interface SpreadsheetStructuralStep {
  sheetId: string;
  axis: "row" | "col";
  op: "insert" | "delete";
  at: number;
  count: number;
}

interface CellLike {
  value?: unknown;
  formula?: unknown;
}

interface SheetLike {
  id: string;
  name: string;
  chartSheet?: unknown;
  cells: Record<string, CellLike | undefined>;
}

/** A cell's content as the user typed it: its formula, else its value. */
export function cellContentText(cell: CellLike | undefined): string {
  if (!cell) return "";
  if (typeof cell.formula === "string" && cell.formula !== "") return cell.formula;
  const value = cell.value;
  if (value === undefined || value === null) return "";
  return typeof value === "string" ? value : String(value);
}

function parseKey(key: string): { r: number; c: number } | null {
  const separator = key.indexOf(":");
  if (separator < 1) return null;
  const r = Number(key.slice(0, separator));
  const c = Number(key.slice(separator + 1));
  if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) return null;
  return { r, c };
}

function clip(text: string): string {
  if (text.length <= CHANGE_LOG_MAX_VALUE_CHARS) return text;
  return `${text.slice(0, CHANGE_LOG_MAX_VALUE_CHARS - 1)}…`;
}

/**
 * The cells whose content differs between two cell maps, in reading order.
 * `keys`, when the commit already knows which cells it wrote, limits the
 * comparison to them; otherwise every key of either map is compared.
 */
/** The keys two cell maps differ at, when one is recorded as following from the other (an edit; Undo goes the other way). */
function changesBetween(before: object, after: object): readonly string[] | null {
  return cellMapChangesSince(after, before) ?? cellMapChangesSince(before, after);
}

export function diffCellContent(
  before: Record<string, CellLike | undefined>,
  after: Record<string, CellLike | undefined>,
  keys?: Iterable<string>,
): SpreadsheetChangeCell[] {
  if (before === after) return [];
  keys ??= changesBetween(before, after) ?? undefined;
  const changed: SpreadsheetChangeCell[] = [];
  const visit = (key: string) => {
    const previous = before[key];
    const next = after[key];
    if (previous === next) return;
    const beforeText = cellContentText(previous);
    const afterText = cellContentText(next);
    if (beforeText === afterText) return;
    const position = parseKey(key);
    if (!position) return;
    changed.push({ ...position, before: beforeText, after: afterText });
  };
  if (keys) {
    for (const key of new Set(keys)) visit(key);
  } else {
    for (const key in after) visit(key);
    for (const key in before) if (!(key in after)) visit(key);
  }
  changed.sort((left, right) => left.r - right.r || left.c - right.c);
  return changed;
}

/**
 * The rectangle around every cell that differs between two cell maps, its
 * format included: what Undo and Redo select, as Excel selects the cells a
 * step it takes back or brings back changed. Null when no cell differs.
 */
export function changedCellBounds(
  before: Record<string, unknown>,
  after: Record<string, unknown>,
): { r1: number; c1: number; r2: number; c2: number } | null {
  if (before === after) return null;
  let r1 = Infinity;
  let c1 = Infinity;
  let r2 = -Infinity;
  let c2 = -Infinity;
  const visit = (key: string) => {
    const previous = before[key];
    const next = after[key];
    if (previous === next || JSON.stringify(previous) === JSON.stringify(next)) return;
    const position = parseKey(key);
    if (!position) return;
    if (position.r < r1) r1 = position.r;
    if (position.c < c1) c1 = position.c;
    if (position.r > r2) r2 = position.r;
    if (position.c > c2) c2 = position.c;
  };
  const known = changesBetween(before, after);
  if (known) {
    for (const key of known) visit(key);
  } else {
    for (const key in after) visit(key);
    for (const key in before) if (!(key in after)) visit(key);
  }
  return r1 === Infinity ? null : { r1, c1, r2, c2 };
}

function defaultChangeId(): string {
  const random = globalThis.crypto?.randomUUID?.();
  if (random) return random.replace(/-/g, "");
  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 12)}`;
}

function structuralRecord(
  step: SpreadsheetStructuralStep,
  sheet: SheetLike,
  newId: () => string,
): SpreadsheetChangeRecord {
  const last = step.at + Math.max(1, step.count) - 1;
  const rows = step.axis === "row";
  return {
    clientChangeId: newId(),
    sheetId: sheet.id,
    sheetName: sheet.name,
    kind: `${step.op}_${rows ? "rows" : "columns"}` as SpreadsheetChangeKind,
    r1: rows ? step.at : 0,
    c1: rows ? 0 : step.at,
    r2: rows ? Math.min(last, MAX_PARSE_ROW) : MAX_PARSE_ROW,
    c2: rows ? MAX_PARSE_COL : Math.min(last, MAX_PARSE_COL),
    cellCount: 0,
    cells: [],
    truncated: false,
  };
}

/**
 * The change records one commit makes: one per worksheet whose cell content
 * changed, or the row/column step a structural command named for its sheet
 * (a row insert shifts every cell below it, which is one change, not
 * thousands). Sheets added or removed by the commit carry no cell change.
 */
export function buildChangeRecords(input: {
  before: readonly SheetLike[];
  after: readonly SheetLike[];
  kind?: "edit" | "undo" | "redo";
  structural?: SpreadsheetStructuralStep | null;
  /** Cells the commit wrote, per sheet id, when it knows them. */
  knownKeys?: ReadonlyMap<string, readonly string[]>;
  newId?: () => string;
}): SpreadsheetChangeRecord[] {
  const newId = input.newId ?? defaultChangeId;
  const kind = input.kind ?? "edit";
  const beforeById = new Map<string, SheetLike>();
  for (const sheet of input.before) {
    if (!beforeById.has(sheet.id)) beforeById.set(sheet.id, sheet);
  }
  const records: SpreadsheetChangeRecord[] = [];
  for (const sheet of input.after) {
    const previous = beforeById.get(sheet.id);
    if (!previous || sheet.chartSheet) continue;
    if (
      input.structural &&
      kind === "edit" &&
      input.structural.sheetId === sheet.id &&
      input.structural.count > 0
    ) {
      records.push(structuralRecord(input.structural, sheet, newId));
      continue;
    }
    if (previous.cells === sheet.cells) continue;
    const cells = diffCellContent(
      previous.cells,
      sheet.cells,
      input.knownKeys?.get(sheet.id),
    );
    if (cells.length === 0) continue;
    let r1 = Infinity;
    let c1 = Infinity;
    let r2 = -Infinity;
    let c2 = -Infinity;
    for (const cell of cells) {
      if (cell.r < r1) r1 = cell.r;
      if (cell.c < c1) c1 = cell.c;
      if (cell.r > r2) r2 = cell.r;
      if (cell.c > c2) c2 = cell.c;
    }
    const listed = cells.slice(0, CHANGE_LOG_MAX_CELLS).map((cell) => ({
      ...cell,
      before: clip(cell.before),
      after: clip(cell.after),
    }));
    records.push({
      clientChangeId: newId(),
      sheetId: sheet.id,
      sheetName: sheet.name,
      kind,
      r1,
      c1,
      r2,
      c2,
      cellCount: cells.length,
      cells: listed,
      truncated: cells.length > listed.length,
    });
  }
  return records;
}

const STRUCTURAL_KINDS = new Set<SpreadsheetChangeKind>([
  "insert_rows",
  "delete_rows",
  "insert_columns",
  "delete_columns",
]);

export function isStructuralChange(entry: { kind: SpreadsheetChangeKind }): boolean {
  return STRUCTURAL_KINDS.has(entry.kind);
}

/**
 * The entry as a filter shows it, or null when the filter hides it. A range
 * keeps only the bulk card's cells inside it (Excel narrows a bulk edit the
 * same way); a row/column change shows when its lines cross the range.
 */
export function narrowChange<T extends SpreadsheetChangeEntry>(
  entry: T,
  filter: SpreadsheetChangeFilter,
): T | null {
  if (filter.scope === "workbook") return entry;
  if (entry.sheetId !== filter.sheetId) return null;
  if (filter.scope === "sheet") return entry;
  if (
    entry.r1 > filter.r2 ||
    entry.r2 < filter.r1 ||
    entry.c1 > filter.c2 ||
    entry.c2 < filter.c1
  ) {
    return null;
  }
  if (isStructuralChange(entry)) return entry;
  const cells = entry.cells.filter(
    (cell) =>
      cell.r >= filter.r1 &&
      cell.r <= filter.r2 &&
      cell.c >= filter.c1 &&
      cell.c <= filter.c2,
  );
  if (cells.length === 0 && !entry.truncated) return null;
  if (cells.length === entry.cells.length) return entry;
  return {
    ...entry,
    cells,
    cellCount: entry.truncated ? entry.cellCount : cells.length,
  };
}

/**
 * The pane's list: server entries and this session's pending ones, each edit
 * once (a confirmed copy replaces its pending one), narrowed by the filter,
 * newest first.
 */
export function mergeChangeEntries(
  lists: ReadonlyArray<readonly SpreadsheetChangeEntry[]>,
  filter: SpreadsheetChangeFilter = { scope: "workbook" },
): SpreadsheetChangeEntry[] {
  const byClientId = new Map<string, SpreadsheetChangeEntry>();
  for (const list of lists) {
    for (const entry of list) {
      const known = byClientId.get(entry.clientChangeId);
      if (!known || (known.pending && !entry.pending)) {
        byClientId.set(entry.clientChangeId, entry);
      }
    }
  }
  const merged: SpreadsheetChangeEntry[] = [];
  for (const entry of byClientId.values()) {
    const shown = narrowChange(entry, filter);
    if (shown) merged.push(shown);
  }
  merged.sort(
    (left, right) =>
      right.createdAt - left.createdAt ||
      (right.id < left.id ? -1 : right.id > left.id ? 1 : 0),
  );
  return merged;
}

/** "A1", "A1:B4", "5:7" (rows) or "C:D" (columns). */
export function changeRangeLabel(entry: {
  kind: SpreadsheetChangeKind;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}): string {
  if (entry.kind === "insert_rows" || entry.kind === "delete_rows") {
    return `${entry.r1 + 1}:${entry.r2 + 1}`;
  }
  if (entry.kind === "insert_columns" || entry.kind === "delete_columns") {
    return `${columnToLetters(entry.c1)}:${columnToLetters(entry.c2)}`;
  }
  const start = `${columnToLetters(entry.c1)}${entry.r1 + 1}`;
  if (entry.r1 === entry.r2 && entry.c1 === entry.c2) return start;
  return `${start}:${columnToLetters(entry.c2)}${entry.r2 + 1}`;
}

/** A cell's address on its own: "B3". */
export function changeCellLabel(cell: { r: number; c: number }): string {
  return `${columnToLetters(cell.c)}${cell.r + 1}`;
}

/** A sheet name as a reference writes it: quoted when it needs quotes. */
export function quotedSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}[0-9]+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

/** "Sheet1!A1:B4", the location a card names. */
export function changeLocationLabel(entry: SpreadsheetChangeRecord): string {
  return `${quotedSheetName(entry.sheetName || entry.sheetId)}!${changeRangeLabel(entry)}`;
}

/**
 * The Range box: "B2:D9", "$B$2", "Sheet2!A1:C3" or "'My sheet'!A:A". A
 * reference without a sheet means the active sheet. Null when it does not
 * read as a range or names no sheet of this workbook.
 */
export function parseChangeRangeFilter(
  text: string,
  sheets: ReadonlyArray<{ id: string; name: string }>,
  activeSheetId: string,
): Extract<SpreadsheetChangeFilter, { scope: "range" }> | null {
  const trimmed = text.trim();
  if (!trimmed) return null;
  let sheetId = activeSheetId;
  let reference = trimmed;
  const bang = trimmed.lastIndexOf("!");
  if (bang >= 0) {
    let name = trimmed.slice(0, bang).trim();
    reference = trimmed.slice(bang + 1);
    if (name.startsWith("'") && name.endsWith("'") && name.length >= 2) {
      name = name.slice(1, -1).replace(/''/g, "'");
    }
    const folded = name.toLocaleLowerCase();
    const sheet = sheets.find((candidate) => candidate.name.toLocaleLowerCase() === folded);
    if (!sheet) return null;
    sheetId = sheet.id;
  }
  const area = parseA1Range(reference.replace(/\$/g, ""));
  if (!area) return null;
  return { scope: "range", sheetId, ...area };
}
