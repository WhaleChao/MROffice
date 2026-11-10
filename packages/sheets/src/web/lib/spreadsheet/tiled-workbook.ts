/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiFetchRaw, apiGet, apiPatch, isApiRequestError } from "@mrnote/core/web/lib/api.ts";
import { type ExcelDateSystemArg } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

/** One cell as a tile stores it (see the design note). */
export interface TileCell {
  /** The value as Excel holds it before formatting. */
  v?: string;
  /** n number, b logical, e error, d ISO date; absent: text. */
  k?: "n" | "b" | "e" | "d";
  /** The formula, with its "=". */
  f?: string;
  /** The cell's cellXfs index in the workbook's style table. */
  s?: number;
  /** An array formula's range (on its first cell). */
  a?: string;
  /** This editor must not change the cell. */
  x?: 1;
}

export type TiledReadOnlyKind = "merge" | "table" | "pivot" | "array";

export interface TiledSheet {
  id: string;
  name: string;
  hidden: boolean;
  rows: number;
  cols: number;
  cells: number;
  tileRows: number;
  defaultRowHeight: number;
  defaultColWidth: number;
  /** `[first, last, pixels]` runs of the file's column widths. */
  colWidths: [number, number, number][];
  freeze: { rows: number; cols: number } | null;
  dimension: string | null;
  /** Why the whole sheet cannot be edited here, if it cannot. */
  readOnly: string | null;
  readOnlyRanges: [number, number, number, number, TiledReadOnlyKind][];
}

export type TiledCellStyle = Record<string, unknown> & { numberFormat?: string };

export interface TiledManifest {
  schema: string;
  date1904: ExcelDateSystemArg;
  styles: (TiledCellStyle | null)[];
  styleCount: number;
  sheets: TiledSheet[];
  revision: number;
}

export interface TiledWorkbookOpening {
  page_id: string;
  title: string;
  manifest: TiledManifest;
  can_edit: boolean;
  read_only_reason: string | null;
  original: { filename?: string; size_bytes?: number } | null;
  cell_count: number;
  /** The sheet the workbook opens on (what it first shows is read before the
   * page's own read has answered); absent from an older server. */
  active_sheet_id?: string | null;
}

export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLUMNS = 16_384;
export const EXCEL_MAX_CELL_TEXT = 32_767;
export const EXCEL_MAX_FORMULA = 8_192;
/** The server's bounds (workbook_tiles.store). */
export const MAX_EDITS_PER_SAVE = 50_000;
export const MAX_TILES_PER_READ = 16;
/** Tiles kept in memory, all sheets together. */
export const TILE_CACHE_LIMIT = 96;
const TILE_FRAME_MAGIC = "MRTILES1";

function base(pageId: string): string {
  return `/api/v1/pages/${encodeURIComponent(pageId)}/workbook-tiles`;
}

/** The workbook's manifest, or `null` when the page is not a tiled workbook. */
export async function fetchTiledWorkbook(
  pageId: string,
  signal?: AbortSignal,
  options?: { beforeMountKnown?: boolean },
): Promise<TiledWorkbookOpening | null> {
  try {
    return await apiGet<TiledWorkbookOpening>(base(pageId), { signal }, options);
  } catch (error) {
    if (isApiRequestError(error) && error.status === 404) return null;
    throw error;
  }
}

async function inflate(bytes: Uint8Array): Promise<string> {
  const body = new Response(bytes as BodyInit).body;
  if (!body) return "";
  return new Response(body.pipeThrough(new DecompressionStream("deflate"))).text();
}

export interface FetchedTile {
  tile: number;
  revision: number;
  cells: Record<string, TileCell>;
}

/** Read a framed tile answer: magic, a big-endian length, a JSON index and
 * each tile's stored zlib payload. */
export async function parseTileFrame(buffer: ArrayBuffer): Promise<FetchedTile[]> {
  const bytes = new Uint8Array(buffer);
  const magic = new TextDecoder().decode(bytes.subarray(0, 8));
  if (magic !== TILE_FRAME_MAGIC) throw new Error("Not a tile answer");
  const length = new DataView(buffer, 8, 4).getUint32(0, false);
  const index = JSON.parse(new TextDecoder().decode(bytes.subarray(12, 12 + length))) as {
    tiles: { tile: number; revision: number; bytes: number }[];
  };
  let offset = 12 + length;
  const out: FetchedTile[] = [];
  for (const entry of index.tiles) {
    let cells: Record<string, TileCell> = {};
    if (entry.bytes > 0) {
      const text = await inflate(bytes.subarray(offset, offset + entry.bytes));
      const parsed = JSON.parse(text) as { cells?: Record<string, TileCell> };
      cells = parsed.cells ?? {};
      offset += entry.bytes;
    }
    out.push({ tile: entry.tile, revision: entry.revision, cells });
  }
  return out;
}

export async function fetchTiles(
  pageId: string,
  sheetId: string,
  tiles: readonly number[],
  signal?: AbortSignal,
): Promise<FetchedTile[]> {
  const query = new URLSearchParams({ sheet: sheetId, tiles: tiles.join(",") });
  const response = await apiFetchRaw(`${base(pageId)}/tiles?${query.toString()}`, { signal });
  return parseTileFrame(await response.arrayBuffer());
}

export interface CellEdit {
  sheet: string;
  row: number;
  col: number;
  cell: TileCell | null;
}

export interface SaveAnswer {
  revision: number;
  tiles: [string, number, number][];
  extents: Record<string, [number, number]>;
}

export function saveCells(pageId: string, edits: CellEdit[], baseRevision: number): Promise<SaveAnswer> {
  return apiPatch<SaveAnswer>(`${base(pageId)}/cells`, { edits, base_revision: baseRevision });
}

export interface ChangesAnswer {
  revision: number;
  tiles: [string, number, number][];
  reload?: boolean;
  extents: Record<string, [number, number]>;
}

export function fetchChanges(pageId: string, since: number, signal?: AbortSignal): Promise<ChangesAnswer> {
  return apiGet<ChangesAnswer>(`${base(pageId)}/changes?since=${since}`, { signal });
}

export interface FindMatch {
  sheet: string;
  row: number;
  col: number;
  text: string;
}

export function findCells(
  pageId: string,
  query: {
    q: string;
    sheet: string;
    afterRow: number;
    afterCol: number;
    matchCase?: boolean;
    wholeCell?: boolean;
    formulas?: boolean;
  },
  signal?: AbortSignal,
): Promise<{ matches: FindMatch[]; done: boolean }> {
  const params = new URLSearchParams({
    q: query.q,
    sheet: query.sheet,
    after_row: String(query.afterRow),
    after_col: String(query.afterCol),
  });
  if (query.matchCase) params.set("match_case", "true");
  if (query.wholeCell) params.set("whole_cell", "true");
  if (query.formulas) params.set("formulas", "true");
  return apiGet(`${base(pageId)}/find?${params.toString()}`, { signal });
}

// ── references ──────────────────────────────────────────────────────────────

export function columnName(index: number): string {
  let value = index + 1;
  let out = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

export function columnIndex(letters: string): number {
  let value = 0;
  for (const char of letters.toUpperCase()) value = value * 26 + (char.charCodeAt(0) - 64);
  return value - 1;
}

export function cellName(row: number, col: number): string {
  return `${columnName(col)}${row + 1}`;
}

/** `{row, col}` (zero-based) of `B12`, or null. */
export function parseCellName(text: string): { row: number; col: number } | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(text.trim());
  if (!match) return null;
  const row = Number(match[2]) - 1;
  const col = columnIndex(match[1]);
  if (row < 0 || row >= EXCEL_MAX_ROWS || col < 0 || col >= EXCEL_MAX_COLUMNS) return null;
  return { row, col };
}

export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/** `A1:C3`, `B2` (a range of one) or null. */
export function parseRangeName(text: string): CellRange | null {
  const [first, second] = text.split(":");
  const start = parseCellName(first ?? "");
  const end = second === undefined ? start : parseCellName(second);
  if (!start || !end) return null;
  return {
    top: Math.min(start.row, end.row),
    left: Math.min(start.col, end.col),
    bottom: Math.max(start.row, end.row),
    right: Math.max(start.col, end.col),
  };
}

/** The read-only range holding a cell, if any: a merged area's first cell is
 * editable, everything else in one is not. */
export function readOnlyRangeAt(
  sheet: TiledSheet,
  row: number,
  col: number,
): TiledReadOnlyKind | null {
  for (const [top, left, bottom, right, kind] of sheet.readOnlyRanges) {
    if (row < top || row > bottom || col < left || col > right) continue;
    if (kind === "merge" && row === top && col === left) continue;
    return kind;
  }
  return null;
}

// ── the workbook in this browser ────────────────────────────────────────────

/** `error`: not saved yet, retrying; `refused`: the server refused the
 * cells for good (the cells show what is stored again). */
export type TiledSaveStatus = "saved" | "unsaved" | "saving" | "error" | "refused";

interface CachedTile {
  revision: number;
  cells: Record<string, TileCell>;
}

interface UndoEntry {
  edits: { sheet: string; row: number; col: number; before: TileCell | null; after: TileCell | null }[];
}

function editKey(sheet: string, row: number, col: number): string {
  return `${sheet}|${row}:${col}`;
}

function tileKey(sheet: string, tile: number): string {
  return `${sheet}#${tile}`;
}

export interface TiledWorkbookStoreOptions {
  canEdit: boolean;
  /** Milliseconds a typed cell waits before it is saved. */
  saveDelayMs?: number;
  fetchTilesImpl?: typeof fetchTiles;
  saveCellsImpl?: typeof saveCells;
  fetchChangesImpl?: typeof fetchChanges;
}

/**
 * The tiles in view, the cells typed and not yet saved, Undo and Redo, and
 * the save queue. Framework-free: the window subscribes and re-renders on
 * `version`.
 */
export class TiledWorkbookStore {
  readonly pageId: string;
  manifest: TiledManifest;
  readonly canEdit: boolean;
  version = 0;
  status: TiledSaveStatus = "saved";
  lastError: string | null = null;
  knownRevision: number;
  private readonly sheetsById = new Map<string, TiledSheet>();
  private readonly tiles = new Map<string, CachedTile>();
  private readonly loading = new Map<string, Promise<void>>();
  /** Typed and not sent yet / sent and not answered yet. */
  private readonly queued = new Map<string, CellEdit>();
  private readonly sending = new Map<string, CellEdit>();
  private readonly undoStack: UndoEntry[] = [];
  private readonly redoStack: UndoEntry[] = [];
  private readonly listeners = new Set<() => void>();
  private saveTimer: ReturnType<typeof setTimeout> | null = null;
  private saving: Promise<void> | null = null;
  private retryDelayMs = 2000;
  private disposed = false;
  private readonly options: Required<Omit<TiledWorkbookStoreOptions, "canEdit">>;

  constructor(pageId: string, manifest: TiledManifest, options: TiledWorkbookStoreOptions) {
    this.pageId = pageId;
    this.manifest = manifest;
    this.canEdit = options.canEdit;
    this.knownRevision = manifest.revision;
    this.options = {
      saveDelayMs: options.saveDelayMs ?? 400,
      fetchTilesImpl: options.fetchTilesImpl ?? fetchTiles,
      saveCellsImpl: options.saveCellsImpl ?? saveCells,
      fetchChangesImpl: options.fetchChangesImpl ?? fetchChanges,
    };
    for (const sheet of manifest.sheets) this.sheetsById.set(sheet.id, sheet);
  }

  dispose(): void {
    this.disposed = true;
    if (this.saveTimer) clearTimeout(this.saveTimer);
    this.listeners.clear();
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getVersion = (): number => this.version;

  private changed(): void {
    this.version += 1;
    for (const listener of this.listeners) listener();
  }

  sheet(id: string): TiledSheet | undefined {
    return this.sheetsById.get(id);
  }

  style(index: number | undefined): TiledCellStyle | null {
    if (index === undefined || index <= 0) return this.manifest.styles[0] ?? null;
    return this.manifest.styles[index] ?? null;
  }

  saveStatus(): TiledSaveStatus {
    return this.status;
  }

  get pendingCount(): number {
    return this.queued.size + this.sending.size;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /** The cell as it is here now: a typed edit, else its tile; `undefined`
   * when the tile is not loaded yet, `null` for an empty cell. */
  cell(sheetId: string, row: number, col: number): TileCell | null | undefined {
    const key = editKey(sheetId, row, col);
    const typed = this.queued.get(key) ?? this.sending.get(key);
    if (typed) return typed.cell;
    const sheet = this.sheetsById.get(sheetId);
    if (!sheet) return null;
    const cached = this.tiles.get(tileKey(sheetId, Math.floor(row / sheet.tileRows)));
    if (!cached) return undefined;
    return cached.cells[`${row}:${col}`] ?? null;
  }

  isLoaded(sheetId: string, row: number): boolean {
    const sheet = this.sheetsById.get(sheetId);
    if (!sheet) return true;
    return this.tiles.has(tileKey(sheetId, Math.floor(row / sheet.tileRows)));
  }

  /** Load the tiles of rows `[firstRow, lastRow]` (bounded), returning when
   * they are in. */
  async loadRows(sheetId: string, firstRow: number, lastRow: number): Promise<void> {
    const sheet = this.sheetsById.get(sheetId);
    if (!sheet) return;
    const first = Math.max(0, Math.floor(firstRow / sheet.tileRows));
    const last = Math.min(
      Math.floor((EXCEL_MAX_ROWS - 1) / sheet.tileRows),
      Math.floor(Math.max(firstRow, lastRow) / sheet.tileRows),
    );
    const waiting: Promise<void>[] = [];
    const missing: number[] = [];
    for (let tile = first; tile <= last; tile += 1) {
      const key = tileKey(sheetId, tile);
      const cached = this.tiles.get(key);
      if (cached) {
        // Most recently used last.
        this.tiles.delete(key);
        this.tiles.set(key, cached);
        continue;
      }
      const inflight = this.loading.get(key);
      if (inflight) waiting.push(inflight);
      else missing.push(tile);
    }
    for (let start = 0; start < missing.length; start += MAX_TILES_PER_READ) {
      const batch = missing.slice(start, start + MAX_TILES_PER_READ);
      const request = this.options
        .fetchTilesImpl(this.pageId, sheetId, batch)
        .then((fetched) => {
          if (this.disposed) return;
          for (const tile of fetched) this.storeTile(sheetId, tile.tile, tile.revision, tile.cells);
          this.changed();
        })
        .finally(() => {
          for (const tile of batch) this.loading.delete(tileKey(sheetId, tile));
        });
      for (const tile of batch) this.loading.set(tileKey(sheetId, tile), request);
      waiting.push(request);
    }
    await Promise.all(waiting);
  }

  private storeTile(sheetId: string, tile: number, revision: number, cells: Record<string, TileCell>): void {
    const key = tileKey(sheetId, tile);
    this.tiles.delete(key);
    this.tiles.set(key, { revision, cells });
    while (this.tiles.size > TILE_CACHE_LIMIT) {
      const oldest = this.tiles.keys().next().value as string | undefined;
      if (oldest === undefined) break;
      this.tiles.delete(oldest);
    }
  }

  /** Change cells: shown at once, saved shortly after, one Undo step. */
  apply(edits: CellEdit[], { record = true }: { record?: boolean } = {}): void {
    if (!this.canEdit || edits.length === 0) return;
    const entry: UndoEntry = { edits: [] };
    for (const edit of edits) {
      const before = this.cell(edit.sheet, edit.row, edit.col) ?? null;
      entry.edits.push({ sheet: edit.sheet, row: edit.row, col: edit.col, before, after: edit.cell });
      this.queued.set(editKey(edit.sheet, edit.row, edit.col), edit);
      const sheet = this.sheetsById.get(edit.sheet);
      if (sheet && edit.cell) {
        sheet.rows = Math.max(sheet.rows, edit.row + 1);
        sheet.cols = Math.max(sheet.cols, edit.col + 1);
      }
    }
    if (record) {
      this.undoStack.push(entry);
      if (this.undoStack.length > 200) this.undoStack.shift();
      this.redoStack.length = 0;
    }
    this.status = "unsaved";
    this.scheduleSave();
    this.changed();
  }

  undo(): UndoEntry | null {
    const entry = this.undoStack.pop();
    if (!entry) return null;
    this.apply(
      entry.edits.map((edit) => ({ sheet: edit.sheet, row: edit.row, col: edit.col, cell: edit.before })),
      { record: false },
    );
    this.redoStack.push(entry);
    return entry;
  }

  redo(): UndoEntry | null {
    const entry = this.redoStack.pop();
    if (!entry) return null;
    this.apply(
      entry.edits.map((edit) => ({ sheet: edit.sheet, row: edit.row, col: edit.col, cell: edit.after })),
      { record: false },
    );
    this.undoStack.push(entry);
    return entry;
  }

  private scheduleSave(delay = this.options.saveDelayMs): void {
    if (this.saveTimer) clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => {
      this.saveTimer = null;
      void this.flush();
    }, delay);
  }

  /** Send what was typed; resolves once nothing typed is left unsent. */
  async flush(): Promise<void> {
    if (this.saving) {
      await this.saving;
      if (this.queued.size === 0) return;
    }
    if (this.queued.size === 0 || this.disposed) return;
    const batch = [...this.queued.entries()].slice(0, MAX_EDITS_PER_SAVE);
    for (const [key, edit] of batch) {
      this.queued.delete(key);
      this.sending.set(key, edit);
    }
    this.status = "saving";
    this.changed();
    const run = (async () => {
      try {
        const answer = await this.options.saveCellsImpl(
          this.pageId,
          batch.map(([, edit]) => edit),
          this.knownRevision,
        );
        const othersSaved = answer.revision !== this.knownRevision + 1;
        for (const [key, edit] of batch) {
          if (this.sending.get(key) === edit) this.sending.delete(key);
          const sheet = this.sheetsById.get(edit.sheet);
          if (!sheet) continue;
          const cached = this.tiles.get(tileKey(edit.sheet, Math.floor(edit.row / sheet.tileRows)));
          if (cached) {
            if (edit.cell) cached.cells[`${edit.row}:${edit.col}`] = edit.cell;
            else delete cached.cells[`${edit.row}:${edit.col}`];
          }
        }
        for (const [sheetId, tile, revision] of answer.tiles) {
          const cached = this.tiles.get(tileKey(sheetId, tile));
          if (cached && !othersSaved) cached.revision = revision;
        }
        this.growExtents(answer.extents);
        if (othersSaved) await this.poll();
        else this.knownRevision = answer.revision;
        this.retryDelayMs = 2000;
        this.lastError = null;
        this.status = this.queued.size + this.sending.size > 0 ? "unsaved" : "saved";
      } catch (error) {
        const permanent = isApiRequestError(error) && error.status !== undefined && error.status >= 400 && error.status < 500 && error.status !== 408 && error.status !== 429;
        for (const [key, edit] of batch) {
          if (this.sending.get(key) !== edit) continue;
          this.sending.delete(key);
          // A newer typing of the same cell stays; this one waits again
          // (unless the server refused it for good).
          if (!permanent && !this.queued.has(key)) this.queued.set(key, edit);
        }
        this.lastError = error instanceof Error ? error.message : String(error);
        this.status = permanent ? "refused" : "error";
        if (permanent) {
          // Show what is stored again where the refused cells were.
          for (const [, edit] of batch) {
            const sheet = this.sheetsById.get(edit.sheet);
            if (sheet) this.tiles.delete(tileKey(edit.sheet, Math.floor(edit.row / sheet.tileRows)));
          }
        } else {
          this.scheduleSave(this.retryDelayMs);
          this.retryDelayMs = Math.min(30_000, this.retryDelayMs * 2);
        }
      } finally {
        this.saving = null;
        this.changed();
      }
    })();
    this.saving = run;
    await run;
    // Read through a call: the save above changed it.
    const settled = this.saveStatus();
    if (this.queued.size > 0 && settled !== "error" && settled !== "refused") await this.flush();
  }

  private growExtents(extents: Record<string, [number, number]> | undefined): void {
    for (const [sheetId, extent] of Object.entries(extents ?? {})) {
      const sheet = this.sheetsById.get(sheetId);
      if (!sheet || !Array.isArray(extent)) continue;
      sheet.rows = Math.max(sheet.rows, extent[0]);
      sheet.cols = Math.max(sheet.cols, extent[1]);
    }
  }

  /** Learn what others saved: changed tiles in the cache are read again. */
  async poll(signal?: AbortSignal): Promise<void> {
    const answer = await this.options.fetchChangesImpl(this.pageId, this.knownRevision, signal);
    if (this.disposed) return;
    this.growExtents(answer.extents);
    const stale: string[] = [];
    if (answer.reload) {
      stale.push(...this.tiles.keys());
    } else {
      for (const [sheetId, tile, revision] of answer.tiles) {
        const key = tileKey(sheetId, tile);
        const cached = this.tiles.get(key);
        if (cached && cached.revision < revision) stale.push(key);
      }
    }
    for (const key of stale) this.tiles.delete(key);
    this.knownRevision = Math.max(this.knownRevision, answer.revision);
    if (stale.length > 0) this.changed();
  }
}
