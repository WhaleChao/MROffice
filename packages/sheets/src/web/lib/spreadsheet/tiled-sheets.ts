/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A large workbook in the ordinary spreadsheet editor, its plain cells read
 * as the grid needs them (a "structured tiled workbook").
 *
 * The page body of such a workbook is its structure, as the ordinary import
 * reads it (drawings, merges, widths, formats, names, print settings) and the
 * cells the body holds itself (*body-owned*: a merged area's, a comment's or
 * a link's cell, a chart's source, ...). Every other cell is *tile-owned*:
 * stored in tiles beside the page, one per sheet and block of rows
 * (apps/api/app/services/workbook_tiles). The editor's model is unchanged: a
 * sheet's `cells` map. This adapter keeps in that map the body's cells plus
 * the tiles around what is on screen (the viewport and a margin), reads more
 * as the grid scrolls and lets far ones go, and turns a save of the model
 * into the body's own save (its cells filtered to the body-owned ones) and a
 * cell patch for the tiles.
 *
 * A tile's cells come in the editor's own form (the server converts them as
 * the ordinary import does); a cell keeps the style object of every other
 * cell with the same format (lib/spreadsheet/cell-style-table).
 */

import { SOURCE_XF_STYLE_KEY } from "./style-source-xf";
import { BLANK_CELL_FORMATS_KEY, blankCellFormatRanges } from "@mrnote/core/web/lib/spreadsheet/blank-cell-formats.ts";
import { apiGet, apiPatch, apiPost, isApiRequestError, withPreparedJsonBytes } from "@mrnote/core/web/lib/api.ts";
import { cellMapRowPermutation, nearestKnownCellMap, recordCellMapEdit, recordCellMapEditOver } from "@mrnote/core/web/lib/spreadsheet/cell-map-lineage.ts";
import { relocateSortedCell } from "./sort";
import { createCellStyleInterner } from "@mrnote/core/web/lib/spreadsheet/cell-style-table.ts";
import type { SpreadsheetCell } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import type { TiledWorkbookOpening } from "./tiled-workbook";
import { lineStyleAt, type LineStyleMap } from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";
import { matchAllBehind } from "@mrnote/core/web/lib/regexp-behind.ts";

/** The tile manifest's sheet entry as a structured workbook uses it. */
export interface TiledSheetEntry {
  id: string;
  name: string;
  rows: number;
  cols: number;
  cells: number;
  tileRows: number;
}

export interface StructuredTiledOpening extends TiledWorkbookOpening {
  structured: true;
}

export function isStructuredTiledOpening(opening: unknown): opening is StructuredTiledOpening {
  return Boolean(opening && typeof opening === "object" && (opening as { structured?: unknown }).structured === true);
}

/** How many of a workbook's other sheets are read at once. */
const OTHER_SHEET_READERS = 4;
/** How often what those reads brought is handed to the window. */
const OTHER_SHEETS_APPLY_MS = 800;

/** The editor fields a tile cell holds (the server's `CORE_FIELDS`). */
const CORE_FIELDS = new Set([
  "value", "kind", "stringValue", "formula", "style", "locked", "formulaHidden", "display", "chartValueCache",
]);
/** Fields a tile cell is read with that say nothing a save must keep. */
const CACHE_FIELDS = new Set(["display", "chartValueCache"]);
/** Tiles one read asks for (the server's `MAX_TILES_PER_READ`). */
const TILES_PER_READ = 32;
/** Cells one save sends (the server's `MAX_EDITS_PER_SAVE`). */
const EDITS_PER_SAVE = 50_000;
/** Tiles kept beyond the ones around the viewport, most recent first. */
const SPARE_TILES = 8;
/** Cells the tiles kept in memory may hold before far ones are let go. */
export const RESIDENT_CELL_BUDGET = 250_000;
/** A sheet holding at most this many cells is read whole when it is shown:
 * every command then sees all of it, as in an ordinary workbook page (the
 * whole page's own per-sheet ceiling). A larger sheet is read around what is
 * on screen, and the commands that need all of it are refused there. */
export const WHOLE_SHEET_CELLS = 200_000;
/** A larger sheet is read whole once a command asks for all of it (a sort, a
 * filter, cells moved), up to this: ONLYOFFICE holds such a sheet in its tab
 * too. Past it those commands stay refused. */
export const WHOLE_ON_REQUEST_CELLS = 3_000_000;
/** A workbook whose tiles hold at most this many cells in all is read whole
 * (every sheet, shown or not) and kept: every command, the Copilot's, a
 * sheet's rename or deletion and a formula reading another sheet then see
 * all of it, as on an ordinary workbook page. Within the resident budget. */
export const WHOLE_WORKBOOK_CELLS = 200_000;







let wholeCellsAsked: number | null | undefined;
function askedWholeCells(): number | null {
  if (wholeCellsAsked !== undefined) return wholeCellsAsked;
  wholeCellsAsked = null;
  if (typeof window !== "undefined") {
    let asked = /[?&]tiledWholeCells=(\d+)/.exec(window.location.search)?.[1] ?? null;
    if (asked === null) {
      try {
        asked = window.localStorage.getItem("mrai.tiledWholeCells");
      } catch {
        asked = null;
      }
    }
    if (asked !== null && /^\d+$/.test(asked)) wholeCellsAsked = Number(asked);
  }
  return wholeCellsAsked;
}

type CellMap = Record<string, SpreadsheetCell>;
interface SheetLike {
  id: string;
  cells: CellMap;
  [key: string]: unknown;
}

export interface TiledCellEdit {
  sheet: string;
  row: number;
  col: number;
  cell: Record<string, unknown> | null;
}

interface ResidentTile {
  index: number;
  revision: number;
  /** When it was read: a model older than this has not seen it. */
  generation: number;
  /** Each cell as last read or saved (the save compares against it). */
  baseline: Map<string, SpreadsheetCell>;
  /** Whether its cells are in the model yet. */
  applied: boolean;
  usedAt: number;
  /** Read for the first time in this sheet's life here: the model holds
   * none of its cells but the body's, loose ones and edits (all kept), so
   * putting it in needs no sweep of the model for cells it no longer has. */
  firstRead?: boolean;
  /** Read again while the model showed an earlier reading of it: that
   * reading, until this one is in the model. What the model holds there
   * unlike it is the window's own, typed since the last collect (night
   * coedit R421-X12: a cell typed as the tile arrived was swept away and
   * never saved); collect compares against it meanwhile. */
  shown?: Map<string, SpreadsheetCell>;
}

interface SheetState {
  entry: TiledSheetEntry;
  tiles: Map<number, ResidentTile>;
  /** Positions the page body holds. */
  bodyKeys: Set<string>;
  /** Tile-owned cells the model holds outside any resident tile (a paste
   * far below), as last saved. */
  loose: Map<string, SpreadsheetCell>;
  /** Every tile ever read for this sheet (``ResidentTile.firstRead``). */
  everRead?: Set<number>;
}

/** A sort of a tiled sheet's rows waiting to be saved as its order
 * (`workbook-tiles/row-sort`), with what the tiles held there before it (put
 * back, and the range saved cell by cell, when the server refuses it). */
interface QueuedRowSort {
  sheet: SheetLike;
  r1: number;
  c1: number;
  c2: number;
  order: readonly number[];
  previous: Map<string, SpreadsheetCell>;
  /** The positions the body held before the sort moved them. */
  previousBodyKeys: Set<string>;
}

export interface EditorTilesAnswer {
  revision: number;
  sheet: string;
  styles: Record<string, unknown>[];
  tiles: { tile: number; revision: number; cells: Record<string, Record<string, unknown>> }[];
}

interface SaveAnswer {
  revision: number;
  /** Every revision the save made (its cells', then the formulas it
   * calculated). */
  revisions?: number[];
  tiles: [string, number, number][];
  extents: Record<string, [number, number]>;
  /** The values of the formulas the saved cells feed, calculated over the
   * stored cells (`recalc`), by sheet and `row:col`; null when the server
   * could not answer. */
  recalculated?: Record<string, Record<string, unknown>> | null;
}

interface ChangesAnswer {
  revision: number;
  tiles: [string, number, number][];
  reload?: boolean;
}

function base(pageId: string): string {
  return `/api/v1/pages/${encodeURIComponent(pageId)}/workbook-tiles`;
}

/** The row of a `row:col` key, without the pair `parseKey` allocates. */
function rowOfKey(key: string): number {
  return parseInt(key, 10);
}

function parseKey(key: string): [number, number] {
  const separator = key.indexOf(":");
  return [Number(key.slice(0, separator)), Number(key.slice(separator + 1))];
}

/** What a save compares: the cell without the caches a save stamps on a
 * value's display. A formula's cache is its value, and counts -- as Excel
 * keeps a number, to 15 significant digits: the browser's 90030.59999999999
 * is the file's 90030.6, not an edit (rewriting every such formula's tile
 * on the first save made each tile dirty for nothing). */
function comparable(cell: SpreadsheetCell | undefined): string {
  if (!cell) return "";
  let plain: Record<string, unknown> = cell;
  const formula = typeof cell.formula === "string" && cell.formula.trim().startsWith("=");
  if (!formula && ("display" in cell || "chartValueCache" in cell)) {
    plain = {};
    for (const [field, value] of Object.entries(cell)) if (!CACHE_FIELDS.has(field)) plain[field] = value;
  }
  const cache = formula ? (cell.chartValueCache as { value?: unknown } | undefined) : undefined;
  if (cache && typeof cache.value === "number" && Number.isFinite(cache.value)) {
    // The shown text follows from the value (and the format, compared too).
    plain = {};
    for (const [field, value] of Object.entries(cell)) if (!CACHE_FIELDS.has(field)) plain[field] = value;
    plain.cachedNumber = Number(cache.value.toPrecision(15));
  }
  return JSON.stringify(plain);
}

/** A cell as typed: without any cache a calculation stamps on it. */
/** Whether two states of a body cell differ only in what a calculation stamps. */
function sameBesidesResult(before: SpreadsheetCell | undefined, cell: SpreadsheetCell | undefined): boolean {
  if (!before || !cell) return false;
  const inputs = (value: unknown) => JSON.stringify(value, (key, item) => (key === "display" || key === "chartValueCache" ? undefined : item));
  return inputs(before) === inputs(cell);
}

function inputOf(cell: SpreadsheetCell | undefined): string {
  if (!cell) return "";
  const plain: Record<string, unknown> = {};
  for (const [field, value] of Object.entries(cell)) if (!CACHE_FIELDS.has(field)) plain[field] = value;
  return JSON.stringify(plain);
}

/** Whether a model cell is only what its row or column format puts at a
 * position with no cell of its own (`materializeLineStyles`): an empty text
 * cell carrying exactly the line's look. Such a cell is not the body's and
 * not an edit; a tile cell read there takes its place. */
function isLineFormatBlank(sheet: SheetLike, key: string, cell: SpreadsheetCell | undefined): boolean {
  if (!cell || cell.value !== "" || cell.formula !== undefined) return false;
  for (const field of Object.keys(cell)) {
    if (field !== "value" && field !== "kind" && field !== "style" && field !== "locked") return false;
  }
  if (cell.kind !== undefined && cell.kind !== "text") return false;
  const [row, col] = parseKey(key);
  const line = lineStyleAt(sheet.rowStyles as LineStyleMap | undefined, sheet.columnStyles as LineStyleMap | undefined, row, col);
  if (!line) return false;
  if ((cell.locked === false) !== (line.locked === false)) return false;
  return styleIsLineLook(cell.style, line);
}

/** A line's blanks share one style object: each pair is compared once, not
 * once a blank over every collect (a fifth of a second of an open). */
const lineLooks = new WeakMap<object, WeakMap<object, boolean>>();

function styleIsLineLook(style: SpreadsheetCell["style"], line: NonNullable<ReturnType<typeof lineStyleAt>>): boolean {
  const lineLook = () => JSON.stringify({ ...line, locked: undefined });
  if (!style) return lineLook() === "{}";
  let known = lineLooks.get(line);
  if (!known) lineLooks.set(line, known = new WeakMap());
  let same = known.get(style);
  if (same === undefined) known.set(style, same = JSON.stringify(style) === lineLook());
  return same;
}

/** Each sheet's positions the stored page body holds a cell for: its own
 * `cells` and its formatted-blank ranges (a merge's covered cells, Excel's
 * styled empties), not the line formats the editor fills into the grid.
 *
 * A range left out of this read as cells nobody had saved: opening the
 * workbook wrote them to the tiles, and a window closed meanwhile asked to
 * save changes the person never made (AI70H H49R B09, 84 cells of the ONS CPI
 * tables). */
export function storedBodyCellKeys(content: unknown): Map<string, Set<string>> | null {
  const sheets = content && typeof content === "object" ? (content as { sheets?: unknown }).sheets : null;
  if (!Array.isArray(sheets)) return null;
  const out = new Map<string, Set<string>>();
  for (const sheet of sheets) {
    if (!sheet || typeof sheet !== "object") continue;
    const { id, cells, rowCount, columnCount, [BLANK_CELL_FORMATS_KEY]: groups } = sheet as {
      id?: unknown; cells?: unknown; rowCount?: unknown; columnCount?: unknown; [BLANK_CELL_FORMATS_KEY]?: unknown;
    };
    if (typeof id !== "string") continue;
    const keys = new Set(cells && typeof cells === "object" ? Object.keys(cells) : []);
    for (const { rect } of blankCellFormatRanges(groups, Number(rowCount) || 0, Number(columnCount) || 0)) {
      for (let row = rect.r1; row <= rect.r2; row += 1) {
        for (let col = rect.c1; col <= rect.c2; col += 1) keys.add(`${row}:${col}`);
      }
    }
    out.set(id, keys);
  }
  return out;
}

/** A cell that only states a format: no value, no formula, nothing else. */
function isFormatOnlyCell(cell: SpreadsheetCell | undefined): boolean {
  if (!cell || cell.value !== "" || cell.formula !== undefined) return false;
  return Object.keys(cell).every((field) => ["value", "kind", "style", "locked"].includes(field));
}

function sameCell(a: SpreadsheetCell | undefined, b: SpreadsheetCell | undefined): boolean {
  return a === b || comparable(a) === comparable(b);
}

/** Whether a cell can be a tile's: only fields a tile holds, and any rule
 * list (ignoredErrors) still the one it was read with. */
function tileOwnable(cell: SpreadsheetCell, read: SpreadsheetCell | undefined): boolean {
  for (const field of Object.keys(cell)) {
    if (CORE_FIELDS.has(field)) continue;
    if (field === "ignoredErrors" && read && JSON.stringify(cell.ignoredErrors) === JSON.stringify(read.ignoredErrors)) continue;
    return false;
  }
  return true;
}

/** A cell as a tile save sends it: no derived rule list; caches only where
 * they carry a formula's value. */
function tilePayload(cell: SpreadsheetCell): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const formula = typeof cell.formula === "string" && cell.formula.trim().startsWith("=");
  for (const [field, value] of Object.entries(cell)) {
    if (!CORE_FIELDS.has(field)) continue;
    if (CACHE_FIELDS.has(field) && !formula) continue;
    out[field] = value;
  }
  return out;
}

export class TiledSheets {
  /** The sheets past `WHOLE_SHEET_CELLS` a command asked to read whole. */
  private readonly wholeOnRequest = new Set<string>();
  /** Rows a derived result is computed from (a PivotTable report reading its
   * source's live cells), by sheet: their tiles are kept once read
   * (`keepRows`), so the report is not dropped back to partial rows. */
  private keptRows = new Map<string, ReadonlyArray<readonly [number, number]>>();
  readonly pageId: string;
  revision: number;
  /** Advances when tiles are read: a model is reconciled against what it has not seen. */
  generation = 0;
  private readonly sheets = new Map<string, SheetState>();
  private readonly intern = createCellStyleInterner();
  private readonly reading = new Map<string, Promise<void>>();
  private readonly seen = new WeakMap<object, number>();
  private clock = 0;
  private pending: TiledCellEdit[] = [];
  /** Body cells changed since the last save (any sheet's): the server reads
   * them when it calculates the formulas the save feeds. */
  private touched: TiledCellEdit[] = [];
  /** Each sheet's body cells as last saved or loaded. */
  private readonly bodyBaseline = new Map<string, Map<string, SpreadsheetCell>>();
  /** Sheets whose body has a calculated result stamped since it was loaded. */
  private readonly bodyResultSheets = new Set<string>();
  /** Each sheet's extent as its body states it (`bodySheets`). */
  private readonly bodyExtents = new Map<string, { rowCount: unknown; columnCount: unknown }>();
  /** While a line edit shows in the window ahead of the server (`holdCollect`). */
  private collectHeld = false;
  /** Sheets whose body has more than a calculated result changed since it was
   * loaded (`bodyEdited`). */
  private readonly bodyEditedSheets = new Set<string>();
  /** Counts the changes to what `collect` compares a sheet's cells with (the
   * tiles read, applied or let go, the body taken, a sheet retired) made
   * anywhere but by `collect` itself. */
  private tileState = 0;
  /** The cell maps `collect` last brought the tiles level with, by sheet: a
   * later map of the same lineage is compared at the keys changed since,
   * not at all 730,000 of a large sheet's cells on each save. */
  private readonly collected = new WeakMap<object, { sheetId: string; tileState: number; rowStyles: unknown; columnStyles: unknown }>();
  private inFlight: Promise<boolean> | null = null;
  private readonly ownRevisions = new Set<number>();
  /** Tiles this window saved into while another member's save landed between
   * its revision and the answer's: read again at the next poll, whatever
   * revision they now carry (the answer's revision is this window's own, and
   * the other member's cells in them would be passed over). */
  private readonly rereadOwnTiles = new Map<string, Set<number>>();
  private readonly listeners = new Set<() => void>();
  private failure: string | null = null;
  /** The refused save's HTTP status (0: no answer) and error code. */
  private failureStatus: number | null = null;
  private failureCode: string | null = null;
  /** Counts refused flushes (a retry that failed the same way is new). */
  private failures = 0;
  /** Formulas typed here whose values the server is asked (the browser's
   * calculation does not run over a sheet it holds only part of). */
  private readonly unanswered = new Map<string, Map<string, string>>();
  private readonly answered = new Map<string, Map<string, { formula: string | null; value: unknown }>>();

  constructor(opening: StructuredTiledOpening) {
    this.pageId = opening.page_id;
    this.revision = opening.manifest.revision ?? 0;
    for (const entry of opening.manifest.sheets as unknown as TiledSheetEntry[]) {
      this.sheets.set(entry.id, { entry, tiles: new Map(), bodyKeys: new Set(), loose: new Map() });
    }
  }

  /** Called when tiles arrive or saves settle (the window re-renders). */
  subscribe(listener: () => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit(): void {
    for (const listener of this.listeners) listener();
  }

  isTiledSheet(sheetId: string): boolean {
    return this.sheets.has(sheetId);
  }

  entry(sheetId: string): TiledSheetEntry | undefined {
    return this.sheets.get(sheetId)?.entry;
  }

  get saveFailure(): string | null {
    return this.failure;
  }

  /** The last refused save, as its message, HTTP status and code. */
  get saveFailureDetail(): { message: string; status: number | null; code: string | null; attempt: number } | null {
    return this.failure === null
      ? null
      : { message: this.failure, status: this.failureStatus, code: this.failureCode, attempt: this.failures };
  }

  /** The edits not sent yet (a recovery copy keeps them), or null. */
  pendingSnapshot(): { edits: TiledCellEdit[]; touched: TiledCellEdit[]; revision: number } | null {
    if (!this.pending.length && !this.touched.length) return null;
    return { edits: [...this.pending], touched: [...this.touched], revision: this.revision };
  }

  /** A save of the pending edits is on its way. */
  get saving(): boolean {
    return this.inFlight !== null;
  }

  get hasPendingEdits(): boolean {
    return this.pending.length > 0 || this.touched.length > 0 || this.rowSorts.length > 0 || this.inFlight !== null;
  }

  // ── the body ───────────────────────────────────────────────────────────

  /** A loaded body's sheets: the positions the stored body holds a cell
   * for (`stored`, from `storedBodyCellKeys`) are the body-owned ones. A
   * blank the editor filled in for a row or column format, or for a
   * formatted-blank range, is not: the tile's cell there is the cell (taking
   * every such position as the body's hid whole formatted columns of data,
   * and the next body save wrote them back as blanks over the tiles). */
  acceptBody(sheets: readonly SheetLike[], stored?: Map<string, Set<string>> | null): void {
    this.tileState += 1;
    for (const sheet of sheets) {
      this.bodyEditedSheets.delete(sheet.id);
      this.bodyResultSheets.delete(sheet.id);
      this.bodyExtents.set(sheet.id, { rowCount: sheet.rowCount, columnCount: sheet.columnCount });
      const state = this.sheets.get(sheet.id);
      const cells = sheet.cells ?? {};
      if (!state) {
        this.bodyBaseline.set(sheet.id, new Map(Object.entries(cells)));
        continue;
      }
      const own = stored?.get(sheet.id);
      state.bodyKeys = own ? new Set([...own].filter((key) => cells[key] !== undefined)) : new Set(Object.keys(cells));
      this.bodyBaseline.set(sheet.id, new Map([...state.bodyKeys].map((key) => [key, cells[key]])));
    }
  }

  /** The body cells (of any sheet) whose input changed since the last
   * save: a calculation stamping a value is not a change. */
  private collectTouched(sheets: readonly SheetLike[], since?: ReadonlyMap<string, readonly string[] | null>): void {
    const touched: TiledCellEdit[] = [];
    for (const sheet of sheets) {
      const state = this.sheets.get(sheet.id);
      const baseline = this.bodyBaseline.get(sheet.id);
      if (!baseline) {
        this.bodyBaseline.set(sheet.id, new Map(Object.entries(sheet.cells)));
        continue;
      }
      const owned = (key: string) => !state || state.bodyKeys.has(key);
      const compare = (key: string): void => {
        if (!owned(key)) return;
        const cell = sheet.cells[key];
        const before = baseline.get(key);
        if (before === cell) return;
        baseline.set(key, cell);
        if (!sameBesidesResult(before, cell)) this.bodyEditedSheets.add(sheet.id);
        else this.bodyResultSheets.add(sheet.id);
        if (before !== undefined && inputOf(before) === inputOf(cell)) return;
        const [row, col] = parseKey(key);
        touched.push({ sheet: sheet.id, row, col, cell: tilePayload(cell) });
      };
      const gone = (key: string): void => {
        baseline.delete(key);
        this.bodyEditedSheets.add(sheet.id);
        const [row, col] = parseKey(key);
        touched.push({ sheet: sheet.id, row, col, cell: null });
      };
      // Only the keys changed since the last collect (see `collect`).
      const changed = since?.get(sheet.id);
      if (changed) {
        for (const key of changed) {
          if (Object.hasOwn(sheet.cells, key)) compare(key);
          if (sheet.cells[key] === undefined && baseline.has(key)) gone(key);
        }
        continue;
      }
      for (const key of Object.keys(sheet.cells)) compare(key);
      for (const key of [...baseline.keys()]) {
        if (sheet.cells[key] !== undefined) continue;
        gone(key);
      }
    }
    if (!touched.length) return;
    const replaced = new Set(touched.map((edit) => `${edit.sheet}\u0000${edit.row}:${edit.col}`));
    this.touched = [...this.touched.filter((edit) => !replaced.has(`${edit.sheet}\u0000${edit.row}:${edit.col}`)), ...touched];
  }

  /**
   * Sheets an Undo or Redo brought back, each formula still as the tiles hold
   * it keeping the value the tiles hold: the step holds what it was calculated
   * to then. Compared as an edit, every formula the undone cell fed went back to
   * the server with its old value, 4,719 cells, some of them blank, until the
   * calculation came round again (xlsx-rounds R133-X02).
   */
  withResultsRead<S extends SheetLike>(sheets: readonly S[]): S[] {
    const isFormula = (cell: SpreadsheetCell | undefined) => typeof cell?.formula === "string" && cell.formula.trim().startsWith("=");
    return sheets.map((sheet) => {
      const state = this.sheets.get(sheet.id);
      if (!state) return sheet;
      let cells: CellMap | null = null;
      const keys: string[] = [];
      for (const tile of state.tiles.values()) {
        if (!tile.applied) continue;
        for (const [key, read] of tile.baseline) {
          if (!isFormula(read)) continue;
          const cell = sheet.cells[key];
          if (cell === read || !isFormula(cell) || inputOf(cell) !== inputOf(read)) continue;
          cells ??= { ...sheet.cells };
          cells[key] = read;
          keys.push(key);
        }
      }
      if (!cells) return sheet;
      recordCellMapEdit(sheet.cells, cells, keys);
      return { ...sheet, cells };
    }) as S[];
  }

  /** Whether a formula of this sheet's body took a new result since it was
   * loaded, as of the last `collect`: its cells go up with the next save. */
  bodyResultsMoved(sheetId: string): boolean {
    return this.bodyResultSheets.has(sheetId);
  }

  /** The window shows a line edit the server has not made yet: until it has
   * (`rereadAfterServerMove`) or has refused (`false`), nothing is collected. */
  holdCollect(held: boolean): void {
    this.collectHeld = held;
  }

  /** Whether this sheet's body was edited since it was loaded, as of the last
   * `collect`. A tile arriving or a cell typed into a tile is not an edit of
   * the body, nor is a calculation stamping a result: a save has only moved
   * results to stamp on such a sheet (`withDisplayCache`'s `resultsOnly`). */
  bodyEdited(sheetId: string): boolean {
    return this.bodyEditedSheets.has(sheetId);
  }

  /** The sheets as the page body holds them: tile-owned cells left out. */
  bodySheets<S extends SheetLike>(sheets: readonly S[]): S[] {
    return sheets.map((sheet) => {
      const state = this.sheets.get(sheet.id);
      if (!state) return sheet;
      const cells: CellMap = {};
      for (const key of state.bodyKeys) {
        const cell = sheet.cells[key];
        if (cell !== undefined) cells[key] = cell;
      }
      // Rows and columns grown by a cell typed into the tiles are the tiles'
      // (the server grows the sheet's extent with the save, `apply` grows the
      // model from the cells): not a change of the body. Appending below a
      // table sent the body through the page's write-back every time, 36 s
      // under the workspace's lock on a 60-chart workbook.
      const stated = this.bodyEditedSheets.has(sheet.id) ? undefined : this.bodyExtents.get(sheet.id);
      return stated && typeof stated.rowCount === "number" && typeof stated.columnCount === "number" &&
        Number(sheet.rowCount) >= stated.rowCount && Number(sheet.columnCount) >= stated.columnCount
        ? { ...sheet, cells, rowCount: stated.rowCount, columnCount: stated.columnCount }
        : { ...sheet, cells };
    });
  }

  // ── reading ────────────────────────────────────────────────────────────

  /** The tiles holding rows `first..last` of a sheet. */
  tilesFor(sheetId: string, first: number, last: number): number[] {
    const state = this.sheets.get(sheetId);
    if (!state) return [];
    const rows = state.entry.tileRows;
    const top = Math.max(0, Math.floor(first / rows));
    const bottom = Math.max(top, Math.floor(Math.max(first, last) / rows));
    const out: number[] = [];
    for (let tile = top; tile <= bottom; tile += 1) out.push(tile);
    return out;
  }

  /** The cells one of the sheet's tiles holds on average: what a read of a
   * tile costs to fetch and to put in the model. */
  averageTileCells(sheetId: string): number {
    const entry = this.sheets.get(sheetId)?.entry;
    if (!entry) return 0;
    const tiles = Math.max(1, Math.ceil(Math.max(1, entry.rows) / Math.max(1, entry.tileRows)));
    return entry.cells / tiles;
  }

  /** The cells the sheet holds, as its manifest counts them. */
  cellsOf(sheetId: string): number {
    return this.sheets.get(sheetId)?.entry.cells ?? 0;
  }

  /** Whether the sheet is read whole when shown (`WHOLE_SHEET_CELLS`). */
  readsWhole(sheetId: string): boolean {
    const entry = this.sheets.get(sheetId)?.entry;
    return entry !== undefined && (entry.cells <= (askedWholeCells() ?? WHOLE_SHEET_CELLS) || this.wholeOnRequest.has(sheetId));
  }

  /** A command needs all of this sheet: it is read whole from now on, rather
   * than around the screen. False when it is past what a tab holds. */
  askWhole(sheetId: string): boolean {
    const entry = this.sheets.get(sheetId)?.entry;
    if (!entry || entry.cells > WHOLE_ON_REQUEST_CELLS) return false;
    this.wholeOnRequest.add(sheetId);
    return true;
  }

  /** Whether the whole workbook is read and kept (`WHOLE_WORKBOOK_CELLS`). */
  readsWholeWorkbook(): boolean {
    let cells = 0;
    for (const state of this.sheets.values()) cells += Math.max(0, Number(state.entry.cells) || 0);
    return cells <= (askedWholeCells() ?? WHOLE_WORKBOOK_CELLS);
  }

  






  async readOtherSheets(shown: string, stillWanted: () => boolean, arrived: () => void): Promise<void> {
    const queue = this.sheetIds().filter((id) => id !== shown && !this.isWhollyResident(id));
    let next = 0;
    let handed = Date.now();
    const reader = async (): Promise<void> => {
      while (next < queue.length && stillWanted()) {
        const sheet = queue[next];
        next += 1;
        await this.ensure(sheet, this.allTilesOf(sheet));
        if (stillWanted() && Date.now() - handed >= OTHER_SHEETS_APPLY_MS) {
          handed = Date.now();
          arrived();
        }
      }
    };
    try {
      await Promise.all(Array.from({ length: OTHER_SHEET_READERS }, reader));
    } finally {
      if (stillWanted()) arrived();
    }
  }

  /** Every tile of a sheet (by its extent). */
  allTilesOf(sheetId: string): number[] {
    const entry = this.sheets.get(sheetId)?.entry;
    return entry ? this.tilesFor(sheetId, 0, Math.max(0, entry.rows - 1)) : [];
  }

  /** A deleted sheet: no longer a tiled one here. Its tiles are retired by
   * the server with the body save that drops it (workbook_tiles.retire); a
   * sheet made later under its id is an ordinary one, its cells the body's. */
  retire(sheetId: string): void {
    if (!this.sheets.delete(sheetId)) return;
    this.tileState += 1;
    this.pending = this.pending.filter((edit) => edit.sheet !== sheetId);
    this.touched = this.touched.filter((edit) => edit.sheet !== sheetId);
    this.bodyBaseline.delete(sheetId);
    this.unanswered.delete(sheetId);
    this.answered.delete(sheetId);
    for (const key of [...this.reading.keys()]) if (key.startsWith(`${sheetId}:`)) this.reading.delete(key);
  }

  /** The tiled sheets, in the manifest's order. */
  sheetIds(): string[] {
    return [...this.sheets.keys()];
  }

  /** The rows to keep around rows `first..last` of the sheet shown. */
  windowRows(sheetId: string, first: number, last: number): [number, number] {
    const entry = this.sheets.get(sheetId)?.entry;
    if (entry && this.readsWhole(sheetId)) return [0, Math.max(0, entry.rows - 1)];
    const span = Math.max(64, last - first);
    return [Math.max(0, first - span), last + span];
  }

  /** Whether every cell of the sheet is in the model (a command reading the
   * whole sheet may run). */
  isWhollyResident(sheetId: string): boolean {
    const state = this.sheets.get(sheetId);
    if (!state) return true;
    const last = Math.floor(Math.max(0, state.entry.rows - 1) / state.entry.tileRows);
    for (let tile = 0; tile <= last; tile += 1) {
      const resident = state.tiles.get(tile);
      if (!resident || !resident.applied) return false;
    }
    return true;
  }

  /** Whether every cell of rows `first..last` of the sheet is in the model
   * (true for a sheet that is not tiled). */
  rowsResident(sheetId: string, first: number, last: number): boolean {
    const state = this.sheets.get(sheetId);
    if (!state) return true;
    const bottom = Math.min(last, Math.max(0, state.entry.rows - 1));
    if (bottom < first) return true;
    for (const tile of this.tilesFor(sheetId, first, bottom)) {
      if (!state.tiles.get(tile)?.applied) return false;
    }
    return true;
  }

  /** The tiles holding rows `first..last` of the sheet that are not yet in
   * the model, within its extent. */
  tilesMissingFor(sheetId: string, first: number, last: number): number[] {
    const state = this.sheets.get(sheetId);
    if (!state) return [];
    const bottom = Math.min(last, Math.max(0, state.entry.rows - 1));
    if (bottom < first) return [];
    return this.tilesFor(sheetId, first, bottom).filter((tile) => !state.tiles.get(tile)?.applied);
  }

  /** The rows derived results read, which `evict` keeps once read (replaces
   * the previous set). */
  keepRows(reads: ReadonlyArray<{ readonly sheetId: string; readonly first: number; readonly last: number }>): void {
    const kept = new Map<string, Array<readonly [number, number]>>();
    for (const { sheetId, first, last } of reads) {
      if (!this.sheets.has(sheetId)) continue;
      const list = kept.get(sheetId) ?? [];
      list.push([first, last]);
      kept.set(sheetId, list);
    }
    this.keptRows = kept;
  }

  isApplied(sheetId: string, tile: number): boolean {
    return this.sheets.get(sheetId)?.tiles.get(tile)?.applied ?? false;
  }

  isResident(sheetId: string, tile: number): boolean {
    return this.sheets.get(sheetId)?.tiles.has(tile) ?? false;
  }

  /** Read the tiles not yet resident. Resolves when they are (not yet in
   * the model: `apply` puts them there). */
  async ensure(sheetId: string, tiles: readonly number[], signal?: AbortSignal): Promise<void> {
    const state = this.sheets.get(sheetId);
    if (!state) return;
    const now = ++this.clock;
    for (const tile of tiles) {
      const resident = state.tiles.get(tile);
      if (resident) resident.usedAt = now;
    }
    const missing = tiles.filter((tile) => !state.tiles.has(tile) && !this.reading.has(`${sheetId}:${tile}`));
    const waits = tiles.map((tile) => this.reading.get(`${sheetId}:${tile}`)).filter(Boolean) as Promise<void>[];
    for (let start = 0; start < missing.length; start += TILES_PER_READ) {
      const batch = missing.slice(start, start + TILES_PER_READ);
      const read = this.read(sheetId, batch, signal);
      for (const tile of batch) this.reading.set(`${sheetId}:${tile}`, read);
      waits.push(read);
    }
    await Promise.all(waits);
  }

  /** The first tiles of a link's page, read ahead of the page's own read
   * (`prepareTiledSheets`): the manifest answered from the viewer's workspace,
   * so these need not wait to learn which one the notebook is in. */
  readsBeforeMountKnown = false;

  private async read(sheetId: string, tiles: readonly number[], signal?: AbortSignal, refresh = false): Promise<void> {
    try {
      const answer = await apiGet<EditorTilesAnswer>(
        `${base(this.pageId)}/tiles?sheet=${encodeURIComponent(sheetId)}&tiles=${tiles.join(",")}&format=editor`,
        { signal },
        this.readsBeforeMountKnown ? { beforeMountKnown: true } : undefined,
      );
      this.accept(sheetId, answer, refresh);
    } finally {
      for (const tile of tiles) this.reading.delete(`${sheetId}:${tile}`);
    }
  }

  /** The tiles an answer holds, resident from now on (as a read of them
   * would have made them); a tile already resident stays as it is unless
   * `refresh`. The manifest's `first_tiles` come in this way. */
  accept(sheetId: string, answer: EditorTilesAnswer, refresh = false): void {
    const state = this.sheets.get(sheetId);
    if (!state) return;
    const styles = answer.styles.map((style) => this.intern(style));
    this.generation += 1;
    this.tileState += 1;
    const everRead = (state.everRead ??= new Set());
    for (const tile of answer.tiles) {
      if (!refresh && state.tiles.has(tile.tile)) continue;
      const firstRead = !refresh && !everRead.has(tile.tile);
      everRead.add(tile.tile);
      const held = state.tiles.get(tile.tile);
      const shown = held ? (held.applied ? held.baseline : held.shown) : undefined;
      const baseline = new Map<string, SpreadsheetCell>();
      for (const [key, raw] of Object.entries(tile.cells)) {
        let cell = raw;
        if (typeof raw.styleRef === "number") {
          const { styleRef, ...rest } = raw;
          cell = { ...rest, style: styles[styleRef as number] };
        }
        baseline.set(key, cell as unknown as SpreadsheetCell);
      }
      state.tiles.set(tile.tile, {
        index: tile.tile,
        revision: tile.revision,
        generation: this.generation,
        baseline,
        applied: false,
        usedAt: ++this.clock,
        ...(firstRead ? { firstRead: true } : {}),
        ...(shown ? { shown } : {}),
      });
    }
    this.emit();
  }

  /**
   * The server moved this workbook's lines (`workbook-tiles/line-edit`, or its
   * undo): every cell of every sheet may be under another key, and every tile
   * this adapter holds says where they were. The manifest and the tiles it held
   * are read again, into maps the model is not shown until the window takes the
   * page's new body (`acceptBody`, `reapply`, `apply`). Throws when they cannot
   * be read, with what it held left as it was: the caller reloads.
   */
  async rereadAfterServerMove(signal?: AbortSignal): Promise<void> {
    const opening = await apiGet<StructuredTiledOpening>(base(this.pageId), { signal });
    const entries = opening.manifest.sheets as unknown as TiledSheetEntry[];
    const wanted = new Map<string, number[]>();
    const next = new Map<string, Map<number, ResidentTile>>();
    for (const entry of entries) {
      const state = this.sheets.get(entry.id);
      const rows = Math.max(1, entry.tileRows);
      const last = Math.floor(Math.max(0, entry.rows - 1) / rows);
      // Those it held, as far as the sheet now goes, and the rows a whole
      // sheet has grown by.
      const tiles = new Set([...(state?.tiles.keys() ?? [])].filter((tile) => tile <= last));
      if (state && state.tiles.size > 0 && (entry.cells <= (askedWholeCells() ?? WHOLE_SHEET_CELLS) || this.wholeOnRequest.has(entry.id))) {
        for (let tile = 0; tile <= last; tile += 1) tiles.add(tile);
      }
      wanted.set(entry.id, [...tiles].sort((a, b) => a - b));
      next.set(entry.id, new Map());
    }
    for (const [sheetId, tiles] of wanted) {
      for (let start = 0; start < tiles.length; start += TILES_PER_READ) {
        const batch = tiles.slice(start, start + TILES_PER_READ);
        const answer = await apiGet<EditorTilesAnswer>(
          `${base(this.pageId)}/tiles?sheet=${encodeURIComponent(sheetId)}&tiles=${batch.join(",")}&format=editor`,
          { signal },
        );
        const styles = answer.styles.map((style) => this.intern(style));
        for (const tile of answer.tiles) {
          const baseline = new Map<string, SpreadsheetCell>();
          for (const [key, raw] of Object.entries(tile.cells)) {
            let cell = raw;
            if (typeof raw.styleRef === "number") {
              const { styleRef, ...rest } = raw;
              cell = { ...rest, style: styles[styleRef as number] };
            }
            baseline.set(key, cell as unknown as SpreadsheetCell);
          }
          next.get(sheetId)!.set(tile.tile, {
            index: tile.tile,
            revision: tile.revision,
            generation: this.generation + 1,
            baseline,
            applied: false,
            usedAt: ++this.clock,
          });
        }
      }
    }
    // All of it read: nothing awaits from here on, so no render sees half.
    this.generation += 1;
    this.tileState += 1;
    this.revision = opening.manifest.revision ?? this.revision;
    for (const sheetId of [...this.sheets.keys()]) {
      if (!next.has(sheetId)) this.sheets.delete(sheetId);
    }
    for (const entry of entries) {
      const state = this.sheets.get(entry.id);
      if (state) {
        state.entry = entry;
        state.tiles = next.get(entry.id)!;
        state.loose.clear();
      } else {
        this.sheets.set(entry.id, { entry, tiles: next.get(entry.id)!, bodyKeys: new Set(), loose: new Map() });
      }
    }
    this.unanswered.clear();
    this.answered.clear();
    this.partsWritten.clear();
    this.collectHeld = false;
    this.emit();
  }

  /**
   * Whether the server's tiles store what `expected` says at each of its
   * positions, read anew (the resident tiles are left as they are): a
   * caller confirming that a save reached the tiles. The tiles hold a cell
   * the way the workbook import reads it (lib comment above), so a stored
   * cell holds the expected one when it has the same formula (or, without
   * one, the same value) and every style property the expected cell states;
   * an expected blank holds where no cell, or an empty one, is stored. False
   * when the sheet is not tiled here.
   */
  async storedHolds(sheetId: string, expected: ReadonlyMap<string, SpreadsheetCell | undefined>, signal?: AbortSignal): Promise<boolean> {
    const state = this.sheets.get(sheetId);
    if (!state) return false;
    const rows = state.entry.tileRows;
    const tiles = [...new Set([...expected.keys()].map((key) => Math.floor(parseKey(key)[0] / rows)))].sort((a, b) => a - b);
    const stored = new Map<string, SpreadsheetCell>();
    for (let start = 0; start < tiles.length; start += TILES_PER_READ) {
      const batch = tiles.slice(start, start + TILES_PER_READ);
      const answer = await apiGet<EditorTilesAnswer>(
        `${base(this.pageId)}/tiles?sheet=${encodeURIComponent(sheetId)}&tiles=${batch.join(",")}&format=editor`,
        { signal },
      );
      for (const tile of answer.tiles) {
        for (const [key, raw] of Object.entries(tile.cells)) {
          if (!expected.has(key)) continue;
          let cell = raw;
          if (typeof raw.styleRef === "number") {
            const { styleRef, ...rest } = raw;
            cell = { ...rest, style: answer.styles[styleRef as number] };
          }
          stored.set(key, cell as unknown as SpreadsheetCell);
        }
      }
    }
    const blank = (cell: SpreadsheetCell | undefined) =>
      !cell || ((cell.value === undefined || cell.value === null || cell.value === "") && typeof cell.formula !== "string");
    for (const [key, want] of expected) {
      const have = stored.get(key);
      if (!want || blank(want)) {
        if (!blank(have)) return false;
        continue;
      }
      if (!have) return false;
      const formula = (cell: SpreadsheetCell) =>
        typeof cell.formula === "string" && cell.formula.trim().startsWith("=") ? cell.formula.trim() : null;
      if (formula(want) !== formula(have)) return false;
      if (formula(want) === null && String(want.value ?? "") !== String(have.value ?? "")) return false;
      const wantStyle = (want.style ?? {}) as Record<string, unknown>;
      const haveStyle = (have.style ?? {}) as Record<string, unknown>;
      for (const [name, value] of Object.entries(wantStyle)) {
        if (value === undefined || value === null || name === SOURCE_XF_STYLE_KEY) continue;
        if (JSON.stringify(haveStyle[name] ?? null) !== JSON.stringify(value)) return false;
      }
      for (const name of ["locked", "formulaHidden"] as const)
        if (want[name] !== undefined && want[name] !== have[name]) return false;
    }
    return true;
  }

  /** A body read anew replaced the model: every resident tile goes back in. */
  /** True when each of `sheets` holds every tile read so far (its cells are
   * what an `apply` after the last read made, or later). */
  holdsEveryTile(sheets: readonly SheetLike[]): boolean {
    return sheets.every((sheet) => !this.sheets.has(sheet.id) || (this.seen.get(sheet.cells) ?? -1) >= this.generation);
  }

  reapply(): void {
    this.tileState += 1;
    // The model is made again from a body: no earlier reading is what it shows.
    for (const state of this.sheets.values()) {
      for (const tile of state.tiles.values()) {
        tile.applied = false;
        delete tile.shown;
      }
    }
  }

  /** The model with every tile read since in it. A position the body owns,
   * or one with an edit not yet saved, keeps the model's cell. */
  apply<S extends SheetLike>(sheets: readonly S[]): S[] {
    let changed = false;
    const out = sheets.map((sheet) => {
      const state = this.sheets.get(sheet.id);
      if (!state) return sheet;
      let cells: CellMap | null = null;
      const taken = new Map<number, ReadonlyMap<string, unknown>>();
      const changedKeys: string[] = [];
      const dirty = this.dirtyKeys(sheet.id);
      // Cells the window typed or cleared since the last collect in a tile
      // read again: the model's, as an edit waiting is (`collect` saves them).
      const own = new Set<string>();
      const shownByTile = new Map<number, Map<string, SpreadsheetCell>>();
      for (const tile of state.tiles.values()) {
        if (!tile.applied && tile.shown) shownByTile.set(tile.index, tile.shown);
        delete tile.shown;
      }
      if (shownByTile.size) {
        const rows = state.entry.tileRows;
        const isOwn = (key: string, shown: Map<string, SpreadsheetCell>): boolean => {
          if (state.bodyKeys.has(key)) return false;
          const cell = sheet.cells[key];
          const before = shown.get(key);
          if (before === undefined) return cell !== undefined && !isLineFormatBlank(sheet, key, cell);
          // What a calculation stamped since is not the window's edit.
          return cell === undefined || (cell !== before && !sameBesidesResult(before, cell));
        };
        for (const shown of shownByTile.values()) for (const key of shown.keys()) if (isOwn(key, shown)) own.add(key);
        // Cells typed where the earlier reading had none: one pass over the
        // model for every tile read again.
        for (const key of Object.keys(sheet.cells)) {
          const shown = shownByTile.get(Math.floor(rowOfKey(key) / rows));
          if (shown && !shown.has(key) && isOwn(key, shown)) own.add(key);
        }
      }
      let sweep = false;
      for (const tile of state.tiles.values()) {
        if (tile.applied) continue;
        tile.applied = true;
        if (!tile.firstRead) sweep = true;
        tile.firstRead = false;
        this.tileState += 1;
        cells ??= { ...sheet.cells };
        taken.set(tile.index, tile.baseline);
        for (const [key, cell] of tile.baseline) {
          if (dirty.has(key) || own.has(key)) continue;
          if (state.bodyKeys.has(key)) {
            // A formatted blank of the body over a cell the tile holds (typed
            // there, or written by a save before ranges stayed the body's):
            // the tile's is the cell.
            if (!isFormatOnlyCell(cells[key])) continue;
            state.bodyKeys.delete(key);
            this.bodyBaseline.get(sheet.id)?.delete(key);
          }
          cells[key] = cell;
          changedKeys.push(key);
          state.loose.delete(key);
        }
      }
      if (!cells) return sheet;
      // A cell the model held in those tiles that a tile no longer has (another
      // member cleared it) goes, unless it is the model's own edit. One pass
      // over the map for all of them: once per tile, ten tiles over 250,000
      // cells were half a second of an open.
      // Tiles read for the first time only (a scroll into rows never read):
      // the model holds nothing of theirs to sweep, and the sweep over every
      
      const rows = state.entry.tileRows;
      if (sweep) {
        for (const key of Object.keys(cells)) {
          const baseline = taken.get(Math.floor(rowOfKey(key) / rows));
          if (!baseline || baseline.has(key) || state.bodyKeys.has(key) || dirty.has(key) || own.has(key) || state.loose.has(key)) continue;
          delete cells[key];
          changedKeys.push(key);
        }
      }
      changed = true;
      recordCellMapEdit(sheet.cells, cells, changedKeys);
      // A cell the tiles hold past the extent the body states shows.
      let rowCount = Number(sheet.rowCount) || 0;
      let columnCount = Number(sheet.columnCount) || 0;
      for (const baseline of taken.values()) {
        for (const key of baseline.keys()) {
          if (cells[key] === undefined) continue;
          const [row, col] = parseKey(key);
          if (row >= rowCount) rowCount = row + 1;
          if (col >= columnCount) columnCount = col + 1;
        }
      }
      return rowCount !== sheet.rowCount || columnCount !== sheet.columnCount
        ? { ...sheet, cells, ...(sheet.rowCount === undefined ? {} : { rowCount, columnCount }) }
        : { ...sheet, cells };
    });
    if (!changed) return sheets as S[];
    for (const sheet of out) this.seen.set(sheet.cells, this.generation);
    return out;
  }

  /**
   * Sheets that stood as placeholders and are filled in now (SHEET-DEFER-DESIGN.md):
   * each body is taken as loaded (`stored`, from `storedBodyCellKeys` of the
   * stored content with them in place), and every tile read of it goes in,
   * the ones read while its placeholder stood included. A filled sheet is a
   * new cell map, which `reconcile` takes as current: without this the
   * window showed a deferred sheet's few body cells and none of its tiles'.
   */
  adoptFilledSheets<S extends SheetLike>(sheets: readonly S[], stored?: Map<string, Set<string>> | null): S[] {
    this.acceptBody(sheets, stored);
    for (const sheet of sheets) {
      for (const tile of this.sheets.get(sheet.id)?.tiles.values() ?? []) {
        tile.applied = false;
        delete tile.shown;
      }
    }
    this.tileState += 1;
    const out = this.apply(sheets);
    for (const sheet of out) if (this.sheets.has(sheet.id)) this.seen.set(sheet.cells, this.generation);
    return out;
  }

  /** A model the window is about to show: one it has not seen (a command's)
   * is current; one it has (an undo restoring an earlier model) gets the
   * tiles read since that model was current, and loses the tiles let go
   * since (they are read again when they come into view). */
  reconcile<S extends SheetLike>(sheets: readonly S[]): S[] {
    let changed = false;
    const out = sheets.map((sheet) => {
      const state = this.sheets.get(sheet.id);
      if (!state) return sheet;
      const seenAt = this.seen.get(sheet.cells);
      if (seenAt === undefined) {
        this.seen.set(sheet.cells, this.generation);
        return sheet;
      }
      if (seenAt >= this.generation) return sheet;
      let cells: CellMap | null = null;
      for (const tile of state.tiles.values()) {
        if (tile.generation <= seenAt) continue;
        for (const [key, cell] of tile.baseline) {
          if (state.bodyKeys.has(key)) continue;
          if (sheet.cells[key] !== undefined && !isLineFormatBlank(sheet, key, sheet.cells[key])) continue;
          cells ??= { ...sheet.cells };
          cells[key] = cell;
        }
      }
      if (!cells) {
        this.seen.set(sheet.cells, this.generation);
        return sheet;
      }
      changed = true;
      this.seen.set(cells, this.generation);
      return { ...sheet, cells };
    });
    return changed ? out : (sheets as S[]);
  }

  /** Let go of the tiles of a sheet far from rows `first..last` (and of
   * every other sheet's), keeping a few recent ones and every tile with an
   * edit not yet saved. The model without their cells. */
  evict<S extends SheetLike>(sheets: readonly S[], activeSheetId: string, first: number, last: number): S[] {
    // A workbook read whole keeps every tile.
    if (this.readsWholeWorkbook()) return sheets as S[];
    const keep = new Set(this.tilesFor(activeSheetId, first, last).map((tile) => `${activeSheetId}:${tile}`));
    for (const [sheetId, ranges] of this.keptRows) {
      for (const [from, to] of ranges) for (const tile of this.tilesFor(sheetId, from, to)) keep.add(`${sheetId}:${tile}`);
    }
    const candidates: { sheet: string; tile: ResidentTile }[] = [];
    let resident = 0;
    for (const [sheetId, state] of this.sheets) {
      for (const tile of state.tiles.values()) {
        resident += tile.baseline.size;
        if (!keep.has(`${sheetId}:${tile.index}`)) candidates.push({ sheet: sheetId, tile });
      }
    }
    candidates.sort((a, b) => b.tile.usedAt - a.tile.usedAt);
    const drop: { sheet: string; tile: ResidentTile }[] = [];
    let spare = 0;
    for (const candidate of candidates) {
      if (spare < SPARE_TILES && resident <= RESIDENT_CELL_BUDGET) {
        spare += 1;
        continue;
      }
      if (this.tileHasPending(candidate.sheet, candidate.tile.index)) continue;
      drop.push(candidate);
      resident -= candidate.tile.baseline.size;
    }
    if (!drop.length) return sheets as S[];
    const bySheet = new Map<string, Set<number>>();
    this.tileState += 1;
    for (const { sheet, tile } of drop) {
      this.sheets.get(sheet)?.tiles.delete(tile.index);
      if (!bySheet.has(sheet)) bySheet.set(sheet, new Set());
      bySheet.get(sheet)!.add(tile.index);
    }
    this.generation += 1;
    const out = sheets.map((sheet) => {
      const gone = bySheet.get(sheet.id);
      const state = this.sheets.get(sheet.id);
      if (!gone || !state) return sheet;
      const rows = state.entry.tileRows;
      const cells: CellMap = {};
      const changedKeys: string[] = [];
      for (const key in sheet.cells) {
        const cell = sheet.cells[key];
        if (gone.has(Math.floor(rowOfKey(key) / rows)) && !state.bodyKeys.has(key) && !state.loose.has(key)) {
          changedKeys.push(key);
          continue;
        }
        cells[key] = cell;
      }
      recordCellMapEditOver(sheet.cells, cells, changedKeys);
      return { ...sheet, cells };
    });
    for (const sheet of out) this.seen.set(sheet.cells, this.generation);
    return out;
  }

  /** Cells the tiles kept in memory hold. */
  residentCells(): number {
    let total = 0;
    for (const state of this.sheets.values()) for (const tile of state.tiles.values()) total += tile.baseline.size;
    return total;
  }

  // ── saving ─────────────────────────────────────────────────────────────

  private dirtyKeys(sheetId: string): Set<string> {
    const out = new Set<string>();
    for (const edit of this.pending) if (edit.sheet === sheetId) out.add(`${edit.row}:${edit.col}`);
    return out;
  }

  private tileHasPending(sheetId: string, tile: number): boolean {
    const rows = this.sheets.get(sheetId)?.entry.tileRows ?? 1;
    return this.pending.some((edit) => edit.sheet === sheetId && Math.floor(edit.row / rows) === tile) || this.inFlight !== null;
  }

  /**
   * What the model changed in the tiles since they were read or saved: each
   * tile-owned cell that differs from its tile's (a new one, an edited one,
   * a cleared one). A cell that gained what a tile cannot hold (a comment, a
   * link) becomes the body's, and leaves its tile. The edits join the
   * pending ones, and the tiles' record moves to the model.
   */
  collect(sheets: readonly SheetLike[]): number {
    // The model shows lines the server has not moved yet: compared with the
    // tiles, every moved cell would read as an edit.
    if (this.collectHeld) return 0;
    // A tiled sheet the workbook no longer has (deleted by any command) is
    // retired here too: a later sheet under its id is an ordinary one.
    if (sheets.length) {
      const present = new Set(sheets.map((sheet) => sheet.id));
      for (const sheetId of [...this.sheets.keys()]) if (!present.has(sheetId)) this.retire(sheetId);
    }
    // Each sheet's keys changed since a map of its lineage was last
    // collected with the tiles as they are (null: every key is compared).
    const since = new Map<string, readonly string[] | null>();
    for (const sheet of sheets) since.set(sheet.id, this.changedSinceCollected(sheet));
    // A sort of a sheet's rows is saved as the order it put them in: the
    // tiles' cells there are moved here as the server will move them, and
    // only the positions that then differ are compared below.
    let sorted = false;
    for (const sheet of sheets) {
      const differing = this.takeRowSort(sheet);
      if (!differing) continue;
      sorted = true;
      const known = since.get(sheet.id);
      if (known) since.set(sheet.id, [...known, ...differing]);
    }
    // What collect last read is no longer what the tiles hold.
    if (sorted) this.tileState += 1;
    const edits: TiledCellEdit[] = [];
    for (const sheet of sheets) {
      const state = this.sheets.get(sheet.id);
      if (!state) continue;
      const rows = state.entry.tileRows;
      const cells = sheet.cells;
      const owner = (key: string): ResidentTile | undefined => state.tiles.get(Math.floor(parseKey(key)[0] / rows));
      const visit = (key: string): void => this.compareCell(sheet, state, key, edits);
      const cleared = (key: string): void => {
        const [row, col] = parseKey(key);
        edits.push({ sheet: sheet.id, row, col, cell: null });
      };
      const changed = since.get(sheet.id);
      if (changed) {
        for (const key of changed) {
          if (Object.hasOwn(cells, key)) visit(key);
          if (cells[key] !== undefined) continue;
          // As the walks below do for a cell the model cleared.
          const tile = owner(key);
          if (tile?.applied && tile.baseline.has(key)) {
            cleared(key);
            tile.baseline.delete(key);
          } else if (tile && !tile.applied && tile.shown?.has(key)) {
            // Cleared while the tile was read again (`ResidentTile.shown`).
            cleared(key);
            tile.shown.delete(key);
            tile.baseline.delete(key);
          }
          if (state.loose.has(key)) {
            cleared(key);
            state.loose.delete(key);
          }
          if (state.bodyKeys.has(key)) {
            state.bodyKeys.delete(key);
            cleared(key);
          }
        }
        continue;
      }
      for (const key of Object.keys(cells)) visit(key);
      // Cells the tiles had that the model cleared.
      for (const tile of state.tiles.values()) {
        const synced = tile.applied ? tile.baseline : tile.shown;
        if (!synced) continue;
        for (const key of [...synced.keys()]) {
          if (cells[key] !== undefined) continue;
          cleared(key);
          synced.delete(key);
          tile.baseline.delete(key);
        }
      }
      for (const key of [...state.loose.keys()]) {
        if (cells[key] !== undefined) continue;
        cleared(key);
        state.loose.delete(key);
      }
      // Body cells the model cleared: the tile under them is cleared too, so
      // the original's cell there never comes back on export.
      for (const key of [...state.bodyKeys]) {
        if (cells[key] !== undefined) continue;
        state.bodyKeys.delete(key);
        cleared(key);
      }
    }
    // The body cells changed, read after the tiles let go of the cells the
    // body takes (above): read before, a sorted header the body took at its
    // new row reached the save only on the next collect, a second save of
    // the 730,000-cell sort calculated again (3.5 s) for 11 header cells.
    this.collectTouched(sheets, since);
    for (const sheet of sheets) {
      this.collected.set(sheet.cells, {
        sheetId: sheet.id,
        tileState: this.tileState,
        rowStyles: sheet.rowStyles,
        columnStyles: sheet.columnStyles,
      });
    }
    if (edits.length) {
      // A later edit of a position replaces an earlier one still waiting.
      const replaced = new Set(edits.map((edit) => `${edit.sheet}\u0000${edit.row}:${edit.col}`));
      this.pending = [...this.pending.filter((edit) => !replaced.has(`${edit.sheet}\u0000${edit.row}:${edit.col}`)), ...edits];
    }
    return edits.length;
  }

  /** Sorts waiting to be saved as their order, sent before any cell edit. */
  private rowSorts: QueuedRowSort[] = [];

  /** A sort recorded for `sheet`'s cells (`recordCellMapRowPermutation`, on
   * them or an earlier map of their lineage), taken as the order to save: the
   * tiles' cells of the range are moved as the server moves them, and the
   * positions that then differ are returned. The body's cells in the range
   * move with their rows: the tiles hold nothing under them that the server
   * would move elsewhere, and they reach the calculation as the save's touched
   * cells, sent with the order. Null when it cannot be (a save on its way, a
   * sort already waiting, edits still waiting in the range, unread tiles
   * there): the range is then saved cell by cell.
   * Seen once either way: after a refusal the tiles are compared level with
   * the sorted cells, and taking the order then would move them twice. */
  private takeRowSort(sheet: SheetLike): string[] | null {
    const state = this.sheets.get(sheet.id);
    if (!state) return null;
    let after: CellMap = sheet.cells;
    let record = cellMapRowPermutation(after);
    if (!record) {
      const found = nearestKnownCellMap(sheet.cells, (cells) => (cellMapRowPermutation(cells) ? cells : undefined));
      if (!found) return null;
      after = found.value as CellMap;
      record = cellMapRowPermutation(after);
      if (!record) return null;
    }
    if (this.takenSorts.has(after)) return null;
    this.takenSorts.add(after);
    if (this.inFlight || this.rowSorts.length) return null;
    const { dataR1: r1, c1, c2, sourceOrder: order } = record.permutation;
    const before = record.before as CellMap;
    const r2 = r1 + order.length - 1;
    const rows = state.entry.tileRows;
    const inRange = (key: string): boolean => {
      const [row, col] = parseKey(key);
      return row >= r1 && row <= r2 && col >= c1 && col <= c2;
    };
    for (let tile = Math.floor(r1 / rows); tile <= Math.floor(r2 / rows); tile += 1) {
      if (!state.tiles.get(tile)?.applied) return null;
    }
    for (const key of state.loose.keys()) if (inRange(key)) return null;
    // An edit waiting there: the tile already reads as it, the server does not.
    for (const edit of [...this.pending, ...this.touched]) {
      if (edit.sheet === sheet.id && edit.row >= r1 && edit.row <= r2 && edit.col >= c1 && edit.col <= c2) return null;
    }
    // The tiles' cells of the range, by source row, taken out.
    const previous = new Map<string, SpreadsheetCell>();
    const bySource = new Map<number, [number, string, SpreadsheetCell][]>();
    for (let index = Math.floor(r1 / rows); index <= Math.floor(r2 / rows); index += 1) {
      const tile = state.tiles.get(index)!;
      for (const [key, cell] of tile.baseline) {
        if (!inRange(key)) continue;
        previous.set(key, cell);
        const [row, col] = parseKey(key);
        let list = bySource.get(row);
        if (!list) bySource.set(row, (list = []));
        list.push([col, key, cell]);
      }
    }
    for (const key of previous.keys()) state.tiles.get(Math.floor(parseKey(key)[0] / rows))!.baseline.delete(key);
    const differing: string[] = [];
    // The body's cells go with their rows, to where the model has them.
    const previousBodyKeys = new Set(state.bodyKeys);
    const destinationOf = new Map<number, number>();
    order.forEach((source, offset) => destinationOf.set(source, r1 + offset));
    const movedBodyKeys: string[] = [];
    for (const key of previousBodyKeys) {
      if (!inRange(key)) continue;
      const [row, col] = parseKey(key);
      state.bodyKeys.delete(key);
      differing.push(key);
      const moved = `${destinationOf.get(row)}:${col}`;
      if (after[moved] !== undefined) movedBodyKeys.push(moved);
    }
    for (const key of movedBodyKeys) {
      state.bodyKeys.add(key);
      differing.push(key);
    }
    // A position the body left holds a tile's cell now, or none: the stored
    // body still says otherwise until its own save lands, and a calculation
    // reads the body over the tiles.
    const bodyBaseline = this.bodyBaseline.get(sheet.id);
    const vacated: TiledCellEdit[] = [];
    for (const key of previousBodyKeys) {
      if (!inRange(key) || state.bodyKeys.has(key)) continue;
      bodyBaseline?.delete(key);
      const [row, col] = parseKey(key);
      const cell = after[key];
      vacated.push({ sheet: sheet.id, row, col, cell: cell === undefined ? null : tilePayload(cell) });
    }
    if (vacated.length) {
      const replaced = new Set(vacated.map((edit) => `${edit.row}:${edit.col}`));
      this.touched = [...this.touched.filter((edit) => edit.sheet !== sheet.id || !replaced.has(`${edit.row}:${edit.col}`)), ...vacated];
    }
    order.forEach((source, offset) => {
      const destination = r1 + offset;
      for (const [col, sourceKey, held] of bySource.get(source) ?? []) {
        const key = `${destination}:${col}`;
        // The model's own moved cell where the tile held the model's cell;
        // otherwise the tile's cell, moved as the server moves it: the model's may be one the
        
        const moved = held === before[sourceKey] && after[key] !== undefined && tileOwnable(after[key], held) ? after[key] : relocateSortedCell(held, destination - source, 0);
        state.tiles.get(Math.floor(destination / rows))!.baseline.set(key, moved);
        if (moved !== after[key]) differing.push(key);
      }
    });
    this.rowSorts.push({ sheet: { ...sheet, cells: after }, r1, c1, c2, order, previous, previousBodyKeys });
    return differing;
  }

  /** Cell maps whose sort was taken (or refused): never taken twice. */
  private readonly takenSorts = new WeakSet<object>();

  /** The server refused a sort's order: the tiles hold what they held before
   * it again, and every position of the range is compared (saved cell by
   * cell, as any other change). */
  private unsort(sort: QueuedRowSort): void {
    const state = this.sheets.get(sort.sheet.id);
    if (!state) return;
    state.bodyKeys = sort.previousBodyKeys;
    const { r1, c1, c2 } = sort;
    const r2 = r1 + sort.order.length - 1;
    const rows = state.entry.tileRows;
    const inRange = (row: number, col: number) => row >= r1 && row <= r2 && col >= c1 && col <= c2;
    const keys = new Set<string>(sort.previous.keys());
    for (let index = Math.floor(r1 / rows); index <= Math.floor(r2 / rows); index += 1) {
      const tile = state.tiles.get(index);
      if (!tile) continue;
      for (const key of [...tile.baseline.keys()]) {
        const [row, col] = parseKey(key);
        if (!inRange(row, col)) continue;
        keys.add(key);
        tile.baseline.delete(key);
      }
    }
    for (const [key, cell] of sort.previous) state.tiles.get(Math.floor(parseKey(key)[0] / rows))?.baseline.set(key, cell);
    for (const key of Object.keys(sort.sheet.cells)) {
      const [row, col] = parseKey(key);
      if (inRange(row, col)) keys.add(key);
    }
    const edits: TiledCellEdit[] = [];
    for (const key of keys) {
      if (sort.sheet.cells[key] !== undefined) {
        this.compareCell(sort.sheet, state, key, edits);
        continue;
      }
      const tile = state.tiles.get(Math.floor(parseKey(key)[0] / rows));
      if (tile?.baseline.has(key)) {
        const [row, col] = parseKey(key);
        edits.push({ sheet: sort.sheet.id, row, col, cell: null });
        tile.baseline.delete(key);
      }
    }
    this.tileState += 1;
    if (edits.length) {
      // Later edits of these positions (made after the sort) stay last.
      const later = new Set(this.pending.map((edit) => `${edit.sheet}\u0000${edit.row}:${edit.col}`));
      this.pending = [...edits.filter((edit) => !later.has(`${edit.sheet}\u0000${edit.row}:${edit.col}`)), ...this.pending];
    }
  }

  /** A cell of the model, compared with what its tile holds: an edit when
   * they differ (and the tile then holds it). */
  private compareCell(sheet: SheetLike, state: SheetState, key: string, edits: TiledCellEdit[]): void {
    const cell = sheet.cells[key];
    if (state.bodyKeys.has(key)) return;
    const tile = state.tiles.get(Math.floor(parseKey(key)[0] / state.entry.tileRows));
    // A tile read again and not yet in the model: the model is level with
    // the reading it showed (`ResidentTile.shown`), and a cell the new one
    // changed is another member's, not an edit of this window's.
    const synced = tile && !tile.applied && tile.shown ? tile.shown : tile?.baseline;
    const read = tile ? synced!.get(key) : state.loose.get(key);
    if (read === cell) return;
    // What a line format shows there, not a cell anyone wrote: never
    // saved over the tile (whose cell may not be read yet).
    if (read === undefined && isLineFormatBlank(sheet, key, cell)) return;
    if (!tileOwnable(cell, read)) {
      // The body takes it; its tile lets go of it.
      state.bodyKeys.add(key);
      if (read !== undefined || tile) {
        const [row, col] = parseKey(key);
        edits.push({ sheet: sheet.id, row, col, cell: null });
        tile?.baseline.delete(key);
        tile?.shown?.delete(key);
        state.loose.delete(key);
      }
      return;
    }
    if (read !== undefined && sameCell(read, cell)) {
      if (synced === tile?.shown && tile) synced!.set(key, cell);
      else if (tile) tile.baseline.set(key, cell);
      else state.loose.set(key, cell);
      return;
    }
    const [row, col] = parseKey(key);
    edits.push({ sheet: sheet.id, row, col, cell: tilePayload(cell) });
    if (tile) {
      tile.baseline.set(key, cell);
      tile.shown?.set(key, cell);
    } else state.loose.set(key, cell);
    this.noteFormula(sheet.id, key, cell);
  }

  /** The keys of `sheet`'s cells changed since `collect` last brought the
   * tiles level with a map of their lineage, the tiles unchanged since but by
   * it; null when no such map is known (every key is compared). */
  private changedSinceCollected(sheet: SheetLike): readonly string[] | null {
    const fits = (cells: object): true | undefined => {
      const at = this.collected.get(cells);
      return at && at.sheetId === sheet.id && at.tileState === this.tileState &&
        at.rowStyles === sheet.rowStyles && at.columnStyles === sheet.columnStyles ? true : undefined;
    };
    if (fits(sheet.cells)) return [];
    return nearestKnownCellMap(sheet.cells, fits)?.keys ?? null;
  }

  private noteFormula(sheetId: string, key: string, cell: SpreadsheetCell): void {
    const formula = typeof cell.formula === "string" && cell.formula.trim().startsWith("=") ? cell.value : null;
    const cache = cell.chartValueCache as { source?: unknown } | undefined;
    if (!formula || (cache && cache.source === formula)) return;
    if (!this.unanswered.has(sheetId)) this.unanswered.set(sheetId, new Map());
    this.unanswered.get(sheetId)!.set(key, formula);
  }

  /** Whether formula values wait to be stamped on the model. */
  get hasAnsweredFormulas(): boolean {
    for (const answers of this.answered.values()) if (answers.size) return true;
    return false;
  }

  get hasUnansweredFormulas(): boolean {
    for (const queue of this.unanswered.values()) if (queue.size) return true;
    return false;
  }

  /** Whether a sheet is read only around the screen, so a formula far from it
   * cannot be calculated here. */
  get hasPartialSheets(): boolean {
    for (const sheetId of this.sheets.keys()) if (!this.readsWhole(sheetId)) return true;
    return false;
  }

  /** Once, the server calculates the formulas the file holds no value for
   * (`poll` then brings in the tiles that changed). */
  async calculateBlanks(): Promise<number> {
    const answer = await apiPost<{ calculated: number }>(`${base(this.pageId)}/calculate-blanks`, {});
    return answer.calculated;
  }

  /** Ask the server the values of the formulas typed since (over the
   * workbook's stored cells). */
  async evaluate(): Promise<boolean> {
    let any = false;
    for (const [sheetId, queue] of this.unanswered) {
      const entries = [...queue.entries()];
      queue.clear();
      for (let start = 0; start < entries.length; start += 64) {
        const batch = entries.slice(start, start + 64);
        try {
          const answer = await apiPost<{ values: Record<string, unknown> | null }>(`${base(this.pageId)}/evaluate`, {
            sheet: sheetId,
            formulas: batch.map(([key, formula]) => [...parseKey(key), formula]),
          });
          if (!answer.values) continue;
          if (!this.answered.has(sheetId)) this.answered.set(sheetId, new Map());
          for (const [key, formula] of batch) {
            if (key in answer.values) {
              this.answered.get(sheetId)!.set(key, { formula, value: answer.values[key] });
              any = true;
            }
          }
        } catch {
          // Unanswered: the cell keeps no value until Excel recalculates.
        }
      }
    }
    return any;
  }

  /** The model with each answered formula's value stamped on its cell, as a
   * saved result is (`stored-formula-results`): what the grid shows. */
  stampResults<S extends SheetLike>(sheets: readonly S[]): S[] {
    let changed = false;
    const out = sheets.map((sheet) => {
      const answers = this.answered.get(sheet.id);
      if (!answers?.size) return sheet;
      let cells: CellMap | null = null;
      const stampedKeys: string[] = [];
      const state = this.sheets.get(sheet.id);
      for (const [key, { formula: asked, value }] of answers) {
        const cell = sheet.cells[key];
        // A formula the server calculated after a save (`asked` null) is the
        // cell's own formula, whatever it is; a typed one must still be it.
        const formula = asked ?? (typeof cell?.formula === "string" ? cell.formula : null);
        if (!cell || formula === null || cell.value !== formula) continue;
        // The value it holds already: no new model, so nothing is drawn or
        // saved again (twice over, 8 s each, after sorting a large sheet).
        if (typeof value === "number" && cell.chartValueCache?.kind === "number" &&
            cell.chartValueCache.value === value && cell.chartValueCache.source === formula) continue;
        const stamped: SpreadsheetCell = { ...cell };
        delete stamped.chartValueCache;
        delete stamped.display;
        if (typeof value === "number") {
          stamped.chartValueCache = { kind: "number", source: formula, value };
        } else if (typeof value === "boolean") {
          stamped.display = value ? "TRUE" : "FALSE";
        } else if (value && typeof value === "object" && typeof (value as { error?: unknown }).error === "string") {
          stamped.display = (value as { error: string }).error;
        } else {
          stamped.display = value === null || value === undefined ? "" : String(value);
        }
        cells ??= { ...sheet.cells };
        cells[key] = stamped;
        stampedKeys.push(key);
        // The server holds this value already (a tile's formula it wrote, a
        // body formula the body save carries): what was read is the stamp.
        if (asked === null) {
          const tile = state?.tiles.get(Math.floor(parseKey(key)[0] / state.entry.tileRows));
          if (tile?.baseline.has(key)) tile.baseline.set(key, stamped);
          else if (state?.loose.has(key)) state.loose.set(key, stamped);
        }
      }
      answers.clear();
      if (!cells) return sheet;
      changed = true;
      // The stamped keys are all that changed: what reads a sheet's cells
      // through their lineage (the stored results, the formula input) reads
      // those again, not 728,000 cells after a sort's save.
      recordCellMapEdit(sheet.cells, cells, stampedKeys);
      return { ...sheet, cells };
    });
    if (!changed) return sheets as S[];
    for (const sheet of out) this.seen.set(sheet.cells, this.generation);
    return out;
  }

  /** The last save's formulas went uncalculated: the server could not answer
   * them in time. Otherwise a save calculates what it feeds, whatever the
   * browser's session refuses of a workbook this large. */
  calculationUnanswered = false;

  /** What a save's answer says: its revisions, its tiles' and what it
   * calculated. */
  private acceptAnswer(answer: SaveAnswer): void {
    if (answer.recalculated !== undefined) this.calculationUnanswered = answer.recalculated === null;
    for (const revision of answer.revisions ?? [answer.revision]) this.ownRevisions.add(revision);
    this.ownRevisions.add(answer.revision);
    // Another member saved between this window's revision and this answer
    // (two windows typing at once): this window has not seen that save, so
    // its revision and its tiles' stay where they were and the next poll
    // reads the others' tiles and these again. Taking the answer's revision
    // passed over the other save for good: the two windows showed different
    // values in one cell, and neither showed the other's next cell (xlsx-rounds
    // night coedit R421-X12, R421-X59).
    let othersBetween = answer.revision - this.revision > 10_000;
    for (let revision = this.revision + 1; !othersBetween && revision <= answer.revision; revision += 1) {
      if (!this.ownRevisions.has(revision)) othersBetween = true;
    }
    // A formula this flush stored and the save calculated (a typed one,
    // or each formula a sort moved) needs no question of its own:
    // asking again, 64 a request, and stamping that answer as new input
    // sent all 2,122 formulas of a sorted 730,000-cell sheet back and
    // had the server calculate them a second time. Not one edited again
    // since (its newer text is still to be sent).
    let unsent: Set<string> | null = null;
    for (const [sheetId, values] of Object.entries(answer.recalculated ?? {})) {
      if (!this.answered.has(sheetId)) this.answered.set(sheetId, new Map());
      const asked = this.unanswered.get(sheetId);
      for (const [key, value] of Object.entries(values)) {
        this.answered.get(sheetId)!.set(key, { formula: null, value });
        if (!asked?.has(key)) continue;
        unsent ??= new Set([...this.pending, ...this.touched].map((edit) => `${edit.sheet}\u0000${edit.row}:${edit.col}`));
        if (!unsent.has(`${sheetId}\u0000${key}`)) asked.delete(key);
      }
    }
    if (othersBetween) {
      for (const [sheetId, tileIndex] of answer.tiles) {
        if (!this.rereadOwnTiles.has(sheetId)) this.rereadOwnTiles.set(sheetId, new Set());
        this.rereadOwnTiles.get(sheetId)!.add(tileIndex);
      }
    } else {
      this.revision = Math.max(this.revision, answer.revision);
      for (const [sheetId, tileIndex, revision] of answer.tiles) {
        const tile = this.sheets.get(sheetId)?.tiles.get(tileIndex);
        if (tile) tile.revision = Math.max(tile.revision, revision);
      }
    }
    for (const [sheetId, extent] of Object.entries(answer.extents ?? {})) {
      const state = this.sheets.get(sheetId);
      if (state) {
        state.entry.rows = Math.max(state.entry.rows, extent[0]);
        state.entry.cols = Math.max(state.entry.cols, extent[1]);
      }
    }
  }

  /** Send the pending edits (in parts of what one save takes). True when
   * every one landed. */
  /** What the parts of a save so far wrote, by sheet. */
  private readonly partsWritten = new Map<string, { top: number; bottom: number; left: number; right: number }>();

  private notePartWritten(edits: readonly TiledCellEdit[]): void {
    for (const { sheet, row, col } of edits) {
      const area = this.partsWritten.get(sheet);
      if (!area) {
        this.partsWritten.set(sheet, { top: row, bottom: row, left: col, right: col });
        continue;
      }
      if (row < area.top) area.top = row;
      if (row > area.bottom) area.bottom = row;
      if (col < area.left) area.left = col;
      if (col > area.right) area.right = col;
    }
  }

  /** The tiles held, by sheet (what a save's answer need cover). */
  private residentTiles(): Record<string, number[]> {
    const out: Record<string, number[]> = {};
    for (const [sheetId, state] of this.sheets) {
      const tiles = new Set(state.tiles.keys());
      // A cell the model holds outside them (a paste far below) too.
      for (const key of state.loose.keys()) tiles.add(Math.floor(parseKey(key)[0] / state.entry.tileRows));
      out[sheetId] = [...tiles];
    }
    return out;
  }

  async flush(): Promise<boolean> {
    while (this.inFlight) await this.inFlight;
    if (!this.pending.length && !this.touched.length && !this.rowSorts.length) return true;
    const send = async (): Promise<boolean> => {
      // A sort goes first, as its order: the edits after it are of the
      // positions it made. Refused, the range is saved cell by cell.
      while (this.rowSorts.length) {
        const sort = this.rowSorts[0];
        try {
          // The body cells changed go with it: the calculation reads them (a
          // save of their own after it would calculate all over again).
          const touched = this.touched.slice(0, EDITS_PER_SAVE);
          const answer = await apiPost<SaveAnswer>(`${base(this.pageId)}/row-sort`, {
            sheet: sort.sheet.id, r1: sort.r1, c1: sort.c1, c2: sort.c2, order: sort.order, base_revision: this.revision,
            resident: this.residentTiles(),
            ...(touched.length ? { touched } : {}),
          });
          this.rowSorts.shift();
          const sentTouched = new Set<TiledCellEdit>(touched);
          this.touched = this.touched.filter((edit) => !sentTouched.has(edit));
          this.acceptAnswer(answer);
          this.failure = null;
          this.failureStatus = null;
          this.failureCode = null;
        } catch (error) {
          this.rowSorts.shift();
          this.unsort(sort);
          // Not an answer at all (offline): the cells wait, as any edit does.
          if (!isApiRequestError(error) || error.status === 0) {
            this.failure = error instanceof Error ? error.message : String(error);
            this.failureStatus = isApiRequestError(error) ? error.status : null;
            this.failureCode = isApiRequestError(error) ? (error.code ?? null) : null;
            this.failures += 1;
            this.emit();
            return false;
          }
        }
      }
      // A save in parts goes tile by tile: each part rewrites whole tiles, and
      // parts in the order the edits were found (a sort's cleared cells after
      // its moved ones) had every part of a sorted 730,000-cell sheet rewrite
      // nearly every tile. One position has one pending edit, so the order
      // changes nothing that is saved.
      if (this.pending.length > EDITS_PER_SAVE) {
        this.pending.sort((a, b) => (a.sheet < b.sheet ? -1 : a.sheet > b.sheet ? 1 : a.row - b.row || a.col - b.col));
      }
      // A save in parts: each part's edits are written out as JSON while the
      // part before it is on its way, not after its answer (a sorted
      // 730,000-cell sheet: 0.3-0.5 s of encoding between each of its parts,
      
      let prepared: { count: number; bytes: Uint8Array<ArrayBuffer> } | null = null;
      while (this.pending.length || this.touched.length) {
        const batch = this.pending.slice(0, EDITS_PER_SAVE);
        // The body cells go with the last part: the calculation reads them,
        // and a part carrying them is calculated.
        const touched = this.pending.length > EDITS_PER_SAVE ? [] : this.touched.slice(0, EDITS_PER_SAVE);
        // One save in parts is calculated once, on its last part, over what
        // the parts before it wrote: the server's answer to each part was
        // only overwritten by the next.
        const more = !touched.length && this.pending.length > batch.length;
        const areas = more || this.partsWritten.size === 0 ? undefined : [...this.partsWritten].map(([sheet, area]) => ({ sheet, ...area }));
        try {
          const rest = {
            format: "editor",
            touched,
            base_revision: this.revision,
            ...(more ? { more } : {}),
            ...(areas ? { saved_areas: areas } : {}),
            // The values the save calculates are stamped on the tiles held
            // here only: the rest are read from the tiles when they are.
            ...(more ? {} : { resident: this.residentTiles() }),
          };
          const body = { ...rest, edits: batch };
          if (more || prepared) {
            const edits = prepared && prepared.count === batch.length ? prepared.bytes : jsonBytes(batch);
            withPreparedJsonBytes(body, joinedJsonBody(edits, rest));
          }
          prepared = null;
          const request = apiPatch<SaveAnswer>(`${base(this.pageId)}/cells`, body);
          const before = this.pending;
          // Written out once the part has left: the request awaits its gates
          // (CSRF, the notebook's mount) before it is sent, and the next
          // part's JSON written right here held this one back by its 0.2-0.4 s
          
          const next = more ? this.pending.slice(batch.length, batch.length + EDITS_PER_SAVE) : null;
          const nextWritten = next
            ? new Promise<Uint8Array<ArrayBuffer>>((resolve) => { setTimeout(() => resolve(jsonBytes(next)), 0); })
            : null;
          const answer = await request;
          const nextEdits = nextWritten ? await nextWritten : null;
          // Nothing collected while the part was on its way: the next part
          // is the one written out above.
          if (nextEdits && this.pending === before) {
            prepared = { count: Math.min(EDITS_PER_SAVE, this.pending.length - batch.length), bytes: nextEdits };
          }
          // The edits this part carried, by identity: an edit of the same
          // cell collected while the part was on its way replaced its entry
          // (`collect`), and counting off the front took that newer edit
          // instead. A cell made bold and then yellow kept only the bold
          
          const sent = new Set<TiledCellEdit>(batch);
          const sentTouched = new Set<TiledCellEdit>(touched);
          this.pending = this.pending.filter((edit) => !sent.has(edit));
          this.touched = this.touched.filter((edit) => !sentTouched.has(edit));
          if (more) this.notePartWritten(batch);
          else this.partsWritten.clear();
          // Only a save's last part is calculated.
          this.acceptAnswer(more ? { ...answer, recalculated: answer.recalculated ?? undefined } : answer);
          this.failure = null;
          this.failureStatus = null;
          this.failureCode = null;
        } catch (error) {
          // What landed stays landed (each part is one save); the rest stays
          // pending for the next flush.
          this.failure = error instanceof Error ? error.message : String(error);
          this.failureStatus = isApiRequestError(error) ? error.status : null;
          this.failureCode = isApiRequestError(error) ? (error.code ?? null) : null;
          this.failures += 1;
          this.emit();
          return false;
        }
      }
      this.emit();
      return true;
    };
    this.inFlight = send();
    try {
      return await this.inFlight;
    } finally {
      this.inFlight = null;
    }
  }

  /**
   * Another member's saves since this window's revision: the resident tiles
   * they changed are read again (`apply` then brings them in; this window's
   * own unsaved edits stay). False when too much changed to say (reload).
   */
  async poll(signal?: AbortSignal): Promise<boolean> {
    if (this.inFlight || this.pending.length || this.rowSorts.length) return true;
    const answer = await apiGet<ChangesAnswer>(`${base(this.pageId)}/changes?since=${this.revision}`, { signal });
    const stale = new Map<string, number[]>();
    if (answer.reload) {
      // Too much changed to list (a version restored): every resident tile
      // is read again.
      for (const [sheetId, state] of this.sheets) stale.set(sheetId, [...state.tiles.keys()]);
    }
    const reread = new Map(this.rereadOwnTiles);
    this.rereadOwnTiles.clear();
    for (const [sheetId, tiles] of answer.reload ? [] : reread) {
      const held = [...tiles].filter((tileIndex) => this.sheets.get(sheetId)?.tiles.has(tileIndex));
      if (held.length) stale.set(sheetId, held);
    }
    for (const [sheetId, tileIndex, revision] of answer.reload ? [] : answer.tiles) {
      if (stale.get(sheetId)?.includes(tileIndex)) continue;
      if (this.ownRevisions.has(revision)) continue;
      const tile = this.sheets.get(sheetId)?.tiles.get(tileIndex);
      if (!tile || tile.revision >= revision) continue;
      if (!stale.has(sheetId)) stale.set(sheetId, []);
      stale.get(sheetId)!.push(tileIndex);
    }
    try {
      for (const [sheetId, tiles] of stale) {
        for (let start = 0; start < tiles.length; start += TILES_PER_READ) {
          await this.read(sheetId, tiles.slice(start, start + TILES_PER_READ), signal, true);
        }
      }
    } catch (error) {
      // Not read: asked again at the next poll.
      for (const [sheetId, tiles] of reread) {
        if (!this.rereadOwnTiles.has(sheetId)) this.rereadOwnTiles.set(sheetId, new Set());
        for (const tileIndex of tiles) this.rereadOwnTiles.get(sheetId)!.add(tileIndex);
      }
      throw error;
    }
    this.revision = Math.max(this.revision, answer.revision);
    return true;
  }
}

export interface TiledFindOptions {
  caseSensitive?: boolean;
  wholeCell?: boolean;
  formulas?: boolean;
}

/** Cells holding `query` across the workbook's tiles, by sheet id, as the
 * server searches them (every tile, not only those read here): up to
 * `limit` matches. */
/** What a conditional format takes over its whole range, as the server read
 * it over the stored cells (`hydrate.range_statistics`). */
export interface TiledRangeStatistics {
  extent: { min: number; max: number } | null;
  numbers: number[] | null;
  count: number;
  mean: number | null;
  stdDev: number;
  counts: Record<string, number> | null;
}

/** The statistics of one rule's areas (`r1,c1,r2,c2` rectangles), or null
 * when the server cannot read that much at once. */
export async function tiledRangeStatistics(
  pageId: string,
  sheetId: string,
  areas: readonly { r1: number; c1: number; r2: number; c2: number }[],
  signal?: AbortSignal,
): Promise<TiledRangeStatistics | null> {
  const query = new URLSearchParams({
    sheet: sheetId,
    areas: areas.map((area) => `${area.r1},${area.c1},${area.r2},${area.c2}`).join(";"),
  });
  const answer = await apiGet<{ statistics: TiledRangeStatistics | null }>(
    `/api/v1/pages/${encodeURIComponent(pageId)}/workbook-tiles/range-stats?${query.toString()}`,
    { signal },
  );
  return answer?.statistics ?? null;
}

export async function findInTiles(
  pageId: string,
  query: string,
  options: TiledFindOptions,
  signal?: AbortSignal,
  limit = 2_000,
): Promise<Map<string, string[]>> {
  const out = new Map<string, string[]>();
  let sheet: string | null = null;
  let afterRow = -1;
  let afterCol = -1;
  let total = 0;
  const seen = new Set<string>();
  for (let page = 0; page < 20 && total < limit; page += 1) {
    const params = new URLSearchParams({
      q: query,
      match_case: String(Boolean(options.caseSensitive)),
      whole_cell: String(Boolean(options.wholeCell)),
      formulas: String(Boolean(options.formulas)),
      after_row: String(afterRow),
      after_col: String(afterCol),
    });
    if (sheet) params.set("sheet", sheet);
    const answer = await apiGet<{
      matches: { sheet: string; row: number; col: number }[];
      done: boolean;
      cursor: { sheet: string; tile: number } | null;
    }>(`${base(pageId)}/find?${params}`, { signal });
    for (const match of answer.matches) {
      const key = `${match.sheet}\u0000${match.row}:${match.col}`;
      if (seen.has(key)) continue;
      seen.add(key);
      if (!out.has(match.sheet)) out.set(match.sheet, []);
      out.get(match.sheet)!.push(`${match.row}:${match.col}`);
      total += 1;
    }
    if (answer.done || !answer.matches.length) break;
    const last = answer.matches[answer.matches.length - 1];
    sheet = last.sheet;
    afterRow = last.row;
    afterCol = last.col;
  }
  return out;
}

/** The opening of a page's structured tiled workbook, or null. */
export async function fetchStructuredTiledWorkbook(pageId: string): Promise<StructuredTiledOpening | null> {
  const { fetchTiledWorkbook } = await import("./tiled-workbook");
  const opening = await fetchTiledWorkbook(pageId);
  return isStructuredTiledOpening(opening) ? opening : null;
}

/** The tiled workbook marker of a page body, when the body is a structured
 * one's (the server's `TILED_SCHEMA_V2`). */
export const STRUCTURED_TILED_SCHEMA = "mrnote.tiled-workbook.v2";

export function structuredTiledMarker(content: unknown): Record<string, unknown> | null {
  if (!content || typeof content !== "object") return null;
  const marker = (content as { tiledWorkbook?: unknown }).tiledWorkbook;
  return marker && typeof marker === "object" && (marker as { schema?: unknown }).schema === STRUCTURED_TILED_SCHEMA
    ? (marker as Record<string, unknown>)
    : null;
}

// ── calculation over the cells read ─────────────────────────────────────────

/** A reference a formula reads: rows `top..bottom` of a sheet (by name, or
 * the formula's own when absent); `dynamic` when what it reads cannot be
 * told from its text (a defined name, INDIRECT, OFFSET). */
interface FormulaReach {
  sheet: string | null;
  top: number;
  bottom: number;
  dynamic: boolean;
}

const REFERENCE =
  /(?:'((?:[^']|'')+)'!|([A-Za-z_À-￿][\w.À-￿]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})(?![\w(])/g;
const DYNAMIC_FUNCTIONS = /\b(?:INDIRECT|OFFSET|INDEX)\s*\(/i;
// A name token: no character of [\w.$'!"#] just before it, checked in code
// (matchAllBehind) since Safari before 16.4 cannot compile a lookbehind.
const NAME_TOKEN = /([A-Za-z_\\À-￿][\w.À-￿]*)(?![\w(!])/g;
const NAME_TOKEN_BEHIND = { char: /[\w.$'!"#]/, not: true } as const;
const EXCEL_MAX_ROWS = 1_048_576;
const reachCache = new Map<string, FormulaReach[]>();

export function formulaReach(formula: string): FormulaReach[] {
  const known = reachCache.get(formula);
  if (known) return known;
  const body = formula.replace(/"(?:[^"]|"")*"/g, '""');
  const out: FormulaReach[] = [];
  const references: [number, number][] = [];
  for (const match of body.matchAll(REFERENCE)) {
    const sheet = match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2] ?? null;
    const reference = match[3].replace(/\$/g, "");
    const [first, second = first] = reference.split(":");
    let top = 0;
    let bottom = EXCEL_MAX_ROWS - 1;
    if (/^\d+$/.test(first)) {
      top = Math.min(Number(first), Number(second)) - 1;
      bottom = Math.max(Number(first), Number(second)) - 1;
    } else if (/\d/.test(first)) {
      const a = Number(first.replace(/^[A-Za-z]+/, "")) - 1;
      const b = Number(second.replace(/^[A-Za-z]+/, "")) - 1;
      top = Math.min(a, b);
      bottom = Math.max(a, b);
    }
    out.push({ sheet, top, bottom, dynamic: false });
    references.push([match.index ?? 0, (match.index ?? 0) + match[0].length]);
  }
  // A name the formula reads (not a function, not a reference, not TRUE).
  const outside = body.replace(REFERENCE, " ");
  let dynamic = DYNAMIC_FUNCTIONS.test(body);
  for (const match of matchAllBehind(outside, NAME_TOKEN, NAME_TOKEN_BEHIND)) {
    const token = match[1].toUpperCase();
    if (token === "TRUE" || token === "FALSE" || /^_XLFN\./.test(token)) continue;
    dynamic = true;
    break;
  }
  if (dynamic) out.push({ sheet: null, top: 0, bottom: EXCEL_MAX_ROWS - 1, dynamic: true });
  if (reachCache.size > 20_000) reachCache.clear();
  reachCache.set(formula, out);
  return out;
}

export interface TrustBinding {
  readonly sheet: { readonly id: string; readonly name: string };
  readonly input: { readonly name: string; getRawCell(row: number, col: number): string | null };
}

export interface TiledCalculationTrust {
  /** The session's results where every cell a formula reads is in the model
   * here; elsewhere the value Excel last calculated (`stored`). */
  trust(
    results: ReadonlyMap<string, unknown>,
    bindings: readonly TrustBinding[],
    stored: ReadonlyMap<string, unknown>,
  ): ReadonlyMap<string, unknown>;
}

export function trustTiledResults(
  tiled: TiledSheets,
  results: ReadonlyMap<string, unknown>,
  bindings: readonly TrustBinding[],
  stored: ReadonlyMap<string, unknown>,
): ReadonlyMap<string, unknown> {
  const partial = bindings.filter((binding) => tiled.isTiledSheet(binding.sheet.id) && !tiled.isWhollyResident(binding.sheet.id));
  if (!partial.length) return results;
  const byName = new Map(bindings.map((binding) => [binding.sheet.name.toLowerCase(), binding.sheet.id]));
  const readable = (sheetId: string, top: number, bottom: number): boolean => {
    if (!tiled.isTiledSheet(sheetId) || tiled.isWhollyResident(sheetId)) return true;
    const entry = tiled.entry(sheetId);
    if (!entry) return true;
    const last = Math.min(bottom, Math.max(0, entry.rows - 1));
    for (const tile of tiled.tilesFor(sheetId, top, last)) if (!tiled.isApplied(sheetId, tile)) return false;
    return true;
  };
  let out: Map<string, unknown> | null = null;
  for (const binding of bindings) {
    const prefix = `${binding.input.name.toLowerCase()}\u0000`;
    for (const key of results.keys()) {
      if (!key.startsWith(prefix)) continue;
      const [row, col] = key.slice(prefix.length).split(":").map(Number);
      const raw = binding.input.getRawCell(row, col);
      if (!raw || !raw.trimStart().startsWith("=")) continue;
      let trusted = true;
      for (const reach of formulaReach(raw)) {
        const sheetId = reach.dynamic ? null : reach.sheet === null ? binding.sheet.id : byName.get(reach.sheet.toLowerCase()) ?? null;
        if (reach.dynamic) {
          trusted = false;
          break;
        }
        if (sheetId === null) continue;
        if (!readable(sheetId, reach.top, reach.bottom)) {
          trusted = false;
          break;
        }
      }
      if (trusted) continue;
      out ??= new Map(results);
      if (stored.has(key)) out.set(key, stored.get(key));
      else out.delete(key);
    }
  }
  return out ?? results;
}






const FIRST_DRAW_ROWS = 64;
/** A sheet read whole has its tiles read before it is drawn up to this many;
 * past it the first screen's are, and the rest follow it in. */
const FIRST_DRAW_TILES = 96;





const FIRST_DRAW_WHOLE_CELLS = 20_000;
const prepared = new WeakMap<object, TiledSheets>();

/**
 * Read what a workbook first shows, before its editor mounts. Its grid was
 * rendered from the body alone and again, whole, when the tiles landed: two
 * seconds of an open of 100,000 cells spent drawing the same screen twice.
 */
export async function prepareTiledSheets(
  opening: StructuredTiledOpening,
  activeSheetId?: string,
  options?: { beforeMountKnown?: boolean },
): Promise<void> {
  // Asked again for the same opening (a read ahead of the window, then the
  // window): the same adapter, whose reads under way are joined.
  const tiled = prepared.get(opening) ?? new TiledSheets(opening);
  prepared.set(opening, tiled);
  const sheetId = activeSheetId && tiled.isTiledSheet(activeSheetId) ? activeSheetId : tiled.sheetIds()[0];
  if (!sheetId) return;
  tiled.readsBeforeMountKnown = options?.beforeMountKnown === true;
  // The first screen's tiles the manifest brought: resident at once, so only
  // what they leave out is read.
  const brought = opening.first_tiles;
  if (brought && brought.sheet === sheetId && tiled.isTiledSheet(sheetId)) tiled.accept(sheetId, brought);
  const all = tiled.allTilesOf(sheetId);
  const first = tiled.ensure(
    sheetId,
    tiled.readsWhole(sheetId) && all.length <= FIRST_DRAW_TILES && tiled.cellsOf(sheetId) <= FIRST_DRAW_WHOLE_CELLS
      ? all
      : tiled.tilesFor(sheetId, 0, FIRST_DRAW_ROWS - 1),
  );
  tiled.readsBeforeMountKnown = false;
  await first;
}

/** The adapter `prepareTiledSheets` read into. Asked while rendering, so it
 * stays until the editor has mounted: a render that suspended on the editor's
 * code is thrown away, and the next one asks again. */
export function preparedTiledSheets(opening: StructuredTiledOpening): TiledSheets | null {
  return prepared.get(opening) ?? null;
}

/** The editor mounted with it: a later editor of this opening reads afresh. */
export function forgetPreparedTiledSheets(opening: StructuredTiledOpening): void {
  prepared.delete(opening);
}

const jsonEncoder = new TextEncoder();

/** `value` as the UTF-8 bytes of its JSON. */
function jsonBytes(value: unknown): Uint8Array<ArrayBuffer> {
  return jsonEncoder.encode(JSON.stringify(value));
}

/** The JSON of `{ ...rest, edits }` with `edits` already written out. */
function joinedJsonBody(edits: Uint8Array<ArrayBuffer>, rest: Record<string, unknown>): Uint8Array<ArrayBuffer> {
  const tail = jsonBytes(rest);
  // `{"edits":` + edits + `,` + the rest's fields (its own `{` dropped)
  const head = jsonBytes("edits");
  const out = new Uint8Array(1 + head.length + 1 + edits.length + 1 + tail.length - 1);
  let at = 0;
  out[at++] = 0x7b; // {
  out.set(head, at); at += head.length;
  out[at++] = 0x3a; // :
  out.set(edits, at); at += edits.length;
  if (tail.length > 2) {
    out[at++] = 0x2c; // ,
    out.set(tail.subarray(1), at); at += tail.length - 1;
  } else {
    out[at++] = 0x7d; // }
  }
  return out.subarray(0, at);
}

