/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a save stamps on a workbook's cells and charts: each formula's and
 * formatted cell's on-screen text (`display`) and typed result
 * (`chartValueCache`), read with the same evaluator and spill overlay the
 * grid draws with, and every chart's model refreshed from those cells. Moved
 * out of SpreadsheetWindow so the window's save and the spreadsheet editor
 * commands (lib/editor-commands/spreadsheet) recalculate with the same code;
 * behaviour is unchanged. The rows an AutoFilter hides are the caller's to
 * give (`filterHiddenRows`); without them no row counts as filtered.
 */

import { cellMapRowPermutation, nearestKnownCellMap, permutedCellKeys, permutedSpanOf, recordCellMapEdit } from "./cell-map-lineage";
import { SPREADSHEET_DEEP_FORMULA_CHAINS } from "../feature-flags";
import type { BackgroundCalculationSheetContext } from "./background-calculation-protocol";
import type { CellMap, SpreadsheetCell } from "./cell-buttons";
import {
  normalizeSpreadsheetPivotChartSource,
  spreadsheetChartWithStoredGaps,
  type SpreadsheetPivotChartSource,
} from "./chart";
import {
  reconcileSpreadsheetChartV2,
  spreadsheetChartLiveFromStoredCells,
  type SpreadsheetChartLike,
  type SpreadsheetChartStoredSheet,
  type SpreadsheetChartV2Live,
} from "./chart-v2-model";
import { chartVisiblePointIndices, projectChartVisibilityRows } from "./chart-visibility";
import { matchingChartValueCache, type ChartValueCache } from "./chart-value-cache";
import { normalizeDataTables, normalizeTableColumnNames, type DataTable } from "./data-table";
import { cellPictureEntries } from "./cell-picture";
import { cellPhoneticReading, readCellPhonetic } from "./cell-phonetic";
import { keptFormulaResultEntries, keptFormulaResultOption } from "./kept-formula-results";
import { displayPrecisionEntries, displayPrecisionOption } from "./display-precision";
import { spreadsheetCellIsText } from "./entry-parse";
import { recoverDeepChain } from "./formula-dependencies";
import {
  columnToLetters,
  columnWidthOption,
  computeSpillOverlay,
  ERRORS,
  formatCellValue,
  isFormulaError,
  isKnownFunction,
  legacyArrayExtentOption,
  makeTableRefResolver,
  parseCellRef,
  FormulaResultUnavailableError,
  type CellValue,
  type ColumnWidthReading,
  type EngineValue,
  type SheetEvaluator,
  type SheetEvaluatorOptions,
  type SpillOverlay,
  type TableDefinition,
  type Workbook,
  type WorkbookSheetInput,
} from "./formula-engine";
import { normalizeManualHidden } from "./manual-hidden";
import { coveredCells, normalizeMerges } from "./merge-cells";
import { normalizeNamedRanges, resolveName as resolveNamedRange, type NamedRange } from "./named-ranges";
import {
  applyCellNumberFormat,
  applyTextFormat,
  formatGeneralNumberWith,
  type NumberFormatSeparators,
} from "./number-format";
import { hiddenRowsFromGroups, normalizeOutlineGroups } from "./outline-groups";
import { isCellContentHidden } from "./protection";
import { worksheetFormulaResultKey } from "./calculation-mode";
import { storedFormulaResult } from "./stored-formula-results";
import { createValueStoreWorkbook } from "./value-store";
import { isWhatIfTableFormula, whatIfTableBlocks, whatIfTableReader } from "./what-if-table-definition";

/** The worksheet fields a recalculation reads. */
export interface DisplayCacheSheet {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  cells: CellMap;
  merges?: unknown;
  namedRanges?: unknown;
  outlineGroups?: unknown;
  manuallyHiddenRows?: unknown;
  dataTables?: unknown;
  /** What-If Data Tables (what-if-table-definition). */
  whatIfTables?: unknown;
  /** Sheet protection: with it, a Hidden cell's formula is not displayable. */
  protected?: unknown;
}

/** The worksheet fields a chart model refresh reads. */
export interface ChartModelSheet {
  id: string;
  name: string;
  cells: CellMap;
  charts?: readonly object[];
  objects?: readonly { type?: unknown }[];
}

/** The rows a sheet's filters hide, given the merge-covered cells and names. */
export type FilterHiddenRows<S> = (
  sheet: S,
  covered: ReadonlySet<string>,
  namedRanges: NamedRange[],
  rowCount: number,
  columnCount: number,
) => Set<number>;

/**
 * Whether a stored cell is a live formula, by the engine's own rule: its value
 * starts with "=" and the cell is not typed as text (a quote prefix or a Text
 * format entry). A stored `formula` field says so too; a cell saved with only
 * `value: "=1+1"` is still a formula, and a quote-prefixed "=A2" never is.
 */
export function spreadsheetCellIsFormula(cell: SpreadsheetCell | undefined): boolean {
  if (!cell) return false;
  if (cell.formula) return true;
  return (
    typeof cell.value === "string" &&
    cell.value.trim().startsWith("=") &&
    !spreadsheetCellIsText(cell)
  );
}

/**
 * Whether a cell is in the way of content running into it, a dynamic array's
 * spill or a long text's overflow alike: it is covered by a merge or holds a
 * value of its own. Any stored character counts, spaces included: Excel
 * reports #SPILL! when the range holds "a cell that contains a space", and a
 * neighbour of spaces stops text overflow, while a truly blank cell, formatted
 * or not, lets both through.
 */
export function spillTargetOccupied(
  cells: CellMap,
  covered: ReadonlySet<string>,
  row: number,
  col: number,
  mergeAnchors?: ReadonlySet<string>,
): boolean {
  const key = `${row}:${col}`;
  // The blank anchor of a merge is as much the merge as its covered cells:
  // an array that would run into it is #SPILL! (live audit B192).
  if (covered.has(key) || mergeAnchors?.has(key)) return true;
  const cell = cells[key];
  const v = cell?.value;
  if (typeof v === "string" && v !== "") return true;
  // The empty string a workbook stored (t="s" naming <si><t/></si>) is a
  // value, not a blank: Excel runs neither text nor an array into it (X-031).
  return v === "" && cell?.stringValue === true && !cell.formula;
}

/**
 * Materialize the sheet's dynamic-array spill overlay (#SPILL! guard) for a
 * given evaluator. A target counts as occupied (blocks a spill) when it
 * carries its own value (spillTargetOccupied) or is covered by a merge; a merged-away
 * anchor never spills. The overlap/claim/no-partial-write semantics live in
 * the engine's computeSpillOverlay; this only wires the sheet model into it.
 */
export function buildSpillOverlay(
  cells: CellMap,
  covered: ReadonlySet<string>,
  evaluator: SheetEvaluator,
  candidates?: readonly string[],
  mergeAnchors?: ReadonlySet<string>,
  /** The sheet's tables: a formula inside one cannot spill (B323). */
  tables?: readonly DataTable[],
): SpillOverlay {
  const anchors: Array<{ row: number; col: number }> = [];
  // Indexed candidates name every stored formula key; the current cell decides.
  const visit = (key: string) => {
    if (!cells[key]?.formula || covered.has(key)) return;
    const [r, c] = key.split(":");
    anchors.push({ row: Number(r), col: Number(c) });
  };
  if (candidates) for (const key of candidates) visit(key);
  else for (const key in cells) if (Object.prototype.hasOwnProperty.call(cells, key)) visit(key);
  return computeSpillOverlay({
    anchors,
    getArray: (row, col) => evaluator.getArray(row, col),
    isOccupied: (row, col) => spillTargetOccupied(cells, covered, row, col, mergeAnchors),
    ...(tables && tables.length > 0
      ? {
          isAnchorBlocked: (row: number, col: number) =>
            tables.some(
              ({ range, totalsRow }) =>
                row >= range.r1 &&
                row <= range.r2 + (totalsRow ? 1 : 0) &&
                col >= range.c1 &&
                col <= range.c2,
            ),
        }
      : {}),
  });
}

/**
 * Whether a cell without a formula of its own is a What-If Data Table's
 * result, which shows what the calculation answers for its =TABLE(...)
 * (the evaluator's input reads it so; what-if-table-definition) rather than
 * the last value it stores.
 */
function computedWhatIfTableCell(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
): boolean {
  if (cell?.formula) return false;
  return isWhatIfTableFormula(evaluator.getFormula(row, col));
}

/**
 * Display string for a cell under a given evaluator + spill overlay: the
 * computed formula result (a blocked spill anchor reads #SPILL!), a spilled
 * element for an empty cell inside an array formula's footprint, and/or the
 * applied number format. This is the non-editing core of cellDisplayValue,
 * shared with the AutoFilter's bootstrap resolution so filter matching sees
 * exactly what the grid renders.
 *
 * `drawing.separators` is for a viewer's own drawing only (the grid): a
 * number is then drawn with that region's decimal and thousands separators,
 * as Excel draws with the system's. Without it the text is the code's own
 * (5,000.00), which is what the display cache, filters, exports and the AI
 * read, so a file reads the same for every viewer.
 */
export function resolveCellDisplay(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
  numberFormatOverride?: string,
  drawing?: { separators?: NumberFormatSeparators },
): string {
  const key = `${row}:${col}`;
  const separators = drawing?.separators;
  // The evaluator's workbook decides how a serial reads as a date.
  const dateSystem = separators
    ? { date1904: evaluator.date1904, separators }
    : { date1904: evaluator.date1904 };
  const fmt =
    numberFormatOverride !== undefined
      ? numberFormatOverride
      : typeof cell?.style?.numberFormat === "string"
        ? cell.style.numberFormat.trim()
        : "";
  if (cell?.formula || computedWhatIfTableCell(cell, row, col, evaluator)) {
    // A spilled array anchor whose footprint is blocked shows #SPILL! (Excel
    // parity) rather than its top-left value.
    if (overlay.blocked.has(key)) return ERRORS.SPILL.code;
    let computed = evaluator.getComputed(row, col);
    // A deep dependency chain can overflow the evaluator's recursion and cache
    // that error before the grid sees it. Plan against the evaluator's exact
    // registered input, then atomically evict/warm/recompute once. A genuine or
    // incomplete error stays an error; successful reads never pay this cost.
    if (SPREADSHEET_DEEP_FORMULA_CHAINS && isFormulaError(computed)) {
      recoverDeepChain(evaluator, { row, col }, computed);
      computed = evaluator.getComputed(row, col);
    }
    // A circular reference with iterative calculation off is no error value
    // in Excel: the cells keep the value they last had (the file's), a new
    // cycle reads 0, and only the status bar warns (SY011; X-307). The
    // evaluator's #CIRCULAR! stays the marker Error Checking lists.
    if (isFormulaError(computed) && computed.code === ERRORS.CIRCULAR.code) {
      const kept = cell?.display;
      if (typeof kept === "string" && kept !== "" && !CELL_ERROR_CODES.has(kept.trim().toUpperCase())) return kept;
      return fmt ? applyCellNumberFormat(0, fmt, dateSystem) : "0";
    }
    if (fmt && typeof computed === "number") {
      return applyCellNumberFormat(computed, fmt, dateSystem);
    }
    // A text or logical result goes through the format's text section, as
    // in Excel: under `"文字:"@` ="甲乙" shows 文字:甲乙, ="" shows 文字:
    // and TRUE shows 文字:TRUE (live audit S-088).
    if (fmt && (typeof computed === "string" || typeof computed === "boolean")) {
      return applyTextFormat(evaluator.getDisplay(row, col), fmt);
    }
    if (separators && typeof computed === "number" && Number.isFinite(computed)) {
      return formatGeneralNumberWith(computed, separators);
    }
    return evaluator.getDisplay(row, col);
  }
  // A non-formula cell that sits inside an array formula's spill range (and has
  // no value of its own) displays the spilled element. This is display-only —
  // the cell is never persisted, so save/load is unaffected.
  if (!cell?.value) {
    const spill = overlay.spilled.get(key);
    if (spill !== undefined) {
      if (fmt && typeof spill === "number")
        return applyCellNumberFormat(spill, fmt, dateSystem);
      if (fmt && (typeof spill === "string" || typeof spill === "boolean"))
        return applyTextFormat(formatCellValue(spill), fmt);
      if (separators && typeof spill === "number" && Number.isFinite(spill)) {
        return formatGeneralNumberWith(spill, separators);
      }
      return formatCellValue(spill);
    }
  }
  // A number format shapes numbers. A constant the cell calls text (typed
  // with a leading apostrophe, or entered in a Text cell) stays the text it
  // is: '007 under Percentage shows 007, as ISTEXT says. A stored number
  // under an apostrophe format (an xf's quotePrefix over a numeric value) is
  // still a number, and Excel shows it formatted: R87-X14's populations
  // "4, 867" (#,##0) and ages "43.1" printed raw, "4867", and as ####.
  const numberIdentity = (candidate: SpreadsheetCell) => candidate.style?.quotePrefix !== true || candidate.kind === "number";
  if (
    fmt &&
    cell?.value &&
    cell.value.trim() !== "" &&
    cell.kind !== "text" &&
    numberIdentity(cell)
  ) {
    const num = Number(cell.value);
    if (Number.isFinite(num)) return applyCellNumberFormat(num, fmt, dateSystem);
  }
  // Any other constant is text (or a logical) and wears the format's text
  // section; a typed error stays the error.
  if (fmt && cell?.value && !CELL_ERROR_CODES.has(cell.value.trim().toUpperCase())) {
    return applyTextFormat(cell.value, fmt);
  }
  // A number stored in an exponent spelling of its own (an import writes
  // 3e-104) is shown as General writes it: Excel shows 3E-104 (X-215).
  if (
    !fmt &&
    cell?.value &&
    cell.kind !== "text" &&
    cell.stringValue !== true &&
    numberIdentity(cell) &&
    STORED_EXPONENT_NUMBER.test(cell.value)
  ) {
    const num = Number(cell.value);
    if (Number.isFinite(num)) return formatGeneralNumberWith(num, separators);
  }
  // A plain number constant in General keeps its digits as stored; only its
  // decimal point is the viewer's.
  if (
    separators &&
    !fmt &&
    cell?.value &&
    cell.kind !== "text" &&
    cell.stringValue !== true &&
    numberIdentity(cell) &&
    PLAIN_DECIMAL_NUMBER.test(cell.value)
  ) {
    return cell.value.trim().replace(".", separators.decimal);
  }
  return cell?.value || "";
}

const PLAIN_DECIMAL_NUMBER = /^\s*-?\d+\.\d+\s*$/;

const STORED_EXPONENT_NUMBER = /^\s*[-+]?(\d+\.?\d*|\.\d+)[eE][-+]?\d+\s*$/;

const CELL_ERROR_CODES: ReadonlySet<string> = new Set(Object.values(ERRORS).map((error) => error.code));

/**
 * Invariant value channel for native chart series. Excel charts bind to the
 * computed cell value, not to the number-formatted text painted in the grid.
 * Keep spill/error behavior aligned with resolveCellDisplay, but deliberately
 * skip applyNumberFormat so currency and percent cells remain numeric.
 */
export function resolveCellChartScalar(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
): CellValue {
  const key = `${row}:${col}`;
  if (cell?.formula || computedWhatIfTableCell(cell, row, col, evaluator)) {
    if (overlay.blocked.has(key)) return ERRORS.SPILL;
    let computed = evaluator.getComputed(row, col);
    if (SPREADSHEET_DEEP_FORMULA_CHAINS && isFormulaError(computed)) {
      recoverDeepChain(evaluator, { row, col }, computed);
      computed = evaluator.getComputed(row, col);
    }
    return computed;
  }
  if (!cell?.value) {
    const spill = overlay.spilled.get(key);
    if (spill !== undefined) {
      return spill;
    }
  }
  return cell?.value || "";
}

/**
 * The rows SUBTOTAL 101-111 and AGGREGATE's "ignore hidden rows" skip: the
 * outline-collapsed rows AND the rows hidden with Hide Rows, one set. Excel
 * treats both as hidden rows (Microsoft SUBTOTAL: 101-111 ignore rows hidden
 * by Hide Rows); a row hidden by a filter is a separate callback.
 */
export function calcManuallyHiddenRows(
  groupHiddenRows: ReadonlySet<number>,
  manuallyHiddenRows: ReadonlySet<number> | Parameters<typeof normalizeManualHidden>[0],
  rowCount?: number,
): Set<number> {
  const manual =
    manuallyHiddenRows instanceof Set
      ? manuallyHiddenRows
      : normalizeManualHidden(manuallyHiddenRows, rowCount ?? 0);
  if (manual.size === 0) return new Set(groupHiddenRows);
  return new Set([...groupHiddenRows, ...manual]);
}

/**
 * Build the formula engine's table registry ({@link TableDefinition}[]) from a
 * sheet's normalized data tables, reading each table's HEADER row texts for
 * column-name matching. `readRaw(row,col)` returns the raw cell text honoring
 * merge cover (the same reader the evaluator uses), so a header column name is
 * the literal text in the header cell. A table with a blank header cell gets an
 * empty name for that column (it simply won't match `[ColName]`, which is the
 * correct Excel behavior for an unnamed column). The structured-ref resolver is
 * built once from these defs via {@link makeTableRefResolver}.
 */
export function tableDefsForSheet(
  rawDataTables: unknown,
  bounds: { rowCount: number; columnCount: number },
  readRaw: (row: number, col: number) => string | null,
): TableDefinition[] {
  const tables = normalizeDataTables(rawDataTables, {
    rowCount: Number(bounds.rowCount) || 0,
    columnCount: Number(bounds.columnCount) || 0,
  });
  return tables.map((t) => {
    const { r1, c1, r2, c2 } = t.range;
    const width = c2 - c1 + 1;
    const columns =
      t.columns?.length === width
        ? t.columns.map((column) => column.name)
        : t.headerRow
          ? normalizeTableColumnNames(
              Array.from({ length: width }, (_, offset) =>
                (readRaw(r1, c1 + offset) ?? "").trim(),
              ),
            )
          : Array.from({ length: width }, (_, offset) => `Column${offset + 1}`);
    return {
      name: t.name,
      range: { r1, c1, r2, c2 },
      hasHeader: t.headerRow,
      hasTotals: t.totalsRow === true,
      columns,
    };
  });
}

const cellsWithEngineFacts = new WeakMap<CellMap, CellMap>();
const textCellKeysByCells = new WeakMap<CellMap, string[]>();

/**
 * The cells the per-cell engine options below read (a picture, furigana, a
 * kept result, a legacy array). A sheet's ordinary cells hold none of these,
 * and each option walking all of them was a tenth of the long task of opening
 * a 250,000-cell sheet: they read this instead, walked once per cell map.
 */
export function cellsCarryingEngineFacts(cells: CellMap): CellMap {
  const known = cellsWithEngineFacts.get(cells);
  if (known) return known;
  const carries = (cell: CellMap[string] | undefined): boolean => {
    const facts = cell as { picture?: unknown; phonetic?: unknown; keptResult?: unknown; arrayRef?: unknown } | undefined;
    return Boolean(facts && (facts.picture || facts.phonetic || facts.keptResult || facts.arrayRef));
  };
  const consider = (carrying: CellMap, key: string) => {
    if (carries(cells[key])) carrying[key] = cells[key];
    else delete carrying[key];
  };
  // A row sort of a map whose list is known: the same cells at new keys.
  const sorted = cellMapRowPermutation(cells);
  const sortedFrom = sorted && cellsWithEngineFacts.get(sorted.before);
  if (sorted && sortedFrom) {
    const moved: CellMap = {};
    for (const key of permutedCellKeys(Object.keys(sortedFrom), permutedSpanOf(sorted.permutation), cells, carries)) moved[key] = cells[key];
    cellsWithEngineFacts.set(cells, moved);
    return moved;
  }
  const prior = nearestKnownCellMap(cells, (before) => cellsWithEngineFacts.get(before) ?? (cellMapRowPermutation(before) ? cellsCarryingEngineFacts(before) : undefined));
  const carrying: CellMap = prior ? { ...prior.value } : {};
  for (const key of prior ? prior.keys : Object.keys(cells)) consider(carrying, key);
  cellsWithEngineFacts.set(cells, carrying);
  return carrying;
}

/**
 * The engine's reading of a sheet's in-cell pictures (Place in Cell): each
 * picture cell is the picture value, with the "row:col" key and Alt Text the
 * background calculation receives in its place. A merge-covered cell reads
 * as blank, as its characters do.
 */
export function cellPictureEngineOption(
  cells: CellMap,
  covered: ReadonlySet<string>,
): {
  entries: Array<[string, string]>;
  option: Pick<SheetEvaluatorOptions, "cellPicture">;
} {
  const entries = cellPictureEntries(cells).filter(([key]) => !covered.has(key));
  if (entries.length === 0) return { entries, option: {} };
  const alts = new Map(entries);
  return {
    entries,
    option: {
      cellPicture: (row: number, col: number) => {
        const alt = alts.get(`${row}:${col}`);
        return alt === undefined ? null : { alt };
      },
    },
  };
}

/**
 * The engine's reading of a sheet's furigana: each text cell with phonetic
 * runs reads, through PHONETIC, as {@link cellPhoneticReading}. Enumerated
 * (key, reading) so the background calculation receives the same data.
 */
export function cellPhoneticEngineOption(
  cells: CellMap,
  covered: ReadonlySet<string>,
): {
  entries: Array<[string, string]>;
  option: Pick<SheetEvaluatorOptions, "cellPhonetic">;
} {
  const entries: Array<[string, string]> = [];
  for (const key of Object.keys(cells)) {
    const cell = cells[key];
    if (!cell?.phonetic || covered.has(key)) continue;
    if (!readCellPhonetic(cell)?.runs) continue;
    entries.push([key, cellPhoneticReading(cell)]);
  }
  if (entries.length === 0) return { entries, option: {} };
  const readings = new Map(entries);
  return {
    entries,
    option: {
      cellPhonetic: (row: number, col: number) =>
        readings.get(`${row}:${col}`) ?? null,
    },
  };
}

/**
 * The formulas a sheet's protection withholds from display: with the sheet
 * protected, each formula cell formatted Hidden ({@link isCellContentHidden},
 * the formula bar's own rule), so FORMULATEXT of it is #N/A. Enumerated
 * "row:col" keys so the background calculation receives the same data;
 * nothing when the sheet is unprotected.
 */
export function formulaHiddenEngineOption(
  sheetProtected: boolean,
  cells: CellMap,
  covered: ReadonlySet<string>,
): {
  keys: string[];
  option: Pick<SheetEvaluatorOptions, "isFormulaHidden">;
} {
  const keys: string[] = [];
  if (!sheetProtected) return { keys, option: {} };
  for (const key of Object.keys(cells)) {
    const cell = cells[key];
    if (covered.has(key) || !spreadsheetCellIsFormula(cell)) continue;
    if (isCellContentHidden({ protected: true }, cell)) keys.push(key);
  }
  if (keys.length === 0) return { keys, option: {} };
  const hidden = new Set(keys);
  return {
    keys,
    option: { isFormulaHidden: (row: number, col: number) => hidden.has(`${row}:${col}`) },
  };
}

export const backgroundInputContexts = new WeakMap<WorkbookSheetInput, BackgroundCalculationSheetContext>();

/**
 * The legacy array formulas of a sheet (an imported `<f t="array" ref>`, the
 * anchor's `arrayRef`): each anchor's "row:col" key with its range's height
 * and width, so the engine lays the result over that range alone. Enumerated
 * so the same list travels to the background calculation worker.
 */
export function legacyArrayEngineOption(
  cells: CellMap,
  covered: ReadonlySet<string>,
): {
  entries: Array<readonly [string, number, number]>;
  option: Pick<SheetEvaluatorOptions, "legacyArrayExtent">;
} {
  const entries: Array<readonly [string, number, number]> = [];
  for (const key of Object.keys(cells)) {
    const ref = (cells[key] as { arrayRef?: unknown } | undefined)?.arrayRef;
    if (typeof ref !== "string" || !ref || covered.has(key) || !spreadsheetCellIsFormula(cells[key])) continue;
    const [from, to = from] = ref.replace(/\$/g, "").split(":");
    const start = parseCellRef(from);
    const end = parseCellRef(to);
    if (!start || !end) continue;
    const [row, col] = key.split(":").map(Number);
    // The range starts at its anchor; one that does not is not this cell's.
    if (Math.min(start.row, end.row) !== row || Math.min(start.col, end.col) !== col) continue;
    entries.push([key, Math.abs(end.row - start.row) + 1, Math.abs(end.col - start.col) + 1]);
  }
  return { entries, option: legacyArrayExtentOption(entries) };
}

/**
 * The "row:col" keys of the cells the host types as TEXT — an apostrophe-
 * prefixed entry, or one typed into a Text (@) cell. Enumerated rather than
 * answered live so the same typing can travel to the background calculation
 * worker, whose host only accepts resolvers it can serialize with their data.
 */
export function textCellKeys(cells: CellMap): string[] {
  const known = textCellKeysByCells.get(cells);
  if (known) return known;
  const isTextCell = (cell: CellMap[string] | undefined) => {
    // An empty cell is blank whatever its format, except the empty STRING a
    // workbook file stored (t="s" naming <si><t/></si>): that is a value, and
    // a formula reading it gets "" (Excel), which the engine reads off this
    // list.
    if (cell?.value === "" && !cell.formula && cell.stringValue !== true) return false;
    return spreadsheetCellIsText(cell);
  };
  const isText = (key: string) => isTextCell(cells[key]);
  // A row sort of a map whose list is known: the same cells at new keys.
  const sorted = cellMapRowPermutation(cells);
  const sortedFrom = sorted && textCellKeysByCells.get(sorted.before);
  if (sorted && sortedFrom) {
    const keys = permutedCellKeys(sortedFrom, permutedSpanOf(sorted.permutation), cells, isTextCell);
    textCellKeysByCells.set(cells, keys);
    return keys;
  }
  // Carried along a recorded edit: only the keys it changed are asked again.
  const prior = nearestKnownCellMap(cells, (before) => textCellKeysByCells.get(before) ?? (cellMapRowPermutation(before) ? textCellKeys(before) : undefined));
  let keys: string[];
  if (prior) {
    // A key that stays text stays where it is: the list is part of what the
    // background calculation compares.
    const changed = new Set(prior.keys);
    keys = prior.value.filter((key) => !changed.has(key) || isText(key));
    const listed = new Set(keys);
    for (const key of changed) if (!listed.has(key) && isText(key)) keys.push(key);
  } else {
    keys = Object.keys(cells).filter(isText);
  }
  textCellKeysByCells.set(cells, keys);
  return keys;
}

/** The text cells less a What-If Data Table's results: a result is its
 *  =TABLE(...) to the engine whatever its last value's type. */
export function whatIfTableTextCells(
  keys: string[],
  whatIf: { readonly isResult: (row: number, col: number) => boolean } | null,
): string[] {
  if (!whatIf || keys.length === 0) return keys;
  return keys.filter((key) => {
    const [row, col] = key.split(":").map(Number);
    return !whatIf.isResult(row, col);
  });
}

export function workbookSheetInputFor<S extends DisplayCacheSheet>(
  sheet: S,
  opts?: {
    // When given, compute this sheet's AutoFilter-hidden rows too (a bootstrap
    // evaluator pass per sheet). The serialize path passes this so an EXPORTED
    // SUBTOTAL's cached display matches the grid exactly; the live display memo
    // skips it for non-active sheets (perf), see formulaEvaluator.
    filterHiddenRows?: FilterHiddenRows<S>;
    /** The sheet's column widths in characters (column-width-reading), for
     *  CELL("width"); without them it reads Excel's default 8. */
    columnWidths?: ColumnWidthReading;
    /** The workbook's sheets still placeholders (deferred-sheets
     *  `pendingSheetNames`): a formula naming one keeps the file's value. */
    pendingSheetNames?: ReadonlySet<string>;
    /** The workbook sets precision as displayed (calcPr@fullPrecision="0"). */
    precisionAsDisplayed?: boolean;
  },
): WorkbookSheetInput {
  const merges = normalizeMerges(sheet.merges);
  const covered = coveredCells(merges);
  const namedRanges = normalizeNamedRanges(sheet.namedRanges);
  const rowCount = Number(sheet.rowCount) || 0;
  const columnCount = Number(sheet.columnCount) || 0;
  const groupHiddenRows = calcManuallyHiddenRows(
    hiddenRowsFromGroups(
      normalizeOutlineGroups(sheet.outlineGroups, { rowCount, columnCount }),
    ),
    sheet.manuallyHiddenRows,
    rowCount,
  );
  const filterHiddenRows = opts?.filterHiddenRows
    ? opts.filterHiddenRows(sheet, covered, namedRanges, rowCount, columnCount)
    : new Set<number>();
  const cells = sheet.cells;
  const readStored = (row: number, col: number): string | null =>
    covered.has(`${String(row)}:${String(col)}`)
      ? null
      : (cells[`${row}:${col}`]?.value ?? null);
  const tableDefs = tableDefsForSheet(
    sheet.dataTables,
    { rowCount, columnCount },
    readStored,
  );
  // A What-If Data Table's results read as its =TABLE(...) to every
  // calculation of the sheet (what-if-table-definition).
  const whatIf = whatIfTableReader(sheet.whatIfTables, { rowCount, columnCount }, readStored);
  const readRaw = whatIf?.readRaw ?? readStored;
  const textCells = whatIfTableTextCells(textCellKeys(cells), whatIf);
  const textCellKeySet = new Set(textCells);
  const carrying = cellsCarryingEngineFacts(cells);
  const pictureCells = cellPictureEngineOption(carrying, covered);
  const phoneticCells = cellPhoneticEngineOption(carrying, covered);
  // Formulas over external data keep the file's last value (X-079).
  const keptResults = keptFormulaResultEntries(
    carrying, covered, opts?.pendingSheetNames?.size ? { sheets: opts.pendingSheetNames, cells } : undefined);
  const formulaHidden = formulaHiddenEngineOption(sheet.protected === true, cells, covered);
  const legacyArrays = legacyArrayEngineOption(carrying, covered);
  const shownPrecision = opts?.precisionAsDisplayed ? displayPrecisionEntries(cells, covered) : [];
  const options: SheetEvaluatorOptions = {
    usedRange: { rowCount, columnCount },
    ...pictureCells.option,
    ...phoneticCells.option,
    ...keptFormulaResultOption(keptResults, readStored),
    ...formulaHidden.option,
    ...legacyArrays.option,
    ...displayPrecisionOption(shownPrecision),
    ...columnWidthOption(opts?.columnWidths),
    // Excel stores a cell's TYPE: an apostrophe-prefixed entry, and one typed
    // into a Text (@) cell, is a string even when it reads as a number. The
    // keys are enumerated (not read live) because the background calculation
    // host only accepts a resolver it can serialize alongside its data.
    ...(textCells.length > 0
      ? {
          isTextCell: (row: number, col: number) =>
            textCellKeySet.has(`${row}:${col}`),
        }
      : {}),
    ...(namedRanges.length > 0
      ? { resolveName: (name: string) => resolveNamedRange(name, namedRanges) }
      : {}),
    ...(filterHiddenRows.size > 0
      ? { isRowHiddenByFilter: (row: number) => filterHiddenRows.has(row) }
      : {}),
    ...(groupHiddenRows.size > 0
      ? { isRowManuallyHidden: (row: number) => groupHiddenRows.has(row) }
      : {}),
    ...(tableDefs.length > 0
      ? { resolveTableRef: makeTableRefResolver(tableDefs) }
      : {}),
    // A plain reference into a dynamic array's spill reads the element the
    // grid shows there, on this sheet as on the active one.
    spillAnchors: () => spillAnchorsOfCells(cells),
    isSpillTargetOccupied: (row: number, col: number) =>
      spillTargetOccupied(cells, covered, row, col),
    ...(whatIf ? { whatIfTables: whatIf.whatIfTables } : {}),
  };
  const input: WorkbookSheetInput = {
    name: typeof sheet.name === "string" && sheet.name ? sheet.name : sheet.id,
    getRawCell: readRaw,
    ...(Object.keys(options).length > 0 ? { options } : {}),
  };
  backgroundInputContexts.set(input, {
    ...(namedRanges.length ? { namedRanges } : {}),
    ...(tableDefs.length ? { tables: tableDefs } : {}),
    ...(whatIf ? { whatIfTables: whatIf.whatIfTables } : {}),
    ...(filterHiddenRows.size ? { filterHiddenRows: [...filterHiddenRows] } : {}),
    ...(groupHiddenRows.size ? { manuallyHiddenRows: [...groupHiddenRows] } : {}),
    ...(textCells.length ? { textCells } : {}),
    ...(pictureCells.entries.length ? { pictureCells: pictureCells.entries } : {}),
    ...(phoneticCells.entries.length ? { phoneticCells: phoneticCells.entries } : {}),
    ...(keptResults.length ? { keptResults } : {}),
    ...(formulaHidden.keys.length ? { formulaHiddenCells: formulaHidden.keys } : {}),
    ...(legacyArrays.entries.length ? { legacyArrays: legacyArrays.entries } : {}),
    ...(shownPrecision.length ? { displayPrecision: shownPrecision } : {}),
    ...(opts?.columnWidths ? { columnWidths: opts.columnWidths } : {}),
    ...(covered.size ? { coveredCells: [...covered] } : {}),
  });
  return input;
}

/**
 * Stamp the DISPLAY CACHE (Option A) onto a sheet's cells for serialization.
 *
 * For each cell that has a `formula` OR a `style.numberFormat`, compute the
 * on-screen string with the SAME resolver the grid renders through
 * (resolveCellDisplay over a fresh evaluator + spill overlay) and write it into
 * a read-only `display` field. Formula cells also receive a typed, unformatted
 * `chartValueCache` from the same evaluator/overlay, bound to canonical `value`.
 * Formatted display text is never a numeric chart-data authority.
 * The plain-text/CSV export (notebooks.py
 * `_cell_text`) prefers this over the raw `value` so a formula cell exports its
 * computed value ("30") and a formatted cell exports its formatted value
 * ("50.00%") — matching what the user sees, exactly like Excel/Sheets "Save as
 * CSV".
 *
 * Gating is deliberate: a plain text/number cell (no formula, no numberFormat)
 * has display === value, so caching nothing keeps it byte-stable (no `display`
 * key). The cache is rebuilt from CURRENT cells on EVERY call, so an edit to a
 * precedent cell refreshes a dependent formula cell's display on the next save
 * (never stamped once). A stale `display` on an incoming cell is always
 * overwritten/dropped here, so it can never go stale across a save.
 */
/** The keys the display cache can touch on one immutable cell map, from the
 * cell index. Every other cell is left byte-identical by the walk below, so
 * naming only these keys changes nothing but the work. */
export interface DisplayCacheCandidates {
  readonly displayCacheKeys: readonly string[];
  readonly formulaKeys: readonly string[];
}

/**
 * The sheets without the caches a save derives on cells (`display`,
 * `chartValueCache`): no part of what an edit wrote. A peer's save brings
 * them back on formula cells, and an Undo that weighed them read the pasted
 * formulas as the peer's and kept them (live audit S-163).
 */
const formulaCellKeysOfMap = new WeakMap<object, readonly string[]>();

/** The cells of a map that hold a formula, in the map's order; kept by map and
 * followed from the map before at the keys that changed. */
function formulaCellKeys(cells: CellMap): readonly string[] {
  const kept = formulaCellKeysOfMap.get(cells);
  if (kept) return kept;
  const near = nearestKnownCellMap(cells, (map) => formulaCellKeysOfMap.get(map));
  let keys: readonly string[];
  if (near) {
    const changed = new Set(near.keys);
    const known = new Set(near.value);
    keys = [
      ...near.value.filter((key) => !changed.has(key) || spreadsheetCellIsFormula(cells[key])),
      ...near.keys.filter((key) => !known.has(key) && spreadsheetCellIsFormula(cells[key])),
    ];
  } else {
    keys = Object.keys(cells).filter((key) => spreadsheetCellIsFormula(cells[key]));
  }
  formulaCellKeysOfMap.set(cells, keys);
  return keys;
}

/** Where a dynamic array may spill from: a walk of every cell at each commit's
 * calculation (12 ms on 100,000 cells), for the same few formulas. */
export function spillAnchorsOfCells(cells: CellMap): { row: number; col: number }[] {
  return formulaCellKeys(cells).map((key) => {
    const [row, col] = key.split(":").map(Number);
    return { row, col };
  });
}

export function sheetsWithoutDerivedCellCaches<S extends { cells: CellMap }>(sheets: readonly S[]): S[] {
  return sheets.map((sheet) => {
    const cells = cellsWithoutDerivedCaches(sheet.cells);
    return cells === sheet.cells ? sheet : { ...sheet, cells };
  });
}

function holdsDerivedCache(cell: CellMap[string] | undefined): boolean {
  return Boolean(cell && ("display" in cell || "chartValueCache" in cell));
}

function withoutDerivedCache(cell: CellMap[string]): CellMap[string] {
  const next = { ...cell };
  delete next.display;
  delete next.chartValueCache;
  return next;
}

const cellsWithoutCachesOfMap = new WeakMap<object, CellMap>();

/**
 * Each Undo asked this of three whole workbooks, 130 ms on 100,000 cells, and the
 * maps it asked of were an edit or two apart. The answer is kept by map, followed
 * from the map before at the keys that changed, and recorded as following from
 * that map's answer, so that comparing the two answers can read those keys alone.
 */
function cellsWithoutDerivedCaches(cells: CellMap): CellMap {
  const known = cellsWithoutCachesOfMap.get(cells);
  if (known) return known;
  const near = nearestKnownCellMap(cells, (map) => {
    const answer = cellsWithoutCachesOfMap.get(map);
    return answer ? { answer, clean: answer === map } : undefined;
  });
  let answer: CellMap;
  if (near?.value.clean && !near.keys.some((key) => holdsDerivedCache(cells[key]))) {
    answer = cells;
  } else if (near) {
    answer = { ...near.value.answer };
    for (const key of near.keys) {
      const cell = cells[key];
      if (cell) answer[key] = holdsDerivedCache(cell) ? withoutDerivedCache(cell) : cell;
      else delete answer[key];
    }
    recordCellMapEdit(near.value.answer, answer, near.keys);
  } else {
    let copy: CellMap | undefined;
    for (const key in cells) {
      const cell = cells[key];
      if (holdsDerivedCache(cell)) (copy ??= { ...cells })[key] = withoutDerivedCache(cell);
    }
    answer = copy ?? cells;
  }
  cellsWithoutCachesOfMap.set(cells, answer);
  return answer;
}

/** The longest `display` a page stores (`spreadsheetCellSchema.display`). */
export const MAX_CELL_DISPLAY_CHARS = 4096;

export function withDisplayCache<S extends DisplayCacheSheet>(
  stampedSheet: S,
  workbook?: Workbook,
  evaluatorFor?: (sheetName: string) => SheetEvaluator,
  candidates?: DisplayCacheCandidates,
  /** `withholdComputed`: store no computed value at all -- no formula's
   * display or number, no spilled value -- for a workbook that imports
   * another page's data (`workbookImportsRanges`): what this page stores
   * is what every reader of it, a print and an export see.
   * `resultsOnly`: a sheet this save did not touch. Only what a change elsewhere
   * can have moved is stamped: a formula whose result differs from the one it
   * stores, and what it spills. A formatted cell's text waits until its sheet
   * is edited: a first save stamping every sheet sent 3 MB for one typed cell
   * (xlsx-rounds R116-X02) and wrote into sheets nobody opened (R133-X02). */
  options?: { readonly withholdComputed?: boolean; readonly resultsOnly?: boolean },
): S {
  const withhold = options?.withholdComputed === true;
  const resultsOnly = options?.resultsOnly === true;
  const sheet = withWhatIfTableValues(stampedSheet, workbook, evaluatorFor);
  const cells = sheet.cells;
  const walk = candidates
    ? (resultsOnly ? candidates.formulaKeys : candidates.displayCacheKeys).filter((key) => Object.hasOwn(cells, key))
    : Object.keys(cells).filter((key) => !resultsOnly || Boolean(cells[key]?.formula));
  // Only build an evaluator when at least one cell can carry a display (formula
  // or number format); otherwise the sheet is byte-identical to its input.
  let hasCacheable = false;
  let hasStaleCache = false;
  for (const key of walk) {
    const cell = cells[key];
    hasStaleCache ||= Boolean(cell && ("display" in cell || "chartValueCache" in cell));
    const fmt = cell?.style?.numberFormat;
    if (cell?.formula || (typeof fmt === "string" && fmt.trim() !== "")) {
      hasCacheable = true;
      break;
    }
  }
  if (!hasCacheable) {
    if (!hasStaleCache) return sheet;
    let cleanCells: CellMap | undefined;
    for (const key of walk) {
      const cell = cells[key];
      if (cell && ("display" in cell || "chartValueCache" in cell)) {
        const next = { ...cell };
        delete next.display;
        delete next.chartValueCache;
        (cleanCells ??= { ...cells })[key] = next;
      }
    }
    return cleanCells ? { ...sheet, cells: cleanCells } : sheet;
  }

  const merges = normalizeMerges(sheet.merges);
  const covered = coveredCells(merges);
  const sheetName =
    typeof sheet.name === "string" && sheet.name ? sheet.name : sheet.id;

  // Prefer the shared WORKBOOK view of this sheet (so a cached display value
  // honors a cross-sheet reference, matching the live grid). A standalone call
  // reads what the cells already carry: each formula's stored result (its
  // stamped display / chart value), and a formula with none keeps what it had
  
  const liveCalculation = Boolean(workbook && evaluatorFor && workbook.hasSheet(sheetName));
  const evaluator: SheetEvaluator =
    workbook && evaluatorFor && workbook.hasSheet(sheetName)
      ? evaluatorFor(sheetName)
      : (() => {
          const getRawCell = (row: number, col: number) =>
            covered.has(`${String(row)}:${String(col)}`) ? null : (cells[`${row}:${col}`]?.value ?? null);
          const results = new Map<string, EngineValue>();
          for (const key of candidates?.formulaKeys ?? Object.keys(cells)) {
            const cell = cells[key];
            const [row, col] = key.split(":").map(Number);
            const stored = cell && Number.isSafeInteger(row) && Number.isSafeInteger(col) ? storedFormulaResult(getRawCell(row, col), cell) : undefined;
            if (stored !== undefined) results.set(worksheetFormulaResultKey(sheetName, row, col), stored);
          }
          return createValueStoreWorkbook([{ name: sheetName, getRawCell }], { results, onMissing: "throw" }).evaluatorFor(sheetName);
        })();
  let overlay: SpillOverlay;
  try {
    overlay = buildSpillOverlay(cells, covered, evaluator, candidates?.formulaKeys);
  } catch (error) {
    if (!(error instanceof FormulaResultUnavailableError)) throw error;
    overlay = { spilled: new Map(), blocked: new Set() };
  }

  // Most cells retain their exact identity. Allocate the map only when a
  // serialized cache actually changes, preserving original own-key order.
  let nextCells: CellMap | undefined;
  const stamping = resultsOnly
    ? [...walk, ...[...overlay.spilled.keys()].filter((key) => Object.hasOwn(cells, key) && !cells[key]?.formula)]
    : walk;
  for (const key of stamping) {
    const cell = cells[key];
    const fmt = cell?.style?.numberFormat;
    const cacheable =
      Boolean(cell?.formula) || (typeof fmt === "string" && fmt.trim() !== "");
    if (!cacheable) {
      // Plain cell: keep byte-stable. Defensively strip a stray `display` a
      // hand-edited/imported file may have left so the no-cache path stays clean.
      if (cell && ("display" in cell || "chartValueCache" in cell)) {
        const rest = { ...cell };
        delete rest.display;
        delete rest.chartValueCache;
        (nextCells ??= { ...cells })[key] = rest;
      }
      continue;
    }
    const [rowText, colText] = key.split(":", 2);
    const row = Number(rowText);
    const col = Number(colText);
    let display: string;
    let chartValueCache: ChartValueCache | undefined;
    try {
      display =
        Number.isFinite(row) && Number.isFinite(col)
          ? resolveCellDisplay(cell, row, col, evaluator, overlay)
          : "";
      if (cell.formula && Number.isFinite(row) && Number.isFinite(col) && evaluator.getFormula(row, col) === cell.value) {
        const value = resolveCellChartScalar(cell, row, col, evaluator, overlay);
        if (typeof value === "number" && Number.isFinite(value)) {
          chartValueCache = { kind: "number", source: cell.value, value };
        }
      }
    } catch (error) {
      if (error instanceof FormulaResultUnavailableError) {
        // Manual calculation owns an ephemeral result ledger. A formula that
        // lacks a result in that ledger must not be evaluated merely because an
        // autosave serializes the page; retain the last durable display (or no
        // display for a newly-authored formula) until an explicit calculation.
        // Preserve a prior raw result only while its canonical source matches.
        if (withhold && ("display" in cell || "chartValueCache" in cell)) {
          const next = { ...cell };
          delete next.display;
          delete next.chartValueCache;
          (nextCells ??= { ...cells })[key] = next;
        } else if ("chartValueCache" in cell && !matchingChartValueCache(cell.chartValueCache, cell.value)) {
          const next = { ...cell };
          delete next.chartValueCache;
          (nextCells ??= { ...cells })[key] = next;
        }
        continue;
      }
      throw error;
    }
    const sameChartValue = chartValueCache
      ? matchingChartValueCache(cell.chartValueCache, cell.value) && cell.chartValueCache.value === chartValueCache.value
      : !("chartValueCache" in cell);
    // A formula whose result is an array is a DYNAMIC ARRAY formula in the
    // file (`<f t="array" ref=... cm="1">`), which only the calculation can
    // tell: the anchor carries the range it spills over (the anchor alone
    // while blocked) for the XLSX writers, as it carries `display` (XLS-A090).
    let spillRef: string | undefined;
    if (!liveCalculation) {
      // Nothing is evaluated on this path: the stamp stays what it was.
      spillRef = (cell as { spillRef?: string }).spillRef;
    } else if ((cell as { spillRef?: string }).spillRef && formulaCallsUnknownFunction(String(cell.formula ?? ""))) {
      // A function this engine cannot evaluate (an add-in's `_xll.`, a name
      // it does not know): its result is not in hand, so the range the file
      // spills it over stays, as a result that is not available keeps its
      
      // over P4:P23 lost its range on the first edit, and the save read it
      // as a plain formula while the file kept the array).
      spillRef = (cell as { spillRef?: string }).spillRef;
    } else if (cell.formula && !(cell as { arrayRef?: unknown }).arrayRef && Number.isFinite(row) && Number.isFinite(col)) {
      try {
        const block = evaluator.getArray(row, col);
        if (block) {
          const anchor = `${columnToLetters(col)}${row + 1}`;
          const height = block.length;
          const width = block[0]?.length ?? 0;
          spillRef =
            overlay.blocked.has(key) || height * width <= 1
              ? anchor
              : `${anchor}:${columnToLetters(col + width - 1)}${row + height}`;
        } else if ((cell as { spillRef?: string }).spillRef === `${columnToLetters(col)}${row + 1}`) {
          // A single-cell dynamic array stays one when its result is one
          // value (or an error): Excel keeps the cell's dynamic-array mark
          // (`cm="1"`) whatever the formula returns. Dropped, the save
          // rewrote the file's single-cell dynamic arrays as plain formulas
          
          // #N/A). A formula that spilled over a range and returns one value
          // now is no longer stamped (XLS-A090).
          spillRef = `${columnToLetters(col)}${row + 1}`;
        }
      } catch (error) {
        if (!(error instanceof FormulaResultUnavailableError)) throw error;
        spillRef = (cell as { spillRef?: string }).spillRef;
      }
    }
    const sameSpill = (cell as { spillRef?: string }).spillRef === spillRef;
    if (withhold && (cell.formula || overlay.spilled.has(key))) {
      // Only the shape of a spill is kept (the XLSX writers need it).
      if ("display" in cell || "chartValueCache" in cell || !sameSpill) {
        const next = { ...cell } as typeof cell & { spillRef?: string };
        delete next.display;
        delete next.chartValueCache;
        if (spillRef) next.spillRef = spillRef;
        else delete next.spillRef;
        (nextCells ??= { ...cells })[key] = next;
      }
      continue;
    }
    if (display.length > MAX_CELL_DISPLAY_CHARS) {
      // Past what a page may store (the page schema's bound): a long text
      // shown as itself is read from its value, so no cache is kept rather
      // than a page no save accepts (xlsx-rounds R6-X08: a 7,780-character
      // abstract under an explicit General format refused every command).
      if ("display" in cell || !sameChartValue || !sameSpill) {
        const next = { ...cell } as typeof cell & { spillRef?: string };
        delete next.display;
        if (chartValueCache) next.chartValueCache = chartValueCache;
        else delete next.chartValueCache;
        if (spillRef) next.spillRef = spillRef;
        else delete next.spillRef;
        (nextCells ??= { ...cells })[key] = next;
      }
      continue;
    }
    // A number the formula already stores, with no text yet: nothing moved.
    if (resultsOnly && !("display" in cell) && chartValueCache && sameChartValue && sameSpill) continue;
    if (cell?.display !== display || !sameChartValue || !sameSpill) {
      const next = { ...cell, display } as typeof cell & { spillRef?: string };
      if (chartValueCache) next.chartValueCache = chartValueCache;
      else delete next.chartValueCache;
      if (spillRef) next.spillRef = spillRef;
      else delete next.spillRef;
      (nextCells ??= { ...cells })[key] = next;
    }
  }
  return nextCells ? { ...sheet, cells: nextCells } : sheet;
}

/**
 * The text and kind a What-If Data Table result keeps as its own value: what
 * a constant cell holding that value holds. Undefined for a value no constant
 * holds (an array, a picture), which keeps the value it had.
 */
function whatIfTableStoredValue(value: CellValue): { value: string; kind: "number" | "text" } | undefined {
  if (value === null || value === "") return { value: "", kind: "text" };
  if (typeof value === "number") return Number.isFinite(value) ? { value: String(value), kind: "number" } : undefined;
  if (typeof value === "boolean") return { value: value ? "TRUE" : "FALSE", kind: "text" };
  if (typeof value === "string") return { value, kind: "text" };
  if (isFormulaError(value)) return { value: value.code, kind: "text" };
  return undefined;
}

/**
 * A sheet whose What-If Data Table results hold the values the calculation
 * answers (Excel saves each result's value with the table), so a reader that
 * does not know tables -- the plain-text export, the AI's reader, an older
 * editor -- reads the current numbers. Only a save with the workbook's
 * results in hand refreshes them; a result without an answer keeps its value.
 */
export function withWhatIfTableValues<S extends DisplayCacheSheet>(
  sheet: S,
  workbook?: Workbook,
  evaluatorFor?: (sheetName: string) => SheetEvaluator,
): S {
  if (!sheet.whatIfTables || !workbook || !evaluatorFor) return sheet;
  const sheetName = typeof sheet.name === "string" && sheet.name ? sheet.name : sheet.id;
  if (!workbook.hasSheet(sheetName)) return sheet;
  const blocks = whatIfTableBlocks(sheet.whatIfTables, {
    rowCount: Number(sheet.rowCount) || 0,
    columnCount: Number(sheet.columnCount) || 0,
  });
  if (blocks.length === 0) return sheet;
  const evaluator = evaluatorFor(sheetName);
  let nextCells: CellMap | undefined;
  for (const block of blocks) {
    for (let row = block.r1; row <= block.r2; row += 1) {
      for (let col = block.c1; col <= block.c2; col += 1) {
        let computed: CellValue;
        try {
          computed = evaluator.getComputed(row, col);
        } catch (error) {
          if (error instanceof FormulaResultUnavailableError) continue;
          throw error;
        }
        const stored = whatIfTableStoredValue(computed);
        if (!stored) continue;
        const key = `${row}:${col}`;
        const cell = sheet.cells[key];
        if ((cell?.value ?? "") === stored.value && (stored.value === "" || cell?.kind === stored.kind)) continue;
        const next: SpreadsheetCell = { ...(cell ?? { value: "" }), value: stored.value, kind: stored.kind };
        if (stored.value === "") delete next.kind;
        (nextCells ??= { ...sheet.cells })[key] = next;
      }
    }
  }
  return nextCells ? { ...sheet, cells: nextCells } : sheet;
}

/**
 * Every chart of a sheet with the one chart model current
 * (lib/spreadsheet/chart-v2-model): a chart without `chartV2` gains it from
 * the stored cells (the migration -- in memory on load, on disk at save, when
 * the cells carry the caches the save stamps first), an agreeing chart has
 * its data cache refreshed, and a chart a flat write touched is re-read from
 * that write. Records that did not change keep their identity.
 */
export function withSheetChartModels<S extends ChartModelSheet>(
  sheet: S,
  sheets: readonly S[],
  pivotChartLive?: (source: SpreadsheetPivotChartSource) => SpreadsheetChartV2Live | null,
): S {
  const stored = sheets as unknown as readonly SpreadsheetChartStoredSheet[];
  const own = sheet as unknown as SpreadsheetChartStoredSheet;
  const sheetIdByName = (name: string) => {
    const matches = sheets.filter((item) => item.name.toLowerCase() === name.toLowerCase());
    return matches.length === 1 ? matches[0].id : undefined;
  };
  // A chart that plots visible cells only (c:plotVisOnly) holds the points of
  // its visible rows, as Excel caches them and the import reads them: the
  // stored cells of its hidden rows are not read into its data. R455-X01:
  // Graphs' two charts over rows 174-181 and 233-245 hidden read 91 points
  // where the file holds 70, and an untouched save was a changed one. Only a
  // projection chartVisiblePointIndices resolves for the data (any chart
  // type, one mask over aligned vectors of one sheet, a filtered sheet not);
  // one it cannot resolve reads every point as before.
  const visiblePoints = (like: SpreadsheetChartLike, live: SpreadsheetChartV2Live | null) => {
    const rows = live?.rows;
    if (!live || !rows || rows.length < 2 || live.headerRow !== true) return live;
    const projection = chartVisiblePointIndices(
      like as Parameters<typeof chartVisiblePointIndices>[0],
      sheets as unknown as Parameters<typeof chartVisiblePointIndices>[1],
      rows.length - 1,
      undefined,
      "data",
    );
    if (projection.state !== "resolved" || projection.sourcePointIndices.length === rows.length - 1) return live;
    return { ...live, rows: projectChartVisibilityRows(rows, projection) };
  };
  const reconcile = <T extends object>(record: T): T => {
    const like = record as unknown as SpreadsheetChartLike;
    // A PivotChart's cache is its PivotTable's report as it is now, so what
    // prints (and what a reader without the table sees) is current.
    const pivotSource = normalizeSpreadsheetPivotChartSource(like.pivotSource);
    try {
      // Stored with its gaps as null (their JSON form).
      return spreadsheetChartWithStoredGaps(
        reconcileSpreadsheetChartV2(
          like,
          pivotSource
            ? (pivotChartLive?.(pivotSource) ?? null)
            : visiblePoints(like, spreadsheetChartLiveFromStoredCells(like, own, stored)),
          { sheetIdByName },
        ),
      ) as unknown as T;
    } catch {
      return spreadsheetChartWithStoredGaps(like) as unknown as T;
    }
  };
  let charts: readonly object[] | undefined = sheet.charts;
  if (charts?.length) {
    const next = charts.map(reconcile);
    if (next.some((chart, index) => chart !== charts![index])) charts = next;
  }
  let objects: readonly { type?: unknown }[] | undefined = sheet.objects;
  if (objects?.some((object) => object.type === "chart")) {
    const next = objects.map((object) => (object.type === "chart" ? reconcile(object) : object));
    if (next.some((object, index) => object !== objects![index])) objects = next;
  }
  if (charts === sheet.charts && objects === sheet.objects) return sheet;
  return {
    ...sheet,
    ...(charts !== sheet.charts ? { charts } : {}),
    ...(objects !== sheet.objects ? { objects } : {}),
  } as S;
}

/** Whether a formula calls a function this engine does not know (outside its
 *  string literals): an add-in's (`_xll.`), a future one, a misspelling. */
export function formulaCallsUnknownFunction(formula: string): boolean {
  const outside = formula.replace(/"(?:[^"]|"")*"/g, '""');
  for (const match of outside.matchAll(/(^|[^\w.!$'])([A-Za-z_][\w.]*)\s*\(/g)) {
    if (!isKnownFunction(match[2]!)) return true;
  }
  return false;
}
