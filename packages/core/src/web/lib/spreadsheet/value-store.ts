/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { worksheetFormulaResultKey } from "./calculation-mode";
import {
  computeSpillOverlay,
  createImageCellValue,
  FormulaResultUnavailableError,
  formatCellValue,
  isArrayValue,
  isFormulaError,
  isLambdaValue,
  isOmittedValue,
  parseCellRef,
  readLiteralCellValue,
  ERRORS,
  type CellValue,
  type EngineValue,
  type SheetEvaluator,
  type Workbook,
  type WorkbookSheetInput,
} from "./formula-engine";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";

/** Whether two result maps hold the same value at every key. An array, a
 * lambda or a picture counts as changed: it is read again, never compared. */
export function sameFormulaResults(a: ReadonlyMap<string, EngineValue>, b: ReadonlyMap<string, EngineValue>): boolean {
  if (a === b) return true;
  if (a.size !== b.size) return false;
  for (const [key, value] of a) {
    const other = b.get(key);
    if (Object.is(value, other)) {
      if (other === undefined && !b.has(key)) return false;
      continue;
    }
    if (!isFormulaError(value) || !isFormulaError(other) || value.code !== other.code) return false;
  }
  return true;
}

export interface ValueStoreWorkbook {
  workbook: Workbook;
  evaluatorFor(sheetName: string): SheetEvaluator;
  memoSnapshot(): ReadonlyMap<string, EngineValue>;
  warmSeeded: number;
}

function scalar(value: EngineValue): CellValue {
  if (isArrayValue(value)) return value.grid[0]?.[0] ?? null;
  if (isLambdaValue(value)) return ERRORS.CALC;
  if (isOmittedValue(value)) return null;
  return value;
}

export function createValueStoreWorkbook(
  sheets: readonly WorkbookSheetInput[],
  options: {
    results: ReadonlyMap<string, EngineValue>;
    date1904?: ExcelDateSystemArg;
    /** A formula without a result: blank (the grid's deferred cell, the
     * default) or FormulaResultUnavailableError (a save keeps the display it had). */
    onMissing?: "blank" | "throw";
  },
): ValueStoreWorkbook {
  const byLower = new Map(sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]));
  const spills = new Map<string, { spilled: Map<string, CellValue>; blocked: Set<string> }>();
  const results = options.results;

  const isFormula = (input: WorkbookSheetInput, row: number, col: number, raw: string | null) =>
    raw !== null && raw.trimStart().startsWith("=") && !input.options?.isTextCell?.(row, col);

  function arrayAt(input: WorkbookSheetInput, row: number, col: number): CellValue[][] | null {
    const value = results.get(worksheetFormulaResultKey(input.name, row, col));
    if (!value || !isArrayValue(value)) return null;
    return value.grid.length === 1 && value.grid[0].length === 1 ? null : value.grid;
  }

  function overlayFor(input: WorkbookSheetInput) {
    const anchors = input.options?.spillAnchors;
    if (!anchors) return undefined;
    const lower = input.name.toLowerCase();
    let overlay = spills.get(lower);
    if (!overlay) {
      overlay = computeSpillOverlay({
        anchors: [...anchors()],
        getArray: (r, c) => arrayAt(input, r, c),
        isOccupied: (r, c) => input.options?.isSpillTargetOccupied
          ? input.options.isSpillTargetOccupied(r, c)
          : (() => { const raw = input.getRawCell(r, c); return raw !== null && raw !== ""; })(),
      });
      spills.set(lower, overlay);
    }
    return overlay;
  }

  function spilledAt(input: WorkbookSheetInput, row: number, col: number): CellValue | undefined {
    return overlayFor(input)?.spilled.get(`${row}:${col}`);
  }

  function computed(input: WorkbookSheetInput, row: number, col: number): CellValue {
    const raw = input.getRawCell(row, col);
    // The empty string a workbook stored is "", not a blank (formula-evaluator).
    if (raw === "" && input.options?.isTextCell?.(row, col)) return "";
    if (raw === null || raw === "") return spilledAt(input, row, col) ?? null;
    if (isFormula(input, row, col, raw)) {
      const value = results.get(worksheetFormulaResultKey(input.name, row, col));
      // A blocked array shows #SPILL!, as the engine reads it (S-049).
      if (value !== undefined && isArrayValue(value) && overlayFor(input)?.blocked.has(`${row}:${col}`)) return ERRORS.SPILL;
      if (value !== undefined) return scalar(value);
      if (options.onMissing === "throw") throw new FormulaResultUnavailableError(input.name, row, col);
      // Not answered yet: blank, as the grid's deferred evaluator shows it.
      return "";
    }
    const picture = input.options?.cellPicture?.(row, col);
    if (picture) return createImageCellValue(input.name, row, col, picture.alt);
    return input.options?.isTextCell?.(row, col) ? raw : readLiteralCellValue(raw);
  }

  const evaluators = new Map<string, SheetEvaluator>();
  function evaluatorFor(sheetName: string): SheetEvaluator {
    const lower = sheetName.toLowerCase();
    const known = evaluators.get(lower);
    if (known) return known;
    const input = byLower.get(lower);
    const evaluator: SheetEvaluator = {
      date1904: options.date1904 ?? false,
      getComputed: (row, col) => (input ? computed(input, row, col) : ERRORS.REF),
      getFormula: (row, col) => {
        const raw = input?.getRawCell(row, col) ?? null;
        return input && isFormula(input, row, col, raw) ? raw : null;
      },
      beginDeepRecovery: () => false,
      getDisplay: (row, col) => formatCellValue(input ? computed(input, row, col) : ERRORS.REF),
      getArray: (row, col) => (input ? arrayAt(input, row, col) : null),
      recomputeWithPrecedents: (row, col) => (input ? computed(input, row, col) : ERRORS.REF),
    };
    evaluators.set(lower, evaluator);
    return evaluator;
  }

  const workbook: Workbook = {
    hasSheet: (sheetName) => byLower.has(sheetName.toLowerCase()),
    getComputed: (sheetName, row, col) => evaluatorFor(sheetName).getComputed(row, col),
    getRange2D: (sheetName, from, to) => {
      const start = parseCellRef(from);
      const end = parseCellRef(to);
      if (!start || !end) return [[ERRORS.REF]];
      const evaluator = evaluatorFor(sheetName);
      const grid: CellValue[][] = [];
      for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row += 1) {
        const line: CellValue[] = [];
        for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col += 1) line.push(evaluator.getComputed(row, col));
        grid.push(line);
      }
      return grid;
    },
    sheetOrder: () => sheets.map((sheet) => sheet.name),
  };
  return { workbook, evaluatorFor, memoSnapshot: () => results, warmSeeded: 0 };
}
