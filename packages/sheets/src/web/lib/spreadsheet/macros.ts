/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure model for the spreadsheet MACRO RECORDER (Excel's View > Macros /
 * Developer > Record Macro, without VBA). A macro is a name plus the list of
 * steps a recording captured; running it replays those steps on the active
 * worksheet.
 *
 * Excel's recorder writes VBA that names the cells it touched: with
 * "Use Relative References" off (the default) a step addresses the cell it
 * was recorded in, with it on each step is an offset from the cell that was
 * active when recording began (R1C1 offsets), so the macro runs from
 * wherever the cursor is. Both modes live here as `mode`; the offsets in a
 * step count from the recording anchor (the sheet origin for an absolute
 * macro).
 *
 * What a recording captures (`recordSheetDiff`): every cell the commit
 * changed (value, formula and formatting), row heights and column widths,
 * hidden rows and columns, merges, freeze panes and row/column styles, as the
 * difference between the sheet before and after the commit. Inserting or
 * deleting rows and columns, sorting and running another macro are recorded
 * as the operation itself by the editor (`lines`, `sort`, `run`), so a replay
 * re-sorts the data it finds rather than pasting the old order.
 *
 * Dependency-free: unit-testable apart from the grid and commit plumbing.
 */

/** Row/column offsets (from the recording anchor) of a rectangle. */
export interface MacroRelativeRect {
  dRow: number;
  dCol: number;
  rows: number;
  cols: number;
}

/** A cell's persisted state, less its derived display fields. `null` clears
 *  the cell. The keys mirror the editor's cell model and pass through
 *  untouched, so a formatting the recorder does not know still replays. */
export type MacroCellPatch = Record<string, unknown>;

export interface MacroSortLevel {
  /** The key line (column for a row sort, row for a column sort) as an
   *  offset from the recording anchor. */
  at: number;
  direction: "asc" | "desc";
  sortOn?: string;
  formatKey?: string;
  placement?: "first" | "last";
  /** Excel's Custom List order, inlined so the macro carries it. */
  orderKeys?: string[];
}

export type MacroStep =
  /** v1 entry step: the text typed into a cell. Entered on replay the way
   *  the user typed it (50% is 0.5 in a percentage format). */
  | { kind?: undefined; dRow: number; dCol: number; value: string }
  /** Cells written by one commit, keyed "dRow:dCol". `formatOnly` marks a
   *  commit that changed formatting and no value (Excel's Selection.Font…). */
  | { kind: "cells"; cells: Record<string, MacroCellPatch | null>; formatOnly?: true }
  | {
      kind: "lines";
      axis: "row" | "col";
      op: "insert" | "delete";
      /** Offset of the first line from the anchor on that axis. */
      at: number;
      count: number;
    }
  /** A row height or column width; `null` returns it to the default. */
  | { kind: "size"; axis: "row" | "col"; at: number; size: number | null }
  | { kind: "defaultSize"; axis: "row" | "col"; size: number | null }
  | { kind: "hidden"; axis: "row" | "col"; at: number; hidden: boolean }
  | { kind: "merges"; add: MacroRelativeRect[]; remove: MacroRelativeRect[] }
  | { kind: "freeze"; rows: number; cols: number }
  | { kind: "lineStyle"; axis: "row" | "col"; at: number; style: unknown | null }
  /** Excel records `Range(...).Select`; replay moves the selection. */
  | { kind: "select"; dRow: number; dCol: number; rows: number; cols: number }
  /** The worksheet AutoFilter after the commit (Excel's Range.AutoFilter):
   *  its range and each column's criteria, columns as offsets from the
   *  anchor. A field set to null is cleared. */
  | {
      kind: "filters";
      autoFilter?: MacroRelativeRect | null;
      filters?: Record<string, unknown> | null;
      colorFilters?: Record<string, unknown> | null;
      conditionFilters?: Record<string, unknown> | null;
    }
  /** Drawing objects (pictures, shapes, charts) the commit added, changed or
   *  removed. Excel records shapes by name at absolute positions
   *  (ActiveSheet.Shapes("Rectangle 1").Left = ...), and so does this: an
   *  added object is copied in on every run, a changed or removed one is
   *  matched by its identity on the sheet the macro runs on. */
  | {
      kind: "objects";
      collection: "objects" | "charts";
      added: Array<Record<string, unknown>>;
      updated: Array<Record<string, unknown>>;
      removed: string[];
    }
  | {
      kind: "sort";
      dRow: number;
      dCol: number;
      rows: number;
      cols: number;
      /** Quick Sort (Data > A-Z / Z-A on one column) rather than the Sort dialog. */
      quick?: true;
      orientation?: "rows" | "columns";
      hasHeader?: boolean;
      caseSensitive?: boolean;
      currentRegion?: boolean;
      levels: MacroSortLevel[];
    }
  /** Excel records `Application.Run "Name"` when a macro runs while recording. */
  | { kind: "run"; name: string };

export type MacroPureStep = Extract<
  MacroStep,
  | { kind?: undefined }
  | { kind: "cells" }
  | { kind: "size" }
  | { kind: "defaultSize" }
  | { kind: "hidden" }
  | { kind: "merges" }
  | { kind: "freeze" }
  | { kind: "lineStyle" }
  | { kind: "filters" }
  | { kind: "objects" }
>;

export type MacroReferenceMode = "absolute" | "relative";

export interface MacroShortcut {
  /** One lowercase Latin letter, Excel's Ctrl+letter. */
  key: string;
  /** Ctrl+Shift+letter (Excel: an uppercase letter typed in the box). */
  shift?: boolean;
}

export interface SpreadsheetMacro {
  name: string;
  steps: MacroStep[];
  /** Absent on a macro recorded before the mode existed: those were relative. */
  mode?: MacroReferenceMode;
  /** The cell that was active when a relative macro was recorded: the
   *  origin its steps and formulas were written against. */
  anchorRow?: number;
  anchorCol?: number;
  shortcut?: MacroShortcut;
  description?: string;
}

/** Caps guarding persisted size / runaway recordings. */
export const MAX_MACRO_STEPS = 5000;
export const MAX_MACROS = 100;
/** Excel: a macro name holds up to 255 characters. */
export const MAX_MACRO_NAME_LENGTH = 255;
export const MAX_MACRO_DESCRIPTION_LENGTH = 1000;
/** Cells one recorded commit may carry (a paste of a large block). */
export const MAX_MACRO_STEP_CELLS = 20_000;
/** A macro that runs a macro that runs a macro: how deep a run expands. */
export const MAX_MACRO_RUN_DEPTH = 8;

/** The mode of a macro, with the pre-mode recordings read as relative. */
export function macroReferenceMode(macro: Pick<SpreadsheetMacro, "mode">): MacroReferenceMode {
  return macro.mode === "absolute" ? "absolute" : "relative";
}

/** The cell a macro's step offsets count from when it runs: the active
 *  cell for a relative macro, the sheet origin for an absolute one. */
export function macroReplayOrigin(
  macro: Pick<SpreadsheetMacro, "mode">,
  active: { row: number; col: number },
): { row: number; col: number } {
  return macroReferenceMode(macro) === "relative" ? active : { row: 0, col: 0 };
}

/** How far a relative macro's formulas move when it runs at `origin`: the
 *  distance from the cell it was recorded at. Zero for an absolute macro. */
export function macroFormulaShift(
  macro: Pick<SpreadsheetMacro, "mode" | "anchorRow" | "anchorCol">,
  origin: { row: number; col: number },
): { rowDelta: number; colDelta: number } {
  if (macroReferenceMode(macro) !== "relative") return { rowDelta: 0, colDelta: 0 };
  return {
    rowDelta: Number.isInteger(macro.anchorRow) ? origin.row - (macro.anchorRow as number) : 0,
    colDelta: Number.isInteger(macro.anchorCol) ? origin.col - (macro.anchorCol as number) : 0,
  };
}

// ---------------------------------------------------------------------------
// Names and shortcuts
// ---------------------------------------------------------------------------

/**
 * Excel's macro-name rules (Record Macro: "Macro name"): it starts with a
 * letter, continues with letters, numbers and underscores (no spaces), holds
 * at most 255 characters, and is not something that reads as a cell
 * reference (A1, R1C1).
 */
export function isValidMacroName(name: string): boolean {
  const trimmed = name.trim();
  if (trimmed.length === 0 || trimmed.length > MAX_MACRO_NAME_LENGTH) return false;
  if (!/^[\p{L}_][\p{L}\p{N}_]*$/u.test(trimmed)) return false;
  if (/^[A-Za-z]{1,3}[0-9]+$/.test(trimmed)) return false; // A1-style
  if (/^R(\[?-?[0-9]+\]?)?C(\[?-?[0-9]+\]?)?$/i.test(trimmed)) return false; // R1C1-style
  return true;
}

/** Excel proposes "Macro1", then the next number no macro in reach uses. */
export function nextDefaultMacroName(taken: Iterable<string>): string {
  const lower = new Set<string>();
  for (const name of taken) lower.add(name.toLowerCase());
  for (let n = 1; n < 100_000; n += 1) {
    const candidate = `Macro${n}`;
    if (!lower.has(candidate.toLowerCase())) return candidate;
  }
  return `Macro${Date.now()}`;
}

/** Excel's shortcut box takes one letter; an uppercase letter means
 *  Ctrl+Shift. Anything else clears the shortcut. */
export function parseMacroShortcutInput(text: string): MacroShortcut | null {
  const letter = text.trim().slice(-1);
  if (!/^[A-Za-z]$/.test(letter)) return null;
  const shift = letter === letter.toUpperCase();
  return shift ? { key: letter.toLowerCase(), shift: true } : { key: letter.toLowerCase() };
}

/** The letter the shortcut box shows: uppercase when Shift is part of it. */
export function macroShortcutInputText(shortcut: MacroShortcut | undefined): string {
  if (!shortcut) return "";
  return shortcut.shift ? shortcut.key.toUpperCase() : shortcut.key;
}

export function macroShortcutLabel(shortcut: MacroShortcut | undefined): string {
  if (!shortcut) return "";
  return `Ctrl+${shortcut.shift ? "Shift+" : ""}${shortcut.key.toUpperCase()}`;
}

function normalizeShortcut(raw: unknown): MacroShortcut | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const r = raw as Record<string, unknown>;
  const key = typeof r.key === "string" ? r.key.toLowerCase() : "";
  if (!/^[a-z]$/.test(key)) return undefined;
  return r.shift === true ? { key, shift: true } : { key };
}

/** The macro a Ctrl(+Shift)+letter press names, if any. Excel gives a macro
 *  shortcut precedence over the built-in command on the same keys. */
export function findMacroByShortcut<T extends Pick<SpreadsheetMacro, "shortcut">>(
  macros: readonly T[],
  press: { key: string; shift: boolean },
): T | undefined {
  const key = press.key.toLowerCase();
  if (!/^[a-z]$/.test(key)) return undefined;
  return macros.find(
    (macro) =>
      macro.shortcut !== undefined &&
      macro.shortcut.key === key &&
      Boolean(macro.shortcut.shift) === press.shift,
  );
}

// ---------------------------------------------------------------------------
// Recording
// ---------------------------------------------------------------------------

/** The slice of the editor's sheet model the recorder and replay read. */
export interface MacroSheet {
  rowCount: number;
  columnCount: number;
  cells: Record<string, object>;
  rowHeights?: Record<string | number, number>;
  columnWidths?: Record<string | number, number>;
  defaultRowHeight?: number;
  defaultColumnWidth?: number;
  merges?: Array<{ row: number; col: number; rowSpan: number; colSpan: number }>;
  manuallyHiddenRows?: number[];
  manuallyHiddenCols?: number[];
  freezeRows?: number;
  freezeCols?: number;
  rowStyles?: Record<string, unknown>;
  columnStyles?: Record<string, unknown>;
  filters?: Record<string, unknown>;
  colorFilters?: Record<string, unknown>;
  conditionFilters?: Record<string, unknown>;
  autoFilterRef?: string;
  objects?: Array<{ id: string }>;
  charts?: Array<{ id: string }>;
}

/** Derived fields a recording never carries: they are recomputed on write. */
const VOLATILE_CELL_FIELDS = new Set(["display", "chartValueCache"]);

function cellPatchOf(cell: object | undefined): MacroCellPatch | null {
  if (!cell) return null;
  const source = cell as Record<string, unknown>;
  const patch: MacroCellPatch = {};
  for (const key of Object.keys(source)) {
    if (VOLATILE_CELL_FIELDS.has(key)) continue;
    if (source[key] === undefined) continue;
    patch[key] = source[key];
  }
  return patch;
}

function stableJson(value: unknown): string {
  return JSON.stringify(value, (_key, inner) => {
    if (inner && typeof inner === "object" && !Array.isArray(inner)) {
      const sorted: Record<string, unknown> = {};
      for (const k of Object.keys(inner as Record<string, unknown>).sort()) {
        const v = (inner as Record<string, unknown>)[k];
        if (v !== undefined) sorted[k] = v;
      }
      return sorted;
    }
    return inner;
  });
}

function samePatch(a: MacroCellPatch | null, b: MacroCellPatch | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return stableJson(a) === stableJson(b);
}

function parseKey(key: string): [number, number] | null {
  const separator = key.indexOf(":");
  if (separator <= 0) return null;
  const row = Number(key.slice(0, separator));
  const col = Number(key.slice(separator + 1));
  return Number.isInteger(row) && Number.isInteger(col) ? [row, col] : null;
}

function spanKey(span: { row: number; col: number; rowSpan: number; colSpan: number }): string {
  return `${span.row}:${span.col}:${span.rowSpan}:${span.colSpan}`;
}

function a1Column(letters: string): number {
  let col = 0;
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col - 1;
}

/** "A1:C9" (or "A1") as 0-based bounds. */
export function macroRangeBounds(
  ref: string | undefined,
): { r1: number; c1: number; r2: number; c2: number } | null {
  if (!ref) return null;
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref.trim());
  if (!match) return null;
  const c1 = a1Column(match[1]);
  const r1 = Number(match[2]) - 1;
  const c2 = match[3] ? a1Column(match[3]) : c1;
  const r2 = match[4] ? Number(match[4]) - 1 : r1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

function shiftColumnKeys(
  map: Record<string, unknown> | undefined,
  delta: number,
): Record<string, unknown> | undefined {
  if (!map) return undefined;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(map)) {
    const col = Number(key);
    if (!Number.isInteger(col)) continue;
    out[String(col + delta)] = value;
  }
  return out;
}

function cloneJson<T>(value: T): T {
  return value === undefined ? value : (JSON.parse(JSON.stringify(value)) as T);
}

export interface MacroRecordResult {
  steps: MacroStep[];
  /** True when a commit changed more cells than one step may carry: the
   *  recording keeps the first MAX_MACRO_STEP_CELLS and says so. */
  truncated: boolean;
}

/**
 * The steps that turn `prev` into `next`, written relative to the anchor.
 * Cells are compared by their persisted state (not identity), so a commit
 * that rewrote a cell map without changing a cell records nothing.
 */
export function recordSheetDiff(
  prev: MacroSheet,
  next: MacroSheet,
  anchor: { row: number; col: number },
): MacroRecordResult {
  const steps: MacroStep[] = [];
  let truncated = false;

  if (prev.cells !== next.cells) {
    const cells: Record<string, MacroCellPatch | null> = {};
    let count = 0;
    let formatOnly = true;
    const keys = new Set<string>([...Object.keys(prev.cells), ...Object.keys(next.cells)]);
    for (const key of keys) {
      const before = prev.cells[key];
      const after = next.cells[key];
      if (before === after) continue;
      const beforePatch = cellPatchOf(before);
      const afterPatch = cellPatchOf(after);
      if (samePatch(beforePatch, afterPatch)) continue;
      const parsed = parseKey(key);
      if (!parsed) continue;
      if (count >= MAX_MACRO_STEP_CELLS) {
        truncated = true;
        break;
      }
      if (
        !beforePatch ||
        !afterPatch ||
        beforePatch.value !== afterPatch.value ||
        beforePatch.formula !== afterPatch.formula
      ) {
        formatOnly = false;
      }
      cells[`${parsed[0] - anchor.row}:${parsed[1] - anchor.col}`] = afterPatch;
      count += 1;
    }
    if (count > 0) steps.push(formatOnly ? { kind: "cells", cells, formatOnly: true } : { kind: "cells", cells });
  }

  const sizeDiff = (
    axis: "row" | "col",
    before: Record<string | number, number> | undefined,
    after: Record<string | number, number> | undefined,
  ) => {
    if (before === after) return;
    const base = axis === "row" ? anchor.row : anchor.col;
    const keys = new Set<string>([...Object.keys(before ?? {}), ...Object.keys(after ?? {})]);
    for (const key of keys) {
      const b = before?.[key];
      const a = after?.[key];
      if (b === a) continue;
      const index = Number(key);
      if (!Number.isInteger(index)) continue;
      steps.push({ kind: "size", axis, at: index - base, size: typeof a === "number" ? a : null });
    }
  };
  sizeDiff("row", prev.rowHeights, next.rowHeights);
  sizeDiff("col", prev.columnWidths, next.columnWidths);
  if (prev.defaultRowHeight !== next.defaultRowHeight) {
    steps.push({ kind: "defaultSize", axis: "row", size: next.defaultRowHeight ?? null });
  }
  if (prev.defaultColumnWidth !== next.defaultColumnWidth) {
    steps.push({ kind: "defaultSize", axis: "col", size: next.defaultColumnWidth ?? null });
  }

  const hiddenDiff = (axis: "row" | "col", before: number[] | undefined, after: number[] | undefined) => {
    if (before === after) return;
    const base = axis === "row" ? anchor.row : anchor.col;
    const b = new Set(before ?? []);
    const a = new Set(after ?? []);
    for (const index of a) if (!b.has(index)) steps.push({ kind: "hidden", axis, at: index - base, hidden: true });
    for (const index of b) if (!a.has(index)) steps.push({ kind: "hidden", axis, at: index - base, hidden: false });
  };
  hiddenDiff("row", prev.manuallyHiddenRows, next.manuallyHiddenRows);
  hiddenDiff("col", prev.manuallyHiddenCols, next.manuallyHiddenCols);

  if (prev.merges !== next.merges) {
    const before = new Map((prev.merges ?? []).map((span) => [spanKey(span), span]));
    const after = new Map((next.merges ?? []).map((span) => [spanKey(span), span]));
    const toRect = (span: { row: number; col: number; rowSpan: number; colSpan: number }): MacroRelativeRect => ({
      dRow: span.row - anchor.row,
      dCol: span.col - anchor.col,
      rows: span.rowSpan,
      cols: span.colSpan,
    });
    const add: MacroRelativeRect[] = [];
    const remove: MacroRelativeRect[] = [];
    for (const [key, span] of after) if (!before.has(key)) add.push(toRect(span));
    for (const [key, span] of before) if (!after.has(key)) remove.push(toRect(span));
    if (add.length > 0 || remove.length > 0) steps.push({ kind: "merges", add, remove });
  }

  if ((prev.freezeRows ?? 0) !== (next.freezeRows ?? 0) || (prev.freezeCols ?? 0) !== (next.freezeCols ?? 0)) {
    steps.push({ kind: "freeze", rows: next.freezeRows ?? 0, cols: next.freezeCols ?? 0 });
  }

  const lineStyleDiff = (
    axis: "row" | "col",
    before: Record<string, unknown> | undefined,
    after: Record<string, unknown> | undefined,
  ) => {
    if (before === after) return;
    const base = axis === "row" ? anchor.row : anchor.col;
    const keys = new Set<string>([...Object.keys(before ?? {}), ...Object.keys(after ?? {})]);
    for (const key of keys) {
      const b = before?.[key];
      const a = after?.[key];
      if (b === a || stableJson(b ?? null) === stableJson(a ?? null)) continue;
      const index = Number(key);
      if (!Number.isInteger(index)) continue;
      steps.push({ kind: "lineStyle", axis, at: index - base, style: a === undefined ? null : a });
    }
  };
  lineStyleDiff("row", prev.rowStyles, next.rowStyles);
  lineStyleDiff("col", prev.columnStyles, next.columnStyles);

  // The AutoFilter: one step carrying every filter field the commit changed.
  const filterStep: Extract<MacroStep, { kind: "filters" }> = { kind: "filters" };
  let filterChanged = false;
  for (const field of ["filters", "colorFilters", "conditionFilters"] as const) {
    if (prev[field] === next[field] || stableJson(prev[field] ?? null) === stableJson(next[field] ?? null)) continue;
    filterStep[field] = next[field] ? (shiftColumnKeys(cloneJson(next[field]), -anchor.col) ?? null) : null;
    filterChanged = true;
  }
  if ((prev.autoFilterRef ?? "") !== (next.autoFilterRef ?? "")) {
    const bounds = macroRangeBounds(next.autoFilterRef);
    filterStep.autoFilter = bounds
      ? {
          dRow: bounds.r1 - anchor.row,
          dCol: bounds.c1 - anchor.col,
          rows: bounds.r2 - bounds.r1 + 1,
          cols: bounds.c2 - bounds.c1 + 1,
        }
      : null;
    filterChanged = true;
  }
  if (filterChanged) steps.push(filterStep);

  // Drawing objects, matched by identity.
  for (const collection of ["objects", "charts"] as const) {
    const before = prev[collection];
    const after = next[collection];
    if (before === after) continue;
    const beforeById = new Map((before ?? []).map((item) => [item.id, item]));
    const afterById = new Map((after ?? []).map((item) => [item.id, item]));
    const added: Array<Record<string, unknown>> = [];
    const updated: Array<Record<string, unknown>> = [];
    const removed: string[] = [];
    for (const [id, item] of afterById) {
      const old = beforeById.get(id);
      if (!old) added.push(cloneJson(item) as unknown as Record<string, unknown>);
      else if (old !== item && stableJson(old) !== stableJson(item)) {
        updated.push(cloneJson(item) as unknown as Record<string, unknown>);
      }
    }
    for (const id of beforeById.keys()) if (!afterById.has(id)) removed.push(id);
    if (added.length > 0 || updated.length > 0 || removed.length > 0) {
      steps.push({ kind: "objects", collection, added, updated, removed });
    }
  }

  return { steps, truncated };
}

/** Append recorded steps, dropping those past the step cap. */
export function appendMacroSteps(
  steps: readonly MacroStep[],
  added: readonly MacroStep[],
): { steps: MacroStep[]; dropped: number } {
  const room = Math.max(0, MAX_MACRO_STEPS - steps.length);
  const kept = added.slice(0, room);
  return { steps: [...steps, ...kept], dropped: added.length - kept.length };
}

/**
 * v1 recording hook: append a typed entry as a step relative to the anchor.
 * Kept for the tests and the legacy shape; the editor records commits with
 * `recordSheetDiff`.
 */
export function recordStep(
  steps: readonly MacroStep[],
  anchorRow: number,
  anchorCol: number,
  row: number,
  col: number,
  value: string,
): MacroStep[] {
  const dRow = row - anchorRow;
  const dCol = col - anchorCol;
  const without = steps.filter((s) => !(s.kind === undefined && s.dRow === dRow && s.dCol === dCol));
  if (without.length >= MAX_MACRO_STEPS) return [...steps];
  return [...without, { dRow, dCol, value }];
}

// ---------------------------------------------------------------------------
// Replay
// ---------------------------------------------------------------------------

export function isPureMacroStep(step: MacroStep): step is MacroPureStep {
  return (
    step.kind === undefined ||
    step.kind === "cells" ||
    step.kind === "size" ||
    step.kind === "defaultSize" ||
    step.kind === "hidden" ||
    step.kind === "merges" ||
    step.kind === "freeze" ||
    step.kind === "lineStyle" ||
    step.kind === "filters" ||
    step.kind === "objects"
  );
}

export interface MacroReplayOptions {
  /** Rebase a formula's relative references by a row/column distance (the
   *  engine's copy semantics). A relative macro run somewhere other than
   *  where it was recorded shifts its formulas by the distance moved, as
   *  Excel's relative recording (R1C1 offsets) does; absolute references
   *  stay. Without it formulas are written as recorded. */
  shiftFormula?: (formula: string, rowDelta: number, colDelta: number) => string;
}

/** Map a macro's v1 entry steps to absolute writes at the target cell. Steps
 *  landing off the top/left edge are dropped. */
export function replayMacro(
  macro: SpreadsheetMacro,
  targetRow: number,
  targetCol: number,
  options: MacroReplayOptions = {},
): Array<{ row: number; col: number; value: string }> {
  const out: Array<{ row: number; col: number; value: string }> = [];
  const { rowDelta, colDelta } = macroFormulaShift(macro, { row: targetRow, col: targetCol });
  for (const step of macro.steps) {
    if (step.kind !== undefined) continue;
    const row = targetRow + step.dRow;
    const col = targetCol + step.dCol;
    if (row < 0 || col < 0) continue;
    const value =
      options.shiftFormula && (rowDelta !== 0 || colDelta !== 0) && step.value.trim().startsWith("=")
        ? options.shiftFormula(step.value, rowDelta, colDelta)
        : step.value;
    out.push({ row, col, value });
  }
  return out;
}

export type MacroReplayPlan =
  | {
      ok: true;
      writes: Array<{ row: number; col: number; value: string }>;
      rowCount: number;
      columnCount: number;
    }
  | { ok: false; reason: "beyond-sheet-end" };

/** Plan a v1 replay against the sheet's extent and the address space. */
export function planMacroReplay(
  macro: SpreadsheetMacro,
  targetRow: number,
  targetCol: number,
  extent: { rowCount: number; columnCount: number },
  limits: { maxRows: number; maxColumns: number },
  options: MacroReplayOptions = {},
): MacroReplayPlan {
  const writes = replayMacro(macro, targetRow, targetCol, options);
  let rowCount = extent.rowCount;
  let columnCount = extent.columnCount;
  for (const write of writes) {
    if (write.row >= limits.maxRows || write.col >= limits.maxColumns) {
      return { ok: false, reason: "beyond-sheet-end" };
    }
    rowCount = Math.max(rowCount, write.row + 1);
    columnCount = Math.max(columnCount, write.col + 1);
  }
  return { ok: true, writes, rowCount, columnCount };
}

export interface MacroFoldContext<T extends MacroSheet> extends MacroReplayOptions {
  origin: { row: number; col: number };
  rowDelta: number;
  colDelta: number;
  maxRows: number;
  maxColumns: number;
  /** True for a cell a replay may not write (sheet protection). */
  isLocked?: (row: number, col: number) => boolean;
  /** Enter a v1 step's text into the cell map as the user typed it. */
  writeEntry: (cells: T["cells"], row: number, col: number, value: string) => void;
  /** A fresh identity for an object a macro adds on replay. */
  newObjectId?: () => string;
  /** The cell map after a merge: the span's value carried to its anchor and
   *  the covered cells cleared. */
  mergeCells?: (
    cells: T["cells"],
    span: { row: number; col: number; rowSpan: number; colSpan: number },
  ) => T["cells"];
}

export type MacroFoldResult<T extends MacroSheet> =
  | { ok: true; sheet: T; touchedCells: string[] }
  | { ok: false; reason: "beyond-sheet-end" | "locked" };

function shiftedPatch(
  patch: MacroCellPatch,
  ctx: Pick<MacroFoldContext<MacroSheet>, "shiftFormula" | "rowDelta" | "colDelta">,
): MacroCellPatch {
  if (!ctx.shiftFormula || (ctx.rowDelta === 0 && ctx.colDelta === 0)) return patch;
  const next: MacroCellPatch = { ...patch };
  for (const field of ["value", "formula"] as const) {
    const text = next[field];
    if (typeof text === "string" && text.trim().startsWith("=")) {
      next[field] = ctx.shiftFormula(text, ctx.rowDelta, ctx.colDelta);
    }
  }
  return next;
}

/**
 * Apply a run of pure steps to a sheet, in order, as one change. Writes past
 * the stored extent grow it; a write past the worksheet's last row or column
 * refuses the whole batch, as does a write to a locked cell.
 */
export function foldMacroSteps<T extends MacroSheet>(
  sheet: T,
  steps: readonly MacroPureStep[],
  ctx: MacroFoldContext<T>,
): MacroFoldResult<T> {
  let next: T = { ...sheet };
  // The cell map is copied on the first cell write, so a batch that only
  // sizes lines keeps the map's identity (the editor's change hints read it).
  let cells: T["cells"] = sheet.cells;
  const ownCells = (): T["cells"] => {
    if (cells === sheet.cells) cells = { ...sheet.cells };
    return cells;
  };
  let rowCount = sheet.rowCount;
  let columnCount = sheet.columnCount;
  const touched: string[] = [];
  const reach = (row: number, col: number): boolean => {
    if (row >= ctx.maxRows || col >= ctx.maxColumns) return false;
    rowCount = Math.max(rowCount, row + 1);
    columnCount = Math.max(columnCount, col + 1);
    return true;
  };
  const lineOf = (axis: "row" | "col", at: number) => (axis === "row" ? ctx.origin.row : ctx.origin.col) + at;
  for (const step of steps) {
    if (step.kind === undefined) {
      const row = ctx.origin.row + step.dRow;
      const col = ctx.origin.col + step.dCol;
      if (row < 0 || col < 0) continue;
      if (!reach(row, col)) return { ok: false, reason: "beyond-sheet-end" };
      if (ctx.isLocked?.(row, col)) return { ok: false, reason: "locked" };
      const value =
        ctx.shiftFormula && (ctx.rowDelta !== 0 || ctx.colDelta !== 0) && step.value.trim().startsWith("=")
          ? ctx.shiftFormula(step.value, ctx.rowDelta, ctx.colDelta)
          : step.value;
      ctx.writeEntry(ownCells(), row, col, value);
      touched.push(`${row}:${col}`);
      continue;
    }
    switch (step.kind) {
      case "cells": {
        for (const key of Object.keys(step.cells)) {
          const parsed = parseKey(key);
          if (!parsed) continue;
          const row = ctx.origin.row + parsed[0];
          const col = ctx.origin.col + parsed[1];
          if (row < 0 || col < 0) continue;
          if (!reach(row, col)) return { ok: false, reason: "beyond-sheet-end" };
          if (ctx.isLocked?.(row, col)) return { ok: false, reason: "locked" };
          const patch = step.cells[key];
          const target = `${row}:${col}`;
          const map = ownCells() as Record<string, object>;
          if (patch === null) delete map[target];
          else map[target] = shiftedPatch(patch, ctx);
          touched.push(target);
        }
        break;
      }
      case "size": {
        const index = lineOf(step.axis, step.at);
        if (index < 0) continue;
        const field = step.axis === "row" ? "rowHeights" : "columnWidths";
        const map: Record<string | number, number> = { ...(next[field] ?? {}) };
        if (step.size === null) delete map[index];
        else map[index] = step.size;
        if (Object.keys(map).length > 0) next = { ...next, [field]: map };
        else {
          next = { ...next };
          delete next[field];
        }
        break;
      }
      case "defaultSize": {
        const field = step.axis === "row" ? "defaultRowHeight" : "defaultColumnWidth";
        next = { ...next };
        if (step.size === null) delete next[field];
        else next[field] = step.size;
        break;
      }
      case "hidden": {
        const index = lineOf(step.axis, step.at);
        if (index < 0) continue;
        const field = step.axis === "row" ? "manuallyHiddenRows" : "manuallyHiddenCols";
        const set = new Set(next[field] ?? []);
        if (step.hidden) set.add(index);
        else set.delete(index);
        next = { ...next };
        if (set.size > 0) next[field] = [...set].sort((a, b) => a - b);
        else delete next[field];
        break;
      }
      case "merges": {
        const spans = [...(next.merges ?? [])];
        const toSpan = (rect: MacroRelativeRect) => ({
          row: ctx.origin.row + rect.dRow,
          col: ctx.origin.col + rect.dCol,
          rowSpan: rect.rows,
          colSpan: rect.cols,
        });
        for (const rect of step.remove) {
          const key = spanKey(toSpan(rect));
          const index = spans.findIndex((span) => spanKey(span) === key);
          if (index >= 0) spans.splice(index, 1);
        }
        for (const rect of step.add) {
          const span = toSpan(rect);
          if (span.row < 0 || span.col < 0 || span.rowSpan < 1 || span.colSpan < 1) continue;
          if (!reach(span.row + span.rowSpan - 1, span.col + span.colSpan - 1)) {
            return { ok: false, reason: "beyond-sheet-end" };
          }
          // A new span replaces any it overlaps (Excel merges over them).
          for (let i = spans.length - 1; i >= 0; i -= 1) {
            const other = spans[i];
            const overlaps =
              other.row < span.row + span.rowSpan &&
              span.row < other.row + other.rowSpan &&
              other.col < span.col + span.colSpan &&
              span.col < other.col + other.colSpan;
            if (overlaps) spans.splice(i, 1);
          }
          spans.push(span);
          if (ctx.mergeCells) cells = ctx.mergeCells(ownCells(), span);
        }
        next = { ...next };
        if (spans.length > 0) next.merges = spans;
        else delete next.merges;
        break;
      }
      case "freeze": {
        next = { ...next };
        if (step.rows > 0) next.freezeRows = step.rows;
        else delete next.freezeRows;
        if (step.cols > 0) next.freezeCols = step.cols;
        else delete next.freezeCols;
        break;
      }
      case "lineStyle": {
        const index = lineOf(step.axis, step.at);
        if (index < 0) continue;
        const field = step.axis === "row" ? "rowStyles" : "columnStyles";
        const map: Record<string, unknown> = { ...(next[field] ?? {}) };
        if (step.style === null) delete map[index];
        else map[index] = step.style;
        next = { ...next };
        if (Object.keys(map).length > 0) next[field] = map;
        else delete next[field];
        break;
      }
      case "filters": {
        next = { ...next };
        for (const field of ["filters", "colorFilters", "conditionFilters"] as const) {
          if (!(field in step)) continue;
          const value = step[field];
          const shifted = value ? shiftColumnKeys(cloneJson(value), ctx.origin.col) : undefined;
          // A column left of the sheet's first one has no place to filter.
          for (const key of Object.keys(shifted ?? {})) if (Number(key) < 0) delete shifted![key];
          if (shifted && Object.keys(shifted).length > 0) next[field] = shifted;
          else delete next[field];
        }
        if ("autoFilter" in step) {
          const rect = step.autoFilter;
          if (rect) {
            const r1 = ctx.origin.row + rect.dRow;
            const c1 = ctx.origin.col + rect.dCol;
            if (r1 < 0 || c1 < 0) break;
            if (!reach(r1 + rect.rows - 1, c1 + rect.cols - 1)) return { ok: false, reason: "beyond-sheet-end" };
            next.autoFilterRef = macroRangeLabel(r1, c1, rect.rows, rect.cols);
          } else {
            delete next.autoFilterRef;
          }
        }
        break;
      }
      case "objects": {
        const list = [...((next[step.collection] ?? []) as Array<{ id: string }>)];
        for (const id of step.removed) {
          const index = list.findIndex((item) => item.id === id);
          if (index >= 0) list.splice(index, 1);
        }
        for (const item of step.updated) {
          const index = list.findIndex((existing) => existing.id === item.id);
          if (index >= 0) list[index] = cloneJson(item) as unknown as { id: string };
        }
        for (const item of step.added) {
          const copy = cloneJson(item) as unknown as { id: string };
          if (ctx.newObjectId && list.some((existing) => existing.id === copy.id)) copy.id = ctx.newObjectId();
          list.push(copy);
        }
        next = { ...next };
        if (list.length > 0) (next as MacroSheet)[step.collection] = list;
        else delete (next as MacroSheet)[step.collection];
        break;
      }
    }
  }
  next = { ...next, cells, rowCount, columnCount };
  return { ok: true, sheet: next, touchedCells: touched };
}

/**
 * The steps a run performs, with `run` steps replaced by the named macro's
 * own steps (Excel's `Application.Run`). A macro that cannot be found is
 * skipped; recursion stops at MAX_MACRO_RUN_DEPTH.
 */
export function expandMacroRun(
  macro: SpreadsheetMacro,
  resolve: (name: string) => SpreadsheetMacro | undefined,
  depth = 0,
): MacroStep[] {
  const out: MacroStep[] = [];
  for (const step of macro.steps) {
    if (step.kind === "run") {
      if (depth >= MAX_MACRO_RUN_DEPTH) continue;
      const nested = resolve(step.name);
      if (!nested || nested === macro) continue;
      out.push(...expandMacroRun(nested, resolve, depth + 1));
      continue;
    }
    out.push(step);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Describing steps (the Edit dialog and Step Into)
// ---------------------------------------------------------------------------

export function macroColumnLabel(index: number): string {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

export function macroCellLabel(row: number, col: number): string {
  return `${macroColumnLabel(Math.max(0, col))}${Math.max(0, row) + 1}`;
}

export function macroRangeLabel(row: number, col: number, rows: number, cols: number): string {
  const start = macroCellLabel(row, col);
  if (rows <= 1 && cols <= 1) return start;
  return `${start}:${macroCellLabel(row + rows - 1, col + cols - 1)}`;
}

/** A readable account of a step at `origin`: an id the UI translates plus
 *  its parameters (ranges already spelled in A1). */
export interface MacroStepDescription {
  id:
    | "entry"
    | "cell"
    | "cells"
    | "format"
    | "clear"
    | "insertRows"
    | "insertColumns"
    | "deleteRows"
    | "deleteColumns"
    | "rowHeight"
    | "columnWidth"
    | "defaultRowHeight"
    | "defaultColumnWidth"
    | "hideRow"
    | "hideColumn"
    | "unhideRow"
    | "unhideColumn"
    | "merge"
    | "unmerge"
    | "freeze"
    | "unfreeze"
    | "lineStyle"
    | "select"
    | "sort"
    | "filter"
    | "clearFilter"
    | "addObject"
    | "editObject"
    | "deleteObject"
    | "run";
  params: Record<string, string | number>;
}

function cellsStepBounds(step: Extract<MacroStep, { kind: "cells" }>) {
  let minR = Infinity;
  let minC = Infinity;
  let maxR = -Infinity;
  let maxC = -Infinity;
  let count = 0;
  let allClear = true;
  let single: MacroCellPatch | null | undefined;
  for (const key of Object.keys(step.cells)) {
    const parsed = parseKey(key);
    if (!parsed) continue;
    count += 1;
    minR = Math.min(minR, parsed[0]);
    maxR = Math.max(maxR, parsed[0]);
    minC = Math.min(minC, parsed[1]);
    maxC = Math.max(maxC, parsed[1]);
    if (step.cells[key] !== null) allClear = false;
    single = step.cells[key];
  }
  return { minR, minC, maxR, maxC, count, allClear, single };
}

export function describeMacroStep(
  step: MacroStep,
  origin: { row: number; col: number },
): MacroStepDescription {
  const cell = (dRow: number, dCol: number) => macroCellLabel(origin.row + dRow, origin.col + dCol);
  const range = (dRow: number, dCol: number, rows: number, cols: number) =>
    macroRangeLabel(origin.row + dRow, origin.col + dCol, rows, cols);
  const line = (axis: "row" | "col", at: number) =>
    axis === "row" ? String(origin.row + at + 1) : macroColumnLabel(origin.col + at);
  if (step.kind === undefined) {
    return { id: "entry", params: { cell: cell(step.dRow, step.dCol), value: step.value } };
  }
  switch (step.kind) {
    case "cells": {
      const b = cellsStepBounds(step);
      if (b.count === 0) return { id: "cells", params: { range: "", count: 0 } };
      const rangeLabel = range(b.minR, b.minC, b.maxR - b.minR + 1, b.maxC - b.minC + 1);
      if (b.allClear) return { id: "clear", params: { range: rangeLabel, count: b.count } };
      if (step.formatOnly) return { id: "format", params: { range: rangeLabel, count: b.count } };
      if (b.count === 1 && b.single) {
        const value = b.single.value;
        return {
          id: "cell",
          params: {
            cell: rangeLabel,
            value: typeof value === "string" || typeof value === "number" ? String(value) : "",
          },
        };
      }
      return { id: "cells", params: { range: rangeLabel, count: b.count } };
    }
    case "lines": {
      const id =
        step.op === "insert"
          ? step.axis === "row"
            ? "insertRows"
            : "insertColumns"
          : step.axis === "row"
            ? "deleteRows"
            : "deleteColumns";
      return { id, params: { at: line(step.axis, step.at), count: step.count } };
    }
    case "size":
      return {
        id: step.axis === "row" ? "rowHeight" : "columnWidth",
        params: { at: line(step.axis, step.at), size: step.size === null ? "" : step.size },
      };
    case "defaultSize":
      return {
        id: step.axis === "row" ? "defaultRowHeight" : "defaultColumnWidth",
        params: { size: step.size === null ? "" : step.size },
      };
    case "hidden":
      return {
        id: step.hidden
          ? step.axis === "row"
            ? "hideRow"
            : "hideColumn"
          : step.axis === "row"
            ? "unhideRow"
            : "unhideColumn",
        params: { at: line(step.axis, step.at) },
      };
    case "merges": {
      const first = step.add[0] ?? step.remove[0];
      const label = first ? range(first.dRow, first.dCol, first.rows, first.cols) : "";
      return { id: step.add.length > 0 ? "merge" : "unmerge", params: { range: label } };
    }
    case "freeze":
      return step.rows === 0 && step.cols === 0
        ? { id: "unfreeze", params: {} }
        : { id: "freeze", params: { rows: step.rows, cols: step.cols } };
    case "lineStyle":
      return { id: "lineStyle", params: { at: line(step.axis, step.at) } };
    case "select":
      return { id: "select", params: { range: range(step.dRow, step.dCol, step.rows, step.cols) } };
    case "sort": {
      const first = step.levels[0];
      return {
        id: "sort",
        params: {
          range: range(step.dRow, step.dCol, step.rows, step.cols),
          key: first ? line(step.orientation === "columns" ? "row" : "col", first.at) : "",
          direction: first?.direction ?? "asc",
        },
      };
    }
    case "filters": {
      const rect = step.autoFilter;
      const cleared =
        rect === null ||
        (["filters", "colorFilters", "conditionFilters"] as const).every(
          (field) => !(field in step) || step[field] === null,
        );
      if (rect) return { id: "filter", params: { range: range(rect.dRow, rect.dCol, rect.rows, rect.cols) } };
      return { id: cleared ? "clearFilter" : "filter", params: { range: "" } };
    }
    case "objects":
      return step.added.length > 0
        ? { id: "addObject", params: { count: step.added.length } }
        : step.removed.length > 0 && step.updated.length === 0
          ? { id: "deleteObject", params: { count: step.removed.length } }
          : { id: "editObject", params: { count: step.updated.length + step.removed.length } };
    case "run":
      return { id: "run", params: { name: step.name } };
  }
}

// ---------------------------------------------------------------------------
// Persistence
// ---------------------------------------------------------------------------

function normalizeRect(raw: unknown): MacroRelativeRect | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const dRow = Number(r.dRow);
  const dCol = Number(r.dCol);
  const rows = Number(r.rows);
  const cols = Number(r.cols);
  if (![dRow, dCol, rows, cols].every(Number.isInteger) || rows < 1 || cols < 1) return null;
  return { dRow, dCol, rows, cols };
}

function normalizeAxis(raw: unknown): "row" | "col" | null {
  return raw === "row" || raw === "col" ? raw : null;
}

function normalizeStep(raw: unknown): MacroStep | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const kind = r.kind;
  if (kind === undefined) {
    const dRow = Number(r.dRow);
    const dCol = Number(r.dCol);
    if (!Number.isInteger(dRow) || !Number.isInteger(dCol)) return null;
    if (typeof r.value !== "string") return null;
    return { dRow, dCol, value: r.value };
  }
  switch (kind) {
    case "cells": {
      if (!r.cells || typeof r.cells !== "object") return null;
      const cells: Record<string, MacroCellPatch | null> = {};
      let count = 0;
      for (const key of Object.keys(r.cells as Record<string, unknown>)) {
        if (!parseKey(key)) continue;
        const patch = (r.cells as Record<string, unknown>)[key];
        if (patch === null) cells[key] = null;
        else if (patch && typeof patch === "object" && !Array.isArray(patch)) cells[key] = patch as MacroCellPatch;
        else continue;
        count += 1;
        if (count >= MAX_MACRO_STEP_CELLS) break;
      }
      return r.formatOnly === true ? { kind: "cells", cells, formatOnly: true } : { kind: "cells", cells };
    }
    case "lines": {
      const axis = normalizeAxis(r.axis);
      const at = Number(r.at);
      const count = Number(r.count);
      if (!axis || (r.op !== "insert" && r.op !== "delete")) return null;
      if (!Number.isInteger(at) || !Number.isInteger(count) || count < 1) return null;
      return { kind: "lines", axis, op: r.op, at, count };
    }
    case "size": {
      const axis = normalizeAxis(r.axis);
      const at = Number(r.at);
      if (!axis || !Number.isInteger(at)) return null;
      const size = r.size === null ? null : Number(r.size);
      if (size !== null && !Number.isFinite(size)) return null;
      return { kind: "size", axis, at, size };
    }
    case "defaultSize": {
      const axis = normalizeAxis(r.axis);
      if (!axis) return null;
      const size = r.size === null ? null : Number(r.size);
      if (size !== null && !Number.isFinite(size)) return null;
      return { kind: "defaultSize", axis, size };
    }
    case "hidden": {
      const axis = normalizeAxis(r.axis);
      const at = Number(r.at);
      if (!axis || !Number.isInteger(at)) return null;
      return { kind: "hidden", axis, at, hidden: r.hidden === true };
    }
    case "merges": {
      const add = Array.isArray(r.add) ? r.add.map(normalizeRect).filter((x): x is MacroRelativeRect => x !== null) : [];
      const remove = Array.isArray(r.remove)
        ? r.remove.map(normalizeRect).filter((x): x is MacroRelativeRect => x !== null)
        : [];
      if (add.length === 0 && remove.length === 0) return null;
      return { kind: "merges", add, remove };
    }
    case "freeze": {
      const rows = Number(r.rows);
      const cols = Number(r.cols);
      if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 0 || cols < 0) return null;
      return { kind: "freeze", rows, cols };
    }
    case "lineStyle": {
      const axis = normalizeAxis(r.axis);
      const at = Number(r.at);
      if (!axis || !Number.isInteger(at)) return null;
      return { kind: "lineStyle", axis, at, style: r.style === undefined ? null : r.style };
    }
    case "select": {
      const rect = normalizeRect(r);
      return rect ? { kind: "select", ...rect } : null;
    }
    case "sort": {
      const rect = normalizeRect(r);
      if (!rect) return null;
      const levels: MacroSortLevel[] = [];
      for (const level of Array.isArray(r.levels) ? r.levels : []) {
        if (!level || typeof level !== "object") continue;
        const l = level as Record<string, unknown>;
        const at = Number(l.at);
        if (!Number.isInteger(at) || (l.direction !== "asc" && l.direction !== "desc")) continue;
        const out: MacroSortLevel = { at, direction: l.direction };
        if (typeof l.sortOn === "string") out.sortOn = l.sortOn;
        if (typeof l.formatKey === "string") out.formatKey = l.formatKey;
        if (l.placement === "first" || l.placement === "last") out.placement = l.placement;
        if (Array.isArray(l.orderKeys)) out.orderKeys = l.orderKeys.filter((k): k is string => typeof k === "string");
        levels.push(out);
      }
      if (levels.length === 0) return null;
      const step: MacroStep = { kind: "sort", ...rect, levels };
      if (r.quick === true) step.quick = true;
      if (r.orientation === "rows" || r.orientation === "columns") step.orientation = r.orientation;
      if (typeof r.hasHeader === "boolean") step.hasHeader = r.hasHeader;
      if (typeof r.caseSensitive === "boolean") step.caseSensitive = r.caseSensitive;
      if (typeof r.currentRegion === "boolean") step.currentRegion = r.currentRegion;
      return step;
    }
    case "filters": {
      const step: Extract<MacroStep, { kind: "filters" }> = { kind: "filters" };
      for (const field of ["filters", "colorFilters", "conditionFilters"] as const) {
        if (!(field in r)) continue;
        const value = r[field];
        if (value === null) step[field] = null;
        else if (value && typeof value === "object" && !Array.isArray(value)) step[field] = value as Record<string, unknown>;
      }
      if ("autoFilter" in r) step.autoFilter = r.autoFilter === null ? null : normalizeRect(r.autoFilter);
      return Object.keys(step).length > 1 ? step : null;
    }
    case "objects": {
      const collection = r.collection === "charts" ? "charts" : r.collection === "objects" ? "objects" : null;
      if (!collection) return null;
      const items = (raw: unknown) =>
        Array.isArray(raw)
          ? raw.filter(
              (item): item is Record<string, unknown> =>
                Boolean(item) && typeof item === "object" && typeof (item as { id?: unknown }).id === "string",
            )
          : [];
      const removed = Array.isArray(r.removed) ? r.removed.filter((id): id is string => typeof id === "string") : [];
      const added = items(r.added);
      const updated = items(r.updated);
      if (added.length + updated.length + removed.length === 0) return null;
      return { kind: "objects", collection, added, updated, removed };
    }
    case "run":
      return typeof r.name === "string" && r.name.length > 0 ? { kind: "run", name: r.name } : null;
    default:
      return null;
  }
}

function normalizeMacro(raw: unknown): SpreadsheetMacro | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const name = typeof r.name === "string" ? r.name.trim() : "";
  if (name.length === 0 || name.length > MAX_MACRO_NAME_LENGTH) return null;
  const rawSteps = Array.isArray(r.steps) ? r.steps : [];
  const steps: MacroStep[] = [];
  for (const s of rawSteps) {
    const step = normalizeStep(s);
    if (step) steps.push(step);
    if (steps.length >= MAX_MACRO_STEPS) break;
  }
  // Excel keeps a macro that recorded nothing (an empty Sub).
  const macro: SpreadsheetMacro = { name, steps };
  if (r.mode === "absolute" || r.mode === "relative") macro.mode = r.mode;
  const anchorRow = Number(r.anchorRow);
  const anchorCol = Number(r.anchorCol);
  if (Number.isInteger(anchorRow) && Number.isInteger(anchorCol) && anchorRow >= 0 && anchorCol >= 0) {
    macro.anchorRow = anchorRow;
    macro.anchorCol = anchorCol;
  }
  const shortcut = normalizeShortcut(r.shortcut);
  if (shortcut) macro.shortcut = shortcut;
  if (typeof r.description === "string" && r.description.trim().length > 0) {
    macro.description = r.description.slice(0, MAX_MACRO_DESCRIPTION_LENGTH);
  }
  return macro;
}

/** Normalize a persisted (untrusted) macros list: drop malformed entries,
 *  dedupe by case-insensitive name (last wins), cap the count. */
export function normalizeMacros(raw: unknown): SpreadsheetMacro[] {
  if (!Array.isArray(raw)) return [];
  const byName = new Map<string, SpreadsheetMacro>();
  for (const item of raw) {
    const macro = normalizeMacro(item);
    if (!macro) continue;
    byName.set(macro.name.toLowerCase(), macro);
  }
  return [...byName.values()].slice(0, MAX_MACROS);
}

/** Upsert by case-insensitive name (Excel macro names are case-insensitive). */
export function upsertMacro(
  macros: readonly SpreadsheetMacro[],
  macro: SpreadsheetMacro,
): SpreadsheetMacro[] {
  const key = macro.name.toLowerCase();
  const without = macros.filter((m) => m.name.toLowerCase() !== key);
  return [...without, macro].slice(-MAX_MACROS);
}

export function removeMacro(
  macros: readonly SpreadsheetMacro[],
  name: string,
): SpreadsheetMacro[] {
  const key = name.toLowerCase();
  return macros.filter((m) => m.name.toLowerCase() !== key);
}

export function findMacro(
  macros: readonly SpreadsheetMacro[],
  name: string,
): SpreadsheetMacro | undefined {
  const key = name.toLowerCase();
  return macros.find((m) => m.name.toLowerCase() === key);
}

// ---------------------------------------------------------------------------
// The Personal Macro Workbook
// ---------------------------------------------------------------------------

/** Excel's PERSONAL.XLSB is a per-user store every workbook sees; ours lives
 *  in this browser's localStorage. */
export const PERSONAL_MACROS_STORAGE_KEY = "mrai.spreadsheet.personalMacros.v1";
export const PERSONAL_MACRO_WORKBOOK_NAME = "PERSONAL.XLSB";

export function readPersonalMacros(storage: Pick<Storage, "getItem"> | null | undefined): SpreadsheetMacro[] {
  try {
    const raw = storage?.getItem(PERSONAL_MACROS_STORAGE_KEY);
    if (!raw) return [];
    return normalizeMacros(JSON.parse(raw));
  } catch {
    return [];
  }
}

export function writePersonalMacros(
  storage: Pick<Storage, "setItem" | "removeItem"> | null | undefined,
  macros: readonly SpreadsheetMacro[],
): boolean {
  try {
    if (!storage) return false;
    if (macros.length === 0) storage.removeItem(PERSONAL_MACROS_STORAGE_KEY);
    else storage.setItem(PERSONAL_MACROS_STORAGE_KEY, JSON.stringify(macros));
    return true;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// Assign Macro (XLS-A081)
// ---------------------------------------------------------------------------

/**
 * What a drawing object stores for Excel's Assign Macro: the macro's name for
 * one of this workbook's, `PERSONAL.XLSB!Name` for one in the Personal Macro
 * Workbook (Excel writes `[0]!Name` / `PERSONAL.XLSB!Name` in `@macro`).
 */
export function assignedMacroReference(store: "workbook" | "personal" | "newWorkbook", name: string): string {
  return store === "personal" ? `${PERSONAL_MACRO_WORKBOOK_NAME}!${name}` : name;
}

/** The store and plain name an assigned reference names. */
export function parseAssignedMacroReference(
  reference: string,
): { store: "workbook" | "personal"; name: string } | null {
  const trimmed = reference.trim();
  if (!trimmed) return null;
  const prefix = `${PERSONAL_MACRO_WORKBOOK_NAME}!`;
  if (trimmed.toUpperCase().startsWith(prefix)) {
    const name = trimmed.slice(prefix.length);
    return isValidMacroName(name) ? { store: "personal", name } : null;
  }
  // A module-qualified name (`Sheet1.Macro1`) names the macro after the dot.
  const name = trimmed.includes(".") ? trimmed.slice(trimmed.lastIndexOf(".") + 1) : trimmed;
  return isValidMacroName(name) ? { store: "workbook", name } : null;
}

/**
 * The macro an assigned reference runs, or undefined when it is gone (Excel
 * then says it cannot run the macro): a workbook reference is resolved the
 * way the Macro dialog resolves a name, a personal one in PERSONAL.XLSB.
 */
export function resolveAssignedMacro(
  reference: string,
  resolveWorkbookMacro: (name: string) => SpreadsheetMacro | undefined,
  personal: readonly SpreadsheetMacro[],
): SpreadsheetMacro | undefined {
  const parsed = parseAssignedMacroReference(reference);
  if (!parsed) return undefined;
  return parsed.store === "personal"
    ? findMacro(personal, parsed.name)
    : resolveWorkbookMacro(parsed.name);
}

/** Excel proposes `<ObjectName>_Click` (spaces dropped: "Rectangle1_Click"). */
export function defaultAssignMacroName(objectName: string): string {
  const base = objectName.replace(/[^\p{L}\p{N}_]/gu, "");
  const candidate = `${base || "Object"}_Click`;
  return isValidMacroName(candidate) ? candidate : "Object_Click";
}

/** A click that runs an object's macro: the left button, no Ctrl / Cmd /
 *  Shift (those select the object instead, as in Excel), and no drag. */
export function isAssignedMacroClick(
  press: { x: number; y: number } | null,
  click: { x: number; y: number; button?: number; ctrlKey: boolean; metaKey: boolean; shiftKey: boolean },
): boolean {
  return Boolean(
    press &&
      (click.button ?? 0) === 0 &&
      !click.ctrlKey &&
      !click.metaKey &&
      !click.shiftKey &&
      Math.hypot(click.x - press.x, click.y - press.y) <= 4,
  );
}
