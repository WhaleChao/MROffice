/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for the spreadsheet Scenario Manager (Excel/WPS "Data ▸ What-If
 * Analysis ▸ Scenario Manager").
 *
 * A SCENARIO is a NAMED set of {cell → value} overrides the user can save and
 * re-apply to compare what-if outcomes (e.g. "Best case" sets B2=1000, B3=0.2;
 * "Worst case" sets B2=400, B3=0.05). Only the recipe — the name + the per-cell
 * literal values — is persisted; applying a scenario writes those values into
 * the grid through the normal edit path, so the dependent formulas recompute.
 *
 * Kept dependency-light + deterministic (like goal-seek.ts and named-ranges.ts)
 * so capture / apply / normalize are unit-testable without a React grid. The
 * shared dependencies are parseCellRef from the formula engine, so a stored
 * A1 ref parses to coordinates exactly as a literal cell reference would, and
 * the grid's entry parser, so a Scenario Values box reads like a cell entry.
 */

import {
  entryKeepsText,
  parseSpreadsheetEntry,
  type SpreadsheetEntryOptions,
} from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import {
  isFormulaError,
  parseCellRef,
  type CellValue,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

/** One {cell → value} override inside a scenario. */
export interface ScenarioCell {
  /** The A1 cell reference the override targets ("B2"). Single cell, not a range. */
  ref: string;
  /**
   * The value to put in the cell when the scenario is shown: a canonical
   * number ("1234.56", never the "1,234.56" a formatted cell displays) or the
   * text as entered. Show reads it the way a cell entry is read.
   */
  value: string;
}

/** A named set of cell overrides (one "what-if" case). */
export interface Scenario {
  /** Human-friendly name shown in the manager ("Best case"). */
  name: string;
  /** The per-cell overrides this scenario applies. */
  cells: ScenarioCell[];
  /** Excel's Comment box ("Created by … on …"); absent when empty. */
  comment?: string;
  /** Protection > Prevent changes: on a protected sheet the scenario can be
   * neither edited nor deleted. Only `true` is stored. */
  locked?: boolean;
  /** Protection > Hide: on a protected sheet the scenario is not listed.
   * Only `true` is stored. */
  hidden?: boolean;
}

/** A cell write to apply to the grid (0-based coordinates + the literal value). */
export interface ScenarioWrite {
  row: number;
  col: number;
  value: string;
}

// Bounds — mirror the schema (native-page-schemas.ts) so the runtime normalize
// and the persisted schema agree on the caps. A scenario name is capped at 255
// (Excel's defined-name cap), a ref string at 64 (same as a namedRange ref), a
// value at 32_767 (Excel's per-cell text cap), and the per-scenario cell count
// and the per-sheet scenario count at 256 each (DoS guard on iteration/apply).
export const MAX_SCENARIO_NAME_LENGTH = 255;
export const MAX_SCENARIO_REF_LENGTH = 64;
export const MAX_SCENARIO_VALUE_LENGTH = 32_767;
export const MAX_SCENARIO_CELLS = 256;
export const MAX_SCENARIOS = 256;
/** Excel's Add Scenario accepts at most 32 changing cells ("Each scenario can
 * accommodate up to 32 variable values"). Stored scenarios keep the looser
 * MAX_SCENARIO_CELLS so an imported file is never truncated. */
export const MAX_SCENARIO_CHANGING_CELLS = 32;
/** Excel caps a scenario comment at 255 characters. */
export const MAX_SCENARIO_COMMENT_LENGTH = 255;

/**
 * The text a Scenario Values box starts from for a changing cell whose current
 * value is `value`. Excel shows the cell's VALUE there, not its formatted text,
 * and a formula cell's current result (the formula is replaced by that constant
 * when the scenario is shown): a number in its plain form (1234.56 for a
 * #,##0.00 cell showing "1,234.56"; 15 significant digits, as Excel keeps),
 * TRUE/FALSE, an error code, or the text itself. Text that a cell entry would
 * read as something else ("00123", "true", "=x", "'a") gets a leading
 * apostrophe, so reading the box back with scenarioEntryValue keeps it text.
 */
export function scenarioInputText(
  value: CellValue | undefined,
  options: SpreadsheetEntryOptions,
): string {
  if (value === null || value === undefined) return "";
  if (isFormulaError(value)) return value.code;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") {
    return Number.isFinite(value)
      ? String(Number(value.toPrecision(15)))
      : "#NUM!";
  }
  return entryKeepsText(value, options) ? value : `'${value}`;
}

/**
 * What a Scenario Values box stores, read the way a cell entry is read
 * (parseSpreadsheetEntry, the grid's own entry parser): "1,234.56", "50%",
 * a currency amount and a date are numbers and are stored canonical
 * ("1234.56", "0.5", the amount, the date serial), so the model and the
 * exported <inputCells val> carry the number itself. Everything else is kept
 * as typed — text (with its leading apostrophe, so Show keeps it text),
 * logicals, error codes, and a leading "=" (Show writes it as a literal, see
 * applyScenario).
 */
export function scenarioEntryValue(
  raw: string,
  options: SpreadsheetEntryOptions,
): string {
  if (raw.trim().startsWith("=")) return raw;
  const entry = parseSpreadsheetEntry(raw, options);
  return entry.kind === "number" ? entry.value : raw;
}

/**
 * Snapshot the CURRENT values of `refs` into a named scenario. The
 * `readDisplay(ref)` closure returns the text for that A1 cell (the Scenario
 * Manager passes scenarioInputText over the live evaluator: the cell's value
 * or a formula's result, never a raw formula and never the number-formatted
 * text the grid paints). A ref that does not parse to a single cell is
 * skipped (a range has no single value to snapshot); duplicate refs collapse
 * to one entry (last wins) so a captured scenario never double-writes the
 * same cell. The name is trimmed; the order of the FIRST occurrence of each
 * ref is preserved.
 */
export function captureScenario(
  refs: readonly string[],
  readDisplay: (ref: string) => string,
  name: string,
): Scenario {
  const byRef = new Map<string, ScenarioCell>();
  for (const raw of refs) {
    const parsed = parseCellRef(String(raw ?? ""));
    if (!parsed) continue; // not a single A1 cell -> nothing to snapshot
    // Canonicalize the key so "B2"/"$B$2"/"b2" collapse to one entry; the stored
    // ref is the canonical "B2" form (column letters + 1-based row).
    const canonical = `${columnLetters(parsed.col)}${parsed.row + 1}`;
    if (canonical.length > MAX_SCENARIO_REF_LENGTH) continue;
    const value = String(readDisplay(raw) ?? "").slice(0, MAX_SCENARIO_VALUE_LENGTH);
    // last write of a repeated ref wins, but keep the first-seen position by
    // re-inserting at the same key (Map preserves first-insertion order).
    byRef.set(canonical, { ref: canonical, value });
  }
  const cells = [...byRef.values()].slice(0, MAX_SCENARIO_CELLS);
  return { name: name.trim().slice(0, MAX_SCENARIO_NAME_LENGTH), cells };
}

/**
 * Resolve a scenario into the concrete cell writes to apply to the grid. Each
 * stored A1 ref parses to 0-based coordinates; a ref that no longer parses (it
 * was dropped to "#REF!" by a structural delete, or is otherwise malformed) is
 * skipped rather than crashing. Duplicate coordinates collapse (last wins) so
 * applying never queues two writes for one cell.
 */
export function applyScenario(scenario: Scenario): ScenarioWrite[] {
  const byKey = new Map<string, ScenarioWrite>();
  for (const cell of scenario.cells ?? []) {
    const parsed = parseCellRef(String(cell?.ref ?? ""));
    if (!parsed) continue; // erased / malformed ref -> nothing to write
    byKey.set(`${parsed.row}:${parsed.col}`, {
      row: parsed.row,
      col: parsed.col,
      value: literalScenarioValue(String(cell?.value ?? "")),
    });
  }
  return [...byKey.values()];
}

/**
 * A scenario pins a LITERAL what-if input. A stored value that begins with "="
 * (from an imported / hand-edited / AI-written file, or typed into a
 * Scenario Values box — the prefill is a computed value) would otherwise be
 * written as a live formula by the grid's edit path. Strip the leading "="
 * so applying a scenario can never inject a formula into the cell it pins.
 */
function literalScenarioValue(value: string): string {
  return value.startsWith("=") ? value.replace(/^=+/, "") : value;
}

/**
 * Coerce a single raw entry into a normalized Scenario, or null when it has no
 * usable name or no usable cells. Mirrors normalizeNamedRange: the name is
 * trimmed + length-clamped; each cell's ref must parse to a single A1 cell (it
 * is canonicalized to "B2" form) and its value is string-coerced + clamped; an
 * unparseable ref drops its cell. A scenario with a non-empty name but zero
 * surviving cells is KEPT (an empty scenario is harmless — applying it writes
 * nothing — and Excel allows a scenario that overrides no cells), but a missing/
 * blank name drops the whole scenario (it could never be selected by name).
 */
export function normalizeScenario(raw: unknown): Scenario | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const name = typeof r.name === "string" ? r.name.trim() : "";
  if (name.length === 0) return null;
  const rawCells = Array.isArray(r.cells) ? r.cells : [];
  const byRef = new Map<string, ScenarioCell>();
  for (const item of rawCells) {
    if (!item || typeof item !== "object") continue;
    const c = item as Record<string, unknown>;
    const parsed = parseCellRef(typeof c.ref === "string" ? c.ref : "");
    if (!parsed) continue; // not a single A1 cell -> drop
    const canonical = `${columnLetters(parsed.col)}${parsed.row + 1}`;
    if (canonical.length > MAX_SCENARIO_REF_LENGTH) continue;
    const value =
      typeof c.value === "string"
        ? c.value.slice(0, MAX_SCENARIO_VALUE_LENGTH)
        : c.value == null
          ? ""
          : String(c.value).slice(0, MAX_SCENARIO_VALUE_LENGTH);
    byRef.set(canonical, { ref: canonical, value });
  }
  const scenario: Scenario = {
    name: name.slice(0, MAX_SCENARIO_NAME_LENGTH),
    cells: [...byRef.values()].slice(0, MAX_SCENARIO_CELLS),
  };
  const comment =
    typeof r.comment === "string"
      ? r.comment.slice(0, MAX_SCENARIO_COMMENT_LENGTH)
      : "";
  if (comment.trim() !== "") scenario.comment = comment;
  if (r.locked === true) scenario.locked = true;
  if (r.hidden === true) scenario.hidden = true;
  return scenario;
}

/**
 * Normalize a raw scenarios array off a persisted sheet: drop malformed entries
 * and de-duplicate by case-insensitive name (last one wins, matching the
 * add-overwrites-existing behavior used by named ranges). The list length is
 * clamped to MAX_SCENARIOS. Returns a fresh array; an absent/non-array input
 * yields [] so a sheet without scenarios round-trips cleanly.
 */
export function normalizeScenarios(raw: unknown): Scenario[] {
  if (!Array.isArray(raw)) return [];
  const byName = new Map<string, Scenario>();
  for (const item of raw) {
    const entry = normalizeScenario(item);
    if (!entry) continue;
    byName.set(entry.name.toUpperCase(), entry);
  }
  return [...byName.values()].slice(0, MAX_SCENARIOS);
}

/**
 * Add (or overwrite) a scenario by case-insensitive name. De-duplicates so a
 * re-captured scenario with the same name REPLACES the old one (Excel prompts to
 * overwrite; we replace). The new/updated entry is appended at the end (a fresh
 * capture moves to the bottom of the list, like named ranges). Returns a fresh
 * array; the input is never mutated.
 */
export function upsertScenario(
  scenarios: readonly Scenario[],
  scenario: Scenario,
): Scenario[] {
  const key = scenario.name.trim().toUpperCase();
  const kept = scenarios.filter((s) => s.name.trim().toUpperCase() !== key);
  return [...kept, scenario].slice(0, MAX_SCENARIOS);
}

/**
 * Replace the scenario named `originalName` in place (Edit Scenario keeps the
 * scenario's position in the list, even when it is renamed). Falls back to an
 * append when the original is gone. Returns a fresh array.
 */
export function replaceScenario(
  scenarios: readonly Scenario[],
  originalName: string,
  scenario: Scenario,
): Scenario[] {
  const key = originalName.trim().toUpperCase();
  const at = scenarios.findIndex((s) => s.name.trim().toUpperCase() === key);
  if (at < 0) return upsertScenario(scenarios, scenario);
  const next = scenarios.slice();
  next[at] = scenario;
  return next;
}

/**
 * The single-cell refs a Changing cells box names, in reading order: each
 * area expands row by row, repeats collapse to their first occurrence.
 * Returns canonical "B2" refs; the caller enforces
 * MAX_SCENARIO_CHANGING_CELLS (the expansion stops one past it so a whole
 * column cannot build a million-entry list).
 */
export function scenarioChangingCellRefs(
  areas: readonly { r1: number; c1: number; r2: number; c2: number }[],
): string[] {
  const refs: string[] = [];
  const seen = new Set<string>();
  for (const area of areas) {
    for (let row = area.r1; row <= area.r2; row += 1) {
      for (let col = area.c1; col <= area.c2; col += 1) {
        const ref = `${columnLetters(col)}${row + 1}`;
        if (seen.has(ref)) continue;
        seen.add(ref);
        refs.push(ref);
        if (refs.length > MAX_SCENARIO_CHANGING_CELLS) return refs;
      }
    }
  }
  return refs;
}

/**
 * A scenario's cells as Excel's Changing cells text, a run of neighbouring
 * cells written as one range: "$B$2:$B$4,$D$2". Reading it back through the
 * Changing cells parser yields the same cells in the same order.
 */
export function formatScenarioChangingCells(
  cells: readonly ScenarioCell[],
): string {
  const absolute = (row: number, col: number) =>
    `$${columnLetters(col)}$${row + 1}`;
  const parts: string[] = [];
  let run: {
    row: number;
    col: number;
    endRow: number;
    endCol: number;
    axis: "row" | "col" | null;
  } | null = null;
  const flush = () => {
    if (!run) return;
    const start = absolute(run.row, run.col);
    parts.push(
      run.endRow === run.row && run.endCol === run.col
        ? start
        : `${start}:${absolute(run.endRow, run.endCol)}`,
    );
    run = null;
  };
  for (const cell of cells) {
    const parsed = parseCellRef(cell.ref);
    if (!parsed) {
      flush();
      parts.push(cell.ref);
      continue;
    }
    if (run) {
      const down =
        run.axis !== "row" &&
        parsed.col === run.endCol &&
        parsed.row === run.endRow + 1;
      const right =
        run.axis !== "col" &&
        parsed.row === run.endRow &&
        parsed.col === run.endCol + 1;
      if (down || right) {
        run.endRow = parsed.row;
        run.endCol = parsed.col;
        run.axis = down ? "col" : "row";
        continue;
      }
      flush();
    }
    run = {
      row: parsed.row,
      col: parsed.col,
      endRow: parsed.row,
      endCol: parsed.col,
      axis: null,
    };
  }
  flush();
  return parts.join(",");
}

/**
 * Scenario Manager's Merge… (Microsoft Support, "Switch between various sets
 * of values by using scenarios"): every scenario of the source sheet is copied
 * onto the target with its changing cells, values, comment and protection. A
 * name the target already carries gets a number appended ("Best 2", "Best 3"),
 * one of the suffixes Excel uses, so nothing is overwritten; Edit… renames it.
 * Copies stop at MAX_SCENARIOS. Returns a fresh list and how many were added.
 */
export function mergeScenarios(
  target: readonly Scenario[],
  source: readonly Scenario[],
): { scenarios: Scenario[]; added: number } {
  const next = target.slice();
  const taken = new Set(next.map((entry) => entry.name.trim().toUpperCase()));
  let added = 0;
  for (const scenario of source) {
    if (next.length >= MAX_SCENARIOS) break;
    const base = scenario.name.trim();
    let name = base;
    for (let suffix = 2; taken.has(name.toUpperCase()); suffix += 1) {
      const tail = ` ${suffix}`;
      name = `${base.slice(0, MAX_SCENARIO_NAME_LENGTH - tail.length)}${tail}`;
    }
    taken.add(name.toUpperCase());
    next.push({
      ...scenario,
      name,
      cells: scenario.cells.map((cell) => ({ ...cell })),
    });
    added += 1;
  }
  return { scenarios: next, added };
}

/** Remove a scenario by case-insensitive name. Returns a fresh array. */
export function removeScenario(
  scenarios: readonly Scenario[],
  name: string,
): Scenario[] {
  const key = name.trim().toUpperCase();
  return scenarios.filter((s) => s.name.trim().toUpperCase() !== key);
}

/** "A1" letters for a 0-based column index (0 -> "A", 26 -> "AA"). */
function columnLetters(index: number): string {
  let label = "";
  let value = index + 1;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}
