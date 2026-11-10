/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What Review > Spelling (F7) walks in a worksheet, in Excel's order, and how
 * a correction is written back.
 *
 * Microsoft Support, "Check spelling in a worksheet": with a single cell
 * selected Excel checks the entire worksheet, "including the comments, page
 * headers, footers, and graphics"; with several cells selected only those
 * cells; never cells that hold formulas, though words selected in a formula
 * being edited are checked. `Range.CheckSpelling` (Microsoft Learn) adds that
 * the cells' notes go with the cells. The check starts at the active cell and
 * runs row by row to the end of the sheet; when it did not start at A1, Excel
 * asks whether to continue at the beginning of the sheet. Hidden rows and
 * columns are checked like any other.
 *
 * The walk is lazy: the plan is a list of targets (cheap: a sort of the
 * sheet's text-cell keys), and each target's text is read from the live sheet
 * only when the checker reaches it, so a correction or a collaborator's edit
 * made meanwhile is what gets checked. Chart text is not walked: a chart's
 * titles are edited through its chart model, which this command does not
 * drive. Threaded comments are conversations, not the notes Excel checks.
 */

import { replaceRangeInRuns } from "@mrnote/slides/web/lib/presentation-run-ranges.ts";
import {
  adaptSpellingReplacement,
  findSpellingIssue,
  spellingOccurrences,
  type SpellingCheckContext,
  type SpellingIssue,
  type SpellingTextRange,
} from "@mrnote/core/web/lib/proofing/spelling-checker.ts";
import type { SpellingOptions } from "@mrnote/core/web/lib/proofing/proofing-preferences.ts";

import { migrateLegacyComment, setNote } from "@mrnote/core/web/lib/spreadsheet/comments.ts";
import { spreadsheetCellIsText } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import { isExcelNumericText } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { headerFooterRunsForEditedText } from "./header-footer-rich-text";
import { shapeTextPatch } from "./object-buttons";
import { normalizePrintSettings, type PrintHeaderFooter, type PrintSettings } from "@mrnote/core/web/lib/spreadsheet/print-settings.ts";

export interface SpellingCellLike {
  value?: unknown;
  formula?: unknown;
  kind?: unknown;
  runs?: ReadonlyArray<{ text: string; style?: unknown }>;
  style?: { quotePrefix?: unknown; [key: string]: unknown };
  comment?: unknown;
  [key: string]: unknown;
}

export interface SpellingObjectLike {
  id: string;
  type?: unknown;
  shape?: unknown;
}

export interface SpellingSheetLike {
  cells: Record<string, SpellingCellLike | undefined>;
  objects?: readonly SpellingObjectLike[];
  printSettings?: unknown;
}

export type HeaderFooterField =
  | "header"
  | "footer"
  | "firstHeader"
  | "firstFooter"
  | "evenHeader"
  | "evenFooter";
export type HeaderFooterAlign = "left" | "center" | "right";

const HEADER_FOOTER_FIELDS: readonly HeaderFooterField[] = [
  "header",
  "footer",
  "firstHeader",
  "firstFooter",
  "evenHeader",
  "evenFooter",
];
const HEADER_FOOTER_ALIGNS: readonly HeaderFooterAlign[] = ["left", "center", "right"];

export type SpellingTarget =
  | {
      kind: "cell";
      key: string;
      row: number;
      col: number;
      /** The cell being edited: its entry, a formula's source included. */
      entry?: true;
      /** Only these characters (words selected in the entry). */
      range?: SpellingTextRange;
    }
  | { kind: "note"; key: string; row: number; col: number }
  | { kind: "shape"; objectId: string }
  | { kind: "header-footer"; field: HeaderFooterField; align: HeaderFooterAlign };

export type SpellingScope =
  | { kind: "sheet"; start: { row: number; col: number } }
  | { kind: "range"; rect: { r1: number; c1: number; r2: number; c2: number } }
  | {
      kind: "entry";
      row: number;
      col: number;
      range?: SpellingTextRange;
      /** The entry's text when the range was taken; a commit that stored
       *  something else (a formula's canonical form) drops the range. */
      draft?: string;
    };

export interface SpellingPlan {
  /** From the start cell to the end of the sheet, then the sheet's notes,
   *  shapes and headers / footers. */
  forward: readonly SpellingTarget[];
  /** The cells before the start cell, checked after Excel's prompt. */
  wrapped: readonly SpellingTarget[];
  /** Excel's "continue checking at the beginning of the sheet?". */
  askBeforeWrapping: boolean;
}

/**
 * The scope Excel checks for a selection: one cell (a merged cell counts as
 * one) is the whole sheet from the active cell; several cells are just those;
 * a cell being edited is its own text.
 */
export function spellingScopeFor(input: {
  selection: { r1: number; c1: number; r2: number; c2: number };
  active: { row: number; col: number };
  singleCell: boolean;
  entry?: { row: number; col: number; range?: SpellingTextRange; draft?: string } | null;
}): SpellingScope {
  if (input.entry) {
    const range = input.entry.range;
    return {
      kind: "entry",
      row: input.entry.row,
      col: input.entry.col,
      ...(range && range.end > range.start
        ? {
            range,
            ...(input.entry.draft !== undefined ? { draft: input.entry.draft } : {}),
          }
        : {}),
    };
  }
  if (input.singleCell) return { kind: "sheet", start: { ...input.active } };
  const { r1, c1, r2, c2 } = input.selection;
  return {
    kind: "range",
    rect: {
      r1: Math.min(r1, r2),
      c1: Math.min(c1, c2),
      r2: Math.max(r1, r2),
      c2: Math.max(c1, c2),
    },
  };
}

const ERROR_LITERAL = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|GETTING_DATA|BUSY!|FIELD!|BLOCKED!|CONNECT!|UNKNOWN!|EXTERNAL!)$/i;

function textValue(cell: SpellingCellLike | undefined): string {
  return typeof cell?.value === "string" ? cell.value : "";
}

/** A cell entry typed as text: Text format, a quote prefix, a text kind. */
function textIdentity(cell: SpellingCellLike): boolean {
  return cell.kind === "text" || cell.style?.quotePrefix === true;
}

export function isSpellingFormulaCell(cell: SpellingCellLike | undefined): boolean {
  if (!cell) return false;
  if (typeof cell.formula === "string" && cell.formula.trim().startsWith("=")) return true;
  if (textIdentity(cell)) return false;
  return textValue(cell).trim().startsWith("=");
}

/**
 * Whether a cell holds a text constant, the only cell content Excel checks:
 * not a formula, a number, a date, a logical or an error value.
 */
export function isSpellingTextCell(cell: SpellingCellLike | undefined): boolean {
  if (!cell) return false;
  const value = textValue(cell);
  if (!value.trim() || isSpellingFormulaCell(cell)) return false;
  // Typed as text ('TRUE, an @ cell, a file string). A plain TRUE or #N/A is
  // stored with kind "text" too, yet it is a logical or an error (S-041).
  if (spreadsheetCellIsText(cell as Parameters<typeof spreadsheetCellIsText>[0])) return true;
  if (cell.kind === "number") return false;
  const trimmed = value.trim();
  if (ERROR_LITERAL.test(trimmed)) return false;
  if (/^(?:TRUE|FALSE)$/i.test(trimmed)) return false;
  return !isExcelNumericText(trimmed);
}

function cellSource(cell: SpellingCellLike): string {
  if (typeof cell.formula === "string" && cell.formula.trim().startsWith("=")) {
    return cell.formula;
  }
  return textValue(cell);
}

function noteText(cell: SpellingCellLike | undefined): string | null {
  const note = migrateLegacyComment(cell?.comment);
  if (note?.kind !== "note" || !note.text.trim() || note.nativeBlocker) return null;
  return note.text;
}

function shapeText(object: SpellingObjectLike | undefined): string | null {
  if (!object || object.type !== "shape") return null;
  const shape = object.shape as { text?: unknown } | undefined;
  return typeof shape?.text === "string" && shape.text.trim() ? shape.text : null;
}

function printSettingsOf(sheet: SpellingSheetLike): PrintSettings {
  return normalizePrintSettings(sheet.printSettings) ?? {};
}

function headerFooterText(
  settings: PrintSettings,
  field: HeaderFooterField,
  align: HeaderFooterAlign,
): string | null {
  const band = settings[field] as PrintHeaderFooter | undefined;
  const text = band?.[align];
  return typeof text === "string" && text.trim() ? text : null;
}

function parseKey(key: string): { row: number; col: number } | null {
  const separator = key.indexOf(":");
  if (separator < 1) return null;
  const row = Number(key.slice(0, separator));
  const col = Number(key.slice(separator + 1));
  return Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0
    ? { row, col }
    : null;
}

interface KeyedCell {
  key: string;
  row: number;
  col: number;
}

function sortedKeys(
  cells: SpellingSheetLike["cells"],
  keep: (cell: SpellingCellLike, at: KeyedCell) => boolean,
): KeyedCell[] {
  const found: KeyedCell[] = [];
  for (const key of Object.keys(cells)) {
    const cell = cells[key];
    if (!cell) continue;
    const at = parseKey(key);
    if (!at) continue;
    const entry = { key, ...at };
    if (keep(cell, entry)) found.push(entry);
  }
  return found.sort((left, right) => left.row - right.row || left.col - right.col);
}

/**
 * The targets a check visits, in Excel's order. `skipCellKeys` are cells
 * whose shown value is not their own (a spilled array's cells).
 */
export function buildSpellingPlan(
  sheet: SpellingSheetLike,
  scope: SpellingScope,
  options: { skipCellKeys?: { has(key: string): boolean } } = {},
): SpellingPlan {
  const skip = options.skipCellKeys;
  const textCell = (cell: SpellingCellLike, at: KeyedCell) =>
    !skip?.has(at.key) && isSpellingTextCell(cell);
  if (scope.kind === "entry") {
    const key = `${scope.row}:${scope.col}`;
    const cell = sheet.cells[key];
    const source = cell ? cellSource(cell) : "";
    const range =
      scope.range && (scope.draft === undefined || scope.draft === source)
        ? scope.range
        : undefined;
    // A formula's words are checked only when they are selected.
    const checkable =
      cell !== undefined &&
      source.trim() !== "" &&
      (!isSpellingFormulaCell(cell) || range !== undefined);
    return {
      forward: checkable
        ? [
            {
              kind: "cell",
              key,
              row: scope.row,
              col: scope.col,
              entry: true,
              ...(range ? { range } : {}),
            },
          ]
        : [],
      wrapped: [],
      askBeforeWrapping: false,
    };
  }
  if (scope.kind === "range") {
    const { r1, c1, r2, c2 } = scope.rect;
    const inside = (at: KeyedCell) =>
      at.row >= r1 && at.row <= r2 && at.col >= c1 && at.col <= c2;
    const cells = sortedKeys(sheet.cells, (cell, at) => inside(at) && textCell(cell, at));
    const notes = sortedKeys(sheet.cells, (cell, at) => inside(at) && noteText(cell) !== null);
    return {
      forward: [
        ...cells.map((at): SpellingTarget => ({ kind: "cell", ...at })),
        ...notes.map((at): SpellingTarget => ({ kind: "note", ...at })),
      ],
      wrapped: [],
      askBeforeWrapping: false,
    };
  }
  const { row: startRow, col: startCol } = scope.start;
  const cells = sortedKeys(sheet.cells, textCell);
  const fromStart = (at: KeyedCell) =>
    at.row > startRow || (at.row === startRow && at.col >= startCol);
  const notes = sortedKeys(sheet.cells, (cell) => noteText(cell) !== null);
  const shapes = (sheet.objects ?? []).filter((object) => shapeText(object) !== null);
  const settings = printSettingsOf(sheet);
  const bands: SpellingTarget[] = [];
  for (const field of HEADER_FOOTER_FIELDS) {
    for (const align of HEADER_FOOTER_ALIGNS) {
      if (headerFooterText(settings, field, align) !== null) {
        bands.push({ kind: "header-footer", field, align });
      }
    }
  }
  const wrapped = cells
    .filter((at) => !fromStart(at))
    .map((at): SpellingTarget => ({ kind: "cell", ...at }));
  return {
    forward: [
      ...cells.filter(fromStart).map((at): SpellingTarget => ({ kind: "cell", ...at })),
      ...notes.map((at): SpellingTarget => ({ kind: "note", ...at })),
      ...shapes.map((object): SpellingTarget => ({ kind: "shape", objectId: object.id })),
      ...bands,
    ],
    wrapped,
    askBeforeWrapping: wrapped.length > 0 && (startRow !== 0 || startCol !== 0),
  };
}

/** The target's text as stored now, or null when it no longer holds any to
 *  check (the cell became a formula, the note was deleted). */
export function readSpellingTarget(
  sheet: SpellingSheetLike,
  target: SpellingTarget,
): string | null {
  switch (target.kind) {
    case "cell": {
      const cell = sheet.cells[target.key];
      if (!cell) return null;
      if (target.entry) {
        const source = cellSource(cell);
        return source.trim() ? source : null;
      }
      return isSpellingTextCell(cell) ? textValue(cell) : null;
    }
    case "note":
      return noteText(sheet.cells[target.key]);
    case "shape":
      return shapeText(sheet.objects?.find((object) => object.id === target.objectId));
    case "header-footer":
      return headerFooterText(printSettingsOf(sheet), target.field, target.align);
  }
}

/** Header / footer field codes (&[Page], &[Tab]) are never words. */
function maskFieldCodes(text: string): string {
  return text.replace(/&\[[^\]]*\]/g, (code) => " ".repeat(code.length));
}

function checkedText(target: SpellingTarget, text: string): string {
  return target.kind === "header-footer" ? maskFieldCodes(text) : text;
}

function targetRange(target: SpellingTarget): SpellingTextRange | undefined {
  return target.kind === "cell" ? target.range : undefined;
}

/** Where the walk stands: which pass, which target, which character. */
export interface SpellingCursor {
  pass: "forward" | "wrapped";
  index: number;
  offset: number;
}

export const SPELLING_START: SpellingCursor = { pass: "forward", index: 0, offset: 0 };

export interface SpellingFinding extends SpellingIssue {
  target: SpellingTarget;
  /** The cursor that found it (re-finding from here yields it again). */
  cursor: SpellingCursor;
}

export type SpellingStep =
  | { status: "issue"; finding: SpellingFinding }
  /** The forward pass ended and Excel asks before checking from A1. */
  | { status: "wrap"; cursor: SpellingCursor }
  | { status: "done" }
  /** The time budget ran out; resume from `cursor`. */
  | { status: "paused"; cursor: SpellingCursor };

/**
 * The next issue from `cursor`, reading each target from `sheet`. Stops early
 * with "paused" once `deadline()` says so, so a caller can yield to the
 * browser between slices of a large sheet.
 */
export function findNextSpellingIssue(input: {
  sheet: SpellingSheetLike;
  plan: SpellingPlan;
  cursor: SpellingCursor;
  context: SpellingCheckContext;
  /** True when the slice should end. Checked between targets. */
  deadline?: () => boolean;
}): SpellingStep {
  const { sheet, plan, context, deadline } = input;
  let { pass, index, offset } = input.cursor;
  for (;;) {
    const targets = pass === "forward" ? plan.forward : plan.wrapped;
    while (index < targets.length) {
      const target = targets[index];
      const text = readSpellingTarget(sheet, target);
      if (text !== null) {
        const issue = findSpellingIssue(
          checkedText(target, text),
          offset,
          context,
          targetRange(target),
        );
        if (issue) {
          return {
            status: "issue",
            finding: { ...issue, target, cursor: { pass, index, offset } },
          };
        }
      }
      index += 1;
      offset = 0;
      if (deadline?.() && index < targets.length) {
        return { status: "paused", cursor: { pass, index, offset } };
      }
    }
    if (pass === "wrapped" || plan.wrapped.length === 0) return { status: "done" };
    if (plan.askBeforeWrapping) {
      return { status: "wrap", cursor: { pass: "wrapped", index: 0, offset: 0 } };
    }
    pass = "wrapped";
    index = 0;
    offset = 0;
  }
}

function sameTarget(left: SpellingTarget, right: SpellingTarget): boolean {
  if (left.kind !== right.kind) return false;
  switch (left.kind) {
    case "cell":
    case "note":
      return left.key === (right as typeof left).key;
    case "shape":
      return left.objectId === (right as typeof left).objectId;
    case "header-footer":
      return (
        left.field === (right as typeof left).field &&
        left.align === (right as typeof left).align
      );
  }
}

export interface SpellingTextEdit {
  target: SpellingTarget;
  /** The target's whole text before and after the edit. */
  before: string;
  after: string;
  /** The keys of the cells the edit rewrites (for the commit's hint). */
  cellKey?: string;
}

function writeCell<S extends SpellingSheetLike>(
  sheet: S,
  target: Extract<SpellingTarget, { kind: "cell" }>,
  before: string,
  after: string,
  start: number,
  end: number,
  replacement: string,
): S | null {
  const cell = sheet.cells[target.key];
  if (!cell) return null;
  const formula = isSpellingFormulaCell(cell);
  if (formula && !target.entry) return null;
  const next: SpellingCellLike = { ...cell };
  delete next.display;
  delete next.chartValueCache;
  if (formula) {
    next.value = after;
    next.formula = after.trim();
    delete next.runs;
  } else {
    next.value = after;
    if (cell.runs && cell.runs.length > 0) {
      const runs = replaceRangeInRuns(cell.runs, start, end, replacement);
      if (runs.map((run) => run.text).join("") === after && runs.length > 0) next.runs = runs;
      else delete next.runs;
    }
  }
  if (cellSource(cell) !== before) return null;
  return { ...sheet, cells: { ...sheet.cells, [target.key]: next } };
}

/**
 * The sheet with `[start, end)` of the target's text replaced, or null when
 * the target no longer reads `expected` there (it changed since it was
 * checked) or cannot take the edit.
 */
export function writeSpellingTarget<S extends SpellingSheetLike>(
  sheet: S,
  target: SpellingTarget,
  start: number,
  end: number,
  replacement: string,
  expected: string,
): { sheet: S; edit: SpellingTextEdit } | null {
  const before = readSpellingTarget(sheet, target);
  if (before === null || before.slice(start, end) !== expected) return null;
  const after = before.slice(0, start) + replacement + before.slice(end);
  const next = replaceSpellingTargetText(sheet, target, before, after, { start, end, replacement });
  return next
    ? {
        sheet: next,
        edit: {
          target,
          before,
          after,
          ...(target.kind === "cell" || target.kind === "note" ? { cellKey: target.key } : {}),
        },
      }
    : null;
}

/**
 * The sheet with the target's whole text set to `after`, when it still reads
 * `before` (Undo Last puts a correction back this way).
 */
export function replaceSpellingTargetText<S extends SpellingSheetLike>(
  sheet: S,
  target: SpellingTarget,
  before: string,
  after: string,
  splice?: { start: number; end: number; replacement: string },
): S | null {
  if (readSpellingTarget(sheet, target) !== before) return null;
  if (before === after) return sheet;
  // The common prefix / suffix is the edited span, for rich-text runs.
  const span = splice ?? editedSpan(before, after);
  switch (target.kind) {
    case "cell":
      return writeCell(sheet, target, before, after, span.start, span.end, span.replacement);
    case "note": {
      const cell = sheet.cells[target.key];
      const note = migrateLegacyComment(cell?.comment);
      if (!cell || note?.kind !== "note") return null;
      const nextNote = setNote(after, note.author, note);
      if (!nextNote) return null;
      return { ...sheet, cells: { ...sheet.cells, [target.key]: { ...cell, comment: nextNote } } };
    }
    case "shape": {
      const objects = sheet.objects ?? [];
      const object = objects.find((candidate) => candidate.id === target.objectId);
      if (!object) return null;
      const patch = shapeTextPatch(object, after);
      if (!patch) return null;
      return {
        ...sheet,
        objects: objects.map((candidate) =>
          candidate === object ? { ...candidate, ...patch } : candidate,
        ),
      };
    }
    case "header-footer": {
      const settings = printSettingsOf(sheet);
      const band = (settings[target.field] as PrintHeaderFooter | undefined) ?? {};
      const runsKey = `${target.align}Runs` as const;
      const runs = headerFooterRunsForEditedText(before, band[runsKey], after);
      const nextBand: PrintHeaderFooter = { ...band, [target.align]: after };
      if (runs) nextBand[runsKey] = runs;
      else delete nextBand[runsKey];
      const nextSettings = normalizePrintSettings({ ...settings, [target.field]: nextBand }) ?? {};
      if ((nextSettings[target.field]?.[target.align] ?? "") !== after.trim()) return null;
      return {
        ...sheet,
        printSettings: Object.keys(nextSettings).length > 0 ? nextSettings : undefined,
      };
    }
  }
}

function editedSpan(
  before: string,
  after: string,
): { start: number; end: number; replacement: string } {
  let prefix = 0;
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix]) {
    prefix += 1;
  }
  let suffix = 0;
  while (
    suffix < before.length - prefix &&
    suffix < after.length - prefix &&
    before[before.length - 1 - suffix] === after[after.length - 1 - suffix]
  ) {
    suffix += 1;
  }
  return {
    start: prefix,
    end: before.length - suffix,
    replacement: after.slice(prefix, after.length - suffix),
  };
}

/**
 * Delete for a repeated word: the second word and the space before it go
 * ("the the cat" -> "the cat"). Excel is documented to leave that space
 * behind; Word removes it, and a double space is never what was meant.
 */
export function repeatedWordDeletion(
  text: string,
  issue: Pick<SpellingIssue, "start" | "end">,
): { start: number; end: number } {
  let start = issue.start;
  while (start > 0 && /[^\S\n\r]/u.test(text[start - 1])) start -= 1;
  return { start, end: issue.end };
}

/**
 * Change All: every occurrence of `word` in the plan's targets replaced (each
 * in the case it was written), as one edit of the sheet. `current` is the
 * finding the dialog stands on; the returned offset is where the walk resumes
 * in that target, just past its corrected word.
 */
export function changeAllInSpellingPlan<S extends SpellingSheetLike>(input: {
  sheet: S;
  plan: SpellingPlan;
  word: string;
  replacement: string;
  options: SpellingOptions;
  current: SpellingFinding;
  /** Targets that may not be written (locked cells); they are left alone. */
  canWrite?: (target: SpellingTarget) => boolean;
}): { sheet: S; edits: SpellingTextEdit[]; resumeOffset: number } {
  let sheet = input.sheet;
  const edits: SpellingTextEdit[] = [];
  let resumeOffset = input.current.start + input.replacement.length;
  for (const target of [...input.plan.forward, ...input.plan.wrapped]) {
    if (input.canWrite && !input.canWrite(target)) continue;
    const before = readSpellingTarget(sheet, target);
    if (before === null) continue;
    const occurrences = spellingOccurrences(
      checkedText(target, before),
      input.word,
      input.options,
      targetRange(target),
    );
    if (occurrences.length === 0) continue;
    const replacementFor = (word: string) =>
      adaptSpellingReplacement(input.current.word, input.replacement, word);
    // Last occurrence first, one splice each, so the earlier offsets hold and
    // each word keeps the formatting of its own run.
    let working: S | null = sheet;
    for (let index = occurrences.length - 1; index >= 0 && working; index -= 1) {
      const occurrence = occurrences[index];
      working =
        writeSpellingTarget(
          working,
          target,
          occurrence.start,
          occurrence.end,
          replacementFor(occurrence.word),
          before.slice(occurrence.start, occurrence.end),
        )?.sheet ?? null;
    }
    if (!working) continue;
    const after = readSpellingTarget(working, target);
    if (after === null || after === before) continue;
    if (sameTarget(target, input.current.target)) {
      let shift = 0;
      for (const occurrence of occurrences) {
        const replacement = replacementFor(occurrence.word);
        if (occurrence.start === input.current.start) {
          resumeOffset = occurrence.start + shift + replacement.length;
          break;
        }
        shift += replacement.length - (occurrence.end - occurrence.start);
      }
    }
    sheet = working;
    edits.push({
      target,
      before,
      after,
      ...(target.kind === "cell" || target.kind === "note" ? { cellKey: target.key } : {}),
    });
  }
  return { sheet, edits, resumeOffset };
}

/** A human-readable place for the dialog: the cell, note, shape or band. */
export function spellingTargetCell(target: SpellingTarget): { row: number; col: number } | null {
  return target.kind === "cell" || target.kind === "note"
    ? { row: target.row, col: target.col }
    : null;
}
