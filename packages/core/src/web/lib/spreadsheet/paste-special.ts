/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Paste-special (Excel 选择性粘贴) — the pure transforms.
 *
 * The window captures a RICH internal clipboard on copy (raw cell text, the
 * cell's formula/text identity, the typed value it showed, its protection and
 * its style bag); these helpers derive the grids the paste modes write. Foreign system-clipboard payloads may
 * contribute a bounded interoperable HTML table through
 * clipboard-html-import.ts; proprietary native workbook objects are not
 * available to browser JavaScript.
 */

import type { AreaLinkSource } from "./hyperlink-areas";
import type { TextRun } from "./cell-rich-text";
import {
  clonePhonetic,
  type SpreadsheetCellPhonetic,
} from "./cell-phonetic";
import { PASTE_MAX_COLS, PASTE_MAX_ROWS } from "./paste-limits";
import {
  cloneCellCommentForPaste,
  type CellComment,
} from "./comments";
import {
  formatSheetNameForFormula,
  isFormulaError,
  isPictureValue,
  readLiteralCellValue,
  shiftFormula,
  transposeFormula,
  type CellValue,
} from "./formula-engine";
import type { Sparkline } from "./sparkline";
import { sparklineCopiedBy } from "./sparkline-buttons";

/**
 * The typed value a copied cell held at copy time: a formula's result, a
 * dynamic array's spilled element, or a literal as the engine reads it. Paste
 * Values writes from this, never from the display string, so General's
 * ten-decimal display cannot round the data and "00123" stays text.
 */
export type ClipboardScalar =
  | { readonly type: "blank" }
  | { readonly type: "number"; readonly value: number }
  | { readonly type: "text"; readonly value: string }
  | { readonly type: "boolean"; readonly value: boolean }
  | { readonly type: "error"; readonly code: string };

export interface RichClipboardCell {
  /** Raw stored text (a formula keeps its leading `=`). */
  raw: string;
  /** Computed display string at copy time (formulas baked). */
  computed: string;
  /**
   * Whether `raw` is a live formula. A quote-prefixed or Text-format literal
   * reading "=A2" is `false`: it pastes as the same text, never rebased. Absent
   * on a foreign (HTML) snapshot, where a leading `=` decides.
   */
  formula?: boolean;
  /** The literal's stored type (the sheet cell's `kind`). */
  kind?: "text" | "number";
  /**
   * The literal is a string the workbook file typed as one (the sheet cell's
   * `stringValue`): text whose characters would otherwise read as a logical,
   * error or number ("FALSE"). It travels with the value, so a copy or a fill
   * of it is the same string (live audit B404).
   */
  stringValue?: true;
  /** Typed value at copy time (see ClipboardScalar). */
  scalar?: ClipboardScalar;
  /**
   * Cell protection (Format Cells > Protection > Locked). A copy of MRAI's own
   * cells always says; absent on a foreign snapshot, whose paste leaves the
   * destination's lock alone.
   */
  locked?: boolean;
  /** Format Cells > Protection > Hidden; true only (absent = not hidden). */
  formulaHidden?: true;
  /** The cell's style bag, if any (opaque here; the sheet owns its shape). */
  style?: Record<string, unknown>;
  /**
   * A cell of a formal table: the format it showed, the table style resolved
   * underneath its own (direct) formats. Excel pastes this effective format
   * as the source formatting wherever the destination is not part of a
   * table; a destination inside a table takes that table's own style, so
   * there only `style` (the direct formats) travels (live audit B157).
   * Absent outside a table and on a foreign snapshot.
   */
  styleOutsideTable?: Record<string, unknown>;
  /** Character-level formatting for a literal rich-text cell. */
  runs?: TextRun[];
  /** The cell's phonetic guide (furigana), which travels with its text. */
  phonetic?: SpreadsheetCellPhonetic;
  /** External/internal hyperlink target plus imported Excel display metadata. */
  hyperlink?: string;
  hyperlinkTooltip?: string;
  hyperlinkDisplay?: string;
  /** A link the page holds inert (`inertHyperlink`, a relative path or
   *  file: address): a full paste takes it along, Paste Values does not. */
  inertHyperlink?: string;
  /** The area the copied link is (hyperlink-areas `areaLinkSourceOf`): a paste
   *  inside that same area deletes it and takes the link beneath instead. */
  hyperlinkFrom?: AreaLinkSource;
  /** Complete Note or threaded-comment state. */
  comment?: CellComment;
  /**
   * The picture placed in the copied cell (Place in Cell). It is the cell's
   * value, so it pastes and fills with the value; `raw` holds its `#VALUE!`
   * fallback. Absent on a foreign snapshot.
   */
  picture?: { readonly src: string; readonly alt?: string };
  /**
   * The picture a FORMULA's result showed at copy time (an IMAGE formula's
   * fetched picture). Only Paste Values uses it: Excel pastes an IMAGE
   * result as values as a picture placed in the cell (MS-XLSX CalcOrigin
   * "Standalone ... copy/paste as value"); every other paste keeps the
   * formula.
   */
  resultPicture?: { readonly src: string; readonly alt?: string };
  /**
   * The sparkline drawn in the copied cell and the sheet it reads its data
   * from. Excel copies a sparkline with its cell (Paste, Fill Down, the fill
   * handle); a paste or fill moves its data range like a relative reference.
   * The model's data range names cells of its own sheet only, so the paste
   * writes it only onto the sheet it was copied from. Absent on a foreign
   * snapshot.
   */
  sparkline?: ClipboardSparkline;
}

export interface ClipboardSparkline {
  readonly sheetId: string;
  readonly value: Sparkline;
}

export type RichClipboard = RichClipboardCell[][];

/**
 * The copied cell's sparkline `rowDelta`/`colDelta` cells away
 * (sparklineCopiedBy), or none when its moved data range leaves the grid.
 */
function shiftClipboardSparkline(
  sparkline: ClipboardSparkline | undefined,
  rowDelta: number,
  colDelta: number,
): ClipboardSparkline | undefined {
  if (!sparkline) return undefined;
  const value = sparklineCopiedBy(sparkline.value, rowDelta, colDelta);
  return value ? { sheetId: sparkline.sheetId, value } : undefined;
}

/**
 * The clipboard's bounds are the paste bounds every other entry uses
 * (paste-limits: 100,000 rows × Excel's 16,384 columns), so a 201-column
 * internal copy pastes as an external one does (B047); the old private
 * 1000 × 200 cap refused it without a word.
 */
const MAX_ROWS = PASTE_MAX_ROWS;
const MAX_COLS = PASTE_MAX_COLS;

/** Clamp a clipboard to the paste bounds (defensive, like the TSV path). */
function clamp(clip: RichClipboard): RichClipboard {
  return clip.slice(0, MAX_ROWS).map((row) => row.slice(0, MAX_COLS));
}

function cloneStyle(
  style: Record<string, unknown> | undefined,
): Record<string, unknown> | undefined {
  return style
    ? (JSON.parse(JSON.stringify(style)) as Record<string, unknown>)
    : undefined;
}

function cloneRuns(runs: TextRun[] | undefined): TextRun[] | undefined {
  return runs?.map((run) => ({
    text: run.text,
    ...(run.style ? { style: { ...run.style } } : {}),
  }));
}

/**
 * Detach one rich clipboard cell from the live workbook model.
 *
 * Copy is a snapshot operation: editing the source after Ctrl/Cmd+C must not
 * mutate what a later paste writes. Strings are immutable; the style bag,
 * rich-text runs and threaded-comment graph are cloned here.
 */
export function detachRichClipboardCell(
  cell: RichClipboardCell,
): RichClipboardCell {
  return {
    raw: cell.raw,
    computed: cell.computed,
    ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
    ...(cell.kind !== undefined ? { kind: cell.kind } : {}),
    ...(cell.stringValue === true ? { stringValue: true as const } : {}),
    ...(cell.scalar !== undefined ? { scalar: { ...cell.scalar } } : {}),
    ...(cell.locked !== undefined ? { locked: cell.locked } : {}),
    ...(cell.formulaHidden === true ? { formulaHidden: true as const } : {}),
    style: cloneStyle(cell.style),
    ...(cell.styleOutsideTable
      ? { styleOutsideTable: cloneStyle(cell.styleOutsideTable) }
      : {}),
    runs: cloneRuns(cell.runs),
    ...(cell.phonetic ? { phonetic: clonePhonetic(cell.phonetic) } : {}),
    hyperlink: cell.hyperlink,
    hyperlinkTooltip: cell.hyperlinkTooltip,
    hyperlinkDisplay: cell.hyperlinkDisplay,
    ...(cell.inertHyperlink ? { inertHyperlink: cell.inertHyperlink } : {}),
    ...(cell.hyperlinkFrom
      ? {
          hyperlinkFrom: {
            ...cell.hyperlinkFrom,
            ...(cell.hyperlinkFrom.range ? { range: { ...cell.hyperlinkFrom.range } } : {}),
            ...(cell.hyperlinkFrom.beneath ? { beneath: { ...cell.hyperlinkFrom.beneath } } : {}),
          },
        }
      : {}),
    comment: cloneCellCommentForPaste(cell.comment),
    ...(cell.picture
      ? {
          picture: {
            src: cell.picture.src,
            ...(cell.picture.alt ? { alt: cell.picture.alt } : {}),
          },
        }
      : {}),
    ...(cell.resultPicture
      ? {
          resultPicture: {
            src: cell.resultPicture.src,
            ...(cell.resultPicture.alt ? { alt: cell.resultPicture.alt } : {}),
          },
        }
      : {}),
    ...(() => {
      const sparkline = shiftClipboardSparkline(cell.sparkline, 0, 0);
      return sparkline ? { sparkline } : {};
    })(),
  };
}

/**
 * Whether a cut source cell carries anything standard paste transfers.
 *
 * Metadata-only cells (for example an empty reviewed cell with a comment or a
 * formatted blank) must move and clear just like non-empty values.
 */
export function richClipboardCellHasTransferableContent(
  cell: RichClipboardCell | undefined,
): boolean {
  if (!cell) return false;
  return Boolean(
    cell.raw !== "" ||
      cell.locked === false ||
      cell.formulaHidden === true ||
      (cell.style && Object.keys(cell.style).length > 0) ||
      (cell.runs && cell.runs.length > 0) ||
      cell.phonetic?.visible === true ||
      cell.hyperlink ||
      cell.hyperlinkTooltip ||
      cell.hyperlinkDisplay ||
      cell.inertHyperlink ||
      cell.comment ||
      cell.sparkline,
  );
}

/**
 * Whether a copied cell is blank for Paste Special > Skip blanks: it holds no
 * value or formula. Formatting, a note or a link do not make it non-blank;
 * Excel's Skip blanks keeps the paste area as it is wherever the copy area is
 * blank, so none of those facets are carried from such a cell either. A cell
 * holding only spaces is text, not blank.
 */
export function richClipboardCellIsBlank(
  cell: RichClipboardCell | undefined,
): boolean {
  return !cell || (cell.raw === "" && !clipboardCellIsFormula(cell));
}

/** Whether a copied cell is a live formula (its identity, not its characters). */
export function clipboardCellIsFormula(cell: RichClipboardCell): boolean {
  return cell.formula ?? cell.raw.trim().startsWith("=");
}

/** The ClipboardScalar of an engine value. */
export function clipboardScalarFromValue(
  value: CellValue | undefined,
): ClipboardScalar {
  if (value === null || value === undefined) return { type: "blank" };
  // A picture placed in a cell (or one IMAGE returned) travels as its
  // cell's `#VALUE!` fallback; the rich snapshot carries the picture itself
  // beside it.
  if (isPictureValue(value)) return { type: "error", code: value.code };
  if (isFormulaError(value)) return { type: "error", code: value.code };
  if (typeof value === "number") {
    return Number.isFinite(value)
      ? { type: "number", value }
      : { type: "error", code: "#NUM!" };
  }
  if (typeof value === "boolean") return { type: "boolean", value };
  return { type: "text", value };
}

/**
 * The typed value a copied cell pastes as a value. A foreign or legacy
 * snapshot without one falls back to reading its display string.
 */
export function clipboardCellScalar(cell: RichClipboardCell): ClipboardScalar {
  if (cell.scalar) return cell.scalar;
  return cell.computed === ""
    ? { type: "blank" }
    : clipboardScalarFromValue(readLiteralCellValue(cell.computed));
}

/**
 * One pasted cell's content, typed: the characters to store and the type they
 * must keep. The sheet's writer stores `raw`; a literal takes `kind`, and
 * `forceText` keeps a text whose characters would otherwise read back as a
 * number, logical, error or formula a text (the cell's quote prefix, the only
 * text-type marker a General cell has).
 */
export interface PastedContent {
  /** Stored characters; "" clears the cell's content. */
  readonly raw: string;
  /** A live formula (raw keeps its leading `=`). */
  readonly formula: boolean;
  /** A literal's stored type. */
  readonly kind: "text" | "number";
  /** The literal is text that its characters alone would not keep text. */
  readonly forceText: boolean;
  /**
   * The picture the copied cell held as its value (Place in Cell): Excel's
   * rich value pastes with values, so Paste Values keeps it beside `raw`,
   * the picture's #VALUE! fallback.
   */
  readonly picture?: { readonly src: string; readonly alt?: string };
}

const BLANK_CONTENT: PastedContent = Object.freeze({
  raw: "",
  formula: false,
  kind: "text",
  forceText: false,
});

/** Whether a text's characters alone would read back as something else. */
function textNeedsTypeGuard(text: string): boolean {
  return (
    text !== "" &&
    (text.trim().startsWith("=") ||
      typeof readLiteralCellValue(text) !== "string")
  );
}

/**
 * The shortest text that reads back as exactly `value` (JavaScript's
 * round-trip form: 1/3 is 0.3333333333333333, not General's 0.3333333333).
 */
function numberText(value: number): string {
  return Object.is(value, -0) ? "0" : String(value);
}

/**
 * The literal a typed value pastes as. An empty text result has no stored
 * form distinct from a blank cell in this model, so it pastes blank.
 */
export function scalarContent(scalar: ClipboardScalar): PastedContent {
  switch (scalar.type) {
    case "blank":
      return BLANK_CONTENT;
    case "number":
      return {
        raw: numberText(scalar.value),
        formula: false,
        kind: "number",
        forceText: false,
      };
    case "boolean":
      return {
        raw: scalar.value ? "TRUE" : "FALSE",
        formula: false,
        kind: "text",
        forceText: false,
      };
    case "error":
      return { raw: scalar.code, formula: false, kind: "text", forceText: false };
    case "text":
      return scalar.value === ""
        ? BLANK_CONTENT
        : {
            raw: scalar.value,
            formula: false,
            kind: "text",
            forceText: textNeedsTypeGuard(scalar.value),
          };
  }
}

/**
 * A copied cell's own content, as Paste Formulas writes it: a formula moved by
 * the given delta, a literal as stored with the type it had.
 */
export function clipboardCellContent(
  cell: RichClipboardCell,
  rowDelta = 0,
  colDelta = 0,
): PastedContent {
  if (clipboardCellIsFormula(cell)) {
    return {
      raw:
        rowDelta !== 0 || colDelta !== 0
          ? shiftFormula(cell.raw.trim(), rowDelta, colDelta)
          : cell.raw,
      formula: true,
      kind: "text",
      forceText: false,
    };
  }
  if (cell.raw === "") return BLANK_CONTENT;
  const scalar =
    cell.scalar ?? clipboardScalarFromValue(readLiteralCellValue(cell.raw));
  return {
    raw: cell.raw,
    formula: false,
    kind: cell.kind ?? (scalar.type === "number" ? "number" : "text"),
    forceText:
      scalar.type === "text" &&
      (cell.style?.quotePrefix === true || textNeedsTypeGuard(cell.raw)),
  };
}

/**
 * Standard copy/paste payload (Excel/WPS plain Paste).
 *
 * A copied formula is rebased by the displacement between the source and
 * destination anchors, so copying B1=`=A1+$A$1` to B3 yields
 * `=A3+$A$1`. A CUT move passes `translateFormulas=false`: moving a formula
 * keeps its references unchanged. Styles are deep-cloned and an absent source
 * style remains explicit `undefined`, allowing the destination style to be
 * cleared instead of accidentally merged.
 */
export function standardPasteGrid(
  clip: RichClipboard,
  rowDelta: number,
  colDelta: number,
  translateFormulas = true,
): RichClipboard {
  return clamp(clip).map((row) =>
    row.map((cell) => {
      const nextRaw =
        translateFormulas && clipboardCellIsFormula(cell)
          ? shiftFormula(cell.raw.trim(), rowDelta, colDelta)
          : cell.raw;
      const detached = detachRichClipboardCell(cell);
      detached.raw = nextRaw;
      // The cell's sparkline reads from the shifted range, as its formula
      // does; a Cut's move keeps it, and the paste re-points it as the
      // moved formulas are.
      if (translateFormulas) {
        detached.sparkline = shiftClipboardSparkline(cell.sparkline, rowDelta, colDelta);
        if (!detached.sparkline) delete detached.sparkline;
      }
      return detached;
    }),
  );
}

/**
 * Values-only paste: each copied cell's typed value at copy time (a formula's
 * full-precision result, a spilled element, a literal), with the type it had.
 */
export function valuesOnlyGrid(clip: RichClipboard): PastedContent[][] {
  return clamp(clip).map((row) =>
    row.map((cell) => {
      const content = scalarContent(clipboardCellScalar(cell));
      if (cell.picture && content.raw === cell.raw) return { ...content, picture: cell.picture };
      // An IMAGE formula's picture pastes as a value: placed in the cell.
      return cell.resultPicture && content.raw === "#VALUE!"
        ? { ...content, picture: cell.resultPicture }
        : content;
    }),
  );
}

/**
 * The copied cells with rows and columns swapped (ragged rows pad with blank
 * cells). The cells are the snapshot's own; callers that write them detach.
 */
export function transposeClipboard(clip: RichClipboard): RichClipboard {
  const bounded = clamp(clip);
  if (bounded.length === 0) return [];
  const rows = bounded.length;
  const cols = Math.max(...bounded.map((row) => row.length));
  const out: RichClipboard = [];
  for (let c = 0; c < cols && c < MAX_ROWS; c += 1) {
    const line: RichClipboardCell[] = [];
    for (let r = 0; r < rows && r < MAX_COLS; r += 1) {
      line.push(bounded[r]?.[c] ?? { raw: "", computed: "" });
    }
    out.push(line);
  }
  return out;
}

/** Transposed values paste: rows become columns (values, formulas baked). */
export function transposedValuesGrid(clip: RichClipboard): PastedContent[][] {
  return valuesOnlyGrid(transposeClipboard(clip));
}

/**
 * Full transpose paste: rows become columns while formulas stay live.
 *
 * Each source cell moves to a different destination coordinate, so its relative
 * references shift by that cell's own source→destination delta (not one shared
 * rectangular delta). All non-value metadata is detached exactly like a normal
 * paste.
 */
export function transposedPasteGrid(
  clip: RichClipboard,
  sourceBaseRow: number,
  sourceBaseCol: number,
  destinationBaseRow: number,
  destinationBaseCol: number,
): RichClipboard {
  const bounded = clamp(clip);
  if (bounded.length === 0) return [];
  const rows = Math.min(MAX_ROWS, bounded.length);
  const cols = Math.min(
    MAX_COLS,
    Math.max(0, ...bounded.slice(0, rows).map((row) => row.length)),
  );
  const out: RichClipboard = [];
  for (let sourceCol = 0; sourceCol < cols; sourceCol += 1) {
    const line: RichClipboardCell[] = [];
    for (let sourceRow = 0; sourceRow < rows; sourceRow += 1) {
      const source = bounded[sourceRow]?.[sourceCol] ?? {
        raw: "",
        computed: "",
      };
      const detached = detachRichClipboardCell(source);
      const rowDelta = destinationBaseRow + sourceCol - (sourceBaseRow + sourceRow);
      const colDelta = destinationBaseCol + sourceRow - (sourceBaseCol + sourceCol);
      if (clipboardCellIsFormula(source)) {
        detached.raw = transposeFormula(
          source.raw.trim(),
          sourceBaseRow + sourceRow,
          sourceBaseCol + sourceCol,
          sourceBaseRow,
          sourceBaseCol,
          destinationBaseRow,
          destinationBaseCol,
        );
      }
      detached.sparkline = shiftClipboardSparkline(source.sparkline, rowDelta, colDelta);
      if (!detached.sparkline) delete detached.sparkline;
      line.push(detached);
    }
    out.push(line);
  }
  return out;
}

/**
 * Formats-only paste: the style bag (deep-copied) for each source cell, to be
 * applied positionally at the paste anchor. `undefined` entries mean the
 * source cell had no style — the target's style is CLEARED to match (Excel
 * semantics: pasting formats replaces, it does not merge).
 */
export function formatsGrid(
  clip: RichClipboard,
  destinationInTable?: (rowOffset: number, colOffset: number) => boolean,
): Array<Array<Record<string, unknown> | undefined>> {
  return clamp(clip).map((row, r) =>
    row.map((cell, c) =>
      cloneStyle(
        richClipboardCellSourceStyle(cell, destinationInTable?.(r, c) ?? false),
      ),
    ),
  );
}

/**
 * The source formatting a copied cell pastes with: a formal table's cell
 * brings the format it showed (its table style under its direct formats)
 * to a destination outside every table, and only its direct formats into a
 * table, whose own style then applies (Excel; live audit B157).
 */
export function richClipboardCellSourceStyle(
  cell: RichClipboardCell | undefined,
  destinationInTable: boolean,
): Record<string, unknown> | undefined {
  if (!cell) return undefined;
  return !destinationInTable && cell.styleOutsideTable
    ? cell.styleOutsideTable
    : cell.style;
}

/**
 * Number-format-only facet for Excel's "Formulas and number formats" / "Values
 * and number formats" Paste Special variants. An absent source format is
 * intentional authority to clear the destination format; every other style
 * field remains owned by the destination cell.
 */
export function numberFormatsGrid(
  clip: RichClipboard,
): Array<Array<string | undefined>> {
  return clamp(clip).map((row) =>
    row.map((cell) => {
      const numberFormat = cell.style?.["numberFormat"];
      if (typeof numberFormat !== "string") return undefined;
      const normalized = numberFormat.trim();
      return normalized && normalized.toLowerCase() !== "general"
        ? numberFormat
        : undefined;
    }),
  );
}

/**
 * Comments-only paste: a positional, detached snapshot of each source cell's
 * complete threaded-comment facet. `undefined` is intentional authority — the
 * corresponding destination comment is cleared while every non-comment facet
 * stays untouched. Cloning again at apply time keeps clipboard and worksheet
 * ownership separate even across repeated Paste Special operations.
 */
export function commentsGrid(
  clip: RichClipboard,
): Array<Array<CellComment | undefined>> {
  return clamp(clip).map((row) =>
    row.map((cell) => cloneCellCommentForPaste(cell.comment)),
  );
}

/**
 * Formulas-only paste (Excel "Formulas"): the RAW stored text of every source
 * cell — a formula keeps its leading `=`, a literal stays literal with its
 * type (a text "=A2" stays that text) — with NO formatting carried (unlike a
 * plain paste, which also copies the style). A Copy supplies the
 * source-to-destination delta so relative references move exactly like a
 * normal copied formula; a Cut supplies zero deltas because moving a formula
 * keeps its references. The "values" sibling instead bakes results.
 */
export function formulasGrid(
  clip: RichClipboard,
  rowDelta = 0,
  colDelta = 0,
): PastedContent[][] {
  return clamp(clip).map((row) =>
    row.map((cell) => clipboardCellContent(cell, rowDelta, colDelta)),
  );
}

// ── Paste Link (Excel "粘贴链接" / Paste Link) ──────────────────────────────────
//
// Paste Link writes, at the destination, a grid of LIVE reference formulas back
// to the copied source cells: copying A1:B2 and Paste-Link at D5 fills D5:E6 with
// =Sheet!A1, =Sheet!B1, =Sheet!A2, =Sheet!B2. The destination then tracks the
// source — editing the source updates the link. The references are absolute to
// the source ORIGIN (top-left of the copied range) offset by each cell's row/col
// position in the clipboard, so the shape is preserved at the new anchor.

/** A1-style column label for a 0-based column index (0→A, 25→Z, 26→AA). */
function columnLabel(index: number): string {
  let n = Math.max(0, Math.trunc(index));
  let label = "";
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

/**
 * The sheet prefix of a cross-sheet reference, written by the formula engine's
 * one quoting rule (formatSheetNameForFormula), so a Paste Link reads back as
 * the same sheet every other formula writer names: bare for a name spelled in
 * any script's letters (`回款!`, `Données!`, `Sheet2!`), quoted with `'`
 * doubled for spaces, punctuation, a leading digit, or a name that is also a
 * cell, an R1C1 reference or a boolean (`'FY2024'!`, `'R2C3'!`). A blank name
 * yields "" so the caller emits a same-sheet `=A1` reference (no prefix).
 */
export function sheetRefPrefix(sheetName: string): string {
  const name = sheetName ?? "";
  if (name.trim() === "") return "";
  return `${formatSheetNameForFormula(name)}!`;
}

/**
 * Authenticate the rectangular rich snapshot Paste Link is allowed to consume.
 * A malformed or over-budget snapshot is rejected whole: silently linking only
 * a bounded prefix would point at a different range than the user copied.
 */
export function pasteLinkSnapshotDimensions(
  clip: RichClipboard | null,
): { rows: number; cols: number } | null {
  if (!clip) return null;
  const rows = clip.length;
  const cols = clip[0]?.length ?? 0;
  if (rows < 1 || rows > MAX_ROWS || cols < 1 || cols > MAX_COLS) return null;
  for (let row = 0; row < rows; row += 1) {
    if (!Array.isArray(clip[row]) || clip[row].length !== cols) return null;
  }
  return { rows, cols };
}

/**
 * Build the Paste Link grid: for a clipboard of `rows`×`cols` cells copied from a
 * source whose top-left is (`originRow`, `originCol`) on sheet `sheetName`, return
 * the grid of reference formulas to write at the destination. Each entry r,c
 * points at source cell (originRow+r, originCol+c). As Excel 16 writes them, a
 * single copied cell links absolutely (`=$A$1`, `=Sheet1!$B$1`) and a copied
 * range links relatively (`=A1`, `=Sheet1!A1`), so a linked block filled on
 * keeps following its source. When `sheetName` is empty the references are
 * same-sheet (no prefix), as Excel writes a link pasted on its own sheet. The result is
 * positional (same shape as the clipboard) so the caller writes it through the
 * normal paste path. Invalid or over-budget input fails closed. PURE.
 */
export function pasteLinkGrid(
  rows: number,
  cols: number,
  originRow: number,
  originCol: number,
  sheetName: string,
  /** The source row of each pasted row, when the Copy left rows out (the
   *  visible rows of a filtered table); otherwise rows run from originRow. */
  sourceRows?: readonly number[],
): string[][] {
  if (
    !Number.isSafeInteger(rows) ||
    !Number.isSafeInteger(cols) ||
    !Number.isSafeInteger(originRow) ||
    !Number.isSafeInteger(originCol) ||
    rows < 1 ||
    rows > MAX_ROWS ||
    cols < 1 ||
    cols > MAX_COLS ||
    originRow < 0 ||
    originCol < 0 ||
    (sourceRows !== undefined &&
      (sourceRows.length !== rows || sourceRows.some((row) => !Number.isSafeInteger(row) || row < 0)))
  ) {
    return [];
  }
  const prefix = sheetRefPrefix(sheetName);
  const out: string[][] = [];
  for (let r = 0; r < rows; r += 1) {
    const line: string[] = [];
    for (let c = 0; c < cols; c += 1) {
      const lock = rows === 1 && cols === 1 ? "$" : "";
      line.push(
        `=${prefix}${lock}${columnLabel(originCol + c)}${lock}${(sourceRows?.[r] ?? originRow + r) + 1}`,
      );
    }
    out.push(line);
  }
  return out;
}

/** The arithmetic operations Excel's Paste Special "Operation" group offers. */
export type PasteArithmeticOp = "add" | "subtract" | "multiply" | "divide";

/**
 * Parse a cell's COMPUTED display string to a number for an arithmetic paste.
 * Blank counts as 0 (Excel treats an empty operand as zero); a non-numeric,
 * non-blank string returns null so the caller can surface "#VALUE!" the way Excel
 * does when you arithmetic-paste over text. Leading/trailing spaces are ignored.
 *
 * Common number-FORMAT decoration is stripped before parsing so a formatted
 * DISPLAY string (the snapshot stores display values, and a formatted destination
 * cell renders the same way) still reads as its underlying number — Excel's
 * arithmetic paste operates on the underlying value, not the rendered text:
 *  - thousands separators ("1,234" -> 1234)
 *  - a leading currency symbol ("$1,234" / "￥99" / "€5")
 *  - a trailing percent ("50%" -> 0.5, matching the cell's stored fraction)
 *  - accounting negatives in parentheses ("(1,234)" -> -1234)
 * A string that is still non-numeric after stripping (real text) returns null.
 */
/** A formula error code as text: the display of an error operand. */
const ERROR_CODE_TEXT =
  /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|CONNECT!|BLOCKED!|UNKNOWN!)$/;

/**
 * The number a text operand IS, or null when it is not one. Only the literal
 * forms Excel reads as numbers convert: a numeric literal (sign, decimals,
 * exponent), a percentage, an accounting negative in parentheses, one leading
 * or trailing currency sign, and thousands separators in proper groups. Any
 * other character makes it text — SKU12 and 12kg are not 12 and never
 * combine as 12 (B045); they are #VALUE!, as =B1+3 is for them.
 */
function operandNumber(text: string): number | null {
  const trimmed = text.trim();
  if (trimmed === "") return 0; // blank operand = 0
  // Fast path: a clean numeric literal (incl. exponent / sign) parses directly.
  const direct = Number(trimmed);
  if (Number.isFinite(direct)) return direct;
  // Accounting negative: a fully-parenthesized value is negative.
  let body = trimmed;
  let sign = 1;
  if (/^\(.*\)$/.test(body)) {
    sign = -1;
    body = body.slice(1, -1).trim();
  }
  if (/^[+-]/.test(body)) {
    if (body.startsWith("-")) sign = -sign;
    body = body.slice(1).trim();
  }
  const isPercent = body.endsWith("%");
  if (isPercent) body = body.slice(0, -1).trim();
  // One currency sign, before or after the digits.
  body = body.replace(/^[$€£¥₩₹￥]\s*/u, "").replace(/\s*[$€£¥₩₹￥]$/u, "");
  // Thousands separators only in proper groups: 1,200 and 1,200,000.50.
  if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(body)) body = body.replace(/,/g, "");
  if (!/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(body)) return null;
  const n = Number(body);
  if (!Number.isFinite(n)) return null;
  return sign * (isPercent ? n / 100 : n);
}

/** The formula operator of each Paste Special operation. */
const ARITHMETIC_OPERATOR: Record<PasteArithmeticOp, string> = {
  add: "+",
  subtract: "-",
  multiply: "*",
  divide: "/",
};

/** A number as a formula literal (1E-07, not JavaScript's 1e-7). */
function formulaNumberLiteral(value: number): string {
  return String(value).toUpperCase();
}

/** Combine a destination and source operand under an arithmetic op. */
function applyArithmetic(
  dest: number,
  src: number,
  op: PasteArithmeticOp,
): number {
  switch (op) {
    case "add":
      return dest + src;
    case "subtract":
      return dest - src;
    case "multiply":
      return dest * src;
    case "divide":
      // Excel surfaces #DIV/0! for a divide-by-zero; the caller maps NaN/Infinity
      // from this to that sentinel (kept out of the number domain here).
      return dest / src;
  }
}

/**
 * Arithmetic Paste Special (Excel "Operation: Add/Subtract/Multiply/Divide"):
 * for each source cell, combine the DESTINATION cell's current computed value
 * with the source's computed value under `op`, producing the new literal string
 * to write at that position. The result grid is positional (same shape as the
 * clipboard), so the caller writes it through the normal paste path.
 *
 * `destValueAt(rowOffset, colOffset)` returns the destination cell's COMPUTED
 * display string (0-based offsets from the paste anchor); a missing/blank cell
 * yields "" → operand 0. Edge cases match Excel:
 *  - both operands blank/0 -> "0" (Excel writes the numeric result, not blank);
 *    UNLESS `skipBlanks` is on AND the SOURCE is blank, in which case the entry is
 *    `null` (the destination is left untouched).
 *  - a non-numeric text operand (source OR dest) -> "#VALUE!".
 *  - divide by zero -> "#DIV/0!".
 * A `null` entry means "do not write this cell" (Skip Blanks); the caller skips
 * those positions so the destination keeps its existing value.
 */
export function arithmeticGrid(
  clip: RichClipboard,
  op: PasteArithmeticOp,
  destValueAt: (rowOffset: number, colOffset: number) => string,
  skipBlanks = false,
  /** The destination cell's formula ("=..."), or null for a constant. */
  destFormulaAt?: (rowOffset: number, colOffset: number) => string | null,
): Array<Array<string | null>> {
  return clamp(clip).map((row, r) =>
    row.map((cell, c) => {
      const scalar = clipboardCellScalar(cell);
      // Skip Blanks: a blank SOURCE cell never overwrites the destination.
      if (skipBlanks && scalar.type === "blank") return null;
      // A FORMULA destination stays a formula: Excel wraps it and applies the
      // operation to the copied number, so =C55*2 with 20 added becomes
      // =(C55*2)+20 and keeps following C55 (Excel for Mac measured, under
      // Paste All and Paste Values alike; audit
      // P-PASTE-ARITHMETIC-FORMULA-RETENTION).
      const destFormula = destFormulaAt?.(r, c) ?? null;
      if (destFormula !== null && destFormula.trim().startsWith("=")) {
        const operand =
          scalar.type === "number"
            ? scalar.value
            : scalar.type === "blank"
              ? 0
              : scalar.type === "error"
                ? null
                : operandNumber(scalar.type === "text" ? scalar.value : cell.computed);
        if (operand !== null && Number.isFinite(operand)) {
          const expression = destFormula.trim().slice(1);
          return `=(${expression})${ARITHMETIC_OPERATOR[op]}${formulaNumberLiteral(operand)}`;
        }
      }
      // An error on EITHER side stays that error, as it does in the formula
      // `=dest+src` (left to right: the destination's first) — a #N/A
      // destination never turns into #VALUE! (B044). The copied value itself,
      // not its display: 1/3 combines as 1/3.
      const destText = destValueAt(r, c);
      if (ERROR_CODE_TEXT.test(destText.trim())) return destText.trim();
      if (scalar.type === "error") return scalar.code;
      const src =
        scalar.type === "number"
          ? scalar.value
          : scalar.type === "blank"
            ? 0
            : operandNumber(
                scalar.type === "text" ? scalar.value : cell.computed,
              );
      const dest = operandNumber(destText);
      if (src === null || dest === null) return "#VALUE!";
      const result = applyArithmetic(dest, src, op);
      if (!Number.isFinite(result)) {
        // divide-by-zero (or overflow) -> Excel's #DIV/0!.
        return op === "divide" ? "#DIV/0!" : "#NUM!";
      }
      return String(result);
    }),
  );
}

/**
 * Apply Skip Blanks to a plain VALUE/FORMULA grid (no arithmetic): a blank source
 * entry becomes `null` so the caller leaves that destination cell untouched,
 * while non-blank entries pass through. Used to layer Skip Blanks onto the
 * values-only / formulas-only paste modes. A grid with skipBlanks off is returned
 * unchanged. Only an empty entry is blank: spaces are text.
 *
 * Blank means the SOURCE cell is empty. With `source` (the copied cells, in
 * the grid's orientation) a formula whose value is "" is not blank: Values
 * paste writes its empty string over the destination, as Excel does.
 */
export function applySkipBlanks<T extends string | PastedContent>(
  grid: T[][],
  skipBlanks: boolean,
  source?: RichClipboard,
): Array<Array<T | null>> {
  if (!skipBlanks) return grid;
  return grid.map((row, r) =>
    row.map((cell, c) => {
      const sourceCell = source?.[r]?.[c];
      const blank =
        sourceCell !== undefined
          ? sourceCell.raw === ""
          : (typeof cell === "string" ? cell : cell.raw) === "";
      return blank ? null : cell;
    }),
  );
}

// ── Paste Special dialog (Excel "选择性粘贴") ──────────────────────────────────
//
// Excel's dialog is one PASTE radio group (which facet of the copied cells
// lands), one OPERATION radio group (arithmetic against the destination), the
// Skip blanks and Transpose checkboxes, and a Paste Link button (Microsoft
// Support, "Copy and paste specific cell contents"). All twelve of Excel's
// Paste choices are offered (XLS-A063). "All using Source theme" pastes All,
// which is what it does inside one workbook; between two workbooks with
// different themes, a theme-bound colour still resolves against the
// destination's theme (the clipboard carries no source theme to re-bind
// it to). "All merging conditional formats" is All whose copied
// conditional formats join the destination's instead of replacing the rules
// that intersect it.

export type PasteSpecialPasteMode =
  | "all"
  | "formulas"
  | "values"
  | "formats"
  | "comments"
  | "validation"
  | "all-source-theme"
  | "all-except-borders"
  | "column-widths"
  | "formulas-number-formats"
  | "values-number-formats"
  | "all-merge-conditional-formats";

/** The dialog's draft: what the user has ticked, before it is applied. */
export interface PasteSpecialDraft {
  readonly paste: PasteSpecialPasteMode;
  readonly operation: PasteArithmeticOp | "none";
  readonly skipBlanks: boolean;
  readonly transpose: boolean;
}

/** Excel opens on All / None with both checkboxes clear. */
export const PASTE_SPECIAL_DEFAULT_DRAFT: PasteSpecialDraft = Object.freeze({
  paste: "all",
  operation: "none",
  skipBlanks: false,
  transpose: false,
});

/** Which companions the engine can honour for the current Paste choice. */
export interface PasteSpecialCompanions {
  /** Operation radios: the copied values combined with the destination. */
  readonly operation: boolean;
  /** Skip blanks: a blank source cell leaves the destination cell alone. */
  readonly skipBlanks: boolean;
  /** Transpose: rows become columns. */
  readonly transpose: boolean;
  /** Paste Link: live references to the copied cells instead of content. */
  readonly pasteLink: boolean;
}

/**
 * What the dialog may offer alongside a Paste choice.
 *
 * Arithmetic bakes values (arithmeticGrid), so an operation is offered under
 * All and Values; Excel also combines under Formulas, which is not expressed
 * here. Skip blanks is offered under All (the full paste leaves every facet
 * of the cell under a blank copied cell alone), Formulas and Values (the
 * nullable writers), with or without an operation. Excel also offers it with
 * the metadata-only choices; they do not take it here. Transpose exists for
 * the full paste (the ribbon's Transpose), for values and for formulas. Excel
 * enables Paste Link only for All / All except borders with no operation,
 * Skip blanks or Transpose, and this mirrors that. The metadata-only choices
 * (Formats, Comments, Validation, Column widths, the two number-format pairs)
 * take none of them.
 */
export function pasteSpecialCompanions(
  rawDraft: PasteSpecialDraft,
): PasteSpecialCompanions {
  // All using Source theme takes All's companions; All merging conditional
  // formats is the full paste with Skip blanks only (XLS-A063).
  if (rawDraft.paste === "all-merge-conditional-formats") {
    return { operation: false, skipBlanks: true, transpose: false, pasteLink: false };
  }
  const draft: PasteSpecialDraft =
    rawDraft.paste === "all-source-theme" ? { ...rawDraft, paste: "all" } : rawDraft;
  const operation = draft.paste === "all" || draft.paste === "values";
  const activeOperation = operation ? draft.operation : "none";
  const skipBlanks =
    draft.paste === "all" ||
    draft.paste === "formulas" ||
    draft.paste === "values";
  const transpose =
    draft.paste === "all" ||
    draft.paste === "values" ||
    draft.paste === "formulas";
  const pasteLink =
    (draft.paste === "all" || draft.paste === "all-except-borders") &&
    activeOperation === "none" &&
    !(skipBlanks && draft.skipBlanks) &&
    !(transpose && draft.transpose);
  return { operation, skipBlanks, transpose, pasteLink };
}

/**
 * Settle a draft after a change: a companion the new Paste choice cannot
 * carry is cleared rather than left ticked behind a disabled control, so what
 * the dialog shows is exactly what Apply will do.
 */
export function settlePasteSpecialDraft(
  draft: PasteSpecialDraft,
): PasteSpecialDraft {
  const companions = pasteSpecialCompanions(draft);
  return {
    paste: draft.paste,
    operation: companions.operation ? draft.operation : "none",
    skipBlanks: companions.skipBlanks && draft.skipBlanks,
    transpose: companions.transpose && draft.transpose,
  };
}

// --- Paste Options button ----------------------------------------------------
// After a paste Excel shows a Paste Options (Ctrl) button at the pasted
// range's bottom-right corner; each choice redoes that paste another way
// (Microsoft Support, "Paste options in Excel" and "Copy and paste specific
// cell contents"). The choices come in three groups, each with the access key
// Excel gives it. A paste of MRAI's own Copy offers what MRAI can do of
// Excel's list; "Keep Source Formatting" is "All using Source theme", which
// inside one workbook is Paste itself, and Keep Source Column Widths and
// Linked Picture have no MRAI command, so those three are left out. A paste
// from another app offers Keep Source Formatting / Match Destination
// Formatting for formatted (HTML) content, and Match Destination Formatting /
// Use Text Import Wizard for plain text.

export type PasteOptionsChoice =
  | "paste"
  | "formulas"
  | "formulas-number-formats"
  | "no-borders"
  | "transpose"
  | "values"
  | "values-number-formats"
  | "values-source-formatting"
  | "formatting"
  | "paste-link"
  | "picture"
  | "keep-source-formatting"
  | "match-destination-formatting"
  | "text-import";

export type PasteOptionsGroup = "paste" | "values" | "other";

export interface PasteOptionsEntry {
  readonly choice: PasteOptionsChoice;
  readonly group: PasteOptionsGroup;
  /** Excel's access key: Ctrl opens the menu, then this letter picks it. */
  readonly accessKey: string;
}

/** The choices after a paste of MRAI's own Copy, in Excel's order. */
export const OWN_COPY_PASTE_OPTIONS: readonly PasteOptionsEntry[] = [
  { choice: "paste", group: "paste", accessKey: "P" },
  { choice: "formulas", group: "paste", accessKey: "F" },
  { choice: "formulas-number-formats", group: "paste", accessKey: "O" },
  { choice: "no-borders", group: "paste", accessKey: "B" },
  { choice: "transpose", group: "paste", accessKey: "T" },
  { choice: "values", group: "values", accessKey: "V" },
  { choice: "values-number-formats", group: "values", accessKey: "A" },
  { choice: "values-source-formatting", group: "values", accessKey: "E" },
  { choice: "formatting", group: "other", accessKey: "R" },
  { choice: "paste-link", group: "other", accessKey: "N" },
  { choice: "picture", group: "other", accessKey: "U" },
];

/** A formatted (HTML) paste from another application. */
export const FOREIGN_HTML_PASTE_OPTIONS: readonly PasteOptionsEntry[] = [
  { choice: "keep-source-formatting", group: "paste", accessKey: "K" },
  { choice: "match-destination-formatting", group: "paste", accessKey: "M" },
];

/** A plain-text paste; the wizard is offered for text that landed in one
 * column, the case Text to Columns can still split. */
export const FOREIGN_TEXT_PASTE_OPTIONS: readonly PasteOptionsEntry[] = [
  { choice: "match-destination-formatting", group: "paste", accessKey: "M" },
  { choice: "text-import", group: "paste", accessKey: "U" },
];

/** The Paste Special dialog choice a Paste Options choice is, when it is one. */
export function pasteOptionsSpecialDraft(
  choice: PasteOptionsChoice,
): PasteSpecialDraft | null {
  switch (choice) {
    case "formulas":
    case "formulas-number-formats":
    case "values":
    case "values-number-formats":
      return { ...PASTE_SPECIAL_DEFAULT_DRAFT, paste: choice };
    default:
      return null;
  }
}

/**
 * Values & Source Formatting: the ordinary rich paste payload with each
 * formula, and each spilled element of a dynamic array, replaced by the typed
 * result it had when copied. Literal cells keep their stored value, so number
 * formats still read them the same way.
 */
export function valuesWithSourceFormatting(clip: RichClipboard): RichClipboard {
  return clip.map((row) =>
    row.map((cell) => {
      const scalar = clipboardCellScalar(cell);
      const derived =
        clipboardCellIsFormula(cell) ||
        (cell.raw === "" && scalar.type !== "blank");
      if (!derived) return cell;
      const content = scalarContent(scalar);
      const baked = detachRichClipboardCell(cell);
      baked.raw = content.raw;
      baked.formula = false;
      baked.kind = content.kind;
      const withTextIdentity = (source: Record<string, unknown> | undefined) => {
        const style = { ...(source ?? {}) };
        if (content.forceText) style.quotePrefix = true;
        else delete style.quotePrefix;
        return Object.keys(style).length > 0 ? style : undefined;
      };
      baked.style = withTextIdentity(baked.style);
      if (baked.styleOutsideTable) {
        baked.styleOutsideTable = withTextIdentity(baked.styleOutsideTable);
      }
      return baked;
    }),
  );
}

/**
 * Lay one grid per tile origin into a single grid covering `range` (the
 * Paste Special fill of a selection that is a whole number of copies, live
 * audit S-104). `tileAt` builds the grid for the tile whose top-left is the
 * given origin, so a formula grid can shift its references per tile. Cells
 * a tile leaves `undefined` stay holes in the result.
 */
export function layPasteTiles<T>(
  origins: ReadonlyArray<{ row: number; col: number }>,
  range: { r1: number; c1: number },
  tileAt: (origin: { row: number; col: number }) => T[][],
): T[][] {
  const out: T[][] = [];
  for (const origin of origins) {
    const tile = tileAt(origin);
    tile.forEach((line, r) => {
      const target = (out[origin.row - range.r1 + r] ??= []);
      line.forEach((cell, c) => {
        target[origin.col - range.c1 + c] = cell;
      });
    });
  }
  return out;
}
