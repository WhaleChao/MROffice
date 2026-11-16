/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The cell writes behind the Home ribbon's buttons and the grid's entry:
 * the stored cell shape, writing a value (typed or set), and patching a
 * cell's format. Moved out of SpreadsheetWindow so the window and the
 * spreadsheet editor commands (lib/editor-commands/spreadsheet) run the same
 * code; behaviour is unchanged.
 */

import { cellLookAt, xlsxDefaultLookEnabled, type DefaultLookSheet } from "./default-look";
import type { KeptFormulaResult } from "./kept-formula-results";
import { stepOfficeFontSize } from "../../components/office-ribbon/officeFontSizeCatalog";
import { dropRecordFontOverridden, recordFontDrawn } from "./cell-record-font";
import {
  CELL_BASE_PADDING_X_PX,
  BORDER_EDGES,
  borderPatchForCell,
  borderPresetPatchesForRange,
  setBorderEdgePatch,
  type BorderMode,
  type SpreadsheetBorderPreset,
  type BorderEdge,
  type BorderLineStyle,
  type CellBorderEdges,
} from "./cell-format";
import {
  formatCellsBorderPatchForCell,
  type CellBorderEdgesPatch,
  type FormatCellsBorderLinesPatch,
} from "./format-cells-border";
import {
  applyRunFormat as applyCellRunFormat,
  normalizeCellRuns,
  runsForWrittenValue,
  textFromRuns,
  type RunStyle as CellRunStyle,
  type TextRun as CellTextRun,
} from "./cell-rich-text";
import { normalizeMerges } from "./merge-cells";
import type { ChartValueCache } from "./chart-value-cache";
import { hasComment, isThreadComment, type CellComment } from "./comments";
import { readCellPicture } from "./cell-picture";
import {
  normalizeCellPhonetic,
  phoneticAfterValueWrite,
  type SpreadsheetCellPhonetic,
} from "./cell-phonetic";
import { parseSpreadsheetEntry, spreadsheetCellIsText } from "./entry-parse";
import type { SpreadsheetFillPattern, SpreadsheetGradientFill } from "./fill-pattern";
import { isExcelNumericText, literalReadsAsNonText } from "./formula-engine";
import { enteredFormulaNumberFormat } from "./formula-result-format";
import {
  autoFitColumnWidth,
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
  type SizeMap,
} from "./grid-sizing";
import { lineStyleAt, materializeLineStyles, patchLineStyles, writeLineBlanksIn, type LineStyleMap } from "./line-styles";
import { styleWithOwnEdgeUnderTable } from "./table-own-borders";
import { nativeFontSizeAfterPatch, nativeFontSizeCss, type NativeFontSize } from "./native-font-size";
import { officePixelsToPoints } from "../../components/office-ribbon/officeTypographyCatalog";
import {
  applyNumberFormat,
  formatGeneralNumber,
  normalizeNumberFormatCode,
  numberFormatFamily,
} from "./number-format";
import { themeColorsAfterPatch, validThemeColors, type SpreadsheetThemeColor, type SpreadsheetThemeColors } from "./theme-colors";
import { type ExcelDateSystemArg } from "./excel-1900-calendar";

export interface SpreadsheetCell {
  value: string;
  // Optional RICH-TEXT runs (Excel inline rich text): per-substring formatting
  // within one cell. When present, `value` stays the concatenation of the runs
  // (textFromRuns) so formulas / sort / CSV read it unchanged; the grid renders
  // the runs as styled spans. Absent for a plain cell (byte-stable). Reuses the
  // presentation run model via lib/spreadsheet/cell-rich-text.
  runs?: CellTextRun[];
  // Optional PHONETIC GUIDE (furigana; Home > Font > Phonetic Guide): the
  // readings over the text, the cell's phonetic settings and Show Phonetic
  // Field. Absent on a plain cell. See lib/spreadsheet/cell-phonetic.
  phonetic?: SpreadsheetCellPhonetic;
  kind?: "text" | "number";
  /** The workbook file typed this value as a string (t="s"/inlineStr) and its
   *  characters would otherwise read as a logical, error or number. */
  stringValue?: true;
  formula?: string;
  // Read-only DISPLAY CACHE (Option A): the resolved on-screen string for this
  // cell (computed formula result and/or applied number format). Written ONLY by
  // the serializer (withDisplayCache) for cells that have a formula or a
  // numberFormat, so plain cells stay byte-stable. The backend CSV/plain-text
  // export prefers this over `value`. Never read by the grid render (which
  // recomputes via resolveCellDisplay); purely an export cache.
  display?: string;
  // Unformatted numeric formula result from the same serialization evaluator.
  // Never canonical cell data; `source` binds it to this exact cell.value.
  chartValueCache?: ChartValueCache;
  // The file's last value of a formula over external data (an OLAP cube,
  // RTD, another workbook), bound to the formula it belongs to: the engine
  // answers it while the cell holds that formula (kept-formula-results).
  keptResult?: KeptFormulaResult;
  hyperlink?: string;
  // Excel hyperlink extras captured on xlsx import: the "ScreenTip" text (shown
  // via the link indicator's title attr) and a display text that differs from
  // the cell value. Both ride along for export fidelity; the editor deletes
  // them whenever the hyperlink itself is changed or cleared.
  hyperlinkTooltip?: string;
  hyperlinkDisplay?: string;
  style?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    // The underline's kind: double, or Excel's single / double accounting
    // rules, which run the whole width of the cell. Absent means the plain
    // single line the boolean already meant (lib/spreadsheet/cell-text-effects).
    underlineStyle?: "double" | "singleAccounting" | "doubleAccounting";
    // Format Cells > Font > Superscript / Subscript on the whole cell
    // (font/vertAlign); at most one. Absent = on the baseline.
    superscript?: boolean;
    subscript?: boolean;
    // Strikethrough (Excel font effect). Composes with underline.
    strikethrough?: boolean;
    // On-screen font size in px (Excel "Font Size"); also carried by named cell
    // styles (Title / Heading). A finite positive number or absent.
    fontSize?: number;
    nativeFontSize?: NativeFontSize | null;
    // Font family / stack (Excel "Font"). A name from the toolbar's whitelist.
    fontFamily?: string;
    // Theme font binding: "heading" (Title/Heading cell styles, the theme's
    // major font) or "body" (its minor font). Direct fontFamily remains
    // stronger. Absence means the workbook's default font — the theme body
    // font in an ordinary workbook, the explicit face (Arial, no scheme) in
    // one whose default font is not theme-bound.
    themeFontRole?: "heading" | "body";
    themeColors?: SpreadsheetThemeColors;
    color?: string;
    backgroundColor?: string;
    // Native SpreadsheetML pattern fill. backgroundColor remains the
    // underlying bgColor used by color filters and HTML fallbacks.
    fillPattern?: SpreadsheetFillPattern;
    // Native SpreadsheetML gradient fill. A header band with a gradient used to
    // arrive white, since the model could only hold one flat colour.
    gradientFill?: SpreadsheetGradientFill;
    numberFormat?: string;
    // The entry was typed after an apostrophe (SpreadsheetML xf quotePrefix):
    // the value is text without the apostrophe, which only editing shows.
    quotePrefix?: boolean;
    // The file's format record this cell came from (lib/spreadsheet/style-source-xf).
    sourceXf?: number;
    border?: boolean;
    // Per-side borders used by the "outline" border mode (the legacy `border`
    // boolean stays the "all"/full-box flag).
    borderTop?: boolean;
    borderRight?: boolean;
    borderBottom?: boolean;
    borderLeft?: boolean;
    // Rich per-edge borders (color + line style), layered OVER the legacy
    // booleans. Present only when a non-default border is set (byte-stable).
    borderEdges?: CellBorderEdges;
    textAlign?: "left" | "center" | "right";
    verticalAlign?: "top" | "middle" | "bottom";
    wrapText?: boolean;
    // Text rotation in degrees (Excel "Orientation": -90..90). Rendered as a CSS
    // transform on the cell's text via textRotationCss. Absent / 0 = horizontal
    // (byte-stable default).
    textRotation?: number;
    // Excel's Vertical Text (XLSX textRotation 255), exclusive with an angle.
    verticalText?: boolean;
    // Left-indent level (Excel "Indent": 0..250). Rendered as leading padding
    // via indentPaddingPx, layered on the grid's base cell padding. Absent / 0 =
    // no indent.
    indent?: number;
    // Shrink-to-fit (Excel): scale the font down so the unwrapped text fits the
    // column. Suppressed when wrapText is on (Excel rule). Absent = off.
    shrinkToFit?: boolean;
    // Center Across Selection (Excel horizontal alignment): the anchor cell's
    // text is centered visually across the blank cells to its right WITHOUT
    // merging them. Absent = off.
    centerAcross?: boolean;
    // Fill (Excel horizontal alignment): the contents are drawn repeated
    // until the cell is full; the value itself is unchanged. Absent = off.
    horizontalFill?: boolean;
    // Excel's Justify / Distributed (Indent) horizontal and Justify /
    // Distributed vertical alignments, beside textAlign / verticalAlign, and
    // the "Justify distributed" box (lib/spreadsheet/cell-format). Absent = off.
    horizontalJustify?: "justify" | "distributed";
    verticalJustify?: "justify" | "distributed";
    justifyLastLine?: boolean;
    // The cell's text direction (alignment/@readingOrder). Absent = Context:
    // the first strong character decides, Excel's default.
    readingOrder?: "ltr" | "rtl";
    // In-cell FORM CONTROLS (Excel / WPS "Insert > Form Controls"). Each flags
    // the cell to render as that control instead of editable text; the control
    // STATE lives in the cell value (see lib/spreadsheet/{checkbox,form-controls}).
    // MUTUALLY EXCLUSIVE — a cell carries at most one (insert sets the chosen
    // flag true and the others false, which updateCellStyle strips). Absent = off.
    checkbox?: boolean;
    radio?: boolean;
    spinner?: boolean;
    slider?: boolean;
    [key: string]: unknown;
  };
  // Excel Notes and threaded Comments are distinct native identities. Legacy
  // `{text, author}` receipts normalize to Notes; only an explicit `thread`
  // receipt becomes a conversation.
  comment?: CellComment;
  // Cell lock flag (Excel/WPS sheet protection). Locked by default; only an
  // explicit `false` unlocks. Meaningful only when the sheet is protected. See
  // lib/spreadsheet/protection. A workflow guard, not a security boundary.
  locked?: boolean;
  // Format Cells > Protection > Hidden: on a protected sheet the formula bar
  // and edit box withhold the contents and a copy carries the value. Stored
  // only when true; travels with the cell's formats like `locked`.
  formulaHidden?: boolean;
  rowspan?: number;
  colspan?: number;
  [key: string]: unknown;
}

export type CellMap = Record<string, SpreadsheetCell>;

/**
 * A blank on a formatted row or column that Clear Formats set back to the
 * default format keeps an empty record of its own, `{ value: "", style: {} }`
 * (SpreadsheetML's cell with the default style index): a cell with a record
 * reads its own format, never its line's, so the cleared bold and fill do not
 * come back when the cell is typed into or the page is reopened (live audit
 * B255). Only such a position keeps one; elsewhere an empty record is pruned.
 */
export const EXPLICIT_DEFAULT_BLANK_CELL: Readonly<SpreadsheetCell> = Object.freeze({
  value: "",
  kind: "text",
  style: {},
});

/** Excel's default font size, in points. */
export const SPREADSHEET_DEFAULT_FONT_PT = 11;
/** A newly chosen point size is known exactly; legacy imported CSS mirrors
 * remain separate and must not quantize a new authoring command. Cell font
 * sizes are stored in CSS pixels. */
export function spreadsheetPtToPx(pt: number): number {
  return pt / .75;
}

/** The worksheet fields the cell writes below read and write. */
export interface CellButtonsSheet {
  cells: CellMap;
  rowCount: number;
  columnCount: number;
  rowStyles?: LineStyleMap;
  columnStyles?: LineStyleMap;
  nativeGrid?: unknown;
  columnWidths?: SizeMap;
  bestFitColumns?: Record<string, true>;
  defaultColumnWidth?: number;
}


export function cellKind(value: string): "text" | "number" {
  // Excel-correct: a JS-only literal (hex 0x1F, 0b/0o, Infinity/NaN) is TEXT, not
  // a number — must agree with coerceLiteral's value-typing so the cell's kind
  // (alignment) matches how a formula reads it.
  return isExcelNumericText(value) ? "number" : "text";
}

export function hasCellMetadata(cell: Partial<SpreadsheetCell> | undefined): boolean {
  return Boolean(
    cell?.formula ||
    cell?.hyperlink ||
    Object.keys(cell?.style || {}).length > 0 ||
    hasComment(cell?.comment) ||
    Boolean(
      isThreadComment(cell?.comment) &&
        cell.comment.draft &&
        cell.comment.draft.trim(),
    ) ||
    cell?.rowspan ||
    cell?.colspan ||
    // An explicit unlock (locked === false) carves out an editable region under
    // sheet protection, so a blank-but-unlocked cell must be kept, not pruned.
    cell?.locked === false ||
    // Format Cells > Protection > Hidden is a format of its own.
    cell?.formulaHidden === true,
  );
}

// Write one cell value INTO an already-cloned cells map (mutates `nextCells`).
// Factored out of updateCellValue so a bulk write (paste / CSV import / fill)
// can clone the cells map ONCE and apply every write into it, instead of
// allocating a fresh sheet+cells object per cell (200k allocations for a
// 1000×200 paste). The per-cell semantics are byte-identical to updateCellValue.
// Only an empty value clears the cell: spaces, tabs and line breaks are text
// (Excel stores "   " as a three-character string; LEN and COUNTA see it).
export function writeCellValueInto(
  nextCells: Record<string, SpreadsheetCell>,
  row: number,
  col: number,
  value: string,
  // The runs an in-place edit of the cell's own text leaves (F2), kept when
  // they still spell `value` (live audit B216); every other write drops them.
  editedRuns?: CellTextRun[],
): void {
  const key = `${row}:${col}`;
  if (value !== "") {
    const previous = nextCells[key] || {};
    const nextCell: SpreadsheetCell = {
      ...previous,
      value,
      kind: cellKind(value),
    };
    // The file's string type belonged to the value it typed; a new value is
    // read afresh (live audit B312).
    delete nextCell.stringValue;
    // A new entry replaces a picture placed in the cell (Excel: the value is
    // the picture, so typing over it replaces it). Writing the cell's own
    // fallback characters back (F2 then Enter) is no new entry.
    if (!(value === (previous as Partial<SpreadsheetCell>).value && readCellPicture(previous))) {
      delete nextCell.picture;
    }
    if (value.trim().startsWith("=")) {
      nextCell.formula = value.trim();
    } else {
      delete nextCell.formula;
    }
    // A freshly typed / pasted / filled value invalidates any inline rich-text
    // `runs` carried over from the cell's PREVIOUS value (their concatenation no
    // longer equals `value`), so keep them only while they still match — else
    // drop them — preserving the invariant textFromRuns(runs) === value. Rich
    // formatting is otherwise only (re)applied via the in-cell dialog, which
    // holds the value constant, never through this value-write path.
    const keptRuns =
      editedRuns && textFromRuns(editedRuns) === value
        ? editedRuns
        : runsForWrittenValue(nextCell.runs, value);
    if (keptRuns) nextCell.runs = keptRuns;
    else {
      delete nextCell.runs;
      // New text takes the cell's own font (cell-record-font).
      if (nextCell.style) nextCell.style = recordFontDrawn(nextCell.style);
    }
    // The phonetic guide describes the text it was read for: a write keeps
    // Show Phonetic Field and the cell's phonetic settings, and the readings
    // of the characters it left untouched (lib/spreadsheet/cell-phonetic).
    const keptPhonetic = phoneticAfterValueWrite(
      (previous as Partial<SpreadsheetCell>).phonetic,
      typeof (previous as Partial<SpreadsheetCell>).value === "string"
        ? ((previous as Partial<SpreadsheetCell>).value as string)
        : "",
      value,
      {
        editedInPlace: true,
        textual: !nextCell.formula && nextCell.kind !== "number",
      },
    );
    if (keptPhonetic) nextCell.phonetic = keptPhonetic;
    else delete nextCell.phonetic;
    nextCells[key] = {
      ...nextCell,
    };
  } else {
    const previous = nextCells[key];
    // Clearing a cell clears its FORMULA too — the user deleted the content.
    // This branch used to spread `previous` wholesale into the retained cell,
    // so a ghost formula survived invisibly: the grid and formula bar showed
    // empty, but the sort/row-ops formula scans (and recalc) still saw it —
    // e.g. column sort kept refusing after every visible formula was cleared.
    const stripped = previous ? { ...previous } : undefined;
    if (stripped) {
      delete stripped.formula;
      delete stripped.picture;
    }
    if (hasCellMetadata(stripped)) {
      // Cleared to empty: stale runs can't match "" either — drop them so the
      // emptied-but-retained cell (kept for its other metadata) stays invariant.
      const cleared: SpreadsheetCell = { ...stripped, value: "", kind: "text" };
      delete cleared.runs;
      // No text, no readings: only Show Phonetic Field stays the cell's.
      const clearedPhonetic = normalizeCellPhonetic(cleared.phonetic, "", false);
      if (clearedPhonetic) cleared.phonetic = clearedPhonetic;
      else delete cleared.phonetic;
      nextCells[key] = cleared;
    } else {
      delete nextCells[key];
    }
  }
}

/**
 * Write one ENTRY (typed, or pasted as plain text) the way Excel enters it:
 * parseSpreadsheetEntry reads "50%", "1,234", "$1,000", "2026/9/21", "10:30",
 * "(123)", "1e3", "true" and "1 1/2" into a value and the format they imply,
 * which a General cell takes. A Text (@) cell keeps what was typed, and a
 * leading apostrophe forces text, remembered as the cell's quote prefix.
 * Every other writer (fill, sort, AI, dialogs) keeps writeCellValueInto.
 */
export function writeSpreadsheetEntryInto(
  nextCells: Record<string, SpreadsheetCell>,
  row: number,
  col: number,
  raw: string,
  locale: string,
  editedRuns?: CellTextRun[],
  date1904: ExcelDateSystemArg = false,
  // Typed into a cell or the formula bar: a numeral reads with the interface
  // language's separators (German 5,5 is 5.5; entry-parse typedNumerals).
  typedNumerals = false,
): void {
  const key = `${row}:${col}`;
  const before = nextCells[key];
  const format = normalizeNumberFormatCode(before?.style?.numberFormat);
  // A cell already formatted as Text takes "=1+1" as the text =1+1 (Excel:
  // the Text format decides before the entry is read; live audit B116).
  const textFormat = format === "@";
  // A percentage cell reads 10 as 10% (Excel's automatic percent entry).
  const percentFormat = format !== null && numberFormatFamily(format) === "percentage";
  const entry =
    raw.trim().startsWith("=") && !textFormat
      ? null
      : parseSpreadsheetEntry(raw, { locale, textFormat, date1904, percentFormat, typedNumerals });
  if (before?.style?.quotePrefix !== undefined) {
    // A new entry replaces the old one's quote prefix.
    const style = { ...before.style };
    delete style.quotePrefix;
    const stripped: SpreadsheetCell = { ...before, style };
    if (Object.keys(style).length === 0) delete stripped.style;
    nextCells[key] = stripped;
  }
  writeCellValueInto(nextCells, row, col, entry?.value ?? raw, editedRuns);
  const written = nextCells[key];
  if (!entry && written?.formula) {
    // A formula entered into a General cell takes the number format its
    // formula implies (lib/spreadsheet/formula-result-format.ts): + and −
    // pass on the first non-General operand's format (=C4+1 of a currency
    // amount is currency), while *, /, ^ and & pass nothing on (=C4*1.1
    // stays General). Applied in the same write, so the entry stays one
    // Undo step.
    const inherited = enteredFormulaNumberFormat(
      written.formula,
      (r, c) => nextCells[`${r}:${c}`]?.style?.numberFormat,
      written.style?.numberFormat,
    );
    if (inherited) {
      nextCells[key] = {
        ...written,
        style: { ...(written.style ?? {}), numberFormat: inherited },
      };
    }
    return;
  }
  if (!entry || !written || !raw.trim()) return;
  const next: SpreadsheetCell = { ...written, kind: entry.kind };
  // '=1+1 is the text "=1+1", not a formula; so is =1+1 under the Text format.
  if (entry.quotePrefix || textFormat) delete next.formula;
  // Excel's AutoCorrect turns a typed Internet or network path into a
  // hyperlink (on by default): https://…, www.…, ftp://…, mailto: and an
  // e-mail address become the link they name. A quote-prefixed entry and a
  // Text-formatted cell are left as text, as in Excel.
  const autoLink = entry.kind === "text" && !entry.quotePrefix && !textFormat ? autoHyperlinkTarget(raw) : null;
  if (autoLink) next.hyperlink = autoLink;
  if (entry.quotePrefix || (entry.numberFormat && format === "")) {
    next.style = {
      ...(written.style ?? {}),
      ...(entry.quotePrefix ? { quotePrefix: true } : {}),
      ...(entry.numberFormat && format === ""
        ? { numberFormat: entry.numberFormat }
        : {}),
    };
  }
  nextCells[key] = next;
}

/**
 * The hyperlink Excel's AutoCorrect gives a typed entry, or null: a URL with
 * a scheme as typed, `www.` with http:// in front, an e-mail address as
 * mailto:. Only the whole entry counts (a URL inside a sentence stays text).
 */
export function autoHyperlinkTarget(text: string): string | null {
  const trimmed = text.trim();
  if (trimmed === "" || /\s/.test(trimmed) || trimmed.length > 2048) return null;
  if (/^(https?|ftp|file):\/\/[^\s]+$/i.test(trimmed)) return trimmed;
  if (/^mailto:[^\s@]+@[^\s@]+$/i.test(trimmed)) return trimmed;
  if (/^www\.[^\s.]+(\.[^\s.]+)+$/i.test(trimmed)) return `http://${trimmed}`;
  if (/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(trimmed)) return `mailto:${trimmed}`;
  return null;
}

/** A plain enumerable copy of a cell map. A `for...in` copy of a
 * dictionary-mode object of tens of thousands of keys runs in about half the
 * time of object spread in V8 and yields the same own string keys; cell maps
 * are plain objects with no inherited enumerable properties. */
export function copyCellMap(cells: CellMap): Record<string, SpreadsheetCell> {
  const next: Record<string, SpreadsheetCell> = {};
  for (const key in cells) next[key] = cells[key];
  return next;
}

/**
 * Excel widens a column the user has never sized when a number or a date typed
 * into it does not fit, so ##### is something you meet on a column you set
 * yourself, not on one you just typed into. The file format calls such a column
 * best-fit — "has never been manually set by the user" — and says the width
 * "should automatically resize to display the number", growing only, never
 * shrinking (ECMA-376, `col/@bestFit`; Microsoft Support, "Change the column
 * width and row height" for the 8.43-character default).
 *
 * Only numbers and dates widen: text is drawn over its empty neighbours
 * instead. The fit uses the same estimate AutoFit itself writes, and a General
 * cell is measured by General's own eleven-character limit — past that Excel
 * shows scientific notation rather than a wider column.
 *
 * A column whose width the user set (dragged, or typed into Column Width) is
 * left alone, and so is an imported one: that ##### is Excel's answer too. A
 * width that came from fitting — this widening itself, AutoFit Column Width, a
 * double-clicked divider — keeps the column best-fit (`bestFitColumns`), so a
 * later, longer number widens it again, as Excel keeps doing until the user
 * sizes the column.
 */
export function widenBestFitColumnForEntry<S extends CellButtonsSheet>(
  sheet: S,
  row: number,
  col: number,
): S {
  // Imported grid metrics are the file's own column widths; entry never
  // rewrites them.
  if (sheet.nativeGrid) return sheet;
  const storedWidth = sheet.columnWidths?.[col];
  if (storedWidth !== undefined && sheet.bestFitColumns?.[col] !== true) {
    return sheet;
  }
  const cell = sheet.cells[`${row}:${col}`];
  if (!cell || cell.formula || cell.kind !== "number") return sheet;
  const number = Number(cell.value);
  if (!Number.isFinite(number)) return sheet;
  const code = normalizeNumberFormatCode(cell.style?.numberFormat);
  const drawn = code
    ? applyNumberFormat(number, code)
    : formatGeneralNumber(number).slice(0, number < 0 ? 12 : 11);
  const defaultWidth = sheet.defaultColumnWidth ?? DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX;
  // Whether it fits is Excel's question, in Excel's metric: a digit of the
  // default 11 pt font is 7 px and a column keeps 5 px of margin, so the
  // 8.43-character default (64 px) holds 8 digits. The editor's own wider
  // estimate widened that column for a six-digit number Excel shows as it is
  // (xlsx-rounds R16-X28 E1: "429980" wrote a <col> the file never had).
  const fontPt = cell.style?.fontSize ? officePixelsToPoints(cell.style.fontSize) : SPREADSHEET_DEFAULT_FONT_PT;
  const digitPx = Math.max(1, Math.round((7 * fontPt) / SPREADSHEET_DEFAULT_FONT_PT));
  // General fits a decimal to its column by rounding the decimals (the grid
  // draws it so, general-number-fit), so only its whole part has to fit: Excel
  // shows 198917.89 as 198917.9 in the default column, and 3.14159265358979 as
  // 3.141593 (xlsx-rounds R247-X11 E1: H2's 198917.89 wrote a <col> the file
  // never had).
  const fitted =
    !code && /^-?\d+\.\d+$/.test(drawn) ? `${number < 0 ? "-" : ""}${Math.round(Math.abs(number))}` : drawn;
  if (fitted.length * digitPx + 5 <= (storedWidth ?? defaultWidth)) return sheet;
  // The cell's text box is the column less its padding on both sides.
  const width = autoFitColumnWidth(drawn, defaultWidth, { paddingPx: 2 * CELL_BASE_PADDING_X_PX });
  // Growing only, never shrinking.
  if (width <= (storedWidth ?? defaultWidth)) return sheet;
  return withBestFitColumns(
    { ...sheet, columnWidths: { ...sheet.columnWidths, [col]: width } },
    [col],
    true,
  );
}

/** Mark columns best-fit (a fitted width) or not (a width the user set). */
export function withBestFitColumns<S extends CellButtonsSheet>(
  sheet: S,
  columns: readonly number[],
  bestFit: boolean,
): S {
  const current = sheet.bestFitColumns;
  if (columns.every((col) => (current?.[col] === true) === bestFit)) return sheet;
  const marks: Record<string, true> = { ...current };
  for (const col of columns) {
    if (bestFit) marks[String(col)] = true;
    else delete marks[String(col)];
  }
  const next: S = { ...sheet };
  if (Object.keys(marks).length > 0) next.bestFitColumns = marks;
  else delete next.bestFitColumns;
  return next;
}

export function updateCellValue<S extends CellButtonsSheet>(
  sheet: S,
  row: number,
  col: number,
  value: string,
  // Set for a user's entry: read it as Excel does, in this locale.
  entryLocale?: string,
  editedRuns?: CellTextRun[],
  // The workbook's date system, for a typed date.
  entryDate1904: ExcelDateSystemArg = false,
  // The entry's digits were typed: read with the interface language's
  // separators (writeSpreadsheetEntryInto).
  entryTypedNumerals = false,
): S {
  const nextCells = copyCellMap(sheet.cells);
  if (entryLocale === undefined) writeCellValueInto(nextCells, row, col, value, editedRuns);
  else writeSpreadsheetEntryInto(nextCells, row, col, value, entryLocale, editedRuns, entryDate1904, entryTypedNumerals);
  // A lazily revealed blank coordinate is session-only until the user writes
  // there. The first real edit extends persisted bounds in this SAME sheet
  // transaction, so existing Undo/autosave/native-writer authority stays the
  // only content seam. Clearing an untouched far cell does not grow anything.
  const hasAddressedCell = Object.prototype.hasOwnProperty.call(
    nextCells,
    `${row}:${col}`,
  );
  // A cell created on a formatted row or column takes the line's format, as
  // Excel keeps a row's format on a value entered anywhere in it (B029); the
  // entry's own number format wins over the line's.
  if (hasAddressedCell && sheet.cells[`${row}:${col}`] === undefined) {
    const inherited = lineStyleAt(sheet.rowStyles, sheet.columnStyles, row, col);
    if (inherited) {
      const written = nextCells[`${row}:${col}`];
      nextCells[`${row}:${col}`] = withLineFormat(written, inherited);
    }
  }
  let next: S = {
    ...sheet,
    cells: nextCells,
    rowCount: hasAddressedCell
      ? Math.max(sheet.rowCount, Math.min(MAX_GRID_ROWS, row + 1))
      : sheet.rowCount,
    columnCount: hasAddressedCell
      ? Math.max(sheet.columnCount, Math.min(MAX_GRID_COLUMNS, col + 1))
      : sheet.columnCount,
  };
  // The grid grew: the blank positions of the new band read their line
  // formats like the rest of the grid.
  if (next.rowCount > sheet.rowCount || next.columnCount > sheet.columnCount) {
    next = materializeLineStyles(next, { rowCount: sheet.rowCount, columnCount: sheet.columnCount });
  }
  return entryLocale === undefined
    ? next
    : widenBestFitColumnForEntry(next, row, col);
}

/**
 * Many values written into a sheet in order, as `updateCellValue` writes
 * each one (no entry reading: the text is stored as given) -- one copy of
 * the cell map for all of them instead of one per value. A column of
 * 20,000 formulas a Copilot writes into a 60,000-row sheet copied its
 * 120,000 cells once per formula (minutes); the result is the same sheet.
 */
export function updateCellValues<S extends CellButtonsSheet>(
  sheet: S,
  writes: readonly { row: number; col: number; value: string }[],
): S {
  const nextCells = copyCellMap(sheet.cells);
  let rowCount = sheet.rowCount;
  let columnCount = sheet.columnCount;
  for (const { row, col, value } of writes) {
    const key = `${row}:${col}`;
    const had = nextCells[key] !== undefined;
    writeCellValueInto(nextCells, row, col, value);
    if (!Object.prototype.hasOwnProperty.call(nextCells, key)) continue;
    // A cell created on a formatted row or column takes the line's format,
    // as in updateCellValue.
    if (!had) {
      const inherited = lineStyleAt(sheet.rowStyles, sheet.columnStyles, row, col);
      if (inherited) {
        nextCells[key] = withLineFormat(nextCells[key], inherited);
      }
    }
    rowCount = Math.max(rowCount, Math.min(MAX_GRID_ROWS, row + 1));
    columnCount = Math.max(columnCount, Math.min(MAX_GRID_COLUMNS, col + 1));
  }
  let next: S = { ...sheet, cells: nextCells, rowCount, columnCount };
  if (next.rowCount > sheet.rowCount || next.columnCount > sheet.columnCount) {
    next = materializeLineStyles(next, { rowCount: sheet.rowCount, columnCount: sheet.columnCount });
  }
  return next;
}

/**
 * A cell's runs after a whole-cell font command: Excel applies Bold, Size,
 * Color and the rest to every character of the cell, so the same patch is
 * laid over each run (a run's own 24 px stops overriding a 20 pt cell; the
 * properties the command does not touch stay). Undefined when the runs no
 * longer carry any formatting of their own.
 */
export function cellRunsAfterFontPatch(
  runs: CellTextRun[] | undefined,
  value: string,
  patch: Record<string, unknown>,
): CellTextRun[] | undefined | "unchanged" {
  const current = runs && normalizeCellRuns(runs, value);
  if (!current) return "unchanged";
  const runPatch = Object.fromEntries(
    Object.entries(patch)
      .filter(([field]) =>
        [
          "bold",
          "italic",
          "underline",
          "strikethrough",
          "fontFamily",
          "fontSize",
          "color",
          // Format Cells > Font > Effects on the whole cell sets every run's
          // position too, the other one off (xlsx-rounds R406-X11 WALK2: δ¹⁸O
          // made subscript kept "18" superscript and the other runs plain).
          "superscript",
          "subscript",
        ].includes(field),
      )
      .map(([field, fieldValue]) => [
        field === "strikethrough" ? "strike" : field,
        fieldValue,
      ]),
  ) as Partial<CellRunStyle>;
  if (Object.keys(runPatch).length === 0) return "unchanged";
  return normalizeCellRuns(
    applyCellRunFormat(current, 0, value.length, runPatch),
    value,
  );
}

export function updateCellStyle<S extends CellButtonsSheet>(
  sheet: S,
  row: number,
  col: number,
  patch: NonNullable<SpreadsheetCell["style"]>,
): S {
  const key = `${row}:${col}`;
  const nextCells = copyCellMap(sheet.cells);
  const previous = nextCells[key] || { value: "", kind: "text" as const };
  const nextStyle = {
    ...(previous.style || {}),
    ...patch,
    themeColors: themeColorsAfterPatch(previous.style || {}, patch),
    nativeFontSize: nativeFontSizeAfterPatch(previous.style?.nativeFontSize, patch),
  };
  dropRecordFontOverridden(nextStyle as Record<string, unknown>, patch, previous.style as Record<string, unknown> | undefined);
  for (const [styleKey, value] of Object.entries(nextStyle)) {
    if (
      value === false ||
      value === "" ||
      value === undefined ||
      (value === null && styleKey !== "nativeFontSize")
    ) {
      delete nextStyle[styleKey as keyof typeof nextStyle];
    }
  }
  const nextCell: SpreadsheetCell = {
    ...previous,
    value: String(previous.value || ""),
    kind: previous.kind || cellKind(String(previous.value || "")),
  };
  if (Object.keys(nextStyle).length > 0) {
    nextCell.style = nextStyle;
  } else {
    delete nextCell.style;
  }
  // A number format changes how a value shows, never the value a formula
  // reads (Microsoft Support, "Available number formats in Excel"): text
  // entered under the Text format stays text after the cell is formatted as
  // Number or General, until it is entered again (live audit S-166). Only
  // characters that would read as something else need the marker, as the
  // workbook reader gives it: "2-3" is text either way, and a marker on it
  // was a change the export's reading did not keep (xlsx-rounds R376-X12).
  if (
    Object.prototype.hasOwnProperty.call(patch, "numberFormat") &&
    nextCell.value !== "" &&
    (literalReadsAsNonText(nextCell.value) || nextCell.value.trim().startsWith("=")) &&
    spreadsheetCellIsText(previous) &&
    !spreadsheetCellIsText(nextCell)
  ) {
    nextCell.stringValue = true;
  }
  // A whole-cell font command reaches the characters as well (live audit
  // B215): the runs' own size or weight no longer outlast it.
  const nextRuns = cellRunsAfterFontPatch(previous.runs, nextCell.value, patch);
  if (nextRuns !== "unchanged") {
    if (nextRuns) nextCell.runs = nextRuns;
    else delete nextCell.runs;
  }
  if (nextCell.value || hasCellMetadata(nextCell)) {
    nextCells[key] = nextCell;
  } else {
    delete nextCells[key];
  }
  return { ...sheet, cells: nextCells };
}

/** A patch laid over xf 0's look: the look's keys first, the patch's over
 *  them, theme bindings of both (the patch's win). */
function withDefaultLookUnder(
  look: NonNullable<SpreadsheetCell["style"]>,
  patch: NonNullable<SpreadsheetCell["style"]>,
): NonNullable<SpreadsheetCell["style"]> {
  const bindings = { ...(look.themeColors ?? {}), ...(patch.themeColors ?? {}) };
  return { ...look, ...patch, ...(Object.keys(bindings).length > 0 ? { themeColors: bindings } : {}) };
}

/** A style patch, or one worked out from each cell's own style: Increase /
 * Decrease Indent steps every cell from its own level, as Excel does. */
export type CellStylePatchInput =
  | NonNullable<SpreadsheetCell["style"]>
  | ((own: SpreadsheetCell["style"] | undefined) => NonNullable<SpreadsheetCell["style"]>);

/**
 * A sheet with Home's / Format Cells' style patch applied to a selection:
 * each cell of it, the Border tab's lines, and, for a whole-row or
 * whole-column selection, the line's own format. Protection and the Wrap
 * Text row growth are the caller's.
 */
export function cellStylePatchedSheet<S extends CellButtonsSheet>(
  sheet: S,
  selRect: { r1: number; c1: number; r2: number; c2: number },
  patchInput: CellStylePatchInput,
  options: {
    /** Format Cells > Border: the lines of the selection as one block
     * (XLS-A060). Only the lines the author drew or removed are present. */
    borderLinesPatch?: FormatCellsBorderLinesPatch;
  } | undefined,
  gridRowCount: number,
  gridColumnCount: number,
): S | null {
  // A format is patched over what the position wears, its line's format too.
  writeLineBlanksIn(sheet as unknown as Parameters<typeof writeLineBlanksIn>[0], selRect);
  const defaultLookStyle = xlsxDefaultLookEnabled() ? (sheet as DefaultLookSheet).defaultStyle : undefined;
  const defaultLook = defaultLookStyle && Object.keys(defaultLookStyle).length > 0
    ? defaultLookStyle as NonNullable<SpreadsheetCell["style"]>
    : undefined;
  let nextSheet = sheet;
  // One cell's style patch with its share of the Border tab's lines applied.
  const withEdgesPatch = (
    r: number,
    c: number,
    edgesPatch: CellBorderEdgesPatch,
    base: NonNullable<SpreadsheetCell["style"]>,
  ): NonNullable<SpreadsheetCell["style"]> => {
    const ownEdges = nextSheet.cells[`${r}:${c}`]?.style?.borderEdges;
    const bindings = {...validThemeColors(nextSheet.cells[`${r}:${c}`]?.style ?? {}), ...(base.themeColors ?? {})};
    const nextEdges: CellBorderEdges = { ...(ownEdges ?? {}) };
    for (const edge of BORDER_EDGES) {
      const edgePatch = edgesPatch[edge];
      if (edgePatch === undefined) continue;
      if (edgePatch.style === null) {
        delete nextEdges[edge];
        delete bindings[`borderEdges.${edge}.color`];
        continue;
      }
      let nextEdge = nextEdges[edge] ? { ...nextEdges[edge] } : undefined;
      if (edgePatch.style !== undefined) {
        nextEdge = {
          ...(nextEdge ?? {}),
          style: edgePatch.style,
        };
      }
      if (!nextEdge) continue;
      if (Object.prototype.hasOwnProperty.call(edgePatch, "color")) {
        delete bindings[`borderEdges.${edge}.color`];
        if (edgePatch.themeColor) bindings[`borderEdges.${edge}.color`] = edgePatch.themeColor;
        if (edgePatch.color === null || edgePatch.color === undefined) {
          delete nextEdge.color;
        } else {
          nextEdge.color = edgePatch.color;
        }
      }
      // A diagonal names its directions; neither flag is the down line, so
      // a down-only diagonal says so explicitly.
      for (const flag of ["up", "down"] as const) {
        if (edgePatch[flag] === undefined) continue;
        if (edgePatch[flag]) nextEdge[flag] = true;
        else delete nextEdge[flag];
      }
      nextEdges[edge] = nextEdge;
    }
    return {
      ...base,
      themeColors: bindings,
      borderEdges:
        Object.keys(nextEdges).length > 0 ? nextEdges : undefined,
    };
  };
  const lines = options?.borderLinesPatch;
  const block = {
    r1: Math.min(selRect.r1, selRect.r2),
    r2: Math.max(selRect.r1, selRect.r2),
    c1: Math.min(selRect.c1, selRect.c2),
    c2: Math.max(selRect.c1, selRect.c2),
  };
  for (let r = selRect.r1; r <= selRect.r2; r += 1) {
    for (let c = selRect.c1; c <= selRect.c2; c += 1) {
      const cellEdges = lines
        ? formatCellsBorderPatchForCell(
            r,
            c,
            block,
            lines,
            nextSheet.cells[`${r}:${c}`]?.style?.borderEdges as CellBorderEdges | undefined,
          )
        : undefined;
      // A position wearing xf 0's look (xlsxDefaultLook) is formatted over
      // that look, as Excel's new xf is xf 0 plus the change: bold on a cell
      // of a vertically centred workbook is centred and bold.
      const overDefault = defaultLook && cellLookAt(nextSheet, r, c).source === "default" ? defaultLook : undefined;
      const patch =
        typeof patchInput === "function"
          ? patchInput(overDefault ?? nextSheet.cells[`${r}:${c}`]?.style)
          : patchInput;
      const cellPatch =
        cellEdges && Object.keys(cellEdges).length > 0
          ? withEdgesPatch(r, c, cellEdges, patch)
          : patch;
      nextSheet = updateCellStyle(nextSheet, r, c, overDefault ? withDefaultLookUnder(overDefault, cellPatch) : cellPatch);
      for (const [side, edgePatch] of Object.entries(cellEdges ?? {})) {
        if (!edgePatch || (side !== "top" && side !== "right" && side !== "bottom" && side !== "left")) continue;
        const pen = edgePatch as { style?: string | null; color?: string | null; themeColor?: object };
        nextSheet = withOwnEdgeUnderTable(nextSheet, r, c, side, pen.style
          ? { style: pen.style, color: pen.color ?? null, theme: pen.themeColor ?? null }
          : null);
      }
    }
  }
  // A whole-row / whole-column selection (the header click) is a LINE
  // format, not the cells of the moment: it is stored on the line so the
  // columns (rows) the sheet grows into inherit it (B029). The line's
  // style is merged by the same rules as a cell's. A whole-sheet selection
  // is stored on the columns, which every later row reads.
  {
    const wholeRows = selRect.c1 === 0 && selRect.c2 >= gridColumnCount - 1;
    const wholeCols = selRect.r1 === 0 && selRect.r2 >= gridRowCount - 1;
    const mergeLineStyle = (lineOwn: LineStyleMap[string] | undefined) => {
      // A line with no format of its own is xf 0: its new format is xf 0's look plus the change.
      const existing = lineOwn ?? (defaultLook as LineStyleMap[string] | undefined);
      return updateCellStyle(
        { ...nextSheet, cells: existing ? { "0:0": { value: "", kind: "text", style: existing } } : {} },
        0,
        0,
        typeof patchInput === "function"
          ? patchInput(existing as SpreadsheetCell["style"] | undefined)
          : patchInput,
      ).cells["0:0"]?.style as LineStyleMap[string] | undefined;
    };
    if (wholeCols && (!wholeRows || nextSheet.columnCount <= nextSheet.rowCount)) {
      nextSheet = patchLineStyles(nextSheet, "col", selRect.c1, selRect.c2, mergeLineStyle);
    } else if (wholeRows) {
      nextSheet = patchLineStyles(nextSheet, "row", selRect.r1, selRect.r2, mergeLineStyle);
    }
  }
  return nextSheet;
}

/**
 * A sheet whose cell at (row, col) wears xf 0's look (xlsxDefaultLook) as
 * that look stated on the cell, so a border drawn on it is drawn over the
 * look, as Excel's new xf is xf 0 plus the borders: a bordered cell of a
 * vertically centred workbook stays centred (xlsx-rounds R308-X25, R309-X21,
 * R309-X30, R310-X08 E24). Any other sheet comes back as it was.
 */
function withDefaultLookStated<S extends CellButtonsSheet>(sheet: S, row: number, col: number): S {
  const look = xlsxDefaultLookEnabled() ? (sheet as DefaultLookSheet).defaultStyle : undefined;
  if (!look || Object.keys(look).length === 0) return sheet;
  if (cellLookAt(sheet as DefaultLookSheet, row, col).source !== "default") return sheet;
  return updateCellStyle(sheet, row, col, withDefaultLookUnder(look as NonNullable<SpreadsheetCell["style"]>, {}));
}

/** A merged area draws its right and bottom lines off its anchor's
 *  `borderEdges` (the importer states the area's last cells' lines there), so
 *  a preset's lines on an area inside the rectangle are also the anchor's: its
 *  bottom is what the preset gives the area's last row, its right what it gives
 *  the last column, and an edge the preset leaves there untouched stays as it
 *  was (xlsx-rounds R420-X05: Thick Bottom over AK5:AK7 with AK6:AK7 merged
 *  went to AK7 only and the merged cell kept drawing its medium line). */
function mergeAnchorsWithTheirFarLines(
  patches: ReturnType<typeof borderPresetPatchesForRange>,
  merges: unknown,
  rect: { r1: number; c1: number; r2: number; c2: number },
): ReturnType<typeof borderPresetPatchesForRange> {
  const spans = normalizeMerges(merges);
  if (spans.length === 0) return patches;
  const top = Math.min(rect.r1, rect.r2);
  const bottom = Math.max(rect.r1, rect.r2);
  const left = Math.min(rect.c1, rect.c2);
  const right = Math.max(rect.c1, rect.c2);
  const byCell = new Map(patches.map((patch) => [`${patch.row}:${patch.col}`, patch]));
  const out = [...patches];
  for (const span of spans) {
    if (span.rowSpan <= 1 && span.colSpan <= 1) continue;
    const lastRow = span.row + span.rowSpan - 1;
    const lastCol = span.col + span.colSpan - 1;
    if (span.row < top || lastRow > bottom || span.col < left || lastCol > right) continue;
    const anchorKey = `${span.row}:${span.col}`;
    const edges = { ...(byCell.get(anchorKey)?.edges ?? {}) };
    for (const [side, key, size] of [
      ["bottom", `${lastRow}:${span.col}`, span.rowSpan],
      ["right", `${span.row}:${lastCol}`, span.colSpan],
    ] as const) {
      if (size <= 1) continue;
      const far = byCell.get(key)?.edges;
      if (far && Object.prototype.hasOwnProperty.call(far, side)) edges[side] = far[side];
      else delete edges[side];
    }
    const index = out.findIndex((patch) => patch.row === span.row && patch.col === span.col);
    if (Object.keys(edges).length === 0) {
      if (index >= 0) out.splice(index, 1);
      continue;
    }
    const patch = { row: span.row, col: span.col, edges };
    if (index >= 0) out[index] = patch;
    else out.push(patch);
  }
  return out;
}

/** Home > Borders, one gallery preset over a rectangle: each cell's edges
 *  from the shared preset walk, in the chosen line style and colour (the
 *  thick/double presets keep their own line). Protection is the caller's. */
export function borderPresetPatchedSheet<S extends CellButtonsSheet>(
  sheet: S,
  preset: SpreadsheetBorderPreset,
  rect: { r1: number; c1: number; r2: number; c2: number },
  selectedLineStyle: BorderLineStyle,
  selectedColor: string | null,
  selectedThemeColor: SpreadsheetThemeColor | undefined,
): S {
  let nextSheet = sheet;
  for (const { row, col, edges } of mergeAnchorsWithTheirFarLines(
    borderPresetPatchesForRange(preset, rect),
    (sheet as { merges?: unknown }).merges,
    rect,
  )) {
    nextSheet = withDefaultLookStated(nextSheet, row, col);
    let own = nextSheet.cells[`${row}:${col}`]?.style?.borderEdges as
      CellBorderEdges | undefined;
    for (const [edge, line] of Object.entries(edges) as Array<
      [BorderEdge, BorderLineStyle | null]
    >) {
      const keepsPresetLineStyle = [
        "thickBox",
        "bottomDouble",
        "thickBottom",
        "topAndThickBottom",
        "topAndDoubleBottom",
      ].includes(preset);
      const patch = setBorderEdgePatch(
        own,
        edge,
        line === null
          ? null
          : {
              style: keepsPresetLineStyle ? line : selectedLineStyle,
              ...(selectedColor ? { color: selectedColor } : {}),
            },
      );
      own = patch.borderEdges;
      const bindings = validThemeColors(nextSheet.cells[`${row}:${col}`]?.style ?? {});
      delete bindings[`borderEdges.${edge}.color`];
      if (line !== null && selectedThemeColor) bindings[`borderEdges.${edge}.color`] = selectedThemeColor;
      nextSheet = updateCellStyle(nextSheet, row, col, {...patch, themeColors:bindings});
      nextSheet = withOwnEdgeUnderTable(nextSheet, row, col, edge, line === null ? null : {
        style: keepsPresetLineStyle ? line : selectedLineStyle,
        color: selectedColor,
        theme: line !== null && selectedThemeColor ? selectedThemeColor : null,
      });
    }
  }
  return nextSheet;
}



function withOwnEdgeUnderTable<S extends CellButtonsSheet>(
  sheet: S,
  row: number,
  col: number,
  side: string,
  edge: Parameters<typeof styleWithOwnEdgeUnderTable>[2],
): S {
  const key = `${row}:${col}`;
  const cell = sheet.cells[key];
  if (!cell?.style) return sheet;
  const style = styleWithOwnEdgeUnderTable(cell.style, side, edge);
  if (style === cell.style) return sheet;
  return { ...sheet, cells: { ...sheet.cells, [key]: { ...cell, style } } };
}

/** All / Outside / No Borders as the legacy per-side flags over a rectangle. */
export function borderModePatchedSheet<S extends CellButtonsSheet>(
  sheet: S,
  mode: BorderMode,
  selRect: { r1: number; c1: number; r2: number; c2: number },
): S {
  let nextSheet = sheet;
  for (let r = selRect.r1; r <= selRect.r2; r += 1) {
    for (let c = selRect.c1; c <= selRect.c2; c += 1) {
      nextSheet = updateCellStyle(
        withDefaultLookStated(nextSheet, r, c),
        r,
        c,
        borderPatchForCell(mode, r, c, selRect),
      );
    }
  }
  return nextSheet;
}

/**
 * Grow Font / Shrink Font (Excel A▲ A▼): each cell of the rectangle steps from
 * ITS OWN size along Excel's size list (a cell without one counts as the
 * grid's default). Null when a cell's stored size is malformed; the cells
 * already at an end of the list are left out.
 */
export function fontSizeStepPlan(
  sheet: { cells: CellMap },
  selRect: { r1: number; c1: number; r2: number; c2: number },
  direction: 1 | -1,
  nativeGrid: boolean,
): Array<{ row: number; column: number; fontSize: number | undefined }> | null {
  const planned: Array<{
    row: number;
    column: number;
    fontSize: number | undefined;
  }> = [];
  for (let row = selRect.r1; row <= selRect.r2; row += 1) {
    for (let column = selRect.c1; column <= selRect.c2; column += 1) {
      const rawStyle = sheet.cells[`${row}:${column}`]?.style as unknown;
      if (
        rawStyle !== undefined &&
        (rawStyle === null ||
          typeof rawStyle !== "object" ||
          Array.isArray(rawStyle))
      ) {
        return null;
      }
      const style = rawStyle as Record<string, unknown> | undefined;
      const rawSize = style?.fontSize;
      if (
        rawSize !== undefined &&
        (typeof rawSize !== "number" ||
          !Number.isFinite(rawSize) ||
          rawSize <= 0)
      ) {
        return null;
      }
      const current = officePixelsToPoints(nativeFontSizeCss(rawSize as number | undefined, style?.nativeFontSize) ?? spreadsheetPtToPx(SPREADSHEET_DEFAULT_FONT_PT));
      // Excel's own size list, past its top and back down by the same steps
      
      const stepped = stepOfficeFontSize(current, direction, "excel");
      const next = stepped === current ? undefined : stepped;
      if (next === undefined) continue;
      planned.push({
        row,
        column,
        fontSize: !nativeGrid && next === SPREADSHEET_DEFAULT_FONT_PT ? undefined : spreadsheetPtToPx(next),
      });
    }
  }
  return planned;
}

/** A cell created on a formatted row or column, with the line's format: its
 *  look as the cell's style and its protection as the cell's own `locked`
 *  (a line's default protection is not a style key -- materializeLineStyles
 *  reads it the same way). Spread into the style, an unlocked column's
 *  `locked: false` came back from the file as the cell's own flag and the
 *  saved page disagreed with its re-read (xlsx-rounds R3-X02). */
function withLineFormat<C extends { style?: unknown; locked?: unknown }>(written: C, inherited: Record<string, unknown>): C {
  const { locked, ...visual } = JSON.parse(JSON.stringify(inherited)) as Record<string, unknown>;
  return {
    ...written,
    style: { ...visual, ...((written.style as Record<string, unknown> | undefined) ?? {}) },
    ...(locked === false && written.locked === undefined ? { locked: false } : {}),
  } as C;
}
