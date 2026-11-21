/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Editing > Fill and the writers that end in the shared paste: the
 * fill handle's drag (Auto Fill), Fill Down / Right / Up / Left, Fill >
 * Series, Flash Fill and Data > Text to Columns, as pure plans over a sheet.
 * Each plan decides what is written; the shared paste transaction
 * (paste-buttons) or the plan's own sheet transaction writes it. Moved out of
 * SpreadsheetWindow so the ribbon and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps the dialogs, the
 * notices, the overwrite confirmation, the commit and the selection.
 */

import { writeLineBlanksIn } from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";
import {
  cellKind,
  hasCellMetadata,
  type CellButtonsSheet,
  type CellMap,
  type SpreadsheetCell,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { placePictureInCell, readCellPicture } from "@mrnote/core/web/lib/spreadsheet/cell-picture.ts";
import { whatIfTablePartRefused, whatIfTablesAfterOverwrite } from "@mrnote/core/web/lib/spreadsheet/what-if-table-definition.ts";
import { runsForWrittenValue } from "@mrnote/core/web/lib/spreadsheet/cell-rich-text.ts";
import { withRecordFontIfPlain } from "@mrnote/core/web/lib/spreadsheet/cell-record-font.ts";
import {
  classifyDelimitedImportCell,
  delimitedImportNumberOptionsAreValid,
  type DelimitedImportDateOrder,
  type DelimitedImportNumberOptions,
} from "./csv-import";
import {
  MAX_CONDITIONAL_FORMATS,
  normalizeRules as normalizeConditionalRules,
  planRulesForFill,
  type ConditionalFormatRule,
} from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import type { DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import {
  normalizeRules as normalizeValidationRules,
  planRulesForValidationFill,
  type ValidationRule,
} from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import { resolveCellChartScalar, spreadsheetCellIsFormula } from "@mrnote/core/web/lib/spreadsheet/display-cache.ts";
import { spreadsheetCellIsText, styleWithOwnTextIdentity } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import {
  autoFillLineIsDates,
  autoFillLineOptions,
  computeFillValues,
  planAutoFillSource,
  planFillSeries,
  type AutoFillValueMode,
  type FillDirection,
  type FillSeriesRange,
  type SeriesDateUnit,
  type SeriesIn,
  type SeriesType,
} from "./fill-series";
import type { FlashFillExample } from "@/lib/spreadsheet/flash-fill";
import {
  readLiteralCellValue,
  shiftFormula,
  type CellValue,
  type SheetEvaluator,
  type SpillOverlay,
} from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { computeEdgeFill, type EdgeFillDirection } from "./grid-ops";
import { sheetWithLinksFilled } from "@mrnote/core/web/lib/spreadsheet/hyperlink-areas.ts";
import { normalizeMerges, rangeHasMerge, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import {
  clipboardScalarFromValue,
  scalarContent,
  type ClipboardSparkline,
  type RichClipboard,
  type RichClipboardCell,
} from "@mrnote/core/web/lib/spreadsheet/paste-special.ts";
import { MAX_SPARKLINES, type Sparkline } from "@mrnote/core/web/lib/spreadsheet/sparkline.ts";
import { sparklineCopiedBy } from "@mrnote/core/web/lib/spreadsheet/sparkline-buttons.ts";
import type { PasteGridOptions, PasteGridRequest } from "./paste-buttons";
import { withConditionalFormatFileCopyRecorded } from "@mrnote/core/web/lib/spreadsheet/structural.ts";
import {
  splitColumnByDelimiters,
  splitColumnByFixedWidth,
  type ColumnSplitResult,
  type TextToColumnsColumnFormat,
  type TextToColumnsQualifier,
} from "@mrnote/core/web/lib/spreadsheet/text-to-columns.ts";
import { type ExcelDateSystemArg } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

type Rect = { r1: number; c1: number; r2: number; c2: number };

// --------------------------------------------------------------- cell reads

/**
 * The typed value Copy snapshots for Paste Values: what the grid shows,
 * unformatted and at full precision. A formula reads its result (a blocked
 * array #SPILL!), an empty cell inside a dynamic array's footprint its spilled
 * element (the same spill overlay the grid draws, so a copied A1:B3 of
 * =SEQUENCE(3,2) carries all six values), and a literal its stored value read
 * the way the engine reads it: a quote-prefixed or Text-format cell is text,
 * any other cell's characters are typed. Nothing is written back.
 */
export function resolveCellClipboardScalar(
  cell: SpreadsheetCell | undefined,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
): CellValue {
  if (cell?.formula || (!cell?.value && overlay.spilled.has(`${row}:${col}`))) {
    return resolveCellChartScalar(cell, row, col, evaluator, overlay);
  }
  // A formula stored as its value alone reads its result like any other.
  if (spreadsheetCellIsFormula(cell)) return evaluator.getComputed(row, col);
  const value = cell?.value ?? "";
  if (value === "") return null;
  return spreadsheetCellIsText(cell) ? value : readLiteralCellValue(value);
}

/**
 * The cell a Fill Down / Right copies from (row, col): the stored cell, or for
 * a dynamic array's result, which has no stored cell of its own, the value the
 * cell shows, as Paste Values does, instead of reading it as a blank that
 * clears the target (live audit B223).
 */
export function cellForFill(
  cells: CellMap,
  row: number,
  col: number,
  evaluator: SheetEvaluator,
  overlay: SpillOverlay,
): SpreadsheetCell | undefined {
  const key = `${row}:${col}`;
  const stored = cells[key];
  if (stored?.value || !overlay.spilled.has(key)) return stored;
  const content = scalarContent(
    clipboardScalarFromValue(
      resolveCellClipboardScalar(stored, row, col, evaluator, overlay),
    ),
  );
  if (!content.raw) return stored;
  return {
    ...(stored ?? {}),
    value: content.raw,
    kind: content.kind,
    ...(content.forceText
      ? { style: { ...(stored?.style ?? {}), quotePrefix: true } }
      : {}),
  };
}

// ---------------------------------------------------- the fill handle's drag

// Build the cell a drag-fill should write at a target coordinate. The fill
// carries the SOURCE cell's style/number-format/hyperlink (Excel copies
// formatting when filling) but takes a fresh value/formula: a shifted formula
// when `formula` is set, otherwise the series/constant `value`. A blank value on
// a source with no carried metadata clears the target.
export function fillCellFrom(
  source: SpreadsheetCell | undefined,
  value: string,
  formula: string | null,
): SpreadsheetCell | null {
  const carried: Partial<SpreadsheetCell> = {};
  if (source?.style && Object.keys(source.style).length > 0) {
    carried.style = { ...source.style };
  }
  // Protection is one of the formats a fill copies.
  if (source?.locked === false) carried.locked = false;
  if (source?.formulaHidden === true) carried.formulaHidden = true;
  if (source?.hyperlink) {
    carried.hyperlink = source.hyperlink;
    // The ScreenTip/display extras belong to the carried link, so they travel
    // with it (a fill-copied link keeps its tooltip, like Excel).
    if (source.hyperlinkTooltip)
      carried.hyperlinkTooltip = source.hyperlinkTooltip;
    if (source.hyperlinkDisplay)
      carried.hyperlinkDisplay = source.hyperlinkDisplay;
  }
  if (formula) {
    // A formula has no runs: the source's record font is the cell's.
    return withRecordFontIfPlain({ ...carried, value, kind: cellKind(value), formula });
  }
  // Spaces are text; only an empty value is blank.
  if (value !== "") {
    // A picture placed in the cell fills as the value it is (Place in Cell).
    const picture = readCellPicture(source);
    if (picture && value === source?.value) {
      return placePictureInCell({ ...carried, value, kind: cellKind(value) }, picture);
    }
    // A copied constant keeps its in-cell formatting (red bold "AB" of
    // "ABcd", live audit B262); a series step writes new text, whose old
    // runs no longer match it.
    const runs = runsForWrittenValue(source?.runs, value);
    // A copied string the workbook typed as one ("FALSE") stays that string
    // (live audit B404); a series step is a new value, read afresh.
    const typedString =
      source?.stringValue === true && value === source.value && source.kind !== "number";
    // A series step's new text has no runs: the record's font is drawn.
    return withRecordFontIfPlain({ ...carried, value, kind: cellKind(value), ...(typedString ? { stringValue: true as const } : {}), ...(runs ? { runs: runs.map((run) => ({ ...run, ...(run.style ? { style: { ...run.style } } : {}) })) } : {}) });
  }
  // Empty value: keep a cell only if it carries formatting (mirrors
  // updateCellValue's blank-cell handling), else drop it.
  if (Object.keys(carried).length > 0) {
    return { ...carried, value: "", kind: "text" };
  }
  return null;
}

// Auto Fill Options > Fill Formatting Only: the target cell keeps its value
// and takes the source cell's formats (none when the source has none).
export function fillFormatsFrom(
  source: SpreadsheetCell | undefined,
  target: SpreadsheetCell | undefined,
): SpreadsheetCell | null {
  const next: SpreadsheetCell = {
    ...(target ?? { value: "", kind: "text" as const }),
  };
  delete next.style;
  if (source?.style && Object.keys(source.style).length > 0) {
    next.style = { ...source.style };
  }
  delete next.locked;
  if (source?.locked === false) next.locked = false;
  delete next.formulaHidden;
  if (source?.formulaHidden === true) next.formulaHidden = true;
  return next.value !== "" || hasCellMetadata(next) ? next : null;
}

/** Auto Fill Options: "auto" is a plain drag and "toggle" a Ctrl+drag (Copy
 *  Cells and Fill Series swapped); anything else is a choice from the button.
 *  Fill Formatting Only writes the source formats alone; Fill Without
 *  Formatting fills values and keeps the target formats. */
export type AutoFillMode = AutoFillValueMode | "toggle" | "formats" | "values";

/** What the Auto Fill Options button offers after a fill. */
export interface AutoFillOffer {
  /** Never "auto": the fill resolved it to Copy Cells or Fill Series. */
  choice: Exclude<AutoFillValueMode, "auto"> | "formats" | "values";
  /** Fill Series would extend the source (else Excel leaves it out). */
  series: boolean;
  /** The source is dates: Fill Days / Weekdays / Months / Years. */
  dates: boolean;
  /** A fill down one column, which Flash Fill can redo. */
  flashFill: boolean;
}

export interface AutoFillContext<S> {
  /** Sheet protection: is a cell of `range` locked (pause-aware in the window). */
  rangeHasLockedCell: (sheet: S, range: Rect) => boolean;
  /** The sheet's merges as the window holds them; read from the sheet when absent. */
  merges?: readonly MergeSpan[];
  /** The workbook's Custom Lists (File > Options), which a drag extends. */
  customLists?: ReadonlyArray<readonly string[]>;
  /** The workbook reads serials in the 1904 date system. */
  date1904: ExcelDateSystemArg;
}

export type AutoFillOutcome<S> =
  | { ok: false; reason: "none" | "locked" | "merge" | "what-if-table" }
  | ({ ok: true; sheet: S } & AutoFillOffer);

/**
 * Drag-to-fill (Excel/WPS parity): fill from a source rectangle into the
 * dragged-out range, in any of the four directions. The dominant axis (the
 * larger overshoot magnitude past the nearest source edge) decides whether we
 * fill vertically (extend each column's series) or horizontally (extend each
 * row's series); the SIGN of that overshoot decides DOWN/RIGHT (target beyond
 * the bottom/right edge) vs UP/LEFT (target beyond the top/left edge). For
 * up/left the source line is read from the near edge outward and the produced
 * series is written stepping away from the source, so a series extends
 * naturally (e.g. dragging 2,1 up yields 0,-1,…). Within a line each cell
 * keeps its own identity: a formula source cell always fills as that
 * formula with its relative references moved (Copy Cells and Fill Series
 * alike, as in Excel), and the line's other cells seed computeFillValues
 * (constant copy, numeric step, or date step) with the formula positions
 * read as blanks, so a formula's text is never stepped as a numbered label.
 * The source cell's formatting is carried onto each filled cell. Targets are
 * clamped to the grid bounds so a drag can never address out-of-grid cells.
 */
export function autoFilledSheet<
  S extends CellButtonsSheet & {
    merges?: MergeSpan[];
    sparklines?: Record<string, Sparkline>;
    conditionalFormats?: ConditionalFormatRule[];
    validations?: ValidationRule[];
    whatIfTables?: unknown;
  },
>(
  sheet: S,
  src: Rect,
  target: { row: number; col: number },
  mode: AutoFillMode,
  context: AutoFillContext<S>,
): AutoFillOutcome<S> {
  const mergeSpans = context.merges ?? normalizeMerges(sheet.merges);
  // SIGNED overshoot past the nearest source edge on each axis: positive =>
  // below/right of the source, negative => above/left of it, 0 => the target
  // is still within the source span on that axis (no fill on that axis).
  const vertExtent =
    target.row > src.r2
      ? target.row - src.r2
      : Math.min(0, target.row - src.r1);
  const horizExtent =
    target.col > src.c2
      ? target.col - src.c2
      : Math.min(0, target.col - src.c1);
  if (vertExtent === 0 && horizExtent === 0) return { ok: false, reason: "none" };
  // Constrain to a single axis (Excel snaps the fill to the larger overshoot).
  const vertical = Math.abs(vertExtent) >= Math.abs(horizExtent);
  const extent = vertical ? vertExtent : horizExtent;
  const direction: FillDirection = vertical
    ? extent >= 0
      ? "down"
      : "up"
    : extent >= 0
      ? "right"
      : "left";
  const count = Math.abs(extent);

  const nextCells: CellMap = { ...sheet.cells };
  // A dragged cell takes its sparkline along (Excel: "copy a sparkline to
  // other cells in a column or row later by dragging"), its data range moved
  // like a relative reference. Fill Formatting Only / Without Formatting
  // leave the targets' sparklines as they are.
  const carried: { sparklines: Record<string, Sparkline> | null } = { sparklines: null };
  const carrySparkline = (
    targetRow: number,
    targetCol: number,
    sourceRow: number,
    sourceCol: number,
  ) => {
    if (mode === "formats" || mode === "values") return;
    const source = sheet.sparklines?.[`${sourceRow}:${sourceCol}`];
    if (!source) return;
    const key = `${targetRow}:${targetCol}`;
    const current = carried.sparklines ?? sheet.sparklines ?? {};
    if (!(key in current) && Object.keys(current).length >= MAX_SPARKLINES) return;
    const copied = sparklineCopiedBy(source, targetRow - sourceRow, targetCol - sourceCol);
    if (!copied) return;
    carried.sparklines ??= { ...(sheet.sparklines ?? {}) };
    carried.sparklines[key] = copied;
  };
  const writeCell = (
    row: number,
    col: number,
    cell: SpreadsheetCell | null,
  ) => {
    const key = `${row}:${col}`;
    if (cell) nextCells[key] = cell;
    else delete nextCells[key];
  };

  // Client-side bounds: never write past the worksheet's last row and
  // column. The used range is not a bound: a fill into the blank rows the
  // grid shows below it grows the sheet (live audit B264).
  const rowLimit = MAX_GRID_ROWS - 1;
  const colLimit = MAX_GRID_COLUMNS - 1;
  let maxRow = sheet.rowCount - 1;
  let maxCol = sheet.columnCount - 1;

  // Protection: refuse the whole fill (Excel all-or-nothing) if any TARGET
  // cell is locked on a protected sheet. The checked box is the fill
  // extension only (excludes the source), clamped to the grid so an
  // overshooting drag does not implicate out-of-grid (default-locked) cells.
  const fillTarget = vertical
    ? direction === "down"
      ? {
          r1: src.r2 + 1,
          r2: Math.min(src.r2 + count, rowLimit),
          c1: src.c1,
          c2: src.c2,
        }
      : {
          r1: Math.max(src.r1 - count, 0),
          r2: src.r1 - 1,
          c1: src.c1,
          c2: src.c2,
        }
    : direction === "right"
      ? {
          r1: src.r1,
          r2: src.r2,
          c1: src.c2 + 1,
          c2: Math.min(src.c2 + count, colLimit),
        }
      : {
          r1: src.r1,
          r2: src.r2,
          c1: Math.max(src.c1 - count, 0),
          c2: src.c1 - 1,
        };
  if (
    fillTarget.r1 <= fillTarget.r2 &&
    fillTarget.c1 <= fillTarget.c2 &&
    context.rangeHasLockedCell(sheet, fillTarget)
  ) {
    return { ok: false, reason: "locked" };
  }
  // A fill writes cells one by one, so a merged cell in its way would
  // take values into the cells it covers, unseen and unsummed; Excel
  // refuses to fill unmerged cells into a merge, as Paste does (live
  // audit B268). A source that carries merges still fills its shape.
  if (
    fillTarget.r1 <= fillTarget.r2 &&
    fillTarget.c1 <= fillTarget.c2 &&
    rangeHasMerge([...mergeSpans], fillTarget) &&
    !rangeHasMerge([...mergeSpans], src)
  ) {
    return { ok: false, reason: "merge" };
  }
  // A What-If Data Table's results change only as one block: a fill over
  // part of them is Excel's "Cannot change part of a data table."; one that
  // covers them all ends the table (its cells take the fill).
  const fillTargetValid = fillTarget.r1 <= fillTarget.r2 && fillTarget.c1 <= fillTarget.c2;
  if (fillTargetValid) writeLineBlanksIn(sheet as unknown as Parameters<typeof writeLineBlanksIn>[0], fillTarget);
  if (fillTargetValid && whatIfTablePartRefused(sheet.whatIfTables, fillTarget)) {
    return { ok: false, reason: "what-if-table" };
  }
  const whatIfTablesLeft = fillTargetValid
    ? whatIfTablesAfterOverwrite(sheet.whatIfTables, [fillTarget])
    : null;

  // What the Auto Fill Options button offers and checks follows from the
  // lines that are not all formulas (formulas always fill shifted).
  let copies = true;
  let extendable = false;
  let dates = true;
  let valueLines = 0;
  // A typed date is a serial under a date format, so the line's number
  // formats go with its values.
  const fillDateSystem = { date1904: context.date1904 };
  const customLists = context.customLists;
  const lineMode = (
    values: string[],
    formats: (string | undefined)[],
  ): AutoFillValueMode => {
    const line = autoFillLineOptions(
      values,
      count,
      direction,
      customLists,
      formats,
      fillDateSystem,
    );
    valueLines += 1;
    copies &&= line.copies;
    extendable ||= line.extendable;
    dates &&= autoFillLineIsDates(values, formats, fillDateSystem);
    if (mode === "toggle") return line.copies ? "series" : "copy";
    return mode === "formats" || mode === "values" ? "auto" : mode;
  };
  const writeFilled = (
    row: number,
    col: number,
    srcCell: SpreadsheetCell | undefined,
    value: string,
    formula: string | null,
  ) => {
    const existing = sheet.cells[`${row}:${col}`];
    writeCell(
      row,
      col,
      mode === "formats"
        ? fillFormatsFrom(srcCell, existing)
        : fillCellFrom(mode === "values" ? existing : srcCell, value, formula),
    );
  };

  if (vertical) {
    const fillingDown = direction === "down";
    for (let c = src.c1; c <= src.c2; c += 1) {
      // Read the source line from the edge nearest the fill toward the far
      // edge: top->bottom for a down fill, bottom->top for an up fill, so the
      // series extends naturally away from the source in the drag direction.
      const lineCells: (SpreadsheetCell | undefined)[] = [];
      if (fillingDown) {
        for (let r = src.r1; r <= src.r2; r += 1) {
          lineCells.push(sheet.cells[`${r}:${c}`]);
        }
      } else {
        for (let r = src.r2; r >= src.r1; r -= 1) {
          lineCells.push(sheet.cells[`${r}:${c}`]);
        }
      }
      const allFormula =
        lineCells.length > 0 && lineCells.every(spreadsheetCellIsFormula);
      // A formula position seeds the series as a blank: its own target
      // takes the moved formula below, never a stepped copy of its text.
      const lineValues = lineCells.map((cell) =>
        spreadsheetCellIsFormula(cell) ? "" : (cell?.value ?? ""),
      );
      const lineFormats = lineCells.map((cell) => cell?.style?.numberFormat);
      const seriesValues = allFormula
        ? []
        : computeFillValues(
            lineValues,
            count,
            direction,
            customLists,
            lineMode(lineValues, lineFormats),
            lineFormats,
            fillDateSystem,
          );
      for (let i = 0; i < count; i += 1) {
        // i counts away from the source: row just past the near edge first.
        const targetRow = fillingDown ? src.r2 + 1 + i : src.r1 - 1 - i;
        if (targetRow < 0 || targetRow > rowLimit) continue;
        // The source cell this target maps to (cycled across the block).
        const srcIndex = i % lineCells.length;
        const srcCell = lineCells[srcIndex];
        carrySparkline(targetRow, c, fillingDown ? src.r1 + srcIndex : src.r2 - srcIndex, c);
        if (spreadsheetCellIsFormula(srcCell)) {
          const srcRow = fillingDown
            ? src.r1 + srcIndex
            : src.r2 - srcIndex;
          const formula = shiftFormula(
            srcCell?.value ?? "",
            targetRow - srcRow,
            0,
          );
          writeFilled(targetRow, c, srcCell, formula, formula);
        } else {
          writeFilled(targetRow, c, srcCell, seriesValues[i] ?? "", null);
        }
        if (targetRow > maxRow) maxRow = targetRow;
      }
    }
  } else {
    const fillingRight = direction === "right";
    for (let r = src.r1; r <= src.r2; r += 1) {
      // Read the source line from the edge nearest the fill: left->right for
      // a right fill, right->left for a left fill.
      const lineCells: (SpreadsheetCell | undefined)[] = [];
      if (fillingRight) {
        for (let c = src.c1; c <= src.c2; c += 1) {
          lineCells.push(sheet.cells[`${r}:${c}`]);
        }
      } else {
        for (let c = src.c2; c >= src.c1; c -= 1) {
          lineCells.push(sheet.cells[`${r}:${c}`]);
        }
      }
      const allFormula =
        lineCells.length > 0 && lineCells.every(spreadsheetCellIsFormula);
      // A formula position seeds the series as a blank: its own target
      // takes the moved formula below, never a stepped copy of its text.
      const lineValues = lineCells.map((cell) =>
        spreadsheetCellIsFormula(cell) ? "" : (cell?.value ?? ""),
      );
      const lineFormats = lineCells.map((cell) => cell?.style?.numberFormat);
      const seriesValues = allFormula
        ? []
        : computeFillValues(
            lineValues,
            count,
            direction,
            customLists,
            lineMode(lineValues, lineFormats),
            lineFormats,
            fillDateSystem,
          );
      for (let i = 0; i < count; i += 1) {
        const targetCol = fillingRight ? src.c2 + 1 + i : src.c1 - 1 - i;
        if (targetCol < 0 || targetCol > colLimit) continue;
        const srcIndex = i % lineCells.length;
        const srcCell = lineCells[srcIndex];
        carrySparkline(r, targetCol, r, fillingRight ? src.c1 + srcIndex : src.c2 - srcIndex);
        if (spreadsheetCellIsFormula(srcCell)) {
          const srcCol = fillingRight
            ? src.c1 + srcIndex
            : src.c2 - srcIndex;
          const formula = shiftFormula(
            srcCell?.value ?? "",
            0,
            targetCol - srcCol,
          );
          writeFilled(r, targetCol, srcCell, formula, formula);
        } else {
          writeFilled(r, targetCol, srcCell, seriesValues[i] ?? "", null);
        }
        if (targetCol > maxCol) maxCol = targetCol;
      }
    }
  }

  // The dragged cells' conditional formats reach over what they filled, and
  // their validation with their content (live audit S-034): Fill Formatting
  // Only carries the formats alone, Fill Without Formatting neither.
  const filled =
    direction === "down"
      ? { ...src, r2: src.r2 + count }
      : direction === "up"
        ? { ...src, r1: src.r1 - count }
        : direction === "right"
          ? { ...src, c2: src.c2 + count }
          : { ...src, c1: src.c1 - count };
  const ruleBounds = { maxRow: MAX_GRID_ROWS - 1, maxCol: MAX_GRID_COLUMNS - 1 };
  const conditionalPlan =
    mode === "values"
      ? null
      : planRulesForFill(normalizeConditionalRules(sheet.conditionalFormats, ruleBounds), src, filled, ruleBounds);
  const validationPlan =
    mode === "values" || mode === "formats"
      ? null
      : planRulesForValidationFill(normalizeValidationRules(sheet.validations), src, filled, shiftFormula, ruleBounds);
  const filledSheet: S = {
    ...sheet,
    cells: nextCells,
    ...(carried.sparklines ? { sparklines: carried.sparklines } : {}),
    rowCount: Math.max(sheet.rowCount, maxRow + 1),
    columnCount: Math.max(sheet.columnCount, maxCol + 1),
  };
  if (conditionalPlan?.changed) {
    if (conditionalPlan.rules.length > 0) filledSheet.conditionalFormats = conditionalPlan.rules;
    else delete filledSheet.conditionalFormats;
  }
  if (validationPlan?.changed) {
    if (validationPlan.rules.length > 0) filledSheet.validations = validationPlan.rules;
    else delete filledSheet.validations;
  }
  if (whatIfTablesLeft) {
    if (whatIfTablesLeft.length > 0) filledSheet.whatIfTables = whatIfTablesLeft;
    else delete filledSheet.whatIfTables;
  }
  // Where the sheet holds links over large areas, the filled cells' links
  // are the fill's own matter (hyperlink-areas, CRITERIA XC-199). Fill
  // Formatting Only and Fill Without Formatting carry no link.
  const sourceRows = src.r2 - src.r1 + 1;
  const sourceCols = src.c2 - src.c1 + 1;
  const linkedSheet =
    mode === "formats" || mode === "values"
      ? filledSheet
      : sheetWithLinksFilled(
          sheet,
          filledSheet,
          vertical
            ? {
                r1: extent > 0 ? src.r2 + 1 : src.r1 + extent,
                r2: extent > 0 ? src.r2 + extent : src.r1 - 1,
                c1: src.c1,
                c2: src.c2,
              }
            : {
                r1: src.r1,
                r2: src.r2,
                c1: extent > 0 ? src.c2 + 1 : src.c1 + extent,
                c2: extent > 0 ? src.c2 + extent : src.c1 - 1,
              },
          (row, col) =>
            vertical
              ? {
                  row: extent > 0 ? src.r1 + ((row - src.r2 - 1) % sourceRows) : src.r2 - ((src.r1 - 1 - row) % sourceRows),
                  col,
                }
              : {
                  row,
                  col: extent > 0 ? src.c1 + ((col - src.c2 - 1) % sourceCols) : src.c2 - ((src.c1 - 1 - col) % sourceCols),
                },
        );
  // Past the page's bound the file's own rules over the source reach over
  
  const recordedSheet =
    conditionalPlan && fillReachesFileRules(sheet.conditionalFormats)
      ? withConditionalFormatFileCopyRecorded(linkedSheet, linkedSheet, fillFileCopy(src, filled))
      : linkedSheet;
  return {
    ok: true,
    sheet: recordedSheet,
    choice:
      mode === "toggle"
        ? copies ? "series" : "copy"
        : mode === "auto"
          ? copies ? "copy" : "series"
          : mode,
    series: extendable,
    dates: valueLines > 0 && dates,
    flashFill:
      direction === "down" && src.c1 === src.c2 && src.c1 > 0 && valueLines > 0,
  };
}






export function fillReachesFileRules(rules: readonly unknown[] | undefined): boolean {
  return Array.isArray(rules) && rules.length >= MAX_CONDITIONAL_FORMATS;
}

/** The record of a fill of `source` over `filled` (which holds it) for the
 * write-back to reach the file's own rules over the source, as the page's
 * reach (`withConditionalFormatFileCopyRecorded`, `fill`). */
export function fillFileCopy(source: Rect, filled: Rect) {
  return {
    rect: [source.r1, source.c1, source.r2, source.c2] as [number, number, number, number],
    at: [source.r1, source.c1] as [number, number],
    page: true,
    fill: [filled.r1, filled.c1, filled.r2, filled.c2] as [number, number, number, number],
  };
}








export function edgeFillRulePlans(
  request: Pick<PasteGridRequest, "baseRow" | "baseCol" | "rows">,
  direction: EdgeFillDirection,
  conditionalFormats: readonly ConditionalFormatRule[],
  validations: readonly ValidationRule[],
  bounds: { maxRow: number; maxCol: number },
): Pick<PasteGridOptions, "conditionalFormatPlan" | "validationPlan" | "conditionalFormatFileCopy"> {
  const destination = {
    r1: request.baseRow,
    c1: request.baseCol,
    r2: request.baseRow + request.rows.length - 1,
    c2: request.baseCol + Math.max(1, request.rows[0]?.length ?? 1) - 1,
  };
  const source =
    direction === "down"
      ? { ...destination, r1: destination.r1 - 1, r2: destination.r1 - 1 }
      : direction === "up"
        ? { ...destination, r1: destination.r2 + 1, r2: destination.r2 + 1 }
        : direction === "right"
          ? { ...destination, c1: destination.c1 - 1, c2: destination.c1 - 1 }
          : { ...destination, c1: destination.c2 + 1, c2: destination.c2 + 1 };
  const filled = {
    r1: Math.min(source.r1, destination.r1),
    c1: Math.min(source.c1, destination.c1),
    r2: Math.max(source.r2, destination.r2),
    c2: Math.max(source.c2, destination.c2),
  };
  const conditionalFormatPlan = planRulesForFill(conditionalFormats, source, filled, bounds);
  const validationPlan = planRulesForValidationFill(validations, source, filled, shiftFormula, bounds);
  return {
    ...(conditionalFormatPlan ? { conditionalFormatPlan } : {}),
    ...(validationPlan ? { validationPlan } : {}),
    ...(conditionalFormatPlan && fillReachesFileRules(conditionalFormats)
      ? { conditionalFormatFileCopy: { source: null, ...fillFileCopy(source, filled) } }
      : {}),
  };
}

// ------------------------------------------------- Fill Down / Right / Up / Left

/**
 * Home > Fill Down / Up / Right / Left (Ctrl/Cmd+D, Ctrl/Cmd+R): the edge
 * row or column on the side the fill comes from is copied into the rest of
 * the selection. Excel copies "the contents and format" of the source cell:
 * a formula moves its relative references, any other value (a text reading
 * "=A2" included) is copied as it is, and the target takes the source's
 * formats, number format and protection, replacing its own; its note and
 * hyperlink stay. A selection one row (one column) deep fills from the row
 * above (the column beside) it, as a single cell does. The source cell's
 * sparkline (`sparklineAt`) is copied too, its data range moved like a
 * relative reference (Excel: "copy a sparkline to other cells ... using Fill
 * Down"). The plan is the paste the window writes (one Undo step, the
 * protection guard); null when the fill has nothing to copy.
 */
export function edgeFillRequest(
  selection: Rect,
  direction: EdgeFillDirection,
  rowCount: number,
  cellAt: (row: number, col: number) => SpreadsheetCell | undefined,
  sparklineAt?: (row: number, col: number) => ClipboardSparkline | undefined,
): PasteGridRequest | null {
  const sel = { ...selection };
  const vertical = direction === "down" || direction === "up";
  if (vertical ? sel.r1 === sel.r2 : sel.c1 === sel.c2) {
    if (direction === "down") {
      if (sel.r1 === 0) return null;
      sel.r1 -= 1;
    } else if (direction === "up") {
      if (sel.r2 + 1 >= rowCount) return null;
      sel.r2 += 1;
    } else if (direction === "right") {
      if (sel.c1 === 0) return null;
      sel.c1 -= 1;
    } else {
      // Fill Left from beyond the last filled column would only blank the
      // selection; it fires only when a cell to the right has content.
      let sourced = false;
      for (let r = sel.r1; r <= sel.r2 && !sourced; r += 1) {
        sourced = Boolean(cellAt(r, sel.c2 + 1) || sparklineAt?.(r, sel.c2 + 1));
      }
      if (!sourced) return null;
      sel.c2 += 1;
    }
  }
  const fill = computeEdgeFill(sel, direction, (row, col) => {
    const cell = cellAt(row, col);
    return cell && { value: cell.value, formula: spreadsheetCellIsFormula(cell) };
  });
  if (!fill || fill.grid.length === 0 || fill.grid[0].length === 0) return null;
  const cells: RichClipboard = fill.grid.map((line, lineIndex) =>
    line.map((entry, entryIndex): RichClipboardCell => {
      const source = cellAt(entry.sourceRow, entry.sourceCol);
      const sourceSparkline = sparklineAt?.(entry.sourceRow, entry.sourceCol);
      const sparkline = sourceSparkline
        ? sparklineCopiedBy(
            sourceSparkline.value,
            fill.row + lineIndex - entry.sourceRow,
            fill.col + entryIndex - entry.sourceCol,
          )
        : null;
      return {
        raw: entry.value,
        computed: entry.value,
        formula: entry.formula,
        ...(!entry.formula && source?.kind ? { kind: source.kind } : {}),
        // A string the workbook typed as one ("FALSE") fills as that string
        // (live audit B404).
        ...(!entry.formula && source?.stringValue === true ? { stringValue: true as const } : {}),
        // Protection is one format with two halves, Locked and Hidden; the
        // fill carries both, as a copy does (live audit B394).
        locked: source?.locked !== false,
        ...(source?.formulaHidden === true ? { formulaHidden: true as const } : {}),
        ...(source?.style ? { style: source.style } : {}),
        ...(source?.runs && !entry.formula ? { runs: source.runs } : {}),
        // The filled text brings its furigana and Show Phonetic Field.
        ...(source?.phonetic && !entry.formula ? { phonetic: source.phonetic } : {}),
        // Fill Down / Right copies a picture placed in the cell, as a value.
        ...(() => {
          const picture = entry.formula ? null : readCellPicture(source);
          return picture ? { picture } : {};
        })(),
        ...(sourceSparkline && sparkline
          ? { sparkline: { sheetId: sourceSparkline.sheetId, value: sparkline } }
          : {}),
      };
    }),
  );
  return {
    baseRow: fill.row,
    baseCol: fill.col,
    rows: fill.grid.map((line) => line.map((entry) => entry.value)),
    clipboardCells: cells,
    // Fill Down from a merged cell wrote its text into the cells the merge
    // covers (xlsx-rounds WALK2 R293-X15: B4:B6 merged, Fill Down over B4:B8
    // put the heading in B5 and B6, which never show and did not save).
    // A merged cell the fill's range cuts refuses it, as Excel does
    // (`refuseMergesCutIn`, R423-X56 WALK1).
    options: { fillContentAndFormats: true, fillLinksFrom: direction, refuseValuesUnderMerges: true, refuseMergesCutIn: { ...sel } },
  };
}

// ----------------------------------------------------------- Fill > Series

/** What Fill > Series reads from its dialog. */
export interface SeriesFillInput {
  /** The selection (0-based, inclusive). */
  range: FillSeriesRange;
  seriesIn: SeriesIn;
  type: Exclude<SeriesType, "autoFill">;
  /** Step value (NaN when the box does not read as a number). */
  step: number;
  /** Stop value text, null for none. */
  stop: string | null;
  dateUnit: SeriesDateUnit;
  /** Trend (Linear / Growth fitted to the starting values). */
  trend: boolean;
  rowLimit: number;
  columnLimit: number;
  date1904: ExcelDateSystemArg;
}

export type SeriesFillError = "need-range" | "no-values" | "bad-step";

/**
 * Fill > Series (Excel "Home > Fill > Series…"): every row or every column of
 * the selection (Series in) is filled from its own first cell with the step
 * and stop value, or refitted from its starting values with Trend; a single
 * cell with a stop value runs on until it passes the stop. planFillSeries
 * decides WHAT is written; the plan is one paste, so undo + autosave + sheet
 * protection apply once. The filled cells take the seed's formats, as
 * Excel's Series does: a date entered normally (a serial under a date format)
 * fills as dates, not as bare serial numbers (live audit B153). Text identity
 * is the written value's own, never the seed's quote prefix. The seed's
 * formats travel as a fill's do (Fill Down's contract): the target keeps its
 * own note and hyperlink whatever the seed looks like (live audit B276).
 */
export function seriesFillRequest(
  input: SeriesFillInput,
  /** A seed as a number or date sees it: a formula's result. */
  seedValueAt: (row: number, col: number) => string,
  seedCellAt: (row: number, col: number) => SpreadsheetCell | undefined,
): { ok: false; error: SeriesFillError } | { ok: true; request: PasteGridRequest; count: number } {
  const { range, seriesIn } = input;
  const trend = input.trend && input.type !== "date";
  if (!trend && !Number.isFinite(input.step)) return { ok: false, error: "bad-step" };
  const writes = planFillSeries(
    {
      date1904: input.date1904,
      range,
      seriesIn,
      type: input.type,
      step: Number.isFinite(input.step) ? input.step : 1,
      stop: input.stop,
      dateUnit: input.dateUnit,
      trend,
      rowLimit: input.rowLimit,
      columnLimit: input.columnLimit,
    },
    seedValueAt,
  );
  if (writes.length === 0) {
    return {
      ok: false,
      error:
        range.r1 === range.r2 && range.c1 === range.c2 && input.stop === null
          ? "need-range"
          : "no-values",
    };
  }
  // One rectangular grid over every write; the cells the series does not
  // reach stay holes, which the paste leaves untouched.
  let baseRow = Infinity;
  let baseCol = Infinity;
  let lastRow = -Infinity;
  let lastCol = -Infinity;
  for (const write of writes) {
    baseRow = Math.min(baseRow, write.row);
    baseCol = Math.min(baseCol, write.col);
    lastRow = Math.max(lastRow, write.row);
    lastCol = Math.max(lastCol, write.col);
  }
  const grid: string[][] = Array.from(
    { length: lastRow - baseRow + 1 },
    () => new Array<string>(lastCol - baseCol + 1),
  );
  for (const write of writes) {
    grid[write.row - baseRow][write.col - baseCol] = write.value;
  }
  let seedStyled = false;
  const seedFormats: RichClipboard = grid.map((row) => new Array(row.length));
  for (const write of writes) {
    const seed =
      seriesIn === "columns"
        ? seedCellAt(range.r1, write.col)
        : seedCellAt(write.row, range.c1);
    const style = styleWithOwnTextIdentity(seed?.style, undefined);
    if (style && Object.keys(style).length > 0) seedStyled = true;
    seedFormats[write.row - baseRow][write.col - baseCol] = {
      raw: write.value,
      computed: write.value,
      formula: false,
      ...(style && Object.keys(style).length > 0 ? { style } : {}),
    };
  }
  return {
    ok: true,
    request: {
      baseRow,
      baseCol,
      rows: grid,
      ...(seedStyled ? { clipboardCells: seedFormats } : {}),
      // A merged cell the series' range cuts refuses it, as Excel refuses a
      // Fill (`refuseMergesCutIn`; R427-X85 WALK1: Fill > Series down G10:G15
      // wrote the anchor of the merged G15:G17 and the block lost its bottom
      // edge on the page while the file kept it).
      options: {
        fillContentAndFormats: true,
        refuseValuesUnderMerges: true,
        refuseMergesCutIn: {
          r1: Math.min(range.r1, baseRow),
          c1: Math.min(range.c1, baseCol),
          r2: Math.max(range.r2, lastRow),
          c2: Math.max(range.c2, lastCol),
        },
      },
    },
    count: writes.length,
  };
}

/**
 * Fill > Series > AutoFill: the fill handle's own fill from the selection's
 * filled leading rows (columns) into the rest of it. Null when the selection
 * has no starting values.
 */
export function seriesAutoFillPlan(
  range: FillSeriesRange,
  seriesIn: SeriesIn,
  isBlank: (row: number, col: number) => boolean,
): { source: Rect; target: { row: number; col: number }; fillTarget: Rect } | null {
  const plan = planAutoFillSource(range, seriesIn, isBlank);
  if (!plan) return null;
  const fillTarget =
    seriesIn === "columns"
      ? { ...range, r1: plan.source.r2 + 1 }
      : { ...range, c1: plan.source.c2 + 1 };
  return { source: plan.source, target: plan.target, fillTarget };
}

// --------------------------------------------------------------- Flash Fill

export type FlashFillError = "no-source" | "no-examples" | "no-targets" | "no-pattern";

/**
 * Flash Fill (Excel/WPS "Data > Flash Fill", Ctrl+E): learn a string transform
 * from the rows the user has ALREADY filled in the active column and apply it
 * to the empty rows below. The "examples" are the non-empty cells of the
 * active column; each example's inputs are the cells to its LEFT in the same
 * row (columns 0..col-1), read as the cells SHOW them (a formula's result,
 * never its source text, live audit B274). Flash Fill learns from the block
 * of data the active cell sits in (its current region: the rows joined to it
 * without a blank row between), as Excel does; a table's header row is never
 * an example. A target is an empty active-column cell with a left input,
 * below the first example. Only the resolved cells are written, in ONE paste
 * over the span (atomic, one undo step, all-or-nothing protection); a gap row
 * re-writes its own value, so it stays as it was. What Flash Fill writes is
 * text it extracted, never a formula.
 *
 * The learning itself (the string-program synthesis) runs on the server:
 * engine `tools`, `sheet.flashFill` (`lib/spreadsheet/flash-fill.ts`; the
 * headless command host composes the same steps in
 * `lib/spreadsheet/fill-buttons-engine.ts`). This is the PLAN -- the
 * examples and the targets' inputs -- and {@link flashFillRequestFromResolved}
 * is the APPLY of the answer.
 */
export interface FlashFillPlan {
  col: number;
  examples: FlashFillExample[];
  targetRows: number[];
  /** Each target row's inputs (the cells to its left, as shown). */
  targetInputs: string[][];
}

export function flashFillPlan(
  cells: CellMap,
  active: { row: number; col: number },
  rowCount: number,
  display: (row: number, col: number) => string,
  dataTables: readonly DataTable[],
): { ok: false; error: FlashFillError } | ({ ok: true } & FlashFillPlan) {
  const col = active.col;
  // Flash Fill needs at least one source column to the LEFT to learn from.
  if (col === 0) return { ok: false, error: "no-source" };
  const inputsAt = (row: number): string[] => {
    const inputs: string[] = [];
    for (let c = 0; c < col; c += 1) inputs.push(display(row, c));
    return inputs;
  };
  const lastRow = rowCount - 1;
  const examples: FlashFillExample[] = [];
  const candidateTargets: number[] = [];
  let firstExampleRow = -1;
  const rowHasContent = (row: number) =>
    String(cells[`${row}:${col}`]?.value ?? "").trim() !== "" ||
    inputsAt(row).some((v) => v.trim() !== "");
  let regionTop = active.row;
  while (regionTop > 0 && rowHasContent(regionTop - 1)) regionTop -= 1;
  let regionBottom = active.row;
  while (regionBottom < lastRow && rowHasContent(regionBottom + 1)) regionBottom += 1;
  // A table's header row names its columns; it is never an example of the
  // transformation ("Full Name" -> "First Name" teaches nothing).
  const isTableHeaderRow = (row: number) =>
    dataTables.some(
      (table) =>
        table.headerRow &&
        row === table.range.r1 &&
        col >= table.range.c1 &&
        col <= table.range.c2,
    );
  for (let row = regionTop; row <= regionBottom; row += 1) {
    if (isTableHeaderRow(row)) continue;
    const own = String(cells[`${row}:${col}`]?.value ?? "");
    const inputs = inputsAt(row);
    if (own.trim() !== "") {
      examples.push({ inputs, output: own });
      if (firstExampleRow < 0) firstExampleRow = row;
    } else if (inputs.some((v) => v.trim() !== "")) {
      candidateTargets.push(row);
    }
  }
  // Keep only targets below the first example row.
  const targetRows = candidateTargets.filter((row) => row > firstExampleRow);
  if (examples.length === 0) return { ok: false, error: "no-examples" };
  if (targetRows.length === 0) return { ok: false, error: "no-targets" };
  return { ok: true, col, examples, targetRows, targetInputs: targetRows.map((row) => inputsAt(row)) };
}

/**
 * Flash Fill's write: the resolved value of each target row (null: the row
 * stays blank; the whole answer null: no pattern) as ONE paste over the span.
 */
export function flashFillRequestFromResolved(
  cells: CellMap,
  plan: FlashFillPlan,
  resolved: readonly (string | null)[] | null,
): { ok: false; error: FlashFillError } | { ok: true; request: PasteGridRequest; count: number } {
  if (!resolved) return { ok: false, error: "no-pattern" };
  const { col, targetRows } = plan;
  const writes: { row: number; value: string }[] = [];
  for (let i = 0; i < targetRows.length; i += 1) {
    const value = resolved[i];
    if (value === null) continue; // un-fillable row stays blank
    writes.push({ row: targetRows[i], value });
  }
  if (writes.length === 0) return { ok: false, error: "no-pattern" };
  const startRow = writes[0].row;
  const endRow = writes[writes.length - 1].row;
  const valueByRow = new Map<number, string>();
  for (const w of writes) valueByRow.set(w.row, w.value);
  const grid: string[][] = [];
  for (let row = startRow; row <= endRow; row += 1) {
    grid.push([
      valueByRow.get(row) ??
        String(cells[`${row}:${col}`]?.value ?? ""),
    ]);
  }
  const typedContent = grid.map((line) =>
    line.map((value) => ({
      raw: value,
      formula: false,
      kind: "text" as const,
      forceText: value.trim().startsWith("="),
    })),
  );
  return {
    ok: true,
    request: { baseRow: startRow, baseCol: col, rows: grid, options: { typedContent } },
    count: writes.length,
  };
}

// ---------------------------------------------------------- Text to Columns

/** Data > Text to Columns, as its wizard leaves it on Finish. */
export interface TextToColumnsInput {
  /** The selection: one column, rows r1..r2 (r2 clamped to the sheet). */
  range: Rect;
  /** The source column's cells, as the split reads them (a formula's shown result). */
  sourceCells: Record<string, { value: string }>;
  split:
    | { mode: "delimited"; delimiters: readonly string[]; qualifier: TextToColumnsQualifier; consecutive: boolean }
    | { mode: "fixedWidth"; breaks: readonly number[] };
  /** Top-left output cell; null when the destination does not read. */
  destination: { row: number; col: number } | null;
  /** Column data format per parsed field (missing = General). */
  columnFormats: readonly TextToColumnsColumnFormat[];
  /** Date order per parsed field (missing = the locale's). */
  dateOrders: readonly DelimitedImportDateOrder[];
  /** How many fields the formats cover (the wizard's preview width);
   *  the whole selection's widest split when absent. */
  fieldCount?: number;
  locale: string;
  numberOptions: DelimitedImportNumberOptions;
  date1904: ExcelDateSystemArg;
  /** Sheet protection: is a cell of `range` locked. */
  rangeHasLockedCell: (range: Rect) => boolean;
  /** A cell's stored value, for the overwrite check. */
  valueAt: (row: number, col: number) => string;
}

export type TextToColumnsError =
  | "single-column"
  | "bad-destination"
  | "number-separators"
  | "bad-breaks"
  | "bad-delimiter"
  | "no-split"
  | "all-skipped"
  | "out-of-bounds"
  | "locked";

/**
 * Text to Columns: the pure splitters plan the bounded writes; the plan adds
 * destination / protection validation and builds the one paste the window
 * commits. `occupied` says a non-empty cell other than the source would be
 * overwritten (the window asks first; a command proceeds as Range.TextToColumns
 * does with DisplayAlerts off).
 */
export function textToColumnsRequest(
  input: TextToColumnsInput,
): { ok: false; error: TextToColumnsError } | { ok: true; request: PasteGridRequest; width: number; occupied: boolean } {
  const { range, destination } = input;
  if (range.c1 !== range.c2) return { ok: false, error: "single-column" };
  const col = range.c1;
  const r1 = range.r1;
  const r2 = range.r2;
  if (
    !destination ||
    destination.row < 0 ||
    destination.col < 0 ||
    destination.row + Math.max(0, r2 - r1) >= MAX_GRID_ROWS ||
    destination.col >= MAX_GRID_COLUMNS
  ) {
    return { ok: false, error: "bad-destination" };
  }
  if (!delimitedImportNumberOptionsAreValid(input.locale, input.numberOptions)) {
    return { ok: false, error: "number-separators" };
  }
  const split = input.split;
  const splitWith = (formats?: readonly TextToColumnsColumnFormat[]): ColumnSplitResult =>
    split.mode === "fixedWidth"
      ? splitColumnByFixedWidth(
          input.sourceCells,
          col,
          { r1, r2 },
          split.breaks,
          MAX_GRID_COLUMNS,
          formats ? { destination, columnFormats: formats } : {},
        )
      : splitColumnByDelimiters(
          input.sourceCells,
          col,
          { r1, r2 },
          {
            delimiters: split.delimiters,
            textQualifier: split.qualifier,
            consecutiveDelimiters: split.consecutive,
            ...(formats ? { destination, columnFormats: formats } : {}),
          },
          MAX_GRID_COLUMNS,
        );
  if (split.mode === "fixedWidth" && split.breaks.length === 0) {
    return { ok: false, error: "bad-breaks" };
  }
  if (split.mode === "delimited" && split.delimiters.length === 0) {
    return { ok: false, error: "bad-delimiter" };
  }
  const fieldCount =
    input.fieldCount ??
    (() => {
      const whole = splitWith();
      return Math.max(1, whole.sourceWidth ?? whole.width);
    })();
  const formats = Array.from(
    { length: fieldCount },
    (_, index) => input.columnFormats[index] ?? "general",
  );
  const dateOrders = Array.from(
    { length: fieldCount },
    (_, index) => input.dateOrders[index] ?? "locale",
  );
  const result = splitWith(formats);
  if ((result.sourceWidth ?? result.width) <= 1) return { ok: false, error: "no-split" };
  if (result.width <= 0 || result.writes.length === 0) {
    const everyFieldSkipped = Array.from(
      { length: result.sourceWidth ?? formats.length },
      (_, index) => formats[index] ?? "general",
    ).every((format) => format === "skip");
    return { ok: false, error: everyFieldSkipped ? "all-skipped" : "no-split" };
  }
  const requestedWidth = Array.from(
    { length: result.sourceWidth ?? result.width },
    (_, index) => formats[index] ?? "general",
  ).filter((format) => format !== "skip").length;
  if (requestedWidth > MAX_GRID_COLUMNS - destination.col) {
    return { ok: false, error: "out-of-bounds" };
  }
  if (
    input.rangeHasLockedCell({
      r1: destination.row,
      r2: destination.row + (r2 - r1),
      c1: destination.col,
      c2: destination.col + result.width - 1,
    })
  ) {
    return { ok: false, error: "locked" };
  }
  // Overwriting destination data needs the person's yes, except the source
  // cell itself when the default in-place destination is retained. A cell of
  // spaces is text (LEN 3, not blank), as much data to warn about as any
  // other (live audit B272).
  const sourceKeys = new Set<string>();
  for (let row = r1; row <= r2; row += 1) sourceKeys.add(`${row}:${col}`);
  let occupied = false;
  outer: for (let row = destination.row; row <= destination.row + (r2 - r1); row += 1) {
    for (let targetCol = destination.col; targetCol < destination.col + result.width; targetCol += 1) {
      if (sourceKeys.has(`${row}:${targetCol}`)) continue;
      if (input.valueAt(row, targetCol) !== "") {
        occupied = true;
        break outer;
      }
    }
  }
  // A row-major grid the shared paste consumes. Group writes by row, place
  // each part at (col - destination col); gaps (a row that split into fewer
  // parts) stay "" so a shorter row never stamps over a longer row's column:
  // the paste clears a "" cell, which is exactly the trailing-clear Excel
  // does within the fan-out.
  const grid = Array.from(
    { length: Math.max(1, r2 - r1 + 1) },
    () => new Array<string>(result.width).fill(""),
  );
  let wrote = false;
  for (const w of result.writes) {
    const gr = w.row - destination.row;
    const gc = w.col - destination.col;
    if (gr >= 0 && gr < grid.length && gc >= 0 && gc < result.width) {
      grid[gr][gc] = w.value;
      wrote = true;
    }
  }
  if (!wrote) return { ok: false, error: "no-split" };
  const emittedDateOrders = dateOrders
    .filter((_dateOrder, index) => formats[index] !== "skip")
    .slice(0, result.width);
  const typedGrid = grid.map((row) =>
    row.map((value, index) =>
      classifyDelimitedImportCell(
        value,
        input.locale,
        result.columnFormats?.[index] ?? "general",
        emittedDateOrders[index] ?? "locale",
        input.numberOptions,
        { date1904: input.date1904 },
      ),
    ),
  );
  return {
    ok: true,
    request: {
      baseRow: destination.row,
      baseCol: destination.col,
      rows: grid,
      options: {
        textToColumnsFormats: result.columnFormats,
        delimitedImportCells: typedGrid,
        refuseValuesUnderMerges: true,
      },
    },
    width: result.width,
    occupied,
  };
}
