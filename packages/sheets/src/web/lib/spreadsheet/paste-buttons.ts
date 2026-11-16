/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The window's shared paste write (Paste, Paste Special, Fill Down / Right /
 * Up / Left, Fill > Series, Flash Fill, Text to Columns, CSV import, What-If
 * Data Table and every other bulk writer) as one pure sheet transaction: a 2D
 * block of values written from a top-left cell, growing the sheet to fit,
 * with the copied cells' formats, a Cut's move and its formula follow-up,
 * authoritative merges, Skip blanks and Table AutoExpansion. Moved out of
 * SpreadsheetWindow so the window and the spreadsheet editor commands run the
 * same code; behaviour is unchanged. The window keeps the notices, the commit
 * (undo, autosave), the AutoCorrect Options offer and the selection move.
 */

import { styleStatesAFormat } from "./style-source-xf";
import { lineBlankAt, writeLineBlanksIn } from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";
import {
  cellKind,
  copyCellMap,
  hasCellMetadata,
  writeCellValueInto,
  writeSpreadsheetEntryInto,
  type CellMap,
  type SpreadsheetCell,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { normalizeCellPicture, placePictureInCell } from "@mrnote/core/web/lib/spreadsheet/cell-picture.ts";
import { clonePhonetic, phoneticAfterValueWrite } from "@mrnote/core/web/lib/spreadsheet/cell-phonetic.ts";
import { normalizeCellRuns, textFromRuns } from "@mrnote/core/web/lib/spreadsheet/cell-rich-text.ts";
import { withRecordFontIfPlain } from "@mrnote/core/web/lib/spreadsheet/cell-record-font.ts";
import { cloneCellCommentForPaste } from "@mrnote/core/web/lib/spreadsheet/comments.ts";
import {
  normalizeRule as normalizeConditionalRule,
  planRulesForCellMove,
  ruleOverlaps as conditionalRuleOverlaps,
  type ConditionalFormatRule,
} from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import type { DelimitedImportCellValue } from "./csv-import";
import {
  clearRulesInRange as clearValidationRulesInRange,
  normalizeRule as normalizeValidationRule,
  planRulesForValidationPaste,
  ruleForCell,
  type ValidationRule,
} from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import { styleWithOwnTextIdentity } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import { adjustFormulaForCellMove, adjustFormulaForCrossSheetMove, shiftFormula } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import {
  coveredCells,
  addMerge,
  normalizeMerges,
  cellsAfterMergesRemoved,
  removeMergesInRange,
  spansOverlap,
  type MergeSpan,
} from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { tableCellRole, type DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { mergeRectTouchesTable } from "./merge-buttons";
import {
  richClipboardCellIsBlank,
  richClipboardCellSourceStyle,
  type PastedContent,
  type RichClipboard,
} from "@mrnote/core/web/lib/spreadsheet/paste-special.ts";
import { rewriteConditionalFormatReferences, rewriteSheetFormulaReferences, rewriteValidationReferences } from "./sheet-formula-references";
import { MAX_SPARKLINES, type Sparkline } from "@mrnote/core/web/lib/spreadsheet/sparkline.ts";
import { spreadsheetTableAutoCorrectSettings } from "./table-autocorrect";
import { planDataTableResize, type TableButtonSheet } from "./table-buttons";
import type { TextToColumnsColumnFormat } from "@mrnote/core/web/lib/spreadsheet/text-to-columns.ts";
import { whatIfTablePartRefused, whatIfTablesAfterOverwrite, type WhatIfTableDefinition } from "@mrnote/core/web/lib/spreadsheet/what-if-table-definition.ts";
import { type ExcelDateSystemArg } from "@mrnote/core/web/lib/spreadsheet/excel-1900-calendar.ts";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** The parts of a worksheet a paste reads and writes. */
export interface PasteButtonSheet extends TableButtonSheet {
  conditionalFormats?: ConditionalFormatRule[];
  validations?: ValidationRule[];
  sparklines?: Record<string, Sparkline>;
  /** What-If Data Tables (what-if-table-definition). */
  whatIfTables?: WhatIfTableDefinition[];
}

/** The style keys Paste > No Borders keeps from the destination. */
export const RICH_PASTE_BORDER_STYLE_KEYS = [
  "border",
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
  "borderEdges",
] as const;

export interface PasteGridOptions {
  /**
   * Plain text from another program, or Ctrl+Enter's entry: each value is
   * read as a typed entry in this locale (writeSpreadsheetEntryInto), so a
   * pasted "50%" or "2026/9/21" lands as the value and format Excel gives it.
   * `typedNumerals` (Ctrl+Enter's typed entry only, never a paste): its digits
   * read with the interface language's separators (5,5 in German).
   */
  parseEntries?: { locale: string; date1904?: ExcelDateSystemArg; typedNumerals?: boolean };
  /**
   * The whole destination rectangle is selected after the paste, as Excel
   * leaves it (Microsoft Support, "Move or copy cells and cell contents"), so
   * the status bar totals what just landed. The Paste commands and a text
   * import into the sheet (whose destination may differ from the selection)
   * set it: fills, Text to Columns and what-if writes leave the selection
   * where the user put it. (Read by the window, not by the write.)
   */
  selectPasted?: boolean;
  /** Excel Paste > No Borders keeps every destination border facet verbatim. */
  preserveDestinationBorders?: boolean;
  /** Do not create content history/save work for an identical resulting sheet. */
  skipIfUnchanged?: boolean;
  /** An explicit empty snapshot clears merges fully contained by the target. */
  authoritativeMerges?: boolean;
  /**
   * A write that carries no merges of its own (Fill > Series, Text to
   * Columns) is refused when it would put a value in a cell an existing merge
   * covers: such a cell holds nothing, the value never shows and the file
   * does not keep it (xlsx-rounds WALK R287-X04 F14:F15 under F13:F15,
   * R287-X07 B1:C1 under A1:H1). Excel refuses both ("all the merged cells
   * need to be the same size"). A write to the merge's own top-left stands.
   */
  refuseValuesUnderMerges?: boolean;
  /**
   * The range a write reads and writes (a Fill's source line and the cells it
   * fills): refused when an existing merged cell lies partly inside it, as
   * Excel refuses ("We can't do that to a merged cell.", Windows Excel 16.0
   * build 20527, xlsx-rounds/probe/night3/fillright-merge: J38:J45 copied to
   * K38 across I39:K39). Written anyway, Fill Right over K38:K45 gave K39 the
   * borders of J39 and the merge I39:K39 lost its right edge in the file
   * (R423-X56 WALK1).
   */
  refuseMergesCutIn?: { r1: number; c1: number; r2: number; c2: number };
  /** Sheet-level formatting planned before any cell or merge mutation. */
  conditionalFormatPlan?: {
    rules: ConditionalFormatRule[];
    changed: boolean;
  };
  /** Sheet-level validation planned before any cell or merge mutation. */
  validationPlan?: { rules: ValidationRule[]; changed: boolean };
  /** Text-to-Columns disposition by emitted destination-column offset. */
  textToColumnsFormats?: readonly TextToColumnsColumnFormat[];
  /**
   * Paste Special > Skip blanks with a rich paste: a copied cell with no value
   * leaves its destination cell exactly as it is (value, style, runs, link,
   * comment), and the paste is refused if a range facet (merge, conditional
   * format, validation) would change on such a cell.
   */
  skipBlankSources?: boolean;
  /** Typed value dispositions resolved once at the delimited-import boundary. */
  delimitedImportCells?: readonly (readonly DelimitedImportCellValue[])[];
  /**
   * Paste Values: each cell's content with the type the clipboard recorded
   * (writePastedContentInto), so a copied result is stored as the number or
   * text it was instead of being read again from its characters.
   */
  typedContent?: readonly (readonly PastedContent[])[];
  /**
   * Fill Down / Right / Up / Left (Excel "copies the contents and format"):
   * the rich cells replace the destination's value, formats and protection,
   * while the destination's own note and hyperlink stay.
   */
  fillContentAndFormats?: boolean;
  /**
   * A paste whose first row lies directly below a table grows the table to
   * take it (Table AutoExpansion, planTablePasteExpansion).
   */
  expandTables?: boolean;
  /**
   * A Cut -> Paste MOVE: the source rectangle and where it landed. Every
   * formula on the sheet then follows the cells that moved (Excel: =C4*2 reads
   * =G4*2 once C4 has been cut to G4), including a moved cell's own reference
   * to another moved cell. A reference to a cell the move overwrote becomes
   * #REF!, and one that names a cell which did not move is left alone.
   */
  movedFrom?: {
    range: { r1: number; c1: number; r2: number; c2: number };
    rowDelta: number;
    colDelta: number;
    sheetName?: string;
    /**
     * The cut came from ANOTHER sheet: `clearKeys` are the source cells to
     * blank over there, and every sheet's references to the source follow
     * the move to this sheet, in the same commit.
     */
    fromSheet?: {
      id: string;
      name: string;
      clearKeys: readonly string[];
      /** The source sheet's conditional formats and validation without the
       * moved cells (they travelled in the plans above); omitted: as they
       * are. Their references then follow the move like the rest. */
      conditionalFormats?: readonly ConditionalFormatRule[];
      validations?: readonly ValidationRule[];
    };
  };
}

/**
 * A native chart's data rectangle follows a Cut -> Paste that moved the whole
 * of it, as a formula's range reference does (live audit S-033: a chart over
 * N140:O143 kept reading the emptied cells after they moved to R140:S143).
 * `hostSheetId` is the sheet the charts are on (a chart without a source
 * sheet reads its own); a move to another sheet names that sheet as the
 * chart's source. Returns the sheet itself when no chart follows.
 */
export function followMovedChartRanges<S extends object>(
  sheet: S,
  hostSheetId: string,
  move: {
    fromSheetId: string;
    toSheetId: string;
    range: { r1: number; c1: number; r2: number; c2: number };
    rowDelta: number;
    colDelta: number;
  },
): S {
  const charts = (sheet as { charts?: unknown }).charts;
  if (!Array.isArray(charts) || charts.length === 0) return sheet;
  const inside = (rect: unknown): rect is { r1: number; c1: number; r2: number; c2: number } => {
    if (!rect || typeof rect !== "object") return false;
    const { r1, c1, r2, c2 } = rect as Record<string, unknown>;
    return [r1, c1, r2, c2].every((value) => Number.isInteger(value)) &&
      Math.min(r1 as number, r2 as number) >= move.range.r1 &&
      Math.max(r1 as number, r2 as number) <= move.range.r2 &&
      Math.min(c1 as number, c2 as number) >= move.range.c1 &&
      Math.max(c1 as number, c2 as number) <= move.range.c2;
  };
  const shifted = (rect: { r1: number; c1: number; r2: number; c2: number }) => ({
    ...rect,
    r1: rect.r1 + move.rowDelta,
    c1: rect.c1 + move.colDelta,
    r2: rect.r2 + move.rowDelta,
    c2: rect.c2 + move.colDelta,
  });
  let changed = false;
  const next = charts.map((chart) => {
    if (!chart || typeof chart !== "object" || (chart as { type?: unknown }).type === "chart") return chart;
    const record = chart as Record<string, unknown>;
    const source = typeof record.sourceSheetId === "string" && record.sourceSheetId ? record.sourceSheetId : hostSheetId;
    if (source !== move.fromSheetId || !inside(record.range)) return chart;
    const moved: Record<string, unknown> = { ...record, range: shifted(record.range) };
    const across = move.toSheetId !== move.fromSheetId;
    if (across) moved.sourceSheetId = move.toSheetId;
    const binding = record.binding;
    if (binding && typeof binding === "object" && !Array.isArray(binding)) {
      const nextBinding: Record<string, unknown> = { ...(binding as Record<string, unknown>) };
      if (inside(nextBinding.range)) nextBinding.range = shifted(nextBinding.range);
      if (across) nextBinding.sheetId = move.toSheetId;
      moved.binding = nextBinding;
    }
    changed = true;
    return moved;
  });
  return changed ? ({ ...sheet, charts: next } as S) : sheet;
}

/**
 * A PivotTable's source range follows a Cut -> Paste that moved the whole of
 * it, as the chart's does (followMovedChartRanges): the report keeps
 * summarising the data at its new place. A Scenario report has no range.
 */
export function followMovedPivotSources<S extends object>(
  sheet: S,
  hostSheetId: string,
  move: {
    fromSheetId: string;
    fromSheetName: string;
    toSheetId: string;
    range: { r1: number; c1: number; r2: number; c2: number };
    rowDelta: number;
    colDelta: number;
  },
): S {
  const record = sheet as { pivot?: unknown; morePivots?: unknown };
  const across = move.toSheetId !== move.fromSheetId;
  const follow = (pivot: unknown): unknown => {
    if (!pivot || typeof pivot !== "object" || Array.isArray(pivot)) return pivot;
    const config = pivot as Record<string, unknown>;
    if (typeof config.source !== "string" || !config.source || config.cacheSource === "scenario") return pivot;
    const sourceSheet =
      typeof config.sourceSheetId === "string" && config.sourceSheetId ? config.sourceSheetId : hostSheetId;
    if (sourceSheet !== move.fromSheetId) return pivot;
    const source = adjustFormulaForCellMove(config.source, {
      from: move.range,
      rowDelta: move.rowDelta,
      colDelta: move.colDelta,
      sheetName: move.fromSheetName,
    });
    // Landing on the source breaks it on its own sheet, as a formula's
    // reference does; across sheets nothing landed there.
    if (source === config.source || (across && source.includes("#REF!"))) return pivot;
    return { ...config, source, ...(across && !source.includes("#REF!") ? { sourceSheetId: move.toSheetId } : {}) };
  };
  const pivot = follow(record.pivot);
  const morePivots = Array.isArray(record.morePivots) ? record.morePivots.map(follow) : record.morePivots;
  const moreChanged =
    Array.isArray(record.morePivots) &&
    (morePivots as unknown[]).some((entry, index) => entry !== (record.morePivots as unknown[])[index]);
  if (pivot === record.pivot && !moreChanged) return sheet;
  return {
    ...sheet,
    ...(pivot !== record.pivot ? { pivot } : {}),
    ...(moreChanged ? { morePivots } : {}),
  } as S;
}

/**
 * A cut source cell after its content moved away: the value, formula, display
 * cache, inline runs, style, hyperlink extras and threaded comment leave with
 * it; `dropLocked` when the cell's protection travelled with its formats.
 */
export function clearMovedSourceCell(
  existing: SpreadsheetCell,
  dropLocked: boolean,
): SpreadsheetCell {
  const cleared: SpreadsheetCell = { ...existing, value: "", kind: "text" };
  delete cleared.formula;
  // The string type belonged to the value that moved away (B404).
  delete cleared.stringValue;
  delete cleared.picture;
  delete cleared.display;
  delete cleared.chartValueCache;
  delete cleared.runs;
  delete cleared.style;
  delete cleared.hyperlink;
  delete cleared.hyperlinkTooltip;
  delete cleared.hyperlinkDisplay;
  delete cleared.comment;
  if (dropLocked) {
    delete cleared.locked;
    delete cleared.formulaHidden;
  }
  return cleared;
}

export function richPasteReplacementStyle(
  sourceStyle: Record<string, unknown> | undefined,
  destinationStyle: SpreadsheetCell["style"],
  preserveDestinationBorders: boolean,
): SpreadsheetCell["style"] | undefined {
  const replacement = sourceStyle
    ? (JSON.parse(JSON.stringify(sourceStyle)) as Record<string, unknown>)
    : {};
  if (preserveDestinationBorders) {
    const destination = destinationStyle as Record<string, unknown> | undefined;
    for (const key of RICH_PASTE_BORDER_STYLE_KEYS) {
      // Source borders are outside this command's transfer mask. Delete them
      // first, then restore only the destination's own exact border authority.
      delete replacement[key];
      if (
        destination &&
        Object.prototype.hasOwnProperty.call(destination, key)
      ) {
        const value = destination[key];
        replacement[key] =
          value && typeof value === "object"
            ? JSON.parse(JSON.stringify(value))
            : value;
      }
    }
  }
  return Object.keys(replacement).length > 0
    ? (replacement as SpreadsheetCell["style"])
    : undefined;
}

/**
 * Write one pasted cell whose type the clipboard already knows (Paste Values,
 * Paste Formulas): the characters are stored as they are and keep the type
 * they had, instead of being read again from their characters. A formula
 * stays a formula; a literal is never one, even when it reads "=A2"; a text
 * that would read back as a number, logical or error keeps the cell's quote
 * prefix, and any other content drops it (a new value replaces the old one's
 * quote prefix, as a new entry does).
 */
export function writePastedContentInto(
  nextCells: Record<string, SpreadsheetCell>,
  row: number,
  col: number,
  content: PastedContent,
): void {
  const key = `${row}:${col}`;
  writeCellValueInto(nextCells, row, col, content.raw);
  const written = nextCells[key];
  if (!written) return;
  const next: SpreadsheetCell = { ...written };
  if (!content.formula) {
    delete next.formula;
    next.kind = content.kind;
  }
  const style = { ...(next.style ?? {}) };
  if (content.forceText) style.quotePrefix = true;
  else delete style.quotePrefix;
  if (Object.keys(style).length > 0) next.style = style;
  else delete next.style;
  // A copied picture placed in a cell pastes with its value (Excel's rich
  // value travels with Paste Values).
  const picture = content.picture ? normalizeCellPicture(content.picture) : null;
  const landed =
    picture && next.value === content.raw ? placePictureInCell(next, picture) : next;
  if (landed.value !== "" || hasCellMetadata(landed)) nextCells[key] = landed;
  else delete nextCells[key];
}

/**
 * Table AutoExpansion for a paste: a block pasted into the row directly below
 * a table without a Totals Row, starting within its columns, grows the table
 * down over every pasted row through the Resize Table transaction typing uses
 * (planTableAutoExpansion), so a table the resize refuses does not grow.
 * Calculated columns fill only the new rows' empty cells, so pasted values
 * stay, as exceptions. Null when no table grows.
 */
function tablePasteExpansionCandidate(
  sheet: TableButtonSheet,
  pasted: Pick<Rect, "r1" | "c1">,
): DataTable | undefined {
  return (sheet.dataTables ?? []).find(
    ({ range, totalsRow }) =>
      !totalsRow &&
      pasted.r1 === range.r2 + 1 &&
      pasted.c1 >= range.c1 &&
      pasted.c1 <= range.c2,
  );
}

export function planTablePasteExpansion<S extends TableButtonSheet>(
  sheet: S,
  pasted: Rect,
  rangeHasLockedCell: (sheet: S, range: Rect) => boolean,
): { sheet: S; tableId: string } | null {
  const table = tablePasteExpansionCandidate(sheet, pasted);
  if (!table) return null;
  const grown = planDataTableResize(
    sheet,
    table.id,
    { ...table.range, r2: pasted.r2 },
    rangeHasLockedCell,
  );
  return grown && grown !== sheet ? { sheet: grown, tableId: table.id } : null;
}

/**
 * What a Cut -> Paste on one sheet does to the sheet's rules, planned before
 * the paste: the moved cells' conditional formats move with them (a rule
 * reaching past them is split, live audit S-029) and their data validation
 * leaves the source and covers the destination (live audit B261); every
 * other rule's references to the moved cells follow them, a dependent
 * list's controlling cell included. A plan the planners cannot make is
 * left out, and the rules stay as they were. `movedValidations` is the
 * validation snapshot the Cut took.
 */
export function sameSheetCutRulePlans(args: {
  conditionalFormats: readonly ConditionalFormatRule[];
  validations: readonly ValidationRule[];
  movedValidations: Parameters<typeof planRulesForValidationPaste>[1] | null;
  movedRange: Rect;
  row: number;
  col: number;
  sheetName: string;
  bounds: { maxRow: number; maxCol: number };
}): {
  validationPlan?: { rules: ValidationRule[]; changed: boolean };
  conditionalFormatPlan?: { rules: ConditionalFormatRule[]; changed: boolean };
} {
  const { conditionalFormats, validations, movedValidations, movedRange, row, col, bounds } = args;
  const operation = {
    from: movedRange,
    rowDelta: row - movedRange.r1,
    colDelta: col - movedRange.c1,
    sheetName: args.sheetName,
  };
  const followMove = (formula: string) => adjustFormulaForCellMove(formula, operation);
  const destinationRange = {
    r1: row,
    c1: col,
    r2: row + movedRange.r2 - movedRange.r1,
    c2: col + movedRange.c2 - movedRange.c1,
  };
  const out: ReturnType<typeof sameSheetCutRulePlans> = {};
  const conditionalPlan = planRulesForCellMove(
    [...conditionalFormats],
    movedRange,
    operation.rowDelta,
    operation.colDelta,
    (rule) => rewriteConditionalFormatReferences(rule, followMove),
    bounds,
  );
  if (conditionalPlan) out.conditionalFormatPlan = conditionalPlan;
  // The rules left behind keep reading what they read: the source and the
  // destination are taken out with their remainders re-anchored, then what
  // names the moved cells follows them.
  const withoutSource = clearValidationRulesInRange([...validations], movedRange, shiftFormula);
  const retained = withoutSource
    ? clearValidationRulesInRange(withoutSource, destinationRange, shiftFormula)
    : null;
  const followed = retained?.map((rule) => {
    const next = rewriteValidationReferences(rule, followMove);
    if (
      next.kind !== "dependent" ||
      next.source.row < movedRange.r1 ||
      next.source.row > movedRange.r2 ||
      next.source.col < movedRange.c1 ||
      next.source.col > movedRange.c2
    ) {
      return next;
    }
    return {
      ...next,
      source: { row: next.source.row + operation.rowDelta, col: next.source.col + operation.colDelta },
    };
  });
  const plan = followed && movedValidations
    ? planRulesForValidationPaste(followed, movedValidations, movedRange, destinationRange, followMove, bounds)
    : null;
  if (plan) {
    out.validationPlan = {
      rules: plan.rules,
      changed: JSON.stringify(plan.rules) !== JSON.stringify(validations),
    };
  }
  return out;
}

/** One bulk write: values from a top-left cell, and what travels with them. */
export interface PasteGridRequest {
  baseRow: number;
  baseCol: number;
  rows: string[][];
  /**
   * Cut→paste MOVE (Excel parity): source cell keys ("r:c") to clear in the
   * SAME commit AFTER the paste writes land, so the move is one undo step.
   * The clear blanks the VALUE but keeps formatting (Clear Contents), via the
   * same hasCellMetadata prune the single-cell write uses. Keys already inside
   * the pasted destination are excluded by the caller (cutSourceClearKeys), so
   * a moved value is never erased.
   */
  clearSourceKeys?: readonly string[];
  /**
   * Standard copy/paste carries the detached positional rich-cell grid.
   * When present it REPLACES the destination's transferred metadata:
   * style, inline runs, hyperlink extras and threaded comment. Other bulk
   * writers omit this argument and retain destination metadata as before.
   */
  clipboardCells?: RichClipboard;
  /**
   * Foreign Office HTML rowspan/colspan metadata, relative to the pasted
   * grid anchor. When present, source merges become authoritative model
   * spans instead of being silently flattened to covered blank cells.
   */
  clipboardMerges?: readonly MergeSpan[];
  options?: PasteGridOptions;
}







function clipboardCellsCoveredInTheAnchorsFormat(
  cells: RichClipboard | undefined,
  merges: readonly MergeSpan[] | undefined,
): RichClipboard | undefined {
  if (!cells || !merges?.length) return cells;
  let next = cells;
  for (const span of merges) {
    if (span.coveredStyles) {
      for (const [key, own] of Object.entries(span.coveredStyles)) {
        const [r, c] = key.split(":").map(Number);
        if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= span.rowSpan || c >= span.colSpan) continue;
        if (r === 0 && c === 0) continue;
        const row = span.row + r;
        const col = span.col + c;
        const slot = next[row]?.[col];
        if (row >= next.length || styleStatesAFormat(slot?.style)) continue;
        if (next === cells) next = [...cells];
        const copied = [...(next[row] ?? [])];
        copied[col] = { ...(slot ?? { raw: "", computed: "" }), style: structuredClone(own) as NonNullable<typeof slot>["style"] };
        next[row] = copied;
      }
      continue;
    }
    const style = cells[span.row]?.[span.col]?.style;
    if (!style) continue;
    const format = Object.fromEntries(
      Object.entries(style).filter(([key]) => !key.startsWith("border")),
    );
    for (let r = 0; r < span.rowSpan; r += 1) {
      for (let c = 0; c < span.colSpan; c += 1) {
        if (r === 0 && c === 0) continue;
        const row = span.row + r;
        const col = span.col + c;
        const slot = next[row]?.[col];
        if (row >= next.length || styleStatesAFormat(slot?.style)) continue;
        if (next === cells) next = [...cells];
        const copied = [...(next[row] ?? [])];
        copied[col] = { ...(slot ?? { raw: "", computed: "" }), style: format };
        next[row] = copied;
      }
    }
  }
  return next;
}

/** What the paste reads of the sheet beyond its cells. */
export interface PasteGridContext<S> {
  /** Sheet protection: is a cell of `range` locked (pause-aware in the window). */
  rangeHasLockedCell: (sheet: S, range: Rect) => boolean;
  /** The sheet's merges, conditional formats and validation rules as the
   *  window holds them (normalized); read from the sheet when absent. */
  merges?: readonly MergeSpan[];
  conditionalFormats?: readonly ConditionalFormatRule[];
  validations?: readonly ValidationRule[];
}

/** Why a paste wrote nothing: a locked cell, a merge it would cut through,
 *  a Skip blanks facet it would change, (skipIfUnchanged) no change, or part
 *  of a What-If Data Table's results ("Cannot change part of a data table."). */
export type PasteGridRefusal = "locked" | "merge" | "skip-blanks" | "unchanged" | "what-if-table";

export type PasteGridOutcome<S> =
  | { ok: false; reason: PasteGridRefusal }
  | {
      ok: true;
      /** The pasted sheet, before any Table AutoExpansion. */
      sheet: S;
      /** The sheet with the table grown over the paste, when one grows. */
      expansion: { sheet: S; tableId: string } | null;
      /** The rest of the workbook after a Cut's move (formulas follow it). */
      followOtherSheets?: (candidate: S) => S;
      /** The rectangle the values landed in. */
      destination: Rect;
    };

/**
 * The sheet's sparklines after a paste lands the copied ones (null when
 * nothing changes). A Cut's move takes each one off its source cell and
 * re-points its data range as the moved formulas are (`moveRange`); a copy
 * already carries its range shifted. Copies past the sheet's sparkline cap
 * are not added.
 */
function pastedSparklines(
  current: Record<string, Sparkline> | undefined,
  writes: ReadonlyArray<{ key: string; sourceKey: string; value: Sparkline }>,
  moveRange: ((range: string) => string) | null,
): Record<string, Sparkline> | null {
  if (writes.length === 0) return null;
  const next: Record<string, Sparkline> = { ...(current ?? {}) };
  if (moveRange) {
    for (const write of writes) delete next[write.sourceKey];
  }
  for (const write of writes) {
    if (!(write.key in next) && Object.keys(next).length >= MAX_SPARKLINES) continue;
    const range =
      moveRange && write.value.range !== "" ? moveRange(write.value.range) : write.value.range;
    if (range.includes("#REF!")) continue;
    next[write.key] = { ...write.value, range };
  }
  return JSON.stringify(next) === JSON.stringify(current ?? {}) ? null : next;
}

/** Whether a position with no cell shows a line format other than the
 *  default (its row's, else its column's; {@link lineBlankAt}). */
function lineFormatDiffersFromDefault(sheet: object, row: number, col: number): boolean {
  const { rowCount, columnCount, columnStyles } = sheet as { rowCount?: number; columnCount?: number; columnStyles?: Record<string, object> };
  if (row >= (rowCount ?? 0) && col < (columnCount ?? 0)) {
    // Below the sheet's last row a column's format still stands (Excel formats
    // the whole column): a paste that grows the sheet lands on it (xlsx-rounds
    // R289-X12 E14: default blanks pasted at row 82 of a 78-row sheet showed,
    // and saved, column B's 10 pt wrapped top-left).
    const column = columnStyles?.[String(col)];
    return Boolean(column && Object.keys(column).length > 0);
  }
  const line = lineBlankAt(sheet, row, col);
  return Boolean(line?.style && Object.keys(line.style).length > 0);
}

/**
 * The paste as one sheet transaction. Refuses the whole write (Excel never
 * writes part of a paste) over a locked cell, across a merge the destination
 * cuts through (authoritative merges), or when Skip blanks would change a
 * kept cell's merge, conditional format or validation.
 */
export function pastedGridSheet<S extends PasteButtonSheet>(
  sheet: S,
  request: PasteGridRequest,
  context: PasteGridContext<S>,
): PasteGridOutcome<S> {
  const { baseRow, baseCol, rows, clearSourceKeys, clipboardMerges } = request;
  const clipboardCells = clipboardCellsCoveredInTheAnchorsFormat(request.clipboardCells, clipboardMerges);
  const options = request.options ?? {};
  const mergeSpans = context.merges ?? normalizeMerges(sheet.merges);
  const conditionalFormats = context.conditionalFormats ?? sheet.conditionalFormats ?? [];
  const validations = context.validations ?? sheet.validations ?? [];
  // Sheet protection (Excel/WPS): block a multi-cell paste/fill/import that
  // would land on ANY locked cell (Excel refuses the whole paste, never a
  // partial write). Shared by paste, CSV import, fill-down and paste-special.
  const maxR = baseRow + Math.max(0, rows.length - 1);
  const maxC =
    baseCol + Math.max(0, ...rows.map((cols) => cols.length - 1), 0);
  if (
    context.rangeHasLockedCell(sheet, {
      r1: baseRow,
      r2: maxR,
      c1: baseCol,
      c2: maxC,
    })
  ) {
    return { ok: false, reason: "locked" };
  }
  const cutIn = options.refuseMergesCutIn;
  if (
    cutIn &&
    mergeSpans.some(
      (span) =>
        span.row <= cutIn.r2 && span.row + span.rowSpan - 1 >= cutIn.r1 &&
        span.col <= cutIn.c2 && span.col + span.colSpan - 1 >= cutIn.c1 &&
        !(span.row >= cutIn.r1 && span.row + span.rowSpan - 1 <= cutIn.r2 &&
          span.col >= cutIn.c1 && span.col + span.colSpan - 1 <= cutIn.c2),
    )
  ) {
    return { ok: false, reason: "merge" };
  }
  if (options.refuseValuesUnderMerges && mergeSpans.length > 0) {
    const covered = coveredCells(mergeSpans);
    for (let r = 0; r < rows.length; r += 1) {
      for (let c = 0; c < rows[r].length; c += 1) {
        if (rows[r][c] !== "" && rows[r][c] !== undefined && covered.has(`${baseRow + r}:${baseCol + c}`)) {
          return { ok: false, reason: "merge" };
        }
      }
    }
  }
  // A What-If Data Table's results change only whole (Excel: "Cannot change
  // part of a data table."): a paste over all of them, or a Cut of all of
  // them, replaces the table's values and ends the table.
  const destinationRect = { r1: baseRow, c1: baseCol, r2: maxR, c2: maxC };
  writeLineBlanksIn(sheet as unknown as Parameters<typeof writeLineBlanksIn>[0], destinationRect);
  const movedSourceRect = options.movedFrom && !options.movedFrom.fromSheet ? options.movedFrom.range : null;
  if (
    whatIfTablePartRefused(sheet.whatIfTables, destinationRect) ||
    (movedSourceRect && whatIfTablePartRefused(sheet.whatIfTables, movedSourceRect))
  ) {
    return { ok: false, reason: "what-if-table" };
  }
  const whatIfTablesAfter = whatIfTablesAfterOverwrite(
    sheet.whatIfTables,
    movedSourceRect ? [destinationRect, movedSourceRect] : [destinationRect],
  );
  let nextMerges: readonly MergeSpan[] = mergeSpans;
  // A Cut -> Paste MOVES its cells, so a merged cell the destination
  // would cut through refuses the whole move, as a Copy does (live audit
  // B137).
  const appliesAuthoritativeMerges = Boolean(
    clipboardMerges &&
    (clipboardMerges.length > 0 || options.authoritativeMerges || options.movedFrom !== undefined),
  );
  if (appliesAuthoritativeMerges && clipboardMerges) {
    const destinationSpan: MergeSpan = {
      row: baseRow,
      col: baseCol,
      rowSpan: maxR - baseRow + 1,
      colSpan: maxC - baseCol + 1,
    };
    // Excel/WPS reject changing only part of an existing merged cell. Keep
    // this paste atomic: a merge crossing the destination boundary blocks
    // all values, styles and merge writes instead of tearing the old span.
    const crossesExistingMerge = mergeSpans.some(
      (existing) =>
        spansOverlap(existing, destinationSpan) &&
        !(
          existing.row >= destinationSpan.row &&
          existing.col >= destinationSpan.col &&
          existing.row + existing.rowSpan <=
            destinationSpan.row + destinationSpan.rowSpan &&
          existing.col + existing.colSpan <=
            destinationSpan.col + destinationSpan.colSpan
        ),
    );
    if (crossesExistingMerge) {
      return { ok: false, reason: "merge" };
    }
    // A move takes its merged cells along: the source is left unmerged, as
    // Excel leaves it (measured: cutting merged J10:K11 to N10 unmerges J10;
    // audit P-CUT-OVERWRITE-TARGET-REF).
    const mergesLeft = movedSourceRect
      ? mergeSpans.filter(
          (span) =>
            !(
              span.row >= movedSourceRect.r1 &&
              span.col >= movedSourceRect.c1 &&
              span.row + span.rowSpan - 1 <= movedSourceRect.r2 &&
              span.col + span.colSpan - 1 <= movedSourceRect.c2
            ),
        )
      : mergeSpans;
    nextMerges = removeMergesInRange([...mergesLeft], {
      r1: baseRow,
      r2: maxR,
      c1: baseCol,
      c2: maxC,
    });
    for (const relative of clipboardMerges) {
      const translated: MergeSpan = {
        row: baseRow + relative.row,
        col: baseCol + relative.col,
        rowSpan: relative.rowSpan,
        colSpan: relative.colSpan,
      };
      if (
        translated.row < baseRow ||
        translated.col < baseCol ||
        translated.row + translated.rowSpan - 1 > maxR ||
        translated.col + translated.colSpan - 1 > maxC
      ) {
        continue;
      }
      // A table holds no merged cell (Merge is greyed in one, user audit
      // S-019): a merged cell pasted into a table lands as its cells. A
      // merge pasted over a table's header row covered the header cells
      // the table's column names are read from (sheet model fuzz).
      if (mergeRectTouchesTable(sheet, {
        r1: translated.row,
        c1: translated.col,
        r2: translated.row + translated.rowSpan - 1,
        c2: translated.col + translated.colSpan - 1,
      })) {
        continue;
      }
      nextMerges = addMerge([...nextMerges], translated);
    }
  }
  // Paste Special > Skip blanks (Excel): wherever the copied cell holds
  // no value, the destination cell keeps everything it has. Its cell
  // facets are simply not written below. Merges, conditional formats and
  // validation are planned for the whole rectangle, so if one of them
  // would change on such a cell the paste is refused whole rather than
  // overwriting what Skip blanks promises to keep.
  const skippedBlankKeys = new Set<string>();
  if (options.skipBlankSources && clipboardCells) {
    rows.forEach((cols, r) => {
      cols.forEach((_value, c) => {
        if (richClipboardCellIsBlank(clipboardCells[r]?.[c])) {
          skippedBlankKeys.add(`${baseRow + r}:${baseCol + c}`);
        }
      });
    });
  }
  if (skippedBlankKeys.size > 0) {
    const destination = { r1: baseRow, c1: baseCol, r2: maxR, c2: maxC };
    const conditionalBefore = options.conditionalFormatPlan
      ? conditionalFormats.filter((rule) =>
          conditionalRuleOverlaps(rule, destination),
        )
      : [];
    const conditionalAfter = options.conditionalFormatPlan
      ? options.conditionalFormatPlan.rules.filter((rule) =>
          conditionalRuleOverlaps(rule, destination),
        )
      : [];
    const inDestination = (range: Rect) =>
      range.r1 <= maxR &&
      range.r2 >= baseRow &&
      range.c1 <= maxC &&
      range.c2 >= baseCol;
    const validationBefore = options.validationPlan
      ? validations.filter((rule) => inDestination(rule.range))
      : [];
    const validationAfter = options.validationPlan
      ? options.validationPlan.rules.filter((rule) =>
          inDestination(rule.range),
        )
      : [];
    const facetAt = (
      merges: readonly MergeSpan[],
      conditional: readonly ConditionalFormatRule[],
      validation: readonly ValidationRule[],
      row: number,
      col: number,
    ) => {
      const cell = { r1: row, c1: col, r2: row, c2: col };
      return JSON.stringify([
        merges.find((span) =>
          spansOverlap(span, { row, col, rowSpan: 1, colSpan: 1 }),
        ) ?? null,
        conditional.filter((rule) => conditionalRuleOverlaps(rule, cell)),
        ruleForCell(validation, row, col),
      ]);
    };
    for (const key of skippedBlankKeys) {
      const [row, col] = key.split(":").map(Number);
      if (
        facetAt(mergeSpans, conditionalBefore, validationBefore, row, col) !==
        facetAt(nextMerges, conditionalAfter, validationAfter, row, col)
      ) {
        return { ok: false, reason: "skip-blanks" };
      }
    }
  }
  // Clone the cells map ONCE and apply every write into it (vs. one fresh
  // sheet+cells object per cell — 200k allocations for a 1000×200 paste).
  const nextCells: CellMap = copyCellMap(sheet.cells);
  let maxRow = baseRow;
  let maxCol = baseCol;
  let cellsChanged = false;
  // The sparklines the copied cells bring (Excel copies a sparkline with its
  // cell). Written after the cells, onto the sheet they were copied from: a
  // destination the copy has no sparkline for keeps its own, as Excel's
  // ordinary paste leaves a sparkline it does not replace.
  const sparklineWrites: Array<{ key: string; sourceKey: string; value: Sparkline }> = [];
  // A copied table cell's source formatting depends on where it lands
  // (richClipboardCellSourceStyle): a cell of a table - or of the row a
  // paste below a table grows it over - shows that table's own style.
  const growingTable =
    options.expandTables && spreadsheetTableAutoCorrectSettings().expandTables
      ? tablePasteExpansionCandidate(sheet, { r1: baseRow, c1: baseCol })
      : undefined;
  // A Cut that moves a whole table on this sheet moves the table (Excel 16:
  // Table1 on A1:B3 cut to E1 is Table1 on E1:F3; one row cut out of a table
  // leaves the table where it is). Its cells then keep their own formats, the
  // table's style showing under them.
  const movedTableIds = new Set<string>();
  const tablesAfterMove =
    options.movedFrom && !options.movedFrom.fromSheet
      ? (sheet.dataTables ?? []).map((table) => {
          const moved = options.movedFrom!;
          const { r1, c1, r2, c2 } = table.range;
          if (
            table.nativeResizeBlocker ||
            Object.keys(table.columnFilters ?? {}).length > 0 ||
            r1 < moved.range.r1 ||
            c1 < moved.range.c1 ||
            r2 > moved.range.r2 ||
            c2 > moved.range.c2
          ) {
            return table;
          }
          movedTableIds.add(table.id);
          return {
            ...table,
            range: { r1: r1 + moved.rowDelta, c1: c1 + moved.colDelta, r2: r2 + moved.rowDelta, c2: c2 + moved.colDelta },
          };
        })
      : sheet.dataTables ?? [];
  const destinationInTable = (row: number, col: number): boolean =>
    tablesAfterMove.some((table) => tableCellRole(table, row, col) !== null) ||
    (growingTable !== undefined &&
      row > growingTable.range.r2 &&
      col >= growingTable.range.c1 &&
      col <= growingTable.range.c2);
  rows.forEach((cols, r) => {
    cols.forEach((val, c) => {
      const tr = baseRow + r;
      const tc = baseCol + c;
      const key = `${tr}:${tc}`;
      if (skippedBlankKeys.has(key)) return;
      const previousJson = options.skipIfUnchanged
        ? JSON.stringify(sheet.cells[key] ?? null)
        : null;
      const typedContent = options.typedContent?.[r]?.[c];
      if (typedContent) {
        writePastedContentInto(nextCells, tr, tc, typedContent);
      } else if (options.parseEntries) {
        writeSpreadsheetEntryInto(nextCells, tr, tc, val, options.parseEntries.locale, undefined, options.parseEntries.date1904, options.parseEntries.typedNumerals === true);
      } else writeCellValueInto(nextCells, tr, tc, val);
      const delimitedImportCell = options.delimitedImportCells?.[r]?.[c];
      const importedWrite = nextCells[key];
      if (delimitedImportCell && importedWrite) {
        // The cell's text identity is this conversion's, not the target's
        // old entry's: a Text column keeps its field text, a General
        // column's TRUE is the logical value wherever it lands (live
        // audit B273).
        const importedStyle = styleWithOwnTextIdentity(
          importedWrite.style,
          delimitedImportCell.literalText ? { quotePrefix: true } : undefined,
        );
        const typedImportWrite: SpreadsheetCell = {
          ...importedWrite,
          value: delimitedImportCell.value,
          kind: delimitedImportCell.kind,
          ...(delimitedImportCell.numberFormat
            ? {
                style: {
                  ...(importedStyle ?? {}),
                  numberFormat: delimitedImportCell.numberFormat,
                },
              }
            : importedStyle && Object.keys(importedStyle).length > 0
              ? { style: importedStyle }
              : {}),
        };
        if (!delimitedImportCell.numberFormat && !(importedStyle && Object.keys(importedStyle).length > 0)) {
          delete typedImportWrite.style;
        }
        if (delimitedImportCell.literalText) {
          delete typedImportWrite.formula;
        }
        delete typedImportWrite.display;
        delete typedImportWrite.chartValueCache;
        // The old entry's runs belong to its text: a field converted to
        // something else (Text to Columns of "2021\n" is the number 2021)
        // is a plain cell, as Excel writes it and the file reads back
        // (xlsx-rounds R323-X12 WALK1: a number cell still holding text runs).
        if (
          typedImportWrite.runs &&
          (typedImportWrite.kind !== "text" || textFromRuns(typedImportWrite.runs) !== typedImportWrite.value)
        ) {
          delete typedImportWrite.runs;
        }
        nextCells[key] = withRecordFontIfPlain(typedImportWrite);
      }
      const textToColumnsFormat = options.textToColumnsFormats?.[c];
      if (textToColumnsFormat) {
        const written = nextCells[key] ?? {
          value: "",
          kind: "text" as const,
        };
        const nextStyle = { ...(written.style ?? {}) };
        if (textToColumnsFormat === "text") {
          nextStyle.numberFormat = "@";
          const textCell: SpreadsheetCell = {
            ...written,
            value: val,
            kind: "text",
            style: nextStyle,
          };
          delete textCell.formula;
          delete textCell.display;
          delete textCell.chartValueCache;
          delete textCell.runs;
          nextCells[key] = textCell;
        } else {
          // Excel's General column "converts date values to dates", and a
          // converted date is shown as a date, exactly as the Date column
          // shows it. Both dispositions keep the date format the shared
          // import classifier resolved (the one Data > From Text/CSV
          // writes); any other value resets the cell to General.
          if (delimitedImportCell?.numberFormat === "yyyy-mm-dd") {
            nextStyle.numberFormat = delimitedImportCell.numberFormat;
          } else {
            delete nextStyle.numberFormat;
          }
          const generalCell: SpreadsheetCell = {
            ...written,
            ...(Object.keys(nextStyle).length > 0
              ? { style: nextStyle }
              : {}),
          };
          if (Object.keys(nextStyle).length === 0) {
            delete generalCell.style;
          }
          delete generalCell.display;
          if (generalCell.value !== "" || hasCellMetadata(generalCell)) {
            nextCells[key] = generalCell;
          } else {
            delete nextCells[key];
          }
        }
      }
      if (clipboardCells) {
        const source = clipboardCells[r]?.[c];
        if (source?.sparkline && source.sparkline.sheetId === sheet.id) {
          sparklineWrites.push({
            key,
            sourceKey: options.movedFrom
              ? `${options.movedFrom.range.r1 + r}:${options.movedFrom.range.c1 + c}`
              : key,
            value: source.sparkline.value,
          });
        }
        const written = nextCells[key] ?? {
          value: val,
          kind: cellKind(val),
        };
        const replaced: SpreadsheetCell = { ...written };
        const keepNotesAndLinks = options.fillContentAndFormats === true;

        delete replaced.style;
        delete replaced.runs;
        delete replaced.phonetic;
        if (!keepNotesAndLinks) {
          delete replaced.hyperlink;
          delete replaced.hyperlinkTooltip;
          delete replaced.hyperlinkDisplay;
          delete replaced.comment;
        }
        // The copied cell's identity, not its characters: a literal
        // (a quote-prefixed "=A2" included) is never made a formula, and
        // keeps the stored type it had.
        if (source?.formula === false) {
          delete replaced.formula;
          replaced.kind = source.kind ?? cellKind(val);
          // ... and a string the workbook typed as one ("FALSE") stays that
          // string, not the logical its characters spell (live audit B404).
          if (source.stringValue === true && val === source.raw && replaced.kind !== "number") {
            replaced.stringValue = true;
          } else {
            delete replaced.stringValue;
          }
        }
        // Protection is part of the copied formats (Format Cells >
        // Protection); a foreign snapshot that does not say leaves the
        // destination's lock alone.
        if (source?.locked !== undefined) {
          delete replaced.locked;
          if (source.locked === false) replaced.locked = false;
          // Hidden is the other half of the same Protection format.
          delete replaced.formulaHidden;
          if (source.formulaHidden === true) replaced.formulaHidden = true;
        }

        const replacementStyle = richPasteReplacementStyle(
          source?.styleOutsideTable
            ? richClipboardCellSourceStyle(source, destinationInTable(tr, tc))
            : source?.style,
          written.style,
          options.preserveDestinationBorders === true,
        );
        if (replacementStyle) {
          replaced.style = replacementStyle;
        }
        const copiedRuns = source?.runs
          ? normalizeCellRuns(
              source.runs.map((run) => ({
                text: run.text,
                ...(run.style ? { style: { ...run.style } } : {}),
              })),
              val,
            )
          : undefined;
        if (copiedRuns) replaced.runs = copiedRuns;
        // The copied text brings its furigana and Show Phonetic Field.
        const copiedPhonetic = source?.phonetic
          ? phoneticAfterValueWrite(clonePhonetic(source.phonetic), source.raw, val, {
              editedInPlace: true,
              textual: !replaced.formula && replaced.kind !== "number",
            })
          : undefined;
        if (copiedPhonetic) replaced.phonetic = copiedPhonetic;
        if (source?.hyperlink && !keepNotesAndLinks) {
          replaced.hyperlink = source.hyperlink;
          if (source.hyperlinkTooltip) {
            replaced.hyperlinkTooltip = source.hyperlinkTooltip;
          }
          if (source.hyperlinkDisplay) {
            replaced.hyperlinkDisplay = source.hyperlinkDisplay;
          }
        }
        if (source?.comment && !keepNotesAndLinks) {
          replaced.comment = cloneCellCommentForPaste(source.comment);
        }
        // A picture placed in the copied cell is its value: it lands in
        // the destination with the value it stands for (Place in Cell).
        const copiedPicture = source?.picture
          ? normalizeCellPicture(source.picture)
          : null;
        const landed =
          copiedPicture && replaced.value === source?.raw
            ? placePictureInCell(replaced, copiedPicture)
            : replaced;

        // An empty string is a value (ISBLANK false, COUNTA 1), and a copy of
        
        if (landed.value !== "" || landed.stringValue === true || hasCellMetadata(landed)) {
          nextCells[key] = landed;
        } else if (lineFormatDiffersFromDefault(sheet, tr, tc)) {
          // A copied blank of the default format over a position its row's or
          // column's format describes: Excel's paste writes the copied cell's
          // format there, General included, and the position stops showing
          
          // pasted into a date column showed, and saved, the dates' format).
          // Kept as the explicit default blank; a position whose line format
          // is already the default stays no cell.
          nextCells[key] = { value: "", kind: "text", style: {} };
        } else {
          delete nextCells[key];
        }
      }
      if (
        options.skipIfUnchanged &&
        JSON.stringify(nextCells[key] ?? null) !== previousJson
      ) {
        cellsChanged = true;
      }
      if (tr > maxRow) maxRow = tr;
      if (tc > maxCol) maxCol = tc;
    });
  });
  // Clear the cut source (move) AFTER the paste writes, so a destination that
  // overlaps the source keeps the pasted value (the caller already excluded
  // overlapping keys). The value, formula, display cache, inline runs and
  // style, hyperlink extras and threaded comment move away together.
  if (clearSourceKeys && clearSourceKeys.length > 0) {
    for (const key of clearSourceKeys) {
      const existing = nextCells[key];
      if (!existing) continue;
      const previousJson = options.skipIfUnchanged
        ? JSON.stringify(existing)
        : null;
      // The cell's protection moved with its formats.
      const cleared = clearMovedSourceCell(existing, clipboardCells !== undefined);
      if (hasCellMetadata(cleared)) nextCells[key] = cleared;
      else delete nextCells[key];
      if (
        options.skipIfUnchanged &&
        JSON.stringify(nextCells[key] ?? null) !== previousJson
      ) {
        cellsChanged = true;
      }
    }
  }
  // Formulas follow the cells a Cut -> Paste moved. This runs after the
  // paste writes and the source clear, so it sees each cell's final text,
  // and it walks only the sheet's formula cells.
  const crossSheetMove = options.movedFrom?.fromSheet;
  if (options.movedFrom) {
    const moved = options.movedFrom;
    const operation = {
      from: moved.range,
      rowDelta: moved.rowDelta,
      colDelta: moved.colDelta,
      sheetName: moved.sheetName,
    };
    const toSheetName = sheet.name || sheet.id;
    for (const key of Object.keys(nextCells)) {
      const cell = nextCells[key];
      const formula = cell?.formula;
      if (typeof formula !== "string" || !formula.startsWith("=")) continue;
      let next: string;
      if (crossSheetMove) {
        // A formula inside the pasted rectangle came over with the move:
        // its unqualified references still name the source sheet's cells.
        const [rowText, colText] = key.split(":");
        const row = Number(rowText);
        const col = Number(colText);
        next = adjustFormulaForCrossSheetMove(formula, {
          from: moved.range,
          rowDelta: moved.rowDelta,
          colDelta: moved.colDelta,
          fromSheet: crossSheetMove.name,
          toSheet: toSheetName,
          hostSheet: toSheetName,
          movedFormula:
            row >= baseRow && row <= maxR && col >= baseCol && col <= maxC,
        });
      } else {
        next = adjustFormulaForCellMove(formula, operation);
      }
      if (next === formula) continue;
      nextCells[key] = { ...cell, formula: next, value: next };
      cellsChanged = true;
    }
  }
  // A move on this sheet is followed by everything else that names the
  // moved cells: the sheet's own defined names, validation sources,
  // table templates, conditional formats, sparklines and charts (its
  // cells were re-pointed above), and the qualified references on the
  // other sheets (live audit B269).
  const sameSheetMove = options.movedFrom && !crossSheetMove ? options.movedFrom : null;
  const sameSheetOperation = sameSheetMove
    ? {
        from: sameSheetMove.range,
        rowDelta: sameSheetMove.rowDelta,
        colDelta: sameSheetMove.colDelta,
        sheetName: sameSheetMove.sheetName,
      }
    : null;
  const chartMove = options.movedFrom
    ? {
        fromSheetId: crossSheetMove?.id ?? sheet.id,
        fromSheetName: crossSheetMove?.name ?? (options.movedFrom.sheetName || sheet.name || sheet.id),
        toSheetId: sheet.id,
        range: options.movedFrom.range,
        rowDelta: options.movedFrom.rowDelta,
        colDelta: options.movedFrom.colDelta,
      }
    : null;
  // Charts and PivotTables over the moved block read it at its new place.
  const followMovedSources = (candidate: S, hostSheetId: string): S =>
    chartMove
      ? followMovedPivotSources(followMovedChartRanges(candidate, hostSheetId, chartMove), hostSheetId, chartMove)
      : candidate;
  const movedCarriers = sameSheetOperation
    ? rewriteSheetFormulaReferences({ ...sheet, cells: {} }, (formula) =>
        adjustFormulaForCellMove(formula, sameSheetOperation),
      )
    : crossSheetMove && options.movedFrom
      ? // This sheet's own names, rules, tables, sparklines and charts that
        // named the source cells on the other sheet follow them here too.
        rewriteSheetFormulaReferences({ ...sheet, cells: {} }, (formula) =>
          adjustFormulaForCrossSheetMove(formula, {
            from: options.movedFrom!.range,
            rowDelta: options.movedFrom!.rowDelta,
            colDelta: options.movedFrom!.colDelta,
            fromSheet: crossSheetMove.name,
            toSheet: sheet.name || sheet.id,
            hostSheet: sheet.name || sheet.id,
          }),
        )
      : null;
  const carriers = chartMove
    ? followMovedSources(movedCarriers ?? sheet, sheet.id)
    : movedCarriers;
  const nextRowCount = Math.max(sheet.rowCount, maxRow + 1);
  const nextColumnCount = Math.max(sheet.columnCount, maxCol + 1);
  const mergesChanged =
    appliesAuthoritativeMerges &&
    JSON.stringify(nextMerges) !== JSON.stringify(mergeSpans);
  const carriedSparklines = pastedSparklines(
    (movedCarriers ?? sheet).sparklines,
    sparklineWrites,
    sameSheetOperation
      ? (range) => adjustFormulaForCellMove(range, sameSheetOperation)
      : null,
  );
  const sheetMetadataChanged =
    whatIfTablesAfter !== null ||
    carriedSparklines !== null ||
    movedTableIds.size > 0 ||
    mergesChanged ||
    (carriers !== null && carriers !== sheet) ||
    options.conditionalFormatPlan?.changed === true ||
    options.validationPlan?.changed === true;
  if (options.skipIfUnchanged && !cellsChanged && !sheetMetadataChanged) {
    return { ok: false, reason: "unchanged" };
  }
  const nextSheet: S = {
    ...(carriers ?? sheet),
    cells: nextCells,
    rowCount: nextRowCount,
    columnCount: nextColumnCount,
  };
  if (carriedSparklines) {
    if (Object.keys(carriedSparklines).length > 0) nextSheet.sparklines = carriedSparklines;
    else delete nextSheet.sparklines;
  }
  if (whatIfTablesAfter) {
    if (whatIfTablesAfter.length > 0) nextSheet.whatIfTables = whatIfTablesAfter;
    else delete nextSheet.whatIfTables;
  }
  if (movedTableIds.size > 0 && nextSheet.dataTables) {
    const rangeOf = new Map(tablesAfterMove.map((table) => [table.id, table.range]));
    nextSheet.dataTables = nextSheet.dataTables.map((table) =>
      movedTableIds.has(table.id) ? { ...table, range: rangeOf.get(table.id)! } : table,
    );
  }
  if (appliesAuthoritativeMerges) {
    nextSheet.cells = cellsAfterMergesRemoved(nextSheet.cells, mergeSpans, nextMerges);
    if (nextMerges.length > 0) nextSheet.merges = [...nextMerges];
    else delete nextSheet.merges;
  }
  // The plans hold the rules the editor reads; a rule it does not read (a
  // list validation whose source is INDIRECT("Yes_No"), a rule kind it does
  // not model) is the file's, kept as the sheet holds it: a paste replaced
  // the whole list with the plan and every such rule of the sheet was gone
  
  // extends rules where it pastes.
  const unreadValidations = (Array.isArray(sheet.validations) ? sheet.validations : []).filter(
    (rule) => normalizeValidationRule(rule) === null,
  );
  const unreadConditionalFormats = (Array.isArray(sheet.conditionalFormats) ? sheet.conditionalFormats : []).filter(
    (rule) => normalizeConditionalRule(rule) === null,
  );
  if (options.conditionalFormatPlan) {
    const rules = [...options.conditionalFormatPlan.rules, ...unreadConditionalFormats];
    if (rules.length > 0) {
      nextSheet.conditionalFormats = rules;
    } else {
      delete nextSheet.conditionalFormats;
    }
  }
  if (options.validationPlan) {
    const rules = [...options.validationPlan.rules, ...unreadValidations];
    if (rules.length > 0) {
      nextSheet.validations = rules;
    } else {
      delete nextSheet.validations;
    }
  }
  // A paste directly below a table grows it (Table AutoExpansion); the
  // AutoCorrect Options button then offers the paste without the growth.
  const expansion =
    options.expandTables && spreadsheetTableAutoCorrectSettings().expandTables
      ? planTablePasteExpansion(
          nextSheet,
          { r1: baseRow, c1: baseCol, r2: maxR, c2: maxC },
          context.rangeHasLockedCell,
        )
      : null;
  // A cross-sheet move writes the OTHER sheets in the same commit: the
  // source sheet loses the moved cells, and every sheet's formula sources
  // (cells, names, validations, tables, conditional formats, sparklines,
  // charts) that named the source now name the destination.
  const followOtherSheets = sameSheetOperation
    ? (candidate: S): S =>
        candidate.id === sheet.id
          ? candidate
          : followMovedSources(
              rewriteSheetFormulaReferences(candidate, (formula) =>
                adjustFormulaForCellMove(formula, { ...sameSheetOperation, qualifiedOnly: true }),
              ),
              candidate.id,
            )
    : crossSheetMove && options.movedFrom
    ? (() => {
        const moved = options.movedFrom;
        return (candidate: S): S => {
          let next = candidate;
          if (candidate.id === crossSheetMove.id) {
            const cells: CellMap = { ...candidate.cells };
            for (const key of crossSheetMove.clearKeys) {
              const existing = cells[key];
              if (!existing) continue;
              const cleared = clearMovedSourceCell(existing, clipboardCells !== undefined);
              if (hasCellMetadata(cleared)) cells[key] = cleared;
              else delete cells[key];
            }
            next = { ...candidate, cells };
            for (const field of ["conditionalFormats", "validations"] as const) {
              const rules = crossSheetMove[field];
              if (!rules) continue;
              const record = next as unknown as Record<string, unknown>;
              if (rules.length > 0) record[field] = [...rules];
              else delete record[field];
            }
          }
          const hostSheet = candidate.name || candidate.id;
          return followMovedSources(
            rewriteSheetFormulaReferences(next, (formula) =>
              adjustFormulaForCrossSheetMove(formula, {
                from: moved.range,
                rowDelta: moved.rowDelta,
                colDelta: moved.colDelta,
                fromSheet: crossSheetMove.name,
                toSheet: sheet.name || sheet.id,
                hostSheet,
              }),
            ),
            candidate.id,
          );
        };
      })()
    : undefined;
  return {
    ok: true,
    sheet: nextSheet,
    expansion,
    ...(followOtherSheets ? { followOtherSheets } : {}),
    destination: { r1: baseRow, c1: baseCol, r2: maxR, c2: maxC },
  };
}
