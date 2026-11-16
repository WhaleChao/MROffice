/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Editing > Clear (All, Formats, Contents, Comments and Notes,
 * Hyperlinks) as one pure sheet transaction. Moved out of SpreadsheetWindow so
 * the ribbon and the spreadsheet editor commands run the same code; behaviour
 * is unchanged.
 */

import { clearRangeBy, type ClearVariant } from "./block-clear";
import {
  EXPLICIT_DEFAULT_BLANK_CELL,
  hasCellMetadata,
  type CellButtonsSheet,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  clearRulesForFormatClear,
  normalizeRules as normalizeConditionalRules,
  type ConditionalFormatRule,
} from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { clearRulesForClearAll, normalizeRules, type ValidationRule } from "@mrnote/core/web/lib/spreadsheet/data-validation.ts";
import type { DataTable } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { shiftFormula } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { lineStyleAt, patchLineStyles } from "@mrnote/core/web/lib/spreadsheet/line-styles.ts";
import { cellsAfterMergesRemoved, normalizeMerges, rangeHasMerge, removeMergesInRange, type MergeSpan } from "@mrnote/core/web/lib/spreadsheet/merge-cells.ts";
import { whatIfTablePartRefused, whatIfTablesAfterOverwrite } from "@mrnote/core/web/lib/spreadsheet/what-if-table-definition.ts";

export interface ClearButtonSheet extends CellButtonsSheet {
  merges?: MergeSpan[];
  dataTables?: DataTable[];
  conditionalFormats?: ConditionalFormatRule[];
  validations?: ValidationRule[];
  /** What-If Data Tables, checked where they are read (what-if-table-definition). */
  whatIfTables?: unknown;
}

export interface ClearButtonContext {
  /** Sheet protection is on (and not paused). */
  protectionEffective: boolean;
  /** The sheet's Insert hyperlinks permission (true when unprotected). */
  insertHyperlinksAllowed?: boolean;
  /** A cell of the range is locked under effective protection. */
  rangeHasLockedCell(range: { r1: number; c1: number; r2: number; c2: number }): boolean;
  gridRowCount: number;
  gridColumnCount: number;
}

/**
 * The sheet after Clear `variant` on `range`: "protected" when protection
 * refuses it, "what-if-table" when it would clear part of a What-If Data
 * Table's results (Excel: "Cannot change part of a data table."; clearing
 * all of them ends the table), null when nothing would change.
 */
export function clearedSheetForRange<S extends ClearButtonSheet>(
  sheet: S,
  range: { r1: number; c1: number; r2: number; c2: number },
  variant: ClearVariant,
  context: ClearButtonContext,
): S | "protected" | "what-if-table" | null {
  const clearsValues = variant === "contents" || variant === "all";
  if (clearsValues && whatIfTablePartRefused(sheet.whatIfTables, range)) return "what-if-table";
  const whatIfTablesLeft = clearsValues ? whatIfTablesAfterOverwrite(sheet.whatIfTables, [range]) : null;
  const mergeSpans = normalizeMerges(sheet.merges);
  // As stored, a whole-column rule to the grid's last row: read through the
  // sheet's used rows, clearing part of it cut the rest short (R316-X13).
  const conditionalFormats = normalizeConditionalRules(sheet.conditionalFormats, {
    maxRow: MAX_GRID_ROWS - 1,
    maxCol: MAX_GRID_COLUMNS - 1,
  });
  const validations = normalizeRules(sheet.validations);
  const wouldClearComment =
    (variant === "comments" || variant === "all") &&
    Object.entries(sheet.cells).some(([key, cell]) => {
      if (!cell?.comment) return false;
      const [row, col] = key.split(":").map(Number);
      return (
        row >= range.r1 &&
        row <= range.r2 &&
        col >= range.c1 &&
        col <= range.c2
      );
    });
  if (context.protectionEffective && wouldClearComment) return "protected";
  // Removing links is the Insert hyperlinks permission's other half, as in
  // the link dialog.
  if (variant === "removeHyperlinks" && context.insertHyperlinksAllowed === false) return "protected";
  if (context.rangeHasLockedCell(range)) return "protected";
  const clearsFormats = variant === "formats" || variant === "all";
  // Excel's Clear Formats and Clear All unmerge every merge area the
  // selection touches, the whole area even where the selection covers part of
  // it; that area may reach cells outside the selection, so they must be
  // editable too.
  const clearedMerges = clearsFormats && rangeHasMerge(mergeSpans, range);
  if (
    clearedMerges &&
    mergeSpans.some(
      (span) =>
        rangeHasMerge([span], range) &&
        context.rangeHasLockedCell({
          r1: span.row,
          r2: span.row + span.rowSpan - 1,
          c1: span.col,
          c2: span.col + span.colSpan - 1,
        }),
    )
  ) {
    return "protected";
  }
  const { cells: nextCells, changed: cellsChanged } = clearRangeBy(
    sheet.cells,
    range,
    variant,
    hasCellMetadata,
    { keepUnlocked: context.protectionEffective },
  );
  // A cleared Totals Row cell gives up its recipe with its content, as a
  // single-cell empty commit does; otherwise the old total keeps owning the
  // blank cell and the column's total cannot be changed or switched off
  // (live audit B173).
  let clearedTables: DataTable[] | undefined;
  if ((variant === "contents" || variant === "all") && sheet.dataTables?.length) {
    let changed = false;
    const next = sheet.dataTables.map((table) => {
      if (!table.totalsRow) return table;
      const totalsRowIndex = table.range.r2 + 1;
      if (totalsRowIndex < range.r1 || totalsRowIndex > range.r2) return table;
      const totals = { ...(table.totals ?? {}) };
      const totalLabels = { ...(table.totalLabels ?? {}) };
      const totalFormulas = { ...(table.totalFormulas ?? {}) };
      let touched = false;
      for (let col = Math.max(range.c1, table.range.c1); col <= Math.min(range.c2, table.range.c2); col += 1) {
        const offset = col - table.range.c1;
        if (offset in totals || offset in totalLabels || offset in totalFormulas) touched = true;
        delete totals[offset];
        delete totalLabels[offset];
        delete totalFormulas[offset];
      }
      if (!touched) return table;
      changed = true;
      const nextTable: DataTable = { ...table };
      if (Object.keys(totals).length > 0) nextTable.totals = totals; else delete nextTable.totals;
      if (Object.keys(totalLabels).length > 0) nextTable.totalLabels = totalLabels; else delete nextTable.totalLabels;
      if (Object.keys(totalFormulas).length > 0) nextTable.totalFormulas = totalFormulas; else delete nextTable.totalFormulas;
      return nextTable;
    });
    if (changed) clearedTables = next;
  }
  const conditionalClear = clearsFormats
    ? clearRulesForFormatClear(conditionalFormats, range)
    : { rules: conditionalFormats, changed: false };
  // Clear All also takes data validation off the cleared cells (Clear
  // Formats keeps it); a rule reaching past the selection keeps the rest.
  const validationClear =
    variant === "all"
      ? clearRulesForClearAll(validations, range, shiftFormula)
      : { rules: validations, changed: false };
  if (!validationClear) return null;
  // A whole-row / whole-column Clear Formats takes the LINE format off too
  // (B029); otherwise the columns added later would still be formatted.
  let lineStylesCleared = false;
  const clearLineStyles = (candidate: S): S => {
    if (!clearsFormats) return candidate;
    const wholeRows = range.c1 === 0 && range.c2 >= context.gridColumnCount - 1;
    const wholeCols = range.r1 === 0 && range.r2 >= context.gridRowCount - 1;
    let next = candidate;
    if (wholeRows) next = patchLineStyles(next, "row", range.r1, range.r2, () => undefined);
    if (wholeCols) next = patchLineStyles(next, "col", range.c1, range.c2, () => undefined);
    lineStylesCleared = next !== candidate;
    return next;
  };
  const lineCleared = clearLineStyles(sheet);
  // A cleared cell of a line that keeps its format holds the default
  // format as its own, so it does not read the line's again (B255).
  if (clearsFormats && (lineCleared.rowStyles || lineCleared.columnStyles)) {
    for (const key of Object.keys(sheet.cells)) {
      if (nextCells[key] !== undefined) continue;
      const separator = key.indexOf(":");
      const row = Number(key.slice(0, separator));
      const col = Number(key.slice(separator + 1));
      if (row < range.r1 || row > range.r2 || col < range.c1 || col > range.c2) continue;
      if (!lineStyleAt(lineCleared.rowStyles, lineCleared.columnStyles, row, col)) continue;
      nextCells[key] = { ...EXPLICIT_DEFAULT_BLANK_CELL, style: {} };
    }
  }
  if (
    !cellsChanged &&
    !conditionalClear.changed &&
    !clearedMerges &&
    !validationClear.changed &&
    !lineStylesCleared &&
    !clearedTables &&
    !whatIfTablesLeft
  ) {
    return null;
  }

  // Cell styles, conditional formats, merges and validation are persisted
  // facets of one Excel Clear command. Build one sheet so history/autosave
  // never expose a state in which only half of the command has landed.
  const nextSheet: S = { ...lineCleared, cells: nextCells };
  if (clearedTables) nextSheet.dataTables = clearedTables;
  if (whatIfTablesLeft) {
    if (whatIfTablesLeft.length > 0) nextSheet.whatIfTables = whatIfTablesLeft;
    else delete nextSheet.whatIfTables;
  }
  if (conditionalClear.changed) {
    if (conditionalClear.rules.length > 0) {
      nextSheet.conditionalFormats = conditionalClear.rules;
    } else {
      delete nextSheet.conditionalFormats;
    }
  }
  if (clearedMerges) {
    const nextMerges = removeMergesInRange(mergeSpans, range);
    nextSheet.cells = cellsAfterMergesRemoved(nextSheet.cells, mergeSpans, nextMerges);
    if (nextMerges.length > 0) nextSheet.merges = nextMerges;
    else delete nextSheet.merges;
  }
  if (validationClear.changed) {
    if (validationClear.rules.length > 0) {
      nextSheet.validations = validationClear.rules;
    } else {
      delete nextSheet.validations;
    }
  }
  return nextSheet;
}
