/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { defaultTableColumnName } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { computeFillValues } from "./fill-series";

/**
 * Excel's AutoCorrect > AutoFormat As You Type options for tables (Microsoft
 * Support, "AutoCorrect features in Excel" / "Undo automatic formatting in
 * Excel"):
 *
 * - `expandTables` is "Include new rows and columns in table": an entry typed
 *   in the row directly below a table, or the column directly right of it,
 *   makes the table grow to include it.
 * - `calculatedColumns` is "Fill formulas in tables to create calculated
 *   columns": a formula typed in a table column fills the whole column.
 *
 * Both are on by default. In Excel they are per-user application settings that
 * the AutoCorrect Options button's "Stop Automatically ..." choices turn off,
 * never workbook content. MRAI has no per-user settings store for them, so the
 * choice lasts for the browser session and is never written into a workbook.
 */
export interface SpreadsheetTableAutoCorrectSettings {
  expandTables: boolean;
  calculatedColumns: boolean;
}

const session: SpreadsheetTableAutoCorrectSettings = {
  expandTables: true,
  calculatedColumns: true,
};

/** A snapshot of the session's table AutoCorrect settings. */
export function spreadsheetTableAutoCorrectSettings(): SpreadsheetTableAutoCorrectSettings {
  return { ...session };
}

/** "Stop Automatically Expanding Tables" / "... Creating Calculated Columns". */
export function stopSpreadsheetTableAutoCorrect(
  option: keyof SpreadsheetTableAutoCorrectSettings,
): void {
  session[option] = false;
}

/** Restore Excel's defaults (both on); a new session starts this way. */
export function resetSpreadsheetTableAutoCorrectSession(): void {
  session.expandTables = true;
  session.calculatedColumns = true;
}

/**
 * The name a table gives a column that AutoExpansion adds at its right edge.
 * Excel continues a naming pattern it recognises in the last two column names
 * ("Qtr 1", "Qtr 2" -> "Qtr 3", Microsoft Support "Resize a table by adding or
 * removing rows and columns"); otherwise it uses the first free "ColumnN".
 * Names are case-insensitively unique within a table.
 */
export function tableAutoExpansionColumnName(
  existingNames: readonly string[],
  // The localized "ColumnN" text ("列1" in Chinese); Excel's English form
  // when omitted.
  defaultName?: (index: number) => string,
): string {
  const used = new Set(existingNames.map((name) => name.toLowerCase()));
  if (existingNames.length >= 2) {
    const [continued] = computeFillValues(existingNames.slice(-2), 1, "right");
    const candidate = continued?.trim() ?? "";
    if (
      candidate &&
      candidate.length <= 255 &&
      !used.has(candidate.toLowerCase())
    ) {
      return candidate;
    }
  }
  return defaultTableColumnName(existingNames, defaultName);
}

/**
 * Excel grows the references that a table's own data range fills when the
 * table grows (Microsoft Support, "Overview of Excel tables" and "Resize a
 * table by adding or removing rows and columns"): =SUM(C4:C10) under a table
 * whose C column holds C4:C10 reads =SUM(C4:C11) once the table has taken row
 * 11, so a total under a table keeps totalling it.
 *
 * Only a range that covers a table column's WHOLE data range is extended, and
 * only inside this sheet's own A1 references: a range that stops short of the
 * last row, starts above the first data row, or names another sheet is left
 * exactly as the user wrote it, as are table references (Table1[Col]), which
 * already follow the table. Absolute markers ($C$4:$C$10) are kept.
 */
export function extendRangeRefsForTableGrowth(
  formula: string,
  span: {
    firstColumn: number;
    lastColumn: number;
    firstDataRow: number;
    lastDataRow: number;
    newLastDataRow: number;
  },
): string {
  if (!formula.startsWith("=") || span.newLastDataRow <= span.lastDataRow) {
    return formula;
  }
  const columnIndex = (letters: string): number => {
    let index = 0;
    for (const letter of letters.toUpperCase()) {
      index = index * 26 + (letter.charCodeAt(0) - 64);
    }
    return index - 1;
  };
  // Quoted text is never a reference: rewrite only what lies between strings.
  const parts = formula.split(/("(?:[^"]|"")*")/);
  for (let index = 0; index < parts.length; index += 1) {
    if (index % 2 === 1) continue;
    parts[index] = parts[index].replace(
      /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}):(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g,
      (
        match,
        startColAbs: string,
        startCol: string,
        startRowAbs: string,
        startRow: string,
        endColAbs: string,
        endCol: string,
        endRowAbs: string,
        endRow: string,
        offset: number,
        whole: string,
      ) => {
        // A qualified or name-embedded reference is not ours to touch.
        const before = offset > 0 ? whole[offset - 1] : "";
        if (before === "!" || /[A-Za-z0-9_$.]/.test(before)) return match;
        const left = columnIndex(startCol);
        const right = columnIndex(endCol);
        if (
          Math.min(left, right) < span.firstColumn ||
          Math.max(left, right) > span.lastColumn ||
          Number(startRow) !== span.firstDataRow + 1 ||
          Number(endRow) !== span.lastDataRow + 1
        ) {
          return match;
        }
        return `${startColAbs}${startCol}${startRowAbs}${startRow}:${endColAbs}${endCol}${endRowAbs}${span.newLastDataRow + 1}`;
      },
    );
  }
  return parts.join("");
}
