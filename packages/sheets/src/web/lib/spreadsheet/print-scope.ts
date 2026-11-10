/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * File > Print > Settings, as Excel offers them: Print Active Sheets (the
 * default: the active sheet and any sheets grouped with it), Print Entire
 * Workbook, Print Selection (the selected cells on the active sheet, with
 * its page setup and titles but never its print area), and Ignore Print
 * Area. The choice travels with the print's PDF view; the server prints the
 * saved workbook through it (`apply_print_scope`).
 */
import type { SpreadsheetPdfExportView } from "@mrnote/core/web/lib/spreadsheet/pdf-export-view.ts";

export type SpreadsheetPrintWhat = "activeSheets" | "workbook" | "selection";

export const SPREADSHEET_PRINT_WHAT: readonly SpreadsheetPrintWhat[] = ["activeSheets", "workbook", "selection"];

export interface SpreadsheetPrintScope {
  what: SpreadsheetPrintWhat;
  ignorePrintArea: boolean;
}

export const DEFAULT_SPREADSHEET_PRINT_SCOPE: SpreadsheetPrintScope = { what: "activeSheets", ignorePrintArea: false };

export type SpreadsheetPrintScopeView = Pick<SpreadsheetPdfExportView, "printSheetIds" | "printSelection" | "ignorePrintArea">;

/** The view fields one print sends for `scope`. */
export function spreadsheetPrintScopeView(
  scope: SpreadsheetPrintScope,
  place: {
    activeSheetId: string;
    /** The active sheet's group in workbook order (the active sheet alone when none). */
    activeSheetIds: readonly string[];
    selection: { r1: number; c1: number; r2: number; c2: number };
  },
): SpreadsheetPrintScopeView {
  if (scope.what === "selection") {
    const { r1, c1, r2, c2 } = place.selection;
    return { printSheetIds: [place.activeSheetId], printSelection: { r1, c1, r2, c2 } };
  }
  const ignore = scope.ignorePrintArea ? { ignorePrintArea: true as const } : {};
  if (scope.what === "workbook") return ignore;
  const ids = place.activeSheetIds.includes(place.activeSheetId) ? place.activeSheetIds : [place.activeSheetId];
  return { printSheetIds: [...ids], ...ignore };
}
