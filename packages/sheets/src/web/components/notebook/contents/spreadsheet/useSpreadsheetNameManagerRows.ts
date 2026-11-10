"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useState, useMemo, useEffect } from "react";
import { useTranslations } from "@mrnote/core/host/next-intl.tsx";
import { type FormulaEvaluationLocation } from "../../../../lib/spreadsheet/FormulaEvaluationDialog";
import { type EngineValue, type CellValue, type SheetEvaluator, type WorkbookSheetInput } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { isListedNamedRange, type NamedRange, namedRangeIdentity } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { nameManagerReferenceRange, nameManagerConstantValue, type NameManagerReferenceRange, NAME_MANAGER_VALUE_CELL_LIMIT, formatNameManagerValue, nameManagerValueIsError, nameManagerRowMatches, type NameManagerFilter, nameManagerRefIsBroken } from "../../../../lib/spreadsheet/name-manager-view";
import { formatSheetNameForFormula } from "@mrnote/core/web/lib/spreadsheet/name-syntax.ts";
import { nameRefShownFrom } from "@mrnote/core/web/lib/spreadsheet/name-relative-refs.ts";
import { normalizeDataTables } from "@mrnote/core/web/lib/spreadsheet/data-table.ts";
import { type PivotRange, type PivotMatrix } from "@mrnote/core/web/lib/spreadsheet/pivot.ts";
import { type PivotChartReport } from "@mrnote/core/web/lib/spreadsheet/pivot-output.ts";
import { type SheetModel } from "./spreadsheetSheetModel";
import { columnLabel } from "./spreadsheetRangeA1";
import { namedRangeUiIdentity } from "./spreadsheetWindowHelpers";
import { type Workbook } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { type PivotOutputPlan } from "@mrnote/core/web/lib/spreadsheet/pivot-output.ts";

/** What useSpreadsheetNameManagerRows reads from SpreadsheetWindow's render scope. */
export interface SpreadsheetNameManagerRowsInputs {
  /** The active cell a relative name's Refers to and Value read from. */
  activeCell: { row: number; col: number };
  evaluateExpressionsInSession: (expressions: readonly string[], location: FormulaEvaluationLocation) => Promise<Array<CellValue | undefined>>;
  formulaWorkbook: { evaluator: SheetEvaluator; evaluatorFor: (sheetName: string) => SheetEvaluator; workbook: Workbook; memoSnapshot: () => ReadonlyMap<string, EngineValue>; pivotPlans: Map<string, PivotOutputPlan>; pivotNativeFootprints: Map<string, PivotRange[]>; pivotChartReports: Map<string, PivotChartReport>; pivotChartMatrices: Map<string, PivotMatrix>; inputs: WorkbookSheetInput[]; pivotSheetIds: string[]; };
  locale: string;
  nameManagerFilter: NameManagerFilter;
  nameManagerOpen: boolean;
  nameManagerSort: "asc" | "desc";
  sheet: SheetModel;
  sheets: SheetModel[];
  t: ReturnType<typeof useTranslations<"console-notebooks">>;
  workbookNamedRanges: NamedRange[];
}

/** Moved verbatim from SpreadsheetWindow (called at the same position, so hook order is unchanged). */
export function useSpreadsheetNameManagerRows({
  activeCell,
  evaluateExpressionsInSession,
  formulaWorkbook,
  locale,
  nameManagerFilter,
  nameManagerOpen,
  nameManagerSort,
  sheet,
  sheets,
  t,
  workbookNamedRanges,
}: SpreadsheetNameManagerRowsInputs) {
  // Name Manager's rows (XLS-A066): every listed defined name and every
  // table, each with Excel's Value, whether that value is an error, its
  // Comment, then the Filter menu and the name order applied. A cell or range
  // reads the workbook's computed cells, a constant reads as itself, and any
  // other formula asks the calculation session; until (or unless) it answers,
  // the Value is Excel's `{...}` placeholder.
  const [nameManagerSessionValues, setNameManagerSessionValues] = useState<{
    readonly sheets: readonly SheetModel[];
    readonly at: string;
    readonly values: ReadonlyMap<string, EngineValue | undefined>;
  } | null>(null);
  const nameManagerFormulaEntries = useMemo(
    () =>
      nameManagerOpen
        ? workbookNamedRanges.filter(
            (entry) =>
              isListedNamedRange(entry) &&
              !entry.ref.includes("#REF!") &&
              nameManagerReferenceRange(entry.ref) === null &&
              nameManagerConstantValue(entry.ref) === undefined,
          )
        : [],
    [nameManagerOpen, workbookNamedRanges],
  );
  useEffect(() => {
    if (nameManagerFormulaEntries.length === 0) return;
    const at = `${activeCell.row}:${activeCell.col}`;
    if (nameManagerSessionValues?.sheets === sheets && nameManagerSessionValues.at === at) return;
    let live = true;
    const groups = new Map<string, NamedRange[]>();
    for (const entry of nameManagerFormulaEntries) {
      const sheetId =
        entry.scope === "worksheet" && entry.scopeSheetId
          ? entry.scopeSheetId
          : sheet.id;
      groups.set(sheetId, [...(groups.get(sheetId) ?? []), entry]);
    }
    void Promise.all(
      [...groups].map(([sheetId, entries]) =>
        evaluateExpressionsInSession(
          // A relative named formula is read from the active cell.
          entries.map((entry) => nameRefShownFrom(entry.ref, activeCell).replace(/^=\s*/, "")),
          { sheetId, row: 0, col: 0 },
        ).then((values) =>
          entries.map(
            (entry, index) =>
              [namedRangeIdentity(entry), values[index] as EngineValue | undefined] as const,
          ),
        ),
      ),
    ).then(
      (pairs) => {
        if (live) setNameManagerSessionValues({ sheets, at, values: new Map(pairs.flat()) });
      },
      () => undefined,
    );
    return () => {
      live = false;
    };
  }, [
    activeCell,
    evaluateExpressionsInSession,
    nameManagerFormulaEntries,
    nameManagerSessionValues,
    sheet.id,
    sheets,
  ]);
  const nameManagerRows = useMemo(() => {
    if (!nameManagerOpen) return [];
    const sheetKey = (candidate: SheetModel) => candidate.name || candidate.id;
    const readRange = (
      range: NameManagerReferenceRange,
      fallback: SheetModel | undefined,
    ): { value: EngineValue | CellValue[][]; truncated: boolean } | null => {
      const owner =
        range.sheetName === undefined
          ? fallback
          : sheets.find(
              (candidate) =>
                sheetKey(candidate).toLowerCase() ===
                range.sheetName!.toLowerCase(),
            );
      if (!owner) return null;
      const grid: CellValue[][] = [];
      let budget = NAME_MANAGER_VALUE_CELL_LIMIT;
      let truncated = false;
      for (let row = range.r1; row <= range.r2 && !truncated; row += 1) {
        const cells: CellValue[] = [];
        for (let col = range.c1; col <= range.c2; col += 1) {
          if (budget <= 0) {
            truncated = true;
            break;
          }
          budget -= 1;
          let value: CellValue = null;
          try {
            value = formulaWorkbook.workbook.getComputed(sheetKey(owner), row, col);
          } catch {
            value = null;
          }
          cells.push(value);
        }
        grid.push(cells);
      }
      if (grid.length === 1 && grid[0].length === 1 && !truncated) {
        return { value: grid[0][0], truncated };
      }
      return { value: grid, truncated };
    };
    const refError = { __formulaError: true as const, code: "#REF!" };
    const rows: Array<{
      readonly identity: string;
      readonly testId: string;
      readonly kind: "defined" | "table";
      readonly entry: NamedRange | null;
      readonly name: string;
      readonly value: string;
      readonly refersTo: string;
      readonly scope: "workbook" | "worksheet";
      readonly scopeLabel: string;
      readonly comment: string;
      readonly hasError: boolean;
    }> = [];
    const sessionValues =
      nameManagerSessionValues?.sheets === sheets &&
      nameManagerSessionValues.at === `${activeCell.row}:${activeCell.col}`
        ? nameManagerSessionValues.values
        : null;
    // A hidden name is left out of Name Manager, as in Excel; it stays in
    // `workbookNamedRanges`, so it still resolves and still counts as taken.
    for (const entry of workbookNamedRanges) {
      if (!isListedNamedRange(entry)) continue;
      const identity = namedRangeIdentity(entry);
      const scope = entry.scope ?? "workbook";
      const owner =
        scope === "worksheet"
          ? sheets.find((candidate) => candidate.id === entry.scopeSheetId)
          : sheet;
      let value: EngineValue | undefined;
      let text: string;
      // A relative name reads from the active cell, its Refers to included.
      const shownRef = nameRefShownFrom(entry.ref, activeCell);
      const broken = nameManagerRefIsBroken(shownRef);
      const range = broken ? null : nameManagerReferenceRange(shownRef);
      if (broken) {
        value = refError;
        text = refError.code;
      } else if (range) {
        const read = readRange(range, owner);
        // A range is an array of values, never an error itself.
        value = !read
          ? refError
          : Array.isArray(read.value)
            ? undefined
            : (read.value as EngineValue);
        text = read
          ? formatNameManagerValue(read.value, read.truncated)
          : refError.code;
      } else {
        const constant = nameManagerConstantValue(shownRef);
        value = constant !== undefined ? constant : sessionValues?.get(identity);
        text = formatNameManagerValue(value);
      }
      rows.push({
        identity,
        testId: namedRangeUiIdentity(entry),
        kind: "defined",
        entry,
        name: entry.name,
        value: text,
        refersTo: shownRef.startsWith("=") ? shownRef : `=${shownRef}`,
        scope,
        scopeLabel:
          scope === "worksheet"
            ? owner?.name ?? t("spreadsheet.help.worksheet")
            : t("spreadsheet.help.workbook"),
        comment: entry.comment ?? "",
        hasError: nameManagerValueIsError(value, shownRef),
      });
    }
    // Table names are workbook names too; Name Manager lists them beside the
    // defined names and refers each to its whole range.
    for (const candidate of sheets) {
      const tables = normalizeDataTables(candidate.dataTables, {
        rowCount: candidate.rowCount,
        columnCount: candidate.columnCount,
      });
      for (const table of tables) {
        const { r1, c1, r2, c2 } = table.range;
        const read = readRange({ r1, c1, r2, c2 }, candidate);
        rows.push({
          identity: `\u0001table\u0000${candidate.id}\u0000${table.id}`,
          testId: `table--${encodeURIComponent(table.name.toLowerCase())}`,
          kind: "table",
          entry: null,
          name: table.name,
          value: read
            ? formatNameManagerValue(read.value, read.truncated)
            : refError.code,
          refersTo: `=${formatSheetNameForFormula(sheetKey(candidate))}!$${columnLabel(c1)}$${r1 + 1}:$${columnLabel(c2)}$${r2 + 1}`,
          scope: "workbook",
          scopeLabel: t("spreadsheet.help.workbook"),
          comment: "",
          hasError: false,
        });
      }
    }
    return rows
      .filter((row) => nameManagerRowMatches(row, nameManagerFilter))
      .sort((left, right) => {
        const comparison = left.name.localeCompare(right.name, locale, {
          sensitivity: "base",
          numeric: true,
        });
        return nameManagerSort === "asc" ? comparison : -comparison;
      });
  }, [
    activeCell,
    formulaWorkbook.workbook,
    locale,
    nameManagerFilter,
    nameManagerOpen,
    nameManagerSessionValues,
    nameManagerSort,
    sheet,
    sheets,
    t,
    workbookNamedRanges,
  ]);
  return {
    nameManagerRows,
  };
}
