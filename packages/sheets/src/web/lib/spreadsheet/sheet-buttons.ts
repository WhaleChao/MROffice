/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The sheet tab menu and Home > Cells > Insert / Delete / Format sheet
 * entries (Insert Sheet, Delete Sheet, Rename, Move or Copy, Hide / Unhide,
 * Tab Color) as pure workbook transactions. Moved out of SpreadsheetWindow so
 * the tabs and the spreadsheet editor commands run the same code; behaviour
 * is unchanged. The window keeps the protection notices, the rename box, the
 * active tab, the selection and the scroll.
 */

import { normalizeSpreadsheetPivotChartSource, type SpreadsheetPivotChartSource } from "@mrnote/core/web/lib/spreadsheet/chart.ts";
import { normalizeTabColor } from "./fill-series";
import { DEFAULT_SPREADSHEET_COLUMN_COUNT, DEFAULT_SPREADSHEET_ROW_COUNT, NEW_SHEET_DEFAULT_COLUMN_WIDTH_PX } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import type { NamedRange } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { resolvePivotChartPivot } from "@mrnote/core/web/lib/spreadsheet/pivot-chart.ts";
import { staticChartFromPivotChart } from "./pivot-chart-static";
import { renamePivotTimelineConnections } from "./pivot-timeline";
import {
  adjustSheetMoveReferences,
  invalidateDeletedSheetReferences,
  renameSheetReferences,
} from "./sheet-formula-references";
import {
  defaultSheetName,
  nextDefaultSheetNumber,
  sheetNameProblem,
  type SheetNameProblem,
} from "./sheet-name-rules";
import { workbookNamedRangeCatalog } from "./name-buttons";
import { copySheetForWorkbook, reorderSheets } from "./sheet-ops";

/** The most worksheets a workbook holds (the API's own limit). */
export const MAX_SHEETS = 200;

/** The worksheet fields the tab commands read. */
export interface WorkbookButtonSheet {
  id: string;
  name: string;
  hidden?: boolean;
  sheetState?: string;
  tabColor?: string;
  tabColorNativeBlocker?: string;
  namedRanges?: unknown;
  charts?: readonly unknown[];
  objects?: readonly unknown[];
  [key: string]: unknown;
}

/**
 * The workbook-level names a sheet created NOW must carry. Workbook names are
 * mirrored into every sheet (see {@link workbookNamedRangeCatalog}), and each
 * sheet's formulas resolve names from its own mirror — so a sheet inserted
 * after a name was defined starts with the same mirror, or =GlobalName on it
 * is #NAME? while Name Manager still lists the name (live audit B034; Excel:
 * a workbook-scope name is recognised on every worksheet).
 */
export function workbookNameMirrorsForNewSheet(
  workbook: ReadonlyArray<{ id: string; namedRanges?: unknown }>,
): { namedRanges?: NamedRange[] } {
  // A divergent mirror is flagged by the catalog itself and never copied.
  const mirrors = workbookNamedRangeCatalog(workbook).filter(
    (entry) =>
      (entry.scope ?? "workbook") === "workbook" &&
      entry.nativeBlocker !== "ambiguous-identity",
  );
  return mirrors.length ? { namedRanges: mirrors } : {};
}

/**
 * The fresh blank worksheet Insert Sheet adds (the tab menu's Insert, the +
 * button, Home > Insert > Insert Sheet): Excel's "Sheet<N>", N continuing past
 * the highest number any spelling carries (sheet names are unique without
 * regard to case), a matching unused id, and the workbook's names. Null at
 * the sheet limit.
 */
export function newWorksheetFor(
  workbook: ReadonlyArray<{ id: string; name: string; namedRanges?: unknown }>,
): {
  id: string;
  name: string;
  rowCount: number;
  columnCount: number;
  /** Excel's default column width for a new sheet (XLS-A095). */
  defaultColumnWidth: number;
  cells: Record<string, never>;
  namedRanges?: NamedRange[];
} | null {
  if (workbook.length >= MAX_SHEETS) return null;
  const existingIds = new Set(workbook.map((candidate) => candidate.id));
  const existingNames = new Set(
    workbook.map((candidate) => candidate.name.toLowerCase()),
  );
  let suffix = nextDefaultSheetNumber(existingNames, workbook.length + 1);
  let id = `sheet-${suffix}`;
  let name = defaultSheetName(suffix);
  while (existingIds.has(id) || existingNames.has(name.toLowerCase())) {
    suffix += 1;
    id = `sheet-${suffix}`;
    name = defaultSheetName(suffix);
  }
  return {
    id,
    name,
    rowCount: DEFAULT_SPREADSHEET_ROW_COUNT,
    columnCount: DEFAULT_SPREADSHEET_COLUMN_COUNT,
    defaultColumnWidth: NEW_SHEET_DEFAULT_COLUMN_WIDTH_PX,
    cells: {},
    ...workbookNameMirrorsForNewSheet(workbook),
  };
}

/**
 * Rename sheet `index` (every rename entry lands here). The typed value is
 * trimmed; an unchanged name is null. A name Excel refuses (blank, over 31
 * characters, \ / ? * : [ ], an apostrophe at either end, "History", or
 * another sheet's name in any case) is the problem. The rename follows every
 * sheet's references to the OLD name (formulas, names, validation sources,
 * charts, PivotTable timelines), the renamed sheet's own included.
 */
export function renamedSheetWorkbook<S extends WorkbookButtonSheet>(
  workbook: readonly S[],
  index: number,
  rawName: string,
): S[] | { problem: SheetNameProblem } | null {
  const target = workbook[index];
  if (!target) return null;
  const name = rawName.trim();
  if (name === target.name) return null;
  const problem = sheetNameProblem(
    name,
    workbook.flatMap((candidate, i) => (i === index ? [] : [candidate.name])),
  );
  if (problem) return { problem };
  const oldName = target.name;
  return workbook.map((candidate, i) => {
    const named = i === index ? { ...candidate, name } : candidate;
    return renamePivotTimelineConnections(
      renameSheetReferences(named, oldName, name),
      oldName,
      name,
    );
  });
}

/**
 * Set (or clear, with null or "") sheet `index`'s tab color, normalized to
 * #rrggbb. Null when nothing changes; "native-blocked" for an imported tab
 * color this editor cannot write back (theme/indexed/tinted).
 */
export function tabColoredWorkbook<S extends WorkbookButtonSheet>(
  workbook: readonly S[],
  index: number,
  rawColor: string | null,
): S[] | "native-blocked" | null {
  const target = workbook[index];
  if (!target) return null;
  if (target.tabColorNativeBlocker) return "native-blocked";
  const color = rawColor === null ? null : normalizeTabColor(rawColor);
  if ((target.tabColor ?? null) === color) return null;
  return workbook.map((candidate, i) => {
    if (i !== index) return candidate;
    if (color) return { ...candidate, tabColor: color };
    // Clearing: drop the key entirely so the no-color render path stays
    // byte-identical to a sheet that never had a tab color.
    const next: S = { ...candidate };
    delete next.tabColor;
    return next;
  });
}

/**
 * Whether Unhide can show this sheet. A workbook sheet is visible, hidden or
 * very hidden (SpreadsheetML `sheet/@state`); Excel's Unhide lists only the
 * hidden ones, and a very hidden sheet does not appear in the Unhide dialog
 * (Microsoft Support, "Hide or unhide worksheets"): only a macro sets it
 * visible again. It stays hidden everywhere else as an ordinary hidden sheet
 * does (tabs, sheet navigation), and its state is kept for the export.
 */
export function sheetUnhideable(sheet: {
  hidden?: boolean;
  sheetState?: string;
}): boolean {
  return sheet.hidden === true && sheet.sheetState !== "veryHidden";
}

/** The sheet as Unhide leaves it: visible, with neither visibility key. */
export function unhiddenSheet<S extends { hidden?: boolean; sheetState?: string }>(
  sheet: S,
): S {
  const next: S = { ...sheet };
  delete next.hidden;
  delete next.sheetState;
  return next;
}

/**
 * Hide / unhide sheet `index` (Excel "Hide" / "Unhide"). Hiding keeps the
 * data and cross-sheet references; Excel requires at least one visible
 * sheet, so hiding the last visible one is null, as is a sheet already so,
 * and so is unhiding a very hidden sheet (`sheetUnhideable`). Unhiding drops
 * the keys so a visible sheet stays byte-identical.
 */
export function sheetHiddenWorkbook<S extends WorkbookButtonSheet>(
  workbook: readonly S[],
  index: number,
  hidden: boolean,
): S[] | null {
  const target = workbook[index];
  if (!target) return null;
  if (Boolean(target.hidden) === hidden) return null;
  if (hidden && workbook.filter((s) => !s.hidden).length <= 1) return null;
  if (!hidden && !sheetUnhideable(target)) return null;
  return workbook.map((candidate, i) => {
    if (i !== index) return candidate;
    return hidden ? { ...candidate, hidden: true } : unhiddenSheet(candidate);
  });
}

/** The tab that becomes active when sheet `index` is hidden: the first other
 *  visible one, or -1. */
export function visibleSheetAfterHiding(
  workbook: ReadonlyArray<{ hidden?: boolean }>,
  index: number,
): number {
  return workbook.findIndex((s, i) => i !== index && !s.hidden);
}

/**
 * Delete sheet `index` (Microsoft Support, "Insert or delete a worksheet"). A
 * workbook keeps at least one sheet and, as in Excel, at least one VISIBLE
 * sheet ("last-visible"). Every reference to the deleted sheet becomes
 * Excel's #REF! across every formula source of the remaining sheets, and a
 * PivotChart elsewhere whose PivotTable goes with the sheet becomes a
 * standard chart of what it last showed (`pivotChartRows`; without them,
 * "pivot-chart": such a delete cannot be made faithfully).
 */
export function deletedSheetWorkbook<S extends WorkbookButtonSheet>(
  workbook: readonly S[],
  index: number,
  pivotChartRows: ((source: SpreadsheetPivotChartSource | undefined) => string[][]) | null,
): S[] | "last-visible" | "pivot-chart" | null {
  if (workbook.length <= 1) return null;
  if (index < 0 || index >= workbook.length) return null;
  const target = workbook[index];
  if (
    !target.hidden &&
    workbook.filter((candidate) => !candidate.hidden).length <= 1
  ) {
    return "last-visible";
  }
  const sheetOrder = workbook.map((candidate) => candidate.name);
  
  const deletedTables = (Array.isArray(target.dataTables) ? (target.dataTables as unknown[]) : [])
    .map((table) => (table && typeof table === "object" ? (table as { name?: unknown }).name : undefined))
    .filter((name): name is string => typeof name === "string" && name !== "");
  const pivotHosts = workbook as unknown as ReadonlyArray<{ id: string }>;
  const orphaned = (source: SpreadsheetPivotChartSource | undefined) =>
    resolvePivotChartPivot(pivotHosts, source)?.sheet.id === target.id;
  type Chart = Parameters<typeof staticChartFromPivotChart>[0] & { pivotSource?: unknown; type?: unknown };
  const chartSource = (chart: Chart) => normalizeSpreadsheetPivotChartSource(chart.pivotSource);
  let blocked = false;
  const staticChart = (chart: Chart, source: SpreadsheetPivotChartSource | undefined) => {
    if (!pivotChartRows) {
      blocked = true;
      return chart;
    }
    return staticChartFromPivotChart(chart, pivotChartRows(source));
  };
  const nextSheets = workbook
    .filter((_, i) => i !== index)
    .map((candidate) =>
      invalidateDeletedSheetReferences(candidate, target.name, sheetOrder, deletedTables),
    )
    .map((candidate) => {
      const charts = (candidate.charts as Chart[] | undefined)?.map((chart) =>
        orphaned(chart.pivotSource as SpreadsheetPivotChartSource | undefined)
          ? staticChart(chart, chart.pivotSource as SpreadsheetPivotChartSource | undefined)
          : chart,
      );
      const objects = (candidate.objects as Chart[] | undefined)?.map((object) => {
        const source = object.type === "chart" ? chartSource(object) : undefined;
        return orphaned(source) ? staticChart(object, source) : object;
      });
      const touched =
        charts?.some((chart, i) => chart !== candidate.charts?.[i]) ||
        objects?.some((object, i) => object !== candidate.objects?.[i]);
      return touched ? { ...candidate, ...(charts ? { charts } : {}), ...(objects ? { objects } : {}) } : candidate;
    });
  return blocked ? "pivot-chart" : nextSheets;
}

/**
 * The tab active after sheet `index` was deleted, when the deleted one was
 * the active one: the next visible tab to its right, else the nearest to its
 * left.
 */
export function activeSheetAfterDelete(
  nextSheets: ReadonlyArray<{ hidden?: boolean }>,
  index: number,
): number {
  // `index` now addresses the sheet that followed the deleted one.
  let nextActive = nextSheets.findIndex(
    (candidate, i) => i >= index && !candidate.hidden,
  );
  if (nextActive === -1) {
    for (let i = Math.min(index, nextSheets.length) - 1; i >= 0; i -= 1) {
      if (!nextSheets[i].hidden) {
        nextActive = i;
        break;
      }
    }
  }
  return Math.max(0, nextActive);
}

/**
 * Move sheet `from` to index `to` (Move Left / Right, Move or Copy, a drag).
 * A 3-D span follows Excel's endpoint rules: an endpoint moved past the
 * other one leaves the span, which then runs over the sheets that remain
 * (live audit B112). Null for a no-op move (clamped / same slot).
 */
export function movedSheetWorkbook<S extends WorkbookButtonSheet>(
  workbook: S[],
  from: number,
  to: number,
): S[] | null {
  const reordered = reorderSheets(workbook, from, to);
  if (reordered === workbook) return null;
  const orderBefore = workbook.map((candidate) => candidate.name || candidate.id);
  const orderAfter = reordered.map((candidate) => candidate.name || candidate.id);
  const movedName = workbook[from]?.name || workbook[from]?.id || "";
  return reordered.map((candidate) =>
    adjustSheetMoveReferences(candidate, orderBefore, orderAfter, movedName),
  );
}

/**
 * Where Move or Copy's "Before sheet" slot puts a MOVED sheet: taking the
 * sheet out first shifts every later slot one to the left, so landing in
 * front of itself or of its right-hand neighbour leaves the order as it is.
 * `beforeIndex` is the chosen sheet's index, or the sheet count for "(move
 * to end)". A copy is inserted in front of the chosen sheet as it stands.
 */
export function movedSheetIndex(from: number, beforeIndex: number): number {
  return from < beforeIndex ? beforeIndex - 1 : beforeIndex;
}

/**
 * Copy sheet `index` (Duplicate, Move or Copy > Create a copy) in front of
 * slot `insertAt`, or right after the source: a legal, unique "<name> (2)";
 * the copy's self-references point at the COPY; its tables get their own ids
 * and names. Null at the sheet limit.
 */
export function copiedSheetWorkbook<S extends WorkbookButtonSheet>(
  workbook: readonly S[],
  index: number,
  insertAt?: number,
): { sheets: S[]; index: number } | null {
  if (workbook.length >= MAX_SHEETS) return null;
  const source = workbook[index];
  if (!source) return null;
  const clone = copySheetForWorkbook(source, workbook);
  const insertedIndex =
    insertAt === undefined
      ? index + 1
      : Math.max(0, Math.min(workbook.length, Math.trunc(insertAt)));
  return {
    sheets: [
      ...workbook.slice(0, insertedIndex),
      clone,
      ...workbook.slice(insertedIndex),
    ],
    index: insertedIndex,
  };
}
