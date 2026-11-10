/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How many drawing objects a workbook page may carry: a share of the page's
 * own size, not a count.
 *
 * The cap used to be 256 objects. Measured on the corpus, a shape costs ~430
 * bytes of page JSON, a chart ~6.7 KB and a preserved SmartArt diagram ~30 KB —
 * so 256 was 1% of the page as shapes and most of it as diagrams, and a
 * workbook of 300 plain shapes arrived with 44 missing for no reason the page
 * could feel. The bound that binds is now the bytes the objects put on the
 * page; the count survives only as a bound on the importer's own work, the
 * same 10,000 the import manifest uses (`PAGE_IMPORT_MAX_PROJECTED_OBJECTS`).
 * The server projection reads the same two numbers
 * (`_MAX_XLSX_DRAWING_OBJECT_BYTES`, `_MAX_XLSX_DRAWING_OBJECTS`), and the
 * whole-page envelope still has the last word on save
 * (docs/spreadsheet-page-limits.md).
 *
 * The rule and the measurement now live in `lib/page-object-budget`, because
 * a deck's slide turned out to need exactly this budget with its own two
 * numbers. This module is the workbook's binding of it, and stays the import
 * that every workbook caller uses.
 */

import {
  pageObjectBudgetSpent,
  pageObjectBytes,
  pageObjectsWithinBudget,
  type PageObjectBudgetLimits,
} from "../page-object-budget";

export const SPREADSHEET_MAX_OBJECTS = 10_000;
export const SPREADSHEET_OBJECT_BYTE_BUDGET = 8 * 1024 * 1024;

const SPREADSHEET_OBJECT_LIMITS: PageObjectBudgetLimits = {
  maxObjects: SPREADSHEET_MAX_OBJECTS,
  byteBudget: SPREADSHEET_OBJECT_BYTE_BUDGET,
};

/** Bytes the list's objects put on the page, measured once per object. */
export function spreadsheetObjectBytes(objects: readonly unknown[]): number {
  return pageObjectBytes(objects, SPREADSHEET_OBJECT_LIMITS);
}

/** Whether the list has used its share: no further object is projected. */
export function spreadsheetObjectBudgetSpent(
  objects: readonly unknown[],
): boolean {
  return pageObjectBudgetSpent(objects, SPREADSHEET_OBJECT_LIMITS);
}

/** The longest prefix of `objects` that fits the budget. */
export function spreadsheetObjectsWithinBudget<T>(objects: readonly T[]): T[] {
  return pageObjectsWithinBudget(objects, SPREADSHEET_OBJECT_LIMITS);
}
