/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { clearRangeContents } from "./block-clear";
import { hasCellMetadata, updateCellStyle, updateCellValue, type CellButtonsSheet } from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import { checkboxSeedValue, isCheckboxCell, isCheckboxChecked } from "./checkbox";

/**
 * Delete over a selection holding checkbox cells (Microsoft, "Using check
 * boxes in Excel"): when any box is checked, every box is unchecked (its cell
 * stays FALSE); when all are unchecked, the boxes are removed together with
 * their values, so the cells are empty ("If you remove the checkboxes from
 * cells, the TRUE and FALSE values are also removed" — Exceljet; Microsoft's
 * way to keep the values is Clear Formats). Other cells of the selection clear
 * as Clear Contents does. Null when the selection holds no checkbox.
 * Live audit B293, P-CHECKBOX-DELETE-RETAINED-BOOL.
 */
export function sheetAfterCheckboxDelete<S extends CellButtonsSheet>(
  sheet: S,
  rect: { r1: number; c1: number; r2: number; c2: number },
): S | null {
  const checkboxKeys: Array<[number, number]> = [];
  for (let r = rect.r1; r <= rect.r2; r += 1) {
    for (let c = rect.c1; c <= rect.c2; c += 1) {
      if (isCheckboxCell(sheet.cells[`${r}:${c}`]?.style as Record<string, unknown> | undefined)) checkboxKeys.push([r, c]);
    }
  }
  if (checkboxKeys.length === 0) return null;
  const anyChecked = checkboxKeys.some(([r, c]) => isCheckboxChecked(sheet.cells[`${r}:${c}`]?.value));
  const { cells: clearedCells } = clearRangeContents(sheet.cells, rect, hasCellMetadata);
  let next: S = { ...sheet, cells: clearedCells };
  for (const [r, c] of checkboxKeys) {
    const original = sheet.cells[`${r}:${c}`];
    if (!original) continue;
    if (anyChecked) {
      next = { ...next, cells: { ...next.cells, [`${r}:${c}`]: original } };
      next = updateCellValue(next, r, c, checkboxSeedValue(""));
    } else {
      // The cleared cell (its value gone) loses its box too.
      next = updateCellStyle(next, r, c, { checkbox: false });
    }
  }
  return next;
}
