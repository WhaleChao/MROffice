/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helper for the "Pick From Drop-down List" command (Excel Alt+Down /
 * right-click → "Pick From Drop-down List…").
 *
 * Unlike a data-validation LIST rule (a fixed allowed set) or the in-cell ghost
 * autocomplete (a single best-match suffix), this lists the DISTINCT existing
 * values already entered ELSEWHERE in the SAME column, so the user can pick a
 * value they've used before without retyping. It works on ANY cell, validated
 * or not.
 *
 * Excel scans the contiguous vertical block of populated cells the active cell
 * belongs to (stopping at the first blank above and below). We mirror that: the
 * candidate set is the run of non-blank cells directly above and below the
 * active row (the active cell itself is skipped — you're picking a value for
 * it). Values are trimmed, de-duplicated (first-seen order, top-to-bottom), and
 * capped so a pathological column can't build an unbounded popover.
 */

/** A cell as far as pick-list enumeration cares: just its display text. */
export interface PickListCell {
  value?: string;
}

/** A `"row:col"`-keyed view of the sheet's resolved-display cells. */
export type PickListCellMap = Record<string, PickListCell | undefined>;

/** Max distinct options a pick-list popover may offer (DoS bound on the list). */
export const MAX_PICK_LIST_VALUES = 1_000;

/** Trim a cell's display value to a string ("" when blank/absent). */
function cellText(cell: PickListCell | undefined): string {
  return (cell?.value ?? "").trim();
}

/**
 * The distinct existing values for the pick-list popover at (activeRow, col).
 *
 * Scans UP from activeRow-1 until a blank cell (or the top) and DOWN from
 * activeRow+1 until a blank cell (or rowCount), collecting non-blank trimmed
 * values. The two directions are merged with the upward values first (nearer
 * the top of the contiguous block), then de-duplicated in first-seen order so
 * the list reads top-to-bottom like Excel. The active cell's own value is never
 * included. Returns at most MAX_PICK_LIST_VALUES entries; an empty result means
 * "nothing to pick" (an isolated cell with no populated neighbours).
 *
 * @param cells the sheet's resolved-display cell map (`"row:col"` -> { value })
 * @param activeRow the 0-based row of the cell the popover is opening on
 * @param col the 0-based column to enumerate
 * @param rowCount the sheet's row count (scan stays within `[0, rowCount)`)
 */
export function pickListValues(
  cells: PickListCellMap,
  activeRow: number,
  col: number,
  rowCount: number,
): string[] {
  const rows = Number.isFinite(rowCount) ? Math.max(0, Math.floor(rowCount)) : 0;
  if (
    !Number.isInteger(activeRow) ||
    !Number.isInteger(col) ||
    activeRow < 0 ||
    col < 0 ||
    rows <= 0
  ) {
    return [];
  }
  const above: string[] = [];
  // Walk up from the cell directly above until a blank (or the top).
  for (let r = activeRow - 1; r >= 0; r -= 1) {
    const text = cellText(cells[`${r}:${col}`]);
    if (text === "") break;
    above.push(text);
  }
  above.reverse(); // restore top-to-bottom order
  const below: string[] = [];
  // Walk down from the cell directly below until a blank (or the bottom).
  for (let r = activeRow + 1; r < rows; r += 1) {
    const text = cellText(cells[`${r}:${col}`]);
    if (text === "") break;
    below.push(text);
  }
  const seen = new Set<string>();
  const order: string[] = [];
  for (const value of [...above, ...below]) {
    if (seen.has(value)) continue;
    seen.add(value);
    order.push(value);
    if (order.length >= MAX_PICK_LIST_VALUES) break;
  }
  return order;
}
