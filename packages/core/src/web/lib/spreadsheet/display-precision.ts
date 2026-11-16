/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { spreadsheetCellIsText } from "./entry-parse";
import type { SheetEvaluatorOptions } from "./formula-engine";
import { displayedPrecisionRounder } from "./number-format";

type PrecisionCell = { value?: unknown; formula?: unknown; kind?: unknown; style?: { numberFormat?: unknown } };

/** The number a constant cell holds, or null (a formula, text, a blank). */
function constantNumber(cell: PrecisionCell): number | null {
  if (cell.formula || typeof cell.value !== "string" || cell.value === "" || cell.value.startsWith("=")) return null;
  if (spreadsheetCellIsText(cell as Parameters<typeof spreadsheetCellIsText>[0])) return null;
  const number = Number(cell.value);
  return Number.isFinite(number) && cell.value.trim() !== "" ? number : null;
}

/**
 * Set precision as displayed turned on (Excel: File > Options > Advanced >
 * When calculating this workbook; "Data will permanently lose accuracy"):
 * every number a cell holds becomes the number its format shows, for good
 * (Microsoft Learn, "Floating-point arithmetic may give inaccurate result in
 * Excel": the option "forces the value of each number in the worksheet to be
 * the displayed value", and the lost digits cannot be recovered). General,
 * dates, fractions and text formats keep the number (displayedPrecisionRounder).
 * The same map when no cell changes.
 */







export function cellsAtKeysRoundedAsDisplayed<C extends PrecisionCell>(
  cells: Readonly<Record<string, C>>,
  keys: readonly string[],
): Readonly<Record<string, C>> {
  const touched: Record<string, C> = {};
  for (const key of keys) {
    const cell = cells[key];
    if (cell) touched[key] = cell;
  }
  const rounded = cellsRoundedAsDisplayed(touched);
  return rounded === touched ? cells : { ...cells, ...rounded };
}

export function cellsRoundedAsDisplayed<C extends PrecisionCell>(cells: Readonly<Record<string, C>>): Readonly<Record<string, C>> {
  let out: Record<string, C> | null = null;
  for (const key in cells) {
    const cell = cells[key]!;
    const format = cell.style?.numberFormat;
    if (typeof format !== "string") continue;
    const number = constantNumber(cell);
    const round = number === null ? null : displayedPrecisionRounder(format);
    if (number === null || !round) continue;
    const shown = round(number);
    if (shown === number || !Number.isFinite(shown)) continue;
    out ??= { ...cells };
    out[key] = { ...cell, value: String(shown === 0 ? 0 : Number(shown.toPrecision(15))) };
  }
  return out ?? cells;
}

/**
 * Set precision as displayed (calcPr@fullPrecision="0"): the formula and
 * number cells whose number format rounds the number they keep, as "row:col"
 * and the format -- plain data, so the background calculation reads the same.
 */
export function displayPrecisionEntries(
  cells: Readonly<Record<string, PrecisionCell | undefined>>,
  covered?: ReadonlySet<string>,
): Array<readonly [string, string]> {
  const out: Array<readonly [string, string]> = [];
  for (const key in cells) {
    const cell = cells[key];
    if (!cell || covered?.has(key)) continue;
    // A formula's number, and a number typed since the option was turned on
    // (which Excel keeps as shown too).
    if (!cell.formula && !(typeof cell.value === "string" && cell.value.startsWith("=")) && constantNumber(cell) === null) continue;
    const format = cell.style?.numberFormat;
    if (typeof format === "string" && displayedPrecisionRounder(format)) out.push([key, format]);
  }
  return out;
}

/** The engine's displayPrecision over {@link displayPrecisionEntries}. */
export function displayPrecisionOption(
  entries: ReadonlyArray<readonly [string, string]> | undefined,
): Pick<SheetEvaluatorOptions, "displayPrecision"> {
  if (!entries?.length) return {};
  const rounders = new Map<string, (value: number) => number>();
  for (const [key, format] of entries) {
    const round = displayedPrecisionRounder(format);
    if (round) rounders.set(key, round);
  }
  return rounders.size ? { displayPrecision: (row: number, col: number) => rounders.get(`${row}:${col}`) ?? null } : {};
}
