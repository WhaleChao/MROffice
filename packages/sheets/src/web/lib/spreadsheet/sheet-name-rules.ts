/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { spreadsheetDefaultNames } from "@mrnote/core/web/lib/spreadsheet/default-names.ts";
/**
 * Excel's rules for a worksheet name — the one check every entry that names a
 * sheet (tab double-click, the tab menu's Rename, Home > Format > Rename Sheet,
 * a chart moved to a new sheet) runs before it commits.
 *
 * Microsoft Support, "Rename a worksheet": a worksheet name cannot be blank,
 * contain more than 31 characters, contain any of / \ ? * : [ ], begin or end
 * with an apostrophe (one in the middle is fine), or be "History", a word Excel
 * reserves internally. Excel also refuses a name another sheet of the workbook
 * already has, compared without regard to case ("That name is already taken").
 *
 * The characters matter beyond Excel's say-so: `[` `]` open an external
 * workbook reference and `:` spans sheets (`Q1:Q3!A1`), so a sheet named
 * `Q1:Q3` would be read back by the formula engine as a 3-D reference.
 */

/** Excel's longest worksheet name, in characters. */
export const MAX_SHEET_NAME_LENGTH = 31;

/** The characters a worksheet name may not contain. */
export const SHEET_NAME_FORBIDDEN_CHARACTERS = [
  "\\",
  "/",
  "?",
  "*",
  ":",
  "[",
  "]",
] as const;

/** The name Excel keeps for itself (the shared-workbook change history). */
export const RESERVED_SHEET_NAME = "History";

/** Why a worksheet name is refused, one per rule. */
export type SheetNameProblem =
  | "blank"
  | "tooLong"
  | "forbiddenCharacter"
  | "apostropheEdge"
  | "reserved"
  | "duplicate";

/**
 * The first rule `name` breaks, or null when Excel would accept it. `name` is
 * checked as given — callers trim what the user typed first. `otherNames` are
 * the names of the workbook's OTHER sheets (leave the renamed sheet out, so
 * changing only the case of its own name stays allowed).
 */
export function sheetNameProblem(
  name: string,
  otherNames: Iterable<string> = [],
): SheetNameProblem | null {
  if (name.trim() === "") return "blank";
  if (name.length > MAX_SHEET_NAME_LENGTH) return "tooLong";
  if (SHEET_NAME_FORBIDDEN_CHARACTERS.some((ch) => name.includes(ch))) {
    return "forbiddenCharacter";
  }
  if (name.startsWith("'") || name.endsWith("'")) return "apostropheEdge";
  const lower = name.toLowerCase();
  if (lower === RESERVED_SHEET_NAME.toLowerCase()) return "reserved";
  for (const other of otherNames) {
    if (other.toLowerCase() === lower) return "duplicate";
  }
  return null;
}

/**
 * The default name of the `n`th worksheet: Excel's "Sheet1", "Sheet2", … with
 * no space (Microsoft Support, "Insert or delete a worksheet"), which a formula
 * can also reference without quotes (`Sheet2!A1`, not `'Sheet 2'!A1`).
 */
export function defaultSheetName(n: number, locale?: string | null): string {
  // In the interface language, as that language's Excel names it
  // ("Tabelle1" in German): `spreadsheetDefaultNames`.
  return `${spreadsheetDefaultNames(locale).sheet}${n}`;
}

/**
 * Pages made before MRAI followed Excel hold "Sheet 1"-style names. Both
 * spellings count as the same default number, compared without regard to case.
 */
const DEFAULT_SHEET_NAME_NUMBER = /^sheet ?(\d+)$/i;

/** A default name in the interface language ("Tabelle3") counts too. */
function defaultSheetNameNumber(name: string): RegExpExecArray | null {
  const own = spreadsheetDefaultNames().sheet.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return DEFAULT_SHEET_NAME_NUMBER.exec(name) ?? new RegExp(`^${own} ?(\\d+)$`, "i").exec(name);
}

/**
 * The number the next default-named sheet takes: past the highest number any
 * sheet already carries as "Sheet<N>" or "Sheet <N>", and never below
 * `atLeast`. A workbook holding "Sheet 1" therefore gets "Sheet2" next, never a
 * "Sheet1" that differs from its neighbour only by a space. Existing sheets are
 * never renamed.
 */
export function nextDefaultSheetNumber(
  existingNames: Iterable<string>,
  atLeast = 1,
): number {
  let next = Math.max(1, Math.trunc(atLeast));
  for (const name of existingNames) {
    const match = defaultSheetNameNumber(name.trim());
    if (!match) continue;
    const taken = Number(match[1]);
    if (Number.isSafeInteger(taken) && taken >= next) next = taken + 1;
  }
  return next;
}

/** The message key that states a {@link SheetNameProblem} to the user. */
export const SHEET_NAME_PROBLEM_MESSAGE_KEYS: Readonly<
  Record<SheetNameProblem, string>
> = {
  blank: "spreadsheet.sheetNameBlank",
  tooLong: "spreadsheet.sheetNameTooLong",
  forbiddenCharacter: "spreadsheet.sheetNameForbiddenCharacter",
  apostropheEdge: "spreadsheet.sheetNameApostropheEdge",
  reserved: "spreadsheet.sheetNameReserved",
  duplicate: "spreadsheet.duplicateSheetName",
};
