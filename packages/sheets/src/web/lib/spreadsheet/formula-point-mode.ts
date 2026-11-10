/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Point mode (Microsoft Support, "Create a simple formula in Excel": instead
 * of typing a reference, select the cell): while a formula is being typed and
 * the caret stands where a reference may go — just after `=`, `(`, `,`, an
 * operator or `:` — a click on a cell writes that cell's reference at the
 * caret, a drag writes the range, and a further click replaces the reference
 * the last click wrote. Anywhere else a click ends the entry as before.
 * Live audit P-002.
 */
import { columnToLetters } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";

export interface PointedReference {
  /** Where the pointed reference starts in the editor's text. */
  readonly start: number;
  /** The reference text the last point wrote. */
  readonly text: string;
}

const BEFORE_REFERENCE = new Set(["=", "(", ",", ";", "+", "-", "*", "/", "^", "&", "<", ">", ":", "{"]);
const AFTER_REFERENCE = new Set([")", ",", ";", "+", "-", "*", "/", "^", "&", "<", ">", "=", "}"]);

/** Whether `index` falls inside a "text" literal of the formula. */
function insideTextLiteral(value: string, index: number): boolean {
  let quotes = 0;
  for (let i = 0; i < index; i += 1) if (value[i] === '"') quotes += 1;
  return quotes % 2 === 1;
}

/**
 * The span of the editor's text a click on a cell replaces with a reference,
 * or null when the click is not a point (no formula, or the caret is not at a
 * place a reference can go).
 */
export function pointModeInsertion(
  value: string,
  selectionStart: number,
  selectionEnd: number,
  last: PointedReference | null,
): { start: number; end: number } | null {
  if (!value.trimStart().startsWith("=")) return null;
  const start = Math.min(selectionStart, selectionEnd);
  const end = Math.max(selectionStart, selectionEnd);
  if (
    last &&
    start === end &&
    end === last.start + last.text.length &&
    value.slice(last.start, end) === last.text
  ) {
    return { start: last.start, end };
  }
  if (insideTextLiteral(value, start)) return null;
  let before = start - 1;
  while (before >= 0 && value[before] === " ") before -= 1;
  if (before < 0 || !BEFORE_REFERENCE.has(value[before])) return null;
  let after = end;
  while (after < value.length && value[after] === " ") after += 1;
  if (after < value.length && !AFTER_REFERENCE.has(value[after])) return null;
  return { start, end };
}

/** The A1 reference of a pointed cell or of the range a drag covers. */
export function pointModeReference(
  anchor: { row: number; col: number },
  cell: { row: number; col: number },
): string {
  const a1 = (row: number, col: number) => `${columnToLetters(col)}${row + 1}`;
  if (anchor.row === cell.row && anchor.col === cell.col) return a1(cell.row, cell.col);
  return `${a1(Math.min(anchor.row, cell.row), Math.min(anchor.col, cell.col))}:${a1(Math.max(anchor.row, cell.row), Math.max(anchor.col, cell.col))}`;
}
