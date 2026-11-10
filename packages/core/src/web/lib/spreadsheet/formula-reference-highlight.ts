/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The references of a formula being typed, each with its colour.
 *
 * Excel's range finder (Microsoft Support, "Move or copy a formula" /
 * "Display the relationships between formulas and cells": "When you edit a
 * formula, Excel color-codes each cell reference or range ... and places a
 * matching border around the cells"): every reference in the entry is drawn
 * in a colour and the cells it names wear a border of the same colour on the
 * sheet. The colours run in a fixed order and start over after the last; the
 * same reference written twice wears the same colour both times.
 *
 * This reads the ENTRY, which is rarely a complete formula while it is being
 * typed (`=SUM(C5:C7`), so it does not go through the evaluator's tokenizer:
 * it scans for A1 references outside string literals, with their optional
 * sheet qualifier, and leaves everything else alone. Structured references,
 * defined names and R1C1 spellings are not coloured.
 */

/** Excel's range-finder colours, in the order references take them. */
export const FORMULA_REFERENCE_COLORS = [
  "#2f6fd0",
  "#e0352b",
  "#8a3fc7",
  "#1e9b4f",
  "#d6338c",
  "#e58a00",
  "#13969b",
] as const;

export interface FormulaReferenceSpan {
  /** Where the reference's text sits in the entry, [start, end). */
  readonly start: number;
  readonly end: number;
  /** The sheet the reference names (unquoted), or undefined for the entry's own. */
  readonly sheet?: string;
  /** Zero-based, inclusive, top-left first. A whole column or row runs to
   * `Infinity` on the open axis. */
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  readonly color: string;
}

const MAX_COLUMN = 16_384;
const MAX_ROW = 1_048_576;
const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_\u0080-￿][\w.\u0080-￿]*))!`;
const CELL = String.raw`\$?([A-Za-z]{1,3})\$?(\d{1,7})`;
const REFERENCE = new RegExp(
  String.raw`(?:${SHEET})?(?:${CELL}(?::${CELL})?|\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})|\$?(\d{1,7}):\$?(\d{1,7}))`,
  "y",
);

function columnIndex(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) value = value * 26 + (letter.charCodeAt(0) - 64);
  return value - 1;
}

const isNameChar = (char: string | undefined): boolean =>
  char !== undefined && /[\w.\u0080-￿]/.test(char);

/**
 * Every reference of `entry` with its colour, in the order written; empty
 * for an entry that is not a formula.
 */
export function formulaReferenceSpans(entry: string): FormulaReferenceSpan[] {
  if (!entry.startsWith("=")) return [];
  const spans: FormulaReferenceSpan[] = [];
  const colorOf = new Map<string, string>();
  let index = 1;
  while (index < entry.length) {
    const char = entry[index];
    if (char === '"') {
      // A string literal, a doubled quote inside it included.
      index += 1;
      while (index < entry.length && (entry[index] !== '"' || entry[index + 1] === '"')) {
        index += entry[index] === '"' ? 2 : 1;
      }
      index += 1;
      continue;
    }
    // A reference starts where a name could not continue from the left.
    if (isNameChar(entry[index - 1]) && char !== "'") {
      index += 1;
      continue;
    }
    REFERENCE.lastIndex = index;
    const match = REFERENCE.exec(entry);
    if (!match) {
      // Skip the whole name so `LOG10` or `Sales2024` is not read from its middle.
      if (isNameChar(char)) {
        while (index < entry.length && isNameChar(entry[index])) index += 1;
      } else index += 1;
      continue;
    }
    const end = index + match[0].length;
    const next = entry[end];
    // `LOG10(`, `A1B`, `Sheet1!A1x`: a longer name, not a reference.
    if (isNameChar(next) || next === "(" || next === "!") {
      while (index < entry.length && isNameChar(entry[index])) index += 1;
      if (index < end) index = end;
      continue;
    }
    const sheet = match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2];
    let r1: number, c1: number, r2: number, c2: number;
    if (match[3] !== undefined) {
      c1 = columnIndex(match[3]);
      r1 = Number(match[4]) - 1;
      c2 = match[5] !== undefined ? columnIndex(match[5]) : c1;
      r2 = match[6] !== undefined ? Number(match[6]) - 1 : r1;
    } else if (match[7] !== undefined) {
      c1 = columnIndex(match[7]);
      c2 = columnIndex(match[8]!);
      r1 = 0;
      r2 = Number.POSITIVE_INFINITY;
    } else {
      r1 = Number(match[9]) - 1;
      r2 = Number(match[10]) - 1;
      c1 = 0;
      c2 = Number.POSITIVE_INFINITY;
    }
    const valid =
      r1 >= 0 && c1 >= 0 &&
      (r2 === Number.POSITIVE_INFINITY || r2 < MAX_ROW) && r1 < MAX_ROW &&
      (c2 === Number.POSITIVE_INFINITY || c2 < MAX_COLUMN) && c1 < MAX_COLUMN;
    if (valid) {
      const top = Math.min(r1, r2), bottom = Math.max(r1, r2);
      const left = Math.min(c1, c2), right = Math.max(c1, c2);
      const key = `${(sheet ?? "").toLowerCase()}!${top}:${left}:${bottom}:${right}`;
      let color = colorOf.get(key);
      if (!color) {
        color = FORMULA_REFERENCE_COLORS[colorOf.size % FORMULA_REFERENCE_COLORS.length]!;
        colorOf.set(key, color);
      }
      spans.push({ start: index, end, ...(sheet !== undefined ? { sheet } : {}), r1: top, c1: left, r2: bottom, c2: right, color });
    }
    index = end;
  }
  return spans;
}

/** The entry cut into runs for a coloured rendering: a reference's text with
 * its colour, everything between without one. */
export function formulaReferenceRuns(
  entry: string,
  spans: readonly FormulaReferenceSpan[] = formulaReferenceSpans(entry),
): { text: string; color?: string }[] {
  const runs: { text: string; color?: string }[] = [];
  let cursor = 0;
  for (const span of spans) {
    if (span.start > cursor) runs.push({ text: entry.slice(cursor, span.start) });
    runs.push({ text: entry.slice(span.start, span.end), color: span.color });
    cursor = span.end;
  }
  if (cursor < entry.length) runs.push({ text: entry.slice(cursor) });
  return runs;
}
