/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// The cells a formula reaches through defined names and table (structured)
// references (XLS-A062: Go To Special > Precedents / Dependents follow them,
// as Excel's do: "=SUM(Sales)" has the cells Sales names as precedents, and
// "=SUM(Table1[Amount])" the Amount column's data cells).
//
// Same sheet only, like the rest of Go To Special: a name or table on another
// sheet adds nothing.
// ---------------------------------------------------------------------------

import { tokenize, type StructuredSpec } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { resolveName, type NamedRange } from "@mrnote/core/web/lib/spreadsheet/named-ranges.ts";
import { shiftNameRefRelativeAxes } from "@mrnote/core/web/lib/spreadsheet/name-relative-refs.ts";

export interface IndirectRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

interface IndirectTable {
  name: string;
  range: IndirectRange;
  headerRow: boolean;
  totalsRow?: boolean;
  columns?: ReadonlyArray<{ name: string }>;
}

export interface IndirectReferenceContext {
  sheetId: string;
  sheetName: string;
  names: readonly NamedRange[];
  tables: readonly IndirectTable[];
  /** A header cell's text, for a table without stored column names. */
  headerText?: (row: number, col: number) => string;
}

function columnIndex(letters: string): number {
  let value = 0;
  for (const ch of letters.toUpperCase()) value = value * 26 + (ch.charCodeAt(0) - 64);
  return value - 1;
}

const A1_AREA = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/;

/** A name's reference as a same-sheet area, or null (another sheet, a formula). */
function nameArea(ref: string, sheetName: string): IndirectRange | null {
  const text = ref.trim().replace(/^=/, "");
  const match = A1_AREA.exec(text);
  if (!match) return null;
  const sheet = match[1]?.replace(/''/g, "'") ?? match[2];
  if (sheet !== undefined && sheet.toUpperCase() !== sheetName.toUpperCase()) return null;
  const r1 = Number(match[4]) - 1;
  const c1 = columnIndex(match[3]);
  const r2 = match[6] ? Number(match[6]) - 1 : r1;
  const c2 = match[5] ? columnIndex(match[5]) : c1;
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

function tableColumn(table: IndirectTable, name: string, context: IndirectReferenceContext): number | null {
  const key = name.trim().toUpperCase();
  const width = table.range.c2 - table.range.c1 + 1;
  for (let offset = 0; offset < width; offset += 1) {
    const stored = table.columns?.[offset]?.name;
    const label = stored ?? (table.headerRow ? context.headerText?.(table.range.r1, table.range.c1 + offset) : undefined);
    if (label !== undefined && label.trim().toUpperCase() === key) return table.range.c1 + offset;
  }
  return null;
}

function tableArea(
  table: IndirectTable,
  spec: StructuredSpec,
  row: number,
  context: IndirectReferenceContext,
): IndirectRange | null {
  const header = table.headerRow ? table.range.r1 : null;
  const dataTop = table.headerRow ? table.range.r1 + 1 : table.range.r1;
  const dataBottom = table.range.r2;
  const totals = table.totalsRow ? table.range.r2 + 1 : null;
  let r1: number;
  let r2: number;
  switch (spec.band) {
    case "headers":
      if (header === null) return null;
      r1 = r2 = header;
      break;
    case "totals":
      if (totals === null) return null;
      r1 = r2 = totals;
      break;
    case "all":
      r1 = header ?? dataTop;
      r2 = totals ?? dataBottom;
      break;
    case "headersData":
      r1 = header ?? dataTop;
      r2 = dataBottom;
      break;
    case "dataTotals":
      r1 = dataTop;
      r2 = totals ?? dataBottom;
      break;
    case "thisrow":
      r1 = r2 = row;
      break;
    default:
      r1 = dataTop;
      r2 = dataBottom;
  }
  if (r2 < r1) return null;
  let c1 = table.range.c1;
  let c2 = table.range.c2;
  if (spec.column !== null) {
    const first = tableColumn(table, spec.column, context);
    const last = spec.columnEnd !== undefined ? tableColumn(table, spec.columnEnd, context) : first;
    if (first === null || last === null) return null;
    c1 = Math.min(first, last);
    c2 = Math.max(first, last);
  }
  return { r1, c1, r2, c2 };
}

/**
 * The same-sheet areas `formula` (written in the cell at row, col) reaches
 * through defined names and structured references.
 */
export function indirectReferenceAreas(
  formula: string,
  row: number,
  col: number,
  context: IndirectReferenceContext,
): IndirectRange[] {
  const body = formula.trim().replace(/^=/, "");
  const tokens = tokenize(body);
  if (!tokens) return [];
  const areas: IndirectRange[] = [];
  tokens.forEach((token, index) => {
    if (token.type === "ident") {
      // A function name is followed by "("; TRUE / FALSE are not names.
      if (tokens[index + 1]?.type === "lparen") return;
      if (/^(TRUE|FALSE)$/i.test(token.value)) return;
      const stored = resolveName(token.value, context.names, context.sheetId);
      // A relative name reaches the cells relative to this formula's cell.
      const ref = stored ? shiftNameRefRelativeAxes(stored, row, col) : null;
      const area = ref ? nameArea(ref, context.sheetName) : null;
      if (area) areas.push(area);
      return;
    }
    if (token.type === "tableref" && token.spec) {
      const table = token.table
        ? context.tables.find((candidate) => candidate.name.toUpperCase() === token.table!.toUpperCase())
        : context.tables.find((candidate) =>
            row >= candidate.range.r1 && row <= candidate.range.r2 + (candidate.totalsRow ? 1 : 0) &&
            col >= candidate.range.c1 && col <= candidate.range.c2);
      const area = table ? tableArea(table, token.spec, row, context) : null;
      if (area) areas.push(area);
    }
  });
  return areas;
}
