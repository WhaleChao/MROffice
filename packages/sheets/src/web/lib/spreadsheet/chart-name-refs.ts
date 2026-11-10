/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A chart series reference written as a defined name (`Sheet!ALL_IMP`), as
 * the range Excel plots for it.
 *
 * Excel lets a series' values, categories or name be a defined name, and a
 * dynamic one (`OFFSET(Sheet!$C$3,0,0,COUNTA(Sheet!$C:$C)-1)`) is the usual
 * way to grow a chart with its data: xlsx-populate XP004's combo chart reads
 * `ClickThroughRateSheet!ALL_IMP` and Excel draws C3:C7 (X-281). The chart's
 * reader only knows cell ranges, so it showed the "cannot verify the source"
 * placeholder. A name that is a range, or an OFFSET of one whose arguments
 * evaluate to numbers, becomes that range here; anything else stays as
 * written (and is refused as before).
 */
import type { CellValue } from "./formula-evaluator";
import { columnToLetters, parseCellRef } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { parseChartRefAreas, parseChartRefRange } from "@mrnote/core/web/lib/spreadsheet/chart.ts";

export interface ChartNameRefSource {
  /** The sheet a reference without a sheet name belongs to. */
  readonly defaultSheet: string;
  /** The definition of `name` seen from `sheetName` (a sheet-scoped name
   * first, then the workbook's), without a leading "=", or null. */
  readonly lookupName: (sheetName: string, name: string) => string | null;
  /** A cell's current value. */
  readonly cellValue: (sheetName: string, row: number, col: number) => CellValue;
  /** The last row index holding a cell of `sheetName` (a whole-column
   * argument reads no further). */
  readonly lastRow: (sheetName: string) => number;
}

const NAME_REF_RE = /^(?:(?:'((?:[^']|'')+)'|([^'!:]+))!)?([A-Za-z_\\][A-Za-z0-9_.\\]*)$/;

function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Top-level comma-separated arguments of a function call's argument text. */
function splitArguments(text: string): string[] | null {
  const parts: string[] = [];
  let depth = 0;
  let quote: string | null = null;
  let start = 0;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quote) {
      if (char === quote) quote = null;
      continue;
    }
    if (char === '"' || char === "'") quote = char;
    else if (char === "(") depth += 1;
    else if (char === ")") depth -= 1;
    else if (char === "," && depth === 0) {
      parts.push(text.slice(start, index).trim());
      start = index + 1;
    }
    if (depth < 0) return null;
  }
  if (depth !== 0 || quote) return null;
  parts.push(text.slice(start).trim());
  return parts;
}

// The OFFSET arguments a dynamic chart name uses are whole numbers: numbers,
// a cell, COUNTA / COUNT / ROWS / COLUMNS of a range or a whole column, and
// + - * / with parentheses. They are read here; anything else is refused as

// here it put the evaluator in the public bundle and every production web

type Token = { kind: "num"; value: number } | { kind: "op"; value: string } | { kind: "ref"; text: string } | { kind: "fn"; name: string };

const TOKEN_RE = /\s*(?:(\d+(?:\.\d+)?)|([+\-*/(),])|((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Za-z]{1,3}\$?\d*(?::\$?[A-Za-z]{1,3}\$?\d*)?|\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})|([A-Za-z]+)\s*(?=\())/y;

function tokens(text: string): Token[] | null {
  const out: Token[] = [];
  TOKEN_RE.lastIndex = 0;
  let index = 0;
  while (index < text.length) {
    if (!text.slice(index).trim()) break;
    TOKEN_RE.lastIndex = index;
    const match = TOKEN_RE.exec(text);
    if (!match || match.index !== index) return null;
    index = TOKEN_RE.lastIndex;
    if (match[1] !== undefined) out.push({ kind: "num", value: Number(match[1]) });
    else if (match[2] !== undefined) out.push({ kind: "op", value: match[2] });
    else if (match[3] !== undefined) out.push({ kind: "ref", text: match[3] });
    else if (match[4] !== undefined) out.push({ kind: "fn", name: match[4].toUpperCase() });
    else return null;
    if (out.length > 64) return null;
  }
  return out;
}

interface Area { readonly sheet: string; readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number }

function area(text: string, sheetName: string, source: ChartNameRefSource): Area | null {
  const bang = text.lastIndexOf("!");
  const sheet = bang >= 0 ? text.slice(0, bang).replace(/^'|'$/g, "").replace(/''/g, "'") : sheetName;
  const [from, to = from] = text.slice(bang + 1).replace(/\$/g, "").split(":");
  const whole = (part: string) => /^[A-Za-z]{1,3}$/.test(part);
  if (whole(from) && whole(to)) {
    const a = parseCellRef(`${from}1`);
    const b = parseCellRef(`${to}1`);
    if (!a || !b) return null;
    return { sheet, r1: 0, c1: Math.min(a.col, b.col), r2: Math.max(0, source.lastRow(sheet)), c2: Math.max(a.col, b.col) };
  }
  const a = parseCellRef(from);
  const b = parseCellRef(to);
  if (!a || !b) return null;
  return { sheet, r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
}

function values(span: Area, source: ChartNameRefSource): CellValue[] {
  const out: CellValue[] = [];
  const last = Math.min(span.r2, Math.max(0, source.lastRow(span.sheet)));
  for (let row = span.r1; row <= last; row += 1) {
    for (let col = span.c1; col <= span.c2; col += 1) out.push(source.cellValue(span.sheet, row, col));
  }
  return out;
}

function wholeNumber(expression: string, sheetName: string, source: ChartNameRefSource): number | null {
  const list = expression ? tokens(expression) : null;
  if (!list || !list.length) return null;
  let at = 0;
  const peek = () => list[at];
  const isOp = (value: string) => peek()?.kind === "op" && (peek() as { value: string }).value === value;
  const expr = (): number | null => {
    let left = term();
    while (left !== null && (isOp("+") || isOp("-"))) {
      const op = (list[at++] as { value: string }).value;
      const right = term();
      left = right === null ? null : op === "+" ? left + right : left - right;
    }
    return left;
  };
  const term = (): number | null => {
    let left = unary();
    while (left !== null && (isOp("*") || isOp("/"))) {
      const op = (list[at++] as { value: string }).value;
      const right = unary();
      left = right === null || (op === "/" && right === 0) ? null : op === "*" ? left * right : left / right;
    }
    return left;
  };
  const unary = (): number | null => {
    if (isOp("-")) { at += 1; const value = unary(); return value === null ? null : -value; }
    if (isOp("+")) { at += 1; return unary(); }
    return primary();
  };
  const primary = (): number | null => {
    const token = list[at++];
    if (!token) return null;
    if (token.kind === "num") return token.value;
    if (token.kind === "op" && token.value === "(") {
      const value = expr();
      if (!isOp(")")) return null;
      at += 1;
      return value;
    }
    if (token.kind === "ref") {
      const span = area(token.text, sheetName, source);
      if (!span || span.r1 !== span.r2 || span.c1 !== span.c2) return null;
      const value = source.cellValue(span.sheet, span.r1, span.c1);
      const number = typeof value === "number" ? value : typeof value === "string" && value.trim() !== "" ? Number(value) : NaN;
      return Number.isFinite(number) ? number : null;
    }
    if (token.kind === "fn" && ["COUNTA", "COUNT", "ROWS", "COLUMNS"].includes(token.name)) {
      if (!isOp("(")) return null;
      at += 1;
      const ref = list[at++];
      if (!ref || ref.kind !== "ref" || !isOp(")")) return null;
      at += 1;
      const span = area(ref.text, sheetName, source);
      if (!span) return null;
      if (token.name === "ROWS") return span.r2 - span.r1 + 1;
      if (token.name === "COLUMNS") return span.c2 - span.c1 + 1;
      const cells = values(span, source);
      return token.name === "COUNT"
        ? cells.filter((value) => typeof value === "number").length
        : cells.filter((value) => value !== null && value !== undefined && value !== "").length;
    }
    return null;
  };
  const value = expr();
  if (value === null || at !== list.length || !Number.isFinite(value)) return null;
  return Math.trunc(value);
}

function rangeText(sheetName: string, r1: number, c1: number, r2: number, c2: number): string {
  const corner = (row: number, col: number) => `$${columnToLetters(col)}$${row + 1}`;
  const area = r1 === r2 && c1 === c2 ? corner(r1, c1) : `${corner(r1, c1)}:${corner(r2, c2)}`;
  return `${quoteSheet(sheetName)}!${area}`;
}

/** `text` as a cell range when it is a defined name Excel plots as one;
 * otherwise `text` unchanged. */
export function resolveChartNameRef(text: string | undefined, source: ChartNameRefSource): string | undefined {
  if (!text) return text;
  const trimmed = text.trim();
  if (parseChartRefAreas(trimmed)) return text;
  const match = NAME_REF_RE.exec(trimmed);
  if (!match) return text;
  const scope = (match[1] ? match[1].replace(/''/g, "'") : match[2])?.trim() || source.defaultSheet;
  const definition = source.lookupName(scope, match[3])?.trim().replace(/^=/, "");
  if (!definition) return text;
  const direct = parseChartRefRange(definition);
  if (direct) {
    const sheet = direct.sheetName ?? scope;
    return rangeText(sheet, direct.range.r1, direct.range.c1, direct.range.r2, direct.range.c2);
  }
  const offset = /^OFFSET\s*\(([\s\S]*)\)$/i.exec(definition);
  const args = offset ? splitArguments(offset[1]) : null;
  if (!args || args.length < 3 || args.length > 5) return text;
  const base = parseChartRefRange(args[0]);
  if (!base) return text;
  const sheet = base.sheetName ?? scope;
  const rows = wholeNumber(args[1], sheet, source);
  const cols = wholeNumber(args[2], sheet, source);
  const height = args[3] ? wholeNumber(args[3], sheet, source) : base.range.r2 - base.range.r1 + 1;
  const width = args[4] ? wholeNumber(args[4], sheet, source) : base.range.c2 - base.range.c1 + 1;
  if (rows === null || cols === null || height === null || width === null || height < 1 || width < 1) return text;
  const r1 = base.range.r1 + rows;
  const c1 = base.range.c1 + cols;
  if (r1 < 0 || c1 < 0) return text;
  return rangeText(sheet, r1, c1, r1 + height - 1, c1 + width - 1);
}
