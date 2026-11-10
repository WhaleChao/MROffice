/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The number format a formula's cell takes from the formula itself.
 *
 * Excel formats the result of a formula entered into a General cell from what
 * the formula DOES, not merely from the first cell it happens to name.
 * Microsoft Support, "Cell linked to text-formatted cell shows formula not
 * value" (KB 135506), gives the underlying rule — "when you link cells in
 * Microsoft Excel, the format of the original cell is applied to the linked
 * cell" — and AutoSum ("Use AutoSum to sum numbers") enters such a formula.
 * The shape of the propagation was MEASURED in Excel, with A1 a date, A2 0%,
 * A3 #,##0.00 and A4 General:
 *
 *   =DATE(..)                              system short date
 *   =TIME(..)                              h:mm AM/PM
 *   =TODAY()                               date
 *   =NOW()                                 date + time
 *   =EDATE / EOMONTH / WORKDAY /
 *   =DATEVALUE / TIMEVALUE                 General (the serial shows)
 *   =A1  =A1+1  =A1+TIME(1,0,0)            date
 *   =A1-A1                                 General
 *   =A2  =A2+A4                            0%
 *   =A2*2  =A4*A2  =A3/A4  =A2/A2          General
 *   =A3+A4  =A4+A3                         #,##0.00
 *   =SUM(A3:A4)  =SUM(A4,A3)               #,##0.00
 *   =ROUND(A3,0)                           #,##0.00
 *   =IF(TRUE,A1,0)  =MAX(A1,A1)            General
 *   =YEAR(A1)  =A1&""  =DOLLAR(..)         General
 *   =PMT(..)                               currency, viewer's region symbol
 *
 * So: `+` and `-` pass on the first NON-General operand whichever side it is
 * on (with date − date falling back to General, because the difference is a
 * count of days); `*`, `/`, `^` and `&` pass nothing on; SUM and the ROUND
 * family inherit from the cells they read; most other functions do not;
 * DATE/TIME/TODAY/NOW and the financial functions imply a format of their own.
 * Where the measurement is silent this answers General rather than guess.
 *
 * Two limits keep the rule from doing harm:
 *  - a cell that already has its own number format keeps it;
 *  - Text (`@`) is never passed on. KB 135506 documents it as the cause of a
 *    formula showing as text instead of its value, which is a trap, not a
 *    presentation.
 */

import { normalizeNumberFormatCode, numberFormatFamily } from "./number-format";
import {
  extractPrecedents,
  extractSheetQualifiedPrecedents,
} from "./sheet-interactions";

/** Built-ins 14 and 22, drawn in the viewer's regional short date. */
const SYSTEM_SHORT_DATE = "m/d/yyyy";
const SYSTEM_SHORT_DATE_TIME = "m/d/yyyy h:mm";
/** Built-in 18, the regional short time. */
const SYSTEM_TIME = "h:mm AM/PM";
/** Built-in 7, the regional currency with decimals (see entry-parse's regions). */
const ZH_CURRENCY = "¥#,##0.00;[Red]¥-#,##0.00";
const EN_CURRENCY = "$#,##0.00_);[Red]($#,##0.00)";

/**
 * The format to give a cell whose own format is `destinationFormat` when its
 * formula's first reference has `referencedFormat`, or null to leave it alone.
 */
export function inheritedNumberFormat(
  referencedFormat: unknown,
  destinationFormat: unknown,
): string | null {
  if (normalizeNumberFormatCode(destinationFormat) !== "") return null;
  return passedOnFormat(referencedFormat);
}

/** A cell's format as something to pass on: General and Text pass nothing. */
function passedOnFormat(format: unknown): string | null {
  const code = normalizeNumberFormatCode(format);
  if (!code || code === "@") return null;
  return code;
}

/**
 * The first cell `formula` references — the top-left cell of its first range —
 * or null when it names none. `=SUM(C4:C10)` reads C4. A formula that names
 * another sheet answers null: which reference comes first is then not known
 * here, and a wrong format is worse than General.
 */
export function firstReferencedCell(
  formula: string,
): { row: number; col: number } | null {
  if (extractSheetQualifiedPrecedents(formula, 1).length > 0) return null;
  return extractPrecedents(formula, 1)[0] ?? null;
}

/**
 * Functions whose result carries a format of its own, whatever they read.
 * DATE/TODAY show the system short date, NOW the short date with time, TIME
 * the short time, and the financial functions an amount in the viewer's
 * currency (MEASURED: =PMT(..) in a General cell shows as currency). Only the
 * functions actually measured are listed — the rest stay General.
 */
function ownNumberFormat(name: string, locale: string): string | null {
  switch (name) {
    case "DATE":
    case "TODAY":
      return SYSTEM_SHORT_DATE;
    case "NOW":
      return SYSTEM_SHORT_DATE_TIME;
    case "TIME":
      return SYSTEM_TIME;
    case "PMT":
    case "PV":
    case "FV":
    case "NPV":
    case "IPMT":
    case "PPMT":
      return locale.toLowerCase().startsWith("zh") ? ZH_CURRENCY : EN_CURRENCY;
    default:
      return null;
  }
}

/**
 * Functions that pass on the format of the cells they read. SUM is a running
 * total of its arguments and the ROUND family re-presents one — MEASURED:
 * =SUM(A3:A4), =SUM(A4,A3) and =ROUND(A3,0) all show #,##0.00, while
 * =MAX(A1,A1), =IF(TRUE,A1,0) and =YEAR(A1) stay General.
 */
const INHERITING_FUNCTIONS = new Set(["SUM", "ROUND", "ROUNDUP", "ROUNDDOWN"]);

/** True for a code Excel draws as a date or a clock time. */
function isDateOrTimeFormat(code: string): boolean {
  const family = numberFormatFamily(code);
  return family === "date" || family === "time";
}

/** A single A1 cell reference and nothing else. */
const BARE_CELL_REF_RE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}$/;

/** The whole text is one function call: its name and its argument list. */
function wholeCall(text: string): { name: string; args: string } | null {
  const head = /^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/.exec(text);
  if (!head || !text.endsWith(")")) return null;
  let depth = 0;
  for (let i = head[0].length - 1; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      i++;
      while (i < text.length && text[i] !== '"') i++;
      continue;
    }
    if (ch === "(") depth++;
    else if (ch === ")") {
      depth--;
      if (depth === 0) {
        return i === text.length - 1
          ? { name: head[1].toUpperCase(), args: text.slice(head[0].length, i) }
          : null;
      }
    }
  }
  return null;
}

/** An operator after which a `+`/`-` is a SIGN rather than an addition. */
const SIGN_POSITION = new Set(["+", "-", "*", "/", "^", "&", "(", ",", "<", ">", "="]);

/**
 * The formula body's TOP-LEVEL `+`/`-` operands, or null when it carries an
 * operator Excel passes no format through (`*`, `/`, `^`, `&`, `%`, or a
 * comparison, whose result is a logical value).
 */
function additiveOperands(
  body: string,
): { operands: string[]; hasMinus: boolean } | null {
  const operands: string[] = [];
  let current = "";
  let depth = 0;
  let hasMinus = false;
  let previous = "";
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"' || ch === "'") {
      const quote = ch;
      current += ch;
      i++;
      while (i < body.length && body[i] !== quote) current += body[i++];
      if (i < body.length) current += body[i];
      previous = quote;
      continue;
    }
    if (ch === "(" || ch === "{" || ch === "[") depth++;
    else if (ch === ")" || ch === "}" || ch === "]") depth--;
    else if (depth === 0) {
      if ("*/^&%=<>".includes(ch)) return null;
      if ((ch === "+" || ch === "-") && previous !== "" && !SIGN_POSITION.has(previous)) {
        operands.push(current);
        current = "";
        if (ch === "-") hasMinus = true;
        previous = ch;
        continue;
      }
    }
    current += ch;
    if (ch.trim() !== "") previous = ch;
  }
  operands.push(current);
  return { operands, hasMinus };
}

/** The first non-General format among the cells `args` reads. */
function referencedFormat(
  args: string,
  formatAt: (row: number, col: number) => unknown,
): string | null {
  const formula = `=${args}`;
  if (extractSheetQualifiedPrecedents(formula, 1).length > 0) return null;
  for (const cell of extractPrecedents(formula, 64)) {
    const code = passedOnFormat(formatAt(cell.row, cell.col));
    if (code) return code;
  }
  return null;
}

/** The format one `+`/`-` operand contributes, or null for General. */
function operandNumberFormat(
  text: string,
  formatAt: (row: number, col: number) => unknown,
  locale: string,
): string | null {
  const operand = text.trim();
  if (!operand) return null;
  if (BARE_CELL_REF_RE.test(operand)) {
    const cell = firstReferencedCell(`=${operand}`);
    return cell ? passedOnFormat(formatAt(cell.row, cell.col)) : null;
  }
  const call = wholeCall(operand);
  if (!call) return null;
  const own = ownNumberFormat(call.name, locale);
  if (own) return own;
  return INHERITING_FUNCTIONS.has(call.name)
    ? referencedFormat(call.args, formatAt)
    : null;
}

/**
 * The format a formula TYPED into a cell gives it, or null to leave the cell
 * as it is (always the case for a cell that already has a format of its own).
 * `locale` picks the region's currency for the financial functions.
 */
export function enteredFormulaNumberFormat(
  formula: string,
  formatAt: (row: number, col: number) => unknown,
  destinationFormat: unknown,
  locale = "en-US",
): string | null {
  if (normalizeNumberFormatCode(destinationFormat) !== "") return null;
  const text = (formula ?? "").trim();
  if (!text.startsWith("=")) return null;
  const body = text.slice(1).trim();
  if (!body) return null;
  // A formula naming another sheet: which reference comes first is not known
  // here, and a wrong format is worse than General.
  if (extractSheetQualifiedPrecedents(text, 1).length > 0) return null;

  const split = additiveOperands(body);
  if (!split) return null;
  const formats = split.operands.map((operand) =>
    operandNumberFormat(operand, formatAt, locale),
  );
  const first = formats.find((code) => code !== null) ?? null;
  if (first === null) return null;
  // Date − date is a count of days, not a date — MEASURED: =A1-A1 on a
  // date-formatted A1 shows General.
  if (
    split.hasMinus &&
    formats.filter((code) => code !== null && isDateOrTimeFormat(code)).length >= 2
  ) {
    return null;
  }
  return first;
}
