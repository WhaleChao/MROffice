/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The AcroForm scripting subset an order form actually uses.
 *
 * A PDF with computed totals carries JavaScript in each field's additional
 * actions: /AA /C recalculates, /AA /F formats for display. Without running any
 * of it, a filled form shows the values the author last saved — a subtotal that
 * does not move when the quantity does. Silent wrong numbers are worse than
 * blank ones.
 *
 * Rather than embed a JavaScript engine, this recognises the shapes Acrobat's
 * own form wizard emits and nothing else:
 *
 *   AFSimple_Calculate("SUM", new Array("qty", "price"))
 *   AFNumber_Format(2, 0, 0, 0, "¥", true)
 *   AFDate_FormatEx("yyyy-mm-dd")
 *   event.value = this.getField("qty").value * this.getField("price").value
 *
 * Anything it does not recognise returns null, and the caller leaves the field
 * exactly as it was. That is the important property: an unrecognised script must
 * never produce a number, because a plausible-looking wrong total is the failure
 * this exists to prevent. No eval, no Function, no sandbox to escape.
 */

/** The AFSimple_Calculate operations Acrobat's wizard emits. */
export type SimpleCalcOp = "SUM" | "PRD" | "AVG" | "MIN" | "MAX";

export interface SimpleCalculation {
  kind: "simple";
  op: SimpleCalcOp;
  /** Field names the operation reads, in the order the script lists them. */
  fields: string[];
}

/** A parsed arithmetic expression over field values and literals. */
export type ArithNode =
  | { kind: "field"; name: string }
  | { kind: "number"; value: number }
  | { kind: "binary"; op: "+" | "-" | "*" | "/"; left: ArithNode; right: ArithNode };



export type CalculationValue = number | string;

export interface ExpressionCalculation {
  kind: "expression";
  expression: ArithNode;
}

export type Calculation = SimpleCalculation | ExpressionCalculation;

const SIMPLE_OPS: Record<string, SimpleCalcOp> = {
  SUM: "SUM",
  PRD: "PRD",
  AVG: "AVG",
  MIN: "MIN",
  MAX: "MAX",
};

const LINE_TERMINATOR = /[\n\r\u2028\u2029]/;










function stripComments(source: string): string {
  let out = "";
  let at = 0;
  while (at < source.length) {
    const ch = source[at]!;
    const next = source[at + 1];
    if (ch === '"' || ch === "'" || ch === "`") {
      let end = at + 1;
      while (end < source.length && source[end] !== ch) {
        if (ch !== "`" && LINE_TERMINATOR.test(source[end]!)) break;
        end += source[end] === "\\" ? 2 : 1;
      }
      const stop = Math.min(source[end] === ch ? end + 1 : end, source.length);
      out += source.slice(at, stop);
      at = stop;
      continue;
    }
    if (ch === "/" && next === "*") {
      const close = source.indexOf("*/", at + 2);
      if (close < 0) return out + source.slice(at);
      out += " ";
      at = close + 2;
      continue;
    }
    if (ch === "/" && next === "/") {
      while (at < source.length && !LINE_TERMINATOR.test(source[at]!)) at += 1;
      out += " ";
      continue;
    }
    out += ch;
    at += 1;
  }
  return out;
}

/**
 * A quoted string literal's content, or null for an escape outside the
 * subset. \\uXXXX, an escaped quote and an escaped backslash are what a
 * field name needs; anything else (\\n, \\x41, a line continuation) names
 * nothing this module can vouch for.
 */
function decodeStringLiteral(raw: string): string | null {
  let body = "";
  for (let at = 1; at < raw.length - 1; ) {
    const ch = raw[at]!;
    if (ch !== "\\") {
      body += ch;
      at += 1;
      continue;
    }
    const next = raw[at + 1];
    const hex = raw.slice(at + 2, at + 6);
    if (next === "u" && /^[0-9a-fA-F]{4}$/.test(hex) && at + 6 <= raw.length - 1) {
      body += String.fromCharCode(Number.parseInt(hex, 16));
      at += 6;
      continue;
    }
    if ((next !== "\\" && next !== '"' && next !== "'") || at + 2 > raw.length - 1) return null;
    body += next;
    at += 2;
  }
  return body;
}

/** Read a quoted string literal's content; "" when it is outside the subset. */
function readStringLiteral(raw: string): string {
  return decodeStringLiteral(raw) ?? "";
}

const SIMPLE_CALL_HEAD =
  /^AFSimple_Calculate\s*\(\s*("[A-Z]+"|'[A-Z]+')\s*,\s*(?:new\s+Array\s*\(|\[)/;
const JS_SPACE = /\s/;













function readFieldNameList(
  text: string,
  start: number,
  close: ")" | "]",
): { names: string[]; end: number } | null {
  const names: string[] = [];
  let at = start;
  const skipSpace = () => {
    while (at < text.length && JS_SPACE.test(text[at]!)) at += 1;
  };
  skipSpace();
  if (text[at] === close) return { names, end: at + 1 };
  for (;;) {
    skipSpace();
    const quote = text[at];
    if (quote !== '"' && quote !== "'") return null;
    let body = "";
    at += 1;
    for (;;) {
      const ch = text[at];
      if (ch === undefined || ch === "\n" || ch === "\r") return null;
      if (ch === quote) break;
      if (ch === "\\") {
        const next = text[at + 1];
        const hex = text.slice(at + 2, at + 6);
        if (next === "u" && /^[0-9a-fA-F]{4}$/.test(hex)) {
          body += String.fromCharCode(Number.parseInt(hex, 16));
          at += 6;
          continue;
        }
        // Only the escapes a field name needs; anything else (\n, \x41, a
        // line continuation) is outside the subset.
        if (next !== "\\" && next !== '"' && next !== "'") return null;
        body += next;
        at += 2;
        continue;
      }
      body += ch;
      at += 1;
    }
    at += 1;
    names.push(body);
    skipSpace();
    if (text[at] === ",") {
      at += 1;
      continue;
    }
    if (text[at] === close) return { names, end: at + 1 };
    return null;
  }
}

/**
 * Parse a field's /AA /C calculation script.
 *
 * Recognises AFSimple_Calculate and a single `event.value = <arithmetic>`
 * assignment over `this.getField("name").value` terms. Returns null for
 * anything else — including a script that merely LOOKS arithmetic but contains
 * a call, a conditional or an array index.
 */
export function parseCalculationScript(source: unknown): Calculation | null {
  if (typeof source !== "string" || source.trim() === "") return null;
  const text = stripComments(source).trim();

  const head = SIMPLE_CALL_HEAD.exec(text);
  if (head) {
    const op = SIMPLE_OPS[readStringLiteral(head[1]).toUpperCase()];
    if (!op) return null;
    const close = head[0].endsWith("[") ? "]" : ")";
    const list = readFieldNameList(text, head[0].length, close);
    if (!list) return null;
    if (!/^\s*\)\s*;?$/.test(text.slice(list.end))) return null;
    const fields = list.names.filter((name) => name !== "");
    if (fields.length === 0) return null;
    return { kind: "simple", op, fields };
  }

  const assignment = /^event\.value\s*=\s*([^;]+?)\s*;?$/.exec(text);
  if (!assignment) return null;
  const expression = parseArithmetic(assignment[1]);
  return expression ? { kind: "expression", expression } : null;
}

/** Bounds that keep a hostile script from exhausting the stack; a real form's
 *  expression is a few dozen tokens deep at most. The server's port
 *  (form_calculation.py) applies the same bounds, so both refuse alike. */
const MAX_ARITHMETIC_TOKENS = 4096;
const MAX_ARITHMETIC_NESTING = 100;

/** A quoted literal on one line; an escaped quote does not end it. */
const STRING_LITERAL = `"(?:[^"\\\\\\n\\r]|\\\\.)*"|'(?:[^'\\\\\\n\\r]|\\\\.)*'`;
const FIELD_REFERENCE = new RegExp(
  `^this\\.getField\\s*\\(\\s*(${STRING_LITERAL})\\s*\\)\\s*\\.value$`,
);
/** A JavaScript numeric literal as source text: a radix form, or a decimal
 *  with an optional fraction and exponent. */
const NUMERIC_LITERAL =
  "0[xX][0-9a-fA-F]+|0[oO][0-7]+|0[bB][01]+|(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][+-]?\\d+)?";
const NUMERIC_LITERAL_TOKEN = new RegExp(`^(?:${NUMERIC_LITERAL})$`);









function numericLiteralValue(token: string): number | null {
  if (/^0[xXoObB]/.test(token)) return Number(token);
  if (/^0\d/.test(token)) {
    if (!/^\d+$/.test(token)) return null;
    return /^[0-7]+$/.test(token) ? Number.parseInt(token, 8) : Number(token);
  }
  return Number(token);
}

/**
 * Parse the arithmetic subset: field references, numeric literals, + - * / and
 * parentheses. Deliberately tiny — a recursive-descent parser over a token list,
 * so an unexpected token fails the whole parse instead of being skipped.
 */
export function parseArithmetic(source: string): ArithNode | null {
  const tokens: string[] = [];
  const pattern = new RegExp(
    `this\\.getField\\s*\\(\\s*(?:${STRING_LITERAL})\\s*\\)\\s*\\.value|${NUMERIC_LITERAL}|[-+*/()]|\\S`,
    "g",
  );
  let match: RegExpExecArray | null = pattern.exec(source);
  while (match !== null) {
    tokens.push(match[0]);
    if (tokens.length > MAX_ARITHMETIC_TOKENS) return null;
    match = pattern.exec(source);
  }
  if (tokens.length === 0) return null;

  let index = 0;
  let nesting = 0;
  const peek = () => tokens[index];
  const take = () => tokens[index++];

  const parsePrimary = (): ArithNode | null => {
    const token = take();
    if (token === undefined) return null;
    if (token === "(" || token === "-") {
      if (nesting >= MAX_ARITHMETIC_NESTING) return null;
      nesting += 1;
      const node = parseNested(token);
      nesting -= 1;
      return node;
    }
    const field = FIELD_REFERENCE.exec(token);
    if (field) {
      const name = decodeStringLiteral(field[1]!);
      return name ? { kind: "field", name } : null;
    }
    if (NUMERIC_LITERAL_TOKEN.test(token)) {
      const value = numericLiteralValue(token);
      return value === null ? null : { kind: "number", value };
    }
    return null;
  };

  const parseNested = (token: "(" | "-"): ArithNode | null => {
    if (token === "(") {
      const inner = parseSum();
      if (!inner || take() !== ")") return null;
      return inner;
    }
    const operand = parsePrimary();
    return operand
      ? { kind: "binary", op: "-", left: { kind: "number", value: 0 }, right: operand }
      : null;
  };

  const parseProduct = (): ArithNode | null => {
    let left = parsePrimary();
    if (!left) return null;
    while (peek() === "*" || peek() === "/") {
      const op = take() as "*" | "/";
      const right = parsePrimary();
      if (!right) return null;
      left = { kind: "binary", op, left, right };
    }
    return left;
  };

  const parseSum = (): ArithNode | null => {
    let left = parseProduct();
    if (!left) return null;
    while (peek() === "+" || peek() === "-") {
      const op = take() as "+" | "-";
      const right = parseProduct();
      if (!right) return null;
      left = { kind: "binary", op, left, right };
    }
    return left;
  };

  const tree = parseSum();
  // Any leftover token means the script said something this parser does not
  // model; producing a number from half of it would be a guess.
  return tree && index === tokens.length ? tree : null;
}

/** A field's value as a number; a blank or unparseable value reads as 0, which
 *  is what Acrobat's own AFSimple_Calculate does. */
export function numericFieldValue(
  value: unknown,
  separatorStyle?: 0 | 1 | 2 | 3,
): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value !== "string") return 0;
  // Tolerate the separators a formatted field displays. Styles 2/3 use a
  // decimal comma; treating it as a grouping separator turns 1.234,50 into
  // 1.2345, so the field's own format script remains the parsing authority.
  const negativeParentheses = /^\s*\(.*\)\s*$/.test(value);
  let cleaned = value.replace(/[^0-9.,\-+eE]/g, "");
  if (separatorStyle === 2) {
    // Writer values remain canonical dot-decimal strings even when the widget
    // displays a decimal comma. Only a value that actually contains a comma is
    // display-formatted and therefore allowed to reinterpret dots as groups.
    if (cleaned.includes(",")) {
      cleaned = cleaned.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (separatorStyle === 3) {
    if (cleaned.includes(",")) cleaned = cleaned.replace(/,/g, ".");
  } else {
    cleaned = cleaned.replace(/,/g, "");
  }
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return 0;
  return negativeParentheses ? -Math.abs(parsed) : parsed;
}









function simpleCalculationNumber(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (typeof value !== "string") return 0;
  const parsed = Number.parseFloat(value.trim().replace(",", "."));
  return Number.isFinite(parsed) ? parsed : 0;
}

/**
 * What a script's `getField(name).value` reads: the number when the whole
 * text is one (its first comma read as a decimal point), otherwise the text
 * itself, and "" for a blank field. A date field stays text. Null for a value
 * a script cannot read as either (a multi-selection, an unknown checkbox
 * state).
 */
function scriptFieldValue(value: unknown, textual: boolean): CalculationValue | null {
  if (typeof value === "number") return value;
  if (value === undefined || value === null) return "";
  if (typeof value !== "string") return null;
  if (value === "" || textual) return value;
  const normalised = value.trim().replace(",", ".");
  return Number.isNaN(Number(normalised)) ? value : Number.parseFloat(normalised);
}

/** The longest text a calculation may produce (a field's own text bound). */
const MAX_CALCULATED_TEXT = 4000;

/** Whether a field already holds `next`, so writing it would change nothing. A
 *  blank field reads as 0 here, as it always has: a blank total over blank
 *  inputs is not rewritten to "0". */
function holdsCalculatedValue(current: unknown, next: CalculationValue): boolean {
  if (typeof next === "string") {
    return typeof current === "number" ? String(current) === next : current === next;
  }
  if (typeof current === "number") return current === next;
  if (current === undefined || current === null) return next === 0;
  if (typeof current !== "string") return false;
  const text = current.trim().replace(",", ".");
  return text === "" ? next === 0 : Number(text) === next;
}

/** What `evaluateCalculation` knows beyond the values themselves. */
export interface CalculationContext {
  


  known?: ReadonlySet<string>;
  /** Fields whose value a script reads as text whatever it looks like (date
   *  fields). */
  textual?: ReadonlySet<string>;
}











export function expandCalculationReference(
  name: string,
  known: Iterable<string>,
): string[] {
  const children: string[] = [];
  const prefix = `${name}.`;
  for (const candidate of known) {
    if (candidate === name) return [name];
    if (candidate.startsWith(prefix)) children.push(candidate);
  }
  return children.length > 0 ? children : [name];
}







function roundSimpleResult(result: number): number {
  return Math.round(1e6 * result) / 1e6;
}

/**
 * Evaluate a parsed calculation against the current field values. Returns
 * null — no value — for a divide by zero, any non-finite result, or an
 * expression that reads something that is not a field.
 *
 * AFSimple_Calculate follows AForm: a name that is no field (and no parent of
 * one) is skipped, an existing blank field counts as 0, and a list that names
 * no field at all gives 0. An `event.value` expression follows the language:
 * `+` joins when either side is text, the other operators read both sides as
 * numbers.
 */
export function evaluateCalculation(
  calc: Calculation,
  values: ReadonlyMap<string, unknown>,
  context: CalculationContext = {},
): CalculationValue | null {
  const exists = (name: string) => context.known?.has(name) ?? values.has(name);
  if (calc.kind === "simple") {
    const numbers = calc.fields
      .flatMap((name) =>
        exists(name)
          ? [name]
          : expandCalculationReference(name, context.known ?? values.keys()).filter(exists),
      )
      .map((name) => simpleCalculationNumber(values.get(name)));
    if (numbers.length === 0) return 0;
    let result: number;
    switch (calc.op) {
      case "SUM":
        result = numbers.reduce((a, b) => a + b, 0);
        break;
      case "PRD":
        result = numbers.reduce((a, b) => a * b, 1);
        break;
      case "AVG":
        result = numbers.reduce((a, b) => a + b, 0) / numbers.length;
        break;
      case "MIN":
        result = Math.min(...numbers);
        break;
      case "MAX":
        result = Math.max(...numbers);
        break;
      default:
        return null;
    }
    // A sum that overflows is not a total anyone can use.
    if (!Number.isFinite(result)) return null;
    const rounded = roundSimpleResult(result);
    return Number.isFinite(rounded) ? rounded : null;
  }
  const walk = (node: ArithNode): CalculationValue | null => {
    if (node.kind === "number") return node.value;
    if (node.kind === "field") {
      if (!exists(node.name)) return null;
      return scriptFieldValue(values.get(node.name), context.textual?.has(node.name) ?? false);
    }
    const left = walk(node.left);
    const right = walk(node.right);
    if (left === null || right === null) return null;
    if (node.op === "+") {
      if (typeof left === "number" && typeof right === "number") return left + right;
      const joined = String(left) + String(right);
      return joined.length <= MAX_CALCULATED_TEXT ? joined : null;
    }
    const a = Number(left);
    const b = Number(right);
    switch (node.op) {
      case "-":
        return a - b;
      case "*":
        return a * b;
      case "/":
        // Acrobat shows a blank rather than Infinity for a divide by zero.
        return b === 0 ? null : a / b;
    }
  };
  const result = walk(calc.expression);
  if (typeof result === "string") return result;
  return result !== null && Number.isFinite(result) ? result : null;
}

/** A parsed AFNumber_Format / AFPercent_Format specification. */
export interface NumberFormatSpec {
  kind: "number" | "percent";
  decimals: number;
  /** 0: 1,234.56 · 1: 1234.56 · 2: 1.234,56 · 3: 1234,56 */
  separatorStyle: 0 | 1 | 2 | 3;
  /** 0: -1 · 1: red -1 · 2: (1) · 3: red (1) */
  negativeStyle: 0 | 1 | 2 | 3;
  currency: string;
  /** True when the symbol goes before the number. */
  prependCurrency: boolean;
}

const SUPPORTED_DATE_FORMAT_PATTERNS = [
  "m/d/yy",
  "m/d/yyyy",
  "mm/dd/yy",
  "mm/dd/yyyy",
  "d/m/yy",
  "d/m/yyyy",
  "dd/mm/yy",
  "dd/mm/yyyy",
  "yy-mm-dd",
  "yyyy-mm-dd",
  "yy/mm/dd",
  "yyyy/mm/dd",
  "d-mmm-yy",
  "d-mmm-yyyy",
  "dd-mmm-yy",
  "dd-mmm-yyyy",
  "mmm d, yyyy",
  "mmmm d, yyyy",
] as const;

export type DateFormatPattern =
  (typeof SUPPORTED_DATE_FORMAT_PATTERNS)[number];

export interface DateFormatSpec {
  kind: "date";
  pattern: DateFormatPattern;
}

export type FieldFormatSpec = NumberFormatSpec | DateFormatSpec;

/** Parse a field's /AA /F format script. */
export function parseFormatScript(source: unknown): FieldFormatSpec | null {
  if (typeof source !== "string") return null;
  const text = stripComments(source).trim();
  const dateCall =
    /^AFDate_FormatEx\s*\(\s*("[^"\r\n]*"|'[^'\r\n]*')\s*\)\s*;?$/.exec(
      text,
    );
  if (dateCall) {
    const pattern = readStringLiteral(dateCall[1]).toLowerCase();
    if (
      (SUPPORTED_DATE_FORMAT_PATTERNS as readonly string[]).includes(pattern)
    ) {
      return { kind: "date", pattern: pattern as DateFormatPattern };
    }
    return null;
  }
  const head = /^AF(Number|Percent)_Format\s*\(/.exec(text);
  if (!head) return null;
  const call = readLiteralArguments(text, head[0].length);
  if (!call || !/^\s*;?\s*$/.test(text.slice(call.end))) return null;
  const args = call.args;
  const percent = head[1] === "Percent";
  // AFNumber_Format(nDec, sepStyle, negStyle, currStyle, strCurrency,
  // bCurrencyPrepend); AFPercent_Format(nDec, sepStyle[, bPercentPrepend]).
  // Any other shape is not this helper: report it rather than read the
  
  if (percent ? args.length < 2 || args.length > 3 : args.length !== 6) {
    return null;
  }
  const numberAt = (index: number) => {
    const arg = args[index];
    return arg?.kind === "number" ? arg.value : Number.NaN;
  };
  const decimals = numberAt(0);
  const sep = percent || args.length > 1 ? numberAt(1) : 0;
  const neg = percent ? 0 : numberAt(2);
  if (!Number.isFinite(decimals) || decimals < 0 || decimals > 20) return null;
  if (!Number.isFinite(sep) || !Number.isFinite(neg)) return null;
  let currency = "";
  let prependCurrency = false;
  if (!percent) {
    const currencyArg = args[4]!;
    const prependArg = args[5]!;
    if (currencyArg.kind !== "string" || prependArg.kind === "string") {
      return null;
    }
    currency = currencyArg.value;
    prependCurrency =
      prependArg.kind === "boolean" ? prependArg.value : prependArg.value !== 0;
  }
  return {
    kind: percent ? "percent" : "number",
    decimals: Math.trunc(decimals),
    separatorStyle: ([0, 1, 2, 3].includes(sep) ? sep : 0) as 0 | 1 | 2 | 3,
    negativeStyle: ([0, 1, 2, 3].includes(neg) ? neg : 0) as 0 | 1 | 2 | 3,
    currency,
    prependCurrency,
  };
}

type LiteralArgument =
  | { kind: "string"; value: string }
  | { kind: "number"; value: number }
  | { kind: "boolean"; value: boolean };

/**
 * Read a call's argument list of plain literals (strings, numbers,
 * true/false) starting just past its `(`, up to and including the `)`.
 * A comma or parenthesis inside a quoted string belongs to the string; an
 * escaped quote does not end it. Anything else fails the parse.
 */
function readLiteralArguments(
  text: string,
  start: number,
): { args: LiteralArgument[]; end: number } | null {
  const args: LiteralArgument[] = [];
  let at = start;
  const skipSpace = () => {
    while (at < text.length && JS_SPACE.test(text[at]!)) at += 1;
  };
  skipSpace();
  if (text[at] === ")") return { args, end: at + 1 };
  for (;;) {
    skipSpace();
    const ch = text[at];
    if (ch === '"' || ch === "'") {
      let close = at + 1;
      while (close < text.length && text[close] !== ch) {
        if (text[close] === "\\") close += 1;
        if (text[close] === "\r" || text[close] === "\n") return null;
        close += 1;
      }
      if (close >= text.length) return null;
      const value = decodeStringLiteral(text.slice(at, close + 1));
      if (value === null) return null;
      args.push({ kind: "string", value });
      at = close + 1;
    } else {
      const literal = /^(?:-?(?:\d+(?:\.\d*)?|\.\d+)|true|false)/.exec(
        text.slice(at),
      );
      if (!literal) return null;
      const raw = literal[0];
      args.push(
        raw === "true" || raw === "false"
          ? { kind: "boolean", value: raw === "true" }
          : { kind: "number", value: Number(raw) },
      );
      at += raw.length;
    }
    skipSpace();
    if (text[at] === ",") {
      at += 1;
      continue;
    }
    if (text[at] === ")") return { args, end: at + 1 };
    return null;
  }
}

function numericValueForSpec(value: unknown, spec: NumberFormatSpec): number {
  const numeric = numericFieldValue(value, spec.separatorStyle);
  return spec.kind === "percent" && typeof value === "string" && value.includes("%")
    ? numeric / 100
    : numeric;
}

interface DateParts {
  year: number;
  month: number;
  day: number;
}

function validDateParts(
  year: number,
  month: number,
  day: number,
): DateParts | null {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1) {
    return null;
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
    ? { year, month, day }
    : null;
}

const ENGLISH_MONTHS = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december",
] as const;

/** Month names a date entry may use: English (what the AFDate helpers
 * read) and the reader's own locale, long and short. */
function monthNameTable(locale: string): Map<string, number> {
  const table = new Map<string, number>();
  const add = (name: string, month: number) => {
    const key = name.toLocaleLowerCase(locale).replace(/\.$/u, "").trim();
    if (key && !table.has(key)) table.set(key, month);
  };
  ENGLISH_MONTHS.forEach((name, index) => {
    add(name, index + 1);
    add(name.slice(0, 3), index + 1);
  });
  add("sept", 9);
  for (const width of ["long", "short"] as const) {
    let format: Intl.DateTimeFormat;
    try {
      format = new Intl.DateTimeFormat(locale, { month: width, timeZone: "UTC" });
    } catch {
      continue;
    }
    for (let month = 1; month <= 12; month += 1) {
      add(format.format(new Date(Date.UTC(2000, month - 1, 1))), month);
    }
  }
  return table;
}

function parseDateValue(
  value: unknown,
  pattern: DateFormatPattern,
  locale = "en-US",
): DateParts | null {
  if (typeof value !== "string") return null;
  const raw = value.trim();
  if (raw === "") return null;

  const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(raw);
  if (iso) {
    return validDateParts(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  }

  const tokens = pattern.match(/yyyy|mmmm|mmm|yy|mm|dd|m|d|[^ymd]+/g);
  if (!tokens) return null;
  // A pattern with a month name reads a month name, "Aug" or "August" alike:
  // the same day typed as a name used to skip the format entirely
  
  const months = tokens.some((token) => token === "mmm" || token === "mmmm")
    ? monthNameTable(locale)
    : null;
  const captures: Array<"year" | "month" | "day" | "monthName"> = [];
  const source = tokens
    .map((token) => {
      if (token === "yyyy") {
        captures.push("year");
        return "(\\d{4})";
      }
      if (token === "yy") {
        captures.push("year");
        return "(\\d{2})";
      }
      if (token === "mmmm" || token === "mmm") {
        captures.push("monthName");
        return "([^\\s\\d,./-]+\\.?)";
      }
      if (token === "mm" || token === "m") {
        captures.push("month");
        return "(\\d{1,2})";
      }
      if (token === "dd" || token === "d") {
        captures.push("day");
        return "(\\d{1,2})";
      }
      // Separators: any run of spaces stands for the pattern's space, and
      // the comma after a day may carry or drop its space.
      return token
        .replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        .replace(/ +/g, "\\s*");
    })
    .join("");
  const matched = new RegExp(`^${source}$`, "u").exec(raw);
  if (!matched) return null;
  const parts: Partial<DateParts> = {};
  for (const [index, capture] of captures.entries()) {
    const text = matched[index + 1] ?? "";
    if (capture === "monthName") {
      const month = months?.get(
        text.toLocaleLowerCase(locale).replace(/\.$/u, ""),
      );
      if (!month) return null;
      parts.month = month;
    } else {
      parts[capture] = Number(text);
    }
  }
  let year = parts.year ?? 0;
  if (year < 100) year += year < 50 ? 2000 : 1900;
  return validDateParts(year, parts.month ?? 0, parts.day ?? 0);
}

function localizedMonth(
  parts: DateParts,
  width: "short" | "long",
  locale: string,
): string {
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  try {
    return new Intl.DateTimeFormat(locale, {
      month: width,
      timeZone: "UTC",
    }).format(date);
  } catch {
    return new Intl.DateTimeFormat("en-US", {
      month: width,
      timeZone: "UTC",
    }).format(date);
  }
}

function formatDateWithSpec(
  value: unknown,
  spec: DateFormatSpec,
  locale: string,
): string {
  const raw = typeof value === "string" ? value : "";
  const parts = parseDateValue(value, spec.pattern, locale);
  if (!parts) return raw;
  return spec.pattern.replace(
    /yyyy|mmmm|mmm|yy|mm|dd|m|d/g,
    (token) => {
      switch (token) {
        case "yyyy":
          return String(parts.year).padStart(4, "0");
        case "yy":
          return String(parts.year % 100).padStart(2, "0");
        case "mmmm":
          return localizedMonth(parts, "long", locale);
        case "mmm":
          return localizedMonth(parts, "short", locale);
        case "mm":
          return String(parts.month).padStart(2, "0");
        case "m":
          return String(parts.month);
        case "dd":
          return String(parts.day).padStart(2, "0");
        default:
          return String(parts.day);
      }
    },
  );
}







export type PdfKeystrokeRule =
  | {
      kind: "special";
      /** AFSpecial_Keystroke: 0 ZIP, 1 ZIP+4, 2 phone, 3 social security. */
      special: "zip" | "zip4" | "phone" | "ssn";
    }
  | { kind: "number"; decimalSeparator: "." | "," }
  | { kind: "date"; pattern: DateFormatPattern }
  /** AFTime_KeystrokeEx: a time of day in the pattern (HH 24-hour, h with
   * tt am/pm; ss seconds). */
  | { kind: "time"; pattern: string }
  /** An email address (the editor's 电子邮件 validation script). */
  | { kind: "email" }
  /** AFSpecial_KeystrokeEx mask: 9 digit, A letter, O letter or digit,
   * X any character, anything else itself. */
  | { kind: "mask"; mask: string };

const SPECIAL_KEYSTROKES = ["zip", "zip4", "phone", "ssn"] as const;
/** Digit counts each special entry accepts when it is complete. */
const SPECIAL_DIGITS: Record<(typeof SPECIAL_KEYSTROKES)[number], number[]> = {
  zip: [5],
  zip4: [9],
  phone: [7, 10],
  ssn: [9],
};
const SPECIAL_SEPARATORS = /[\s\-().+]/gu;

/** Parse a field's /AA /K keystroke script into a rule the reader checks. */
export function parseKeystrokeScript(source: unknown): PdfKeystrokeRule | null {
  if (typeof source !== "string") return null;
  const text = stripComments(source).trim();
  const head = /^(AFSpecial_Keystroke|AFSpecial_KeystrokeEx|AFNumber_Keystroke|AFPercent_Keystroke|AFDate_KeystrokeEx|AFTime_KeystrokeEx)\s*\(/.exec(
    text,
  );
  if (!head) return null;
  const call = readLiteralArguments(text, head[0].length);
  if (!call || !/^\s*;?\s*$/.test(text.slice(call.end))) return null;
  const [first, second] = call.args;
  switch (head[1]) {
    case "AFSpecial_Keystroke": {
      if (call.args.length !== 1 || first?.kind !== "number") return null;
      const special = SPECIAL_KEYSTROKES[first.value];
      return Number.isInteger(first.value) && special
        ? { kind: "special", special }
        : null;
    }
    case "AFSpecial_KeystrokeEx":
      return call.args.length === 1 &&
        first?.kind === "string" &&
        first.value.length > 0 &&
        first.value.length <= 256
        ? { kind: "mask", mask: first.value }
        : null;
    case "AFTime_KeystrokeEx":
      return call.args.length === 1 && first?.kind === "string" && /^(?:HH|h):MM(?::ss)?(?: tt)?$/u.test(first.value)
        ? { kind: "time", pattern: first.value }
        : null;
    case "AFDate_KeystrokeEx": {
      if (call.args.length !== 1 || first?.kind !== "string") return null;
      const pattern = first.value.toLowerCase();
      return (SUPPORTED_DATE_FORMAT_PATTERNS as readonly string[]).includes(
        pattern,
      )
        ? { kind: "date", pattern: pattern as DateFormatPattern }
        : null;
    }
    default: {
      // AFNumber_Keystroke(nDec, sepStyle, negStyle, currStyle, strCurrency,
      // bCurrencyPrepend); AFPercent_Keystroke(nDec, sepStyle).
      if (call.args.length < 2 || first?.kind !== "number") return null;
      if (second?.kind !== "number") return null;
      return {
        kind: "number",
        decimalSeparator: second.value === 2 || second.value === 3 ? "," : ".",
      };
    }
  }
}

function maskSlotAccepts(slot: string, ch: string): boolean {
  switch (slot) {
    case "9":
      return /^\p{Nd}$/u.test(ch);
    case "A":
      return /^\p{L}$/u.test(ch);
    case "O":
      return /^[\p{L}\p{Nd}]$/u.test(ch);
    case "X":
      return true;
    default:
      return ch === slot;
  }
}

/**
 * Whether `value` can still become a valid entry: what a keystroke may
 * leave in the field. A ZIP field takes "123" on the way to "12345" but
 * never "A"; the whole entry is judged at commit ({@link keystrokeCommitOk}).
 */
export function keystrokePartialOk(
  rule: PdfKeystrokeRule,
  value: string,
): boolean {
  if (value === "") return true;
  switch (rule.kind) {
    case "special": {
      if (/[^\p{Nd}\s\-().+]/u.test(value)) return false;
      const digits = value.replace(SPECIAL_SEPARATORS, "").length;
      return digits <= Math.max(...SPECIAL_DIGITS[rule.special]);
    }
    case "number": {
      const body = value.trim().replace(/%$/u, "").trim().replace(/^[-+]/u, "");
      const parts = body.split(rule.decimalSeparator);
      if (parts.length > 2) return false;
      const group = rule.decimalSeparator === "." ? "," : ".";
      const integer = parts[0]!.split(group).join("");
      return /^[0-9]*$/u.test(integer) && /^[0-9]*$/u.test(parts[1] ?? "");
    }
    case "date":
      return !/[\p{C}]/u.test(value);
    case "time":
      return /^[0-9:\s]*(?:[aApP][mM]?)?$/u.test(value);
    case "email":
      return !/[\p{C}\s]/u.test(value);
    case "mask": {
      const chars = Array.from(value);
      const slots = Array.from(rule.mask);
      if (chars.length > slots.length) return false;
      return chars.every((ch, index) => maskSlotAccepts(slots[index]!, ch));
    }
  }
}

/** Which message names what a keystroke rule needs. */
export function keystrokeRuleMessageKey(
  rule: PdfKeystrokeRule,
): "zip" | "zip4" | "phone" | "ssn" | "number" | "date" | "time" | "email" | "mask" {
  return rule.kind === "special" ? rule.special : rule.kind;
}

/** Whether `value` is a complete entry the rule accepts; empty clears. */
export function keystrokeCommitOk(
  rule: PdfKeystrokeRule,
  value: string,
  locale = "en-US",
): boolean {
  if (value.trim() === "") return true;
  if (!keystrokePartialOk(rule, value)) return false;
  switch (rule.kind) {
    case "special":
      return SPECIAL_DIGITS[rule.special].includes(
        value.replace(SPECIAL_SEPARATORS, "").length,
      );
    case "number": {
      const normalized = value
        .replace(/[\s%]/gu, "")
        .split(rule.decimalSeparator === "." ? "," : ".")
        .join("")
        .replace(",", ".");
      return normalized !== "" && Number.isFinite(Number(normalized));
    }
    case "date":
      return parseDateValue(value, rule.pattern, locale) !== null;
    case "time": {
      // 9:05, 09:05:30, 9:05 pm: hours in range for the clock the
      // pattern names, minutes and seconds under 60.
      const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([aApP][mM]?)?$/u.exec(value.trim());
      if (!match) return false;
      const twelve = rule.pattern.includes("tt");
      const hours = Number(match[1]);
      if (Boolean(match[4]) !== twelve) return false;
      if (twelve ? hours < 1 || hours > 12 : hours > 23) return false;
      return Number(match[2]) < 60 && (match[3] === undefined || Number(match[3]) < 60) &&
        (match[3] !== undefined) === rule.pattern.includes("ss");
    }
    case "email":
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(value.trim());
    case "mask":
      return Array.from(value).length === Array.from(rule.mask).length;
  }
}

/** Render a raw field value through a safely parsed format specification. */
export function formatWithSpec(
  value: unknown,
  spec: FieldFormatSpec,
  locale = "en-US",
): string {
  if (spec.kind === "date") return formatDateWithSpec(value, spec, locale);
  if (typeof value === "number" && !Number.isFinite(value)) return "";
  // AForm's AFNumber_Format and AFPercent_Format show nothing for a value
  // AFMakeNumber reads no number from -- blank, spaces, words -- as WPS and
  // Acrobat leave such a field blank; "0.00" read as a zero typed in (user
  // audit C0010). A calculation still counts it as 0 (numericFieldValue).
  if (typeof value === "string" && !/[0-9]/.test(value)) return "";
  const numeric = numericValueForSpec(value, spec);
  if (!Number.isFinite(numeric)) return "";
  const scaled = spec.kind === "percent" ? numeric * 100 : numeric;
  const negative = scaled < 0;
  const fixed = Math.abs(scaled).toFixed(spec.decimals);
  const [intPart, fracPart] = fixed.split(".");
  const groupWith = (sep: string) =>
    intPart.replace(/\B(?=(\d{3})+(?!\d))/g, sep);
  let body: string;
  switch (spec.separatorStyle) {
    case 1:
      body = fracPart ? `${intPart}.${fracPart}` : intPart;
      break;
    case 2:
      body = fracPart ? `${groupWith(".")},${fracPart}` : groupWith(".");
      break;
    case 3:
      body = fracPart ? `${intPart},${fracPart}` : intPart;
      break;
    default:
      body = fracPart ? `${groupWith(",")}.${fracPart}` : groupWith(",");
  }
  if (spec.kind === "percent") body = `${body}%`;
  if (spec.currency) {
    body = spec.prependCurrency ? `${spec.currency}${body}` : `${body}${spec.currency}`;
  }
  if (!negative) return body;
  // Styles 2 and 3 wrap a negative in parentheses; the colour of 1 and 3 is a
  // display concern the caller applies, not part of the string.
  return spec.negativeStyle === 2 || spec.negativeStyle === 3
    ? `(${body})`
    : `-${body}`;
}

/** One field's scripts, as read from its /AA dictionary. */
export interface FieldScripts {
  name: string;
  calculate?: unknown;
  format?: unknown;
}

export interface CalculationDependencyAnalysis {
  /** Fields that participate directly in one or more dependency cycles. */
  cycleFields: string[];
  /** Cycle fields plus calculated fields that transitively depend on them. */
  cycleAffectedFields: string[];
}

/** Why a calculated field could not be brought up to date. */
export type CalculationIncompleteReason =
  /** It sits in, or reads from, a calculation dependency cycle. */
  | "cycle"
  /** Its own script is outside the subset this module models. */
  | "unsupported-script"
  /** A calculated field it reads (directly or through other calculations)
   *  could not be brought up to date — an unsupported script, or no value —
   *  so its inputs are stale. */
  | "unsupported-input"
  /** Its own calculation produced no finite number, e.g. a divide by zero. */
  | "no-value";

/** A calculated field left on its existing value, which is therefore stale. */
export interface IncompleteCalculation {
  name: string;
  reason: CalculationIncompleteReason;
}

function referencedFields(calculation: Calculation): string[] {
  if (calculation.kind === "simple") return calculation.fields;
  const names: string[] = [];
  const visit = (node: ArithNode): void => {
    if (node.kind === "field") {
      names.push(node.name);
      return;
    }
    if (node.kind === "binary") {
      visit(node.left);
      visit(node.right);
    }
  };
  visit(calculation.expression);
  return names;
}

/** The calculation dependency graph shared by analysis, recalculation and
 *  impact queries. */
interface CalculationGraph {
  /** Every field with a calculate script, by rank (/CO, then document order). */
  calculated: string[];
  /** The script each field runs; a repeated name keeps its last entry. */
  byName: Map<string, FieldScripts>;
  /** Supported calculations by field name. */
  parsed: Map<string, Calculation>;
  /** Calculated fields whose script this module does not model, by rank. */
  unsupported: string[];
  /** Supported field -> the distinct fields its calculation reads. */
  references: Map<string, string[]>;
  cycleFields: string[];
  cycleAffectedFields: string[];
  /** Supported, cycle-free calculations in dependency order. */
  order: string[];
}

/** Fields on a dependency cycle (a strongly connected component of more than
 *  one field, or a field that reads itself). Iterative, so a long chain cannot
 *  exhaust the stack. */
function cycleMembers(
  names: ReadonlyArray<string>,
  edges: ReadonlyMap<string, ReadonlyArray<string>>,
): Set<string> {
  const index = new Map<string, number>();
  const low = new Map<string, number>();
  const onStack = new Set<string>();
  const stack: string[] = [];
  const members = new Set<string>();
  let counter = 0;
  const open = (name: string) => {
    index.set(name, counter);
    low.set(name, counter);
    counter += 1;
    stack.push(name);
    onStack.add(name);
  };
  for (const root of names) {
    if (index.has(root)) continue;
    open(root);
    const work: Array<{ name: string; next: number }> = [{ name: root, next: 0 }];
    while (work.length > 0) {
      const frame = work[work.length - 1];
      const deps = edges.get(frame.name) ?? [];
      if (frame.next < deps.length) {
        const dep = deps[frame.next];
        frame.next += 1;
        if (!index.has(dep)) {
          open(dep);
          work.push({ name: dep, next: 0 });
        } else if (onStack.has(dep)) {
          low.set(frame.name, Math.min(low.get(frame.name)!, index.get(dep)!));
        }
        continue;
      }
      work.pop();
      if (work.length > 0) {
        const parent = work[work.length - 1].name;
        low.set(parent, Math.min(low.get(parent)!, low.get(frame.name)!));
      }
      if (low.get(frame.name) !== index.get(frame.name)) continue;
      const component: string[] = [];
      let member: string;
      do {
        member = stack.pop()!;
        onStack.delete(member);
        component.push(member);
      } while (member !== frame.name);
      if (component.length > 1 || deps.includes(frame.name)) {
        for (const name of component) members.add(name);
      }
    }
  }
  return members;
}

function heapPush(
  heap: string[],
  name: string,
  rank: ReadonlyMap<string, number>,
): void {
  heap.push(name);
  let at = heap.length - 1;
  while (at > 0) {
    const parent = (at - 1) >> 1;
    if (rank.get(heap[parent])! <= rank.get(heap[at])!) break;
    [heap[parent], heap[at]] = [heap[at], heap[parent]];
    at = parent;
  }
}

function heapPop(heap: string[], rank: ReadonlyMap<string, number>): string {
  const top = heap[0];
  const last = heap.pop()!;
  if (heap.length > 0) {
    heap[0] = last;
    let at = 0;
    for (;;) {
      const left = at * 2 + 1;
      const right = left + 1;
      let smallest = at;
      if (left < heap.length && rank.get(heap[left])! < rank.get(heap[smallest])!) {
        smallest = left;
      }
      if (right < heap.length && rank.get(heap[right])! < rank.get(heap[smallest])!) {
        smallest = right;
      }
      if (smallest === at) break;
      [heap[smallest], heap[at]] = [heap[at], heap[smallest]];
      at = smallest;
    }
  }
  return top;
}

/**
 * Build the dependency graph for the safe calculation subset.
 *
 * The evaluation order is topological: a field runs only after every
 * calculated field it reads. /CO — then document order — decides only between
 * fields that are ready at the same time, so a complete and consistent /CO is
 * reproduced exactly, while a missing, partial or contradictory one still gives
 * the dependency-correct answer in a single pass. Cycle members and everything
 * downstream of them are left out of the order entirely.
 *
 * The server's derived-write pass (form_calculation.py) ports this graph; keep
 * the two in step.
 */
function buildCalculationGraph(
  scripts: ReadonlyArray<FieldScripts>,
  order?: ReadonlyArray<string>,
): CalculationGraph {
  const byName = new Map<string, FieldScripts>();
  const firstSeen: string[] = [];
  for (const script of scripts) {
    if (!byName.has(script.name)) firstSeen.push(script.name);
    byName.set(script.name, script);
  }
  const documentOrder = firstSeen.filter((name) => !!byName.get(name)?.calculate);
  const calculatedSet = new Set(documentOrder);
  const rank = new Map<string, number>();
  for (const name of order ?? []) {
    if (calculatedSet.has(name) && !rank.has(name)) rank.set(name, rank.size);
  }
  for (const name of documentOrder) {
    if (!rank.has(name)) rank.set(name, rank.size);
  }
  const calculated = [...documentOrder].sort(
    (a, b) => rank.get(a)! - rank.get(b)!,
  );

  const parsed = new Map<string, Calculation>();
  const unsupported: string[] = [];
  const references = new Map<string, string[]>();
  for (const name of calculated) {
    const calculation = parseCalculationScript(byName.get(name)!.calculate);
    if (!calculation) {
      unsupported.push(name);
      continue;
    }
    parsed.set(name, calculation);
    // A parent-group reference depends on the calculated fields beneath it
    
    const refs = referencedFields(calculation).flatMap((ref) =>
      calculation.kind === "simple" && !byName.has(ref)
        ? [ref, ...expandCalculationReference(ref, calculated)]
        : [ref],
    );
    references.set(name, [...new Set(refs)]);
  }

  // Only a supported calculation has known inputs; an unsupported one is a
  // source whose value simply cannot be refreshed.
  const edges = new Map<string, string[]>();
  const dependents = new Map<string, string[]>();
  for (const [name, refs] of references) {
    const deps = refs.filter((ref) => parsed.has(ref));
    edges.set(name, deps);
    for (const dep of deps) {
      const list = dependents.get(dep);
      if (list) list.push(name);
      else dependents.set(dep, [name]);
    }
  }

  const members = cycleMembers([...parsed.keys()], edges);
  const affected = new Set(members);
  const frontier = [...members];
  while (frontier.length > 0) {
    for (const dependent of dependents.get(frontier.pop()!) ?? []) {
      if (affected.has(dependent)) continue;
      affected.add(dependent);
      frontier.push(dependent);
    }
  }

  const pending = new Map<string, number>();
  const ready: string[] = [];
  for (const name of parsed.keys()) {
    if (affected.has(name)) continue;
    const count = edges.get(name)!.length;
    pending.set(name, count);
    if (count === 0) heapPush(ready, name, rank);
  }
  const evaluation: string[] = [];
  while (ready.length > 0) {
    const name = heapPop(ready, rank);
    evaluation.push(name);
    for (const dependent of dependents.get(name) ?? []) {
      const left = pending.get(dependent);
      if (left === undefined) continue;
      pending.set(dependent, left - 1);
      if (left === 1) heapPush(ready, dependent, rank);
    }
  }

  return {
    calculated,
    byName,
    parsed,
    unsupported,
    references,
    cycleFields: documentOrder.filter((name) => members.has(name)),
    cycleAffectedFields: documentOrder.filter((name) => affected.has(name)),
    order: evaluation,
  };
}

/**
 * Report the dependency cycles in the safe calculation subset.
 *
 * A cyclic result is not merely delayed: repeated evaluation can manufacture a
 * plausible intermediate number. Cycle members and every calculated field
 * downstream from them therefore remain on their existing source/reader value.
 */
export function analyzeCalculationDependencies(
  scripts: ReadonlyArray<FieldScripts>,
): CalculationDependencyAnalysis {
  const graph = buildCalculationGraph(scripts);
  return {
    cycleFields: graph.cycleFields,
    cycleAffectedFields: graph.cycleAffectedFields,
  };
}

export interface RecalculationResult {
  /** Field name -> the newly computed raw value: a number, or the text a
   *  script's `+` built from a text operand. */
  values: Map<string, CalculationValue>;
  /** Field name -> the display string its format script asks for. For a field
   *  listed in `incompleteFields` this formats its existing, stale value. */
  display: Map<string, string>;
  /** Fields whose script this module does not model, left untouched. */
  unsupported: string[];
  /** Fields participating directly in a calculation dependency cycle. */
  cycleFields: string[];
  /** Cycle fields and downstream calculations left on their existing values. */
  cycleAffectedFields: string[];
  /**
   * Every calculated field that could NOT be brought up to date, and why, by
   * rank (/CO, then document order). Such a field keeps its existing value,
   * which must not be presented as a computed result. Empty means every
   * calculated field in the form was computed from current inputs.
   */
  incompleteFields: IncompleteCalculation[];
  /**
   * Computed fields whose new value differs from the one they hold (a number
   * against the number their current text reads as, after the field's own
   * format normalisation; text against text). This is the
   * rule the server's derived write for read-only calculated fields uses, so a
   * caller can predict what the saved file will hold.
   */
  changedFields: string[];
  /** Supported, cycle-free calculations in the dependency order they ran. */
  evaluationOrder: string[];
}

/**
 * Recalculate every scripted field in one dependency-ordered pass.
 *
 * `order` is the AcroForm's /CO calculation order when the document declares
 * one. It only breaks ties between fields whose inputs are equally ready: a
 * total that sums subtotals always runs after them, whether or not /CO says
 * so, and a complete, consistent /CO is followed exactly. Calculations are pure
 * functions of their inputs, so any dependency-respecting order gives the same
 * numbers.
 *
 * A calculated field is left on its existing value — and listed in
 * `incompleteFields` — when it is part of or downstream of a cycle, when its
 * own script is unsupported, when it produces no value, or when a calculated
 * field it reads is in one of those states.
 *
 * `knownFields` names every field of the form, including the ones `values`
 * holds nothing for; a name in neither (nor among the scripted fields) is not
 * a field.
 */
export function recalculateFields(
  scripts: ReadonlyArray<FieldScripts>,
  values: ReadonlyMap<string, unknown>,
  order?: ReadonlyArray<string>,
  locale = "en-US",
  knownFields?: Iterable<string>,
): RecalculationResult {
  const current = new Map<string, unknown>(values);
  const computed = new Map<string, CalculationValue>();
  const display = new Map<string, string>();
  const graph = buildCalculationGraph(scripts, order);
  const formatByName = new Map<string, FieldFormatSpec>();
  const textual = new Set<string>();
  for (const [name, script] of graph.byName) {
    const spec = parseFormatScript(script.format);
    if (!spec) continue;
    formatByName.set(name, spec);
    if (spec.kind === "date") textual.add(name);
    else if (current.has(name)) {
      current.set(name, numericValueForSpec(current.get(name), spec));
    }
  }
  const context: CalculationContext = {
    known: new Set([...values.keys(), ...graph.byName.keys(), ...(knownFields ?? [])]),
    textual,
  };

  const stale = new Map<string, CalculationIncompleteReason>();
  for (const name of graph.unsupported) stale.set(name, "unsupported-script");
  for (const name of graph.cycleAffectedFields) stale.set(name, "cycle");
  const changedFields: string[] = [];
  for (const name of graph.order) {
    if (graph.references.get(name)!.some((ref) => stale.has(ref))) {
      stale.set(name, "unsupported-input");
      continue;
    }
    const next = evaluateCalculation(graph.parsed.get(name)!, current, context);
    if (next === null) {
      stale.set(name, "no-value");
      continue;
    }
    if (!holdsCalculatedValue(current.get(name), next)) changedFields.push(name);
    computed.set(name, next);
    current.set(name, next);
  }

  for (const [name, spec] of formatByName) {
    if (current.get(name) === undefined && !computed.has(name)) continue;
    const raw = computed.has(name) ? computed.get(name)! : current.get(name);
    display.set(name, formatWithSpec(raw, spec, locale));
  }

  return {
    values: computed,
    display,
    unsupported: graph.unsupported,
    cycleFields: graph.cycleFields,
    cycleAffectedFields: graph.cycleAffectedFields,
    incompleteFields: graph.calculated
      .filter((name) => stale.has(name))
      .map((name) => ({ name, reason: stale.get(name)! })),
    changedFields,
    evaluationOrder: graph.order,
  };
}

export interface CalculationImpact {
  /** Calculated fields downstream of the change that were brought up to
   *  date, in evaluation order. */
  recomputed: string[];
  /** Calculated fields downstream of the change left on a stale value, and
   *  why. A field with an unsupported script is always listed: nothing can
   *  tell whether a script this module cannot read depends on the change. */
  stale: IncompleteCalculation[];
}

/**
 * Which calculated fields a change to `changedNames` reaches, split into the
 * ones `result` brought up to date and the ones it left stale.
 *
 * `result` must be the recalculation run over the same `scripts` after the
 * change.
 */
export function calculationImpact(
  scripts: ReadonlyArray<FieldScripts>,
  changedNames: Iterable<string>,
  result: RecalculationResult,
): CalculationImpact {
  const graph = buildCalculationGraph(scripts);
  const readers = new Map<string, string[]>();
  for (const [name, refs] of graph.references) {
    for (const ref of refs) {
      const list = readers.get(ref);
      if (list) list.push(name);
      else readers.set(ref, [name]);
    }
  }
  const reached = new Set<string>(graph.unsupported);
  const frontier = [...changedNames, ...graph.unsupported];
  while (frontier.length > 0) {
    for (const reader of readers.get(frontier.pop()!) ?? []) {
      if (reached.has(reader)) continue;
      reached.add(reader);
      frontier.push(reader);
    }
  }
  return {
    recomputed: result.evaluationOrder.filter(
      (name) => reached.has(name) && result.values.has(name),
    ),
    stale: result.incompleteFields.filter((field) => reached.has(field.name)),
  };
}

export interface ReadOnlyCalculationOutcome {
  /** Field name -> the exact /V string the server's derived write stores:
   *  `String(number)`, only for fields whose number changed. */
  values: Map<string, string>;
  /** Read-only calculated fields whose shown value is stale, and why. */
  stale: IncompleteCalculation[];
}

/**
 * What a save will do to read-only calculated fields.
 *
 * A read-only field refuses direct writes, so a host must not send these
 * values as edits. The server recomputes the same fields from the saved inputs
 * with the same rules and writes them itself; a host can show `values` in the
 * meantime and mark `stale` fields as not updated.
 */
export function readOnlyCalculationOutcome(
  result: RecalculationResult,
  isReadOnly: (name: string) => boolean,
): ReadOnlyCalculationOutcome {
  const values = new Map<string, string>();
  for (const name of result.changedFields) {
    if (isReadOnly(name)) values.set(name, String(result.values.get(name)!));
  }
  return {
    values,
    stale: result.incompleteFields.filter((field) => isReadOnly(field.name)),
  };
}
