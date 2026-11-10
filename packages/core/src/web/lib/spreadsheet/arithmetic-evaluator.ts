/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** A cell's value as the grid knows it; `undefined` = not known yet. */
export type ArithmeticValue = number | string | boolean | null | { readonly __formulaError: true; readonly code: string };

const error = (code: string) => ({ __formulaError: true as const, code });
const isError = (value: unknown): value is { readonly __formulaError: true; readonly code: string } =>
  typeof value === "object" && value !== null && (value as { __formulaError?: unknown }).__formulaError === true;

const MAX_LENGTH = 512;
const MAX_TOKENS = 256;

type Token =
  | { kind: "num"; value: number }
  | { kind: "err"; code: string }
  | { kind: "ref"; row: number; col: number }
  | { kind: "op"; value: string };

// The error literals the engine's tokenizer reads (formula-engine
// ERROR_LITERALS), spelled exactly; one running on into an address or a name
// (`#REF!A1`, a reference to a deleted sheet) is the server's.
const TOKEN = /\s+|\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})(?![\w.(!])|((?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)|(<>|<=|>=|[-+*/^%&=<>()])|(#DIV\/0!|#VALUE!|#NAME\?|#N\/A|#NUM!|#NULL!|#REF!|#SPILL!|#ERROR!|#CIRCULAR!|#CALC!)(?![\w$.(!:\\'])/y;

function tokenize(body: string): Token[] | null {
  const tokens: Token[] = [];
  let offset = 0;
  while (offset < body.length) {
    TOKEN.lastIndex = offset;
    const match = TOKEN.exec(body);
    if (!match || match.index !== offset || match[0].length === 0) return null;
    offset = TOKEN.lastIndex;
    if (match[1]) {
      let col = 0;
      for (const char of match[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64;
      tokens.push({ kind: "ref", row: Number(match[2]) - 1, col: col - 1 });
    } else if (match[3]) tokens.push({ kind: "num", value: Number(match[3]) });
    else if (match[4]) tokens.push({ kind: "op", value: match[4] });
    else if (match[5]) tokens.push({ kind: "err", code: match[5] });
    if (tokens.length > MAX_TOKENS) return null;
  }
  return tokens;
}

function toNumber(value: ArithmeticValue): number | { __formulaError: true; code: string } {
  if (isError(value)) return value;
  if (value === null || value === "") return 0;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "number") return value;
  const trimmed = value.trim();
  const parsed = trimmed === "" ? NaN : Number(trimmed);
  return Number.isFinite(parsed) ? parsed : error("#VALUE!");
}

function toText(value: ArithmeticValue): string | { __formulaError: true; code: string } {
  if (isError(value)) return value;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return Number.isInteger(value) ? String(value) : String(Number(value.toPrecision(15)));
  return value;
}

const rank = (value: ArithmeticValue) => (typeof value === "number" || value === null ? 0 : typeof value === "string" ? 1 : 2);

function compare(op: string, a: ArithmeticValue, b: ArithmeticValue): ArithmeticValue {
  if (isError(a)) return a;
  if (isError(b)) return b;
  // A blank compares as the other side's empty value.
  const left = a === null ? (typeof b === "string" ? "" : typeof b === "boolean" ? false : 0) : a;
  const right = b === null ? (typeof a === "string" ? "" : typeof a === "boolean" ? false : 0) : b;
  let order: number;
  if (rank(left) !== rank(right)) order = rank(left) - rank(right);
  else if (typeof left === "string") order = left.toLowerCase().localeCompare((right as string).toLowerCase());
  else order = Number(left) - Number(right);
  switch (op) {
    case "=": return order === 0;
    case "<>": return order !== 0;
    case "<": return order < 0;
    case ">": return order > 0;
    case "<=": return order <= 0;
    default: return order >= 0;
  }
}

function arithmetic(op: string, a: ArithmeticValue, b: ArithmeticValue): ArithmeticValue {
  if (op === "&") {
    const left = toText(a);
    if (isError(left)) return left;
    const right = toText(b);
    return isError(right) ? right : left + right;
  }
  const left = toNumber(a);
  if (isError(left)) return left;
  const right = toNumber(b);
  if (isError(right)) return right;
  let out: number;
  switch (op) {
    case "+": out = left + right; break;
    case "-": out = left - right; break;
    case "*": out = left * right; break;
    case "/": if (right === 0) return error("#DIV/0!"); out = left / right; break;
    default:
      if (left === 0 && right === 0) return error("#NUM!");
      if (left === 0 && right < 0) return error("#DIV/0!");
      out = left ** right;
      // A negative base to the reciprocal of an odd integer is its real root,
      // as the engine's powerOf answers: (-8)^(1/3) is -2 in Excel.
      if (left < 0 && !Number.isInteger(right) && Number.isFinite(right)) {
        const n = Math.round(1 / right);
        if (n % 2 !== 0 && 1 / n === right) out = -((-left) ** right);
      }
  }
  return Number.isFinite(out) ? out : error("#NUM!");
}

/**
 * The value of `formula` (with its `=`) over `read` (a same-sheet cell's
 * value, `undefined` while not known), or `undefined` when the formula is
 * outside the arithmetic subset or reads a value not known yet.
 */
export function evaluateArithmeticFormula(
  formula: string,
  read: (row: number, col: number) => ArithmeticValue | undefined,
): ArithmeticValue | undefined {
  const text = formula.trim();
  if (!text.startsWith("=") || text.length > MAX_LENGTH) return undefined;
  const parsed = tokenize(text.slice(1));
  if (!parsed || parsed.length === 0) return undefined;
  const tokens: Token[] = parsed;
  let at = 0;
  let unknown = false;
  const peek = () => tokens[at];
  const isOp = (value: string) => peek()?.kind === "op" && (peek() as { value: string }).value === value;

  // comparison < concatenation < additive < multiplicative < power < unary < percent < primary
  function primary(): ArithmeticValue {
    const token = tokens[at++];
    if (!token) throw new SyntaxError();
    if (token.kind === "num") return token.value;
    if (token.kind === "err") return error(token.code);
    if (token.kind === "ref") {
      const value = read(token.row, token.col);
      if (value === undefined) { unknown = true; return null; }
      return value;
    }
    if (token.value === "(") {
      const value = comparison();
      if (!isOp(")")) throw new SyntaxError();
      at += 1;
      return value;
    }
    throw new SyntaxError();
  }
  function percent(): ArithmeticValue {
    let value = primary();
    while (isOp("%")) { at += 1; value = arithmetic("/", value, 100); }
    return value;
  }
  function unary(): ArithmeticValue {
    if (isOp("-")) { at += 1; return arithmetic("*", unary(), -1); }
    if (isOp("+")) { at += 1; return unary(); }
    return percent();
  }
  function power(): ArithmeticValue {
    let value = unary();
    while (isOp("^")) { at += 1; value = arithmetic("^", value, unary()); }
    return value;
  }
  function multiplicative(): ArithmeticValue {
    let value = power();
    while (isOp("*") || isOp("/")) { const op = (tokens[at++] as { value: string }).value; value = arithmetic(op, value, power()); }
    return value;
  }
  function additive(): ArithmeticValue {
    let value = multiplicative();
    while (isOp("+") || isOp("-")) { const op = (tokens[at++] as { value: string }).value; value = arithmetic(op, value, multiplicative()); }
    return value;
  }
  function concatenation(): ArithmeticValue {
    let value = additive();
    while (isOp("&")) { at += 1; value = arithmetic("&", value, additive()); }
    return value;
  }
  function comparison(): ArithmeticValue {
    let value = concatenation();
    while (["=", "<>", "<", ">", "<=", ">="].some(isOp)) {
      const op = (tokens[at++] as { value: string }).value;
      value = compare(op, value, concatenation());
    }
    return value;
  }
  try {
    const value = comparison();
    if (at !== tokens.length) return undefined;
    if (unknown) return undefined;
    // A formula whose result is a blank cell is 0, as the engine answers
    // (evaluateFormulaRaw): =B1 over an empty B1 shows 0 in Excel (X-351).
    return value === null ? 0 : value;
  } catch {
    return undefined;
  }
}
