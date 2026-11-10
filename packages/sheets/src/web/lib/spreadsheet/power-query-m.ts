/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------

/** A date: whole days since 1899-12-30 (Excel's serial). */
export class MDate {
  constructor(readonly serial: number) {}
}
/** A datetime: fractional days since 1899-12-30. */
export class MDateTime {
  constructor(readonly serial: number) {}
}
/** A time of day: fraction of a day in [0, 1). */
export class MTime {
  constructor(readonly fraction: number) {}
}
/** A duration in days (may be fractional or negative). */
export class MDuration {
  constructor(readonly days: number) {}
}
export class MRecord {
  constructor(
    readonly names: string[],
    private readonly thunks: Array<() => MValue>,
    private cache: Array<MValue | typeof UNSET> = thunks.map(() => UNSET),
  ) {}
  static of(names: string[], values: MValue[]): MRecord {
    return new MRecord(names, values.map((value) => () => value), values.slice());
  }
  has(name: string): boolean {
    return this.names.includes(name);
  }
  get(name: string): MValue {
    const index = this.names.indexOf(name);
    if (index < 0) throw new PowerQueryError("Expression.Error", `The field '${name}' of the record wasn't found.`);
    return this.at(index);
  }
  at(index: number): MValue {
    const cached = this.cache[index];
    if (cached !== UNSET) return cached as MValue;
    const value = this.thunks[index]();
    this.cache[index] = value;
    return value;
  }
  values(): MValue[] {
    return this.names.map((_, index) => this.at(index));
  }
}
export class MTable {
  constructor(
    readonly columns: string[],
    readonly rows: MValue[][],
  ) {}
}
export class MFunction {
  constructor(
    readonly name: string,
    readonly call: (args: MValue[]) => MValue,
    readonly minArgs = 0,
    readonly maxArgs = Infinity,
    /** Comparer.OrdinalIgnoreCase and friends, recognised by name. */
    readonly ignoreCase = false,
  ) {}
}
export class MType {
  constructor(
    readonly kind: string,
    readonly nullable = false,
    /** Int64.Type, Currency.Type, Percentage.Type: a number with a facet. */
    readonly facet?: string,
    readonly fields?: Array<[string, MType]>,
  ) {}
}
export type MValue =
  | null
  | boolean
  | number
  | string
  | MDate
  | MDateTime
  | MTime
  | MDuration
  | MValue[]
  | MRecord
  | MTable
  | MFunction
  | MType;

const UNSET = Symbol("unset");

/** An error the M program raises (Error.Record: reason, message). */
export class PowerQueryError extends Error {
  constructor(
    readonly reason: string,
    message: string,
  ) {
    super(message);
    this.name = "PowerQueryError";
  }
}
/** Something this interpreter does not run; `feature` names it. */
export class PowerQueryUnsupportedError extends Error {
  constructor(readonly feature: string) {
    super(`unsupported: ${feature}`);
    this.name = "PowerQueryUnsupportedError";
  }
}

const fail = (message: string, reason = "Expression.Error"): never => {
  throw new PowerQueryError(reason, message);
};
const unsupported = (feature: string): never => {
  throw new PowerQueryUnsupportedError(feature);
};

// ---------------------------------------------------------------------------
// Lexer
// ---------------------------------------------------------------------------

interface Token {
  kind: "ident" | "qident" | "text" | "number" | "op" | "keyword" | "eof";
  text: string;
  start: number;
  end: number;
}

const KEYWORDS = new Set([
  "and", "as", "each", "else", "error", "false", "if", "in", "is", "let", "meta",
  "not", "null", "or", "otherwise", "section", "shared", "then", "true", "try", "type",
]);
const IDENT_START = /[\p{L}_]/u;
const IDENT_PART = /[\p{L}\p{N}_\p{Mn}\p{Mc}\p{Pc}]/u;

function readTextLiteral(source: string, index: number): [string, number] {
  let out = "";
  index += 1;
  while (index < source.length) {
    const char = source[index];
    if (char === '"') {
      if (source[index + 1] === '"') {
        out += '"';
        index += 2;
        continue;
      }
      return [out, index + 1];
    }
    if (char === "#" && source[index + 1] === "(") {
      const end = source.indexOf(")", index + 2);
      if (end < 0) fail("Invalid text escape.");
      for (const raw of source.slice(index + 2, end).split(",")) {
        const code = raw.trim();
        if (code === "cr") out += "\r";
        else if (code === "lf") out += "\n";
        else if (code === "tab") out += "\t";
        else if (code === "#") out += "#";
        else if (/^[0-9A-Fa-f]{4}$|^[0-9A-Fa-f]{8}$/.test(code)) out += String.fromCodePoint(parseInt(code, 16));
        else fail("Invalid text escape.");
      }
      index = end + 1;
      continue;
    }
    out += char;
    index += 1;
  }
  return fail("Unterminated text literal.");
}

export function tokenizeM(source: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;
  while (index < source.length) {
    const char = source[index];
    if (/\s/.test(char) || char === "﻿") {
      index += 1;
      continue;
    }
    if (source.startsWith("//", index)) {
      const end = source.indexOf("\n", index);
      index = end < 0 ? source.length : end + 1;
      continue;
    }
    if (source.startsWith("/*", index)) {
      const end = source.indexOf("*/", index + 2);
      if (end < 0) fail("Unterminated comment.");
      index = end + 2;
      continue;
    }
    const start = index;
    if (char === '"') {
      const [value, end] = readTextLiteral(source, index);
      tokens.push({ kind: "text", text: value, start, end });
      index = end;
      continue;
    }
    if (char === "#" && source[index + 1] === '"') {
      const [value, end] = readTextLiteral(source, index + 1);
      tokens.push({ kind: "qident", text: value, start, end });
      index = end;
      continue;
    }
    if (char === "#" && /[A-Za-z]/.test(source[index + 1] ?? "")) {
      let end = index + 1;
      while (end < source.length && /[A-Za-z]/.test(source[end])) end += 1;
      tokens.push({ kind: "ident", text: source.slice(index, end), start, end });
      index = end;
      continue;
    }
    if (/[0-9]/.test(char) || (char === "." && /[0-9]/.test(source[index + 1] ?? ""))) {
      const match = /^(?:0[xX][0-9A-Fa-f]+|(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)/.exec(source.slice(index));
      const text = match![0];
      tokens.push({ kind: "number", text, start, end: index + text.length });
      index += text.length;
      continue;
    }
    if (IDENT_START.test(char)) {
      let end = index + 1;
      for (;;) {
        while (end < source.length && IDENT_PART.test(source[end])) end += 1;
        // A dot continues the identifier when a letter follows (Table.Group).
        if (source[end] === "." && IDENT_START.test(source[end + 1] ?? "")) {
          end += 2;
          continue;
        }
        break;
      }
      const text = source.slice(index, end);
      tokens.push({ kind: KEYWORDS.has(text) ? "keyword" : "ident", text, start, end });
      index = end;
      continue;
    }
    const operator = ["...", "=>", "<=", ">=", "<>", "..", "??"].find((op) => source.startsWith(op, index));
    const text = operator ?? char;
    if (!operator && !"=<>+-*/&(){}[],;?@!".includes(char)) fail(`Unexpected character '${char}'.`);
    tokens.push({ kind: "op", text, start, end: index + text.length });
    index += text.length;
  }
  tokens.push({ kind: "eof", text: "", start: source.length, end: source.length });
  return tokens;
}

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

type Node =
  | { t: "lit"; value: MValue }
  | { t: "id"; name: string }
  | { t: "let"; bindings: Array<[string, Node]>; body: Node }
  | { t: "if"; cond: Node; then: Node; else: Node }
  | { t: "fn"; params: Array<{ name: string; optional: boolean }>; body: Node }
  | { t: "bin"; op: string; left: Node; right: Node }
  | { t: "un"; op: string; operand: Node }
  | { t: "list"; items: Array<Node | { range: [Node, Node] }> }
  | { t: "record"; fields: Array<[string, Node]> }
  | { t: "field"; target: Node; name: string; optional: boolean }
  | { t: "project"; target: Node; names: string[]; optional: boolean }
  | { t: "item"; target: Node; key: Node; optional: boolean }
  | { t: "call"; fn: Node; args: Node[] }
  | { t: "try"; body: Node; otherwise: Node | null }
  | { t: "error"; value: Node }
  | { t: "type"; type: MType }
  | { t: "is"; value: Node; type: MType }
  | { t: "as"; value: Node; type: MType };

class Parser {
  private index = 0;
  constructor(
    private readonly source: string,
    private readonly tokens: Token[],
  ) {}

  private peek(offset = 0): Token {
    return this.tokens[Math.min(this.index + offset, this.tokens.length - 1)];
  }
  private next(): Token {
    const token = this.peek();
    if (this.index < this.tokens.length - 1) this.index += 1;
    return token;
  }
  private is(text: string, offset = 0): boolean {
    const token = this.peek(offset);
    return (token.kind === "op" || token.kind === "keyword") && token.text === text;
  }
  private expect(text: string): Token {
    if (!this.is(text)) fail(`Token '${text}' expected near '${this.source.slice(this.peek().start, this.peek().start + 30)}'.`);
    return this.next();
  }

  parseDocument(): Node {
    const node = this.expression();
    if (this.peek().kind !== "eof") fail(`Unexpected '${this.peek().text}' after the expression.`);
    return node;
  }

  private name(): string {
    const token = this.next();
    if (token.kind === "ident" || token.kind === "qident") return token.text;
    return fail(`A name was expected near '${token.text}'.`);
  }

  expression(): Node {
    if (this.is("let")) {
      this.next();
      const bindings: Array<[string, Node]> = [];
      do {
        const name = this.name();
        this.expect("=");
        bindings.push([name, this.expression()]);
      } while (this.is(",") && this.next());
      this.expect("in");
      return { t: "let", bindings, body: this.expression() };
    }
    if (this.is("if")) {
      this.next();
      const cond = this.expression();
      this.expect("then");
      const then = this.expression();
      this.expect("else");
      return { t: "if", cond, then, else: this.expression() };
    }
    if (this.is("each")) {
      this.next();
      return { t: "fn", params: [{ name: "_", optional: false }], body: this.expression() };
    }
    if (this.is("try")) {
      this.next();
      const body = this.expression();
      if (this.is("otherwise")) {
        this.next();
        return { t: "try", body, otherwise: this.expression() };
      }
      return { t: "try", body, otherwise: null };
    }
    if (this.is("error")) {
      this.next();
      return { t: "error", value: this.expression() };
    }
    if (this.is("(") && this.looksLikeFunction()) return this.functionLiteral();
    return this.coalesce();
  }

  private looksLikeFunction(): boolean {
    let depth = 0;
    for (let offset = 0; ; offset += 1) {
      const token = this.peek(offset);
      if (token.kind === "eof") return false;
      if (token.kind === "op" && (token.text === "(" || token.text === "[" || token.text === "{")) depth += 1;
      if (token.kind === "op" && (token.text === ")" || token.text === "]" || token.text === "}")) {
        depth -= 1;
        if (depth === 0) {
          // `(params) =>` or `(params) as type =>`
          if (this.is("=>", offset + 1)) return true;
          if (this.is("as", offset + 1)) {
            let after = offset + 2;
            if (this.peek(after).text === "nullable") after += 1;
            return this.is("=>", after + 1);
          }
          return false;
        }
      }
    }
  }

  private functionLiteral(): Node {
    this.expect("(");
    const params: Array<{ name: string; optional: boolean }> = [];
    while (!this.is(")")) {
      let optional = false;
      if (this.peek().kind === "ident" && this.peek().text === "optional" && this.peek(1).kind !== "op") {
        this.next();
        optional = true;
      }
      params.push({ name: this.name(), optional });
      if (this.is("as")) {
        this.next();
        this.primitiveType();
      }
      if (!this.is(")")) this.expect(",");
    }
    this.expect(")");
    if (this.is("as")) {
      this.next();
      this.primitiveType();
    }
    this.expect("=>");
    return { t: "fn", params, body: this.expression() };
  }

  private coalesce(): Node {
    let left = this.logicalOr();
    while (this.is("??")) {
      this.next();
      left = { t: "bin", op: "??", left, right: this.logicalOr() };
    }
    return left;
  }
  private logicalOr(): Node {
    let left = this.logicalAnd();
    while (this.is("or")) {
      this.next();
      left = { t: "bin", op: "or", left, right: this.logicalAnd() };
    }
    return left;
  }
  private logicalAnd(): Node {
    let left = this.isAs();
    while (this.is("and")) {
      this.next();
      left = { t: "bin", op: "and", left, right: this.isAs() };
    }
    return left;
  }
  private isAs(): Node {
    let left = this.equality();
    while (this.is("is") || this.is("as")) {
      const op = this.next().text;
      const type = this.primitiveType();
      left = op === "is" ? { t: "is", value: left, type } : { t: "as", value: left, type };
    }
    return left;
  }
  private equality(): Node {
    let left = this.relational();
    while (this.is("=") || this.is("<>")) {
      const op = this.next().text;
      left = { t: "bin", op, left, right: this.relational() };
    }
    return left;
  }
  private relational(): Node {
    let left = this.additive();
    while (this.is("<") || this.is(">") || this.is("<=") || this.is(">=")) {
      const op = this.next().text;
      left = { t: "bin", op, left, right: this.additive() };
    }
    return left;
  }
  private additive(): Node {
    let left = this.multiplicative();
    while (this.is("+") || this.is("-") || this.is("&")) {
      const op = this.next().text;
      left = { t: "bin", op, left, right: this.multiplicative() };
    }
    return left;
  }
  private multiplicative(): Node {
    let left = this.metadata();
    while (this.is("*") || this.is("/")) {
      const op = this.next().text;
      left = { t: "bin", op, left, right: this.metadata() };
    }
    return left;
  }
  private metadata(): Node {
    const value = this.unary();
    if (this.is("meta")) {
      this.next();
      this.unary();
    }
    return value;
  }
  private unary(): Node {
    if (this.is("+") || this.is("-") || this.is("not")) {
      const op = this.next().text;
      return { t: "un", op, operand: this.unary() };
    }
    if (this.is("type")) {
      this.next();
      return { t: "type", type: this.typeExpression() };
    }
    return this.postfix(this.primary());
  }

  /** `nullable text`, `number`, `table`… after `as` / `is`. */
  private primitiveType(): MType {
    if (this.peek().text === "nullable") {
      this.next();
      const inner = this.primitiveType();
      return new MType(inner.kind, true, inner.facet, inner.fields);
    }
    const token = this.next();
    if (token.kind === "keyword" && token.text === "null") return new MType("null", true);
    if (token.kind === "keyword" && token.text === "type") return new MType("type");
    if (token.kind !== "ident") return fail("A type was expected.");
    return primitiveTypeNamed(token.text);
  }

  /** After the `type` keyword. */
  private typeExpression(): MType {
    if (this.peek().text === "nullable") {
      this.next();
      const inner = this.typeExpression();
      return new MType(inner.kind, true, inner.facet, inner.fields);
    }
    if (this.is("[")) return this.recordType();
    if (this.is("{")) {
      this.next();
      this.typeExpression();
      this.expect("}");
      return new MType("list");
    }
    if (this.peek().kind === "ident" && this.peek().text === "table") {
      this.next();
      if (this.is("[")) return new MType("table", false, undefined, this.recordType().fields);
      return new MType("table");
    }
    if (this.peek().kind === "ident" && this.peek().text === "function") return unsupported("function types");
    if (this.is("(")) {
      this.next();
      const inner = this.typeExpression();
      this.expect(")");
      return inner;
    }
    return this.primitiveType();
  }

  private recordType(): MType {
    this.expect("[");
    const fields: Array<[string, MType]> = [];
    while (!this.is("]")) {
      if (this.is("...")) {
        this.next();
        continue;
      }
      const name = this.fieldName();
      let type = new MType("any");
      if (this.is("=")) {
        this.next();
        type = this.typeExpression();
      }
      fields.push([name, type]);
      if (!this.is("]")) this.expect(",");
    }
    this.expect("]");
    return new MType("record", false, undefined, fields);
  }

  /** A generalized identifier inside [ ]: `Avg Price`, `#"x"`, `Targets.Target`. */
  private fieldName(): string {
    const first = this.peek();
    if (first.kind === "qident") {
      this.next();
      return first.text;
    }
    if (first.kind !== "ident" && first.kind !== "keyword" && first.kind !== "number") return fail("A field name was expected.");
    let last = first;
    this.next();
    while (
      (this.peek().kind === "ident" || this.peek().kind === "keyword" || this.peek().kind === "number") &&
      !this.is("]") && !this.is("=")
    ) {
      last = this.next();
    }
    return this.source.slice(first.start, last.end).replace(/\s+/g, " ");
  }

  private primary(): Node {
    const token = this.peek();
    if (token.kind === "number") {
      this.next();
      return { t: "lit", value: token.text.startsWith("0x") || token.text.startsWith("0X") ? parseInt(token.text, 16) : Number(token.text) };
    }
    if (token.kind === "text") {
      this.next();
      return { t: "lit", value: token.text };
    }
    if (token.kind === "keyword") {
      if (token.text === "null") return this.next(), { t: "lit", value: null };
      if (token.text === "true") return this.next(), { t: "lit", value: true };
      if (token.kind === "keyword" && token.text === "false") return this.next(), { t: "lit", value: false };
    }
    if (token.kind === "ident" || token.kind === "qident") {
      this.next();
      return { t: "id", name: token.text };
    }
    if (this.is("@")) {
      this.next();
      return { t: "id", name: this.name() };
    }
    if (this.is("(")) {
      this.next();
      const inner = this.expression();
      this.expect(")");
      return inner;
    }
    if (this.is("{")) {
      this.next();
      const items: Array<Node | { range: [Node, Node] }> = [];
      while (!this.is("}")) {
        const item = this.expression();
        if (this.is("..")) {
          this.next();
          items.push({ range: [item, this.expression()] });
        } else items.push(item);
        if (!this.is("}")) this.expect(",");
      }
      this.expect("}");
      return { t: "list", items };
    }
    if (this.is("[")) {
      // Implicit field access `[Name]` / projection `[[A],[B]]`, or a record.
      if (this.is("[", 1)) return { t: "project", target: { t: "id", name: "_" }, names: this.projection(), optional: this.optionalMark() };
      this.next();
      if (this.is("]")) {
        this.next();
        return { t: "record", fields: [] };
      }
      const name = this.fieldName();
      if (this.is("]")) {
        this.next();
        return { t: "field", target: { t: "id", name: "_" }, name, optional: this.optionalMark() };
      }
      const fields: Array<[string, Node]> = [];
      this.expect("=");
      fields.push([name, this.expression()]);
      while (this.is(",")) {
        this.next();
        const next = this.fieldName();
        this.expect("=");
        fields.push([next, this.expression()]);
      }
      this.expect("]");
      return { t: "record", fields };
    }
    return fail(`Unexpected '${token.text || "end of formula"}'.`);
  }

  private optionalMark(): boolean {
    if (this.is("?")) {
      this.next();
      return true;
    }
    return false;
  }

  private projection(): string[] {
    this.expect("[");
    const names: string[] = [];
    while (!this.is("]")) {
      this.expect("[");
      names.push(this.fieldName());
      this.expect("]");
      if (!this.is("]")) this.expect(",");
    }
    this.expect("]");
    return names;
  }

  private postfix(node: Node): Node {
    for (;;) {
      if (this.is("(")) {
        this.next();
        const args: Node[] = [];
        while (!this.is(")")) {
          args.push(this.expression());
          if (!this.is(")")) this.expect(",");
        }
        this.expect(")");
        node = { t: "call", fn: node, args };
      } else if (this.is("{")) {
        this.next();
        const key = this.expression();
        this.expect("}");
        node = { t: "item", target: node, key, optional: this.optionalMark() };
      } else if (this.is("[")) {
        if (this.is("[", 1)) {
          node = { t: "project", target: node, names: this.projection(), optional: false };
          node.optional = this.optionalMark();
          continue;
        }
        this.next();
        const name = this.fieldName();
        this.expect("]");
        node = { t: "field", target: node, name, optional: this.optionalMark() };
      } else return node;
    }
  }
}

function primitiveTypeNamed(name: string): MType {
  const known = [
    "any", "anynonnull", "binary", "date", "datetime", "datetimezone", "duration", "function",
    "list", "logical", "none", "number", "record", "table", "text", "time", "type",
  ];
  if (!known.includes(name)) return fail(`The type '${name}' is not known.`);
  return new MType(name);
}

const parseCache = new Map<string, Node>();
export function parseM(formula: string): Node {
  const cached = parseCache.get(formula);
  if (cached) return cached;
  const node = new Parser(formula, tokenizeM(formula)).parseDocument();
  if (parseCache.size > 256) parseCache.clear();
  parseCache.set(formula, node);
  return node;
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

class Env {
  private readonly values = new Map<string, () => MValue>();
  constructor(readonly parent: Env | null) {}
  define(name: string, thunk: () => MValue): void {
    this.values.set(name, thunk);
  }
  private find(name: string): (() => MValue) | undefined {
    return this.values.get(name) ?? this.parent?.find(name);
  }
  lookup(name: string): MValue {
    const thunk = this.find(name);
    if (thunk) return thunk();
    if (/^(Web|Sql|File|Folder|OData|SharePoint|Odbc|OleDb|Oracle|MySQL|PostgreSQL|AnalysisServices|AzureStorage|Salesforce|Exchange|ActiveDirectory|Access|Pdf|Json|Xml|Csv|Excel|Binary|Lines|Html)\./.test(name))
      return unsupported(name);
    if (/^[A-Z#][A-Za-z]*\.[A-Za-z.]+$|^#[a-z]+$/.test(name)) return unsupported(name);
    return fail(`The name '${name}' wasn't recognized. Make sure it's spelled correctly.`);
  }
}

function lazy(compute: () => MValue): () => MValue {
  let state: 0 | 1 | 2 = 0;
  let value: MValue = null;
  return () => {
    if (state === 2) return value;
    if (state === 1) fail("A cyclic reference was encountered during evaluation.");
    state = 1;
    try {
      value = compute();
    } catch (error) {
      state = 0;
      throw error;
    }
    state = 2;
    return value;
  };
}

function evaluate(node: Node, env: Env): MValue {
  switch (node.t) {
    case "lit":
      return node.value;
    case "id":
      return env.lookup(node.name);
    case "let": {
      const scope = new Env(env);
      for (const [name, value] of node.bindings) scope.define(name, lazy(() => evaluate(value, scope)));
      return evaluate(node.body, scope);
    }
    case "if": {
      const cond = evaluate(node.cond, env);
      if (typeof cond !== "boolean") fail(`We cannot convert the value ${describe(cond)} to type Logical.`);
      return evaluate(cond ? node.then : node.else, env);
    }
    case "fn": {
      const { params, body } = node;
      const required = params.filter((param) => !param.optional).length;
      return new MFunction("function", (args) => {
        const scope = new Env(env);
        params.forEach((param, index) => {
          const value = index < args.length ? args[index] : null;
          scope.define(param.name, () => value);
        });
        return evaluate(body, scope);
      }, required, params.length);
    }
    case "bin":
      return binary(node.op, node.left, node.right, env);
    case "un": {
      const value = evaluate(node.operand, env);
      if (node.op === "not") {
        if (value === null) return null;
        if (typeof value !== "boolean") fail(`We cannot apply operator 'not' to ${describe(value)}.`);
        return !value;
      }
      if (value === null) return null;
      if (node.op === "-") {
        if (typeof value === "number") return -value;
        if (value instanceof MDuration) return new MDuration(-value.days);
      } else if (typeof value === "number" || value instanceof MDuration) return value;
      return fail(`We cannot apply operator ${node.op} to ${describe(value)}.`);
    }
    case "list": {
      const out: MValue[] = [];
      for (const item of node.items) {
        if ("range" in item) {
          const from = toNumber(evaluate(item.range[0], env));
          const to = toNumber(evaluate(item.range[1], env));
          if (to - from > 1_000_000) unsupported("lists over a million items");
          for (let value = from; value <= to; value += 1) out.push(value);
        } else out.push(evaluate(item, env));
      }
      return out;
    }
    case "record": {
      const names = node.fields.map(([name]) => name);
      if (new Set(names).size !== names.length) fail("A record field name is repeated.");
      let record: MRecord | null = null;
      const scope = new Env(env);
      const thunks = node.fields.map(([, value]) => lazy(() => evaluate(value, scope)));
      record = new MRecord(names, thunks);
      names.forEach((name, index) => scope.define(name, () => record!.at(index)));
      return record;
    }
    case "field": {
      const target = evaluate(node.target, env);
      return fieldAccess(target, node.name, node.optional);
    }
    case "project": {
      const target = evaluate(node.target, env);
      if (target instanceof MRecord) {
        return MRecord.of(node.names, node.names.map((name) => (target.has(name) ? target.get(name) : node.optional ? null : fail(`The field '${name}' of the record wasn't found.`))));
      }
      if (target instanceof MTable) return selectColumns(target, node.names, node.optional ? 2 : 0);
      return fail(`We cannot apply field access to ${describe(target)}.`);
    }
    case "item":
      return itemAccess(evaluate(node.target, env), evaluate(node.key, env), node.optional);
    case "call": {
      const fn = evaluate(node.fn, env);
      if (!(fn instanceof MFunction)) return fail(`We cannot convert ${describe(fn)} to Function.`);
      const args = node.args.map((arg) => evaluate(arg, env));
      return invoke(fn, args);
    }
    case "try":
      try {
        const value = evaluate(node.body, env);
        if (node.otherwise) return value;
        return MRecord.of(["HasError", "Value"], [false, value]);
      } catch (error) {
        if (!(error instanceof PowerQueryError)) throw error;
        if (node.otherwise) return evaluate(node.otherwise, env);
        return MRecord.of(["HasError", "Error"], [true, MRecord.of(["Reason", "Message", "Detail"], [error.reason, error.message, null])]);
      }
    case "error": {
      const value = evaluate(node.value, env);
      if (typeof value === "string") return fail(value, "Expression.Error");
      if (value instanceof MRecord) {
        const reason = value.has("Reason") ? value.get("Reason") : "Expression.Error";
        const message = value.has("Message") ? value.get("Message") : "";
        return fail(String(message ?? ""), String(reason ?? "Expression.Error"));
      }
      return fail("An error was raised.");
    }
    case "type":
      return node.type;
    case "is":
      return conformsTo(evaluate(node.value, env), node.type);
    case "as": {
      const value = evaluate(node.value, env);
      if (!conformsTo(value, node.type)) fail(`We cannot convert ${describe(value)} to type ${node.type.kind}.`);
      return value;
    }
  }
}

export function invoke(fn: MFunction, args: MValue[]): MValue {
  if (args.length < fn.minArgs || args.length > fn.maxArgs)
    fail(`${fn.name === "function" ? "The function" : fn.name} expects ${fn.minArgs === fn.maxArgs ? fn.minArgs : `${fn.minArgs} to ${fn.maxArgs}`} arguments but received ${args.length}.`);
  return fn.call(args);
}

function conformsTo(value: MValue, type: MType): boolean {
  if (value === null) return type.nullable || ["any", "null"].includes(type.kind);
  switch (type.kind) {
    case "any":
      return true;
    case "anynonnull":
      return true;
    case "number":
      return typeof value === "number";
    case "text":
      return typeof value === "string";
    case "logical":
      return typeof value === "boolean";
    case "date":
      return value instanceof MDate;
    case "datetime":
      return value instanceof MDateTime;
    case "time":
      return value instanceof MTime;
    case "duration":
      return value instanceof MDuration;
    case "list":
      return Array.isArray(value);
    case "record":
      return value instanceof MRecord;
    case "table":
      return value instanceof MTable;
    case "function":
      return value instanceof MFunction;
    case "type":
      return value instanceof MType;
    default:
      return false;
  }
}

function describe(value: MValue): string {
  if (value === null) return "null";
  if (typeof value === "string") return JSON.stringify(value);
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) return "List";
  if (value instanceof MRecord) return "Record";
  if (value instanceof MTable) return "Table";
  if (value instanceof MFunction) return "Function";
  if (value instanceof MDate) return formatDate(value.serial, "yyyy-MM-dd");
  if (value instanceof MDateTime) return formatDate(value.serial, "yyyy-MM-dd HH:mm:ss");
  if (value instanceof MType) return "Type";
  return "value";
}

function fieldAccess(target: MValue, name: string, optional: boolean): MValue {
  if (target instanceof MRecord) {
    if (!target.has(name)) return optional ? null : fail(`The field '${name}' of the record wasn't found.`);
    return target.get(name);
  }
  if (target instanceof MTable) {
    const index = target.columns.indexOf(name);
    if (index < 0) return optional ? null : fail(`The column '${name}' of the table wasn't found.`, "Expression.Error");
    return target.rows.map((row) => row[index]);
  }
  if (target === null && optional) return null;
  return fail(`We cannot apply field access to the type ${describe(target)}.`);
}

function rowRecord(table: MTable, row: MValue[]): MRecord {
  return MRecord.of(table.columns, row);
}

function itemAccess(target: MValue, key: MValue, optional: boolean): MValue {
  if (Array.isArray(target)) {
    const index = toNumber(key);
    if (index >= 0 && index < target.length && Number.isInteger(index)) return target[index];
    return optional ? null : fail("There weren't enough elements in the enumeration to complete the operation.");
  }
  if (target instanceof MTable) {
    if (typeof key === "number") {
      if (key >= 0 && key < target.rows.length && Number.isInteger(key)) return rowRecord(target, target.rows[key]);
      return optional ? null : fail("There weren't enough elements in the enumeration to complete the operation.");
    }
    if (key instanceof MRecord) {
      const matches = target.rows.filter((row) =>
        key.names.every((name) => {
          const index = target.columns.indexOf(name);
          if (index < 0) fail(`The column '${name}' of the table wasn't found.`);
          return equals(row[index], key.get(name));
        }),
      );
      if (matches.length === 1) return rowRecord(target, matches[0]);
      if (matches.length === 0) return optional ? null : fail("The key didn't match any rows in the table.");
      return fail("The key matched more than one row in the table.");
    }
  }
  return fail(`We cannot apply indexing to ${describe(target)}.`);
}

function binary(op: string, leftNode: Node, rightNode: Node, env: Env): MValue {
  if (op === "and" || op === "or") {
    const left = evaluate(leftNode, env);
    if (left !== null && typeof left !== "boolean") fail(`We cannot apply operator ${op} to ${describe(left)}.`);
    if (op === "and" && left === false) return false;
    if (op === "or" && left === true) return true;
    const right = evaluate(rightNode, env);
    if (right !== null && typeof right !== "boolean") fail(`We cannot apply operator ${op} to ${describe(right)}.`);
    if (op === "and") {
      if (right === false) return false;
      return left === null || right === null ? null : true;
    }
    if (right === true) return true;
    return left === null || right === null ? null : false;
  }
  if (op === "??") {
    const left = evaluate(leftNode, env);
    return left !== null ? left : evaluate(rightNode, env);
  }
  const left = evaluate(leftNode, env);
  const right = evaluate(rightNode, env);
  return applyBinary(op, left, right);
}

function applyBinary(op: string, left: MValue, right: MValue): MValue {
  switch (op) {
    case "=":
      return equals(left, right);
    case "<>":
      return !equals(left, right);
    case "<":
    case ">":
    case "<=":
    case ">=": {
      if (left === null || right === null) return null;
      const order = compareStrict(left, right);
      return op === "<" ? order < 0 : op === ">" ? order > 0 : op === "<=" ? order <= 0 : order >= 0;
    }
    case "&": {
      if (left === null || right === null) return null;
      if (typeof left === "string" && typeof right === "string") return left + right;
      if (Array.isArray(left) && Array.isArray(right)) return [...left, ...right];
      if (left instanceof MRecord && right instanceof MRecord) return combineRecords([left, right]);
      if (left instanceof MTable && right instanceof MTable) return combineTables([left, right]);
      if (left instanceof MDate && right instanceof MTime) return new MDateTime(left.serial + right.fraction);
      return fail(`We cannot apply operator & to types ${describe(left)} and ${describe(right)}.`);
    }
  }
  if (left === null || right === null) return null;
  if (typeof left === "number" && typeof right === "number") {
    if (op === "+") return left + right;
    if (op === "-") return left - right;
    if (op === "*") return left * right;
    if (op === "/") return left / right;
  }
  if (op === "+" || op === "-") {
    const sign = op === "+" ? 1 : -1;
    if (left instanceof MDate && right instanceof MDuration) return new MDate(Math.floor(left.serial + sign * right.days));
    if (left instanceof MDateTime && right instanceof MDuration) return new MDateTime(left.serial + sign * right.days);
    if (left instanceof MDuration && right instanceof MDuration) return new MDuration(left.days + sign * right.days);
    if (op === "+" && left instanceof MDuration && (right instanceof MDate || right instanceof MDateTime)) return applyBinary("+", right, left);
    if (op === "-" && left instanceof MDate && right instanceof MDate) return new MDuration(left.serial - right.serial);
    if (op === "-" && left instanceof MDateTime && right instanceof MDateTime) return new MDuration(left.serial - right.serial);
  }
  if (op === "*" && left instanceof MDuration && typeof right === "number") return new MDuration(left.days * right);
  if (op === "*" && typeof left === "number" && right instanceof MDuration) return new MDuration(left * right.days);
  if (op === "/" && left instanceof MDuration && typeof right === "number") return new MDuration(left.days / right);
  return fail(`We cannot apply operator ${op} to types ${describe(left)} and ${describe(right)}.`);
}

// ---------------------------------------------------------------------------
// Equality / ordering (M's default: ordinal text, nulls first)
// ---------------------------------------------------------------------------

function typeRank(value: MValue): number {
  if (value === null) return 0;
  if (typeof value === "boolean") return 1;
  if (typeof value === "number") return 2;
  if (value instanceof MTime) return 3;
  if (value instanceof MDate) return 4;
  if (value instanceof MDateTime) return 5;
  if (value instanceof MDuration) return 6;
  if (typeof value === "string") return 7;
  return 8;
}

export function equals(left: MValue, right: MValue, ignoreCase = false): boolean {
  if (left === null || right === null) return left === right;
  if (typeof left === "string" && typeof right === "string")
    return ignoreCase ? left.toLowerCase() === right.toLowerCase() : left === right;
  if (typeof left === "number" && typeof right === "number") return left === right;
  if (typeof left === "boolean" && typeof right === "boolean") return left === right;
  if (left instanceof MDate && right instanceof MDate) return left.serial === right.serial;
  if (left instanceof MDateTime && right instanceof MDateTime) return left.serial === right.serial;
  if (left instanceof MTime && right instanceof MTime) return left.fraction === right.fraction;
  if (left instanceof MDuration && right instanceof MDuration) return left.days === right.days;
  if (Array.isArray(left) && Array.isArray(right))
    return left.length === right.length && left.every((item, index) => equals(item, right[index], ignoreCase));
  if (left instanceof MRecord && right instanceof MRecord)
    return left.names.length === right.names.length && left.names.every((name) => right.has(name) && equals(left.get(name), right.get(name), ignoreCase));
  if (left instanceof MTable && right instanceof MTable)
    return equals(left.columns, right.columns) && left.rows.length === right.rows.length && left.rows.every((row, index) => equals(row, right.rows[index], ignoreCase));
  return false;
}

function ordinal(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function compareStrict(left: MValue, right: MValue): number {
  const rank = typeRank(left);
  if (rank !== typeRank(right) || rank === 8)
    return fail(`We cannot apply operator < to types ${describe(left)} and ${describe(right)}.`);
  return compareValues(left, right);
}

/** Ordering for sorting: null first, then by type, values within a type. */
export function compareValues(left: MValue, right: MValue, ignoreCase = false): number {
  const leftRank = typeRank(left);
  const rightRank = typeRank(right);
  if (leftRank !== rightRank) return leftRank - rightRank;
  if (left === null) return 0;
  if (typeof left === "string") {
    const a = ignoreCase ? left.toLowerCase() : left;
    const b = ignoreCase ? (right as string).toLowerCase() : (right as string);
    return ordinal(a, b);
  }
  if (typeof left === "number" || typeof left === "boolean") return Number(left) - Number(right);
  if (left instanceof MDate || left instanceof MDateTime) return left.serial - (right as MDate).serial;
  if (left instanceof MTime) return left.fraction - (right as MTime).fraction;
  if (left instanceof MDuration) return left.days - (right as MDuration).days;
  return 0;
}

/** A hashable identity for grouping / distinct, consistent with `equals`. */
function keyOf(value: MValue, ignoreCase = false): string {
  if (value === null) return "n";
  if (typeof value === "string") return "t" + (ignoreCase ? value.toLowerCase() : value);
  if (typeof value === "number") return "d" + (Object.is(value, -0) ? 0 : value);
  if (typeof value === "boolean") return "b" + value;
  if (value instanceof MDate) return "D" + value.serial;
  if (value instanceof MDateTime) return "T" + value.serial;
  if (value instanceof MTime) return "h" + value.fraction;
  if (value instanceof MDuration) return "u" + value.days;
  if (Array.isArray(value)) return "[" + value.map((item) => keyOf(item, ignoreCase)).join(",") + "]";
  if (value instanceof MRecord) return "{" + value.names.map((name) => name + ":" + keyOf(value.get(name), ignoreCase)).join(",") + "}";
  return unsupported("grouping by tables or functions");
}

// ---------------------------------------------------------------------------
// Conversions
// ---------------------------------------------------------------------------

function toNumber(value: MValue): number {
  if (typeof value !== "number") fail(`We cannot convert the value ${describe(value)} to type Number.`);
  return value as number;
}
function toText(value: MValue): string {
  if (typeof value !== "string") fail(`We cannot convert the value ${describe(value)} to type Text.`);
  return value as string;
}
function toTable(value: MValue): MTable {
  if (!(value instanceof MTable)) fail(`We cannot convert the value ${describe(value)} to type Table.`);
  return value as MTable;
}
function toList(value: MValue): MValue[] {
  if (!Array.isArray(value)) fail(`We cannot convert the value ${describe(value)} to type List.`);
  return value as MValue[];
}
function toFunction(value: MValue): MFunction {
  if (!(value instanceof MFunction)) fail(`We cannot convert the value ${describe(value)} to type Function.`);
  return value as MFunction;
}
function toRecord(value: MValue): MRecord {
  if (!(value instanceof MRecord)) fail(`We cannot convert the value ${describe(value)} to type Record.`);
  return value as MRecord;
}
function names(value: MValue): string[] {
  if (typeof value === "string") return [value];
  return toList(value).map(toText);
}
function optionalNumber(value: MValue | undefined, fallback: number): number {
  return value === undefined || value === null ? fallback : toNumber(value);
}

interface Culture {
  name: string;
  decimal: string;
  group: string;
  shortDate: string;
  longTime: string;
  english: boolean;
}
const CULTURES: Record<string, Culture> = {
  "en-us": { name: "en-US", decimal: ".", group: ",", shortDate: "M/d/yyyy", longTime: "h:mm:ss tt", english: true },
  "en-gb": { name: "en-GB", decimal: ".", group: ",", shortDate: "dd/MM/yyyy", longTime: "HH:mm:ss", english: true },
  "en-au": { name: "en-AU", decimal: ".", group: ",", shortDate: "d/MM/yyyy", longTime: "h:mm:ss tt", english: true },
  "en-ca": { name: "en-CA", decimal: ".", group: ",", shortDate: "yyyy-MM-dd", longTime: "h:mm:ss tt", english: true },
  "de-de": { name: "de-DE", decimal: ",", group: ".", shortDate: "dd.MM.yyyy", longTime: "HH:mm:ss", english: false },
  "fr-fr": { name: "fr-FR", decimal: ",", group: " ", shortDate: "dd/MM/yyyy", longTime: "HH:mm:ss", english: false },
  "zh-cn": { name: "zh-CN", decimal: ".", group: ",", shortDate: "yyyy/M/d", longTime: "H:mm:ss", english: false },
  "ja-jp": { name: "ja-JP", decimal: ".", group: ",", shortDate: "yyyy/MM/dd", longTime: "H:mm:ss", english: false },
};
function cultureOf(name: string | null | undefined, fallback: Culture): Culture {
  if (!name) return fallback;
  return CULTURES[name.toLowerCase()] ?? unsupported(`culture ${name}`);
}

/** .NET's shortest round-trip double text ("R"), as M's Number.ToText writes it. */
export function numberText(value: number, culture: Culture): string {
  if (Number.isNaN(value)) return "NaN";
  if (value === Infinity) return "Infinity";
  if (value === -Infinity) return "-Infinity";
  const magnitude = Math.abs(value);
  let text: string;
  if (magnitude !== 0 && (magnitude >= 1e15 || magnitude < 1e-4)) {
    const [mantissa, exponent] = value.toExponential().split("e");
    const power = Number(exponent);
    text = `${mantissa}E${power < 0 ? "-" : "+"}${String(Math.abs(power)).padStart(2, "0")}`;
  } else text = String(value === 0 ? 0 : value);
  return culture.decimal === "." ? text : text.replace(".", culture.decimal);
}

function parseNumberText(text: string, culture: Culture): number {
  let normalized = text.trim();
  if (!normalized) fail(`We couldn't convert to Number.`, "DataFormat.Error");
  const percent = normalized.endsWith("%");
  if (percent) normalized = normalized.slice(0, -1).trim();
  const group = culture.group === " " ? /[   ]/g : new RegExp(`\\${culture.group}`, "g");
  normalized = normalized.replace(group, "");
  if (culture.decimal !== ".") normalized = normalized.replace(culture.decimal, ".");
  if (/^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$/.test(normalized)) {
    const value = Number(normalized);
    return percent ? value / 100 : value;
  }
  if (/^\(([0-9.]+)\)$/.test(normalized)) return -Number(normalized.slice(1, -1));
  return fail(`We couldn't convert to Number.`, "DataFormat.Error");
}

/** Round half to even at `digits` decimals, as .NET Math.Round. */
function roundToEven(value: number, digits = 0): number {
  const scale = 10 ** digits;
  const scaled = value * scale;
  const floor = Math.floor(scaled);
  const diff = scaled - floor;
  let rounded: number;
  if (Math.abs(diff - 0.5) < 1e-9 * Math.max(1, Math.abs(scaled))) rounded = floor % 2 === 0 ? floor : floor + 1;
  else rounded = Math.round(scaled);
  return rounded / scale;
}
function roundAway(value: number, digits = 0): number {
  const scale = 10 ** digits;
  const scaled = value * scale;
  const floor = Math.floor(scaled);
  const diff = scaled - floor;
  if (Math.abs(diff - 0.5) < 1e-9 * Math.max(1, Math.abs(scaled))) return (scaled < 0 ? floor : floor + 1) / scale;
  return Math.round(scaled) / scale;
}

// Dates (serial days since 1899-12-30, the proleptic Gregorian calendar).
const EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86_400_000;
function serialOf(year: number, month: number, day: number): number {
  return Math.round((Date.UTC(year, month - 1, day) - EPOCH) / DAY_MS);
}
function partsOf(serial: number): { year: number; month: number; day: number; hour: number; minute: number; second: number; ms: number; dow: number } {
  const whole = Math.floor(serial);
  const date = new Date(EPOCH + whole * DAY_MS);
  const fraction = Math.round((serial - whole) * DAY_MS);
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hour: Math.floor(fraction / 3_600_000),
    minute: Math.floor((fraction % 3_600_000) / 60_000),
    second: Math.floor((fraction % 60_000) / 1000),
    ms: fraction % 1000,
    dow: date.getUTCDay(),
  };
}
function addMonths(serial: number, months: number): number {
  const parts = partsOf(serial);
  const total = parts.year * 12 + (parts.month - 1) + months;
  const year = Math.floor(total / 12);
  const month = (total % 12) + 1;
  const days = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return serialOf(year, month, Math.min(parts.day, days)) + (serial - Math.floor(serial));
}

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/** A .NET custom date/time format (yyyy, MM, dd, HH, mm, ss, tt, quoted literals). */
export function formatDate(serial: number, format: string, culture: Culture = CULTURES["en-us"]): string {
  const parts = partsOf(serial);
  let out = "";
  for (let index = 0; index < format.length; ) {
    const char = format[index];
    if (char === "'" || char === '"') {
      const end = format.indexOf(char, index + 1);
      out += format.slice(index + 1, end < 0 ? format.length : end);
      index = end < 0 ? format.length : end + 1;
      continue;
    }
    if (char === "\\") {
      out += format[index + 1] ?? "";
      index += 2;
      continue;
    }
    let run = 1;
    while (format[index + run] === char) run += 1;
    const pad = (value: number, width: number) => String(value).padStart(width, "0");
    switch (char) {
      case "y":
        out += run <= 2 ? pad(parts.year % 100, run) : pad(parts.year, run);
        break;
      case "M":
        if (run >= 3 && !culture.english) unsupported(`month names in ${culture.name}`);
        out += run === 1 ? String(parts.month) : run === 2 ? pad(parts.month, 2) : run === 3 ? MONTHS[parts.month - 1].slice(0, 3) : MONTHS[parts.month - 1];
        break;
      case "d":
        if (run >= 3 && !culture.english) unsupported(`day names in ${culture.name}`);
        out += run === 1 ? String(parts.day) : run === 2 ? pad(parts.day, 2) : run === 3 ? DAYS[parts.dow].slice(0, 3) : DAYS[parts.dow];
        break;
      case "H":
        out += pad(parts.hour, Math.min(run, 2));
        break;
      case "h":
        out += pad(parts.hour % 12 || 12, Math.min(run, 2));
        break;
      case "m":
        out += pad(parts.minute, Math.min(run, 2));
        break;
      case "s":
        out += pad(parts.second, Math.min(run, 2));
        break;
      case "f":
        out += pad(parts.ms, 3).slice(0, run);
        break;
      case "t":
        if (!culture.english) unsupported(`AM/PM designators in ${culture.name}`);
        out += (parts.hour < 12 ? "AM" : "PM").slice(0, run === 1 ? 1 : 2);
        break;
      default:
        out += char.repeat(run);
    }
    index += run;
  }
  return out;
}

function parseDateText(text: string, culture: Culture): number {
  const trimmed = text.trim();
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?$/.exec(trimmed);
  if (match) {
    const [, y, m, d, hh, mm, ss] = match;
    return serialOf(+y, +m, +d) + ((+(hh ?? 0)) * 3600 + (+(mm ?? 0)) * 60 + (+(ss ?? 0))) / 86400;
  }
  match = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?$/i.exec(trimmed);
  if (match) {
    const [, a, b, y, hh, mm, ss, ampm] = match;
    const monthFirst = culture.shortDate.startsWith("M");
    const month = monthFirst ? +a : +b;
    const day = monthFirst ? +b : +a;
    let hour = +(hh ?? 0);
    if (ampm) hour = (hour % 12) + (ampm.toUpperCase() === "PM" ? 12 : 0);
    return serialOf(+y, month, day) + (hour * 3600 + +(mm ?? 0) * 60 + +(ss ?? 0)) / 86400;
  }
  return fail(`We couldn't parse the input provided as a Date value.`, "DataFormat.Error");
}

// ---------------------------------------------------------------------------
// Table helpers
// ---------------------------------------------------------------------------

function columnIndex(table: MTable, name: string): number {
  const index = table.columns.indexOf(name);
  if (index < 0) fail(`The column '${name}' of the table wasn't found.`);
  return index;
}

/** MissingField.Error = 0, Ignore = 1, UseNull = 2. */
function selectColumns(table: MTable, wanted: string[], missing = 0): MTable {
  const kept: Array<number | null> = [];
  const columns: string[] = [];
  for (const name of wanted) {
    const index = table.columns.indexOf(name);
    if (index < 0) {
      if (missing === 1) continue;
      if (missing !== 2) fail(`The column '${name}' of the table wasn't found.`);
    }
    kept.push(index < 0 ? null : index);
    columns.push(name);
  }
  return new MTable(columns, table.rows.map((row) => kept.map((index) => (index === null ? null : row[index]))));
}

function uniqueColumns(columns: string[]): void {
  if (new Set(columns).size !== columns.length) {
    const seen = new Set<string>();
    const duplicate = columns.find((name) => (seen.has(name) ? true : (seen.add(name), false)));
    fail(`The field '${duplicate}' already exists in the record.`);
  }
}

function combineTables(tables: MTable[]): MTable {
  const columns: string[] = [];
  for (const table of tables) for (const name of table.columns) if (!columns.includes(name)) columns.push(name);
  const rows: MValue[][] = [];
  for (const table of tables) {
    const map = columns.map((name) => table.columns.indexOf(name));
    for (const row of table.rows) rows.push(map.map((index) => (index < 0 ? null : row[index])));
  }
  return new MTable(columns, rows);
}
function combineRecords(records: MRecord[]): MRecord {
  const fieldNames: string[] = [];
  const values = new Map<string, MValue>();
  for (const record of records)
    for (const name of record.names) {
      if (!fieldNames.includes(name)) fieldNames.push(name);
      values.set(name, record.get(name));
    }
  return MRecord.of(fieldNames, fieldNames.map((name) => values.get(name) ?? null));
}

function rowsCall(table: MTable, fn: MFunction): (row: MValue[]) => MValue {
  return (row) => invoke(fn, [rowRecord(table, row)]);
}

/** A count, or a condition kept while true (Table.FirstN / Table.Skip). */
function countOrCondition(table: MTable, value: MValue | undefined, fallback: number): number {
  if (value === undefined || value === null) return fallback;
  if (value instanceof MFunction) {
    const test = rowsCall(table, value);
    let count = 0;
    while (count < table.rows.length && test(table.rows[count]) === true) count += 1;
    return count;
  }
  return Math.max(0, Math.floor(toNumber(value)));
}

function sortCriteria(table: MTable, criteria: MValue): (left: MValue[], right: MValue[]) => number {
  const list = typeof criteria === "string" || criteria instanceof MFunction ? [criteria] : toList(criteria);
  // {"A", Order.Descending} is one criterion, not two columns.
  const items = list.length === 2 && typeof list[0] === "string" && typeof list[1] === "number" ? [list] : list;
  const comparers = items.map((item): ((left: MValue[], right: MValue[]) => number) => {
    if (typeof item === "string") {
      const index = columnIndex(table, item);
      return (left, right) => compareValues(left[index], right[index]);
    }
    if (item instanceof MFunction) {
      const key = rowsCall(table, item);
      if (item.maxArgs >= 2 && item.minArgs === 2)
        return (left, right) => toNumber(invoke(item, [rowRecord(table, left), rowRecord(table, right)]));
      return (left, right) => compareValues(key(left), key(right));
    }
    const pair = toList(item);
    const order = pair[1] === undefined ? 0 : pair[1];
    if (pair[0] instanceof MFunction) {
      const key = rowsCall(table, pair[0]);
      const sign = order === 1 ? -1 : 1;
      return (left, right) => sign * compareValues(key(left), key(right));
    }
    const index = columnIndex(table, toText(pair[0]));
    if (order instanceof MFunction) {
      const ignore = order.ignoreCase;
      return (left, right) => compareValues(left[index], right[index], ignore);
    }
    const sign = toNumber(order) === 1 ? -1 : 1;
    return (left, right) => sign * compareValues(left[index], right[index]);
  });
  return (left, right) => {
    for (const compare of comparers) {
      const result = compare(left, right);
      if (result !== 0) return result;
    }
    return 0;
  };
}

function convertTo(value: MValue, type: MType, culture: Culture): MValue {
  if (value === null) return null;
  switch (type.kind) {
    case "any":
    case "anynonnull":
      return value;
    case "text":
      return textFrom(value, culture);
    case "number": {
      const number = numberFrom(value, culture);
      if (number === null) return null;
      if (type.facet === "Int64") return roundToEven(number);
      if (type.facet === "Currency") return roundToEven(number, 4);
      return number;
    }
    case "logical":
      if (typeof value === "boolean") return value;
      if (typeof value === "number") return value !== 0;
      if (typeof value === "string") {
        if (/^true$/i.test(value.trim())) return true;
        if (/^false$/i.test(value.trim())) return false;
      }
      return fail(`We couldn't convert to Logical.`, "DataFormat.Error");
    case "date":
      return dateFrom(value, culture);
    case "datetime":
      return dateTimeFrom(value, culture);
    case "time":
      if (value instanceof MTime) return value;
      if (value instanceof MDateTime) return new MTime(value.serial - Math.floor(value.serial));
      if (typeof value === "number") return new MTime(value - Math.floor(value));
      return unsupported("conversion to time");
    case "duration":
      if (value instanceof MDuration) return value;
      if (typeof value === "number") return new MDuration(value);
      return unsupported("conversion to duration");
    case "table":
    case "list":
    case "record":
      if (!conformsTo(value, type)) fail(`We cannot convert ${describe(value)} to type ${type.kind}.`);
      return value;
    default:
      return unsupported(`conversion to type ${type.kind}`);
  }
}

function textFrom(value: MValue, culture: Culture): MValue {
  if (value === null) return null;
  if (typeof value === "string") return value;
  if (typeof value === "number") return numberText(value, culture);
  if (typeof value === "boolean") return value ? "true" : "false";
  if (value instanceof MDate) return formatDate(value.serial, culture.shortDate, culture);
  if (value instanceof MDateTime) return formatDate(value.serial, `${culture.shortDate} ${culture.longTime}`, culture);
  if (value instanceof MTime) return formatDate(value.fraction, culture.longTime, culture);
  if (value instanceof MDuration) {
    const sign = value.days < 0 ? "-" : "";
    const total = Math.abs(value.days);
    const days = Math.floor(total);
    const parts = partsOf(total - days);
    return `${sign}${days}.${String(parts.hour).padStart(2, "0")}:${String(parts.minute).padStart(2, "0")}:${String(parts.second).padStart(2, "0")}`;
  }
  return fail(`We cannot convert a value of type ${describe(value)} to type Text.`);
}
function numberFrom(value: MValue, culture: Culture): number | null {
  if (value === null) return null;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string") return parseNumberText(value, culture);
  if (value instanceof MDate || value instanceof MDateTime) return value.serial;
  if (value instanceof MTime) return value.fraction;
  if (value instanceof MDuration) return value.days;
  return fail(`We cannot convert a value of type ${describe(value)} to type Number.`);
}
function dateFrom(value: MValue, culture: Culture): MValue {
  if (value === null) return null;
  if (value instanceof MDate) return value;
  if (value instanceof MDateTime) return new MDate(Math.floor(value.serial));
  if (typeof value === "number") return new MDate(Math.floor(value));
  if (typeof value === "string") return new MDate(Math.floor(parseDateText(value, culture)));
  return fail(`We cannot convert the value ${describe(value)} to type Date.`);
}
function dateTimeFrom(value: MValue, culture: Culture): MValue {
  if (value === null) return null;
  if (value instanceof MDateTime) return value;
  if (value instanceof MDate) return new MDateTime(value.serial);
  if (typeof value === "number") return new MDateTime(value);
  if (typeof value === "string") return new MDateTime(parseDateText(value, culture));
  return fail(`We cannot convert the value ${describe(value)} to type DateTime.`);
}

function serialPart(value: MValue): number | null {
  if (value === null) return null;
  if (value instanceof MDate || value instanceof MDateTime) return value.serial;
  return fail(`We cannot convert the value ${describe(value)} to type Date.`);
}
function sameDateKind(source: MValue, serial: number): MValue {
  return source instanceof MDateTime ? new MDateTime(serial) : new MDate(Math.floor(serial));
}

function nonNull(list: MValue[]): MValue[] {
  return list.filter((item) => item !== null);
}

// ---------------------------------------------------------------------------
// Library
// ---------------------------------------------------------------------------

export interface PowerQueryWorkbookItem {
  name: string;
  columns: string[];
  rows: MValue[][];
}
export interface PowerQueryOptions {
  /** The query package's culture (Config/Package.xml), default en-US. */
  culture?: string;
  /** DateTime.LocalNow / FixedLocalNow; default the current local time. */
  now?: Date;
}

type Impl = (args: MValue[]) => MValue;

function buildLibrary(workbook: PowerQueryWorkbookItem[], options: PowerQueryOptions): Map<string, MValue> {
  const culture = cultureOf(options.culture ?? "en-US", CULTURES["en-us"]);
  const lib = new Map<string, MValue>();
  const def = (name: string, min: number, max: number, impl: Impl) => lib.set(name, new MFunction(name, impl, min, max));
  const constant = (name: string, value: MValue) => lib.set(name, value);
  const argCulture = (value: MValue | undefined) => (typeof value === "string" ? cultureOf(value, culture) : value instanceof MRecord && value.has("Culture") ? cultureOf(value.get("Culture") as string, culture) : culture);

  // Enumerations.
  ["Ascending", "Descending"].forEach((name, index) => constant(`Order.${name}`, index));
  ["Inner", "LeftOuter", "RightOuter", "FullOuter", "LeftAnti", "RightAnti"].forEach((name, index) => constant(`JoinKind.${name}`, index));
  ["Error", "Ignore", "UseNull"].forEach((name, index) => constant(`MissingField.${name}`, index));
  ["First", "Last", "All"].forEach((name, index) => constant(`Occurrence.${name}`, index));
  ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"].forEach((name, index) => constant(`Day.${name}`, index));
  ["Error", "Ignore", "List"].forEach((name, index) => constant(`ExtraValues.${name}`, index));
  ["Up", "Down", "AwayFromZero", "TowardZero", "ToEven"].forEach((name, index) => constant(`RoundingMode.${name}`, index));
  constant("GroupKind.Global", 1);
  constant("GroupKind.Local", 0);
  constant("JoinAlgorithm.Dynamic", 0);
  constant("Number.PI", Math.PI);
  constant("Number.E", Math.E);
  constant("Number.PositiveInfinity", Infinity);
  constant("Number.NegativeInfinity", -Infinity);
  constant("#infinity", Infinity);
  constant("#nan", NaN);
  for (const [name, kind, facet] of [
    ["Int64.Type", "number", "Int64"], ["Int32.Type", "number", "Int64"], ["Int16.Type", "number", "Int64"],
    ["Int8.Type", "number", "Int64"], ["Currency.Type", "number", "Currency"], ["Percentage.Type", "number", undefined],
    ["Decimal.Type", "number", undefined], ["Double.Type", "number", undefined], ["Single.Type", "number", undefined],
    ["Number.Type", "number", undefined], ["Text.Type", "text", undefined], ["Date.Type", "date", undefined],
    ["DateTime.Type", "datetime", undefined], ["Time.Type", "time", undefined], ["Duration.Type", "duration", undefined],
    ["Logical.Type", "logical", undefined], ["Any.Type", "any", undefined], ["Table.Type", "table", undefined],
    ["List.Type", "list", undefined], ["Record.Type", "record", undefined],
  ] as Array<[string, string, string | undefined]>)
    constant(name, new MType(kind, false, facet));
  lib.set("Comparer.Ordinal", new MFunction("Comparer.Ordinal", ([a, b]) => compareValues(a, b), 2, 2));
  lib.set("Comparer.OrdinalIgnoreCase", new MFunction("Comparer.OrdinalIgnoreCase", ([a, b]) => compareValues(a, b, true), 2, 2, true));
  lib.set("Replacer.ReplaceValue", new MFunction("Replacer.ReplaceValue", ([value, old, replacement]) => (equals(value, old) ? replacement : value), 3, 3));
  lib.set("Replacer.ReplaceText", new MFunction("Replacer.ReplaceText", ([value, old, replacement]) => {
    if (value === null) return null;
    const text = toText(value);
    const from = toText(old);
    return from ? text.split(from).join(toText(replacement ?? "")) : text;
  }, 3, 3));

  // Excel.
  def("Excel.CurrentWorkbook", 0, 0, () =>
    new MTable(["Content", "Name"], workbook.map((item) => [new MTable(item.columns.slice(), item.rows), item.name])),
  );
  // #table / literals.
  def("#table", 2, 2, ([columns, rows]) => {
    const rowList = toList(rows).map((row) => toList(row));
    let columnNames: string[];
    if (typeof columns === "number") columnNames = Array.from({ length: columns }, (_, index) => `Column${index + 1}`);
    else if (columns instanceof MType) columnNames = (columns.fields ?? []).map(([name]) => name);
    else columnNames = names(columns);
    uniqueColumns(columnNames);
    rowList.forEach((row) => row.length !== columnNames.length && fail("The number of values in a row doesn't match the number of columns."));
    return new MTable(columnNames, rowList.map((row) => row.slice()));
  });
  def("#date", 3, 3, ([y, m, d]) => new MDate(serialOf(toNumber(y), toNumber(m), toNumber(d))));
  def("#datetime", 6, 6, ([y, m, d, hh, mm, ss]) =>
    new MDateTime(serialOf(toNumber(y), toNumber(m), toNumber(d)) + (toNumber(hh) * 3600 + toNumber(mm) * 60 + toNumber(ss)) / 86400));
  def("#time", 3, 3, ([hh, mm, ss]) => new MTime((toNumber(hh) * 3600 + toNumber(mm) * 60 + toNumber(ss)) / 86400));
  def("#duration", 4, 4, ([d, hh, mm, ss]) => new MDuration(toNumber(d) + (toNumber(hh) * 3600 + toNumber(mm) * 60 + toNumber(ss)) / 86400));

  // ---- Table
  def("Table.TransformColumnTypes", 2, 3, ([value, transforms, cultureArg]) => {
    const table = toTable(value);
    const local = argCulture(cultureArg);
    let list = toList(transforms);
    if (list.length === 2 && typeof list[0] === "string" && list[1] instanceof MType) list = [list];
    const plan = list.map((entry) => {
      const pair = toList(entry);
      const type = pair[1];
      if (!(type instanceof MType)) fail("A type was expected.");
      return [columnIndex(table, toText(pair[0])), type as MType] as const;
    });
    return new MTable(table.columns.slice(), table.rows.map((row) => {
      const next = row.slice();
      for (const [index, type] of plan) next[index] = convertTo(next[index], type, local);
      return next;
    }));
  });
  def("Table.TransformColumns", 2, 4, ([value, transforms, defaultTransform, missing]) => {
    const table = toTable(value);
    let list = toList(transforms);
    if (list.length >= 2 && typeof list[0] === "string" && list[1] instanceof MFunction) list = [list];
    const plan: Array<[number, MFunction, MType | undefined]> = [];
    for (const entry of list) {
      const pair = toList(entry);
      const index = table.columns.indexOf(toText(pair[0]));
      if (index < 0) {
        if (missing === 1 || missing === 2) continue;
        fail(`The column '${pair[0]}' of the table wasn't found.`);
      }
      plan.push([index, toFunction(pair[1]), pair[2] instanceof MType ? pair[2] : undefined]);
    }
    const fallback = defaultTransform instanceof MFunction ? defaultTransform : null;
    return new MTable(table.columns.slice(), table.rows.map((row) => row.map((cell, index) => {
      const own = plan.find(([target]) => target === index);
      if (own) return invoke(own[1], [cell]);
      return fallback ? invoke(fallback, [cell]) : cell;
    })));
  });
  def("Table.TransformColumnNames", 2, 3, ([value, fn]) => {
    const table = toTable(value);
    const columns = table.columns.map((name) => toText(invoke(toFunction(fn), [name])));
    uniqueColumns(columns);
    return new MTable(columns, table.rows);
  });
  def("Table.SelectRows", 2, 2, ([value, condition]) => {
    const table = toTable(value);
    const test = rowsCall(table, toFunction(condition));
    return new MTable(table.columns, table.rows.filter((row) => {
      const result = test(row);
      if (result !== null && typeof result !== "boolean") fail(`We cannot convert the value ${describe(result)} to type Logical.`);
      return result === true;
    }));
  });
  def("Table.RemoveColumns", 2, 3, ([value, columns, missing]) => {
    const table = toTable(value);
    const removed = names(columns);
    for (const name of removed) if (!table.columns.includes(name) && !(missing === 1 || missing === 2)) fail(`The column '${name}' of the table wasn't found.`);
    return selectColumns(table, table.columns.filter((name) => !removed.includes(name)));
  });
  def("Table.SelectColumns", 2, 3, ([value, columns, missing]) =>
    selectColumns(toTable(value), names(columns), typeof missing === "number" ? missing : 0));
  def("Table.RenameColumns", 2, 3, ([value, renames, missing]) => {
    const table = toTable(value);
    let list = toList(renames);
    if (list.length === 2 && typeof list[0] === "string") list = [list];
    const columns = table.columns.slice();
    for (const entry of list) {
      const [from, to] = toList(entry).map(toText);
      const index = columns.indexOf(from);
      if (index < 0) {
        if (missing === 1 || missing === 2) continue;
        fail(`The column '${from}' of the table wasn't found.`);
      }
      columns[index] = to;
    }
    uniqueColumns(columns);
    return new MTable(columns, table.rows);
  });
  def("Table.ReorderColumns", 2, 3, ([value, order, missing]) => {
    const table = toTable(value);
    const wanted = names(order).filter((name) => {
      if (table.columns.includes(name)) return true;
      if (missing === 1 || missing === 2) return false;
      return fail(`The column '${name}' of the table wasn't found.`);
    });
    const slots = table.columns.map((name, index) => (wanted.includes(name) ? index : -1)).filter((index) => index >= 0);
    const columns = table.columns.slice();
    slots.forEach((slot, position) => (columns[slot] = wanted[position]));
    return selectColumns(table, columns);
  });
  def("Table.DuplicateColumn", 3, 4, ([value, column, newName]) => {
    const table = toTable(value);
    const index = columnIndex(table, toText(column));
    const columns = [...table.columns, toText(newName)];
    uniqueColumns(columns);
    return new MTable(columns, table.rows.map((row) => [...row, row[index]]));
  });
  def("Table.AddColumn", 3, 4, ([value, name, fn, type]) => {
    const table = toTable(value);
    const columns = [...table.columns, toText(name)];
    uniqueColumns(columns);
    const compute = rowsCall(table, toFunction(fn));
    void type;
    return new MTable(columns, table.rows.map((row) => [...row, compute(row)]));
  });
  def("Table.AddIndexColumn", 2, 5, ([value, name, start, step]) => {
    const table = toTable(value);
    const columns = [...table.columns, toText(name)];
    uniqueColumns(columns);
    const from = optionalNumber(start, 0);
    const by = optionalNumber(step, 1);
    return new MTable(columns, table.rows.map((row, index) => [...row, from + index * by]));
  });
  def("Table.Sort", 2, 2, ([value, criteria]) => {
    const table = toTable(value);
    const compare = sortCriteria(table, criteria);
    return new MTable(table.columns, table.rows.map((row, index) => [row, index] as const)
      .sort((left, right) => compare(left[0], right[0]) || left[1] - right[1])
      .map(([row]) => row));
  });
  def("Table.Distinct", 1, 2, ([value, criteria]) => {
    const table = toTable(value);
    let indexes = table.columns.map((_, index) => index);
    let ignoreCase = false;
    if (criteria !== undefined && criteria !== null) {
      if (criteria instanceof MFunction) ignoreCase = criteria.ignoreCase;
      else {
        let list = typeof criteria === "string" ? [criteria] : toList(criteria);
        if (list.length === 2 && typeof list[0] === "string" && list[1] instanceof MFunction) {
          ignoreCase = list[1].ignoreCase;
          list = [list[0]];
        }
        indexes = list.map((name) => columnIndex(table, toText(name)));
      }
    }
    const seen = new Set<string>();
    return new MTable(table.columns, table.rows.filter((row) => {
      const key = indexes.map((index) => keyOf(row[index], ignoreCase)).join("\u0000");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    }));
  });
  def("Table.Group", 3, 5, ([value, keys, aggregates, groupKind, comparer]) => {
    const table = toTable(value);
    const keyNames = names(keys);
    const keyIndexes = keyNames.map((name) => columnIndex(table, name));
    let aggList = toList(aggregates);
    if (aggList.length >= 2 && typeof aggList[0] === "string" && aggList[1] instanceof MFunction) aggList = [aggList];
    const aggs = aggList.map((entry) => {
      const parts = toList(entry);
      return [toText(parts[0]), toFunction(parts[1])] as const;
    });
    const ignoreCase = comparer instanceof MFunction && comparer.ignoreCase;
    const local = groupKind === 0;
    const groups: Array<{ key: MValue[]; rows: MValue[][] }> = [];
    const byKey = new Map<string, { key: MValue[]; rows: MValue[][] }>();
    let previous: string | null = null;
    for (const row of table.rows) {
      const key = keyIndexes.map((index) => row[index]);
      const id = key.map((part) => keyOf(part, ignoreCase)).join("\u0000");
      if (local) {
        if (id !== previous) groups.push({ key, rows: [] });
        groups[groups.length - 1].rows.push(row);
        previous = id;
        continue;
      }
      let group = byKey.get(id);
      if (!group) {
        group = { key, rows: [] };
        byKey.set(id, group);
        groups.push(group);
      }
      group.rows.push(row);
    }
    const columns = [...keyNames, ...aggs.map(([name]) => name)];
    uniqueColumns(columns);
    return new MTable(columns, groups.map((group) => {
      const sub = new MTable(table.columns, group.rows);
      return [...group.key, ...aggs.map(([, fn]) => invoke(fn, [sub]))];
    }));
  });
  def("Table.RowCount", 1, 1, ([value]) => toTable(value).rows.length);
  def("Table.ColumnCount", 1, 1, ([value]) => toTable(value).columns.length);
  def("Table.ColumnNames", 1, 1, ([value]) => toTable(value).columns.slice());
  def("Table.HasColumns", 2, 2, ([value, columns]) => names(columns).every((name) => toTable(value).columns.includes(name)));
  def("Table.Column", 2, 2, ([value, column]) => fieldAccess(toTable(value), toText(column), false));
  def("Table.IsEmpty", 1, 1, ([value]) => toTable(value).rows.length === 0);
  def("Table.Buffer", 1, 2, ([value]) => toTable(value));
  def("Table.StopFolding", 1, 1, ([value]) => toTable(value));
  def("Table.ToRows", 1, 1, ([value]) => toTable(value).rows.map((row) => row.slice()));
  def("Table.ToRecords", 1, 1, ([value]) => { const table = toTable(value); return table.rows.map((row) => rowRecord(table, row)); });
  def("Table.ToList", 1, 2, ([value, combiner]) => {
    const table = toTable(value);
    const fn = combiner instanceof MFunction ? combiner : null;
    return table.rows.map((row) => (fn ? invoke(fn, [row]) : row.map((cell) => textFrom(cell, culture)).join(",")));
  });
  def("Table.FromRows", 1, 2, ([rows, columns]) => {
    const list = toList(rows).map((row) => toList(row).slice());
    const width = list[0]?.length ?? 0;
    const columnNames = columns === undefined || columns === null ? Array.from({ length: width }, (_, index) => `Column${index + 1}`) : typeof columns === "number" ? Array.from({ length: columns }, (_, index) => `Column${index + 1}`) : columns instanceof MType ? (columns.fields ?? []).map(([name]) => name) : names(columns);
    list.forEach((row) => row.length !== columnNames.length && fail("The number of values in a row doesn't match the number of columns."));
    return new MTable(columnNames, list);
  });
  def("Table.FromRecords", 1, 3, ([records, columns]) => {
    const list = toList(records).map(toRecord);
    const columnNames = columns !== undefined && columns !== null ? names(columns) : list[0]?.names.slice() ?? [];
    return new MTable(columnNames, list.map((record) => columnNames.map((name) => (record.has(name) ? record.get(name) : fail(`The field '${name}' of the record wasn't found.`)))));
  });
  def("Table.FromList", 1, 5, ([list, splitter, columns]) => {
    const items = toList(list);
    if (splitter !== undefined && splitter !== null && !(splitter instanceof MFunction && splitter.name === "Splitter.SplitByNothing")) {
      const split = toFunction(splitter);
      const rows = items.map((item) => toList(invoke(split, [item])));
      const width = typeof columns === "number" ? columns : columns ? names(columns).length : Math.max(0, ...rows.map((row) => row.length));
      const columnNames = columns && typeof columns !== "number" ? names(columns) : Array.from({ length: width }, (_, index) => `Column${index + 1}`);
      return new MTable(columnNames, rows.map((row) => columnNames.map((_, index) => row[index] ?? null)));
    }
    const columnNames = columns && typeof columns !== "number" ? names(columns) : ["Column1"];
    return new MTable(columnNames, items.map((item) => [item]));
  });
  lib.set("Splitter.SplitByNothing", new MFunction("Splitter.SplitByNothing", () => new MFunction("Splitter.SplitByNothing", ([value]) => [value], 1, 1), 0, 0));
  def("Splitter.SplitTextByDelimiter", 1, 2, ([delimiter]) =>
    new MFunction("splitter", ([value]) => (value === null ? [null] : toText(value).split(toText(delimiter))), 1, 1));
  def("Combiner.CombineTextByDelimiter", 1, 2, ([delimiter]) =>
    new MFunction("combiner", ([value]) => toList(value).filter((item) => item !== null).map((item) => textFrom(item, culture) as string).join(toText(delimiter)), 1, 1));
  def("Table.First", 1, 2, ([value, fallback]) => { const table = toTable(value); return table.rows.length ? rowRecord(table, table.rows[0]) : fallback ?? null; });
  def("Table.Last", 1, 2, ([value, fallback]) => { const table = toTable(value); return table.rows.length ? rowRecord(table, table.rows[table.rows.length - 1]) : fallback ?? null; });
  def("Table.FirstN", 2, 2, ([value, count]) => { const table = toTable(value); return new MTable(table.columns, table.rows.slice(0, countOrCondition(table, count, 1))); });
  def("Table.Skip", 1, 2, ([value, count]) => { const table = toTable(value); return new MTable(table.columns, table.rows.slice(countOrCondition(table, count, 1))); });
  def("Table.RemoveFirstN", 1, 2, ([value, count]) => { const table = toTable(value); return new MTable(table.columns, table.rows.slice(countOrCondition(table, count, 1))); });
  def("Table.LastN", 2, 2, ([value, count]) => {
    const table = toTable(value);
    if (count instanceof MFunction) {
      const test = rowsCall(table, count);
      let start = table.rows.length;
      while (start > 0 && test(table.rows[start - 1]) === true) start -= 1;
      return new MTable(table.columns, table.rows.slice(start));
    }
    const n = Math.max(0, Math.floor(toNumber(count)));
    return new MTable(table.columns, n ? table.rows.slice(-n) : []);
  });
  def("Table.RemoveLastN", 1, 2, ([value, count]) => {
    const table = toTable(value);
    if (count instanceof MFunction) {
      const test = rowsCall(table, count);
      let end = table.rows.length;
      while (end > 0 && test(table.rows[end - 1]) === true) end -= 1;
      return new MTable(table.columns, table.rows.slice(0, end));
    }
    const n = count === undefined || count === null ? 1 : Math.max(0, Math.floor(toNumber(count)));
    return new MTable(table.columns, table.rows.slice(0, Math.max(0, table.rows.length - n)));
  });
  def("Table.Range", 2, 3, ([value, offset, count]) => {
    const table = toTable(value);
    const start = toNumber(offset);
    return new MTable(table.columns, table.rows.slice(start, count === undefined || count === null ? undefined : start + toNumber(count)));
  });
  def("Table.RemoveRows", 2, 3, ([value, offset, count]) => {
    const table = toTable(value);
    const start = toNumber(offset);
    const n = count === undefined || count === null ? 1 : toNumber(count);
    return new MTable(table.columns, [...table.rows.slice(0, start), ...table.rows.slice(start + n)]);
  });
  def("Table.ReverseRows", 1, 1, ([value]) => { const table = toTable(value); return new MTable(table.columns, table.rows.slice().reverse()); });
  def("Table.AlternateRows", 4, 4, ([value, offset, skip, take]) => {
    const table = toTable(value);
    const start = toNumber(offset), s = toNumber(skip), t = toNumber(take);
    const rows = table.rows.slice(0, start);
    for (let index = start; index < table.rows.length; index += s + t) rows.push(...table.rows.slice(index + s, index + s + t));
    return new MTable(table.columns, rows);
  });
  const fill = (direction: 1 | -1) => ([value, columns]: MValue[]) => {
    const table = toTable(value);
    const indexes = names(columns).map((name) => columnIndex(table, name));
    const rows = table.rows.map((row) => row.slice());
    const order = direction === 1 ? rows : rows.slice().reverse();
    for (const index of indexes) {
      let last: MValue = null;
      for (const row of order) {
        if (row[index] === null) row[index] = last;
        else last = row[index];
      }
    }
    return new MTable(table.columns, rows);
  };
  def("Table.FillDown", 2, 2, fill(1));
  def("Table.FillUp", 2, 2, fill(-1));
  def("Table.ReplaceValue", 5, 5, ([value, old, replacement, replacer, columns]) => {
    const table = toTable(value);
    const fn = toFunction(replacer);
    const indexes = names(columns).map((name) => columnIndex(table, name));
    if (old instanceof MFunction || replacement instanceof MFunction) {
      return new MTable(table.columns, table.rows.map((row) => {
        const record = rowRecord(table, row);
        const from = old instanceof MFunction ? invoke(old, [record]) : old;
        const to = replacement instanceof MFunction ? invoke(replacement, [record]) : replacement;
        return row.map((cell, index) => (indexes.includes(index) ? invoke(fn, [cell, from, to]) : cell));
      }));
    }
    return new MTable(table.columns, table.rows.map((row) => row.map((cell, index) => (indexes.includes(index) ? invoke(fn, [cell, old, replacement]) : cell))));
  });
  def("Table.ReplaceErrorValues", 2, 2, ([value]) => toTable(value));
  def("Table.RemoveRowsWithErrors", 1, 2, ([value]) => toTable(value));
  def("Table.SelectRowsWithErrors", 1, 2, ([value]) => { const table = toTable(value); return new MTable(table.columns, []); });
  def("Table.Combine", 1, 2, ([tables, columns]) => {
    const combined = combineTables(toList(tables).map(toTable));
    return columns === undefined || columns === null ? combined : selectColumns(combined, names(columns), 2);
  });
  def("Table.PrefixColumns", 2, 2, ([value, prefix]) => {
    const table = toTable(value);
    return new MTable(table.columns.map((name) => `${toText(prefix)}.${name}`), table.rows);
  });
  def("Table.Transpose", 1, 2, ([value, columns]) => {
    const table = toTable(value);
    const width = table.rows.length;
    const columnNames = columns === undefined || columns === null ? Array.from({ length: width }, (_, index) => `Column${index + 1}`) : names(columns);
    return new MTable(columnNames, table.columns.map((_, col) => table.rows.map((row) => row[col])));
  });
  def("Table.PromoteHeaders", 1, 2, ([value, options]) => {
    const table = toTable(value);
    if (!table.rows.length) return table;
    const all = options instanceof MRecord && options.has("PromoteAllScalars") && options.get("PromoteAllScalars") === true;
    const optionCulture = options instanceof MRecord && options.has("Culture") ? cultureOf(toText(options.get("Culture")), culture) : culture;
    const raw = table.rows[0].map((cell, index) => {
      if (typeof cell === "string" && cell !== "") return cell;
      if (all && cell !== null && cell !== "" && !(cell instanceof MRecord) && !(cell instanceof MTable) && !Array.isArray(cell)) return textFrom(cell, optionCulture) as string;
      return table.columns[index];
    });
    const columns: string[] = [];
    for (const name of raw) {
      let candidate = name;
      for (let suffix = 1; columns.includes(candidate); suffix += 1) candidate = `${name}_${suffix}`;
      columns.push(candidate);
    }
    return new MTable(columns, table.rows.slice(1));
  });
  def("Table.DemoteHeaders", 1, 1, ([value]) => {
    const table = toTable(value);
    return new MTable(table.columns.map((_, index) => `Column${index + 1}`), [table.columns.slice(), ...table.rows]);
  });
  const join = (nested: boolean) => (args: MValue[]) => {
    const [leftValue, leftKeys, rightValue, rightKeys] = args;
    const newColumn = nested ? toText(args[4]) : null;
    const kind = (nested ? args[5] : args[4]) ?? 0;
    const left = toTable(leftValue);
    const right = toTable(rightValue);
    const leftIndexes = names(leftKeys).map((name) => columnIndex(left, name));
    const rightIndexes = names(rightKeys).map((name) => columnIndex(right, name));
    if (leftIndexes.length !== rightIndexes.length) fail("The number of keys doesn't match.");
    const byKey = new Map<string, number[]>();
    right.rows.forEach((row, index) => {
      const key = rightIndexes.map((col) => keyOf(row[col])).join("\u0000");
      const list = byKey.get(key);
      if (list) list.push(index);
      else byKey.set(key, [index]);
    });
    const used = new Set<number>();
    const pairs: Array<[MValue[] | null, number[]]> = [];
    for (const row of left.rows) {
      const key = leftIndexes.map((col) => keyOf(row[col])).join("\u0000");
      const matches = byKey.get(key) ?? [];
      matches.forEach((index) => used.add(index));
      if (kind === 0 && matches.length === 0) continue;
      if (kind === 2 && matches.length === 0) continue;
      if (kind === 4 && matches.length > 0) continue;
      if (kind === 5) continue;
      pairs.push([row, kind === 4 ? [] : matches]);
    }
    if (kind === 2 || kind === 3 || kind === 5)
      right.rows.forEach((_, index) => !used.has(index) && pairs.push([null, [index]]));
    if (typeof kind !== "number" || kind < 0 || kind > 5) fail("The join kind isn't valid.");
    if (nested) {
      const columns = [...left.columns, newColumn!];
      uniqueColumns(columns);
      return new MTable(columns, pairs.map(([row, matches]) => [
        ...(row ?? left.columns.map(() => null)),
        new MTable(right.columns, matches.map((index) => right.rows[index])),
      ]));
    }
    const columns = [...left.columns, ...right.columns];
    uniqueColumns(columns);
    const rows: MValue[][] = [];
    for (const [row, matches] of pairs) {
      const base = row ?? left.columns.map(() => null);
      if (!matches.length) rows.push([...base, ...right.columns.map(() => null)]);
      for (const index of matches) rows.push([...base, ...right.rows[index]]);
    }
    return new MTable(columns, rows);
  };
  def("Table.NestedJoin", 5, 7, join(true));
  def("Table.Join", 4, 6, join(false));
  def("Table.ExpandTableColumn", 3, 4, ([value, column, columnNames, newNames]) => {
    const table = toTable(value);
    const index = columnIndex(table, toText(column));
    const wanted = names(columnNames);
    const renamed = newNames === undefined || newNames === null ? wanted : names(newNames);
    const columns = [...table.columns.slice(0, index), ...renamed, ...table.columns.slice(index + 1)];
    uniqueColumns(columns);
    const rows: MValue[][] = [];
    for (const row of table.rows) {
      const nestedValue = row[index];
      const before = row.slice(0, index);
      const after = row.slice(index + 1);
      if (nestedValue === null) {
        rows.push([...before, ...wanted.map(() => null), ...after]);
        continue;
      }
      const nested = toTable(nestedValue);
      const map = wanted.map((name) => nested.columns.indexOf(name));
      if (!nested.rows.length) rows.push([...before, ...wanted.map(() => null), ...after]);
      for (const inner of nested.rows) rows.push([...before, ...map.map((col) => (col < 0 ? null : inner[col])), ...after]);
    }
    return new MTable(columns, rows);
  });
  def("Table.ExpandRecordColumn", 3, 4, ([value, column, fieldNames, newNames]) => {
    const table = toTable(value);
    const index = columnIndex(table, toText(column));
    const wanted = names(fieldNames);
    const renamed = newNames === undefined || newNames === null ? wanted : names(newNames);
    const columns = [...table.columns.slice(0, index), ...renamed, ...table.columns.slice(index + 1)];
    uniqueColumns(columns);
    return new MTable(columns, table.rows.map((row) => {
      const record = row[index];
      const values = wanted.map((name) => (record === null ? null : toRecord(record).has(name) ? toRecord(record).get(name) : null));
      return [...row.slice(0, index), ...values, ...row.slice(index + 1)];
    }));
  });
  def("Table.ExpandListColumn", 2, 2, ([value, column]) => {
    const table = toTable(value);
    const index = columnIndex(table, toText(column));
    const rows: MValue[][] = [];
    for (const row of table.rows) {
      const list = row[index];
      if (list === null || toList(list).length === 0) rows.push([...row.slice(0, index), null, ...row.slice(index + 1)]);
      else for (const item of toList(list)) rows.push([...row.slice(0, index), item, ...row.slice(index + 1)]);
    }
    return new MTable(table.columns, rows);
  });
  def("Table.Pivot", 4, 5, ([value, pivotValues, attribute, valueColumn, aggregate]) => {
    const table = toTable(value);
    const attributeIndex = columnIndex(table, toText(attribute));
    const valueIndex = columnIndex(table, toText(valueColumn));
    const pivots = toList(pivotValues);
    const pivotNames = pivots.map((item) => toText(textFrom(item, culture)));
    const otherIndexes = table.columns.map((_, index) => index).filter((index) => index !== attributeIndex && index !== valueIndex);
    const columns = [...otherIndexes.map((index) => table.columns[index]), ...pivotNames];
    uniqueColumns(columns);
    const groups = new Map<string, { key: MValue[]; cells: MValue[][] }>();
    for (const row of table.rows) {
      const key = otherIndexes.map((index) => row[index]);
      const id = key.map((part) => keyOf(part)).join("\u0000");
      let group = groups.get(id);
      if (!group) {
        group = { key, cells: pivots.map(() => []) };
        groups.set(id, group);
      }
      const slot = pivots.findIndex((item) => equals(item, row[attributeIndex]));
      if (slot >= 0) group.cells[slot].push(row[valueIndex]);
    }
    const fn = aggregate instanceof MFunction ? aggregate : null;
    const ordered = [...groups.values()].sort((left, right) => {
      for (let index = 0; index < left.key.length; index += 1) {
        const order = compareValues(left.key[index], right.key[index]);
        if (order) return order;
      }
      return 0;
    });
    return new MTable(columns, ordered.map((group) => [
      ...group.key,
      ...group.cells.map((cells) => {
        if (!cells.length) return null;
        if (fn) return invoke(fn, [cells]);
        if (cells.length > 1) fail("There were too many elements in the enumeration to complete the operation.");
        return cells[0];
      }),
    ]));
  });
  const unpivot = (args: MValue[], other: boolean) => {
    const table = toTable(args[0]);
    const listed = names(args[1]);
    listed.forEach((name) => columnIndex(table, name));
    const attribute = toText(args[2]);
    const valueName = toText(args[3]);
    const pivoted = table.columns.map((name, index) => [name, index] as const).filter(([name]) => (other ? !listed.includes(name) : listed.includes(name)));
    const kept = table.columns.map((name, index) => [name, index] as const).filter(([name]) => (other ? listed.includes(name) : !listed.includes(name)));
    const columns = [...kept.map(([name]) => name), attribute, valueName];
    uniqueColumns(columns);
    const rows: MValue[][] = [];
    for (const row of table.rows)
      for (const [name, index] of pivoted)
        if (row[index] !== null) rows.push([...kept.map(([, keep]) => row[keep]), name, row[index]]);
    return new MTable(columns, rows);
  };
  def("Table.UnpivotOtherColumns", 4, 4, (args) => unpivot(args, true));
  def("Table.Unpivot", 4, 4, (args) => unpivot(args, false));
  def("Table.Max", 2, 3, ([value, column, fallback]) => {
    const table = toTable(value);
    const index = columnIndex(table, toText(column));
    if (!table.rows.length) return fallback ?? null;
    return rowRecord(table, table.rows.reduce((best, row) => (compareValues(row[index], best[index]) > 0 ? row : best)));
  });
  def("Table.Min", 2, 3, ([value, column, fallback]) => {
    const table = toTable(value);
    const index = columnIndex(table, toText(column));
    if (!table.rows.length) return fallback ?? null;
    return rowRecord(table, table.rows.reduce((best, row) => (compareValues(row[index], best[index]) < 0 ? row : best)));
  });

  // ---- List
  const numbers = (value: MValue) => nonNull(toList(value));
  def("List.Sum", 1, 2, ([value]) => {
    const list = numbers(value);
    if (!list.length) return null;
    if (list.every((item) => item instanceof MDuration)) return new MDuration(list.reduce((sum: number, item) => sum + (item as MDuration).days, 0));
    return list.reduce((sum: number, item) => sum + toNumber(item), 0);
  });
  def("List.Product", 1, 2, ([value]) => { const list = numbers(value); return list.length ? list.reduce((product: number, item) => product * toNumber(item), 1) : null; });
  def("List.Average", 1, 2, ([value]) => { const list = numbers(value); return list.length ? list.reduce((sum: number, item) => sum + toNumber(item), 0) / list.length : null; });
  def("List.Median", 1, 2, ([value]) => {
    const list = numbers(value).slice().sort((a, b) => compareValues(a, b));
    if (!list.length) return null;
    const middle = Math.floor(list.length / 2);
    return list.length % 2 ? list[middle] : (toNumber(list[middle - 1]) + toNumber(list[middle])) / 2;
  });
  def("List.Count", 1, 1, ([value]) => toList(value).length);
  def("List.NonNullCount", 1, 1, ([value]) => numbers(value).length);
  const extreme = (sign: 1 | -1) => ([value, fallback]: MValue[]) => {
    const list = numbers(value);
    if (!list.length) return fallback ?? null;
    return list.reduce((best, item) => (sign * compareStrict(item, best) > 0 ? item : best));
  };
  def("List.Max", 1, 4, extreme(1));
  def("List.Min", 1, 4, extreme(-1));
  def("List.Distinct", 1, 2, ([value, comparer]) => {
    const ignoreCase = comparer instanceof MFunction && comparer.ignoreCase;
    const seen = new Set<string>();
    return toList(value).filter((item) => {
      const key = keyOf(item, ignoreCase);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  });
  def("List.First", 1, 2, ([value, fallback]) => { const list = toList(value); return list.length ? list[0] : fallback ?? null; });
  def("List.Last", 1, 2, ([value, fallback]) => { const list = toList(value); return list.length ? list[list.length - 1] : fallback ?? null; });
  def("List.FirstN", 2, 2, ([value, count]) => toList(value).slice(0, toNumber(count)));
  def("List.LastN", 2, 2, ([value, count]) => { const n = toNumber(count); return n ? toList(value).slice(-n) : []; });
  def("List.Skip", 1, 2, ([value, count]) => toList(value).slice(optionalNumber(count, 1)));
  def("List.RemoveFirstN", 1, 2, ([value, count]) => toList(value).slice(optionalNumber(count, 1)));
  def("List.RemoveLastN", 1, 2, ([value, count]) => { const list = toList(value); return list.slice(0, Math.max(0, list.length - optionalNumber(count, 1))); });
  def("List.Range", 2, 3, ([value, offset, count]) => { const start = toNumber(offset); return toList(value).slice(start, count === undefined || count === null ? undefined : start + toNumber(count)); });
  def("List.Reverse", 1, 1, ([value]) => toList(value).slice().reverse());
  def("List.IsEmpty", 1, 1, ([value]) => toList(value).length === 0);
  def("List.Contains", 2, 3, ([value, item, comparer]) => toList(value).some((entry) => equals(entry, item, comparer instanceof MFunction && comparer.ignoreCase)));
  def("List.ContainsAny", 2, 3, ([value, items, comparer]) => toList(items).some((item) => toList(value).some((entry) => equals(entry, item, comparer instanceof MFunction && comparer.ignoreCase))));
  def("List.ContainsAll", 2, 3, ([value, items, comparer]) => toList(items).every((item) => toList(value).some((entry) => equals(entry, item, comparer instanceof MFunction && comparer.ignoreCase))));
  def("List.PositionOf", 2, 4, ([value, item]) => toList(value).findIndex((entry) => equals(entry, item)));
  def("List.Select", 2, 2, ([value, fn]) => toList(value).filter((item) => invoke(toFunction(fn), [item]) === true));
  def("List.Transform", 2, 2, ([value, fn]) => toList(value).map((item) => invoke(toFunction(fn), [item])));
  def("List.Accumulate", 3, 3, ([value, seed, fn]) => toList(value).reduce((state, item) => invoke(toFunction(fn), [state, item]), seed));
  def("List.Sort", 1, 2, ([value, criteria]) => {
    const list = toList(value).slice();
    let sign = 1;
    let ignoreCase = false;
    if (typeof criteria === "number") sign = criteria === 1 ? -1 : 1;
    else if (criteria instanceof MFunction) {
      if (criteria.minArgs === 2) return list.map((item, index) => [item, index] as const).sort((a, b) => toNumber(invoke(criteria, [a[0], b[0]])) || a[1] - b[1]).map(([item]) => item);
      return list.map((item, index) => [item, index, invoke(criteria, [item])] as const).sort((a, b) => compareValues(a[2], b[2]) || a[1] - b[1]).map(([item]) => item);
    } else if (Array.isArray(criteria)) {
      if (criteria[1] === 1) sign = -1;
      if (criteria[1] instanceof MFunction) ignoreCase = criteria[1].ignoreCase;
    }
    return list.map((item, index) => [item, index] as const).sort((a, b) => sign * compareValues(a[0], b[0], ignoreCase) || a[1] - b[1]).map(([item]) => item);
  });
  def("List.Combine", 1, 1, ([value]) => toList(value).flatMap((item) => toList(item)));
  def("List.RemoveNulls", 1, 1, ([value]) => numbers(value));
  def("List.RemoveMatchingItems", 2, 3, ([value, items]) => toList(value).filter((item) => !toList(items).some((other) => equals(item, other))));
  def("List.RemoveItems", 2, 2, ([value, items]) => toList(value).filter((item) => !toList(items).some((other) => equals(item, other))));
  def("List.Numbers", 2, 3, ([start, count, step]) => Array.from({ length: Math.max(0, toNumber(count)) }, (_, index) => toNumber(start) + index * optionalNumber(step, 1)));
  def("List.Repeat", 2, 2, ([value, count]) => Array.from({ length: toNumber(count) }, () => toList(value)).flat());
  def("List.Zip", 1, 1, ([value]) => {
    const lists = toList(value).map(toList);
    const width = Math.max(0, ...lists.map((list) => list.length));
    return Array.from({ length: width }, (_, index) => lists.map((list) => list[index] ?? null));
  });
  def("List.Dates", 3, 3, ([start, count, step]) => {
    const from = serialPart(start)!;
    return Array.from({ length: toNumber(count) }, (_, index) => sameDateKind(start, from + index * (step as MDuration).days));
  });
  def("List.Union", 1, 2, ([value]) => {
    const out: MValue[] = [];
    for (const list of toList(value)) for (const item of toList(list)) if (!out.some((other) => equals(other, item))) out.push(item);
    return out;
  });
  def("List.Intersect", 1, 2, ([value]) => {
    const lists = toList(value).map(toList);
    if (!lists.length) return [];
    return lists[0].filter((item, index, self) => self.findIndex((other) => equals(other, item)) === index && lists.every((list) => list.some((other) => equals(other, item))));
  });
  def("List.Difference", 2, 3, ([value, other]) => toList(value).filter((item) => !toList(other).some((entry) => equals(entry, item))));
  def("List.Generate", 3, 4, () => unsupported("List.Generate"));

  // ---- Record
  def("Record.Field", 2, 2, ([record, name]) => toRecord(record).get(toText(name)));
  def("Record.FieldOrDefault", 2, 3, ([record, name, fallback]) => { const value = toRecord(record); return value.has(toText(name)) ? value.get(toText(name)) : fallback ?? null; });
  def("Record.FieldNames", 1, 1, ([record]) => toRecord(record).names.slice());
  def("Record.FieldValues", 1, 1, ([record]) => toRecord(record).values());
  def("Record.FieldCount", 1, 1, ([record]) => toRecord(record).names.length);
  def("Record.HasFields", 2, 2, ([record, fields]) => names(fields).every((name) => toRecord(record).has(name)));
  def("Record.ToTable", 1, 1, ([record]) => { const value = toRecord(record); return new MTable(["Name", "Value"], value.names.map((name) => [name, value.get(name)])); });
  def("Record.ToList", 1, 1, ([record]) => toRecord(record).values());
  def("Record.FromList", 2, 2, ([values, fields]) => MRecord.of(names(fields), toList(values)));
  def("Record.AddField", 3, 4, ([record, name, value]) => combineRecords([toRecord(record), MRecord.of([toText(name)], [value])]));
  def("Record.Combine", 1, 1, ([records]) => combineRecords(toList(records).map(toRecord)));
  def("Record.RemoveFields", 2, 3, ([record, fields]) => { const value = toRecord(record); const drop = names(fields); const kept = value.names.filter((name) => !drop.includes(name)); return MRecord.of(kept, kept.map((name) => value.get(name))); });
  def("Record.SelectFields", 2, 3, ([record, fields]) => { const value = toRecord(record); const kept = names(fields); return MRecord.of(kept, kept.map((name) => value.get(name))); });

  // ---- Text
  const text = (value: MValue) => (value === null ? null : toText(value));
  const textFn = (name: string, min: number, max: number, impl: (value: string, args: MValue[]) => MValue) =>
    def(name, min, max, ([value, ...rest]) => (value === null ? null : impl(toText(value), rest)));
  def("Text.From", 1, 2, ([value, cultureArg]) => textFrom(value, argCulture(cultureArg)));
  textFn("Text.Length", 1, 1, (value) => value.length);
  textFn("Text.Upper", 1, 2, (value) => value.toUpperCase());
  textFn("Text.Lower", 1, 2, (value) => value.toLowerCase());
  textFn("Text.Proper", 1, 2, (value) => value.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, before: string, letter: string) => before + letter.toUpperCase()));
  const trimChars = (args: MValue[]) => (args[0] === undefined || args[0] === null ? null : typeof args[0] === "string" ? [args[0]] : toList(args[0]).map(toText));
  const trimWith = (value: string, chars: string[] | null, start: boolean, end: boolean) => {
    const test = (char: string) => (chars ? chars.includes(char) : /\s/.test(char));
    let from = 0, to = value.length;
    if (start) while (from < to && test(value[from])) from += 1;
    if (end) while (to > from && test(value[to - 1])) to -= 1;
    return value.slice(from, to);
  };
  textFn("Text.Trim", 1, 2, (value, args) => trimWith(value, trimChars(args), true, true));
  textFn("Text.TrimStart", 1, 2, (value, args) => trimWith(value, trimChars(args), true, false));
  textFn("Text.TrimEnd", 1, 2, (value, args) => trimWith(value, trimChars(args), false, true));
  textFn("Text.Clean", 1, 1, (value) => value.replace(/[\u0000-\u001f\u007f-\u009f]/g, ""));
  textFn("Text.Start", 2, 2, (value, [count]) => value.slice(0, toNumber(count)));
  textFn("Text.End", 2, 2, (value, [count]) => { const n = toNumber(count); return n ? value.slice(-n) : ""; });
  textFn("Text.Middle", 2, 3, (value, [offset, count]) => {
    const start = toNumber(offset);
    if (start > value.length) fail("The 'offset' argument is out of range.");
    return value.slice(start, count === undefined || count === null ? undefined : start + toNumber(count));
  });
  textFn("Text.Range", 2, 3, (value, [offset, count]) => {
    const start = toNumber(offset);
    const n = count === undefined || count === null ? value.length - start : toNumber(count);
    if (start < 0 || start > value.length || start + n > value.length) fail("The 'offset' or 'count' argument is out of range.");
    return value.slice(start, start + n);
  });
  const textCompare = (comparer: MValue | undefined) => comparer instanceof MFunction && comparer.ignoreCase;
  textFn("Text.Contains", 2, 3, (value, [sub, comparer]) => (textCompare(comparer) ? value.toLowerCase().includes(toText(sub).toLowerCase()) : value.includes(toText(sub))));
  textFn("Text.StartsWith", 2, 3, (value, [sub, comparer]) => (textCompare(comparer) ? value.toLowerCase().startsWith(toText(sub).toLowerCase()) : value.startsWith(toText(sub))));
  textFn("Text.EndsWith", 2, 3, (value, [sub, comparer]) => (textCompare(comparer) ? value.toLowerCase().endsWith(toText(sub).toLowerCase()) : value.endsWith(toText(sub))));
  textFn("Text.Replace", 3, 3, (value, [old, replacement]) => { const from = toText(old); return from ? value.split(from).join(toText(replacement)) : value; });
  textFn("Text.Remove", 2, 2, (value, [chars]) => { const drop = typeof chars === "string" ? [chars] : toList(chars).map(toText); return [...value].filter((char) => !drop.includes(char)).join(""); });
  textFn("Text.Select", 2, 2, (value, [chars]) => { const keep = typeof chars === "string" ? [chars] : toList(chars).map(toText); return [...value].filter((char) => keep.includes(char)).join(""); });
  textFn("Text.Insert", 3, 3, (value, [offset, insert]) => value.slice(0, toNumber(offset)) + toText(insert) + value.slice(toNumber(offset)));
  textFn("Text.Repeat", 2, 2, (value, [count]) => value.repeat(toNumber(count)));
  textFn("Text.Reverse", 1, 1, (value) => [...value].reverse().join(""));
  textFn("Text.ToList", 1, 1, (value) => [...value]);
  textFn("Text.Split", 2, 2, (value, [separator]) => value.split(toText(separator)));
  textFn("Text.SplitAny", 2, 2, (value, [separators]) => { const set = toText(separators); return value.split(new RegExp(`[${set.replace(/[\\\]^-]/g, "\\$&")}]`)); });
  textFn("Text.PadStart", 2, 3, (value, [count, pad]) => value.padStart(toNumber(count), pad === undefined || pad === null ? " " : toText(pad)));
  textFn("Text.PadEnd", 2, 3, (value, [count, pad]) => value.padEnd(toNumber(count), pad === undefined || pad === null ? " " : toText(pad)));
  textFn("Text.PositionOf", 2, 4, (value, [sub, occurrence, comparer]) => {
    const haystack = textCompare(comparer) ? value.toLowerCase() : value;
    const needle = textCompare(comparer) ? toText(sub).toLowerCase() : toText(sub);
    if (occurrence === 1) return haystack.lastIndexOf(needle);
    if (occurrence === 2) {
      const all: number[] = [];
      for (let index = haystack.indexOf(needle); index >= 0 && needle; index = haystack.indexOf(needle, index + needle.length)) all.push(index);
      return all;
    }
    return haystack.indexOf(needle);
  });
  const delimiterIndex = (value: string, delimiter: string, index: MValue | undefined): number => {
    let skip = 0;
    let fromEnd = false;
    if (typeof index === "number") skip = index;
    else if (Array.isArray(index)) {
      skip = toNumber(index[0]);
      fromEnd = index[1] === 1;
    }
    if (fromEnd) {
      let position = value.length;
      for (let count = 0; count <= skip; count += 1) {
        position = value.lastIndexOf(delimiter, position - 1);
        if (position < 0) return -1;
      }
      return position;
    }
    let position = -delimiter.length;
    for (let count = 0; count <= skip; count += 1) {
      position = value.indexOf(delimiter, position + delimiter.length);
      if (position < 0) return -1;
    }
    return position;
  };
  textFn("Text.BeforeDelimiter", 2, 3, (value, [delimiter, index]) => {
    const position = delimiterIndex(value, toText(delimiter), index);
    return position < 0 ? value : value.slice(0, position);
  });
  textFn("Text.AfterDelimiter", 2, 3, (value, [delimiter, index]) => {
    const d = toText(delimiter);
    const position = delimiterIndex(value, d, index);
    return position < 0 ? "" : value.slice(position + d.length);
  });
  textFn("Text.BetweenDelimiters", 3, 5, (value, [start, end, startIndex, endIndex]) => {
    const s = toText(start);
    const position = delimiterIndex(value, s, startIndex);
    if (position < 0) return "";
    const rest = value.slice(position + s.length);
    const close = delimiterIndex(rest, toText(end), endIndex);
    return close < 0 ? rest : rest.slice(0, close);
  });
  def("Text.Combine", 1, 2, ([list, separator]) => nonNull(toList(list)).map((item) => toText(item)).join(separator === undefined || separator === null ? "" : toText(separator)));
  def("Text.Format", 2, 3, ([format, values]) => toText(format).replace(/#\{([^}]+)\}/g, (_, key: string) => {
    const value = Array.isArray(values) ? values[Number(key)] : toRecord(values).get(key);
    return (textFrom(value ?? null, culture) as string | null) ?? "";
  }));
  void text;

  // ---- Number
  const num = (name: string, min: number, max: number, impl: (value: number, args: MValue[]) => MValue) =>
    def(name, min, max, ([value, ...rest]) => (value === null ? null : impl(toNumber(value), rest)));
  def("Number.From", 1, 2, ([value, cultureArg]) => numberFrom(value, argCulture(cultureArg)));
  def("Number.FromText", 1, 2, ([value, cultureArg]) => (value === null ? null : parseNumberText(toText(value), argCulture(cultureArg))));
  for (const name of ["Int64.From", "Int32.From", "Int16.From", "Int8.From"])
    def(name, 1, 3, ([value, cultureArg]) => { const n = numberFrom(value, argCulture(cultureArg)); return n === null ? null : roundToEven(n); });
  for (const name of ["Decimal.From", "Double.From", "Single.From", "Percentage.From"])
    def(name, 1, 2, ([value, cultureArg]) => numberFrom(value, argCulture(cultureArg)));
  def("Currency.From", 1, 3, ([value, cultureArg]) => { const n = numberFrom(value, argCulture(cultureArg)); return n === null ? null : roundToEven(n, 4); });
  num("Number.Round", 1, 3, (value, [digits, mode]) => {
    const d = optionalNumber(digits, 0);
    const m = optionalNumber(mode, 4);
    if (m === 4) return roundToEven(value, d);
    if (m === 2) return roundAway(value, d);
    const scale = 10 ** d;
    if (m === 0) return Math.ceil(value * scale) / scale;
    if (m === 1) return Math.floor(value * scale) / scale;
    return Math.trunc(value * scale) / scale;
  });
  num("Number.RoundUp", 1, 2, (value, [digits]) => { const scale = 10 ** optionalNumber(digits, 0); return Math.ceil(value * scale) / scale; });
  num("Number.RoundDown", 1, 2, (value, [digits]) => { const scale = 10 ** optionalNumber(digits, 0); return Math.floor(value * scale) / scale; });
  num("Number.RoundTowardZero", 1, 2, (value, [digits]) => { const scale = 10 ** optionalNumber(digits, 0); return Math.trunc(value * scale) / scale; });
  num("Number.RoundAwayFromZero", 1, 2, (value, [digits]) => { const scale = 10 ** optionalNumber(digits, 0); const scaled = value * scale; return (scaled < 0 ? Math.floor(scaled) : Math.ceil(scaled)) / scale; });
  num("Number.Abs", 1, 1, (value) => Math.abs(value));
  num("Number.Sign", 1, 1, (value) => Math.sign(value));
  num("Number.Sqrt", 1, 1, (value) => Math.sqrt(value));
  num("Number.Exp", 1, 1, (value) => Math.exp(value));
  num("Number.Ln", 1, 1, (value) => Math.log(value));
  num("Number.Log10", 1, 1, (value) => Math.log10(value));
  num("Number.Log", 1, 2, (value, [base]) => (base === undefined || base === null ? Math.log(value) : Math.log(value) / Math.log(toNumber(base))));
  num("Number.Power", 2, 2, (value, [power]) => (power === null ? null : value ** toNumber(power)));
  num("Number.Mod", 2, 3, (value, [divisor]) => { if (divisor === null) return null; const d = toNumber(divisor); return d === 0 ? fail("Division by zero.") : value % d; });
  num("Number.IntegerDivide", 2, 3, (value, [divisor]) => { if (divisor === null) return null; const d = toNumber(divisor); return d === 0 ? fail("Division by zero.") : Math.trunc(value / d); });
  num("Number.IsEven", 1, 1, (value) => value % 2 === 0);
  num("Number.IsOdd", 1, 1, (value) => Math.abs(value % 2) === 1);
  num("Number.IsNaN", 1, 1, (value) => Number.isNaN(value));
  num("Number.ToText", 1, 3, (value, [format, cultureArg]) => {
    const local = argCulture(cultureArg);
    if (format === undefined || format === null || format === "") return numberText(value, local);
    const spec = /^([FfNnPpDd])(\d{0,2})$/.exec(toText(format));
    if (!spec) return unsupported(`Number.ToText format "${format}"`);
    const letter = spec[1].toUpperCase();
    const digits = spec[2] === "" ? (letter === "D" ? 0 : 2) : Number(spec[2]);
    if (letter === "D") return (value < 0 ? "-" : "") + String(Math.abs(Math.trunc(value))).padStart(digits, "0");
    const scaled = letter === "P" ? value * 100 : value;
    const rounded = roundAway(scaled, digits);
    const [digitsText, frac = ""] = Math.abs(rounded).toFixed(digits).split(".");
    let whole = digitsText;
    if (letter === "N" || letter === "P") whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, local.group);
    const body = frac ? `${whole}${local.decimal}${frac}` : whole;
    const sign = rounded < 0 ? "-" : "";
    if (letter === "P") {
      if (!local.english) unsupported(`percent format in ${local.name}`);
      return `${sign}${body}%`;
    }
    return sign + body;
  });
  def("Logical.From", 1, 2, ([value]) => convertTo(value, new MType("logical"), culture));
  def("Logical.FromText", 1, 1, ([value]) => convertTo(value, new MType("logical"), culture));
  def("Logical.ToText", 1, 1, ([value]) => (value === null ? null : value ? "true" : "false"));
  def("Value.Equals", 2, 3, ([a, b]) => equals(a, b));
  def("Value.Compare", 2, 3, ([a, b]) => Math.sign(compareValues(a, b)));
  def("Value.Is", 2, 2, ([value, type]) => conformsTo(value, type as MType));
  def("Value.FromText", 1, 2, ([value]) => {
    if (value === null) return null;
    const t = toText(value);
    if (/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(t.trim())) return Number(t);
    if (/^(true|false)$/i.test(t.trim())) return /^true$/i.test(t.trim());
    return t;
  });

  // ---- Date / DateTime / Duration / Time
  const dateFn = (name: string, max: number, impl: (serial: number, source: MValue, args: MValue[]) => MValue) =>
    def(name, 1, max, ([value, ...rest]) => (value === null ? null : impl(serialPart(value)!, value, rest)));
  def("Date.From", 1, 2, ([value, cultureArg]) => dateFrom(value, argCulture(cultureArg)));
  def("Date.FromText", 1, 2, ([value, cultureArg]) => (value === null ? null : new MDate(Math.floor(parseDateText(toText(value), argCulture(cultureArg))))));
  def("DateTime.From", 1, 2, ([value, cultureArg]) => dateTimeFrom(value, argCulture(cultureArg)));
  def("DateTime.FromText", 1, 2, ([value, cultureArg]) => (value === null ? null : new MDateTime(parseDateText(toText(value), argCulture(cultureArg)))));
  dateFn("DateTime.Date", 1, (serial) => new MDate(Math.floor(serial)));
  dateFn("DateTime.Time", 1, (serial) => new MTime(serial - Math.floor(serial)));
  dateFn("Date.Year", 1, (serial) => partsOf(serial).year);
  dateFn("Date.Month", 1, (serial) => partsOf(serial).month);
  dateFn("Date.Day", 1, (serial) => partsOf(serial).day);
  dateFn("Date.DayOfYear", 1, (serial) => { const p = partsOf(serial); return Math.floor(serial) - serialOf(p.year, 1, 1) + 1; });
  dateFn("Date.DaysInMonth", 1, (serial) => { const p = partsOf(serial); return new Date(Date.UTC(p.year, p.month, 0)).getUTCDate(); });
  dateFn("Date.IsLeapYear", 1, (serial) => { const y = partsOf(serial).year; return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; });
  dateFn("Date.QuarterOfYear", 1, (serial) => Math.ceil(partsOf(serial).month / 3));
  dateFn("Date.DayOfWeek", 2, (serial, _source, [first]) => { const start = optionalNumber(first, 0); return (partsOf(serial).dow - start + 7) % 7; });
  dateFn("Date.DayOfWeekName", 2, (serial, _source, [cultureArg]) => { if (!argCulture(cultureArg).english) unsupported("day names outside English"); return DAYS[partsOf(serial).dow]; });
  dateFn("Date.MonthName", 2, (serial, _source, [cultureArg]) => { if (!argCulture(cultureArg).english) unsupported("month names outside English"); return MONTHS[partsOf(serial).month - 1]; });
  dateFn("Date.WeekOfYear", 2, (serial, _source, [first]) => {
    const start = optionalNumber(first, 0);
    const p = partsOf(serial);
    const jan1 = serialOf(p.year, 1, 1);
    const offset = (partsOf(jan1).dow - start + 7) % 7;
    return Math.floor((Math.floor(serial) - jan1 + offset) / 7) + 1;
  });
  dateFn("Date.WeekOfMonth", 2, (serial, _source, [first]) => {
    const start = optionalNumber(first, 0);
    const p = partsOf(serial);
    const day1 = serialOf(p.year, p.month, 1);
    const offset = (partsOf(day1).dow - start + 7) % 7;
    return Math.floor((Math.floor(serial) - day1 + offset) / 7) + 1;
  });
  dateFn("Date.AddDays", 2, (serial, source, [days]) => sameDateKind(source, serial + toNumber(days)));
  dateFn("Date.AddWeeks", 2, (serial, source, [weeks]) => sameDateKind(source, serial + 7 * toNumber(weeks)));
  dateFn("Date.AddMonths", 2, (serial, source, [months]) => sameDateKind(source, addMonths(serial, toNumber(months))));
  dateFn("Date.AddQuarters", 2, (serial, source, [quarters]) => sameDateKind(source, addMonths(serial, 3 * toNumber(quarters))));
  dateFn("Date.AddYears", 2, (serial, source, [years]) => sameDateKind(source, addMonths(serial, 12 * toNumber(years))));
  dateFn("Date.StartOfMonth", 1, (serial, source) => { const p = partsOf(serial); return sameDateKind(source, serialOf(p.year, p.month, 1)); });
  dateFn("Date.EndOfMonth", 1, (serial, source) => { const p = partsOf(serial); const last = serialOf(p.year, p.month + 1, 1) - 1; return source instanceof MDateTime ? new MDateTime(last + (DAY_MS - 1) / DAY_MS) : new MDate(last); });
  dateFn("Date.StartOfYear", 1, (serial, source) => sameDateKind(source, serialOf(partsOf(serial).year, 1, 1)));
  dateFn("Date.EndOfYear", 1, (serial, source) => { const last = serialOf(partsOf(serial).year, 12, 31); return source instanceof MDateTime ? new MDateTime(last + (DAY_MS - 1) / DAY_MS) : new MDate(last); });
  dateFn("Date.StartOfQuarter", 1, (serial, source) => { const p = partsOf(serial); return sameDateKind(source, serialOf(p.year, Math.floor((p.month - 1) / 3) * 3 + 1, 1)); });
  dateFn("Date.EndOfQuarter", 1, (serial, source) => { const p = partsOf(serial); const last = serialOf(p.year, Math.floor((p.month - 1) / 3) * 3 + 4, 1) - 1; return source instanceof MDateTime ? new MDateTime(last + (DAY_MS - 1) / DAY_MS) : new MDate(last); });
  dateFn("Date.StartOfWeek", 2, (serial, source, [first]) => { const start = optionalNumber(first, 0); const back = (partsOf(serial).dow - start + 7) % 7; return sameDateKind(source, Math.floor(serial) - back); });
  dateFn("Date.EndOfWeek", 2, (serial, source, [first]) => { const start = optionalNumber(first, 0); const back = (partsOf(serial).dow - start + 7) % 7; return sameDateKind(source, Math.floor(serial) - back + 6); });
  dateFn("Date.ToText", 3, (serial, source, [format, cultureArg]) => {
    let local = culture;
    let pattern: string | null = null;
    if (format instanceof MRecord) {
      if (format.has("Culture")) local = argCulture(format.get("Culture"));
      if (format.has("Format")) pattern = format.get("Format") as string | null;
    } else {
      if (typeof format === "string") pattern = format;
      local = argCulture(cultureArg);
    }
    if (pattern === null) return textFrom(source instanceof MDateTime ? new MDate(Math.floor(serial)) : source, local);
    if (pattern.length === 1) return unsupported(`standard date format "${pattern}"`);
    return formatDate(serial, pattern, local);
  });
  dateFn("DateTime.ToText", 3, (serial, source, [format, cultureArg]) => {
      const pattern = format instanceof MRecord ? (format.has("Format") ? (format.get("Format") as string | null) : null) : (format as string | null | undefined) ?? null;
      const local = format instanceof MRecord && format.has("Culture") ? argCulture(format.get("Culture")) : argCulture(cultureArg);
      if (pattern === null) return textFrom(source, local);
      if (pattern.length === 1) return unsupported(`standard date format "${pattern}"`);
      return formatDate(serial, pattern, local);
    });
  dateFn("Time.Hour", 1, (serial) => partsOf(serial).hour);
  dateFn("Time.Minute", 1, (serial) => partsOf(serial).minute);
  dateFn("Time.Second", 1, (serial) => partsOf(serial).second);
  def("DateTime.LocalNow", 0, 0, () => {
    const now = options.now ?? new Date();
    return new MDateTime((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds()) - EPOCH) / DAY_MS);
  });
  lib.set("DateTime.FixedLocalNow", lib.get("DateTime.LocalNow")!);
  const duration = (name: string, impl: (days: number) => number) =>
    def(name, 1, 1, ([value]) => (value === null ? null : impl((value as MDuration).days)));
  duration("Duration.Days", (days) => Math.trunc(days));
  duration("Duration.Hours", (days) => Math.trunc((days - Math.trunc(days)) * 24));
  duration("Duration.Minutes", (days) => Math.trunc(((days * 24) % 1) * 60));
  duration("Duration.Seconds", (days) => Math.round(((days * 1440) % 1) * 60 * 1000) / 1000);
  duration("Duration.TotalDays", (days) => days);
  duration("Duration.TotalHours", (days) => days * 24);
  duration("Duration.TotalMinutes", (days) => days * 1440);
  duration("Duration.TotalSeconds", (days) => days * 86400);
  def("Duration.From", 1, 1, ([value]) => (value === null ? null : value instanceof MDuration ? value : new MDuration(toNumber(value))));
  def("Time.From", 1, 2, ([value]) => convertTo(value, new MType("time"), culture));
  return lib;
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

export interface PowerQueryMember {
  name: string;
  formula: string;
}

/**
 * Evaluate the section member `target` with `members` as the workbook's
 * other queries (Excel's Section1) and `workbook` as `Excel.CurrentWorkbook()`.
 */
export function evaluatePowerQuery(
  members: readonly PowerQueryMember[],
  target: string,
  workbook: readonly PowerQueryWorkbookItem[],
  options: PowerQueryOptions = {},
): MValue {
  const library = buildLibrary([...workbook], options);
  const global = new Env(null);
  for (const [name, value] of library) global.define(name, () => value);
  const section = new Env(global);
  for (const member of members) section.define(member.name, lazy(() => evaluate(parseM(member.formula), section)));
  const member = members.find((entry) => entry.name === target);
  if (!member) return fail(`The query '${target}' wasn't found.`);
  return section.lookup(target);
}

/** The library functions this interpreter runs (for the pane's support check). */
export function powerQuerySupportedFunctions(): string[] {
  return [...buildLibrary([], {}).entries()].filter(([, value]) => value instanceof MFunction).map(([name]) => name).sort();
}

/** One cell of a loaded query result, in the sheet's own terms. */
export type PowerQuerySheetValue =
  | null
  | { kind: "number"; value: number; format?: "date" | "datetime" | "time" }
  | { kind: "text"; value: string }
  | { kind: "logical"; value: boolean };

export interface PowerQuerySheetTable {
  columns: string[];
  rows: PowerQuerySheetValue[][];
}

/**
 * What Excel writes into a table when it loads a query result: column names
 * unique without regard to case (a later duplicate gets the next free number,
 * one counter for the whole load: "Cherry2", "banana3"), dates as serial
 * numbers with a date format, nested values refused.
 */
export function powerQueryResultToSheet(value: MValue): PowerQuerySheetTable {
  if (!(value instanceof MTable)) {
    if (value instanceof MRecord) return powerQueryResultToSheet(new MTable(value.names, [value.values()]));
    if (Array.isArray(value)) return powerQueryResultToSheet(new MTable(["List"], value.map((item) => [item])));
    return powerQueryResultToSheet(new MTable(["Column1"], [[value]]));
  }
  const columns: string[] = [];
  let counter = 2;
  for (const raw of value.columns) {
    let name = raw;
    if (columns.some((other) => other.toLowerCase() === name.toLowerCase())) {
      while (columns.some((other) => other.toLowerCase() === `${raw}${counter}`.toLowerCase())) counter += 1;
      name = `${raw}${counter}`;
      counter += 1;
    }
    columns.push(name);
  }
  const rows = value.rows.map((row) => row.map((cell): PowerQuerySheetValue => {
    if (cell === null) return null;
    if (typeof cell === "number") {
      if (!Number.isFinite(cell)) return unsupported("loading a non-finite number (NaN or infinity)");
      return { kind: "number", value: cell };
    }
    if (typeof cell === "string") return cell === "" ? null : { kind: "text", value: cell };
    if (typeof cell === "boolean") return { kind: "logical", value: cell };
    if (cell instanceof MDate) return { kind: "number", value: cell.serial, format: "date" };
    if (cell instanceof MDateTime) return { kind: "number", value: cell.serial, format: "datetime" };
    if (cell instanceof MTime) return { kind: "number", value: cell.fraction, format: "time" };
    if (cell instanceof MDuration) return unsupported("loading a duration value");
    return unsupported("loading nested tables, records or lists into cells");
  }));
  return { columns, rows };
}
