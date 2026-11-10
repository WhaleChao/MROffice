/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Computing a formula's numeric value from its LaTeX.
 *
 * A recognised formula carries LaTeX — that is what recognition produces and
 * the editor round-trips. For the arithmetic subset (fractions, powers,
 * roots, trig, logs, constants) the value is derivable right here, no CAS
 * service required, and showing it next to the source turns the formula panel
 * into a calculator for exactly the formulas a textbook page contains.
 *
 * Everything outside the subset returns a REASON, not a guess: symbolic
 * variables, sums, integrals and matrices are real mathematics this module
 * does not do, and a wrong number under a formula is worse than no number.
 * An equation like `E = mc^2` is not computable; `x = 3 + 4` computes its
 * right-hand side, because "what does this come to" is the question a reader
 * is asking of it.
 */
import { matchAllBehind } from "@mrnote/core/web/lib/regexp-behind.ts";

export interface FormulaValue {
  ok: true;
  value: number;
  /** Trimmed for display: up to 10 significant digits, no trailing zeros. */
  display: string;
}

export interface FormulaNotComputable {
  ok: false;
  /** Why not, as a stable code the UI can translate. */
  reason:
    | "empty"
    | "symbolic"
    | "unsupported"
    | "malformed"
    | "not_finite";
}

export type FormulaComputation = FormulaValue | FormulaNotComputable;

const CONSTANTS: Record<string, number> = {
  "\\pi": Math.PI,
  e: Math.E,
};

const FUNCTIONS: Record<string, (x: number) => number> = {
  "\\sin": Math.sin,
  "\\cos": Math.cos,
  "\\tan": Math.tan,
  "\\arcsin": Math.asin,
  "\\arccos": Math.acos,
  "\\arctan": Math.atan,
  "\\ln": Math.log,
  "\\log": Math.log10,
  "\\exp": Math.exp,
  "\\sinh": Math.sinh,
  "\\cosh": Math.cosh,
  "\\tanh": Math.tanh,
};

/** Commands that read as noise for arithmetic and are safe to drop. */
const IGNORED = ["\\mathrm"];

/** Commands that mark real mathematics beyond this module. */
const SYMBOLIC_MARKERS = [
  "\\sum",
  "\\int",
  "\\prod",
  "\\lim",
  "\\infty",
  "\\partial",
  "\\nabla",
  "\\begin",
  "\\vec",
  "\\hat",
];

class Parser {
  private text: string;
  private at = 0;
  private variables: Readonly<Record<string, number>>;

  constructor(text: string, variables: Readonly<Record<string, number>> = {}) {
    this.text = text;
    this.variables = variables;
  }

  private skipSpace(): void {
    while (this.at < this.text.length && /\s/.test(this.text[this.at])) {
      this.at += 1;
    }
  }

  peek(): string {
    this.skipSpace();
    return this.text[this.at] ?? "";
  }

  /** The command starting at the cursor, e.g. `\frac`, or null. */
  peekCommand(): string | null {
    this.skipSpace();
    if (this.text[this.at] !== "\\") return null;
    const match = /^\\[a-zA-Z]+|^\\./.exec(this.text.slice(this.at));
    return match ? match[0] : null;
  }

  consume(literal: string): boolean {
    this.skipSpace();
    if (this.text.startsWith(literal, this.at)) {
      this.at += literal.length;
      return true;
    }
    return false;
  }

  done(): boolean {
    this.skipSpace();
    return this.at >= this.text.length;
  }

  fail(reason: FormulaNotComputable["reason"]): never {
    throw reason;
  }

  /** expression := term (('+'|'-') term)* */
  expression(): number {
    let value = this.term();
    for (;;) {
      if (this.consume("+")) value += this.term();
      else if (this.consume("-")) value -= this.term();
      else return value;
    }
  }

  /** term := factor (('*'|explicit division|adjacency) factor)* */
  term(): number {
    let value = this.factor();
    for (;;) {
      if (
        this.consume("\\cdot") ||
        this.consume("\\times") ||
        this.consume("*")
      ) {
        value *= this.factor();
        continue;
      }
      if (this.consume("\\div") || this.consume("/")) {
        value /= this.factor();
        continue;
      }
      // Adjacency is multiplication: `2\pi`, `3(4+1)`, `\frac{1}{2}\pi`.
      const next = this.peek();
      const command = this.peekCommand();
      if (
        next === "(" ||
        next === "{" ||
        /[0-9.]/.test(next) ||
        (command !== null &&
          (command === "\\frac" ||
            command === "\\sqrt" ||
            command === "\\pi" ||
            command in FUNCTIONS))
      ) {
        value *= this.factor();
        continue;
      }
      return value;
    }
  }

  /** factor := ('-'|'+')* power */
  factor(): number {
    if (this.consume("-")) return -this.factor();
    if (this.consume("+")) return this.factor();
    return this.power();
  }

  /** power := atom ('^' exponent)* — right-associative, as maths reads it. */
  power(): number {
    const base = this.atom();
    if (this.consume("^")) {
      const exponent = this.consume("{")
        ? this.braced()
        : // A bare exponent is one token: x^23 means x^2 * 3 in LaTeX.
          this.singleTokenNumber();
      return base ** exponent;
    }
    return base;
  }

  private braced(): number {
    const value = this.expression();
    if (!this.consume("}")) this.fail("malformed");
    return value;
  }

  private singleTokenNumber(): number {
    this.skipSpace();
    const char = this.text[this.at] ?? "";
    if (/[0-9]/.test(char)) {
      this.at += 1;
      return Number(char);
    }
    if (this.peekCommand() === "\\pi") {
      this.consume("\\pi");
      return Math.PI;
    }
    if (this.consume("(")) {
      const value = this.expression();
      if (!this.consume(")")) this.fail("malformed");
      return value;
    }
    this.fail("unsupported");
  }

  atom(): number {
    for (;;) {
      const dropped = IGNORED.find((cmd) => this.consume(cmd));
      if (!dropped) break;
      // \mathrm{...} etc. wrap their argument; unwrap the brace group.
      if (dropped === "\\mathrm" && this.consume("{")) {
        const value = this.expression();
        if (!this.consume("}")) this.fail("malformed");
        return value;
      }
    }

    const command = this.peekCommand();
    if (command) {
      if (SYMBOLIC_MARKERS.includes(command)) this.fail("symbolic");
      if (command === "\\frac" || command === "\\dfrac" || command === "\\tfrac") {
        this.consume(command);
        if (!this.consume("{")) this.fail("malformed");
        const numerator = this.braced();
        if (!this.consume("{")) this.fail("malformed");
        return numerator / this.braced();
      }
      if (command === "\\sqrt") {
        this.consume(command);
        let degree = 2;
        if (this.consume("[")) {
          degree = this.expression();
          if (!this.consume("]")) this.fail("malformed");
        }
        if (!this.consume("{")) this.fail("malformed");
        const radicand = this.braced();
        return radicand ** (1 / degree);
      }
      if (command === "\\pi") {
        this.consume(command);
        return CONSTANTS["\\pi"];
      }
      if (command in FUNCTIONS) {
        this.consume(command);
        // Optional power on the function name: \sin^2 x = (\sin x)^2.
        let functionPower = 1;
        if (this.consume("^")) {
          functionPower = this.consume("{")
            ? this.braced()
            : this.singleTokenNumber();
        }
        const argument = this.functionArgument();
        return FUNCTIONS[command](argument) ** functionPower;
      }
      this.fail("unsupported");
    }

    const char = this.peek();
    if (char === "(") {
      this.consume("(");
      const value = this.expression();
      if (!this.consume(")")) this.fail("malformed");
      return value;
    }
    if (char === "{") {
      this.consume("{");
      return this.braced();
    }
    if (/[0-9.]/.test(char)) {
      const match = /^[0-9]*\.?[0-9]+/.exec(this.text.slice(this.at));
      if (!match) this.fail("malformed");
      this.at += match[0].length;
      return Number(match[0]);
    }
    if (/[a-zA-Z]/.test(char)) {
      const following = this.text[this.at + 1] ?? "";
      const standsAlone = !/[a-zA-Z]/.test(following);
      // A bound variable resolves; `x` in a plot is a number-to-be.
      if (standsAlone && char in this.variables) {
        this.at += 1;
        return this.variables[char];
      }
      if (char === "e" && standsAlone) {
        this.at += 1;
        return CONSTANTS.e;
      }
      this.fail("symbolic");
    }
    this.fail("malformed");
  }

  /** A function's argument: a group, a parenthesised expression, or one factor. */
  private functionArgument(): number {
    if (this.consume("(")) {
      const value = this.expression();
      if (!this.consume(")")) this.fail("malformed");
      return value;
    }
    if (this.consume("{")) return this.braced();
    return this.factor();
  }
}

function trimmedDisplay(value: number): string {
  const fixed = Number(value.toPrecision(10));
  return String(fixed);
}

/**
 * Evaluate `latex` numerically, or say why not.
 *
 * `lhs = rhs` computes the right-hand side when the left is a bare symbol —
 * "x = 3 + 4" is asking what x comes to. Any other `=` is a statement, not a
 * question, and is refused as symbolic.
 */
export function computeFormulaValue(
  latex: string,
  options: { variables?: Readonly<Record<string, number>> } = {},
): FormulaComputation {
  const source = (latex ?? "").trim();
  if (!source) return { ok: false, reason: "empty" };

  let body = source;
  const equals = source.indexOf("=");
  if (equals >= 0) {
    const lhs = source.slice(0, equals).trim();
    // One letter (Greek included), at most a simple subscript. Anything wider
    // is an expression, and `\sum_{i=1}` would otherwise split at ITS equals
    // and read as a definition of `\sum_{i`.
    if (/^(\\?[a-zA-Z]+|[a-zA-Z](_\{?[a-zA-Z0-9]+\}?)?)$/.test(lhs) &&
        !/^\\(sum|int|prod|lim)/.test(lhs)) {
      body = source.slice(equals + 1).trim();
      if (body.includes("=")) return { ok: false, reason: "symbolic" };
    } else {
      return { ok: false, reason: "symbolic" };
    }
  }

  // Layout commands are pure typography and can sit anywhere — including
  // between an expression and its closing bracket, where a token-level skip
  // cannot reach them. Stripped before parsing.
  const stripped = body.replace(
    /\\(left|right|,|;|:|!|quad|qquad|displaystyle)(?![a-zA-Z])/g,
    "",
  );

  try {
    const parser = new Parser(stripped, options.variables ?? {});
    const value = parser.expression();
    if (!parser.done()) {
      // `2x` parses the 2 and leaves the x: the leftover being a letter or a
      // command means symbols, which is a different honesty than "syntax this
      // module does not know".
      return {
        ok: false,
        reason: /[a-zA-Z\\]/.test(parser.peek()) ? "symbolic" : "unsupported",
      };
    }
    if (!Number.isFinite(value)) return { ok: false, reason: "not_finite" };
    return { ok: true, value, display: trimmedDisplay(value) };
  } catch (reason) {
    if (
      reason === "symbolic" ||
      reason === "unsupported" ||
      reason === "malformed" ||
      reason === "not_finite"
    ) {
      return { ok: false, reason };
    }
    return { ok: false, reason: "malformed" };
  }
}

/**
 * The single free variable of `latex`, or null.
 *
 * Standalone letters, commands stripped first so `\sin`'s letters do not
 * count; `e` is the constant, never a variable. More than one distinct letter
 * means a surface, not a curve, and this module plots curves.
 */
export function freeVariableOf(latex: string): string | null {
  const withoutLhs = (() => {
    const equals = latex.indexOf("=");
    if (equals < 0) return latex;
    return latex.slice(equals + 1);
  })();
  const stripped = withoutLhs.replace(/\\[a-zA-Z]+/g, " ");
  const letters = new Set<string>();
  for (const match of matchAllBehind(stripped, /([a-zA-Z])(?![a-zA-Z])/g, { char: /[a-zA-Z]/, not: true })) {
    if (match[1] !== "e") letters.add(match[1]);
  }
  if (letters.size !== 1) return null;
  return [...letters][0];
}

export interface CompiledFormula {
  variable: string;
  evaluate: (value: number) => number | null;
}

/**
 * Compile a one-variable formula into a callable, or say no.
 *
 * Probed at a handful of points before being handed out: a formula that
 * fails everywhere is not a function worth plotting, and discovering that in
 * the middle of a render loop is the wrong place.
 */
export function compileFormulaFunction(latex: string): CompiledFormula | null {
  const variable = freeVariableOf(latex);
  if (!variable) return null;
  const evaluate = (value: number): number | null => {
    const result = computeFormulaValue(latex, {
      variables: { [variable]: value },
    });
    return result.ok ? result.value : null;
  };
  const probes = [-2, -0.5, 0.5, 1, 2];
  if (!probes.some((probe) => evaluate(probe) !== null)) return null;
  return { variable, evaluate };
}

export interface PlotSegment {
  points: Array<{ x: number; y: number }>;
}

/**
 * Sample a compiled formula for plotting.
 *
 * Non-finite and failed points BREAK the segment rather than being skipped
 * silently: connecting across a pole draws the vertical asymptote line that
 * every naive plotter draws through tan(x).
 */
export function samplePlotSegments(
  compiled: CompiledFormula,
  options: { from?: number; to?: number; samples?: number } = {},
): PlotSegment[] {
  const from = options.from ?? -5;
  const to = options.to ?? 5;
  const samples = Math.max(16, Math.min(2000, options.samples ?? 240));
  if (!(to > from)) return [];
  const segments: PlotSegment[] = [];
  let current: Array<{ x: number; y: number }> = [];
  for (let i = 0; i <= samples; i += 1) {
    const x = from + ((to - from) * i) / samples;
    const y = compiled.evaluate(x);
    if (y === null || !Number.isFinite(y)) {
      if (current.length > 1) segments.push({ points: current });
      current = [];
      continue;
    }
    current.push({ x, y });
  }
  if (current.length > 1) segments.push({ points: current });

  // Second pass: an asymptote the sampling STRADDLES never yields a
  // non-finite point — 1/x at 41 samples over [-2, 2] is finite everywhere
  // sampled and swings from -20 to +20 across one step. A sign-changing jump
  // bigger than half the whole y-range is that signature, and connecting it
  // draws the vertical line every naive plotter puts through tan(x).
  let minY = Infinity;
  let maxY = -Infinity;
  for (const segment of segments) {
    for (const point of segment.points) {
      if (point.y < minY) minY = point.y;
      if (point.y > maxY) maxY = point.y;
    }
  }
  const range = maxY - minY;
  if (!(range > 0)) return segments;
  const split: PlotSegment[] = [];
  for (const segment of segments) {
    let run: Array<{ x: number; y: number }> = [];
    for (const point of segment.points) {
      const previous = run[run.length - 1];
      if (
        previous &&
        Math.abs(point.y - previous.y) > range / 2 &&
        Math.sign(point.y) !== Math.sign(previous.y)
      ) {
        if (run.length > 1) split.push({ points: run });
        run = [];
      }
      run.push(point);
    }
    if (run.length > 1) split.push({ points: run });
  }
  return split;
}
