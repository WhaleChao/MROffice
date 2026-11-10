/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Step-by-step formula evaluation ("Evaluate Formula" / 公式求值).
 *
 * Excel and WPS both let you walk a formula one sub-expression at a time: the
 * next thing that will be evaluated is underlined, and pressing Evaluate
 * replaces it with its result until only the answer is left. This module is the
 * PURE half of that: given the text of a formula it says WHICH span evaluates
 * next and how to substitute a result back in. The host supplies the actual
 * value (from the same evaluator the grid uses), so a step's result can never
 * disagree with what the cell shows.
 *
 * Which span comes next, in Excel's order:
 *
 *   - a reference that is not swallowed by a function call — `A1`, `Sheet2!B3`,
 *     `A1:B4`, a defined name — evaluated on its own, leftmost first;
 *   - the innermost function call, i.e. one whose body has no further `(`;
 *   - a bare group `( … )` once everything inside it is already a literal;
 *   - finally the whole remaining expression, when only literals and operators
 *     are left.
 *
 * Refs inside a bare group stay visible (Excel steps through `(A1+B1)*2` one ref
 * at a time); refs inside a function call do not (`SUM(A1:A3)` evaluates as one
 * step). Deviation from Excel is possible in exotic nesting, and the reported
 * value is always the engine's, so the walk is a faithful explanation rather
 * than a second implementation of evaluation.
 */

import {
  BARE_NAME_PART_CLASS,
  BARE_NAME_START_CLASS,
  DEFINED_NAME_START_CLASS,
} from "@mrnote/core/web/lib/spreadsheet/name-syntax.ts";

/** Longest formula this module will walk (matches the engine's own bound). */
export const MAX_EVALUATION_FORMULA_LENGTH = 8_192;
/** Hard stop so a pathological formula cannot loop the host. */
export const MAX_EVALUATION_STEPS = 256;

export type EvaluationTargetKind = "reference" | "call" | "group" | "whole";

export interface EvaluationTarget {
  /** Offset of the span's first character within the expression. */
  start: number;
  /** Offset just past the span's last character. */
  end: number;
  kind: EvaluationTargetKind;
  /** The span's text, for display and for the host to evaluate. */
  text: string;
}

interface Token {
  kind: "string" | "ref" | "name" | "number" | "punct" | "error" | "space";
  start: number;
  end: number;
  text: string;
}

// A bare sheet name and a defined name are spelled with the formula engine's
// own character set (./name-syntax): a letter of any script, `回款!C4`, `金额`.
const BARE_SHEET_NAME = `[${BARE_NAME_START_CLASS}][${BARE_NAME_PART_CLASS}]*`;
const REFERENCE_RE = new RegExp(
  `^(?:'[^']*'|${BARE_SHEET_NAME})!\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7}(?::\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7})?`,
  "u",
);
const LOCAL_REFERENCE_RE =
  /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?/;
// A defined name's later characters include the `?` and `\` Excel stores
// (./name-syntax isDefinedNameCharacter; xlsx-rounds R268-X15).
const NAME_RE = new RegExp(
  `^[${DEFINED_NAME_START_CLASS}][${BARE_NAME_PART_CLASS}?\\\\]*`,
  "u",
);
const QUALIFIED_CELL_RE = new RegExp(
  `^(?:'([^']*)'|(${BARE_SHEET_NAME}))!(\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7})$`,
  "u",
);
const NUMBER_RE = /^[0-9]+(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?/;
const ERROR_RE =
  /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/;

/** Split an expression into the few token classes the walk needs. */
function tokenize(expression: string): Token[] | null {
  const tokens: Token[] = [];
  let index = 0;
  while (index < expression.length) {
    const rest = expression.slice(index);
    const whitespace = /^\s+/.exec(rest);
    if (whitespace) {
      tokens.push({
        kind: "space",
        start: index,
        end: index + whitespace[0].length,
        text: whitespace[0],
      });
      index += whitespace[0].length;
      continue;
    }
    if (rest.startsWith('"')) {
      // Excel escapes a quote by doubling it.
      let cursor = 1;
      while (cursor < rest.length) {
        if (rest[cursor] === '"') {
          if (rest[cursor + 1] === '"') {
            cursor += 2;
            continue;
          }
          cursor += 1;
          break;
        }
        cursor += 1;
      }
      if (cursor > rest.length) return null;
      tokens.push({
        kind: "string",
        start: index,
        end: index + cursor,
        text: rest.slice(0, cursor),
      });
      index += cursor;
      continue;
    }
    const error = ERROR_RE.exec(rest);
    if (error) {
      tokens.push({
        kind: "error",
        start: index,
        end: index + error[0].length,
        text: error[0],
      });
      index += error[0].length;
      continue;
    }
    const qualified = REFERENCE_RE.exec(rest);
    if (qualified) {
      tokens.push({
        kind: "ref",
        start: index,
        end: index + qualified[0].length,
        text: qualified[0],
      });
      index += qualified[0].length;
      continue;
    }
    const number = NUMBER_RE.exec(rest);
    if (number) {
      tokens.push({
        kind: "number",
        start: index,
        end: index + number[0].length,
        text: number[0],
      });
      index += number[0].length;
      continue;
    }
    const name = NAME_RE.exec(rest);
    if (name) {
      // A name followed by `(` is a FUNCTION HEAD, never a reference. Without
      // this, `LOG10(` reads as column LOG row 10 — letters then digits is a
      // valid reference shape — and the walk would offer to evaluate the
      // function's own name as a cell.
      const after = rest.slice(name[0].length).trimStart();
      const local = after.startsWith("(")
        ? null
        : LOCAL_REFERENCE_RE.exec(rest);
      const isReference = local !== null && local[0].length === name[0].length;
      tokens.push({
        kind: isReference ? "ref" : "name",
        start: index,
        end: index + name[0].length,
        text: name[0],
      });
      index += name[0].length;
      continue;
    }
    tokens.push({
      kind: "punct",
      start: index,
      end: index + 1,
      text: rest[0],
    });
    index += 1;
  }
  return tokens;
}

interface OpenParen {
  /** Index of the `(` token. */
  tokenIndex: number;
  /** Offset of the call's name when this paren opens a function call. */
  callStart: number | null;
  /** True once another `(` has been seen before this one closes. */
  hasNestedParen: boolean;
  /** True while nothing but literals/operators has been seen inside. */
  literalOnly: boolean;
}

function significantBefore(tokens: Token[], index: number): Token | null {
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    if (tokens[cursor].kind !== "space") return tokens[cursor];
  }
  return null;
}

/**
 * The span that evaluates next, or null when the expression is a bare literal
 * (nothing left to do). Returns `kind: "whole"` when only literals and
 * operators remain — the host evaluates the whole thing in one last step.
 */
export function nextEvaluationTarget(
  expression: string,
): EvaluationTarget | null {
  const trimmed = expression.trim();
  if (!trimmed || trimmed.length > MAX_EVALUATION_FORMULA_LENGTH) return null;
  const tokens = tokenize(expression);
  if (!tokens) return null;

  const stack: OpenParen[] = [];
  let leftmostRef: Token | null = null;
  let innermostCall: EvaluationTarget | null = null;
  let literalGroup: EvaluationTarget | null = null;
  let sawOperator = false;

  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.kind === "space") continue;
    const insideCall = stack.some((frame) => frame.callStart !== null);

    if (token.kind === "punct" && token.text === "(") {
      const previous = significantBefore(tokens, index);
      const isCall = previous?.kind === "name";
      for (const frame of stack) {
        frame.hasNestedParen = true;
        frame.literalOnly = false;
      }
      stack.push({
        tokenIndex: index,
        callStart: isCall ? previous.start : null,
        hasNestedParen: false,
        literalOnly: true,
      });
      continue;
    }

    if (token.kind === "punct" && token.text === ")") {
      const frame = stack.pop();
      if (!frame) return null;
      if (!frame.hasNestedParen) {
        if (frame.callStart !== null && !innermostCall) {
          innermostCall = {
            start: frame.callStart,
            end: token.end,
            kind: "call",
            text: expression.slice(frame.callStart, token.end),
          };
        } else if (
          frame.callStart === null &&
          frame.literalOnly &&
          !literalGroup
        ) {
          literalGroup = {
            start: tokens[frame.tokenIndex].start,
            end: token.end,
            kind: "group",
            text: expression.slice(tokens[frame.tokenIndex].start, token.end),
          };
        }
      }
      continue;
    }

    if (token.kind === "ref") {
      for (const frame of stack) frame.literalOnly = false;
      if (!insideCall && !leftmostRef) leftmostRef = token;
      continue;
    }

    if (token.kind === "name") {
      const next = tokens
        .slice(index + 1)
        .find((candidate) => candidate.kind !== "space");
      const isCallHead = next?.kind === "punct" && next.text === "(";
      if (!isCallHead) {
        // A defined name behaves like a reference.
        for (const frame of stack) frame.literalOnly = false;
        if (!insideCall && !leftmostRef) leftmostRef = token;
      }
      continue;
    }

    if (token.kind === "punct" && token.text !== "," && token.text !== ")") {
      sawOperator = true;
    }
  }

  if (stack.length > 0) return null;

  const candidates = [
    leftmostRef
      ? {
          start: leftmostRef.start,
          end: leftmostRef.end,
          kind: "reference" as const,
          text: leftmostRef.text,
        }
      : null,
    innermostCall,
    literalGroup,
  ].filter((candidate): candidate is EvaluationTarget => candidate !== null);

  if (candidates.length > 0) {
    return candidates.reduce((best, candidate) =>
      candidate.start < best.start ? candidate : best,
    );
  }

  // Only literals and operators are left: one final step for the whole thing.
  if (sawOperator) {
    const start = expression.length - expression.trimStart().length;
    return {
      start,
      end: start + trimmed.length,
      kind: "whole",
      text: trimmed,
    };
  }
  return null;
}

/** Substitute an evaluated literal back into the expression. */
export function applyEvaluationStep(
  expression: string,
  target: EvaluationTarget,
  literal: string,
): string {
  if (
    target.start < 0 ||
    target.end > expression.length ||
    target.start >= target.end
  ) {
    return expression;
  }
  return (
    expression.slice(0, target.start) + literal + expression.slice(target.end)
  );
}

/**
 * Render an evaluated value the way the dialog shows it: text in quotes,
 * booleans upper-case, errors verbatim — i.e. re-parseable as a literal so the
 * next step evaluates the substituted expression unchanged.
 */
export function formatEvaluatedLiteral(value: unknown): string {
  if (value === null || value === undefined || value === "") return '""';
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  // The engine's error and array values are objects: an error reads as its
  // code (#NUM!, #DIV/0!), an array as Excel's array constant {1,2;3,4}. Both
  // used to read "[object Object]" (S-075).
  if (typeof value === "object") {
    const record = value as { __formulaError?: unknown; code?: unknown; __array?: unknown; grid?: unknown };
    if (record.__formulaError === true && typeof record.code === "string") return record.code;
    if (record.__array === true && Array.isArray(record.grid)) {
      return `{${(record.grid as unknown[][])
        .map((row) => (Array.isArray(row) ? row.map(formatEvaluatedLiteral).join(",") : ""))
        .join(";")}}`;
    }
    return "#VALUE!";
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? String(value) : "#NUM!";
  }
  const text = String(value);
  if (ERROR_RE.test(text) && ERROR_RE.exec(text)?.[0].length === text.length) {
    return text;
  }
  if (/^-?[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?$/.test(text)) return text;
  if (text === "TRUE" || text === "FALSE") return text;
  return `"${text.replaceAll('"', '""')}"`;
}

/** True when the span is a single-cell reference the user can step into. */
export function stepIntoTarget(
  target: EvaluationTarget,
): { sheet: string | null; ref: string } | null {
  if (target.kind !== "reference") return null;
  const text = target.text.trim();
  const qualified = QUALIFIED_CELL_RE.exec(text);
  if (qualified) {
    return { sheet: qualified[1] ?? qualified[2] ?? null, ref: qualified[3] };
  }
  if (/^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(text)) {
    return { sheet: null, ref: text };
  }
  return null;
}
