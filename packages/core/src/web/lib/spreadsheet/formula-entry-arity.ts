/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel refuses to enter a formula that calls a function with too few or too
 * many arguments: Enter shows "You've entered too few arguments for this
 * function." (or "too many") and the cell stays in Edit mode. The evaluator
 * gives such a call #VALUE! (formula-evaluator, case "call"); this is the
 * entry-time check, so a typed =AVERAGE() is not committed as #VALUE!.
 */
import { Parser, stripFutureFunctionPrefix, tokenize, type Node } from "./formula-engine";
import { functionArity, SELF_CHECKED_ARITY } from "./function-signatures";

export interface FormulaArityProblem {
  readonly name: string;
  readonly kind: "too-few" | "too-many";
}

function children(node: Node): Node[] {
  switch (node.kind) {
    case "call":
      return node.args;
    case "apply":
      return [node.callee, ...node.args];
    case "unary":
      return [node.operand];
    case "binary":
    case "rangeExpr":
      return [node.left, node.right];
    case "array":
      return node.rows.flat();
    default:
      return [];
  }
}

/** The first call in an entered formula whose argument count Excel refuses. */
export function formulaArityProblem(text: string): FormulaArityProblem | null {
  const trimmed = text.trim();
  if (!trimmed.startsWith("=") || trimmed.length < 2) return null;
  const tokens = tokenize(trimmed.slice(1));
  if (!tokens) return null;
  let root: Node;
  try {
    root = new Parser(tokens).parse();
  } catch {
    return null;
  }
  const pending: Node[] = [root];
  while (pending.length > 0) {
    const node = pending.pop()!;
    if (node.kind === "call") {
      const name = stripFutureFunctionPrefix(node.name.toUpperCase());
      if (!SELF_CHECKED_ARITY.has(name)) {
        const arity = functionArity(name);
        if (arity && node.args.length < arity.min) return { name, kind: "too-few" };
        if (arity && node.args.length > arity.max) return { name, kind: "too-many" };
      }
    }
    pending.push(...children(node));
  }
  return null;
}

/**
 * Whether an entered formula calls TABLE(). Only a What-If Data Table writes
 * that call (Data > What-If Analysis > Data Table); Excel refuses it typed
 * into a cell with "That function isn't valid." (measured: Excel 365 leaves
 * a cell given =TABLE(A1,B1) or =1+TABLE(,B1) empty; a "TABLE(" in a string
 * is text).
 */
export function formulaCallsDataTable(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed.startsWith("=") || !/table/i.test(trimmed)) return false;
  const tokens = tokenize(trimmed.slice(1));
  if (!tokens) return false;
  return tokens.some(
    (token, index) =>
      token.type === "ident" &&
      token.value.toUpperCase() === "TABLE" &&
      tokens[index + 1]?.type === "lparen",
  );
}
