/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// Spreadsheet formula intellisense (Excel/WPS parity)
//
// PURE logic — no React, no DOM. Powers the cell editor's two formula aids:
//
//   1. Function-name autocomplete: while typing a function prefix after "=",
//      offer the matching engine functions; selecting one inserts "NAME(".
//   2. Argument hint: while the caret sits INSIDE a function call's argument
//      list, show that function's signature and which argument the caret is on.
//
// The set of functions is derived from the formula-engine FUNCTIONS registry
// (via getFunctionNames) so it can never drift from what the engine evaluates.
// Argument lists come from the shared function-signatures table, the same one
// the engine reads to decide how many arguments a call may carry.
// ---------------------------------------------------------------------------

import { getFunctionNames } from "@mrnote/core/web/lib/spreadsheet/formula-engine.ts";
import { functionSignatureSpec } from "@mrnote/core/web/lib/spreadsheet/function-signatures.ts";

/**
 * Every function name the engine supports, sorted (uppercase). Memoized: the
 * registry is static for the life of the module, so we snapshot once. Callers
 * that mutate the result get a fresh copy each call.
 */
let CACHED_NAMES: readonly string[] | null = null;
export function functionNames(): string[] {
  if (!CACHED_NAMES) CACHED_NAMES = Object.freeze(getFunctionNames());
  return [...CACHED_NAMES];
}

/** Reset the memoized name snapshot (test seam; not used in production). */
export function __resetFunctionNamesCache(): void {
  CACHED_NAMES = null;
}

/**
 * Functions the engine evaluates for compatibility but MRAI does not offer
 * for authoring, because Excel for the web does not: Microsoft Support states
 * for each FORECAST.ETS page that the function "is not available in Excel for
 * the web, iOS, or Android" (live audit B396). A workbook that already holds
 * such a formula still calculates, keeps its formula and gets its argument
 * hint ({@link signatureFor}); only the authoring surfaces — the Insert
 * Function catalog, the Formulas library menus and name autocomplete — leave
 * them out.
 */
const NOT_OFFERED_FOR_WEB_AUTHORING: ReadonlySet<string> = new Set([
  "FORECAST.ETS",
  "FORECAST.ETS.CONFINT",
  "FORECAST.ETS.SEASONALITY",
  "FORECAST.ETS.STAT",
]);

/** Whether a function is offered for writing a new formula (see above). */
export function isAuthoringFunction(name: string): boolean {
  return !NOT_OFFERED_FOR_WEB_AUTHORING.has(name.toUpperCase());
}

/**
 * The engine functions offered for AUTHORING, sorted: {@link functionNames}
 * minus those Excel for the web does not support. The Insert Function
 * catalog and autocomplete read this; evaluation and argument hints keep the
 * full registry.
 */
export function authoringFunctionNames(): string[] {
  return functionNames().filter(isAuthoringFunction);
}

/**
 * The signature of a function: its display form plus the ordered argument
 * names. `variadic` marks a function whose last argument repeats (Excel's
 * `…`), so the arg hint keeps highlighting the last slot past the final comma.
 */
export interface FunctionSignature {
  readonly name: string;
  /** Ordered argument labels, e.g. ["number1", "[number2]", "…"]. */
  readonly args: readonly string[];
  /** Whether the final argument repeats (SUM(number1, [number2], …)). */
  readonly variadic: boolean;
}

/**
 * Look up a function's signature. Engine functions return their curated arg
 * labels; a registered name without an entry (which the signature-coverage test
 * forbids) would still get a bounded generic signature so the hint names the
 * function; a name that is not a function at all returns null.
 */
export function signatureFor(name: string): FunctionSignature | null {
  const upper = name.toUpperCase();
  const sig = functionSignatureSpec(upper);
  if (sig) {
    return { name: upper, args: sig.args, variadic: sig.variadic };
  }
  // Only fall back for names the engine actually evaluates, so a typo never
  // produces a phantom hint.
  if (functionNames().includes(upper)) {
    return { name: upper, args: ["argument", "…"], variadic: true };
  }
  return null;
}

/**
 * Render a signature as a compact one-line string:
 *   "SUM(number1, number2, …)"
 * Marking the active argument is left to the UI (it has the argIndex).
 */
export function formatSignature(sig: FunctionSignature): string {
  return `${sig.name}(${sig.args.join(", ")})`;
}

/**
 * Match function names against a (case-insensitive) prefix, ranked so the most
 * useful completions come first:
 *   1. exact match (typed the whole name)
 *   2. prefix matches, shorter names first, then alphabetical
 * An empty prefix returns the full list (alphabetical). A prefix that matches
 * nothing returns []. Dotted names (STDEV.S) match on their literal text.
 */
export function matchFunctions(prefix: string, limit = 12): string[] {
  const all = authoringFunctionNames();
  const p = prefix.trim().toUpperCase();
  if (p === "") return all.slice(0, limit);
  const matches = all.filter((n) => n.startsWith(p));
  matches.sort((a, b) => {
    if (a === p && b !== p) return -1;
    if (b === p && a !== p) return 1;
    if (a.length !== b.length) return a.length - b.length;
    return a < b ? -1 : a > b ? 1 : 0;
  });
  return matches.slice(0, limit);
}

// ---------------------------------------------------------------------------
// Formula context: where is the caret?
// ---------------------------------------------------------------------------

export type FormulaContextKind = "name" | "args" | "none";

export interface FormulaContext {
  /**
   * "name" -> the caret is on a function-name token; offer name completions.
   * "args" -> the caret is inside a function call's argument list; show a hint.
   * "none" -> not in a formula, or in a spot where neither aid applies.
   */
  kind: FormulaContextKind;
  /** For "name": the identifier prefix under the caret (uppercase-insensitive,
   *  returned verbatim). For "args"/"none": "". */
  token: string;
  /** For "args": the enclosing function's name (uppercase). Otherwise omitted. */
  funcName?: string;
  /** For "args": the 0-based index of the argument the caret is on. */
  argIndex?: number;
  /**
   * For "name": the [start, end) range of the identifier token in `text` so the
   * caller can splice in the chosen completion. Otherwise omitted.
   */
  replaceStart?: number;
  replaceEnd?: number;
}

const NONE: FormulaContext = { kind: "none", token: "" };

const isIdentChar = (c: string): boolean =>
  (c >= "A" && c <= "Z") ||
  (c >= "a" && c <= "z") ||
  (c >= "0" && c <= "9") ||
  c === "_" ||
  c === ".";
const isIdentStart = (c: string): boolean =>
  (c >= "A" && c <= "Z") || (c >= "a" && c <= "z") || c === "_";

interface OpenCall {
  /** The function name that opened this paren (uppercase), or "" for a bare
   *  grouping paren like "=(1+2)". */
  name: string;
  /** Top-level comma count seen so far inside this paren -> current arg index. */
  argIndex: number;
}

/**
 * Decide what intellisense (if any) applies for `text` with the caret at
 * `caretPos`. The scanner walks character-by-character up to the caret,
 * tracking a stack of open function calls (so nesting works), skipping over
 * quoted strings (so a comma/paren inside a literal never moves the arg index
 * or opens a call), and counting top-level commas per open call.
 *
 * Rules (Excel-ish):
 *  - Not starting with "=" (after optional leading spaces) -> "none".
 *  - Caret inside a quoted string -> "none" (don't pop a completion list over
 *    a string literal).
 *  - Caret immediately after / within an identifier that is NOT yet followed by
 *    "(" -> "name" (offer completions for that prefix). The identifier may be a
 *    bare function name being typed at top level or as an argument.
 *  - Otherwise, if the caret is inside an open call's argument list -> "args"
 *    with that call's name + current arg index.
 *  - Inside a bare grouping paren with no enclosing call -> "none".
 */
export function parseFormulaContext(
  text: string,
  caretPos: number,
): FormulaContext {
  const caret = Math.max(0, Math.min(caretPos, text.length));

  // Must be a formula. Allow leading whitespace before "=".
  let firstNonSpace = 0;
  while (firstNonSpace < text.length && text[firstNonSpace] === " ")
    firstNonSpace++;
  if (text[firstNonSpace] !== "=") return NONE;
  // Caret before/at the "=" itself: nothing to complete.
  if (caret <= firstNonSpace) return NONE;

  const stack: OpenCall[] = [];
  let inString = false;
  // The function name awaiting its "(" — the most recent identifier that abuts
  // a "(". Local to this call, so no state ever leaks between invocations.
  let pendingName = "";
  let arrayDepth = 0;
  let i = firstNonSpace + 1;

  while (i < caret) {
    const c = text[i];
    if (inString) {
      if (c === '"') {
        // Doubled "" is an escaped quote and stays inside the string.
        if (text[i + 1] === '"') {
          i += 2;
          continue;
        }
        inString = false;
      }
      i++;
      continue;
    }
    if (c === '"') {
      inString = true;
      i++;
      continue;
    }
    if (isIdentStart(c)) {
      // Consume the whole identifier, then look at what follows.
      let j = i + 1;
      while (j < text.length && isIdentChar(text[j])) j++;
      // An identifier immediately followed by "(" is a function call. Record the
      // name; the "(" pushes a frame carrying it on the next loop turn.
      if (text[j] === "(") {
        pendingName = text.slice(i, j).toUpperCase();
      } else {
        // A bare name (operand) is not a call; clear any pending name.
        pendingName = "";
      }
      i = j;
      continue;
    }
    if (c === "(") {
      stack.push({ name: pendingName, argIndex: 0 });
      pendingName = "";
      i++;
      continue;
    }
    if (c === ")") {
      // Guard against an unbalanced closer (e.g. "=)A1") — match the comma
      // branch's stack.length check rather than relying on Array.pop()'s
      // silent-undefined behavior.
      if (stack.length > 0) stack.pop();
      pendingName = "";
      i++;
      continue;
    }
    // An array constant's commas and semicolons separate its elements, not
    // the call's arguments: =INDEX({10,20}, is still on INDEX's first
    // argument until the brace closes (live audit B442).
    if (c === "{") {
      arrayDepth += 1;
      pendingName = "";
      i++;
      continue;
    }
    if (c === "}") {
      if (arrayDepth > 0) arrayDepth -= 1;
      i++;
      continue;
    }
    if (c === ",") {
      if (arrayDepth === 0 && stack.length > 0) stack[stack.length - 1].argIndex += 1;
      i++;
      continue;
    }
    // Any other char (operators, refs, digits, spaces) clears a pending name.
    pendingName = "";
    i++;
  }

  // Caret inside an unterminated string literal -> suppress completions.
  if (inString) return NONE;

  // --- Name completion check --------------------------------------------
  // If the char immediately before the caret is part of an identifier, the user
  // is typing a name token; offer completions for its prefix UNLESS that token
  // is already a completed function call (followed by "(").
  if (caret > 0 && isIdentChar(text[caret - 1])) {
    // Find the identifier's start (scan back over ident chars).
    let start = caret - 1;
    while (start > 0 && isIdentChar(text[start - 1])) start--;
    // The token must START with an identifier-start char (not a pure number or
    // ".5"); a numeric literal is not a function name.
    if (isIdentStart(text[start])) {
      // Find the token's end (it may extend past the caret) and check the char
      // after it: if it's "(", the function is already applied — show its arg
      // hint instead of re-completing the name.
      let end = caret;
      while (end < text.length && isIdentChar(text[end])) end++;
      if (text[end] !== "(") {
        const token = text.slice(start, caret);
        return {
          kind: "name",
          token,
          replaceStart: start,
          replaceEnd: end,
        };
      }
    }
  }

  // --- Argument hint check ----------------------------------------------
  // The caret is inside the innermost open call (if that call has a name).
  for (let s = stack.length - 1; s >= 0; s--) {
    if (stack[s].name) {
      return {
        kind: "args",
        token: "",
        funcName: stack[s].name,
        argIndex: stack[s].argIndex,
      };
    }
  }

  return NONE;
}

/**
 * Convenience for the UI: given the current cell text + caret, splice the chosen
 * function name in as "NAME(" over the name token, returning the new text and
 * the caret position to place just after the "(". When the context isn't a
 * name-completion, the text is returned unchanged with the caret untouched.
 */
export function applyFunctionCompletion(
  text: string,
  caretPos: number,
  name: string,
): { text: string; caret: number } {
  const ctx = parseFormulaContext(text, caretPos);
  if (
    ctx.kind !== "name" ||
    ctx.replaceStart === undefined ||
    ctx.replaceEnd === undefined
  ) {
    return { text, caret: caretPos };
  }
  const upper = name.toUpperCase();
  const typed = text.slice(ctx.replaceStart, ctx.replaceEnd).toUpperCase();
  // =TRUE and =FALSE are complete formulas: Enter commits them, it does not
  // turn them into the open call "TRUE(" (live audit B170). A zero-argument
  // function is completed as a whole call, caret after it.
  if ((upper === "TRUE" || upper === "FALSE") && typed === upper) {
    return { text, caret: caretPos };
  }
  const insert = ZERO_ARGUMENT_FUNCTIONS.has(upper) ? `${upper}()` : `${upper}(`;
  const next =
    text.slice(0, ctx.replaceStart) + insert + text.slice(ctx.replaceEnd);
  return { text: next, caret: ctx.replaceStart + insert.length };
}

/** Functions Excel completes as NAME() because they take no argument. */
const ZERO_ARGUMENT_FUNCTIONS = new Set([
  "TRUE", "FALSE", "PI", "NOW", "TODAY", "RAND", "NA", "SHEETS",
]);

/** True when the token under the caret already is a complete formula that
 *  Enter should commit rather than complete (a boolean literal). */
export function functionCompletionLeavesTextUnchanged(
  text: string,
  caretPos: number,
  name: string,
): boolean {
  return applyFunctionCompletion(text, caretPos, name).text === text;
}
