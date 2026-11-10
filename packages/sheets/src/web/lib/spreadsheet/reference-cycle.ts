/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// F4 reference cycling (Excel parity)
//
// While editing a formula, pressing F4 cycles the A1 reference under (or
// adjacent to) the caret through the four anchor states in Excel order:
//
//   A1  ->  $A$1  ->  A$1  ->  $A1  ->  A1
//
// This module is PURE: no React, no DOM. It reuses the A1 token shape from
// formula-engine.ts — column letters {1,3} + digits {1,7}, with optional `$`
// before each part. Crucially, a token that is immediately followed by `(` is
// a function call (e.g. LOG10) and must NOT be treated as a cell reference.
// ---------------------------------------------------------------------------

// Same A1 shape as formula-engine's CELL_REF_RE / SHIFT_REF_RE, but with each
// anchor part captured so we can re-emit any of the four states. Matches the
// whole token (anchored).
const REF_PARTS_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

// Characters that can appear inside a candidate A1 token while scanning a
// formula string. Mirrors formula-engine's tokenizer "word" rule (letters,
// digits, `$`, `_`, `.`), so we slice the same span the engine would, then
// validate it against REF_PARTS_RE (which rejects `_`/`.`-bearing idents).
const TOKEN_CHAR_RE = /[A-Za-z0-9$_.]/;

/**
 * Cycle a single A1 reference string through the next anchor state.
 *
 *   "A1"   -> "$A$1"
 *   "$A$1" -> "A$1"
 *   "A$1"  -> "$A1"
 *   "$A1"  -> "A1"
 *
 * Returns the input unchanged if it is not a valid A1 reference.
 */
export function cycleReferenceAnchors(ref: string): string {
  const m = REF_PARTS_RE.exec(ref);
  if (!m) return ref;
  const colAbs = Boolean(m[1]);
  const letters = m[2];
  const rowAbs = Boolean(m[3]);
  const digits = m[4];

  // Order: (rel,rel) -> (abs,abs) -> (rel-col,abs-row) -> (abs-col,rel-row).
  // i.e. A1 -> $A$1 -> A$1 -> $A1 -> A1.
  let nextColAbs: boolean;
  let nextRowAbs: boolean;
  if (!colAbs && !rowAbs) {
    nextColAbs = true;
    nextRowAbs = true; // A1 -> $A$1
  } else if (colAbs && rowAbs) {
    nextColAbs = false;
    nextRowAbs = true; // $A$1 -> A$1
  } else if (!colAbs && rowAbs) {
    nextColAbs = true;
    nextRowAbs = false; // A$1 -> $A1
  } else {
    nextColAbs = false;
    nextRowAbs = false; // $A1 -> A1
  }

  return `${nextColAbs ? "$" : ""}${letters}${nextRowAbs ? "$" : ""}${digits}`;
}

/** Result of cycling a reference under the caret. */
export interface CycleAtCaretResult {
  /** The full formula text with the targeted reference rewritten. */
  formula: string;
  /** New caret position, kept at the end of the rewritten reference. */
  caret: number;
}

interface TokenSpan {
  start: number;
  end: number; // exclusive
  text: string;
}

/**
 * Find the maximal token span (letters/digits/$/_/.) that contains the caret,
 * or — when the caret sits directly between two tokens — the token immediately
 * to the LEFT of the caret (Excel behaviour: F4 after typing a ref acts on
 * that ref). Returns null if no token touches the caret.
 */
function tokenSpanAtCaret(src: string, caret: number): TokenSpan | null {
  const n = src.length;
  if (n === 0) return null;
  const clamped = Math.max(0, Math.min(caret, n));

  const isTokenChar = (idx: number) =>
    idx >= 0 && idx < n && TOKEN_CHAR_RE.test(src[idx]);

  // Anchor index: a position inside the token. Prefer the char at the caret;
  // if that isn't a token char (caret at a boundary / end), look one to the
  // left so a caret right after a ref still targets that ref.
  let anchor = -1;
  if (isTokenChar(clamped)) {
    anchor = clamped;
  } else if (isTokenChar(clamped - 1)) {
    anchor = clamped - 1;
  }
  if (anchor < 0) return null;

  let start = anchor;
  while (start > 0 && isTokenChar(start - 1)) start--;
  let end = anchor + 1;
  while (end < n && isTokenChar(end)) end++;

  return { start, end, text: src.slice(start, end) };
}

/**
 * Whether `index` of a formula falls inside a double-quoted string literal
 * (a doubled "" inside the string is an escaped quote, not its end).
 */
function insideStringLiteral(formula: string, index: number): boolean {
  let inside = false;
  for (let i = 0; i < index; i += 1) {
    if (formula[i] !== '"') continue;
    if (inside && formula[i + 1] === '"') {
      i += 1;
      continue;
    }
    inside = !inside;
  }
  return inside;
}

/**
 * Cycle the A1 reference under/adjacent to the caret within a formula.
 *
 * - Only operates when `formula` starts with `=` (a real formula).
 * - The targeted token must be a valid A1 cell reference AND must NOT be
 *   immediately followed by `(` (so function names like LOG10 are skipped).
 * - Returns the rewritten formula plus the new caret (placed at the end of
 *   the rewritten reference), or `null` if nothing was cycled.
 */
export function cycleReferenceAtCaret(
  formula: string,
  caret: number,
): CycleAtCaretResult | null {
  const result = cycleReferenceInSelection(formula, caret, caret);
  return result ? { formula: result.formula, caret: result.selectionEnd } : null;
}

/** Result of cycling the reference a selection (or a caret) is on. */
export interface CycleInSelectionResult {
  formula: string;
  /** The rewritten reference's span, to keep it selected (equal, a caret at
   * its end, when the press came from a caret). */
  selectionStart: number;
  selectionEnd: number;
}

/** A valid A1 reference token at [start, end): not a function name. */
function referenceTokenAt(formula: string, start: number, end: number): TokenSpan | null {
  const text = formula.slice(start, end);
  if (!REF_PARTS_RE.test(text)) return null;
  let after = end;
  while (after < formula.length && /\s/.test(formula[after])) after++;
  if (formula[after] === "(") return null;
  return { start, end, text };
}

/** The token span that ends right before `index` (exclusive), if any. */
function tokenEndingAt(formula: string, index: number): TokenSpan | null {
  let start = index;
  while (start > 0 && TOKEN_CHAR_RE.test(formula[start - 1])) start--;
  return start < index ? referenceTokenAt(formula, start, index) : null;
}

/** The token span that starts at `index`, if any. */
function tokenStartingAt(formula: string, index: number): TokenSpan | null {
  let end = index;
  while (end < formula.length && TOKEN_CHAR_RE.test(formula[end])) end++;
  return end > index ? referenceTokenAt(formula, index, end) : null;
}

/**
 * F4 on what the formula's selection is on (a caret is an empty selection).
 *
 * A RANGE reference (C2:C3) is one reference: both ends take the next state
 * of its first end together, C2:C3 -> $C$2:$C$3 -> C$2:C$3 -> $C2:$C3 ->
 * C2:C3, whichever end the selection or caret is on -- "a cell reference or
 * range is selected in a formula" (Microsoft Support, "Keyboard shortcuts in
 * Excel", F4). Before, a selection over a whole C2:C3 locked C2 alone (live
 * audit S-135), and the copy down moved the unlocked end. A selection that
 * covers no reference (text in quotes, a function name) changes nothing.
 */
export function cycleReferenceInSelection(
  formula: string,
  selectionStart: number,
  selectionEnd: number,
): CycleInSelectionResult | null {
  if (!formula.startsWith("=")) return null;
  const from = Math.max(0, Math.min(selectionStart, selectionEnd));
  const to = Math.min(formula.length, Math.max(selectionStart, selectionEnd));
  // The reference the selection starts on (a caret: the one under or just
  // before it).
  const touched = tokenSpanAtCaret(formula, from);
  if (!touched) return null;
  // Text inside a string literal ("A1" as a SUMIF criterion) is not a
  // reference: F4 leaves it as typed.
  if (insideStringLiteral(formula, touched.start)) return null;
  const own = referenceTokenAt(formula, touched.start, touched.end);
  if (!own) return null;
  // Its range partner across a colon, either side.
  const before = formula[own.start - 1] === ":" ? tokenEndingAt(formula, own.start - 1) : null;
  const after = !before && formula[own.end] === ":" ? tokenStartingAt(formula, own.end + 1) : null;
  const first = before ?? own;
  const last = before ? own : after ?? own;
  const next = cycleReferenceAnchors(first.text);
  if (next === first.text) return null;
  const parts = REF_PARTS_RE.exec(next)!;
  const state = (ref: string) => {
    const m = REF_PARTS_RE.exec(ref)!;
    return `${parts[1]}${m[2]}${parts[3]}${m[4]}`;
  };
  const rewritten = first === last ? next : `${next}:${state(last.text)}`;
  const nextFormula = formula.slice(0, first.start) + rewritten + formula.slice(last.end);
  const end = first.start + rewritten.length;
  return {
    formula: nextFormula,
    selectionStart: from === to ? end : first.start,
    selectionEnd: end,
  };
}
