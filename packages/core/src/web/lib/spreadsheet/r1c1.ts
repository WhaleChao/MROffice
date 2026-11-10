/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// ---------------------------------------------------------------------------
// R1C1 reference style (Excel parity)
// ---------------------------------------------------------------------------
//
// A1 is the default cell-reference style: columns are letters (A, B, …, AA) and
// rows are 1-based numbers (A1, $B$3). R1C1 numbers BOTH axes and encodes a
// reference RELATIVE to the cell that contains the formula:
//
//   • R1C1                — ABSOLUTE row 1, column 1 (the A1 `$A$1`).
//   • R[2]C[-1]           — RELATIVE: two rows below, one column left of the
//                           formula cell.
//   • RC                  — the formula cell itself (relative row 0, col 0).
//   • R[1]C               — one row below, same column.
//   • R3C                 — ABSOLUTE row 3, RELATIVE same column.
//
// This module is a pure DISPLAY / INPUT layer: the persisted formula model is
// always A1 (so `referenceStyle` never touches stored content). `a1RefToR1C1`
// converts ONE address relative to an anchor cell; `formulaA1ToR1C1` /
// `formulaR1C1ToA1` rewrite every address inside a whole formula STRING while
// leaving string literals, function names, numbers and operators untouched.
//
// Everything here is a pure function (no DOM, no engine state) so it is unit
// testable in isolation.

import { columnToLetters, formulaWithoutStoragePrefixes, lettersToColumn } from "./formula-engine";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "./grid-sizing";

/** Workbook-level cell-reference display style. */
export type ReferenceStyle = "a1" | "r1c1";

/** Normalize an unknown persisted value to a valid {@link ReferenceStyle}. */
export function normalizeReferenceStyle(value: unknown): ReferenceStyle {
  return value === "r1c1" ? "r1c1" : "a1";
}

/** A parsed A1 address with per-axis absolute flags. row/col are 0-indexed. */
interface ParsedA1 {
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

// An A1 single-cell address with optional per-axis `$`. Anchored so it only
// matches a STANDALONE token (the scanner below already isolates the candidate).
const A1_ADDRESS_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/;

function parseA1Address(ref: string): ParsedA1 | null {
  const m = A1_ADDRESS_RE.exec(ref);
  if (!m) return null;
  const col = lettersToColumn(m[2]);
  const row = Number(m[4]) - 1;
  if (col < 0 || row < 0 || !Number.isFinite(row)) return null;
  if (col >= MAX_GRID_COLUMNS || row >= MAX_GRID_ROWS) return null;
  // m[1] is the `$` before the COLUMN letters; m[3] is the `$` before the ROW
  // digits — keep the axes matched to their own absolute flag.
  return { row, col, colAbs: m[1] === "$", rowAbs: m[3] === "$" };
}

// Render one R1C1 axis component. `letter` is "R" or "C". An ABSOLUTE component
// is the 1-based index with no brackets (R1 / C1). A RELATIVE component is the
// signed offset in brackets (R[2] / C[-1]); a zero relative offset collapses to
// the bare letter (R / C) — Excel's shorthand for "same row/column".
function r1c1Component(letter: "R" | "C", index: number, abs: boolean): string {
  if (abs) return `${letter}${index + 1}`;
  return index === 0 ? letter : `${letter}[${index}]`;
}

/**
 * Convert ONE A1 address to its R1C1 form relative to `anchor` (the cell that
 * owns the formula, 0-indexed). Absolute ($-prefixed) axes become absolute R1C1
 * components (R1 / C1); relative axes become signed offsets (R[n] / C[m], or the
 * bare R / C when the offset is 0). Returns the input unchanged if it is not a
 * parseable A1 address.
 */
export function a1RefToR1C1(
  ref: string,
  anchor: { row: number; col: number },
): string {
  const p = parseA1Address(ref);
  if (!p) return ref;
  const rowPart = r1c1Component(
    "R",
    p.rowAbs ? p.row : p.row - anchor.row,
    p.rowAbs,
  );
  const colPart = r1c1Component(
    "C",
    p.colAbs ? p.col : p.col - anchor.col,
    p.colAbs,
  );
  return `${rowPart}${colPart}`;
}

// R1C1 single-cell address: an optional R component then an optional C component
// (at least one must be present). Each is either absolute (R1 / C1), a bracketed
// signed relative offset (R[2] / C[-1]), or the bare letter (R / C = relative 0).
const R1C1_ADDRESS_RE =
  /^(?:R(\d+|\[-?\d+\])?)?(?:C(\d+|\[-?\d+\])?)?$/;

// Decode ONE R1C1 axis component capture into { index, abs } relative to
// `anchorIndex`, or null when the (absolute) index is out of grid bounds. A
// MISSING component (capture === undefined and the letter absent) is signalled
// by the caller; here `raw === undefined` with the letter PRESENT means the bare
// letter (relative 0).
function decodeR1C1Component(
  raw: string | undefined,
  anchorIndex: number,
  max: number,
): { index: number; abs: boolean } | null {
  if (raw === undefined) {
    // Bare "R" / "C": relative, zero offset → the anchor's own row/col.
    return { index: anchorIndex, abs: false };
  }
  if (raw.startsWith("[")) {
    const delta = Number(raw.slice(1, -1));
    const index = anchorIndex + delta;
    if (index < 0 || index >= max) return null;
    return { index, abs: false };
  }
  // Absolute 1-based.
  const index = Number(raw) - 1;
  if (index < 0 || index >= max) return null;
  return { index, abs: true };
}

/**
 * Convert ONE R1C1 address back to A1, relative to `anchor` (0-indexed). The
 * inverse of {@link a1RefToR1C1}: absolute components (R1 / C1) become $-prefixed
 * A1 axes; relative components resolve against the anchor. Returns the input
 * unchanged if it is not a parseable R1C1 address or resolves out of bounds.
 */
export function r1c1RefToA1(
  ref: string,
  anchor: { row: number; col: number },
): string {
  const trimmed = ref.trim();
  // Both components absent ("" or a stray) is not a real address.
  if (trimmed === "") return ref;
  const m = R1C1_ADDRESS_RE.exec(trimmed.toUpperCase());
  if (!m) return ref;
  const hasR = /^R/.test(trimmed.toUpperCase());
  const hasC = /C/.test(trimmed.toUpperCase().replace(/^R(\d+|\[-?\d+\])?/, ""));
  // A standalone "R" / "C" with NO partner and NO digits is ambiguous with a
  // named range / boolean; require at least one component that is unambiguously
  // an address (has the partner letter, a digit, or brackets).
  if (!hasR && !hasC) return ref;
  const row = hasR
    ? decodeR1C1Component(m[1], anchor.row, MAX_GRID_ROWS)
    : { index: anchor.row, abs: false };
  const col = hasC
    ? decodeR1C1Component(m[2], anchor.col, MAX_GRID_COLUMNS)
    : { index: anchor.col, abs: false };
  if (!row || !col) return ref;
  const colLetters = columnToLetters(col.index);
  return `${col.abs ? "$" : ""}${colLetters}${row.abs ? "$" : ""}${row.index + 1}`;
}

// A candidate A1 address inside a formula: optional $, 1–3 letters, optional $,
// 1–7 digits. The scanner only rewrites a match that is NOT immediately followed
// by "(" (a function call like LOG10(…)) and NOT part of a longer identifier.
const A1_SCAN_RE = /\$?[A-Za-z]{1,3}\$?[0-9]{1,7}/g;

// A candidate R1C1 address inside a formula. Matches the R/C grammar; the
// scanner re-validates each hit with `r1c1RefToA1` and skips non-addresses.
const R1C1_SCAN_RE = /R(?:\d+|\[-?\d+\])?C(?:\d+|\[-?\d+\])?|R\[-?\d+\]|C\[-?\d+\]/gi;

// True when the char can be part of an identifier (so a match flanked by one is
// inside a longer name, e.g. the "A1" in "NamedA1" or "A1B", not a real ref).
// Letters of every script count: in the unquoted sheet name `Données1!A1` the
// "es1" is part of the name, exactly as "eet1" is part of `Sheet1`.
function isIdentChar(ch: string | undefined): boolean {
  return ch !== undefined && /[\p{L}\p{M}\p{N}_.$]/u.test(ch);
}

// Walk a formula body, invoking `rewrite` on each ADDRESS-position token while
// copying string literals ("…"), and any other char, verbatim. `rewrite` returns
// the (possibly converted) text for the matched address; returning the input
// unchanged leaves it as-is. This preserves the formula's exact non-address text
// (whitespace, operators, function names, numbers) — only addresses change.
function scanFormulaAddresses(
  body: string,
  re: RegExp,
  rewrite: (address: string) => string,
): string {
  let out = "";
  let i = 0;
  const n = body.length;
  while (i < n) {
    const ch = body[i];
    // Copy a "…" string literal verbatim (with "" escapes) — never rewrite
    // inside it.
    if (ch === '"') {
      out += ch;
      i += 1;
      while (i < n) {
        out += body[i];
        if (body[i] === '"') {
          // Doubled quote = escaped quote, stay in the string.
          if (body[i + 1] === '"') {
            out += body[i + 1];
            i += 2;
            continue;
          }
          i += 1;
          break;
        }
        i += 1;
      }
      continue;
    }
    // A single quote only opens a quoted SHEET NAME ('Q1 A1'!B2): copy it
    // verbatim ('' escapes included), so an address-looking part of the name
    // is never rewritten.
    if (ch === "'") {
      out += ch;
      i += 1;
      while (i < n) {
        out += body[i];
        if (body[i] === "'") {
          if (body[i + 1] === "'") {
            out += body[i + 1];
            i += 2;
            continue;
          }
          i += 1;
          break;
        }
        i += 1;
      }
      continue;
    }
    re.lastIndex = i;
    const m = re.exec(body);
    if (m && m.index === i) {
      const token = m[0];
      const after = body[i + token.length];
      const before = body[i - 1];
      // Skip a function call (LOG10() — followed by "(") and a token that is
      // really a slice of a longer identifier / number (flanked by ident chars).
      if (after === "(" || isIdentChar(after) || isIdentChar(before)) {
        out += token;
        i += token.length;
        continue;
      }
      out += rewrite(token);
      i += token.length;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

/**
 * Rewrite every A1 address in a formula STRING to R1C1, relative to `anchor`
 * (the formula's own cell, 0-indexed). String literals, function names, numbers
 * and operators are preserved verbatim. A leading "=" is preserved. Pure
 * display transform — the stored formula stays A1.
 */
export function formulaA1ToR1C1(
  formula: string,
  anchor: { row: number; col: number },
): string {
  if (!formula) return formula;
  const hasEq = formula.startsWith("=");
  const body = hasEq ? formula.slice(1) : formula;
  const out = scanFormulaAddresses(body, A1_SCAN_RE, (addr) =>
    a1RefToR1C1(addr, anchor),
  );
  return `${hasEq ? "=" : ""}${out}`;
}

/**
 * Rewrite every R1C1 address in a formula STRING back to A1, relative to
 * `anchor` (0-indexed). The inverse of {@link formulaA1ToR1C1}, used to convert
 * R1C1 INPUT into the canonical A1 form before storage. Non-address tokens that
 * merely look R1C1-ish (a function name, a named range) are left untouched
 * because `r1c1RefToA1` returns them unchanged. A leading "=" is preserved.
 */
export function formulaR1C1ToA1(
  formula: string,
  anchor: { row: number; col: number },
): string {
  if (!formula) return formula;
  const hasEq = formula.startsWith("=");
  const body = hasEq ? formula.slice(1) : formula;
  const out = scanFormulaAddresses(body, R1C1_SCAN_RE, (addr) =>
    r1c1RefToA1(addr, anchor),
  );
  return `${hasEq ? "=" : ""}${out}`;
}

/** The grid COLUMN header label for a 0-indexed column under `style`. */
export function columnHeaderLabel(
  col: number,
  style: ReferenceStyle,
): string {
  return style === "r1c1" ? String(col + 1) : columnToLetters(col);
}

/**
 * The DISPLAY form of a cell's stored value under `style` (formula bar + in-cell
 * editor). In r1c1 mode a FORMULA (leading "=") is rewritten to R1C1 relative to
 * `anchor`; everything else — a1 mode, or a plain text/number value (no "=") — is
 * returned verbatim. The leading-"=" gate is essential: a plain text value like
 * "A1 budget" must never be treated as a reference.
 */
export function displayFormula(
  value: string,
  anchor: { row: number; col: number },
  style: ReferenceStyle,
): string {
  // A file's storage prefixes (_xlfn. / _xlws. / _xlpm.) are not shown
  // (X-169); the stored formula keeps them.
  const shown = value.startsWith("=") ? formulaWithoutStoragePrefixes(value) : value;
  return style === "r1c1" && shown.startsWith("=")
    ? formulaA1ToR1C1(shown, anchor)
    : shown;
}

/**
 * The canonical A1 form to STORE for a value typed under `style` — the inverse of
 * {@link displayFormula}. In r1c1 mode a typed FORMULA is converted back to A1
 * relative to `anchor`; a1 mode and non-formula values pass through unchanged, so
 * the persisted model is always A1.
 */
export function inputFormula(
  value: string,
  anchor: { row: number; col: number },
  style: ReferenceStyle,
): string {
  return style === "r1c1" && value.startsWith("=")
    ? formulaR1C1ToA1(value, anchor)
    : value;
}

/** The Name Box address of a 0-indexed cell under `style` (R3C2 / B3). */
export function cellAddressLabel(
  row: number,
  col: number,
  style: ReferenceStyle,
): string {
  return style === "r1c1" ? `R${row + 1}C${col + 1}` : `${columnToLetters(col)}${row + 1}`;
}

// A whole row (R3, R1:R3) or column (C2, C1:C4) in R1C1, absolute or relative.
const R1C1_ROWS_RE = /^R(\d+|\[-?\d+\])?(?::R(\d+|\[-?\d+\])?)?$/;
const R1C1_COLUMNS_RE = /^C(\d+|\[-?\d+\])?(?::C(\d+|\[-?\d+\])?)?$/;

/**
 * What the Name Box (or Go To) means by `text` typed in R1C1 style, in A1
 * form, or null when it is no R1C1 reference: R3C2, R1C1:R3C3, R[1]C, a whole
 * row R3 / R1:R3, a whole column C2 / C1:C4, each optionally after `Sheet!`.
 * Relative parts count from `anchor`, the active cell.
 */
export function r1c1NameBoxToA1(
  text: string,
  anchor: { row: number; col: number },
): string | null {
  const trimmed = text.trim();
  const bang = trimmed.lastIndexOf("!");
  const prefix = bang >= 0 ? trimmed.slice(0, bang + 1) : "";
  const body = (bang >= 0 ? trimmed.slice(bang + 1) : trimmed).toUpperCase();
  if (body === "") return null;
  const rows = R1C1_ROWS_RE.exec(body);
  if (rows) {
    const from = decodeR1C1Component(rows[1], anchor.row, MAX_GRID_ROWS);
    const to = body.includes(":") ? decodeR1C1Component(rows[2], anchor.row, MAX_GRID_ROWS) : from;
    return from && to ? `${prefix}${from.index + 1}:${to.index + 1}` : null;
  }
  const columns = R1C1_COLUMNS_RE.exec(body);
  if (columns) {
    const from = decodeR1C1Component(columns[1], anchor.col, MAX_GRID_COLUMNS);
    const to = body.includes(":") ? decodeR1C1Component(columns[2], anchor.col, MAX_GRID_COLUMNS) : from;
    return from && to ? `${prefix}${columnToLetters(from.index)}:${columnToLetters(to.index)}` : null;
  }
  const corners = body.split(":");
  if (corners.length > 2) return null;
  const converted: string[] = [];
  for (const corner of corners) {
    // A cell names both axes; a lone R or C part is a row or column, above.
    if (!/^R(\d+|\[-?\d+\])?C(\d+|\[-?\d+\])?$/.test(corner)) return null;
    const a1 = r1c1RefToA1(corner, anchor);
    if (a1 === corner) return null;
    converted.push(a1);
  }
  return `${prefix}${converted.join(":")}`;
}
