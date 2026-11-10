/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { BARE_NAME_PART_CLASS } from "./name-syntax";

/** Excel's grid: the wrap a relative name's offsets take. */
export const NAME_GRID_ROWS = 1_048_576;
export const NAME_GRID_COLUMNS = 16_384;

const NAME_CHARACTER_CLASS = `${BARE_NAME_PART_CLASS}?\\\\`;

/** One A1 reference (a cell corner or a whole row/column line) in a name's text. */
interface NameRefToken {
  start: number;
  end: number;
  kind: "cell" | "columns" | "rows";
  /** Follows a sheet qualifier (`Sheet1!A1`, or the corner after `Sheet1!A1:`). */
  qualified: boolean;
  /** Whether a qualified first corner precedes this one across a `:`. */
  secondCorner: boolean;
  text: string;
}

const wrap = (value: number, size: number) => ((value % size) + size) % size;

function columnNumber(letters: string): number {
  let value = 0;
  for (const letter of letters.toUpperCase()) value = value * 26 + (letter.charCodeAt(0) - 64);
  return value - 1;
}

function columnLetters(index: number): string {
  let value = index + 1;
  let out = "";
  while (value > 0) {
    const rest = (value - 1) % 26;
    out = String.fromCharCode(65 + rest) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

/**
 * The references in a name's text, read as a formula reads them: text in
 * "..." is a string and a quoted sheet name is a name, neither a reference;
 * `LOG10(` is a function and `金额A1` one name.
 */
function nameRefTokens(text: string): NameRefToken[] {
  const blank = (literal: string) => " ".repeat(literal.length);
  const code = text
    .replace(/"(?:[^"]|"")*"/g, blank)
    .replace(/'(?:[^']|'')*'(?=!)/g, (literal) => `'${"x".repeat(Math.max(0, literal.length - 2))}'`);
  const pattern = new RegExp(
    `(^|[^${NAME_CHARACTER_CLASS}])(?:` +
      // a whole-column line (A:C), a whole-row line (2:5) ...
      `(\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3})(?![${NAME_CHARACTER_CLASS}(])` +
      `|(\\$?[0-9]{1,7}:\\$?[0-9]{1,7})(?![${NAME_CHARACTER_CLASS}(])` +
      // ... or one cell corner.
      `|(\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7})(?![${NAME_CHARACTER_CLASS}(])` +
      `)`,
    "gu",
  );
  const tokens: NameRefToken[] = [];
  for (const match of code.matchAll(pattern)) {
    const start = (match.index ?? 0) + match[1].length;
    const body = match[2] ?? match[3] ?? match[4] ?? "";
    const end = start + body.length;
    const before = code.slice(0, start);
    const previous = tokens[tokens.length - 1];
    const secondCorner =
      Boolean(match[4]) &&
      before.endsWith(":") &&
      previous?.kind === "cell" &&
      previous.end === start - 1;
    tokens.push({
      start,
      end,
      kind: match[2] ? "columns" : match[3] ? "rows" : "cell",
      qualified: before.endsWith("!") || (secondCorner && previous?.qualified === true),
      secondCorner,
      text: text.slice(start, end),
    });
  }
  return tokens;
}

function hasRelativeAxis(token: NameRefToken): boolean {
  if (token.kind === "cell") {
    const match = /^(\$?)[A-Za-z]{1,3}(\$?)[0-9]{1,7}$/.exec(token.text);
    return Boolean(match) && (match![1] !== "$" || match![2] !== "$");
  }
  return token.text.split(":").some((part) => !part.startsWith("$"));
}

/** The token with its relative axes moved by (rows, columns), wrapped. */
function shiftedToken(token: NameRefToken, rows: number, columns: number): string {
  if (token.kind === "cell") {
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$/.exec(token.text);
    if (!match) return token.text;
    const column = match[1] === "$" ? match[2].toUpperCase() : columnLetters(wrap(columnNumber(match[2]) + columns, NAME_GRID_COLUMNS));
    const row = match[3] === "$" ? match[4] : String(wrap(Number(match[4]) - 1 + rows, NAME_GRID_ROWS) + 1);
    return `${match[1]}${column}${match[3]}${row}`;
  }
  return token.text
    .split(":")
    .map((part) => {
      if (part.startsWith("$")) return part;
      return token.kind === "columns"
        ? columnLetters(wrap(columnNumber(part) + columns, NAME_GRID_COLUMNS))
        : String(wrap(Number(part) - 1 + rows, NAME_GRID_ROWS) + 1);
    })
    .join(":");
}

/**
 * Whether a stored name reads relative to the cell that uses it: one of its
 * sheet-qualified references has a relative axis.
 */
export function nameRefIsRelative(storedRef: string): boolean {
  return nameRefTokens(storedRef).some((token) => token.qualified && hasRelativeAxis(token));
}

/**
 * A stored name's text with every sheet-qualified relative axis moved by
 * (rows, columns), wrapped at the grid's edges: the stored A1-based offsets
 * read from a cell (`+cell`) for a formula that uses the name, or for the
 * active cell Name Manager shows them from.
 */
export function shiftNameRefRelativeAxes(storedRef: string, rows: number, columns: number): string {
  if (!rows && !columns) return storedRef;
  const tokens = nameRefTokens(storedRef).filter((token) => token.qualified && hasRelativeAxis(token));
  if (!tokens.length) return storedRef;
  let out = "";
  let cursor = 0;
  for (const token of tokens) {
    out += storedRef.slice(cursor, token.start) + shiftedToken(token, rows, columns);
    cursor = token.end;
  }
  return out + storedRef.slice(cursor);
}

/**
 * A stored name's reference as Name Manager's Refers to shows it with `active`
 * the active cell (relative axes read from it), or as stored without one.
 */
export function nameRefShownFrom(
  storedRef: string,
  active: { row: number; col: number } | null | undefined,
): string {
  return active ? shiftNameRefRelativeAxes(storedRef, active.row, active.col) : storedRef;
}

/** Whether the typed text holds a reference with a relative axis anywhere. */
export function typedNameRefHasRelativeAxis(typed: string): boolean {
  return nameRefTokens(typed).some(hasRelativeAxis);
}

/**
 * What Define Name / Name Manager store for a typed reference with the cell
 * `base` active: each unqualified reference takes `sheetPrefix` (`Sheet1!`,
 * `'My Sheet'!`; a range's second corner rides on its first), and each
 * relative axis becomes its offset from that cell, wrapped. Absolute axes and
 * already-qualified references keep their text.
 */
export function storedNameRefFromTyped(
  typed: string,
  sheetPrefix: string,
  base: { row: number; col: number },
): string {
  const tokens = nameRefTokens(typed);
  let out = "";
  let cursor = 0;
  for (const token of tokens) {
    const prefix = token.qualified || token.secondCorner ? "" : sheetPrefix;
    const body = hasRelativeAxis(token) ? shiftedToken(token, -base.row, -base.col) : token.text;
    out += typed.slice(cursor, token.start) + prefix + body;
    cursor = token.end;
  }
  return out + typed.slice(cursor);
}

/**
 * The references in a structural edit (insert / delete rows, columns or
 * cells) that must not move: a relative axis is an offset, which Excel leaves
 * as it is. Each such reference (with its sheet and its range's other corner)
 * is swapped for an inert placeholder before the edit's own adjuster runs and
 * put back after.
 */
export function protectRelativeNameRefs(
  storedRef: string,
  axis: "row" | "col" | "both",
): { text: string; restore: (adjusted: string) => string } | null {
  const tokens = nameRefTokens(storedRef);
  const relativeOnAxis = (token: NameRefToken) => {
    if (!token.qualified) return false;
    if (token.kind === "cell") {
      const match = /^(\$?)[A-Za-z]{1,3}(\$?)[0-9]{1,7}$/.exec(token.text);
      if (!match) return false;
      if (axis === "row") return match[2] !== "$";
      if (axis === "col") return match[1] !== "$";
      return match[1] !== "$" || match[2] !== "$";
    }
    if (axis !== "both" && (token.kind === "rows") !== (axis === "row")) return false;
    return hasRelativeAxis(token);
  };
  // One reference: a corner, or a range's two corners.
  const references: Array<{ first: NameRefToken; last: NameRefToken }> = [];
  for (const token of tokens) {
    const open = references[references.length - 1];
    if (token.secondCorner && open && open.last.end === token.start - 1) open.last = token;
    else references.push({ first: token, last: token });
  }
  const spans: Array<{ start: number; end: number }> = [];
  for (const reference of references) {
    const corners = reference.first === reference.last ? [reference.first] : [reference.first, reference.last];
    if (!corners.some(relativeOnAxis)) continue;
    // Its sheet qualifier goes with it.
    const before = storedRef.slice(0, reference.first.start);
    const sheet = /(?:'(?:[^']|'')*'|[^\s,;()=+\-*/^&<>!:"]+)(?::(?:'(?:[^']|'')*'|[^\s,;()=+\-*/^&<>!:"]+))?!$/.exec(before);
    spans.push({ start: reference.first.start - (sheet ? sheet[0].length : 0), end: reference.last.end });
  }
  if (!spans.length) return null;
  const originals: string[] = [];
  let text = "";
  let cursor = 0;
  for (const span of spans) {
    text += storedRef.slice(cursor, span.start) + `__MRAI_RELNAME_${originals.length}__`;
    originals.push(storedRef.slice(span.start, span.end));
    cursor = span.end;
  }
  text += storedRef.slice(cursor);
  return {
    text,
    restore: (adjusted) =>
      adjusted.replace(/__MRAI_RELNAME_(\d+)__/g, (whole, index: string) => originals[Number(index)] ?? whole),
  };
}
