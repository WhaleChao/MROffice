/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Manual in-cell line breaks (Excel / WPS "Alt+Enter") — pure text helpers.
 *
 * Excel lets a cell hold MANUAL line breaks: pressing Alt+Enter while editing
 * inserts a newline INSIDE the cell, and the cell then renders on multiple
 * lines. The stored value carries a literal "\n" between the lines; the formula
 * engine treats the whole thing as one text value (a leading "=" still makes it
 * a formula, and a multi-line value never parses as a number).
 *
 * The grid cell is a single-line `<input>`, which cannot DISPLAY a newline, so
 * the renderer needs to know (a) whether a non-active cell should switch to a
 * wrapped multi-line OVERLAY and (b) how to split the value into display lines.
 * Editing itself stays in the `<input>` (the value still round-trips the "\n");
 * the caller inserts the break at the caret with {@link insertLineBreakAt}.
 *
 * Everything here is framework-free, total, and deterministic so it is unit
 * testable without a grid. A value with no manual break is reported as such, so
 * the single-line default render path is unchanged byte-for-byte.
 */

/** Max manual line breaks honored in one cell (DoS bound on the overlay split). */
export const MAX_CELL_LINES = 250;

/**
 * Normalize newline conventions to "\n": CRLF ("\r\n") and lone CR ("\r") both
 * collapse to a single "\n" so a value pasted/imported from Windows or old-Mac
 * text reads as the same manual breaks the editor produces. A value without any
 * carriage return / line feed is returned unchanged (and the SAME reference is
 * not guaranteed — callers treat it as a value).
 */
export function normalizeNewlines(value: string): string {
  if (value.indexOf("\r") === -1) return value;
  return value.replace(/\r\n?/g, "\n");
}

/**
 * True when `value` contains at least one manual line break: a line feed
 * ("\n", alone or after a CR). A carriage return on its own is no break --
 * Excel draws it as nothing, also with Wrap Text on (RO050_x000D: 银手镯_x000D_用户C
 * runs on in one paragraph, X-247). A formula value (leading "=") is still
 * reported by its raw text — the renderer only consults this for NON-formula
 * display, and a formula result is shown by the engine, never split here.
 */
export function hasManualLineBreak(value: string | undefined): boolean {
  if (typeof value !== "string" || value === "") return false;
  return value.indexOf("\n") !== -1;
}

/** `value` as its lines are drawn: CRLF is one break, a lone CR nothing
 * (X-247). */
function displayNewlines(value: string): string {
  if (value.indexOf("\r") === -1) return value;
  return value.replace(/\r\n/g, "\n").replace(/\r/g, "");
}

/**
 * The text a cell draws when its Wrap Text is explicitly off: Excel draws its
 * line breaks (CR, LF, CRLF) as nothing, on one line ("ABC\r\nDEF" shows
 * ABCDEF, MEASURED in Excel for Mac 16). Any other cell keeps its text.
 */
/**
 * A cell's Wrap Text as its text is drawn. A formula's result draws its line
 * breaks only with Wrap Text on, as in Excel (S26_tdf169326: =A1 over
 * "Hello\nWorld" shows HelloWorld on one line, X-148); a cell stating nothing
 * otherwise keeps the editor's older reading of its own Alt+Enter entries,
 * which drew every break.
 */
export function cellWrapTextSetting(
  cell: { value?: unknown; formula?: unknown; style?: { wrapText?: boolean } } | null | undefined,
): boolean | undefined {
  const stated = cell?.style?.wrapText;
  if (stated !== undefined) return stated;
  const formula =
    (typeof cell?.formula === "string" && cell.formula.startsWith("=")) ||
    (typeof cell?.value === "string" && cell.value.startsWith("=") && cell.value.length > 1);
  return formula ? false : undefined;
}








export function autoRowMeasuredText(
  text: string,
  cell: { value?: unknown; formula?: unknown; style?: { wrapText?: boolean } } | null | undefined,
): { text: string; singleLine: boolean } {
  if (cell?.style?.wrapText || cellWrapTextSetting(cell) !== false) return { text, singleLine: false };
  return { text: cellTextOnOneLine(text, false), singleLine: true };
}

/** The characters a cell shows on one line: Excel draws the first 1,024 of a
 * longer text that does not wrap (R250-X06's J571, a 6,407-character
 * TEXTJOIN, prints its first 1,024 across 11 pages); wrapped text shows whole. */
export const EXCEL_SHOWN_LINE_CHARACTERS = 1024;

/** A text that does not wrap, cut to the characters Excel shows of it. */
export function cellLineShown(text: string, wrapText: boolean | undefined): string {
  return wrapText === true || text.length <= EXCEL_SHOWN_LINE_CHARACTERS ? text : text.slice(0, EXCEL_SHOWN_LINE_CHARACTERS);
}

/** The faces' missing glyph (.notdef) advances, em: GDI's default character. */
const EXCEL_NOTDEF_EM: Readonly<Record<string, number>> = {
  arial: 0.75, "arial nova": 0.75, calibri: 0.5068359375, "times new roman": 0.77783203125, "courier new": 0.60009765625,
};

/**
 * How far the line breaks of a one-line text (Wrap Text stated off) reach
 * past its drawn characters, em: Windows Excel draws them as nothing but
 * measures each as the face's missing glyph when it lays out the print range,
 * the ones the text ends with too (lookwin p236: R462-X23's Info A3, bold
 * Arial 12 ending in four breaks 14 pt short of D, prints a band of D with
 * one trailing break and not with none). A face not listed reaches as
 * Arial's. The API's writer measures the face's own (sheet_extent).
 */
export function excelOneLineBreaksReachEm(text: string, family: string | undefined): number {
  let breaks = 0;
  for (let index = text.indexOf("\n"); index >= 0; index = text.indexOf("\n", index + 1)) breaks += 1;
  if (!breaks) return 0;
  return breaks * (EXCEL_NOTDEF_EM[(family ?? "").trim().toLowerCase()] ?? EXCEL_NOTDEF_EM.arial ?? 0.75);
}

export function cellTextOnOneLine<T>(text: T, wrapText: boolean | undefined): T {
  if (wrapText !== false || typeof text !== "string" || !/[\r\n]/.test(text)) return text;
  return text.replace(/[\r\n]/g, "") as T;
}

/**
 * How one line of cell text wider than its cell is anchored where it cannot
 * spill: centred text stays centred and is cut at both edges (S20's
 * "IGITEM_NU" of CONFIGITEM_NUMBER, X-154); right-aligned text keeps its end
 * at the right edge and is cut at the left (CR096_tPivot5_Static's
 * "ItemsTotal}}", X-239). Left-aligned text overflows to the right only, as
 * an inline line does (no style). Right-to-left text keeps the inline flow.
 */
export function clippedCellLineAnchorCss(
  textAlign: unknown,
  dir: string | undefined,
): { display: "flex"; justifyContent: "center" | "flex-end" } | Record<string, never> {
  if (textAlign === "center") return { display: "flex", justifyContent: "center" };
  if (textAlign === "right" && dir !== "rtl") return { display: "flex", justifyContent: "flex-end" };
  return {};
}

/**
 * Split a cell value into its manual display lines (Excel multi-line render).
 * Newlines are normalized first, then the value is split on "\n". A value with
 * no break yields a single-element array (the value itself). The line count is
 * capped at {@link MAX_CELL_LINES}; any overflow lines are joined back onto the
 * last kept line so no text is lost (it just stops adding new visual rows).
 * An empty string yields `[""]` (one empty line) — callers that want the
 * single-line path should gate on {@link hasManualLineBreak} first.
 */
export function splitCellLines(value: string): string[] {
  const normalized = displayNewlines(value);
  if (normalized.indexOf("\n") === -1) return [normalized];
  const parts = normalized.split("\n");
  if (parts.length <= MAX_CELL_LINES) return parts;
  const head = parts.slice(0, MAX_CELL_LINES - 1);
  const tail = parts.slice(MAX_CELL_LINES - 1).join("\n");
  head.push(tail);
  return head;
}

/**
 * Insert a manual line break into `value` at the caret, replacing any selected
 * range [selStart, selEnd). Returns the new value AND the caret offset just
 * after the inserted "\n" so the editor can restore the caret. Out-of-range /
 * inverted selection bounds are clamped into [0, value.length] and ordered, so a
 * caller passing a stale or null-ish selection still produces a sane result
 * (the break lands at the clamped position). Mirrors what a textarea would do on
 * Alt+Enter, but works against the single-line `<input>` value.
 */
export function insertLineBreakAt(
  value: string,
  selStart: number,
  selEnd: number,
): { value: string; caret: number } {
  const len = value.length;
  const clamp = (n: number): number => {
    if (!Number.isFinite(n)) return len;
    return Math.max(0, Math.min(len, Math.floor(n)));
  };
  const a = clamp(selStart);
  const b = clamp(selEnd);
  const start = Math.min(a, b);
  const end = Math.max(a, b);
  const next = value.slice(0, start) + "\n" + value.slice(end);
  return { value: next, caret: start + 1 };
}
