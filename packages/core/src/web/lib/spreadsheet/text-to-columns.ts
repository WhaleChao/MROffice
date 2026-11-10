/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helper for "Text to Columns" (Excel / WPS Data tab parity).
 *
 * Splits each cell of ONE selected column and fans the parts out across the
 * columns to the right, mirroring Excel's Text-to-Columns. Two modes are
 * supported: DELIMITED (one or more delimiter characters, qualifier and
 * consecutive-delimiter rules) and FIXED WIDTH (character-break positions).
 *
 * This is a TOTAL, side-effect-free function: an empty selection, a single cell,
 * blanks, and a value with no delimiter present all return cleanly (a no-delimiter
 * cell keeps its value whole in its own column). It returns the cell writes the
 * caller should apply (row/col -> value) plus the maximum fan-out width, so the
 * caller can warn before clobbering adjacent columns and clamp to the schema
 * column bound. The caller writes these through the normal grid edit path, which
 * is what gives undo / autosave / sheet-protection for free.
 */

/** A single cell write the split produces: a target coordinate and its value. */
export interface ColumnSplitWrite {
  row: number;
  col: number;
  value: string;
}

/** The vertical span the split reads from (inclusive row range). */
export interface ColumnSplitRange {
  r1: number;
  r2: number;
}

/** The result of splitting one column on a delimiter. */
export interface ColumnSplitResult {
  /** The cell writes to apply (row/col -> value), in row-major order. */
  writes: ColumnSplitWrite[];
  /** The widest fan-out across the split rows (1 = nothing actually split). */
  width: number;
  /** Widest parsed source row before any columns marked "skip" are removed. */
  sourceWidth?: number;
  /** Format of each emitted destination column, after skipped fields collapse. */
  columnFormats?: TextToColumnsColumnFormat[];
}

/** Excel Text-to-Columns text qualifier choices. */
export type TextToColumnsQualifier = "double" | "single" | "none";

/** Per-field output choices from the final wizard step. */
export type TextToColumnsColumnFormat =
  | "general"
  | "text"
  | "date"
  | "skip";

export interface DelimitedColumnSplitOptions {
  /** One-character delimiter choices. Multiple entries are active together. */
  delimiters: readonly string[];
  /** Double quote is Excel's default. */
  textQualifier?: TextToColumnsQualifier;
  /** Collapse a run of any active delimiter characters into one boundary. */
  consecutiveDelimiters?: boolean;
  /** Top-left output cell. Defaults to the source range's first row/column. */
  destination?: { row: number; col: number };
  /** Format per parsed field. Missing entries are General. */
  columnFormats?: readonly TextToColumnsColumnFormat[];
}

interface FixedWidthColumnSplitOptions {
  destination?: { row: number; col: number };
  columnFormats?: readonly TextToColumnsColumnFormat[];
}

const MAX_DELIMITERS = 8;

function normalizedDelimiters(raw: readonly string[]): Set<string> {
  const delimiters = new Set<string>();
  for (const candidate of raw) {
    const value = String(candidate ?? "").slice(0, 1);
    if (!value) continue;
    delimiters.add(value);
    if (delimiters.size >= MAX_DELIMITERS) break;
  }
  return delimiters;
}

function qualifierCharacter(
  kind: TextToColumnsQualifier | undefined,
): string | null {
  if (kind === "single") return "'";
  if (kind === "none") return null;
  return '"';
}

/**
 * Parse one cell using Excel's delimited Text-to-Columns rules.
 *
 * Unlike the CSV helper, this accepts several delimiter characters at once and
 * lets the author choose double quote, single quote, or no qualifier. A doubled
 * qualifier inside a qualified field is unescaped to one literal character.
 */
export function parseTextToColumnsFields(
  value: string,
  options: Pick<
    DelimitedColumnSplitOptions,
    "delimiters" | "textQualifier" | "consecutiveDelimiters"
  >,
): string[] {
  const delimiters = normalizedDelimiters(options.delimiters);
  if (delimiters.size === 0) return [value];
  const qualifier = qualifierCharacter(options.textQualifier);
  const fields: string[] = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < value.length; index += 1) {
    const char = value[index]!;
    if (qualifier && char === qualifier) {
      if (quoted && value[index + 1] === qualifier) {
        field += qualifier;
        index += 1;
      } else if (quoted) {
        quoted = false;
      } else if (field.length === 0) {
        quoted = true;
      } else {
        // A qualifier only opens at the start of a field. Quotes embedded in
        // ordinary text are data, not a reason to suppress later delimiters.
        field += char;
      }
      continue;
    }
    if (!quoted && delimiters.has(char)) {
      fields.push(field);
      field = "";
      if (options.consecutiveDelimiters) {
        while (index + 1 < value.length && delimiters.has(value[index + 1]!)) {
          index += 1;
        }
      }
      continue;
    }
    field += char;
  }
  fields.push(field);
  return fields;
}

function normalizedColumnFormat(
  raw: TextToColumnsColumnFormat | undefined,
): TextToColumnsColumnFormat {
  return raw === "text" || raw === "date" || raw === "skip"
    ? raw
    : "general";
}

function emittedColumnFormats(
  sourceWidth: number,
  requested: readonly TextToColumnsColumnFormat[] | undefined,
  maxParts: number,
): TextToColumnsColumnFormat[] {
  const output: TextToColumnsColumnFormat[] = [];
  for (let index = 0; index < sourceWidth; index += 1) {
    const format = normalizedColumnFormat(requested?.[index]);
    if (format === "skip") continue;
    output.push(format);
    if (output.length >= maxParts) break;
  }
  return output;
}

/**
 * Full delimited Text-to-Columns authority. It supports Excel's simultaneous
 * delimiter switches, qualifier choice, consecutive-delimiter option,
 * destination cell, and per-field General/Text/Date/Skip disposition.
 */
export function splitColumnByDelimiters(
  cells: Record<string, { value?: string } | undefined>,
  sourceColumn: number,
  range: ColumnSplitRange,
  options: DelimitedColumnSplitOptions,
  maxColumns = Infinity,
): ColumnSplitResult {
  const writes: ColumnSplitWrite[] = [];
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const destinationRow = Math.max(
    0,
    Math.trunc(options.destination?.row ?? r1),
  );
  const destinationColumn = Math.max(
    0,
    Math.trunc(options.destination?.col ?? sourceColumn),
  );
  const maxParts = Math.max(1, maxColumns - destinationColumn);
  let sourceWidth = 1;
  let width = 0;

  for (let row = r1; row <= r2; row += 1) {
    const raw = cells[`${row}:${sourceColumn}`]?.value ?? "";
    if (raw === "") continue;
    const fields = parseTextToColumnsFields(raw, options);
    sourceWidth = Math.max(sourceWidth, fields.length);
    let outputColumn = 0;
    for (let index = 0; index < fields.length; index += 1) {
      const format = normalizedColumnFormat(options.columnFormats?.[index]);
      if (format === "skip") continue;
      if (outputColumn >= maxParts) break;
      writes.push({
        row: destinationRow + (row - r1),
        col: destinationColumn + outputColumn,
        value: fields[index]!,
      });
      outputColumn += 1;
    }
    width = Math.max(width, outputColumn);
  }

  const columnFormats = emittedColumnFormats(
    sourceWidth,
    options.columnFormats,
    maxParts,
  );
  return {
    writes,
    width,
    sourceWidth,
    columnFormats: columnFormats.slice(0, width),
  };
}

/**
 * Split the cells of `columnIndex` (over the inclusive row range `range`) on
 * `delimiter`, returning the writes that lay the parts out left-to-right starting
 * at `columnIndex`. Part 0 overwrites the source column; parts 1..n land in the
 * columns to the right. The split is quote-aware (reuses `parseDelimited`), so a
 * `"a,b"` token stays whole.
 *
 * Per row, every produced part is written — including trailing empties from a
 * `"a,,c"` value (so the slots line up) — but a part beyond the row's own fan-out
 * is NOT written, so a shorter row never stamps blanks over a longer row's data.
 * A cell with no delimiter present yields a single part (its own value, unchanged)
 * and contributes width 1. An empty cell is skipped entirely (no write), matching
 * how Excel leaves a blank source cell blank.
 *
 * The fan-out is clamped to `maxColumns` (the sheet's hard column-count bound): a
 * destination can never reach column `maxColumns`, so at most `maxColumns -
 * columnIndex` parts are emitted — both the writes AND the reported `width` are
 * capped. This MUST be enforced here (not just in the caller's filter): the caller
 * sizes its paste grid by `width`, and an unclamped width near the right edge would
 * grow the sheet's columnCount past the bound and autosave would reject the page.
 *
 * @param cells the sheet's `row:col` -> { value } map (only `value` is read)
 * @param columnIndex the single column to split (source = destination part 0)
 * @param range the inclusive row span to operate on
 * @param delimiter the field separator (a single character: "," "\t" ";" " " …)
 * @param maxColumns the sheet's column-count bound; no part lands at or past it
 *   (default `Infinity` = unclamped, for callers that bound the result themselves)
 */
export function splitColumnByDelimiter(
  cells: Record<string, { value?: string } | undefined>,
  columnIndex: number,
  range: ColumnSplitRange,
  delimiter: string,
  maxColumns = Infinity,
): ColumnSplitResult {
  return splitColumnByDelimiters(
    cells,
    columnIndex,
    range,
    {
      delimiters: delimiter ? [delimiter] : [],
      textQualifier: "double",
      consecutiveDelimiters: false,
    },
    maxColumns,
  );
}

/**
 * Normalize a list of FIXED-WIDTH break positions into a clean, sorted, de-duped
 * set of POSITIVE integer character offsets. A break at offset `k` cuts a string
 * BEFORE character index `k` (so breaks [3, 7] split "abcdefghij" into "abc" |
 * "defg" | "hij"). Offsets <= 0 and non-finite / non-integer values are dropped
 * (a break at the very start makes no column), and duplicates collapse so two
 * sliders at the same spot don't create an empty column. Exported for the UI's
 * break-position editor (it dedupes the user's drag positions the same way the
 * split does).
 */
export function normalizeFixedWidthBreaks(
  breaks: readonly number[],
): number[] {
  const seen = new Set<number>();
  for (const raw of breaks) {
    const n = Math.trunc(Number(raw));
    if (Number.isFinite(n) && n > 0) seen.add(n);
  }
  return [...seen].sort((a, b) => a - b);
}

function splitFixedWidthFieldsAtCuts(
  value: string,
  cuts: readonly number[],
): string[] {
  const fields: string[] = [];
  let previous = 0;
  for (const cut of cuts) {
    fields.push(value.slice(previous, cut));
    previous = cut;
  }
  fields.push(value.slice(previous));
  return fields;
}

/**
 * Slice one raw record with the same normalized fixed-width boundary authority
 * used by Text to Columns. File import consumes this helper so preview and cell
 * transforms cannot disagree about sorting, duplicate breaks, short rows, or
 * whitespace preservation.
 */
export function splitFixedWidthFields(
  value: string,
  breaks: readonly number[],
): string[] {
  return splitFixedWidthFieldsAtCuts(value, normalizeFixedWidthBreaks(breaks));
}

/**
 * Split the cells of `columnIndex` (over the inclusive row range `range`) at the
 * given FIXED-WIDTH break positions (Excel's "Fixed width" Text-to-Columns
 * mode), returning the writes that lay the field slices out left-to-right
 * starting at `columnIndex`. Slice 0 (chars before the first break) overwrites
 * the source column; later slices land to the right.
 *
 * The breaks are character offsets where a new field BEGINS: breaks [3, 7] cut a
 * value into [0,3), [3,7), [7,end). Breaks are normalized first (sorted, de-
 * duped, positive integers only); a break at or past a row's own length simply
 * yields an empty trailing slice for that row, exactly as Excel pads short rows.
 * Whitespace inside a slice is preserved verbatim (Excel does not trim fixed-
 * width fields). An empty source cell is skipped entirely (no write), matching
 * the delimited path and Excel's blank-cell behaviour.
 *
 * Like the delimited split, the fan-out is clamped to `maxColumns` (the sheet's
 * hard column-count bound) so the result can never grow the sheet past the cap
 * (which would make autosave reject the page): at most `maxColumns - columnIndex`
 * slices are emitted, and the reported `width` is capped to match. With no valid
 * breaks the result is width 1 (every value kept whole in its own column), which
 * the caller treats as "nothing to split".
 *
 * @param cells the sheet's `row:col` -> { value } map (only `value` is read)
 * @param columnIndex the single column to split (source = destination slice 0)
 * @param range the inclusive row span to operate on
 * @param breaks character offsets where each new field begins (normalized here)
 * @param maxColumns the sheet's column-count bound; no slice lands at or past it
 *   (default `Infinity` = unclamped, for callers that bound the result themselves)
 */
export function splitColumnByFixedWidth(
  cells: Record<string, { value?: string } | undefined>,
  columnIndex: number,
  range: ColumnSplitRange,
  breaks: readonly number[],
  maxColumns = Infinity,
  options: FixedWidthColumnSplitOptions = {},
): ColumnSplitResult {
  const writes: ColumnSplitWrite[] = [];
  let width = 1;
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const cuts = normalizeFixedWidthBreaks(breaks);
  // The widest fan-out that still fits (mirrors the delimited path): slices run
  // from columnIndex up to (but not including) maxColumns. Always >= 1.
  const destinationRow = Math.max(
    0,
    Math.trunc(options.destination?.row ?? r1),
  );
  const destinationColumn = Math.max(
    0,
    Math.trunc(options.destination?.col ?? columnIndex),
  );
  const maxParts = Math.max(1, maxColumns - destinationColumn);
  const sourceWidth = cuts.length + 1;
  const columnFormats = emittedColumnFormats(
    sourceWidth,
    options.columnFormats,
    maxParts,
  );

  for (let row = r1; row <= r2; row += 1) {
    const raw = cells[`${row}:${columnIndex}`]?.value ?? "";
    if (raw === "") continue;
    // A cut beyond the value length yields an empty trailing slice (Excel pads
    // short rows). Code-unit indices match the UI character ruler.
    const allParts = splitFixedWidthFieldsAtCuts(raw, cuts);
    // Clamp to the column bound (the tail is lost, as Excel also caps at the
    // sheet's last column).
    let outputColumn = 0;
    for (let index = 0; index < allParts.length; index += 1) {
      const format = normalizedColumnFormat(options.columnFormats?.[index]);
      if (format === "skip") continue;
      if (outputColumn >= maxParts) break;
      writes.push({
        row: destinationRow + (row - r1),
        col: destinationColumn + outputColumn,
        value: allParts[index]!,
      });
      outputColumn += 1;
    }
    if (outputColumn > width) width = outputColumn;
  }

  return {
    writes,
    width,
    sourceWidth,
    columnFormats: columnFormats.slice(0, width),
  };
}
