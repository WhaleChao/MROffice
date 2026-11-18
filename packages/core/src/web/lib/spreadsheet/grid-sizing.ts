/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx, pxToPt } from "../office-units";
import OFFICE_DIGIT_ADVANCES from "./office-digit-advances.json";
import OFFICE_DIGIT_WHOLE_EM from "./office-digit-whole-em.json";
import OFFICE_ROW_METRICS from "./office-row-metrics.json";
import OFFICE_ROW_HEIGHTS from "./office-row-heights.json";
import OFFICE_ROW_PARTS from "./office-row-parts.json";
import OFFICE_ROW_SCRIPTS from "./office-row-scripts.json";
import { excelFontRecordOf, excelWindowsFaceKey, type ExcelFontRecordClass } from "../office-fonts/excel-windows-faces";
import { advancesOf } from "./office-hinted-wrap";

/**
 * Pure helpers for spreadsheet column-width / row-height resizing (Excel/WPS
 * parity).
 *
 * The sheet model can override the default column width / row height per line
 * via two optional maps keyed by stringified line index: `columnWidths` and
 * `rowHeights`. A missing entry means "use the default". This module keeps all
 * the size math out of the React component so it can be unit-tested in
 * isolation, and so the renderer (which sizes each cell) and the virtualization
 * window (which decides the visible band from the scroll offset) agree on the
 * model.
 *
 * Sizes are clamped to a sane min/max so a drag can never collapse a line to
 * zero (lost forever) or blow it up past the schema bound. Cumulative offsets
 * let variable-size virtualization stay correct: the first visible line is the
 * one whose running offset crosses the scroll position, and the rendered table
 * is positioned at that line's offset so the body never jumps.
 */

/** A per-line size override map keyed by stringified line index. */
export type SizeMap = Record<number | string, number>;

// Shared blank-sheet geometry. Excel's default row height is 15 points, which
// is 20 CSS pixels at 96 DPI; the column width is the reference desktop grid's
// compact 80-pixel band. Imported workbook defaults and persisted per-line
// overrides remain authoritative over these editor-created-sheet fallbacks.
export const DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX = 80;
/**
 * The default column width a NEW sheet is created with: Excel's 8.43
 * characters of the Normal font, 64 px at 96 DPI (XLS-A095). It is stamped on
 * the new sheet as its `defaultColumnWidth`, so sheets saved before keep the
 * 80-px fallback above (and every object anchored on them stays put) while a
 * new workbook or sheet opens on Excel's grid. The xlsx export writes it as
 * the sheet's defaultColWidth.
 */
export const NEW_SHEET_DEFAULT_COLUMN_WIDTH_PX = 64;
export const DEFAULT_SPREADSHEET_ROW_HEIGHT_PX = 20;
export const DEFAULT_SPREADSHEET_HEADER_WIDTH_PX = 48;
export const DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX = 20;
export const DEFAULT_SPREADSHEET_ROW_COUNT = 50;
export const DEFAULT_SPREADSHEET_COLUMN_COUNT = 20;

// A column is as wide as Column Width accepts: 0 to 255 characters, where 0
// hides it. These are that range in pixels, the same one an imported grid
// already keeps (nativeColumnCssClamp), so a width typed as 1 or as 255 is
// the width the column gets instead of being moved to an older, narrower
// drag range without a word (live audit B372). A row runs from a single text
// line up to a tall multi-line block.
export const MIN_COLUMN_WIDTH_PX = 1;
export const MAX_COLUMN_WIDTH_PX = 4096;
export const MIN_ROW_HEIGHT_PX = 18;
export const MAX_ROW_HEIGHT_PX = 512;

// Excel's full worksheet address space. These are ADDRESS bounds, not collection
// budgets: a sparse sheet may contain XFD1048576 without allocating every line
// before it.
export const MAX_GRID_COLUMNS = 16_384;
export const MAX_GRID_ROWS = 1_048_576;

// Persisted size overrides remain independently bounded. Reusing the row
// address ceiling as an entry-count limit would let a hostile sparse workbook
// allocate a million-entry map even though almost every line uses the default
// size. The row budget is the importer's own (page_projection
// _MAX_ROW_SIZE_OVERRIDES, 200,000; past it the import reports
// xlsx_line_sizes_not_projected): at 10,000 here, a sheet with more sized or
// hidden rows had them set aside on open and a row edit left them behind

// outline groups. A column map is bounded by the 16,384 columns themselves, a small map,
// and a workbook may size every one of them: P07_62629 states <col> widths out
// to XFD (7,757 of them not the default), where a 1,000-entry budget set the
// sheet's widths aside whole (X-065). The importer's own budget is the column
// count too (page_projection _MAX_COLUMN_SIZE_OVERRIDES).
export const MAX_GRID_COLUMN_SIZE_OVERRIDES = MAX_GRID_COLUMNS;
export const MAX_GRID_ROW_SIZE_OVERRIDES = 200_000;

/** The tallest row a file states: 409.5 pt, 546 px. */
export const MAX_NATIVE_ROW_HEIGHT_PX = 546;




export const MAX_DRAWING_LAYER_X_PX = DEFAULT_SPREADSHEET_HEADER_WIDTH_PX + MAX_GRID_COLUMNS * MAX_COLUMN_WIDTH_PX;
export const MAX_DRAWING_LAYER_Y_PX = DEFAULT_SPREADSHEET_HEADER_HEIGHT_PX + MAX_GRID_ROWS * MAX_NATIVE_ROW_HEIGHT_PX;

const SPARSE_AXIS_RESOLVER = Symbol("spreadsheet-sparse-axis-resolver");

interface SparseAxisResolverMetadata {
  offsetOfLine(index: number): number;
  visibleLine?(index: number, count: number, direction: 1 | -1): number;
  defaultSize: number;
  overrides: ReadonlyMap<number, number>;
  hidden?: ReadonlySet<number>;
  /** Hidden lines held as inclusive runs (``maskHiddenAxisLines``). */
  hiddenSpans?: ReadonlyArray<readonly [number, number]>;
  intrinsic?: SparseAxisSizeResolver;
}

export type SparseAxisSizeResolver = ((line: number) => number) & {
  [SPARSE_AXIS_RESOLVER]: SparseAxisResolverMetadata;
};

/**
 * Build a line-size resolver whose cumulative offset is O(log K), where K is
 * the number of persisted overrides. This keeps virtual navigation to row
 * 1,048,576 bounded while preserving the existing function-shaped API used by
 * the renderer and drag logic.
 */
export function createSparseAxisSizeResolver(options: {
  overrides: SizeMap | undefined;
  defaultSize: number;
  clampSize: (value: unknown, fallback: number) => number;
  /** The live drag preview: one line, or every line one drag is sizing. */
  transientOverride?:
    | { index: number; size: number }
    | readonly { index: number; size: number }[]
    | null;
}): SparseAxisSizeResolver {
  const fallback = options.clampSize(options.defaultSize, options.defaultSize);
  const transient = options.transientOverride;
  const entries = new Map<number, number>();
  for (const [rawIndex, rawSize] of Object.entries(options.overrides ?? {})) {
    const index = Number(rawIndex);
    if (!Number.isInteger(index) || index < 0) continue;
    entries.set(index, options.clampSize(rawSize, fallback));
  }
  for (const line of transient
    ? Array.isArray(transient)
      ? transient
      : [transient as { index: number; size: number }]
    : []) {
    if (!Number.isInteger(line.index) || line.index < 0) continue;
    entries.set(line.index, options.clampSize(line.size, fallback));
  }
  const sorted = [...entries.entries()].sort(([left], [right]) => left - right);
  const indices = sorted.map(([index]) => index);
  const prefixDeltas: number[] = [];
  let delta = 0;
  for (const [, size] of sorted) {
    delta += size - fallback;
    prefixDeltas.push(delta);
  }
  const upperBound = (index: number): number => {
    let low = 0;
    let high = indices.length;
    while (low < high) {
      const middle = low + Math.floor((high - low) / 2);
      if (indices[middle] < index) low = middle + 1;
      else high = middle;
    }
    return low;
  };
  const resolver = ((line: number) => {
    if (!Number.isInteger(line) || line < 0) return fallback;
    return entries.get(line) ?? fallback;
  }) as SparseAxisSizeResolver;
  resolver[SPARSE_AXIS_RESOLVER] = {
    defaultSize: fallback,
    overrides: entries,
    offsetOfLine(index: number): number {
      const bound = Math.max(0, Math.floor(index));
      const position = upperBound(bound);
      return (
        bound * fallback + (position > 0 ? prefixDeltas[position - 1] : 0)
      );
    },
  };
  return resolver;
}

/** View-only collapse. Stored dimensions remain in the base resolver. */
export function maskHiddenAxisLines(
  base: SparseAxisSizeResolver,
  hiddenLines: Iterable<number>,
  



  hiddenSpans: ReadonlyArray<readonly [number, number]> = [],
): SparseAxisSizeResolver {
  const baseMetadata = base[SPARSE_AXIS_RESOLVER];
  const intrinsic = baseMetadata.intrinsic ?? base;
  const source = intrinsic[SPARSE_AXIS_RESOLVER];
  // Runs of hidden lines, merged: the base's and these.
  const spans: Array<[number, number]> = [];
  for (const [first, last] of [...(baseMetadata.hiddenSpans ?? []), ...hiddenSpans]
    .filter(([first, last]) => Number.isSafeInteger(first) && Number.isSafeInteger(last) && first >= 0 && first <= last)
    .sort(([left], [right]) => left - right)) {
    const previous = spans[spans.length - 1];
    if (previous && first <= previous[1] + 1) previous[1] = Math.max(previous[1], last);
    else spans.push([first, last]);
  }
  const spanFirsts = spans.map(([first]) => first);
  const spanPrefix: number[] = [];
  let spanned = 0;
  for (const [first, last] of spans) { spanned += last - first + 1; spanPrefix.push(spanned); }
  const lowerBound = (values: readonly number[], index: number) => {
    let low = 0, high = values.length;
    while (low < high) { const middle = (low + high) >>> 1; if (values[middle] < index) low = middle + 1; else high = middle; }
    return low;
  };
  /** The span holding `index`, or undefined. */
  const spanAt = (index: number): [number, number] | undefined => {
    const span = spans[lowerBound(spanFirsts, index + 1) - 1];
    return span && index <= span[1] ? span : undefined;
  };
  /** How many spanned lines lie before `bound`. */
  const spannedBefore = (bound: number): number => {
    const position = lowerBound(spanFirsts, bound);
    if (position === 0) return 0;
    const [first, last] = spans[position - 1];
    return (position > 1 ? spanPrefix[position - 2] : 0) + Math.min(last, bound - 1) - first + 1;
  };
  const hidden = new Set([...(baseMetadata.hidden ?? []), ...hiddenLines]
    .filter(index => Number.isSafeInteger(index) && index >= 0 && !spanAt(index)));
  if (hidden.size === 0 && spans.length === 0) return base;
  const isHidden = (index: number) => hidden.has(index) || spanAt(index) !== undefined;
  const indices = [...hidden].sort((a, b) => a - b);
  const runs: Array<{ first: number; last: number }> = [];
  for (const [first, last] of [...indices.map((index): [number, number] => [index, index]), ...spans].sort(([a], [b]) => a - b)) {
    const previous = runs[runs.length - 1];
    if (previous && previous.last + 1 >= first) previous.last = Math.max(previous.last, last);
    else runs.push({ first, last });
  }
  const overrides = [...source.overrides].filter(([index]) => !isHidden(index)).sort(([a], [b]) => a - b);
  const overrideIndices = overrides.map(([index]) => index), prefixDeltas: number[] = [];
  let delta = 0;
  for (const [, size] of overrides) { delta += size - source.defaultSize; prefixDeltas.push(delta); }
  const positive = overrides.filter(([, size]) => size > 0).map(([index]) => index);
  const resolver = ((index: number) => isHidden(index) ? 0 : intrinsic(index)) as SparseAxisSizeResolver;
  resolver[SPARSE_AXIS_RESOLVER] = {
    defaultSize: source.defaultSize, overrides: source.overrides, hidden, hiddenSpans: spans, intrinsic,
    offsetOfLine(index) {
      const bound = Math.max(0, Math.floor(index)), position = lowerBound(overrideIndices, bound);
      // Count visible default lines instead of subtracting two large fractional
      // sums: a fully hidden interval must have exactly zero extent for hits.
      return (bound - lowerBound(indices, bound) - spannedBefore(bound)) * source.defaultSize + (position ? prefixDeltas[position - 1] : 0);
    },
    visibleLine(index, count, direction) {
      if (index < 0 || index >= count) return index;
      if (source.defaultSize <= 0) {
        const position = direction === 1 ? lowerBound(positive, index) : lowerBound(positive, index + 1) - 1;
        const found = positive[position];
        return found === undefined ? (direction === 1 ? count : -1) : found;
      }
      let candidate = index;
      while (candidate >= 0 && candidate < count) {
        let low = 0, high = runs.length;
        while (low < high) { const middle = (low + high) >>> 1; if (runs[middle].first <= candidate) low = middle + 1; else high = middle; }
        const run = runs[low - 1];
        if (run && candidate <= run.last) { candidate = direction === 1 ? run.last + 1 : run.first - 1; continue; }
        if (intrinsic(candidate) > 0) return candidate;
        candidate += direction;
      }
      return candidate;
    },
  };
  return resolver;
}

export function nextVisibleAxisLine(index:number,count:number,size:(line:number)=>number,direction:1|-1=1):number {
  const sparse=(size as Partial<SparseAxisSizeResolver>)[SPARSE_AXIS_RESOLVER];
  if(sparse?.visibleLine)return sparse.visibleLine(index,count,direction);
  let candidate=index;
  while(candidate>=0&&candidate<count&&size(candidate)<=0)candidate+=direction;
  return candidate;
}

export function hasHiddenAxisLines(size:(line:number)=>number):boolean {
  return Boolean((size as Partial<SparseAxisSizeResolver>)[SPARSE_AXIS_RESOLVER]?.visibleLine);
}

/**
 * Clamp a raw (possibly fractional / out-of-range / non-finite) column width to
 * the supported pixel range, rounding to a whole pixel. Non-finite input falls
 * back to the supplied default (itself clamped) so a corrupt drag never yields
 * NaN.
 */
export function clampColumnWidth(value: unknown, fallback: number): number {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return clampColumnWidth(fallback, MIN_COLUMN_WIDTH_PX);
  }
  return Math.min(MAX_COLUMN_WIDTH_PX, Math.max(MIN_COLUMN_WIDTH_PX, Math.round(n)));
}

/**
 * Clamp a raw row height to the supported pixel range (see clampColumnWidth).
 */
export function clampRowHeight(value: unknown, fallback: number): number {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return clampRowHeight(fallback, MIN_ROW_HEIGHT_PX);
  }
  return Math.min(MAX_ROW_HEIGHT_PX, Math.max(MIN_ROW_HEIGHT_PX, Math.round(n)));
}

// Excel's own units for a line's size. A column width counts characters of the
// Normal style font's maximum digit width, plus 5 pixels of cell margin and
// gridline (ISO/IEC 29500-1 §18.3.1.13, CT_Col); a row height counts points.
// The Column Width dialog and the drag readout show characters, which are
// `(pixels - 5) / 7` for the 7-pixel digit of the default font (Calibri or
// Aptos Narrow 11 at 96 DPI, where 8.43 characters are 64 pixels); a row is
// `points / 0.75` pixels. The model keeps pixels.
//
// The FILE's width is a different number: `CT_Col@width` and
// `sheetFormatPr@defaultColWidth` hold the characters PLUS the 5-pixel margin,
// in characters. Desktop Excel stores a column set to 8 as 8.7109375 at a
// 7-pixel digit (8.83203125 at a 6-pixel one), so reading the file's width as
// dialog characters and adding the margin again made every sized column 5
// pixels wider than Excel draws it. Read the file's width with
// columnFileWidthToPixels (§18.3.1.13's own truncation).
export const DEFAULT_MAX_DIGIT_WIDTH_PX = 7;
export const COLUMN_WIDTH_PADDING_PX = 5;
export const MAX_COLUMN_WIDTH_CHARACTERS = 255;
export const MAX_ROW_HEIGHT_POINTS = 409;

/** Pixels of a column `characters` wide (see DEFAULT_MAX_DIGIT_WIDTH_PX). */
export function columnCharactersToPixels(
  characters: number,
  digitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
  dpi = 96,
): number {
  const device = characters * digitWidth + COLUMN_WIDTH_PADDING_PX;
  return dpi === 96 ? device : (device * 96) / dpi;
}

/**
 * The maximum digit width, in 96dpi pixels, of a workbook grid font: the unit
 * every column width of that workbook is counted in (Calibri and Aptos Narrow
 * 11 are 7; 等线 11, Chinese Excel's Normal font, is 8, which is why its
 * default column is 72 pixels, "8.38"). The widest digit's advance comes from
 * a table measured off the font files Office ships, keyed by every name a
 * face carries (Excel writes the author's: 等线, 游ゴシック); the API keeps an
 * identical copy, so the page, its print and its export count alike. A face
 * missing from the table counts like Calibri, on both sides.
 */
const WHOLE_EM_DIGIT_FACES: ReadonlySet<string> = new Set(OFFICE_DIGIT_WHOLE_EM as readonly string[]);

export function officeDigitWidth(
  font: { readonly family?: string; readonly sizePt?: number; readonly record?: ExcelFontRecordClass | null } | undefined,
): number {
  const size = font?.sizePt;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) {
    return DEFAULT_MAX_DIGIT_WIDTH_PX;
  }
  const advances = OFFICE_DIGIT_ADVANCES as Readonly<Record<string, number>>;
  // Read under the face Windows Excel draws the name in: one it lacks
  // (Helvetica Neue) is Arial's digits, or its record's class's (lookwin
  // faces probe).
  const own = (font?.family ?? "").trim().toLowerCase();
  const key = excelWindowsFaceKey(font?.family, font?.record);
  // The East Asian fixed-pitch faces whose digits are half the em (宋体,
  // ＭＳ Ｐゴシック ...) carry no hinted widths: GDI draws them at the whole-
  // pixel em, half of which rounds up -- 宋体 11 is a 15-pixel em with 8-pixel
  // digits, Chinese Excel 2007's "8.38 (72 像素)" default column (X-090). The
  // API's office_digit_width reads the same list.
  if (WHOLE_EM_DIGIT_FACES.has(own) || WHOLE_EM_DIGIT_FACES.has(key)) {
    const em = Math.floor((size * 96) / 72 + 0.5);
    return Math.max(1, Math.floor(em / 2 + 0.5));
  }
  const ratio = advances[key] ?? advances.calibri;
  return Math.max(1, Math.round((ratio * size * 96) / 72));
}

/**
 * One indent level of an imported grid on the screen, px: Windows Excel's is
 * three spaces of the Normal font, each its whole-pixel hinted advance
 * (lookwin p87-lw; the API's office_indent_step, printed=False). Calibri 11
 * is 9 px a level, Arial 10 is 12, Times New Roman 14 is 15. A face the
 * hinted table lacks takes half its digit advance for the space.
 */
export function officeIndentStepPx(
  font: { readonly family?: string; readonly sizePt?: number; readonly record?: ExcelFontRecordClass | null } | undefined,
): number {
  const size = font?.sizePt;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return 3 * 3;
  const hinted = advancesOf({ family: font?.family ?? "", sizePt: size, record: font?.record })?.(" ");
  if (hinted !== undefined && hinted > 0) return 3 * hinted;
  const advances = OFFICE_DIGIT_ADVANCES as Readonly<Record<string, number>>;
  const ratio = advances[excelWindowsFaceKey(font?.family, font?.record)] ?? advances.calibri;
  return 3 * Math.max(1, Math.round((ratio / 2) * size * 96 / 72));
}

/**
 * How much wider Excel draws a run of digits than the browser in the same
 * face and size: the face's maximum digit width on Excel's grid
 * ({@link officeDigitWidth}, whole pixels) over the digit's plain advance at
 * that size. Calibri 11 is 7 / 7.43 (Excel's hinted digits are narrower:
 * S26_pivottable_date_field_filter's "2016/ 1/ 6. 0:00" fits its 101-pixel
 * column in Excel and measured 96.5 pixels in the browser, X-101); 宋体 11 is
 * 8 / 7.33. A number's fit is judged at Excel's widths.
 */
export function officeDigitWidthScale(
  font: { readonly family?: string; readonly sizePt?: number; readonly record?: ExcelFontRecordClass | null } | undefined,
): number {
  const size = font?.sizePt;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return 1;
  const advances = OFFICE_DIGIT_ADVANCES as Readonly<Record<string, number>>;
  const ratio = advances[excelWindowsFaceKey(font?.family, font?.record)] ?? advances.calibri;
  const plain = (ratio * size * 96) / 72;
  return plain > 0 ? officeDigitWidth(font) / plain : 1;
}

/**
 * The automatic row height, in 96dpi pixels, of a grid font: what Windows
 * Excel gives a row that holds nothing taller. Its line box from the face's
 * own metrics at GDI's whole-pixel em (Windows metrics with the hhea line gap
 * beyond them, or the typo metrics a face asks for, as Aptos does), plus one
 * pixel: ceil(ascent) + round(descent) + floor(leading) + 1. It reproduces
 * every Normal font default found in real Excel files: Calibri and Aptos
 * Narrow 11 are 20 (15 pt), Arial and Times New Roman 10 are 17, 宋体 11 is
 * 18 and 12 is 19 (14.25 pt), 新細明體 11 is 19, 游ゴシック 11 is 25 (18.75 pt),
 * 맑은 고딕 11 is 22. The integers come from office-row-metrics.json (the API
 * keeps an identical copy); a face missing from it measures like Calibri.
 */
/** Whether `family` is a face office-row-metrics.json has Excel's own metrics
 * for (officeRowHeight measures any other one like Calibri). */
export function officeRowMetricsKnown(family: string | null | undefined, sizePt?: number, record?: ExcelFontRecordClass | null): boolean {
  const key = excelWindowsFaceKey(family, record);
  if (key === "" || key.startsWith("_")) return false;
  return Object.prototype.hasOwnProperty.call(OFFICE_ROW_METRICS, key) ||
    (sizePt !== undefined && (officeRowHeightMeasured(key, sizePt) !== undefined || officeRowPartsMeasured(key, sizePt, false) !== undefined));
}

/** A grid font as the row rules read it: its face, size and weight. */
export interface OfficeRowFont {
  readonly family?: string;
  readonly sizePt?: number;
  readonly bold?: boolean;
  /** The font record's class where Windows lacks its face (excelFontRecordOf). */
  readonly record?: ExcelFontRecordClass | null;
  /** A superscript or subscript face: its line is taller (officeScriptExtra). */
  readonly script?: "super" | "sub";
  /** A top or centred position's face: it stands at its own line, never on
   * the row's baseline (officeFloorRowHeight). */
  readonly standing?: boolean;
}

/** A face the table lacks: the share of its em the measured faces' scripts
 * add (lookwin p109-lw, twelve faces at 6-48 pt). */
const SCRIPT_EM_SHARE = { super: 0.17, sub: 0.2 } as const;

/**
 * The px a superscript adds to a line's ascent ("super") or a subscript to
 * its descent ("sub") in Windows Excel's automatic row (lookwin p109-lw:
 * Excel's own rows on resave): Arial 12 stands 20 px, superscript 24,
 * subscript 26, and a superscript beside Calibri 11 stands 25 (its descent
 * is still Arial's). The face's own sizes (office-row-scripts.json, the API's
 * office_row_scripts.json), else the nearest measured size's share of the em.
 * The API's office_script_extra is the same rule.
 */
export function officeScriptExtra(font: OfficeRowFont): number {
  const script = font.script;
  const size = font.sizePt;
  if ((script !== "super" && script !== "sub") || typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return 0;
  const at = script === "super" ? 0 : 1;
  const faces = (OFFICE_ROW_SCRIPTS as { faces: Readonly<Record<string, Readonly<Record<string, readonly number[]>>>> }).faces;
  const sizes = faces[(font.family ?? "").trim().toLowerCase()] ?? faces[excelWindowsFaceKey(font.family, font.record)];
  let share: number = SCRIPT_EM_SHARE[script];
  if (sizes && Object.keys(sizes).length) {
    const found = sizes[String(Math.round(size * 2) / 2)];
    if (found && found.length === 2) return found[at] ?? 0;
    let nearest = "";
    for (const key in sizes) if (!nearest || Math.abs(Number(key) - size) < Math.abs(Number(nearest) - size)) nearest = key;
    const parts = sizes[nearest];
    if (parts && parts.length === 2) share = (parts[at] ?? 0) / ((Number(nearest) * 96) / 72);
  }
  return Math.max(0, Math.round((share * size * 96) / 72));
}

/** Windows Excel's own split of a line, [ascent, descent] px with the height
 * ascent + descent + 1, where its printed rows were measured
 * (office-row-parts.json, lookwin probes p44/p44d): Verdana 10 is 17 px (the
 * formula 19), Arial 11 is 19 and bold 20, Calibri 10 is 13 + 3. The API's
 * office_row_parts.json is the same table. */
function officeRowPartsMeasured(key: string, sizePt: number, bold: boolean): readonly [number, number] | undefined {
  const faces = (OFFICE_ROW_PARTS as { faces: Readonly<Record<string, Readonly<Record<string, readonly number[]>>>> }).faces;
  const size = String(Math.round(sizePt * 2) / 2);
  const parts = (bold ? faces[`${key} bold`]?.[size] : undefined) ?? faces[key]?.[size];
  return parts && parts.length === 2 ? [parts[0], parts[1]] : undefined;
}

function officeRowPartsOf(font: OfficeRowFont | undefined): readonly [number, number] | undefined {
  const size = font?.sizePt;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return undefined;
  const own = (font?.family ?? "").trim().toLowerCase();
  // Measured under the name itself first (a table name the face key aliases), then the face Excel draws it in.
  return officeRowPartsMeasured(own, size, font?.bold === true) ?? officeRowPartsMeasured(excelWindowsFaceKey(font?.family, font?.record), size, font?.bold === true);
}

/** Windows Excel's own row height for `family` at `sizePt` where real files
 * say the metric formula is off (office-row-heights.json): ＭＳ Ｐゴシック 11
 * is 18 px (13.5 pt, the formula 16), Arial 12 is 20, Times New Roman 12 is
 * 21. The API's office_row_heights.json is the same table. A size arrives
 * from a cell's pixels (14.67 px is 11.0025 pt), so it is read to the half
 * point Excel's sizes come in: as "11.0025" it never met the table's "11". */
function officeRowHeightMeasured(key: string, sizePt: number): number | undefined {
  const sizes = (OFFICE_ROW_HEIGHTS as Readonly<Record<string, unknown>>)[key];
  if (!sizes || typeof sizes !== "object") return undefined;
  const pixels = (sizes as Readonly<Record<string, unknown>>)[String(Math.round(sizePt * 2) / 2)];
  return typeof pixels === "number" ? pixels : undefined;
}

export function officeRowHeight(
  font: OfficeRowFont | undefined,
): number {
  if (font?.script) return officeRowHeight({ ...font, script: undefined }) + officeScriptExtra(font);
  const split = officeRowPartsOf(font);
  if (split) return split[0] + split[1] + 1;
  const size = font?.sizePt;
  const metrics = OFFICE_ROW_METRICS as Readonly<Record<string, readonly number[]>>;
  // The face Windows Excel draws the name in: a Normal font it lacks rows
  // like Arial (lookwin faces probe), not like Calibri.
  const key = excelWindowsFaceKey(font?.family, font?.record);
  const [ascent, descent, leading, unitsPerEm] = metrics[key] ?? metrics.calibri;
  if (typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return 20;
  // A height Windows Excel was measured at under the name itself (its real
  // files' charset in it: 明朝, Osaka) before the substitute's.
  const measured = officeRowHeightMeasured((font?.family ?? "").trim().toLowerCase(), size) ?? officeRowHeightMeasured(key, size);
  if (measured !== undefined) return measured;
  const em = Math.ceil((size * 96) / 72 - 1e-9);
  return (
    Math.ceil((ascent * em) / unitsPerEm) +
    Math.floor((2 * descent * em + unitsPerEm) / (2 * unitsPerEm)) +
    Math.floor((leading * em) / unitsPerEm) +
    1
  );
}

/**
 * The two halves of {@link officeRowHeight}'s line, px: the descent (with the
 * line gap below it) from the face's metrics, and the ascent the rest of the
 * height (where the measured table corrects the formula, the correction is
 * the ascent's). Undefined for a face the metrics do not know.
 */
export function officeRowParts(
  font: OfficeRowFont,
): { ascent: number; descent: number } | undefined {
  if (font.script) {
    const parts = officeRowParts({ ...font, script: undefined });
    if (!parts) return undefined;
    const extra = officeScriptExtra(font);
    return font.script === "super" ? { ascent: parts.ascent + extra, descent: parts.descent } : { ascent: parts.ascent, descent: parts.descent + extra };
  }
  const split = officeRowPartsOf(font);
  if (split) return { ascent: split[0], descent: split[1] };
  const key = excelWindowsFaceKey(font.family, font.record);
  const metrics = (OFFICE_ROW_METRICS as Readonly<Record<string, readonly number[]>>)[key];
  const size = font.sizePt;
  if (!metrics || typeof size !== "number" || !Number.isFinite(size) || size <= 0 || size > 409) return undefined;
  const [, descent, leading, unitsPerEm] = metrics;
  const em = Math.ceil((size * 96) / 72 - 1e-9);
  const below = Math.floor((2 * descent * em + unitsPerEm) / (2 * unitsPerEm)) + Math.floor((leading * em) / unitsPerEm);
  return { ascent: officeRowHeight(font) - 1 - below, descent: below };
}

/**
 * Windows Excel's automatic height of one line set in several faces at once:
 * the tallest ascent over the deepest descent, not the tallest of their
 * heights. A row fits its cells' faces and the face of its positions with no
 * cell (the row's format, else the columns', else the workbook's font): an
 * Arial 12 row of a Calibri 11 workbook is 21 px, though each alone is 20
 * (7,417 rows of Windows-written files; 20 px where every empty position is
 * formatted in another face). Undefined when a face is one the metrics do not
 * know.
 */
/**
 * The automatic default row of a sheet whose file states its default without
 * customHeight, px: one line set in the faces of a row's empty positions,
 * each formatted column's font and, while a column is left unformatted, the
 * grid font (Windows Excel recomputes it on open: lookwin probe p26 draws a
 * Calibri 11 sheet stating 14.4 pt at 15 pt; R79-X23, every column Verdana 8,
 * stays 10.5 pt). Undefined when a face is one the metrics do not know.
 * Column styles carry their size in px, as the model stores them.
 */
export function officeEmptyRowHeight(
  font: { readonly family: string; readonly sizePt: number; readonly record?: ExcelFontRecordClass | null },
  columnStyles: Readonly<Record<string, { readonly fontFamily?: unknown; readonly fontSize?: unknown; readonly fontRecord?: unknown } | undefined>> | undefined,
): number | undefined {
  const faces = new Map<string, { family: string; sizePt: number; record?: ExcelFontRecordClass | null }>();
  let styled = 0;
  for (const key in columnStyles ?? {}) {
    if (!/^\d+$/.test(key)) continue;
    styled += 1;
    const style = columnStyles?.[key];
    const own = typeof style?.fontFamily === "string" && style.fontFamily.trim() ? style.fontFamily.trim() : undefined;
    const family = own ?? font.family;
    const record = own ? excelFontRecordOf(style?.fontRecord) : font.record;
    const sizePt = typeof style?.fontSize === "number" && style.fontSize > 0 ? style.fontSize * 0.75 : font.sizePt;
    faces.set(`${family}\u0000${sizePt}\u0000${JSON.stringify(record ?? null)}`, { family, sizePt, record });
  }
  if (styled < 16_384) faces.set(`${font.family}\u0000${font.sizePt}\u0000${JSON.stringify(font.record ?? null)}`, { family: font.family, sizePt: font.sizePt, record: font.record });
  for (const face of faces.values()) if (!officeRowMetricsKnown(face.family, face.sizePt, face.record)) return undefined;
  return officeCombinedRowHeight([...faces.values()]);
}

/**
 * One line of a row's empty positions and empty cells: the bottom-aligned
 * faces stand on one baseline, their tallest ascent over their deepest
 * descent; a `standing` (top or centred) face at its own line (lookwin
 * p168-lw: blank or "Hxgy" Arial 12 and Courier New 10 cells fit 21 px both
 * bottom-aligned, 20 -- the taller line -- centred, top or one of each). The
 * API's office_floor_row_height. Undefined for a face the metrics lack.
 */
export function officeFloorRowHeight(fonts: readonly OfficeRowFont[]): number | undefined {
  const bottom = fonts.filter((font) => !font.standing);
  const combined = bottom.length ? officeCombinedRowHeight(bottom) : 0;
  const standing = officeStandingRowHeight(fonts);
  return combined === undefined || standing === undefined ? undefined : Math.max(combined, standing) || undefined;
}

/** The tallest line of the `standing` faces (0 for none), undefined for one
 * the metrics lack. */
export function officeStandingRowHeight(fonts: readonly OfficeRowFont[]): number | undefined {
  let tallest = 0;
  for (const font of fonts) {
    if (!font.standing) continue;
    const height = officeCombinedRowHeight([font]);
    if (height === undefined) return undefined;
    tallest = Math.max(tallest, height);
  }
  return tallest;
}

export function officeCombinedRowHeight(
  fonts: readonly OfficeRowFont[],
): number | undefined {
  const parts = officeCombinedRowParts(fonts);
  return parts ? parts.ascent + parts.descent + 1 : undefined;
}

/** The ascent and descent of one line set in several faces at once
 * ({@link officeCombinedRowHeight}'s halves). */
export function officeCombinedRowParts(
  fonts: readonly OfficeRowFont[],
): { ascent: number; descent: number } | undefined {
  let ascent = 0;
  let descent = 0;
  for (const font of fonts) {
    const parts = officeRowParts(font);
    if (!parts) return undefined;
    ascent = Math.max(ascent, parts.ascent);
    descent = Math.max(descent, parts.descent);
  }
  return fonts.length ? { ascent, descent } : undefined;
}

/**
 * The automatic height of a cell's wrapped lines in its row, px, as Windows
 * Excel stands them (lookwin probes p44b/p44c). `own` is the lines' own
 * height (each line at its faces), `last` the halves of its last line and
 * `floor` the halves of the row's other faces (its empty positions'). A
 * bottom-aligned cell (Excel's default) stands its lines on the row's
 * baseline: the row is the lines' height above that last line's baseline
 * over the deepest descent of them all -- two lines of Calibri 8 (15 px
 * each) in a Calibri 11 workbook are 31 px, not 30. A top, center, justify
 * or distributed cell keeps its lines' height, never less than one line of
 * the row's other faces (R75-X01: two top-aligned lines of Calibri 10 are 34).
 */
export function officeStackedLinesHeight(
  own: number,
  last: { ascent: number; descent: number },
  floor: { ascent: number; descent: number } | undefined,
  baseline: boolean,
): number {
  if (!floor) return own;
  if (!baseline) return Math.max(own, floor.ascent + floor.descent + 1);
  const above = own - last.descent - 1;
  return Math.max(above, floor.ascent) + Math.max(last.descent, floor.descent) + 1;
}

/**
 * Windows Excel's extra automatic height of the rows its heavy lines touch: a
 * thick or double line on a row's top edge adds a pixel, a medium or heavier
 * one on its bottom edge another (the rows Excel marks thickTop / thickBot;
 * 7,400 Windows-written rows, xlsx-rounds rowborder_edges.py: both 3 px +2 in
 * 257 of 260, a 2 px bottom alone +1 in 1,989 of 2,166). An edge is shared: a
 * row's top edge carries its cells' top lines and the row above's bottom lines.
 * `lines` are cells' top and bottom line widths (borderLineWidthPx) by row.
 */
/**
 * The extra automatic height of the rows a file marks thickTop (1) / thickBot
 * (2) (`sheet.rowThickEdges`, row -> marks, the import's reading
 * xlsxRowThickEdges): Windows Excel grows a row a pixel for each mark,
 * whatever lines its cells draw (lookwin p80-lw: a medium bottom on an
 * unmarked row adds nothing, the mark alone a pixel; R105-X05). Null when the
 * sheet carries no marks (a page imported before): officeBorderRowExtras.
 */
export function officeRowThickMarkExtras(marks: unknown): Record<string, number> | undefined | null {
  if (!marks || typeof marks !== "object" || Array.isArray(marks)) return null;
  let extras: Record<string, number> | undefined;
  for (const [row, flags] of Object.entries(marks as Record<string, unknown>)) {
    if (!/^\d+$/.test(row) || typeof flags !== "number" || !Number.isInteger(flags)) continue;
    const extra = (flags & 1) + ((flags >> 1) & 1);
    if (extra) (extras ??= {})[row] = extra;
  }
  return extras;
}

export function officeBorderRowExtras(
  lines: Iterable<{ readonly row: number; readonly top: number; readonly bottom: number }>,
): Record<string, number> | undefined {
  const topEdge = new Map<number, number>();
  const bottomEdge = new Map<number, number>();
  const widen = (edges: Map<number, number>, row: number, width: number) => {
    if (row >= 0 && width > (edges.get(row) ?? 0)) edges.set(row, width);
  };
  for (const { row, top, bottom } of lines) {
    widen(topEdge, row, top);
    widen(bottomEdge, row - 1, top);
    widen(bottomEdge, row, bottom);
    widen(topEdge, row + 1, bottom);
  }
  let extras: Record<string, number> | undefined;
  for (const row of new Set([...topEdge.keys(), ...bottomEdge.keys()])) {
    const extra = Number((topEdge.get(row) ?? 0) >= 3) + Number((bottomEdge.get(row) ?? 0) >= 2);
    if (extra) (extras ??= {})[String(row)] = extra;
  }
  return extras;
}

/**
 * Device pixels of a column whose FILE width (`CT_Col@width`, which already
 * includes the 5-pixel margin) is `width`: §18.3.1.13's
 * `Truncate(((256 * width + Truncate(128 / MDW)) / 256) * MDW)`. For any width
 * Excel wrote this is exactly the pixel count it was written from.
 */
export function columnFileWidthToPixels(
  width: number,
  digitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
): number {
  return Math.floor(((256 * width + Math.floor(128 / digitWidth)) / 256) * digitWidth);
}

/**
 * Device pixels of the default column of a sheet whose file names no
 * `defaultColWidth`, from its `baseColWidth` characters plus the 5-pixel
 * margin. Windows Excel rounds that up to a multiple of 8 pixels, which is
 * why a base of 8 at the 7-pixel digit is 64 pixels, "8.43" characters, and not
 * 61; Excel for Mac does not round (a new Mac workbook's base of 10 at its
 * 6-point digit is 65 points, measured).
 */
export function columnBaseWidthToPixels(
  baseColumnWidth: number,
  digitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
  dpi = 96,
): number {
  const device = baseColumnWidth * digitWidth + COLUMN_WIDTH_PADDING_PX;
  return dpi === 72 ? device : Math.ceil(device / 8) * 8;
}

/** Characters shown for a column `pixels` wide (never below zero). */
export function columnPixelsToCharacters(
  pixels: number,
  digitWidth = DEFAULT_MAX_DIGIT_WIDTH_PX,
  dpi = 96,
): number {
  const device = dpi === 96 ? pixels : (pixels * dpi) / 96;
  return Math.max(0, (device - COLUMN_WIDTH_PADDING_PX) / digitWidth);
}

export function rowPixelsToPoints(pixels: number): number {
  return pxToPt(pixels);
}

export function rowPointsToPixels(points: number): number {
  return ptToPx(points);
}

/** Excel reads sizes to two decimals (8.43 characters, 14.25 points). */
export function roundSizeReading(value: number): number {
  return Math.round(value * 100) / 100;
}

/**
 * The value typed in Column Width (characters, 0 to 255) or Row Height
 * (points, 0 to 409), or null when it is not a number in that range. Zero is a
 * valid entry: Excel hides the line.
 */
export function parseLineSizeEntry(raw: unknown, max: number): number | null {
  const text =
    typeof raw === "number" ? String(raw) : typeof raw === "string" ? raw.trim() : "";
  if (text === "") return null;
  const value = Number(text);
  if (!Number.isFinite(value) || value < 0 || value > max) return null;
  return value;
}

/**
 * The resolved width for a column: its persisted override (clamped) when set,
 * otherwise the sheet default. Negative / non-integer indices resolve to the
 * default. The `defaultWidth` is itself clamped so a wild default never escapes
 * the bounds.
 */
export function resolvedColumnWidth(
  index: number,
  widths: SizeMap | undefined,
  defaultWidth: number,
): number {
  const fallback = clampColumnWidth(defaultWidth, MIN_COLUMN_WIDTH_PX);
  if (!widths || !Number.isInteger(index) || index < 0) return fallback;
  const raw = widths[index];
  if (raw === undefined) return fallback;
  return clampColumnWidth(raw, fallback);
}

/**
 * The resolved height for a row: its persisted override (clamped) when set,
 * otherwise the sheet default (see resolvedColumnWidth).
 */
export function resolvedRowHeight(
  index: number,
  heights: SizeMap | undefined,
  defaultHeight: number,
): number {
  const fallback = clampRowHeight(defaultHeight, MIN_ROW_HEIGHT_PX);
  if (!heights || !Number.isInteger(index) || index < 0) return fallback;
  const raw = heights[index];
  if (raw === undefined) return fallback;
  return clampRowHeight(raw, fallback);
}

/**
 * The cumulative pixel offset of line `index` measured from the first line's
 * leading edge (i.e. the sum of the resolved sizes of every line before it).
 * `offsetOfLine(0, ...)` is always 0; `offsetOfLine(count, ...)` is the total
 * size of the whole axis. Out-of-range indices clamp to [0, count].
 *
 * @param resolveSize resolves the size of a single line by index (already clamped)
 */
export function offsetOfLine(
  index: number,
  count: number,
  resolveSize: (line: number) => number,
): number {
  const safeCount = Math.max(0, Math.floor(count));
  const bound = Math.min(Math.max(0, Math.floor(index)), safeCount);
  const sparse = (resolveSize as Partial<SparseAxisSizeResolver>)[
    SPARSE_AXIS_RESOLVER
  ];
  if (sparse) return sparse.offsetOfLine(bound);
  let offset = 0;
  for (let i = 0; i < bound; i += 1) offset += resolveSize(i);
  return offset;
}

/**
 * The total pixel size of an axis: the sum of every line's resolved size.
 * Equivalent to offsetOfLine(count, count, resolveSize) but named for clarity.
 */
export function totalAxisSize(
  count: number,
  resolveSize: (line: number) => number,
): number {
  return offsetOfLine(Math.max(0, Math.floor(count)), count, resolveSize);
}

/**
 * The index of the first line whose trailing edge is past `scroll` pixels — i.e.
 * the first line at least partially visible at the given scroll offset. Used to
 * derive the virtualization window start from the scroll position when lines can
 * have variable sizes (the fixed-size shortcut `floor(scroll / size)` no longer
 * holds). Clamped to [0, count - 1] (or 0 for an empty axis); a non-positive
 * scroll resolves to line 0.
 */
export function lineIndexAtOffset(
  scroll: number,
  count: number,
  resolveSize: (line: number) => number,
): number {
  const safeCount = Math.max(0, Math.floor(count));
  if (safeCount === 0) return 0;
  const target = Number.isFinite(scroll) ? Math.max(0, scroll) : 0;
  const sparse = (resolveSize as Partial<SparseAxisSizeResolver>)[
    SPARSE_AXIS_RESOLVER
  ];
  if (sparse) {
    let low = 0;
    let high = safeCount - 1;
    while (low < high) {
      const middle = low + Math.floor((high - low) / 2);
      if (sparse.offsetOfLine(middle + 1) > target) high = middle;
      else low = middle + 1;
    }
    return low;
  }
  let acc = 0;
  for (let i = 0; i < safeCount; i += 1) {
    acc += resolveSize(i);
    // Strictly-greater so a scroll landing exactly on a boundary picks the next
    // line (its leading edge is the boundary), matching the fixed-size floor().
    if (acc > target) return i;
  }
  return safeCount - 1;
}

/**
 * Auto-fit a column to its content width: the widest cell text drives the size,
 * clamped to the bounds, with a small fixed padding so text never touches the
 * cell edge. When no cell has text the column resets to its (clamped) default —
 * Excel's double-click on an empty column collapses to a minimal width, but a
 * reset-to-default is the safe, predictable behaviour here. The text width is
 * estimated from a per-character pixel budget (no DOM measuring in this pure
 * layer); the renderer can pass the longest visible string per column.
 */
// A CJK / fullwidth code point occupies ~1.7 ASCII-character widths at the grid
// font (~13px vs the 8px ASCII budget). Slight over-fit beats clipping.
const CJK_DISPLAY_WIDTH = 1.7;

/** True for East-Asian Wide / Fullwidth code points (the chars that render ~2x). */
function isWideCodePoint(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x115f) || // Hangul Jamo
    (cp >= 0x2e80 && cp <= 0x303e) || // CJK radicals, Kangxi, symbols/punct
    (cp >= 0x3041 && cp <= 0x33ff) || // Hiragana, Katakana, CJK symbols
    (cp >= 0x3400 && cp <= 0x4dbf) || // CJK Ext A
    (cp >= 0x4e00 && cp <= 0x9fff) || // CJK Unified Ideographs
    (cp >= 0xa000 && cp <= 0xa4cf) || // Yi
    (cp >= 0xac00 && cp <= 0xd7a3) || // Hangul Syllables
    (cp >= 0xf900 && cp <= 0xfaff) || // CJK Compatibility Ideographs
    (cp >= 0xfe30 && cp <= 0xfe4f) || // CJK Compatibility Forms
    (cp >= 0xff00 && cp <= 0xff60) || // Fullwidth Forms
    (cp >= 0xffe0 && cp <= 0xffe6) || // Fullwidth signs
    (cp >= 0x1f300 && cp <= 0x1faff) || // Emoji & pictographs (wide)
    (cp >= 0x20000 && cp <= 0x3fffd) // CJK Ext B+ (supplementary ideographs)
  );
}

/**
 * Display width of a string in "ASCII-character" units: each East-Asian wide /
 * fullwidth code point counts as {@link CJK_DISPLAY_WIDTH}, every other as 1.
 * Iterates by code POINT (so an astral CJK Ext-B glyph counts once, not twice).
 * For ASCII-only text this equals `text.length`, so non-CJK callers are
 * byte-identical. PURE.
 */
export function eastAsianTextWidth(text: string): number {
  if (typeof text !== "string" || text.length === 0) return 0;
  let width = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    width += isWideCodePoint(cp) ? CJK_DISPLAY_WIDTH : 1;
  }
  return width;
}

export function autoFitColumnWidth(
  longestText: string,
  defaultWidth: number,
  options?: {
    pxPerChar?: number;
    paddingPx?: number;
    /** The widest cell's text measured in its OWN font (the grid's canvas
     * measurement); the per-character estimate below is only a fallback. */
    measure?: (text: string) => number;
  },
): number {
  const fallback = clampColumnWidth(defaultWidth, MIN_COLUMN_WIDTH_PX);
  const text = typeof longestText === "string" ? longestText : "";
  if (text.length === 0) return fallback;
  const pxPerChar = options?.pxPerChar ?? 8;
  const paddingPx = options?.paddingPx ?? 16;
  if (options?.measure) {
    const measured = options.measure(text);
    if (Number.isFinite(measured) && measured > 0) {
      return clampColumnWidth(Math.ceil(measured) + paddingPx, fallback);
    }
  }
  // East-Asian-width-aware: a CJK / fullwidth glyph renders ~1.7x an ASCII one at
  // the grid font, so sizing by the raw code-unit count under-fits (clips) a
  // Chinese/Japanese/Korean column. ASCII-only text is unchanged (factor 1).
  return clampColumnWidth(
    eastAsianTextWidth(text) * pxPerChar + paddingPx,
    fallback,
  );
}

/**
 * The wrapped line count with REAL glyph widths, for the row heights AutoFit
 * writes. The per-character budget below is an estimate and over-counts the
 * lines of proportional text, which left an auto-fitted row taller than its
 * own text; `measure` (the grid's canvas measurement, in the cell's own font)
 * counts the lines the grid actually draws.
 *
 * The wrap mirrors the wrapped-text layer's CSS: lines break at whitespace,
 * the whitespace at a break hangs rather than starting the next line, a word
 * wider than the line breaks anywhere inside it (`overflow-wrap: anywhere`),
 * and each "\n" starts a new paragraph. Always >= 1.
 */
function measuredWrappedLineCount(
  text: string,
  usableWidthPx: number,
  measure: (segment: string) => number,
): number {
  let lines = 0;
  for (const paragraph of text.split("\n")) {
    lines += 1; // every paragraph occupies at least one line
    if (paragraph === "") continue;
    let current = "";
    for (const token of paragraph.match(/\s+|\S+/g) ?? []) {
      let rest = token;
      // Bounded by the token's length: each turn either places the rest of
      // the token or consumes at least one character of it.
      for (let guard = rest.length + 1; rest.length > 0 && guard > 0; guard -= 1) {
        if (measure(current + rest) <= usableWidthPx) {
          current += rest;
          break;
        }
        if (current !== "") {
          // Break here. Whitespace at the break hangs off the line's end.
          lines += 1;
          current = "";
          if (token.trim() === "") {
            rest = "";
            break;
          }
          continue;
        }
        // A word wider than a whole line: fill this line, continue below.
        let cut = 1;
        while (
          cut < rest.length &&
          measure(rest.slice(0, cut + 1)) <= usableWidthPx
        ) {
          cut += 1;
        }
        current = rest.slice(0, cut);
        rest = rest.slice(cut);
        if (rest.length > 0) {
          lines += 1;
          current = "";
        }
      }
    }
  }
  return Math.max(1, lines);
}

/**
 * Estimate how many wrapped display lines a string occupies inside a column of
 * `columnWidthPx`, given a per-character pixel budget. Used by the wrap-text row
 * auto-grow: a cell whose text wraps onto N lines needs ~N line-heights of room.
 *
 * The wrap is a greedy word/char wrap with the SAME `pxPerChar` budget the
 * column-width auto-fit uses, so the row-height estimate and the column-width
 * estimate stay consistent. Explicit "\n" breaks (Excel Alt+Enter / a formula
 * result with a newline) always start a new line. A word longer than the usable
 * width is broken across lines (Excel's `overflow-wrap: anywhere`). An empty
 * string is one line. The result is always >= 1.
 *
 * @param text the cell's displayed text (may contain "\n")
 * @param columnWidthPx the resolved (spanned) column width in px
 * @param options.pxPerChar per-character width budget (default 8, matches auto-fit)
 * @param options.paddingPx leading+trailing cell padding subtracted from the width
 */
export function wrappedLineCount(
  text: string,
  columnWidthPx: number,
  options?: {
    pxPerChar?: number;
    paddingPx?: number;
    measure?: (segment: string) => number;
  },
): number {
  const raw = typeof text === "string" ? text : "";
  if (options?.measure) {
    const paddingPx = Math.max(0, options.paddingPx ?? 18);
    const width = Number.isFinite(columnWidthPx) ? columnWidthPx : 0;
    return measuredWrappedLineCount(
      raw,
      Math.max(1, width - paddingPx),
      options.measure,
    );
  }
  const pxPerChar = Math.max(1, options?.pxPerChar ?? 8);
  const paddingPx = Math.max(0, options?.paddingPx ?? 18);
  const width = Number.isFinite(columnWidthPx) ? columnWidthPx : 0;
  // Usable text width is the column minus padding; floor of one char so a
  // pathologically narrow column still wraps one char per line (never 0 → ∞).
  const usable = Math.max(pxPerChar, width - paddingPx);
  const charsPerLine = Math.max(1, Math.floor(usable / pxPerChar));
  let lines = 0;
  // CJK-aware floor: the word-wrap loop below counts by raw code units, which
  // OVER-fits a CJK column (each glyph is ~1.7 ASCII widths) and under-counts the
  // lines, clipping wrapped Chinese/Japanese text. Track a parallel display-width
  // estimate per paragraph and take the max. `charsPerLine` is already in
  // ASCII-char-width units (usable/pxPerChar), the same units eastAsianTextWidth
  // returns. For ASCII text this estimate is <= the word-wrap count (which only
  // adds lines at word boundaries), so non-CJK results stay byte-identical.
  let displayLines = 0;
  // Split on explicit newlines first: each hard break is its own paragraph that
  // wraps independently. An empty paragraph still occupies one line.
  for (const paragraph of raw.split("\n")) {
    if (paragraph.length === 0) {
      lines += 1;
      displayLines += 1;
      continue;
    }
    displayLines += Math.max(
      1,
      Math.ceil(eastAsianTextWidth(paragraph) / charsPerLine),
    );
    // Greedy word wrap: accumulate words onto the current line until the next
    // word would overflow; a single word wider than the line is hard-split.
    let lineLen = 0;
    const words = paragraph.split(/(\s+)/); // keep whitespace tokens
    for (const token of words) {
      let remaining = token;
      while (remaining.length > 0) {
        const room = charsPerLine - lineLen;
        if (remaining.length <= room) {
          lineLen += remaining.length;
          remaining = "";
        } else if (room > 0 && remaining.length > charsPerLine) {
          // A word longer than a whole line: fill the rest of this line, wrap.
          lineLen = charsPerLine;
          remaining = remaining.slice(room);
          lines += 1;
          lineLen = 0;
        } else {
          // Word fits on a fresh line: wrap then place it (or as much as fits).
          lines += 1;
          lineLen = 0;
          if (remaining.length <= charsPerLine) {
            lineLen = remaining.length;
            remaining = "";
          } else {
            remaining = remaining.slice(charsPerLine);
            lines += 1;
          }
        }
      }
    }
    lines += 1; // the last (or only) line of this paragraph
  }
  return Math.max(1, lines, displayLines);
}

/**
 * The row height needed to show `text` wrapped inside `columnWidthPx`: the
 * wrapped line count times the per-line height plus vertical padding, clamped to
 * the row-height bounds. Mirrors `autoFitColumnWidth` for the row axis. A single
 * line resolves to the (clamped) default height so a one-line wrapped cell keeps
 * the byte-stable default.
 *
 * @param defaultHeight the sheet's default row height (fallback / one-line height)
 * @param options.lineHeightPx height of one wrapped text line (default 18)
 * @param options.paddingPx vertical cell padding added on top of the lines
 */
export function wrappedTextHeightPx(
  text: string,
  columnWidthPx: number,
  defaultHeight: number,
  options?: {
    pxPerChar?: number;
    paddingPx?: number;
    lineHeightPx?: number;
    verticalPaddingPx?: number;
    measure?: (segment: string) => number;
  },
): number {
  const fallback = clampRowHeight(defaultHeight, MIN_ROW_HEIGHT_PX);
  const lines = wrappedLineCount(text, columnWidthPx, {
    pxPerChar: options?.pxPerChar,
    paddingPx: options?.paddingPx,
    measure: options?.measure,
  });
  if (lines <= 1) return fallback;
  const lineHeightPx = Math.max(1, options?.lineHeightPx ?? 18);
  const verticalPaddingPx = Math.max(0, options?.verticalPaddingPx ?? 6);
  // Excel's AutoFit Row Height is the lines' own height plus a small padding
  // (Microsoft Support, "Change the column width and row height"), so the
  // fractional line box is carried, not dropped.
  const needed = Math.ceil(lines * lineHeightPx) + verticalPaddingPx;
  // Never shrink below the default: a wrapped cell is at least one default row.
  return clampRowHeight(Math.max(needed, fallback), fallback);
}

/**
 * The row height AutoFit Row Height gives one cell, as the grid draws its
 * text: wrapped lines, one line in a font taller than the row, a line turned
 * by an angle (the row fits its rotated box) or letters stacked upright. The
 * automatic row heights measure the same shapes on the page; the height
 * AutoFit fixes must not be shorter (live audit S-051: a 90° label fixed at
 * 20px showed one character).
 */
export function autoFitCellHeightPx(
  text: string,
  columnWidthPx: number,
  defaultHeight: number,
  metrics: {
    lineHeightPx: number;
    verticalPaddingPx: number;
    fontSizePx: number;
    measure: (segment: string) => number;
  },
  layout: { wraps: boolean; rotation?: number; stacked?: boolean },
): number {
  const fallback = clampRowHeight(defaultHeight, MIN_ROW_HEIGHT_PX);
  // The automatic row's own allowance around a single measured line.
  const fit = (needed: number) => clampRowHeight(Math.max(fallback, Math.ceil(needed)), fallback);
  if (layout.stacked) {
    // One upright letter per em, down the cell.
    const letters = [...text.replace(/\n/g, "")].length;
    return fit(letters * metrics.fontSizePx + 4);
  }
  const rotation = Number(layout.rotation ?? 0);
  if (rotation) {
    const radians = (rotation * Math.PI) / 180;
    const longest = Math.max(0, ...text.split("\n").map((line) => metrics.measure(line)));
    return fit(Math.abs(Math.sin(radians)) * longest + Math.abs(Math.cos(radians)) * metrics.lineHeightPx + 4);
  }
  if (layout.wraps) {
    return wrappedTextHeightPx(text, columnWidthPx, defaultHeight, metrics);
  }
  return fit(metrics.lineHeightPx + metrics.verticalPaddingPx);
}

/**
 * Produce the next size map after a resize, dropping a back-to-default override
 * so the persisted model stays minimal (a column dragged back to the default
 * width should not keep a redundant entry — important for "double-click =
 * reset" to actually clear the override). Returns `undefined` when the result is
 * empty so the sheet field can be omitted entirely.
 *
 * @param clamp the axis clamp (clampColumnWidth / clampRowHeight)
 * @param defaultSize the sheet default for the axis (used to detect "= default")
 */
export function setLineSize(
  map: SizeMap | undefined,
  index: number,
  size: number,
  defaultSize: number,
  clamp: (value: unknown, fallback: number) => number,
): SizeMap | undefined {
  if (!Number.isInteger(index) || index < 0) return map;
  const fallback = clamp(defaultSize, 1);
  const next: SizeMap = {};
  for (const [key, value] of Object.entries(map || {})) {
    const k = Number(key);
    if (!Number.isInteger(k) || k < 0 || k === index) continue;
    next[k] = clamp(value, fallback);
  }
  const clamped = clamp(size, fallback);
  // Keep the override only when it differs from the default; otherwise the line
  // already renders at the default and the entry is redundant.
  if (clamped !== fallback) next[index] = clamped;
  return Object.keys(next).length > 0 ? next : undefined;
}

/**
 * Set MANY line sizes in one pass (the multi-selection autofit / resize: a whole
 * span of columns or rows updated as a SINGLE model change = one undo step).
 * Mirrors {@link setLineSize}: each requested override is clamped, an override
 * equal to the default is dropped (the line already renders at the default), and
 * a carried entry the overrides replace is dropped. Last write wins per index;
 * invalid indices are ignored. Returns undefined when nothing remains.
 */
export function setLineSizes(
  map: SizeMap | undefined,
  sizes: ReadonlyArray<{ index: number; size: number }>,
  defaultSize: number,
  clamp: (value: unknown, fallback: number) => number,
): SizeMap | undefined {
  const fallback = clamp(defaultSize, 1);
  const requested = new Map<number, number>();
  for (const entry of sizes) {
    const index = entry?.index;
    if (!Number.isInteger(index) || index < 0) continue;
    requested.set(index, clamp(entry.size, fallback));
  }
  const next: SizeMap = {};
  for (const [key, value] of Object.entries(map || {})) {
    const k = Number(key);
    if (!Number.isInteger(k) || k < 0 || requested.has(k)) continue;
    next[k] = clamp(value, fallback);
  }
  for (const [index, clamped] of requested) {
    if (clamped !== fallback) next[index] = clamped;
  }
  return Object.keys(next).length > 0 ? next : undefined;
}

/**
 * The letter spacing that stacks Vertical Text (textRotation 255) a character
 * to a line as Excel does: each character's advance down the cell is the
 * automatic row height Excel gives its face (officeRowHeight), not its em --
 * measured on R20-X14's survey headers, a character's pitch is a row of that
 * font. `fontSizePx` at 96 dpi. The browser's own line box when the face is
 * one the tables do not know.
 */
export function stackedTextLetterSpacing(fontFamily: unknown, fontSizePx: unknown): string {
  const family = typeof fontFamily === "string" ? fontFamily.split(",")[0]?.trim().replace(/^["']|["']$/g, "").trim() : "";
  const size = typeof fontSizePx === "number" ? fontSizePx : Number.parseFloat(String(fontSizePx ?? ""));
  if (!family || !(size > 0) || !officeRowMetricsKnown(family, size * 0.75)) return "calc(1lh - 1em)";
  return `${Math.max(0, officeRowHeight({ family, sizePt: size * 0.75 }) - size)}px`;
}
