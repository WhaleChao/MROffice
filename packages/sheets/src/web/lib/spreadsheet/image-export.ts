/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Spreadsheet image export (PNG) — "输出为图片" — WPS / Excel parity.
 *
 * Two targets:
 *   1. **Chart → PNG**: accepts the chart's `<svg>` element; serializes +
 *      rasters via the shared `lib/export/svg-to-png` pipeline.
 *   2. **Range / sheet → PNG**: renders a rectangular cell region (the user's
 *      selection or the whole used range) to a canvas: gridlines, cell
 *      backgrounds, and text drawn from the cell-model data.
 *
 * Architecture (unit-testability):
 *   - `computeRangeLayout` — PURE: geometry + style data for every cell in the
 *     range; no canvas, no DOM.  Fully unit-tested.
 *   - `drawRangeToCanvas` — thin browser-only wrapper that calls
 *     `computeRangeLayout` then issues canvas 2D API calls.
 *   - `exportChartSvgToPng` — delegates to `exportSvgElementToPng`.
 *   - `exportRangeToPng` — orchestrates the canvas path + download.
 *
 * The layout draws what the caller resolved: the sheet hands it each cell as
 * the grid paints it (its table style, conditional format, automatic General
 * alignment and font), its merges, hidden lines and the drawings over the
 * range, so the picture is the range as shown on screen, as Excel's Copy as
 * Picture is. It stays one file and is unit-tested without a browser.
 */

import {
  DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX,
  DEFAULT_SPREADSHEET_ROW_HEIGHT_PX,
  resolvedColumnWidth,
  resolvedRowHeight,
  type SizeMap,
} from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";
import { MAX_IMAGE_SRC_LENGTH } from "@mrnote/core/web/lib/spreadsheet/sheet-image.ts";
import { pictureObjectDisplaySrc } from "@mrnote/core/web/lib/page-picture-objects.ts";
import {
  effectiveVerticalAlign,
  indentPaddingPx,
  resolveBorderCss,
  SPREADSHEET_WRAP_LINE_HEIGHT,
  type CellBorderStyle,
} from "@mrnote/core/web/lib/spreadsheet/cell-format.ts";
import { normalizeCellRuns } from "@mrnote/core/web/lib/spreadsheet/cell-rich-text.ts";
import type { NumberFormatLayoutPart } from "@mrnote/core/web/lib/spreadsheet/number-format.ts";
import { nativeFontSizeCss } from "@mrnote/core/web/lib/spreadsheet/native-font-size.ts";
import {
  cellPhoneticGuideLayout,
  phoneticRubyFontPx,
  type PhoneticGuideLayout,
} from "@mrnote/core/web/lib/spreadsheet/cell-phonetic.ts";
import {
  SPREADSHEET_SCRIPT_FONT_SCALE,
  cellAccountingUnderline,
  cellHorizontalFill,
  cellResolvedTextDirection,
  cellTextDirection,
  cellUnderlineKind,
  cellVerticalScript,
  horizontalFillText,
  scriptBaselineShiftPx,
  type SpreadsheetUnderlineKind,
} from "@mrnote/core/web/lib/spreadsheet/cell-text-effects.ts";
import { cssFontFamilyList } from "@mrnote/core/web/lib/css-font-family.ts";
import { splitBehind } from "@mrnote/core/web/lib/regexp-behind.ts";

// Pieces that each end after white space: `split(/(?<=\s)/)` without the
// lookbehind, which Safari before 16.4 cannot compile.
const AFTER_SPACE = { char: /\s/ } as const;

// ---------------------------------------------------------------------------
// Inlined PNG-download utilities (mirrors lib/export/svg-to-png helpers;
// inlined here so this module has zero dependency on lib/export/svg-to-png and
// can be unit-tested in isolation — the test suite cannot follow imports that
// do not exist on the running checkout's disk at test time).
// ---------------------------------------------------------------------------

/** Derive a safe PNG filename from a label, same rules as svg-to-png. */
function _toPngFilename(label: string | undefined, fallback = "image"): string {
  const cleaned = (label ?? "")
    .replace(/\.png$/i, "")
    .replace(/[\\/:*?"<>| -]/g, " ")
    .replace(/\s+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120);
  const stem = cleaned.length > 0 ? cleaned : fallback;
  return `${stem}.png`;
}

/** Trigger a browser download of a PNG Blob via a transient <a download>. */
function _downloadPngBlob(filename: string, blob: Blob): boolean {
  if (
    typeof document === "undefined" ||
    typeof URL === "undefined" ||
    typeof URL.createObjectURL !== "function"
  ) {
    return false;
  }
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.rel = "noopener";
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
  return true;
}

// Options type for SVG export — mirrors SvgToPngOptions in svg-to-png.ts.
export interface SvgToPngOptions {
  width?: number;
  height?: number;
  scale?: number;
  background?: string | null;
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** One piece of in-cell rich text as the picture draws it; an absent field is
 *  the cell's own. */
export interface ExportTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  /** Px, as the cell's own fontSize. */
  fontSize?: number;
  fontFamily?: string;
  /** The 2 of H2O: raised or lowered off the line, smaller (B399). */
  superscript?: boolean;
  subscript?: boolean;
}

/** Minimal cell model — mirrors the SpreadsheetCell subset we need. */
export interface ExportCell {
  /** Display text (already resolved: formula result / formatted value). */
  text: string;
  /** The cell's rich text when the grid draws it (B338): the runs concatenate
   *  to `text`. */
  runs?: ExportTextRun[];
  style?: {
    bold?: boolean;
    italic?: boolean;
    color?: string;
    backgroundColor?: string;
    textAlign?: "left" | "center" | "right";
    /** Top / Middle / Bottom; Bottom, Excel's default, when absent. */
    verticalAlign?: "top" | "middle" | "bottom";
    /** The cell's own face and size — an exported picture must draw the cells
     *  in their fonts, as Excel/WPS do, not in one hardcoded UI font. */
    fontFamily?: string;
    fontSize?: number;
    /** The CSS font-family list the grid paints the cell in (its face, else
     *  the theme or workbook default, then the app's fallbacks). Drawn in
     *  place of `fontFamily`, so the picture and the screen pick one face. */
    fontStack?: string;
    /** The edges a table style lays under the cell's own borders (the
     *  header's and the total row's rules): a cell's own edge wins. */
    tableBorderEdges?: CellBorderStyle["borderEdges"];
    /** Wrapped text breaks at the cell's width in the picture as on the grid. */
    wrapText?: boolean;
    /** The cell's borders, legacy box flags and rich edges alike, drawn in the
     *  picture over the gridlines (live audit B297). */
    border?: boolean;
    borderTop?: boolean;
    borderRight?: boolean;
    borderBottom?: boolean;
    borderLeft?: boolean;
    borderEdges?: CellBorderStyle["borderEdges"];
    /** The cell's own text effects, drawn as the grid draws them (live audit
     *  B397-B400, lib/spreadsheet/cell-text-effects): its underline kind, a
     *  whole-cell script, Fill alignment and its reading order. */
    underline?: boolean;
    underlineStyle?: "double" | "singleAccounting" | "doubleAccounting";
    superscript?: boolean;
    subscript?: boolean;
    horizontalFill?: boolean;
    centerAcross?: boolean;
    readingOrder?: "ltr" | "rtl";
    /** Format Cells > Alignment Justify / Distributed (XA, drawn in the
     *  picture since XD2): horizontal, vertical, and Justify distributed. */
    horizontalJustify?: "justify" | "distributed";
    verticalJustify?: "justify" | "distributed";
    justifyLastLine?: boolean;
    indent?: number;
  };
  /** The picture the cell shows as its value (Place in Cell, or a formula
   *  returning one), drawn in place of its `#VALUE!` fallback text. */
  picture?: ExportCellPicture;
  /** Insert > Checkbox: the box the grid draws in place of the cell's
   *  TRUE/FALSE, checked or not (user audit S-117). */
  checkbox?: { checked: boolean };
  /** The phonetic guide the cell shows (Show Phonetic Field), the grid's own
   *  layout of it (live audit B388). */
  phonetic?: PhoneticGuideLayout;
  /** A number format that fills (`* ` in the accounting formats): what comes
   *  before the fill sits at the cell's left edge, what follows it at the
   *  right, as the grid lays it out (live audit S-035). */
  numberLayout?: readonly NumberFormatLayoutPart[];
}

/** An in-cell picture as the export draws it: fitted in the cell keeping its
 *  aspect ratio, placed along `position` (CSS object-position keywords, from
 *  {@link cellPictureObjectPosition}), exactly as the grid draws it. */
export interface ExportCellPicture {
  src: string;
  position?: string;
}

/** One drawn border edge: a solid or doubled band of `width` px in `color`. */
export interface BorderEdgeLine {
  width: number;
  color: string;
  double: boolean;
}

/** Resolved geometry + paint data for a single cell in the export layout. */
export interface CellLayout {
  /** Pixel offset from the left edge of the canvas (after the row-header). */
  x: number;
  /** Pixel offset from the top edge of the canvas (after the col-header). */
  y: number;
  width: number;
  height: number;
  text: string;
  bold: boolean;
  italic: boolean;
  /** Resolved px size for this cell's text (EXPORT_FONT_SIZE_PX when unset). */
  fontSize: number;
  /** Resolved family stack, already quoted/sanitized; the default when unset. */
  fontFamily: string;
  textColor: string;
  /** `null` = transparent (use canvas background). */
  backgroundColor: string | null;
  textAlign: "left" | "center" | "right";
  /** Where the text block stands in the cell's height (Bottom by default). */
  verticalAlign: "top" | "middle" | "bottom";
  /** The text's hard lines (a CHAR(10) break is one line per part); a wrapped
   *  cell breaks each further at draw time (live audit B280). */
  lines: string[];
  /** Rich text, each run's look resolved against the cell's (B338). */
  runs?: LayoutTextRun[];
  wrapText: boolean;
  borders: {
    top?: BorderEdgeLine;
    right?: BorderEdgeLine;
    bottom?: BorderEdgeLine;
    left?: BorderEdgeLine;
  };
  /** The cell's picture value, drawn fitted in the cell instead of text. */
  picture?: ExportCellPicture;
  /** A check box, drawn centred in the cell instead of its TRUE/FALSE. */
  checkbox?: { checked: boolean };
  /** The shown phonetic guide, drawn in the band over the text, which then
   *  sits in the rest of the cell. */
  phonetic?: LayoutPhoneticGuide;
  /** The cell's underline: under the characters, or for the accounting kinds
   *  a rule across the whole cell under the last line. */
  underline?: SpreadsheetUnderlineKind;
  /** A whole-cell script: the text drawn smaller, raised or lowered. */
  script?: "superscript" | "subscript";
  /** Fill alignment: the text repeated as many whole times as fit. */
  fill?: true;
  /** A number format's fill: parts before it at the left, after it at the
   *  right, the fill character repeated between (ExportCell.numberLayout). */
  numberLayout?: readonly NumberFormatLayoutPart[];
  /** The base direction the text is laid out in, set when the cell has one
   *  of its own or its text is right to left. */
  direction?: "ltr" | "rtl";
  /** Horizontal Justify / Distributed: the text wraps and its lines are
   *  spread across the cell (Distributed: every line, with the indent on
   *  both sides; Justify: all but a paragraph's last unless
   *  `justifyLastLine`). */
  horizontalSpread?: "justify" | "distributed";
  justifyLastLine?: boolean;
  /** Indent in px, on both sides for Distributed. */
  spreadIndent?: number;
  /** Vertical Justify (lines from the top edge to the bottom one, a single
   *  line at the top) / Distributed (spread evenly, a single line centred). */
  verticalSpread?: "justify" | "distributed";
}

/** A cell's phonetic guide with its font resolved for the canvas. */
export interface LayoutPhoneticGuide {
  segments: PhoneticGuideLayout["segments"];
  alignment: PhoneticGuideLayout["alignment"];
  bandHeight: number;
  /** Canvas font shorthand of the readings. */
  font: string;
  color: string;
}

/** A run of a cell's rich text with every field resolved. */
export interface LayoutTextRun {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  fontSize: number;
  fontFamily: string;
  color: string;
  /** How far a raised (negative) or lowered run's baseline moves, px; its
   *  `fontSize` is already the script's smaller size. */
  baselineShift?: number;
}

/** A drawing over the range (a chart, picture or shape), as a picture of
 *  itself placed at its frame. */
export interface RangePictureDrawing {
  /** The drawing rendered as an image (a PNG data URL). */
  src: string;
  /** The frame, in CSS px from the range's top-left corner as the sheet runs
   *  left to right; it may stand partly outside the range, and the picture
   *  shows only the part over the range, as Excel's does. */
  left: number;
  top: number;
  width: number;
  height: number;
}

/** A drawing placed on the picture's canvas. */
export interface DrawingLayout {
  src: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Pure layout result for a range of cells. */
export interface RangeLayout {
  /** Total canvas width in CSS pixels. */
  canvasWidth: number;
  /** Total canvas height in CSS pixels. */
  canvasHeight: number;
  cells: CellLayout[];
  /** The cells' left/right text inset, px ({@link EXPORT_CELL_PADDING_X}
   *  when absent). */
  cellPaddingX?: number;
  /** The drawings over the range, bottom first, drawn over the cells. */
  drawings?: DrawingLayout[];
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Cell width (px) of a column with no override when the caller gives no sheet
 * default: a new sheet's grid default, so the picture matches the grid. */
export const EXPORT_DEFAULT_COL_WIDTH = DEFAULT_SPREADSHEET_COLUMN_WIDTH_PX;
/** Cell height (px) of a row with no override when the caller gives no sheet
 * default: a new sheet's grid default. */
export const EXPORT_DEFAULT_ROW_HEIGHT = DEFAULT_SPREADSHEET_ROW_HEIGHT_PX;
/** Grid-line colour used in the canvas export. */
export const EXPORT_GRID_LINE_COLOR = "#d0d0d0";
/** Default text colour when the cell carries none. */
export const EXPORT_DEFAULT_TEXT_COLOR = "#1a1a1a";
/** Font family used when drawing text on the export canvas. Canvas cannot
 *  resolve CSS var(), so the CJK coverage is spelled out. */
export const EXPORT_FONT_FAMILY =
  'system-ui, "PingFang SC", "Noto Sans SC", "Microsoft YaHei", sans-serif';

/** A cell's face, quoted and sanitized, backed by the default stack — the same
 *  two rules the grid renderer applies (quote + fallback), so the picture and
 *  the screen agree on what an imported face falls back to. */
export function exportFontStack(face: unknown): string {
  const trimmed = typeof face === "string" ? face.trim() : "";
  const safe = trimmed.replace(/["'\\;{}()]/g, "").slice(0, 200).trim();
  return safe ? `"${safe}", ${EXPORT_FONT_FAMILY}` : EXPORT_FONT_FAMILY;
}

/** A CSS font-family list (the grid's) as the canvas can take it: each name
 *  kept as written, a token the canvas cannot resolve (a `var()`) or one that
 *  could break out of the font shorthand dropped; the default stack when
 *  nothing is left. */
export function exportCssFontStack(stack: unknown): string {
  if (typeof stack !== "string") return EXPORT_FONT_FAMILY;
  // Split at the commas between names, not those inside a quoted name.
  const names: string[] = [];
  let current = "";
  let quote: string | null = null;
  for (const ch of stack.slice(0, 2_000)) {
    if (quote) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ",") {
      names.push(current.trim());
      current = "";
      continue;
    }
    current += ch;
  }
  if (!quote) names.push(current.trim());
  const kept = names.filter((name) => name !== "" && !/[\\;{}()<>]/.test(name));
  return kept.length > 0 ? kept.join(", ") : EXPORT_FONT_FAMILY;
}
/** Base font size (px) for cell text. */
export const EXPORT_FONT_SIZE_PX = 12;
/** Horizontal text padding inside each cell. */
export const EXPORT_CELL_PADDING_X = 4;
/** Top / bottom inset of a vertically spread text block (XA). */
const EXPORT_CELL_PADDING_Y = 2;
/** Browser raster guardrails for a pasted range picture. The pixel-area cap
 * bounds the backing RGBA allocation (~64 MiB before PNG encoding), while the
 * edge cap stays within the common cross-browser canvas limit. */
export const MAX_RANGE_PICTURE_CANVAS_EDGE = 16_384;
export const MAX_RANGE_PICTURE_CANVAS_PIXELS = 16_777_216;

// ---------------------------------------------------------------------------
// Pure geometry + layout computation (unit-testable without a browser)
// ---------------------------------------------------------------------------

/** Input describing a rectangular cell region to export. */
export interface RangeLayoutInput {
  /** Inclusive row/column bounds. */
  r1: number;
  r2: number;
  c1: number;
  c2: number;
  /**
   * Flat map `"row:col" → ExportCell`.  Missing entries are treated as blank
   * cells (white background, empty text).
   */
  cells: Record<string, ExportCell | undefined>;
  /**
   * Per-column width overrides from the sheet model.  Absent columns use
   * `defaultColumnWidth`.
   */
  columnWidths?: SizeMap;
  /**
   * Per-row height overrides from the sheet model.  Absent rows use
   * `defaultRowHeight`.
   */
  rowHeights?: SizeMap;
  /** The sheet's own width of a column without an override (an imported or
   * customised default), {@link EXPORT_DEFAULT_COL_WIDTH} when absent. */
  defaultColumnWidth?: number;
  /** The sheet's own height of a row without an override,
   * {@link EXPORT_DEFAULT_ROW_HEIGHT} when absent. */
  defaultRowHeight?: number;
  /**
   * The sheet's merged ranges. A merge's anchor is drawn as one cell over the
   * part of the merge inside the range (a centred title across two columns is
   * centred across both, with no gridline between them), and the cells it
   * covers are not drawn at all.
   */
  merges?: ReadonlyArray<{ row: number; col: number; rowSpan: number; colSpan: number }>;
  /** Rows and columns the sheet hides (filters, outlines, manual hiding): they
   *  are not in the picture, as they are not on the grid (live audit B279). */
  hiddenRows?: ReadonlySet<number>;
  hiddenColumns?: ReadonlySet<number>;
  /** The sheet runs right to left (sheetView rightToLeft): the picture keeps
   *  the screen's column order, the range's first column on the right. The
   *  text itself is never mirrored. */
  rightToLeft?: boolean;
  /** The cells' left/right text inset in px: the grid's own, so text stands
   *  and wraps where it does on screen ({@link EXPORT_CELL_PADDING_X} when
   *  absent). */
  cellPaddingX?: number;
  /** The drawings over the sheet (charts, pictures, shapes), bottom first,
   *  each as a picture of itself at its frame: the ones over the range are
   *  drawn over its cells and cut at its edges, as Excel's Copy as Picture
   *  of a range shows the objects on it. */
  drawings?: ReadonlyArray<RangePictureDrawing>;
}

/** The border edges of a cell as the picture draws them: its own, over the
 *  edges a table style lays under it. */
function exportBorders(style: ExportCell["style"]): CellLayout["borders"] {
  if (!style) return {};
  const parse = (value: string | undefined): BorderEdgeLine | undefined => {
    const m = value && /^(\d+(?:\.\d+)?)px\s+(\S+)\s+(.+)$/.exec(value);
    if (!m) return undefined;
    return { width: Number(m[1]), color: m[3], double: m[2] === "double" };
  };
  const edgesOf = (css: ReturnType<typeof resolveBorderCss>): CellLayout["borders"] => {
    const box = parse(css.border);
    return {
      top: parse(css.borderTop) ?? box,
      right: parse(css.borderRight) ?? box,
      bottom: parse(css.borderBottom) ?? box,
      left: parse(css.borderLeft) ?? box,
    };
  };
  const own = edgesOf(resolveBorderCss(style));
  const table = style.tableBorderEdges
    ? edgesOf(resolveBorderCss({ borderEdges: style.tableBorderEdges }))
    : {};
  const out: CellLayout["borders"] = {};
  const top = own.top ?? table.top;
  const right = own.right ?? table.right;
  const bottom = own.bottom ?? table.bottom;
  const left = own.left ?? table.left;
  if (top) out.top = top;
  if (right) out.right = right;
  if (bottom) out.bottom = bottom;
  if (left) out.left = left;
  return out;
}

/**
 * Compute the full pixel layout for a rectangular cell range without any
 * canvas / DOM interaction.
 *
 * The result is a flat list of `CellLayout` objects (one per cell in row-major
 * order) plus the total canvas size.
 */
export function computeRangeLayout(input: RangeLayoutInput): RangeLayout {
  const { r1, r2, c1, c2, cells, columnWidths, rowHeights } = input;
  const positive = (value: number | undefined, fallback: number) =>
    typeof value === "number" && Number.isFinite(value) && value > 0 ? value : fallback;
  const defaultColumnWidth = positive(input.defaultColumnWidth, EXPORT_DEFAULT_COL_WIDTH);
  const defaultRowHeight = positive(input.defaultRowHeight, EXPORT_DEFAULT_ROW_HEIGHT);

  // Cumulative column offsets within the export (relative to canvas x=0).
  const colWidthOf = (col: number) =>
    resolvedColumnWidth(col, columnWidths, defaultColumnWidth);
  const rowHeightOf = (row: number) =>
    resolvedRowHeight(row, rowHeights, defaultRowHeight);

  const rowVisible = (row: number) => !input.hiddenRows?.has(row);
  const colVisible = (col: number) => !input.hiddenColumns?.has(col);
  // A hidden line has no extent in the picture, exactly as on the grid.
  const visibleColWidth = (col: number) => (colVisible(col) ? colWidthOf(col) : 0);
  const visibleRowHeight = (row: number) => (rowVisible(row) ? rowHeightOf(row) : 0);

  // Build cumulative column x-offsets for [c1..c2].
  const colOffsets: number[] = [];
  let xCursor = 0;
  for (let c = c1; c <= c2; c++) {
    colOffsets.push(xCursor);
    xCursor += visibleColWidth(c);
  }
  const totalWidth = xCursor;

  // Build cumulative row y-offsets for [r1..r2].
  const rowOffsets: number[] = [];
  let yCursor = 0;
  for (let r = r1; r <= r2; r++) {
    rowOffsets.push(yCursor);
    yCursor += visibleRowHeight(r);
  }
  const totalHeight = yCursor;

  // Each merge clipped to the range: its first cell inside the range draws
  // the whole clipped span and the rest are covered.
  const spanAt = new Map<string, { rows: number; cols: number }>();
  const covered = new Set<string>();
  for (const merge of input.merges ?? []) {
    const mr1 = Math.max(r1, merge.row);
    const mc1 = Math.max(c1, merge.col);
    const mr2 = Math.min(r2, merge.row + merge.rowSpan - 1);
    const mc2 = Math.min(c2, merge.col + merge.colSpan - 1);
    if (mr1 > mr2 || mc1 > mc2 || (mr1 === mr2 && mc1 === mc2)) continue;
    spanAt.set(`${mr1}:${mc1}`, { rows: mr2 - mr1 + 1, cols: mc2 - mc1 + 1 });
    for (let r = mr1; r <= mr2; r++) {
      for (let c = mc1; c <= mc2; c++) {
        if (r !== mr1 || c !== mc1) covered.add(`${r}:${c}`);
      }
    }
  }
  const cellLayouts: CellLayout[] = [];
  for (let ri = 0; ri <= r2 - r1; ri++) {
    for (let ci = 0; ci <= c2 - c1; ci++) {
      const row = r1 + ri;
      const col = c1 + ci;
      const key = `${row}:${col}`;
      if (covered.has(key)) continue;
      if (!rowVisible(row) || !colVisible(col)) continue;
      const span = spanAt.get(key);
      let width = colWidthOf(col);
      let height = rowHeightOf(row);
      for (let extra = 1; span && extra < span.cols; extra++) width += visibleColWidth(col + extra);
      for (let extra = 1; span && extra < span.rows; extra++) height += visibleRowHeight(row + extra);
      const cell = cells[key];
      const style = cell?.style ?? {};
      // A picture value draws the picture, never its `#VALUE!` fallback.
      const picture = cell?.picture?.src ? cell.picture : undefined;
      const checkbox = cell?.checkbox ? { checked: cell.checkbox.checked === true } : undefined;
      const text = picture || checkbox ? "" : (cell?.text ?? "");
      const fontSize =
        typeof style.fontSize === "number" &&
        Number.isFinite(style.fontSize) &&
        style.fontSize > 0
          ? style.fontSize
          : EXPORT_FONT_SIZE_PX;
      // The grid's own font list when the sheet hands it, else the face.
      const fontFamily = style.fontStack
        ? exportCssFontStack(style.fontStack)
        : exportFontStack(style.fontFamily);
      const textColor = style.color || EXPORT_DEFAULT_TEXT_COLOR;
      const script = cellVerticalScript(style);
      const runs =
        cell?.runs && cell.runs.map((run) => run.text).join("") === text
          ? cell.runs.map((run): LayoutTextRun => {
              const size =
                typeof run.fontSize === "number" && Number.isFinite(run.fontSize) && run.fontSize > 0
                  ? run.fontSize
                  : fontSize;
              // The run's own script, else the cell's (as the grid's
              // partsWithCellScript), at the script's smaller size.
              const runScript = run.superscript ? "superscript" : run.subscript ? "subscript" : script;
              return {
                text: run.text,
                bold: run.bold ?? Boolean(style.bold),
                italic: run.italic ?? Boolean(style.italic),
                underline: Boolean(run.underline),
                strike: Boolean(run.strike),
                fontSize: runScript ? size * SPREADSHEET_SCRIPT_FONT_SCALE : size,
                fontFamily: run.fontFamily ? exportFontStack(run.fontFamily) : fontFamily,
                color: run.color || textColor,
                ...(runScript ? { baselineShift: scriptBaselineShiftPx(runScript, size) } : {}),
              };
            })
          : undefined;
      // Center Across Selection keeps the text's own underline, as the grid
      // draws it there (no rule across the cell).
      const underline = cellUnderlineKind(style);
      const accounting = style.centerAcross ? cellAccountingUnderline(style) : null;
      const shownUnderline = accounting ?? underline;
      const direction =
        text && cellTextDirection(style, text) ? cellResolvedTextDirection(style, text) : undefined;
      const guide = !picture && text ? cell?.phonetic : undefined;
      const phonetic: LayoutPhoneticGuide | undefined = guide
        ? {
            segments: guide.segments.map((segment) => ({ ...segment })),
            alignment: guide.alignment,
            bandHeight: guide.bandHeightPx,
            font: [
              guide.font.italic ? "italic" : "",
              guide.font.bold ? "bold" : "",
              `${phoneticRubyFontPx(guide.font, style.fontSize, fontSize)}px`,
              guide.font.family ? exportFontStack(guide.font.family) : fontFamily,
            ]
              .filter(Boolean)
              .join(" "),
            color: guide.font.color ?? textColor,
          }
        : undefined;
      cellLayouts.push({
        // Right to left, the range's first column is on the picture's right.
        x: input.rightToLeft ? totalWidth - colOffsets[ci] - width : colOffsets[ci],
        y: rowOffsets[ri],
        width,
        height,
        text,
        lines: text.split(/\r?\n/),
        wrapText: Boolean(style.wrapText),
        ...(runs ? { runs } : {}),
        borders: exportBorders(style),
        bold: Boolean(style.bold),
        italic: Boolean(style.italic),
        fontSize,
        fontFamily,
        textColor,
        backgroundColor: style.backgroundColor || null,
        textAlign: style.textAlign ?? "left",
        verticalAlign: effectiveVerticalAlign(style),
        ...(picture ? { picture: { ...picture } } : {}),
        ...(checkbox ? { checkbox } : {}),
        ...(phonetic ? { phonetic } : {}),
        ...(text && shownUnderline ? { underline: shownUnderline } : {}),
        ...(text && script && !runs ? { script } : {}),
        ...(text && !runs && cellHorizontalFill(style) ? { fill: true as const } : {}),
        ...(text && !runs && !style.wrapText && cell?.numberLayout && !text.includes("\n")
          ? { numberLayout: cell.numberLayout }
          : {}),
        ...(direction ? { direction } : {}),
        ...(text && !runs && (style.horizontalJustify === "justify" || style.horizontalJustify === "distributed")
          ? {
              horizontalSpread: style.horizontalJustify,
              ...(style.justifyLastLine ? { justifyLastLine: true } : {}),
              ...(style.horizontalJustify === "distributed" && style.indent
                ? { spreadIndent: indentPaddingPx(style.indent) }
                : {}),
            }
          : {}),
        ...(text && !runs && (style.verticalJustify === "justify" || style.verticalJustify === "distributed")
          ? { verticalSpread: style.verticalJustify }
          : {}),
      });
    }
  }

  // The drawings over the range, at their frames from its corner (mirrored
  // with the columns on a right-to-left sheet); one wholly outside it is not
  // in the picture, and the canvas's edges cut the ones that cross them.
  const drawings: DrawingLayout[] = [];
  for (const drawing of input.drawings ?? []) {
    const { src, left, top, width, height } = drawing;
    if (
      typeof src !== "string" ||
      src === "" ||
      ![left, top, width, height].every(Number.isFinite) ||
      !(width > 0) ||
      !(height > 0) ||
      left >= totalWidth ||
      top >= totalHeight ||
      left + width <= 0 ||
      top + height <= 0
    ) {
      continue;
    }
    drawings.push({
      src,
      x: input.rightToLeft ? totalWidth - left - width : left,
      y: top,
      width,
      height,
    });
  }
  const cellPaddingX = input.cellPaddingX;

  return {
    canvasWidth: Math.max(1, totalWidth),
    canvasHeight: Math.max(1, totalHeight),
    cells: cellLayouts,
    ...(typeof cellPaddingX === "number" && Number.isFinite(cellPaddingX) && cellPaddingX >= 0
      ? { cellPaddingX }
      : {}),
    ...(drawings.length > 0 ? { drawings } : {}),
  };
}

/**
 * Build an `ExportCell` map from the raw cell model data and pre-resolved
 * display strings.  The caller provides the display string for each cell
 * (already formatted — e.g. via `resolveCellDisplay`).
 *
 * This is also PURE and unit-testable: the caller supplies all resolved values.
 */
export function buildExportCellMap(
  cells: Record<
    string,
    | {
        display?: string;
        value?: string;
        style?: {
          bold?: boolean;
          italic?: boolean;
          color?: string;
          backgroundColor?: string;
          textAlign?: "left" | "center" | "right";
          verticalAlign?: "top" | "middle" | "bottom";
          fontFamily?: string;
          fontSize?: number;
          fontStack?: string;
          tableBorderEdges?: CellBorderStyle["borderEdges"];
          wrapText?: boolean;
          underline?: boolean;
          underlineStyle?: "double" | "singleAccounting" | "doubleAccounting";
          superscript?: boolean;
          subscript?: boolean;
          horizontalFill?: boolean;
          centerAcross?: boolean;
          readingOrder?: "ltr" | "rtl";
          horizontalJustify?: "justify" | "distributed";
          verticalJustify?: "justify" | "distributed";
          justifyLastLine?: boolean;
          indent?: number;
          border?: boolean;
          borderTop?: boolean;
          borderRight?: boolean;
          borderBottom?: boolean;
          borderLeft?: boolean;
          borderEdges?: CellBorderStyle["borderEdges"];
        };
        runs?: unknown;
        phonetic?: unknown;
        kind?: unknown;
        formula?: unknown;
      }
    | undefined
  >,
  /** Pre-resolved display text keyed by "row:col". */
  displayText: Record<string, string>,
  /** The picture each cell shows as its value, keyed by "row:col" (the
   *  grid's own resolution: placed in the cell or returned by a formula). */
  pictures: Readonly<Record<string, ExportCellPicture | null | undefined>> = {},
  /** The number-format layout the grid draws a cell with, keyed by "row:col";
   *  kept only while it is the text shown (not its #### stand-in). */
  numberLayouts: Readonly<Record<string, readonly NumberFormatLayoutPart[] | undefined>> = {},
  /** The check boxes the grid draws, keyed by "row:col": whether each is
   *  checked (its computed value, as the grid reads it). */
  checkboxes: Readonly<Record<string, boolean | undefined>> = {},
): Record<string, ExportCell> {
  const out: Record<string, ExportCell> = {};
  // Collect all keys from both sources.
  const allKeys = new Set([
    ...Object.keys(cells),
    ...Object.keys(displayText),
    ...Object.keys(pictures),
  ]);
  for (const key of allKeys) {
    const cell = cells[key];
    const picture = pictures[key];
    const text = picture ? "" : (displayText[key] ?? cell?.display ?? cell?.value ?? "");
    // The grid draws a cell's rich text while its value is what it shows.
    const runs = cell?.runs ? normalizeCellRuns(cell.runs, cell.value ?? "") : undefined;
    // The guide reads over the cell's own text, so it is drawn while that
    // text is what the cell shows, as on the grid.
    const phonetic =
      !picture && cell?.value === text
        ? cellPhoneticGuideLayout(cell, cell.style?.fontFamily)
        : null;
    const layoutParts = numberLayouts[key];
    const numberLayout =
      !picture &&
      layoutParts &&
      layoutParts.filter((part) => part.kind === "fill").length === 1 &&
      layoutParts.map((part) => (part.kind === "text" ? part.text : "")).join("").replace(/\s/g, "") ===
        text.replace(/\s/g, "")
        ? layoutParts
        : undefined;
    const checked = picture ? undefined : checkboxes[key];
    out[key] = {
      ...(picture ? { picture: { ...picture } } : {}),
      ...(typeof checked === "boolean" ? { checkbox: { checked } } : {}),
      ...(phonetic ? { phonetic } : {}),
      ...(numberLayout ? { numberLayout: numberLayout.map((part) => ({ ...part })) } : {}),
      text,
      ...(runs && cell?.value === text
        ? {
            runs: runs.map((run): ExportTextRun => {
              const style = run.style ?? {};
              const exportRun: ExportTextRun = { text: run.text };
              if (typeof style.bold === "boolean") exportRun.bold = style.bold;
              if (typeof style.italic === "boolean") exportRun.italic = style.italic;
              if (style.underline) exportRun.underline = true;
              if (style.strike) exportRun.strike = true;
              if (typeof style.color === "string" && style.color) exportRun.color = style.color;
              const size = nativeFontSizeCss(style.fontSize, style.nativeFontSize);
              if (typeof size === "number" && Number.isFinite(size)) exportRun.fontSize = size;
              if (typeof style.fontFamily === "string" && style.fontFamily) {
                exportRun.fontFamily = style.fontFamily;
              }
              if (style.superscript) exportRun.superscript = true;
              else if (style.subscript) exportRun.subscript = true;
              return exportRun;
            }),
          }
        : {}),
      style: cell?.style
        ? {
            bold: cell.style.bold,
            italic: cell.style.italic,
            fontFamily: cell.style.fontFamily,
            fontSize: cell.style.fontSize,
            color: cell.style.color,
            backgroundColor: cell.style.backgroundColor,
            textAlign: cell.style.textAlign,
            verticalAlign: cell.style.verticalAlign,
            fontStack: cell.style.fontStack,
            tableBorderEdges: cell.style.tableBorderEdges,
            wrapText: cell.style.wrapText,
            underline: cell.style.underline,
            underlineStyle: cell.style.underlineStyle,
            superscript: cell.style.superscript,
            subscript: cell.style.subscript,
            horizontalFill: cell.style.horizontalFill,
            centerAcross: cell.style.centerAcross,
            readingOrder: cell.style.readingOrder,
            horizontalJustify: cell.style.horizontalJustify,
            verticalJustify: cell.style.verticalJustify,
            justifyLastLine: cell.style.justifyLastLine,
            indent: cell.style.indent,
            border: cell.style.border,
            borderTop: cell.style.borderTop,
            borderRight: cell.style.borderRight,
            borderBottom: cell.style.borderBottom,
            borderLeft: cell.style.borderLeft,
            borderEdges: cell.style.borderEdges,
          }
        : undefined,
    };
  }
  return out;
}

// ---------------------------------------------------------------------------
// Canvas draw (browser-only; structured as a separate function for mocking)
// ---------------------------------------------------------------------------

/**
 * Interface for a 2D canvas context — the subset we use — so tests can supply
 * a lightweight mock without needing a real canvas.
 *
 * `fillStyle` / `strokeStyle` accept the full CanvasRenderingContext2D union
 * (`string | CanvasGradient | CanvasPattern`) so a real context is assignable
 * here without a cast. Tests only ever assign plain strings, which is allowed
 * since `string` is assignable to the union.
 */
export interface CanvasCtx2D {
  fillStyle: string | CanvasGradient | CanvasPattern;
  strokeStyle: string | CanvasGradient | CanvasPattern;
  lineWidth: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  /** Optional: the base direction text is laid out in (a cell's reading
   *  order). */
  direction?: CanvasDirection;
  fillRect(x: number, y: number, w: number, h: number): void;
  strokeRect(x: number, y: number, w: number, h: number): void;
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  /** Optional: a wrapped cell breaks its lines by measured width. */
  measureText?(text: string): { width: number };
  /** Optional: draws the cells' picture values and the drawings. */
  drawImage?(image: CanvasImageSource, dx: number, dy: number, dw: number, dh: number): void;
  /** Optional: keep a cell's text inside the cell, as the grid clips it. */
  save?(): void;
  restore?(): void;
  beginPath?(): void;
  rect?(x: number, y: number, w: number, h: number): void;
  clip?(): void;
  /** Optional: a check box's tick. */
  moveTo?(x: number, y: number): void;
  lineTo?(x: number, y: number): void;
  stroke?(): void;
}

/** The side of a drawn check box, the browser control's 13 px. */
const EXPORT_CHECKBOX_SIZE = 13;
/** A checked box's fill, the grid control's default accent. */
const EXPORT_CHECKBOX_ACCENT = "#0075ff";

/** A check box centred in its cell, as the grid draws the cell's control. */
function drawCellCheckbox(ctx: CanvasCtx2D, cell: CellLayout, checked: boolean): void {
  const size = Math.min(EXPORT_CHECKBOX_SIZE, cell.width - 2, cell.height - 2);
  if (!(size > 3)) return;
  const x = Math.round(cell.x + (cell.width - size) / 2) + 0.5;
  const y = Math.round(cell.y + (cell.height - size) / 2) + 0.5;
  ctx.lineWidth = 1;
  if (checked) {
    ctx.fillStyle = EXPORT_CHECKBOX_ACCENT;
    ctx.fillRect(x - 0.5, y - 0.5, size + 1, size + 1);
    if (ctx.beginPath && ctx.moveTo && ctx.lineTo && ctx.stroke) {
      ctx.strokeStyle = "#ffffff";
      ctx.lineWidth = Math.max(1.5, size / 7);
      ctx.beginPath();
      ctx.moveTo(x + size * 0.22, y + size * 0.52);
      ctx.lineTo(x + size * 0.42, y + size * 0.72);
      ctx.lineTo(x + size * 0.78, y + size * 0.3);
      ctx.stroke();
      ctx.lineWidth = 1;
    }
    return;
  }
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(x, y, size, size);
  ctx.strokeStyle = "#767676";
  ctx.strokeRect(x, y, size, size);
}

/** A cell picture decoded and ready to draw, keyed by its `src`. */
export interface LoadedCellPicture {
  image: CanvasImageSource;
  width: number;
  height: number;
}

/**
 * Where a picture lands in its cell: inside the cell's 1px inset, scaled to
 * fit keeping its aspect ratio and placed along the object-position keywords,
 * the box the grid's `object-fit: contain` image takes. Null when there is
 * nothing to draw.
 */
export function fitCellPicture(
  cell: { x: number; y: number; width: number; height: number },
  natural: { width: number; height: number },
  position = "center bottom",
): { x: number; y: number; width: number; height: number } | null {
  const boxX = cell.x + 1;
  const boxY = cell.y + 1;
  const boxW = cell.width - 2;
  const boxH = cell.height - 2;
  if (
    !(boxW > 0) ||
    !(boxH > 0) ||
    !(natural.width > 0) ||
    !(natural.height > 0) ||
    !Number.isFinite(natural.width) ||
    !Number.isFinite(natural.height)
  ) {
    return null;
  }
  const scale = Math.min(boxW / natural.width, boxH / natural.height);
  const width = natural.width * scale;
  const height = natural.height * scale;
  const [horizontal = "center", vertical = "center"] = position.trim().split(/\s+/);
  const x =
    horizontal === "left"
      ? boxX
      : horizontal === "right"
        ? boxX + boxW - width
        : boxX + (boxW - width) / 2;
  const y =
    vertical === "top"
      ? boxY
      : vertical === "bottom"
        ? boxY + boxH - height
        : boxY + (boxH - height) / 2;
  return { x, y, width, height };
}

/**
 * Decode every picture a layout draws (the cells' picture values and the
 * drawings over the range), once per distinct source. Resolves to null when
 * any of them cannot be decoded, so a picture is never silently left out of
 * the copy. Browser-only (`Image`).
 */
export async function loadLayoutPictures(
  layout: RangeLayout,
): Promise<Map<string, LoadedCellPicture> | null> {
  const sources = new Set<string>();
  for (const cell of layout.cells) if (cell.picture?.src) sources.add(cell.picture.src);
  for (const drawing of layout.drawings ?? []) sources.add(drawing.src);
  const loaded = new Map<string, LoadedCellPicture>();
  if (sources.size === 0) return loaded;
  if (typeof Image === "undefined") return null;
  try {
    await Promise.all(
      [...sources].map(async (src) => {
        const image = new Image();
        await new Promise<void>((resolve, reject) => {
          image.onload = () => resolve();
          image.onerror = () => reject(new Error("cell picture failed to load"));
          // A picture stored beside the page is drawn from its picture route.
          image.src = pictureObjectDisplaySrc(src);
        });
        if (typeof image.decode === "function") await image.decode();
        if (!(image.naturalWidth > 0) || !(image.naturalHeight > 0)) {
          throw new Error("cell picture has no size");
        }
        loaded.set(src, { image, width: image.naturalWidth, height: image.naturalHeight });
      }),
    );
  } catch {
    return null;
  }
  return loaded;
}

/** Break one hard line of a wrapped cell at the width available, by words
 *  where the text has them and by characters where it does not (CJK). */
function wrapLine(ctx: CanvasCtx2D, line: string, maxWidth: number): string[] {
  if (!ctx.measureText || !line) return [line];
  const fits = (text: string) => ctx.measureText!(text).width <= maxWidth;
  if (fits(line)) return [line];
  const out: string[] = [];
  const units = line.includes(" ") ? splitBehind(line, /(?:)/, AFTER_SPACE) : Array.from(line);
  let current = "";
  for (const unit of units) {
    const candidate = current + unit;
    if (current && !fits(candidate.trimEnd())) {
      out.push(current.trimEnd());
      current = unit;
    } else {
      current = candidate;
    }
  }
  if (current) out.push(current.trimEnd());
  return out;
}

/**
 * Draw one line spread across `width` from `left` (XA Justify / Distributed):
 * the extra space goes between words, or between characters when the line
 * has no spaces (East Asian text) or the alignment is Distributed and the
 * line is one word. A single character is centred.
 */
function drawSpreadLine(
  ctx: CanvasCtx2D,
  line: string,
  left: number,
  width: number,
  y: number,
  distributed: boolean,
  measure: (text: string) => number,
): void {
  const text = line.trim();
  const words = text.split(/\s+/).filter(Boolean);
  const perCharacter = words.length < 2 && (distributed || !text.includes(" "));
  const units = perCharacter ? Array.from(text) : words;
  const previousAlign = ctx.textAlign;
  ctx.textAlign = "left";
  if (units.length < 2) {
    ctx.textAlign = "center";
    ctx.fillText(text, left + width / 2, y, width);
    ctx.textAlign = previousAlign;
    return;
  }
  const widths = units.map((unit) => measure(unit));
  const used = widths.reduce((sum, value) => sum + value, 0);
  const gap = Math.max(0, (width - used) / (units.length - 1));
  let x = left;
  units.forEach((unit, index) => {
    ctx.fillText(unit, x, y);
    x += widths[index] + gap;
  });
  ctx.textAlign = previousAlign;
}

function runFont(run: LayoutTextRun): string {
  return [run.italic ? "italic" : "", run.bold ? "bold" : "", `${run.fontSize}px`, cssFontFamilyList(run.fontFamily)]
    .filter(Boolean)
    .join(" ");
}

/** A run's drawn width; an estimate where the context cannot measure. */
function runWidth(ctx: CanvasCtx2D, run: LayoutTextRun): number {
  if (!ctx.measureText) return run.text.length * run.fontSize * 0.6;
  ctx.font = runFont(run);
  return ctx.measureText(run.text).width;
}

/** A cell's rich text as drawn lines of runs: one per hard line, and a
 *  wrapped cell breaks each at its width by words, or characters (CJK). */
function richTextLines(
  ctx: CanvasCtx2D,
  runs: readonly LayoutTextRun[],
  wrap: boolean,
  maxWidth: number,
): LayoutTextRun[][] {
  const hard: LayoutTextRun[][] = [[]];
  for (const run of runs) {
    run.text.split(/\r?\n/).forEach((part, index) => {
      if (index > 0) hard.push([]);
      if (part) hard[hard.length - 1].push({ ...run, text: part });
    });
  }
  if (!wrap || !ctx.measureText) return hard;
  const out: LayoutTextRun[][] = [];
  for (const line of hard) {
    const spaced = line.some((run) => /\s/.test(run.text));
    let current: LayoutTextRun[] = [];
    let width = 0;
    for (const run of line) {
      for (const unit of spaced ? splitBehind(run.text, /(?:)/, AFTER_SPACE) : Array.from(run.text)) {
        const piece = { ...run, text: unit };
        const unitWidth = runWidth(ctx, { ...piece, text: unit.trimEnd() });
        if (current.length > 0 && width + unitWidth > maxWidth) {
          out.push(current);
          current = [];
          width = 0;
        }
        const last = current[current.length - 1];
        if (last && runFont(last) === runFont(piece) && last.color === piece.color && last.underline === piece.underline && last.strike === piece.strike && last.baselineShift === piece.baselineShift) {
          current[current.length - 1] = { ...last, text: last.text + unit };
        } else {
          current.push(piece);
        }
        width += runWidth(ctx, piece);
      }
    }
    out.push(current);
  }
  return out.map((line) =>
    line.length > 0
      ? [...line.slice(0, -1), { ...line[line.length - 1], text: line[line.length - 1].text.trimEnd() }]
      : line,
  );
}

/** The middle of a cell's first text line: its block of `lineCount` lines
 *  stands at the top, middle or bottom of the cell (its vertical alignment,
 *  inset as the grid insets it), in what is left under a shown phonetic
 *  guide's band. */
function firstLineMiddle(cell: CellLayout, lineCount: number, lineHeight: number): number {
  const band = cell.phonetic ? Math.min(cell.phonetic.bandHeight, cell.height / 2) : 0;
  const block = lineCount * lineHeight;
  const blockTop =
    cell.verticalAlign === "top"
      ? cell.y + band + EXPORT_CELL_PADDING_Y
      : cell.verticalAlign === "middle"
        ? cell.y + band + (cell.height - band - block) / 2
        : cell.y + cell.height - EXPORT_CELL_PADDING_Y - block;
  return blockTop + lineHeight / 2;
}

/**
 * Draw a cell's phonetic guide in the band over its text: each reading over
 * its base characters (Left / Center / Distributed), or all of them run
 * together from where the text starts (No control) — the grid's layout of the
 * guide (live audit B388). Base widths are measured in the cell's own font.
 */
function drawPhoneticGuide(
  ctx: CanvasCtx2D,
  cell: CellLayout,
  guide: LayoutPhoneticGuide,
  padX = EXPORT_CELL_PADDING_X,
): void {
  const baseFont = [cell.italic ? "italic" : "", cell.bold ? "bold" : "", `${cell.fontSize}px`, cell.fontFamily]
    .filter(Boolean)
    .join(" ");
  ctx.font = baseFont;
  const measure = (text: string) =>
    ctx.measureText ? ctx.measureText(text).width : text.length * cell.fontSize * 0.6;
  const widths = guide.segments.map((segment) => measure(segment.base));
  const total = widths.reduce((sum, width) => sum + width, 0);
  let x =
    cell.textAlign === "right"
      ? cell.x + cell.width - padX - total
      : cell.textAlign === "center"
        ? cell.x + cell.width / 2 - total / 2
        : cell.x + padX;
  const y = cell.y + Math.min(guide.bandHeight, cell.height / 2) / 2;
  ctx.font = guide.font;
  ctx.fillStyle = guide.color;
  ctx.textBaseline = "middle";
  if (guide.alignment === "noControl") {
    ctx.textAlign = "left";
    ctx.fillText(guide.segments.map((segment) => segment.ruby ?? "").join(""), x, y);
    return;
  }
  guide.segments.forEach((segment, index) => {
    const width = widths[index];
    if (segment.ruby) {
      const chars = Array.from(segment.ruby);
      if (guide.alignment === "distributed" && chars.length > 1) {
        // Spread over the base, one reading character per equal share.
        ctx.textAlign = "center";
        chars.forEach((char, at) => ctx.fillText(char, x + ((at + 0.5) * width) / chars.length, y));
      } else if (guide.alignment === "center") {
        ctx.textAlign = "center";
        ctx.fillText(segment.ruby, x + width / 2, y);
      } else {
        ctx.textAlign = "left";
        ctx.fillText(segment.ruby, x, y);
      }
    }
    x += width;
  });
}

/** Draw a cell's rich text run by run, each in its own face and colour, the
 *  lines aligned and stacked as a plain cell's are (B338). */
function drawRichText(
  ctx: CanvasCtx2D,
  cell: CellLayout,
  runs: readonly LayoutTextRun[],
  padX = EXPORT_CELL_PADDING_X,
): void {
  const maxWidth = Math.max(1, cell.width - padX * 2);
  const lines = richTextLines(ctx, runs, cell.wrapText, maxWidth);
  const tallest = Math.max(cell.fontSize, ...runs.map((run) => run.fontSize));
  const lineHeight = tallest * SPREADSHEET_WRAP_LINE_HEIGHT;
  const firstY = firstLineMiddle(cell, lines.length, lineHeight);
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  lines.forEach((line, index) => {
    const widths = line.map((run) => runWidth(ctx, run));
    const total = widths.reduce((sum, width) => sum + width, 0);
    // As fillText's maxWidth does for a plain cell, a line wider than the
    // cell is drawn condensed to fit it.
    const squeeze = total > maxWidth ? maxWidth / total : 1;
    const drawn = total * squeeze;
    let x =
      cell.textAlign === "right"
        ? cell.x + cell.width - padX - drawn
        : cell.textAlign === "center"
          ? cell.x + cell.width / 2 - drawn / 2
          : cell.x + padX;
    const y = firstY + index * lineHeight;
    line.forEach((run, runIndex) => {
      const width = widths[runIndex] * squeeze;
      // A raised or lowered run (the 2 of H2O) moves off the line's middle.
      const runY = y + (run.baselineShift ?? 0);
      ctx.font = runFont(run);
      ctx.fillStyle = run.color;
      ctx.fillText(run.text, x, runY, Math.max(1, width));
      const rule = Math.max(1, Math.round(run.fontSize / 14));
      if (run.underline) ctx.fillRect(x, runY + Math.round(run.fontSize * 0.45), width, rule);
      if (run.strike) ctx.fillRect(x, runY, width, rule);
      x += width;
    });
  });
  // The runs carry their own underlines; the cell's accounting rule runs
  // across it under the last line.
  if (cell.underline === "singleAccounting" || cell.underline === "doubleAccounting") {
    drawCellUnderline(ctx, cell, [{ x: cell.x, y: firstY + (lines.length - 1) * lineHeight, width: cell.width }], tallest);
  }
}

/**
 * A cell's own underline, as the grid draws it (live audit B400): Single and
 * Double under each line's characters; Single and Double Accounting a rule
 * across the whole cell, under the last line and below its descenders.
 * `lines` are the drawn lines' extents, `y` each one's middle.
 */
function drawCellUnderline(
  ctx: CanvasCtx2D,
  cell: CellLayout,
  lines: ReadonlyArray<{ x: number; y: number; width: number }>,
  fontPx: number,
): void {
  const kind = cell.underline;
  if (!kind || lines.length === 0) return;
  const rule = Math.max(1, Math.round(fontPx / 14));
  const double = kind === "double" || kind === "doubleAccounting";
  ctx.fillStyle = cell.textColor;
  if (kind === "singleAccounting" || kind === "doubleAccounting") {
    const y = lines[lines.length - 1].y + Math.round(fontPx * 0.6);
    ctx.fillRect(cell.x, y, cell.width, rule);
    if (double) ctx.fillRect(cell.x, y + rule * 2, cell.width, rule);
    return;
  }
  for (const line of lines) {
    const y = line.y + Math.round(fontPx * 0.45);
    ctx.fillRect(line.x, y, line.width, rule);
    if (double) ctx.fillRect(line.x, y + rule * 2, line.width, rule);
  }
}

/**
 * Draw a `RangeLayout` onto a canvas context.  PURE in the sense that all
 * geometry is pre-computed; only the 2D API calls live here.
 *
 * Exported so unit-tests can supply a mock `CanvasCtx2D` and assert on the
 * draw calls via a spy/stub pattern.
 */
export function drawLayoutToCtx(
  layout: RangeLayout,
  ctx: CanvasCtx2D,
  /** The layout's cell pictures, decoded ({@link loadLayoutPictures}). */
  pictures?: ReadonlyMap<string, LoadedCellPicture>,
): void {
  const { canvasWidth, canvasHeight, cells } = layout;
  const padX =
    typeof layout.cellPaddingX === "number" && Number.isFinite(layout.cellPaddingX) && layout.cellPaddingX >= 0
      ? layout.cellPaddingX
      : EXPORT_CELL_PADDING_X;
  const canClip = Boolean(ctx.save && ctx.restore && ctx.beginPath && ctx.rect && ctx.clip);

  // White background for the whole export area.
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvasWidth, canvasHeight);

  // Cell backgrounds + text.
  for (const cell of cells) {
    // Background (only when non-null / non-white).
    if (cell.backgroundColor && cell.backgroundColor !== "#ffffff") {
      ctx.fillStyle = cell.backgroundColor;
      ctx.fillRect(cell.x, cell.y, cell.width, cell.height);
    }

    // A picture value, fitted in the cell as the grid shows it.
    const loaded = cell.picture ? pictures?.get(cell.picture.src) : undefined;
    if (cell.picture && loaded && ctx.drawImage) {
      const box = fitCellPicture(cell, loaded, cell.picture.position);
      if (box) ctx.drawImage(loaded.image, box.x, box.y, box.width, box.height);
    }
    if (cell.checkbox) drawCellCheckbox(ctx, cell, cell.checkbox.checked);

    // Text, in the cell's own base direction (B397), kept inside the cell
    // as the grid keeps it: a line taller than its row is cut at the edges.
    const clipped = canClip && Boolean(cell.text);
    if (clipped) {
      ctx.save!();
      ctx.beginPath!();
      ctx.rect!(cell.x, cell.y, cell.width, cell.height);
      ctx.clip!();
    }
    const previousDirection = ctx.direction;
    if (cell.direction && cell.text) ctx.direction = cell.direction;
    if (cell.text && cell.runs) {
      drawRichText(ctx, cell, cell.runs, padX);
    } else if (cell.text) {
      // A whole-cell script is drawn smaller, raised or lowered (B399).
      const drawSize = cell.script ? cell.fontSize * SPREADSHEET_SCRIPT_FONT_SCALE : cell.fontSize;
      const shift = cell.script ? scriptBaselineShiftPx(cell.script, cell.fontSize) : 0;
      const fontParts: string[] = [];
      if (cell.italic) fontParts.push("italic");
      if (cell.bold) fontParts.push("bold");
      fontParts.push(`${drawSize}px`);
      fontParts.push(cell.fontFamily);

      ctx.font = fontParts.join(" ");
      ctx.fillStyle = cell.textColor;
      ctx.textBaseline = "middle";
      // Fill runs its copies from the cell's left edge (B398).
      const textAlign = cell.fill ? "left" : cell.textAlign;
      ctx.textAlign = textAlign as CanvasTextAlign;

      const textX =
        textAlign === "right"
          ? cell.x + cell.width - padX
          : textAlign === "center"
            ? cell.x + cell.width / 2
            : cell.x + padX;
      const maxWidth = Math.max(1, cell.width - padX * 2);
      const measure = (text: string) =>
        ctx.measureText ? ctx.measureText(text).width : text.length * drawSize * 0.6;
      // A hard line break stays a line break in the picture, and a wrapped
      // cell breaks at its width (live audit B280); the lines stack at the
      // grid's pitch, the block at the cell's top, middle or bottom.
      const hardLines = cell.lines.length ? cell.lines : [cell.text];
      // Justify / Distributed wrap like Wrap Text (XA): their lines are
      // spread across the cell below.
      const spread = cell.horizontalSpread;
      const spreadIndent = spread === "distributed" ? Math.min(cell.spreadIndent ?? 0, maxWidth / 3) : 0;
      const spreadWidth = Math.max(1, maxWidth - spreadIndent * 2);
      const paragraphEnds = new Set<number>();
      const lines = cell.fill && hardLines.length === 1
        ? [horizontalFillText(cell.text, measure(cell.text), maxWidth)]
        : cell.wrapText || spread
          ? hardLines.flatMap((line) => {
              const wrapped = wrapLine(ctx, line, spread ? spreadWidth : maxWidth);
              return wrapped;
            })
          : hardLines;
      if (spread) {
        let at = -1;
        for (const line of hardLines) {
          at += wrapLine(ctx, line, spreadWidth).length;
          paragraphEnds.add(at);
        }
      }
      const lineHeight = cell.fontSize * SPREADSHEET_WRAP_LINE_HEIGHT;
      const lineCount = lines.length;
      // Vertical Justify / Distributed spread the lines over the cell's
      // height; otherwise the block stands by the vertical alignment.
      const lineY = (index: number): number => {
        const top = cell.y + EXPORT_CELL_PADDING_Y;
        const usable = Math.max(lineHeight, cell.height - EXPORT_CELL_PADDING_Y * 2);
        if (cell.verticalSpread === "justify") {
          if (lineCount === 1) return top + lineHeight / 2;
          return top + lineHeight / 2 + (index * (usable - lineHeight)) / (lineCount - 1);
        }
        if (cell.verticalSpread === "distributed") {
          const gap = Math.max(0, (usable - lineCount * lineHeight) / (lineCount + 1));
          return top + gap * (index + 1) + lineHeight * index + lineHeight / 2;
        }
        return firstLineMiddle(cell, lineCount, lineHeight) + index * lineHeight;
      };
      const firstY = lineY(0);
      lines.forEach((line, index) => {
        const stretch =
          spread === "distributed" ||
          (spread === "justify" && (!paragraphEnds.has(index) || cell.justifyLastLine));
        if (spread && stretch && line.trim() !== "") {
          drawSpreadLine(
            ctx,
            line,
            cell.x + padX + spreadIndent,
            spreadWidth,
            lineY(index) + shift,
            spread === "distributed",
            measure,
          );
          return;
        }
        if (cell.numberLayout && lines.length === 1) {
          drawNumberLayoutLine(ctx, cell.numberLayout, cell.x + padX, cell.x + cell.width - padX, lineY(index) + shift, measure);
          ctx.textAlign = textAlign as CanvasTextAlign;
          return;
        }
        ctx.fillText(line, textX, lineY(index) + shift, maxWidth);
      });
      drawCellUnderline(
        ctx,
        cell,
        lines.map((line, index) => {
          const width = Math.min(measure(line), maxWidth);
          return {
            x: textAlign === "right" ? textX - width : textAlign === "center" ? textX - width / 2 : textX,
            y: firstY + index * lineHeight,
            width,
          };
        }),
        cell.fontSize,
      );
    }
    if (cell.direction && cell.text) ctx.direction = previousDirection ?? "inherit";
    if (cell.text && cell.phonetic) drawPhoneticGuide(ctx, cell, cell.phonetic, padX);
    if (clipped) ctx.restore!();
  }

  // Gridlines (drawn last so they appear over backgrounds).
  ctx.strokeStyle = EXPORT_GRID_LINE_COLOR;
  ctx.lineWidth = 0.5;
  for (const cell of cells) {
    ctx.strokeRect(cell.x + 0.5, cell.y + 0.5, cell.width - 1, cell.height - 1);
  }

  // The cells' own borders, over the gridlines, as bands along each edge: a
  // double edge is two bands with a gap (live audit B297).
  const band = (edge: BorderEdgeLine, x: number, y: number, w: number, h: number) => {
    ctx.fillStyle = edge.color;
    if (edge.double && edge.width >= 3) {
      const thin = Math.max(1, Math.floor(edge.width / 3));
      if (w >= h) {
        ctx.fillRect(x, y, w, thin);
        ctx.fillRect(x, y + h - thin, w, thin);
      } else {
        ctx.fillRect(x, y, thin, h);
        ctx.fillRect(x + w - thin, y, thin, h);
      }
      return;
    }
    ctx.fillRect(x, y, w, h);
  };
  for (const cell of cells) {
    const { top, right, bottom, left } = cell.borders;
    if (top) band(top, cell.x, cell.y, cell.width, top.width);
    if (bottom) band(bottom, cell.x, cell.y + cell.height - bottom.width, cell.width, bottom.width);
    if (left) band(left, cell.x, cell.y, left.width, cell.height);
    if (right) band(right, cell.x + cell.width - right.width, cell.y, right.width, cell.height);
  }

  // The drawings over the range, on top of the cells, bottom first; the
  // canvas's edges cut a drawing that runs past the range.
  for (const drawing of layout.drawings ?? []) {
    const loaded = pictures?.get(drawing.src);
    if (loaded && ctx.drawImage) {
      ctx.drawImage(loaded.image, drawing.x, drawing.y, drawing.width, drawing.height);
    }
  }
}

// ---------------------------------------------------------------------------
// Browser-level export functions
// ---------------------------------------------------------------------------

/**
 * Raster a detached range layout into a PNG data URL for Excel's Paste as
 * Picture command. Unlike download/copy sinks this is synchronous, allowing
 * the caller to validate the complete image object before its one worksheet
 * mutation. Every browser/canvas failure is fail-closed, including tainted
 * canvases and oversized data URLs rejected by the sheet-image model.
 */
export function rangeLayoutToPngDataUrl(
  layout: RangeLayout,
  scale = 2,
): string | null {
  if (
    typeof document === "undefined" ||
    typeof document.createElement !== "function" ||
    !Number.isFinite(layout.canvasWidth) ||
    !Number.isFinite(layout.canvasHeight) ||
    layout.canvasWidth <= 0 ||
    layout.canvasHeight <= 0 ||
    !Number.isFinite(scale) ||
    scale < 1 ||
    scale > 4
  ) {
    return null;
  }

  const pixelWidth = Math.round(layout.canvasWidth * scale);
  const pixelHeight = Math.round(layout.canvasHeight * scale);
  if (
    !Number.isSafeInteger(pixelWidth) ||
    !Number.isSafeInteger(pixelHeight) ||
    pixelWidth < 1 ||
    pixelHeight < 1 ||
    pixelWidth > MAX_RANGE_PICTURE_CANVAS_EDGE ||
    pixelHeight > MAX_RANGE_PICTURE_CANVAS_EDGE ||
    pixelWidth * pixelHeight > MAX_RANGE_PICTURE_CANVAS_PIXELS
  ) {
    return null;
  }

  try {
    const canvas = document.createElement("canvas");
    canvas.width = pixelWidth;
    canvas.height = pixelHeight;
    const ctx = canvas.getContext("2d");
    if (!ctx || typeof canvas.toDataURL !== "function") return null;
    ctx.scale(scale, scale);
    drawLayoutToCtx(layout, ctx);
    const dataUrl = canvas.toDataURL("image/png");
    if (
      typeof dataUrl !== "string" ||
      dataUrl.length > MAX_IMAGE_SRC_LENGTH ||
      dataUrl.slice(0, 22).toLowerCase() !== "data:image/png;base64,"
    ) {
      return null;
    }
    return dataUrl;
  } catch {
    return null;
  }
}

/**
 * Raster the pre-computed layout to a PNG and trigger a download.
 * BROWSER-ONLY — `document.createElement("canvas")` must be available.
 * Returns `true` when a download was triggered, `false` on capability failure.
 *
 * @param layout  Output of {@link computeRangeLayout}.
 * @param label   Human-readable label for the filename (sanitized).
 * @param scale   Device-pixel multiplier (default 2 for crisp high-DPI PNG).
 */
export async function downloadRangeLayout(
  layout: RangeLayout,
  label: string,
  scale = 2,
): Promise<boolean> {
  if (
    typeof document === "undefined" ||
    typeof document.createElement !== "function"
  ) {
    return false;
  }
  const { canvasWidth, canvasHeight } = layout;
  const pictures = await loadLayoutPictures(layout);
  if (!pictures) return false;
  const s = Math.max(1, Math.min(4, scale));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(canvasWidth * s));
  canvas.height = Math.max(1, Math.round(canvasHeight * s));
  const ctx = canvas.getContext("2d");
  if (!ctx) return false;

  ctx.scale(s, s);
  drawLayoutToCtx(layout, ctx, pictures);

  return new Promise<boolean>((resolve) => {
    if (typeof canvas.toBlob !== "function") {
      resolve(false);
      return;
    }
    canvas.toBlob((blob) => {
      if (!blob) {
        resolve(false);
        return;
      }
      resolve(_downloadPngBlob(_toPngFilename(label, "spreadsheet"), blob));
    }, "image/png");
  });
}

/**
 * Serialize a chart `<svg>` to a data URL and raster + download as PNG.
 * Delegates to the shared `lib/export/svg-to-png` pipeline via a dynamic
 * import so this module has zero static dependency on it (keeps it
 * unit-testable in isolation and avoids the file-not-found error if the file
 * doesn't exist on the current checkout).
 *
 * @param svgEl   The `<svg>` rendered by MiniChart (live DOM element).
 * @param label   Human-readable chart label for the filename.
 * @param opts    Optional width/height/scale/background overrides.
 */
export async function exportChartSvgToPng(
  svgEl: SVGSVGElement,
  label: string,
  opts: SvgToPngOptions = {},
): Promise<boolean> {
  // Dynamic import so tests that don't exercise this path are unaffected.
  const { exportSvgElementToPng } = await import("@mrnote/core/web/lib/export/svg-to-png.ts");
  return exportSvgElementToPng(svgEl, label, {
    scale: 2,
    background: "#ffffff",
    ...opts,
  });
}

/**
 * One line of a number format that fills: the parts before the fill from the
 * left edge, the parts after it ending at the right edge, and the fill
 * character repeated as many whole times as fit between them (a space fill
 * is just the gap). A reserve (`_x`) holds x's width and draws nothing.
 */
function drawNumberLayoutLine(
  ctx: CanvasCtx2D,
  parts: readonly NumberFormatLayoutPart[],
  left: number,
  right: number,
  y: number,
  measure: (text: string) => number,
): void {
  const fillIndex = parts.findIndex((part) => part.kind === "fill");
  const widthOf = (part: NumberFormatLayoutPart) => measure(part.kind === "text" ? part.text : part.character);
  ctx.textAlign = "left";
  let x = left;
  for (const part of parts.slice(0, fillIndex)) {
    if (part.kind === "text" && part.text) ctx.fillText(part.text, x, y);
    x += widthOf(part);
  }
  let end = right;
  for (const part of parts.slice(fillIndex + 1).reverse()) {
    end -= widthOf(part);
    if (part.kind === "text" && part.text) ctx.fillText(part.text, end, y);
  }
  const fill = parts[fillIndex];
  if (fill?.kind === "fill" && fill.character.trim()) {
    const glyph = measure(fill.character);
    const count = glyph > 0 ? Math.max(0, Math.floor((end - x + 1e-7) / glyph)) : 0;
    if (count > 0) ctx.fillText(fill.character.repeat(count), x, y);
  }
}

/**
 * Convenience: given a pre-built `ExportCell` map and range bounds, compute the
 * layout + download as PNG.  This is what SpreadsheetWindow calls.
 */
export async function exportRangeToPng(
  params: RangeLayoutInput,
  label: string,
): Promise<boolean> {
  const layout = computeRangeLayout(params);
  return downloadRangeLayout(layout, label);
}

/**
 * Copy the range to the SYSTEM CLIPBOARD as a PNG (Excel Home ▸ Copy ▸
 * Copy as Picture). Same layout + draw as the download path; only the sink
 * differs. Returns false where the async Clipboard API is unavailable
 * (insecure context, or a browser without ClipboardItem) — the caller
 * surfaces that as a toast rather than failing silently.
 *
 * `drawings` are the drawings over the range while they are still being
 * rendered: the clipboard write starts at once, inside the click that asked
 * for it (the permission to write lasts only a moment), and the picture is
 * drawn when they are ready. Null fails the copy rather than leaving them
 * out.
 */
export async function copyRangeAsPictureToClipboard(
  params: RangeLayoutInput,
  options: {
    drawings?: Promise<ReadonlyArray<RangePictureDrawing> | null>;
    scale?: number;
  } = {},
): Promise<boolean> {
  if (
    typeof document === "undefined" ||
    typeof navigator === "undefined" ||
    !navigator.clipboard ||
    typeof ClipboardItem === "undefined"
  ) {
    return false;
  }
  const s = Math.max(1, Math.min(4, options.scale ?? 2));
  const png = (async (): Promise<Blob> => {
    const drawings = options.drawings ? await options.drawings : params.drawings;
    if (options.drawings && !drawings) throw new Error("range picture drawings unavailable");
    const layout = computeRangeLayout(drawings ? { ...params, drawings } : params);
    // The cells' picture values and the drawings are decoded before anything
    // is drawn; one that cannot be is a failed copy, never a picture with a
    // hole or `#VALUE!`.
    const pictures = await loadLayoutPictures(layout);
    if (!pictures) throw new Error("range picture images unavailable");
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(layout.canvasWidth * s));
    canvas.height = Math.max(1, Math.round(layout.canvasHeight * s));
    const ctx = canvas.getContext("2d");
    if (!ctx || typeof canvas.toBlob !== "function") throw new Error("no canvas");
    ctx.scale(s, s);
    drawLayoutToCtx(layout, ctx, pictures);
    const blob = await new Promise<Blob | null>((resolve) => {
      canvas.toBlob((result) => resolve(result), "image/png");
    });
    if (!blob) throw new Error("range picture not encoded");
    return blob;
  })();
  // Its failure is read below; never reported as unhandled meanwhile.
  png.catch(() => undefined);
  try {
    await navigator.clipboard.write([new ClipboardItem({ "image/png": png })]);
    return true;
  } catch {
    // A browser that takes only a finished Blob gets the finished one.
    try {
      const blob = await png;
      await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
      return true;
    } catch {
      return false;
    }
  }
}

// ---------------------------------------------------------------------------
// Chart SVG discovery helper (browser-only, uses querySelector)
// ---------------------------------------------------------------------------

/**
 * Find the first `<svg>` inside the chart canvas container for the given chart
 * id (`data-chart-id` attribute on the wrapper div rendered by FloatingChart).
 *
 * Returns `null` when the container or SVG is not found.
 */
export function findChartSvg(chartId: string): SVGSVGElement | null {
  if (typeof document === "undefined") return null;
  const container = document.querySelector<HTMLElement>(
    `[data-chart-id="${CSS.escape(chartId)}"] [data-testid="spreadsheet-chart-canvas-${CSS.escape(chartId)}"]`,
  );
  if (!container) return null;
  return container.querySelector<SVGSVGElement>("svg");
}
