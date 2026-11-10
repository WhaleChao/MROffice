/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** The look the grid derives for a cell, in the cell style's own keys. */
export interface SpreadsheetPdfDerivedStyle {
  backgroundColor?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  wrapText?: boolean;
  textAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  borderEdges?: Partial<Record<"top" | "right" | "bottom" | "left", { style: string; color: string }>>;
}
/**
 * A cell the grid shows that the saved workbook does not store as a cell (a
 * PivotTable report cell, a dynamic-array spill) or paints beyond the cell's
 * own style (a table style band). The server lays it over the saved cell; a
 * derived text never replaces a saved one.
 */
/**
 * What a matching conditional format paints on the cell: unlike a table
 * style's paint (`s`), it OVERRIDES the cell's own format, as Excel's does.
 * `backgroundColor: "none"` is a rule's No Fill.
 */
export interface SpreadsheetPdfConditionalStyle {
  backgroundColor?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
}
/** A data bar as the grid draws it (`dataBarPaintCss`), in percent of the cell. */
export interface SpreadsheetPdfDataBar {
  /** Left edge and width, 0..100. */
  s: number;
  w: number;
  /** #rrggbb fill (the negative colour for a negative bar). */
  c: string;
  /** Gradient fading to white away from the axis. */
  g?: true;
  /** Negative bar: fades the other way. */
  n?: true;
  /** Border colour. */
  b?: string;
  /** Zero axis position (0..100) and colour, where drawn. */
  a?: number;
  ac?: string;
}
export interface SpreadsheetPdfDerivedCell {
  r: number;
  c: number;
  t?: string;
  k?: "text" | "number";
  s?: SpreadsheetPdfDerivedStyle;
  /** A conditional format's style, over the cell's own. */
  cf?: SpreadsheetPdfConditionalStyle;
  bar?: SpreadsheetPdfDataBar;
  /** An icon set's icon: an index into the sheet's `glyphs`. */
  icon?: number;
  /** The cell's text as a matching rule's number format shows it
   * (Format > Number of a conditional format): replaces the saved display. */
  ft?: string;
  /** The cell's value is not shown (Show Icon Only, Show Bar Only). */
  hide?: true;
  /** The cell's sparkline: an index into the sheet's `glyphs`. */
  spark?: number;
  /** Its number format's text colour ([Red]), as the grid draws it. */
  nc?: string;
  /** Its number format's layout: text (0), a reserved character's width
   * (1, `_x`) and the fill (2, `*x`), in order (`numberFormatLayout`). */
  nl?: ReadonlyArray<readonly [0 | 1 | 2, string]>;
}
/** Disposable view description; values/styles always come from the saved page. */
export interface SpreadsheetPdfSheetView {
  sheetId: string;
  hiddenRowRanges: [number, number][];
  /** Ordered ascending OR descending runs covering the entire source row domain. */
  rowOrderRanges?: [number, number][];
  derivedCells?: SpreadsheetPdfDerivedCell[];
  /** SVG pictures the derived cells draw (icon-set icons, sparklines), each
   * the grid's own glyph component rendered to markup. */
  glyphs?: string[];
  /** The sheet's slicers and timelines as it shows them: each frame in CSS
   * px from the sheet's top-left cell, and its picture (a `glyphs` index). */
  panels?: SpreadsheetPdfPanel[];
}
export interface SpreadsheetPdfPanel {
  x: number;
  y: number;
  w: number;
  h: number;
  glyph: number;
}
export const MAX_SPREADSHEET_PDF_PANELS = 200;
export interface SpreadsheetPdfExportView {
  version: 1;
  sheets: SpreadsheetPdfSheetView[];
  /** File > Print's Print what (`lib/spreadsheet/print-scope`): the sheets
   * printed (every sheet when absent), the selection on the one printed
   * sheet, and Ignore Print Area. */
  printSheetIds?: string[];
  printSelection?: { r1: number; c1: number; r2: number; c2: number };
  ignorePrintArea?: true;
  /** Some of the grid's extras (derived cells, glyphs) were left out past
   * the view's budget: the server's old drawing, which paints them from this
   * view, says so on the export instead of printing fewer in silence. The
   * editor's own print derives them itself and needs none of them. */
  budgetTrimmed?: true;
}
export class SpreadsheetPdfViewError extends Error {
  constructor(
    readonly code: "too_complex" | "source_changed" | "view_unavailable",
    /** A `source_changed` the server answered with the page's newer saved
     * version (page_version_conflict): the export may run again at it. */
    readonly newerContentVersion?: number,
  ) {
    super(code);
    this.name = "SpreadsheetPdfViewError";
  }
}
export const MAX_SPREADSHEET_PDF_VIEW_RUNS = 20_000;
export const MAX_SPREADSHEET_PDF_DERIVED_CELLS = 150_000;
/** The derived cells' share of the export request (JSON characters): with
 * the glyphs' it stays well inside the API's 20 MiB JSON body. */
export const MAX_SPREADSHEET_PDF_DERIVED_CHARS = 12_000_000;
export const MAX_SPREADSHEET_PDF_FILTER_WORK = 1_000_000;
/** Glyphs one PDF view may carry across every sheet, and their size. */
export const MAX_SPREADSHEET_PDF_GLYPHS = 5_000;
export const MAX_SPREADSHEET_PDF_GLYPH_CHARS = 20_000;
export const MAX_SPREADSHEET_PDF_GLYPH_TOTAL_CHARS = 4_000_000;

export function spreadsheetPdfRowRuns(indices: readonly number[], descending = false): [number, number][] {
  const result: [number, number][] = [];
  for (const value of indices) {
    if (!Number.isSafeInteger(value) || value < 0 || value >= 1_048_576) throw new SpreadsheetPdfViewError('source_changed');
    const last = result.at(-1);
    const direction = last && last[0] !== last[1] ? Math.sign(last[1] - last[0]) : 0;
    if (last && (value - last[1] === (direction || 1) || (descending && direction === 0 && value === last[1] - 1))) last[1] = value;
    else result.push([value, value]);
    if (result.length > MAX_SPREADSHEET_PDF_VIEW_RUNS) throw new SpreadsheetPdfViewError('too_complex');
  }
  return result;
}

/**
 * The view within its budget, the grid's extras degraded rather than the PDF
 * refused: derived cells past the workbook's budget print as the saved cells
 * they lie over (a conditional format's paint, a table band, a PivotTable
 * report beyond it), and glyphs past theirs (icon-set icons, sparklines,
 * slicer pictures) are left out with the cells and panels that drew them.
 * Hidden rows and the row order are what the PDF prints, never trimmed: a
 * view too large there is still refused (`assertSpreadsheetPdfViewBudget`).
 */
export function fitSpreadsheetPdfViewBudget(view: SpreadsheetPdfExportView): SpreadsheetPdfExportView {
  let cellsLeft = MAX_SPREADSHEET_PDF_DERIVED_CELLS;
  let cellCharsLeft = MAX_SPREADSHEET_PDF_DERIVED_CHARS;
  let glyphsLeft = MAX_SPREADSHEET_PDF_GLYPHS;
  let charsLeft = MAX_SPREADSHEET_PDF_GLYPH_TOTAL_CHARS;
  const sheets = view.sheets.map((sheet) => {
    const cells = sheet.derivedCells ?? [];
    const glyphs = sheet.glyphs ?? [];
    // The glyphs kept: a leading run within the budget, so every index a
    // kept cell or panel names stays the same.
    let kept = 0;
    while (kept < glyphs.length && glyphsLeft > 0) {
      const size = glyphs[kept].length;
      if (size > MAX_SPREADSHEET_PDF_GLYPH_CHARS || size > charsLeft) break;
      charsLeft -= size;
      glyphsLeft -= 1;
      kept += 1;
    }
    const lost = (index: number | undefined) => index !== undefined && index >= kept;
    let keptCells = 0;
    while (keptCells < cells.length && cellsLeft > 0) {
      const size = JSON.stringify(cells[keptCells]).length + 1;
      if (size > cellCharsLeft) break;
      cellCharsLeft -= size;
      cellsLeft -= 1;
      keptCells += 1;
    }
    let trimmedCells = keptCells < cells.length ? cells.slice(0, keptCells) : cells;
    if (kept < glyphs.length) {
      trimmedCells = trimmedCells.map((cell) => {
        if (!lost(cell.icon) && !lost(cell.spark)) return cell;
        const { icon, spark, ...rest } = cell;
        return { ...rest, ...(lost(icon) ? {} : { icon }), ...(lost(spark) ? {} : { spark }) };
      });
    }
    if (trimmedCells === cells && kept === glyphs.length) return sheet;
    const panels = sheet.panels?.filter((panel) => !lost(panel.glyph));
    const { derivedCells: _cells, glyphs: _glyphs, panels: _panels, ...rest } = sheet;
    void _cells; void _glyphs; void _panels;
    return {
      ...rest,
      ...(trimmedCells.length ? { derivedCells: trimmedCells } : {}),
      ...(kept ? { glyphs: glyphs.slice(0, kept) } : {}),
      ...(panels?.length ? { panels } : {}),
    };
  });
  return sheets.every((sheet, index) => sheet === view.sheets[index]) ? view : { ...view, sheets, budgetTrimmed: true };
}

export function assertSpreadsheetPdfViewBudget(view: SpreadsheetPdfExportView): void {
  if (view.sheets.length > 200 || view.sheets.reduce((sum, sheet) => sum + sheet.hiddenRowRanges.length + (sheet.rowOrderRanges?.length ?? 0), 0) > MAX_SPREADSHEET_PDF_VIEW_RUNS) {
    throw new SpreadsheetPdfViewError('too_complex');
  }
  if (view.sheets.reduce((sum, sheet) => sum + (sheet.derivedCells?.length ?? 0), 0) > MAX_SPREADSHEET_PDF_DERIVED_CELLS) {
    throw new SpreadsheetPdfViewError('too_complex');
  }
  const glyphs = view.sheets.flatMap((sheet) => sheet.glyphs ?? []);
  if (glyphs.length > MAX_SPREADSHEET_PDF_GLYPHS || glyphs.some((glyph) => glyph.length > MAX_SPREADSHEET_PDF_GLYPH_CHARS) ||
    glyphs.reduce((sum, glyph) => sum + glyph.length, 0) > MAX_SPREADSHEET_PDF_GLYPH_TOTAL_CHARS) {
    throw new SpreadsheetPdfViewError('too_complex');
  }
}
