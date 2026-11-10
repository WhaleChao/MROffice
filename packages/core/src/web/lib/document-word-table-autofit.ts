/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How Word sizes and places an imported table, beyond its stored grid
 * (`source.wordTableLayout`, projected by the server from the source table).
 *
 * - `contentFit`: Word sizes the columns from their content on every layout
 *   (AutoFit to contents) and ignores `w:tblGrid`; `"pct"` tables also have a
 *   percentage preferred width (`contentFitPct`, fiftieths of a percent).
 * - `textAligned`: the document is in Word 2010 (or earlier) compatibility,
 *   so the table's first column starts one left cell margin before its
 *   indent and the cell text lines up with the surrounding paragraphs.
 *   Word 2013 compatibility puts the table's outer border edge at the
 *   indent instead.
 */
export interface WordTableLayout {
  readonly contentFit: "auto" | "pct" | null;
  readonly contentFitPct: number | null;
  readonly textAligned: boolean;
  /**
   * A percentage preferred width (fiftieths of a percent) on a table whose
   * columns keep their preferred widths: Word scales those columns so the
   * table is that share of the text column.
   */
  readonly preferredWidthPct?: number;
  





  readonly preferredWidthTwips?: number;
  /**
   * `w:tblLayout w:type="fixed"`: the grid keeps its own widths and a grid
   * wider than the text column runs into the right margin, as Word draws it.
   */
  readonly fixedLayout?: true;
  /**
   * `w:bidiVisual`: the columns run from the right, the first cell at the
   * right edge; the text in them keeps its own direction.
   */
  readonly bidiVisual?: true;
}

function preferredWidthPct(raw: Record<string, unknown>): { preferredWidthPct?: number } {
  const value = raw.preferredWidthPct;
  // Over 5000 (100%) the table runs past the column, as Word draws it (R282-D21).
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 20000
    ? { preferredWidthPct: value }
    : {};
}

function preferredWidthTwips(raw: Record<string, unknown>): { preferredWidthTwips?: number } {
  const value = raw.preferredWidthTwips;
  return typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 31680
    ? { preferredWidthTwips: value }
    : {};
}

function record(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

export function readWordTableLayout(tableAttrs: {
  readonly source?: unknown;
}): WordTableLayout | null {
  const source = record(tableAttrs.source);
  if (source?.adapter !== "docx") return null;
  const raw = record(source.wordTableLayout);
  if (!raw || raw.schemaVersion !== 1) return null;
  const { contentFit, contentFitPct, textAligned } = raw;
  if (
    !(contentFit === null || contentFit === "auto" || contentFit === "pct") ||
    typeof textAligned !== "boolean"
  ) {
    return null;
  }
  const bidi = raw.bidiVisual === true ? { bidiVisual: true as const } : {};
  if (contentFit === "pct") {
    if (
      typeof contentFitPct !== "number" ||
      !Number.isInteger(contentFitPct) ||
      contentFitPct < 1 ||
      contentFitPct > 5000
    ) {
      return null;
    }
    return { contentFit, contentFitPct, textAligned, ...bidi };
  }
  if (contentFitPct !== null) return null;
  return {
    contentFit,
    contentFitPct: null,
    textAligned,
    ...bidi,
    ...(contentFit === null ? preferredWidthPct(raw) : {}),
    ...(contentFit === null ? preferredWidthTwips(raw) : {}),
    ...(contentFit === null && raw.fixedLayout === true ? { fixedLayout: true as const } : {}),
  };
}

export interface WordAutofitColumn {
  /** Narrowest content box: the longest unbreakable piece plus cell margins. */
  readonly minPx: number;
  /** Widest content box: the widest one-line content plus cell margins. */
  readonly maxPx: number;
  /** The cell margins and borders inside `minPx` (they never shrink). */
  readonly chromePx?: number;
}

export type WordAutofitTableWidth =
  | { readonly type: "auto" }
  | { readonly type: "pct"; readonly fraction: number };

function finiteNonNegative(value: number): number {
  return Number.isFinite(value) && value > 0 ? value : 0;
}

export function wordAutofitColumnWidths(
  columns: readonly WordAutofitColumn[],
  availablePx: number,
  tableWidth: WordAutofitTableWidth = { type: "auto" },
): number[] {
  if (columns.length === 0) return [];
  const available = finiteNonNegative(availablePx);
  const min = columns.map((column) => finiteNonNegative(column.minPx));
  const max = columns.map((column, index) => Math.max(min[index], finiteNonNegative(column.maxPx)));
  const sumMin = min.reduce((sum, value) => sum + value, 0);
  const sumMax = max.reduce((sum, value) => sum + value, 0);

  if (tableWidth.type === "pct") {
    // A percentage narrower than the columns' narrowest contents grows to
    // them while the column has room: Word cuts no word to keep the
    // preference (Word 16.86, mode 15, corpus R83-D04 CN12a: a 25% table of
    // three columns drew 40.3 / 88.6 / 57.6pt, each its longest word plus its
    // margins, 186.5pt for the 112.8pt it asked).
    const target = Math.max(
      available * Math.min(1, Math.max(0, tableWidth.fraction)),
      Math.min(sumMin, available),
    );
    if (sumMax <= target && sumMax > 0) {
      return max.map((value) => (value * target) / sumMax);
    }
    return distribute(min, max, sumMin, sumMax, target, columns);
  }
  if (sumMax <= available) return max;
  return distribute(min, max, sumMin, sumMax, available, columns);
}

























export function wordMinContentFloorWidths(
  widths: readonly number[],
  columns: readonly WordAutofitColumn[],
  drawnTotal?: number,
  growTo?: number,
): number[] | null {
  if (widths.length === 0 || widths.length !== columns.length) return null;
  const grid = widths.map(finiteNonNegative);
  const min = columns.map((column) => finiteNonNegative(column.minPx));
  const asked = grid.reduce((sum, value) => sum + value, 0);
  const shrunk = drawnTotal !== undefined && drawnTotal > 0 && drawnTotal < asked;
  if (!shrunk && !min.some((value, index) => value > grid[index] + 0.5)) return null;
  const floors = grid.map((value, index) => Math.max(value, min[index]));
  // `growTo`: the table states no width of its own (w:tblW auto), so the
  // widened column widens the table and the others keep theirs, as far as the
  // room it stands in; past that the others give the rest back as above.
  // Word 16.86 fixtures (mode 15, a 9026-twip column, "Incomprehensibilities"
  // of 2270 twips with its margins in a 1000-twip cell; corpus
  // docx-tablewidths/fixtures-1110): tblW auto over 1000 + 3000 draws
  // 2270 + 3000 (b1), over 1000 + 3000 + 3000 draws 2270 + 3000 + 3000 (b4),
  // centred the same (b5); over 1000 + 8026, the whole column, 2270 + 6744
  // (b2); tblW 4000 dxa over 1000 + 3000 keeps 4000: 2270 + 1728 (b3). Corpus
  // R282-D29 (auto, centred, mode 14): 1020 widened to 1062, the others kept.
  const floorTotal = floors.reduce((sum, value) => sum + value, 0);
  const target = shrunk
    ? (drawnTotal ?? asked)
    : growTo !== undefined && growTo > asked
      ? Math.min(floorTotal, growTo)
      : asked;
  const excess = floorTotal - target;
  const slack = floors.map((value, index) => (min[index] > grid[index] ? 0 : value - min[index]));
  const totalSlack = slack.reduce((sum, value) => sum + value, 0);
  if (totalSlack >= excess) {
    return floors.map((value, index) => value - (totalSlack > 0 ? (slack[index] * excess) / totalSlack : 0));
  }
  const sumMin = min.reduce((sum, value) => sum + value, 0);
  const max = columns.map((column, index) => Math.max(min[index], finiteNonNegative(column.maxPx)));
  return distribute(min, max, sumMin, max.reduce((sum, value) => sum + value, 0), target, columns);
}














export function wordAdjacentFloorWidths(
  tables: readonly { readonly grid: readonly number[]; readonly columns: readonly WordAutofitColumn[] }[],
): number[][] | null {
  const totals = tables.map((table) => table.grid.reduce((sum, width) => sum + finiteNonNegative(width), 0));
  if (tables.length < 2 || totals.some((total) => !(total > 0) || Math.abs(total - totals[0]) > 1)) return null;
  if (tables.some((table) => table.grid.length !== table.columns.length)) return null;
  const edgesOf = tables.map((table) => {
    let at = 0;
    return table.grid.map((width) => (at += finiteNonNegative(width)));
  });
  const union: number[] = [];
  for (const edge of edgesOf.flat().sort((a, b) => a - b)) {
    if (!union.length || edge - union[union.length - 1] > 0.5) union.push(edge);
  }
  const edgeIndex = (edge: number) => union.findIndex((value) => Math.abs(value - edge) <= 0.5);
  const widths = union.map((edge, index) => edge - (index > 0 ? union[index - 1] : 0));
  const cells = tables.flatMap((table, tableIndex) => table.columns.map((column, index) => ({
    from: index > 0 ? edgeIndex(edgesOf[tableIndex][index - 1]) + 1 : 0,
    to: edgeIndex(edgesOf[tableIndex][index]),
    minPx: finiteNonNegative(column.minPx),
    maxPx: finiteNonNegative(column.maxPx),
  })));
  // Narrow cells first: a wide one then asks only for what they leave it.
  cells.sort((a, b) => a.to - a.from - (b.to - b.from));
  const columns = widths.map(() => ({ minPx: 0, maxPx: 0 }));
  for (const cell of cells) {
    const span = columns.slice(cell.from, cell.to + 1);
    const room = widths.slice(cell.from, cell.to + 1).reduce((sum, width) => sum + width, 0);
    const held = span.reduce((sum, column) => sum + column.minPx, 0);
    if (cell.minPx <= held || !(room > 0)) continue;
    span.forEach((column, offset) => {
      column.minPx += ((cell.minPx - held) * widths[cell.from + offset]) / room;
    });
  }
  for (const [index, column] of columns.entries()) column.maxPx = Math.max(column.minPx, widths[index]);
  const shared = wordMinContentFloorWidths(widths, columns);
  if (!shared) return null;
  return tables.map((_, tableIndex) => {
    let from = 0;
    return edgesOf[tableIndex].map((edge) => {
      const to = edgeIndex(edge);
      const width = shared.slice(from, to + 1).reduce((sum, value) => sum + value, 0);
      from = to + 1;
      return width;
    });
  });
}

/** The table fills `target`; each column keeps its narrowest content and
 * shares the rest in proportion to how much wider its content can be. */
function distribute(
  min: readonly number[],
  max: readonly number[],
  sumMin: number,
  sumMax: number,
  target: number,
  columns: readonly WordAutofitColumn[],
): number[] {
  // Narrower than every unbreakable piece: each column keeps its margins
  // and borders and its narrowest content shrinks by one factor.
  if (target <= sumMin) {
    const chrome = columns.map((column, index) => Math.min(min[index], finiteNonNegative(column.chromePx ?? 0)));
    const sumChrome = chrome.reduce((sum, value) => sum + value, 0);
    const content = sumMin - sumChrome;
    if (!(content > 0) || target <= sumChrome) return [...min];
    const factor = (target - sumChrome) / content;
    return min.map((value, index) => chrome[index] + (value - chrome[index]) * factor);
  }
  const flexible = sumMax - sumMin;
  if (!(flexible > 0)) return [...min];
  const ratio = (target - sumMin) / flexible;
  return min.map((value, index) => value + (max[index] - value) * ratio);
}
