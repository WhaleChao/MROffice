/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Excel's Cut / Copy / Paste of a chart on a worksheet, the part that does
 * not need the window: what the clipboard keeps of a copied chart and the
 * chart a paste makes of it.
 *
 * Excel (Microsoft Support, "Move or copy a chart", "Copy an Excel chart to
 * another Office program"): a pasted chart is a new chart of the workbook
 * with the copied chart's type, size, titles and formatting, still plotting
 * the same cells. A chart pasted on another worksheet keeps reading the
 * worksheet its data is on (`=SERIES(Sheet1!$B$1,...)`), so it is the same
 * chart in a new place, never a chart of the destination sheet's cells. The
 * paste lands with its top-left corner at the active cell and gets a name of
 * its own (Excel numbers it "Chart n"). The system clipboard receives a
 * picture of the chart as well, which every other application pastes; that
 * half lives in `object-clipboard.ts`.
 */

import type { SpreadsheetChart } from "@mrnote/core/web/lib/spreadsheet/chart.ts";

/** A pixel frame in the drawing layer (the grid stage, headers included). */
export interface SheetChartClipboardFrame {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** One copied chart and where it was drawn when it was copied. */
export interface SheetChartClipboardItem {
  chart: SpreadsheetChart;
  frame: SheetChartClipboardFrame;
}

/**
 * The copy the clipboard keeps: a deep snapshot, so editing (or deleting)
 * the original after Copy never changes what Paste makes.
 */
export function sheetChartClipboardSnapshot(chart: SpreadsheetChart): SpreadsheetChart {
  return JSON.parse(JSON.stringify(chart)) as SpreadsheetChart;
}

export interface PastedSheetChartPlacement {
  /** The new chart's identity. */
  id: string;
  /** The frame the paste puts it in (top-left at the active cell). */
  left: number;
  top: number;
  /** The cell it floats over: the active cell of the paste. */
  anchor: { row: number; col: number };
  /** The sheet the chart was copied from. */
  copiedFromSheetId: string;
  /** The sheet it is pasted on. */
  targetSheetId: string;
}

/**
 * The chart one paste makes of a copied chart: a new chart with a new
 * identity (no group membership, no place in the drawing stack yet -- it
 * lands on top -- and no name, which the writer numbers as Excel does),
 * bound to the same data. Its data sheet is the copied chart's own data
 * sheet; it is stated only when that is not the sheet the chart lands on.
 */
export function pastedSheetChart(
  copied: SpreadsheetChart,
  placement: PastedSheetChartPlacement,
): SpreadsheetChart {
  const {
    id: _id,
    groupId: _groupId,
    stackOrder: _stackOrder,
    name: _name,
    hidden: _hidden,
    sourceSheetId,
    ...rest
  } = sheetChartClipboardSnapshot(copied);
  const dataSheetId =
    typeof sourceSheetId === "string" && sourceSheetId.trim()
      ? sourceSheetId.trim()
      : placement.copiedFromSheetId;
  const elsewhere = dataSheetId !== placement.targetSheetId;
  const next: SpreadsheetChart = {
    ...rest,
    id: placement.id,
    anchor: { row: placement.anchor.row, col: placement.anchor.col },
    x: Math.max(0, placement.left),
    y: Math.max(0, placement.top),
    ...(elsewhere ? { sourceSheetId: dataSheetId } : {}),
  };
  // The binding names the same data sheet the flat field does.
  if (next.binding && !next.binding.literal) {
    const binding = { ...next.binding };
    if (elsewhere) binding.sheetId = dataSheetId;
    else delete binding.sheetId;
    next.binding = binding;
  }
  return next;
}

/**
 * Where each copied frame lands: the copy keeps its size and its members
 * keep their places relative to one another, the copy's top-left corner at
 * `target` (the active cell's top-left), as Excel pastes a drawing.
 */
export function pastedClipboardFrames(
  frames: readonly SheetChartClipboardFrame[],
  target: { left: number; top: number },
): SheetChartClipboardFrame[] {
  if (frames.length === 0) return [];
  const left = Math.min(...frames.map((frame) => frame.left));
  const top = Math.min(...frames.map((frame) => frame.top));
  return frames.map((frame) => ({
    left: target.left + (frame.left - left),
    top: target.top + (frame.top - top),
    width: frame.width,
    height: frame.height,
  }));
}
