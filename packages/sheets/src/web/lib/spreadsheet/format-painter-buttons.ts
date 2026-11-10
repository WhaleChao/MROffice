/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home > Format Painter on cells: what the brush picks up from a source range
 * and what painting it onto a destination writes (styles tiled over the
 * destination, the unlocked flag, and the conditional formats that cover the
 * source). Moved out of SpreadsheetWindow so the window and the spreadsheet
 * editor commands share it; behaviour is unchanged.
 */

import {
  copyCellMap,
  hasCellMetadata,
  type CellButtonsSheet,
  type SpreadsheetCell,
} from "@mrnote/core/web/lib/spreadsheet/cell-buttons.ts";
import {
  planRulesForFormatPainter,
  snapshotRulesForFormatPaste,
  type ConditionalFormatRule,
} from "@mrnote/core/web/lib/spreadsheet/conditional-format.ts";
import { styleWithOwnTextIdentity } from "@mrnote/core/web/lib/spreadsheet/entry-parse.ts";
import {
  formatPainterStyleAt,
  formatPainterStylesEqual,
  snapshotStyleMatrix,
  type FormatPainterStyleMatrix,
} from "./format-painter";
import { MAX_GRID_COLUMNS as MAX_SHEET_COLS, MAX_GRID_ROWS as MAX_SHEET_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** What the brush holds after picking up a range of cells. */
export type FormatPainterCellsSnapshot = NonNullable<ReturnType<typeof formatPainterCellsSnapshot>>;

export function formatPainterCellsSnapshot(
  sheet: CellButtonsSheet,
  range: Rect,
  conditionalFormats: readonly ConditionalFormatRule[],
) {
  const sourceRange = Object.freeze({ ...range });
  const matrix = snapshotStyleMatrix(
    sourceRange,
    { rows: sheet.rowCount, columns: sheet.columnCount },
    (row, column) => sheet.cells[`${row}:${column}`]?.style,
  );
  // Protection is one of the cell's formats (Format Cells > Protection):
  // an unlocked source makes its target unlocked, as Paste Formatting does.
  const unlocked = Object.freeze(
    Array.from({ length: sourceRange.r2 - sourceRange.r1 + 1 }, (_, dr) =>
      Object.freeze(
        Array.from(
          { length: sourceRange.c2 - sourceRange.c1 + 1 },
          (_, dc) =>
            sheet.cells[`${sourceRange.r1 + dr}:${sourceRange.c1 + dc}`]?.locked === false,
        ),
      ),
    ),
  );
  const hidden = Object.freeze(
    Array.from({ length: sourceRange.r2 - sourceRange.r1 + 1 }, (_, dr) =>
      Object.freeze(
        Array.from(
          { length: sourceRange.c2 - sourceRange.c1 + 1 },
          (_, dc) =>
            sheet.cells[`${sourceRange.r1 + dr}:${sourceRange.c1 + dc}`]?.formulaHidden === true,
        ),
      ),
    ),
  );
  return matrix
    ? {
        kind: "cells" as const,
        sourceRange,
        conditionalFormats: snapshotRulesForFormatPaste(
          conditionalFormats,
          sourceRange,
        ),
        matrix,
        unlocked,
        hidden,
      }
    : null;
}

/** What a paint reads of the brush: the window's own snapshot (read-only
 *  rules) and {@link formatPainterCellsSnapshot}'s both qualify. */
export interface FormatPainterPaintSource {
  readonly matrix: FormatPainterStyleMatrix;
  readonly unlocked: readonly (readonly boolean[])[];
  readonly sourceRange: Readonly<Rect>;
  readonly conditionalFormats: readonly ConditionalFormatRule[];
  readonly hidden?: readonly (readonly boolean[])[];
}

/**
 * The sheet after painting `snapshot` onto `destination`; the same sheet when
 * nothing changes, null when the paint is refused (a style the matrix cannot
 * give, or a conditional format plan that does not fit).
 */
export function sheetAfterFormatPaint<S extends CellButtonsSheet & { conditionalFormats?: ConditionalFormatRule[] }>(
  sheet: S,
  snapshot: FormatPainterPaintSource,
  destination: Rect,
  conditionalFormats: readonly ConditionalFormatRule[],
): S | null {
  const conditionalFormatPlan = planRulesForFormatPainter(
    conditionalFormats,
    snapshot.conditionalFormats,
    snapshot.sourceRange,
    destination,
    { maxRow: MAX_SHEET_ROWS - 1, maxCol: MAX_SHEET_COLS - 1 },
  );
  if (!conditionalFormatPlan) return null;
  const nextCells = copyCellMap(sheet.cells);
  let cellsChanged = false;
  for (let row = destination.r1; row <= destination.r2; row += 1) {
    for (let column = destination.c1; column <= destination.c2; column += 1) {
      const painted = formatPainterStyleAt(
        snapshot.matrix,
        row - destination.r1,
        column - destination.c1,
      );
      if (!painted) return null;
      const key = `${row}:${column}`;
      const previous = nextCells[key];
      // The look travels; the quote prefix (the entry's text identity) does
      // not, and the target keeps its own (live audit B086).
      const style = styleWithOwnTextIdentity(painted, previous?.style) ?? {};
      const nextStyle =
        Object.keys(style).length > 0
          ? (style as NonNullable<SpreadsheetCell["style"]>)
          : undefined;
      const sourceRows = snapshot.unlocked.length;
      const sourceColumns = snapshot.unlocked[0]?.length ?? 0;
      const unlocked =
        sourceRows > 0 && sourceColumns > 0
          ? snapshot.unlocked[(row - destination.r1) % sourceRows][
              (column - destination.c1) % sourceColumns
            ]
          : previous?.locked === false;
      const hidden =
        sourceRows > 0 && sourceColumns > 0 && snapshot.hidden
          ? snapshot.hidden[(row - destination.r1) % sourceRows][
              (column - destination.c1) % sourceColumns
            ] === true
          : previous?.formulaHidden === true;
      if (
        formatPainterStylesEqual(previous?.style, nextStyle) &&
        (previous?.locked === false) === unlocked &&
        (previous?.formulaHidden === true) === hidden
      ) {
        continue;
      }
      cellsChanged = true;
      const next: SpreadsheetCell = {
        ...(previous ?? { value: "", kind: "text" as const }),
      };
      if (nextStyle) next.style = nextStyle;
      else delete next.style;
      if (unlocked) next.locked = false;
      else delete next.locked;
      if (hidden) next.formulaHidden = true;
      else delete next.formulaHidden;
      if (next.value || hasCellMetadata(next)) nextCells[key] = next;
      else delete nextCells[key];
    }
  }
  if (cellsChanged || conditionalFormatPlan.changed) {
    return {
      ...sheet,
      cells: nextCells,
      rowCount: Math.max(sheet.rowCount, destination.r2 + 1),
      columnCount: Math.max(sheet.columnCount, destination.c2 + 1),
      ...(conditionalFormatPlan.changed
        ? { conditionalFormats: conditionalFormatPlan.rules }
        : {}),
    };
  }
  return sheet;
}
