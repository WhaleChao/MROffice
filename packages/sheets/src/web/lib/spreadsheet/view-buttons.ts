/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * View > Window > Freeze Panes and Split as pure sheet transactions. Moved out
 * of SpreadsheetWindow so the ribbon and the spreadsheet editor commands run
 * the same code; behaviour is unchanged. The window keeps the notices, the
 * active cell, the on-screen measuring Split starts from and the pane scroll.
 */

import {
  freezeConfigForAction,
  hasFreeze,
  hasSplit,
  readFreezeConfig,
  readSplitConfig,
  normalizeSplitPaneViewState,
  type FreezeAction,
  type SplitPaneViewState,
} from "@mrnote/core/web/lib/spreadsheet/freeze-panes.ts";
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from "@mrnote/core/web/lib/spreadsheet/grid-sizing.ts";

export interface PaneSheet {
  rowCount: number;
  columnCount: number;
  freezeRows?: number;
  freezeCols?: number;
  splitRow?: number;
  splitCol?: number;
  splitPaneState?: SplitPaneViewState;
  rightToLeft?: boolean;
}

/**
 * Apply a Freeze Panes menu action at the active cell. Freeze and split are
 * mutually exclusive: a freeze removes any split so a single divider axis
 * governs the layout, and the split keys are dropped entirely so the
 * no-split save stays byte-identical.
 */
export function frozenSheet<S extends PaneSheet>(
  sheet: S,
  action: FreezeAction,
  active: { row: number; col: number },
): S {
  // Freeze Panes on a split window freezes at the split: everything above
  // and left of the bars (G1 §6). The other commands freeze their lines.
  const freeze = readFreezeConfig({
    freezeRows: sheet.freezeRows,
    freezeCols: sheet.freezeCols,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  });
  const split = hasFreeze(freeze)
    ? { splitRow: 0, splitCol: 0 }
    : readSplitConfig({
        splitRow: sheet.splitRow,
        splitCol: sheet.splitCol,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
      });
  const freezeAt =
    action === "freezePanes" && hasSplit(split)
      ? { row: split.splitRow, col: split.splitCol }
      : { row: active.row, col: active.col };
  // The lines a freeze may hold are the worksheet's, not the rows and
  // columns that happen to be stored: Freeze Top Row on a sheet with one
  // used row still pins row 1 while the empty rows below scroll (live audit
  // B373). The stored extent grows to the frozen band and one line after
  // it, so every reader that clamps a freeze to the sheet agrees.
  const { freezeRows, freezeCols } = freezeConfigForAction(
    action,
    freezeAt,
    { rowCount: MAX_GRID_ROWS, columnCount: MAX_GRID_COLUMNS },
  );
  const next = {
    ...sheet,
    freezeRows,
    freezeCols,
    rowCount: Math.max(sheet.rowCount, freezeRows > 0 ? freezeRows + 1 : 0),
    columnCount: Math.max(sheet.columnCount, freezeCols > 0 ? freezeCols + 1 : 0),
  };
  delete next.splitRow;
  delete next.splitCol;
  delete next.splitPaneState;
  return next;
}

/** The sheet without any split (View > Split on a split window). */
export function splitRemovedSheet<S extends PaneSheet>(sheet: S): S {
  const next = { ...sheet };
  delete next.splitRow;
  delete next.splitCol;
  delete next.splitPaneState;
  return next;
}

/**
 * View > Split placed: bars before row `split.splitRow` and column
 * `split.splitCol` (0 = no bar on that axis), the trailing pane starting at
 * `topLeftCell` (its A1 is the stored pane's topLeftCell). Split and freeze
 * are exclusive: splitting removes the freeze.
 */
export function splitPlacedSheet<S extends PaneSheet>(
  sheet: S,
  split: { splitRow: number; splitCol: number },
  topLeftCell: { row: number; col: number },
  active: { row: number; col: number },
  columnLabel: (index: number) => string,
): S {
  const nextSheet = { ...sheet };
  delete nextSheet.freezeRows;
  delete nextSheet.freezeCols;
  if (split.splitRow > 0) nextSheet.splitRow = split.splitRow;
  else delete nextSheet.splitRow;
  if (split.splitCol > 0) nextSheet.splitCol = split.splitCol;
  else delete nextSheet.splitCol;
  nextSheet.splitPaneState = normalizeSplitPaneViewState(
    { topLeftCell: `${columnLabel(topLeftCell.col)}${topLeftCell.row + 1}` },
    split,
    { rowCount: sheet.rowCount, columnCount: sheet.columnCount },
    { row: active.row, col: active.col },
  );
  return nextSheet;
}

/**
 * A split bar moved (dragged, typed, or removed with 0): the requested lines
 * replace the current ones on their axes. Refused (null) on a frozen sheet or
 * for a line outside the sheet. The pane state follows the new lines.
 */
export function splitLinesSheet<S extends PaneSheet>(
  sheet: S,
  requested: { splitRow?: number; splitCol?: number },
  active: { row: number; col: number },
): S | null {
  const liveFreeze = readFreezeConfig({
    freezeRows: sheet.freezeRows,
    freezeCols: sheet.freezeCols,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  });
  if (hasFreeze(liveFreeze)) return null;
  const liveSplit = readSplitConfig({
    splitRow: sheet.splitRow,
    splitCol: sheet.splitCol,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
  });
  const nextSplit = {
    splitRow: requested.splitRow ?? liveSplit.splitRow,
    splitCol: requested.splitCol ?? liveSplit.splitCol,
  };
  if (
    nextSplit.splitRow < 0 ||
    nextSplit.splitRow >= sheet.rowCount ||
    nextSplit.splitCol < 0 ||
    nextSplit.splitCol >= sheet.columnCount
  ) {
    return null;
  }
  const nextSheet = { ...sheet };
  if (nextSplit.splitRow > 0) nextSheet.splitRow = nextSplit.splitRow;
  else delete nextSheet.splitRow;
  if (nextSplit.splitCol > 0) nextSheet.splitCol = nextSplit.splitCol;
  else delete nextSheet.splitCol;
  if (hasSplit(nextSplit)) {
    nextSheet.splitPaneState = normalizeSplitPaneViewState(
      sheet.splitPaneState,
      nextSplit,
      { rowCount: sheet.rowCount, columnCount: sheet.columnCount },
      { row: active.row, col: active.col },
    );
  } else {
    delete nextSheet.splitPaneState;
  }
  return nextSheet;
}
