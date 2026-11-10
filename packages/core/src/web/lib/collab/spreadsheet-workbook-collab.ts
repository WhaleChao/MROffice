/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PAGE_AUTHORITY_OUTAGE_GRACE_MS } from "./page-writable";
import type {
  WorkbookPresenceInput,
  WorkbookRemotePresence,
} from "./workbook-presence";

/**
 * The only editor-owned wiring contract for Workbook Sharded Collaboration V2.
 * Keeping it outside SpreadsheetWindow lets the Collaboration layer evolve
 * without duplicating workbook semantics in that large UI module.
 */
export const SPREADSHEET_WORKBOOK_COLLAB_OPTIONS = Object.freeze({
  collabJson: true,
  collabJsonStrategy: "workbook-sharded-v2" as const,
  // `sheetsUnfilled` is one save's own mark (buildContent: this window still
  // held a sheet as its placeholder), never the workbook's: it stays out of
  // the shared shards, where every other window would have read it as theirs.
  collabPrivateKeys: Object.freeze(["activeSheetId", "sheetsUnfilled"] as const),
  // A page read that fails while the window is offline keeps the last
  // confirmed verdict a while (as the code and PDF windows do): unconfirmed,
  // the rooms were reopened on reconnect over the window's offline edits.
  pageAuthorityOutageGraceMs: PAGE_AUTHORITY_OUTAGE_GRACE_MS,
});

export function spreadsheetWorkbookPresenceInput(input: {
  readonly sheetId: string;
  readonly active: { readonly row: number; readonly col: number };
  readonly range: {
    readonly r1: number;
    readonly c1: number;
    readonly r2: number;
    readonly c2: number;
  };
  readonly editingCell: {
    readonly sheetId: string;
    readonly row: number;
    readonly col: number;
  } | null;
}): WorkbookPresenceInput {
  return {
    sheetId: input.sheetId,
    active: input.active,
    range: input.range,
    editing:
      input.editingCell?.sheetId === input.sheetId &&
      input.editingCell.row === input.active.row &&
      input.editingCell.col === input.active.col,
  };
}

export function spreadsheetRemotePresenceForSheet(
  remote: readonly WorkbookRemotePresence[],
  sheetId: string,
  rowCount: number,
  columnCount: number,
): WorkbookRemotePresence[] {
  return remote.filter(
    (entry) =>
      entry.sheetId === sheetId &&
      entry.range.r2 < rowCount &&
      entry.range.c2 < columnCount,
  );
}
