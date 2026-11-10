/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Maximum sparse row history per immutable cell map before a full refresh. */
export const SPREADSHEET_ROW_REVISION_LIMIT = 1024;

interface RowRevisions {
  readonly base: number;
  readonly rows: ReadonlyMap<number, number>;
}

// Shared across registries so replacing a registry cannot reuse a visible token.
let nextRevision = 0;
function freshRevision(): number {
  if (nextRevision >= Number.MAX_SAFE_INTEGER) throw new Error("Row revision exhausted");
  return ++nextRevision;
}
function freshRows(): RowRevisions {
  return { base: freshRevision(), rows: new Map() };
}

/**
 * Rendering hints for immutable cell maps, independent of calculation journals.
 * Values contain numbers only: no source, sheet, parent receipt or closure keeps
 * an obsolete model alive. Each source receives one immutable receipt, including
 * sources first encountered during an abandoned render.
 *
 * This does not prove that an edit is local. The caller must omit knownLiteralKeys
 * unless they cover EVERY changed cell and all cross-row effects are excluded
 * (formulas, merges, conditional formats, filters, layout, permissions, etc.).
 * Other render context must remain a separate conservative invalidation key.
 */
export class SpreadsheetRowRevisions {
  #sources = new WeakMap<object, RowRevisions>();

  #receipt(source: object): RowRevisions {
    let receipt = this.#sources.get(source);
    if (!receipt) {
      receipt = freshRows();
      this.#sources.set(source, receipt);
    }
    return receipt;
  }

  revision(source: object, row: number): number {
    if (!Number.isSafeInteger(row) || row < 0) throw new RangeError("Invalid row");
    const receipt = this.#receipt(source);
    return receipt.rows.get(row) ?? receipt.base;
  }

  /** Call for accepted immutable edits, before rendering the new cell map. */
  record(before: object, after: object, knownLiteralKeys?: readonly string[] | null): void {
    // Never revise a source already observed, including Undo/Redo destinations.
    if (before === after || this.#sources.has(after)) return;
    if (!knownLiteralKeys) {
      this.#sources.set(after, freshRows());
      return;
    }
    const dirtyRows = new Set<number>();
    for (const key of knownLiteralKeys) {
      const match = /^(0|[1-9]\d*):(0|[1-9]\d*)$/.exec(key);
      const row = match ? Number(match[1]) : NaN;
      const col = match ? Number(match[2]) : NaN;
      if (!Number.isSafeInteger(row) || !Number.isSafeInteger(col)) {
        this.#sources.set(after, freshRows());
        return;
      }
      dirtyRows.add(row);
      // The preceding row may paint the changed row's shared border.
      if (row > 0) dirtyRows.add(row - 1);
      if (dirtyRows.size > SPREADSHEET_ROW_REVISION_LIMIT) {
        this.#sources.set(after, freshRows());
        return;
      }
    }
    const previous = this.#receipt(before);
    if (dirtyRows.size === 0) {
      this.#sources.set(after, previous);
      return;
    }
    const rows = new Map(previous.rows);
    const revision = freshRevision();
    for (const row of dirtyRows) rows.set(row, revision);
    // Bound copying and retained history regardless of workbook size or session
    // length. A fresh base safely invalidates every row when the bound is hit.
    this.#sources.set(after, rows.size > SPREADSHEET_ROW_REVISION_LIMIT
      ? freshRows()
      : { base: previous.base, rows });
  }
}
