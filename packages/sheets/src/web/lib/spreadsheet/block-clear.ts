/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for two daily-frequency range operations (Excel/WPS parity) kept
 * out of the React component so the math is unit-tested in isolation:
 *
 *  - BLOCK CLEAR (Delete / Backspace over a selection): wipe the VALUE of every
 *    cell in a rectangle while KEEPING its formatting — exactly Excel's Delete
 *    key, which clears contents (value + formula) but leaves the cell's style,
 *    number format, borders, comment, span and lock flag intact ("Clear
 *    Contents", not "Clear All"). A cell that becomes blank AND carries no other
 *    metadata is pruned from the map so a cleared plain cell is byte-identical to
 *    a never-touched empty cell. A cell that keeps metadata (a style, a comment,
 *    a merge span, an explicit unlock) is retained with an empty value.
 *
 *  - CUT-THEN-PASTE MOVE (Ctrl/Cmd+X then paste): after the cut block is pasted
 *    at the destination, the SOURCE rectangle is cleared — unless it overlaps the
 *    pasted destination, in which case the overlapping cells must be left as the
 *    just-pasted values (clearing them would erase the move). `cutSourceClearSet`
 *    computes exactly which source keys to clear given the destination top-left,
 *    so the component can clear the source in the SAME commit as the paste (one
 *    undo step) without erasing any moved cell.
 *
 * Both reuse the SAME blank-cell pruning rule as the component's
 * `writeCellValueInto`, threaded in as `cellHasMetadata`, so the component and
 * these helpers always agree on when a blanked cell is dropped vs. retained.
 */

import { normalizeCellPhonetic } from "@mrnote/core/web/lib/spreadsheet/cell-phonetic.ts";
import { withRecordFontIfPlain } from "@mrnote/core/web/lib/spreadsheet/cell-record-font.ts";

/** A rectangular cell range in inclusive, 0-based grid coordinates. */
export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/**
 * The minimal structural shape these helpers need off a cell. The component's
 * full SpreadsheetCell satisfies it; tests can pass a bare `{ value, ... }`.
 */
export interface BlockClearCell {
  value: string;
  [key: string]: unknown;
}

/** Normalize a range so r1<=r2 and c1<=c2 (an anchor/active pair may be inverted). */
export function normalizeCellRange(range: CellRange): CellRange {
  return {
    r1: Math.min(range.r1, range.r2),
    r2: Math.max(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    c2: Math.max(range.c1, range.c2),
  };
}

/** True when (row, col) lies inside the (normalized) inclusive rectangle. */
export function rangeContainsCell(
  range: CellRange,
  row: number,
  col: number,
): boolean {
  const r = normalizeCellRange(range);
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

/**
 * Clear the VALUE of every cell inside `range` in a CLONE of `cells`, keeping
 * each cell's formatting/metadata (Excel "Clear Contents"). A cell that becomes
 * blank and carries no other metadata (per `cellHasMetadata`) is removed from the
 * map; one that keeps metadata is retained with `value: ""` and its formula
 * dropped. The input map is never mutated. Returns `{ cells, changed }` so the
 * caller can skip the commit (no undo entry / no autosave) when the selection was
 * already empty — keeping a Delete on blank cells byte-stable.
 *
 * `cellHasMetadata` is the SAME predicate the component uses for a blank write
 * (`hasCellMetadata`), threaded in so this module imports no component internals.
 */
/**
 * The keys of `cells` that fall inside `range`, in row-major order.
 *
 * Both clear paths only ever touch a cell that exists, so walking the map and
 * filtering is exactly equivalent to walking the rectangle -- and it is what
 * keeps Ctrl+A on an imported workbook from costing three and a half seconds:
 * the declared grid is ten million cells, the sheet inside it is rarely more
 * than a few thousand. A small rectangle is walked instead: Delete on one cell of
 * 100,000 read and sorted through every key to find it.
 */
function keysInRange<T>(
  cells: Record<string, T>,
  range: { r1: number; c1: number; r2: number; c2: number },
): string[] {
  if ((range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1) <= 4096) {
    const inside: string[] = [];
    for (let row = range.r1; row <= range.r2; row += 1) {
      for (let col = range.c1; col <= range.c2; col += 1) {
        const key = `${row}:${col}`;
        if (Object.prototype.hasOwnProperty.call(cells, key)) inside.push(key);
      }
    }
    return inside;
  }
  const keys: Array<[number, number, string]> = [];
  for (const key of Object.keys(cells)) {
    const separator = key.indexOf(":");
    if (separator < 0) continue;
    const row = Number(key.slice(0, separator));
    const col = Number(key.slice(separator + 1));
    if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
    if (row < range.r1 || row > range.r2) continue;
    if (col < range.c1 || col > range.c2) continue;
    keys.push([row, col, key]);
  }
  keys.sort((left, right) =>
    left[0] === right[0] ? left[1] - right[1] : left[0] - right[0],
  );
  return keys.map(([, , key]) => key);
}

export function clearRangeContents<T extends BlockClearCell>(
  cells: Record<string, T>,
  range: CellRange,
  cellHasMetadata: (cell: T) => boolean,
): { cells: Record<string, T>; changed: boolean } {
  const r = normalizeCellRange(range);
  const next: Record<string, T> = { ...cells };
  let changed = false;
  for (const key of keysInRange(cells, r)) {
    {
      const existing = next[key];
      if (!existing) continue; // already empty -> nothing to clear
      // Build the cleared cell: blank value, NO formula, but every other field
      // (style/comment/span/lock) preserved.
      const cleared = { ...existing, value: "", kind: "text" } as T;
      delete (cleared as Record<string, unknown>).formula;
      // The cell's link goes with its contents, as under Clear All: Windows
      // Excel's Clear Contents (the Delete key) deletes the hyperlink of a
      // cleared cell and keeps it under Clear Formats (xlsx-rounds
      // probe/n35-p948/a pd3, pd4, pd5; CRITERIA XC-202). The link's blue and
      // underline are the cell's format and stay.
      const hadLink = (existing as Record<string, unknown>).hyperlink !== undefined;
      delete (cleared as Record<string, unknown>).hyperlink;
      delete (cleared as Record<string, unknown>).hyperlinkTooltip;
      delete (cleared as Record<string, unknown>).hyperlinkDisplay;
      // A picture placed in the cell is its content (Place in Cell).
      delete (cleared as Record<string, unknown>).picture;
      delete (cleared as Record<string, unknown>).display;
      delete (cleared as Record<string, unknown>).chartValueCache;
      // The rich text's runs are its text: they go with it (xlsx-rounds R308-X10
      // WALK1: Clear Contents on A3:A4 left both texts' runs on the blanks).
      delete (cleared as Record<string, unknown>).runs;
      const drawn = withRecordFontIfPlain(cleared as { runs?: unknown; style?: unknown });
      (cleared as Record<string, unknown>).style = drawn.style;
      if (drawn.style === undefined) delete (cleared as Record<string, unknown>).style;
      // The readings go with the text; Show Phonetic Field stays the cell's.
      const phonetic = normalizeCellPhonetic(
        (cleared as Record<string, unknown>).phonetic,
        "",
        false,
      );
      if (phonetic) (cleared as Record<string, unknown>).phonetic = phonetic;
      else delete (cleared as Record<string, unknown>).phonetic;
      if (cellHasMetadata(cleared)) {
        // Only count a real change when the cell actually differs (a cell that
        // was already blank with metadata but no formula is unchanged).
        if (existing.value !== "" || existing.formula !== undefined || hadLink) {
          changed = true;
        }
        next[key] = cleared;
      } else {
        // No metadata left -> prune so a cleared plain cell == never-set cell.
        delete next[key];
        changed = true;
      }
    }
  }
  return { cells: next, changed };
}

/**
 * The six Excel/WPS "Clear" variants (the Home > Clear submenu). `contents`
 * matches the Delete key (`clearRangeContents`); the others strip ONE facet of a
 * cell and leave the rest:
 *  - contents:    value + formula + the cell's hyperlink (keep style /
 *                 comment / merge / lock / validation)
 *  - formats:     the whole cell format: style + numberFormat, the merge and
 *                 the Locked flag (keep value / formula / comment / link /
 *                 validation)
 *  - comments:    the threaded comment (keep everything else)
 *  - hyperlinks:  the hyperlink (keep everything else)
 *  - removeHyperlinks: hyperlink + cell formatting of the linked cells only
 *                 (keep value/formula/comment; unlinked cells untouched)
 *  - all:         value + formula + formats (merge and Locked included) +
 *                 comment + hyperlink; the caller also drops validation
 *
 * Excel counts the merge and the protection flags as format: Range.ClearFormats
 * and Range.Clear (what Clear Formats and Clear All run) unmerge the selected
 * merge areas and turn an unlocked cell back to Locked. The merge list and the
 * validation rules live on the sheet, so the caller trims those; this module
 * handles the per-cell part (the unlock, and a span an older import baked onto
 * the anchor cell).
 */
export type ClearVariant =
  | "contents"
  | "formats"
  | "comments"
  | "hyperlinks"
  | "removeHyperlinks"
  | "all";

export interface ClearRangeOptions {
  /**
   * Keep an explicit unlock through Clear Formats / Clear All. The caller sets
   * it while sheet protection is in force: which cells are locked is decided
   * with the sheet unprotected, so a command run under protection must not
   * take the user's own editable cells away from them.
   */
  keepUnlocked?: boolean;
}

/**
 * Apply a Clear variant to a CLONE of a single cell, returning the cleared cell
 * (which may be byte-identical when the facet was already absent). `style` is the
 * whole style bag, so `formats`/`all` drop it entirely — a formats clear removes
 * EVERY visual attribute (Excel's Clear Formats wipes number format + font +
 * borders + fill in one go), matching how `pasteFormatsOnly` REPLACES the style.
 */
function applyClearVariantToCell<T extends BlockClearCell>(
  cell: T,
  variant: ClearVariant,
  options: ClearRangeOptions,
): T {
  // Remove Hyperlinks deletes the selection's hyperlinks (Hyperlinks.Delete)
  // and the format of the cells they sit in; a cell with no link keeps its
  // look, so Ctrl+A then Remove Hyperlinks leaves the rest of the sheet alone.
  if (variant === "removeHyperlinks" && !cell.hyperlink) return cell;
  const next = { ...cell } as Record<string, unknown>;
  if (variant === "contents" || variant === "all") {
    next.value = "";
    next.kind = "text";
    delete next.formula;
    // A picture placed in the cell is its content (Place in Cell).
    delete next.picture;
    // Drop the export display cache too — it baked the now-cleared value/format.
    delete next.display;
    delete next.chartValueCache;
    // The rich text's runs are its text (clearRangeContents).
    delete next.runs;
    if (next.style && typeof next.style === "object") next.style = withRecordFontIfPlain(next as { style?: unknown }).style as typeof next.style;
  }
  if (
    variant === "formats" ||
    variant === "removeHyperlinks" ||
    variant === "all"
  ) {
    // The quote prefix is the entry's text identity, not a format: an '001
    // stays the text 001 (ISTEXT is TRUE, SUM ignores it) after its formats
    // are cleared, as a cell's value type does in Excel (live audit B086).
    const style = next.style as { quotePrefix?: unknown } | undefined;
    const textIdentity = variant !== "all" && style?.quotePrefix === true;
    delete next.style;
    if (textIdentity) next.style = { quotePrefix: true };
    delete next.display; // the display cache also reflects the number format
    // The rich-text runs are formatting too: an imported "alpha" bolded inside
    // "alpha beta" loses its weight with the rest of the cell's formats (live
    // audit B135). A run's other marks stay; runs left with only text go.
    if (Array.isArray(next.runs)) {
      const stripped = next.runs.map((run) => {
        if (!run || typeof run !== "object") return run;
        const rest = { ...(run as Record<string, unknown>) };
        delete rest.style;
        return rest;
      });
      const plain = stripped.every(
        (run) =>
          run !== null &&
          typeof run === "object" &&
          Object.keys(run as Record<string, unknown>).every((key) => key === "text"),
      );
      if (plain) delete next.runs;
      else next.runs = stripped;
    }
  }
  if (variant === "formats" || variant === "all") {
    // The merge and the Locked and Hidden flags are part of the cell format
    // in Excel; under protection they stay (see keepUnlocked).
    delete next.rowspan;
    delete next.colspan;
    if (next.locked === false && !options.keepUnlocked) delete next.locked;
    if (next.formulaHidden === true && !options.keepUnlocked) delete next.formulaHidden;
  }
  if (variant === "comments" || variant === "all") {
    delete next.comment;
  }
  if (variant === "all") {
    // Clear All leaves nothing of the cell, its phonetic guide included.
    delete next.phonetic;
  }
  if (
    variant === "hyperlinks" ||
    variant === "removeHyperlinks" ||
    variant === "all"
  ) {
    delete next.hyperlink;
    // The imported ScreenTip/display extras describe the link — clear together.
    delete next.hyperlinkTooltip;
    delete next.hyperlinkDisplay;
  }
  return next as T;
}

/**
 * Apply a Clear variant across every cell in `range` on a CLONE of `cells`,
 * pruning any cell that ends up blank AND metadata-free (per `cellHasMetadata`)
 * so a fully-cleared plain cell is byte-identical to a never-touched empty cell —
 * the SAME rule `clearRangeContents` and `writeCellValueInto` use. Returns
 * `{ cells, changed }` so the caller skips the commit (no undo / no autosave) when
 * the variant touched nothing. The input map is never mutated.
 *
 * `contents` delegates to `clearRangeContents` (the Delete-key path) so the two
 * never drift; the other variants reuse the same per-cell clone + prune here.
 */
export function clearRangeBy<T extends BlockClearCell>(
  cells: Record<string, T>,
  range: CellRange,
  variant: ClearVariant,
  cellHasMetadata: (cell: T) => boolean,
  options: ClearRangeOptions = {},
): { cells: Record<string, T>; changed: boolean } {
  if (variant === "contents") {
    return clearRangeContents(cells, range, cellHasMetadata);
  }
  const r = normalizeCellRange(range);
  const next: Record<string, T> = { ...cells };
  let changed = false;
  for (const key of keysInRange(cells, r)) {
    {
      const existing = next[key];
      if (!existing) continue; // nothing in this cell to clear
      const cleared = applyClearVariantToCell(existing, variant, options);
      // Compare on the facets this variant can touch so an already-clean cell is
      // not counted as changed (keeps a no-op clear byte-stable / undo-free).
      const facetUnchanged =
        cleared.value === existing.value &&
        cleared.locked === existing.locked &&
        cleared.formulaHidden === existing.formulaHidden &&
        cleared.rowspan === existing.rowspan &&
        cleared.colspan === existing.colspan &&
        (cleared as Record<string, unknown>).formula ===
          (existing as Record<string, unknown>).formula &&
        (cleared as Record<string, unknown>).style ===
          (existing as Record<string, unknown>).style &&
        (cleared as Record<string, unknown>).comment ===
          (existing as Record<string, unknown>).comment &&
        (cleared as Record<string, unknown>).hyperlink ===
          (existing as Record<string, unknown>).hyperlink &&
        (cleared as Record<string, unknown>).hyperlinkTooltip ===
          (existing as Record<string, unknown>).hyperlinkTooltip &&
        (cleared as Record<string, unknown>).hyperlinkDisplay ===
          (existing as Record<string, unknown>).hyperlinkDisplay &&
        // The in-cell formats (a bold word inside the text) and the phonetic
        // guide are facets too: a cell formatted only inside its text was
        // taken as unchanged and kept its bold (live audit S-036).
        (cleared as Record<string, unknown>).picture ===
          (existing as Record<string, unknown>).picture &&
        (cleared as Record<string, unknown>).phonetic ===
          (existing as Record<string, unknown>).phonetic &&
        JSON.stringify((cleared as Record<string, unknown>).runs) ===
          JSON.stringify((existing as Record<string, unknown>).runs);
      if (facetUnchanged) continue;
      changed = true;
      if (cleared.value === "" && !cellHasMetadata(cleared)) {
        delete next[key]; // blank + metadata-free -> prune to empty
      } else {
        next[key] = cleared;
      }
    }
  }
  return { cells: next, changed };
}

/**
 * The set of SOURCE cell keys to clear after a CUT block is pasted at
 * destination top-left (`destRow`, `destCol`). Every non-blank source cell is a
 * clear candidate EXCEPT those that fall inside the pasted destination rectangle
 * (clearing them would erase the just-moved value). The destination rectangle is
 * the source rectangle translated by (destRow - srcR1, destCol - srcC1).
 *
 * `isSourceFilled(row, col)` reports whether the source cell actually held
 * content worth clearing (a blank source cell needs no clear and never enters the
 * set), so a sparse cut block only clears the cells it moved.
 *
 * Returned as an array of "r:c" keys (stable row-major order) so the caller can
 * fold the clears into the same commit as the paste for a single undo step.
 */
export function cutSourceClearKeys(
  source: CellRange,
  destRow: number,
  destCol: number,
  isSourceFilled: (row: number, col: number) => boolean,
): string[] {
  const s = normalizeCellRange(source);
  const dRow = destRow - s.r1;
  const dCol = destCol - s.c1;
  const dest: CellRange = {
    r1: s.r1 + dRow,
    r2: s.r2 + dRow,
    c1: s.c1 + dCol,
    c2: s.c2 + dCol,
  };
  const keys: string[] = [];
  for (let row = s.r1; row <= s.r2; row += 1) {
    for (let col = s.c1; col <= s.c2; col += 1) {
      // A source cell that lands inside the pasted destination keeps its moved
      // value — never clear it.
      if (rangeContainsCell(dest, row, col)) continue;
      if (!isSourceFilled(row, col)) continue;
      keys.push(`${row}:${col}`);
    }
  }
  return keys;
}
