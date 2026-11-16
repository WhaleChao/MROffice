/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure helpers for SLICERS (Excel / WPS parity).
 *
 * A slicer is a floating point-and-click filter panel bound to ONE stable
 * Table/Pivot field. It lists that field's distinct display values as toggle
 * chips and writes the SAME native filter authority used by the source object.
 * It is never a second worksheet-column filter system. Legacy source-less
 * entries remain readable for backward compatibility, but new authoring must
 * name a ListObject/tableColumn or Pivot field identity.
 *
 * The sheet model stores a small `slicers` array (one entry per panel). Each
 * slicer carries a display column plus stable source identity, optional title,
 * and position; SELECTED values are DERIVED from that source's filter. A
 * persisted `selected` is tolerated/normalized for forward compatibility but
 * remains advisory — the native filter wins.
 *
 * This module is deterministic, and leans only on the shared pure sort/filter
 * comparator (so a slicer and an AutoFilter list agree), so it can be tested in
 * isolation: it carries the slicer MODEL, the normalize/persist coercion (a
 * persisted value can come from an imported file, an AI write, or a hand-edited
 * JSON, so every field is clamped/dropped rather than trusted), and the three
 * pure mappings the React layer needs — distinct values, selection<->filter, and
 * toggle. The React layer reads the cell values, calls these helpers, and
 * renders the result; it owns no slicer math.
 */

import { orderFilterValues } from "./sort-filter";

/**
 * A sheet-level slicer panel. `column` is the 0-based worksheet column used for
 * placement/backward compatibility; `source` is the stable filter authority.
 * `title` is an optional header label (the renderer falls back
 * to the column header / letter when absent). `selected` is an OPTIONAL cached
 * copy of the chosen display values — empty/absent means "all" (no filter); it
 * is advisory only (the column filter is authoritative), persisted so an
 * importer/AI can round-trip intent. `x`/`y` are the panel's pixel offset within
 * the grid stage (optional; the renderer lays panels out by index when absent).
 */
export interface Slicer {
  id: string;
  column: number;
  /**
   * Stable native filter authority. Legacy source-less slicers remain readable
   * as worksheet-column filters, but new authoring must name a Table/Pivot
   * field before it can claim Excel slicer semantics.
   */
  source?: SlicerSource;
  title?: string;
  selected?: string[];
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  columns?: number;
  buttonHeight?: number;
  styleId?: SlicerStyleId;
  showHeader?: boolean;
}

export type SlicerSource =
  | {
      kind: "table";
      tableId: string;
      tableColumnId: number;
    }
  | {
      kind: "pivot";
      pivotOutput: string;
      fieldId: string;
      /** Which PivotTable of the host sheet; absent means the sheet's first. */
      pivotName?: string;
      /**
       * Further PivotTables the slicer filters (Excel's Slicer > Report
       * Connections), each by host sheet id and name. They read the same
       * PivotCache as the first, and one click filters them all.
       */
      connections?: PivotSlicerConnection[];
    };

/** One more PivotTable a PivotTable slicer filters (Report Connections). */
export interface PivotSlicerConnection {
  pivotOutput: string;
  pivotName: string;
}

/** Report Connections a slicer may carry beyond its first PivotTable. */
export const MAX_PIVOT_SLICER_CONNECTIONS = 63;

export const SLICER_STYLE_IDS = [
  "light1",
  "light2",
  "light3",
  "dark1",
  "dark2",
  "dark3",
] as const;
export type SlicerStyleId = (typeof SLICER_STYLE_IDS)[number];

export const DEFAULT_SLICER_WIDTH = 168;
export const DEFAULT_SLICER_HEIGHT = 220;
export const DEFAULT_SLICER_COLUMNS = 1;
export const DEFAULT_SLICER_BUTTON_HEIGHT = 28;
export const MIN_SLICER_WIDTH = 120;
export const MAX_SLICER_WIDTH = 640;
export const MIN_SLICER_HEIGHT = 120;
export const MAX_SLICER_HEIGHT = 720;
export const MAX_SLICER_COLUMNS = 8;
export const MIN_SLICER_BUTTON_HEIGHT = 20;
export const MAX_SLICER_BUTTON_HEIGHT = 64;










const STORED_SLICER_MIN_PX = 1;
const STORED_SLICER_MAX_PX = 100_000;
const STORED_SLICER_MAX_BUTTON_PX = 10_000;

/** Max slicers a sheet may carry (DoS bound on iteration / overlay render). */
export const MAX_SLICERS = 64;
/** How far one more slicer steps down-right when its spot is taken. */
export const SLICER_CASCADE_OFFSET = 24;

/**
 * Where a new slicer lands. Excel puts a slicer beside the object it filters
 * (the first one to the right of the table, level with its top row) and
 * cascades the next ones a little down and to the right so none covers
 * another (Microsoft Support, "Use slicers to filter data"). `anchor` is that
 * first spot in the grid's pixel frame; a spot within one cascade step of an
 * existing slicer's origin steps on until it is free.
 */
export function cascadeSlicerOrigin(
  existing: readonly Pick<Slicer, "x" | "y">[],
  anchor: { x: number; y: number },
): { x: number; y: number } {
  let x = Math.max(0, Math.round(anchor.x));
  let y = Math.max(0, Math.round(anchor.y));
  const taken = existing.filter(
    (slicer): slicer is { x: number; y: number } =>
      typeof slicer.x === "number" && typeof slicer.y === "number",
  );
  for (let step = 0; step <= MAX_SLICERS; step += 1) {
    const covered = taken.some(
      (slicer) =>
        Math.abs(slicer.x - x) < SLICER_CASCADE_OFFSET &&
        Math.abs(slicer.y - y) < SLICER_CASCADE_OFFSET,
    );
    if (!covered) break;
    x += SLICER_CASCADE_OFFSET;
    y += SLICER_CASCADE_OFFSET;
  }
  return { x, y };
}
/** Max length of a slicer id / title (DoS bound on a hand-edited value). */
export const MAX_SLICER_ID_LENGTH = 200;
export const MAX_SLICER_TITLE_LENGTH = 200;
/** Max distinct values a slicer caches in `selected` (mirrors the filter cap). */
export const MAX_SLICER_SELECTED = 10_000;
/** Max length of a single cached selected value (mirrors the filter cap). */
export const MAX_SLICER_VALUE_LENGTH = 1_000;
/** Pixel clamp for a persisted panel offset so a stale value can't fly off. */
const MAX_SLICER_OFFSET = 100_000;

/** A cell as far as slicer value-enumeration cares: just its display text. */
export interface SlicerCell {
  value?: string;
}

/**
 * A row-major view of the cells a slicer enumerates: a map keyed by the same
 * `"row:col"` ref shape the sheet's `cells` map uses. The display value is the
 * RESOLVED text (formula results already resolved by the caller) so the slicer's
 * values match what the user sees, exactly like the AutoFilter dropdown.
 */
export type SlicerCellMap = Record<string, SlicerCell | undefined>;

/**
 * Compute the ordered set of distinct display values for a column, for the
 * slicer's chip list. Excel treats row 0 as the header, so values are read from
 * the DATA rows (row >= 1) of `[0, rowCount)`. Each value is trimmed; blanks
 * collapse to a single empty-string entry (the renderer shows it as "(Blanks)").
 *
 * The chips come back in Excel's slicer order (Microsoft Support, "Use slicers
 * to filter data"): a slicer sorts its items ascending and keeps the items with
 * no data last, which for a worksheet column is the blanks bucket. That is the
 * SAME order the AutoFilter dropdown lists, through the shared comparator, so
 * a slicer and its column's dropdown read alike.
 *
 * @param cells the sheet's resolved-display cell map (`"row:col"` -> { value })
 * @param columnIndex the 0-based column to enumerate
 * @param rowCount the sheet's row count (the header row 0 is skipped)
 */
export function distinctColumnValues(
  cells: SlicerCellMap,
  columnIndex: number,
  rowCount: number,
): string[] {
  const seen = new Set<string>();
  const order: string[] = [];
  const rows = Number.isFinite(rowCount) ? Math.max(0, Math.floor(rowCount)) : 0;
  for (let r = 1; r < rows; r += 1) {
    const value = (cells[`${r}:${columnIndex}`]?.value ?? "").trim();
    if (!seen.has(value)) {
      seen.add(value);
      order.push(value);
    }
  }
  return orderFilterValues(order);
}

/**
 * Map a slicer's selection to the column-filter representation the existing
 * engine expects — the SAME shape `setColumnFilter` accepts and `applyFilter`
 * materializes: a `string[]` of allowed display values, or `null` for "no
 * filter" (all rows visible).
 *
 *  - empty selection  -> null  (no filter: the column shows every row)
 *  - all values chosen -> null  (everything on === no constraint, Excel parity)
 *  - a proper subset  -> the chosen values, deduped, in `allValues` order
 *
 * `allValues` is the column's distinct value list (from `distinctColumnValues`).
 * Any selected value NOT in `allValues` (e.g. it referenced a now-deleted row)
 * is dropped so the resulting filter can never hide every row via a phantom
 * value. The order follows `allValues` so the persisted filter is deterministic.
 */
export function slicerSelectionToFilter(
  selected: readonly string[],
  allValues: readonly string[],
): string[] | null {
  if (selected.length === 0) return null;
  const chosen = new Set(selected);
  const out: string[] = [];
  for (const value of allValues) {
    if (chosen.has(value)) out.push(value);
  }
  // Everything selected (or no real subset) imposes no constraint -> no filter.
  if (out.length === 0) return null;
  if (out.length === allValues.length) return null;
  return out;
}

/**
 * Derive a slicer's selection FROM the column filter (the source of truth). The
 * filter is the engine's `string[]` of allowed values (or null/undefined when
 * the column has no filter). A column with no filter means "all selected", which
 * the slicer represents as the empty selection (every chip lit). A filter lists
 * exactly the selected values; we intersect with `allValues` (in its order) so a
 * stale value on a since-deleted row never shows as a phantom selected chip.
 */
/**
 * The slicer's state from its column filter: which of the current values are
 * selected AND whether the column is filtered at all. A filter whose values
 * have all disappeared from the column (the only Q1 row was edited to QX)
 * still hides every row; it must read as "filtered, nothing matches", not as
 * "all selected" (live audit B172).
 */
export function slicerSelectionState(
  filter: readonly string[] | null | undefined,
  allValues: readonly string[],
): { selected: string[]; filtered: boolean } {
  return {
    selected: filterToSlicerSelection(filter, allValues),
    filtered: Boolean(filter && filter.length > 0),
  };
}

export function filterToSlicerSelection(
  filter: readonly string[] | null | undefined,
  allValues: readonly string[],
): string[] {
  if (!filter || filter.length === 0) return [];
  const allowed = new Set(filter.map(String));
  const out: string[] = [];
  for (const value of allValues) {
    if (allowed.has(value)) out.push(value);
  }
  return out;
}

/**
 * Toggle a single value in a slicer's selection: remove it when present, add it
 * (appended) when absent. Returns a NEW array (the input is not mutated) so the
 * caller can feed it straight into `slicerSelectionToFilter`. Duplicates in the
 * incoming selection collapse to one (a normalize pass), so toggling is stable.
 */
export function toggleSlicerValue(
  selected: readonly string[],
  value: string,
): string[] {
  const out: string[] = [];
  let removed = false;
  const seen = new Set<string>();
  for (const v of selected) {
    if (v === value) {
      removed = true;
      continue;
    }
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  if (!removed) out.push(value);
  return out;
}

/**
 * Excel slicer click semantics. An ordinary click selects only the clicked
 * item. A Ctrl/Command+click, or any click while the header's Multi-Select
 * button is on (`additive`), toggles the clicked item and leaves every other
 * item as it was. The empty array is the unfiltered state, in which every item
 * is selected, so a toggle from there deselects the clicked item and keeps the
 * rest (`allValues` is the slicer's item list). Deselecting the only selected
 * item selects every item again: the result is empty, i.e. the filter clears.
 */
export function slicerSelectionForWebClick(
  selected: readonly string[],
  value: string,
  additive: boolean,
  allValues: readonly string[],
): string[] {
  if (!additive) return [value];
  return toggleSlicerValue(selected.length === 0 ? allValues : selected, value);
}

/**
 * The slicer header's keyboard shortcuts while focus is inside a slicer:
 * Alt+S toggles Multi-Select and Alt+C clears the filter. The letter is read
 * from `key`; the physical `code` covers macOS Option, which types "ß" / "ç",
 * and non-Latin layouts.
 */
export function slicerHeaderShortcut(
  event: Pick<
    KeyboardEvent,
    "key" | "code" | "altKey" | "ctrlKey" | "metaKey" | "shiftKey" | "isComposing"
  >,
): "multiSelect" | "clearFilter" | null {
  if (
    !event.altKey ||
    event.ctrlKey ||
    event.metaKey ||
    event.shiftKey ||
    event.isComposing
  ) {
    return null;
  }
  const letter = /^[a-z]$/i.test(event.key)
    ? event.key.toLowerCase()
    : event.code === "KeyS"
      ? "s"
      : event.code === "KeyC"
        ? "c"
        : "";
  if (letter === "s") return "multiSelect";
  if (letter === "c") return "clearFilter";
  return null;
}

/** Clamp a persisted panel offset to a finite, non-negative pixel value. */
function clampOffset(raw: unknown): number | undefined {
  if (typeof raw !== "number" || !Number.isFinite(raw)) return undefined;
  const v = Math.max(0, Math.min(MAX_SLICER_OFFSET, Math.round(raw)));
  return v;
}

function clampInteger(
  raw: unknown,
  minimum: number,
  maximum: number,
): number | undefined {
  if (typeof raw !== "number" || !Number.isFinite(raw)) return undefined;
  return Math.max(minimum, Math.min(maximum, Math.round(raw)));
}

/**
 * Coerce a raw persisted entry into a safe Slicer, or null when malformed
 * (missing/over-long id, non-integer or negative column). `selected` is
 * defensively capped/trimmed; `title`/`x`/`y` are dropped unless valid. Extra
 * fields are ignored.
 *
 * @param raw the persisted entry
 * @param columnCount the sheet's column count, so a stale slicer on a deleted
 *   column is dropped (it would otherwise reference past the grid)
 */
export function normalizeSlicer(
  raw: unknown,
  columnCount: number,
): Slicer | null {
  if (!raw || typeof raw !== "object") return null;
  const r = raw as Record<string, unknown>;
  const id = typeof r.id === "string" ? r.id.trim() : "";
  if (id === "" || id.length > MAX_SLICER_ID_LENGTH) return null;
  const column = Number(r.column);
  const cols = Number.isFinite(columnCount) ? Math.floor(columnCount) : 0;
  if (!Number.isInteger(column) || column < 0 || column >= cols) return null;
  const out: Slicer = { id, column };
  if (r.source && typeof r.source === "object" && !Array.isArray(r.source)) {
    const source = r.source as Record<string, unknown>;
    if (
      source.kind === "table" &&
      typeof source.tableId === "string" &&
      source.tableId.trim().length > 0 &&
      source.tableId.trim().length <= MAX_SLICER_ID_LENGTH &&
      Number.isSafeInteger(source.tableColumnId) &&
      Number(source.tableColumnId) > 0
    ) {
      out.source = {
        kind: "table",
        tableId: source.tableId.trim(),
        tableColumnId: Number(source.tableColumnId),
      };
    } else if (
      source.kind === "pivot" &&
      typeof source.pivotOutput === "string" &&
      source.pivotOutput.trim().length > 0 &&
      source.pivotOutput.trim().length <= 64 &&
      typeof source.fieldId === "string" &&
      source.fieldId.trim().length > 0 &&
      source.fieldId.trim().length <= MAX_SLICER_ID_LENGTH
    ) {
      const connections: PivotSlicerConnection[] = [];
      for (const raw of Array.isArray(source.connections)
        ? source.connections.slice(0, MAX_PIVOT_SLICER_CONNECTIONS)
        : []) {
        const entry = raw && typeof raw === "object" ? (raw as Record<string, unknown>) : null;
        const pivotOutput = typeof entry?.pivotOutput === "string" ? entry.pivotOutput.trim() : "";
        const pivotName = typeof entry?.pivotName === "string" ? entry.pivotName.trim() : "";
        if (
          pivotOutput.length > 0 &&
          pivotOutput.length <= 64 &&
          pivotName.length > 0 &&
          pivotName.length <= 255 &&
          !connections.some(
            (known) =>
              known.pivotOutput === pivotOutput &&
              known.pivotName.toLowerCase() === pivotName.toLowerCase(),
          )
        ) {
          connections.push({ pivotOutput, pivotName });
        }
      }
      out.source = {
        kind: "pivot",
        pivotOutput: source.pivotOutput.trim(),
        fieldId: source.fieldId.trim(),
        ...(typeof source.pivotName === "string" &&
        source.pivotName.trim().length > 0 &&
        source.pivotName.trim().length <= 255
          ? { pivotName: source.pivotName.trim() }
          : {}),
        ...(connections.length > 0 ? { connections } : {}),
      };
    }
  }
  if (typeof r.title === "string") {
    const title = r.title.trim();
    if (title !== "" && title.length <= MAX_SLICER_TITLE_LENGTH) {
      out.title = title;
    }
  }
  if (Array.isArray(r.selected)) {
    const selected: string[] = [];
    const seen = new Set<string>();
    for (const value of r.selected) {
      if (selected.length >= MAX_SLICER_SELECTED) break;
      if (typeof value !== "string") continue;
      if (value.length > MAX_SLICER_VALUE_LENGTH) continue;
      if (seen.has(value)) continue;
      seen.add(value);
      selected.push(value);
    }
    if (selected.length > 0) out.selected = selected;
  }
  const x = clampOffset(r.x);
  if (x !== undefined) out.x = x;
  const y = clampOffset(r.y);
  if (y !== undefined) out.y = y;
  const width = clampInteger(r.width, STORED_SLICER_MIN_PX, STORED_SLICER_MAX_PX);
  if (width !== undefined) out.width = width;
  const height = clampInteger(r.height, STORED_SLICER_MIN_PX, STORED_SLICER_MAX_PX);
  if (height !== undefined) out.height = height;
  const columns = clampInteger(r.columns, 1, MAX_SLICER_COLUMNS);
  if (columns !== undefined) out.columns = columns;
  const buttonHeight = clampInteger(
    r.buttonHeight,
    STORED_SLICER_MIN_PX,
    STORED_SLICER_MAX_BUTTON_PX,
  );
  if (buttonHeight !== undefined) out.buttonHeight = buttonHeight;
  if (
    typeof r.styleId === "string" &&
    (SLICER_STYLE_IDS as readonly string[]).includes(r.styleId)
  ) {
    out.styleId = r.styleId as SlicerStyleId;
  }
  if (typeof r.showHeader === "boolean") out.showHeader = r.showHeader;
  return out;
}

/**
 * The identity of one PivotTable slicer's source: the PivotTable (its host
 * sheet plus its name — absent names the sheet's first PivotTable) and the
 * cache field, which Excel matches case-insensitively. Two PivotTables on one
 * sheet may each have a slicer on a field of the same name; only a second
 * slicer on the SAME PivotTable field is a duplicate. Shared by the
 * normalizer and the Insert Slicer writer so both agree.
 */
export function pivotSlicerSourceKey(source: {
  pivotOutput: string;
  pivotName?: string;
  fieldId: string;
}): string {
  return JSON.stringify([
    "pivot",
    source.pivotOutput,
    (source.pivotName ?? "").trim().toLowerCase(),
    source.fieldId.trim().toLowerCase(),
  ]);
}

/**
 * Normalize a raw `slicers` array off a persisted sheet: drop malformed entries,
 * drop a second slicer on the same stable source field (legacy entries still
 * dedupe by worksheet column), and cap the entry count (DoS bound). Returns a
 * fresh array; empty when nothing valid remains so the no-slicer path stays
 * byte-identical.
 */
export function normalizeSlicers(raw: unknown, columnCount: number): Slicer[] {
  if (!Array.isArray(raw)) return [];
  const out: Slicer[] = [];
  const usedSources = new Set<string>();
  const usedIds = new Set<string>();
  for (const entry of raw) {
    if (out.length >= MAX_SLICERS) break;
    const slicer = normalizeSlicer(entry, columnCount);
    if (!slicer) continue;
    const sourceKey =
      slicer.source?.kind === "table"
        ? `table:${slicer.source.tableId}:${slicer.source.tableColumnId}`
        : slicer.source?.kind === "pivot"
          ? pivotSlicerSourceKey(slicer.source)
          : `legacy-column:${slicer.column}`;
    if (usedSources.has(sourceKey)) continue;
    if (usedIds.has(slicer.id)) continue;
    usedSources.add(sourceKey);
    usedIds.add(slicer.id);
    out.push(slicer);
  }
  return out;
}

// ---------------------------------------------------------------------------
// PivotTable slicers
//
// A PivotTable slicer (`source.kind === "pivot"`) names its PivotTable by the
// worksheet that hosts it (`pivotOutput` holds that sheet's id — a worksheet
// holds at most one PivotTable) and its field by the PivotCache field name
// (`fieldId`, the value Excel writes as the slicer cache's `sourceName`). The
// filter itself lives on the PivotTable (`itemFilters`, or `filters` when the
// field sits in the Filters area); the slicer only shows and edits it.
// ---------------------------------------------------------------------------

/** Column of `fieldId` in a PivotCache header row, or null when absent. */
export function resolvePivotSlicerField(
  header: readonly unknown[],
  fieldId: string,
): number | null {
  const wanted = fieldId.trim();
  if (wanted === "") return null;
  let folded: number | null = null;
  for (let index = 0; index < header.length; index += 1) {
    const name = String(header[index] ?? "").trim();
    if (name === wanted) return index;
    // Excel treats cache field names case-insensitively; keep the first match.
    if (folded === null && name.toLowerCase() === wanted.toLowerCase()) {
      folded = index;
    }
  }
  return folded;
}

const PIVOT_SLICER_COLLATOR = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: "base",
});

/**
 * The items a PivotTable slicer shows: the distinct display values of one
 * PivotCache field, in ascending order (Excel's default slicer sort order).
 * The blank item, when present, sorts last as Excel does.
 */
export function pivotSlicerFieldValues(
  matrix: readonly (readonly unknown[])[],
  columnIndex: number,
): string[] {
  const seen = new Set<string>();
  for (let row = 1; row < matrix.length; row += 1) {
    const value = String(matrix[row]?.[columnIndex] ?? "").trim();
    seen.add(value);
  }
  return [...seen].sort((left, right) => {
    if (left === "" || right === "") return left === "" ? (right === "" ? 0 : 1) : -1;
    return PIVOT_SLICER_COLLATOR.compare(left, right);
  });
}
