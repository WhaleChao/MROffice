/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { cellMapPermutedSourceKey, cellMapRowPermutation, nearestKnownCellMap, permutedCellKeys, permutedSpanOf, type PermutedSpan } from "@mrnote/core/web/lib/spreadsheet/cell-map-lineage.ts";
import { hasManualLineBreak } from "./cell-text";
import { hasNote, isThreadComment, type CellComment, type ThreadComment } from "@mrnote/core/web/lib/spreadsheet/comments.ts";
import { createSelectionObserver, sameObservation, type SelectionObservationRange, type StoredSelectionSummary } from "./selection-observations";

interface IndexedCell {
  readonly comment?: CellComment;
  readonly hyperlink?: unknown;
  readonly formula?: unknown;
  readonly value?: string;
  readonly display?: string;
  readonly style?: { readonly wrapText?: boolean; readonly backgroundColor?: string; readonly numberFormat?: unknown; readonly fontSize?: number | false; readonly textRotation?: unknown; readonly verticalText?: boolean;
    readonly borderEdges?: { readonly top?: { readonly style?: unknown }; readonly bottom?: { readonly style?: unknown } } };
  readonly locked?: unknown;
  readonly rowspan?: number;
  readonly colspan?: number;
  readonly phonetic?: unknown;
}
type Cells = Readonly<Record<string, IndexedCell | undefined>>;
interface IndexedSpan {
  readonly row: number;
  readonly column: number;
  readonly rowSpan: number;
  readonly colSpan: number;
}
interface IndexedFacets {
  readonly row: number;
  readonly column: number;
  readonly comments: boolean;
  readonly hyperlinks: boolean;
}
/** The font the default row height is made for; a taller font grows the row
 * even for a one-line literal, as Excel's automatic row height does. */
export const DEFAULT_CELL_FONT_SIZE_PT = 11;
export function cellFontTallerThanDefault(style: { readonly fontSize?: number | false } | undefined): boolean {
  return typeof style?.fontSize === "number" && style.fontSize > DEFAULT_CELL_FONT_SIZE_PT;
}
/** Whether a cell's text wraps at its column: Wrap Text, or a horizontal
 * Justify / Distributed alignment, which wraps whether or not Wrap Text is on
 * (Excel; the grid draws them so, XLS-A037/A038) -- and so grows its row. */
export function cellTextWraps(
  style: { readonly wrapText?: boolean; readonly horizontalJustify?: unknown } | undefined,
): boolean {
  return Boolean(style?.wrapText) || style?.horizontalJustify === "justify" || style?.horizontalJustify === "distributed";
}
/** Turned or stacked text (a non-zero Orientation angle, Vertical Text): the
 * row grows to the turned line's height, so the cell's fill stands behind all
 * of it (Excel). */
export function cellTextTurned(
  style: { readonly textRotation?: unknown; readonly verticalText?: boolean } | undefined,
): boolean {
  if (style?.verticalText === true) return true;
  const angle = Number(style?.textRotation ?? 0);
  return Number.isFinite(angle) && Math.round(angle) !== 0;
}
function autoHeightCandidate(cell: IndexedCell | undefined): boolean {
  return Boolean(cellTextWraps(cell?.style) || cell?.value?.startsWith("=") ||
    hasManualLineBreak(cell?.value) || hasManualLineBreak(cell?.display) ||
    cellFontTallerThanDefault(cell?.style) || cellTextTurned(cell?.style));
}
/** Any cell a formula scan may care about; consumers re-check the current cell. */
const formulaCandidate = (cell: IndexedCell | undefined) =>
  Boolean(cell?.formula) || Boolean(cell?.value?.startsWith("="));
/** Cells with a stored fill; sort/filter-by-color reads the current color. */
const fillCandidate = (cell: IndexedCell | undefined) => Boolean(cell?.style?.backgroundColor);
/** Cells carrying phonetic data (furigana); the row band reads the current cell. */
const phoneticCandidate = (cell: IndexedCell | undefined) => typeof cell?.phonetic === "object" && cell.phonetic !== null;
/** Excel's width of a border line style in grid pixels: the medium styles
 * draw 2, thick and double 3, every other 1; no line, 0. */
export function borderLineWidthPx(style: unknown): number {
  if (typeof style !== "string" || style === "" || style === "none") return 0;
  if (style === "thick" || style === "double") return 3;
  return style.startsWith("medium") || style === "slantDashDot" ? 2 : 1;
}
/** Cells with a top or bottom line of 2 px or more: the lines that make a
 * Windows Excel row taller (officeBorderRowExtras). */
const heavyEdgeCandidate = (cell: IndexedCell | undefined) =>
  borderLineWidthPx(cell?.style?.borderEdges?.top?.style) >= 2 || borderLineWidthPx(cell?.style?.borderEdges?.bottom?.style) >= 2;
/** Cells the persisted display cache can touch: a formula or a number format
 * earns a cache, and a stray cache on any other cell is stripped. */
const displayCacheCandidate = (cell: IndexedCell | undefined) => {
  if (!cell) return false;
  const format = cell.style?.numberFormat;
  return Boolean(cell.formula) || (typeof format === "string" && format.trim() !== "") ||
    "display" in cell || "chartValueCache" in cell;
};
const rowSpanOf = (cell: IndexedCell | undefined) => Math.max(1, Number(cell?.rowspan || 1));
const colSpanOf = (cell: IndexedCell | undefined) => Math.max(1, Number(cell?.colspan || 1));
/** A stored nonblank string value: what the used-range walk counts. */
const nonblank = (cell: IndexedCell | undefined) => typeof cell?.value === "string" && cell.value.trim() !== "";
/** The lenient parser of the large-selection observation walk. */
function parseLenient(key: string): [number, number] | null {
  const separator = key.indexOf(":");
  if (separator <= 0 || separator === key.length - 1) return null;
  const row = Number(key.slice(0, separator)), col = Number(key.slice(separator + 1));
  return Number.isSafeInteger(row) && Number.isSafeInteger(col) ? [row, col] : null;
}
/** The canonical parser of storedUsedRangeRect. */
function parseCanonical(key: string): [number, number] | null {
  const parsed = parseLenient(key);
  return parsed && parsed[0] >= 0 && parsed[1] >= 0 && key === `${parsed[0]}:${parsed[1]}` ? parsed : null;
}
const expand = (range: SelectionObservationRange | null, row: number, col: number): SelectionObservationRange =>
  range === null ? { r1: row, r2: row, c1: col, c2: col }
    : row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2 ? range
      : { r1: Math.min(range.r1, row), r2: Math.max(range.r2, row), c1: Math.min(range.c1, col), c2: Math.max(range.c2, col) };
interface IndexedThread {
  readonly key: string;
  readonly row: number;
  readonly col: number;
  readonly comment: ThreadComment;
}
export interface SpreadsheetCellIndex {
  readonly autoHeightKeys: readonly string[];
  readonly formulaKeys: readonly string[];
  readonly fillKeys: readonly string[];
  readonly displayCacheKeys: readonly string[];
  /** Every stored cell carrying phonetic data. */
  readonly phoneticKeys: readonly string[];
  /** Every stored cell with a top or bottom border of 2 px or more. */
  readonly heavyEdgeKeys: readonly string[];
  readonly spans: readonly IndexedSpan[];
  readonly facets: readonly IndexedFacets[];
  readonly notes: number;
  readonly threads: readonly IndexedThread[];
  /** Bounds of the canonical nonblank stored values (storedUsedRangeRect), null
   * when there are none, undefined once an accepted edit removed or blanked a
   * nonblank cell: the consumer then computes it from the map. */
  readonly usedRange: SelectionObservationRange | null | undefined;
  /** Every stored cell observed once for the formatting families. */
  readonly observation: StoredSelectionSummary;
}

/** One sparse index for cell metadata previously found in separate full scans.
 * formulaKeys names every stored cell that carries a formula field or an "="
 * value and fillKeys every cell with a stored fill; scans that only care about
 * those cells walk the list instead of the map.
 * Values retain only relevant keys, numeric spans and comment data, never the
 * source map or parent index. Accepted edits reuse unchanged metadata; unknown
 * sources or changed metadata rebuild in their current own-key order. */
/** Changed keys an index is brought up to date at, past which it is built again. */
const INHERITED_KEYS = 4096;
/** As `permutedCellKeys`, for the entries that name a cell's place. Null
 * when one in the span is not the cell at its canonical key: where that one
 * went, only a whole walk can say. */
function permutedEntries<T extends object>(entries: readonly T[], span: PermutedSpan, at: (entry: T) => readonly [number, number], holds: (entry: T, row: number, col: number) => boolean, placed: (entry: T, row: number) => T): readonly T[] | null {
  const { permutation: { dataR1, c1, c2 }, r2, destination } = span;
  if (entries.length === 0) return entries;
  const outside: T[] = [];
  const inside: { entry: T; row: number; col: number }[] = [];
  for (const entry of entries) {
    const [row, col] = at(entry);
    if (row < dataR1 || row > r2 || col < c1 || col > c2) { outside.push(entry); continue; }
    if (!holds(entry, row, col)) return null;
    const moved = destination[row - dataR1];
    inside.push({ entry: moved === row ? entry : placed(entry, moved), row: moved, col });
  }
  if (inside.length === 0) return entries;
  inside.sort((a, b) => a.row - b.row || a.col - b.col);
  for (const { entry } of inside) outside.push(entry);
  return outside;
}

/** The used range after a row sort. Cells keep their columns, so only its
 * row ends can move, and only where they touch the span: an end inside the
 * span holds while that row still has a value. Otherwise undefined: the
 * consumer computes it from the map, as after an edit that blanked a cell. */
function usedRangeAfterRowSort(prior: SelectionObservationRange | null | undefined, span: PermutedSpan, before: Cells, after: Cells): SelectionObservationRange | null | undefined {
  if (prior === undefined) return undefined;
  const { permutation, r2, relocatedAt } = span;
  const { dataR1 } = permutation;
  let range = prior;
  if (range && range.r1 <= r2 && range.r2 >= dataR1) {
    const used = range;
    const holds = (row: number) => {
      for (let col = used.c1; col <= used.c2; col += 1) if (nonblank(after[`${row}:${col}`])) return true;
      return false;
    };
    if (range.r1 > dataR1 || range.r2 < r2) return undefined;
    if ((range.r1 === dataR1 && !holds(dataR1)) || (range.r2 === r2 && !holds(r2))) return undefined;
  }
  // A rewritten formula's value is its text: a cell that stored none now holds one.
  for (const [row, col] of relocatedAt) {
    if (!nonblank(before[cellMapPermutedSourceKey(permutation, row, col)])) range = expand(range, row, col);
  }
  return range;
}

/** The observed identities and their counts stay: a moved cell keeps its style
 * and lock. The bounds grow to the span's rows, which its cells may now fill. */
function observationAfterRowSort(prior: StoredSelectionSummary, span: PermutedSpan): StoredSelectionSummary {
  const { bounds } = prior;
  const { dataR1 } = span.permutation;
  if (!bounds || (bounds.r1 <= dataR1 && bounds.r2 >= span.r2)) return prior;
  return Object.freeze({ ...prior, bounds: { ...bounds, r1: Math.min(bounds.r1, dataR1), r2: Math.max(bounds.r2, span.r2) } });
}

export class SpreadsheetCellIndexes {
  #sources = new WeakMap<Cells, SpreadsheetCellIndex>();
  #counts = new WeakMap<Cells, number>();
  count(cells: Cells): number {
    let count = this.#counts.get(cells);
    if (count === undefined) { count = Object.keys(cells).length; this.#counts.set(cells, count); }
    return count;
  }
  get(cells: Cells): SpreadsheetCellIndex {
    const cached = this.#sources.get(cells);
    if (cached) return cached;
    // A row sort of a map indexed here: the same cells at new keys.
    const sorted = this.#afterRowSort(cells);
    if (sorted) return sorted;
    // A map recorded as an edit of one indexed here (a calculation's stamped
    // values: 2,122 formulas of a sorted 730,000-cell sheet, whose whole
    // index took 0.65 s) is brought up to date at the keys it changed.
    const found = nearestKnownCellMap(cells, (known: Cells) => (this.#sources.has(known) || this.#afterRowSort(known) ? known : undefined));
    if (found && found.keys.length <= INHERITED_KEYS) {
      this.inherit(found.value, cells, found.keys);
      const inherited = this.#sources.get(cells);
      if (inherited) return inherited;
    }
    let notes = 0;
    const threads: IndexedThread[] = [];
    const autoHeightKeys: string[] = [];
    const formulaKeys: string[] = [];
    const fillKeys: string[] = [];
    const displayCacheKeys: string[] = [];
    const phoneticKeys: string[] = [];
    const heavyEdgeKeys: string[] = [];
    const spans: IndexedSpan[] = [];
    const facets: IndexedFacets[] = [];
    const keys = Object.keys(cells);
    this.#counts.set(cells, keys.length);
    const observer = createSelectionObserver();
    let usedRange: SelectionObservationRange | null = null, bounds: SelectionObservationRange | null = null, observed = 0;
    for (const key of keys) {
      const cell = cells[key];
      if (autoHeightCandidate(cell)) autoHeightKeys.push(key);
      if (formulaCandidate(cell)) formulaKeys.push(key);
      if (fillCandidate(cell)) fillKeys.push(key);
      if (displayCacheCandidate(cell)) displayCacheKeys.push(key);
      if (phoneticCandidate(cell)) phoneticKeys.push(key);
      if (heavyEdgeCandidate(cell)) heavyEdgeKeys.push(key);
      const lenient = parseLenient(key);
      if (lenient) {
        observed++; observer.observe(cell); bounds = expand(bounds, lenient[0], lenient[1]);
        if (nonblank(cell) && parseCanonical(key)) usedRange = expand(usedRange, lenient[0], lenient[1]);
      }
      const rowSpan = rowSpanOf(cell), colSpan = colSpanOf(cell);
      if (rowSpan > 1 || colSpan > 1) {
        const [rowPart, colPart] = key.split(":");
        const row = Number(rowPart), column = Number(colPart);
        if (Number.isFinite(row) && Number.isFinite(column)) spans.push({ row, column, rowSpan, colSpan });
      }
      if (cell?.comment || cell?.hyperlink) {
        const separator = key.indexOf(":");
        if (separator > 0 && separator < key.length - 1) {
          const row = Number(key.slice(0, separator)), column = Number(key.slice(separator + 1));
          if (Number.isSafeInteger(row) && Number.isSafeInteger(column)) {
            facets.push({ row, column, comments: Boolean(cell.comment), hyperlinks: Boolean(cell.hyperlink) });
          }
        }
      }
      const comment = cell?.comment;
      if (hasNote(comment)) notes++;
      if (!isThreadComment(comment) || comment.thread.length === 0) continue;
      const [rowText, colText] = key.split(":");
      const row = Number(rowText), col = Number(colText);
      if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
      threads.push({ key, row, col, comment });
    }
    const result: SpreadsheetCellIndex = { notes, threads, autoHeightKeys, formulaKeys, fillKeys, displayCacheKeys, phoneticKeys, heavyEdgeKeys, spans, facets,
      usedRange, observation: observer.summary(observed, bounds) };
    this.#sources.set(cells, result);
    return result;
  }
  /** The index of a map a row sort made from one indexed here, without
   * walking it (0.45 s over 730,000 cells, one of many such walks a sort
   * sets off): it equals the index built whole, but that its observation
   * bounds may be wider (allowed, they only ever grow) and its used range may
   * be undefined where blank rows of the span moved past its ends. Null when
   * only a whole walk can say. */
  #afterRowSort(cells: Cells): SpreadsheetCellIndex | null {
    const record = cellMapRowPermutation(cells);
    const prior = record && this.#sources.get(record.before);
    if (!record || !prior) return null;
    const { before, permutation } = record;
    const span = permutedSpanOf(permutation);
    const threads = permutedEntries(prior.threads, span, (entry) => [entry.row, entry.col],
      (entry, row, col) => entry.key === `${row}:${col}`,
      (entry, row) => ({ ...entry, key: `${row}:${entry.col}`, row }));
    const spans = threads && permutedEntries(prior.spans, span, (entry) => [entry.row, entry.column],
      (entry, row, col) => { const cell = before[`${row}:${col}`]; return rowSpanOf(cell) === entry.rowSpan && colSpanOf(cell) === entry.colSpan; },
      (entry, row) => ({ ...entry, row }));
    const facets = spans && permutedEntries(prior.facets, span, (entry) => [entry.row, entry.column],
      (entry, row, col) => { const cell = before[`${row}:${col}`]; return Boolean(cell?.comment || cell?.hyperlink); },
      (entry, row) => ({ ...entry, row }));
    if (!threads || !spans || !facets) return null;
    const result: SpreadsheetCellIndex = {
      notes: prior.notes,
      threads, spans, facets,
      autoHeightKeys: permutedCellKeys(prior.autoHeightKeys, span, cells, autoHeightCandidate),
      formulaKeys: permutedCellKeys(prior.formulaKeys, span, cells, formulaCandidate),
      fillKeys: permutedCellKeys(prior.fillKeys, span, cells, fillCandidate),
      displayCacheKeys: permutedCellKeys(prior.displayCacheKeys, span, cells, displayCacheCandidate),
      phoneticKeys: permutedCellKeys(prior.phoneticKeys, span, cells, phoneticCandidate),
      heavyEdgeKeys: permutedCellKeys(prior.heavyEdgeKeys, span, cells, heavyEdgeCandidate),
      usedRange: usedRangeAfterRowSort(prior.usedRange, span, before, cells),
      observation: observationAfterRowSort(prior.observation, span),
    };
    this.#sources.set(cells, result);
    const count = this.#counts.get(before);
    if (count !== undefined && !this.#counts.has(cells)) this.#counts.set(cells, count);
    return result;
  }
  /** knownKeys must describe every changed cell in an accepted transaction.
   * The keyed facets are shared unchanged when no changed key alters them; the
   * observation summary and used range follow the changed keys exactly, so an
   * accepted edit never enumerates the new map. */
  inherit(before: Cells, after: Cells, knownKeys: readonly string[]): void {
    if (before === after || this.#sources.has(after)) return;
    const prior = this.#sources.get(before);
    let sameMetadata = true, delta = 0;
    let usedRange = prior?.usedRange, observation = prior?.observation;
    let observer: ReturnType<typeof createSelectionObserver> | null = null;
    for (const key of new Set(knownKeys)) {
      const had = Object.prototype.propertyIsEnumerable.call(before, key);
      const has = Object.prototype.propertyIsEnumerable.call(after, key);
      delta += Number(has) - Number(had);
      const a = had ? before[key] : undefined, b = has ? after[key] : undefined;
      if (a?.comment !== b?.comment || Boolean(a?.hyperlink) !== Boolean(b?.hyperlink) ||
          autoHeightCandidate(a) !== autoHeightCandidate(b) || formulaCandidate(a) !== formulaCandidate(b) ||
          fillCandidate(a) !== fillCandidate(b) || displayCacheCandidate(a) !== displayCacheCandidate(b) ||
          phoneticCandidate(a) !== phoneticCandidate(b) || heavyEdgeCandidate(a) !== heavyEdgeCandidate(b) ||
          !Object.is(rowSpanOf(a), rowSpanOf(b)) || !Object.is(colSpanOf(a), colSpanOf(b))) sameMetadata = false;
      if (!prior || !sameMetadata) continue;
      const lenient = parseLenient(key);
      if (!lenient || (had === has && a === b)) continue;
      // Counted identities move one cell out and one cell in exactly; the
      // bounds only ever grow, which keeps the summary a safe superset. A
      // value-only edit contributes the same style and lock, so it keeps the
      // very same summary and index.
      if (observation && !(had && has && sameObservation(a, b))) {
        observer ??= createSelectionObserver(observation);
        if (had) observer.unobserve(a);
        if (has) observer.observe(b);
        observation = observer.summary(observation.count + Number(has) - Number(had),
          has ? expand(observation.bounds, lenient[0], lenient[1]) : observation.bounds);
      }
      if (usedRange !== undefined && parseCanonical(key)) {
        if (nonblank(a) && !nonblank(b)) usedRange = undefined;
        else if (!nonblank(a) && nonblank(b)) usedRange = expand(usedRange, lenient[0], lenient[1]);
      }
    }
    const count = this.#counts.get(before);
    if (count !== undefined && !this.#counts.has(after)) this.#counts.set(after, count + delta);
    if (prior && sameMetadata) {
      this.#sources.set(after, usedRange === prior.usedRange && observation === prior.observation
        ? prior : { ...prior, usedRange, observation: observation ?? prior.observation });
    }
  }
}
