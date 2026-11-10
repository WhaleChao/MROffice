"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Boxes drawn over ranges of the grid, measured off the cells it has on
 * screen: Excel's range finder while a formula is entered
 * (`SpreadsheetFormulaReferenceOverlay`) and a selected chart's source data
 * (`SpreadsheetChartSourceOverlay`) both outline cells this way.
 *
 * The host is a layer laid over the grid stage (its parent element): a box
 * is in the stage's own pixels, so it stays on its cells through the
 * stage's zoom. A range that runs past the rendered rows or
 * columns is boxed as far as they go (the edge it runs on past is not its
 * own, `edges`), one wholly out of view has no box, and the boxes are
 * measured again as the grid scrolls (frozen panes keep their cells in
 * place while the rest moves) or resizes.
 */

import { useLayoutEffect, useState, type RefObject } from "react";

export interface SheetRangeSpan {
  /** Stable identity of the range among those drawn. */
  readonly key: string;
  /** Zero-based, inclusive, top-left first (`Infinity` runs to the end). */
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export interface SheetRangeBox<T extends SheetRangeSpan> {
  readonly span: T;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /** Which edges are the range's own (false: it runs on out of view). */
  readonly edges: readonly [top: boolean, right: boolean, bottom: boolean, left: boolean];
}

const CELL_SELECTOR = '[data-testid="spreadsheet-cell"][data-cell]';

function cellKey(element: HTMLElement): { row: number; col: number } | null {
  return cellKeyOf(element.dataset.cell);
}

function cellKeyOf(key: string | null | undefined): { row: number; col: number } | null {
  if (!key) return null;
  const separator = key.indexOf(":");
  const row = Number(key.slice(0, separator));
  const col = Number(key.slice(separator + 1));
  return Number.isInteger(row) && Number.isInteger(col) ? { row, col } : null;
}

function measureBoxes<T extends SheetRangeSpan>(stage: HTMLElement, spans: readonly T[]): SheetRangeBox<T>[] {
  const stageRect = stage.getBoundingClientRect();
  // The stage may be zoomed: its children's own pixels are layout pixels.
  const scale = stage.offsetWidth > 0 ? stageRect.width / stage.offsetWidth : 1;
  const cells = new Map<string, HTMLElement>();
  let minRow = Infinity, maxRow = -1, minCol = Infinity, maxCol = -1;
  // A ghost cell (the canvas text layer's, without an editor) is its own address.
  for (const editor of stage.querySelectorAll<HTMLElement>(`${CELL_SELECTOR}, td[data-ghost-cell]`)) {
    const ghost = editor.hasAttribute("data-ghost-cell");
    const at = ghost ? cellKeyOf(editor.getAttribute("data-range-cell")) : cellKey(editor);
    const cell = ghost ? editor : editor.closest("td");
    if (!at || !cell) continue;
    cells.set(`${at.row}:${at.col}`, cell);
    if (at.row < minRow) minRow = at.row;
    if (at.row > maxRow) maxRow = at.row;
    if (at.col < minCol) minCol = at.col;
    if (at.col > maxCol) maxCol = at.col;
  }
  const corner = (row: number, col: number, toward: 1 | -1, limitRow: number, limitCol: number) => {
    // The cell at (row, col), or the nearest rendered one further into the
    // range (a cell a merge covers, or a hidden line, has no element).
    for (let r = row; toward > 0 ? r <= limitRow : r >= limitRow; r += toward) {
      for (let c = col; toward > 0 ? c <= limitCol : c >= limitCol; c += toward) {
        const cell = cells.get(`${r}:${c}`);
        if (cell) return cell;
      }
    }
    return undefined;
  };
  const next: SheetRangeBox<T>[] = [];
  for (const span of spans) {
    const r1 = Math.max(span.r1, minRow), r2 = Math.min(span.r2, maxRow);
    const c1 = Math.max(span.c1, minCol), c2 = Math.min(span.c2, maxCol);
    if (r1 > r2 || c1 > c2) continue;
    const first = corner(r1, c1, 1, r2, c2);
    const last = corner(r2, c2, -1, r1, c1);
    if (!first || !last) continue;
    const a = first.getBoundingClientRect(), b = last.getBoundingClientRect();
    const left = Math.min(a.left, b.left), right = Math.max(a.right, b.right);
    const top = Math.min(a.top, b.top), bottom = Math.max(a.bottom, b.bottom);
    next.push({
      span,
      left: (left - stageRect.left) / scale,
      top: (top - stageRect.top) / scale,
      width: (right - left) / scale,
      height: (bottom - top) / scale,
      edges: [span.r1 >= minRow, span.c2 <= maxCol, span.r2 <= maxRow, span.c1 >= minCol],
    });
  }
  return next;
}

const sameBoxes = <T extends SheetRangeSpan>(current: readonly SheetRangeBox<T>[], next: readonly SheetRangeBox<T>[]) =>
  current.length === next.length &&
  current.every((box, index) => {
    const other = next[index]!;
    return box.span === other.span &&
      Math.abs(box.left - other.left) < 0.25 && Math.abs(box.top - other.top) < 0.25 &&
      Math.abs(box.width - other.width) < 0.25 && Math.abs(box.height - other.height) < 0.25 &&
      box.edges.every((edge, side) => edge === other.edges[side]);
  });

/**
 * The boxes of `spans` over the grid stage that is `hostRef`'s parent,
 * measured again when the spans change, `layoutKey` changes (the grid laid
 * out again: its scroll window, sizes, zoom), the stage scrolls or resizes.
 */
/**
 * The grid stage an overlay host draws over: its parent, or its parent's
 * parent when the host stands in a stage slot (a `display: contents` box
 * that keeps the stage's own children fixed while the overlay comes and
 * goes, see SpreadsheetWindow).
 */
export function sheetOverlayStage(host: HTMLElement | null | undefined): HTMLElement | null {
  const parent = host?.parentElement ?? null;
  return parent?.hasAttribute("data-spreadsheet-stage-slot") ? parent.parentElement : parent;
}

export function useSheetRangeBoxes<T extends SheetRangeSpan>(
  hostRef: RefObject<HTMLElement | null>,
  spans: readonly T[],
  layoutKey: unknown,
): readonly SheetRangeBox<T>[] {
  const [boxes, setBoxes] = useState<readonly SheetRangeBox<T>[]>([]);
  useLayoutEffect(() => {
    const host = hostRef.current;
    const stage = sheetOverlayStage(host);
    if (!host || !stage || spans.length === 0) {
      setBoxes((current) => (current.length ? [] : current));
      return;
    }
    const measure = () => {
      const next = measureBoxes(stage, spans);
      setBoxes((current) => (sameBoxes(current, next) ? current : next));
    };
    measure();
    // Frozen panes keep their cells in place while the rest scrolls, and the
    // grid renders other rows as it goes: measure again as it moves.
    let frame = 0;
    const schedule = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(measure);
    };
    const scroller = stage.parentElement;
    scroller?.addEventListener("scroll", schedule, { passive: true });
    const resize = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(schedule);
    resize?.observe(stage);
    return () => {
      cancelAnimationFrame(frame);
      scroller?.removeEventListener("scroll", schedule);
      resize?.disconnect();
    };
  }, [hostRef, spans, layoutKey]);
  return boxes;
}

/** The grid cell under a point of the screen (the frozen one over a body
 * cell scrolled beneath it; a merge's anchor for any cell it covers), or
 * null off the cells. */
export function sheetCellAtPoint(stage: HTMLElement, clientX: number, clientY: number): { row: number; col: number } | null {
  if (typeof document.elementsFromPoint !== "function") return null;
  for (const element of document.elementsFromPoint(clientX, clientY)) {
    if (!(element instanceof HTMLElement) || !stage.contains(element)) continue;
    const cell = element.closest("td");
    const editor = cell?.querySelector<HTMLElement>(CELL_SELECTOR);
    const at = editor ? cellKey(editor) : cell?.hasAttribute("data-ghost-cell") ? cellKeyOf(cell.getAttribute("data-range-cell")) : null;
    if (at) return at;
  }
  return null;
}
