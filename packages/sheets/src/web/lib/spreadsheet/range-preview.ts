/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PAGE_LOCATE_FLASH_MS, PAGE_LOCATED_ATTRIBUTE } from "@mrnote/core/web/lib/page-locate.ts";

import type { CellRange as RangeRect } from "./block-clear";
import { RANGE_CELL_SELECTOR } from "./selection-glide";

/** The lines that carry the selection's border. Excel draws ONE border around
 * a selected range; when the range starts or ends on a hidden row or column
 * the border falls on the first or last line of it that can be seen. Computed
 * from the rectangle and the hidden lines alone, never from neighbouring
 * cells: a virtualized grid has most of them unmounted. */
export function visibleSelectionEdges(
  range: RangeRect,
  rowHidden: (row: number) => boolean,
  colHidden: (col: number) => boolean,
): RangeRect {
  let { r1, r2, c1, c2 } = range;
  while (r1 < r2 && rowHidden(r1)) r1 += 1;
  while (r2 > r1 && rowHidden(r2)) r2 -= 1;
  while (c1 < c2 && colHidden(c1)) c1 += 1;
  while (c2 > c1 && colHidden(c2)) c2 -= 1;
  return { r1, r2, c1, c2 };
}

/** Which sides of one selected cell lie on the range's perimeter, as the
 * space-separated tokens of `data-selection-edges` (top, bottom, start, end).
 * `start` / `end` are logical, so the rectangle is right on an RTL sheet. A
 * merge anchor reaches its edges through its span. */
export function selectionEdgeTokens(
  edges: RangeRect,
  row: number,
  col: number,
  rowSpan = 1,
  colSpan = 1,
): string {
  const tokens: string[] = [];
  if (row <= edges.r1) tokens.push("top");
  if (row + rowSpan - 1 >= edges.r2) tokens.push("bottom");
  if (col <= edges.c1) tokens.push("start");
  if (col + colSpan - 1 >= edges.c2) tokens.push("end");
  return tokens.join(" ");
}

/** How a row or column heading answers the selection (Excel): "partial" when
 * the selection crosses its line, "full" when the whole line is selected. */
export function headerSelectionMark(index: number, from: number, to: number, whole: boolean): "full" | "partial" | null {
  if (index < from || index > to) return null;
  return whole ? "full" : "partial";
}

export interface SheetExtent { rows: number; cols: number }

/** Brings every mounted heading to the mark its own render would produce.
 * Row headings live in rows that do not re-render for a selection change. */
function writeHeaderMarks(root: HTMLElement | null, range: RangeRect | null, extent: SheetExtent | null, remember?: (header: HTMLElement) => void) {
  if (!root) return;
  const wholeRows = range !== null && extent !== null && extent.cols > 0 && range.c1 <= 0 && range.c2 >= extent.cols - 1;
  const wholeCols = range !== null && extent !== null && extent.rows > 0 && range.r1 <= 0 && range.r2 >= extent.rows - 1;
  for (const header of root.querySelectorAll<HTMLElement>("th[data-header-row], th[data-header-col]")) {
    const isRow = header.dataset.headerRow !== undefined;
    const index = Number(isRow ? header.dataset.headerRow : header.dataset.headerCol);
    const mark = range === null ? null
      : isRow ? headerSelectionMark(index, range.r1, range.r2, wholeRows) : headerSelectionMark(index, range.c1, range.c2, wholeCols);
    remember?.(header);
    if (mark === null) {
      if (header.hasAttribute("data-selection-header")) header.removeAttribute("data-selection-header");
    } else if (header.getAttribute("data-selection-header") !== mark) header.setAttribute("data-selection-header", mark);
  }
}

type EdgeResolver = (range: RangeRect) => RangeRect;
const MARK_ATTRIBUTES = ["data-selected", "aria-selected", "data-selection-edges"] as const;

function cellSpan(cell: HTMLElement): [number, number] {
  const span = cell.dataset.rangeSpan;
  if (!span) return [1, 1];
  const [rowSpan, colSpan] = span.split(":").map(Number);
  return [rowSpan || 1, colSpan || 1];
}

function writeEdges(cell: HTMLElement, selected: boolean, edges: RangeRect, row: number, col: number) {
  const tokens = selected ? selectionEdgeTokens(edges, row, col, ...cellSpan(cell)) : "";
  if (!tokens) {
    if (cell.hasAttribute("data-selection-edges")) cell.removeAttribute("data-selection-edges");
  } else if (cell.getAttribute("data-selection-edges") !== tokens) cell.setAttribute("data-selection-edges", tokens);
}

const FILL_HANDLE_CLASS = "spreadsheet-window__fill-handle";
/** On the window root while a dragged-out range carries the fill handle. */
export const FILL_HANDLE_PREVIEW_ATTRIBUTE = "data-fill-handle-preview";
const FILL_HANDLE_GHOST_ATTRIBUTE = "data-fill-handle-ghost";

/** The cell holding the range's bottom end corner: the perimeter cell that
 * draws both the bottom and the end side (a merge reaches it through its
 * span, a hidden edge line hands it to the last line that shows). */
function isCornerCell(cell: HTMLElement): boolean {
  const tokens = (cell.getAttribute("data-selection-edges") ?? "").split(" ");
  return tokens.includes("bottom") && tokens.includes("end");
}

/** Selection paint is transient DOM state; final selection still belongs to
 * the existing worksheet command host. No cell values or focus are changed.
 *
 * The fill handle rides the dragged-out range (Excel): the rendered handle
 * belongs to the committed selection and only moves when the drag commits,
 * so while the preview paints, that one steps back and a look-alike sits in
 * the preview's corner cell, written in the same pass as the frame's edges.
 * While the frame glides, the gliding box carries the handle instead. */
export function createRangePreview(getRoot: () => HTMLElement | null, resolveEdges: EdgeResolver = (next) => next, getExtent: () => SheetExtent | null = () => null, onPainted: (edges: RangeRect) => void = () => {}) {
  let frame: number | null = null;
  let range: RangeRect | null = null;
  const originals = new Map<HTMLElement, (string | null)[]>();
  const headerOriginals = new Map<HTMLElement, string | null>();
  // Whether the selection wears a fill handle at all (not on a PivotTable's
  // output, a preview surface...): read once, from the committed render, when
  // the drag paints first; its row may scroll out of the rendered ones later.
  let handleWanted: boolean | null = null;
  let handleRoot: HTMLElement | null = null;
  let ghost: HTMLSpanElement | null = null;
  const dropHandle = () => {
    ghost?.remove();
    handleRoot?.removeAttribute(FILL_HANDLE_PREVIEW_ATTRIBUTE);
    handleRoot = null;
  };
  const placeHandle = (root: HTMLElement, corner: HTMLElement | null) => {
    if (handleWanted === null) handleWanted = root.querySelector(`.${FILL_HANDLE_CLASS}:not([${FILL_HANDLE_GHOST_ATTRIBUTE}])`) !== null;
    if (!handleWanted) return;
    if (handleRoot !== root) {
      dropHandle();
      handleRoot = root;
    }
    root.setAttribute(FILL_HANDLE_PREVIEW_ATTRIBUTE, "true");
    if (!corner) {
      // The corner is off the rendered rows: no handle until it comes back.
      ghost?.remove();
      return;
    }
    if (!ghost) {
      ghost = root.ownerDocument.createElement("span");
      ghost.className = FILL_HANDLE_CLASS;
      ghost.setAttribute(FILL_HANDLE_GHOST_ATTRIBUTE, "true");
      ghost.setAttribute("aria-hidden", "true");
      ghost.dataset.testid = "spreadsheet-fill-handle-preview";
    }
    // Last in the cell, where React appends nothing of its own before it.
    if (ghost.parentElement !== corner || ghost.nextSibling) corner.append(ghost);
  };
  const paint = () => {
    frame = null;
    if (!range) return;
    const root = getRoot();
    const edges = resolveEdges(range);
    let corner: HTMLElement | null = null;
    for (const cell of root?.querySelectorAll<HTMLElement>(RANGE_CELL_SELECTOR) ?? []) {
      const [row, col] = cell.dataset.rangeCell!.split(":").map(Number);
      const selected = row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
      if (!originals.has(cell)) originals.set(cell, MARK_ATTRIBUTES.map((attr) => cell.getAttribute(attr)));
      // A cell outside the range carries no data-selected at all, as the
      // cell render and the committed paint leave it: written as "false",
      // the attribute changed on every mounted cell when a drag began and
      // again when it ended, and each change restyled the cell (~3,000 on a
      // full window). Only the cells whose marks change are touched.
      if (selected) {
        if (cell.getAttribute("data-selected") !== "true") cell.setAttribute("data-selected", "true");
        if (cell.getAttribute("aria-selected") !== "true") cell.setAttribute("aria-selected", "true");
      } else {
        if (cell.hasAttribute("data-selected")) cell.removeAttribute("data-selected");
        if (cell.getAttribute("data-active") !== "true" && cell.hasAttribute("aria-selected")) cell.removeAttribute("aria-selected");
      }
      writeEdges(cell, selected, edges, row, col);
      if (selected && isCornerCell(cell)) corner = cell;
    }
    if (root) placeHandle(root, corner);
    writeHeaderMarks(getRoot(), range, getExtent(), (header) => {
      if (!headerOriginals.has(header)) headerOriginals.set(header, header.getAttribute("data-selection-header"));
    });
    onPainted(edges);
  };
  return {
    schedule(next: RangeRect) { range = next; if (frame === null) frame = requestAnimationFrame(paint); },
    paintNow() { if (frame !== null) cancelAnimationFrame(frame); frame = null; paint(); },
    clear() {
      if (frame !== null) cancelAnimationFrame(frame);
      frame = null; range = null;
      dropHandle();
      handleWanted = null;
      for (const [cell, values] of originals) {
        for (const [index, attr] of MARK_ATTRIBUTES.entries()) {
          if (values[index] === null) cell.removeAttribute(attr); else cell.setAttribute(attr, values[index]!);
        }
      }
      originals.clear();
      for (const [header, value] of headerOriginals) {
        if (value === null) header.removeAttribute("data-selection-header"); else header.setAttribute("data-selection-header", value);
      }
      headerOriginals.clear();
    },
  };
}

/** Committed selection marks, written directly to the rendered cells. Rows do
 * not re-render for a selection change; this brings every mounted cell to the
 * same attributes the cell render itself would produce, so a later React
 * render of a row only ever writes the same truth. The look (one border around
 * the range, a tint on every cell but the active one) is the stylesheet's,
 * keyed on these attributes; the active cell keeps aria-selected. */
export function createSelectionPaint(getRoot: () => HTMLElement | null, onPainted: (edges: RangeRect | null) => void = () => {}) {
  return {
    paint(range: RangeRect | null, edges: RangeRect | null = range) {
      for (const cell of getRoot()?.querySelectorAll<HTMLElement>(RANGE_CELL_SELECTOR) ?? []) {
        const [row, col] = cell.dataset.rangeCell!.split(":").map(Number);
        const selected = range !== null && row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
        if (selected) {
          if (cell.getAttribute("data-selected") !== "true") cell.setAttribute("data-selected", "true");
          if (cell.getAttribute("aria-selected") !== "true") cell.setAttribute("aria-selected", "true");
        } else {
          if (cell.hasAttribute("data-selected")) cell.removeAttribute("data-selected");
          if (cell.getAttribute("data-active") === "true") {
            if (cell.getAttribute("aria-selected") !== "true") cell.setAttribute("aria-selected", "true");
          } else if (cell.hasAttribute("aria-selected")) cell.removeAttribute("aria-selected");
        }
        writeEdges(cell, selected, edges ?? range ?? { r1: row, r2: row, c1: col, c2: col }, row, col);
      }
      onPainted(edges ?? range);
    },
    /** A press made this the selection while its paint waits for the drag to
     * end (or not to start): the frame may go there already. */
    pressed(edges: RangeRect) {
      onPainted(edges);
    },
    /** The cells were laid out again where the selection stays (a Page
     * Layout page re-planned, zoomed or resized): its frame follows them. */
    relaid(edges: RangeRect | null) {
      onPainted(edges);
    },
    /** The headings of the selection's rows and columns, a single active cell
     * included. */
    paintHeaders(range: RangeRect | null, extent: SheetExtent | null) {
      writeHeaderMarks(getRoot(), range, extent);
    },
  };
}

export interface FlashLocatedRangeOptions {
  /** False while the grid still shows another sheet (a jump across sheets
   * commits the target sheet a render later). */
  ready?: () => boolean;
  /** Frames to wait for the range's cells to mount (virtualized rows land
   * after the scroll). */
  frames?: number;
}

/** Frames a jump waits for its cells: about half a second at 60 Hz. */
const FLASH_LOCATED_RANGE_FRAMES = 30;

/** The jump flash: a Name Box / Go To commit, a Find All row, a Navigation
 * pane item or an AI "Locate" tints the cells it landed on once, the Docs
 * locate flash (styles/office-motion.css, keyed on data-located). Only the
 * mounted cells of the range are marked; the grid's own navigation owns the
 * scroll, so nothing here moves the viewport. Returns a cancel that clears
 * the marks at once. */
export function flashLocatedRange(
  getRoot: () => HTMLElement | null,
  range: RangeRect,
  options: FlashLocatedRangeOptions = {},
): () => void {
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  let marked: HTMLElement[] = [];
  let frame: number | null = null;
  let timer: ReturnType<typeof setTimeout> | null = null;
  const unmark = () => {
    for (const cell of marked) cell.removeAttribute(PAGE_LOCATED_ATTRIBUTE);
    marked = [];
  };
  const attempt = (framesLeft: number) => {
    frame = null;
    const root = getRoot();
    const cells =
      root && (options.ready?.() ?? true)
        ? Array.from(root.querySelectorAll<HTMLElement>(RANGE_CELL_SELECTOR)).filter(
            (cell) => {
              const [row, col] = cell.dataset.rangeCell!.split(":").map(Number);
              return row >= r1 && row <= r2 && col >= c1 && col <= c2;
            },
          )
        : [];
    if (cells.length === 0) {
      if (framesLeft > 0) frame = requestAnimationFrame(() => attempt(framesLeft - 1));
      return;
    }
    // Restart the flash on cells a previous jump is still tinting.
    for (const cell of cells) cell.removeAttribute(PAGE_LOCATED_ATTRIBUTE);
    void cells[0].offsetWidth;
    for (const cell of cells) cell.setAttribute(PAGE_LOCATED_ATTRIBUTE, "true");
    marked = cells;
    timer = setTimeout(() => {
      timer = null;
      unmark();
    }, PAGE_LOCATE_FLASH_MS);
  };
  attempt(options.frames ?? FLASH_LOCATED_RANGE_FRAMES);
  return () => {
    if (frame !== null) cancelAnimationFrame(frame);
    if (timer !== null) clearTimeout(timer);
    frame = null;
    timer = null;
    unmark();
  };
}
