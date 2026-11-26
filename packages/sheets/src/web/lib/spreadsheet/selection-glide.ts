/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The selection's frame in motion. Excel never teleports it: clicked to
 * another cell, the frame flies there -- the farther, the faster -- and a
 * range dragged out or pulled back grows and shrinks under the pointer
 * instead of stepping a cell at a time.
 *
 * At rest the frame is the cells' own: each cell on the perimeter draws its
 * sides (`data-selection-edges`, the active cell's ring, styles/native-page-
 * window.css), which keeps it right on merged, hidden, virtualized and frozen
 * cells. When the painted selection moves, one box over the grid's stage
 * eases from where the frame was to where it now is, and the cells' frame
 * and fill handle are held back (`data-selection-glide` on the stage) until
 * the box arrives. A new target while it travels redirects it from where it
 * is, so a drag reads as one continuous stretch.
 *
 * Every leg takes the same short time on an ease-out curve, so a long move
 * runs faster than a short one and none trails a tail of creeping pixels. No glide (the frame just shows where it is) when
 * motion is reduced, on a frozen cell, or when the perimeter is not all
 * mounted (a jump past the rendered rows): a box that could not land on the
 * cells' frame would end in a jump of its own.
 */

import { reducedMotionNow } from "@mrnote/core/web/lib/reduce-motion.ts";

/** A selected range in rows and columns, its hidden edge lines resolved. */
export interface GlideRange {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export interface GlideFrame {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

export interface SelectionGlideOptions {
  /** The stage the frame is drawn in (the cells' positioned ancestor): the
   * grid's stage, or in Page Layout the page holding the range's first cell
   * (selectionStageIn). */
  readonly getStage: (range?: GlideRange | null) => HTMLElement | null;
  /** False where the frame must stay put (split panes, right-to-left, a preview surface). */
  readonly enabled: () => boolean;
  /** What the frame belongs to (the sheet and the view): another one starts it afresh. */
  readonly scope: () => unknown;
}

/** A selectable cell of either view: the grid's <td>, or a Page Layout
 * page's cell box. Both carry the same marks (data-range-cell,
 * data-selected, data-active, data-selection-edges). */
export const RANGE_CELL_SELECTOR = "td[data-range-cell], [data-page-layout-cell][data-range-cell]";
const rangeCellAt = (stage: ParentNode, key: string) =>
  stage.querySelector<HTMLElement>(`td[data-range-cell="${key}"], [data-page-layout-cell][data-range-cell="${key}"]`);

/** A Page Layout page's cell grid: the stage its selection frame is drawn in. */
export const PAGE_LAYOUT_STAGE_ATTRIBUTE = "data-page-layout-stage";

/**
 * Where a selection's frames are drawn inside the sheet's viewport: the grid's
 * one stage in Normal view and Page Break Preview; in Page Layout, the page
 * that shows the range's first cell -- the page whose cell has the keyboard
 * when the cell appears on more than one (a print title), else the first.
 */
export function selectionStageIn(viewport: HTMLElement | null, range?: GlideRange | null): HTMLElement | null {
  if (!viewport) return null;
  const grid = viewport.querySelector<HTMLElement>(":scope > .spreadsheet-window__grid-stage");
  if (grid) return grid;
  const pages = `[${PAGE_LAYOUT_STAGE_ATTRIBUTE}]`;
  const focused = viewport.ownerDocument.activeElement;
  const key = range ? `${range.r1}:${range.c1}` : null;
  if (focused instanceof HTMLElement && viewport.contains(focused)) {
    const page = focused.closest<HTMLElement>(pages);
    if (page && (!key || rangeCellAt(page, key))) return page;
  }
  if (key) {
    const cell = rangeCellAt(viewport, key);
    const page = cell?.closest<HTMLElement>(pages);
    if (page) return page;
  }
  return viewport.querySelector<HTMLElement>('[data-page-layout-cell][data-active="true"]')?.closest<HTMLElement>(pages)
    ?? null;
}

/** One leg of the flight, whatever its length. */
const GLIDE_DURATION_MS = 110;
/** Closer than this on every side, the box is where the cells draw the frame. */
const GLIDE_SETTLE_PX = 0.5;
const easeOutCubic = (p: number) => 1 - (1 - p) ** 3;
/** How far the cells' frame reaches past a perimeter cell's box (measured). */
const FRAME_OUTSET_PX = 0.5;
/** A Page Layout cell's box holds its own right and bottom gridline and none
 * of the line above or before it. The frame sits on the gridlines as the
 * grid's does: one pixel out over the line before and above the range, none
 * past the line the range's last cells hold. */
const PAGE_FRAME_OUTSET_BEFORE_PX = 1;
const PAGE_FRAME_OUTSET_AFTER_PX = 0;
export const SELECTION_GLIDE_ATTRIBUTE = "data-selection-glide";
/** On a stage whose frame is the landed box itself (Page Layout): the cells'
 * own frame stays under it. */
export const SELECTION_FRAME_ATTRIBUTE = "data-selection-frame";
const BOX_CLASS = "spreadsheet-window__selection-glide";

/** The frame the cells draw, in the stage's own (unzoomed) pixels: one
 * pixel past the perimeter cells on every side, as their ::after layer.
 * With the range, from its corner cells (what the grid is about to paint);
 * without, or with a corner a merge covers, from the painted perimeter. */
export function measureSelectionFrame(stage: HTMLElement, range?: GlideRange | null): GlideFrame | null {
  const stageRect = stage.getBoundingClientRect();
  const scale = stage.offsetWidth > 0 ? stageRect.width / stage.offsetWidth : 1;
  if (!(scale > 0)) return null;
  // A cell's box takes half of each collapsed gridline; its frame layer
  // reaches 1px past the padding box, so half a pixel past this box.
  const page = stage.hasAttribute(PAGE_LAYOUT_STAGE_ATTRIBUTE);
  const before = page ? PAGE_FRAME_OUTSET_BEFORE_PX : FRAME_OUTSET_PX;
  const after = page ? PAGE_FRAME_OUTSET_AFTER_PX : FRAME_OUTSET_PX;
  const inStage = (left: number, top: number, right: number, bottom: number): GlideFrame => ({
    left: (left - stageRect.left) / scale - before,
    top: (top - stageRect.top) / scale - before,
    width: (right - left) / scale + before + after,
    height: (bottom - top) / scale + before + after,
  });
  if (range) {
    const first = rangeCellAt(stage, `${range.r1}:${range.c1}`);
    const last = rangeCellAt(stage, `${range.r2}:${range.c2}`);
    if (first && last) {
      if ([first, last].some((cell) => cell.hasAttribute("data-frozen") || cell.hasAttribute("data-frozen-pane"))) return null;
      const a = first.getBoundingClientRect();
      const b = last.getBoundingClientRect();
      return inStage(Math.min(a.left, b.left), Math.min(a.top, b.top), Math.max(a.right, b.right), Math.max(a.bottom, b.bottom));
    }
  }
  let top = Infinity;
  let bottom = -Infinity;
  let start = Infinity;
  let end = -Infinity;
  const perimeter = stage.querySelectorAll<HTMLElement>(
    "td[data-range-cell][data-selection-edges], [data-page-layout-cell][data-range-cell][data-selection-edges]");
  const cells = perimeter.length > 0
    ? Array.from(perimeter)
    : Array.from(stage.querySelectorAll<HTMLElement>(
      'td[data-range-cell][data-active="true"], [data-page-layout-cell][data-range-cell][data-active="true"]'));
  if (cells.length === 0) return null;
  for (const cell of cells) {
    // A frozen cell is pinned while the stage scrolls under it.
    if (cell.hasAttribute("data-frozen") || cell.hasAttribute("data-frozen-pane")) return null;
    const rect = cell.getBoundingClientRect();
    const tokens = perimeter.length > 0 ? (cell.dataset.selectionEdges ?? "").split(" ") : ["top", "bottom", "start", "end"];
    if (tokens.includes("top")) top = Math.min(top, rect.top);
    if (tokens.includes("bottom")) bottom = Math.max(bottom, rect.bottom);
    if (tokens.includes("start")) start = Math.min(start, rect.left);
    if (tokens.includes("end")) end = Math.max(end, rect.right);
  }
  // Every side must be on the page, or the box would land short of the frame.
  if (![top, bottom, start, end].every(Number.isFinite)) return null;
  return inStage(start, top, end, bottom);
}

/** A range's frame from its corner cells alone (null when one is not
 * rendered, or frozen). */
export function measureRangeCorners(stage: HTMLElement, range: GlideRange): GlideFrame | null {
  const first = rangeCellAt(stage, `${range.r1}:${range.c1}`);
  const last = rangeCellAt(stage, `${range.r2}:${range.c2}`);
  if (!first || !last) return null;
  return measureSelectionFrame(stage, range);
}

function settled(a: GlideFrame, b: GlideFrame): boolean {
  return (
    Math.abs(a.left - b.left) < GLIDE_SETTLE_PX &&
    Math.abs(a.top - b.top) < GLIDE_SETTLE_PX &&
    Math.abs(a.width - b.width) < GLIDE_SETTLE_PX &&
    Math.abs(a.height - b.height) < GLIDE_SETTLE_PX
  );
}

/** A box over the stage that eases from one frame to the next. It knows
 * where the frame is even while hidden, so the next leg sets out from there. */
function createGlidingBox(options: { className: string; testId: string; onLanded: () => void }) {
  let box: HTMLDivElement | null = null;
  let current: GlideFrame | null = null;
  let target: GlideFrame | null = null;
  let from: GlideFrame | null = null;
  let startedAt = 0;
  let frame = 0;

  const place = (at: GlideFrame) => {
    if (!box) return;
    box.style.transform = `translate(${at.left}px, ${at.top}px)`;
    box.style.width = `${Math.max(0, at.width)}px`;
    box.style.height = `${Math.max(0, at.height)}px`;
  };
  const stop = () => {
    if (frame) cancelAnimationFrame(frame);
    frame = 0;
    target = null;
    from = null;
  };
  const step = (now: number) => {
    if (!from || !target) {
      stop();
      return;
    }
    const to = target;
    const p = Math.min(1, Math.max(0, now - startedAt) / GLIDE_DURATION_MS);
    const k = easeOutCubic(p);
    current = {
      left: from.left + (to.left - from.left) * k,
      top: from.top + (to.top - from.top) * k,
      width: from.width + (to.width - from.width) * k,
      height: from.height + (to.height - from.height) * k,
    };
    if (p >= 1 || settled(current, to)) {
      current = to;
      place(current);
      stop();
      options.onLanded();
      return;
    }
    place(current);
    frame = requestAnimationFrame(step);
  };

  return {
    get element() {
      return box;
    },
    get moving() {
      return frame !== 0;
    },
    get current() {
      return current;
    },
    /** The box, in this stage. */
    in(host: HTMLElement): HTMLDivElement {
      if (!box) {
        box = host.ownerDocument.createElement("div");
        box.className = options.className;
        box.setAttribute("aria-hidden", "true");
        box.dataset.testid = options.testId;
        box.hidden = true;
      }
      if (box.parentElement !== host) host.append(box);
      return box;
    },
    /** Where the frame is now, without moving anything. */
    rest(at: GlideFrame) {
      stop();
      current = at;
      place(at);
    },
    /** True when a leg starts (or one in flight is redirected). */
    glideTo(to: GlideFrame): boolean {
      if (!current) {
        current = to;
        place(to);
        return false;
      }
      if (target ? settled(target, to) : settled(current, to)) return false;
      target = to;
      from = current;
      startedAt = performance.now();
      if (!frame) {
        place(current);
        frame = requestAnimationFrame(step);
      }
      return true;
    },
    hide() {
      stop();
      if (box) box.hidden = true;
    },
    forget() {
      stop();
      current = null;
      if (box) box.hidden = true;
    },
    dispose() {
      stop();
      box?.remove();
      box = null;
      current = null;
    },
  };
}

export function createSelectionGlide(options: SelectionGlideOptions) {
  let stage: HTMLElement | null = null;
  let scope: unknown = undefined;
  // A Page Layout page draws its frame with this box, landed or in flight:
  // its cells clip what they paint, so a frame of their own could not reach
  // out over the gridlines as the grid's does. Elsewhere the box shows only
  // in flight and the cells draw the frame again where it lands.
  const keepsFrame = (host: HTMLElement | null) => host?.hasAttribute(PAGE_LAYOUT_STAGE_ATTRIBUTE) === true;
  const glide = createGlidingBox({
    className: BOX_CLASS,
    testId: "spreadsheet-selection-glide",
    // Landed: the cells draw the frame again (a page keeps the box).
    onLanded: () => {
      stage?.removeAttribute(SELECTION_GLIDE_ATTRIBUTE);
      if (!keepsFrame(stage)) glide.hide();
    },
  });
  const land = () => {
    glide.hide();
    stage?.removeAttribute(SELECTION_GLIDE_ATTRIBUTE);
    stage?.removeAttribute(SELECTION_FRAME_ATTRIBUTE);
  };

  return {
    /** The selection moved (painted onto the cells, or pressed and about to
     * be), or its cells were laid out again: carry the frame there. */
    follow(range?: GlideRange | null) {
      const host = options.getStage(range);
      const owner = options.scope();
      if (host !== stage || owner !== scope) {
        land();
        glide.forget();
        stage = host;
        scope = owner;
      }
      const keep = keepsFrame(host);
      if (!host || !options.enabled() || (reducedMotionNow() && !keep)) {
        land();
        glide.forget();
        return;
      }
      const next = measureSelectionFrame(host, range);
      if (!next) {
        land();
        glide.forget();
        return;
      }
      if (!glide.moving && glide.current && settled(glide.current, next) && (!keep || glide.element?.hidden === false)) return;
      const box = glide.in(host);
      // Without motion the frame is placed where it belongs at once.
      if (reducedMotionNow()) glide.rest(next);
      else if (!glide.glideTo(next) && !keep) return;
      // The box draws the fill handle only where the cells draw one.
      box.toggleAttribute("data-handle", Boolean(host.querySelector(".spreadsheet-window__fill-handle")));
      box.hidden = false;
      if (glide.moving) host.setAttribute(SELECTION_GLIDE_ATTRIBUTE, "true");
      if (keep) host.setAttribute(SELECTION_FRAME_ATTRIBUTE, "true");
    },
    /** Drop any glide in flight and forget where the frame was. */
    reset() {
      land();
      glide.forget();
    },
    dispose() {
      land();
      glide.dispose();
      stage = null;
    },
  };
}

/**
 * A frame that IS the drawing while it shows (a fill-handle drag's target, a
 * range dragged by its border to a new place): it glides from one range to
 * the next as the pointer crosses cells, and goes when the drag ends.
 */
export function createRangeFrame(options: {
  readonly getStage: (range?: GlideRange | null) => HTMLElement | null;
  readonly className: string;
  readonly testId: string;
  /** Set on the stage while the frame shows (the cells' own marks step back). */
  readonly stageAttribute: string;
}) {
  let stage: HTMLElement | null = null;
  const glide = createGlidingBox({ className: options.className, testId: options.testId, onLanded: () => {} });
  const hide = () => {
    glide.forget();
    stage?.removeAttribute(options.stageAttribute);
  };
  return {
    show(range: GlideRange) {
      const host = options.getStage(range);
      if (host !== stage) {
        hide();
        stage = host;
      }
      if (!host) return;
      const next = measureRangeCorners(host, range);
      // A corner off the rendered cells (or frozen): keep the frame where it is.
      if (!next) return;
      const box = glide.in(host);
      if (box.hidden || reducedMotionNow()) glide.rest(next);
      else glide.glideTo(next);
      box.hidden = false;
      host.setAttribute(options.stageAttribute, "true");
    },
    hide,
    dispose() {
      hide();
      glide.dispose();
      stage = null;
    },
  };
}

export type SelectionGlide = ReturnType<typeof createSelectionGlide>;
