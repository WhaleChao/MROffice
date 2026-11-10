/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Object alignment snapping ("smart guides") for the editors that place
 * objects in CSS px: a sheet's drawing layer, a Word page. The geometry is
 * the deck's (`presentation-align`: the moving selection's edges and centre
 * to every other object's, equal spacing in a gap, a resize's dragged edge);
 * this module adapts it to px surfaces and draws the guides.
 *
 * The rules (Office's Smart Guides / Word's alignment guides, which match
 * Excalidraw's object snapping):
 * - a move snaps the selection's left / centre / right (top / middle /
 *   bottom) to another object's, or to a frame's (the page, its margins, the
 *   sheet's corner); an ellipse's snap points are its edge midpoints, which
 *   line up on the same coordinates as its box and only change where a
 *   guide line ends;
 * - an axis no edge or centre took may centre the selection in a gap
 *   between two objects or repeat a gap already there beside a neighbour
 *   (never against a frame);
 * - the snap distance is 8 screen px, so a caller passes its units per
 *   screen px ({@link objectSnapThreshold});
 * - a resize snaps its dragged edges to edges and centres only, and never
 *   for a rotated single object (the caller does not ask);
 * - a new object's press and far corner snap as points ({@link snapObjectPoint});
 * - Alt (Command on a Mac) held while dragging suspends every snap
 *   ({@link objectSnapSuspended}, the deck's override key);
 * - guides are thin accent lines spanning the objects that line up, and a
 *   gap guide is a short line with end ticks between two objects; they go at
 *   pointer up ({@link paintObjectSnapGuides} with null).
 */

import { clipToVisibleEditor } from "../components/console/editor/native-document/visibleEditorClip";
import {
  snapMovingSelection,
  snapResizePointer,
  type BlockGeometry,
  type SpacingGuide,
} from "./presentation-align";
import { presentationDragSnapOverride } from "./presentation-drag";
import { isMacPlatform } from "./mac-platform";

/** The snap distance in screen px. */
export const OBJECT_SNAP_DISTANCE_PX = 8;

/** The deck's geometry rounds to whole units (EMU there): px are carried in
 * thousandths so a fractional box keeps its place. */
const UNIT = 1000;
/** Two coordinates closer than this (scene units) line up. */
const ALIGNED_EPSILON = 0.01;

/** A box on the surface, in the caller's units. */
export interface ObjectSnapRect {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /** An ellipse or a diamond: its snap points are its edge midpoints. */
  readonly round?: boolean;
}

/** What a gesture snaps against, in the caller's units. */
export interface ObjectSnapScene {
  /** The other objects (never the moving ones, never ink strokes). */
  readonly others: readonly ObjectSnapRect[];
  /** Boxes lined up with but never spaced against: the page and its margin
   * box, a sheet's corner (a zero-size box). */
  readonly frames?: readonly ObjectSnapRect[];
  /** The snap distance in the caller's units ({@link objectSnapThreshold}). */
  readonly threshold: number;
}

/** An alignment guide: a line at `position` (x for vertical, y for
 * horizontal) from `from` to `to` along the other axis. */
export interface ObjectSnapLine {
  readonly orientation: "vertical" | "horizontal";
  readonly position: number;
  readonly from: number;
  readonly to: number;
}

/** A gap guide: each segment one gap, along `axis` from `from` to `to`, at
 * `at` across it. */
export interface ObjectSnapGap {
  readonly axis: "x" | "y";
  readonly segments: ReadonlyArray<{ readonly from: number; readonly to: number; readonly at: number }>;
}

export interface ObjectSnapGuides {
  readonly lines: readonly ObjectSnapLine[];
  readonly gaps: readonly ObjectSnapGap[];
}

export interface ObjectSnapMove {
  readonly dx: number;
  readonly dy: number;
  readonly guides: ObjectSnapGuides;
  readonly snappedX: boolean;
  readonly snappedY: boolean;
}

/** The preset shapes whose snap points are their edge midpoints (their
 * corners are empty space): an ellipse, a diamond. */
const ROUND_PRESETS: ReadonlySet<string> = new Set(["ellipse", "diamond", "flowChartConnector", "flowChartDecision"]);

/** Whether a preset shape snaps by its edge midpoints. */
export function officeShapeSnapsRound(preset: string | null | undefined): boolean {
  return typeof preset === "string" && ROUND_PRESETS.has(preset);
}

/** The snap distance in a surface's units, from its units per screen px
 * (1 / zoom for a surface laid out in CSS px and drawn at `zoom`). */
export function objectSnapThreshold(unitsPerScreenPx: number): number {
  const ratio = Number.isFinite(unitsPerScreenPx) && unitsPerScreenPx > 0 ? unitsPerScreenPx : 1;
  return OBJECT_SNAP_DISTANCE_PX * ratio;
}

/** Whether the drag's modifiers suspend snapping for now (Alt; Command on a
 * Mac, where Option is the copy key), as Office's override does. */
export function objectSnapSuspended(
  event: { readonly altKey?: boolean; readonly metaKey?: boolean },
  mac: boolean = isMacPlatform(),
): boolean {
  return presentationDragSnapOverride(event, mac);
}

const scaled = (rect: ObjectSnapRect): BlockGeometry => ({
  left: Math.round(rect.left * UNIT),
  top: Math.round(rect.top * UNIT),
  width: Math.max(0, Math.round(rect.width * UNIT)),
  height: Math.max(0, Math.round(rect.height * UNIT)),
});

const unionOf = (rects: readonly ObjectSnapRect[]): ObjectSnapRect => {
  const left = Math.min(...rects.map((rect) => rect.left));
  const top = Math.min(...rects.map((rect) => rect.top));
  return {
    left,
    top,
    width: Math.max(...rects.map((rect) => rect.left + rect.width)) - left,
    height: Math.max(...rects.map((rect) => rect.top + rect.height)) - top,
    round: rects.length === 1 ? rects[0].round : false,
  };
};

/**
 * The places along the other axis where `rect` touches the line at
 * `coord` on `axis` (empty when it does not): a box's edge from corner to
 * corner, its centre line at its centre point; an ellipse's edge at its
 * midpoint, its centre line through its two midpoints.
 */
function touchesOf(rect: ObjectSnapRect, axis: "x" | "y", coord: number, frame = false): number[] {
  const start = axis === "x" ? rect.left : rect.top;
  const size = axis === "x" ? rect.width : rect.height;
  const crossStart = axis === "x" ? rect.top : rect.left;
  const crossSize = axis === "x" ? rect.height : rect.width;
  const crossMid = crossStart + crossSize / 2;
  const near = (value: number) => Math.abs(value - coord) <= ALIGNED_EPSILON;
  const out: number[] = [];
  const edge = near(start) || near(start + size);
  const centre = near(start + size / 2);
  // A frame's line (a margin, the page's centre) runs its whole length.
  if (frame && (edge || centre)) return [crossStart, crossStart + crossSize];
  if (edge) out.push(...(rect.round ? [crossMid] : [crossStart, crossStart + crossSize]));
  if (centre) out.push(...(rect.round ? [crossStart, crossStart + crossSize] : [crossMid]));
  return out;
}

/**
 * Every alignment line the moving box makes on `axis` with the scene: each
 * of its left / centre / right that another box (or frame) lines up with,
 * drawn across all of them.
 */
function alignmentLines(
  moving: ObjectSnapRect,
  scene: ObjectSnapScene,
  axis: "x" | "y",
): ObjectSnapLine[] {
  const start = axis === "x" ? moving.left : moving.top;
  const size = axis === "x" ? moving.width : moving.height;
  const coords = [start, start + size / 2, start + size];
  const lines: ObjectSnapLine[] = [];
  for (const coord of coords) {
    if (lines.some((line) => Math.abs(line.position - coord) <= ALIGNED_EPSILON)) continue;
    const touches = [
      ...scene.others.flatMap((rect) => touchesOf(rect, axis, coord)),
      ...(scene.frames ?? []).flatMap((rect) => touchesOf(rect, axis, coord, true)),
    ];
    if (touches.length === 0) continue;
    const all = [...touches, ...touchesOf(moving, axis, coord)];
    lines.push({
      orientation: axis === "x" ? "vertical" : "horizontal",
      position: coord,
      from: Math.min(...all),
      to: Math.max(...all),
    });
  }
  return lines;
}

const unscaledGap = (guide: SpacingGuide): ObjectSnapGap => ({
  axis: guide.axis,
  segments: guide.segments.map((segment) => ({
    from: segment.from / UNIT,
    to: segment.to / UNIT,
    at: segment.at / UNIT,
  })),
});

/**
 * Snap a move of `moving` (the boxes at the gesture's start) by the
 * requested `(dx, dy)`. The selection snaps as one box, its union. `lock`
 * is the axis Shift holds the drag on: the other axis never snaps.
 */
export function snapObjectMove(
  moving: readonly ObjectSnapRect[],
  dx: number,
  dy: number,
  scene: ObjectSnapScene,
  options: { readonly gaps?: boolean; readonly lock?: "horizontal" | "vertical" | null } = {},
): ObjectSnapMove {
  const none: ObjectSnapMove = { dx, dy, guides: { lines: [], gaps: [] }, snappedX: false, snappedY: false };
  if (moving.length === 0 || !(scene.threshold > 0)) return none;
  const threshold = scene.threshold * UNIT;
  const snapped = snapMovingSelection(
    moving.map(scaled),
    Math.round(dx * UNIT),
    Math.round(dy * UNIT),
    scene.others.map(scaled),
    null,
    {
      // The axis Shift holds the drag off takes no snap.
      x: options.lock === "vertical" ? 0 : threshold,
      y: options.lock === "horizontal" ? 0 : threshold,
    },
    { spacing: options.gaps ?? true, frames: (scene.frames ?? []).map(scaled) },
  );
  const nextDx = snapped.snappedX ? snapped.dx / UNIT : dx;
  const nextDy = snapped.snappedY ? snapped.dy / UNIT : dy;
  const union = unionOf(moving);
  const placed = { ...union, left: union.left + nextDx, top: union.top + nextDy };
  return {
    dx: nextDx,
    dy: nextDy,
    guides: {
      lines: [
        ...(snapped.snappedX ? alignmentLines(placed, scene, "x") : []),
        ...(snapped.snappedY ? alignmentLines(placed, scene, "y") : []),
      ],
      gaps: snapped.spacing.map(unscaledGap),
    },
    snappedX: snapped.snappedX,
    snappedY: snapped.snappedY,
  };
}

/**
 * Snap one point (a new object's press or far corner) to the scene's edges
 * and centres.
 */
export function snapObjectPoint(
  point: { readonly x: number; readonly y: number },
  scene: ObjectSnapScene,
): { x: number; y: number; guides: ObjectSnapGuides } {
  const result = snapObjectMove([{ left: point.x, top: point.y, width: 0, height: 0 }], 0, 0, scene, {
    gaps: false,
  });
  // The geometry gives a box at least one unit wide: a point lands within
  // that unit of its line, so it is put on the line itself.
  const onLine = (value: number, axis: "x" | "y") => {
    const rects = [...scene.others, ...(scene.frames ?? [])];
    for (const rect of rects) {
      const start = axis === "x" ? rect.left : rect.top;
      const size = axis === "x" ? rect.width : rect.height;
      for (const coord of [start, start + size / 2, start + size]) {
        if (Math.abs(coord - value) <= 2 / UNIT) return coord;
      }
    }
    return value;
  };
  const x = result.snappedX ? onLine(point.x + result.dx, "x") : point.x;
  const y = result.snappedY ? onLine(point.y + result.dy, "y") : point.y;
  const placed = { left: x, top: y, width: 0, height: 0 };
  return {
    x,
    y,
    guides: {
      lines: [
        ...(result.snappedX ? alignmentLines(placed, scene, "x") : []),
        ...(result.snappedY ? alignmentLines(placed, scene, "y") : []),
      ],
      gaps: [],
    },
  };
}

/**
 * Snap a resize, worked through the pointer: `evaluate` answers the box the
 * caller's own resize math makes for a pointer offset (its minimum size,
 * locked ratio and resize from the centre included), in scene units; the
 * answer is the pointer offset that puts the dragged edges on an edge or a
 * centre of the scene. Points only: no equal size, no equal spacing.
 */
export function snapObjectResize(request: {
  readonly evaluate: (dx: number, dy: number) => ObjectSnapRect | null;
  readonly dx: number;
  readonly dy: number;
  /** The compass handle (`n`, `se`, ...). */
  readonly handle: string;
  readonly scene: ObjectSnapScene;
  readonly centered?: boolean;
  readonly lockAspect?: boolean;
}): { dx: number; dy: number; guides: ObjectSnapGuides } {
  const { scene } = request;
  const none = { dx: request.dx, dy: request.dy, guides: { lines: [], gaps: [] } };
  if (!(scene.threshold > 0)) return none;
  // Pointer offsets and boxes both go to thousandths of the caller's unit.
  const evaluate = (dx: number, dy: number) => {
    const box = request.evaluate(dx / UNIT, dy / UNIT);
    return box ? scaled(box) : null;
  };
  const snapped = snapResizePointer({
    evaluate,
    dx: request.dx * UNIT,
    dy: request.dy * UNIT,
    handle: request.handle,
    others: [...scene.others, ...(scene.frames ?? [])].map(scaled),
    slideBounds: null,
    threshold: scene.threshold * UNIT,
    objects: true,
    smartGuides: false,
    centered: request.centered,
    lockAspect: request.lockAspect,
  });
  if (!snapped.snappedX && !snapped.snappedY) return none;
  const result = { ...snapped, dx: snapped.dx / UNIT, dy: snapped.dy / UNIT };
  const box = request.evaluate(result.dx, result.dy);
  if (!box) return none;
  return {
    dx: result.dx,
    dy: result.dy,
    guides: {
      lines: [
        ...(result.snappedX ? alignmentLines(box, scene, "x") : []),
        ...(result.snappedY ? alignmentLines(box, scene, "y") : []),
      ],
      gaps: [],
    },
  };
}

// ---------------------------------------------------------------------------
// Drawing the guides.
// ---------------------------------------------------------------------------

const OVERLAY_ATTRIBUTE = "data-office-object-snap-guides";
const FALLBACK_ACCENT = "#0d9488";
const TICK_PX = 6;

/** Where a scene point is on screen (client px). */
export type ObjectSnapToClient = (x: number, y: number) => { x: number; y: number };

/**
 * Draw `guides` over the page (one fixed overlay per document, outside every
 * editor's own DOM, so no editor sees a foreign node), or take them away
 * with null. `toClient` maps the scene to the screen; `colorFrom` is the
 * element whose console accent the guides take.
 */
export function paintObjectSnapGuides(
  doc: Document,
  guides: ObjectSnapGuides | null,
  toClient?: ObjectSnapToClient,
  colorFrom?: Element | null,
): void {
  const existing = doc.querySelector<HTMLElement>(`[${OVERLAY_ATTRIBUTE}]`);
  if (!guides || !toClient || (guides.lines.length === 0 && guides.gaps.length === 0)) {
    existing?.remove();
    return;
  }
  const overlay = existing ?? doc.createElement("div");
  if (!existing) {
    overlay.setAttribute(OVERLAY_ATTRIBUTE, "true");
    overlay.setAttribute("aria-hidden", "true");
    overlay.dataset.testid = "office-object-snap-guides";
    Object.assign(overlay.style, {
      position: "fixed",
      inset: "0",
      pointerEvents: "none",
      zIndex: "2147483000",
      overflow: "hidden",
    });
    doc.body.append(overlay);
  }
  const view = doc.defaultView;
  const accent =
    (colorFrom && view ? view.getComputedStyle(colorFrom).getPropertyValue("--console-accent").trim() : "") ||
    FALLBACK_ACCENT;
  const parts: HTMLElement[] = [];
  const segment = (x1: number, y1: number, x2: number, y2: number, kind: string) => {
    const part = doc.createElement("div");
    part.dataset.snapGuide = kind;
    const vertical = Math.abs(x2 - x1) < Math.abs(y2 - y1);
    Object.assign(part.style, {
      position: "absolute",
      left: `${Math.min(x1, x2) - (vertical ? 0.5 : 0)}px`,
      top: `${Math.min(y1, y2) - (vertical ? 0 : 0.5)}px`,
      width: vertical ? "1px" : `${Math.max(1, Math.abs(x2 - x1))}px`,
      height: vertical ? `${Math.max(1, Math.abs(y2 - y1))}px` : "1px",
      background: accent,
    });
    parts.push(part);
  };
  for (const line of guides.lines) {
    const vertical = line.orientation === "vertical";
    const a = vertical ? toClient(line.position, line.from) : toClient(line.from, line.position);
    const b = vertical ? toClient(line.position, line.to) : toClient(line.to, line.position);
    segment(a.x, a.y, b.x, b.y, vertical ? "vertical" : "horizontal");
  }
  for (const gap of guides.gaps) {
    for (const piece of gap.segments) {
      const along = gap.axis === "x";
      const a = along ? toClient(piece.from, piece.at) : toClient(piece.at, piece.from);
      const b = along ? toClient(piece.to, piece.at) : toClient(piece.at, piece.to);
      segment(a.x, a.y, b.x, b.y, "gap");
      // The end ticks, across the gap.
      for (const end of [a, b]) {
        if (along) segment(end.x, end.y - TICK_PX / 2, end.x, end.y + TICK_PX / 2, "gap-tick");
        else segment(end.x - TICK_PX / 2, end.y, end.x + TICK_PX / 2, end.y, "gap-tick");
      }
    }
  }
  // The overlay is on the body, over every window: cut it to what the
  // surface's scroller and window leave visible, so a guide running past
  // them is not drawn over the ribbon, another window or the desk (the Docs
  // drawn caret's rule, visibleEditorClip).
  const width = view?.innerWidth ?? 0;
  const height = view?.innerHeight ?? 0;
  const shown =
    colorFrom instanceof HTMLElement && width > 0 && height > 0
      ? clipToVisibleEditor(colorFrom, { left: 0, top: 0, width, height })
      : { left: 0, top: 0, width, height };
  if (!shown) {
    overlay.remove();
    return;
  }
  overlay.style.clipPath =
    width > 0 && height > 0
      ? `inset(${shown.top}px ${Math.max(0, width - shown.left - shown.width)}px ${Math.max(0, height - shown.top - shown.height)}px ${shown.left}px)`
      : "";
  overlay.replaceChildren(...parts);
}
