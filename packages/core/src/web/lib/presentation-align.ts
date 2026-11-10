/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry math for PowerPoint-parity alignment of presentation blocks on
 * the editor canvas:
 *  - smart alignment guides + edge/center snapping while a block is dragged
 *    ({@link computeSnapGuides}),
 *  - multi-select alignment to a shared edge / center ({@link alignBlocks}),
 *  - even spacing distribution ({@link distributeBlocks}).
 *
 * Dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * reused by both the editor component and tests. All values are in the same EMU
 * coordinate space the rest of the presentation stack uses (`left`/`top`/
 * `width`/`height` on a block); the caller commits the returned geometry through
 * the existing block-geometry update path so undo/redo + autosave work (one
 * history step per operation).
 *
 * Mirrors the conventions of the sibling pure module `presentation-drag.ts`
 * (shared {@link BlockGeometry} shape, integer-EMU rounded output, frozen
 * allowed lists, small composable pure functions, no mutation of inputs).
 */

import { normalizeGeometry, type BlockGeometry } from "./presentation-drag";

export type { BlockGeometry } from "./presentation-drag";

/** The slide canvas extent, in EMU (matches `presentation-drag`'s bounds). */
export interface SlideBounds {
  width: number;
  height: number;
}

/**
 * The six alignment edges. `left`/`hcenter`/`right` act on the horizontal axis
 * (x), `top`/`vcenter`/`bottom` on the vertical axis (y). Named to mirror the
 * PowerPoint Align menu (Align Left / Center / Right, Align Top / Middle /
 * Bottom — "middle"/"center" map to `vcenter`/`hcenter`).
 */
export const ALIGN_EDGES = [
  "left",
  "hcenter",
  "right",
  "top",
  "vcenter",
  "bottom",
] as const;

export type AlignEdge = (typeof ALIGN_EDGES)[number];

const ALIGN_EDGE_SET: ReadonlySet<string> = new Set(ALIGN_EDGES);

/** The two distribution axes (even spacing). */
export const DISTRIBUTE_AXES = ["horizontal", "vertical"] as const;

export type DistributeAxis = (typeof DISTRIBUTE_AXES)[number];

const DISTRIBUTE_AXIS_SET: ReadonlySet<string> = new Set(DISTRIBUTE_AXES);

/** Default snap threshold (EMU). ~0.6% of a 9.14M-EMU-wide 16:9 slide. */
export const DEFAULT_SNAP_THRESHOLD = 50_000;

/** Type guard: is `value` exactly one of the six alignment edges? */
export function isAlignEdge(value: unknown): value is AlignEdge {
  return typeof value === "string" && ALIGN_EDGE_SET.has(value);
}

/** Type guard: is `value` exactly one of the two distribute axes? */
export function isDistributeAxis(value: unknown): value is DistributeAxis {
  return typeof value === "string" && DISTRIBUTE_AXIS_SET.has(value);
}

/** Round to an integer, guarding against NaN/Infinity (→ 0). */
function safeRound(value: number): number {
  return Number.isFinite(value) ? Math.round(value) : 0;
}

/**
 * A single smart-guide line emitted by {@link computeSnapGuides}. `vertical`
 * guides run top→bottom at x === `position`; `horizontal` guides run left→right
 * at y === `position`. The `kind` records which feature produced the match so
 * the renderer / tests can distinguish a slide-centre snap from a block-edge
 * snap. Positions are integer EMU in slide space.
 */
export interface SnapGuide {
  /** Line orientation. A `vertical` line snaps the x axis; `horizontal` the y. */
  orientation: "vertical" | "horizontal";
  /** The line's coordinate in EMU (x for vertical, y for horizontal). */
  position: number;
  /** What the moving block snapped to (for styling / debugging / tests). */
  kind: "slide-edge" | "slide-center" | "block-edge" | "block-center";
}

export interface SnapResult {
  /** The moving block's geometry after snapping (size preserved, integer EMU). */
  snappedGeom: BlockGeometry;
  /** Guide lines to draw for the current frame (deduped, may be empty). */
  guides: SnapGuide[];
}

/**
 * Per-axis snap thresholds (EMU). The x value gates horizontal-axis snapping
 * (vertical guide lines), the y value gates vertical-axis snapping (horizontal
 * guide lines). Used so the on-screen pixel tolerance stays symmetric on slides
 * whose px→EMU scale differs between axes (a non-uniformly-scaled canvas).
 */
export interface SnapThresholds {
  /** Horizontal-axis (x) threshold in EMU. */
  x: number;
  /** Vertical-axis (y) threshold in EMU. */
  y: number;
}

/** Normalize a scalar OR per-axis threshold argument into a {@link SnapThresholds}. */
function resolveThresholds(
  threshold: number | SnapThresholds,
): SnapThresholds {
  if (typeof threshold === "number") {
    return { x: threshold, y: threshold };
  }
  return {
    x: Number.isFinite(threshold?.x) ? threshold.x : Number.NaN,
    y: Number.isFinite(threshold?.y) ? threshold.y : Number.NaN,
  };
}

/** The three horizontal reference coordinates of a rect: left, h-center, right. */
function xTargets(geom: BlockGeometry): [number, number, number] {
  return [geom.left, geom.left + geom.width / 2, geom.left + geom.width];
}

/** The three vertical reference coordinates of a rect: top, v-center, bottom. */
function yTargets(geom: BlockGeometry): [number, number, number] {
  return [geom.top, geom.top + geom.height / 2, geom.top + geom.height];
}

/**
 * Internal: find the best (smallest-delta) single-axis snap for the moving
 * block. `movingCoords` are the moving block's three reference coords on this
 * axis (left/center/right or top/center/bottom); `candidates` are
 * `{ coord, kind }` reference lines to snap onto. Returns the winning origin
 * delta to apply on that axis plus the guide to draw, or null when nothing is
 * within `threshold`.
 *
 * Snapping moves the block ORIGIN by the matched delta; because all three of
 * the moving block's reference coords are rigidly offset from the origin, the
 * returned delta lines whichever reference coord matched up with its candidate.
 */
function bestAxisSnap(
  movingCoords: readonly [number, number, number],
  candidates: ReadonlyArray<{ coord: number; kind: SnapGuide["kind"] }>,
  threshold: number,
  orientation: SnapGuide["orientation"],
): { delta: number; guide: SnapGuide } | null {
  let best: { delta: number; guide: SnapGuide; distance: number } | null = null;
  for (const moving of movingCoords) {
    for (const candidate of candidates) {
      const distance = Math.abs(moving - candidate.coord);
      if (distance > threshold) continue;
      if (best && distance >= best.distance) continue;
      best = {
        distance,
        delta: candidate.coord - moving,
        guide: {
          orientation,
          position: safeRound(candidate.coord),
          kind: candidate.kind,
        },
      };
    }
  }
  if (!best) return null;
  return { delta: best.delta, guide: best.guide };
}

/**
 * Smart alignment guides + snapping for a dragged block. Compares the moving
 * block's left / center / right (x) and top / center / bottom (y) reference
 * lines against:
 *  - the slide's left/right edges and horizontal center (x),
 *  - the slide's top/bottom edges and vertical center (y),
 *  - every OTHER block's left/center/right and top/center/bottom.
 *
 * On each axis independently it snaps to the single closest candidate within
 * `threshold` (EMU), moving the block's ORIGIN so the matched reference line
 * lands exactly on the candidate, and emits a guide line for that match. The
 * block's size is never changed. The origin is NOT clamped to the slide here
 * (the caller's commit path already clamps to the persisted geometry limits, and
 * snapping to a slide edge keeps it on-canvas); only finite, integer-EMU output
 * is guaranteed.
 *
 * `threshold` defaults to {@link DEFAULT_SNAP_THRESHOLD} and may be either a
 * single scalar (same tolerance on both axes — kept for tests and simple
 * callers) or a per-axis {@link SnapThresholds} `{ x, y }` so a canvas whose
 * px→EMU scale differs between axes can keep a symmetric on-screen pixel
 * tolerance. On each axis a non-positive / non-finite threshold disables
 * snapping for THAT axis (returning the normalized origin on it, no guide).
 */
export function computeSnapGuides(
  movingGeom: Partial<BlockGeometry> | null | undefined,
  others: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  /** Null for a surface with no slide of its own to line up with (a sheet,
   * a page passed as one of `others`): only the other boxes snap. */
  slideBounds: SlideBounds | null,
  threshold: number | SnapThresholds = DEFAULT_SNAP_THRESHOLD,
): SnapResult {
  const moving = normalizeGeometry(movingGeom);
  const { x: thresholdX, y: thresholdY } = resolveThresholds(threshold);
  const xEnabled = thresholdX > 0 && Number.isFinite(thresholdX);
  const yEnabled = thresholdY > 0 && Number.isFinite(thresholdY);
  if (!xEnabled && !yEnabled) {
    return { snappedGeom: moving, guides: [] };
  }
  const boundsWidth = slideBounds && Number.isFinite(slideBounds.width)
    ? slideBounds.width
    : 0;
  const boundsHeight = slideBounds && Number.isFinite(slideBounds.height)
    ? slideBounds.height
    : 0;

  // Horizontal (x) candidates: slide edges + center, then every other block's
  // left / h-center / right.
  const xCandidates: Array<{ coord: number; kind: SnapGuide["kind"] }> = slideBounds
    ? [
        { coord: 0, kind: "slide-edge" },
        { coord: boundsWidth / 2, kind: "slide-center" },
        { coord: boundsWidth, kind: "slide-edge" },
      ]
    : [];
  const yCandidates: Array<{ coord: number; kind: SnapGuide["kind"] }> = slideBounds
    ? [
        { coord: 0, kind: "slide-edge" },
        { coord: boundsHeight / 2, kind: "slide-center" },
        { coord: boundsHeight, kind: "slide-edge" },
      ]
    : [];
  for (const other of others) {
    if (!other) continue;
    const rect = normalizeGeometry(other);
    const [xLeft, xCenter, xRight] = xTargets(rect);
    xCandidates.push(
      { coord: xLeft, kind: "block-edge" },
      { coord: xCenter, kind: "block-center" },
      { coord: xRight, kind: "block-edge" },
    );
    const [yTop, yCenter, yBottom] = yTargets(rect);
    yCandidates.push(
      { coord: yTop, kind: "block-edge" },
      { coord: yCenter, kind: "block-center" },
      { coord: yBottom, kind: "block-edge" },
    );
  }

  const xSnap = xEnabled
    ? bestAxisSnap(xTargets(moving), xCandidates, thresholdX, "vertical")
    : null;
  const ySnap = yEnabled
    ? bestAxisSnap(yTargets(moving), yCandidates, thresholdY, "horizontal")
    : null;

  const snappedGeom: BlockGeometry = {
    left: safeRound(moving.left + (xSnap?.delta ?? 0)),
    top: safeRound(moving.top + (ySnap?.delta ?? 0)),
    width: moving.width,
    height: moving.height,
  };

  const guides: SnapGuide[] = [];
  if (xSnap) guides.push(xSnap.guide);
  if (ySnap) guides.push(ySnap.guide);
  return { snappedGeom, guides };
}

/**
 * Smart-guide snapping for everything a drag moves at once (PowerPoint snaps
 * the SELECTION as one object: a group, or several objects dragged together,
 * line up by their combined bounding box, never member by member, and never
 * against their own start positions). `moving` are the visible boxes (a
 * rotated object's turned box) of every moving object at the gesture's start;
 * `(dx, dy)` is the drag's requested offset. Answers the snapped offset and
 * the guides to draw; `others` must not contain any moving object.
 */
export function snapMovingSelection(
  moving: ReadonlyArray<BlockGeometry>,
  dx: number,
  dy: number,
  others: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  slideBounds: SlideBounds | null,
  threshold: number | SnapThresholds = DEFAULT_SNAP_THRESHOLD,
  /** False for "Snap objects to other objects" alone: edges and centres
   * line up, but the even gap is a Smart Guide's (and nothing is drawn by
   * the caller). `frames` are boxes the selection lines up with but never
   * spaces against (a page's margins). */
  options: {
    readonly spacing?: boolean;
    readonly frames?: ReadonlyArray<Partial<BlockGeometry> | null | undefined>;
  } = {},
): {
  dx: number;
  dy: number;
  guides: SnapGuide[];
  spacing: SpacingGuide[];
  snappedX: boolean;
  snappedY: boolean;
} {
  if (moving.length === 0) {
    return { dx, dy, guides: [], spacing: [], snappedX: false, snappedY: false };
  }
  const left = Math.min(...moving.map((rect) => rect.left));
  const top = Math.min(...moving.map((rect) => rect.top));
  const union = normalizeGeometry({
    left: left + dx,
    top: top + dy,
    width: Math.max(...moving.map((rect) => rect.left + rect.width)) - left,
    height: Math.max(...moving.map((rect) => rect.top + rect.height)) - top,
  });
  const snapped = computeSnapGuides(
    union,
    options.frames?.length ? [...others, ...options.frames] : others,
    slideBounds,
    threshold,
  );
  let nextLeft = snapped.snappedGeom.left;
  let nextTop = snapped.snappedGeom.top;
  let snappedX = snapped.guides.some((guide) => guide.orientation === "vertical");
  let snappedY = snapped.guides.some((guide) => guide.orientation === "horizontal");
  // Smart Guides also show equal spacing (Microsoft Support: they "aid you
  // in aligning and spacing one object in relation to others"): an axis no
  // edge or centre lined up may take the even gap.
  const { x: thresholdX, y: thresholdY } = resolveThresholds(threshold);
  const rects = others.filter(Boolean).map((other) => normalizeGeometry(other));
  const spacing: SpacingGuide[] = [];
  const withSpacing = options.spacing ?? true;
  if (withSpacing && !snappedX && thresholdX > 0) {
    const even = equalSpacingSnap({ ...union, left: nextLeft, top: nextTop }, rects, "x", thresholdX);
    if (even) {
      nextLeft += even.delta;
      snappedX = true;
      spacing.push(even.guide);
    }
  }
  if (withSpacing && !snappedY && thresholdY > 0) {
    const even = equalSpacingSnap({ ...union, left: nextLeft, top: nextTop }, rects, "y", thresholdY);
    if (even) {
      nextTop += even.delta;
      snappedY = true;
      spacing.push(even.guide);
    }
  }
  return {
    dx: dx + (nextLeft - union.left),
    dy: dy + (nextTop - union.top),
    guides: snapped.guides,
    spacing,
    snappedX,
    snappedY,
  };
}

/**
 * An equal-spacing Smart Guide: the gaps it shows as arrows, each a segment
 * along `axis` from `from` to `to`, drawn across the other axis at `at`.
 */
export interface SpacingGuide {
  axis: "x" | "y";
  segments: Array<{ from: number; to: number; at: number }>;
  /** "size": an equal-size Smart Guide (a resize matched another object's
   * width or height), each segment one object's extent. Absent = spacing. */
  kind?: "spacing" | "size";
}

/**
 * The equal-spacing snap of a moving box along one axis, among the objects
 * that share its band on the other axis: centred between its two nearest
 * neighbours, or a gap that two neighbouring objects already keep repeated
 * beside the nearest one. Null when none is within `threshold`.
 */
export function equalSpacingSnap(
  moving: BlockGeometry,
  others: ReadonlyArray<BlockGeometry>,
  axis: "x" | "y",
  threshold: number,
): { delta: number; guide: SpacingGuide } | null {
  const start = (r: BlockGeometry) => (axis === "x" ? r.left : r.top);
  const size = (r: BlockGeometry) => (axis === "x" ? r.width : r.height);
  const end = (r: BlockGeometry) => start(r) + size(r);
  const crossStart = (r: BlockGeometry) => (axis === "x" ? r.top : r.left);
  const crossEnd = (r: BlockGeometry) => crossStart(r) + (axis === "x" ? r.height : r.width);
  const band = others.filter(
    (r) => crossStart(r) < crossEnd(moving) && crossEnd(r) > crossStart(moving),
  );
  const center = start(moving) + size(moving) / 2;
  const before = band
    .filter((r) => end(r) <= center && start(r) + size(r) / 2 < center)
    .sort((a, b) => start(a) - start(b));
  const after = band
    .filter((r) => start(r) >= center && start(r) + size(r) / 2 > center)
    .sort((a, b) => start(a) - start(b));
  const at = crossStart(moving) + (crossEnd(moving) - crossStart(moving)) / 2;
  const gapsOf = (list: BlockGeometry[]) =>
    list.slice(1).flatMap((next, index) => {
      const gap = start(next) - end(list[index]);
      return gap > 0 ? [{ gap, from: end(list[index]), to: start(next) }] : [];
    });
  const candidates: Array<{ target: number; segments: SpacingGuide["segments"] }> = [];
  const near = before.length ? before.reduce((a, b) => (end(b) > end(a) ? b : a)) : null;
  const far = after.length ? after.reduce((a, b) => (start(b) < start(a) ? b : a)) : null;
  if (near && far && start(far) - end(near) > size(moving)) {
    const target = (end(near) + start(far) - size(moving)) / 2;
    candidates.push({
      target,
      segments: [
        { from: end(near), to: target, at },
        { from: target + size(moving), to: start(far), at },
      ],
    });
  }
  for (const reference of [...gapsOf(before), ...gapsOf(after)]) {
    const segment = { from: reference.from, to: reference.to, at };
    if (near) {
      const target = end(near) + reference.gap;
      candidates.push({ target, segments: [segment, { from: end(near), to: target, at }] });
    }
    if (far) {
      const target = start(far) - reference.gap - size(moving);
      candidates.push({
        target,
        segments: [segment, { from: target + size(moving), to: start(far), at }],
      });
    }
  }
  let best: { delta: number; segments: SpacingGuide["segments"] } | null = null;
  for (const candidate of candidates) {
    const delta = candidate.target - start(moving);
    if (Math.abs(delta) > threshold) continue;
    if (best && Math.abs(delta) >= Math.abs(best.delta)) continue;
    best = { delta, segments: candidate.segments };
  }
  if (!best) return null;
  const delta = safeRound(best.delta);
  return {
    delta,
    guide: {
      axis,
      segments: best.segments.map((segment) => ({
        from: safeRound(segment.from),
        to: safeRound(segment.to),
        at: safeRound(segment.at),
      })),
    },
  };
}

/**
 * Align a set of blocks to a shared edge / center. The reference line is taken
 * from the bounding box that encloses ALL the input rects (PowerPoint behaviour:
 * "Align Left" moves every block's left edge to the group's leftmost edge;
 * "Align Center" centers them on the group's horizontal midline; etc.). Only the
 * relevant axis origin (`left` or `top`) changes; size is preserved. Output is
 * integer EMU. Order is preserved.
 *
 * No-op safety: 0 or 1 input blocks are returned normalized but otherwise
 * unchanged (nothing to align against), and an unknown `edge` returns the inputs
 * normalized unchanged.
 */
export function alignBlocks(
  geoms: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  edge: AlignEdge,
): BlockGeometry[] {
  const rects = geoms.map((geom) => normalizeGeometry(geom));
  if (rects.length < 2 || !isAlignEdge(edge)) return rects;

  // Group bounding box.
  const minLeft = Math.min(...rects.map((r) => r.left));
  const maxRight = Math.max(...rects.map((r) => r.left + r.width));
  const minTop = Math.min(...rects.map((r) => r.top));
  const maxBottom = Math.max(...rects.map((r) => r.top + r.height));
  const groupHCenter = (minLeft + maxRight) / 2;
  const groupVCenter = (minTop + maxBottom) / 2;

  return rects.map((rect) => {
    switch (edge) {
      case "left":
        return { ...rect, left: safeRound(minLeft) };
      case "right":
        return { ...rect, left: safeRound(maxRight - rect.width) };
      case "hcenter":
        return { ...rect, left: safeRound(groupHCenter - rect.width / 2) };
      case "top":
        return { ...rect, top: safeRound(minTop) };
      case "bottom":
        return { ...rect, top: safeRound(maxBottom - rect.height) };
      case "vcenter":
        return { ...rect, top: safeRound(groupVCenter - rect.height / 2) };
      default:
        return rect;
    }
  });
}

/**
 * Align each block to the SLIDE bounds (PowerPoint / WPS "Align to Slide" mode of
 * the Align menu). Unlike {@link alignBlocks} — which aligns a multi-selection to
 * the GROUP bounding box and therefore needs >= 2 blocks — this aligns every
 * input block independently to the slide rect, so it is the affordance that lets
 * the align buttons act on a SINGLE selected object: "Align Center" centers it
 * horizontally on the slide, "Align Middle" vertically, "Align Left/Right/Top/
 * Bottom" pins it to that slide edge.
 *
 * Per edge, using slide width `W` / height `H` and the block's own width `w` /
 * height `h`:
 *   left → 0,  right → W - w,  hcenter → (W - w) / 2,
 *   top  → 0,  bottom → H - h, vcenter → (H - h) / 2.
 * Only the relevant axis origin (`left` / `top`) changes; size is preserved.
 * Output is integer EMU; order is preserved; inputs are never mutated.
 *
 * No-op safety: an unknown `edge` or a non-finite slide bound on the relevant
 * axis returns the inputs normalized unchanged (so a stray call can never move a
 * block off-canvas to NaN). A zero/empty input list returns `[]`. Works for any
 * count >= 1 (the single-object case is the whole point).
 */
export function alignBlocksToBounds(
  geoms: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  edge: AlignEdge,
  slideBounds: SlideBounds,
): BlockGeometry[] {
  const rects = geoms.map((geom) => normalizeGeometry(geom));
  if (!isAlignEdge(edge)) return rects;
  const horizontal = edge === "left" || edge === "hcenter" || edge === "right";
  const bound = horizontal ? slideBounds?.width : slideBounds?.height;
  // A degenerate slide bound on the active axis disables the move (return inputs
  // unchanged) rather than snapping every block to 0/NaN.
  if (!Number.isFinite(bound)) return rects;
  const extent = bound as number;

  return rects.map((rect) => {
    switch (edge) {
      case "left":
        return { ...rect, left: 0 };
      case "right":
        return { ...rect, left: safeRound(extent - rect.width) };
      case "hcenter":
        return { ...rect, left: safeRound((extent - rect.width) / 2) };
      case "top":
        return { ...rect, top: 0 };
      case "bottom":
        return { ...rect, top: safeRound(extent - rect.height) };
      case "vcenter":
        return { ...rect, top: safeRound((extent - rect.height) / 2) };
      default:
        return rect;
    }
  });
}

/**
 * Distribute blocks so the GAPS between adjacent blocks are even along `axis`
 * (PowerPoint "Distribute Horizontally / Vertically"). The two extreme blocks
 * (leftmost+rightmost, or topmost+bottommost) stay pinned; the interior blocks
 * are repositioned so the inter-block gaps are equal. Only the relevant axis
 * origin changes; size is preserved. Output is integer EMU, returned in the
 * SAME order as the input (the internal sort is local and does not reorder the
 * result array).
 *
 * No-op safety: fewer than 3 blocks are returned normalized but unchanged
 * (with 2 or fewer there is no interior block to redistribute), and an unknown
 * `axis` returns the inputs normalized unchanged. If total content exceeds the
 * span (negative gap) the even (possibly negative) gap is still applied so the
 * operation is deterministic.
 */
export function distributeBlocks(
  geoms: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  axis: DistributeAxis,
): BlockGeometry[] {
  const rects = geoms.map((geom) => normalizeGeometry(geom));
  if (rects.length < 3 || !isDistributeAxis(axis)) return rects;

  const horizontal = axis === "horizontal";
  const start = (r: BlockGeometry) => (horizontal ? r.left : r.top);
  const size = (r: BlockGeometry) => (horizontal ? r.width : r.height);

  // Order by CENTRE, as PowerPoint does (measured in PowerPoint for Mac,
  
  // one -- centre 160 -- first and the 100-wide one last, moving the
  // 200-wide one between them to 135). The first and last stay; the span
  // runs from the first's near edge to the last's far edge.
  const order = rects
    .map((rect, index) => ({ rect, index }))
    .sort((a, b) => start(a.rect) + size(a.rect) / 2 - (start(b.rect) + size(b.rect) / 2));

  const first = order[0].rect;
  const last = order[order.length - 1].rect;
  const span = start(last) + size(last) - start(first); // outer extent
  const totalSize = order.reduce((sum, entry) => sum + size(entry.rect), 0);
  const gapCount = order.length - 1;
  const gap = (span - totalSize) / gapCount;

  // Result starts as a normalized copy of the inputs (preserving order); we then
  // overwrite the interior blocks' axis origin. Extremes are left pinned.
  const result = rects.map((rect) => ({ ...rect }));
  let cursor = start(first) + size(first);
  for (let i = 1; i < order.length - 1; i += 1) {
    cursor += gap;
    const target = order[i];
    const origin = safeRound(cursor);
    if (horizontal) {
      result[target.index] = { ...target.rect, left: origin };
    } else {
      result[target.index] = { ...target.rect, top: origin };
    }
    cursor += size(target.rect);
  }
  return result;
}












export function distributeBlocksToBounds(
  geoms: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  axis: DistributeAxis,
  slideBounds: SlideBounds,
): BlockGeometry[] {
  const rects = geoms.map((geom) => normalizeGeometry(geom));
  if (rects.length === 0 || !isDistributeAxis(axis)) return rects;
  const horizontal = axis === "horizontal";
  const extent = horizontal ? slideBounds?.width : slideBounds?.height;
  if (!Number.isFinite(extent)) return rects;
  const span = extent as number;
  const start = (r: BlockGeometry) => (horizontal ? r.left : r.top);
  const size = (r: BlockGeometry) => (horizontal ? r.width : r.height);
  const place = (rect: BlockGeometry, origin: number): BlockGeometry =>
    horizontal ? { ...rect, left: safeRound(origin) } : { ...rect, top: safeRound(origin) };
  const order = rects
    .map((rect, index) => ({ rect, index }))
    .sort((a, b) => start(a.rect) + size(a.rect) / 2 - (start(b.rect) + size(b.rect) / 2));
  const totalSize = order.reduce((sum, entry) => sum + size(entry.rect), 0);
  // Objects wider together than the slide: PowerPoint puts the first at the
  // slide's near edge and the last at its far edge, the gaps between them
  
  // slide at 0 / 154 / 308 / 462 / 616 / 770 pt.
  const overflows = totalSize > span && order.length > 1;
  const gap = overflows ? (span - totalSize) / (order.length - 1) : (span - totalSize) / (order.length + 1);
  const result = rects.map((rect) => ({ ...rect }));
  let cursor = overflows ? 0 : gap;
  for (const entry of order) {
    result[entry.index] = place(entry.rect, cursor);
    cursor += size(entry.rect) + gap;
  }
  return result;
}

/**
 * PowerPoint's Arrange ▸ Auto Fix ("automatically align, uniformly resize,
 * and distribute elements, and straighten the connectors between them"):
 * two or more elements that do not overlap. Read as rows (or, when every row
 * holds one element, as a column): the elements are made one size along the
 * cross axis -- each keeping its own proportions, so a picture is never
 * distorted -- centred on their row's (column's) median line, and spread
 * evenly between the first and the last; a grid's columns line up too.
 * Connectors anchored to them follow through the caller's re-snap, which is
 * what straightens them.
 *
 * Returns the inputs unchanged (normalized) when Auto Fix does not apply:
 * fewer than two, or any two overlapping ({@link canAutoFixBlocks}).
 */
export function autoFixBlocks(
  geoms: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
): BlockGeometry[] {
  const rects = geoms.map((geom) => normalizeGeometry(geom));
  if (!canAutoFixBlocks(rects)) return rects;
  const median = (values: number[]) => {
    const sorted = [...values].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  };
  const centerX = (r: BlockGeometry) => r.left + r.width / 2;
  const centerY = (r: BlockGeometry) => r.top + r.height / 2;
  // Rows: elements whose vertical centres lie within half the median height.
  const tolerance = median(rects.map((r) => r.height)) / 2;
  const byY = rects.map((rect, index) => ({ rect, index })).sort((a, b) => centerY(a.rect) - centerY(b.rect));
  const rows: Array<typeof byY> = [];
  for (const entry of byY) {
    const row = rows[rows.length - 1];
    if (row && Math.abs(centerY(entry.rect) - median(row.map((e) => centerY(e.rect)))) <= tolerance) {
      row.push(entry);
    } else {
      rows.push([entry]);
    }
  }
  const asColumn = rows.every((row) => row.length === 1);
  const result = rects.map((rect) => ({ ...rect }));
  // Uniform cross-axis size, proportions kept.
  const crossSize = median(rects.map((r) => (asColumn ? r.width : r.height)));
  const sized = rects.map((r) => {
    const ratio = r.height > 0 ? r.width / r.height : 1;
    return asColumn
      ? { width: crossSize, height: ratio > 0 ? crossSize / ratio : r.height }
      : { width: crossSize * ratio, height: crossSize };
  });
  const spread = (
    entries: Array<{ index: number }>,
    axis: "x" | "y",
    from: number,
    to: number,
  ) => {
    const sizes = entries.map((entry) => (axis === "x" ? sized[entry.index].width : sized[entry.index].height));
    const total = sizes.reduce((sum, value) => sum + value, 0);
    const gap = entries.length > 1 ? Math.max(0, (to - from - total) / (entries.length - 1)) : 0;
    let cursor = entries.length > 1 ? from : from + (to - from - total) / 2;
    entries.forEach((entry, i) => {
      if (axis === "x") result[entry.index].left = safeRound(cursor);
      else result[entry.index].top = safeRound(cursor);
      cursor += sizes[i] + gap;
    });
  };
  rects.forEach((_, index) => {
    result[index].width = safeRound(sized[index].width);
    result[index].height = safeRound(sized[index].height);
  });
  if (asColumn) {
    const line = median(rects.map(centerX));
    const ordered = byY.map((entry) => entry);
    for (const entry of ordered) result[entry.index].left = safeRound(line - sized[entry.index].width / 2);
    const first = ordered[0].rect;
    const last = ordered[ordered.length - 1].rect;
    spread(ordered, "y", first.top, last.top + last.height);
    return result;
  }
  const left = Math.min(...rects.map((r) => r.left));
  const right = Math.max(...rects.map((r) => r.left + r.width));
  const grid = rows.length > 1 && rows.every((row) => row.length === rows[0].length);
  for (const row of rows) {
    const line = median(row.map((entry) => centerY(entry.rect)));
    for (const entry of row) result[entry.index].top = safeRound(line - sized[entry.index].height / 2);
    const ordered = [...row].sort((a, b) => centerX(a.rect) - centerX(b.rect));
    if (grid) {
      spread(ordered, "x", left, right);
    } else {
      const first = ordered[0].rect;
      const last = ordered[ordered.length - 1].rect;
      spread(ordered, "x", first.left, last.left + last.width);
    }
  }
  if (rows.length > 1) {
    // The rows themselves, evenly spaced between the first and the last.
    const lines = rows.map((row) => ({
      top: Math.min(...row.map((entry) => result[entry.index].top)),
      height: Math.max(...row.map((entry) => result[entry.index].top + result[entry.index].height)) -
        Math.min(...row.map((entry) => result[entry.index].top)),
      row,
    }));
    const from = lines[0].top;
    const to = lines[lines.length - 1].top + lines[lines.length - 1].height;
    const total = lines.reduce((sum, line) => sum + line.height, 0);
    const gap = Math.max(0, (to - from - total) / (lines.length - 1));
    let cursor = from;
    for (const line of lines) {
      const shift = cursor - line.top;
      for (const entry of line.row) result[entry.index].top = safeRound(result[entry.index].top + shift);
      cursor += line.height + gap;
    }
  }
  return result;
}

/** Whether Auto Fix applies: two or more elements, no two overlapping. */
export function canAutoFixBlocks(
  geoms: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
): boolean {
  const rects = geoms.map((geom) => normalizeGeometry(geom));
  if (rects.length < 2) return false;
  for (let i = 0; i < rects.length; i += 1) {
    for (let j = i + 1; j < rects.length; j += 1) {
      const a = rects[i];
      const b = rects[j];
      if (
        a.left < b.left + b.width &&
        b.left < a.left + a.width &&
        a.top < b.top + b.height &&
        b.top < a.top + a.height
      ) {
        return false;
      }
    }
  }
  return true;
}

/**
 * Snapping for a resize's dragged edges (PowerPoint's Smart Guides and
 * "Snap objects to other objects" while sizing a shape from a handle).
 *
 * On each axis whose edge the handle drags (`e`/`w` for x, `s`/`n` for y)
 * the edge goes to the closest of, within `threshold`:
 * - an alignment line: the slide's edges and centre, and every other
 *   object's edges and centre (a guide line drawn at it);
 * - with `size`, the extent that makes the box exactly as wide (tall) as
 *   another object -- PowerPoint's equal-size Smart Guide, drawn as an arrow
 *   across the resized box and every object of that size;
 * - with `spacing`, the place that makes the gap to the next object in the
 *   box's band equal to a gap already on the slide (two neighbours' gap, or
 *   the box's own gap on its pinned side) -- the equal-spacing Smart Guide.
 * The pinned edges never move, and an edge that would pass the pinned one
 * (below `minSize`) is left where the pointer put it. `snappedX` / `snappedY`
 * say which axes took a place, so the grid can take the rest.
 */
export function snapResizedEdges(
  geom: BlockGeometry,
  handle: string,
  others: ReadonlyArray<Partial<BlockGeometry> | null | undefined>,
  /** Null: no slide to line up with, only `others`. */
  slideBounds: SlideBounds | null,
  threshold: number | SnapThresholds = DEFAULT_SNAP_THRESHOLD,
  options: {
    readonly spacing?: boolean;
    readonly size?: boolean;
    readonly minSize?: number;
    /** A resize from the centre (Ctrl on Windows, Option on a Mac): the
     * centre stays, the far edge mirrors the dragged one. */
    readonly centered?: boolean;
  } = {},
): {
  geom: BlockGeometry;
  guides: SnapGuide[];
  spacing: SpacingGuide[];
  snappedX: boolean;
  snappedY: boolean;
} {
  const box = normalizeGeometry(geom);
  const { x: thresholdX, y: thresholdY } = resolveThresholds(threshold);
  const rects = others.filter(Boolean).map((other) => normalizeGeometry(other));
  const minSize = Math.max(1, options.minSize ?? 1);
  const guides: SnapGuide[] = [];
  const spacing: SpacingGuide[] = [];
  let { left, top, width, height } = box;
  const axisSnap = (axis: "x" | "y", fromEnd: boolean, limit: number) => {
    if (!(limit > 0) || !Number.isFinite(limit)) return null;
    const start = (r: BlockGeometry) => (axis === "x" ? r.left : r.top);
    const size = (r: BlockGeometry) => (axis === "x" ? r.width : r.height);
    const end = (r: BlockGeometry) => start(r) + size(r);
    const crossStart = (r: BlockGeometry) => (axis === "x" ? r.top : r.left);
    const crossSize = (r: BlockGeometry) => (axis === "x" ? r.height : r.width);
    const current: BlockGeometry = { left, top, width, height };
    const centered = options.centered === true;
    // What stays: the far edge, or the centre for a resize from the centre.
    const pinned = centered
      ? start(current) + size(current) / 2
      : fromEnd ? start(current) : end(current);
    // A span from the pinned point is the whole size, or half of it.
    const spanFactor = centered ? 2 : 1;
    const edge = fromEnd ? end(current) : start(current);
    const bound = axis === "x" ? slideBounds?.width : slideBounds?.height;
    const extent = Number.isFinite(bound) ? (bound as number) : 0;
    type Candidate = { edge: number; line?: SnapGuide; guide?: SpacingGuide };
    const candidates: Candidate[] = [];
    const orientation = axis === "x" ? "vertical" : "horizontal";
    const lines: Array<{ coord: number; kind: SnapGuide["kind"] }> = slideBounds
      ? [
          { coord: 0, kind: "slide-edge" },
          { coord: extent / 2, kind: "slide-center" },
          { coord: extent, kind: "slide-edge" },
        ]
      : [];
    for (const rect of rects) {
      lines.push(
        { coord: start(rect), kind: "block-edge" },
        { coord: start(rect) + size(rect) / 2, kind: "block-center" },
        { coord: end(rect), kind: "block-edge" },
      );
    }
    for (const line of lines) {
      candidates.push({
        edge: line.coord,
        line: { orientation, position: safeRound(line.coord), kind: line.kind },
      });
    }
    // The arrows sit just outside each box, on its far side across the axis.
    const offset = (axis === "x" ? thresholdY : thresholdX) || limit;
    const arrowAt = (r: BlockGeometry) => crossStart(r) + crossSize(r) + offset / 2;
    if (options.size) {
      for (const rect of rects) {
        const target = size(rect);
        if (!(target >= minSize)) continue;
        const nextEdge = fromEnd ? pinned + target / spanFactor : pinned - target / spanFactor;
        const matches = rects.filter((other) => Math.abs(size(other) - target) < 1);
        const resized = centered
          ? { from: pinned - target / 2, to: pinned + target / 2 }
          : fromEnd ? { from: pinned, to: nextEdge } : { from: nextEdge, to: pinned };
        candidates.push({
          edge: nextEdge,
          guide: {
            axis,
            kind: "size",
            segments: [
              { ...resized, at: arrowAt(current) },
              ...matches.map((other) => ({ from: start(other), to: end(other), at: arrowAt(other) })),
            ],
          },
        });
      }
    }
    if (options.spacing) {
      const crossEnd = (r: BlockGeometry) => crossStart(r) + crossSize(r);
      const band = rects.filter(
        (r) => crossStart(r) < crossEnd(current) && crossEnd(r) > crossStart(current),
      );
      const at = crossStart(current) + crossSize(current) / 2;
      // The neighbour the dragged edge faces, and the one behind the pinned edge.
      const ahead = fromEnd
        ? band.filter((r) => start(r) >= pinned + minSize).sort((a, b) => start(a) - start(b))[0]
        : band.filter((r) => end(r) <= pinned - minSize).sort((a, b) => end(b) - end(a))[0];
      // From the centre the far edge moves too: no gap of its own to repeat.
      const behind = centered
        ? undefined
        : fromEnd
          ? band.filter((r) => end(r) <= pinned).sort((a, b) => end(b) - end(a))[0]
          : band.filter((r) => start(r) >= pinned).sort((a, b) => start(a) - start(b))[0];
      if (ahead) {
        const sorted = [...band].sort((a, b) => start(a) - start(b));
        const gaps: Array<{ gap: number; from: number; to: number }> = [];
        sorted.slice(1).forEach((next, index) => {
          const gap = start(next) - end(sorted[index]);
          if (gap > 0) gaps.push({ gap, from: end(sorted[index]), to: start(next) });
        });
        if (behind) {
          const gap = fromEnd ? pinned - end(behind) : start(behind) - pinned;
          if (gap > 0) {
            gaps.push(fromEnd ? { gap, from: end(behind), to: pinned } : { gap, from: pinned, to: start(behind) });
          }
        }
        for (const reference of gaps) {
          const nextEdge = fromEnd ? start(ahead) - reference.gap : end(ahead) + reference.gap;
          const own = fromEnd ? { from: nextEdge, to: start(ahead) } : { from: end(ahead), to: nextEdge };
          candidates.push({
            edge: nextEdge,
            guide: {
              axis,
              segments: [
                { from: reference.from, to: reference.to, at },
                { ...own, at },
              ],
            },
          });
        }
      }
    }
    let best: (Candidate & { distance: number }) | null = null;
    for (const candidate of candidates) {
      const distance = Math.abs(candidate.edge - edge);
      if (distance > limit) continue;
      const span = fromEnd ? candidate.edge - pinned : pinned - candidate.edge;
      if (span * spanFactor < minSize) continue;
      if (best && distance >= best.distance) continue;
      best = { ...candidate, distance };
    }
    if (!best) return null;
    const snappedEdge = safeRound(best.edge);
    const span = fromEnd ? snappedEdge - pinned : pinned - snappedEdge;
    const nextSize = span * spanFactor;
    const nextStart = centered ? pinned - span : fromEnd ? pinned : snappedEdge;
    if (axis === "x") {
      left = safeRound(nextStart);
      width = safeRound(nextSize);
    } else {
      top = safeRound(nextStart);
      height = safeRound(nextSize);
    }
    if (best.line) guides.push(best.line);
    if (best.guide) {
      spacing.push({
        ...best.guide,
        segments: best.guide.segments.map((segment) => ({
          from: safeRound(segment.from),
          to: safeRound(segment.to),
          at: safeRound(segment.at),
        })),
      });
    }
    return true;
  };
  const snappedX = handle.includes("e")
    ? Boolean(axisSnap("x", true, thresholdX))
    : handle.includes("w")
      ? Boolean(axisSnap("x", false, thresholdX))
      : false;
  const snappedY = handle.includes("s")
    ? Boolean(axisSnap("y", true, thresholdY))
    : handle.includes("n")
      ? Boolean(axisSnap("y", false, thresholdY))
      : false;
  return { geom: { left, top, width, height }, guides, spacing, snappedX, snappedY };
}

/**
 * Object and grid snapping for any resize, worked through the pointer: the
 * caller's `evaluate` answers the box the user SEES for a pointer offset (a
 * rotated object's turned box, a multi-selection's union, a picture's
 * crop-aware frame), and this finds the pointer offset that puts the box's
 * dragged edge where {@link snapResizedEdges} (or, failing that, the grid)
 * wants it. So every resize -- pictures, rotated objects, several objects
 * scaled at once, a resize from the centre, a locked ratio -- snaps by the
 * same rules as a plain one, and the caller's own resize math (minimum
 * size, ratio, mirroring) still decides the result.
 *
 * Per axis the dragged edge is the one of the seen box that follows the
 * pointer most (both, from the centre: the handle's side). An axis the
 * pointer does not drive is left alone; with a locked ratio only the first
 * axis snaps (the other follows it). A target the resize cannot reach (it
 * would pass the minimum size) leaves the axis as the pointer put it.
 */
export function snapResizePointer(request: {
  readonly evaluate: (dx: number, dy: number) => BlockGeometry | null;
  readonly dx: number;
  readonly dy: number;
  /** The handle in the box's own frame, for a tie between two edges. */
  readonly handle: string;
  readonly others: ReadonlyArray<Partial<BlockGeometry> | null | undefined>;
  readonly slideBounds: SlideBounds | null;
  readonly threshold: number | SnapThresholds;
  /** Snap objects to other objects, or Smart Guides: edges and centres. */
  readonly objects: boolean;
  /** Smart Guides: also equal size and equal spacing (and guides drawn). */
  readonly smartGuides: boolean;
  readonly centered?: boolean;
  readonly lockAspect?: boolean;
  /** Snap objects to grid's spacing, EMU; 0 / absent for no grid. */
  readonly gridSpacing?: number;
}): { dx: number; dy: number; guides: SnapGuide[]; spacing: SpacingGuide[]; snappedX: boolean; snappedY: boolean } {
  let { dx, dy } = request;
  const guides: SnapGuide[] = [];
  const spacing: SpacingGuide[] = [];
  const snapped = { x: false, y: false };
  const { x: thresholdX, y: thresholdY } = resolveThresholds(request.threshold);
  for (const axis of ["x", "y"] as const) {
    if (request.lockAspect && (snapped.x || snapped.y)) break;
    const limit = axis === "x" ? thresholdX : thresholdY;
    const step = Math.max(1, Math.abs(limit) / 4 || 100);
    const box = request.evaluate(dx, dy);
    const nudged = axis === "x" ? request.evaluate(dx + step, dy) : request.evaluate(dx, dy + step);
    if (!box || !nudged) continue;
    const startOf = (r: BlockGeometry) => (axis === "x" ? r.left : r.top);
    const endOf = (r: BlockGeometry) => startOf(r) + (axis === "x" ? r.width : r.height);
    const slopeStart = (startOf(nudged) - startOf(box)) / step;
    const slopeEnd = (endOf(nudged) - endOf(box)) / step;
    if (Math.max(Math.abs(slopeStart), Math.abs(slopeEnd)) < 0.2) continue;
    const fromEnd =
      Math.abs(slopeEnd) > Math.abs(slopeStart) + 0.05
        ? true
        : Math.abs(slopeStart) > Math.abs(slopeEnd) + 0.05
          ? false
          : !request.handle.includes(axis === "x" ? "w" : "n");
    const slope = fromEnd ? slopeEnd : slopeStart;
    const edgeOf = fromEnd ? endOf : startOf;
    const edge = edgeOf(box);
    let target: number | null = null;
    let found: ReturnType<typeof snapResizedEdges> | null = null;
    if (request.objects) {
      const side = axis === "x" ? (fromEnd ? "e" : "w") : fromEnd ? "s" : "n";
      const result = snapResizedEdges(
        box,
        side,
        request.others,
        request.slideBounds,
        { x: thresholdX, y: thresholdY },
        {
          spacing: request.smartGuides,
          size: request.smartGuides,
          minSize: 1,
          centered: request.centered,
        },
      );
      if (axis === "x" ? result.snappedX : result.snappedY) {
        target = edgeOf(result.geom);
        found = result;
      }
    }
    if (target === null && request.gridSpacing && request.gridSpacing > 0) {
      target = Math.round(Math.round(edge / request.gridSpacing) * request.gridSpacing);
    }
    if (target === null || target === edge) {
      if (target === edge) snapped[axis] = true;
      if (found) {
        guides.push(...found.guides);
        spacing.push(...found.spacing);
      }
      continue;
    }
    // The resize rounds to whole EMU and the measured slope carries that
    // rounding: step again from where the edge landed until it is exactly
    // on the target, so an equal size or gap is equal to the EMU, as
    // PowerPoint makes it (a single step left a width one EMU short).
    let nextDx = dx;
    let nextDy = dy;
    let check: BlockGeometry | null = box;
    for (let pass = 0; pass < 4 && check && edgeOf(check) !== target; pass += 1) {
      const delta = (target - edgeOf(check)) / slope;
      if (axis === "x") nextDx += delta;
      else nextDy += delta;
      check = request.evaluate(nextDx, nextDy);
    }
    // The resize's own limits (minimum size, the pasteboard) may stop short:
    // a target it cannot reach is not taken.
    if (!check || Math.abs(edgeOf(check) - target) > Math.max(2, Math.abs(target - edge) / 2)) continue;
    dx = nextDx;
    dy = nextDy;
    snapped[axis] = true;
    if (found) {
      guides.push(...found.guides);
      spacing.push(...found.spacing);
    }
  }
  return { dx, dy, guides, spacing, snappedX: snapped.x, snappedY: snapped.y };
}
