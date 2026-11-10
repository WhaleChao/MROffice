/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry math for direct manipulation (drag-to-move + handle-resize) of
 * presentation blocks on the editor canvas.
 *
 * Dependency-free (no React / no DOM) so it can be unit tested in isolation and
 * reused by both the editor component and tests. All values are in the same EMU
 * coordinate space the rest of the presentation stack uses (`left`/`top`/
 * `width`/`height` on a block); the caller converts pointer pixel deltas into
 * EMU deltas before calling these helpers, and commits the returned geometry
 * through the existing block-geometry update path so undo/redo + autosave work.
 *
 * Invariants every helper upholds:
 *  - The block never leaves the slide bounds (clamped on every axis).
 *  - Width / height never go below `minSize` and are never negative or zero.
 *  - Integer EMU out (rounded) so persisted geometry stays clean.
 *
 * Mirrors the conventions of the sibling pure modules
 * (`presentation-shapes.ts`, `presentation-block-fx.ts`): frozen allowed lists,
 * a type guard / normalizer, and small composable functions.
 */

/** A block's positioned rectangle, in EMU. Matches `ShapeGeometry`. */
export interface BlockGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * The area a gesture keeps an object in, in EMU: by default the slide
 * (`0..width`, `0..height`); with `originX`/`originY` an area that starts
 * elsewhere -- the pasteboard around the slide ({@link presentationPasteboard}).
 */
export interface DragBounds {
  width: number;
  height: number;
  /** Where the area starts on x (default 0, the slide's left edge). */
  originX?: number;
  /** Where the area starts on y (default 0, the slide's top edge). */
  originY?: number;
}

/**
 * The pasteboard: the grey area around the slide where PowerPoint lets an
 * object be dragged, nudged or placed partly or wholly off the slide (it is
 * drawn there in Normal view and cut off only in the show and the
 * thumbnails). MRAI's reaches one slide's width and height past each edge,
 * so an object put there can always be seen and taken back.
 */
export function presentationPasteboard(slide: DragBounds): DragBounds {
  const width = Number.isFinite(slide?.width) ? slide.width : 0;
  const height = Number.isFinite(slide?.height) ? slide.height : 0;
  return { originX: -width, originY: -height, width: width * 3, height: height * 3 };
}

/** The area's origin, and the area moved to start at 0. */
function localFrame(bounds: DragBounds): { x: number; y: number; local: DragBounds } {
  const x = Number.isFinite(bounds?.originX) ? Math.round(bounds.originX as number) : 0;
  const y = Number.isFinite(bounds?.originY) ? Math.round(bounds.originY as number) : 0;
  return { x, y, local: { width: bounds?.width, height: bounds?.height } };
}

function shifted<T extends Partial<BlockGeometry> | null | undefined>(geom: T, dx: number, dy: number): T {
  if (!geom || (!dx && !dy)) return geom;
  return {
    ...geom,
    ...(Number.isFinite(geom.left) ? { left: (geom.left as number) + dx } : {}),
    ...(Number.isFinite(geom.top) ? { top: (geom.top as number) + dy } : {}),
  } as T;
}

/** Run a slide-framed gesture function in an area that starts elsewhere. */
function inFrame(
  bounds: DragBounds,
  geom: Partial<BlockGeometry> | null | undefined,
  run: (local: Partial<BlockGeometry> | null | undefined, area: DragBounds) => BlockGeometry,
): BlockGeometry {
  const frame = localFrame(bounds);
  if (!frame.x && !frame.y) return run(geom, bounds);
  return shifted(run(shifted(geom, -frame.x, -frame.y), frame.local), frame.x, frame.y);
}

/**
 * The eight resize handles, named by the edge(s)/corner they grab (compass
 * directions). Edge handles move one side; corner handles move two. `n`/`w`
 * handles move the top/left origin as the size changes; `s`/`e` only change the
 * size.
 */
export const RESIZE_HANDLES = [
  "n",
  "s",
  "e",
  "w",
  "ne",
  "nw",
  "se",
  "sw",
] as const;

export type ResizeHandle = (typeof RESIZE_HANDLES)[number];

const RESIZE_HANDLE_SET: ReadonlySet<string> = new Set(RESIZE_HANDLES);

/** Default minimum block extent (EMU) — keeps a resized block grabbable. */
export const DEFAULT_MIN_BLOCK_SIZE = 90_000;

/** Type guard: is `value` exactly one of the eight handle identifiers? */
export function isResizeHandle(value: unknown): value is ResizeHandle {
  return typeof value === "string" && RESIZE_HANDLE_SET.has(value);
}

/** Clamp a number into `[min, max]`. Returns `min` when `min > max`. */
function clamp(value: number, min: number, max: number): number {
  if (max < min) return min;
  return Math.min(max, Math.max(min, value));
}

/** Round to an integer, guarding against NaN/Infinity (→ 0). */
function safeRound(value: number): number {
  return Number.isFinite(value) ? Math.round(value) : 0;
}

/**
 * Normalize a raw geometry-ish value into a finite, on-positive rect. Missing /
 * non-finite fields collapse to 0 (left/top) or 1 (width/height) so callers can
 * pass a partially-populated block without exploding. Width/height are forced
 * to at least 1 EMU here; {@link applyResize} enforces the real `minSize`.
 */
export function normalizeGeometry(
  geom: Partial<BlockGeometry> | null | undefined,
): BlockGeometry {
  const left = Number.isFinite(geom?.left) ? safeRound(geom!.left!) : 0;
  const top = Number.isFinite(geom?.top) ? safeRound(geom!.top!) : 0;
  const width = Number.isFinite(geom?.width) ? Math.max(1, safeRound(geom!.width!)) : 1;
  const height = Number.isFinite(geom?.height)
    ? Math.max(1, safeRound(geom!.height!))
    : 1;
  return { left, top, width, height };
}

/**
 * Move a block by `(dx, dy)` EMU, clamping its origin so the whole rectangle
 * stays inside `bounds`. The size is preserved; only `left`/`top` change. If the
 * block is wider/taller than the slide (degenerate imported geometry) the clamp
 * still keeps `left`/`top` non-negative (max bound collapses to 0).
 */
export function applyMove(
  geom: Partial<BlockGeometry> | null | undefined,
  dx: number,
  dy: number,
  bounds: DragBounds,
): BlockGeometry {
  return inFrame(bounds, geom, (local, area) => applyMoveInSlide(local, dx, dy, area));
}

function applyMoveInSlide(
  geom: Partial<BlockGeometry> | null | undefined,
  dx: number,
  dy: number,
  bounds: DragBounds,
): BlockGeometry {
  const base = normalizeGeometry(geom);
  // An axis the move does not touch keeps its value, and an axis that already
  // sits past the slide's edge (a native deck may lay a shape slightly off the
  // page) is not pulled back by a move along the other one; the clamp only
  // stops a move from taking the block further out (PPT-B387).
  const along = (value: number, delta: number, max: number): number => {
    if (delta === 0) return value;
    // An oversized block (max < 0) keeps a non-negative origin, as before.
    if (max < 0) return 0;
    return clamp(safeRound(value + delta), Math.min(0, value), Math.max(max, value));
  };
  const left = along(base.left, dx, bounds.width - base.width);
  const top = along(base.top, dy, bounds.height - base.height);
  return { ...base, left, top };
}

/**
 * The line a Shift-drag keeps an object on (PowerPoint, Word, Excel: "hold
 * down Shift as you drag" moves it only horizontally or vertically): the axis
 * the pointer has gone further along, in screen px so a non-uniform slide
 * scale does not tip it. Null without Shift, or before the pointer moved.
 */
export function constrainedMoveAxis(
  pointerDx: number,
  pointerDy: number,
  shift: boolean,
): "horizontal" | "vertical" | null {
  if (!shift) return null;
  const x = Math.abs(Number.isFinite(pointerDx) ? pointerDx : 0);
  const y = Math.abs(Number.isFinite(pointerDy) ? pointerDy : 0);
  if (x === 0 && y === 0) return null;
  return x >= y ? "horizontal" : "vertical";
}

/** The modifier keys a drag reads (a pointer or mouse event). */
export interface DragModifierKeys {
  readonly altKey?: boolean;
  readonly ctrlKey?: boolean;
  readonly metaKey?: boolean;
}

/**
 * Whether a move drag dropped now leaves the original where it was and puts
 * a COPY at the drop (Microsoft Support, "Keyboard shortcuts in PowerPoint":
 * Windows "Ctrl+drag" copies the selected object; PowerPoint for Mac
 * "Option+drag"). Read at the drop, as PowerPoint does: the key may be
 * pressed or released during the drag.
 */
export function presentationDragCopies(event: DragModifierKeys, mac: boolean): boolean {
  return mac ? Boolean(event.altKey) : Boolean(event.ctrlKey);
}

/**
 * Whether the drag ignores the snapping for now (Microsoft Support, "Work
 * with gridlines and use snap-to-grid in PowerPoint": Windows "hold down Alt
 * while you're dragging"; on a Mac "hold down Command while dragging" for
 * fine control -- Option is the copy key there).
 */
export function presentationDragSnapOverride(event: DragModifierKeys, mac: boolean): boolean {
  return mac ? Boolean(event.metaKey) : Boolean(event.altKey);
}

/**
 * Whether a handle drag resizes about the object's centre (Microsoft
 * Support, "Change the size of a picture, shape, text box, or WordArt":
 * Windows "press and hold Ctrl while you drag the sizing handle"; PowerPoint
 * for Mac "press and hold OPTION"). On Windows Alt is the snap override
 * only, never the centre key.
 */
export function presentationResizeFromCenter(event: DragModifierKeys, mac: boolean): boolean {
  return mac ? Boolean(event.altKey) : Boolean(event.ctrlKey);
}

/**
 * The one `(dx, dy)` every rect of a set that moves as ONE object takes (a
 * group drag or nudge): the requested delta, clamped so the set's UNION stays
 * inside `bounds`. Clamping each rect on its own would let the edge members
 * lag behind and deform the group. Integer EMU out.
 */
export function clampGroupMoveDelta(
  rects: readonly BlockGeometry[],
  dx: number,
  dy: number,
  bounds: DragBounds,
): { dx: number; dy: number } {
  if (rects.length === 0) return { dx: 0, dy: 0 };
  const { x, y } = localFrame(bounds);
  const unionLeft = Math.min(...rects.map((rect) => rect.left)) - x;
  const unionTop = Math.min(...rects.map((rect) => rect.top)) - y;
  const unionRight = Math.max(...rects.map((rect) => rect.left + rect.width)) - x;
  const unionBottom = Math.max(...rects.map((rect) => rect.top + rect.height)) - y;
  return {
    dx: safeRound(Math.min(Math.max(dx, -unionLeft), bounds.width - unionRight)),
    dy: safeRound(Math.min(Math.max(dy, -unionTop), bounds.height - unionBottom)),
  };
}

/**
 * Resize a block by dragging `handle` by `(dx, dy)` EMU. Edge handles change one
 * dimension; corner handles change both. North/west handles move the origin
 * (`top`/`left`) inward as the opposite edge stays pinned, so the far edge does
 * not drift. Every result is clamped so:
 *  - the moving edge never crosses the pinned edge (width/height ≥ `minSize`),
 *  - the rectangle never extends past the slide bounds,
 *  - width/height are never negative or zero (floored at `minSize`, itself ≥ 1).
 *
 * `minSize` defaults to {@link DEFAULT_MIN_BLOCK_SIZE}; callers may pass a
 * smaller floor but it is always coerced to ≥ 1 EMU.
 */
export function applyResize(
  geom: Partial<BlockGeometry> | null | undefined,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  bounds: DragBounds,
  minSize: number = DEFAULT_MIN_BLOCK_SIZE,
): BlockGeometry {
  return inFrame(bounds, geom, (local, area) =>
    applyResizeInSlide(local, handle, dx, dy, area, minSize),
  );
}

function applyResizeInSlide(
  geom: Partial<BlockGeometry> | null | undefined,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  bounds: DragBounds,
  minSize: number,
): BlockGeometry {
  const base = normalizeGeometry(geom);
  const min = Math.max(1, safeRound(minSize));
  // Pinned (far) edges of the original rect.
  const right = base.left + base.width;
  const bottom = base.top + base.height;

  let { left, top, width, height } = base;

  const movesEast = handle === "e" || handle === "ne" || handle === "se";
  const movesWest = handle === "w" || handle === "nw" || handle === "sw";
  const movesSouth = handle === "s" || handle === "se" || handle === "sw";
  const movesNorth = handle === "n" || handle === "ne" || handle === "nw";

  if (movesEast) {
    // Right edge moves; left pinned. Cap so it never passes the slide right.
    const maxWidth = bounds.width - base.left;
    width = clamp(safeRound(base.width + dx), min, Math.max(min, maxWidth));
  } else if (movesWest) {
    // Left edge moves; right pinned. New left is between 0 and right - min.
    const maxLeft = right - min;
    left = clamp(safeRound(base.left + dx), 0, Math.max(0, maxLeft));
    width = right - left;
  }

  if (movesSouth) {
    const maxHeight = bounds.height - base.top;
    height = clamp(safeRound(base.height + dy), min, Math.max(min, maxHeight));
  } else if (movesNorth) {
    const maxTop = bottom - min;
    top = clamp(safeRound(base.top + dy), 0, Math.max(0, maxTop));
    height = bottom - top;
  }

  return { left, top, width, height };
}

/**
 * Optional modifier keys that change how a resize-handle drag is interpreted,
 * matching PowerPoint / Keynote:
 *  - `lockAspect` (Shift held): preserve the block's original width:height ratio.
 *    For a corner handle the axis with the larger proportional drag drives the
 *    size and the other axis is derived from the locked ratio; for an edge handle
 *    the dragged axis drives and the cross axis grows/shrinks proportionally,
 *    centred on the block's cross-axis midline.
 *  - `fromCenter` (Option held on a Mac, Ctrl on Windows): resize symmetrically about the block CENTER —
 *    the edge opposite the grabbed handle moves by the same amount in the other
 *    direction, so the centre stays fixed and each axis changes by ~2× the drag.
 */
export interface ResizeModifiers {
  lockAspect?: boolean;
  fromCenter?: boolean;
}

/**
 * Resize a block by dragging `handle` by `(dx, dy)` EMU, honouring the Shift
 * (lock-aspect) / Alt (from-centre) modifiers. With NEITHER modifier set this is
 * exactly {@link applyResize} (delegated, byte-for-byte identical) so the default
 * drag path is unchanged; the extra geometry only runs when a modifier is active.
 *
 * The result upholds the same invariants as {@link applyResize}: width/height are
 * never below `minSize` (itself coerced to ≥ 1 EMU) and never negative/zero, the
 * rectangle never extends past `bounds`, and every field is finite integer EMU.
 * The aspect ratio and centre constraints are best-effort under those hard
 * clamps — when honouring a modifier would push the rect out of bounds or below
 * the minimum the clamp wins (PowerPoint behaves the same), so a returned rect is
 * always valid even if it can no longer hold the exact ratio.
 *
 * Pure + deterministic: no DOM, no Date/random; inputs are never mutated.
 */
export function applyResizeWithModifiers(
  geom: Partial<BlockGeometry> | null | undefined,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  bounds: DragBounds,
  modifiers?: ResizeModifiers | null,
  minSize: number = DEFAULT_MIN_BLOCK_SIZE,
): BlockGeometry {
  return inFrame(bounds, geom, (local, area) =>
    applyResizeWithModifiersInSlide(local, handle, dx, dy, area, modifiers, minSize),
  );
}

function applyResizeWithModifiersInSlide(
  geom: Partial<BlockGeometry> | null | undefined,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  bounds: DragBounds,
  modifiers: ResizeModifiers | null | undefined,
  minSize: number,
): BlockGeometry {
  const lockAspect = modifiers?.lockAspect === true;
  const fromCenter = modifiers?.fromCenter === true;
  // No modifier → identical to the plain resize (keeps the default path exact).
  if (!lockAspect && !fromCenter) {
    return applyResizeInSlide(geom, handle, dx, dy, bounds, minSize);
  }

  const base = normalizeGeometry(geom);
  const min = Math.max(1, safeRound(minSize));
  const boundsW = Number.isFinite(bounds?.width) ? bounds.width : 0;
  const boundsH = Number.isFinite(bounds?.height) ? bounds.height : 0;
  const ddx = Number.isFinite(dx) ? dx : 0;
  const ddy = Number.isFinite(dy) ? dy : 0;

  const movesEast = handle === "e" || handle === "ne" || handle === "se";
  const movesWest = handle === "w" || handle === "nw" || handle === "sw";
  const movesSouth = handle === "s" || handle === "se" || handle === "sw";
  const movesNorth = handle === "n" || handle === "ne" || handle === "nw";
  const controlsX = movesEast || movesWest;
  const controlsY = movesSouth || movesNorth;

  // Per-axis width/height change implied by the drag. Dragging the east/south
  // edge outward grows by +d; dragging the west/north edge outward (toward
  // smaller coords) grows by -d. `fromCenter` mirrors the opposite edge so the
  // dimension changes by twice the drag.
  const centerFactor = fromCenter ? 2 : 1;
  let targetW = controlsX
    ? base.width + centerFactor * (movesEast ? ddx : -ddx)
    : base.width;
  let targetH = controlsY
    ? base.height + centerFactor * (movesSouth ? ddy : -ddy)
    : base.height;

  if (lockAspect && base.width > 0 && base.height > 0) {
    const aspect = base.width / base.height; // w per h
    if (controlsX && controlsY) {
      // Corner: the axis dragged proportionally further drives the size.
      const relX = Math.abs(targetW - base.width) / base.width;
      const relY = Math.abs(targetH - base.height) / base.height;
      if (relX >= relY) {
        targetH = targetW / aspect;
      } else {
        targetW = targetH * aspect;
      }
    } else if (controlsX) {
      targetH = targetW / aspect; // edge on x drives, derive height
    } else if (controlsY) {
      targetW = targetH * aspect; // edge on y drives, derive width
    }
  }

  // Floor at the minimum (and never below 1). Aspect can be momentarily broken
  // by the floor; the bounds clamp below may break it further — clamps win.
  targetW = Math.max(min, targetW);
  targetH = Math.max(min, targetH);

  // Anchor each axis: with `fromCenter` the original centre is fixed; otherwise
  // the edge opposite the grabbed handle stays pinned. A pure cross-axis grow
  // (aspect-lock on an edge handle) expands symmetrically about the centre.
  const centerX = base.left + base.width / 2;
  const centerY = base.top + base.height / 2;
  const right = base.left + base.width;
  const bottom = base.top + base.height;

  const placeLeft = (w: number): number => {
    if (fromCenter) return centerX - w / 2;
    if (movesEast) return base.left; // left pinned
    if (movesWest) return right - w; // right pinned
    return centerX - w / 2; // cross-axis grow → centre
  };
  const placeTop = (h: number): number => {
    if (fromCenter) return centerY - h / 2;
    if (movesSouth) return base.top; // top pinned
    if (movesNorth) return bottom - h; // bottom pinned
    return centerY - h / 2;
  };

  // Clamp the size DOWN so the placed rect fits the slide, then re-derive the
  // origin from the (possibly reduced) size so the anchor still holds.
  const fitWidth = (w: number): number => {
    let out = w;
    if (fromCenter) {
      // Symmetric: limited by the nearer side's room ×2.
      const room = 2 * Math.min(centerX, boundsW - centerX);
      out = Math.min(out, Math.max(min, room));
    } else if (movesEast) {
      out = Math.min(out, Math.max(min, boundsW - base.left));
    } else if (movesWest) {
      out = Math.min(out, Math.max(min, right));
    } else {
      const room = 2 * Math.min(centerX, boundsW - centerX);
      out = Math.min(out, Math.max(min, room));
    }
    return Math.max(min, out);
  };
  const fitHeight = (h: number): number => {
    let out = h;
    if (fromCenter) {
      const room = 2 * Math.min(centerY, boundsH - centerY);
      out = Math.min(out, Math.max(min, room));
    } else if (movesSouth) {
      out = Math.min(out, Math.max(min, boundsH - base.top));
    } else if (movesNorth) {
      out = Math.min(out, Math.max(min, bottom));
    } else {
      const room = 2 * Math.min(centerY, boundsH - centerY);
      out = Math.min(out, Math.max(min, room));
    }
    return Math.max(min, out);
  };

  // When aspect is locked, clamping ONE axis to fit bounds must shrink the other
  // by the same ratio so the rect stays on-ratio after the fit (a single pass is
  // enough: re-fitting the now-smaller second axis can only relax its own limit).
  let finalW = fitWidth(targetW);
  let finalH = fitHeight(targetH);
  if (lockAspect && base.width > 0 && base.height > 0) {
    const aspect = base.width / base.height;
    const wScale = finalW / targetW;
    const hScale = finalH / targetH;
    if (wScale <= hScale) {
      // Width was clamped harder → match height to the clamped width.
      finalH = Math.max(min, finalW / aspect);
    } else {
      finalW = Math.max(min, finalH * aspect);
    }
  }

  let left = placeLeft(finalW);
  let top = placeTop(finalH);
  // Final safety clamp of the origin into the slide (covers rounding and the
  // degenerate oversized-block case the same way applyMove/applyResize do).
  left = clamp(left, 0, Math.max(0, boundsW - finalW));
  top = clamp(top, 0, Math.max(0, boundsH - finalH));

  return {
    left: safeRound(left),
    top: safeRound(top),
    width: safeRound(finalW),
    height: safeRound(finalH),
  };
}

/**
 * The handle a pointer holds on a mirrored block, as the block's own box
 * names it. The handles are drawn inside the block's mirrored frame, so on a
 * block with `flipH` the "e" grip sits on the box's LEFT edge: dragging it
 * moves that left edge (the right one pinned), which is the box's "w" resize.
 */
export function mirroredResizeHandle(
  handle: ResizeHandle,
  flip: { flipH?: boolean; flipV?: boolean } | null | undefined,
): ResizeHandle {
  if (!flip?.flipH && !flip?.flipV) return handle;
  const y = handle.includes("n") ? (flip.flipV ? "s" : "n") : handle.includes("s") ? (flip.flipV ? "n" : "s") : "";
  const x = handle.includes("e") ? (flip.flipH ? "w" : "e") : handle.includes("w") ? (flip.flipH ? "e" : "w") : "";
  return `${y}${x}` as ResizeHandle;
}

/** A resize that may have carried a handle through the pinned edge. */
export interface ResizeThroughZero {
  geom: BlockGeometry;
  /** The drag crossed the pinned vertical edge: the block is mirrored on x. */
  flipX: boolean;
  /** The drag crossed the pinned horizontal edge: the block is mirrored on y. */
  flipY: boolean;
}

/**
 * Resize by dragging `handle`, and keep going through the pinned edge:
 * PowerPoint mirrors the shape when a handle is dragged past the opposite
 * side (its text with it) in the same gesture, instead of stopping at the
 * minimum size. On an axis the drag has not crossed this is exactly
 * {@link applyResizeWithModifiers}; on an axis it has, the result is the rect
 * the handle now spans on the far side of the pinned edge (or, from the
 * centre, through it), and the matching flip is reported for the caller to
 * compose with the block's own flags (`flipX` toggles `flipH`). The handle
 * stays the handle the pointer holds, so dragging back un-mirrors.
 *
 * Pure: the crossed axis is solved as the mirrored problem -- the opposite
 * handle on the rect reflected across the pinned edge, dragged to the same
 * place -- so the floor, bounds and modifier rules are the shared ones.
 */
export function applyResizeThroughZero(
  geom: Partial<BlockGeometry> | null | undefined,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  bounds: DragBounds,
  modifiers?: ResizeModifiers | null,
  minSize: number = DEFAULT_MIN_BLOCK_SIZE,
): ResizeThroughZero {
  const base = normalizeGeometry(geom);
  const fromCenter = modifiers?.fromCenter === true;
  const factor = fromCenter ? 2 : 1;
  const ddx = Number.isFinite(dx) ? dx : 0;
  const ddy = Number.isFinite(dy) ? dy : 0;
  const movesEast = handle === "e" || handle === "ne" || handle === "se";
  const movesWest = handle === "w" || handle === "nw" || handle === "sw";
  const movesSouth = handle === "s" || handle === "se" || handle === "sw";
  const movesNorth = handle === "n" || handle === "ne" || handle === "nw";
  // The signed extent the drag asks for on each axis it controls.
  const signedW = movesEast ? base.width + factor * ddx : movesWest ? base.width - factor * ddx : base.width;
  const signedH = movesSouth ? base.height + factor * ddy : movesNorth ? base.height - factor * ddy : base.height;
  const flipX = signedW < 0;
  const flipY = signedH < 0;
  if (!flipX && !flipY) {
    return { geom: applyResizeWithModifiers(base, handle, ddx, ddy, bounds, modifiers, minSize), flipX, flipY };
  }
  // The mirrored problem: on a crossed axis the opposite handle, on the rect
  // reflected across its pinned edge (from the centre the rect is its own
  // reflection), dragged so its extent is |signed|: d' = d ∓ 2w / factor.
  let { left, top } = base;
  let x: "e" | "w" | "" = movesEast ? "e" : movesWest ? "w" : "";
  let y: "s" | "n" | "" = movesSouth ? "s" : movesNorth ? "n" : "";
  let mdx = ddx;
  let mdy = ddy;
  if (flipX) {
    if (!fromCenter) left = movesEast ? base.left - base.width : base.left + base.width;
    mdx = movesEast ? ddx + (2 * base.width) / factor : ddx - (2 * base.width) / factor;
    x = movesEast ? "w" : "e";
  }
  if (flipY) {
    if (!fromCenter) top = movesSouth ? base.top - base.height : base.top + base.height;
    mdy = movesSouth ? ddy + (2 * base.height) / factor : ddy - (2 * base.height) / factor;
    y = movesSouth ? "n" : "s";
  }
  const mirroredHandle = `${y}${x}` as ResizeHandle;
  return {
    geom: applyResizeWithModifiers({ ...base, left, top }, mirroredHandle, mdx, mdy, bounds, modifiers, minSize),
    flipX,
    flipY,
  };
}

/* ------------------------------------------------------------------------- *
 * Rotation + flip transform math (PowerPoint parity).
 *
 * A block's render transform composes its degree `rotation` with the two
 * mirror flags `flipH` (horizontal) / `flipV` (vertical). PowerPoint stores the
 * mirror on the shape's `a:xfrm` and applies the flip BEFORE the rotation: the
 * shape is first reflected in its own local frame, then the whole reflected box
 * is rotated about its center.
 *
 * In a CSS `transform`, functions apply RIGHT-TO-LEFT (the right-most function
 * acts first, in the element's own coordinate space). So to get "flip then
 * rotate" we emit `rotate(R) scaleX(-1) scaleY(-1)` — `rotate` is written FIRST
 * but, because CSS reads right-to-left, the `scale` (flip) is applied first and
 * the rotation second. This is the documented, stable order this whole module
 * (and the tests) rely on:
 *
 *     transform: rotate(<deg>deg) scaleX(-1) scaleY(-1)
 *                └ written first ┘ └──── applied first (flip) ────┘
 *
 * `transform-origin: center center` (set by the caller) makes both the flip and
 * the rotation pivot on the block center, so flipping twice is an exact identity
 * and a flipped+rotated block lands where PowerPoint would put it.
 * ------------------------------------------------------------------------- */

/** A block's mirror flags. Absent / `false` both mean "not flipped". */
export interface BlockFlip {
  flipH?: boolean;
  flipV?: boolean;
}

/**
 * Build the CSS `transform` value that composes a block's `rotation` (degrees)
 * with its `flipH`/`flipV` mirror flags, in the documented PowerPoint order
 * (flip applied first, then rotate — see the block comment above).
 *
 * Returns `undefined` when there is nothing to transform (rotation 0 / absent
 * AND neither flip set) so a plain block serializes/renders byte-for-byte as
 * before (no stray `transform` in the style). Pure: no DOM, no Date/random.
 *
 * The factor order in the string is ALWAYS `rotate(...)` (when non-zero) then
 * `scaleX(-1)` (when flipH) then `scaleY(-1)` (when flipV); each part is omitted
 * when not needed, but the relative order is fixed so the result is
 * deterministic and the rotation always precedes the flips in the string.
 */
export function buildBlockTransform(
  rotation: number | null | undefined,
  flip?: BlockFlip | null,
): string | undefined {
  const deg = Number.isFinite(rotation) ? (rotation as number) : 0;
  const flipH = flip?.flipH === true;
  const flipV = flip?.flipV === true;
  const parts: string[] = [];
  if (deg !== 0) parts.push(`rotate(${deg}deg)`);
  if (flipH) parts.push("scaleX(-1)");
  if (flipV) parts.push("scaleY(-1)");
  return parts.length > 0 ? parts.join(" ") : undefined;
}

/** Toggle one mirror flag, returning the next boolean (absent/false → true). */
export function toggleFlip(current: boolean | undefined): boolean {
  return current !== true;
}

/**
 * One member of a GROUP flipped as a single object (PowerPoint: Flip
 * Horizontal / Vertical on a group mirrors the group, so its members trade
 * places). The member's box is mirrored across the group frame's centre line
 * on that axis and its rotation reflected (a mirror turns R(θ) into R(-θ));
 * the caller toggles the member's own flip flag, which completes the mirror
 * of its content. Pure.
 */
export function mirrorGroupMemberGeometry(
  rect: { left: number; top: number; width: number; height: number },
  rotation: number,
  frame: { left: number; top: number; right: number; bottom: number },
  axis: "flipH" | "flipV",
): { left: number; top: number; rotation: number } {
  return {
    left:
      axis === "flipH"
        ? frame.left + frame.right - (rect.left + rect.width)
        : rect.left,
    top:
      axis === "flipV"
        ? frame.top + frame.bottom - (rect.top + rect.height)
        : rect.top,
    rotation: normalizeRotationDegrees(-rotation),
  };
}

/** Default rotate-handle snap increment (degrees) when Shift is held. */
export const ROTATE_SNAP_DEGREES = 15;

/**
 * Snap a degree value to the nearest multiple of `increment` (default 15°),
 * matching PowerPoint's Shift-rotate behaviour. The result is normalized into
 * `[0, 360)`. Pure + deterministic.
 */
export function snapAngle(
  degrees: number,
  increment: number = ROTATE_SNAP_DEGREES,
): number {
  if (!Number.isFinite(degrees)) return 0;
  const step = increment > 0 ? increment : ROTATE_SNAP_DEGREES;
  const snapped = Math.round(degrees / step) * step;
  return normalizeRotationDegrees(snapped);
}

/** Normalize any degree value into the `[0, 360)` range. Pure. */
export function normalizeRotationDegrees(degrees: number): number {
  if (!Number.isFinite(degrees)) return 0;
  const mod = degrees % 360;
  // `+ 0` collapses a `-0` (e.g. `-360 % 360`) to `+0` so the value serializes
  // cleanly and equals exactly 0.
  return (mod < 0 ? mod + 360 : mod) + 0;
}

/** A 2-D point in canvas pixel space. */
export interface Point2D {
  x: number;
  y: number;
}

/**
 * Compute the rotation angle (degrees, `[0, 360)`) implied by dragging a rotate
 * handle to `pointer`, given the block's `center`. 0° points the handle
 * straight UP (the handle sits above the block, like PowerPoint), and the angle
 * increases CLOCKWISE so it matches the screen-space `rotate(<deg>deg)` the
 * render transform applies.
 *
 * Screen Y grows DOWNWARD, so a pointer directly above the center
 * (`dy = pointer.y - center.y < 0`) yields 0°, to the right yields 90°, below
 * 180°, and to the left 270°. When `snap` is true the result is snapped to the
 * nearest {@link ROTATE_SNAP_DEGREES}° (Shift-rotate). Pure + deterministic —
 * no DOM, no Date/random; a zero-length vector (pointer exactly on center)
 * returns 0.
 */
export function angleFromPointer(
  pointer: Point2D,
  center: Point2D,
  snap = false,
): number {
  const dx = pointer.x - center.x;
  const dy = pointer.y - center.y;
  if (dx === 0 && dy === 0) return 0;
  // atan2(dx, -dy): measures the angle from the UP axis, clockwise-positive in
  // screen space (Y-down). atan2(dx, -dy) == 0 when straight up (dy<0), +90°
  // when to the right (dx>0), etc.
  const radians = Math.atan2(dx, -dy);
  const degrees = normalizeRotationDegrees((radians * 180) / Math.PI);
  return snap ? snapAngle(degrees) : degrees;
}
