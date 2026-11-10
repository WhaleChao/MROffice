/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry for ANCHORED CONNECTORS (PowerPoint / WPS "connector" parity).
 *
 * A connector is NOT a new block type: it is an ordinary `line`/`arrow` `shape`
 * block that has GLUED itself to two sibling shapes by carrying two optional
 * passthrough fields — `startShapeId` + `endShapeId`. When BOTH ids resolve to
 * real sibling shapes the connector is "anchored": its bounding box and its
 * `flipH`/`flipV` mirror flags are recomputed from the referenced shapes'
 * CURRENT rects so the line follows them when they move / resize. A connector
 * with no anchors (or with an unresolved anchor) is left exactly as authored.
 *
 * The anchor point is the CENTER of each shape's rect (simple + robust: it never
 * depends on the connector's own geometry, so recompute is idempotent and never
 * needs a second pass). The connector's box is the rect spanning the two anchor
 * points, and the DIRECTION is carried the way DrawingML carries it: a preset
 * connector always runs `(l,t) → (r,b)` inside its own box, and `@flipH`/`@flipV`
 * mirror that box. So `flipH` is exactly "the start point is to the RIGHT of the
 * end point" and `flipV` is "the start point is BELOW the end point" — the same
 * derivation PowerPoint's own writer uses, and the same one the Docs canvas uses
 * (`_document_shape_connector_route` in `routers/notebooks.py`).
 *
 * LEGACY `lineSlope`. Before the flips, an anchored connector carried a
 * `"tlbr"`/`"trbl"` string naming which diagonal of the box it ran along, and
 * the slide canvas drew a hand-written route from it. `"trbl"` is exactly the
 * horizontal mirror of `(l,t)→(r,b)`, i.e. `flipH` — see
 * {@link legacyLineSlopeFlipsHorizontally}. The field is never written again;
 * it stays READABLE so pages saved before this change keep their on-slide angle
 * and export the bytes they always exported.
 *
 * All geometry is in EMU (914,400 per inch). Every function is TOTAL and
 * side-effect-free; mirrors `presentation-smartart.ts`'s pure-module shape so it
 * is unit-testable in isolation (no React / no DOM).
 */

import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import {
  isPresentationConnectorShapeKind,
  normalizeShapeKind,
  shapePolygonPoints,
} from "./presentation-shapes";
import { PRESENTATION_OBJECT_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";

/**
 * LEGACY, read-only: which diagonal of the connector's bounding box the line ran
 * along. Superseded by `flipH`/`flipV`; kept so decks saved before the change
 * still render and export as they did.
 */
export type LineSlope = "tlbr" | "trbl";

/** The two slope values, for validation / exhaustiveness. */
export const LINE_SLOPES: readonly LineSlope[] = ["tlbr", "trbl"];

export function isLineSlope(value: unknown): value is LineSlope {
  return value === "tlbr" || value === "trbl";
}

/**
 * Whether a LEGACY `lineSlope` mirrors the DrawingML route horizontally.
 *
 * Every connector preset runs `(l,t) → (r,b)` in its own box, so the
 * top-right→bottom-left diagonal IS that route under `@flipH`. `"tlbr"` (and an
 * absent / unrecognised value) is the unmirrored route. The result COMPOSES with
 * the block's own `flipH` — a deck that was anchored (`lineSlope`) and then
 * mirrored by hand (`flipH`) drew both, so both still apply.
 */
export function legacyLineSlopeFlipsHorizontally(value: unknown): boolean {
  return value === "trbl";
}

/** An axis-aligned rectangle in EMU. */
export interface Rect {
  left: number;
  top: number;
  width: number;
  height: number;
  /**
   * The shape's own rotation (degrees, clockwise) and mirrors. A connection
   * site is a point on the shape, so it turns and mirrors with it; absent is
   * an unrotated, unmirrored shape.
   */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  






  outline?: string | null;
}

/**
 * A connector's recomputed geometry: a bounding box plus the DrawingML mirror
 * flags that turn the preset's `(l,t) → (r,b)` route into the real direction.
 */
export interface ConnectorGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
  flipH: boolean;
  flipV: boolean;
}

/**
 * The minimum box size (EMU) a connector is guarded to when its two anchor
 * centers coincide (or are sub-EMU apart). A zero-size box would render nothing
 * and could divide-by-zero downstream, so we floor it to a tiny positive square
 * — small enough to be visually a point, large enough to be a valid block.
 */
const MIN_CONNECTOR_SIZE = 1;

/**
 * The schema's hard EMU bound for block geometry (mirrors
 * PRESENTATION_GEOMETRY_EMU_LIMIT in native-page-schemas.ts and
 * MAX_PRESENTATION_GEOMETRY_EMU in PresentationWindow). Two anchors clamped to the
 * extreme opposite edges can be ~2×this apart, so the span MUST be clamped here or
 * the connector block exceeds the cap and the autosave is rejected (422).
 */
const MAX_CONNECTOR_EMU = PRESENTATION_OBJECT_GEOMETRY_MAX_EMU;

/** Clamp a coordinate to the schema's EMU window. */
function clampCoord(value: number): number {
  return Math.max(-MAX_CONNECTOR_EMU, Math.min(MAX_CONNECTOR_EMU, value));
}

/** Clamp a size to [MIN_CONNECTOR_SIZE, MAX_CONNECTOR_EMU]. */
function clampSize(value: number): number {
  return Math.max(MIN_CONNECTOR_SIZE, Math.min(MAX_CONNECTOR_EMU, value));
}

function centerX(rect: Rect): number {
  return rect.left + rect.width / 2;
}

function centerY(rect: Rect): number {
  return rect.top + rect.height / 2;
}

/**
 * Compute an anchored connector's geometry from its two endpoint shapes' rects.
 *
 * The connector is glued CENTER-to-CENTER: the bounding box is the rect spanning
 * the two centers, and `flipH`/`flipV` mirror the preset's `(l,t) → (r,b)` route
 * so it starts at the start shape. Coincident (or sub-EMU) centers are guarded to
 * a minimal positive box so the result is always a valid, renderable rect.
 *
 * Deterministic + total: identical inputs always yield identical output, and the
 * result depends ONLY on the two source rects (never on the connector's own
 * prior geometry), so applying it repeatedly is a fixed point.
 */
/**
 * Connection-site anchors (PPT's four cardinal sites + the legacy center).
 * Absent field ⇒ "center", which keeps every existing deck byte-identical —
 * side anchors are opt-in per endpoint.
 */
export const CONNECTOR_ANCHORS = [
  "center",
  "top",
  "right",
  "bottom",
  "left",
] as const;
export type ConnectorAnchor = (typeof CONNECTOR_ANCHORS)[number];

const CONNECTOR_ANCHOR_SET: ReadonlySet<string> = new Set(CONNECTOR_ANCHORS);

export function normalizeConnectorAnchor(value: unknown): ConnectorAnchor {
  return typeof value === "string" && CONNECTOR_ANCHOR_SET.has(value)
    ? (value as ConnectorAnchor)
    : "center";
}

/**
 * The point an anchor names on a rect. The site is found on the unrotated
 * shape, mirrored by the shape's flips and turned by its rotation about the
 * shape's centre — DrawingML's `a:xfrm` order — so a connector stays glued to
 * the same site of a rotated or flipped shape, as it does in PowerPoint.
 */
/**
 * Where the midline from the outline's centre (50,50) towards a side meets
 * the outline (0-100 unit space), or null when the outline cannot be read.
 * The farthest crossing along the ray is the one on the shape's silhouette.
 */
function outlineHitTowards(outline: string, sx: number, sy: number): { x: number; y: number } | null {
  const points = outline
    .trim()
    .split(/\s+/)
    .map((pair) => pair.split(",").map(Number))
    .filter((pair): pair is [number, number] => pair.length === 2 && pair.every(Number.isFinite));
  if (points.length < 3) return null;
  let best: { x: number; y: number; t: number } | null = null;
  for (let index = 0; index < points.length; index += 1) {
    const [x1, y1] = points[index];
    const [x2, y2] = points[(index + 1) % points.length];
    // Ray P = (50,50) + t * (sx, sy), t > 0; edge Q = (x1,y1) + u * (x2-x1, y2-y1), 0..1.
    const ex = x2 - x1;
    const ey = y2 - y1;
    const denominator = sx * ey - sy * ex;
    if (Math.abs(denominator) < 1e-9) continue;
    const wx = x1 - 50;
    const wy = y1 - 50;
    const t = (wx * ey - wy * ex) / denominator;
    const u = (wx * sy - wy * sx) / denominator;
    if (t <= 0 || u < -1e-9 || u > 1 + 1e-9) continue;
    if (!best || t > best.t) best = { x: 50 + sx * t, y: 50 + sy * t, t };
  }
  return best ? { x: best.x, y: best.y } : null;
}

export function resolveAnchorPoint(
  rect: Rect,
  anchor: ConnectorAnchor,
): { x: number; y: number } {
  const cx = centerX(rect);
  const cy = centerY(rect);
  let dx = 0;
  let dy = 0;
  switch (anchor) {
    case "top":
      dy = -rect.height / 2;
      break;
    case "bottom":
      dy = rect.height / 2;
      break;
    case "left":
      dx = -rect.width / 2;
      break;
    case "right":
      dx = rect.width / 2;
      break;
    default:
      break;
  }
  if ((dx !== 0 || dy !== 0) && rect.outline) {
    const hit = outlineHitTowards(rect.outline, dx === 0 ? 0 : Math.sign(dx), dy === 0 ? 0 : Math.sign(dy));
    if (hit) {
      dx = ((hit.x - 50) / 100) * rect.width;
      dy = ((hit.y - 50) / 100) * rect.height;
    }
  }
  if (rect.flipH === true) dx = -dx;
  if (rect.flipV === true) dy = -dy;
  const rotation =
    typeof rect.rotation === "number" && Number.isFinite(rect.rotation)
      ? ((rect.rotation % 360) + 360) % 360
      : 0;
  if (rotation === 0 || (dx === 0 && dy === 0)) {
    return { x: cx + dx, y: cy + dy };
  }
  // Quarter turns are exact, so a site on a rotated shape does not pick up
  // floating-point drift that would move the connector by an EMU.
  const quarter = rotation / 90;
  if (Number.isInteger(quarter)) {
    const [x, y] = [
      [dx, dy],
      [-dy, dx],
      [-dx, -dy],
      [dy, -dx],
    ][quarter];
    return { x: cx + x, y: cy + y };
  }
  const radians = (rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}

/**
 * PPT-style auto-routing for a NEW connector: of the 4×4 side pairings pick
 * the one with the shortest span (ties resolve by list order, so the result
 * is deterministic). Centers are excluded — auto always lands on sides; the
 * legacy center behaviour remains the DEFAULT for connectors that never chose.
 */
export function nearestAnchorPair(
  startRect: Rect,
  endRect: Rect,
): { start: ConnectorAnchor; end: ConnectorAnchor } {
  const sides: ConnectorAnchor[] = ["top", "right", "bottom", "left"];
  let best: { start: ConnectorAnchor; end: ConnectorAnchor } = {
    start: "right",
    end: "left",
  };
  let bestDist = Infinity;
  for (const a of sides) {
    const pa = resolveAnchorPoint(startRect, a);
    for (const b of sides) {
      const pb = resolveAnchorPoint(endRect, b);
      const d = (pa.x - pb.x) ** 2 + (pa.y - pb.y) ** 2;
      if (d < bestDist) {
        bestDist = d;
        best = { start: a, end: b };
      }
    }
  }
  return best;
}

export function connectorGeometry(
  startRect: Rect,
  endRect: Rect,
  startAnchor: ConnectorAnchor = "center",
  endAnchor: ConnectorAnchor = "center",
): ConnectorGeometry {
  return connectorGeometryBetween(
    resolveAnchorPoint(startRect, startAnchor),
    resolveAnchorPoint(endRect, endAnchor),
  );
}

/** A connector's box and mirrors for a route from `startPoint` to `endPoint`. */
function connectorGeometryBetween(
  startPoint: { x: number; y: number },
  endPoint: { x: number; y: number },
): ConnectorGeometry {
  const sx = startPoint.x;
  const sy = startPoint.y;
  const ex = endPoint.x;
  const ey = endPoint.y;

  const rawLeft = Math.min(sx, ex);
  const rawTop = Math.min(sy, ey);
  const rawWidth = Math.abs(ex - sx);
  const rawHeight = Math.abs(ey - sy);

  // Round to whole EMU (block geometry is integer EMU everywhere else) and floor
  // the size to a minimal positive box so coincident centers stay renderable.
  // Clamp to the schema's EMU window: two shapes clamped to opposite extreme
  // edges are ~2×MAX apart, so an un-clamped span would exceed the cap and the
  // backend would reject the autosave (and the zod schema would strip the field).
  const left = clampCoord(Math.round(rawLeft));
  const top = clampCoord(Math.round(rawTop));
  const width = clampSize(Math.round(rawWidth));
  const height = clampSize(Math.round(rawHeight));

  // DrawingML's own derivation: the preset route always leaves the box's
  // top-left and arrives at its bottom-right, so a start point RIGHT of the end
  // point is `@flipH` and a start point BELOW the end point is `@flipV`. This is
  // byte-for-byte what PowerPoint writes and what python-pptx's `_add_cxnSp`
  // computes, so an exported connector re-imports with these same two flags.
  // Equal coordinates mean the box is degenerate on that axis (both mirrors
  // collapse to the same segment), and false keeps the block minimal.
  return { left, top, width, height, flipH: sx > ex, flipV: sy > ey };
}

/** Read a numeric geometry field off a block, or null when absent / non-finite. */
function geomNumber(block: PresentationBlock, key: keyof Rect): number | null {
  const value = (block as Record<string, unknown>)[key];
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/**
 * A block's full rect, treating a missing left/top as 0 and a missing size as
 * 0, with the rotation and mirrors its connection sites turn with.
 */
function blockRect(block: PresentationBlock): Rect {
  const shape = (block as { shape?: unknown }).shape;
  return presentationConnectorRectTurn(block, {
    left: geomNumber(block, "left") ?? 0,
    top: geomNumber(block, "top") ?? 0,
    width: geomNumber(block, "width") ?? 0,
    height: geomNumber(block, "height") ?? 0,
    outline:
      block.type === "shape" && typeof shape === "string"
        ? shapePolygonPoints(normalizeShapeKind(shape))
        : null,
  });
}

/** `rect` with the block's own rotation and mirrors, when it states any. */
export function presentationConnectorRectTurn(
  block: PresentationBlock,
  rect: Rect,
): Rect {
  const record = block as Record<string, unknown>;
  const rotation = record.rotation;
  return {
    ...rect,
    ...(typeof rotation === "number" && Number.isFinite(rotation) && rotation !== 0
      ? { rotation }
      : {}),
    ...(record.flipH === true ? { flipH: true } : {}),
    ...(record.flipV === true ? { flipV: true } : {}),
  };
}

/** True for any stable connector shape block. */
function isLinearShape(block: PresentationBlock): boolean {
  if (block.type !== "shape") return false;
  return isPresentationConnectorShapeKind(
    normalizeShapeKind((block as Record<string, unknown>).shape),
  );
}

/** Read a connector's anchor id field, or null when absent / not a non-empty string. */
function anchorId(
  block: PresentationBlock,
  key: "startShapeId" | "endShapeId",
): string | null {
  const value = (block as Record<string, unknown>)[key];
  return typeof value === "string" && value.length > 0 ? value : null;
}

/**
 * Whether a block is an ANCHORED connector: a `line`/`arrow` shape carrying both
 * a `startShapeId` and an `endShapeId`. (Whether those ids currently RESOLVE is a
 * separate check done during recompute.)
 */
export function isAnchoredConnector(block: PresentationBlock): boolean {
  return (
    isLinearShape(block) &&
    anchorId(block, "startShapeId") !== null &&
    anchorId(block, "endShapeId") !== null
  );
}








































export function recomputeConnectors(
  blocks: readonly PresentationBlock[],
  editedIds?: ReadonlySet<string>,
): PresentationBlock[] {
  if (blocks.length === 0) return [];

  // Index sibling rects by id from the ORIGINAL blocks. A connector reads its
  // anchors from here, so it is always recomputed from pre-recompute geometry —
  // guaranteeing a single pass and no connector-chaining feedback. We include
  // EVERY block (defensive: a connector referencing another connector reads that
  // connector's original — not recomputed — rect, so still one pass).
  const rectById = new Map<string, Rect>();
  for (const block of blocks) {
    if (typeof block.id === "string" && block.id.length > 0) {
      rectById.set(block.id, blockRect(block));
    }
  }

  return blocks.map((block) => {
    if (editedIds && editedIds.has(block.id)) return block;
    const startId = anchorId(block, "startShapeId");
    const endId = anchorId(block, "endShapeId");
    if (
      editedIds &&
      !(startId !== null && editedIds.has(startId)) &&
      !(endId !== null && editedIds.has(endId))
    ) {
      return block;
    }
    if (!isAnchoredConnector(block)) return recomputeOneEndConnector(block, rectById);

    // Guarded by isAnchoredConnector, but re-narrow for the type checker.
    if (startId === null || endId === null) return block;

    const startRect = rectById.get(startId);
    const endRect = rectById.get(endId);
    // Dangling: at least one anchor no longer resolves — leave the block as-is.
    if (!startRect || !endRect) return block;

    const record = block as Record<string, unknown>;
    const geom = connectorGeometry(
      startRect,
      endRect,
      normalizeConnectorAnchor(record.startAnchor),
      normalizeConnectorAnchor(record.endAnchor),
    );
    // Deep-fresh ONLY the changed block; carry every other field verbatim.
    const next = {
      ...block,
      left: geom.left,
      top: geom.top,
      width: geom.width,
      height: geom.height,
      flipH: geom.flipH,
      flipV: geom.flipV,
    } as PresentationBlock & Record<string, unknown>;
    // An anchored connector's direction is DERIVED from its two anchor points,
    // so the recompute owns both flags: a false one is dropped (minimal JSON)
    // and the legacy diagonal string is dropped with them.
    if (!geom.flipH) delete next.flipH;
    if (!geom.flipV) delete next.flipV;
    delete next.lineSlope;
    // The route is laid out upright between its two sites. A turn the
    // connector took with a rotated group stays behind after Ungroup, and
    
    delete next.rotation;
    return next;
  });
}

/**
 * A connector glued at ONE end (its other end free on the slide: drawn from a
 * shape into empty space, or copied with only one of the shapes it joined)
 * follows that shape as PowerPoint's does -- "the connectors remain attached
 * to and move with the shapes" (Microsoft Support, "Draw or delete a line or
 * connector") holds end by end. The glued end goes to its site on the
 * shape's current rect; the free end stays where the connector draws it now.
 * The block is returned by reference when nothing moves (or the shape is gone,
 * or the connector is turned: its free end is then not a corner of its box),
 * so a still slide stays byte-stable. The glued point is rounded to whole EMU
 * first, so the free end read back from the box never drifts.
 */
function recomputeOneEndConnector(
  block: PresentationBlock,
  rectById: ReadonlyMap<string, Rect>,
): PresentationBlock {
  if (!isLinearShape(block)) return block;
  const startId = anchorId(block, "startShapeId");
  const endId = anchorId(block, "endShapeId");
  if ((startId === null) === (endId === null)) return block;
  const record = block as Record<string, unknown>;
  const rotation = record.rotation;
  if (typeof rotation === "number" && Number.isFinite(rotation) && rotation % 360 !== 0) return block;
  const glued = rectById.get((startId ?? endId) as string);
  const left = geomNumber(block, "left");
  const top = geomNumber(block, "top");
  const width = geomNumber(block, "width");
  const height = geomNumber(block, "height");
  if (!glued || left === null || top === null || width === null || height === null) return block;
  // Where the connector's ends are drawn now: the preset runs (l,t) -> (r,b)
  // in its box, mirrored by its flips (a legacy diagonal string mirrors it
  // horizontally as well).
  const flipH = (record.flipH === true) !== legacyLineSlopeFlipsHorizontally(record.lineSlope);
  const flipV = record.flipV === true;
  const start = { x: flipH ? left + width : left, y: flipV ? top + height : top };
  const end = { x: flipH ? left : left + width, y: flipV ? top : top + height };
  const site = resolveAnchorPoint(
    glued,
    normalizeConnectorAnchor(startId !== null ? record.startAnchor : record.endAnchor),
  );
  const gluedPoint = { x: Math.round(site.x), y: Math.round(site.y) };
  const geom = startId !== null
    ? connectorGeometryBetween(gluedPoint, end)
    : connectorGeometryBetween(start, gluedPoint);
  if (
    geom.left === left &&
    geom.top === top &&
    geom.width === width &&
    geom.height === height &&
    geom.flipH === flipH &&
    geom.flipV === flipV
  ) {
    // Drawn as it stands (a legacy diagonal string included): unchanged.
    return block;
  }
  const next = {
    ...block,
    left: geom.left,
    top: geom.top,
    width: geom.width,
    height: geom.height,
    flipH: geom.flipH,
    flipV: geom.flipV,
  } as PresentationBlock & Record<string, unknown>;
  if (!geom.flipH) delete next.flipH;
  if (!geom.flipV) delete next.flipV;
  delete next.lineSlope;
  return next;
}

/**
 * Shape Format > Edit Shape > Reroute Connectors (PPT-A124; Microsoft
 * Support, "Draw or delete a line or connector"): each selected connector
 * glued at both ends moves to the closest pair of connection sites of the
 * two shapes it joins, then is redrawn between them. Connectors not glued at
 * both ends, or whose shapes are gone, are left as they are. Returns null
 * when nothing changes.
 */
export function rerouteConnectors(
  blocks: readonly PresentationBlock[],
  connectorIds: ReadonlySet<string>,
): PresentationBlock[] | null {
  const rectById = new Map<string, Rect>();
  for (const block of blocks) {
    if (typeof block.id === "string" && block.id.length > 0) rectById.set(block.id, blockRect(block));
  }
  let changed = false;
  const anchored = blocks.map((block) => {
    if (!connectorIds.has(block.id) || !isAnchoredConnector(block)) return block;
    const startRect = rectById.get(anchorId(block, "startShapeId") ?? "");
    const endRect = rectById.get(anchorId(block, "endShapeId") ?? "");
    if (!startRect || !endRect) return block;
    const pair = nearestAnchorPair(startRect, endRect);
    const record = block as Record<string, unknown>;
    if (normalizeConnectorAnchor(record.startAnchor) === pair.start && normalizeConnectorAnchor(record.endAnchor) === pair.end) {
      return block;
    }
    changed = true;
    return { ...block, startAnchor: pair.start, endAnchor: pair.end } as PresentationBlock;
  });
  if (!changed) return null;
  // Only the connectors the command names are redrawn.
  const redrawn = recomputeConnectors(anchored);
  return anchored.map((block, index) => (connectorIds.has(block.id) ? redrawn[index] : block));
}
