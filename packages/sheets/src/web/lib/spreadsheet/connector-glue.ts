/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Connector glue on the worksheet, on the deck's connector geometry.
 *
 * Office semantics (Microsoft Support, "Draw or delete a line or connector";
 * Excel and PowerPoint): a connector whose end is dropped on a shape's
 * connection site stays joined to it -- moving or resizing the shape re-routes
 * the connector, and a connector moved off the site lets go. The deck keeps
 * this as `startShapeId` / `endShapeId` + `startAnchor` / `endAnchor` on the
 * connector block and recomputes with `presentation-connectors`; the sheet
 * keeps the same four fields on the connector object (`startObjectId` /
 * `endObjectId` naming sibling drawing objects) and uses the same pure
 * `connectorGeometry` / `resolveAnchorPoint`, in the sheet's CSS pixels.
 */

import {
  connectorGeometry,
  nearestAnchorPair,
  normalizeConnectorAnchor,
  resolveAnchorPoint,
  type ConnectorAnchor,
  type Rect,
} from "@mrnote/core/web/lib/presentation-connectors.ts";

/** How near (CSS pixels) a connector's end must land to a site to glue. */
export const SPREADSHEET_CONNECTOR_GLUE_DISTANCE_PX = 10;

const SITES: readonly ConnectorAnchor[] = ["top", "right", "bottom", "left"];

export interface SpreadsheetGlueObject {
  readonly id: string;
  readonly type: string;
  readonly startObjectId?: unknown;
  readonly endObjectId?: unknown;
  readonly startAnchor?: unknown;
  readonly endAnchor?: unknown;
  readonly anchor?: { readonly flipH?: boolean; readonly flipV?: boolean } | undefined;
}

/** The connector's two ends in its box, from the DrawingML flips. */
export function spreadsheetConnectorEnds(
  box: Rect,
  flipH = false,
  flipV = false,
): { start: { x: number; y: number }; end: { x: number; y: number } } {
  const left = box.left;
  const right = box.left + box.width;
  const top = box.top;
  const bottom = box.top + box.height;
  return {
    start: { x: flipH ? right : left, y: flipV ? bottom : top },
    end: { x: flipH ? left : right, y: flipV ? top : bottom },
  };
}

/** The shape site nearest `point` within the glue distance, or null. */
export function spreadsheetNearestGlueSite(
  point: { x: number; y: number },
  candidates: readonly { id: string; rect: Rect }[],
  distance = SPREADSHEET_CONNECTOR_GLUE_DISTANCE_PX,
): { id: string; anchor: ConnectorAnchor } | null {
  let best: { id: string; anchor: ConnectorAnchor; d: number } | null = null;
  for (const candidate of candidates) {
    for (const anchor of SITES) {
      const site = resolveAnchorPoint(candidate.rect, anchor);
      const d = Math.hypot(site.x - point.x, site.y - point.y);
      if (d <= distance && (!best || d < best.d)) best = { id: candidate.id, anchor, d };
    }
  }
  return best ? { id: best.id, anchor: best.anchor } : null;
}

/**
 * The glue fields a connector takes after being dropped at `box`: each end
 * glues to the site it landed on, or lets go.
 */
export function spreadsheetConnectorGlueAfterMove(
  box: Rect,
  flipH: boolean,
  flipV: boolean,
  shapes: readonly { id: string; rect: Rect }[],
): {
  startObjectId?: string;
  startAnchor?: ConnectorAnchor;
  endObjectId?: string;
  endAnchor?: ConnectorAnchor;
} {
  const ends = spreadsheetConnectorEnds(box, flipH, flipV);
  const start = spreadsheetNearestGlueSite(ends.start, shapes);
  const end = spreadsheetNearestGlueSite(ends.end, shapes.filter((shape) => shape.id !== start?.id));
  return {
    ...(start ? { startObjectId: start.id, startAnchor: start.anchor } : {}),
    ...(end ? { endObjectId: end.id, endAnchor: end.anchor } : {}),
  };
}

/**
 * Shape Format > Edit Shape > Reroute Connectors (Excel, Word and PowerPoint
 * alike): each of `connectorIds` glued at both ends to objects that still
 * exist moves to the closest pair of their connection sites -- the deck's
 * `nearestAnchorPair` -- and is redrawn between them. Null when nothing
 * changes.
 */
export function rerouteSpreadsheetConnectors<T extends SpreadsheetGlueObject>(
  objects: readonly T[],
  connectorIds: ReadonlySet<string>,
  rectOf: (object: T) => Rect,
  withBox: (connector: T, box: Rect, flipH: boolean, flipV: boolean) => T,
): readonly T[] | null {
  const byId = new Map(objects.map((object) => [object.id, object]));
  let changed = false;
  const anchored = objects.map((object) => {
    if (!connectorIds.has(object.id)) return object;
    const start = typeof object.startObjectId === "string" ? byId.get(object.startObjectId) : undefined;
    const end = typeof object.endObjectId === "string" ? byId.get(object.endObjectId) : undefined;
    if (!start || !end) return object;
    const pair = nearestAnchorPair(rectOf(start), rectOf(end));
    if (
      normalizeConnectorAnchor(object.startAnchor) === pair.start &&
      normalizeConnectorAnchor(object.endAnchor) === pair.end
    ) {
      return object;
    }
    changed = true;
    return { ...object, startAnchor: pair.start, endAnchor: pair.end };
  });
  if (!changed) return null;
  return recomputeSpreadsheetGluedConnectors(anchored, rectOf, withBox);
}

/** Whether an object is a connector glued at both ends. */
export function spreadsheetConnectorIsGlued(object: SpreadsheetGlueObject): boolean {
  return typeof object.startObjectId === "string" && typeof object.endObjectId === "string";
}

/**
 * Every connector glued at BOTH ends to objects that still exist, re-routed
 * from those objects' current rects (one pass, from the input's rects, as the
 * deck's `recomputeConnectors` does). `rectOf` reads an object's rect;
 * `withBox` writes a connector's new box and flips. Returns the same array
 * when nothing moved.
 */
export function recomputeSpreadsheetGluedConnectors<T extends SpreadsheetGlueObject>(
  objects: readonly T[],
  rectOf: (object: T) => Rect,
  withBox: (connector: T, box: Rect, flipH: boolean, flipV: boolean) => T,
  skipId?: string,
): readonly T[] {
  const byId = new Map(objects.map((object) => [object.id, object]));
  let changed = false;
  const next = objects.map((object) => {
    if (object.id === skipId) return object;
    const startId = typeof object.startObjectId === "string" ? object.startObjectId : null;
    const endId = typeof object.endObjectId === "string" ? object.endObjectId : null;
    if (!startId || !endId) return object;
    const start = byId.get(startId);
    const end = byId.get(endId);
    if (!start || !end) return object;
    const geometry = connectorGeometry(
      rectOf(start),
      rectOf(end),
      normalizeConnectorAnchor(object.startAnchor),
      normalizeConnectorAnchor(object.endAnchor),
    );
    const current = rectOf(object);
    if (
      Math.abs(current.left - geometry.left) < 0.5 &&
      Math.abs(current.top - geometry.top) < 0.5 &&
      Math.abs(current.width - geometry.width) < 0.5 &&
      Math.abs(current.height - geometry.height) < 0.5 &&
      (object.anchor?.flipH === true) === geometry.flipH &&
      (object.anchor?.flipV === true) === geometry.flipV
    ) {
      return object;
    }
    changed = true;
    return withBox(object, geometry, geometry.flipH, geometry.flipV);
  });
  return changed ? next : objects;
}
