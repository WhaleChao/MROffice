/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure geometry math for PowerPoint-parity marquee (rubber-band) multi-select on
 * the presentation editor canvas: pressing on EMPTY slide canvas and dragging
 * draws a selection rectangle, and on release every block that INTERSECTS the
 * marquee becomes selected.
 *
 * PowerPoint (and Keynote) select on INTERSECTION — a block is captured if the
 * marquee overlaps it at all, NOT only when it is fully contained. This module
 * implements that intersection hit-test as a single pure function so it can be
 * unit-tested in isolation and reused by the editor component.
 *
 * Dependency-free (no React / no DOM). All values are in the same EMU coordinate
 * space the rest of the presentation stack uses (`left`/`top`/`width`/`height`);
 * the caller converts the pointer's pixel rect into EMU before calling, exactly
 * as the drag gesture does (`presentation-drag` / `presentation-align`).
 *
 * Mirrors the conventions of the sibling pure modules (`presentation-drag.ts`,
 * `presentation-align.ts`): shared {@link BlockGeometry} shape, finite/normalized
 * inputs, small composable pure functions, no mutation of inputs.
 */

import { type BlockGeometry } from "@mrnote/core/web/lib/presentation-drag.ts";
import { rotatedBlockBounds } from "@mrnote/core/web/lib/presentation-multi-transform.ts";

export type { BlockGeometry } from "@mrnote/core/web/lib/presentation-drag.ts";

/**
 * A block paired with its id, as the hit-test consumes it. The geometry fields
 * are the same EMU rect as everywhere else; any may be missing / non-finite (a
 * partially-populated block) and is normalized away before the test.
 */
export interface MarqueeBlock extends Partial<BlockGeometry> {
  id: string;
  /** The block's rotation about its centre, in degrees; the marquee meets the rotated frame. */
  rotation?: number;
}

/**
 * The marquee rectangle, as the pointer produces it. Because the user can drag
 * in any direction, `width`/`height` MAY be negative (a drag up/left): the
 * helper normalizes the rect (min/max corners) before testing so a negative
 * drag selects the same blocks as the equivalent positive drag.
 */
export interface MarqueeRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** An axis-aligned rect by its four edges (all finite). */
interface Edges {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Coerce to a finite number, mapping NaN/Infinity → 0. */
function finite(value: number | undefined): number {
  return Number.isFinite(value) ? (value as number) : 0;
}

/**
 * Normalize a (possibly negative-extent) marquee rect into ordered edges. A
 * drag up/left produces negative width/height; taking min/max of the two
 * corners yields a well-ordered box (`right >= left`, `bottom >= top`) so the
 * intersection test is direction-agnostic.
 */
function marqueeEdges(rect: MarqueeRect | null | undefined): Edges {
  const left = finite(rect?.left);
  const top = finite(rect?.top);
  const width = finite(rect?.width);
  const height = finite(rect?.height);
  const x0 = left;
  const x1 = left + width;
  const y0 = top;
  const y1 = top + height;
  return {
    left: Math.min(x0, x1),
    right: Math.max(x0, x1),
    top: Math.min(y0, y1),
    bottom: Math.max(y0, y1),
  };
}

/**
 * Normalize a block's geometry into ordered edges. Missing / non-finite fields
 * collapse to 0; a negative width/height (degenerate imported geometry) is
 * ordered the same way the marquee is so the test never depends on input sign.
 */
function blockEdges(block: MarqueeBlock | Partial<BlockGeometry>): Edges {
  // A rotated block is met where it is drawn, not where its unrotated frame
  // was: a marquee over the old blank spot selected nothing visible (PPT-B208).
  const bounds = rotatedBlockBounds(
    {
      left: finite(block.left),
      top: finite(block.top),
      width: finite(block.width),
      height: finite(block.height),
    },
    "rotation" in block ? block.rotation : undefined,
  );
  const left = bounds.left;
  const top = bounds.top;
  const right = left + bounds.width;
  const bottom = top + bounds.height;
  return {
    left: Math.min(left, right),
    right: Math.max(left, right),
    top: Math.min(top, bottom),
    bottom: Math.max(top, bottom),
  };
}

/** True when the marquee rect has zero (or inverted-to-zero) area on EITHER axis. */
function isZeroArea(edges: Edges): boolean {
  return edges.right === edges.left || edges.bottom === edges.top;
}

/**
 * Axis-aligned bounding-box intersection (inclusive of a shared edge). Two rects
 * intersect when they overlap on BOTH axes; a shared boundary line (e.g. the
 * marquee's right edge exactly on the block's left edge) counts as a hit, which
 * matches the marquee visually touching the block. Both rects are assumed
 * already ordered (`right >= left`, `bottom >= top`).
 */
function edgesIntersect(a: Edges, b: Edges): boolean {
  return (
    a.left <= b.right &&
    a.right >= b.left &&
    a.top <= b.bottom &&
    a.bottom >= b.top
  );
}

/**
 * Return the ids of every block whose rect INTERSECTS the marquee rectangle
 * (PowerPoint rubber-band selection — intersect, not fully-contained). Ids are
 * returned in input order; duplicates in the input are preserved positionally
 * (the caller passes the slide's blocks, which have unique ids).
 *
 * Safety / edge behaviour:
 *  - The marquee rect may have NEGATIVE width/height (a drag up/left); it is
 *    normalized to min/max corners first, so a negative drag selects the same
 *    blocks as the mirror-image positive drag.
 *  - A ZERO-AREA marquee (no real drag — width or height collapses to 0) selects
 *    NOTHING and returns `[]`. This is what lets the caller treat a sub-threshold
 *    press as a plain empty-canvas click (clear selection) rather than leaving a
 *    stuck one-block marquee.
 *  - Non-finite / missing fields on the marquee or a block coerce to 0 (no
 *    throw); a `null`/`undefined` marquee or block list returns `[]`.
 *  - A shared edge (touching boundary, zero overlap area) counts as an
 *    intersection (inclusive test).
 *
 * Pure + deterministic: no DOM, no Date/random; inputs are never mutated.
 */
export function blocksInMarquee(
  marquee: MarqueeRect | null | undefined,
  blocks: ReadonlyArray<MarqueeBlock | null | undefined> | null | undefined,
): string[] {
  if (!marquee || !Array.isArray(blocks) || blocks.length === 0) return [];
  const marquEdges = marqueeEdges(marquee);
  // No real drag → no marquee selection (caller falls back to a plain click).
  if (isZeroArea(marquEdges)) return [];
  const ids: string[] = [];
  for (const block of blocks) {
    if (!block || typeof block.id !== "string") continue;
    if (edgesIntersect(marquEdges, blockEdges(block))) {
      ids.push(block.id);
    }
  }
  return ids;
}
