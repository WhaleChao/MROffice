/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Presentation ZOOM tiles — PowerPoint / WPS "Insert ▸ Zoom" parity.
 *
 * A zoom tile is a navigation object placed ON a slide: a labelled thumbnail of a
 * target slide / section. In PRESENT mode, activating it JUMPS to that target.
 *   - SLIDE  zoom: jump to one specific slide.
 *   - SECTION zoom: jump to the first slide of a deck section.
 *   - SUMMARY zoom: a menu slide with one SECTION tile per section. The editor
 *     no longer builds one (PowerPoint for the web cannot create zooms, PPT
 *     A052); a deck that already holds summary tiles still plays them.
 *
 * MODEL CHOICE (documented) — REUSE, not a new block kind:
 * A zoom tile is just a SHAPE block (`type: "shape"`) that ALSO carries an
 * internal `slide://<slideId>` {@link href} plus two additive marker fields:
 *   - `zoomKind`:     "slide" | "section" | "summary" (its navigation intent),
 *   - `zoomTargetId`: the id of the target slide (slide zoom) or section
 *                     (section / summary tile).
 * Reasons:
 *  1. The shape block already round-trips through the deck schema, the slide
 *     renderer, serialize/import, geometry/drag, undo — a zoom tile inherits ALL
 *     of that for free instead of bolting on a new union member everywhere.
 *  2. The `slide://` href reuses the EXISTING present-mode navigation path end to
 *     end: `parseSlideLinkTarget` → `onNavigateToSlide` → `onJumpToSlide(index)`.
 *     A zoom tile therefore navigates with ZERO new present-mode wiring.
 *  3. BYTE-STABILITY: the marker fields are optional + additive and the shape
 *     schema is `.passthrough()`, so a deck with NO zoom tiles serializes
 *     byte-for-byte unchanged (no shape block carries `zoomKind`).
 *
 * RESOLUTION (the testable core): a tile's `zoomKind` + `zoomTargetId` resolve to
 * a destination slide INDEX against the live deck:
 *   - slide:   the index of the slide whose id is `zoomTargetId`.
 *   - section / summary tile: the index of the section's FIRST slide (its anchor),
 *     i.e. the same first-member rule the section grouping uses.
 * A dangling target (deleted slide / removed section) resolves to `null` so the
 * tile is inert rather than jumping somewhere wrong. The `href` is kept in sync
 * with the resolved slide at BUILD time so present-mode navigation (which reads
 * the href) and the editor semantics (which read `zoomKind`/`zoomTargetId`) agree.
 *
 * DETERMINISM: nothing here calls Date / Math.random / crypto. Ids are supplied
 * by the caller; the pure helpers only read, resolve, and project the data handed
 * to them.
 */

import type { PresentationShapeBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { buildSlideLinkHref } from "@mrnote/core/web/lib/presentation-block-fx.ts";
import { DEFAULT_SHAPE_FILL, DEFAULT_SHAPE_STROKE } from "@mrnote/core/web/lib/presentation-shapes.ts";
import type { SlideSection } from "@mrnote/core/web/lib/presentation-sections.ts";

/** The navigation intent a zoom tile carries. */
export type ZoomKind = "slide" | "section" | "summary";

/** The marker fields a shape block carries to BE a zoom tile. */
export interface ZoomTileMarker {
  /** Which navigation kind this tile is. */
  zoomKind: ZoomKind;
  /**
   * The id of the navigation TARGET: a slide id for `slide` zoom, or a SECTION id
   * for `section` / `summary` tiles. Resolved to a slide index by
   * {@link resolveZoomTargetSlideIndex}.
   */
  zoomTargetId: string;
}

/** A zoom target descriptor (kind + the id of its slide / section target). */
export interface ZoomTarget {
  kind: ZoomKind;
  /** Slide id (slide zoom) or section id (section / summary tile). */
  id: string;
}

/** Minimal slide shape the resolver needs: just an id. */
export interface ZoomResolvableSlide {
  id: string;
}

/**
 * Does this (already-typed `shape`) block carry the zoom-tile markers? A block is
 * a zoom tile iff it has a well-formed `zoomKind` AND a non-empty `zoomTargetId`.
 * Tolerant of unknown values (a hand-edited deck can't crash the renderer): any
 * malformed marker simply makes the block a plain shape.
 */
export function readZoomTileMarker(block: unknown): ZoomTileMarker | null {
  if (!block || typeof block !== "object") return null;
  const record = block as Record<string, unknown>;
  const kind = record.zoomKind;
  if (kind !== "slide" && kind !== "section" && kind !== "summary") return null;
  const targetId = record.zoomTargetId;
  if (typeof targetId !== "string" || !targetId.trim()) return null;
  return { zoomKind: kind, zoomTargetId: targetId };
}

/**
 * Resolve a zoom TARGET to its destination slide INDEX in the live deck, or
 * `null` when the target no longer resolves (a deleted slide / removed section /
 * a section whose anchor slide is gone). PURE — the single source of truth both
 * the editor (build the href) and present mode (where to jump) rely on.
 *
 *  - `slide`:            index of the slide whose id === `target.id`.
 *  - `section`/`summary`: index of the section's FIRST slide. The section's
 *    `startSlideId` anchor is its first member (same rule the rail grouping uses);
 *    we return that anchor slide's index. A `summary` target id is treated as a
 *    section id (each summary TILE points at one section).
 */
export function resolveZoomTargetSlideIndex(
  target: ZoomTarget,
  slides: readonly ZoomResolvableSlide[],
  sections: readonly SlideSection[],
): number | null {
  const id = typeof target?.id === "string" ? target.id.trim() : "";
  if (!id) return null;
  if (target.kind === "slide") {
    const index = slides.findIndex((slide) => slide.id === id);
    return index >= 0 ? index : null;
  }
  // section / summary tile: find the section, then its first (anchor) slide.
  const section = sections.find((entry) => entry.id === id);
  if (!section) return null;
  const index = slides.findIndex((slide) => slide.id === section.startSlideId);
  return index >= 0 ? index : null;
}

/**
 * Resolve a zoom tile's marker → the canonical `slide://<slideId>` href for the
 * slide it should navigate to, or `null` when the target does not resolve. This
 * keeps the tile's persisted href in sync with its `zoomKind`/`zoomTargetId` so
 * present-mode navigation (which reads the href) lands on the same slide the
 * editor's marker describes. Built at insert time and re-buildable after a slide
 * mutation moves the target.
 */
export function zoomTileHref(
  marker: ZoomTileMarker,
  slides: readonly ZoomResolvableSlide[],
  sections: readonly SlideSection[],
): string | null {
  const index = resolveZoomTargetSlideIndex(
    { kind: marker.zoomKind, id: marker.zoomTargetId },
    slides,
    sections,
  );
  if (index === null) return null;
  return buildSlideLinkHref(slides[index].id);
}

/**
 * One zoom tile: a rounded rectangle labelled `label`, placed at `geometry`,
 * that jumps to the slide (or the section's first slide) `targetId` names.
 * Null when the target does not resolve to an existing slide (a section with
 * no slides, a missing slide id), so no inert tile is inserted.
 */
export function zoomTileBlock(
  id: string,
  kind: ZoomKind,
  targetId: string,
  label: string,
  geometry: { left: number; top: number; width: number; height: number },
  slides: readonly ZoomResolvableSlide[],
  sections: readonly SlideSection[],
): PresentationShapeBlock | null {
  const href = zoomTileHref({ zoomKind: kind, zoomTargetId: targetId }, slides, sections);
  if (!href) return null;
  return {
    id,
    type: "shape",
    shape: "rounded-rectangle",
    fill: DEFAULT_SHAPE_FILL,
    stroke: DEFAULT_SHAPE_STROKE,
    zoomKind: kind,
    zoomTargetId: targetId,
    href,
    text: label,
    ...geometry,
  };
}
