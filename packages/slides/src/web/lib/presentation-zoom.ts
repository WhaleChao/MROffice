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
import { normalizeSections, type SlideSection } from "@mrnote/core/web/lib/presentation-sections.ts";

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













export const SLIDE_ZOOM_PLACEHOLDER_FILL = "#d9d9d9";
/** `prstClr val="ltGray"` (Office's preset light grey). */
export const SLIDE_ZOOM_OUTLINE = "#d3d3d3";
/** 3175 EMU (1/4 pt) in the deck's px (9525 EMU). */
export const SLIDE_ZOOM_OUTLINE_WIDTH_PX = 3175 / 9525;

export function slideZoomGeometry(slide: { width: number; height: number }): {
  left: number;
  top: number;
  width: number;
  height: number;
} {
  const width = Math.round(slide.width / 4);
  const height = Math.round(slide.height / 4);
  return { left: Math.round((slide.width - width) / 2), top: Math.round((slide.height - height) / 2), width, height };
}

/** The most zooms one Insert Slide / Section Zoom adds (a deck's slide budget bounds them too). */
export const MAX_ZOOMS_PER_INSERT = 100;

/** 12 pt in the deck's units (EMU): the step between zooms inserted together. */
export const ZOOM_CASCADE_STEP_EMU = 152_400;








export function zoomCascadeGeometries(
  slide: { width: number; height: number },
  count: number,
): { left: number; top: number; width: number; height: number }[] {
  const base = slideZoomGeometry(slide);
  return Array.from({ length: Math.max(0, count) }, (_, index) => {
    const offset = Math.round((index - (count - 1) / 2) * ZOOM_CASCADE_STEP_EMU);
    return { ...base, left: base.left + offset, top: base.top + offset };
  });
}

/**
 * One zoom tile: PowerPoint's Slide Zoom or Section Zoom (above, no words:
 * it shows the picture of its target slide, or of its section's first slide;
 * both package writers take either, `slide_zoom_write`); a SUMMARY tile keeps
 * the rounded rectangle labelled `label`. Placed at `geometry`, it jumps to
 * the slide (or the section's first slide) `targetId` names. Null when the
 * target does not resolve to an existing slide (a section with no slides, a
 * missing slide id), so no inert tile is inserted.
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
  if (kind === "slide" || kind === "section") {
    return {
      id,
      type: "shape",
      shape: "rectangle",
      fill: SLIDE_ZOOM_PLACEHOLDER_FILL,
      stroke: SLIDE_ZOOM_OUTLINE,
      strokeWidth: SLIDE_ZOOM_OUTLINE_WIDTH_PX,
      zoomKind: kind,
      zoomTargetId: targetId,
      href,
      ...geometry,
    };
  }
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

/**
 * The tiles of a Summary Zoom inside its frame (the summary slide's body
 * placeholder), as PowerPoint lays them out (office-truth probes zoom-insert
 * and zoom-insert2, n = 1 to 10): `ceil(sqrt(n))` columns and as many rows as
 * the tiles need; each tile 90 % of its cell, at the slide's aspect ratio,
 * whichever of the cell's width or height binds; the tiles 3.75 % of a tile's
 * width apart; the grid centred in the frame, its last row from the left.
 * E.g. three tiles in 457200, 1600200, 8229600 x 4525963 EMU: 3620769 x
 * 2036683 at frame offsets (426141, 158409), (4182689, 158409), (426141,
 * 2330871).
 */
export function summaryZoomTileGeometries(
  frame: { left: number; top: number; width: number; height: number },
  count: number,
  slideAspect: number,
): { left: number; top: number; width: number; height: number }[] {
  if (count <= 0 || !(slideAspect > 0) || !(frame.width > 0) || !(frame.height > 0)) return [];
  const cols = Math.ceil(Math.sqrt(count));
  const rows = Math.ceil(count / cols);
  // In whole EMU, as PowerPoint writes them: the binding side first, the
  // other from it, the gap rounded, the grid's offset floored.
  const cellWidth = frame.width / cols;
  const cellHeight = frame.height / rows;
  let width: number;
  let height: number;
  if (cellWidth <= cellHeight * slideAspect) {
    width = Math.floor(0.9 * cellWidth);
    height = Math.floor(width / slideAspect);
  } else {
    height = Math.floor(0.9 * cellHeight);
    width = Math.floor(height * slideAspect);
  }
  const gap = Math.round(0.0375 * width);
  const x0 = Math.floor((frame.width - (cols * width + (cols - 1) * gap)) / 2);
  const y0 = Math.floor((frame.height - (rows * height + (rows - 1) * gap)) / 2);
  return Array.from({ length: count }, (_, index) => ({
    left: Math.round(frame.left) + x0 + (index % cols) * (width + gap),
    top: Math.round(frame.top) + y0 + Math.floor(index / cols) * (height + gap),
    width,
    height,
  }));
}

/** What Insert Summary Zoom does to a deck's slides and sections (`summaryZoomPlan`). */
export interface SummaryZoomPlan {
  /** Where the summary slide goes: just before the first slide ticked. */
  readonly insertIndex: number;
  /** The slide ids in their new order, the summary slide's among them. */
  readonly slideIds: string[];
  /** The deck's sections after the insert. */
  readonly sections: SlideSection[];
  /** One per slide ticked, in the deck's order: the section its tile zooms to. */
  readonly tileSectionIds: string[];
}

/**
 * Insert > Zoom > Summary Zoom, as PowerPoint does it (office-truth probe
 * zoom-insert2, Microsoft Support "Use Zoom for PowerPoint"): the summary
 * slide goes just before the first slide ticked, in a section of its own
 * named "Summary Section"; every slide ticked that does not begin a section
 * begins one ("Section 1", "Section 2", ... in order), splitting the section
 * it was in; each tile zooms to the section its slide begins. With "Keep
 * unused sections" off, the sections after the summary slide that no tile
 * zooms to are removed -- their slides join the section before; no slide is
 * deleted. A deck with no sections whose summary does not stand first gets
 * the Default Section for the slides before it, as adding a section does.
 * Null when no slide ticked is in the deck.
 */
export function summaryZoomPlan(input: {
  readonly slideIds: readonly string[];
  readonly sections: readonly SlideSection[];
  readonly selectedSlideIds: ReadonlySet<string> | readonly string[];
  readonly keepUnusedSections: boolean;
  readonly summarySlideId: string;
  readonly newSectionId: () => string;
  readonly names: {
    readonly summarySection: string;
    readonly defaultSection: string;
    readonly newSection: (ordinal: number) => string;
  };
}): SummaryZoomPlan | null {
  const selected = new Set(input.selectedSlideIds);
  const ticked = input.slideIds.filter((id) => selected.has(id));
  if (ticked.length === 0 || input.slideIds.includes(input.summarySlideId)) return null;
  const insertIndex = input.slideIds.indexOf(ticked[0]);
  const slideIds = [
    ...input.slideIds.slice(0, insertIndex),
    input.summarySlideId,
    ...input.slideIds.slice(insertIndex),
  ];
  let sections: SlideSection[] = normalizeSections(input.sections, input.slideIds).map((section) => ({ ...section }));
  if (sections.length === 0 && insertIndex > 0) {
    sections.push({ id: input.newSectionId(), name: input.names.defaultSection, startSlideId: input.slideIds[0] });
  }
  sections.push({ id: input.newSectionId(), name: input.names.summarySection, startSlideId: input.summarySlideId });
  const tileSectionIds: string[] = [];
  let ordinal = 0;
  for (const slideId of ticked) {
    let section = sections.find((entry) => entry.startSlideId === slideId);
    if (!section) {
      ordinal += 1;
      section = { id: input.newSectionId(), name: input.names.newSection(ordinal), startSlideId: slideId };
      sections.push(section);
    }
    tileSectionIds.push(section.id);
  }
  if (!input.keepUnusedSections) {
    const zoomed = new Set(tileSectionIds);
    const after = (section: SlideSection) => slideIds.indexOf(section.startSlideId) > insertIndex;
    sections = sections.filter((section) => zoomed.has(section.id) || !after(section));
  }
  return { insertIndex, slideIds, sections: normalizeSections(sections, slideIds), tileSectionIds };
}
