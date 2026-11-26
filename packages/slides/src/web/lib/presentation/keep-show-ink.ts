/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Ink drawn in a slide show, kept on its slides (GAP 14; PowerPoint's "Keep" when a show with ink
 * ends: "Choose Keep to save the ink-pen annotations on the slides", Microsoft Support "Draw on
 * slides during a presentation").
 *
 * A show's stroke (lib/presentation-annotation `InkStroke`: fractions of the slide, its width a
 * fraction of the slide's smaller side) becomes the slide's own ink object, the `shape: "ink"`
 * block the Draw tab, the Record studio, the ink eraser, Review's Delete All Ink, Ink Replay and
 * Ink to Shape all read -- one model, not a second one. Kept ink goes onto every slide it was drawn
 * on in one change: PowerPoint's one Undo took back all three strokes kept on two slides
 * (office-truth/probes/show-keep-ink, K). PowerPoint keeps each stroke as an Ink object (COM type 23,
 * msoInkComment), written as a p14 `p:contentPart` with an InkML part (`application/inkml+xml`) and a
 * PNG fallback; MRAI's ink blocks still go out as freeform shapes (InkML for slides is its own change).
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  inkShapeOpacityForTool,
  isInkTool,
  type AnnotationPoint,
  type InkStroke,
  type InkTool,
} from "@mrnote/core/web/lib/presentation-annotation.ts";
import { inkStrokeToShape } from "@mrnote/core/web/lib/presentation-freeform.ts";
import { MAX_STROKE_WIDTH } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { AUTOFIT_EMU_PER_PX } from "./text-layout-css";
import { MAX_PRESENTATION_BLOCKS_PER_SLIDE } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/**
 * One stroke drawn over a slide (`points` fractions of a slide `bounds` EMU wide and high; `width`
 * a fraction of its smaller side) as the slide's ink block, its width in the slide's design pixels
 * as every slide ink stores it; null for a stroke with nothing to draw.
 */
export function annotationStrokeToInkBlock(
  stroke: { readonly tool: InkTool; readonly color: string; readonly width: number; readonly points: readonly AnnotationPoint[] },
  bounds: { readonly width: number; readonly height: number },
  id: string,
): PresentationBlock | null {
  const minimumDimension = Math.min(bounds.width, bounds.height);
  const strokeWidth = Math.min(
    MAX_STROKE_WIDTH,
    Math.max(1, Math.round(((stroke.width * minimumDimension) / AUTOFIT_EMU_PER_PX) * 100) / 100),
  );
  const shape = inkStrokeToShape(
    stroke.points.map((point) => ({ x: point.x * bounds.width, y: point.y * bounds.height })),
    strokeWidth * AUTOFIT_EMU_PER_PX,
  );
  if (!shape) return null;
  const opacity = inkShapeOpacityForTool(stroke.tool);
  return {
    id,
    type: "shape",
    shape: "ink",
    points: shape.points,
    stroke: stroke.color,
    strokeWidth,
    inkTool: stroke.tool,
    ...(opacity !== undefined ? { opacity } : {}),
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  } as unknown as PresentationBlock;
}

/** Whether a show left ink worth asking about: a pen or highlighter stroke on some slide. */
export function showHasKeepableInk(inkBySlide: Readonly<Record<string, readonly InkStroke[]>>): boolean {
  return Object.values(inkBySlide).some((strokes) => strokes.some((stroke) => isInkTool(stroke.tool)));
}

/**
 * The deck with each slide's show ink (`inkBySlide`, by slide id) added to it as ink blocks, in the
 * order drawn, as far as a slide's block budget goes; null when nothing would be added.
 */
export function slidesWithKeptShowInk<S extends { readonly id: string; readonly blocks: PresentationBlock[] }>(
  slides: readonly S[],
  inkBySlide: Readonly<Record<string, readonly InkStroke[]>>,
  boundsOf: (slide: S) => { readonly width: number; readonly height: number },
  makeId: () => string,
): S[] | null {
  let changed = false;
  const next = slides.map((slide) => {
    const strokes = inkBySlide[slide.id];
    if (!strokes || strokes.length === 0) return slide;
    const bounds = boundsOf(slide);
    const added: PresentationBlock[] = [];
    // Highlighter strokes go under the pen's (truth show-keep-ink K: slide 2's highlighter, drawn after
    // its pen stroke, was kept beneath it), each kind in the order drawn.
    const ordered = [
      ...strokes.filter((stroke) => stroke.tool === "highlighter"),
      ...strokes.filter((stroke) => stroke.tool !== "highlighter"),
    ];
    for (const stroke of ordered) {
      if (!isInkTool(stroke.tool)) continue;
      if (slide.blocks.length + added.length >= MAX_PRESENTATION_BLOCKS_PER_SLIDE) break;
      const block = annotationStrokeToInkBlock(stroke as InkStroke & { tool: InkTool }, bounds, makeId());
      if (block) added.push(block);
    }
    if (added.length === 0) return slide;
    changed = true;
    return { ...slide, blocks: [...slide.blocks, ...added] };
  });
  return changed ? next : null;
}
