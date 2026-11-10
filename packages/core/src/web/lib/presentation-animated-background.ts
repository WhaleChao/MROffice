/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Animate as Background (PPT-A134).
 *
 * Microsoft Support ("Animate the background of your slides in PowerPoint
 * for Mac"): the user clicks the parts of a picture that fills the slide;
 * in the show the background zooms and pans between them while the rest of
 * the slide stays still. A picture carries the points (`animatedBackground`,
 * picture fractions); the server writes PowerPoint for Mac's own timing for
 * them and reads it back.
 *
 * The model: the field, PowerPoint's pan rule and who may carry it; the
 * show's loop is `presentation-animated-background-play`.
 */
import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";

export interface PresentationFocalPoint {
  readonly x: number;
  readonly y: number;
}

export interface PresentationAnimatedBackground {
  readonly focalPoints: readonly PresentationFocalPoint[];
}

export const PRESENTATION_ANIMATED_BACKGROUND_MAX_POINTS = 20;
/** One pixel at 96 per inch: the margin PowerPoint's pan leaves at the edge. */
const EDGE_EMU = 9525;

/** The field, validated; undefined for anything else. */
export function normalizePresentationAnimatedBackground(
  value: unknown,
): PresentationAnimatedBackground | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const points = (value as { focalPoints?: unknown }).focalPoints;
  if (!Array.isArray(points) || points.length < 1 || points.length > PRESENTATION_ANIMATED_BACKGROUND_MAX_POINTS)
    return undefined;
  const focalPoints: PresentationFocalPoint[] = [];
  for (const point of points) {
    const x = (point as { x?: unknown } | null)?.x;
    const y = (point as { y?: unknown } | null)?.y;
    if (
      typeof x !== "number" || typeof y !== "number" ||
      !Number.isFinite(x) || !Number.isFinite(y) ||
      x < 0 || x > 1 || y < 0 || y > 1
    )
      return undefined;
    focalPoints.push({ x, y });
  }
  return { focalPoints };
}

/** How far the zoomed picture may pan before its edge would show. */
function panBound(slideEmu: number): number {
  return 0.25 - EDGE_EMU / slideEmu;
}

/** The pan bringing picture fraction `focal` to the slide's centre: PowerPoint's. */
export function presentationAnimatedBackgroundPan(focal: number, slideEmu: number): number {
  const bound = panBound(slideEmu);
  return Math.round(Math.max(-bound, Math.min(bound, 1.5 * (0.5 - focal))) * 1e5) / 1e5;
}

/**
 * Whether a picture may animate as the background, or why not (an i18n key).
 * Microsoft's article: a picture resized to fill the slide. One picture,
 * unrotated, covering the slide (to 1 px).
 */
export function presentationAnimatedBackgroundRefusal(
  block: PresentationBlock | null | undefined,
  slide: { width: number; height: number },
): string | null {
  if (!block || block.type !== "image") return "presentation.animatedBackground.fillSlide";
  const record = block as Record<string, unknown>;
  const left = Number(record.left);
  const top = Number(record.top);
  const width = Number(record.width);
  const height = Number(record.height);
  const rotation = Number(record.rotation ?? 0);
  if (![left, top, width, height].every(Number.isFinite)) return "presentation.animatedBackground.fillSlide";
  const covers =
    left <= EDGE_EMU &&
    top <= EDGE_EMU &&
    left + width >= slide.width - EDGE_EMU &&
    top + height >= slide.height - EDGE_EMU;
  if (!covers || (Number.isFinite(rotation) && Math.abs(rotation % 360) > 0.01))
    return "presentation.animatedBackground.fillSlide";
  return null;
}

/**
 * Picture Format > Animate as Background on one slide's blocks: the picture
 * gains the background (or loses it, for null) and every other picture on the
 * slide loses its own, as PowerPoint keeps one per slide. Null when nothing
 * changes. `canEdit` is the host's lock rule for the picture.
 */
export function blocksWithAnimatedBackground(
  blocks: readonly PresentationBlock[],
  blockId: string,
  background: PresentationAnimatedBackground | null,
  canEdit: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  let changed = false;
  const next = blocks.map((block) => {
    if (block.id === blockId && block.type === "image" && canEdit(block)) {
      const written = { ...block } as Record<string, unknown>;
      if (background) written.animatedBackground = background;
      else delete written.animatedBackground;
      changed = true;
      return written as PresentationBlock;
    }
    // One background per slide, as PowerPoint keeps it.
    if (background && block.type === "image" && "animatedBackground" in block) {
      changed = true;
      const { animatedBackground: _previous, ...rest } = block as Record<string, unknown>;
      void _previous;
      return rest as PresentationBlock;
    }
    return block;
  });
  return changed ? next : null;
}
