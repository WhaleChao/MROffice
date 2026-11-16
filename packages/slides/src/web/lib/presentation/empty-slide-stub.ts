/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The importer's empty-slide stub: an imported slide with nothing on it is read
 * as one text block of nothing but its id (`<slide id>-block-1`), with no frame
 * and no source. The canvas drew it at a stacked fallback place and the model
 * kept it frameless, so words typed into it were not drawn until the deck was
 * opened again, and the save wrote them where the canvas had not shown them.
 *
 * Its place is that of the words the slide's layout would take: the layout's
 * body / content placeholder (the layout's own placeholder block in the deck's
 * masters, `slide.layoutId`, the frame PowerPoint draws a body in), or, on a
 * layout with none, the place Insert > Text Box gives a new text box. The canvas draws it there from the start, and the
 * first words typed into it state that frame on the block, in the same step:
 * one Undo takes the words and the frame back together.
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { geometryNumber } from "@mrnote/core/web/lib/presentation-block-reads.ts";
import { defaultShapeGeometry, type ShapeGeometry } from "@mrnote/core/web/lib/presentation-shapes.ts";

type Rec = Record<string, unknown>;

const record = (value: unknown): Rec | undefined =>
  value && typeof value === "object" && !Array.isArray(value) ? (value as Rec) : undefined;

/** Placeholder types PowerPoint pours a slide's body text into (`obj` is the default type). */
const BODY_PLACEHOLDER_TYPES = new Set(["body", "obj"]);

/** Whether ``slide`` was read from a file (its source names its package slide). */
function importedSlide(slide: unknown): boolean {
  const source = record(record(slide)?.source);
  return typeof source?.slideId === "string" && source.slideId !== "";
}

/** A text block of an imported slide with no frame and no source of its own. */
export function isFramelessImportedText(slide: unknown, block: PresentationBlock): boolean {
  const value = block as unknown as Rec;
  return (
    importedSlide(slide) &&
    value.type === "text" &&
    value.source === undefined &&
    geometryNumber(block, "left") === null &&
    geometryNumber(block, "top") === null
  );
}

/** A placeholder's frame when it is a body / content placeholder with a whole frame. */
function bodyFrame(kind: unknown, index: unknown, frame: Rec | undefined): ShapeGeometry | null {
  // A placeholder with an index and no type is a content (obj) placeholder.
  const body = typeof kind === "string" ? BODY_PLACEHOLDER_TYPES.has(kind) : index !== undefined;
  const numbers = ["left", "top", "width", "height"].map((key) => frame?.[key]);
  if (!body || !numbers.every((value) => typeof value === "number" && Number.isFinite(value))) return null;
  const [left, top, width, height] = numbers.map((value) => Math.round(value as number));
  return width! > 0 && height! > 0 ? { left: left!, top: top!, width: width!, height: height! } : null;
}

/**
 * The frame of the slide's layout's body placeholder, or the Insert > Text Box
 * default. ``masters``: the deck's masters, whose layouts hold their own
 * placeholders (an empty slide has none of its own to name its layout's by).
 */
export function emptySlideStubFrame(
  slide: unknown,
  bounds: { width: number; height: number },
  masters?: readonly unknown[],
): ShapeGeometry {
  const value = record(slide);
  const layoutId = value?.layoutId;
  if (typeof layoutId === "string" && Array.isArray(masters)) {
    for (const master of masters) {
      const masterValue = record(master);
      if (typeof value?.masterId === "string" && masterValue?.id !== value.masterId) continue;
      const layout = (Array.isArray(masterValue?.layouts) ? masterValue.layouts : [])
        .map(record)
        .find((item) => item?.id === layoutId || item?.sourcePart === layoutId);
      for (const raw of Array.isArray(layout?.blocks) ? layout.blocks : []) {
        const block = record(raw);
        const source = record(block?.source);
        const frame = bodyFrame(
          block?.placeholderType ?? source?.placeholderType,
          block?.placeholderIndex ?? source?.placeholderIndex,
          block,
        );
        if (frame) return frame;
      }
    }
  }
  const cascade = record(record(value?.source)?.layoutCascade);
  for (const raw of Array.isArray(cascade?.placeholders) ? cascade.placeholders : []) {
    const placeholder = record(raw);
    const frame = bodyFrame(placeholder?.placeholderType, placeholder?.placeholderIndex, record(placeholder?.layoutGeometry));
    if (frame) return frame;
  }
  return defaultShapeGeometry("textbox", bounds);
}

/**
 * ``next`` with every imported slide's frameless text block that has just been
 * given words (frameless and empty in ``previous``) stating its frame; the same
 * array when there is none.
 */
export function slidesWithFramedStubWords<S extends { id?: unknown; blocks?: readonly PresentationBlock[] }>(
  next: readonly S[],
  previous: readonly S[],
  boundsOf: (slide: S) => { width: number; height: number },
  masters?: readonly unknown[],
): S[] {
  const before = new Map<unknown, S>(previous.map((slide) => [slide.id, slide]));
  let changed = false;
  const out = next.map((slide) => {
    const prior = before.get(slide.id);
    if (!prior || !Array.isArray(slide.blocks) || !importedSlide(slide)) return slide;
    const priorBlocks = new Map((prior.blocks ?? []).map((block) => [block.id, block]));
    let slideChanged = false;
    const blocks = slide.blocks.map((block) => {
      const was = priorBlocks.get(block.id);
      const words = (block as unknown as Rec).text;
      if (
        !was ||
        typeof words !== "string" ||
        words === "" ||
        !isFramelessImportedText(slide, block) ||
        !isFramelessImportedText(prior, was) ||
        (was as unknown as Rec).text !== ""
      ) {
        return block;
      }
      slideChanged = true;
      return { ...block, ...emptySlideStubFrame(slide, boundsOf(slide), masters) } as PresentationBlock;
    });
    if (!slideChanged) return slide;
    changed = true;
    return { ...slide, blocks };
  });
  return changed ? out : (next as S[]);
}
