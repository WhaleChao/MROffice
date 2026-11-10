/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Slides from Outline (Home > New Slide > Slides from Outline): the slides a
 * parsed outline becomes -- the first a Title Slide, the rest Title and
 * Content, each title in its title placeholder and its bullets as first-level
 * bulleted paragraphs, notes kept -- sized as the deck's slides are. The
 * dialog and the AI's command both build them here.
 */

import type { OutlineSlide } from "./presentation-outline";
import { layoutBlocks, normalizeLayoutId, type LayoutFrameSource } from "@mrnote/core/web/lib/presentation-layouts.ts";
import { deckSlideSizeFields } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import type { SlideParagraph } from "@mrnote/core/web/lib/presentation-paragraphs.ts";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";

export function slidesFromOutline(
  outline: readonly OutlineSlide[],
  options: {
    /** The deck's slide size (its first or active slide). */
    sizeSource: { readonly [key: string]: unknown } | undefined;
    bounds: { width?: number | null; height?: number | null };
    frameSource: LayoutFrameSource;
    newId: () => string;
  },
): SlideModel[] {
  return outline.map((slide, index) => {
    const layout = normalizeLayoutId(index === 0 ? "title" : "title-content");
    const blocks = layoutBlocks(layout, options.bounds, options.frameSource).map((block) => {
      if (block.role === "title") return { ...block, text: slide.title };
      const bulletParagraphs: SlideParagraph[] = slide.bullets.map((bullet) => ({
        text: bullet,
        level: 0,
        bullet: "bullet",
      }));
      return {
        ...block,
        text: slide.bullets.join("\n"),
        ...(bulletParagraphs.length > 0 ? { paragraphs: bulletParagraphs } : {}),
      };
    });
    return {
      id: options.newId(),
      title: slide.title,
      blocks: blocks as unknown as PresentationBlock[],
      layout,
      ...deckSlideSizeFields(options.sizeSource),
      ...(slide.notes ? { notes: slide.notes } : {}),
    } as SlideModel;
  });
}
