/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  layoutBlocksInheritedBySlides,
  masterBlocksInheritedBySlides,
  resolveSlideContext,
  type SlideMaster,
} from "@mrnote/core/web/lib/presentation-master.ts";
import { resolveSlideRenderBlocks } from "../presentation-master-cascade";
import {
  presentationInPaintOrder,
  presentationSlideReordersNativeObjects,
} from "@mrnote/core/web/lib/presentation-paint-order.ts";
import { isBlockNotDrawn } from "@mrnote/core/web/lib/presentation-selection-pane.ts";
import { isImportedInheritedSlideGraphic, isReplacedNativeMasterGraphic } from "@mrnote/core/web/lib/presentation-slide-bg.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import type { SlideModel } from "@mrnote/core/web/lib/presentation/deck-model.ts";

/**
 * The objects a slide draws in the show, in the order it paints them: the
 * master's and the layout's graphics the slide inherits (unless it hides
 * background graphics), then its own, less any object the Selection Pane
 * hides. The static slide and every transition layer read this one list, so
 * a transition draws exactly what the slide does, in the same stacking.
 */
export function presentationSlideRenderBlocks(
  slide: SlideModel,
  masters: readonly SlideMaster[] | undefined,
  theme: DeckTheme | undefined,
): Array<{ block: PresentationBlock; index: number }> {
  const context = resolveSlideContext(slide, masters, theme);
  return presentationInPaintOrder(
    resolveSlideRenderBlocks(
      slide.hideBackgroundGraphics
        ? undefined
        : (masterBlocksInheritedBySlides(context.master) as PresentationBlock[] | undefined),
      slide.blocks,
      slide.hideBackgroundGraphics
        ? undefined
        : (layoutBlocksInheritedBySlides(context.nativeLayout) as PresentationBlock[] | undefined),
    ).filter(
      (block) =>
        !isBlockNotDrawn(block) &&
        !(slide.hideBackgroundGraphics && isImportedInheritedSlideGraphic(block)) &&
        !isReplacedNativeMasterGraphic(block, theme),
    ),
    { reordered: presentationSlideReordersNativeObjects(slide) },
  );
}
