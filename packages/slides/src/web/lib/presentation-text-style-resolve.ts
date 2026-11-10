/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A slide text block's style as its master and the deck theme resolve it,
 * apart from the text commands in presentation-text-buttons: those serve a
 * spreadsheet's cell and shape text too, and resolving through masters
 * brought the deck's masters, themes and layout catalog with them.
 */
import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { pxToPtRounded } from "@mrnote/core/web/lib/office-units.ts";
import { resolveWithMaster, type SlideMaster } from "@mrnote/core/web/lib/presentation-master.ts";
import {
  emptyPlaceholderFontSize,
  placeholderStageFontSize,
  presentationTextUsesPoints,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import { designPxToPoints } from "@mrnote/core/web/lib/presentation/text-units.ts";

/** PowerPoint's size for words no level of their chain sizes. */
const POWERPOINT_UNSIZED_TEXT_PT = 18;

/**
 * `resolveWithMaster` with its size in the block's own unit. The deck
 * theme's sizes are design px (its master's title and body text); a box that
 * keeps points read them as points (the 32 pt body drew at 43 pt). An
 * imported box that is no placeholder inherits the presentation's default
 * text style, which the import resolves: one that states none draws at
 * PowerPoint's 18 pt, not at the master's body size.
 */
export function resolvePresentationTextStyle(
  ...[slide, block, masters, theme]: Parameters<typeof resolveWithMaster>
): ReturnType<typeof resolveWithMaster> {
  const resolved = resolveWithMaster(slide, block, masters, theme);
  if (!block || block.type !== "text" || !presentationTextUsesPoints(block as PresentationBlock)) return resolved;
  const stated = block.style?.fontSize;
  if (typeof stated === "number" && Number.isFinite(stated)) return resolved;
  if (typeof (block as { placeholder?: unknown }).placeholder === "string") {
    return typeof resolved.fontSize === "number"
      ? { ...resolved, fontSize: pxToPtRounded(resolved.fontSize) }
      : resolved;
  }
  return { ...resolved, fontSize: POWERPOINT_UNSIZED_TEXT_PT };
}

/**
 * The size an object's words draw at when they state none, in points: the
 * master's or the placeholder's for a text block, PowerPoint's 18 pt for a
 * shape.
 */
export function presentationHomeFallbackSizePt(
  slide: { masterId?: unknown; layout?: unknown } | undefined,
  block: PresentationTextBlock | PresentationShapeBlock,
  masters: readonly SlideMaster[] | undefined,
  theme: DeckTheme | undefined,
): number | undefined {
  if (!slide) return undefined;
  if (block.type !== "text") return 18;
  const resolved = resolvePresentationTextStyle(slide, block, masters, theme);
  if (typeof resolved.fontSize === "number" && presentationTextUsesPoints(block)) return resolved.fontSize;
  const px =
    typeof resolved.fontSize === "number"
      ? resolved.fontSize
      : emptyPlaceholderFontSize(block, slide.layout) ?? placeholderStageFontSize(block, slide.layout) ?? 17;
  return designPxToPoints(px);
}
