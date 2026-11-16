/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether an ordinary slide shape is drawn with No Fill (`a:noFill`, Shape
 * Fill > No Fill): its geometry paints nothing inside the outline.
 *
 * PowerPoint selects such a shape only by its outline or its words; a click
 * on its empty interior passes through to whatever lies below (a dashed
 * frame drawn over a chart leaves the chart's labels clickable). The edit
 * canvas hit-tests these shapes by geometry for that reason
 * (`presentationShapeGeometryHitClassName`).
 *
 * The paint is read as the renderer reads it (PresentationShapeGraphic): the
 * block's own fill, else the package's theme paint when the block states no
 * fill of its own. A line, connector, ink stroke or open shape has no
 * interior (`resolveShapeStyle` gives it no fill) and is not one; neither is
 * a picture, pattern or gradient fill, a camera (cameo) shape, or a text box
 * (`type: "text"`), which PowerPoint selects by a click anywhere inside it.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { isPresentationCameoBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { readPresentationNativeThemePaint } from "@mrnote/core/web/lib/presentation-native-theme-paint.ts";
import { resolveShapeStyle } from "@mrnote/core/web/lib/presentation-shapes.ts";

export function presentationShapeHasNoFill(block: PresentationBlock): boolean {
  const style = paintedStyleOf(block);
  return style !== null && style.fill === "none" && style.fillTo === null && !style.fillGradient && !style.fillImage && !style.fillPattern;
}

/** Lines, connectors and ink are picked by their stroke already (their own
 * hit paths, or a box no wider than the line). */
const STROKE_PICKED_KINDS = /(^|-)line$|connector|^ink$/;
















export function presentationShapeHasNoInterior(block: PresentationBlock): boolean {
  const style = paintedStyleOf(block);
  if (style === null || style.fill !== null || style.fillTo !== null || style.fillGradient || style.fillImage || style.fillPattern) return false;
  return !STROKE_PICKED_KINDS.test(String((block as unknown as { shape?: unknown }).shape ?? ""));
}

function paintedStyleOf(block: PresentationBlock): ReturnType<typeof resolveShapeStyle> | null {
  if (block.type !== "shape" || isPresentationCameoBlock(block)) return null;
  const raw = block as unknown as Record<string, unknown>;
  const ownsFill =
    raw.fill !== undefined || raw.fillImage !== undefined || raw.fillPattern !== undefined || raw.fillGradient !== undefined;
  const themePaint = ownsFill ? undefined : readPresentationNativeThemePaint(block);
  return resolveShapeStyle({
    shape: raw.shape,
    fill: ownsFill ? raw.fill : themePaint?.fill,
    fillTo: raw.fillTo ?? themePaint?.fillTo,
    fillGradient: raw.fillGradient ?? themePaint?.fillGradient,
    fillImage: raw.fillImage,
    fillPattern: raw.fillPattern,
    points: raw.points,
    customPath: raw.customPath,
    openPathFilled: raw.openPathFilled,
  });
}
