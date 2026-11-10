/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  presentationWebPictureStyle,
  type PresentationWebPictureStyleId,
} from "./presentation-picture-style-catalog";
import {
  canonicalOfficePictureStyleOutline,
  nativePictureStyle as nativeOfficePictureStyle,
  type OfficePictureStyleOutline,
} from "./office-picture-styles/native";
import {
  normalizeImageBorderColor,
  normalizeImageBorderStyle,
  normalizeImageBorderWidth,
} from "./presentation-image-effects";
import type { PresentationImageBlock } from "../components/console/editor/native-document/documentCommands";

// The recipes and their path geometry are Office's, shared by every editor.
export {
  nativePictureStyle,
  pictureStylePath,
  type NativePictureStyle,
} from "./office-picture-styles/native";








export function presentationPictureStyleOutline(block: {
  readonly pictureStylePreset?: unknown;
  readonly borderColor?: unknown;
  readonly borderWidth?: unknown;
  readonly borderStyle?: unknown;
}): OfficePictureStyleOutline | null {
  const style = nativeOfficePictureStyle(block.pictureStylePreset);
  if (!style) return null;
  if (block.borderStyle === "none")
    return canonicalOfficePictureStyleOutline(style.id, "none");
  const widthPx = normalizeImageBorderWidth(block.borderWidth);
  const color = normalizeImageBorderColor(block.borderColor, null);
  if (!widthPx && !color && block.borderStyle === undefined) return null;
  const dash = normalizeImageBorderStyle(block.borderStyle, "solid");
  return canonicalOfficePictureStyleOutline(style.id, {
    widthEmu: widthPx ? Math.round(widthPx * 9525) : (style.line?.width ?? 9525),
    color: color ?? style.line?.color ?? "#000000",
    dash: dash === "none" ? "solid" : dash,
  });
}

/**
 * The line the Picture Border menu shows as current on a styled picture: the
 * border set over the style, else the style's own line (weight in px). Null
 * for an unstyled picture, whose border fields are its line.
 */
export function presentationPictureStyleBorder(block: {
  readonly pictureStylePreset?: unknown;
  readonly borderColor?: unknown;
  readonly borderWidth?: unknown;
  readonly borderStyle?: unknown;
}): { color: string | null; widthPx: number; style: "none" | "solid" | "dashed" | "dotted" } | null {
  const style = nativeOfficePictureStyle(block.pictureStylePreset);
  if (!style) return null;
  const outline = presentationPictureStyleOutline(block);
  if (outline === "none") return { color: null, widthPx: 0, style: "none" };
  if (outline) return { color: outline.color, widthPx: outline.widthEmu / 9525, style: outline.dash };
  return style.line?.width
    ? { color: style.line.color ?? null, widthPx: style.line.width / 9525, style: "solid" }
    : { color: null, widthPx: 0, style: "none" };
}

export const PICTURE_STYLE_LEGACY_FIELDS = [
  "shadow",
  "cornerRadius",
  "borderColor",
  "borderWidth",
  "borderStyle",
] as const;




export const PICTURE_STYLE_EXCLUSIVE_FIELDS = ["shadow", "cornerRadius"] as const;
export function canApplyNativePictureStyle(
  block: PresentationImageBlock,
): boolean {
  if (block.officeGraphicKind === "svg") return false;
  const source = block.source as Record<string, unknown> | undefined;
  if (!source || (!source.slideId && !source.slidePart)) return true;
  return source.pictureStyleDisposition === "editable";
}
export function applyNativePictureStyle(
  block: PresentationImageBlock,
  id: unknown,
): PresentationImageBlock {
  if (!presentationWebPictureStyle(id) || !canApplyNativePictureStyle(block))
    return block;
  if (
    block.pictureStylePreset === id &&
    !PICTURE_STYLE_LEGACY_FIELDS.some((key) => key in block) &&
    !("pictureEffects" in block)
  )
    return block;
  const next: PresentationImageBlock = {
    ...block,
    pictureStylePreset: id as PresentationWebPictureStyleId,
  };
  for (const key of PICTURE_STYLE_LEGACY_FIELDS) delete next[key];
  // A Picture Style is a whole recipe, effects included: PowerPoint drops the
  // effects picked before it.
  delete next.pictureEffects;
  return next;
}

