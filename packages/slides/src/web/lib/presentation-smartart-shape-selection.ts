/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { isNativeSmartArtMember } from "@mrnote/core/web/lib/presentation-native-smartart.ts";
import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/presentation-smartart.ts";
import { isSmartArtMemberBlock } from "./presentation/selection-geometry";
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";

/**
 * A shape of a SmartArt graphic: a native diagram member, a block of an
 * authored graphic, or a member of the graphic's group.
 */
export function isSmartArtGraphicShape(block: PresentationBlock | null | undefined): boolean {
  if (!block) return false;
  return isNativeSmartArtMember(block) || readSmartArtBlockMarker(block) !== null || isSmartArtMemberBlock(block);
}









export function listLevelGreyedForSmartArtShape(
  block: PresentationBlock | null | undefined,
  wordsSelectedIn: string | null | undefined,
): boolean {
  return isSmartArtGraphicShape(block) && wordsSelectedIn !== block?.id;
}
