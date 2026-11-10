/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Which unit a presentation block states its text sizes in (text-units.ts).
//
// Points: an imported block (its native source), and any block marked
// `textFontSizeUnit: "point"` -- every text box, WordArt and equation box the

// placeholder (marked on its `source`, as before). Design px: a box the editor
// made before then. The marker sits on the block itself, not on `source`:
// the server's writers read a block with no `source` as the editor's own.
import type { PresentationBlock } from "../../components/console/editor/native-document/documentCommands";
import { isImportedPptxBlock } from "../presentation-block-reads";
import type { PresentationFontSizeUnit } from "./text-units";

/** The marker a block states its sizes in points with. */
export const POINT_TEXT_UNIT = "point" as const;

/** Whether a block states its text sizes in points. */
export function presentationTextUsesPoints(block: PresentationBlock): boolean {
  if ((block as { textFontSizeUnit?: unknown }).textFontSizeUnit === POINT_TEXT_UNIT) return true;
  if (isImportedPptxBlock(block)) return true;
  const source = (block as { source?: unknown }).source;
  return typeof source === "object" && source !== null && (source as Record<string, unknown>).textFontSizeUnit === POINT_TEXT_UNIT;
}

/** The unit a block states its text sizes in. */
export function presentationFontSizeUnit(block: PresentationBlock): PresentationFontSizeUnit {
  return presentationTextUsesPoints(block) ? "point" : "design-px";
}

/** The block marked as stating its sizes in points (sizes are not converted). */
export function withPointTextUnit<T extends object>(block: T): T & { textFontSizeUnit: typeof POINT_TEXT_UNIT } {
  return { ...block, textFontSizeUnit: POINT_TEXT_UNIT };
}
