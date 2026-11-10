/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { officeSmartArtLayout } from "@mrnote/core/web/lib/office-smartart/catalog.ts";
import { readSmartArtBlockMarker, type SmartArtKind } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { smartArtOutlineFromTextBlock } from "@mrnote/core/web/lib/presentation-conveniences.ts";

export interface PowerPointConvertToSmartArtLayout {
  /** Office's layout id (the last segment of its `uniqueId` urn). */
  readonly office: string;
  /** The catalog layout MRAI builds for it, or null when it cannot yet. */
  readonly layoutId: string | null;
}

export const POWERPOINT_CONVERT_TO_SMARTART_LAYOUTS: readonly PowerPointConvertToSmartArtLayout[] =
  Object.freeze(
    (
      [
        ["vList2", "vertical-bullet-list"],
        ["vList5", "vertical-block-list"],
        ["vList3", null],
        ["target3", null],
        ["hList1", "horizontal-bullet-list"],
        ["orgChart1", "org-chart"],
        ["hierarchy3", null],
        ["hList7", null],
        ["radial2", null],
        ["cycle2", "basic-cycle"],
        ["chevron2", null],
        ["lProcess3", null],
        ["hProcess9", null],
        ["hProcess11", "timeline"],
        ["process1", "basic-process"],
        ["bList2", null],
        ["pyramid2", "pyramid-list"],
        ["matrix3", "basic-matrix"],
        ["venn1", "venn"],
        ["target1", "target"],
      ] as const
    ).map(([office, layoutId]) => ({ office, layoutId })),
  );

/**
 * Whether Convert to SmartArt can act on `block`: a text box, or a shape or
 * placeholder holding text (PowerPoint offers Home > Convert to SmartArt for
 * the body placeholder a layout gave the slide, the usual place a slide's
 * points are; an imported deck's content placeholder is a `shape` here,
 * AI70H R112c B12), with at least one non-blank line, which becomes the
 * diagram's first item. Never a SmartArt's own member, nor a shape that
 * belongs to a group or a native diagram's drawing.
 */
export function presentationBlockConvertsToSmartArt(
  block: PresentationBlock | null | undefined,
): boolean {
  if (!block || (block.type !== "text" && block.type !== "shape")) return false;
  if (readSmartArtBlockMarker(block) !== null) return false;
  if (block.type === "shape") {
    // A shape that is part of something else (a group, a native diagram's
    // drawing) is not the selection PowerPoint converts.
    const record = block as Record<string, unknown>;
    if (typeof record.groupId === "string" && record.groupId.length > 0) return false;
    if (Object.keys(record).some((key) => key.startsWith("nativeSmartArt"))) return false;
  }
  return smartArtOutlineFromTextBlock(block).labels.length > 0;
}

/** The engine kind a catalog layout id builds, or null for an unknown id. */
export function convertToSmartArtKind(layoutId: string): SmartArtKind | null {
  return officeSmartArtLayout(layoutId)?.kind ?? null;
}
