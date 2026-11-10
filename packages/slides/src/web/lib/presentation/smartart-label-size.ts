/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Home's font size on SmartArt labels, the way PowerPoint treats it.
 *
 * A graphic's labels shrink together to fit (Microsoft Support, "Learn more
 * about SmartArt graphics"), so a label is DRAWN smaller than the size it
 * stores. Home reads and steps the drawn size, and a size set by hand belongs
 * to that label alone (`smartArtAuthorTextSize`, written by the label's text
 * adapter): it leaves the shared shrink, so the
 * words draw at exactly the size picked (the same size the hover preview
 * showed) and resizing another shape no longer moves them.
 */
import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { OfficeHomeTextCommand } from "@mrnote/core/web/lib/office-home-object-target.ts";
import type { PresentationHomeObjectTargets } from "@mrnote/core/web/lib/presentation-home-object-text.ts";
import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/presentation-smartart.ts";
import {
  smartArtLabelAtDrawnSize,
  smartArtLabelFitScale,
  smartArtLabelSizeIsOwn,
} from "./text-layout-css";

function isFittedSmartArtLabel(block: PresentationBlock): boolean {
  return (
    block.type === "text" &&
    readSmartArtBlockMarker(block)?.role === "label" &&
    !smartArtLabelSizeIsOwn(block)
  );
}

/** A fitted label at the size it is drawn, or the block itself. */
function atDrawnSize(block: PresentationBlock, blocks: readonly PresentationBlock[]): PresentationBlock {
  if (!isFittedSmartArtLabel(block)) return block;
  const scale = smartArtLabelFitScale(block, blocks);
  return scale !== null && scale < 1 ? smartArtLabelAtDrawnSize(block, scale) : block;
}

/** Home's targets with each fitted SmartArt label at its drawn size, for reading the size box. */
export function homeObjectTargetsAtDrawnSize(
  blocks: readonly PresentationBlock[],
  targets: PresentationHomeObjectTargets,
): PresentationHomeObjectTargets {
  if (!targets.text.some((block) => isFittedSmartArtLabel(block as PresentationBlock))) return targets;
  return {
    ...targets,
    text: targets.text.map((block) => atDrawnSize(block as PresentationBlock, blocks)) as typeof targets.text,
  };
}

/**
 * Runs a Home command, sizing SmartArt labels by hand when it sets or steps
 * the size: each targeted fitted label starts from its drawn size and keeps
 * the size written. Other commands pass straight through.
 */
export function homeObjectBlocksWithSmartArtSize(
  blocks: PresentationBlock[],
  targets: PresentationHomeObjectTargets,
  command: OfficeHomeTextCommand,
  write: (blocks: PresentationBlock[]) => PresentationBlock[] | null,
): PresentationBlock[] | null {
  if (command.kind !== "size" && command.kind !== "step") return write(blocks);
  const ids = new Set(targets.text.map((block) => block.id));
  const labelIds = new Set(
    blocks.filter((block) => ids.has(block.id) && isFittedSmartArtLabel(block)).map((block) => block.id),
  );
  if (labelIds.size === 0) return write(blocks);
  const baked = blocks.map((block) => (labelIds.has(block.id) ? atDrawnSize(block, blocks) : block));
  // Picking the size a label already draws at still makes it the label's
  // own, as in PowerPoint (the box takes the value even when nothing moves).
  const written = write(baked) ?? (command.kind === "size" ? baked : null);
  if (!written) return null;
  // The label adapter marks a size it writes; a pick that changed nothing
  // is marked here.
  return written.map((block) =>
    labelIds.has(block.id)
      ? ({ ...block, smartArtAuthorText: true, smartArtAuthorTextSize: true } as PresentationBlock)
      : block,
  );
}
