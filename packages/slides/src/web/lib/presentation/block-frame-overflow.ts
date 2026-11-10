/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Whether an object's frame cuts its drawing at rest: ONE answer for the
 * editor canvas and for every view drawn from the slide (the slide show, the
 * presenter's slides, Reading View, the thumbnails, the notes page).
 *
 * PowerPoint never cuts an object's paint at its frame: a reflection hangs
 * below the picture, a shadow or a glow around it, an outline straddles the
 * edge, words a text box does not fit run past it (only a stated clip cuts
 * them). The canvas and the show each used to decide this for themselves, and
 * the show cut what the canvas drew: a picture style's reflection and shadow,
 * a text box's WordArt glow and its overflowing words.
 */

import type { PresentationBlock } from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { presentationBlockPaintsOutsideFrame } from "@mrnote/core/web/lib/presentation-shape-style-adapter.ts";
import { asRecord } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import { importedPptxTextOverflow } from "./text-layout-css";

export function presentationBlockFrameOverflow(block: PresentationBlock): "visible" | "hidden" {
  // A table's rules straddle its frame.
  if (block.type === "table") return "visible";
  // A SmartArt drawing member paints its own geometry, which may leave the
  // member's frame.
  if (asRecord(asRecord((block as { source?: unknown }).source)?.nativeSmartArtView)) return "visible";
  if (presentationBlockPaintsOutsideFrame(block)) return "visible";
  // DrawingML's own default for a text body is `overflow`; only a stated clip
  
  if (block.type === "text" && importedPptxTextOverflow(block) === "visible") return "visible";
  return "hidden";
}
