/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export type PresentationChartFrameOperation = "move" | "resize" | "rotate";

type FrameBlock = { type?: unknown; source?: unknown };

/** A lock check is not native frame authority. These imported ChartEx blocks
 * have only an enclosing slide identity and are patchable while unchanged.
 * Complete native frame identities and scratch charts keep their existing path;
 * this predicate grants no chart data/type/style writer permission. */
export function canChangePresentationChartFrame(
  block: FrameBlock,
  operation: PresentationChartFrameOperation,
): boolean {
  if (!["move", "resize", "rotate"].includes(operation)) return false;
  // A chart is a graphic frame, and PowerPoint turns no graphic frame: its
  // selection wears no rotation handle and Size > Rotation is greyed (as a
  
  if (operation === "rotate" && block.type === "chart") return false;
  const source =
    block.source &&
    typeof block.source === "object" &&
    !Array.isArray(block.source)
      ? (block.source as Record<string, unknown>)
      : undefined;
  if (
    block.type !== "chart" ||
    source?.nativeChartDisposition !== "preserve-only"
  )
    return true;
  const slideId = [source.slideId, source.slide_id].find(
    (value) => typeof value === "string" && value.trim(),
  );
  return typeof slideId === "string" && /^sld[1-9][0-9]*$/.test(slideId.trim());
}
