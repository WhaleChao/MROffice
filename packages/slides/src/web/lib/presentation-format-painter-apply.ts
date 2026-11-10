/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Painter, painting a WHOLE target block: the plan of
 * `planFormatPaint` applied onto one slide object, style only. The window's
 * click-to-paint uses it for every target but a text-selection paint (which
 * reads the live caret), and so does the AI's presentation.object.format_painter
 * command: one function, so both paint the same.
 *
 * STYLE-ONLY: the target is spread first and only style props are layered on;
 * geometry / id / type / text / grouping / anchors / media / href pass through.
 */

import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
  PresentationTextStyle,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { tableWithPaintedLook, type FormatPaintPlan } from "@mrnote/core/web/lib/presentation-format-painter.ts";
import { applyRunFormat, normalizeRuns, runsFromText, textFromRuns, type RunStyle } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import { presentationBlockWithFontColor } from "./presentation-run-text-paint";
import { shapeBlockWithPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import {
  normalizeParagraphsWithRuns,
  ownsParagraphRichLayout,
  type SlideParagraphWithRuns,
} from "@mrnote/core/web/lib/presentation-text-buttons.ts";

/** The run patch of a plan with its cleared fields (null) as undefined. */
export function formatPaintRunStylePatch(plan: FormatPaintPlan): Partial<RunStyle> | undefined {
  return plan.runStylePatch
    ? (Object.fromEntries(
        Object.entries(plan.runStylePatch).map(([key, value]) => [key, value === null ? undefined : value]),
      ) as Partial<RunStyle>)
    : undefined;
}

/** The block text style and a uniform run style over the whole text. */
function paintWholeText(textBlock: PresentationTextBlock, plan: FormatPaintPlan): PresentationTextBlock {
  let merged: PresentationTextBlock = textBlock;
  if (plan.textPatch) {
    const nextStyle = { ...(textBlock.style || {}), ...plan.textPatch };
    for (const [k, v] of Object.entries(nextStyle)) {
      if (v === false || v === undefined || v === null || v === "") {
        delete nextStyle[k as keyof PresentationTextStyle];
      }
    }
    merged = { ...merged, style: nextStyle };
  }
  const patch = formatPaintRunStylePatch(plan);
  if (!patch) return merged;
  // Apply as a uniform whole-block run style, mirroring the editor's
  // no-sub-selection run-format path (covers flat + list paragraphs).
  const paragraphs = normalizeParagraphsWithRuns(merged.paragraphs);
  if (paragraphs && ownsParagraphRichLayout(paragraphs)) {
    const nextParagraphs = paragraphs.map((paragraph) => {
      const baseRuns = paragraph.runs?.length ? [...paragraph.runs] : runsFromText(paragraph.text);
      const normalized = normalizeRuns(applyRunFormat(baseRuns, 0, paragraph.text.length, patch));
      const nextParagraph: SlideParagraphWithRuns = { ...paragraph };
      if (normalized) nextParagraph.runs = normalized;
      return nextParagraph;
    });
    return { ...merged, paragraphs: nextParagraphs };
  }
  const baseRuns = normalizeRuns(merged.runs) ?? runsFromText(merged.text);
  const normalized = normalizeRuns(applyRunFormat(baseRuns, 0, merged.text.length, patch));
  const withRuns: PresentationTextBlock = {
    ...merged,
    text: normalized ? textFromRuns(normalized) : merged.text,
  };
  if (normalized) withRuns.runs = normalized;
  else delete (withRuns as { runs?: unknown }).runs;
  return withRuns;
}

/**
 * `block` with `plan` painted onto the whole of it. Families the block's kind
 * cannot take are skipped (the plan already holds only applicable ones). The
 * block itself when nothing changes.
 */
export function paintPresentationBlockFormat(block: PresentationBlock, plan: FormatPaintPlan): PresentationBlock {
  let next: PresentationBlock = block;
  // Shape visual style (fill / gradient / outline / shadow).
  if (plan.shapePatch && block.type === "shape") {
    next = shapeBlockWithPatch(
      block as PresentationShapeBlock,
      plan.shapePatch as Partial<PresentationShapeBlock>,
    ) as unknown as PresentationBlock;
  }
  // Image visual style (shadow / border / corner radius).
  if (plan.imagePatch && block.type === "image") {
    next = { ...next, ...plan.imagePatch };
  }
  // Table look: style and every Table Style Option (PPT-B154).
  if (plan.tablePatch && block.type === "table") {
    next = tableWithPaintedLook(next as unknown as Record<string, unknown>, plan.tablePatch) as unknown as PresentationBlock;
  }
  // Block-level text style and the uniform run style; a shape's own words
  // take the same whole-text formatting.
  if (
    (plan.textPatch || plan.runStylePatch) &&
    (block.type === "text" || (block.type === "shape" && typeof block.text === "string"))
  ) {
    next = paintWholeText(next as unknown as PresentationTextBlock, plan) as unknown as PresentationBlock;
    // The painted colour on WordArt words is their fill: the paint the
    // target wears drew them in its old colour beside the painted one.
    const runPatch = formatPaintRunStylePatch(plan);
    if (runPatch && Object.hasOwn(runPatch, "color")) {
      next = presentationBlockWithFontColor(next, runPatch.color || undefined);
    }
  }
  return next;
}
