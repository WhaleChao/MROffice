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

import { withCommandStatedKeys } from "@mrnote/core/web/lib/presentation-command-stated-keys.ts";
import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
  PresentationTextStyle,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  FORMAT_PAINTER_BODY_BLOCK_KEYS,
  FORMAT_PAINTER_BODY_STYLE_KEYS,
  tableWithPaintedLook,
  type FormatPainterBodySnapshot,
  type FormatPaintPlan,
} from "@mrnote/core/web/lib/presentation-format-painter.ts";
import { applyRunFormat, normalizeRuns, runsFromText, textFromRuns, type RunStyle } from "@mrnote/core/web/lib/presentation-point-runs.ts";
import { pictureBorderIsOpaque } from "./presentation-picture-effects";
import { presentationBlockWithFontColor } from "@mrnote/core/web/lib/presentation-run-text-paint.ts";
import { shapeBlockWithPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import {
  flatRunsFromParagraphs,
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

/**
 * `block` with its first list level's entry (`source.nativeLevelStyles[0]`)
 * as a text box painted whole leaves it (`FormatPaintPlan.levelOnePatch`):
 * what its words, which now leave the painted keys to their level, draw.
 */
function withPaintedLevelOne(block: PresentationTextBlock, plan: FormatPaintPlan): PresentationTextBlock {
  const patch = plan.levelOnePatch;
  if (!patch) return block;
  const source = (block as { source?: unknown }).source;
  const record = source && typeof source === "object" ? (source as Record<string, unknown>) : {};
  const ladder = Array.isArray(record.nativeLevelStyles) ? [...(record.nativeLevelStyles as unknown[])] : [];
  const first: Record<string, unknown> = { ...((ladder[0] && typeof ladder[0] === "object" ? ladder[0] : {}) as object) };
  for (const [key, value] of Object.entries(patch)) {
    if (value === null || value === undefined || value === false) delete first[key];
    else first[key] = value;
  }
  if (typeof first.color === "string") first.colorFromList = true;
  else delete first.colorFromList;
  if (typeof first.fontFamily !== "string") delete first.themeFontRole;
  ladder[0] = first;
  return { ...block, source: { ...record, nativeLevelStyles: ladder } } as PresentationTextBlock;
}

/** The block text style and a uniform run style over the whole text. */
function paintWholeText(textBlock: PresentationTextBlock, plan: FormatPaintPlan): PresentationTextBlock {
  let merged: PresentationTextBlock = withPaintedLevelOne(textBlock, plan);
  if (plan.textPatch) {
    // A painted box's format is its first level's (P2), never a whole-object
    // command's: the keys it paints are no longer named as one's.
    const nextStyle = {
      ...withCommandStatedKeys((textBlock.style || {}) as Record<string, unknown>, { remove: Object.keys(plan.textPatch) }),
      ...plan.textPatch,
    } as PresentationTextStyle;
    // Only what the paint states: a setting of the target's own that is
    // false (`spaceFirstLastPara`) is no format the painter carries, and taken
    
    for (const [k, v] of Object.entries(plan.textPatch)) {
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
    // The box's flat runs say the same words in one list (an imported
    // shape states both): they take the paint too, or they keep the old
    
    // 10.5 pt green after its paragraphs took 20 pt bold).
    if (!Array.isArray((merged as { runs?: unknown }).runs)) return { ...merged, paragraphs: nextParagraphs };
    const flat = flatRunsFromParagraphs(nextParagraphs);
    const painted: PresentationTextBlock = { ...merged, paragraphs: nextParagraphs };
    if (flat) painted.runs = flat;
    else delete (painted as { runs?: unknown }).runs;
    return painted;
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
    const patch: Record<string, unknown> = { ...plan.imagePatch };
    // A line the file states in a form the page cannot stays the file's.
    if (pictureBorderIsOpaque(block)) for (const key of ["borderColor", "borderWidth", "borderStyle", "borderCompound", "borderJoin"]) delete patch[key];
    next = { ...next, ...patch } as PresentationBlock;
    // A painted frame is the source's whole line: one line with round
    // corners where the source states no compound or join.
    if (patch.borderStyle !== undefined) {
      for (const key of ["borderCompound", "borderJoin"] as const) {
        if (!(key in patch) && key in next) {
          const rest = { ...(next as unknown as Record<string, unknown>) };
          delete rest[key];
          next = rest as unknown as PresentationBlock;
        }
      }
    }
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
  // The source's whole text frame (`a:bodyPr`), the target's own taken off
  // first: what the source does not state the target no longer states either
  // (truth c62bodypr B3: its vertical text and three columns went).
  if (plan.bodyPatch && (block.type === "text" || (block.type === "shape" && typeof block.text === "string"))) {
    next = withPaintedBody(next, plan.bodyPatch);
  }
  return next;
}

/**
 * ``block`` wearing ``body``: its frame fields replaced by the source's, and,
 * for an imported target painted from an imported shape, the receipt the save
 * copies that shape's whole `a:bodyPr` from (`source.paintedBodyFrom`, readings
 * row 829): what the page does not model -- the frame's rotation, column gap
 * and order, overflow -- comes with it, checked against the original. A later
 * change to the target's frame stays its own: the save lays it over the copy.
 */
export function withPaintedBody(block: PresentationBlock, body: FormatPainterBodySnapshot): PresentationBlock {
  const next = { ...(block as unknown as Record<string, unknown>) };
  const style = {
    ...((next.style && typeof next.style === "object" && !Array.isArray(next.style) ? next.style : {}) as Record<string, unknown>),
  };
  for (const key of FORMAT_PAINTER_BODY_BLOCK_KEYS) delete next[key];
  for (const key of FORMAT_PAINTER_BODY_STYLE_KEYS) {
    delete style[key];
    // A read page states a text box's anchor, wrap and overflow on the block.
    if (!(FORMAT_PAINTER_BODY_BLOCK_KEYS as readonly string[]).includes(key)) delete next[key];
  }
  for (const [key, value] of Object.entries(body.block)) next[key] = JSON.parse(JSON.stringify(value));
  for (const [key, value] of Object.entries(body.style)) style[key] = JSON.parse(JSON.stringify(value));
  next.style = style;
  const source = next.source && typeof next.source === "object" ? { ...(next.source as Record<string, unknown>) } : null;
  if (source) {
    const imported = typeof source.slideId === "string" && typeof source.shapeId === "string";
    if (imported && body.from && !(body.from.slideId === source.slideId && body.from.shapeId === source.shapeId)) {
      source.paintedBodyFrom = {
        slideId: body.from.slideId,
        shapeId: body.from.shapeId,
        body: { block: body.block, style: body.style },
      };
    } else {
      delete source.paintedBodyFrom;
    }
    next.source = source;
  }
  return next as unknown as PresentationBlock;
}
