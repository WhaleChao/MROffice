/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Insert > WordArt and Shape Format > WordArt Styles on the deck (PPT-A059):
 * the transforms the ribbon and the editor commands share ("button =
 * command"). The styles are PowerPoint's own twenty
 * (`presentation-wordart-gallery`), under the deck's theme.
 *
 * Insert > WordArt adds a text box, as PowerPoint does (`txBox`), holding
 * PowerPoint's placeholder words, centred, at 54 pt, in the chosen style.
 * WordArt Styles gives a selected text box's or shape's words the style: its
 * paint, and the style's own bold and character spacing.
 */
import { CSS_PX_PER_POINT, pointsToDesignPx } from "@mrnote/core/web/lib/presentation/text-units.ts";
import { POINT_TEXT_UNIT } from "@mrnote/core/web/lib/presentation/text-units-block.ts";
import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import type { OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import { presentationThemeGridColors } from "@mrnote/core/web/lib/presentation-theme-color-binding.ts";
import type { DeckTheme } from "@mrnote/core/web/lib/presentation-theme.ts";
import {
  presentationWordArtStyles,
  type PresentationWordArtStyle,
} from "@mrnote/core/web/lib/presentation-wordart-gallery.ts";
import { defaultShapeGeometry } from "@mrnote/core/web/lib/presentation-shapes.ts";
import { slideBounds } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { shapeTextPaintPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER } from "@mrnote/core/web/lib/presentation-smartart-text.ts";
import type { TextRun } from "@mrnote/core/web/lib/presentation-runs.ts";

/** PowerPoint's WordArt size: 54 pt, as a new WordArt box states it. */
export const PRESENTATION_WORDART_FONT_SIZE_PT = 54;
/** The same size in the deck's design px (96 per inch), to measure with. */
export const PRESENTATION_WORDART_FONT_SIZE_PX = pointsToDesignPx(PRESENTATION_WORDART_FONT_SIZE_PT);

/** The gallery's styles under a deck's theme. */
export function presentationDeckWordArtStyles(
  theme: DeckTheme | null | undefined,
): readonly PresentationWordArtStyle[] {
  return presentationWordArtStyles(presentationThemeGridColors(theme));
}

/** A style by its gallery id (`ppt-wordart-<item>`) under a deck's theme. */
export function presentationDeckWordArtStyle(
  theme: DeckTheme | null | undefined,
  id: string,
): PresentationWordArtStyle | null {
  return presentationDeckWordArtStyles(theme).find((style) => style.id === id) ?? null;
}

/** Insert > WordArt: a new text box with PowerPoint's placeholder words in `style`. */
export function wordArtTextBlock(
  id: string,
  style: PresentationWordArtStyle,
  text: string,
  slide: { readonly [key: string]: unknown },
  blocks: readonly PresentationBlock[],
): PresentationTextBlock {
  return {
    id,
    type: "text",
    kind: "textbox",
    text,
    textFontSizeUnit: POINT_TEXT_UNIT,
    style: {
      fontSize: PRESENTATION_WORDART_FONT_SIZE_PT,
      align: "center",
      ...(style.letterSpacingPt ? { letterSpacing: style.letterSpacingPt } : {}),
    },
    // The style's bold rides on the run, where every writer puts `a:rPr@b`
    // (PowerPoint recognises a gallery style only with its bold).
    ...(style.bold ? { runs: [{ text, style: { bold: true } }] } : {}),
    textPaint: style.paint,
    // PowerPoint's WordArt box follows its words (Resize shape to fit text).
    autofit: "resize-shape",
    ...wordArtGeometry(text, style, slide, blocks),
  } as PresentationTextBlock;
}

const EMU_PER_PX = 9525;
/** The canvas's line height for a text box's words (`.canvas` 1.45). */
const WORDART_LINE_HEIGHT = 1.45;
/** A text box's own insets, 0.1" across and 0.05" down each side, in px. */
const WORDART_INSET_X_PX = 9.6 * 2;
const WORDART_INSET_Y_PX = 4.8 * 2;

/** How wide `text` runs at `fontPx`: a CJK character an em, Latin less. */
function estimatedTextWidthPx(text: string, fontPx: number, letterSpacingPt = 0): number {
  let ems = 0;
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    ems +=
      code >= 0x2e80 && code <= 0xffef ? 1 // CJK, kana, hangul, full-width
        : char === " " ? 0.3
          : /[A-Z0-9]/.test(char) ? 0.68
            : 0.56;
  }
  return ems * fontPx + [...text].length * letterSpacingPt * CSS_PX_PER_POINT;
}







function wordArtGeometry(
  text: string,
  style: PresentationWordArtStyle,
  slide: { readonly [key: string]: unknown },
  blocks: readonly PresentationBlock[],
) {
  const bounds = slideBounds(slide);
  const base = defaultShapeGeometry("textbox", bounds, [...blocks]);
  const slideWidth = typeof bounds.width === "number" && bounds.width > 0 ? bounds.width : base.width / 0.34;
  const slideHeight = typeof bounds.height === "number" && bounds.height > 0 ? bounds.height : base.height / 0.22;
  // A little room past the estimate, so a face a touch wider still fits.
  const maxContentPx = (slideWidth * 0.9) / EMU_PER_PX - WORDART_INSET_X_PX - 8;
  const words = estimatedTextWidthPx(text, PRESENTATION_WORDART_FONT_SIZE_PX, style.letterSpacingPt ?? 0);
  const lines = Math.max(1, Math.ceil(words / maxContentPx));
  const width = Math.round((Math.min(words, maxContentPx) + WORDART_INSET_X_PX + 8) * EMU_PER_PX);
  const height = Math.round(
    (lines * PRESENTATION_WORDART_FONT_SIZE_PX * WORDART_LINE_HEIGHT + WORDART_INSET_Y_PX + 4) * EMU_PER_PX,
  );
  const clamp = (value: number, size: number, extent: number) =>
    Math.max(0, Math.min(Math.round(value), Math.max(0, Math.round(extent - size))));
  return {
    left: clamp(base.left + (base.width - width) / 2, width, slideWidth),
    top: clamp(base.top + (base.height - height) / 2, height, slideHeight),
    width,
    height,
  };
}

/** Whether a block's words can wear a WordArt style: a text box, or a shape with words. */
export function presentationBlockTakesWordArt(block: PresentationBlock | null | undefined): boolean {
  if (!block) return false;
  if (block.type === "text") return Boolean((block as PresentationTextBlock).text?.trim());
  if (block.type === "shape") return Boolean((block as PresentationShapeBlock).text?.trim());
  return false;
}

/** The WordArt paint a block's words wear, or null. */
export function presentationBlockTextPaint(block: PresentationBlock | null | undefined): OfficeTextPaint | null {
  if (!block || (block.type !== "text" && block.type !== "shape")) return null;
  return ((block as { textPaint?: OfficeTextPaint }).textPaint ?? null) as OfficeTextPaint | null;
}

/**
 * A text box's or shape's words with a WordArt paint (`null` takes it off).
 * A shape's words lose the colour a run states (Text Fill IS the text's
 * colour, `shapeTextPaintPatch`); a text box's runs keep theirs and draw in
 * the paint (`[data-text-paint]`), so an imported box changes in one place.
 */
export function presentationBlockWithTextPaint(
  block: PresentationBlock,
  paint: OfficeTextPaint | null,
): PresentationBlock {
  if (block.type === "shape") {
    return { ...block, ...shapeTextPaintPatch(block as PresentationShapeBlock, paint) } as PresentationBlock;
  }
  if (block.type !== "text") return block;
  const next = { ...block } as Record<string, unknown>;
  if (paint) next.textPaint = paint;
  else delete next.textPaint;
  return next as PresentationBlock;
}

/**
 * Shape Format > WordArt Styles: the style's paint on a text box's or a
 * shape's words, with the style's own bold and character spacing.
 */
export function presentationBlockWithWordArtStyle(
  block: PresentationBlock,
  style: PresentationWordArtStyle,
): PresentationBlock {
  let next = presentationBlockWithTextPaint(block, style.paint);
  if (style.bold && next.type === "text" && textBlockHasRuns(next as PresentationTextBlock)) {
    // A text box's words are bold run by run, as PowerPoint writes the
    // style's `b` on each run: an imported box's runs are what its in-place
    // patch writes back.
    next = textBlockWithEveryRunBold(next as PresentationTextBlock) as PresentationBlock;
  } else if (style.bold && (next.type === "text" || next.type === "shape")) {
    next = PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER.write(
      next as PresentationTextBlock,
      { bold: true },
    ) as PresentationBlock;
  }
  if (style.letterSpacingPt && next.type === "text") {
    const own = ((next as PresentationTextBlock).style ?? {}) as Record<string, unknown>;
    next = { ...next, style: { ...own, letterSpacing: style.letterSpacingPt } } as PresentationBlock;
  }
  return next;
}

function textBlockHasRuns(block: PresentationTextBlock): boolean {
  return (
    (Array.isArray(block.runs) && block.runs.length > 0) ||
    (Array.isArray(block.paragraphs) &&
      block.paragraphs.some((paragraph) => Array.isArray(paragraph.runs) && paragraph.runs.length > 0))
  );
}

function textBlockWithEveryRunBold(block: PresentationTextBlock): PresentationTextBlock {
  const bold = (runs: readonly TextRun[] | undefined) =>
    runs?.map((run) => ({ ...run, style: { ...(run.style ?? {}), bold: true } }));
  return {
    ...block,
    ...(Array.isArray(block.runs) ? { runs: bold(block.runs) } : {}),
    ...(Array.isArray(block.paragraphs)
      ? {
          paragraphs: block.paragraphs.map((paragraph) =>
            Array.isArray(paragraph.runs) ? { ...paragraph, runs: bold(paragraph.runs) } : paragraph,
          ),
        }
      : {}),
  };
}
