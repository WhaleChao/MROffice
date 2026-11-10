/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Format Shape ▸ Text Options ▸ Text Box on the deck (PPT-A050): what a text
 * box's or a shape's text frame states, and the block with one setting
 * changed.
 *
 * PowerPoint's Text Box page (Microsoft Support, "Change the margins between
 * text and the edge of a text box or shape"; "Wrap text around a circle or
 * other shape"; "Shrink or enlarge text to fit in a text box") holds the
 * `a:bodyPr` settings of ONE text frame:
 *
 * - Vertical alignment: `@anchor` t / ctr / b, plus `@anchorCtr` for the
 *   three "centered" variants (the text block also centred across).
 * - Text direction: `@vert` horz / eaVert / vert / vert270 / wordArtVert.
 * - Do not rotate text: `@upright` (not modelled here).
 * - Do not Autofit / Shrink text on overflow / Resize shape to fit text:
 *   `a:noAutofit` / `a:normAutofit` / `a:spAutoFit`.
 * - Left / Right / Top / Bottom margin: `@lIns/@rIns/@tIns/@bIns`, EMU;
 *   PowerPoint's own defaults are 0.1" left and right, 0.05" top and bottom.
 * - Wrap text in shape: `@wrap` square / none.
 * - Columns...: `@numCol` / `@spcCol` (the Columns dialog).
 *
 * A text block and a shape keep these in the same fields (the shape's label
 * flows in the same frame model): `style.verticalAlign`,
 * `style.anchorCenter`, `style.writingMode`, `autofit`, `textInsets`,
 * `style.textWrap` (false only), `columnCount` / `columnSpacing`. A shape's
 * label is centred when it states no anchor; a text box is at the top.
 *
 * Pure and dependency-light: the pane module renders these, the window
 * commits what the writers return (null: nothing changed or not allowed).
 */

import type {
  PresentationShapeBlock,
  PresentationTextBlock,
  PresentationTextInsets,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  effectiveAutofitMode,
  nativeAutofitResizesShape,
  type AutofitMode,
} from "@mrnote/core/web/lib/presentation-autofit.ts";
import { applyPresentationHomeObjectAnchor } from "@mrnote/core/web/lib/presentation-home-object-text.ts";
import { normalizeWritingMode, type WritingMode } from "@mrnote/core/web/lib/presentation-line-spacing.ts";
import { isImportedPptxBlock } from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import {
  effectiveColumnSpacingEmu,
  normalizeColumnCount,
  normalizeColumnSpacing,
  persistedColumnCount,
} from "@mrnote/core/web/lib/presentation-text-columns.ts";
import {
  DEFAULT_PRESENTATION_TEXT_INSETS,
  PRESENTATION_TEXT_INSET_EDGES,
  type PresentationTextInsetEdge,
} from "./presentation-text-insets";
import { PRESENTATION_GEOMETRY_MAX_EMU } from "@mrnote/core/web/lib/presentation-geometry-limit.ts";
import { readAnchorCenter, shapeLabelVerticalAlign } from "@mrnote/core/web/lib/presentation-vertical-anchor.ts";

export type PresentationTextFrameBlock = PresentationTextBlock | PresentationShapeBlock;

/** Vertical alignment's six choices, PowerPoint's order. */
export const PRESENTATION_TEXT_FRAME_ANCHORS = [
  "top",
  "middle",
  "bottom",
  "top-centered",
  "middle-centered",
  "bottom-centered",
] as const;
export type PresentationTextFrameAnchor = (typeof PRESENTATION_TEXT_FRAME_ANCHORS)[number];

/** Text direction's five choices, PowerPoint's order. */
export const PRESENTATION_TEXT_FRAME_DIRECTIONS = [
  "horizontal-tb",
  "vertical-rl",
  "rot90",
  "rot270",
  "stacked",
] as const satisfies readonly WritingMode[];

/** The three autofit choices (Resize shape to fit text: `presentation-text-fit`). */
export const PRESENTATION_TEXT_FRAME_AUTOFITS = ["none", "shrink", "resize-shape"] as const;
export type PresentationTextFrameAutofit = (typeof PRESENTATION_TEXT_FRAME_AUTOFITS)[number];

export interface PresentationTextFrameOptions {
  readonly anchor: PresentationTextFrameAnchor;
  readonly direction: WritingMode;
  readonly autofit: AutofitMode;
  /** The four margins the frame draws with (PowerPoint's defaults when unstated). */
  readonly insets: PresentationTextInsets;
  readonly wrap: boolean;
  readonly columnCount: number;
  readonly columnSpacingEmu: number;
}

function styleOf(block: PresentationTextFrameBlock): Record<string, unknown> {
  const style = block.style;
  return style && typeof style === "object" && !Array.isArray(style)
    ? (style as Record<string, unknown>)
    : {};
}

function frameInsets(block: PresentationTextFrameBlock): PresentationTextInsets | null {
  const raw = (block as { textInsets?: unknown }).textInsets;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const record = raw as Record<string, unknown>;
  const next = {} as PresentationTextInsets;
  for (const edge of PRESENTATION_TEXT_INSET_EDGES) {
    const value = record[edge];
    if (typeof value !== "number" || !Number.isInteger(value) || value < 0) return null;
    next[edge] = value;
  }
  return next;
}

/** What one text box's or shape's frame states, PowerPoint's defaults filled in. */
export function presentationTextFrameOptions(
  block: PresentationTextFrameBlock,
): PresentationTextFrameOptions {
  const style = styleOf(block);
  const vertical =
    block.type === "shape"
      ? (shapeLabelVerticalAlign(style, block.source) ?? "middle")
      : style.verticalAlign === "middle" || style.verticalAlign === "bottom"
        ? style.verticalAlign
        : "top";
  const centered = readAnchorCenter(style as { anchorCenter?: unknown });
  return {
    anchor: centered ? (`${vertical}-centered` as PresentationTextFrameAnchor) : vertical,
    direction: normalizeWritingMode(style.writingMode) ?? "horizontal-tb",
    autofit: effectiveAutofitMode(block),
    insets: frameInsets(block) ?? { ...DEFAULT_PRESENTATION_TEXT_INSETS },
    wrap: style.textWrap !== false,
    columnCount: normalizeColumnCount((block as { columnCount?: unknown }).columnCount),
    columnSpacingEmu: effectiveColumnSpacingEmu(block as { columnSpacing?: unknown; source?: unknown }),
  };
}

/**
 * Whether the frame's margins can be written: an authored box always; an
 * imported one only when it is the slide's own shape (a placeholder still
 * inheriting its layout's frame has no shape of its own to carry them).
 */
export function presentationTextFrameInsetsWritable(block: PresentationTextFrameBlock): boolean {
  const source = (block as { source?: unknown }).source;
  const record =
    source && typeof source === "object" && !Array.isArray(source)
      ? (source as Record<string, unknown>)
      : null;
  if (record?.inheritedFrom !== undefined) return false;
  if (!isImportedPptxBlock(block)) return true;
  if (!record) return false;
  return (
    typeof record.slideId === "string" &&
    record.slideId !== "" &&
    typeof record.shapeId === "string" &&
    record.shapeId !== "" &&
    record.inheritedFrom === undefined
  );
}

/**
 * The centred variants are drawn on a horizontal text box only
 * (`PresentationAnchorCenterBounds`; the import reads `@anchorCtr` only on
 * horizontal text); a shape's label keeps its own width.
 */
export function presentationTextFrameTakesAnchorCenter(block: PresentationTextFrameBlock): boolean {
  return block.type === "text" && normalizeWritingMode(styleOf(block).writingMode) === undefined;
}

function withStyle(
  block: PresentationTextFrameBlock,
  style: Record<string, unknown>,
): PresentationTextFrameBlock {
  const next = { ...block } as Record<string, unknown>;
  if (Object.keys(style).length > 0) next.style = style;
  else delete next.style;
  return next as PresentationTextFrameBlock;
}

/** Vertical alignment (and its "centered" half). Null when unchanged or not drawable. */
export function presentationTextFrameWithAnchor(
  block: PresentationTextFrameBlock,
  anchor: PresentationTextFrameAnchor,
): PresentationTextFrameBlock | null {
  const centered = anchor.endsWith("-centered");
  if (centered && !presentationTextFrameTakesAnchorCenter(block)) return null;
  const vertical = anchor.replace("-centered", "") as "top" | "middle" | "bottom";
  const anchored = (applyPresentationHomeObjectAnchor([block], vertical)?.[0] ?? block) as PresentationTextFrameBlock;
  const style = { ...styleOf(anchored) };
  if (centered) style.anchorCenter = true;
  else delete style.anchorCenter;
  const next = withStyle(anchored, style);
  return JSON.stringify(next) === JSON.stringify(block) ? null : next;
}

/** Text direction. Horizontal is stated by omission. */
export function presentationTextFrameWithDirection(
  block: PresentationTextFrameBlock,
  direction: WritingMode,
): PresentationTextFrameBlock | null {
  const style = { ...styleOf(block) };
  const value = normalizeWritingMode(direction);
  if ((style.writingMode ?? undefined) === value) return null;
  if (value === undefined) delete style.writingMode;
  else style.writingMode = value;
  return withStyle(block, style);
}

/**
 * Do not Autofit / Shrink text on overflow / Resize shape to fit text, on the
 * block (the canonical place). Do not Autofit is stated only where it turns
 * off an imported box's `a:spAutoFit`; elsewhere it is the default, omitted.
 */
export function presentationTextFrameWithAutofit(
  block: PresentationTextFrameBlock,
  mode: AutofitMode,
): PresentationTextFrameBlock | null {
  if (presentationTextFrameOptions(block).autofit === mode) return null;
  const style = { ...styleOf(block) };
  delete style.autofit;
  const next = withStyle(block, style) as Record<string, unknown>;
  if (mode !== "none") next.autofit = mode;
  else if (nativeAutofitResizesShape(block)) next.autofit = "none";
  else delete next.autofit;
  return next as PresentationTextFrameBlock;
}

/** One margin, in EMU; the other three keep what the frame draws with. */
export function presentationTextFrameWithInset(
  block: PresentationTextFrameBlock,
  edge: PresentationTextInsetEdge,
  emu: number,
): PresentationTextFrameBlock | null {
  if (!Number.isInteger(emu) || emu < 0 || emu > PRESENTATION_GEOMETRY_MAX_EMU) return null;
  if (!presentationTextFrameInsetsWritable(block)) return null;
  const current = presentationTextFrameOptions(block).insets;
  if (frameInsets(block) !== null && current[edge] === emu) return null;
  return { ...block, textInsets: { ...current, [edge]: emu } } as PresentationTextFrameBlock;
}

/** Wrap text in shape: `false` is stated, wrapping (the default) is not. */
export function presentationTextFrameWithWrap(
  block: PresentationTextFrameBlock,
  wrap: boolean,
): PresentationTextFrameBlock | null {
  const style = { ...styleOf(block) };
  if ((style.textWrap !== false) === wrap) return null;
  if (wrap) delete style.textWrap;
  else style.textWrap = false;
  return withStyle(block, style);
}

/** Columns...: the count and the spacing (EMU); one column states neither. */
export function presentationTextFrameWithColumns(
  block: PresentationTextFrameBlock,
  count: number,
  spacingEmu: number | undefined,
): PresentationTextFrameBlock | null {
  const nextCount = persistedColumnCount(count);
  // A spacing the dialog did not touch stays what the box had (stated or
  // not): OK on an untouched dialog used to write the field's 0 over a box
  // drawn with the default gap, piling every column into the first (PPT-B433).
  const spacing =
    spacingEmu === undefined
      ? normalizeColumnSpacing((block as { columnSpacing?: unknown }).columnSpacing)
      : normalizeColumnSpacing(Math.round(spacingEmu));
  const next = { ...block } as Record<string, unknown>;
  delete next.columnCount;
  delete next.columnSpacing;
  if (nextCount !== undefined) {
    next.columnCount = nextCount;
    if (spacing !== undefined) next.columnSpacing = spacing;
  }
  return JSON.stringify(next) === JSON.stringify(block) ? null : (next as PresentationTextFrameBlock);
}
