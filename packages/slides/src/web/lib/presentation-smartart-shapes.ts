/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Convert to Shapes on a slide, and Copy of one SmartArt shape: the members
 * the graphic drew, as ordinary slide objects the way PowerPoint writes them.
 *
 * PowerPoint's result is one shape per item with the item's words INSIDE it
 * (SmartArt probe B8, Basic Block List: five shapes, each `p:txBody` with the
 * label's runs at the size the graphic drew them and `a:bodyPr` insets of
 * 247650 EMU, the distance to the box the words stood in), and the shapes'
 * colours stay the theme's (`p:style` `fillRef`/`lnRef` on `accent1`/`lt1`),
 * so a later Design > Themes repaints them. MRAI draws a node and its label
 * as two members, so the words move into the node's text body with those
 * insets ({@link smartArtWordsIntoTheirShapes}), and every colour that IS
 * one of the deck's theme colours keeps that reference (`themeColors`, the
 * binding the slide's other shapes keep, written as `a:schemeClr`).
 */

import type {
  PresentationBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
  PresentationTextStyle,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import {
  releaseSmartArtMembers,
  smartArtWordsIntoTheirShapes,
  type SmartArtMemberBox,
  type SmartArtTextInShapeInsets,
} from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import {
  officeThemeColorGrid,
  type OfficeThemeGridColors,
} from "@mrnote/core/web/lib/office-theme-color-grid.ts";
import { geometryNumber } from "@mrnote/core/web/lib/presentation-block-reads.ts";
import {
  presentationThemeSlotPicks,
  reconcilePresentationThemeColors,
  type PresentationThemeColorPick,
} from "@mrnote/core/web/lib/presentation-theme-color-binding.ts";
import {
  shapeTextFrameLayout,
  smartArtLabelAtDrawnSize,
  smartArtLabelFitScale,
  smartArtLabelInItsBand,
} from "./presentation/text-layout-css";
import { CSS_PX_PER_POINT } from "@mrnote/core/web/lib/presentation/text-units.ts";
import { presentationTextUsesPoints } from "@mrnote/core/web/lib/presentation/text-units-block.ts";

function memberBox(block: PresentationBlock): SmartArtMemberBox | null {
  const left = geometryNumber(block, "left");
  const top = geometryNumber(block, "top");
  const width = geometryNumber(block, "width");
  const height = geometryNumber(block, "height");
  return left === null || top === null || width === null || height === null
    ? null
    : { left, top, width, height };
}

/** The node's text rectangle (its preset's `a:rect`) at `box`, in EMU. */
function nodeTextArea(node: PresentationBlock, box: SmartArtMemberBox): SmartArtMemberBox {
  const rect = node.type === "shape" ? shapeTextFrameLayout(node as PresentationShapeBlock).textRect : null;
  if (!rect) return box;
  return {
    left: box.left + rect.left * box.width,
    top: box.top + rect.top * box.height,
    width: Math.max(0, rect.right - rect.left) * box.width,
    height: Math.max(0, rect.bottom - rect.top) * box.height,
  };
}

type SizedStyle = { fontSize?: unknown; [key: string]: unknown };

/** A label's sizes restated in the unit its new shape keeps (points, or design px). */
function sizeIn(label: PresentationBlock, node: PresentationBlock): (value: unknown) => unknown {
  const from = presentationTextUsesPoints(label);
  const to = presentationTextUsesPoints(node);
  if (from === to) return (value) => value;
  return (value) =>
    typeof value === "number" && Number.isFinite(value)
      ? Math.round((to ? value / CSS_PX_PER_POINT : value * CSS_PX_PER_POINT) * 100) / 100
      : value;
}

function restyled<S extends SizedStyle | undefined>(style: S, size: (value: unknown) => unknown): S {
  if (!style || style.fontSize === undefined) return style;
  return { ...style, fontSize: size(style.fontSize) } as S;
}

/** The label fields a shape's text body keeps (its words and how they are set). */
const CARRIED_TEXT_FIELDS = ["textPaint", "textWarp", "bulletChar", "numberFormat", "columnCount"] as const;

function wordsInShape(
  node: PresentationBlock,
  label: PresentationBlock,
  insets: SmartArtTextInShapeInsets,
): PresentationBlock | null {
  if (node.type !== "shape" || label.type !== "text") return null;
  const text = label as PresentationTextBlock;
  if (typeof text.text !== "string") return null;
  const size = sizeIn(label, node);
  const { sizeUnit: _unit, ...ownStyle } = (restyled(text.style as SizedStyle | undefined, size) ?? {}) as SizedStyle;
  const runs = Array.isArray(text.runs)
    ? text.runs.map((run) => ({ ...run, ...(run.style ? { style: restyled(run.style as SizedStyle, size) } : {}) }))
    : undefined;
  const paragraphs = Array.isArray(text.paragraphs)
    ? text.paragraphs.map((paragraph) =>
        Array.isArray(paragraph.runs)
          ? {
              ...paragraph,
              runs: paragraph.runs.map((run) => ({
                ...run,
                ...(run.style ? { style: restyled(run.style as SizedStyle, size) } : {}),
              })),
            }
          : paragraph,
      )
    : undefined;
  const carried: Record<string, unknown> = {};
  for (const field of CARRIED_TEXT_FIELDS) {
    if ((text as Record<string, unknown>)[field] !== undefined) carried[field] = (text as Record<string, unknown>)[field];
  }
  return {
    ...node,
    ...carried,
    text: text.text,
    ...(runs ? { runs } : {}),
    ...(paragraphs ? { paragraphs } : {}),
    // A text box's absent alignments are left and top; a shape's would be
    // centred, so the label's are stated.
    style: {
      ...ownStyle,
      align: (ownStyle.align as PresentationTextStyle["align"]) ?? "left",
      verticalAlign: (ownStyle.verticalAlign as PresentationTextStyle["verticalAlign"]) ?? "top",
    },
    textInsets: insets,
  } as PresentationBlock;
}

/** Every theme colour of the deck's grid, the ten slots first. */
function deckThemePicks(colors: OfficeThemeGridColors | null | undefined): PresentationThemeColorPick[] {
  if (!colors) return [];
  const grid = officeThemeColorGrid(colors);
  return [
    ...presentationThemeSlotPicks(colors),
    ...grid.columns.flatMap((column, at) =>
      column.slice(1).map((color, step) => ({ color, ref: grid.refs[at]![step + 1]! })),
    ),
  ];
}

/**
 * The members as the ordinary objects Convert to Shapes (or Copy) leaves:
 * words drawn at the size the graphic drew them, inside their shape where
 * an inset can say where they stood, else in a text box in the band they
 * were drawn in; colours that are the deck's theme colours kept as theme
 * references.
 */
export function presentationSmartArtMembersAsShapes(
  members: readonly PresentationBlock[],
  options: {
    /** The slide's blocks, which the graphic's shared label size is read against. */
    siblings: readonly PresentationBlock[] | undefined;
    keepGroup: boolean;
    /** The deck's Theme Colors; absent, every colour stays the literal it is. */
    themeColors?: OfficeThemeGridColors | null;
  },
): PresentationBlock[] {
  const drawn = members.map((member) => {
    const fit = smartArtLabelFitScale(member, options.siblings);
    return withoutRunColorSchemes(fit !== null && fit < 1 ? smartArtLabelAtDrawnSize(member, fit) : member);
  });
  const inside = smartArtWordsIntoTheirShapes(drawn, {
    box: memberBox,
    turn: (member) => geometryNumber(member, "rotation") ?? 0,
    textArea: nodeTextArea,
    takesWords: (node) => {
      const record = node as Record<string, unknown>;
      return (
        node.type === "shape" &&
        !(typeof record.text === "string" && record.text.trim() !== "") &&
        record.flipH !== true &&
        record.flipV !== true
      );
    },
    merge: wordsInShape,
  });
  const released = releaseSmartArtMembers(inside.map(smartArtLabelInItsBand), { keepGroup: options.keepGroup });
  const picks = deckThemePicks(options.themeColors);
  return picks.length ? released.map((block) => reconcilePresentationThemeColors(undefined, block, picks)) : released;
}

/**
 * A member's runs as an ordinary object keeps them: in the colour they draw,
 * without the SmartArt point's theme reference (`RunStyle.colorScheme`),
 * which only the diagram's writer states back as `a:schemeClr`.
 */
function withoutRunColorSchemes(input: PresentationBlock): PresentationBlock {
  let member = input;
  if ("smartArtTextColorScheme" in (member as object)) {
    const { smartArtTextColorScheme: _label, ...rest } = member as PresentationBlock & { smartArtTextColorScheme?: unknown };
    member = rest as PresentationBlock;
  }
  const runs = (member as { runs?: { text: string; style?: Record<string, unknown> }[] }).runs;
  if (!Array.isArray(runs) || !runs.some((run) => run.style && "colorScheme" in run.style)) return member;
  return {
    ...member,
    runs: runs.map((run) => {
      if (!run.style || !("colorScheme" in run.style)) return run;
      const { colorScheme: _scheme, ...style } = run.style;
      return Object.keys(style).length ? { ...run, style } : { text: run.text };
    }),
  } as PresentationBlock;
}
