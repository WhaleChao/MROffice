/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { OfficeRibbonCommandGate } from "@mrnote/core/web/components/office-ribbon/officeRibbonCommandGate.tsx";
import { selectionGateFromEntries } from "@mrnote/core/web/components/office-ribbon/officeRibbonSelectionGate.ts";
import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { blockGroupId } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import { normalizeShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";

/** The kinds of object a slide selection can be, as the gate tells them apart. */
export type PresentationSelectedStructure =
  | "picture"
  | "shape"
  | "shapeWithoutText"
  | "ink"
  | "chart"
  | "table"
  | "media"
  | "smartArtShape"
  | "smartArtGraphic"
  | "group"
  | "groupWithoutText"
  | "multiple";

/** The block fields the gate reads (a slide block satisfies it). */
export interface PresentationGateBlock {
  readonly id: string;
  readonly type: string;
  readonly shape?: unknown;
  readonly text?: unknown;
  readonly groupId?: unknown;
}

/** What the window knows about its selection, as the gate reads it. */
export interface PresentationSelectionFacts {
  readonly blocks: readonly PresentationGateBlock[];
  readonly selectedBlockIds: ReadonlySet<string>;
  /** The whole SmartArt graphic is selected (its frame). */
  readonly smartArtGraphicSelected: boolean;
  /** A caret is in a block's text (a text block being edited, a shape's words). */
  readonly editingText: boolean;
  /**
   * A cell of the selected table has the focus: Home formats that cell's
   * (or the cell range's) text, the size box included, as it does a word.
   */
  readonly tableCellFocused?: boolean;
}

/** Whether a selected block carries words Home can run-format. */
function carriesWords(block: PresentationGateBlock): boolean {
  return (
    (block.type === "text" || block.type === "shape") &&
    typeof block.text === "string" &&
    block.text.length > 0
  );
}

/**
 * The object the selection is, or null for a text block (frame-selected, the
 * Font and Paragraph groups format all of its text, as in PowerPoint), a
 * focused table cell and while typing, where every command applies.
 */
export function presentationSelectedStructure(
  facts: PresentationSelectionFacts,
): PresentationSelectedStructure | null {
  if (facts.editingText) return null;
  const selected = facts.blocks.filter((block) =>
    facts.selectedBlockIds.has(block.id),
  );
  if (selected.length === 0) return null;
  if (facts.smartArtGraphicSelected) return "smartArtGraphic";
  if (selected.length > 1) {
    const markers = selected.map((block) => readSmartArtBlockMarker(block));
    if (
      markers.every((marker) => marker) &&
      new Set(markers.map((marker) => marker?.groupId)).size === 1
    ) {
      return "smartArtGraphic";
    }
    const worded = selected.some(carriesWords);
    const groups = new Set(
      selected.map((block) => blockGroupId(block as never) ?? null),
    );
    if (groups.size === 1 && !groups.has(null)) {
      return worded ? "group" : "groupWithoutText";
    }
    return "multiple";
  }
  const [block] = selected;
  if (readSmartArtBlockMarker(block)) return "smartArtShape";
  switch (block.type) {
    case "image":
      return "picture";
    case "shape":
      if (normalizeShapeKind(block.shape) === "ink") return "ink";
      return carriesWords(block) ? "shape" : "shapeWithoutText";
    case "chart":
      return "chart";
    case "table":
      return facts.tableCellFocused ? null : "table";
    case "embed":
      return "media";
    default:
      return null;
  }
}

const SHAPE_LABEL = "presentation.ui.selectionGate.shape";
const GROUP_LABEL = "presentation.ui.selectionGate.group";
const OBJECTS_LABEL = "presentation.ui.selectionGate.objects";

/** The catalogue key of each structure's whole "not available" sentence. */
const STRUCTURE_LABEL: Record<PresentationSelectedStructure, string> = {
  picture: "presentation.ui.selectionGate.picture",
  shape: SHAPE_LABEL,
  shapeWithoutText: SHAPE_LABEL,
  ink: "presentation.ui.selectionGate.ink",
  chart: "presentation.ui.selectionGate.chart",
  table: "presentation.ui.selectionGate.table",
  media: "presentation.ui.selectionGate.media",
  smartArtShape: "presentation.ui.selectionGate.smartArtShape",
  smartArtGraphic: "presentation.ui.selectionGate.smartArtGraphic",
  group: GROUP_LABEL,
  groupWithoutText: GROUP_LABEL,
  multiple: OBJECTS_LABEL,
};

const RIBBON = "presentation-ribbon-";

/**
 * The Font Size box's list. The host already disables the rest of the Font
 * and Paragraph groups when no text block is selected; this dropdown stayed
 * enabled and its sizes changed nothing (measured on every object kind).
 */
const FONT_SIZE_LIST = [
  `${RIBBON}text-font-size`,
  `${RIBBON}text-font-size-options`,
  `${RIBBON}text-font-size-option-`,
];

/**
 * What an object WITHOUT words gives the run marks: superscript, subscript
 * and highlight mark runs, and a shape with no text has none to mark (the
 * object path writes them on every run of an object's words, as PowerPoint
 * formats all of a selected object's text). Align Text anchors the text
 * frame itself and applies to an empty shape too.
 */
const OBJECT_RUN_MARKS = [
  `${RIBBON}text-superscript`,
  `${RIBBON}text-subscript`,
  `${RIBBON}text-highlight`,
  `${RIBBON}text-highlight-menu`,
];

/**
 * What a chart's own text does not take from Home: it has no anchor of its
 * own, and its run marks are not written to its text properties.
 */
const OBJECT_TEXT_UNWRITTEN = [
  ...OBJECT_RUN_MARKS,
  `${RIBBON}align-text`,
  `${RIBBON}text-valign-`,
];

/**
 * Strikethrough marks each run, so an object without words gives it nothing
 * to mark (an empty box's own style stores no strikethrough). Bold, Italic,
 * Underline, the font, the size list, Grow / Shrink Font and the colour are
 * stored on the object and take effect on the words typed into it, so they
 * stay.
 */
const RUNS_ONLY = [
  `${RIBBON}text-strike`,
];

/**
 * A chart's own text (c:txPr) takes the face, size, bold, italic and colour;
 * it has no strikethrough, underline, paragraph alignment or anchor to take
 * (measured, pixel-compared for underline).
 */
const CHART_TEXT_UNWRITTEN = [
  ...OBJECT_TEXT_UNWRITTEN,
  `${RIBBON}text-strike`,
  `${RIBBON}text-underline`,
  `${RIBBON}text-align-left`,
  `${RIBBON}text-align-center`,
  `${RIBBON}text-align-right`,
  `${RIBBON}text-align-justify`,
];

/**
 * An ink stroke cannot be typed into (it takes no words), so the Font and
 * Paragraph commands Home leaves enabled either change nothing (Grow /
 * Shrink, Strikethrough, the scripts, Highlight, Align Text) or store a
 * property the stroke never draws (the face, size, Bold, Italic, Underline,
 * font colour, alignment: pixel-compared, unchanged). Like a picture, it
 * gives the Font group nothing to act on.
 */
const INK_WORDLESS = [
  `${RIBBON}text-font-family`,
  `${RIBBON}text-font-family-`,
  ...FONT_SIZE_LIST,
  ...RUNS_ONLY,
  `${RIBBON}text-bold`,
  `${RIBBON}text-italic`,
  `${RIBBON}text-underline`,
  `${RIBBON}text-color`,
  `${RIBBON}text-color-`,
  ...OBJECT_TEXT_UNWRITTEN,
  `${RIBBON}text-align-left`,
  `${RIBBON}text-align-center`,
  `${RIBBON}text-align-right`,
  `${RIBBON}text-align-justify`,
];

/**
 * Test ids of the commands that have no meaning for each object, as
 * measured. A prefix entry ending in "-" covers a menu's every item.
 */
const GATED: Record<PresentationSelectedStructure, readonly string[]> = {
  // A picture and a media frame carry no text: as in PowerPoint, the Font
  // group has nothing to act on.
  picture: FONT_SIZE_LIST,
  media: FONT_SIZE_LIST,
  // A table selected by its border takes Home on every cell (PowerPoint
  // formats all of its cells); a focused cell is not an object at all.
  table: [],
  // A shape's words, every SmartArt node, a chart's own text and every
  // object of a multi-selection take the Font and Paragraph groups, the
  // size, the scripts, the highlight and Align Text included
  // (lib/presentation-home-object-text), as in PowerPoint.
  shape: [],
  shapeWithoutText: [...OBJECT_RUN_MARKS, ...RUNS_ONLY],
  // A group selects its members, so it reads as their multi-selection.
  group: [],
  groupWithoutText: [...OBJECT_RUN_MARKS, ...RUNS_ONLY],
  multiple: [],
  // A click on a SmartArt diagram selects the whole graphic; Hide Selected
  // hides it as one object, as its Selection Pane row does.
  smartArtGraphic: [],
  // One SmartArt node alone: not reachable by a click, left as it was.
  smartArtShape: [],
  chart: CHART_TEXT_UNWRITTEN,
  ink: INK_WORDLESS,
};

/** The gate for `structure`; null for text, where every command applies. */
export function presentationRibbonSelectionGate(
  structure: PresentationSelectedStructure | null,
  t: (key: string) => string,
): OfficeRibbonCommandGate | null {
  if (!structure) return null;
  return selectionGateFromEntries(GATED[structure], t(STRUCTURE_LABEL[structure]));
}

export { RIBBON as PRESENTATION_RIBBON_TEST_ID_PREFIX };
