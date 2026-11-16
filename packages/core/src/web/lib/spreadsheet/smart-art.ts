/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SmartArt (Excel / WPS 智能图形) for the worksheet: project the SHARED Office
 * SmartArt scene onto the sheet's own drawing objects.
 *
 * The layouts themselves come from `lib/office-smartart/{catalog,engine}` — the
 * one catalog Word, Excel and PowerPoint share — so the sheet does not own a
 * second, smaller set of diagrams. What IS sheet-specific is the projection:
 * every scene node becomes an ORDINARY shape object in `sheet.objects[]`, laid
 * out on the cell grid, so a diagram selects, moves, resizes, persists and
 * exports (as native `xdr:sp`) through exactly one object path instead of
 * becoming a second kind of drawing the rest of the editor has to learn about.
 * That also means a user can pull one box out of a diagram and edit it like any
 * other shape.
 *
 * Pure: given a scene it returns the members to insert. The caller owns ids and
 * anchors, because only the live sheet knows its column widths.
 */

import { ptToPx } from "../office-units";
import { MAX_OFFICE_TEXT_LAYOUT_PARAGRAPHS, type OfficeTextLayout } from "../office-text-layout";
import { trimRunsLikeText, type TextRun } from "../presentation-runs";
import {
  SMARTART_PICTURE_PLACEHOLDER_FILL as SMART_ART_PICTURE_PLACEHOLDER_FILL,
  SMARTART_PICTURE_PLACEHOLDER_SRC as SMART_ART_PICTURE_PLACEHOLDER_SRC,
  SMARTART_PICTURE_PLACEHOLDER_STROKE as SMART_ART_PICTURE_PLACEHOLDER_STROKE,
  smartArtLabelInsideHost,
  smartArtLabelSizeIsPx,
  type OfficeSmartArtNode,
  type OfficeSmartArtShapeNode,
  type OfficeSmartArtTextNode,
} from "../office-smartart/model";
import {
  readSmartArtCustom,
  smartArtLayoutBox,
} from "../office-smartart/custom-geometry";
import { smartArtPictureSrc } from "../office-smartart/picture";
import {
  smartArtSlotPictureFields,
  smartArtSlotPictureFormatOfBlock,
} from "../office-smartart/slot-picture-format";
import type { CropRect } from "../presentation-crop-group";
import type { OfficePictureStyleId } from "../office-picture-styles/catalog";
import { smartArtTextFitScale } from "../office-smartart/text-fit";
import {
  readSmartArtTextBox,
  SMARTART_TEXT_BOX_UNIT,
  smartArtLabelTextBox,
  type SmartArtTextBox,
} from "../office-smartart/text-box";
import {
  normalizePresentationCustomPath,
  type PresentationCustomPath,
} from "../presentation-freeform";
import {
  isSpreadsheetPresetShapeKind,
  normalizeShapeStyle,
  normalizeSpreadsheetCustomPoints,
  SPREADSHEET_CUSTOM_SHAPE_KIND,
  type SpreadsheetObjectShapeKind,
  type SpreadsheetPresetShapeKind,
  type SpreadsheetShapeStyle,
} from "./shapes";

/** Excel caps its own SmartArt panes well below this; the bound is a guard. */
export const MAX_SMART_ART_NODES = 12;

/**
 * The frame a fresh worksheet SmartArt graphic is laid out in, in EMU: 5 in by
 * 3 in, Excel's own. Excel drops a new graphic at a fixed size rather than
 * scaling it to the cells under it; a Basic Block List inserted by Excel for
 * Mac fills exactly 4572000 x 2743200 EMU of its drawing
 * (`SmartArt39-MS-part1.xlsx`, `xl/diagrams/drawing1.xml`). MRAI's frame was
 * six default columns by twelve default rows, 5 in by 2.5 in, so every layout
 * was drawn a sixth shorter than Excel draws it.
 */
export const SPREADSHEET_SMARTART_FRAME_EMU = Object.freeze({
  width: 4_572_000,
  height: 2_743_200,
});

/**
 * The largest a worksheet diagram's labels are drawn, in points: the size a
 * fresh insert asks the engine for (its own 32 pt is slide type). The fit
 * below takes every label down from here together; a rebuild starts from
 * here again, so text that shrank for a long label grows back once the label
 * is short again, as Office's does.
 */
export const SPREADSHEET_SMARTART_FONT_PT = 14;

/** EMU per CSS pixel, the sheet object layer's own scale (96 dpi). */
const EMU_PER_CSS_PIXEL = 9525;

/**
 * The labels a SmartArt insertion reads off the selected cells. Reading them
 * from the selection is what makes it one click — the list is almost always
 * already in the sheet, and re-typing it into a pane is the part people skip.
 * Blanks are dropped rather than becoming empty boxes, formulas are skipped
 * (their text is not a label), and the list is bounded so one paste of a long
 * column cannot cover the sheet in shapes.
 */
export function smartArtLabelsFromValues(
  values: readonly string[],
): string[] {
  return values
    .map((value) => String(value ?? "").trim())
    .filter((value) => value !== "" && !value.startsWith("="))
    .slice(0, MAX_SMART_ART_NODES);
}

/** Where one member sits, in CSS pixels relative to the diagram frame. */
export interface SpreadsheetSmartArtMemberBounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * One projected diagram member: a shape kind, its normalized style (label text
 * included) and its frame-relative box. The identity markers travel with it so
 * a later round can find the members of one diagram again without guessing
 * from visual similarity.
 */
export interface SpreadsheetSmartArtMember {
  kind: SpreadsheetObjectShapeKind;
  shape: SpreadsheetShapeStyle;
  bounds: SpreadsheetSmartArtMemberBounds;
  /**
   * The member's turn in degrees, clockwise, about its own centre -- a cycle's
   * flow arrows point along the ring, a decoration lies at its angle. Absent
   * when the engine drew the member upright.
   */
  rotation?: number;
  /**
   * Drawn mirrored (the anchor's `flipH`): a Right to Left diagram's arrows
   * and chevrons. Only ever `true`; the host puts it on the object's anchor,
   * where the canvas flips the geometry (never the label) and the workbook
   * writer emits `a:xfrm@flipH`.
   */
  flipH?: true;
  /** True for the empty picture slot of an imported Office picture layout. */
  smartArtPicturePlaceholder?: true;
  /**
   * The fill and outline are the AUTHOR's, not the catalog palette's — only on
   * a `node`, only ever `true`. Without it the export can only tell a chosen
   * colour from the palette by disagreement between the diagram's own nodes
   * (`SMARTART-IMPORT.md` §5.3), so a UNIFORM recolour reached Office without
   * its colours. Absent on an object stored before this, which is that rule
   * exactly as it was.
   */
  smartArtAuthorPaint?: true;
  /**
   * The label's COLOUR is the author's (Home > Font Color), not the colour
   * the palette gave it; only ever `true`. The text twin of
   * {@link smartArtAuthorPaint}: a rebuild keeps it and the export writes it
   * onto the data point. (Family, bold, italic, underline and strikethrough
   * need no marker — the engine never sets them, so their presence is the
   * author's statement.)
   */
  smartArtAuthorText?: true;
  /**
   * The label's SIZE is the author's (Home's size box, Grow / Shrink Font),
   * Office's `custT`: it stays out of the graphic's shared fit, a rebuild
   * keeps it and the export writes it. A colour alone never sets it.
   */
  smartArtAuthorTextSize?: true;
  /**
   * SmartArt Format > Change Shape on this one shape: the kind the AUTHOR
   * gave it, which IS the object's own `shape.kind`. A rebuild draws the item
   * in it again (`preserveSmartArtAuthorShape`), the export writes it as
   * `a:prstGeom` on the shape's presentation point, and Reset Shape / Reset
   * Graphic / a layout switch go back to the layout's shape. Only on a `node`.
   */
  smartArtAuthorShape?: SpreadsheetPresetShapeKind;
  /**
   * Picture Format on the picture in this slot (Crop's `a:srcRect`, the
   * Picture Styles pick: `lib/office-smartart/slot-picture-format`), on the
   * object beside its other markers. Only on a slot holding a picture.
   */
  smartArtPictureCrop?: CropRect;
  smartArtPictureStyle?: OfficePictureStyleId;
  /** An organization chart's assistant (Add Assistant), only ever `true`. */
  smartArtAssistant?: true;
  /**
   * An item its layout does not draw (Office's "unused text does not appear,
   * but remains available if you switch layouts"): kept in the outline, never
   * drawn, hit-tested or selected; the export writes its data point only.
   */
  smartArtHidden?: true;
  /** The item a title box belongs to (Name and Title), on a `connector` member. */
  smartArtTitleIndex?: number;
  /** The item a shape drawn for an item without its words belongs to (`model.ts`). */
  smartArtOwnerIndex?: number;
  /** An org chart box's Layout (Office's hierBranch), on its node (`model.ts`). */
  smartArtHierBranch?: "std" | "hang" | "l" | "r";
  /**
   * Where the member's words are laid out, relative to its own box (Office's
   * `dsp:txXfrm`, `lib/office-smartart/text-box`), when that is not the
   * box's middle: a Segmented Cycle wedge's box is the whole disc, and its
   * words sit in the wedge. Kept on the OBJECT beside `smartArtCustom`, so
   * Convert to Shapes and a paste, which drop every `smartArt…` key, drop it
   * with the rest of the diagram's identity. Absent: the whole box.
   */
  smartArtTextBox?: SmartArtTextBox;
  /**
   * The diagram runs Right to Left (Office's `dgm:dir="rev"`), on every member
   * like `smartArtKind`; left to right is stated by omission.
   */
  smartArtDirection?: "rtl";
  /** The Change Colors variation (Office's id), on every member like the kind. */
  smartArtColors?: string;
  /**
   * The Office ALIAS the diagram is drawn as (`smartArtLayoutUrn`,
   * `lib/office-smartart/native-layouts`), on every member like `smartArtKind`.
   */
  smartArtLayoutUrn?: string;
  smartArtKind?: string;
  smartArtRole?: string;
  smartArtItemId?: string;
  smartArtItemIndex?: number;
  smartArtLevel?: number;
}

/**
 * A sheet object's SmartArt text box ({@link SpreadsheetSmartArtMember.smartArtTextBox}),
 * or `null` for none or one that is not a text box (dropped, as an unreadable
 * `smartArtCustom` is: the words then fill the shape's box). The object's own
 * value is handed back, so a renderer sees the same reference every frame.
 */
export function spreadsheetSmartArtTextBoxOf(object: unknown): SmartArtTextBox | null {
  if (!object || typeof object !== "object") return null;
  const value = (object as { smartArtTextBox?: unknown }).smartArtTextBox;
  if (!readSmartArtTextBox(value)) return null;
  // Words at a size the author fixed break where Office breaks them, exactly
  // (SmartArt probe X24ui: 28 pt Georgia "gamm" / "a").
  return (object as { smartArtAuthorTextSize?: unknown }).smartArtAuthorTextSize === true
    ? { ...(value as SmartArtTextBox), exact: true }
    : (value as SmartArtTextBox);
}

function pixels(emu: number): number {
  return emu / EMU_PER_CSS_PIXEL;
}

function memberBounds(node: OfficeSmartArtNode): SpreadsheetSmartArtMemberBounds {
  return {
    // Not clamped at the frame's corner: a box the layout turns a quarter
    // (Vertical Block List's first bar, 233 x 62 pt turned upright) starts
    // above it and, clamped, was pushed down onto the second bar.
    left: pixels(node.left),
    top: pixels(node.top),
    width: Math.max(1, pixels(node.width)),
    height: Math.max(1, pixels(node.height)),
  };
}

/**
 * The sheet shape kind that draws this scene node. The scene speaks DrawingML
 * preset names, which the shared Office shape catalog already registers for the
 * worksheet, so most presets carry over unchanged.
 *
 * A node the engine drew as its OWN outline — a segmented cycle's wedge, a
 * step-up's tread, a funnel's vessel, a pyramid's band — names no preset, and
 * the sheet used to flatten every one of them to a rectangle. It now carries
 * the outline across instead, in the shape's own 0..100 box, so the worksheet
 * draws the same figure the slide does. A node claiming custom geometry but
 * carrying no usable path still falls back to the rectangle at its exact rect
 * rather than disappearing.
 */
function shapeGeometryForNode(node: OfficeSmartArtShapeNode): {
  kind: SpreadsheetObjectShapeKind;
  customPath?: PresentationCustomPath;
  customPoints?: string;
} {
  // The shared envelope first — the one the workbook writer emits as
  // `a:custGeom`, curves included, and the one Docs already carries. The
  // polygon-only `customPoints` was all this host stored, and the exporter
  // does not know it: every diagram with a wedge, a tread, a vessel or a band
  // refused its own export (`shape_kind_unsupported`).
  const customPath = normalizePresentationCustomPath(node.customPath);
  if (customPath) {
    return { kind: SPREADSHEET_CUSTOM_SHAPE_KIND, customPath };
  }
  const customPoints = customOutlinePoints(node);
  if (customPoints) {
    return { kind: SPREADSHEET_CUSTOM_SHAPE_KIND, customPoints };
  }
  return {
    kind: isSpreadsheetPresetShapeKind(node.shape)
      ? (node.shape as SpreadsheetPresetShapeKind)
      : "rectangle",
  };
}

/**
 * An engine path (the shape's own 0..100000 space, per
 * {@link OfficeSmartArtCustomPath}) as the sheet's own 0..100 point list.
 * Curves are not part of the engine's vocabulary, so every command is a corner.
 */
function customOutlinePoints(
  node: OfficeSmartArtShapeNode,
): string | undefined {
  const commands = node.customPath?.commands;
  if (!commands || commands.length < 3) return undefined;
  return normalizeSpreadsheetCustomPoints(
    commands
      .map((command) => `${command.x / 1000},${command.y / 1000}`)
      .join(" "),
  );
}

/** A turn in whole degrees, 0..359 (0 for none). */
function turnOf(rotation: number | undefined): number {
  return typeof rotation === "number" && Number.isFinite(rotation) ? ((Math.round(rotation) % 360) + 360) % 360 : 0;
}

/** A label's size as the px the shape renderer speaks: the catalog's points converted, the Office scene's px kept. */
function labelPixels(style: OfficeSmartArtTextNode["style"]): number | undefined {
  const fontSize = style.fontSize;
  if (!fontSize || !(fontSize > 0)) return undefined;
  if (!smartArtLabelSizeIsPx(style)) return Math.round(ptToPx(fontSize));
  // A scene states the engine's point size to the half pixel (34 pt =
  // 45.33 px, stated 45.5); the label takes the points' own pixels. Rounded
  // up to 46, Picture Lineup's "项目1" grew past the 86 pt text area it was
  // fitted to and broke onto two lines where Excel draws one.
  const points = Math.round(fontSize * 0.75 * 2) / 2;
  return Math.round(ptToPx(points) * 100) / 100;
}

/** The smallest label a squeezed member is given; below it the text is noise. */
const MIN_FITTED_LABEL_PX = 8;

/**
 * The sheet shape's own text box inside a member: its whole box less the body
 * insets `SpreadsheetShapeTextBody` draws with when the package states none
 * (0.1 in left and right, 0.05 in top and bottom). A SmartArt member lays its
 * words over its whole box, not its preset's `a:rect`.
 */
const SHEET_TEXT_INSET_X_PX = 9.6;
/** A text box that is the member's whole box, no insets of the sheet's own. */
const SMARTART_WHOLE_TEXT_BOX: SmartArtTextBox = { left: 0, top: 0, width: SMARTART_TEXT_BOX_UNIT, height: SMARTART_TEXT_BOX_UNIT };
const SHEET_TEXT_INSET_Y_PX = 4.8;
const SHEET_TEXT_LINE_HEIGHT = 1.2;

/** One sheet SmartArt member as it is drawn now, for {@link spreadsheetSmartArtDrawnTextScale}. */
export interface SpreadsheetSmartArtDrawnLabel {
  readonly text: string;
  /** The stored label size, px (the size the graphic was fitted to). */
  readonly textSize: number | undefined;
  /** The member's drawn box, px. */
  readonly width: number;
  readonly height: number;
  readonly textBox: SmartArtTextBox | null;
  /** A size set on Home (`smartArtAuthorTextSize`): Office leaves it as set. */
  readonly authored: boolean;
  /** The words' line spacing ({@link spreadsheetSmartArtLineSpacingOf}). */
  readonly lineSpacing?: number;
}

/**
 * The scale every fitted label of one graphic is drawn at, from the boxes its
 * shapes have NOW.
 *
 * Office keeps the text of a SmartArt graphic's shapes at one size (Microsoft
 * Support: make one shape smaller and "the text size decreases in all three
 * shapes"; make it larger again and the text grows back once every shape has
 * room), except a shape whose size the author customized. The sheet fits its
 * labels when the graphic is built (`fitSheetSmartArtLabels`) and stores that
 * size; a shape resized afterwards kept its words at that size and they ran
 * out of it. Measured with the same model, insets and floor as the build, so
 * a graphic nobody resized draws at 1 and the stored sizes are never rewritten.
 */
export function spreadsheetSmartArtDrawnTextScale(
  members: readonly SpreadsheetSmartArtDrawnLabel[],
): number {
  const labels = members.flatMap((member) => {
    const text = member.text.trim();
    const size = member.textSize;
    if (member.authored || !text || typeof size !== "number" || !(size > 0)) return [];
    if (!(member.width > 0) || !(member.height > 0)) return [];
    const box = member.textBox;
    const width = box ? (member.width * box.width) / SMARTART_TEXT_BOX_UNIT : member.width;
    const height = box ? (member.height * box.height) / SMARTART_TEXT_BOX_UNIT : member.height;
    // A label box of its own is Office's text area, its insets already off.
    return [{
      text,
      width: box ? width : width - 2 * SHEET_TEXT_INSET_X_PX,
      height: box ? height : height - 2 * SHEET_TEXT_INSET_Y_PX,
      fontSize: size,
      ...(member.lineSpacing !== undefined ? { lineHeight: SHEET_TEXT_LINE_HEIGHT * member.lineSpacing } : {}),
    }];
  });
  if (labels.length === 0) return 1;
  const largest = Math.max(...labels.map((label) => label.fontSize));
  return smartArtTextFitScale(labels, {
    lineHeight: SHEET_TEXT_LINE_HEIGHT,
    minScale: Math.min(1, MIN_FITTED_LABEL_PX / largest),
  });
}

/** A label size drawn at a graphic's {@link spreadsheetSmartArtDrawnTextScale}. */
export function spreadsheetSmartArtDrawnTextSize(size: number, scale: number): number {
  return scale >= 1 ? size : Math.max(MIN_FITTED_LABEL_PX, Math.floor(size * scale));
}

/**
 * Office's one size for the whole graphic, on the sheet's members.
 *
 * The engine hands every label the style's one size and lays the scene out in
 * any frame, so a sheet's small circles (a radial cycle's 40 px satellites)
 * wore 19 px words that wrapped into three lines and clipped. Office sizes a
 * graphic's text together: every shape's words shrink to the size the most
 * crowded shape can hold ({@link smartArtTextFitScale}), a word is never
 * broken, and the words never grow past the nominal size. Sizes the author set
 * on Home are applied afterwards by the rebuild and are not touched here.
 */
function fitSheetSmartArtLabels(
  members: SpreadsheetSmartArtMember[],
  nominalPx: ReadonlyMap<SpreadsheetSmartArtMember, number>,
): void {
  const labels = [...nominalPx.entries()].map(([member, fontSize]) => {
    // A member with a text box of its own fits its words to THAT rectangle
    // (a wedge's words to the wedge, not to the whole disc).
    const box = member.smartArtTextBox;
    const width = box
      ? (member.bounds.width * box.width) / SMARTART_TEXT_BOX_UNIT
      : member.bounds.width;
    const height = box
      ? (member.bounds.height * box.height) / SMARTART_TEXT_BOX_UNIT
      : member.bounds.height;
    // Spaced words stand their lines as the shape draws them.
    const spacing = spreadsheetSmartArtLineSpacingOf(member.shape);
    return {
      text: member.shape.text ?? "",
      width: box ? width : width - 2 * SHEET_TEXT_INSET_X_PX,
      height: box ? height : height - 2 * SHEET_TEXT_INSET_Y_PX,
      fontSize,
      ...(spacing !== undefined ? { lineHeight: SHEET_TEXT_LINE_HEIGHT * spacing } : {}),
    };
  });
  const largest = Math.max(0, ...nominalPx.values());
  const scale = smartArtTextFitScale(labels, {
    lineHeight: SHEET_TEXT_LINE_HEIGHT,
    minScale: largest > 0 ? Math.min(1, MIN_FITTED_LABEL_PX / largest) : undefined,
  });
  for (const member of members) {
    const nominal = nominalPx.get(member);
    if (nominal === undefined) continue;
    member.shape = normalizeShapeStyle({
      ...member.shape,
      // At 1 the size the layout gave (a scene's exact 45.33 px); shrunk, a
      // whole pixel under it.
      textSize: scale >= 1 ? nominal : Math.max(MIN_FITTED_LABEL_PX, Math.floor(nominal * scale)),
    });
  }
}

/**
 * The quick style's paint, carried onto the sheet shape UNCHANGED.
 *
 * The engine already resolved the fourteen styles into concrete member
 * properties (`lib/office-smartart/styles`, the contract's §4), so the sheet
 * host neither re-derives nor renames anything: the same diagram is the same
 * paint in a worksheet as in a slide. A style that uses none of these fields
 * (the flat default) adds no keys at all, so an unstyled diagram persists
 * exactly what it did before styles existed.
 */
function shapePaint(node: OfficeSmartArtShapeNode) {
  return {
    // A wash (a grouped list's plate, a Venn's overlapping discs) IS its
    // transparency; dropping it drew those members as solid blocks.
    ...(typeof node.opacity === "number" ? { opacity: node.opacity } : {}),
    ...(node.fillGradient ? { fillGradient: node.fillGradient } : {}),
    ...(node.shapeShadow ? { shapeShadow: node.shapeShadow } : {}),
    ...(node.shape3d ? { shape3d: node.shape3d } : {}),
    ...(node.scene3d ? { scene3d: node.scene3d } : {}),
    ...(node.smartArtStyle ? { smartArtStyle: node.smartArtStyle } : {}),
  };
}

/**
 * A LABEL's own share of the style. A sheet shape draws its label itself, so
 * these ride on the shape the label was merged into (or on the text box a
 * free-floating label becomes) rather than on a second object.
 */
function labelPaint(node: OfficeSmartArtTextNode, text?: string) {
  // A converted label's runs (a red phrase in a black label), when they are
  // its words exactly: the shape draws and the export writes them.
  const runs = (node as { smartArtAuthorText?: unknown; runs?: unknown }).smartArtAuthorText === true
    ? (node as { runs?: unknown }).runs
    : undefined;
  const trimmed = Array.isArray(runs) ? trimRunsLikeText(runs as TextRun[]) : null;
  const textRuns =
    trimmed && trimmed.length && text !== undefined && trimmed.map((run) => run.text).join("") === text
      ? trimmed
      : undefined;
  return {
    ...(node.textShape3d ? { textShape3d: node.textShape3d } : {}),
    ...(node.scene3d ? { scene3d: node.scene3d } : {}),
    ...(node.smartArtStyle ? { smartArtStyle: node.smartArtStyle } : {}),
    ...(textRuns ? { textRuns } : {}),
    ...labelLineSpacing(node, text),
  };
}

/**
 * The line spacing the author gave a label's words (a converted point's
 * `a:lnSpc`, `smartArtAuthorLineSpacing`), as the sheet shape keeps a
 * paragraph's spacing: one `textLayout` paragraph per line of its words, each
 * at that multiple, which the shape draws and the export writes back on the
 * data point (`authorTextLineSpacing`).
 */
function labelLineSpacing(node: OfficeSmartArtTextNode, text?: string): { textLayout?: OfficeTextLayout } {
  const spacing = node.style.lineSpacing;
  if (
    (node as { smartArtAuthorLineSpacing?: unknown }).smartArtAuthorLineSpacing !== true ||
    typeof spacing !== "number" ||
    !Number.isFinite(spacing) ||
    text === undefined
  ) {
    return {};
  }
  const lines = Math.min(text.split("\n").length, MAX_OFFICE_TEXT_LAYOUT_PARAGRAPHS);
  return {
    textLayout: {
      paragraphs: Array.from({ length: lines }, () => ({ level: 0, bullet: "none" as const, lineSpacing: spacing })),
    },
  };
}

/**
 * The line spacing a sheet SmartArt shape's words are drawn at, as a multiple
 * of the sheet's single line: its paragraphs' one spacing ({@link
 * labelLineSpacing}, or Paragraph's own), `undefined` when they state none or
 * differ. The graphic's shared size is fitted at it, as the shape draws it.
 */
export function spreadsheetSmartArtLineSpacingOf(shape: unknown): number | undefined {
  const paragraphs = (shape as { textLayout?: { paragraphs?: readonly { lineSpacing?: unknown }[] } } | null | undefined)
    ?.textLayout?.paragraphs;
  if (!Array.isArray(paragraphs) || !paragraphs.length) return undefined;
  const first = paragraphs[0]?.lineSpacing;
  if (typeof first !== "number" || !Number.isFinite(first) || first < 0.05 || first > 9.6) return undefined;
  return paragraphs.every((paragraph) => paragraph?.lineSpacing === first) ? first : undefined;
}

/**
 * Where the LAYOUT puts a node: its box with the author's move / resize
 * (`smartArtCustom`) taken back off. A rebuild re-applies an item's factors to
 * its shape but leaves its label where the layout puts it, so a label and the
 * moved shape it belongs to are compared, and its text box measured, where the
 * layout drew them both -- and the words then go wherever the shape went.
 */
function smartArtNodeLayoutBox(node: OfficeSmartArtNode) {
  const box = {
    left: node.left,
    top: node.top,
    width: node.width,
    height: node.height,
  };
  const factors = readSmartArtCustom(
    (node as { smartArtCustom?: unknown }).smartArtCustom,
  );
  return factors ? smartArtLayoutBox(box, factors) : box;
}

/**
 * Project a SmartArt scene onto sheet shape members.
 *
 * A scene keeps each item's TEXT in its own centered text node, which is how a
 * slide draws it. A sheet shape carries its own label, so a label is merged
 * into the shape of the same outline item -- its words, and its rectangle as
 * the shape's text box (`smartArtTextBox`) where that is not the shape's
 * middle -- and only a label with no shape of its own (or one belonging to a
 * picture member this host does not project) becomes a separate text box.
 * Connectors stay ordinary filled shapes, exactly as the engine authored them,
 * so they export at their exact rect.
 *
 * Every member keeps the quick style's paint as the engine resolved it —
 * gradient, bounded shadow, node 3-D, the diagram camera, the label's own
 * extrusion and the style marker — so the SAME diagram carries the SAME paint
 * in a worksheet as on a slide (`docs/office-graphics/SMARTART-STYLES.md` §4).
 *
 * An `image` node is projected as its PICTURE PLACEHOLDER: a plain shape at the
 * node's exact rect, carrying the same SmartArt identity as every other member.
 * The worksheet has no picture-custody path, so the insert gallery still does
 * not offer picture layouts (a newly authored diagram would have nothing to put
 * in the slot) — but an IMPORTED Office picture layout must still land as that
 * catalog layout with its full member set, and a placeholder is exactly what
 * Office draws for a picture point holding no picture
 * (docs/office-graphics/SMARTART-IMPORT.md §3).
 */
export function spreadsheetSmartArtMembers(
  nodes: readonly OfficeSmartArtNode[],
): SpreadsheetSmartArtMember[] {
  const shapeByItem = new Map<number, SpreadsheetSmartArtMember>();
  // The engine node behind each host, so a label can be checked against the
  // shape it would be merged into (Docs makes the same check).
  const hostNodeByItem = new Map<number, OfficeSmartArtNode>();
  const members: SpreadsheetSmartArtMember[] = [];
  const identity = (node: OfficeSmartArtNode) => ({
    ...(node.smartArtKind ? { smartArtKind: node.smartArtKind } : {}),
    ...(node.smartArtRole ? { smartArtRole: node.smartArtRole } : {}),
    ...(node.smartArtItemId ? { smartArtItemId: node.smartArtItemId } : {}),
    ...(node.smartArtItemIndex !== undefined
      ? { smartArtItemIndex: node.smartArtItemIndex }
      : {}),
    ...(node.smartArtLevel !== undefined
      ? { smartArtLevel: node.smartArtLevel }
      : {}),
    // Only a node has paint of its own to claim; a label is text over somebody
    // else's shape and a connector is the engine's own drawing.
    ...(node.smartArtRole === "node" && node.smartArtAuthorPaint === true
      ? { smartArtAuthorPaint: true as const }
      : {}),
    // The shape the author gave an item (Change Shape), which its node draws.
    ...(node.type === "shape" &&
    node.smartArtRole === "node" &&
    typeof node.smartArtAuthorShape === "string" &&
    isSpreadsheetPresetShapeKind(node.smartArtAuthorShape)
      ? { smartArtAuthorShape: node.smartArtAuthorShape as SpreadsheetPresetShapeKind }
      : {}),
    // An organization chart's assistant (`dgm:pt type="asst"`).
    ...(node.smartArtAssistant === true ? { smartArtAssistant: true as const } : {}),
    // An item the layout does not draw.
    ...(node.smartArtHidden === true &&
    (node.smartArtRole === "node" ||
      node.smartArtRole === "label" ||
      (node.smartArtRole === "connector" && typeof node.smartArtTitleIndex === "number"))
      ? { smartArtHidden: true as const }
      : {}),
    // The item a title box titles (Name and Title).
    ...(node.smartArtRole === "connector" && typeof node.smartArtTitleIndex === "number"
      ? { smartArtTitleIndex: node.smartArtTitleIndex }
      : {}),
    // The item a shape drawn for it without its words belongs to, which
    // Delete and Cut on that shape take out.
    ...(node.smartArtRole === "connector" && typeof node.smartArtOwnerIndex === "number"
      ? { smartArtOwnerIndex: node.smartArtOwnerIndex }
      : {}),
    // An org chart box's Layout, which the next relayout lays its
    // subordinates out by.
    ...(node.smartArtRole === "node" && node.smartArtHierBranch
      ? { smartArtHierBranch: node.smartArtHierBranch }
      : {}),
    // Right to Left, which every member of the diagram states.
    ...(node.smartArtDirection === "rtl"
      ? { smartArtDirection: "rtl" as const }
      : {}),
    // The Change Colors variation, which paints the next relayout.
    ...(typeof (node as { smartArtColors?: unknown }).smartArtColors === "string"
      ? { smartArtColors: (node as { smartArtColors?: string }).smartArtColors }
      : {}),
    // …and the Office alias it is drawn as, likewise.
    ...(typeof node.smartArtLayoutUrn === "string"
      ? { smartArtLayoutUrn: node.smartArtLayoutUrn }
      : {}),
    // A shape the author moved or resized keeps Office's factors for it
    // through a rebuild (`reapplySmartArtCustom`).
    ...((node as { smartArtCustom?: unknown }).smartArtCustom
      ? { smartArtCustom: (node as { smartArtCustom?: unknown }).smartArtCustom }
      : {}),
  });
  // The scene's layers are its paint order, as the Docs host reads them: a
  // shape the layout puts in front (cycle4's centre arrows, `zOrderOff` 16)
  // stood first in the scene's list and was drawn under the wedges.
  // An Office scene states where Office paints each shape (`smartArtPaint`,
  // the layout's own order): Bubble Picture List's rings each under their own
  // picture and over the one before, Hexagon Cluster's accents over theirs --
  // which the layers cannot say. MRAI's own layouts paint by their layers.
  const painted = nodes.some((node) => typeof node.smartArtPaint === "number");
  // A member that paints only words (no fill, no line) goes over the shapes:
  // it can hide nothing, and under them its words were hidden (Target List's
  // rows behind the white boxes the layout states after them).
  const wordsOnly = (node: OfficeSmartArtNode) =>
    node.type === "shape" &&
    (node.fill ?? "none") === "none" &&
    (!node.stroke || node.stroke === "none" || !(Number(node.strokeWidth) > 0));
  const ordered = [...nodes].sort((left, right) =>
    Number(wordsOnly(left)) - Number(wordsOnly(right)) ||
    (painted && typeof left.smartArtPaint === "number" && typeof right.smartArtPaint === "number"
      ? left.smartArtPaint - right.smartArtPaint
      : Number(left.zIndex ?? 0) - Number(right.zIndex ?? 0)),
  );
  for (const node of ordered) {
    if (node.type === "image") {
      // The picture IN the slot, when the node holds one. The engine's insert
      // placeholder is not a picture anybody chose, so it stays absent and the
      // slot keeps reading as empty; anything else is the author's or the
      // imported Office diagram's, and the sheet keeps it on the member's own
      // style so the canvas draws it and the export writes it back as a blip.
      const picture =
        node.src === SMART_ART_PICTURE_PLACEHOLDER_SRC
          ? null
          : smartArtPictureSrc(node.src);
      // An oval slot stays oval (Circular Picture Callout's pictures), a
      // rounded one rounded (Picture Caption List's roundRect slots).
      const slotKind =
        node.clipShape === "ellipse"
          ? "ellipse"
          : node.smartArtSlotShape === "hexagon"
            ? "hexagon"
            : node.smartArtSlotShape === "roundRect" || (node.cornerRadius ?? 0) > 0
              ? "rounded-rectangle"
              : "rectangle";
      // Office's own paint for the slot where the scene states it: the empty
      // slot filled as Excel prints it (Picture Caption List's slots are the
      // diagram's blue, not a pale box) and the layout's line round it.
      const officeFill =
        typeof node.smartArtSlotFill === "string" && /^#[0-9a-f]{6}$/i.test(node.smartArtSlotFill)
          ? node.smartArtSlotFill
          : null;
      const officeLine =
        typeof node.borderColor === "string" && /^#[0-9a-f]{6}$/i.test(node.borderColor)
          ? node.borderWidth && node.borderWidth > 0
            ? { stroke: node.borderColor, strokeWidth: node.borderWidth }
            : { stroke: "none", strokeWidth: 0 }
          : null;
      members.push({
        kind: slotKind,
        shape: normalizeShapeStyle({
          kind: slotKind,
          fill: officeFill ?? SMART_ART_PICTURE_PLACEHOLDER_FILL,
          // A slot wears the layout's own line round it (fgImgPlace1: the
          // white ring Excel draws round O17's pictures) where the scene
          // states one -- with its picture, or empty on the slot's own fill;
          // otherwise the neutral placeholder paint.
          ...((picture || officeFill) && officeLine
            ? officeLine
            : { stroke: SMART_ART_PICTURE_PLACEHOLDER_STROKE, strokeWidth: 1 }),
          ...(picture ? { src: picture } : {}),
          // The empty slot keeps its neutral paint, but still says which style
          // the diagram carries — the marker is on EVERY member (§4), so a
          // later restyle never has to infer one from how a member looks.
          ...(node.smartArtStyle ? { smartArtStyle: node.smartArtStyle } : {}),
        }),
        bounds: memberBounds(node),
        smartArtPicturePlaceholder: true,
        ...identity(node),
        // The picture's Picture Format, renamed from the slide block's own.
        ...(picture ? smartArtSlotPictureFields(smartArtSlotPictureFormatOfBlock(node)) : {}),
      });
      continue;
    }
    if (node.type !== "shape") continue;
    const { kind, customPath, customPoints } = shapeGeometryForNode(node);
    const member: SpreadsheetSmartArtMember = {
      kind,
      shape: normalizeShapeStyle({
        kind,
        ...(customPath ? { customPath } : {}),
        ...(customPoints ? { customPoints } : {}),
        fill: node.fill,
        stroke: node.stroke ?? node.fill,
        strokeWidth: node.strokeWidth ?? 0,
        // The layout's handles (Block Cycle's arcs each span their own
        // angles, a rounded box its corners): without them every arc drew
        // the preset's default half circle.
        ...(node.adjustments && Object.keys(node.adjustments).length ? { adjustments: { ...node.adjustments } } : {}),
        ...shapePaint(node),
      }),
      bounds: memberBounds(node),
      ...(typeof node.rotation === "number" && Number.isFinite(node.rotation) && node.rotation !== 0
        ? { rotation: node.rotation }
        : {}),
      ...(node.flipH === true ? { flipH: true as const } : {}),
      ...identity(node),
    };
    members.push(member);
    if (
      node.smartArtRole === "node" &&
      node.smartArtItemIndex !== undefined &&
      !shapeByItem.has(node.smartArtItemIndex)
    ) {
      shapeByItem.set(node.smartArtItemIndex, member);
      hostNodeByItem.set(node.smartArtItemIndex, node);
    }
  }
  // Each labelled member's nominal size, before the whole graphic is fitted.
  const nominalPx = new Map<SpreadsheetSmartArtMember, number>();
  for (const node of nodes) {
    if (node.type !== "text") continue;
    const text = node.text.trim();
    if (!text) continue;
    // Only a label drawn INSIDE its item's shape is that shape's text. A
    // caption the layout places beside or above its shape (a timeline's words
    // over its dots) stays a text box of its own; merged, it was squeezed into
    // a 12 px dot and clipped to nothing.
    const hostNode =
      node.smartArtItemIndex === undefined
        ? undefined
        : hostNodeByItem.get(node.smartArtItemIndex);
    const labelBox = smartArtNodeLayoutBox(node);
    const hostBox = hostNode ? smartArtNodeLayoutBox(hostNode) : undefined;
    // A hidden label (an item the layout does not draw) goes into its hidden
    // shape and never into a drawn one, which would paint its words.
    const hiddenLabel = node.smartArtHidden === true;
    const host =
      node.smartArtItemIndex === undefined ||
      !hostNode ||
      !hostBox ||
      (hostNode.smartArtHidden === true) !== hiddenLabel ||
      (!hiddenLabel && !smartArtLabelInsideHost({ ...node, ...labelBox }, { ...hostNode, ...hostBox }))
        ? undefined
        : shapeByItem.get(node.smartArtItemIndex);
    if (host) {
      host.shape = normalizeShapeStyle({
        ...host.shape,
        text,
        ...(node.style.color ? { textColor: node.style.color } : {}),
        // A shape centres its words; a label the layout aligns otherwise
        // (a bullet row flush left after its dot) says so.
        ...(node.style.align && node.style.align !== "center" && node.style.align !== "distribute"
          ? { textAlign: node.style.align }
          : {}),
        ...(node.style.verticalAlign && node.style.verticalAlign !== "middle"
          ? { textVerticalAlign: node.style.verticalAlign }
          : {}),
        ...labelPaint(node, text),
        // A shape the layout turns under a label it keeps level (Inverted
        // Pyramid's layers, turned a half turn) keeps its words upright, as
        // Excel draws them (`a:bodyPr@upright`); turned with the shape they
        // read upside down.
        ...(turnOf(host.rotation) !== 0 && turnOf((node as { rotation?: number }).rotation) === 0
          ? { textUpright: true as const }
          : {}),
      });
      // Only the words are merged; WHERE they sit is the label's own
      // rectangle (Office's `dsp:txXfrm`). A sheet shape centred its text in
      // its box, and a Segmented Cycle wedge's box is the whole disc, so all
      // five labels were drawn on top of each other in its middle.
      // Words kept level on a shape turned a quarter are laid in the box the
      // turned shape covers, its sides exchanged (the sheet's upright text),
      // so their rectangle is measured against that box.
      const hostTurn = turnOf(host.rotation);
      const levelOnTurned =
        ((hostTurn >= 45 && hostTurn < 135) || (hostTurn >= 225 && hostTurn < 315)) &&
        turnOf((node as { rotation?: number }).rotation) === 0;
      const reference = levelOnTurned && hostBox
        ? {
            left: hostBox.left + hostBox.width / 2 - hostBox.height / 2,
            top: hostBox.top + hostBox.height / 2 - hostBox.width / 2,
            width: hostBox.height,
            height: hostBox.width,
          }
        : hostBox;
      const textBox = reference
        ? smartArtLabelTextBox(labelBox, reference, {
            aligned: node.style.align === "left" || node.style.align === "right" || node.style.align === "justify",
            exact: smartArtLabelSizeIsPx(node.style),
          })
        : null;
      if (textBox) host.smartArtTextBox = textBox;
      else delete host.smartArtTextBox;
      const wanted = labelPixels(node.style);
      if (wanted !== undefined && !hiddenLabel) nominalPx.set(host, wanted);
      continue;
    }
    const caption: SpreadsheetSmartArtMember = {
      kind: "text-box",
      shape: normalizeShapeStyle({
        kind: "text-box",
        fill: "none",
        stroke: "none",
        text,
        ...(node.style.color ? { textColor: node.style.color } : {}),
        // A text box reads from the top left on the sheet; a diagram's label
        // box states its own alignment (Alternating Hexagons' words centred
        // in the hexagon they sit over, as Excel draws them).
        ...(node.style.align && node.style.align !== "distribute" ? { textAlign: node.style.align } : {}),
        ...(node.style.verticalAlign ? { textVerticalAlign: node.style.verticalAlign } : {}),
        ...labelPaint(node, text),
      }),
      bounds: memberBounds(node),
      // An Office scene's label box is the words' text area, its insets
      // already off: the caption draws them over all of it. With the sheet's
      // own 0.1 in insets on top, Spiral Picture's 27 px captions kept 8 px
      // for "项目1" and broke it into three lines Excel draws as one.
      ...(smartArtLabelSizeIsPx(node.style) ? { smartArtTextBox: SMARTART_WHOLE_TEXT_BOX } : {}),
      // A label the layout turns (Picture Accent List's titles, read bottom
      // to top up each bar) keeps its turn: without it the long, flat box lay
      // level across the bars and its words printed over them.
      ...(typeof node.rotation === "number" && Number.isFinite(node.rotation) && node.rotation !== 0
        ? { rotation: node.rotation }
        : {}),
      ...identity(node),
    };
    members.push(caption);
    const wanted = labelPixels(node.style);
    if (wanted !== undefined && !hiddenLabel) nominalPx.set(caption, wanted);
  }
  fitSheetSmartArtLabels(members, nominalPx);
  // A member that draws nothing -- no fill, no line, no words -- and is no
  // item's is left out (the frame-sized background Office's definitions hold
  // a process layout in): Excel's own drawing has no such shape, and as the
  // first member it took the diagram's first index and clicks meant for the
  // shapes. An item's unpainted shape stays, or an item whose words were
  // cleared lost its only member and the graphic stopped reading as one
  // outline (the Docs host keeps the same rule).
  return members.filter((member) => {
    if (member.smartArtPicturePlaceholder === true) return true;
    if (typeof member.smartArtItemId === "string") return true;
    const style = member.shape as { fill?: unknown; stroke?: unknown; text?: unknown };
    return style.fill !== "none" || style.stroke !== "none" || (typeof style.text === "string" && style.text.length > 0);
  });
}
