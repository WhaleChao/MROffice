/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Docs SmartArt frame and label size: the constants an inserted diagram is
 * sized by and the one label size an authored diagram's members are drawn at.
 *
 * The browser needs these to DRAW a stored diagram (the node view sizes its
 * labels on every render); the projection of a new layout onto the Docs shape
 * group is the server's (`./document-shapes`, run by the engine host), which
 * re-exports everything here.
 */

import { ptToPx } from "../office-units";
import { SMARTART_TEXT_BOX_UNIT } from "./text-box";
import {
  DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_SIZE_PT,
  DOCUMENT_SHAPE_MAX_EDGE_EMU,
  type DocumentShapeAttrs,
  type DocumentShapeGroupMember,
} from "../document-shapes";
import { DEFAULT_SMARTART_STYLE, OFFICE_SCENE_LABEL_LINE_SPACING, type SmartArtStyle } from "./model";
import { smartArtTextFitScale, smartArtWrappedLineCount } from "./text-fit";

/** EMU per inch. */
const EMU_PER_INCH = 914_400;

/**
 * The frame a fresh diagram occupies on the page: 6 in x 3.5 in, which fits
 * inside Word's default 6.5 in text column with room for the wrap.
 */
export const DOCUMENT_SMARTART_FRAME_EMU = Object.freeze({
  width: Math.round(6 * EMU_PER_INCH),
  height: Math.round(3.5 * EMU_PER_INCH),
});

/**
 * The group's child coordinate space, as a multiple of the frame. It exists
 * only so the 16 px minimum member edge lands at ~1.5 pt once the group is
 * drawn at its page size; nothing about the picture changes with it.
 */
export const DOCUMENT_SMARTART_COORDINATE_SCALE = 8;

export const DOCUMENT_SMARTART_COORDINATE_EMU = Object.freeze({
  width: DOCUMENT_SMARTART_FRAME_EMU.width * DOCUMENT_SMARTART_COORDINATE_SCALE,
  height:
    DOCUMENT_SMARTART_FRAME_EMU.height * DOCUMENT_SMARTART_COORDINATE_SCALE,
});

/**
 * The child coordinate space for ONE diagram frame, and the multiple it uses.
 *
 * The multiple is capped so the space itself stays a legal shape edge. Word's
 * own default SmartArt frame is 7 in wide, and 7 in times the flat 8 is
 * 51_206_400 EMU — past {@link DOCUMENT_SHAPE_MAX_EDGE_EMU}. Every Word
 * diagram therefore used to be refused: the preserved projection reported
 * `frame_out_of_range` and the converted one lost its members to the
 * coordinate clamp, which is why twenty of the thirty-nine catalog layouts
 * came back preserved from a real Word file
 * (`docs/office-graphics/SMARTART-IMPORT-AUDIT.md` §3 R1).
 *
 * A frame that fits keeps the flat 8, so every diagram MRAI has already
 * written keeps the coordinates it has; only a frame too large for it steps
 * down, and never below 1.
 */
export function documentSmartArtCoordinateSpace(frame: {
  readonly widthEmu: number;
  readonly heightEmu: number;
}): { width: number; height: number; scale: number } | null {
  const edge = Math.max(frame.widthEmu, frame.heightEmu);
  if (
    !Number.isFinite(frame.widthEmu) ||
    !Number.isFinite(frame.heightEmu) ||
    frame.widthEmu < 1 ||
    frame.heightEmu < 1 ||
    edge > DOCUMENT_SHAPE_MAX_EDGE_EMU
  ) {
    return null;
  }
  const scale = Math.max(
    1,
    Math.min(
      DOCUMENT_SMARTART_COORDINATE_SCALE,
      Math.floor(DOCUMENT_SHAPE_MAX_EDGE_EMU / edge),
    ),
  );
  return {
    width: Math.round(frame.widthEmu * scale),
    height: Math.round(frame.heightEmu * scale),
    scale,
  };
}

/**
 * The palette a fresh diagram is built with: the SHARED one, the same
 * `DEFAULT_SMARTART_STYLE` a slide and a worksheet build with, and the same
 * one the import conversion already uses (`native-convert`).
 *
 * It used to be Word's own shape palette, because a member label could only be
 * painted in the document text colour and PowerPoint's #4472c4 node would have
 * put a dark label below 4.5:1. A member now carries the label's OWN colour
 * ({@link smartArtLabelPaint}), so that reason is gone — and keeping it would
 * mean a diagram the author inserts and the SAME diagram imported from Word
 * were two different pictures, which is exactly what this feature must not do.
 */
export const DOCUMENT_SMARTART_STYLE: SmartArtStyle = DEFAULT_SMARTART_STYLE;

/** CSS px per EMU at the page's own scale (96 dpi). */
const PX_PER_EMU = 1 / 9525;

/**
 * Where a Docs member draws its words inside its box, as fractions: the
 * `.text` rule's `inset: 18% 14%` and its line height. The fit below measures
 * the very box the canvas paints in.
 */
export const DOCUMENT_SMARTART_LABEL_INSET = Object.freeze({ x: 0.14, y: 0.18 });
export const DOCUMENT_SMARTART_LABEL_LINE_HEIGHT = 1.25;

/**
 * The one size, in page CSS px, that an authored diagram's member labels are
 * drawn at -- or `null` for a group that is not one (a preserved Office
 * diagram draws the sizes its package states).
 *
 * Every member used to draw a flat 14 px whatever its size, which is 10.5 pt
 * where Word's own inserted diagrams carry 24 to 36 pt, and which wrapped
 * "Item 2" into two lines in a Radial Cycle's 56 px circles while the hub kept
 * one. Office sizes a graphic's text together (Microsoft Support, "Learn more
 * about SmartArt graphics": when one shape's text shrinks, all the others
 * shrink to the same size), from the layout's largest size down to what the
 * most crowded shape holds without breaking a word. MRAI's largest is the
 * shared style's size, the one the deck draws; the fit is
 * {@link smartArtTextFitScale}, the sheet's and the deck's too. It is worked
 * out from the group's CURRENT frame, so a diagram resized smaller takes its
 * words down with it, as Word's does, and nothing is stored: Word re-fits the
 * graphic itself when it opens the file.
 */
/** Office's default `a:bodyPr` insets in CSS px (91440 / 45720 EMU at 9525 EMU per px). */
const OFFICE_BODY_INSET_X_PX = 91_440 / 9_525;
const OFFICE_BODY_INSET_Y_PX = 45_720 / 9_525;

export function documentSmartArtLabelFontPx(attrs: {
  readonly smartArtKind?: string | null;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly groupCoordinateWidthEmu?: number | null;
  readonly groupCoordinateHeightEmu?: number | null;
  readonly groupMembers?: readonly DocumentShapeGroupMember[] | null;
}): number | null {
  const members = attrs.groupMembers;
  if (typeof attrs.smartArtKind !== "string" || !members?.length) return null;
  const childWidth = attrs.groupCoordinateWidthEmu ?? attrs.widthEmu;
  const childHeight = attrs.groupCoordinateHeightEmu ?? attrs.heightEmu;
  if (!(childWidth > 0) || !(childHeight > 0)) return null;
  const scaleX = (attrs.widthEmu / childWidth) * PX_PER_EMU;
  const scaleY = (attrs.heightEmu / childHeight) * PX_PER_EMU;
  const nominal = ptToPx(DOCUMENT_SMARTART_STYLE.fontSize);
  // A member whose size the author set on Home keeps it and takes no part in
  // the graphic's one size (Microsoft Support: "if you have previously
  // customized the text size in any shape, then the text in that shape will
  // not resize"); shrinking it must not take the other shapes' words down.
  const authored = (member: DocumentShapeGroupMember) =>
    member.smartArtAuthorTextSize === true ||
    (typeof member.textFontSizePt === "number" &&
      member.textFontSizePt !== DOCUMENT_SHAPE_DEFAULT_TEXT_FONT_SIZE_PT);
  const labels = members
    .filter(
      (member) =>
        member.text.trim() &&
        !member.textBody &&
        !authored(member) &&
        // An item the layout does not draw has no box to fit.
        member.smartArtHidden !== true,
    )
    .map((member) => {
      // A member with a text box of its own fits its words to THAT rectangle
      // (a Segmented Cycle wedge's words to the wedge, not to the whole disc
      // its box spans); the node view draws them there with no further inset.
      // Office lays the words out inside that box less its body insets
      // (`a:bodyPr` defaults, 0.1 in left and right, 0.05 in top and bottom),
      // so they are fitted to the inset box and drawn on the whole box: the
      // slack is what keeps a fit made by estimate from wrapping the last word
      // onto a line the wedge clips.
      // Words the author spaced (`smartArtLineSpacing`) stand their lines as
      // the member draws them: the host's line at that share of PowerPoint's 90%.
      const spacing = member.smartArtLineSpacing;
      const spaced =
        typeof spacing === "number" && Number.isFinite(spacing) && spacing >= 0.05 && spacing <= 9.6
          ? { lineHeight: (DOCUMENT_SMARTART_LABEL_LINE_HEIGHT * spacing) / OFFICE_SCENE_LABEL_LINE_SPACING }
          : {};
      const box = member.smartArtTextBox;
      if (box) {
        const boxWidth = (member.widthEmu * scaleX * box.width) / SMARTART_TEXT_BOX_UNIT;
        const boxHeight = (member.heightEmu * scaleY * box.height) / SMARTART_TEXT_BOX_UNIT;
        return {
          text: member.text,
          width: Math.max(boxWidth * 0.5, boxWidth - 2 * OFFICE_BODY_INSET_X_PX),
          height: Math.max(boxHeight * 0.5, boxHeight - 2 * OFFICE_BODY_INSET_Y_PX),
          fontSize: nominal,
          ...spaced,
        };
      }
      // The words are drawn in exactly this inset box, so the estimate is
      // fitted half a character narrower: line breaking (East Asian
      // first/last-character rules among them) can carry a character to a
      // line the estimate did not count, and the box then clipped its last
      // line (WUA 209: 27 characters estimated at four lines, drawn on five).
      // Words on one line break nowhere, so they keep the whole inset box: Word
      // draws Basic Block List's "项目 1" at its full size on one line.
      const inner = member.widthEmu * scaleX * (1 - 2 * DOCUMENT_SMARTART_LABEL_INSET.x);
      const oneLine = smartArtWrappedLineCount(member.text, nominal, inner) === 1;
      return {
        text: member.text,
        width: oneLine ? inner : Math.max(inner * 0.5, inner - nominal * 0.5),
        height: member.heightEmu * scaleY * (1 - 2 * DOCUMENT_SMARTART_LABEL_INSET.y),
        fontSize: nominal,
        ...spaced,
      };
    });
  if (labels.length === 0) return null;
  const scale = smartArtTextFitScale(labels, {
    lineHeight: DOCUMENT_SMARTART_LABEL_LINE_HEIGHT,
    // 6 pt, the engine's own floor for a fitted label.
    minScale: 8 / nominal,
  });
  return Math.round(nominal * scale * 10) / 10;
}

export type DocumentSmartArtProjectionFailure =
  /** The id is not in the shared catalog. */
  | "unknown-layout"
  /** A picture layout: its picture members have no native group equivalent. */
  | "picture-layout"
  /** The engine produced nothing a native group can hold. */
  | "empty-scene"
  /** More shapes than one native group may carry. */
  | "too-many-members";

export type DocumentSmartArtProjection =
  | { readonly ok: true; readonly attrs: DocumentShapeAttrs; readonly memberCount: number }
  | {
      readonly ok: false;
      readonly reason: DocumentSmartArtProjectionFailure;
      readonly memberCount: number;
    };
