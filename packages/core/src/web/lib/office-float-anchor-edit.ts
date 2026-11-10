/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Word's Layout Options on one floating drawing's `wp:anchor` (`FloatAnchor`),
 * as pure functions: the Wrap Text choices, the Position gallery, Move with
 * Text / Fix Position on Page, a drag, Align and the floating stack steps.
 *
 * Every choice keeps the object where it is on the page unless the choice IS
 * a place (a Position preset, an Align): Word restates the offsets on the new
 * origin rather than letting the object jump. A drawing's place on its page is
 * read from what the float pass painted (`FloatPaint`), whose origins are the
 * points where offset 0 of each `relativeFrom` lands.
 */

import {
  EMU_PER_PX,
  FLOAT_RELATIVE_FROM_H,
  FLOAT_RELATIVE_FROM_V,
  placeFloatOnPage,
  resolveFloatPlacement,
  type FloatAlignH,
  type FloatAlignV,
  type FloatAnchor,
  type FloatPageGeometry,
  type FloatRelativeFromH,
  type FloatRelativeFromV,
  type FloatWrapText,
} from "./office-float-anchor";

/** Word's Wrap Text choices; In Line with Text has no anchor at all. */
export type FloatWrapChoice =
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

export const FLOAT_WRAP_CHOICES: readonly FloatWrapChoice[] = [
  "square",
  "tight",
  "through",
  "topAndBottom",
  "behindText",
  "inFrontOfText",
];

/** The Position gallery's nine "… with Square Text Wrapping" places. */
export type FloatPositionPreset =
  | "topLeft"
  | "topCenter"
  | "topRight"
  | "middleLeft"
  | "middleCenter"
  | "middleRight"
  | "bottomLeft"
  | "bottomCenter"
  | "bottomRight";

export const FLOAT_POSITION_PRESETS: readonly FloatPositionPreset[] = [
  "topLeft",
  "topCenter",
  "topRight",
  "middleLeft",
  "middleCenter",
  "middleRight",
  "bottomLeft",
  "bottomCenter",
  "bottomRight",
];

/** Word's default room beside a Square / Tight / Through object: 0.125 inch. */
const SIDE_DISTANCE_EMU = 114300;

/**
 * What the float pass measured for one placed drawing, in page CSS px: its
 * painted box, where offset 0 of every origin lands (so any axis can be
 * restated on any origin without the drawing moving), and the zoom.
 */
export interface FloatPaint {
  readonly leftPx: number;
  readonly topPx: number;
  readonly widthPx: number;
  readonly heightPx: number;
  readonly originsH: Readonly<Record<FloatRelativeFromH, number>>;
  readonly originsV: Readonly<Record<FloatRelativeFromV, number>>;
  /** Page px per client px is 1 / scale. */
  readonly scaleX: number;
  readonly scaleY: number;
  /** What it was measured against, to place another anchor of the same
   * drawing (`floatPaintFor`). */
  readonly geometry: FloatPageGeometry;
  readonly shift: { readonly leftPx: number; readonly topPx: number };
}

/** The pass's reading of one drawing: its painted box and every origin. */
export function measureFloatPaint(
  anchor: FloatAnchor,
  geometry: FloatPageGeometry,
  scale: { x: number; y: number },
  shift: { leftPx: number; topPx: number } = { leftPx: 0, topPx: 0 },
): FloatPaint {
  // The painted box is kept on the page as Word keeps it; the origins are
  // where offset 0 of each lands, before any such keeping.
  const placement = placeFloatOnPage({ ...anchor, groupTransform: undefined }, geometry);
  const origin = <H extends FloatRelativeFromH, V extends FloatRelativeFromV>(h: H, v: V) =>
    resolveFloatPlacement(
      {
        ...anchor,
        horizontal: { relativeFrom: h, offsetEmu: 0 },
        vertical: { relativeFrom: v, offsetEmu: 0 },
        groupTransform: undefined,
        simplePos: false,
      },
      geometry,
    );
  const originsH = Object.fromEntries(
    FLOAT_RELATIVE_FROM_H.map((h) => [h, origin(h, "page").leftPx + shift.leftPx]),
  ) as Record<FloatRelativeFromH, number>;
  const originsV = Object.fromEntries(
    FLOAT_RELATIVE_FROM_V.map((v) => [v, origin("page", v).topPx + shift.topPx]),
  ) as Record<FloatRelativeFromV, number>;
  return {
    leftPx: placement.leftPx + shift.leftPx,
    topPx: placement.topPx + shift.topPx,
    widthPx: placement.widthPx,
    heightPx: placement.heightPx,
    originsH,
    originsV,
    scaleX: scale.x || 1,
    scaleY: scale.y || 1,
    geometry,
    shift,
  };
}

/** The same measurement for another anchor of the drawing (a shape whose
 * live offsets are not in the anchor the pass placed it by). */
export function floatPaintFor(anchor: FloatAnchor, paint: FloatPaint): FloatPaint {
  return measureFloatPaint(anchor, paint.geometry, { x: paint.scaleX, y: paint.scaleY }, paint.shift);
}

/** The Wrap Text choice an anchor states. */
export function floatWrapChoiceOf(anchor: FloatAnchor): FloatWrapChoice {
  switch (anchor.wrap.kind) {
    case "none":
      return anchor.behindDoc ? "behindText" : "inFrontOfText";
    case "topAndBottom":
      return "topAndBottom";
    default:
      return anchor.wrap.kind;
  }
}

/**
 * The anchor restated for one Wrap Text choice. Square / Tight / Through keep
 * the side the text may use, and a Tight / Through outline the file drew
 * stays the outline; Behind / In Front of Text are `wrapNone` told apart by
 * `behindDoc`. Word gives a wrapped object 0.125" of room beside it and an
 * overlay none.
 */
export function floatAnchorWithWrap(anchor: FloatAnchor, choice: FloatWrapChoice): FloatAnchor {
  const previous = anchor.wrap;
  const text = previous.text ?? "bothSides";
  const wrap: FloatAnchor["wrap"] =
    choice === "square" || choice === "tight" || choice === "through"
      ? {
          kind: choice,
          text,
          ...(choice !== "square" && previous.polygon ? { polygon: previous.polygon } : {}),
          ...(choice !== "square" && previous.polygonEdited ? { polygonEdited: true } : {}),
        }
      : choice === "topAndBottom"
        ? { kind: "topAndBottom" }
        : { kind: "none" };
  const sided = choice === "square" || choice === "tight" || choice === "through";
  return {
    ...anchor,
    wrap,
    behindDoc: choice === "behindText",
    distLEmu: sided && anchor.distLEmu === 0 ? SIDE_DISTANCE_EMU : anchor.distLEmu,
    distREmu: sided && anchor.distREmu === 0 ? SIDE_DISTANCE_EMU : anchor.distREmu,
  };
}

/**
 * Word's anchor for a drawing it takes out of the line (measured, Word 16.86):
 * from the column and the paragraph where it stood, Square's distances,
 * unlocked, laid out in a cell, overlap allowed, above every other floating
 * drawing -- with the chosen wrap.
 */
export function floatAnchorOutOfLine(input: {
  readonly widthEmu: number;
  readonly heightEmu: number;
  /** From the column's left edge to the drawing's. */
  readonly columnOffsetEmu: number;
  /** From the paragraph's top to the drawing's. */
  readonly paragraphOffsetEmu: number;
  readonly relativeHeight: number;
  readonly wrap: FloatWrapChoice;
}): FloatAnchor {
  return floatAnchorWithWrap(
    {
      version: 1,
      horizontal: { relativeFrom: "column", offsetEmu: Math.round(input.columnOffsetEmu) },
      vertical: { relativeFrom: "paragraph", offsetEmu: Math.round(input.paragraphOffsetEmu) },
      widthEmu: Math.round(input.widthEmu),
      heightEmu: Math.round(input.heightEmu),
      wrap: { kind: "square", text: "bothSides" },
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: true,
      hidden: false,
      relativeHeight: input.relativeHeight,
      simplePos: false,
      distTEmu: 0,
      distBEmu: 0,
      distLEmu: SIDE_DISTANCE_EMU,
      distREmu: SIDE_DISTANCE_EMU,
      effectLEmu: 0,
      effectTEmu: 0,
      effectREmu: 0,
      effectBEmu: 0,
      rotation: 0,
      flipH: false,
      flipV: false,
    },
    input.wrap,
  );
}

/**
 * The Position gallery: "Position in Top Left with Square Text Wrapping" and
 * its eight siblings. Word aligns the object to the margins on both axes and
 * wraps text squarely round it; the alignment is the position, so the object
 * keeps it whatever the text does.
 */
export function floatAnchorAtPosition(anchor: FloatAnchor, preset: FloatPositionPreset): FloatAnchor {
  const vertical = preset.startsWith("top") ? "top" : preset.startsWith("middle") ? "center" : "bottom";
  const horizontal = preset.endsWith("Left") ? "left" : preset.endsWith("Center") ? "center" : "right";
  return {
    ...floatAnchorWithWrap(anchor, "square"),
    horizontal: { relativeFrom: "margin", align: horizontal },
    vertical: { relativeFrom: "margin", align: vertical },
    simplePos: false,
  };
}

/** Which Position preset an anchor states, or null for any other place. */
export function floatPositionPresetOf(anchor: FloatAnchor): FloatPositionPreset | null {
  if (anchor.wrap.kind !== "square") return null;
  if (anchor.horizontal.relativeFrom !== "margin" || anchor.vertical.relativeFrom !== "margin") return null;
  const horizontal: Readonly<Record<string, string>> = { left: "Left", center: "Center", right: "Right" };
  const vertical: Readonly<Record<string, string>> = { top: "top", center: "middle", bottom: "bottom" };
  const h = horizontal[anchor.horizontal.align ?? ""];
  const v = vertical[anchor.vertical.align ?? ""];
  return h && v ? (`${v}${h}` as FloatPositionPreset) : null;
}

/**
 * Align (one object, Word's "Align to Margin"): the axis becomes an alignment
 * in the margins; the other axis is left as it is.
 */
export function floatAnchorAligned(
  anchor: FloatAnchor,
  align: "left" | "center" | "right" | "top" | "middle" | "bottom",
): FloatAnchor {
  if (align === "left" || align === "center" || align === "right") {
    return { ...anchor, horizontal: { relativeFrom: "margin", align }, simplePos: false };
  }
  return {
    ...anchor,
    vertical: { relativeFrom: "margin", align: align === "middle" ? "center" : align },
    simplePos: false,
  };
}

/** Move with Text: the object's vertical origin is its paragraph (or line). */
export function floatAnchorMovesWithText(anchor: FloatAnchor): boolean {
  return anchor.vertical.relativeFrom === "paragraph" || anchor.vertical.relativeFrom === "line";
}

/**
 * Move with Text / Fix Position on Page. Word changes the VERTICAL origin only
 * (paragraph, or page) and restates the offset so the object stays where it
 * is on the page; the horizontal axis is untouched.
 */
export function floatAnchorWithMoveWithText(
  anchor: FloatAnchor,
  paint: FloatPaint,
  moveWithText: boolean,
): FloatAnchor {
  if (floatAnchorMovesWithText(anchor) === moveWithText) return anchor;
  const relativeFrom: FloatRelativeFromV = moveWithText ? "paragraph" : "page";
  return {
    ...anchor,
    vertical: {
      relativeFrom,
      offsetEmu: Math.round((paint.topPx - paint.originsV[relativeFrom]) * EMU_PER_PX),
    },
    simplePos: false,
  };
}

/**
 * The anchor after the drawing is dragged by `(dxPx, dyPx)` page px. Word
 * turns a dragged aligned or relative position into an absolute offset on the
 * same origin; an axis already stated by offset just moves.
 */
export function floatAnchorMovedBy(
  anchor: FloatAnchor,
  paint: FloatPaint,
  dxPx: number,
  dyPx: number,
): FloatAnchor {
  const h = anchor.horizontal.relativeFrom;
  const v = anchor.vertical.relativeFrom;
  return {
    ...anchor,
    horizontal:
      dxPx === 0
        ? anchor.horizontal
        : { relativeFrom: h, offsetEmu: Math.round((paint.leftPx + dxPx - paint.originsH[h]) * EMU_PER_PX) },
    vertical:
      dyPx === 0
        ? anchor.vertical
        : { relativeFrom: v, offsetEmu: Math.round((paint.topPx + dyPx - paint.originsV[v]) * EMU_PER_PX) },
    simplePos: false,
  };
}

/** One step of the floating stack (Word's `relativeHeight`): the object goes
 * just above the next object up, or just below the next one down; Front and
 * Back go past every other one. */
export function floatStackHeightAfter(
  current: number,
  heights: readonly number[],
  step: "forward" | "backward" | "front" | "back",
): number | null {
  const others = heights.filter((height) => height !== current);
  if (step === "front" || step === "forward") {
    const above = others.filter((height) => height > current);
    if (above.length === 0) return null;
    return Math.min(0xffff_ffff, (step === "front" ? Math.max(...above) : Math.min(...above)) + 1);
  }
  const below = others.filter((height) => height < current);
  if (below.length === 0) return null;
  return Math.max(0, (step === "back" ? Math.min(...below) : Math.max(...below)) - 1);
}

// ---------------------------------------------------------------------------
// The Layout dialog (More Layout Options): Position and Text Wrapping.
// ---------------------------------------------------------------------------

/** How one axis states its place: Alignment, Book layout, Absolute position
 * or Relative position (percent), as the Layout dialog's four radios. */
export type FloatAxisMode = "align" | "book" | "absolute" | "relative";

export interface FloatLayoutAxisForm<Align extends string, Origin extends string> {
  readonly mode: FloatAxisMode;
  readonly align: Align;
  readonly relativeFrom: Origin;
  /** Absolute position, EMU. */
  readonly offsetEmu: number;
  /** Relative position, percent of the origin's extent. */
  readonly percent: number;
}

export interface FloatLayoutForm {
  readonly horizontal: FloatLayoutAxisForm<FloatAlignH, FloatRelativeFromH>;
  readonly vertical: FloatLayoutAxisForm<FloatAlignV, FloatRelativeFromV>;
  readonly locked: boolean;
  readonly allowOverlap: boolean;
  readonly layoutInCell: boolean;
  readonly wrap: FloatWrapChoice;
  readonly wrapText: FloatWrapText;
  readonly distTEmu: number;
  readonly distBEmu: number;
  readonly distLEmu: number;
  readonly distREmu: number;
}

/** The origins each axis offers per mode, in the Layout dialog's order. */
export const FLOAT_LAYOUT_ORIGINS_H: Readonly<Record<FloatAxisMode, readonly FloatRelativeFromH[]>> = {
  align: ["margin", "page", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"],
  book: ["margin", "page"],
  absolute: ["margin", "page", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"],
  relative: ["margin", "page", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"],
};
export const FLOAT_LAYOUT_ORIGINS_V: Readonly<Record<FloatAxisMode, readonly FloatRelativeFromV[]>> = {
  align: ["margin", "page", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"],
  book: ["margin", "page"],
  absolute: ["margin", "page", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"],
  relative: ["margin", "page", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"],
};

/** Where the drawing stands from `relativeFrom`'s offset 0, in EMU, as the
 * float pass painted it; null when it was not painted. */
export function floatPaintedOffsetEmu(
  paint: FloatPaint | null,
  axis: "horizontal" | "vertical",
  relativeFrom: string,
): number | null {
  if (!paint) return null;
  const origin =
    axis === "horizontal"
      ? paint.originsH[relativeFrom as FloatRelativeFromH]
      : paint.originsV[relativeFrom as FloatRelativeFromV];
  if (typeof origin !== "number") return null;
  const placed = axis === "horizontal" ? paint.leftPx : paint.topPx;
  return Math.round((placed - origin) * EMU_PER_PX);
}

/** The Layout dialog's reading of an anchor (with its painted place, so an
 * aligned axis shows the absolute position it stands at). */
export function floatLayoutFormOf(anchor: FloatAnchor, paint: FloatPaint | null): FloatLayoutForm {
  const axisForm = <Align extends string, Origin extends string>(
    axis: { relativeFrom: Origin; align?: Align; offsetEmu?: number; percentOffset?: number },
    which: "horizontal" | "vertical",
    defaultAlign: Align,
  ): FloatLayoutAxisForm<Align, Origin> => {
    const painted = floatPaintedOffsetEmu(paint, which, axis.relativeFrom);
    const offsetEmu = axis.offsetEmu ?? painted ?? 0;
    if (typeof axis.percentOffset === "number") {
      return { mode: "relative", align: defaultAlign, relativeFrom: axis.relativeFrom, offsetEmu, percent: axis.percentOffset / 1000 };
    }
    if (axis.align) {
      const book = axis.align === "inside" || axis.align === "outside";
      return { mode: book ? "book" : "align", align: axis.align, relativeFrom: axis.relativeFrom, offsetEmu, percent: 0 };
    }
    return { mode: "absolute", align: defaultAlign, relativeFrom: axis.relativeFrom, offsetEmu, percent: 0 };
  };
  return {
    horizontal: axisForm<FloatAlignH, FloatRelativeFromH>(anchor.horizontal, "horizontal", "left"),
    vertical: axisForm<FloatAlignV, FloatRelativeFromV>(anchor.vertical, "vertical", "top"),
    locked: anchor.locked,
    allowOverlap: anchor.allowOverlap,
    layoutInCell: anchor.layoutInCell,
    wrap: floatWrapChoiceOf(anchor),
    wrapText: anchor.wrap.text ?? "bothSides",
    distTEmu: anchor.distTEmu,
    distBEmu: anchor.distBEmu,
    distLEmu: anchor.distLEmu,
    distREmu: anchor.distREmu,
  };
}

/** The anchor the Layout dialog's OK writes. */
export function floatAnchorFromLayoutForm(anchor: FloatAnchor, form: FloatLayoutForm): FloatAnchor {
  const axis = <Align extends string, Origin extends string>(
    value: FloatLayoutAxisForm<Align, Origin>,
  ): { relativeFrom: Origin; align?: Align; offsetEmu?: number; percentOffset?: number } =>
    value.mode === "align" || value.mode === "book"
      ? { relativeFrom: value.relativeFrom, align: value.align }
      : value.mode === "relative"
        ? {
            relativeFrom: value.relativeFrom,
            percentOffset: Math.round(value.percent * 1000),
            offsetEmu: Math.round(value.offsetEmu),
          }
        : { relativeFrom: value.relativeFrom, offsetEmu: Math.round(value.offsetEmu) };
  const wrapped = floatAnchorWithWrap(anchor, form.wrap);
  const sided = form.wrap === "square" || form.wrap === "tight" || form.wrap === "through";
  const overlay = form.wrap === "behindText" || form.wrap === "inFrontOfText";
  return {
    ...wrapped,
    horizontal: axis(form.horizontal),
    vertical: axis(form.vertical),
    wrap: sided ? { ...wrapped.wrap, text: form.wrapText } : wrapped.wrap,
    locked: form.locked,
    allowOverlap: form.allowOverlap,
    layoutInCell: form.layoutInCell,
    distTEmu: overlay ? 0 : Math.max(0, Math.round(form.distTEmu)),
    distBEmu: overlay ? 0 : Math.max(0, Math.round(form.distBEmu)),
    distLEmu: overlay ? 0 : Math.max(0, Math.round(form.distLEmu)),
    distREmu: overlay ? 0 : Math.max(0, Math.round(form.distREmu)),
    simplePos: false,
  };
}
