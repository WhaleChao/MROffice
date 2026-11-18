/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { normalizeOfficeTextWarp, type OfficeTextWarp } from "./office-text-warp";
/**
 * Office floating-object anchors — one model for every host.
 *
 * A DrawingML object in a Word document is either INLINE (`wp:inline`: it sits
 * in its paragraph like a glyph) or ANCHORED (`wp:anchor`: it floats, and the
 * body text either flows around it or ignores it). This module is the anchored
 * half: the exact reading of `wp:anchor` — its two position axes, its extent,
 * its wrap element and its z-order — plus the arithmetic that turns that
 * reading into a CSS placement once a real page geometry is known.
 *
 * It is deliberately host-agnostic. Docs uses it for Word text boxes today;
 * `documentShape` (which already carries `wrapMode` and `relativeHeight` but
 * lays out as an in-flow block), the chart host's `DocumentChartLayout`, and
 * the image node's `dataFloat` all describe SUBSETS of the same Word model and
 * can adopt it without a second design.
 *
 * Reference: ECMA-376 Part 1 §20.4 (`wp:` DrawingML-WordprocessingDrawing) and
 * the Open XML SDK reference for `wp:anchor`, `wp:positionH`/`positionV`,
 * `wp:extent`, `wp:effectExtent`, `wp:wrapNone`/`wrapSquare`/`wrapTight`/
 * `wrapThrough`/`wrapTopAndBottom` and `wp:wrapPolygon`. The semantics MRAI
 * reproduces are recorded in
 * `docs/microsoft-office-replication/WORD-FLOATING-TEXT-BOX.md`.
 */

/** Office's length units (one source: office-units). */
import { EMU_PER_PX, EMU_PER_INCH, EMU_PER_PT, EMU_PER_CM, emuToPx, pxToEmu, emuToPt } from "./office-units";
export { EMU_PER_PX, EMU_PER_INCH, EMU_PER_PT, EMU_PER_CM, emuToPx, pxToEmu, emuToPt };

/** `wp:positionH/@relativeFrom` — ST_RelFromH. */
export const FLOAT_RELATIVE_FROM_H = [
  "character",
  "column",
  "insideMargin",
  "leftMargin",
  "margin",
  "outsideMargin",
  "page",
  "rightMargin",
] as const;
export type FloatRelativeFromH = (typeof FLOAT_RELATIVE_FROM_H)[number];

/** `wp:positionV/@relativeFrom` — ST_RelFromV. */
export const FLOAT_RELATIVE_FROM_V = [
  "bottomMargin",
  "insideMargin",
  "line",
  "margin",
  "outsideMargin",
  "page",
  "paragraph",
  "topMargin",
] as const;
export type FloatRelativeFromV = (typeof FLOAT_RELATIVE_FROM_V)[number];

export const FLOAT_ALIGN_H = ["left", "right", "center", "inside", "outside"] as const;
export type FloatAlignH = (typeof FLOAT_ALIGN_H)[number];

export const FLOAT_ALIGN_V = ["top", "bottom", "center", "inside", "outside"] as const;
export type FloatAlignV = (typeof FLOAT_ALIGN_V)[number];

/** The five wrap elements `wp:anchor` chooses between. */
export const FLOAT_WRAP_KINDS = [
  "none",
  "square",
  "tight",
  "through",
  "topAndBottom",
] as const;
export type FloatWrapKind = (typeof FLOAT_WRAP_KINDS)[number];

/** `@wrapText` on wrapSquare/wrapTight/wrapThrough. */
export const FLOAT_WRAP_TEXT = ["bothSides", "left", "right", "largest"] as const;
export type FloatWrapText = (typeof FLOAT_WRAP_TEXT)[number];

export interface FloatAnchorAxis<Origin extends string, Align extends string> {
  relativeFrom: Origin;
  /** Signed EMU from the origin. Mutually exclusive with `align`. */
  offsetEmu?: number;
  align?: Align;
  /**
   * Word 2010's relative position (`wp14:pctPosHOffset` / `pctPosVOffset`):
   * thousandths of a percent of the origin's extent, signed. Resolved against
   * the band's extent when a geometry is known; `offsetEmu` beside it is the
   * file's absolute place for the page it was saved on, used without one.
   */
  percentOffset?: number;
}

export interface FloatWrap {
  kind: FloatWrapKind;
  text?: FloatWrapText;
  distTEmu?: number;
  distBEmu?: number;
  distLEmu?: number;
  distREmu?: number;
  /**
   * `wp:wrapPolygon`: `wp:start` followed by its `wp:lineTo` points, verbatim.
   * Word writes them in the shape's own relative space (0..21600 on each
   * axis), which is what `floatWrapShapeOutside` assumes; a file using another
   * space simply gets no `shape-outside` rather than a wrong one.
   */
  polygon?: Array<[number, number]>;
  polygonEdited?: boolean;
}

/**
 * Where a frame sits inside the `wpg:wgp` group its anchor places.
 *
 * A group draws its members in a child coordinate space of its own
 * (`wpg:grpSpPr/a:xfrm`'s `a:chOff`/`a:chExt`), and the reader has already
 * mapped this member out of it: `offset*Emu` is the member's top-left in EMU
 * from the GROUP's top-left, and `frame*Emu` is the group's own extent.
 *
 * The group's extent matters because it is the GROUP that Word aligns within
 * a band — `wp:align="center"` centres the group, and the member then sits at
 * its offset inside it. Resolving the alignment against the member's own
 * (smaller) width would put every member in a different, wrong place.
 */
export interface FloatGroupTransform {
  offsetXEmu: number;
  offsetYEmu: number;
  frameWidthEmu: number;
  frameHeightEmu: number;
}

export interface FloatAnchor {
  version: 1;
  horizontal: FloatAnchorAxis<FloatRelativeFromH, FloatAlignH>;
  vertical: FloatAnchorAxis<FloatRelativeFromV, FloatAlignV>;
  /** Present only when the anchor draws a group and this frame is a member. */
  groupTransform?: FloatGroupTransform;
  widthEmu: number;
  heightEmu: number;
  wrap: FloatWrap;
  /** Paints under the text layer. Only `wrapNone` can overlap text at all. */
  behindDoc: boolean;
  locked: boolean;
  layoutInCell: boolean;
  allowOverlap: boolean;
  hidden: boolean;
  /** `@relativeHeight`: the floating z-order. Larger paints later. */
  relativeHeight: number;
  simplePos: boolean;
  simplePosXEmu?: number;
  simplePosYEmu?: number;
  distTEmu: number;
  distBEmu: number;
  distLEmu: number;
  distREmu: number;
  effectLEmu: number;
  effectTEmu: number;
  effectREmu: number;
  effectBEmu: number;
  /** 60000ths of a degree, as `a:xfrm/@rot` states it. */
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  /**
   * How many characters of the host paragraph's text stand before the run
   * holding the anchor, for an anchor placed across from a `character` or
   * down from a `line` (the reader states it only for those). Word measures
   * the character origin from the left edge of the character just before
   * the anchor (the line's start when none is), and the line origin from the
   * top of the line holding it (Word 16.86, wps-position-parity Gchar).
   */
  anchorTextOffset?: number;
}

/** `wps:bodyPr` — how the words sit inside the frame. */
export interface FloatFrameBody {
  anchor: "t" | "ctr" | "b" | "just" | "dist";
  anchorCtr: boolean;
  vert: "horz" | "vert" | "vert270" | "wordArtVert" | "eaVert" | "mongolianVert" | "wordArtVertRtl";
  wrap: "square" | "none";
  insetLEmu: number;
  insetTEmu: number;
  insetREmu: number;
  insetBEmu: number;
  autofit: "none" | "shape" | "normal";
  rotation: number;
  upright: boolean;
  /**
   * The words' WordArt transform (`a:prstTxWarp`): Word's Insert > WordArt is
   * a text box whose words are bent. The ONE warp value of
   * `@/lib/office-text-warp`; absent bends nothing.
   */
  warp?: OfficeTextWarp;
}

/** `wps:txbx` / `wps:linkedTxbx` — the linked text-box chain. */
export interface FloatFrameChain {
  storyId: number;
  sequence: number;
  /** True for the box that holds the chain's single `w:txbxContent`. */
  ownsStory: boolean;
}

const DEFAULT_BODY: FloatFrameBody = {
  anchor: "t",
  anchorCtr: false,
  vert: "horz",
  wrap: "square",
  // DrawingML's own defaults: 0.1 in left/right, 0.05 in top/bottom.
  insetLEmu: 91440,
  insetTEmu: 45720,
  insetREmu: 91440,
  insetBEmu: 45720,
  autofit: "none",
  rotation: 0,
  upright: false,
};

function finiteInt(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? Math.round(value) : fallback;
}

function boolAttr(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function oneOf<T extends string>(value: unknown, allowed: readonly T[], fallback: T): T {
  return typeof value === "string" && (allowed as readonly string[]).includes(value)
    ? (value as T)
    : fallback;
}

/**
 * Read a persisted `floatAnchor` attribute, or `null` when it is not one.
 *
 * Fail-closed on the two fields that decide WHERE the frame goes: an axis with
 * neither an offset nor a valid alignment, and an extent that is no box (a
 * negative edge, or both edges zero), both yield `null` so the frame falls back to the flow instead of being drawn at
 * an invented position.
 */
export function normalizeFloatAnchor(value: unknown): FloatAnchor | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  const horizontal = normalizeAxis(raw.horizontal, FLOAT_RELATIVE_FROM_H, FLOAT_ALIGN_H, "column", "left");
  const vertical = normalizeAxis(raw.vertical, FLOAT_RELATIVE_FROM_V, FLOAT_ALIGN_V, "paragraph", "top");
  if (!horizontal || !vertical) return null;
  const widthEmu = finiteInt(raw.widthEmu, -1);
  const heightEmu = finiteInt(raw.heightEmu, -1);
  // ONE edge may be zero: a floating horizontal line is as tall as nothing
  // (Word states `cy="0"`; VML, two endpoints with the same y), and its anchor
  // still says exactly where it goes. Refusing it sent a floating line back
  // into the flow. Neither edge may be negative and both at zero is no box.
  if (widthEmu < 0 || heightEmu < 0 || (widthEmu === 0 && heightEmu === 0))
    return null;
  const wrap = normalizeFloatWrap(raw.wrap);
  if (!wrap) return null;
  const groupTransform = normalizeFloatGroupTransform(raw.groupTransform);
  return {
    version: 1,
    horizontal,
    vertical,
    ...(groupTransform ? { groupTransform } : {}),
    widthEmu,
    heightEmu,
    wrap,
    behindDoc: boolAttr(raw.behindDoc, false),
    locked: boolAttr(raw.locked, false),
    layoutInCell: boolAttr(raw.layoutInCell, true),
    allowOverlap: boolAttr(raw.allowOverlap, true),
    hidden: boolAttr(raw.hidden, false),
    relativeHeight: Math.max(0, finiteInt(raw.relativeHeight, 0)),
    simplePos: boolAttr(raw.simplePos, false),
    ...(raw.simplePosXEmu === undefined ? {} : { simplePosXEmu: finiteInt(raw.simplePosXEmu, 0) }),
    ...(raw.simplePosYEmu === undefined ? {} : { simplePosYEmu: finiteInt(raw.simplePosYEmu, 0) }),
    distTEmu: Math.max(0, finiteInt(raw.distTEmu, 0)),
    distBEmu: Math.max(0, finiteInt(raw.distBEmu, 0)),
    distLEmu: Math.max(0, finiteInt(raw.distLEmu, 0)),
    distREmu: Math.max(0, finiteInt(raw.distREmu, 0)),
    effectLEmu: finiteInt(raw.effectLEmu, 0),
    effectTEmu: finiteInt(raw.effectTEmu, 0),
    effectREmu: finiteInt(raw.effectREmu, 0),
    effectBEmu: finiteInt(raw.effectBEmu, 0),
    rotation: finiteInt(raw.rotation, 0),
    flipH: boolAttr(raw.flipH, false),
    flipV: boolAttr(raw.flipV, false),
    ...(typeof raw.anchorTextOffset === "number" && Number.isFinite(raw.anchorTextOffset) && raw.anchorTextOffset >= 0
      ? { anchorTextOffset: Math.round(raw.anchorTextOffset) }
      : {}),
  };
}

function normalizeAxis<Origin extends string, Align extends string>(
  value: unknown,
  origins: readonly Origin[],
  aligns: readonly Align[],
  fallbackOrigin: Origin,
  _fallbackAlign: Align,
): FloatAnchorAxis<Origin, Align> | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  const relativeFrom = oneOf(raw.relativeFrom, origins, fallbackOrigin);
  if (typeof raw.percentOffset === "number" && Number.isFinite(raw.percentOffset)) {
    const percentOffset = Math.max(-1_000_000, Math.min(1_000_000, Math.round(raw.percentOffset)));
    return typeof raw.offsetEmu === "number" && Number.isFinite(raw.offsetEmu)
      ? { relativeFrom, percentOffset, offsetEmu: Math.round(raw.offsetEmu) }
      : { relativeFrom, percentOffset };
  }
  if (typeof raw.offsetEmu === "number" && Number.isFinite(raw.offsetEmu)) {
    return { relativeFrom, offsetEmu: Math.round(raw.offsetEmu) };
  }
  if (typeof raw.align === "string" && (aligns as readonly string[]).includes(raw.align)) {
    return { relativeFrom, align: raw.align as Align };
  }
  return null;
}

/**
 * Read a persisted `groupTransform`, or `null` when it is not one.
 *
 * Fail-closed like the anchor itself: a group extent that is not positive
 * cannot resolve an alignment, so the frame falls back to being placed as if
 * it were the anchor's own box rather than being drawn at a guess.
 */
export function normalizeFloatGroupTransform(value: unknown): FloatGroupTransform | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  const frameWidthEmu = finiteInt(raw.frameWidthEmu, 0);
  const frameHeightEmu = finiteInt(raw.frameHeightEmu, 0);
  if (frameWidthEmu <= 0 || frameHeightEmu <= 0) return null;
  if (typeof raw.offsetXEmu !== "number" || !Number.isFinite(raw.offsetXEmu)) return null;
  if (typeof raw.offsetYEmu !== "number" || !Number.isFinite(raw.offsetYEmu)) return null;
  return {
    offsetXEmu: Math.round(raw.offsetXEmu),
    offsetYEmu: Math.round(raw.offsetYEmu),
    frameWidthEmu,
    frameHeightEmu,
  };
}

export function normalizeFloatWrap(value: unknown): FloatWrap | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  const kind = oneOf(raw.kind, FLOAT_WRAP_KINDS, "none");
  const wrap: FloatWrap = { kind };
  if (kind === "square" || kind === "tight" || kind === "through") {
    wrap.text = oneOf(raw.text, FLOAT_WRAP_TEXT, "bothSides");
  }
  for (const key of ["distTEmu", "distBEmu", "distLEmu", "distREmu"] as const) {
    if (typeof raw[key] === "number" && Number.isFinite(raw[key] as number)) {
      wrap[key] = Math.max(0, Math.round(raw[key] as number));
    }
  }
  if (Array.isArray(raw.polygon) && (kind === "tight" || kind === "through")) {
    const points: Array<[number, number]> = [];
    for (const point of raw.polygon) {
      if (!Array.isArray(point) || point.length !== 2) continue;
      const [x, y] = point;
      if (typeof x !== "number" || typeof y !== "number") continue;
      if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
      points.push([Math.round(x), Math.round(y)]);
    }
    if (points.length >= 3) {
      wrap.polygon = points.slice(0, 256);
      if (typeof raw.polygonEdited === "boolean") wrap.polygonEdited = raw.polygonEdited;
    }
  }
  return wrap;
}

export function normalizeFloatFrameBody(value: unknown): FloatFrameBody {
  if (!value || typeof value !== "object") return { ...DEFAULT_BODY };
  const raw = value as Record<string, unknown>;
  return {
    anchor: oneOf(raw.anchor, ["t", "ctr", "b", "just", "dist"] as const, "t"),
    anchorCtr: boolAttr(raw.anchorCtr, false),
    vert: oneOf(
      raw.vert,
      ["horz", "vert", "vert270", "wordArtVert", "eaVert", "mongolianVert", "wordArtVertRtl"] as const,
      "horz",
    ),
    wrap: oneOf(raw.wrap, ["square", "none"] as const, "square"),
    insetLEmu: Math.max(0, finiteInt(raw.insetLEmu, DEFAULT_BODY.insetLEmu)),
    insetTEmu: Math.max(0, finiteInt(raw.insetTEmu, DEFAULT_BODY.insetTEmu)),
    insetREmu: Math.max(0, finiteInt(raw.insetREmu, DEFAULT_BODY.insetREmu)),
    insetBEmu: Math.max(0, finiteInt(raw.insetBEmu, DEFAULT_BODY.insetBEmu)),
    autofit: oneOf(raw.autofit, ["none", "shape", "normal"] as const, "none"),
    rotation: finiteInt(raw.rotation, 0),
    upright: boolAttr(raw.upright, false),
    ...(normalizeOfficeTextWarp(raw.warp) ? { warp: normalizeOfficeTextWarp(raw.warp)! } : {}),
  };
}

/**
 * The INLINE half of the model this module's header names: `wp:inline`'s own
 * `wp:extent`.
 *
 * An in-line text box has no origin, no wrap and no z-order — it takes a place
 * in the flow. What it DOES have is a size, and that size had no home on the
 * page at all: the exporter wrote a hard-coded `2700000 x 900000` extent for
 * every text box that carried no anchor, so an in-line box came back a
 * different size than it went in
 * (docs/office-graphics/SHAPE-IMPORT-AUDIT.md §37.1).
 */
export interface FloatFrameExtent {
  readonly widthEmu: number;
  readonly heightEmu: number;
}

/** ECMA-376 `ST_PositiveCoordinate`, capped at what this product lays out. */
const MAX_FRAME_EXTENT_EMU = 50_292_000;

export function normalizeDocumentTextBoxFrameExtent(
  value: unknown,
): FloatFrameExtent | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const raw = value as Record<string, unknown>;
  const width = finiteInt(raw.widthEmu, -1);
  const height = finiteInt(raw.heightEmu, -1);
  if (width <= 0 || height <= 0) return null;
  if (width > MAX_FRAME_EXTENT_EMU || height > MAX_FRAME_EXTENT_EMU) return null;
  return { widthEmu: width, heightEmu: height };
}

export function parseDocumentTextBoxFrameExtentAttribute(
  value: string | null,
): FloatFrameExtent | null {
  if (!value) return null;
  try {
    return normalizeDocumentTextBoxFrameExtent(JSON.parse(value));
  } catch {
    return null;
  }
}

export function documentTextBoxFrameExtentHtmlAttributes(
  value: unknown,
): Record<string, string> {
  const extent = normalizeDocumentTextBoxFrameExtent(value);
  if (!extent) return {};
  // WIDTH is definite; HEIGHT is a floor, not a ceiling. Word's own in-line
  // box keeps its authored height and lets an overset story spill past it,
  // and a hard `height` here would CLIP words instead — which is the one
  // thing a frame must never do to the text it holds.
  return {
    "data-frame-extent": JSON.stringify(extent),
    style:
      `width:${Number(emuToPx(extent.widthEmu).toFixed(3))}px;` +
      `min-height:${Number(emuToPx(extent.heightEmu).toFixed(3))}px`,
  };
}

export function normalizeFloatFrameChain(value: unknown): FloatFrameChain | null {
  if (!value || typeof value !== "object") return null;
  const raw = value as Record<string, unknown>;
  if (typeof raw.storyId !== "number" || !Number.isFinite(raw.storyId)) return null;
  return {
    storyId: Math.round(raw.storyId),
    sequence: Math.max(0, finiteInt(raw.sequence, 0)),
    ownsStory: boolAttr(raw.ownsStory, false),
  };
}

/**
 * The page geometry an anchor is resolved against, in CSS px.
 *
 * `page` is the paper; `content` is the text area inside the margins;
 * `column` is the current newspaper column (equal to `content` in a
 * single-column section). `flow` is where the anchor's own paragraph/line
 * currently sits, which is what the `paragraph`, `line` and `character`
 * origins are measured from.
 */
export interface FloatPageGeometry {
  pageWidthPx: number;
  pageHeightPx: number;
  marginLeftPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  columnLeftPx?: number;
  columnWidthPx?: number;
  /** Anchor paragraph's top, relative to the page's top-left. */
  flowTopPx?: number;
  /** The `character` origin: the left edge of the character just before
   * the anchor in its line, relative to the page's top-left. */
  flowLeftPx?: number;
  /** The `line` origin: the top of the line holding the anchor, relative to
   * the page's top-left (the paragraph's top, `flowTopPx`, when absent). */
  lineTopPx?: number;
  /** True on a left-hand (even) page of a mirrored-margin section. */
  mirroredEvenPage?: boolean;
  /** Where Word keeps an object on this page (`keepFloatOnPage`); absent
   * where it does not (a table cell). */
  keepOnPage?: FloatKeepOnPage;
}

/**
 * The two facts Word's keep-on-page rule depends on: the document's
 * `compatibilityMode` (null when it states none, which Word lays out by the
 * Word 2010 rules) and whether the anchor is in a header or footer.
 */
export interface FloatKeepOnPage {
  readonly wordCompatibilityMode: number | null;
  readonly headerFooter: boolean;
}

/** Word 2013 and later layout (compatibility mode 15 and up). */
export function floatLayoutIsModern(keep: FloatKeepOnPage): boolean {
  return keep.wordCompatibilityMode !== null && keep.wordCompatibilityMode >= 15;
}

export interface FloatPlacement {
  /** From the page's top-left corner, in CSS px. */
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
}

function horizontalBand(
  origin: FloatRelativeFromH,
  geometry: FloatPageGeometry,
): { startPx: number; extentPx: number } {
  const contentLeft = geometry.marginLeftPx;
  const contentWidth = Math.max(
    0,
    geometry.pageWidthPx - geometry.marginLeftPx - geometry.marginRightPx,
  );
  const mirrored = geometry.mirroredEvenPage === true;
  switch (origin) {
    case "page":
      return { startPx: 0, extentPx: geometry.pageWidthPx };
    case "margin":
      return { startPx: contentLeft, extentPx: contentWidth };
    case "leftMargin":
      return { startPx: 0, extentPx: geometry.marginLeftPx };
    case "rightMargin":
      return {
        startPx: geometry.pageWidthPx - geometry.marginRightPx,
        extentPx: geometry.marginRightPx,
      };
    case "insideMargin":
      // The inside margin is the binding edge: left on an odd (right-hand)
      // page, right on an even one.
      return mirrored
        ? { startPx: geometry.pageWidthPx - geometry.marginRightPx, extentPx: geometry.marginRightPx }
        : { startPx: 0, extentPx: geometry.marginLeftPx };
    case "outsideMargin":
      return mirrored
        ? { startPx: 0, extentPx: geometry.marginLeftPx }
        : {
            startPx: geometry.pageWidthPx - geometry.marginRightPx,
            extentPx: geometry.marginRightPx,
          };
    case "column":
      return {
        startPx: geometry.columnLeftPx ?? contentLeft,
        extentPx: geometry.columnWidthPx ?? contentWidth,
      };
    case "character":
      return { startPx: geometry.flowLeftPx ?? contentLeft, extentPx: contentWidth };
    default:
      return { startPx: contentLeft, extentPx: contentWidth };
  }
}

function verticalBand(
  origin: FloatRelativeFromV,
  geometry: FloatPageGeometry,
): { startPx: number; extentPx: number } {
  const contentTop = geometry.marginTopPx;
  const contentHeight = Math.max(
    0,
    geometry.pageHeightPx - geometry.marginTopPx - geometry.marginBottomPx,
  );
  switch (origin) {
    case "page":
      return { startPx: 0, extentPx: geometry.pageHeightPx };
    case "margin":
      return { startPx: contentTop, extentPx: contentHeight };
    case "topMargin":
      return { startPx: 0, extentPx: geometry.marginTopPx };
    case "bottomMargin":
      return {
        startPx: geometry.pageHeightPx - geometry.marginBottomPx,
        extentPx: geometry.marginBottomPx,
      };
    // Vertically there is no binding edge: Word resolves inside/outside to the
    // top and bottom margin bands.
    case "insideMargin":
      return { startPx: 0, extentPx: geometry.marginTopPx };
    case "outsideMargin":
      return {
        startPx: geometry.pageHeightPx - geometry.marginBottomPx,
        extentPx: geometry.marginBottomPx,
      };
    case "paragraph":
      return { startPx: geometry.flowTopPx ?? contentTop, extentPx: contentHeight };
    case "line":
      return { startPx: geometry.lineTopPx ?? geometry.flowTopPx ?? contentTop, extentPx: contentHeight };
    default:
      return { startPx: contentTop, extentPx: contentHeight };
  }
}

function alignedOffset(
  align: string,
  bandExtentPx: number,
  objectExtentPx: number,
  mirrored: boolean,
): number {
  switch (align) {
    case "left":
    case "top":
      return 0;
    case "right":
    case "bottom":
      return bandExtentPx - objectExtentPx;
    case "center":
      return (bandExtentPx - objectExtentPx) / 2;
    case "inside":
      return mirrored ? bandExtentPx - objectExtentPx : 0;
    case "outside":
      return mirrored ? 0 : bandExtentPx - objectExtentPx;
    default:
      return 0;
  }
}

/**
 * Where a frame lands on a page, in CSS px from the page's top-left corner.
 *
 * `wp:simplePos` wins when the anchor sets `@simplePos="1"`, exactly as the
 * spec says: the two axes are then ignored in favour of `wp:simplePos/@x,@y`.
 */
/** An offset axis's distance from its origin: a percent of the band, else EMU. */
function axisOffsetPx(
  axis: FloatAnchorAxis<string, string>,
  bandExtentPx: number,
): number {
  if (typeof axis.percentOffset === "number" && Number.isFinite(axis.percentOffset)) {
    return (bandExtentPx * axis.percentOffset) / 100_000;
  }
  return emuToPx(axis.offsetEmu ?? 0);
}

export function resolveFloatPlacement(
  anchor: FloatAnchor,
  geometry: FloatPageGeometry,
): FloatPlacement {
  const widthPx = emuToPx(anchor.widthEmu);
  const heightPx = emuToPx(anchor.heightEmu);
  if (anchor.simplePos) {
    return {
      leftPx: emuToPx(anchor.simplePosXEmu ?? 0),
      topPx: emuToPx(anchor.simplePosYEmu ?? 0),
      widthPx,
      heightPx,
    };
  }
  const mirrored = geometry.mirroredEvenPage === true;
  const horizontal = horizontalBand(anchor.horizontal.relativeFrom, geometry);
  const vertical = verticalBand(anchor.vertical.relativeFrom, geometry);
  // What the anchor actually places. For a grouped frame that is the GROUP:
  // Word aligns the group in the band and sizes the wrap by it, then draws
  // this member at its own offset inside it. For an ordinary frame the two
  // are the same box, so the group branch costs nothing.
  const group = anchor.groupTransform;
  const placedWidthPx = group ? emuToPx(group.frameWidthEmu) : widthPx;
  const placedHeightPx = group ? emuToPx(group.frameHeightEmu) : heightPx;
  const leftPx =
    horizontal.startPx +
    (anchor.horizontal.align
      ? alignedOffset(anchor.horizontal.align, horizontal.extentPx, placedWidthPx, mirrored)
      : axisOffsetPx(anchor.horizontal, horizontal.extentPx)) +
    (group ? emuToPx(group.offsetXEmu) : 0);
  const topPx =
    vertical.startPx +
    (anchor.vertical.align
      ? alignedOffset(anchor.vertical.align, vertical.extentPx, placedHeightPx, mirrored)
      : axisOffsetPx(anchor.vertical, vertical.extentPx)) +
    (group ? emuToPx(group.offsetYEmu) : 0);
  return { leftPx, topPx, widthPx, heightPx };
}

































export function keepFloatOnPage(
  placement: FloatPlacement,
  anchor: FloatAnchor,
  geometry: FloatPageGeometry,
): FloatPlacement {
  const keep = geometry.keepOnPage;
  if (!keep || anchor.simplePos || anchor.wrap.kind === "none") return placement;
  const modern = floatLayoutIsModern(keep);
  if (modern && keep.headerFooter) return placement;
  if (!floatMovesWithTextForKeep(anchor, modern)) return placement;
  // The anchor places the group's frame; a member moves with it.
  const group = anchor.groupTransform;
  const frameLeft = placement.leftPx - (group ? emuToPx(group.offsetXEmu) : 0);
  const frameTop = placement.topPx - (group ? emuToPx(group.offsetYEmu) : 0);
  const frameWidth = group ? emuToPx(group.frameWidthEmu) : placement.widthPx;
  const frameHeight = group ? emuToPx(group.frameHeightEmu) : placement.heightPx;
  const within = (start: number, size: number, low: number, high: number) =>
    Math.max(low, Math.min(start, high - size));
  const left = within(frameLeft, frameWidth, 0, geometry.pageWidthPx);
  let top = frameTop;
  if (!modern) top = within(frameTop, frameHeight, 0, geometry.pageHeightPx);
  else if (anchor.vertical.relativeFrom === "paragraph" || anchor.vertical.relativeFrom === "line") {
    top = within(frameTop, frameHeight, geometry.marginTopPx, geometry.pageHeightPx - geometry.marginBottomPx);
  }
  if (left === frameLeft && top === frameTop) return placement;
  return { ...placement, leftPx: placement.leftPx + left - frameLeft, topPx: placement.topPx + top - frameTop };
}

/** Whether Word keeps this object on its page at all (see `keepFloatOnPage`). */
function floatMovesWithTextForKeep(anchor: FloatAnchor, modern: boolean): boolean {
  const vertical = anchor.vertical.relativeFrom;
  if (modern) return vertical === "paragraph" || vertical === "line";
  return vertical === "paragraph" || anchor.horizontal.relativeFrom === "column";
}

/** Where an object is drawn on its page: its anchor's place, kept on the
 * page as Word keeps it. */
export function placeFloatOnPage(anchor: FloatAnchor, geometry: FloatPageGeometry): FloatPlacement {
  return keepFloatOnPage(resolveFloatPlacement(anchor, geometry), anchor, geometry);
}

/**
 * Where the float pass publishes `floatPageRoomPx` for the paginator, on the
 * object's wrapper (one that still carries `data-float-anchor`).
 */
export const FLOAT_PAGE_ROOM_ATTR = "data-float-page-room";

/**
 * A run of Top and Bottom objects laid out whole (floatingTextBoxLayout's
 * `layoutMixedBandRuns`) is charged to the paginator as one block: its first
 * object asks, besides its own height, the room from its own foot down to the
 * lowest object of the run (px, `BAND_RUN_ROOM_ATTR`), and every block after
 * it in the run rides on it (`BAND_RUN_RIDE_ATTR`) at no height. Charged one
 * by one in their stacked order, the run's later objects were sent to the
 * next page before the run could set them beside the first (R308-D01).
 */
export const BAND_RUN_ROOM_ATTR = "data-band-run-room";
export const BAND_RUN_RIDE_ATTR = "data-band-run-ride";

/**
 * The room below its paragraph's top a floating object needs on that page,
 * in px, past which Word moves the paragraph to the next page rather than
 * leave the object past the bottom margin: a body object the text wraps around, placed from
 * its paragraph or line -- in compatibility mode 15 any such object
 * (measured: a 2" picture anchored 0" or 0.5" below a paragraph 1.5" above
 * the bottom margin opens the next page with its paragraph), in mode 14 and
 * earlier only one that leaves too little room beside it for the text it
 * would cover (`FLOAT_TEXT_BESIDE_MIN_PX`; a narrower one stays and is kept
 * on the paper). Null when Word moves nothing for it. An object taller than
 * a fresh page's column moves too, and is kept on that page
 * (`keepFloatOnPage`).
 */
export function floatPageRoomPx(anchor: FloatAnchor, geometry: FloatPageGeometry): number | null {
  const keep = geometry.keepOnPage;
  if (!keep || keep.headerFooter) return null;
  if (anchor.simplePos || anchor.wrap.kind === "none") return null;
  if (anchor.vertical.relativeFrom !== "paragraph" && anchor.vertical.relativeFrom !== "line") return null;
  if (geometry.flowTopPx === undefined) return null;
  const placement = resolveFloatPlacement(anchor, geometry);
  const group = anchor.groupTransform;
  const frameLeft = placement.leftPx - (group ? emuToPx(group.offsetXEmu) : 0);
  const frameTop = placement.topPx - (group ? emuToPx(group.offsetYEmu) : 0);
  const frameWidth = group ? emuToPx(group.frameWidthEmu) : placement.widthPx;
  const frameHeight = group ? emuToPx(group.frameHeightEmu) : placement.heightPx;
  const modern = floatLayoutIsModern(keep);
  if (!modern && floatTextRoomBesidePx(anchor, frameLeft, frameWidth, geometry) >= FLOAT_TEXT_BESIDE_MIN_PX) {
    return null;
  }
  // Word keeps the object on the page by raising it -- inside the margins
  // (mode 15) or the paper (14 and earlier) -- but never above its own
  // paragraph's top: the paragraph moves on only when the object does not
  // fit below that top. Windows Word, docx-look/probes room_probe.py (a box
  // anchored 7.9" down an A4 page, 1" margins): 2" tall at 2" or 3" below the
  // paragraph stays on the page, raised to end on the bottom margin; 3" tall
  // 1" below it, or 4" tall at 0", moves on in mode 15; the same 3" box stays
  // in modes 12 and 14, raised to end on the paper's edge. Counted from the
  // drawn offset (offset plus height), every such box pushed its paragraph
  // on (corpus R372-D01: a cover's credits box 8.7" below an empty paragraph
  // took the cover to page 2).
  // `window.__docsFloatRoomFromOffset = true` asks the older room (offset plus
  // height) for an A/B of one build.
  if (typeof window !== "undefined" && (window as Window & { __docsFloatRoomFromOffset?: boolean }).__docsFloatRoomFromOffset === true) {
    return Math.max(0, frameTop - geometry.flowTopPx + frameHeight);
  }
  const below = Math.min(frameTop - geometry.flowTopPx, 0) + frameHeight;
  return Math.max(0, modern ? below : below - geometry.marginBottomPx);
}















const FLOAT_TEXT_BESIDE_MIN_PX = 1.15 * 96;









export function floatTextStripMinPx(keep: FloatKeepOnPage | undefined): number {
  return (keep && floatLayoutIsModern(keep) ? 0.325 : 0.225) * 96;
}

/** The widest room the text may use beside a wrapped frame in its column, px
 * (none for Top and Bottom, which has no text beside it). */
function floatTextRoomBesidePx(
  anchor: FloatAnchor,
  frameLeftPx: number,
  frameWidthPx: number,
  geometry: FloatPageGeometry,
): number {
  if (anchor.wrap.kind === "topAndBottom") return 0;
  const columnLeft = geometry.columnLeftPx ?? geometry.marginLeftPx;
  const columnRight =
    columnLeft + (geometry.columnWidthPx ?? Math.max(0, geometry.pageWidthPx - geometry.marginLeftPx - geometry.marginRightPx));
  const left = Math.max(0, frameLeftPx - emuToPx(anchor.wrap.distLEmu ?? anchor.distLEmu) - columnLeft);
  const right = Math.max(0, columnRight - (frameLeftPx + frameWidthPx + emuToPx(anchor.wrap.distREmu ?? anchor.distREmu)));
  const text = anchor.wrap.text ?? "bothSides";
  if (text === "left") return left;
  if (text === "right") return right;
  return Math.max(left, right);
}

/**
 * How a frame takes part in the body flow.
 *
 * - `overlay`  — `wrapNone`: the frame is out of flow entirely and text runs
 *   under or over it. This is Word's "In Front Of Text" / "Behind Text".
 * - `float`    — `wrapSquare`/`wrapTight`/`wrapThrough`: the frame sits beside
 *   the text, which re-wraps around it.
 * - `band`     — `wrapTopAndBottom`: no text beside the frame at all; the flow
 *   resumes below it.
 */
export type FloatFlowMode = "overlay" | "float" | "band";

/**
 * Whether a wrapped frame lies wholly in the page's top or bottom margin,
 * its wrap distances included. No body line runs beside or under it, so it
 * takes no room from the body: Word draws R21-D04's page-top banner (a VML
 * rect at the top of the page, 42 pt tall, square wrap) in the margin and
 * starts the body at the margin, where a float in the flow pushed the first
 * page's text down 62 px.
 *
 * So does one wholly in the left or right margin, beside every line, when
 * the placement and geometry say where it stands across: Word draws it at its
 * page point and wraps nothing (corpus R400-D05: a 499 pt grey sidebar box,
 * square-wrapped, 117 pt below the top margin, its paragraph a long abstract;
 * as a float it could not rise above its place after the abstract and was
 * drawn 192 pt low). `window.__docsFloatSideMarginOff = true` asks the older
 * reading (top and bottom margins only) for an A/B of one build.
 */
export function floatWhollyInPageMargin(
  placement: Pick<FloatPlacement, "topPx" | "heightPx"> & Partial<Pick<FloatPlacement, "leftPx" | "widthPx">>,
  distances: { topPx: number; bottomPx: number; leftPx?: number; rightPx?: number },
  geometry: Pick<FloatPageGeometry, "pageHeightPx" | "marginTopPx" | "marginBottomPx"> &
    Partial<Pick<FloatPageGeometry, "pageWidthPx" | "marginLeftPx" | "marginRightPx">>,
): boolean {
  const top = placement.topPx - distances.topPx;
  const bottom = placement.topPx + placement.heightPx + distances.bottomPx;
  if (bottom <= geometry.marginTopPx + 0.5 || top >= geometry.pageHeightPx - geometry.marginBottomPx - 0.5) return true;
  if (
    typeof window !== "undefined" &&
    (window as Window & { __docsFloatSideMarginOff?: boolean }).__docsFloatSideMarginOff === true
  ) {
    return false;
  }
  const { leftPx, widthPx } = placement;
  const { pageWidthPx, marginLeftPx, marginRightPx } = geometry;
  if (
    leftPx === undefined || widthPx === undefined || pageWidthPx === undefined ||
    marginLeftPx === undefined || marginRightPx === undefined || !(marginLeftPx > 0 || marginRightPx > 0)
  ) {
    return false;
  }
  const left = leftPx - (distances.leftPx ?? 0);
  const right = leftPx + widthPx + (distances.rightPx ?? 0);
  return (marginLeftPx > 0 && right <= marginLeftPx + 0.5) || (marginRightPx > 0 && left >= pageWidthPx - marginRightPx - 0.5);
}

export function floatFlowMode(anchor: FloatAnchor): FloatFlowMode {
  if (anchor.wrap.kind === "none") return "overlay";
  if (anchor.wrap.kind === "topAndBottom") return "band";
  return "float";
}

/**
 * Which side a side-wrapped frame floats to.
 *
 * `@wrapText` states which side the TEXT may use, so `left` means text on the
 * left and the frame on the right. `bothSides`/`largest` leave the choice to
 * the position: Word puts the frame on whichever side its own centre is
 * nearest, and text takes the wider remainder.
 */
export function floatSide(
  anchor: FloatAnchor,
  placement: FloatPlacement,
  geometry: FloatPageGeometry,
): "left" | "right" {
  const text = anchor.wrap.text ?? "bothSides";
  if (text === "left") return "right";
  if (text === "right") return "left";
  const contentLeft = geometry.marginLeftPx;
  const contentWidth = Math.max(
    0,
    geometry.pageWidthPx - geometry.marginLeftPx - geometry.marginRightPx,
  );
  const centre = placement.leftPx + placement.widthPx / 2;
  return centre <= contentLeft + contentWidth / 2 ? "left" : "right";
}

/**
 * A `shape-outside` polygon for a tight/through wrap, or `null`.
 *
 * Word's polygon points are in the shape's own 0..21600 relative space, so a
 * point maps to a percentage of the frame's box. A polygon with a point
 * outside that space is not converted: a wrong outline is worse than a plain
 * rectangle, which is what `square` already gives.
 */
export function floatWrapShapeOutside(wrap: FloatWrap): string | null {
  const points = wrap.polygon;
  if (!points || points.length < 3) return null;
  if (points.some(([x, y]) => x < 0 || y < 0 || x > 21600 || y > 21600)) return null;
  const pairs = points.map(
    ([x, y]) => `${((x / 21600) * 100).toFixed(3)}% ${((y / 21600) * 100).toFixed(3)}%`,
  );
  return `polygon(${pairs.join(", ")})`;
}

/**
 * The frame's four wrap distances in px, as the anchor states them.
 *
 * The wrap element's own `dist*` attributes override the anchor's when it
 * carries them. `wrapNone` has no wrap distance at all — nothing is being
 * kept away from the frame.
 */
export function floatWrapDistancesPx(anchor: FloatAnchor): {
  topPx: number;
  bottomPx: number;
  leftPx: number;
  rightPx: number;
} {
  if (anchor.wrap.kind === "none") {
    return { topPx: 0, bottomPx: 0, leftPx: 0, rightPx: 0 };
  }
  return {
    topPx: emuToPx(anchor.wrap.distTEmu ?? anchor.distTEmu),
    bottomPx: emuToPx(anchor.wrap.distBEmu ?? anchor.distBEmu),
    leftPx: emuToPx(anchor.wrap.distLEmu ?? anchor.distLEmu),
    rightPx: emuToPx(anchor.wrap.distREmu ?? anchor.distREmu),
  };
}

/**
 * A compact, stable paint order for the floating stack.
 *
 * `@relativeHeight` is a UInt32 and cannot go into `z-index` as it stands.
 * Word's order is: everything with `behindDoc="1"` under the text layer, then
 * the text, then everything else, each half ordered by `relativeHeight`. The
 * text layer is 0 here, so a behind-text frame gets a negative rank.
 *
 * `allRelativeHeights` is the whole DOCUMENT's floating stack, never one
 * family's: Word has a single stack for shapes, text boxes, pictures and
 * groups alike (`lib/document-shapes.documentFloatingStack`). The parameter is
 * the two fields that decide a place in it, so a floating SHAPE — whose height
 * lives on the shape itself rather than inside its anchor — ranks through this
 * one function too.
 */
export function floatStackZIndex(
  anchor: Pick<FloatAnchor, "relativeHeight" | "behindDoc">,
  allRelativeHeights: readonly number[],
): number {
  const ranks = floatStackRanks(allRelativeHeights);
  const rank = (ranks.get(anchor.relativeHeight) ?? 0) + 1;
  if (!anchor.behindDoc) return rank;
  // Under the text layer the stack keeps its order: the higher of two
  // behind-text frames still paints over the lower one. Negated, the order
  // ran backwards (corpus R308-D01: a journal masthead of nine behind-text
  // VML shapes painted its backmost pale rectangle over the orange banner
  // in front of it). `window.__docsBehindStackOrderOff = true` keeps the
  // older order (A/B).
  if (behindStackOrderOff()) return -rank;
  return rank - Math.max(ranks.size, rank) - 1;
}

function behindStackOrderOff(): boolean {
  return typeof window !== "undefined" && (window as Window & { __docsBehindStackOrderOff?: boolean }).__docsBehindStackOrderOff === true;
}

/** Whether a height of `before` still in `after` stands at another rank there. */
export function floatStackRanksMoved(before: readonly number[], after: readonly number[]): boolean {
  const ranksAfter = floatStackRanks(after);
  for (const [height, rank] of floatStackRanks(before)) {
    const now = ranksAfter.get(height);
    if (now !== undefined && now !== rank) return true;
  }
  return false;
}

// The ranks of a frozen stack (a document's, from documentFloatingStack, which
// never changes): every floating object on the page ranks against the same
// list, so sorting it once per list rather than once per object keeps a page
// of ink strokes from sorting its whole stack for each stroke painted.
const frozenStackRanks = new WeakMap<readonly number[], ReadonlyMap<number, number>>();

function floatStackRanks(heights: readonly number[]): ReadonlyMap<number, number> {
  const cached = frozenStackRanks.get(heights);
  if (cached) return cached;
  const ranks = new Map<number, number>();
  for (const value of [...heights].sort((a, b) => a - b)) {
    if (!ranks.has(value)) ranks.set(value, ranks.size);
  }
  if (Object.isFrozen(heights)) frozenStackRanks.set(heights, ranks);
  return ranks;
}
