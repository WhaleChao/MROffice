/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  EMU_PER_PX,
  normalizeFloatAnchor,
  normalizeFloatFrameBody,
  type FloatAnchor,
  type FloatFrameBody,
} from "./office-float-anchor";
import type { DocumentBorderLine } from "./document-borders";
import { normalizeOfficeShapeShadow, type OfficeShapeShadow } from "./office-shape-paint";

export type DocumentTextBoxDirection = "horizontal" | "vertical";

/** Word's default line on a drawn text box: 0.5 pt black (`a:ln w="6350"`). */
export const DOCUMENT_TEXT_BOX_DEFAULT_OUTLINE: DocumentBorderLine = {
  style: "single",
  sizePt: 0.5,
  color: "#000000",
};

/** `wp:anchor/@distL|distR` Word writes on a new text box (0.125 in). */
const DEFAULT_SIDE_DISTANCE_EMU = 114_300;

/** The smallest frame a drag may leave (the shape minimum, ~0.1 in). */
export const DOCUMENT_TEXT_BOX_MIN_EDGE_EMU = 91_440;
/** What this product lays out (`office-float-anchor` MAX_FRAME_EXTENT_EMU). */
export const DOCUMENT_TEXT_BOX_MAX_EDGE_EMU = 50_292_000;

/**
 * The frame body a new box gets. A vertical box stacks its words top to
 * bottom with East Asian characters upright (`vert="eaVert"`), which is how
 * Word's Vertical Text Box lays out CJK text.
 */
export function documentTextBoxFrameBody(direction: DocumentTextBoxDirection): FloatFrameBody {
  return {
    ...normalizeFloatFrameBody(null),
    vert: direction === "vertical" ? "eaVert" : "horz",
  };
}

function clampEdge(value: number): number {
  return Math.min(DOCUMENT_TEXT_BOX_MAX_EDGE_EMU, Math.max(DOCUMENT_TEXT_BOX_MIN_EDGE_EMU, Math.round(value)));
}

/**
 * The `floatAnchor` of a new box `offset` EMU from its column's left edge and
 * its paragraph's top, `size` EMU large, painting at `relativeHeight`.
 */
export function documentTextBoxFloatAnchor(
  offset: { readonly xEmu: number; readonly yEmu: number },
  size: { readonly widthEmu: number; readonly heightEmu: number },
  relativeHeight: number,
): FloatAnchor {
  return {
    version: 1,
    horizontal: { relativeFrom: "column", offsetEmu: Math.round(offset.xEmu) },
    vertical: { relativeFrom: "paragraph", offsetEmu: Math.round(offset.yEmu) },
    widthEmu: clampEdge(size.widthEmu),
    heightEmu: clampEdge(size.heightEmu),
    wrap: { kind: "none" },
    behindDoc: false,
    locked: false,
    layoutInCell: true,
    allowOverlap: true,
    hidden: false,
    relativeHeight,
    simplePos: false,
    distTEmu: 0,
    distBEmu: 0,
    distLEmu: DEFAULT_SIDE_DISTANCE_EMU,
    distREmu: DEFAULT_SIDE_DISTANCE_EMU,
    effectLEmu: 0,
    effectTEmu: 0,
    effectREmu: 0,
    effectBEmu: 0,
    rotation: 0,
    flipH: false,
    flipV: false,
  };
}

/** The node attributes of a new text box (its content is one empty paragraph). */
export function documentTextBoxFrameAttrs(input: {
  readonly direction: DocumentTextBoxDirection;
  readonly xEmu: number;
  readonly yEmu: number;
  readonly widthEmu: number;
  readonly heightEmu: number;
  readonly relativeHeight: number;
  /** Word's Selection Pane name: "Text Box N". */
  readonly name: string;
}): Record<string, unknown> {
  return {
    floatAnchor: documentTextBoxFloatAnchor(input, input, input.relativeHeight),
    frameBody: documentTextBoxFrameBody(input.direction),
    textBoxOutline: { ...DOCUMENT_TEXT_BOX_DEFAULT_OUTLINE },
    textBoxFill: null,
    frameName: input.name.slice(0, 255),
    // The new box takes its place in the floating stack the same way a shape
    // drawn here does, so Bring Forward / Send Backward may move it.
    zOrderStackComplete: true,
  };
}

/**
 * Where offset (0, 0) of each axis lands on the page, and where the frame is
 * painted, in page CSS px -- what the layout pass measured for one frame.
 */
export interface DocumentTextBoxPaintedFrame {
  readonly bandLeftPx: number;
  readonly bandTopPx: number;
  readonly leftPx: number;
  readonly topPx: number;
}

/**
 * The anchor after the frame is moved by `(dxPx, dyPx)` page px. Word turns a
 * dragged aligned or relative position into an absolute offset on the same
 * origin; an axis already stated by offset just moves.
 */
export function documentTextBoxMovedAnchor(
  anchor: FloatAnchor,
  painted: DocumentTextBoxPaintedFrame,
  dxPx: number,
  dyPx: number,
): FloatAnchor {
  const horizontal = {
    relativeFrom: anchor.horizontal.relativeFrom,
    offsetEmu: Math.round((painted.leftPx + dxPx - painted.bandLeftPx) * EMU_PER_PX),
  };
  const vertical = {
    relativeFrom: anchor.vertical.relativeFrom,
    offsetEmu: Math.round((painted.topPx + dyPx - painted.bandTopPx) * EMU_PER_PX),
  };
  return {
    ...anchor,
    horizontal: dxPx === 0 ? anchor.horizontal : horizontal,
    vertical: dyPx === 0 ? anchor.vertical : vertical,
    simplePos: false,
  };
}

/** The anchor with a new size and (for a top or left grip) a new place. */
export function documentTextBoxResizedAnchor(
  anchor: FloatAnchor,
  painted: DocumentTextBoxPaintedFrame,
  geometry: { readonly leftPx: number; readonly topPx: number; readonly widthPx: number; readonly heightPx: number },
): FloatAnchor {
  const moved = documentTextBoxMovedAnchor(
    anchor,
    painted,
    geometry.leftPx - painted.leftPx,
    geometry.topPx - painted.topPx,
  );
  return {
    ...moved,
    widthEmu: clampEdge(geometry.widthPx * EMU_PER_PX),
    heightEmu: clampEdge(geometry.heightPx * EMU_PER_PX),
  };
}

/** A body text box frame of the Word file (`…#p[N]/txbxframe[k]`): the frames
 * whose own drawing the save writes back (text_box_frame_patch). */
const BODY_FRAME_SOURCE = /^word\/document\.xml#p\[\d+\]\/txbxframe\[\d+\]$/;

/**
 * Whether a frame's box can be edited -- moved, sized, painted, re-aligned:
 * a floating body frame of the Word file, not a member of a group (its group
 * places it) and not locked. Anything else keeps its frame as read.
 */
export function documentTextBoxFrameIsEditable(attrs: Readonly<Record<string, unknown>>): boolean {
  const anchor = normalizeFloatAnchor(attrs.floatAnchor);
  if (!anchor || anchor.groupTransform || anchor.locked || anchor.simplePos) return false;
  const source = attrs.source as { sourceAnchor?: unknown } | null | undefined;
  return typeof source?.sourceAnchor === "string" && BODY_FRAME_SOURCE.test(source.sourceAnchor);
}

/** Word's Wrap Text choices, as the shared Shape Format menu names them. */
export type DocumentTextBoxWrap =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

/** The frame's Wrap Text choice: In Line with Text has no anchor at all. */
export function documentTextBoxWrapOf(floatAnchor: unknown): DocumentTextBoxWrap {
  const anchor = normalizeFloatAnchor(floatAnchor);
  if (!anchor) return "inline";
  switch (anchor.wrap.kind) {
    case "none":
      return anchor.behindDoc ? "behindText" : "inFrontOfText";
    case "square":
    case "tight":
    case "through":
    case "topAndBottom":
      return anchor.wrap.kind;
    default:
      return "inFrontOfText";
  }
}

/** `wp:anchor/@distT|distB` Word writes on a box that wraps text (0.05 in). */
const WRAPPING_TOP_DISTANCE_EMU = 45_720;

/**
 * The frame attributes after a Wrap Text choice. In Line with Text keeps the
 * box's size as `frameExtent` (`wp:inline`) and drops the anchor; any other
 * choice anchors it (an in-line box gets Word's column/paragraph origins at
 * offset 0) and states the wrap element and text layer Word writes for it.
 */
export function documentTextBoxWrapPatch(
  attrs: { readonly floatAnchor?: unknown; readonly frameExtent?: unknown },
  wrap: DocumentTextBoxWrap,
  relativeHeight: number,
): Record<string, unknown> | null {
  const anchor = normalizeFloatAnchor(attrs.floatAnchor);
  const extent = attrs.frameExtent as { widthEmu?: unknown; heightEmu?: unknown } | null | undefined;
  const width = anchor?.widthEmu ?? (typeof extent?.widthEmu === "number" ? extent.widthEmu : null);
  const height = anchor?.heightEmu ?? (typeof extent?.heightEmu === "number" ? extent.heightEmu : null);
  if (width === null || height === null) return null;
  if (wrap === "inline") {
    return anchor ? { floatAnchor: null, frameExtent: { widthEmu: width, heightEmu: height } } : null;
  }
  const base = anchor ?? documentTextBoxFloatAnchor({ xEmu: 0, yEmu: 0 }, { widthEmu: width, heightEmu: height }, relativeHeight);
  const overlay = wrap === "behindText" || wrap === "inFrontOfText";
  const next: FloatAnchor = {
    ...base,
    wrap: overlay
      ? { kind: "none" }
      : wrap === "topAndBottom"
        ? { kind: "topAndBottom" }
        : { kind: wrap, text: "bothSides" },
    behindDoc: wrap === "behindText",
    distTEmu: overlay ? 0 : WRAPPING_TOP_DISTANCE_EMU,
    distBEmu: overlay ? 0 : WRAPPING_TOP_DISTANCE_EMU,
    distLEmu: DEFAULT_SIDE_DISTANCE_EMU,
    distREmu: DEFAULT_SIDE_DISTANCE_EMU,
  };
  return { floatAnchor: next, frameExtent: null };
}

/** Bring Forward / Send Backward on the document's one floating stack. */
export function documentTextBoxZOrderHeight(
  current: number,
  heights: readonly number[],
  action: "bringForward" | "bringToFront" | "sendBackward" | "sendToBack",
): number | null {
  const others = heights.filter((height) => height !== current);
  if (others.length === 0) return null;
  const above = others.filter((height) => height > current);
  const below = others.filter((height) => height < current);
  switch (action) {
    case "bringToFront":
      return above.length ? Math.max(...above) + 1 : null;
    case "bringForward":
      return above.length ? Math.min(...above) + 1 : null;
    case "sendToBack":
      return below.length ? Math.max(0, Math.min(...below) - 1) : null;
    case "sendBackward":
      return below.length ? Math.max(0, Math.max(...below) - 1) : null;
  }
}

/** `bodyPr/@vert` for the Shape Format Text Direction choice, and back. */
export function documentTextBoxVertFor(direction: "horizontal" | "rotate90" | "rotate270" | "stacked"): FloatFrameBody["vert"] {
  return direction === "rotate90" ? "vert" : direction === "rotate270" ? "vert270" : direction === "stacked" ? "wordArtVert" : "horz";
}

export function documentTextBoxDirectionOf(vert: FloatFrameBody["vert"]): "horizontal" | "rotate90" | "rotate270" | "stacked" {
  if (vert === "vert" || vert === "eaVert" || vert === "mongolianVert") return "rotate90";
  if (vert === "vert270") return "rotate270";
  if (vert === "wordArtVert" || vert === "wordArtVertRtl") return "stacked";
  return "horizontal";
}

/** A text box frame's outer shadow (`textBoxShadow`), the shared shape
 * shadow model, or null for anything that is not one. */
export function documentTextBoxShadow(value: unknown): OfficeShapeShadow | null {
  let raw = value;
  if (typeof raw === "string") {
    try {
      raw = JSON.parse(raw);
    } catch {
      return null;
    }
  }
  try {
    return normalizeOfficeShapeShadow(raw) ?? null;
  } catch {
    return null;
  }
}

/** The frame's shadow as its own box shadow: the box's silhouette in the
 * shadow colour, moved `distance` px along `direction` (degrees clockwise
 * from +x, y down; straight down when unstated) and blurred. */
export function documentTextBoxShadowHtmlAttributes(value: unknown): Record<string, string> {
  const shadow = documentTextBoxShadow(value);
  if (!shadow) return {};
  const radians = ((shadow.direction ?? 90) * Math.PI) / 180;
  const x = Math.round(Math.cos(radians) * shadow.distance * 100) / 100 + 0;
  const y = Math.round(Math.sin(radians) * shadow.distance * 100) / 100 + 0;
  const hex = shadow.color.replace("#", "");
  const [r, g, b] = [0, 2, 4].map((at) => Number.parseInt(hex.slice(at, at + 2), 16));
  return {
    "data-text-box-shadow": JSON.stringify(shadow),
    style: `box-shadow:${x}px ${y}px ${shadow.blur}px rgba(${r}, ${g}, ${b}, ${shadow.opacity})`,
  };
}
