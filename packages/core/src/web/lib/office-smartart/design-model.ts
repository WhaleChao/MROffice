/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The SmartArt Design tab's MODEL, with no host and no layout in it.
 *
 * What the browser computes itself for a diagram: reading a diagram's members
 * (marker, frame, colours, style, direction), the outline transforms behind
 * Add / Delete / Promote / Demote / Move and the commands' enabled state, the
 * paint-ownership stamps, Larger / Smaller on one item, and Convert to Shapes
 * / Text. Every one of these reads or rewrites members the page already
 * stores; none of them lays a diagram out.
 *
 * A NEW layout -- insert, change layout, colours, quick style, Reset, any
 * outline change -- is the server's (`./design-ops`, run by the engine host;
 * the browser asks through `./remote`). `./design-ops` re-exports everything
 * here, so server code and the engine's own tests see one module.
 */

import {
  resolveSmartArtColorVariant,
  smartArtColorVariant,
  type SmartArtColorTheme,
} from "./colors";
import { smartArtConnectorsFollowNodes } from "./connector-follow";
import {
  readSmartArtCustom,
  smartArtLayoutBox,
  smartArtMemberAfterBoxEdit,
} from "./custom-geometry";
import {
  DEFAULT_SMARTART_STYLE,
  isSmartArtHiddenMember,
  readSmartArtBlockMarker,
  readSmartArtDirection,
  SMARTART_ASSISTANT_KINDS,
  smartArtAcceptsLevelTwo,
  smartArtLayoutTakesBullets,
  smartArtLayoutTakesHierBranch,
  smartArtSupportsHierarchy,
  type SmartArtHierBranch,
  type OfficeSmartArtBlockLike,
  type OfficeSmartArtHostNode,
  type SmartArtBlockMarker,
  type SmartArtDirection,
  type SmartArtKind,
  type SmartArtNodePaint,
  type SmartArtStyle,
} from "./model";
import {
  MAX_SMARTART_OUTLINE_TEXT_LENGTH,
  addSmartArtOutlineItem,
  demoteSmartArtOutlineItem,
  insertSmartArtOutlineItem,
  mergeSmartArtOutlineItems,
  moveSmartArtOutlineItem,
  promoteSmartArtOutlineItem,
  removeSmartArtOutlineItem,
  updateSmartArtOutlineItemText,
  type SmartArtOutlineInsertPosition,
  type SmartArtOutlineItem,
} from "./outline";
import {
  DEFAULT_SMARTART_STYLE_ID,
  readSmartArtStyleId,
  type SmartArtStyleId,
} from "./styles-model";
import { randomUUID } from "../random-uuid";

/**
 * The two labelled sections PowerPoint's SmartArt Styles gallery shows, in
 * gallery order. The styles themselves live in the shared catalog
 * (`./styles`), so a host adds no list of its own.
 */
export const OFFICE_SMARTART_STYLE_SECTIONS = Object.freeze([
  "bestMatch",
  "threeD",
] as const);

export type OfficeSmartArtStyleSectionId =
  (typeof OFFICE_SMARTART_STYLE_SECTIONS)[number];

/** A diagram's rectangle in its host's own coordinates. */
export interface OfficeSmartArtFrame {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** One diagram: the marker every member agrees on, and the members. */
export interface OfficeSmartArtGroup<T> {
  marker: SmartArtBlockMarker;
  members: T[];
}



function finiteNumber(node: OfficeSmartArtBlockLike, key: string): number | null {
  const value = (node as Record<string, unknown>)[key];
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

/** The host group-id normalisation every editor's own grouping already uses. */
function nodeGroupId(node: OfficeSmartArtBlockLike): string | undefined {
  const value = (node as Record<string, unknown>).groupId;
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/**
 * The whole diagram the given member belongs to: every node sharing its
 * `groupId`, refused unless ALL of them carry the same marker (`groupId` +
 * `kind`). A half-marked group is never treated as a diagram — the tab would
 * otherwise rebuild a set of shapes that only looks like one.
 */
export function readSmartArtGroupMembers<T extends OfficeSmartArtHostNode>(
  nodes: readonly T[],
  primary: T | null | undefined,
): OfficeSmartArtGroup<T> | null {
  const marker = readSmartArtBlockMarker(primary);
  if (!marker) return null;
  const members = nodes.filter((node) => nodeGroupId(node) === marker.groupId);
  if (
    members.length === 0 ||
    members.some((node) => {
      const candidate = readSmartArtBlockMarker(node);
      return (
        candidate?.groupId !== marker.groupId || candidate.kind !== marker.kind
      );
    })
  ) {
    return null;
  }
  return { marker, members };
}

/**
 * How many outline items a diagram's members draw: what a Change Colors
 * variation that paints the shapes in turn spreads its colours across.
 */
export function smartArtDiagramItemCount(members: readonly OfficeSmartArtBlockLike[]): number {
  let count = 0;
  for (const member of members) {
    const record = member as Record<string, unknown>;
    const index = record.smartArtItemIndex;
    if (record.smartArtRole === "node" && typeof index === "number" && Number.isInteger(index)) {
      count = Math.max(count, index + 1);
    }
  }
  return Math.max(1, count);
}

/** The members' bounding box, or null when no member has usable geometry. */
/**
 * The frame an Office-drawn diagram states with its first, unpainted member
 * (the layout's graphic frame): Office lays a diagram out in that frame and
 * lets shapes stand outside it (Vertical Curved List's arc, a turned shape's
 * box), so the members' extent is not the frame. Null when the first member
 * is not such a frame or does not hold the centre of every other member.
 */
function statedSmartArtFrame(members: readonly OfficeSmartArtBlockLike[]): OfficeSmartArtFrame | null {
  const first = members[0] as Record<string, unknown> | undefined;
  if (
    !first ||
    first.type !== "shape" ||
    first.shape !== "rectangle" ||
    first.fill !== "none" ||
    first.stroke !== "none" ||
    (first.smartArtRole !== undefined && first.smartArtRole !== "connector") ||
    first.smartArtCustom !== undefined ||
    (typeof first.text === "string" && first.text !== "")
  ) {
    return null;
  }
  const left = finiteNumber(first as OfficeSmartArtBlockLike, "left");
  const top = finiteNumber(first as OfficeSmartArtBlockLike, "top");
  const width = finiteNumber(first as OfficeSmartArtBlockLike, "width");
  const height = finiteNumber(first as OfficeSmartArtBlockLike, "height");
  if (left === null || top === null || width === null || height === null || width <= 0 || height <= 0) return null;
  for (const member of members.slice(1)) {
    const record = member as Record<string, unknown>;
    if (record.smartArtRole !== "node" && record.smartArtRole !== "label") continue;
    const ml = finiteNumber(member, "left");
    const mt = finiteNumber(member, "top");
    const mw = finiteNumber(member, "width");
    const mh = finiteNumber(member, "height");
    if (ml === null || mt === null || mw === null || mh === null) continue;
    const cx = ml + mw / 2;
    const cy = mt + mh / 2;
    if (cx < left - 1 || cx > left + width + 1 || cy < top - 1 || cy > top + height + 1) return null;
  }
  return { left, top, width, height };
}

/**
 * The member a diagram states its frame with: an unpainted rectangle with no
 * words that belongs to no item (the engine's and the Office conversion's
 * first member). It draws nothing; it only holds the frame.
 */
function isSmartArtFrameMember(member: unknown): boolean {
  if (!member || typeof member !== "object") return false;
  const record = member as Record<string, unknown>;
  return (
    record.type === "shape" &&
    record.shape === "rectangle" &&
    record.fill === "none" &&
    record.stroke === "none" &&
    (record.smartArtRole === undefined || record.smartArtRole === "connector") &&
    record.smartArtItemId === undefined &&
    record.smartArtCustom === undefined &&
    (typeof record.text !== "string" || record.text === "")
  );
}

export function readSmartArtGroupFrame(
  members: readonly OfficeSmartArtBlockLike[],
): OfficeSmartArtFrame | null {
  const stated = statedSmartArtFrame(members);
  if (stated) return stated;
  const boxes = members.flatMap((node) => {
    const left = finiteNumber(node, "left");
    const top = finiteNumber(node, "top");
    const width = finiteNumber(node, "width");
    const height = finiteNumber(node, "height");
    if (left === null || top === null || width === null || height === null) return [];
    // A shape made Larger / Smaller or moved inside the graphic counts at its
    // LAYOUT box: the graphic's frame does not follow one shape's resize, so
    // Reset Graphic re-lays it out in the frame it had (PPT-B082: shrinking
    // an edge shape narrowed the frame and Reset rebuilt every shape smaller).
    const layout = smartArtLayoutBox(
      { left, top, width, height },
      readSmartArtCustom((node as Record<string, unknown>).smartArtCustom),
    );
    return [
      {
        left: layout.left,
        top: layout.top,
        right: layout.left + layout.width,
        bottom: layout.top + layout.height,
      },
    ];
  });
  if (boxes.length === 0) return null;
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.right));
  const bottom = Math.max(...boxes.map((box) => box.bottom));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * The colour set the diagram currently wears, read off its OWN members (never
 * inferred from the style id): a rebuild keeps the author's colours whatever
 * quick style is applied. Each field falls back to the engine default when no
 * member of that role carries it.
 */
/**
 * The colour a connector member carries. A flow arrow is a FILLED shape whose
 * outline is the quick style's line (white for most styles), so its colour is
 * its fill; only a hollow or hairline connector (fill `none`) carries its
 * colour on the stroke. Reading the stroke first turned every rebuild's arrows
 * white -- an Add Shape on a Basic Cycle made the arrows vanish into the page.
 */
export function smartArtConnectorColor(
  fill: unknown,
  stroke: unknown,
  fallback: string,
): string {
  const colour = (value: unknown) =>
    typeof value === "string" && value.length > 0 && value !== "none" && value !== "transparent"
      ? value
      : null;
  return colour(fill) ?? colour(stroke) ?? fallback;
}

/**
 * A member's fill when it is one a viewer sees: a colour, not "none" /
 * "transparent", on a member that is not fully transparent (`opacity` 0).
 */
export function smartArtVisibleFill(fill: unknown, opacity?: unknown): string | undefined {
  if (typeof fill !== "string") return undefined;
  const value = fill.trim();
  if (!value || value === "none" || value === "transparent") return undefined;
  if (typeof opacity === "number" && opacity <= 0) return undefined;
  return value;
}

/**
 * The shape whose paint IS the diagram's colour set: the fill most of the
 * item shapes wear, preferring one the author did not paint by hand.
 *
 * The readers took the FIRST item's shape, so a diagram whose first shape the
 * author had painted red came back from any rebuild (Add Shape, a layout
 * switch) with every other shape red too. A shape painted on its own is that
 * shape's formatting, not the diagram's (Microsoft Support, "Learn more about
 * SmartArt graphics": a shape's own formatting is carried over by itself).
 * One rule for the three hosts; each says how it reads a shape's fill.
 */
export function smartArtPaletteNode<T>(
  nodes: readonly T[],
  fillOf: (node: T) => unknown,
  /**
   * The diagram's other shapes, read only when NO item shape wears a visible
   * fill: Opposing Arrows' items are its words (Office's `revTx`, no paint),
   * so the colour set is on its arrows. Reading a paintless word box instead
   * handed the next restyle "none", or the white of a transparent box, as the
   * diagram's colour.
   */
  others: readonly T[] = [],
): T | undefined {
  const painted = nodes.filter((node) => smartArtVisibleFill(fillOf(node)) !== undefined);
  if (nodes.length && !painted.length) {
    const fallback = others.filter((node) => smartArtVisibleFill(fillOf(node)) !== undefined);
    if (fallback.length) return smartArtPaletteNode(fallback, fillOf);
    return nodes[0];
  }
  nodes = painted;
  if (nodes.length <= 1) return nodes[0];
  const counts = new Map<string, number>();
  for (const node of nodes) {
    const fill = String(fillOf(node) ?? "").toLowerCase();
    counts.set(fill, (counts.get(fill) ?? 0) + 1);
  }
  const claimed = (node: T) =>
    (node as unknown as Record<string, unknown>).smartArtAuthorPaint === true;
  let best: T | undefined;
  let bestScore = -1;
  for (const node of nodes) {
    // Most common fill first; among equals, a shape nobody painted by hand.
    const score =
      (counts.get(String(fillOf(node) ?? "").toLowerCase()) ?? 0) * 2 +
      (claimed(node) ? 0 : 1);
    if (score > bestScore) {
      best = node;
      bestScore = score;
    }
  }
  return best;
}

export function readSmartArtGroupColors(
  members: readonly OfficeSmartArtBlockLike[],
): SmartArtStyle {
  const byRole = (type: string, role: SmartArtBlockMarker["role"]) =>
    members.find(
      (node) =>
        (node as Record<string, unknown>).type === type &&
        readSmartArtBlockMarker(node)?.role === role,
    ) as Record<string, unknown> | undefined;
  const shapes = members.filter(
    (candidate) => (candidate as Record<string, unknown>).type === "shape",
  );
  const node = smartArtPaletteNode(
    shapes.filter((candidate) => readSmartArtBlockMarker(candidate)?.role === "node"),
    (candidate) =>
      smartArtVisibleFill(
        (candidate as Record<string, unknown>).fill,
        (candidate as Record<string, unknown>).opacity,
      ),
    shapes,
  ) as Record<string, unknown> | undefined;
  const connector = byRole("shape", "connector");
  const label = byRole("text", "label");
  const labelStyle = label?.style as Record<string, unknown> | undefined;
  return {
    nodeFill:
      smartArtVisibleFill(node?.fill, node?.opacity) ?? DEFAULT_SMARTART_STYLE.nodeFill,
    nodeStroke:
      typeof node?.stroke === "string"
        ? node.stroke
        : DEFAULT_SMARTART_STYLE.nodeStroke,
    connector: smartArtConnectorColor(
      connector?.fill,
      connector?.stroke,
      DEFAULT_SMARTART_STYLE.connector,
    ),
    textColor:
      typeof labelStyle?.color === "string"
        ? labelStyle.color
        : DEFAULT_SMARTART_STYLE.textColor,
    fontSize:
      typeof labelStyle?.fontSize === "number"
        ? labelStyle.fontSize
        : DEFAULT_SMARTART_STYLE.fontSize,
  };
}

/**
 * The quick style the members agree on, so a gallery can check it. Members that
 * disagree — an older diagram, or one caught mid-edit — report none rather than
 * letting an arbitrary member decide.
 */
export function readSmartArtGroupStyleId(
  members: readonly OfficeSmartArtBlockLike[],
): SmartArtStyleId | null {
  const ids = new Set(members.map((node) => readSmartArtStyleId(node)));
  const [only] = [...ids];
  return ids.size === 1 ? only : null;
}

/**
 * The direction the diagram runs ({@link SmartArtDirection}): Right to Left
 * when its members say so. Every member of a mirrored diagram carries the
 * mark, so one saying it is enough; a host whose members could not all keep
 * it must not flip the diagram back on its next rebuild.
 */
export function readSmartArtGroupDirection(
  members: readonly unknown[],
): SmartArtDirection {
  return members.some((member) => readSmartArtDirection(member) === "rtl")
    ? "rtl"
    : "ltr";
}

/**
 * SmartArt Design > Right to Left, as a toggle (Microsoft Support, "Change the
 * direction of a SmartArt graphic": to return to the original direction,
 * click Right to Left again).
 */
export function smartArtToggledDirection(
  direction: SmartArtDirection,
): SmartArtDirection {
  return direction === "rtl" ? "ltr" : "rtl";
}

export interface RebuildSmartArtGroupInput<T extends OfficeSmartArtHostNode> {
  /** The diagram's current members; their ids are reused so a host's own
   *  per-node state (selection, z-order, picture custody) survives. */
  members: readonly T[];
  /** The diagram id every rebuilt member keeps. */
  groupId: string;
  /** The layout to build — the current one for an outline edit. */
  kind: SmartArtKind;
  /** Where the rebuilt diagram must land; normally the members' own box. */
  frame: OfficeSmartArtFrame;
  /** The outline the rebuilt diagram carries, in order. */
  outline: readonly SmartArtOutlineItem[];
  /** Colours; defaults to the members' own ({@link readSmartArtGroupColors}). */
  colors?: SmartArtStyle;
  /**
   * The quick style to build with. Omitted means the one the members wear:
   * Office carries a graphic's SmartArt Style across Add Shape, Promote and a
   * layout switch alike (Microsoft Support, "Learn more about SmartArt
   * graphics": switching layouts carries "colors, styles, effects, and text
   * formatting" over). It used to mean the engine's default, so every
   * structural command in the deck and the worksheet flattened a 3-D or
   * gradient style back to `simple1`. Pass one to rebuild at another style.
   */
  styleId?: SmartArtStyleId;
  /**
   * Which way the rebuilt diagram runs. Omitted means the way the members run
   * ({@link readSmartArtGroupDirection}), so Add Shape, Promote, a layout
   * switch, a colour or a style keep a Right to Left diagram mirrored; only
   * the Right to Left command itself passes one.
   */
  direction?: SmartArtDirection;
  /**
   * Whether the rebuild may change the SET of items (add / delete). Forwarded
   * to {@link preserveSmartArtPictureCustody}, which refuses to drop a picture
   * the author actually supplied unless the edit is allowed to.
   */
  allowItemChanges?: boolean;
  /** Fresh ids once the members' own are exhausted; defaults to `randomUUID`. */
  newId?: () => string;
  /**
   * Each item's own fill and line for this rebuild (a Change Colors variation
   * that paints the shapes in turn, `./colors`). Wins over the author's paint.
   */
  nodePaint?: readonly (SmartArtNodePaint | null)[];
  /**
   * Whether a shape the author painted keeps its own fill and line through
   * the rebuild (default true). Office carries a shape's colours across Add
   * Shape, Promote and a layout switch (Microsoft Support, "Learn more about
   * SmartArt graphics": "most of your text and other content, colors, styles,
   * effects, and text formatting are automatically carried over"). Change
   * Colors and Reset Graphic pass false: they are the command that repaints.
   */
  keepAuthorPaint?: boolean;
  /**
   * The document's theme colours, for a diagram that wears a colour
   * VARIATION (`smartArtColors`): with them, the rebuild re-spreads the
   * variation over the new item count the way Office does (Add Shape on a
   * Gradient Range diagram re-grades every shape). Without them each item
   * keeps the colour it had and a new item takes the diagram's first.
   */
  /** The workbook's / document's body face (`+mn-lt`), which sets how tall the labels' lines stand. */
  bodyFace?: string | null;
  /** The face the workbook theme gives Chinese words (`SmartArtColorTheme.eastAsianFace`). */
  eastAsianFace?: string | null;
  colorTheme?: Partial<SmartArtColorTheme> | null;
  /**
   * The colour variation (Office's id) a layout drawn from Office's own
   * definition is painted with; `null` is Office's default, absent keeps the
   * one the members wear.
   */
  colorsId?: string | null;
  /**
   * The host draws layouts Office's own definition covers from it (the
   * Presentation host): such a diagram is rebuilt in its frame as Office lays
   * it out and painted from Office's colours and style definitions.
   */
  officeLayout?: boolean;
  /**
   * Reset Graphic: drop every shape's move / resize (`smartArtCustom`) instead
   * of carrying it into the rebuilt layout. A layout SWITCH drops them too.
   */
  resetCustomGeometry?: boolean;
  /**
   * The Office alias this diagram is laid out from (`smartArtLayoutUrn`).
   * Omitted keeps the one the members record ({@link readSmartArtGroupLayoutUrn})
   * while `kind` stays the members' kind; a layout switch to another kind
   * drops it. `null` drops it outright.
   */
  layoutUrn?: string | null;
}

/**
 * The fill and line each outline item's shape was given BY THE AUTHOR, in
 * outline order, for `buildSmartArt`'s `nodePaint`.
 *
 * Only a shape that says the paint is the author's (`smartArtAuthorPaint`)
 * contributes; every other one wears the palette and is rebuilt from it. The
 * member may be a slide block (`fill` / `stroke`), a Docs group member
 * (`fill` / `outline`) or a sheet object (its `shape`'s `fill` / `stroke`).
 * Keyed on the durable item id, which is what survives a relayout.
 */
export function smartArtAuthorNodePaint(
  previous: readonly unknown[],
  outline: readonly SmartArtOutlineItem[],
): (SmartArtNodePaint | null)[] {
  const colour = (value: unknown) =>
    typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined;
  const byItem = new Map<string, SmartArtNodePaint>();
  for (const member of previous) {
    const record = member as Record<string, unknown> | null;
    if (
      !record ||
      typeof record !== "object" ||
      record.smartArtAuthorPaint !== true ||
      record.smartArtRole !== "node" ||
      typeof record.smartArtItemId !== "string" ||
      byItem.has(record.smartArtItemId)
    ) {
      continue;
    }
    const shape =
      record.shape && typeof record.shape === "object"
        ? (record.shape as Record<string, unknown>)
        : record;
    const fill = colour(shape.fill);
    const stroke = colour(shape.stroke ?? shape.outline);
    if (fill || stroke) {
      byItem.set(record.smartArtItemId, {
        ...(fill ? { fill } : {}),
        ...(stroke ? { stroke } : {}),
      });
    }
  }
  return outline.map((item) => byItem.get(item.id) ?? null);
}

/**
 * Reset Shape (Office's shape menu, `ContextMenuSmartArtEditSmartArt`): the
 * members of ONE item without what the author did to that shape -- its own
 * fill and line (`smartArtAuthorPaint`) and its move / resize
 * (`smartArtCustom`). Microsoft Support, "Restore the default layout and
 * color of your SmartArt graphic": right-click a shape and Reset Shape
 * "to remove customizations from" it; Reset Graphic does the same to every
 * shape. The host then rebuilds the diagram at its own layout, which gives
 * that shape the palette and its layout box back while every other shape
 * keeps its own paint and placement.
 */
/** What Reset Shape takes off an item's shapes (`smartArtMembersWithShapeReset`). */
const SMARTART_SHAPE_RESET_KEYS = [
  "smartArtAuthorPaint",
  "smartArtCustom",
  "smartArtAuthorShape",
  "smartArtAuthorFormat",
] as const;

/**
 * Whether Reset Shape has anything to take off the item's shapes: its own
 * fill or line, geometry, effects, move or resize, or the angle it was turned
 * to. Word, PowerPoint and Excel do nothing for a shape that has none; the
 * diagram is not rebuilt (a SmartArt the file holds stays the file's, not
 * converted by a command that changes nothing).
 */
export function smartArtShapeHasOwnReset(members: readonly unknown[], itemId: string): boolean {
  return members.some((member) => {
    if (!member || typeof member !== "object") return false;
    const record = member as Record<string, unknown>;
    // A sheet object keeps its paint on its `shape`, a slide block or a
    // Word member on itself.
    const shape = record.shape && typeof record.shape === "object"
      ? (record.shape as Record<string, unknown>)
      : {};
    if (record.smartArtItemId !== itemId) return false;
    if (record.smartArtRole !== "node" && record.smartArtRole !== "label") return false;
    const own = (value: unknown) =>
      value !== undefined && value !== null && value !== false &&
      !(typeof value === "object" && Object.keys(value as object).length === 0);
    if (SMARTART_SHAPE_RESET_KEYS.some((key) => own(record[key]) || own(shape[key]))) return true;
    const rotation = Number(record.rotation ?? shape.rotation ?? 0);
    return Number.isFinite(rotation) && rotation % 360 !== 0;
  });
}

export function smartArtMembersWithShapeReset<T>(
  members: readonly T[],
  itemId: string,
): T[] {
  // The shape's paint lives on the member (a slide block, a Docs member) or
  // on its `shape` (a sheet object). The reset shape takes the palette the
  // OTHER item shapes wear, so the rebuild cannot read its old colour back as
  // the diagram's.
  const paintOf = (record: Record<string, unknown>) =>
    record.shape && typeof record.shape === "object"
      ? (record.shape as Record<string, unknown>)
      : record;
  const palette = smartArtPaletteNode(
    members.filter((member) => {
      const record = member as unknown as Record<string, unknown>;
      return record.smartArtRole === "node" && record.smartArtItemId !== itemId;
    }),
    (member) => paintOf(member as unknown as Record<string, unknown>).fill,
  );
  const paletteRecord = palette
    ? paintOf(palette as unknown as Record<string, unknown>)
    : null;
  return members.map((member) => {
    const record = member as unknown as Record<string, unknown>;
    if (
      record.smartArtItemId !== itemId ||
      (record.smartArtRole !== "node" && record.smartArtRole !== "label")
    ) {
      return member;
    }
    const {
      smartArtAuthorPaint: _paint,
      smartArtCustom: _custom,
      smartArtAuthorShape: _shape,
      // …and its effects (PPT-A010): "shape geometry, fill, line and effects"
      // are what Reset removes.
      smartArtAuthorFormat: _format,
      // …and the angle it was turned to (`dgm:prSet/@custAng`), which a
      // same-layout rebuild otherwise carries (`reapplySmartArtRotation`).
      rotation: _rotation,
      ...rest
    } = record;
    if (record.smartArtRole !== "node" || !paletteRecord) {
      return rest as unknown as T;
    }
    const repaint = Object.fromEntries(
      (["fill", "stroke", "outline"] as const).flatMap((key) =>
        paletteRecord[key] === undefined ? [] : [[key, paletteRecord[key]]],
      ),
    );
    return (
      record.shape && typeof record.shape === "object"
        ? { ...rest, shape: { ...(record.shape as object), ...repaint } }
        : { ...rest, ...repaint }
    ) as unknown as T;
  });
}

/**
 * The label text formatting a slide label block keeps on its `style`: the
 * WHOLE-LABEL storage. Home > Font on the graphic's border writes these
 * (`lib/presentation-smartart-text`), the engine builds a label with them,
 * the native SmartArt writer states them on every run of the data point.
 * A word formatted inside a label lives on its `runs` (per-run overrides
 * layered over the style); those ride through a rebuild, and a text-pane
 * edit of the label keeps them on the characters it did not change.
 */
export const SMARTART_LABEL_FORMAT_KEYS = [
  "fontFamily",
  "fontSize",
  "bold",
  "italic",
  "underline",
  "strike",
  "color",
] as const;



/** Reset Graphic: the default colour set and the default quick style, together. */
export const SMARTART_RESET_CHANGE: {
  colors: SmartArtStyle;
  styleId: SmartArtStyleId;
  authorPaint: false;
  colorsId: null;
} = {
  colors: DEFAULT_SMARTART_STYLE,
  styleId: DEFAULT_SMARTART_STYLE_ID,
  // Reset is the author saying "the catalog's", which is the one recolour that
  // CLEARS `smartArtAuthorPaint` instead of setting it.
  authorPaint: false,
  // And the colour variation goes back to Office's default (`accent1_2`),
  // which a member states by not naming one.
  colorsId: null,
};

/**
 * Reset Graphic's restyle for one diagram. Office's reset removes the
 * shapes' own fills, lines and effects and the quick style, and keeps the
 * colour variation (`csTypeId`) the diagram wears (Microsoft Support,
 * "Restore the default layout and color of your SmartArt graphic": colour
 * changes come off individual shapes, the graphic's colour set is not on the
 * removed list). So the diagram is repainted from its OWN variation --
 * Office's default `accent1_2` when it names none -- resolved against the
 * document's live theme: the reset used to paint every shape the 2013
 * palette's #4472C4 and drop the variation, on a deck whose accent 1 is
 * #156082. With no theme and no variation, the historical catalog default.
 */
export function smartArtResetChange(
  members: readonly OfficeSmartArtBlockLike[],
  theme?: Partial<SmartArtColorTheme> | null,
): {
  colors: SmartArtStyle;
  styleId: SmartArtStyleId;
  authorPaint: false;
  colorsId: string | null;
  nodePaint: SmartArtNodePaint[] | null;
  colorTheme?: Partial<SmartArtColorTheme> | null;
} {
  const colorsId = readSmartArtGroupColorsId(members);
  const variant = smartArtColorVariant(colorsId ?? "accent1_2", { all: true });
  if (!variant || (!theme && !colorsId)) return { ...SMARTART_RESET_CHANGE, nodePaint: null };
  const resolved = resolveSmartArtColorVariant(variant, smartArtDiagramItemCount(members), theme);
  return {
    ...SMARTART_RESET_CHANGE,
    colors: resolved.colors,
    nodePaint: resolved.nodePaint,
    // A diagram on the default variation keeps stating none, as before.
    colorsId,
    ...(theme ? { colorTheme: theme } : {}),
  };
}

/**
 * Whose paint a `(colors, styleId)` command applies, in the ONE convention all
 * three hosts already spell their Design-tab commands in: Change Colors passes
 * `colors` with a null `styleId`, the Styles gallery passes a `styleId` with
 * null `colors`, and Reset Graphic passes BOTH.
 *
 * So: colours alone are the author's choice; colours WITH a style is Reset,
 * which hands the diagram back to the catalog palette; a style alone says
 * nothing about colour at all and leaves the field exactly as it was.
 * `undefined` is that third answer, and it is not the same as `false`.
 */
export function smartArtRecolourIsAuthors(
  colors: SmartArtStyle | null,
  styleId: SmartArtStyleId | null,
): boolean | undefined {
  if (!colors) return undefined;
  return styleId ? false : true;
}

/**
 * State — or unstate — that a diagram's colours are the AUTHOR'S CHOICE.
 *
 * `smartArtAuthorPaint` is §5.17's field: optional, only ever `true`, only ever
 * on a `node`. Without it the export can only tell a chosen colour from the
 * catalog palette by disagreement between the diagram's own nodes (§5.3), so a
 * diagram whose every node wears ONE chosen colour reached Office without its
 * colours. §5.17 closed the round trip — the reopen states the field because
 * the reader knows Office paints a point only when somebody chose the paint —
 * and this closes the other half: a recolour performed in MRAI states it on the
 * spot, so the FIRST export carries the colours instead of only the one after a
 * reopen.
 *
 * One helper for all three hosts, because all three member shapes spell the
 * role the same way. It is deliberately NOT keyed on `readSmartArtBlockMarker`:
 * that reader needs a `groupId` on the record, which a Docs group MEMBER does
 * not carry (its group node does).
 *
 * `false` REMOVES the field rather than writing `false`, because the page model
 * only ever stores `true` and a page stored before the field simply does not
 * carry it — which is §5.3's heuristic exactly as it was, and why there is no
 * migration.
 */
export function withSmartArtAuthorPaint<T>(
  members: readonly T[],
  authorPaint: boolean,
): T[] {
  return members.map((member) => {
    const record = member as unknown as Record<string, unknown>;
    if (record.smartArtRole !== "node") return member;
    if (authorPaint) {
      return record.smartArtAuthorPaint === true
        ? member
        : ({ ...record, smartArtAuthorPaint: true } as unknown as T);
    }
    if (record.smartArtAuthorPaint === undefined) return member;
    const { smartArtAuthorPaint: _cleared, ...rest } = record;
    return rest as unknown as T;
  });
}

/**
 * Carry "these colours are the AUTHOR's" across a REBUILD.
 *
 * Every structural command — Change Layout, Right to Left, Add / Promote /
 * Demote / Delete — rebuilds the diagram's members from the engine, and an
 * engine-built member claims nothing about whose paint it wears. The rebuild
 * does keep the COLOURS (it is handed the diagram's own `colors`), so the
 * diagram looked identical and the page silently stopped saying the colours
 * were chosen: the very next export dropped them and the author's red came
 * back as the catalog palette on reopen. That is §5.17's defect exactly,
 * reintroduced by any command after the recolour, and measured in all three
 * hosts (SMARTART-IMPORT.md §5.68).
 *
 * Keyed on the DURABLE item id, for the same reason
 * {@link smartArtSlotPicturesByItemId} is: a relayout moves positions and
 * indices, and the item id is what survives it. An item that is gone from the
 * outline takes its claim with it, which is the author deleting the item.
 *
 * One helper for all three hosts, like the stamp itself: it reads `smartArtRole`
 * and `smartArtItemId`, which all three member shapes spell identically.
 *
 * A command that has its OWN answer about whose paint this is — Change Colors
 * (`true`), Reset Graphic (`false`) — applies {@link withSmartArtAuthorPaint}
 * AFTER this, and overrides it.
 */
export function preserveSmartArtAuthorPaint<T>(
  previous: readonly unknown[],
  rebuilt: readonly T[],
): T[] {
  const claimed = new Set<string>();
  for (const member of previous) {
    const record = member as Record<string, unknown> | null;
    if (!record || typeof record !== "object") continue;
    if (record.smartArtAuthorPaint !== true || record.smartArtRole !== "node") {
      continue;
    }
    const itemId = record.smartArtItemId;
    if (typeof itemId === "string" && itemId) claimed.add(itemId);
  }
  if (claimed.size === 0) return [...rebuilt];
  return rebuilt.map((member) => {
    const record = member as unknown as Record<string, unknown>;
    if (record.smartArtRole !== "node") return member;
    return claimed.has(String(record.smartArtItemId))
      ? ({ ...record, smartArtAuthorPaint: true } as unknown as T)
      : member;
  });
}

/**
 * The text a still-focused outline input holds when a structural command fires,
 * so the edit is not lost by the rebuild that follows it.
 */
export interface SmartArtPendingOutlineText {
  itemId: string;
  text: string;
}

function withPendingText(
  items: readonly SmartArtOutlineItem[],
  pending: SmartArtPendingOutlineText | undefined,
): readonly SmartArtOutlineItem[] | null {
  if (!pending) return items;
  return updateSmartArtOutlineItemText(
    items,
    pending.itemId,
    pending.text.slice(0, MAX_SMARTART_OUTLINE_TEXT_LENGTH),
  );
}

/** The outline transforms the tab's four structural commands perform. */
export interface SmartArtOutlineChange {
  items: SmartArtOutlineItem[];
  selectedItemId: string;
}

/**
 * Add Shape: a new item after the selected one, carrying `label`.
 *
 * An add past what the layout draws (Office's capped layouts draw the first
 * N items) is kept, as Office keeps it: the item is in the outline and the
 * Text Pane, and the rebuild keeps it as a hidden member ("unused text does
 * not appear, but remains available if you switch layouts").
 */
export function smartArtOutlineAfterAdd(
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  label: string,
  kind: SmartArtKind,
  pending?: SmartArtPendingOutlineText,
  newId: () => string = () => randomUUID(),
): SmartArtOutlineChange | null {
  const current = withPendingText(items, pending);
  if (!current) return null;
  const id = newId();
  const next = addSmartArtOutlineItem(current, selectedItemId, "after", {
    id,
    text: label,
    level: 0,
  });
  if (!next) return null;
  return { items: next, selectedItemId: id };
}

export type { SmartArtOutlineInsertPosition };

/**
 * The Design tab's Add Shape split button and the right-click Add Shape menu:
 * one new item {@link insertSmartArtOutlineItem | relative to} the selected
 * one. Microsoft's `SmartArtNode.AddNode` names the positions (`After` is the
 * plain Add Shape). Above and Below change the nesting, so a layout that does
 * not nest refuses them the way it refuses Demote. An item past what the
 * layout draws is kept as {@link smartArtOutlineAfterAdd} keeps it.
 */
export function smartArtOutlineAfterInsert(
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  position: SmartArtOutlineInsertPosition,
  label: string,
  kind: SmartArtKind,
  pending?: SmartArtPendingOutlineText,
  newId: () => string = () => randomUUID(),
): SmartArtOutlineChange | null {
  if (
    (position === "above" || position === "below") &&
    !smartArtSupportsHierarchy(kind)
  ) {
    return null;
  }
  if (position === "assistant" && !smartArtLayoutHasAssistant(kind)) return null;
  // Add Bullet follows the layout's definition (`bulletEnabled`), not its
  // nesting: Basic Block List takes bullets in its blocks, an org chart none.
  if (position === "bullet" && !smartArtLayoutTakesBullets(kind)) return null;
  const current = withPendingText(items, pending);
  if (!current) return null;
  const id = newId();
  const next = insertSmartArtOutlineItem(current, selectedItemId, position, {
    id,
    text: label,
    level: 0,
  });
  if (!next) return null;
  return { items: next, selectedItemId: id };
}

/** Move Up / Move Down: reorder among siblings, selection follows the item. */
export function smartArtOutlineAfterMove(
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  direction: "up" | "down",
  kind: SmartArtKind,
  pending?: SmartArtPendingOutlineText,
): SmartArtOutlineChange | null {
  const current = withPendingText(items, pending);
  if (!current) return null;
  const next = moveSmartArtOutlineItem(current, selectedItemId, direction);
  if (!next) return null;
  return { items: next, selectedItemId };
}

/**
 * Organization chart layouts, the only ones Office offers Add Assistant on
 * ("Add Assistant is available only for organization chart layouts. It is not
 * available for hierarchy layouts, such as Hierarchy" -- Microsoft Support,
 * "Create an organization chart"). The hierarchy engine hangs an assistant off
 * its manager's trunk on a row above the manager's other children, and the
 * DiagramML writers write it as `dgm:pt type="asst"`.
 */
export function smartArtLayoutHasAssistant(kind: SmartArtKind): boolean {
  return SMARTART_ASSISTANT_KINDS.has(kind);
}

/**
 * Which structural commands the selected item accepts, computed by running the
 * very transforms the commands run. ONE answer for the Design tab and the
 * right-click menu in every host, so the two can never disagree.
 */
export interface SmartArtOutlineCommandState {
  addAfter: boolean;
  addBefore: boolean;
  addAbove: boolean;
  addBelow: boolean;
  addAssistant: boolean;
  /** Add Bullet (SmartArt Design ▸ Create Graphic). */
  addBullet: boolean;
  /** Layout (Standard / Both / Left / Right Hanging) on an org chart box. */
  orgChartLayout: boolean;
  promote: boolean;
  demote: boolean;
  moveUp: boolean;
  moveDown: boolean;
}

export const SMARTART_NO_OUTLINE_COMMANDS: SmartArtOutlineCommandState =
  Object.freeze({
    addAfter: false,
    addBefore: false,
    addAbove: false,
    addBelow: false,
    addAssistant: false,
    addBullet: false,
    orgChartLayout: false,
    promote: false,
    demote: false,
    moveUp: false,
    moveDown: false,
  });

export function smartArtOutlineCommandState(
  items: readonly SmartArtOutlineItem[] | null | undefined,
  selectedItemId: string | null | undefined,
  kind: SmartArtKind,
): SmartArtOutlineCommandState {
  if (!items?.length || !selectedItemId) return SMARTART_NO_OUTLINE_COMMANDS;
  const probe = () => "__smartart_probe__";
  const insert = (position: SmartArtOutlineInsertPosition) =>
    smartArtOutlineAfterInsert(
      items,
      selectedItemId,
      position,
      "",
      kind,
      undefined,
      probe,
    ) !== null;
  return {
    addAfter: insert("after"),
    addBefore: insert("before"),
    addAbove: insert("above"),
    addBelow: insert("below"),
    addAssistant: insert("assistant"),
    addBullet: insert("bullet"),
    orgChartLayout: smartArtOutlineAfterHierBranch(items, selectedItemId, "std", kind) !== null,
    promote: smartArtOutlineAfterPromote(items, selectedItemId, kind) !== null,
    demote: smartArtOutlineAfterDemote(items, selectedItemId, kind) !== null,
    moveUp: smartArtOutlineAfterMove(items, selectedItemId, "up", kind) !== null,
    moveDown:
      smartArtOutlineAfterMove(items, selectedItemId, "down", kind) !== null,
  };
}

/**
 * Reset > Convert > Convert to Shapes (PowerPoint and Excel; Word has no such
 * command — Microsoft Support, "Convert your SmartArt graphic into shapes").
 *
 * Office splits the graphic into ordinary shapes that keep every visible
 * property (geometry, paint, text) and can never become SmartArt again. In
 * MRAI a member IS an ordinary shape (or text / picture) plus the durable
 * SmartArt identity every field of which is spelled `smartArt…`
 * ({@link OfficeSmartArtNodeIdentity}), so releasing it removes exactly those
 * fields and nothing else: what was drawn stays drawn, and the export no
 * longer writes a diagram for it.
 *
 * `keepGroup` keeps the shared `groupId`, which a host whose page model groups
 * by it (the deck) turns into the group Office's result is; a host with no
 * shape group model (the worksheet) drops it and says so.
 *
 * An item the layout does not draw (`smartArtHidden`) has no shape to become:
 * Office's result is the shapes it drew, and the unused text goes with the
 * diagram. Nor does the graphic's own frame (the unpainted rectangle a
 * diagram states its frame with, {@link statedSmartArtFrame}'s member): a
 * graphic frame is not a shape of the drawing, so PowerPoint's group is the
 * drawn shapes' extent (SmartArt probe B8: the group starts at the first
 * box, 762000,1770062, not at the frame's 762000,1524000), where the frame
 * left behind was an invisible shape the size of the whole graphic over
 * the slide.
 */
export function releaseSmartArtMembers<T>(
  members: readonly T[],
  options: { keepGroup: boolean },
): T[] {
  return members.filter((member) => !isSmartArtHiddenMember(member) && !isSmartArtFrameMember(member)).map((member) => {
    const record = member as unknown as Record<string, unknown>;
    const released: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(record)) {
      if (key.startsWith("smartArt")) continue;
      if (key === "groupId" && !options.keepGroup) continue;
      released[key] = value;
    }
    return released as unknown as T;
  });
}

/** A text body's margins inside its shape's text rectangle, in the host's own units. */
export interface SmartArtTextInShapeInsets {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * What Convert to Shapes (and Copy of one shape) makes of an item's words:
 * the text INSIDE the shape that draws them, as PowerPoint writes it -- one
 * shape per item whose `p:txBody` holds the words, its `a:bodyPr` insets the
 * distance from the shape's edges to the box the graphic drew them in
 * (SmartArt probe B8: Basic Block List's five shapes, each `lIns`..`bIns`
 * 247650, no text box beside them). MRAI draws a node and its label as two
 * members; this pairs them.
 *
 * A label is moved into the one node of its own item whose text rectangle
 * (`textArea`, for a preset that is narrower than the frame: an ellipse's
 * inscribed box) holds the label's box and that turns with it. Anything
 * else -- words standing beside their shape, words wider than the shape's
 * text rectangle, a node that already says something or is mirrored (a
 * flipped shape's words turn upside down in PowerPoint), a turned shape
 * whose words are off its centre, an item with two candidate shapes --
 * keeps its own text box, because no inset can put those
 * words where the graphic drew them. The merged shape takes the node's place
 * in the stacking order; the label's place is left out.
 */
export function smartArtWordsIntoTheirShapes<T>(
  members: readonly T[],
  host: {
    box: (member: T) => SmartArtMemberBox | null;
    /** The member's turn, in degrees (absent is 0). */
    turn: (member: T) => number;
    /** The node's text rectangle at `box`, in the same units. */
    textArea: (node: T, box: SmartArtMemberBox) => SmartArtMemberBox;
    /** Whether the node can take words: it says nothing of its own and is not mirrored. */
    takesWords: (node: T) => boolean;
    merge: (node: T, label: T, insets: SmartArtTextInShapeInsets) => T | null;
  },
  /** How far a label may stand outside the text rectangle and still count as inside it. */
  tolerance = 2,
): T[] {
  const record = (member: T) => member as unknown as Record<string, unknown>;
  const visible = (member: T) => !isSmartArtHiddenMember(member) && !isSmartArtFrameMember(member);
  const sameTurn = (a: T, b: T) => {
    const delta = ((host.turn(a) - host.turn(b)) % 360 + 360) % 360;
    return delta < 0.01 || delta > 359.99;
  };
  const merged = new Map<T, T>();
  const absorbed = new Set<T>();
  for (const label of members) {
    const own = record(label);
    if (own.smartArtRole !== "label" || typeof own.smartArtItemId !== "string" || !visible(label)) continue;
    const labelBox = host.box(label);
    if (!labelBox) continue;
    const candidates = members.filter((node) => {
      const theirs = record(node);
      if (
        theirs.smartArtRole !== "node" ||
        theirs.smartArtItemId !== own.smartArtItemId ||
        !visible(node) ||
        merged.has(node) ||
        !host.takesWords(node) ||
        !sameTurn(node, label)
      ) {
        return false;
      }
      const box = host.box(node);
      if (!box || box.width <= 0 || box.height <= 0) return false;
      const area = host.textArea(node, box);
      // A turned shape turns its words about ITS centre, a label about its
      // own: the same drawing only when the two centres meet.
      if (
        host.turn(node) % 360 !== 0 &&
        (Math.abs(labelBox.left + labelBox.width / 2 - (box.left + box.width / 2)) > tolerance ||
          Math.abs(labelBox.top + labelBox.height / 2 - (box.top + box.height / 2)) > tolerance)
      ) {
        return false;
      }
      return (
        labelBox.left >= area.left - tolerance &&
        labelBox.top >= area.top - tolerance &&
        labelBox.left + labelBox.width <= area.left + area.width + tolerance &&
        labelBox.top + labelBox.height <= area.top + area.height + tolerance
      );
    });
    if (candidates.length !== 1) continue;
    const node = candidates[0]!;
    const area = host.textArea(node, host.box(node)!);
    const inset = (value: number) => Math.max(0, Math.round(value));
    const result = host.merge(node, label, {
      left: inset(labelBox.left - area.left),
      top: inset(labelBox.top - area.top),
      right: inset(area.left + area.width - labelBox.left - labelBox.width),
      bottom: inset(area.top + area.height - labelBox.top - labelBox.height),
    });
    if (!result) continue;
    merged.set(node, result);
    absorbed.add(label);
  }
  if (!merged.size) return [...members];
  return members.flatMap((member) => (absorbed.has(member) ? [] : [merged.get(member) ?? member]));
}

/**
 * Reset > Convert > Convert to Text (PowerPoint only): the Text Pane's outline
 * as a bulleted list, one paragraph per item at the item's own level, which is
 * the reverse of Home > Paragraph > Convert to SmartArt.
 */
export function smartArtOutlineParagraphs(
  outline: readonly SmartArtOutlineItem[],
): { text: string; level: number; bullet: "bullet" }[] {
  return outline.map((item) => ({
    text: item.text,
    level: item.level,
    bullet: "bullet" as const,
  }));
}

/**
 * SmartArt Format > Shapes > Larger / Smaller: one step per press. Microsoft
 * documents the command ("click Larger to make the shape bigger", Microsoft
 * Support, Resize a SmartArt graphic shape) and says the amount depends on the
 * layout (Microsoft Learn, `SmartArtNode.Larger`) without naming it; MRAI's
 * step is ten per cent of the shape's current size.
 */
export const SMARTART_SHAPE_SIZE_STEP = 1.1;

/** A member rectangle in the host's own units. */
export interface SmartArtMemberBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

/**
 * Larger / Smaller on one outline item: its node member (and its own label,
 * where the host keeps one) grow or shrink about the node's centre, so the
 * shape stays where the author sees it. Everything else is untouched, which is
 * the host-independent part of Office's "the remaining shapes may adjust"
 * (MRAI does not re-flow the neighbours; Reset Graphic restores the layout's
 * own sizes).
 *
 * Box access goes through the host's `read` / `write` because the three page
 * models spell a member's rectangle differently (deck EMU fields, a Docs
 * member's child-space EMU, a sheet anchor). `null` when the item has no
 * node, a box is unreadable, or the step would take the shape below
 * `minEdge` — nothing partial is ever returned.
 */
export function resizeSmartArtItem<T>(
  members: readonly T[],
  itemId: string,
  /**
   * The scale: one number for Larger / Smaller, or each axis on its own for
   * SmartArt Format's Height and Width boxes (PPT-A010).
   */
  factor: number | { readonly x: number; readonly y: number },
  read: (member: T) => SmartArtMemberBox | null,
  write: (member: T, box: SmartArtMemberBox) => T,
  minEdge = 1,
  /**
   * The graphic's frame. Office grows a shape INSIDE its graphic, so when one
   * is given the grown item is moved back in from an edge it would cross, and
   * an item that would outgrow the frame is refused. A host whose frame is the
   * members' own union passes it, or every Larger would silently widen the
   * graphic and Reset would re-lay it out in the wider box.
   */
  frame?: SmartArtMemberBox,
  /**
   * The diagram's layout, for a host whose members do not state it (a Docs
   * group keeps it on the group); otherwise read off the members.
   */
  kind?: string,
): T[] | null {
  let fx = typeof factor === "number" ? factor : factor.x;
  let fy = typeof factor === "number" ? factor : factor.y;
  if (!Number.isFinite(fx) || fx <= 0 || !Number.isFinite(fy) || fy <= 0) return null;
  const roleOf = (member: T) =>
    (member as unknown as Record<string, unknown>).smartArtRole;
  const itemOf = (member: T) =>
    (member as unknown as Record<string, unknown>).smartArtItemId;
  const node = members.find(
    (member) => roleOf(member) === "node" && itemOf(member) === itemId,
  );
  const nodeBox = node ? read(node) : null;
  if (!nodeBox) return null;
  const centerX = nodeBox.left + nodeBox.width / 2;
  const centerY = nodeBox.top + nodeBox.height / 2;
  const scaled = (box: SmartArtMemberBox): SmartArtMemberBox => ({
    left: centerX + (box.left - centerX) * fx,
    top: centerY + (box.top - centerY) * fy,
    width: box.width * fx,
    height: box.height * fy,
  });
  const item = members.filter((member) => {
    const role = roleOf(member);
    return itemOf(member) === itemId && (role === "node" || role === "label");
  });
  const boxes = item.map((member) => read(member));
  if (boxes.some((box) => box === null)) return null;
  // Larger / Smaller is ONE factor, both ways: a shape keeps its proportions.
  // Near the frame it grows by the largest factor that still fits on both
  // axes, and not at all when it has no room. Capping each axis on its own
  // let a node as tall as its row keep growing wide, and the Smaller after it
  // shrank both ways: two Larger and two Smaller turned a 100x100 node into
  // 100x82.6 (video audit A03).
  //
  // A shape already at the frame on one axis grows along the other (the
  // per-axis cap below): refusing it left Larger enabled and dead on every
  // Basic Process node, each as tall as its graphic (98edf085c7). Interim,
  // until the layout re-flows the other shapes as Office does (Larger
  // enlarges the selected shape and the layout adjusts the rest). Smaller
  // after such a step shrinks both axes, so that pair does not round-trip.
  if (typeof factor === "number" && frame && fx > 1) {
    let room = fx;
    for (const box of boxes as SmartArtMemberBox[]) {
      if (box.width > 0) room = Math.min(room, frame.width / box.width);
      if (box.height > 0) room = Math.min(room, frame.height / box.height);
    }
    if (room > 1 + 1e-6) {
      fx = room;
      fy = room;
    }
  }
  // A graphic's frame caps how far a shape can grow along each axis: a box
  // already as tall as its row keeps growing wide, and a shape that fills the
  // frame both ways cannot grow at all (Office: "depending on ... the amount of
  // available space").
  const grown = (boxes as SmartArtMemberBox[]).map((box) => {
    const next = scaled(box);
    if (!frame) return next;
    const width = Math.min(next.width, frame.width);
    const height = Math.min(next.height, frame.height);
    return {
      left: next.left + (next.width - width) / 2,
      top: next.top + (next.height - height) / 2,
      width,
      height,
    };
  });
  if (grown.some((box) => box.width < minEdge || box.height < minEdge)) {
    return null;
  }
  const unchanged = (boxes as SmartArtMemberBox[]).every(
    (box, index) =>
      Math.abs(box.width - grown[index].width) < 1e-6 &&
      Math.abs(box.height - grown[index].height) < 1e-6,
  );
  if (unchanged) return null;
  // Keep the item inside the graphic: one shift for the node and its label,
  // so they stay together.
  let dx = 0;
  let dy = 0;
  if (frame) {
    const left = Math.min(...grown.map((box) => box.left));
    const top = Math.min(...grown.map((box) => box.top));
    const right = Math.max(...grown.map((box) => box.left + box.width));
    const bottom = Math.max(...grown.map((box) => box.top + box.height));
    if (right - left > frame.width + 1e-6 || bottom - top > frame.height + 1e-6) {
      return null;
    }
    dx =
      left < frame.left
        ? frame.left - left
        : right > frame.left + frame.width
          ? frame.left + frame.width - right
          : 0;
    dy =
      top < frame.top
        ? frame.top - top
        : bottom > frame.top + frame.height
          ? frame.top + frame.height - bottom
          : 0;
  }
  const byMember = new Map(
    item.map((member, index) => [
      member,
      { ...grown[index], left: grown[index].left + dx, top: grown[index].top + dy },
    ]),
  );
  // The lines that reach the node follow its edges (PPT-B167: an organization
  // chart's drop stopped at the old top, a gap above the smaller shape).
  const newNode = node ? byMember.get(node) : undefined;
  const connectors = new Map<T, SmartArtMemberBox>();
  if (newNode) {
    for (const member of members) {
      if (roleOf(member) !== "connector") continue;
      const line = read(member);
      const followed = line ? smartArtLineFollowingNode(line, nodeBox, newNode) : null;
      if (followed && followed.width >= minEdge && followed.height >= minEdge) {
        connectors.set(member, followed);
      }
    }
  }
  // Office keeps a Larger / Smaller only as the shape's `custScaleX/Y` on its
  // presentation point, so the grown shape carries the factor against its
  // layout box (`./custom-geometry`); the writers put it where Office reads it.
  const grownMembers = members.map((member) => {
    const box = byMember.get(member);
    if (!box) return member;
    const written = write(member, box);
    return smartArtMemberAfterBoxEdit(member, written, read);
  });
  // In a layout whose transitions Office routes (`./connector-follow`) the
  // arrows and lines that join the shape are re-laid onto it. The grown shape
  // keeps its centre, so its connectors keep their direction; one that would
  // have to turn is left to the line rule below. A member's place in the list
  // is its identity across this one edit (an untouched member is the same
  // object in both lists).
  const position = new Map<T, number>();
  members.forEach((member, at) => position.set(member, at));
  grownMembers.forEach((member, at) => position.set(member, at));
  const layout =
    kind ??
    ((members.find((member) => typeof (member as Record<string, unknown>).smartArtKind === "string") as
      | Record<string, unknown>
      | undefined)?.smartArtKind as string | undefined);
  const followed = smartArtConnectorsFollowNodes(layout, members, grownMembers, {
    id: (member) => String(position.get(member) ?? -1),
    box: read,
    rotation: smartArtMemberRotation,
    shape: smartArtMemberShape,
    write: (member, box, rotation) =>
      Math.abs(rotation - smartArtMemberRotation(member)) < 0.01 ? write(member, box) : member,
  });
  return followed.map((member, at) => {
    if (member !== grownMembers[at]) return member;
    const line = connectors.get(members[at]);
    return line ? write(members[at], line) : member;
  });
}

/** A member's turn in clockwise degrees: a block's, a Docs member's or a sheet anchor's. */
function smartArtMemberRotation(member: unknown): number {
  const record = member as Record<string, unknown>;
  const anchor = record.anchor as Record<string, unknown> | undefined;
  const turn = typeof record.rotation === "number" ? record.rotation : anchor?.rotation;
  return typeof turn === "number" && Number.isFinite(turn) ? turn : 0;
}

/** A member's preset shape: a block's `shape`, a sheet shape's `kind`, a Docs `preset`. */
function smartArtMemberShape(member: unknown): unknown {
  const record = member as Record<string, unknown>;
  const shape = record.shape;
  if (typeof shape === "string") return shape;
  if (shape && typeof shape === "object") return (shape as Record<string, unknown>).kind;
  return record.preset;
}

/**
 * A straight connector segment after the node one of its ends touches was
 * resized: a vertical segment ending on the node's top or bottom edge, or a
 * horizontal one on its left or right edge, is stretched or shortened to the
 * node's new edge. Null for a segment that does not end on the node.
 */
export function smartArtLineFollowingNode(
  line: SmartArtMemberBox,
  before: SmartArtMemberBox,
  after: SmartArtMemberBox,
): SmartArtMemberBox | null {
  const thickness = Math.min(line.width, line.height);
  const near = (a: number, b: number) => Math.abs(a - b) <= Math.max(1, thickness * 2);
  const vertical = line.height > line.width;
  if (vertical) {
    const x = line.left + line.width / 2;
    if (x < before.left || x > before.left + before.width) return null;
    const lineBottom = line.top + line.height;
    if (near(lineBottom, before.top)) {
      return { ...line, height: after.top - line.top };
    }
    if (near(line.top, before.top + before.height)) {
      const top = after.top + after.height;
      return { ...line, top, height: lineBottom - top };
    }
    return null;
  }
  const y = line.top + line.height / 2;
  if (y < before.top || y > before.top + before.height) return null;
  const lineRight = line.left + line.width;
  if (near(lineRight, before.left)) {
    return { ...line, width: after.left - line.left };
  }
  if (near(line.left, before.left + before.width)) {
    const left = after.left + after.width;
    return { ...line, left, width: lineRight - left };
  }
  return null;
}

/** Delete: remove one item, selecting the one that takes its place. */
export function smartArtOutlineAfterDelete(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
): SmartArtOutlineChange | null {
  const removedIndex = items.findIndex((item) => item.id === itemId);
  const next = removeSmartArtOutlineItem(items, itemId);
  if (!next || removedIndex < 0) return null;
  return {
    items: next,
    selectedItemId: next[Math.min(removedIndex, next.length - 1)].id,
  };
}

/**
 * The Text Pane's Delete at a line's end / Backspace at the next line's start:
 * `keepId` takes the joined `text` and the line after it goes. The caret
 * stays in `keepId` at the join.
 */
export function smartArtOutlineAfterMerge(
  items: readonly SmartArtOutlineItem[],
  keepId: string,
  goneId: string,
  text: string,
): SmartArtOutlineChange | null {
  const next = mergeSmartArtOutlineItems(
    items,
    keepId,
    goneId,
    text.slice(0, MAX_SMARTART_OUTLINE_TEXT_LENGTH),
  );
  return next ? { items: next, selectedItemId: keepId } : null;
}

/**
 * A paste of several lines into a Text Pane row: each line is a line of the
 * pane, as Enter would make it (Office's Text Pane is an outline, one
 * paragraph per shape). `itemId` keeps `text` (its words before the caret and
 * the first pasted line); every later line becomes a new item right after it,
 * at its level, the last one carrying the row's words after the caret.
 * `null` when the outline cannot hold that many items.
 */
export function smartArtOutlineAfterPaste(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
  text: string,
  lines: readonly string[],
  newId: () => string = () => randomUUID(),
): SmartArtOutlineChange | null {
  let next = updateSmartArtOutlineItemText(
    items,
    itemId,
    text.slice(0, MAX_SMARTART_OUTLINE_TEXT_LENGTH),
  );
  let anchor = itemId;
  for (const line of lines) {
    if (!next) return null;
    const id = newId();
    next = addSmartArtOutlineItem(next, anchor, "after", {
      id,
      text: line.slice(0, MAX_SMARTART_OUTLINE_TEXT_LENGTH),
      level: 0,
    });
    anchor = id;
  }
  return next ? { items: next, selectedItemId: anchor } : null;
}

/**
 * SmartArt Design > Create Graphic > Layout on an organization chart: how the
 * selected box's subordinates hang below it (Microsoft Support, "Create an
 * organization chart": "the layout affects all boxes below the selected
 * box"). Offered on the organization chart layouts only, on a box that is not
 * an assistant (MRAI's assistants have no subordinates).
 */
export function smartArtOutlineAfterHierBranch(
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  branch: SmartArtHierBranch,
  kind: SmartArtKind,
): SmartArtOutlineChange | null {
  if (!smartArtLayoutTakesHierBranch(kind)) return null;
  const target = items.find((item) => item.id === selectedItemId);
  if (!target || target.assistant) return null;
  return {
    items: items.map((item) => (item.id === selectedItemId ? { ...item, hierBranch: branch } : item)),
    selectedItemId,
  };
}

/** Promote: only layouts that actually nest accept it. */
export function smartArtOutlineAfterPromote(
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  kind: SmartArtKind,
  pending?: SmartArtPendingOutlineText,
): SmartArtOutlineChange | null {
  if (!smartArtAcceptsLevelTwo(kind)) return null;
  const current = withPendingText(items, pending);
  if (!current) return null;
  const next = promoteSmartArtOutlineItem(current, selectedItemId);
  if (!next) return null;
  return { items: next, selectedItemId };
}

/** Demote: the mirror of {@link smartArtOutlineAfterPromote}. */
export function smartArtOutlineAfterDemote(
  items: readonly SmartArtOutlineItem[],
  selectedItemId: string,
  kind: SmartArtKind,
  pending?: SmartArtPendingOutlineText,
): SmartArtOutlineChange | null {
  if (!smartArtAcceptsLevelTwo(kind)) return null;
  const current = withPendingText(items, pending);
  if (!current) return null;
  const next = demoteSmartArtOutlineItem(current, selectedItemId);
  if (!next) return null;
  return { items: next, selectedItemId };
}

/** Edit one item's text, capped at the outline's own limit. */
export function smartArtOutlineAfterTextEdit(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
  text: string,
): SmartArtOutlineChange | null {
  const nextText = text.slice(0, MAX_SMARTART_OUTLINE_TEXT_LENGTH);
  const current = items.find((item) => item.id === itemId);
  if (!current || current.text === nextText) return null;
  const next = updateSmartArtOutlineItemText(items, itemId, nextText);
  return next ? { items: next, selectedItemId: itemId } : null;
}

/** The colour variation a diagram's members wear (`smartArtColors`), or null. */
export function readSmartArtGroupColorsId(
  members: readonly unknown[],
): string | null {
  for (const member of members) {
    const value = (member as Record<string, unknown> | null)?.smartArtColors;
    if (typeof value === "string" && value) return value;
  }
  return null;
}

/**
 * The Office ALIAS a diagram's members record (`smartArtLayoutUrn`: the layout
 * it is drawn from when that is not its catalog layout's own), or null. Read
 * from a member stating its role in the diagram, which every member does.
 */
export function readSmartArtGroupLayoutUrn(
  members: readonly unknown[],
): string | null {
  for (const member of members) {
    const value = (member as Record<string, unknown> | null)?.smartArtLayoutUrn;
    if (typeof value === "string" && value) return value;
  }
  return null;
}

/**
 * The colour set a diagram wears, resolved on the host's theme: the colour
 * variation it records (Office's default `accent1_2` when none) over the
 * theme's accents, the editor's default without either. What the Styles
 * gallery draws its tiles in (PowerPoint's tiles follow the diagram's
 * colours). Pure, so a window can ask without the engine (`design-ops`
 * carries the engine and must stay out of the client graph).
 */
export function smartArtDiagramColors(
  members: readonly OfficeSmartArtBlockLike[],
  theme?: Partial<SmartArtColorTheme> | null,
): SmartArtStyle {
  const colorsId = readSmartArtGroupColorsId(members);
  const variant = smartArtColorVariant(colorsId ?? "accent1_2", { all: true });
  if (!variant || (!theme && !colorsId)) return DEFAULT_SMARTART_STYLE;
  return resolveSmartArtColorVariant(variant, smartArtDiagramItemCount(members), theme).colors;
}
