/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * How a SmartArt graphic is selected, moved, sized and typed into: the
 * decisions the three editors share (Presentation, Docs, the worksheet), each
 * host keeping only how it draws and stores them.
 *
 * Office's behaviour (Microsoft Support, "Create a SmartArt graphic from
 * scratch", "Add or delete shapes in a SmartArt graphic"; PowerPoint, Word and
 * Excel agree):
 * - The graphic has ONE frame, the graphic frame. Its border is drawn there --
 *   not round the shapes, which need not fill it (a Target in a wide frame) --
 *   and its border, its handles and the Text Pane control on its left edge are
 *   what a press on the border reaches; a press inside the frame on no shape
 *   selects the graphic.
 * - A double click (or F2, or a typed key) on a shape types into the words of
 *   that shape's item. A connecting shape (an arrow, a ring with no words of
 *   its own) has no line in the Text Pane and opens nothing; a picture
 *   placeholder opens the picture, not words.
 * - Add Shape, Promote, Demote, Move Up / Down and Delete act on "the existing
 *   shape" selected -- the Text Pane's row when the caret is in the pane. With
 *   only the graphic selected, Add Shape adds at the end; the others need an
 *   item and do nothing.
 *
 * Pure and framework-free: the browser and the engine host may both import it.
 */

import {
  SMARTART_NO_OUTLINE_COMMANDS,
  smartArtOutlineCommandState,
  type SmartArtOutlineCommandState,
} from "./design-model";
import { readSmartArtBlockMarker, type SmartArtKind } from "./model";
import type { SmartArtOutlineItem } from "./outline";

/** A box in a host's own units (px on a slide or a sheet, EMU in a document). */
export interface SmartArtBox {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

/** The selection frame's padding outside the graphic frame, in screen px. */
export const SMARTART_FRAME_PAD_PX = 10;

export const SMARTART_FRAME_EDGES = ["top", "right", "bottom", "left"] as const;
export type SmartArtFrameEdge = (typeof SMARTART_FRAME_EDGES)[number];

/**
 * The graphic frame every part of the selection shares -- the border, the
 * press area, a frame drag's live box and the next relayout: the frame the
 * host stores for the graphic when it has one, else the union of its members.
 */
export function smartArtGraphicFrameBox(
  stored: SmartArtBox | null | undefined,
  members: readonly SmartArtBox[],
): SmartArtBox | null {
  if (stored && stored.width > 0 && stored.height > 0) return stored;
  const boxes = members.filter(
    (box) => [box.left, box.top, box.width, box.height].every(Number.isFinite),
  );
  if (!boxes.length) return null;
  const left = Math.min(...boxes.map((box) => box.left));
  const top = Math.min(...boxes.map((box) => box.top));
  const right = Math.max(...boxes.map((box) => box.left + box.width));
  const bottom = Math.max(...boxes.map((box) => box.top + box.height));
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Where the parts of the selection stand around a frame drawn `pad` px
 * outside the graphic frame, in px relative to that padded box, and how they
 * stack. The border strips straddle the padded box's edge and lie wholly
 * outside the graphic's press area (which reaches half a pad out), so a press
 * on the border is always the border's; the grips and the Text Pane control
 * sit over the strips.
 */
export function smartArtFrameChrome(pad: number = SMARTART_FRAME_PAD_PX) {
  const half = Math.floor(pad / 2);
  const edge = (side: SmartArtFrameEdge): Record<string, number> =>
    side === "top" || side === "bottom"
      ? { left: 0, right: 0, height: pad, [side]: -half }
      : { top: 0, bottom: 0, width: pad, [side]: -half };
  const edges: Record<SmartArtFrameEdge, Record<string, number>> = {
    top: edge("top"),
    right: edge("right"),
    bottom: edge("bottom"),
    left: edge("left"),
  };
  return {
    /** The graphic's own press area, outside the graphic frame by this much. */
    pressOutset: half,
    edges,
  };
}

/** The graphic's press area: the graphic frame, half a pad out on every side. */
export function smartArtGraphicPressBox(
  frame: SmartArtBox,
  pad: number = SMARTART_FRAME_PAD_PX,
): SmartArtBox {
  const out = smartArtFrameChrome(pad).pressOutset;
  return {
    left: frame.left - out,
    top: frame.top - out,
    width: frame.width + out * 2,
    height: frame.height + out * 2,
  };
}

/** A member as every host's identity fields spell it. */
export interface SmartArtMemberIdentity {
  readonly smartArtRole?: unknown;
  readonly smartArtItemId?: unknown;
  readonly smartArtPicturePlaceholder?: unknown;
}

/**
 * The outline item whose words a double click, F2 or a typed key on this
 * member opens, or `null` when it opens none (a connecting shape, a picture
 * placeholder, a member of no item).
 */
export function smartArtMemberTextItem(member: SmartArtMemberIdentity): string | null {
  const role = member.smartArtRole;
  if (role !== "node" && role !== "label") return null;
  if (member.smartArtPicturePlaceholder === true) return null;
  const itemId = member.smartArtItemId;
  return typeof itemId === "string" && itemId ? itemId : null;
}

/** The Design-tab and Text Pane commands that act on one item. */
export type SmartArtItemCommand =
  | "add"
  | "promote"
  | "demote"
  | "move-up"
  | "move-down"
  | "delete";

/**
 * The item `command` acts on: the Text Pane's row, else the selected shape's
 * item; with neither, Add Shape adds after the last item and every other
 * command has nothing to act on (`null`).
 */
export function smartArtCommandItemId(input: {
  readonly outline: readonly { readonly id: string }[];
  readonly command: SmartArtItemCommand;
  /** The row the Text Pane's caret is in. */
  readonly paneItemId?: string | null;
  /** The item of the shape selected on the canvas. */
  readonly shapeItemId?: string | null;
}): string | null {
  const { outline } = input;
  for (const candidate of [input.paneItemId, input.shapeItemId]) {
    if (candidate && outline.some((item) => item.id === candidate)) return candidate;
  }
  return input.command === "add" ? (outline[outline.length - 1]?.id ?? null) : null;
}

/**
 * Which item commands are available: every command on the item the pane's
 * row or the selected shape names; with neither, only Add Shape (after the
 * last item), as Office offers with only the graphic selected.
 */
export function smartArtItemCommandState(input: {
  readonly outline: readonly SmartArtOutlineItem[] | null | undefined;
  readonly kind: SmartArtKind;
  readonly paneItemId?: string | null;
  readonly shapeItemId?: string | null;
}): SmartArtOutlineCommandState {
  const outline = input.outline;
  if (!outline?.length) return SMARTART_NO_OUTLINE_COMMANDS;
  const named = smartArtCommandItemId({ ...input, outline, command: "promote" });
  if (named) return smartArtOutlineCommandState(outline, named, input.kind);
  const last = smartArtCommandItemId({ ...input, outline, command: "add" });
  return {
    ...SMARTART_NO_OUTLINE_COMMANDS,
    addAfter: smartArtOutlineCommandState(outline, last, input.kind).addAfter,
  };
}

/**
 * The SmartArt shape Tab (offset 1) or Shift+Tab (-1) moves to inside an
 * activated graphic (Microsoft Support, "Keyboard shortcuts for SmartArt
 * graphics": after Enter activates a graphic, Tab and Shift+Tab select its
 * next and previous elements; Esc returns to the top level): the item shapes
 * in the outline's order, wrapping round, an item the layout does not draw
 * left out. A label stands for its shape, a connecting shape for the item
 * before it. Null when `from` is no member of an authored graphic, or the
 * graphic has no item shape. Offset 0 is Enter on the whole graphic, which
 * activates it on its first shape. Generic over the host's member type: a
 * deck block, a worksheet object.
 */
export function smartArtSteppedMember<T extends { readonly id: string; readonly type?: unknown }>(
  members: readonly T[],
  from: T | null | undefined,
  offset: 1 | -1 | 0,
): T | null {
  const marker = from ? readSmartArtBlockMarker(from) : null;
  if (!from || !marker) return null;
  const shapes = members
    .map((member, order) => ({ member, order, marker: readSmartArtBlockMarker(member) }))
    .filter(
      (entry) =>
        entry.member.type === "shape" &&
        entry.marker?.role === "node" &&
        entry.marker.groupId === marker.groupId &&
        (entry.member as { smartArtHidden?: unknown }).smartArtHidden !== true,
    )
    .sort(
      (a, b) =>
        (a.marker!.itemIndex ?? Number.MAX_SAFE_INTEGER) -
          (b.marker!.itemIndex ?? Number.MAX_SAFE_INTEGER) || a.order - b.order,
    );
  if (!shapes.length) return null;
  if (offset === 0) return shapes[0].member;
  const at = shapes.findIndex(
    (entry) =>
      entry.member.id === from.id ||
      (marker.role === "label" && marker.itemId !== null && entry.marker!.itemId === marker.itemId),
  );
  if (at < 0) {
    // A connecting shape: the next item after it, or the one it follows.
    const before = shapes.filter((entry) => (entry.marker!.itemIndex ?? -1) <= (marker.itemIndex ?? -1));
    const anchor = before.length ? shapes.indexOf(before[before.length - 1]) : offset > 0 ? -1 : 0;
    return shapes[(anchor + (offset > 0 ? 1 : 0) + shapes.length) % shapes.length].member;
  }
  return shapes[(at + offset + shapes.length) % shapes.length].member;
}
