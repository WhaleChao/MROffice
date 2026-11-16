/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SmartArt Format on a slide (PPT-A010): the shapes of ONE diagram the
 * selection names, and the commands PowerPoint's SmartArt Format tab runs on
 * them -- Shape Fill, Shape Outline, Change Shape and Text Fill.
 *
 * PowerPoint applies them to the SHAPES the author selected inside the
 * diagram, or to every shape when the graphic itself is selected (Microsoft
 * Support, "Change the colors of a shape in a SmartArt graphic": select the
 * shape, then Shape Fill). A slide diagram draws an item as two blocks -- its
 * node and its label -- so either one selected names the item.
 *
 * Every write goes through `shapeBlockWithPatch` (paint, shape) or the label
 * text adapter (text colour), which state the authority Office keeps on the
 * shape's own points (`smartArtAuthorPaint`, `smartArtAuthorShape`,
 * `smartArtAuthorText`), so a rebuild keeps it, the export writes it and Reset
 * Shape / Reset Graphic discard it.
 */

import type {
  PresentationBlock,
  PresentationImageBlock,
  PresentationShapeBlock,
  PresentationTextBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { readSmartArtBlockMarker, SMARTART_PICTURE_PLACEHOLDER_SRC } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { pictureFillCropRect, pictureFitCropRect } from "@mrnote/core/web/lib/presentation-crop-group.ts";
import { resetImagePicture } from "@mrnote/core/web/lib/presentation-conveniences.ts";
import { applyNativePictureStyle, canApplyNativePictureStyle } from "@mrnote/core/web/lib/presentation-picture-style-native.ts";
import {
  captureFormatSnapshot,
  isEmptyPlan,
  planFormatPaint,
  type FormatPainterSnapshot,
} from "@mrnote/core/web/lib/presentation-format-painter.ts";
import { paintPresentationBlockFormat } from "./presentation-format-painter-apply";
import { shapeBlockWithPatch } from "@mrnote/core/web/lib/presentation-shape-buttons.ts";
import { PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER } from "@mrnote/core/web/lib/presentation-smartart-text.ts";
import type { ShapeKind } from "@mrnote/core/web/lib/presentation-shapes.ts";
import type { OfficeShapeEffect } from "@mrnote/core/web/components/office-ribbon/OfficeShapeFormatTab.tsx";
import {
  normalizeOfficeShapeEffectPreset,
  officeShapeEffectMenuValue,
} from "@mrnote/core/web/lib/office-shape-effects.ts";
import { presentationPictureSizeEmuFromCmDraft } from "@mrnote/core/web/lib/presentation-command-adapters.ts";
import { presentationShapeEffectPatch, presentationShapeEffects } from "./presentation-shape-effects";
import { smartArtEffectsOwn3d, smartArtNodeWearsStyleBevel } from "@mrnote/core/web/lib/office-smartart/model.ts";
import { withSmartArtLabelCameras } from "@mrnote/core/web/lib/office-smartart/member-camera.ts";
import type {
  OfficePictureEffectKind,
  OfficePictureEffects,
  OfficePictureEffectValues,
} from "@mrnote/core/web/lib/office-picture-effects.ts";
import type { PresentationSmartArtFormatState } from "../components/notebook/contents/presentation/PresentationSmartArtFormatTab";
import type { OfficeShapeStyleRecipe } from "@mrnote/core/web/lib/office-shape-style-catalog.ts";
import { sameOfficeTextPaint, type OfficeTextPaint } from "@mrnote/core/web/lib/office-text-paint.ts";
import type { PresentationWordArtStyle } from "@mrnote/core/web/lib/presentation-wordart-gallery.ts";
import {
  presentationShapeMatchesStyle,
  presentationShapeStylePatch,
} from "@mrnote/core/web/lib/presentation-shape-style-adapter.ts";

/** The items of one diagram a SmartArt Format command acts on. */
export interface PresentationSmartArtFormatTarget {
  readonly groupId: string;
  readonly itemIds: readonly string[];
}

/**
 * The diagram the selection is inside and the items it names: every item when
 * every member is selected (the graphic), else the items whose node or label
 * is selected. Null when the selection is not inside exactly one diagram, or
 * names no item (a connector alone).
 */
export function presentationSmartArtFormatTarget(
  blocks: readonly PresentationBlock[],
  selectedIds: Iterable<string>,
  /**
   * The member the author pressed. A click on a slide diagram selects every
   * member (the frame moves them together) with the pressed one primary, so
   * the pressed shape -- not the whole graphic -- is what Format acts on.
   */
  primaryId?: string | null,
): PresentationSmartArtFormatTarget | null {
  const wanted = new Set(selectedIds);
  const selected = blocks.filter((block) => wanted.has(block.id));
  if (selected.length === 0 || selected.length !== wanted.size) return null;
  const markers = selected.map((block) => readSmartArtBlockMarker(block));
  const groupId = markers[0]?.groupId;
  if (!groupId || markers.some((marker) => marker?.groupId !== groupId)) return null;
  const members = blocks.filter((block) => readSmartArtBlockMarker(block)?.groupId === groupId);
  const whole = members.every((block) => wanted.has(block.id));
  const pressed = whole && primaryId
    ? readSmartArtBlockMarker(blocks.find((block) => block.id === primaryId))
    : null;
  if (pressed?.groupId === groupId && pressed.role !== "connector" && pressed.itemId) {
    return { groupId, itemIds: [pressed.itemId] };
  }
  const itemIds = [
    ...new Set(
      (whole ? members : selected).flatMap((block) => {
        const marker = readSmartArtBlockMarker(block);
        return marker && marker.role !== "connector" && marker.itemId ? [marker.itemId] : [];
      }),
    ),
  ];
  return itemIds.length ? { groupId, itemIds } : null;
}

function targetMembers<T extends PresentationBlock>(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  role: "node" | "label",
  type: T["type"],
): T[] {
  const items = new Set(target.itemIds);
  return blocks.filter((block): block is T => {
    const marker = readSmartArtBlockMarker(block);
    return (
      block.type === type &&
      marker?.groupId === target.groupId &&
      marker.role === role &&
      marker.itemId !== null &&
      items.has(marker.itemId)
    );
  });
}

/**
 * The picture in the ONE targeted item's slot (a picture layout's slot block
 * holding a picture), which Picture Format acts on: PowerPoint shows the
 * Picture Format tab for a picture placeholder that holds a picture. Null for
 * any other target, an empty slot, or a graphic-wide target.
 */
export function presentationSmartArtSlotPicture(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
): PresentationImageBlock | null {
  if (target.itemIds.length !== 1) return null;
  const slots = targetMembers<PresentationImageBlock>(blocks, target, "node", "image").filter(
    (block) => typeof block.src === "string" && block.src !== "" && block.src !== SMARTART_PICTURE_PLACEHOLDER_SRC,
  );
  return slots.length === 1 ? slots[0] : null;
}

/** One Picture Format command on a slot's picture (`lib/office-smartart/slot-picture-format`). */
export type PresentationSmartArtSlotPictureCommand =
  | { readonly kind: "crop"; readonly mode: "fill" | "fit"; readonly natural?: { width: number; height: number } | null }
  | { readonly kind: "style"; readonly style: string }
  | { readonly kind: "reset" };

/**
 * The slide's blocks with Picture Format applied to the targeted slot's
 * picture, through the deck's own picture machinery: Crop > Fill / Fit is
 * the deck's crop math on the slot's frame (its `cropRect`, `a:srcRect`), a
 * Picture Style is `applyNativePictureStyle` (`pictureStylePreset`), Reset
 * Picture is the deck's `resetImagePicture`. Null when nothing changes or the
 * slot may not be written.
 */
export function presentationSmartArtBlocksWithSlotPicture(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  command: PresentationSmartArtSlotPictureCommand,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const slot = presentationSmartArtSlotPicture(blocks, target);
  if (!slot || !canManipulate(slot)) return null;
  let next: PresentationImageBlock = slot;
  if (command.kind === "reset") {
    next = resetImagePicture(slot);
  } else if (command.kind === "style") {
    if (!canApplyNativePictureStyle(slot)) return null;
    next = applyNativePictureStyle(slot, command.style);
  } else {
    const own = slot as unknown as { pixelWidth?: unknown; pixelHeight?: unknown };
    const natural = command.natural ?? (
      Number(own.pixelWidth) > 0 && Number(own.pixelHeight) > 0
        ? { width: Number(own.pixelWidth), height: Number(own.pixelHeight) }
        : null
    );
    if (!natural) return null;
    const crop =
      command.mode === "fill"
        ? pictureFillCropRect(natural.width, natural.height, slot.width, slot.height)
        : pictureFitCropRect(natural.width, natural.height, slot.width, slot.height);
    if (!crop) return null;
    next = { ...slot, cropRect: crop } as PresentationImageBlock;
  }
  if (JSON.stringify(next) === JSON.stringify(slot)) return null;
  return blocks.map((block) => (block === slot ? next : block));
}

/**
 * Format Painter with ONE SmartArt shape as the source (Microsoft Support,
 * "Use the Format Painter": it copies an object's formatting -- for a shape
 * its fill, line and effects -- and the formatting of its words). A slide
 * diagram draws an item as its node and its label, so the shape formatting
 * is the node's and the text formatting the label's; never the item's
 * identity or its geometry. Null for anything else -- the whole graphic (the
 * graphic is one object whose formatting is its own frame's, which the
 * painter does not carry), several items, or nothing to copy.
 */
export function presentationSmartArtPainterSource(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
): { snapshot: FormatPainterSnapshot; sourceId: string } | null {
  if (target.itemIds.length !== 1) return null;
  const node = presentationSmartArtTargetNodes(blocks, target)[0] ?? null;
  const label = presentationSmartArtTargetLabels(blocks, target)[0] ?? null;
  const shape = node ? captureFormatSnapshot(node) : null;
  const words = label ? captureFormatSnapshot(label) : null;
  const snapshot: FormatPainterSnapshot = {
    ...(shape?.shape ? { shape: shape.shape } : {}),
    ...(words?.text ? { text: words.text } : shape?.text ? { text: shape.text } : {}),
    ...(words?.runStyle ? { runStyle: words.runStyle } : shape?.runStyle ? { runStyle: shape.runStyle } : {}),
  };
  if (!snapshot.shape && !snapshot.text && !snapshot.runStyle) return null;
  const sourceId = node?.id ?? label?.id;
  return sourceId ? { snapshot, sourceId } : null;
}

/**
 * Format Painter applied onto ONE SmartArt shape: the clicked member names
 * its item (its node or its label), the shape formatting goes onto the node
 * through `shapeBlockWithPatch` (the shape's paint and effects are the
 * author's from then on, `smartArtAuthorPaint` / `smartArtAuthorFormat`) and
 * the text formatting onto the label, which states what it changed is the
 * author's (`smartArtAuthorText`, `smartArtAuthorTextSize` for a size) as
 * Home's commands on SmartArt words do -- so a relayout keeps the paint, the
 * export writes it and Reset Shape discards it. Null when nothing applies or
 * a member may not be written.
 */
export function presentationSmartArtBlocksWithPaintedFormat(
  blocks: readonly PresentationBlock[],
  clickedId: string,
  snapshot: FormatPainterSnapshot,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const clicked = blocks.find((block) => block.id === clickedId);
  const marker = readSmartArtBlockMarker(clicked);
  if (!marker || marker.role === "connector" || !marker.itemId) return null;
  const target = { groupId: marker.groupId, itemIds: [marker.itemId] };
  const nodes = presentationSmartArtTargetNodes(blocks, target);
  const labels = presentationSmartArtTargetLabels(blocks, target);
  const members = [...nodes, ...labels];
  if (members.length === 0 || members.some((member) => !canManipulate(member))) return null;
  const nodeIds = new Set(nodes.map((node) => node.id));
  const labelIds = new Set(labels.map((label) => label.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (nodeIds.has(block.id)) {
      const plan = planFormatPaint({ ...(snapshot.shape ? { shape: snapshot.shape } : {}) }, block);
      if (isEmptyPlan(plan)) return block;
      const painted = paintPresentationBlockFormat(block, { shapePatch: plan.shapePatch });
      if (painted !== block) changed = true;
      return painted;
    }
    if (labelIds.has(block.id)) {
      const plan = planFormatPaint(
        {
          ...(snapshot.text ? { text: snapshot.text } : {}),
          ...(snapshot.runStyle ? { runStyle: snapshot.runStyle } : {}),
        },
        block,
      );
      if (isEmptyPlan(plan)) return block;
      const painted = paintPresentationBlockFormat(block, plan) as PresentationBlock & Record<string, unknown>;
      if (JSON.stringify(painted) === JSON.stringify(block)) return block;
      changed = true;
      const sized =
        plan.textPatch?.fontSize !== undefined ||
        (plan.runStylePatch as { fontSize?: unknown } | undefined)?.fontSize !== undefined;
      const coloured =
        plan.textPatch?.color !== undefined ||
        (plan.runStylePatch as { color?: unknown } | undefined)?.color !== undefined;
      return {
        ...painted,
        ...(sized || coloured ? { smartArtAuthorText: true } : {}),
        ...(sized ? { smartArtAuthorTextSize: true } : {}),
      } as PresentationBlock;
    }
    return block;
  });
  return changed ? next : null;
}

/** The node shapes of the targeted items. */
export function presentationSmartArtTargetNodes(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
): PresentationShapeBlock[] {
  return targetMembers<PresentationShapeBlock>(blocks, target, "node", "shape");
}

/** The labels (the words) of the targeted items. */
export function presentationSmartArtTargetLabels(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
): PresentationTextBlock[] {
  return targetMembers<PresentationTextBlock>(blocks, target, "label", "text");
}

/** The value every block states for `read`, or null when they differ. */
export function uniformSmartArtValue<T, V>(
  blocks: readonly T[],
  read: (block: T) => V,
): V | null {
  if (blocks.length === 0) return null;
  const first = read(blocks[0]);
  return blocks.every((block) => Object.is(read(block), first)) ? first : null;
}

/**
 * One patch on the targeted items' node shapes -- Shape Fill, Shape Outline,
 * Change Shape -- or null when nothing changes or a shape may not.
 */
export function presentationSmartArtBlocksWithNodePatch(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  patch: Partial<PresentationShapeBlock>,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const nodes = presentationSmartArtTargetNodes(blocks, target);
  if (nodes.length === 0 || nodes.some((node) => !canManipulate(node))) return null;
  const ids = new Set(nodes.map((node) => node.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (!ids.has(block.id)) return block;
    const written = shapeBlockWithPatch(block as PresentationShapeBlock, patch);
    if (written !== block) changed = true;
    return written as PresentationBlock;
  });
  return changed ? next : null;
}

/** Shape Fill with a colour: a solid fill that replaces the style's gradient. */
export function smartArtShapeFillPatch(color: string): Partial<PresentationShapeBlock> {
  return {
    fill: color,
    fillTo: undefined,
    gradientAngle: undefined,
    fillGradient: undefined,
    fillImage: undefined,
    fillPattern: undefined,
  } as Partial<PresentationShapeBlock>;
}

/** Shape Outline with a colour. */
export function smartArtShapeOutlinePatch(color: string): Partial<PresentationShapeBlock> {
  return {
    stroke: color,
    strokeTo: undefined,
    strokeGradientAngle: undefined,
  } as Partial<PresentationShapeBlock>;
}

/** Change Shape. */
export function smartArtChangeShapePatch(shape: ShapeKind): Partial<PresentationShapeBlock> {
  return { shape } as Partial<PresentationShapeBlock>;
}

/** Text Fill: the targeted items' words in `color`, or null when nothing changes. */
export function presentationSmartArtBlocksWithTextFill(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  color: string,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const labels = presentationSmartArtTargetLabels(blocks, target);
  if (labels.length === 0 || labels.some((label) => !canManipulate(label))) return null;
  const ids = new Set(labels.map((label) => label.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (!ids.has(block.id)) return block;
    const written = PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER.write(block as PresentationTextBlock, { color });
    if (written !== block) changed = true;
    return written as PresentationBlock;
  });
  return changed ? next : null;
}

/**
 * Shape Effects on the targeted shapes (PPT-A010): one preset, the deck's own
 * vocabulary for an ordinary shape. The style's exact shadow comes off so the
 * pick is what the shape wears -- "No Effect" included, which Office keeps as
 * an empty effect list on the shape's point.
 */
export function smartArtShapeEffectPatch(
  effect: OfficeShapeEffect,
): Partial<PresentationShapeBlock> {
  return {
    shadow: effect === "shadow" ? true : undefined,
    shapeShadow: undefined,
    effect: normalizeOfficeShapeEffectPreset(effect) ?? undefined,
  } as Partial<PresentationShapeBlock>;
}

/**
 * SmartArt Format > Shape Effects as PowerPoint draws it -- Preset, Shadow,
 * Reflection, Glow, Soft Edges, Bevel, 3-D Rotation, every variation -- on
 * the targeted shapes: each node's own Office effects (`shapeEffects`, the
 * model an ordinary shape and a picture share) after the pick, from what it
 * already wears (the style's shadow included, as PowerPoint's point keeps
 * it). The effects are the author's from then on (`smartArtAuthorFormat`),
 * which the export writes on the item's data point. A pick that gives the
 * shape a 3-D look of its own takes the style's `shape3d` / `scene3d` off it
 * (`smartArtEffectsOwn3d`): Office's point-level 3-D replaces the style's,
 * and the shape's effects then draw its solid alone. Null when nothing
 * changes or a shape may not.
 */
export function presentationSmartArtBlocksWithShapeEffect<K extends OfficePictureEffectKind>(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  kind: K,
  value: OfficePictureEffectValues[K] | null,
  accent1: string,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  return presentationSmartArtBlocksWithNodePatches(
    blocks,
    target,
    (node) => {
      // No Bevel over a SmartArt Style that bevels the shape: Office states
      // the point's own `a:sp3d` with no `a:bevelT`, replacing the style's
      // (`bevel: "none"`; `smartArtShape3dUnderEffects` draws it). A shape
      // whose own bevel took the style's 3-D off keeps the statement too, so
      // the style's bevel does not come back where it is opened.
      const ownBevel = presentationShapeEffects(node, accent1)?.bevel;
      const noStyleBevel =
        kind === "bevel" &&
        value === null &&
        (smartArtNodeWearsStyleBevel(node) ||
          (!node.shape3d && ownBevel !== undefined && ownBevel !== "none"));
      return presentationShapeEffectPatch(node, kind, value, accent1, noStyleBevel);
    },
    canManipulate,
  );
}

/**
 * The effects the targeted nodes' SmartArt Style draws (the first node's, as
 * the menu checks a multi-selection's first shape): its Bevel, so No Bevel
 * reads unchecked and takes the style's off.
 */
export function presentationSmartArtStyleEffects(
  nodes: readonly PresentationShapeBlock[],
): ReadonlySet<OfficePictureEffectKind> {
  return new Set<OfficePictureEffectKind>(
    nodes[0] && smartArtNodeWearsStyleBevel(nodes[0]) ? ["bevel"] : [],
  );
}

/**
 * Each targeted node after its own patch (Format Shape's pane on the shapes
 * selected inside the graphic: an effect joins what each already wears).
 * The shared node writer states the author's authority; a patch whose Office
 * effects give the shape its own 3-D takes the style's 3-D off it
 * (`smartArtEffectsOwn3d`). Null when nothing changes or a shape may not.
 */
export function presentationSmartArtBlocksWithNodePatches(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  patchOf: (node: PresentationShapeBlock) => Partial<PresentationShapeBlock> | null,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const nodes = presentationSmartArtTargetNodes(blocks, target);
  if (nodes.length === 0 || nodes.some((node) => !canManipulate(node))) return null;
  const ids = new Set(nodes.map((node) => node.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (!ids.has(block.id)) return block;
    const node = block as PresentationShapeBlock;
    const stated = patchOf(node);
    if (!stated) return block;
    const patch = { ...stated } as Record<string, unknown>;
    if ("shapeEffects" in patch && smartArtEffectsOwn3d(patch.shapeEffects)) {
      patch.shape3d = undefined;
      patch.scene3d = undefined;
    }
    const written = shapeBlockWithPatch(node, patch as Partial<PresentationShapeBlock>);
    if (written !== block) changed = true;
    return written as PresentationBlock;
  });
  // A shape's own turn moves its words with it (`./office-smartart/member-camera`).
  return changed ? withSmartArtLabelCameras(next) : null;
}

/** The Office effects the targeted nodes wear: the first node's, as the
 * deck's Shape Effects menu checks a multi-selection's first shape. */
export function presentationSmartArtShapeEffects(
  nodes: readonly PresentationShapeBlock[],
  accent1: string,
): OfficePictureEffects | undefined {
  return nodes[0] ? presentationShapeEffects(nodes[0], accent1) : undefined;
}

/**
 * A Shape Styles pick on the targeted items: the deck's own recipe on each
 * node (fill or gradient, line, shadow -- `presentationShapeStylePatch`) and
 * the recipe's text colour on each label, as Office's style sets both on the
 * shape it formats. Null when nothing changes or a member may not.
 */
export function presentationSmartArtBlocksWithStyle(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  recipe: OfficeShapeStyleRecipe,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const nodes = presentationSmartArtTargetNodes(blocks, target);
  if (nodes.length === 0 || nodes.some((node) => !canManipulate(node))) return null;
  const patches = new Map(
    nodes.map((node) => {
      const { style: _style, runs: _runs, ...patch } = presentationShapeStylePatch(node, recipe);
      return [node.id, patch as Partial<PresentationShapeBlock>] as const;
    }),
  );
  let changed = false;
  const painted = blocks.map((block) => {
    const patch = patches.get(block.id);
    if (!patch) return block;
    const written = shapeBlockWithPatch(block as PresentationShapeBlock, patch);
    if (written !== block) changed = true;
    return written as PresentationBlock;
  });
  const lettered = presentationSmartArtBlocksWithTextFill(painted, target, recipe.textColor, canManipulate);
  return lettered ?? (changed ? painted : null);
}

/** The recipe the targeted nodes all wear, by the deck's own match rule. */
export function presentationSmartArtStyleValue(
  nodes: readonly PresentationShapeBlock[],
  labels: readonly PresentationTextBlock[],
  recipes: readonly OfficeShapeStyleRecipe[],
): string | null {
  if (nodes.length === 0) return null;
  const colour = uniformSmartArtValue(labels, (label) => {
    const value = (label.style as { color?: unknown } | undefined)?.color;
    return typeof value === "string" ? value : null;
  });
  return (
    recipes.find((recipe) =>
      nodes.every((node) =>
        presentationShapeMatchesStyle(
          { ...node, runs: undefined, style: { color: colour ?? undefined } } as PresentationShapeBlock,
          recipe,
        ),
      ),
    )?.id ?? null
  );
}

/**
 * WordArt on the targeted items' words (WordArt Styles, Text Outline, Text
 * Effects; Text Fill once the words wear WordArt): the shared paint on each
 * label, `null` to remove it. A solid fill is also the label's colour (Text
 * Fill IS the text colour in PowerPoint) and a colour stated on a run comes
 * off, or it would paint over the new fill.
 */
export function presentationSmartArtBlocksWithTextPaint(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  paint: OfficeTextPaint | null,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const labels = presentationSmartArtTargetLabels(blocks, target);
  if (labels.length === 0 || labels.some((label) => !canManipulate(label))) return null;
  const solid = paint && paint.fillGradient === undefined && /^#[0-9a-f]{6}$/i.test(paint.fill)
    ? paint.fill
    : null;
  const ids = new Set(labels.map((label) => label.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (!ids.has(block.id)) return block;
    const coloured = solid
      ? (PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER.write(block as PresentationTextBlock, { color: solid }) as PresentationTextBlock)
      : (block as PresentationTextBlock);
    const record = { ...coloured } as PresentationTextBlock & { textPaint?: OfficeTextPaint };
    if (paint) record.textPaint = paint;
    else delete record.textPaint;
    if (
      coloured === block &&
      JSON.stringify((block as { textPaint?: unknown }).textPaint ?? null) === JSON.stringify(paint)
    ) {
      return block;
    }
    changed = true;
    return record as PresentationBlock;
  });
  return changed ? next : null;
}

/**
 * SmartArt Format > WordArt Styles (E026): the style's paint on the targeted
 * shapes' words, with the style's bold, as PowerPoint applies a gallery
 * style to the selected shapes' text (or every shape's, the graphic
 * selected). One change, or null when nothing would change.
 */
export function presentationSmartArtBlocksWithWordArtStyle(
  blocks: readonly PresentationBlock[],
  target: PresentationSmartArtFormatTarget,
  style: PresentationWordArtStyle,
  canManipulate: (block: PresentationBlock) => boolean,
): PresentationBlock[] | null {
  const painted =
    presentationSmartArtBlocksWithTextPaint(blocks, target, style.paint, canManipulate) ?? blocks;
  const labels = presentationSmartArtTargetLabels(painted, target);
  if (labels.length === 0 || labels.some((label) => !canManipulate(label))) return null;
  const ids = new Set(labels.map((label) => label.id));
  let changed = painted !== blocks;
  const next = painted.map((block) => {
    if (!style.bold || !ids.has(block.id)) return block;
    const bold = PRESENTATION_SMARTART_LABEL_TEXT_ADAPTER.write(block as PresentationTextBlock, {
      bold: true,
    }) as PresentationBlock;
    if (JSON.stringify(bold) !== JSON.stringify(block)) changed = true;
    return bold;
  });
  return changed ? next : null;
}

/** The WordArt the targeted labels all wear, or null when none or mixed. */
export function presentationSmartArtTextPaint(
  labels: readonly PresentationTextBlock[],
): OfficeTextPaint | null {
  if (labels.length === 0) return null;
  const first = (labels[0] as { textPaint?: OfficeTextPaint }).textPaint ?? null;
  return labels.every((label) => sameOfficeTextPaint((label as { textPaint?: unknown }).textPaint ?? null, first))
    ? first
    : null;
}

/**
 * What the Format tab shows for the targeted shapes: each value the targeted
 * nodes (or labels) share, null where they differ. `size` formats a node's
 * EMU width / height the way the deck's size boxes show it.
 */
export function presentationSmartArtFormatState(
  nodes: readonly PresentationShapeBlock[],
  labels: readonly PresentationTextBlock[],
  recipes: readonly OfficeShapeStyleRecipe[],
  size: {
    width: (node: PresentationShapeBlock) => string;
    height: (node: PresentationShapeBlock) => string;
  },
): PresentationSmartArtFormatState {
  const stroke = (node: PresentationShapeBlock) =>
    typeof node.stroke === "string" ? node.stroke : null;
  const effects = nodes.map((node) =>
    officeShapeEffectMenuValue(node.effect, node.shadow === true),
  );
  return {
    shape: uniformSmartArtValue(nodes, (node) => node.shape ?? null) as ShapeKind | null,
    fill: uniformSmartArtValue(nodes, (node) => (typeof node.fill === "string" ? node.fill : null)),
    outline: uniformSmartArtValue(nodes, stroke),
    outlineWeightPt: uniformSmartArtValue(nodes, (node) =>
      typeof node.strokeWidth === "number" ? node.strokeWidth : null,
    ),
    effect: effects.length && effects.every((value) => value === effects[0]) ? effects[0] : null,
    styleId: presentationSmartArtStyleValue(nodes, labels, recipes),
    textColor: uniformSmartArtValue(labels, (label) => {
      const color = (label.style as { color?: unknown } | undefined)?.color;
      return typeof color === "string" ? color : null;
    }),
    textPaint: presentationSmartArtTextPaint(labels),
    width: uniformSmartArtValue(nodes, size.width),
    height: uniformSmartArtValue(nodes, size.height),
  };
}

/**
 * SmartArt Format's Height / Width box on ONE shape: the per-axis scale that
 * takes its node to `centimetres` (the box's draft, already in centimetres),
 * both axes when the aspect is locked. Null for a value that is not a size.
 */
export function presentationSmartArtSizeFactor(
  node: PresentationShapeBlock,
  axis: "width" | "height",
  centimetres: string,
  lockAspect: boolean,
): { x: number; y: number } | null {
  const target = presentationPictureSizeEmuFromCmDraft(centimetres);
  const current = (node as { width?: unknown; height?: unknown })[axis];
  if (target === null || typeof current !== "number" || !(current > 0)) return null;
  const factor = target / current;
  if (lockAspect) return { x: factor, y: factor };
  return axis === "width" ? { x: factor, y: 1 } : { x: 1, y: factor };
}
