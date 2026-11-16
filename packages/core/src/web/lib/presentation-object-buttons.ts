/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The deck editor's Arrange buttons on a slide's objects, as pure transforms
 * of the slide's blocks: Align, Distribute, Flip, Rotate 90°, Group /
 * Ungroup, Bring Forward / Send Backward / to Front / to Back, Hide (Selection
 * Pane), Delete and Duplicate. The ribbon (`PresentationWindow`) resolves its
 * live selection and commits; an editor command resolves its explicit target
 * and commits; both run these same transforms.
 *
 * Each transform returns the slide's next blocks (or next slide), or null
 * when the button does not apply to those objects.
 */

import type { PresentationBlock } from "../components/console/editor/native-document/documentCommands";
import {
  alignBlocks,
  alignBlocksToBounds,
  distributeBlocks,
  distributeBlocksToBounds,
  type AlignEdge,
  type DistributeAxis,
} from "./presentation-align";
import {
  copyTimelineStep,
  MAX_TIMELINE_STEPS,
  newTimelineStepId,
  pruneSequence,
  timelineAnimationAuthoringCapability,
  type TimelineStep,
} from "./presentation-anim-timeline";
import { reidAndOffsetBlocks } from "./presentation-block-clipboard";
import { normalizePresentationClickAction } from "./presentation-block-fx";
import { canChangePresentationChartFrame } from "./presentation-chart-frame-operations";
import { recomputeConnectors } from "./presentation-connectors";
import {
  assignGroup,
  blockGroupId,
  blockGroupPath,
  blockNativeGroupAuthority,
  blockNativeGroupMoveAuthority,
  blocksInGroup,
  newGroupId,
  stackGroupAtFrontmostMember,
  ungroup,
} from "./presentation-crop-group";
import { mirrorGroupMemberGeometry, toggleFlip, type BlockGeometry } from "./presentation-drag";
import { PRESENTATION_OBJECT_GEOMETRY_MAX_EMU } from "./presentation-geometry-limit";
import { applyMultiRotate, rotatedBlockBounds } from "./presentation-multi-transform";
import { presentationRestackedPlaces } from "./presentation-paint-order";
import { quarterTurn, type QuarterTurnDirection } from "./presentation-rotate-quarter";
import {
  applyPresentationNativeObjectOrder,
  applyPresentationNativeOrderCommand,
  presentationNativeSelectionPaneObjectIsWritable,
} from "./presentation-command-adapters";
import { applyBlockHidden, isBlockHidden, selectionPaneObjectIdentity } from "./presentation-selection-pane";
import { readSmartArtBlockMarker } from "./presentation-smartart";
import { canEditBlockContent, canManipulateBlock, isFrameLockedByFile, isLockedFor } from "./presentation-text-case";
import { isVerifiedAuthoredSmartArtGroup } from "./presentation-authored-smartart";
import { readSmartArtGroupMembers } from "./office-smartart/design-model";
import { MAX_PRESENTATION_SLIDE_BLOCKS } from "./presentation/slide-block-budget";
import { blocksLeftOnSlide } from "./presentation-layouts";
import { geometryNumber, hasLayoutGeometry, isImportedPptxBlock, presentationBlockIsPlaceholder } from "./presentation-block-reads";
export { geometryNumber, hasLayoutGeometry, isImportedPptxBlock, presentationBlockIsPlaceholder };

export const DEFAULT_SLIDE_WIDTH_EMU = 9_144_000;
export const DEFAULT_SLIDE_HEIGHT_EMU = 5_143_500;

/** The fields of a slide these transforms read. */
export interface PresentationObjectSlide {
  id: string;
  blocks: PresentationBlock[];
  source?: unknown;
  animationSequence?: TimelineStep[];
  [key: string]: unknown;
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function clampNumber(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}

/** Read a strict-boolean flag off a block (e.g. `flipH`/`flipV`); else false. */
export function blockFlag(block: PresentationBlock, key: string): boolean {
  return (block as Record<string, unknown>)[key] === true;
}

/**
 * Whether a slide is drawn as a canvas of placed objects rather than the
 * legacy stacked title-and-body form: any placed object makes it one, and so
 * does an empty slide -- one on PowerPoint's Blank layout, which has no
 * placeholder at all (`blocksLeftOnSlide`).
 */
export function slideUsesCanvas(blocks: readonly PresentationBlock[]): boolean {
  return blocks.length === 0 || blocks.some(hasLayoutGeometry);
}









export function slideDrawnAsSlide(slide: { readonly blocks: readonly PresentationBlock[]; readonly [key: string]: unknown }): boolean {
  return slideUsesCanvas(slide.blocks) || slideDimension(slide, "widthEmu") !== null;
}

export function slideDimension(
  slide: { readonly [key: string]: unknown } | undefined,
  key: string,
): number | null {
  if (!slide) return null;
  const value = (slide as Record<string, unknown>)[key];
  return typeof value === "number" && Number.isFinite(value) && value > 0
    ? value
    : null;
}

/**
 * The slide's size: the one the deck states, else the built-in 16:9 default.
 * A slide is never sized by what sits on it: PowerPoint's slide size is the
 * presentation's (`p:sldSz`), an object dragged past an edge stays on the
 * pasteboard around the slide (the frame does not clip it), and the show and
 * the export cut it at the edge. The canvas used to grow to take in every
 * block, so one slide of a deck came out taller than the rest and showed
 * objects the exported file cuts off.
 */
export function slideBounds(slide: { readonly [key: string]: unknown } | undefined): {
  width: number;
  height: number;
} {
  const slideWidth = slideDimension(slide, "widthEmu");
  const slideHeight = slideDimension(slide, "heightEmu");
  if (slideWidth && slideHeight) {
    return { width: slideWidth, height: slideHeight };
  }
  return { width: DEFAULT_SLIDE_WIDTH_EMU, height: DEFAULT_SLIDE_HEIGHT_EMU };
}

/**
 * The slide size a NEW slide of this deck carries: `slide`'s explicit
 * `widthEmu` / `heightEmu`, or nothing for a deck at the built-in default.
 * PowerPoint's slide size is the presentation's, so every slide the editor
 * makes -- New Slide, an AI or outline deck, a design template -- takes the
 * deck's; a slide without it is drawn and exported at the 10 x 5.625 in
 * default, its placeholders laid out for the deck's size.
 */
export function deckSlideSizeFields(
  slide: { readonly [key: string]: unknown } | undefined,
): { widthEmu: number; heightEmu: number } | Record<string, never> {
  const width = slideDimension(slide, "widthEmu");
  const height = slideDimension(slide, "heightEmu");
  return width && height
    ? { widthEmu: Math.round(width), heightEmu: Math.round(height) }
    : {};
}







export function presentationImportedOrderChanged(
  before: readonly PresentationBlock[],
  after: readonly PresentationBlock[],
): boolean {
  // A member of a group the file owns is one of its objects too: its order
  // among the group's members is the file's (`p:grpSp`) to write.
  const imported = (blocks: readonly PresentationBlock[]) =>
    blocks.filter((block) => isImportedPptxBlock(block) || blockClaimsNativeGroup(block)).map((block) => block.id);
  const was = imported(before);
  const now = imported(after);
  return was.length !== now.length || was.some((id, index) => id !== now[index]);
}

/**
 * Whether `block` is a member of a group the file owns (its receipt says so).
 * An object Ungroup took out of such a group keeps the receipt it was imported
 * with -- the save reads the Ungroup off the page against it -- but is a group
 * member no more: Delete, Lock and a move of its own apply to it as to any
 * loose object of the file.
 */
export function blockClaimsNativeGroup(block: PresentationBlock): boolean {
  const record = block as Record<string, unknown>;
  const source = record.source;
  if (!blockGroupId(block)) return false;
  return Boolean(
    record.nativeGroupDisposition === "preserve-only" ||
    source &&
      typeof source === "object" &&
      !Array.isArray(source) &&
      Object.hasOwn(source, "nativeGroupAuthority"),
  );
}

export function exactNativeGroupMembers(
  blocks: PresentationBlock[],
  selected: PresentationBlock,
): PresentationBlock[] | null {
  const authority = blockNativeGroupAuthority(
    selected as unknown as Record<string, unknown>,
  );
  const groupId = blockGroupId(selected);
  if (!authority || !groupId) return null;
  const byShapeId = new Map<string, PresentationBlock>();
  for (const block of blocks) {
    if (blockGroupId(block) !== groupId) continue;
    const candidate = blockNativeGroupAuthority(
      block as unknown as Record<string, unknown>,
    );
    const source = (block as Record<string, unknown>).source;
    const shapeId =
      source && typeof source === "object" && !Array.isArray(source)
        ? (source as Record<string, unknown>).shapeId
        : null;
    if (
      !candidate ||
      candidate.groupShapeId !== authority.groupShapeId ||
      candidate.groupFingerprint !== authority.groupFingerprint ||
      typeof shapeId !== "string" ||
      byShapeId.has(shapeId)
    ) {
      return null;
    }
    byShapeId.set(shapeId, block);
  }
  const ordered = authority.members.map((member) => byShapeId.get(member.shapeId));
  return ordered.length === byShapeId.size && ordered.every(Boolean)
    ? (ordered as PresentationBlock[])
    : null;
}

export function isPresentationCameoBlock(block: PresentationBlock): boolean {
  return block.type === "shape" && block.cameo === true;
}

/**
 * Whether the Selection Pane may hide, show or rename an object the editor
 * made itself: one with a model identity, no native order hint and no
 * imported custody.
 */
export function selectionPaneBlockHasModelMutationAuthority(
  block: PresentationBlock,
): boolean {
  // A paint order (`sourceOrder` / `zIndex`) alone is no native custody: the
  // AI composer writes both on every block it makes, and the pane could
  
  // imported object is one that names its slide and shape, or inherits
  // from a layout or master (`isImportedPptxBlock`).
  return (
    /^(?:model|group):/.test(selectionPaneObjectIdentity(block) ?? "") &&
    !isImportedPptxBlock(block)
  );
}

/** Whether the Selection Pane may change an imported object: its native receipt allows it. */
export function selectionPaneBlockHasNativeMutationAuthority(
  slide: Parameters<typeof presentationNativeSelectionPaneObjectIsWritable>[0],
  block: PresentationBlock,
): boolean {
  return presentationNativeSelectionPaneObjectIsWritable(slide, block.id);
}







export function selectionPaneBlockIsRenamable(
  slide: Parameters<typeof presentationNativeSelectionPaneObjectIsWritable>[0],
  block: PresentationBlock,
): boolean {
  return (
    canEditBlockContent(block) &&
    (selectionPaneBlockHasModelMutationAuthority(block) ||
      selectionPaneBlockHasNativeMutationAuthority(slide, block))
  );
}

/** An object the Drawing geometry commands may move, turn or flip. */
export function isGeometryCommandTarget(block: PresentationBlock): boolean {
  return hasLayoutGeometry(block) && canManipulateBlock(block) && canChangePresentationChartFrame(block, "move");
}

/**
 * A member of a SmartArt graphic the page keeps as the file drew it (its
 * native diagram, not yet converted): the graphic is its frame on the slide,
 * and the save writes a move of all of it as that frame's offset
 * (`split_native_projection_diagram_moves`). Its members are locked against
 * every edit of their own.
 */
export function isPreservedDiagramMember(block: PresentationBlock): boolean {
  const record = block as Record<string, unknown>;
  if (record.nativeGroupDisposition !== "preserve-only") return false;
  if (!blockGroupId(block)?.startsWith("native-smartart-")) return false;
  const source = asRecord(record.source);
  return Boolean(
    source && (asRecord(source.nativeSmartArtView) || typeof source.nativeSmartArtBlocker === "string"),
  );
}

/** The graphic frame (`nativeSmartArtView.frameEmu`) every member of one SmartArt graphic of the file names, or null. */
function preservedDiagramFrame(members: readonly PresentationBlock[]): BlockGeometry | null {
  if (members.length === 0 || !members.every(isPreservedDiagramMember)) return null;
  const frames = members.map((member) => asRecord(asRecord(asRecord((member as Record<string, unknown>).source)?.nativeSmartArtView)?.frameEmu));
  const first = frames[0];
  if (!first || !(["left", "top", "width", "height"] as const).every((key) => Number.isInteger(first[key]))) return null;
  if (!frames.every((frame) => frame && frame.left === first.left && frame.top === first.top && frame.width === first.width && frame.height === first.height)) {
    return null;
  }
  return { left: first.left as number, top: first.top as number, width: first.width as number, height: first.height as number };
}

/**
 * A member of a SmartArt graphic of the file moved with its frame: the frame
 * it names moves by the same offset, so the page always says where the
 * graphic is (the save reads the move from the members and writes it as the
 * frame's offset).
 */
function withPreservedDiagramFrameMoved(block: PresentationBlock, dx: number, dy: number): PresentationBlock {
  const record = block as Record<string, unknown>;
  const source = asRecord(record.source);
  const view = asRecord(source?.nativeSmartArtView);
  const frame = asRecord(view?.frameEmu);
  if (!source || !view || !frame || !Number.isInteger(frame.left) || !Number.isInteger(frame.top)) return block;
  return {
    ...block,
    source: {
      ...source,
      nativeSmartArtView: { ...view, frameEmu: { ...frame, left: (frame.left as number) + dx, top: (frame.top as number) + dy } },
    },
  } as PresentationBlock;
}

/**
 * An object Align and Distribute may move: a geometry target the file does
 * not lock in place, or a SmartArt graphic of the file, which PowerPoint
 * aligns and distributes as ONE object (its graphic frame) without opening
 * it -- every member moves by one offset and keeps its lock.
 */
export function isArrangeTarget(block: PresentationBlock): boolean {
  if (isFrameLockedByFile(block, "move")) return false;
  return isGeometryCommandTarget(block) || (isPreservedDiagramMember(block) && hasLayoutGeometry(block));
}

/** A frame with every edge clamped to the persisted geometry limits. */
export function withClampedBlockGeometry(block: PresentationBlock, geom: BlockGeometry): PresentationBlock {
  return {
    ...block,
    left: clampNumber(Math.round(geom.left), -PRESENTATION_OBJECT_GEOMETRY_MAX_EMU, PRESENTATION_OBJECT_GEOMETRY_MAX_EMU),
    top: clampNumber(Math.round(geom.top), -PRESENTATION_OBJECT_GEOMETRY_MAX_EMU, PRESENTATION_OBJECT_GEOMETRY_MAX_EMU),
    width: clampNumber(Math.round(geom.width), 1, PRESENTATION_OBJECT_GEOMETRY_MAX_EMU),
    height: clampNumber(Math.round(geom.height), 1, PRESENTATION_OBJECT_GEOMETRY_MAX_EMU),
  } as PresentationBlock;
}











export function distributeBox(frame: BlockGeometry, rotation: number): BlockGeometry {
  const turned = ((rotation % 360) + 360) % 360;
  const swapped = (turned >= 45 && turned < 135) || (turned >= 225 && turned < 315);
  if (!swapped) return frame;
  const centreX = frame.left + frame.width / 2;
  const centreY = frame.top + frame.height / 2;
  return { left: centreX - frame.height / 2, top: centreY - frame.width / 2, width: frame.height, height: frame.width };
}

/**
 * Align or Distribute: one batch geometry transform over the targets, each
 * group moved as one object by its bounding box.
 */
export function arrangeBlockGeometry(
  blocks: PresentationBlock[],
  slide: { readonly [key: string]: unknown } | undefined,
  requested: readonly PresentationBlock[],
  transform: (
    rects: BlockGeometry[],
    bounds: { width: number; height: number },
  ) => BlockGeometry[],
  minTargets = 2,
  /** How an object is measured: "visible" (Align: its turned box) or "distribute" (see {@link distributeBox}). */
  measure: "visible" | "distribute" = "visible",
): PresentationBlock[] | null {
  if (requested.length < minTargets) return null;
  // A SmartArt graphic of the file moves whole: every member of it, whichever
  // of them the selection names.
  const chosen = new Set(requested.map((block) => block.id));
  const targets = [...requested];
  for (const block of requested) {
    if (!isPreservedDiagramMember(block)) continue;
    for (const member of blocksInGroup(blocks, blockGroupId(block) as string)) {
      if (chosen.has(member.id)) continue;
      chosen.add(member.id);
      targets.push(member);
    }
  }
  const bounds = slideBounds(slide);
  const frameOf = (block: PresentationBlock): BlockGeometry => ({
    left: geometryNumber(block, "left") ?? 0,
    top: geometryNumber(block, "top") ?? 0,
    width:
      geometryNumber(block, "width") ?? Math.round(bounds.width * 0.35),
    height:
      geometryNumber(block, "height") ??
      Math.round(bounds.height * 0.18),
  });
  // Align and Distribute measure the object as it is seen: a rotated
  // object's visible box, not its unrotated frame (PPT-B219). The frame
  // is then moved by the same delta as its visible box.
  const rectOf = (block: PresentationBlock): BlockGeometry =>
    measure === "distribute"
      ? distributeBox(frameOf(block), geometryNumber(block, "rotation") ?? 0)
      : rotatedBlockBounds(frameOf(block), geometryNumber(block, "rotation") ?? 0);
  // A group is ONE object to Align and Distribute, as in PowerPoint: it
  // moves by its bounding box and its members keep their layout
  // (PPT-B076: centring a group centred each member on top of the
  // others). Every other selected object is its own unit.
  const units: Array<{ members: PresentationBlock[]; rect: BlockGeometry }> = [];
  const unitByGroup = new Map<string, number>();
  for (const block of targets) {
    const groupId = blockGroupId(block);
    const existing = groupId === undefined ? undefined : unitByGroup.get(groupId);
    if (existing !== undefined) {
      units[existing].members.push(block);
      continue;
    }
    if (groupId !== undefined) unitByGroup.set(groupId, units.length);
    units.push({ members: [block], rect: rectOf(block) });
  }
  for (const unit of units) {
    // A SmartArt graphic of the file is its graphic frame, as PowerPoint
    // aligns it, not the box its drawn shapes fill.
    const diagramFrame = preservedDiagramFrame(unit.members);
    if (diagramFrame) {
      unit.rect = diagramFrame;
      continue;
    }
    if (unit.members.length < 2) continue;
    const rects = unit.members.map(rectOf);
    const left = Math.min(...rects.map((rect) => rect.left));
    const top = Math.min(...rects.map((rect) => rect.top));
    unit.rect = {
      left,
      top,
      width: Math.max(...rects.map((rect) => rect.left + rect.width)) - left,
      height: Math.max(...rects.map((rect) => rect.top + rect.height)) - top,
    };
  }
  if (units.length < minTargets) return null;
  const transformed = transform(units.map((unit) => unit.rect), bounds);
  if (transformed.length !== units.length) return null;
  const byId = new Map<string, BlockGeometry>();
  const frameMoves = new Map<string, { dx: number; dy: number }>();
  units.forEach((unit, index) => {
    const next = transformed[index];
    const dx = Math.round(next.left - unit.rect.left);
    const dy = Math.round(next.top - unit.rect.top);
    for (const member of unit.members) {
      const frame = frameOf(member);
      byId.set(member.id, { ...frame, left: frame.left + dx, top: frame.top + dy });
      if (isPreservedDiagramMember(member)) frameMoves.set(member.id, { dx, dy });
    }
  });
  let changed = false;
  const nextBlocks = blocks.map((block) => {
    const geom = byId.get(block.id);
    if (!geom) return block;
    const moved = withClampedBlockGeometry(block, geom);
    const frameMove = frameMoves.get(block.id);
    const next = frameMove ? withPreservedDiagramFrameMoved(moved, frameMove.dx, frameMove.dy) : moved;
    if (
      geometryNumber(next, "left") === geometryNumber(block, "left") &&
      geometryNumber(next, "top") === geometryNumber(block, "top") &&
      geometryNumber(next, "width") === geometryNumber(block, "width") &&
      geometryNumber(next, "height") === geometryNumber(block, "height")
    ) {
      return block;
    }
    changed = true;
    return next;
  });
  if (!changed) return null;
  // Re-snap connectors anchored to moved targets while preserving a
  // directly transformed connector's own geometry.
  return recomputeConnectors(nextBlocks, new Set(byId.keys()));
}

/** Align: to the slide (each object, one or more) or to the selection (two or more). */
export function alignBlocksOnSlide(
  blocks: PresentationBlock[],
  slide: { readonly [key: string]: unknown } | undefined,
  targets: readonly PresentationBlock[],
  edge: AlignEdge,
  toSlide: boolean,
): PresentationBlock[] | null {
  return toSlide
    ? arrangeBlockGeometry(blocks, slide, targets, (rects, bounds) => alignBlocksToBounds(rects, edge, bounds), 1)
    : arrangeBlockGeometry(blocks, slide, targets, (rects) => alignBlocks(rects, edge), 2);
}

/**
 * Distribute, under the same Align to Slide / Align Selected Objects choice
 * as Align (PowerPoint's `RelativeTo`): over the slide, one or more objects
 * (equal gaps, the two at the slide's edges included, so one is centred);
 * over the selection's own span, three or more.
 */
export function distributeBlocksOnSlide(
  blocks: PresentationBlock[],
  slide: { readonly [key: string]: unknown } | undefined,
  targets: readonly PresentationBlock[],
  axis: DistributeAxis,
  toSlide = false,
): PresentationBlock[] | null {
  return toSlide
    ? arrangeBlockGeometry(blocks, slide, targets, (rects, bounds) => distributeBlocksToBounds(rects, axis, bounds), 1, "distribute")
    : arrangeBlockGeometry(blocks, slide, targets, (rects) => distributeBlocks(rects, axis), 3, "distribute");
}

/** Distribute's unit transform for `arrangeBlockGeometry`, slide or selection. */
export function distributeBlocksToSlideOrSelection(
  axis: DistributeAxis,
  toSlide: boolean,
): (rects: BlockGeometry[], bounds: { width: number; height: number }) => BlockGeometry[] {
  return toSlide
    ? (rects, bounds) => distributeBlocksToBounds(rects, axis, bounds)
    : (rects) => distributeBlocks(rects, axis);
}

/** The fewest selected objects Distribute acts on, under Align to Slide or not. */
export function distributeMinimumTargets(toSlide: boolean): number {
  return toSlide ? 1 : 3;
}

/**
 * Flip Horizontal / Vertical. A whole selected group flips as one object:
 * its members are mirrored across the group's frame, as PowerPoint flips a
 * group (PPT-B019); a loose multi-selection (and a SmartArt, whose members its
 * layout places) flips each in place. A flag set back to false is deleted.
 */
export function flipBlocks(
  blocks: PresentationBlock[],
  slide: { readonly [key: string]: unknown } | undefined,
  targets: readonly PresentationBlock[],
  axis: "flipH" | "flipV",
): PresentationBlock[] {
  const targetIds = new Set(targets.map((block) => block.id));
  const bounds = slideBounds(slide);
  const groups = new Map<string, PresentationBlock[]>();
  for (const block of targets) {
    const groupId = blockGroupId(block);
    if (!groupId || readSmartArtBlockMarker(block)) continue;
    groups.set(groupId, [...(groups.get(groupId) ?? []), block]);
  }
  const mirrored = new Map<string, { left: number; top: number; rotation: number }>();
  for (const [groupId, members] of groups) {
    if (
      members.length < 2 ||
      blocks.filter((block) => blockGroupId(block) === groupId).length !== members.length
    ) {
      continue;
    }
    const rects = members.map((member) => ({
      left: geometryNumber(member, "left") ?? 0,
      top: geometryNumber(member, "top") ?? 0,
      width: geometryNumber(member, "width") ?? Math.round(bounds.width * 0.35),
      height: geometryNumber(member, "height") ?? Math.round(bounds.height * 0.18),
    }));
    const frame = {
      left: Math.min(...rects.map((rect) => rect.left)),
      top: Math.min(...rects.map((rect) => rect.top)),
      right: Math.max(...rects.map((rect) => rect.left + rect.width)),
      bottom: Math.max(...rects.map((rect) => rect.top + rect.height)),
    };
    members.forEach((member, index) =>
      mirrored.set(
        member.id,
        mirrorGroupMemberGeometry(
          rects[index],
          geometryNumber(member, "rotation") ?? 0,
          frame,
          axis,
        ),
      ),
    );
  }
  const nextBlocks = blocks.map((block) => {
    if (!targetIds.has(block.id)) return block;
    const next = { ...block } as PresentationBlock &
      Record<string, unknown>;
    const flipped = toggleFlip(blockFlag(block, axis));
    if (flipped) next[axis] = true;
    else delete next[axis];
    const moved = mirrored.get(block.id);
    if (moved) {
      next.left = moved.left;
      next.top = moved.top;
      if (moved.rotation === 0) delete next.rotation;
      else next.rotation = moved.rotation;
    }
    return next;
  });
  return mirrored.size > 0
    ? recomputeConnectors(nextBlocks, new Set(mirrored.keys()))
    : nextBlocks;
}

/**
 * Rotate Right / Left 90°: each target's rotation by a quarter turn (an
 * off-grid angle snaps to the nearest quarter first); a group turns as one
 * about its centre (PPT-B222). An upright result deletes `rotation`.
 */
export function quarterTurnBlocks(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
  direction: QuarterTurnDirection,
): PresentationBlock[] {
  const targetIds = new Set(targets.map((block) => block.id));
  const delta = direction === "right" ? 90 : -90;
  const groupMoves = new Map<string, { left: number; top: number }>();
  const byGroup = new Map<string, PresentationBlock[]>();
  for (const block of targets) {
    const groupId = blockGroupId(block);
    if (groupId) byGroup.set(groupId, [...(byGroup.get(groupId) ?? []), block]);
  }
  for (const members of byGroup.values()) {
    if (members.length < 2) continue;
    const turned = applyMultiRotate(
      members.map((block) => ({
        blockId: block.id,
        startGeom: {
          left: geometryNumber(block, "left") ?? 0,
          top: geometryNumber(block, "top") ?? 0,
          width: geometryNumber(block, "width") ?? 0,
          height: geometryNumber(block, "height") ?? 0,
        },
        startRotation: geometryNumber(block, "rotation") ?? 0,
      })),
      delta,
    );
    for (const [id, geometry] of turned) groupMoves.set(id, { left: geometry.left, top: geometry.top });
  }
  const nextBlocks = blocks.map((block) => {
    if (!targetIds.has(block.id)) return block;
    const current = geometryNumber(block, "rotation") ?? 0;
    const rotated = quarterTurn(current, direction);
    const next = { ...block } as PresentationBlock &
      Record<string, unknown>;
    if (rotated === 0) delete next.rotation;
    else next.rotation = rotated;
    const moved = groupMoves.get(block.id);
    if (moved) {
      next.left = moved.left;
      next.top = moved.top;
    }
    return next;
  });
  return recomputeConnectors(nextBlocks, targetIds);
}





















export type PresentationGroupRefusal = "placeholder" | "table" | "imported" | "object" | "locked";

const IMPORTED_GROUPABLE_TYPES = new Set(["shape", "text", "image", "chart"]);

function blockIsDiagramShape(block: PresentationBlock): boolean {
  const record = block as Record<string, unknown>;
  const groupId = blockGroupId(block);
  return (
    (block.type as string) === "smartart" ||
    Boolean(readSmartArtBlockMarker(block)) ||
    Object.keys(record).some((key) => key.startsWith("smartArt")) ||
    Boolean(groupId && groupId.startsWith("native-smartart-"))
  );
}

function blockInheritsFromLayout(block: PresentationBlock): boolean {
  const source = asRecord((block as Record<string, unknown>).source);
  return source?.inheritedFrom === "layout" || source?.inheritedFrom === "master";
}

/**
 * Whether `block`'s outermost group is one the file owns (imported with the
 * slide), not one made in this editor.
 */
function presentationGroupIsFiles(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): boolean {
  const groupId = blockGroupId(block);
  // The editor's own ids are `grp-`; the import names the file's groups
  // `native-group-` / `group-` (a Regroup restores that id).
  if (!groupId || groupId.startsWith("grp-")) return false;
  return blocks.some(
    (member) =>
      blockGroupId(member) === groupId &&
      (blockClaimsNativeGroup(member) ||
        blockNativeGroupMoveAuthority(member as unknown as Record<string, unknown>) !== null ||
        isImportedPptxBlock(member)),
  );
}

export function presentationGroupRefusal(
  targets: readonly PresentationBlock[],
  slideBlocks?: readonly PresentationBlock[],
): PresentationGroupRefusal | null {
  // Whichever entry asks: the Shape Format tab closed the command for a
  // placeholder while Home > Arrange and the context menu grouped one all
  // the same (PPT-B334).
  if (targets.some(presentationBlockIsPlaceholder)) return "placeholder";
  if (targets.some((block) => block.type === "table")) return "table";
  // Lock writes `noGrp`: a locked object joins no group until it is
  // unlocked, and neither does one whose file says `noGrp` (ECMA-376
  // 20.1.2.2.34). Every Group entry said nothing and grouped nothing.
  if (targets.some((block) => isLockedFor(block, "group"))) return "locked";
  if (
    targets.some(
      (block) =>
        isPresentationCameoBlock(block) ||
        blockInheritsFromLayout(block) ||
        // A diagram of the file is drawn from its own parts; the editor's
        // own SmartArt groups as it always did.
        (blockIsDiagramShape(block) && (isImportedPptxBlock(block) || Boolean(blockGroupId(block)?.startsWith("native-smartart-")))),
    )
  ) {
    return "object";
  }
  if (
    targets.some(
      (block) => isImportedPptxBlock(block) && !IMPORTED_GROUPABLE_TYPES.has(block.type as string),
    )
  ) {
    return "object";
  }
  if (slideBlocks) {
    // A group of the file is taken whole or not at all: the save cannot move
    // some of its members into another group.
    const chosen = new Set(targets.map((block) => block.id));
    for (const block of targets) {
      const groupId = blockGroupId(block);
      if (!groupId || !presentationGroupIsFiles(slideBlocks, block)) continue;
      if (blocksInGroup(slideBlocks, groupId).some((member) => !chosen.has(member.id))) return "imported";
    }
  }
  return null;
}

/**
 * Group: a fresh group id on two or more objects, stacked together. A whole
 * group among them nests inside the new one (PPT-B204). Regroup passes the
 * dissolved group's own id (`groupId`) so the restored group is the same
 * object it was -- for a group of the file, the save then finds the file's
 * own group unchanged and writes nothing.
 */
export function groupBlocks(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
  groupId?: string,
): PresentationBlock[] | null {
  if (targets.length < 2) return null;
  if (presentationGroupRefusal(targets, blocks) !== null) return null;
  const existingGroupId = blockGroupId(targets[0]);
  if (
    existingGroupId &&
    targets.every((block) => blockGroupId(block) === existingGroupId)
  ) {
    return null;
  }
  const targetIds = targets.map((block) => block.id);
  return presentationRestackedPlaces(
    stackGroupAtFrontmostMember(
      assignGroup(blocks, targetIds, groupId ?? newGroupId()),
      targetIds,
    ),
  );
}

/**
 * The outermost groups Ungroup takes apart for `targets`: each group a target
 * belongs to, when every member of a group the file owns is selected (a part
 * of it is not a group to ungroup) and it is not a SmartArt diagram (which
 * has Convert to Shapes) or locked. Null when there is none.
 */
export function presentationUngroupMembers(
  blocks: readonly PresentationBlock[],
  targets: readonly PresentationBlock[],
): PresentationBlock[] | null {
  const chosen = new Set(targets.map((block) => block.id));
  const groupIds = new Set<string>();
  for (const block of targets) {
    const groupId = blockGroupId(block);
    if (groupId) groupIds.add(groupId);
  }
  if (groupIds.size === 0) return null;
  const members: PresentationBlock[] = [];
  for (const groupId of groupIds) {
    const group = blocksInGroup(blocks, groupId);
    if (group.length === 0) return null;
    if (
      group.some(
        (member) =>
          blockIsDiagramShape(member) &&
          (isImportedPptxBlock(member) || groupId.startsWith("native-smartart-")),
      )
    ) {
      return null;
    }
    if (presentationGroupIsFiles(blocks, group[0]) && group.some((member) => !chosen.has(member.id))) {
      return null;
    }
    // A locked group (Lock writes the group's `noUngrp`) stays one.
    if (group.some((member) => isLockedFor(member, "ungroup") || isLockedFor(member, "ungroup", true))) {
      return null;
    }
    members.push(...group);
  }
  return members;
}

/**
 * Ungroup: one level off, as PowerPoint's Ungroup ("dissolves the group,
 * leaving the individual objects in their current location"): a group nested
 * inside stays a group. A group the file owns -- nested, mirrored, filled or
 * turned -- is taken apart the same way; the save lifts its members out of
 * the p:grpSp through the group's own map (`pptx/native_group_ungroup.py`).
 */
export function ungroupBlocks(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
): PresentationBlock[] | null {
  const members = presentationUngroupMembers(blocks, targets);
  if (!members) return null;
  return ungroup(
    blocks,
    members.map((block) => block.id),
  );
}

/**
 * Bring to Front / Send to Back / Bring Forward / Send Backward on the
 * slide's paint order (the blocks' array order). The targets keep their
 * relative order; a one-step move steps over a whole neighbouring object, so
 * an object never lands inside a group it passes (PPT-B205).
 */
export function reorderBlocks(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
  command: "front" | "forward" | "backward" | "back",
): PresentationBlock[] | null {
  // Objects inside a group restack among the group's own objects, the group
  // where it stands: "in a group, you can only reorder it within the group"
  // (Microsoft Support, Manage objects with the Selection Pane). Bring to
  // Front on a member took it out of the group's place, and the save into
  
  const plan = presentationReorderUnits(blocks, targets);
  if (plan.inGroup.size === 0) return reorderSlideObjects(blocks, targets, command);
  const loose = targets.filter((block) => !plan.inGroupTargetIds.has(block.id));
  let next = loose.length ? reorderSlideObjects(blocks, loose, command) ?? [...blocks] : [...blocks];
  let changed = next !== blocks && next.some((block, index) => block !== blocks[index]);
  for (const [parentKey, unitKeys] of plan.inGroup) {
    const restacked = restackWithinGroup(next, parentKey.split("\u0000"), unitKeys, command);
    if (restacked) {
      next = restacked;
      changed = true;
    }
  }
  return changed ? presentationRestackedPlaces(next) : null;
}

/** The groups a block sits in, outermost first (`groupId`, then `groupPath`). */
function blockGroupChain(block: PresentationBlock): string[] {
  const outer = blockGroupId(block);
  return outer ? [outer, ...blockGroupPath(block)] : [];
}

function chainStartsWith(chain: readonly string[], prefix: readonly string[]): boolean {
  return prefix.length <= chain.length && prefix.every((id, index) => chain[index] === id);
}

/** A group's object a block belongs to: one of its own (`b:`) or a group
 * nested in it (`g:`), whose objects then move together. */
function unitKeyIn(block: PresentationBlock, parent: readonly string[]): string {
  const chain = blockGroupChain(block);
  return chain.length > parent.length ? `g:${chain[parent.length]}` : `b:${block.id}`;
}

/**
 * The objects `targets` select, each in the group it is an object of. A
 * target whose whole group (or a whole group nested in it) is selected is
 * that group; a group selected whole at the top is an object of the slide.
 * `inGroup`: per group (its chain joined by NUL) the objects of it to
 * restack inside it.
 */
function presentationReorderUnits(
  blocks: readonly PresentationBlock[],
  targets: readonly PresentationBlock[],
): { inGroup: Map<string, Set<string>>; inGroupTargetIds: Set<string> } {
  const targetIds = new Set(targets.map((block) => block.id));
  const inGroup = new Map<string, Set<string>>();
  const inGroupTargetIds = new Set<string>();
  for (const target of targets) {
    const chain = blockGroupChain(target);
    if (chain.length === 0) continue;
    // The outermost group of the chain selected whole names the object.
    let level = -1;
    for (let depth = 1; depth <= chain.length; depth += 1) {
      const prefix = chain.slice(0, depth);
      if (blocks.every((block) => !chainStartsWith(blockGroupChain(block), prefix) || targetIds.has(block.id))) {
        level = depth;
        break;
      }
    }
    if (level === 1) continue;
    const parent = level === -1 ? chain : chain.slice(0, level - 1);
    const key = parent.join("\u0000");
    const units = inGroup.get(key) ?? new Set<string>();
    units.add(unitKeyIn(target, parent));
    inGroup.set(key, units);
    inGroupTargetIds.add(target.id);
  }
  return { inGroup, inGroupTargetIds };
}

/**
 * Bring to Front / Send to Back / Bring Forward / Send Backward among the
 * objects of one group: its blocks keep the places in the list the group
 * holds, and a one-step move steps over one whole object of the group (a
 * group nested in it moves and is stepped over whole). Null when nothing
 * moves.
 */
function restackWithinGroup(
  blocks: readonly PresentationBlock[],
  parent: readonly string[],
  moving: ReadonlySet<string>,
  command: "front" | "forward" | "backward" | "back",
): PresentationBlock[] | null {
  const slots: number[] = [];
  const units: Array<{ key: string; blocks: PresentationBlock[] }> = [];
  const unitByKey = new Map<string, { key: string; blocks: PresentationBlock[] }>();
  blocks.forEach((block, index) => {
    if (!chainStartsWith(blockGroupChain(block), parent)) return;
    slots.push(index);
    const key = unitKeyIn(block, parent);
    let unit = unitByKey.get(key);
    if (!unit) {
      unit = { key, blocks: [] };
      unitByKey.set(key, unit);
      units.push(unit);
    }
    unit.blocks.push(block);
  });
  let order = [...units];
  if (command === "front" || command === "back") {
    const stay = order.filter((unit) => !moving.has(unit.key));
    const go = order.filter((unit) => moving.has(unit.key));
    order = command === "front" ? [...stay, ...go] : [...go, ...stay];
  } else if (command === "forward") {
    for (let index = order.length - 2; index >= 0; index -= 1) {
      if (moving.has(order[index].key) && !moving.has(order[index + 1].key)) {
        [order[index], order[index + 1]] = [order[index + 1], order[index]];
      }
    }
  } else {
    for (let index = 1; index < order.length; index += 1) {
      if (moving.has(order[index].key) && !moving.has(order[index - 1].key)) {
        [order[index - 1], order[index]] = [order[index], order[index - 1]];
      }
    }
  }
  const placed = order.flatMap((unit) => unit.blocks);
  if (placed.every((block, index) => blocks[slots[index]] === block)) return null;
  const next = [...blocks];
  slots.forEach((slot, index) => {
    next[slot] = placed[index];
  });
  return withGroupAltOnFirstMember(next, parent[0]);
}

/**
 * The group's own Alt Text (`groupAlt` / `groupDecorative`, the `p:grpSp`'s
 * description) rides on the group's first object, as the file reads it back
 * (its first member carries it) and the save looks for it: a restack inside
 * the group that put another object first moves it there.
 */
function withGroupAltOnFirstMember(blocks: PresentationBlock[], groupId: string | undefined): PresentationBlock[] {
  if (!groupId) return blocks;
  const members = blocks.map((block, index) => ({ block, index })).filter(({ block }) => blockGroupId(block) === groupId);
  const carrier = members.find(({ block }) => Object.hasOwn(block, "groupAlt") || Object.hasOwn(block, "groupDecorative"));
  const first = members[0];
  if (!carrier || !first || carrier.index === first.index) return blocks;
  const carried = carrier.block as PresentationBlock & { groupAlt?: unknown; groupDecorative?: unknown };
  const { groupAlt, groupDecorative, ...rest } = carried;
  const lead: Record<string, unknown> = { ...(first.block as Record<string, unknown>) };
  if (Object.hasOwn(carried, "groupAlt")) lead.groupAlt = groupAlt;
  if (Object.hasOwn(carried, "groupDecorative")) lead.groupDecorative = groupDecorative;
  const next = [...blocks];
  next[carrier.index] = rest as PresentationBlock;
  next[first.index] = lead as PresentationBlock;
  return next;
}










export function presentationNativeRestack<S extends { readonly blocks: readonly PresentationBlock[] }>(
  slide: S,
  targets: readonly PresentationBlock[],
  /** Null: `listed` is the order of several commands (a Selection Pane
   * drop), which the journal takes as a whole new order. */
  command: "front" | "forward" | "backward" | "back" | null,
  listed: readonly PresentationBlock[],
): S | "unchanged" | null {
  // On a slide holding a group the journal takes the steps that put the list
  // in the order `reorderBlocks` gives, which never parts a group: a step of
  // its own over one neighbouring object landed an object between a group's
  // members (Send Backward over an imported group, and the save refused it).
  // On a slide that also holds objects the author added (no source: none of
  // the file's shapes), the journal orders the file's objects among
  // themselves and the slide takes the list whole, so an added object stands
  // where the command put it; the journal alone refused such a slide and the
  // list changed without the canvas, the show or the save (z-order audit
  
  const withAuthored = slide.blocks.some((block) => (block as { source?: unknown }).source == null);
  const restacked =
    command === null ||
    withAuthored ||
    presentationReorderIsWithinGroup(slide.blocks, targets) ||
    slide.blocks.some((block) => blockGroupId(block))
      ? applyPresentationNativeObjectOrder(slide as never, listed as never, withAuthored)
    : applyPresentationNativeOrderCommand(
        slide as never,
        targets.map((block) => block.id),
        command,
      );
  if (restacked === null || restacked === "unchanged") return restacked;
  // The journal rebuilds the list from the slide's own blocks: the group's
  // Alt Text goes to the object now first in it, as `reorderBlocks` puts it.
  const restackedSlide = restacked as unknown as S;
  let blocks = [...restackedSlide.blocks] as PresentationBlock[];
  for (const groupId of new Set(targets.map((block) => blockGroupId(block)))) {
    blocks = withGroupAltOnFirstMember(blocks, groupId);
  }
  return { ...restackedSlide, blocks } as S;
}

/** Whether Arrange on `targets` restacks objects inside a group (each
 * among its own group's objects) rather than objects of the slide. */
export function presentationReorderIsWithinGroup(
  blocks: readonly PresentationBlock[],
  targets: readonly PresentationBlock[],
): boolean {
  return presentationReorderUnits(blocks, targets).inGroup.size > 0;
}

function reorderSlideObjects(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
  command: "front" | "forward" | "backward" | "back",
): PresentationBlock[] | null {
  const targetIds = new Set(targets.map((block) => block.id));
  if (command === "front") {
    return presentationRestackedPlaces([
      ...blocks.filter((block) => !targetIds.has(block.id)),
      ...targets,
    ]);
  }
  if (command === "back") {
    return presentationRestackedPlaces([
      ...targets,
      ...blocks.filter((block) => !targetIds.has(block.id)),
    ]);
  }
  const nextBlocks = [...blocks];
  let changed = false;
  // The neighbour an object steps over is a whole object: a group's
  // members stay together, so an object never lands inside the group
  // it passes (PPT-B205: that split the group and emptied the
  // Selection Pane).
  const unitEnd = (start: number, step: 1 | -1): number => {
    const groupId = blockGroupId(nextBlocks[start]);
    if (!groupId) return start;
    let end = start;
    while (
      end + step >= 0 &&
      end + step < nextBlocks.length &&
      !targetIds.has(nextBlocks[end + step].id) &&
      blockGroupId(nextBlocks[end + step]) === groupId
    ) {
      end += step;
    }
    return end;
  };
  if (command === "forward") {
    for (let index = nextBlocks.length - 2; index >= 0; index -= 1) {
      if (
        targetIds.has(nextBlocks[index].id) &&
        !targetIds.has(nextBlocks[index + 1].id)
      ) {
        const end = unitEnd(index + 1, 1);
        const [moved] = nextBlocks.splice(index, 1);
        nextBlocks.splice(end, 0, moved);
        changed = true;
      }
    }
  } else {
    for (let index = 1; index < nextBlocks.length; index += 1) {
      if (
        targetIds.has(nextBlocks[index].id) &&
        !targetIds.has(nextBlocks[index - 1].id)
      ) {
        const end = unitEnd(index - 1, -1);
        const [moved] = nextBlocks.splice(index, 1);
        nextBlocks.splice(end, 0, moved);
        changed = true;
      }
    }
  }
  return changed ? presentationRestackedPlaces(nextBlocks) : null;
}

/** Hide / Show (Selection Pane): one common state, hidden unless every target already is. */
export function toggleBlocksHidden(
  blocks: PresentationBlock[],
  targets: readonly PresentationBlock[],
): PresentationBlock[] {
  const targetIds = new Set(targets.map((block) => block.id));
  const nextHidden = !targets.every(isBlockHidden);
  return blocks.map((block) =>
    targetIds.has(block.id)
      ? applyBlockHidden(block, nextHidden)
      : block,
  );
}

/**
 * Home Delete: the selected objects leave the slide. A sound or video action
 * whose media leaves with them is cleared (a deck that kept it could not be
 * saved, PPT-B240), the animation steps of a deleted object go, and a slide
 * left empty keeps what `blocksLeftOnSlide` says: nothing on a Blank slide,
 * else one empty text box.
 */
export function slideWithoutBlocks<S extends PresentationObjectSlide>(
  slide: S,
  selectedIds: ReadonlySet<string>,
  makeId: () => string,
): S {
  const retained = slide.blocks
    .filter((block) => !selectedIds.has(block.id))
    .map((block) => {
      const action = normalizePresentationClickAction(
        (block as { clickAction?: unknown }).clickAction,
      );
      if (action?.kind !== "play-media" || !selectedIds.has(action.targetBlockId)) return block;
      const { clickAction: _gone, ...rest } = block as PresentationBlock & { clickAction?: unknown };
      void _gone;
      return rest as PresentationBlock;
    });
  const nextBlocks = blocksLeftOnSlide(
    { layout: (slide as { layout?: unknown }).layout },
    retained,
    makeId,
  );
  const nextSlide: S = { ...slide, blocks: nextBlocks };
  if (slide.animationSequence?.length) {
    const validIds = new Set(nextBlocks.map((block) => block.id));
    const pruned = pruneSequence(slide.animationSequence, validIds);
    if (pruned.length > 0) nextSlide.animationSequence = pruned;
    else delete nextSlide.animationSequence;
  }
  return nextSlide;
}

/**
 * Where Paste puts copied objects on a slide (EMU, the same distance right
 * and down), as PowerPoint places them: where they were copied from, unless
 * an object of the same frame already stands there -- the original, on the
 * slide it was copied from, or an earlier paste -- when the copies step
 * right and down past every such object (2% of the slide's width a step).
 * So a paste onto another slide keeps the position, a paste beside the
 * original is offset and each further paste steps on, and a cut object
 * pasted back onto its slide returns to its place.
 */
export function presentationPasteOffset(
  existing: ReadonlyArray<PresentationBlock>,
  source: ReadonlyArray<PresentationBlock>,
  bounds: { width: number; height: number },
): number {
  const step = Math.round(bounds.width * 0.02);
  const frame = (block: PresentationBlock) => {
    const record = block as Record<string, unknown>;
    const values = ["left", "top", "width", "height"].map((key) => record[key]);
    return values.every((value) => typeof value === "number" && Number.isFinite(value))
      ? (values as [number, number, number, number])
      : null;
  };
  const standing = existing.map(frame).filter((value): value is [number, number, number, number] => value !== null);
  const copies = source.map(frame).filter((value): value is [number, number, number, number] => value !== null);
  const near = (a: number, b: number) => Math.abs(a - b) <= 1;
  const taken = (dx: number) =>
    copies.some(([left, top, width, height]) =>
      standing.some(
        ([l, t, w, h]) => near(l, left + dx) && near(t, top + dx) && near(w, width) && near(h, height),
      ),
    );
  let steps = 0;
  while (steps < 64 && taken(steps * step)) steps += 1;
  return steps * step;
}

/**
 * Duplicate: deep clones of `source` appended to the slide, offset a little,
 * with fresh ids and group ids; each animation step of a copied original is
 * appended again for its clone (PPT-B029). `pasted` lists the clones.
 */
export function slideWithBlockClones<S extends PresentationObjectSlide>(
  slide: S,
  source: ReadonlyArray<PresentationBlock>,
  makeId: () => string,
  /** Duplicate's offset (EMU, PPT-E006); omitted for Paste, which places
   *  the copies as PowerPoint does (`presentationPasteOffset`). */
  offset?: { readonly dx: number; readonly dy: number },
): { slide: S; pasted: PresentationBlock[] } | null {
  if (
    source.length === 0 ||
    slide.blocks.length + source.length > MAX_PRESENTATION_SLIDE_BLOCKS
  ) {
    return null;
  }
  const clones = JSON.parse(JSON.stringify(source)) as PresentationBlock[];
  const bounds = slideBounds(slide);
  const pasteOffset = offset ? null : presentationPasteOffset(slide.blocks, source, bounds);
  const pasted = reidAndOffsetBlocks(
    clones as Array<Record<string, unknown> & { id: string }>,
    {
      // Duplicate's offset, or where Paste puts the copies (PowerPoint's
      // place: `presentationPasteOffset`).
      offsetEmu: offset ? offset.dx : pasteOffset!,
      ...(offset ? { offsetYEmu: offset.dy } : {}),
      maxLeft: bounds.width,
      maxTop: bounds.height,
      makeId,
      makeGroupId: newGroupId,
    },
  ) as unknown as PresentationBlock[];
  const nextSlide: S = {
    ...slide,
    blocks: [...slide.blocks, ...pasted],
  };
  // An object's animations travel with it (PPT-B029): each step of this
  // slide's sequence that animates a copied original is appended again for
  // its clone (a trigger that is itself copied follows to its clone), as
  // far as the sequence cap and the authoring gate allow.
  const cloneIdOf = new Map(
    source.map((block, index) => [block.id, pasted[index]?.id]),
  );
  const copiedSteps = (slide.animationSequence ?? []).flatMap(
    (step): TimelineStep[] => {
      const cloneId = cloneIdOf.get(step.blockId);
      const clone = cloneId
        ? pasted.find((block) => block.id === cloneId)
        : undefined;
      if (
        !clone ||
        !timelineAnimationAuthoringCapability({
          slideSource: slide.source,
          blockSource: (clone as { source?: unknown }).source,
          animation: step.animation,
          ...(step.motionPath
            ? { motionPathPreset: step.motionPath.preset }
            : {}),
        }).enabled
      ) {
        return [];
      }
      const copy: TimelineStep = {
        ...copyTimelineStep(step),
        id: newTimelineStepId(),
        blockId: clone.id,
      };
      const triggerClone = step.triggerBlockId
        ? cloneIdOf.get(step.triggerBlockId)
        : undefined;
      if (triggerClone) copy.triggerBlockId = triggerClone;
      return [copy];
    },
  );
  const room =
    MAX_TIMELINE_STEPS - (slide.animationSequence?.length ?? 0);
  if (copiedSteps.length > 0 && room > 0) {
    nextSlide.animationSequence = [
      ...(slide.animationSequence ?? []),
      ...copiedSteps.slice(0, room),
    ];
  }
  return { slide: nextSlide, pasted };
}

export type PresentationSelectionAuthority = Readonly<{
  disposition: "none" | "single" | "multiple" | "invalid";
  blocks: readonly PresentationBlock[];
  primary: PresentationBlock | null;
  selectionType: PresentationBlock["type"] | "mixed" | null;
}>;

/**
 * Resolve the editor's primary id and geometry multi-selection as one typed,
 * fail-closed authority. A selected id must resolve exactly once, and a primary
 * id may never sit outside the multi-selection it claims to represent.
 *
 * Non-positioned blocks intentionally use an empty `selectedBlockIds` set, so
 * their exact `selectedBlockId` remains a valid single selection. Consumers can
 * distinguish a real singleton from homogeneous or mixed multi-selection
 * without falling back to the first block on the slide.
 */
export function resolvePresentationSelectionAuthority(
  blocks: readonly PresentationBlock[],
  selectedBlockId: string | null,
  selectedBlockIds: ReadonlySet<string>,
): PresentationSelectionAuthority {
  const requestedIds =
    selectedBlockIds.size > 0
      ? new Set(selectedBlockIds)
      : selectedBlockId
        ? new Set([selectedBlockId])
        : new Set<string>();
  if (requestedIds.size === 0) {
    return {
      disposition: "none",
      blocks: [],
      primary: null,
      selectionType: null,
    };
  }
  if (!selectedBlockId || !requestedIds.has(selectedBlockId)) {
    return {
      disposition: "invalid",
      blocks: [],
      primary: null,
      selectionType: null,
    };
  }
  const selectedBlocks = blocks.filter((block) => requestedIds.has(block.id));
  if (
    selectedBlocks.length !== requestedIds.size ||
    new Set(selectedBlocks.map((block) => block.id)).size !== requestedIds.size
  ) {
    return {
      disposition: "invalid",
      blocks: [],
      primary: null,
      selectionType: null,
    };
  }
  const primary =
    selectedBlocks.find((block) => block.id === selectedBlockId) ?? null;
  if (!primary) {
    return {
      disposition: "invalid",
      blocks: [],
      primary: null,
      selectionType: null,
    };
  }
  const types = new Set(selectedBlocks.map((block) => block.type));
  return {
    disposition: selectedBlocks.length === 1 ? "single" : "multiple",
    blocks: selectedBlocks,
    primary,
    selectionType: types.size === 1 ? primary.type : "mixed",
  };
}

/** The object the Alt Text pane shows for a selection. */
export type PresentationAltTextPaneTarget = Readonly<{
  kind: "image" | "shape" | "chart" | "smartart";
  /** The block written: the picture, the shape, text box or chart, or the SmartArt
   * graphic's carrier member (which holds `groupAlt` / `groupDecorative`). */
  blockId: string;
  /** Names this object: a new key reseeds the pane's box, and a write whose
   * key no longer matches the live selection writes nothing. */
  key: string;
  description: string;
  decorative: boolean;
  /** Why its alt text cannot be changed here, or null. */
  readOnly: "locked" | "preserved" | null;
}>;

export type PresentationAltTextPaneResolution =
  | Readonly<{ status: "target"; target: PresentationAltTextPaneTarget }>
  | Readonly<{ status: "none" | "multiple" }>;

/**
 * What PowerPoint's Alt Text pane describes for the slide's selection: a
 * picture, a shape (a Cameo included), a text box or a chart selected on its own, or
 * the SmartArt graphic one or more of whose shapes are selected -- the deck
 * keeps one description per SmartArt graphic, on its carrier member, as the
 * PPTX keeps it on the graphic frame. A selection of several objects is
 * "multiple", and anything else (nothing selected, a table, a placeholder)
 * "none". A chart's description is its graphic frame's (`p:cNvPr`); one the
 * import could not read safely is the package's own ("preserved"). `historyNonce` is the pane's Undo / Redo counter,
 * part of the key (see `altTextPaneHistoryNonce`).
 */
export function resolvePresentationAltTextPaneTarget(
  slide: PresentationObjectSlide | undefined,
  selectedBlockId: string | null,
  selectedBlockIds: ReadonlySet<string>,
  historyNonce: number,
): PresentationAltTextPaneResolution {
  if (!slide) return { status: "none" };
  const selection = resolvePresentationSelectionAuthority(
    slide.blocks,
    selectedBlockId,
    selectedBlockIds,
  );
  const primary = selection.primary;
  if (
    !primary ||
    (selection.disposition !== "single" &&
      selection.disposition !== "multiple")
  ) {
    return { status: "none" };
  }
  const key = (kind: PresentationAltTextPaneTarget["kind"], id: string) =>
    `${kind}:${slide.id}:${id}:${historyNonce}`;
  const group = readSmartArtGroupMembers(slide.blocks, primary);
  const carrier = group?.members[0];
  if (group && carrier) {
    const memberIds = new Set(group.members.map((block) => block.id));
    if (selection.blocks.every((block) => memberIds.has(block.id))) {
      const record = carrier as PresentationBlock & {
        groupAlt?: unknown;
        groupDecorative?: unknown;
      };
      return {
        status: "target",
        target: {
          kind: "smartart",
          blockId: carrier.id,
          key: key("smartart", carrier.id),
          description:
            typeof record.groupAlt === "string" ? record.groupAlt : "",
          decorative: record.groupDecorative === true,
          // Lock fixes the graphic's place, never its Alt Text.
          readOnly:
            !isVerifiedAuthoredSmartArtGroup(group.members) && group.members.some(blockClaimsNativeGroup)
              ? "preserved"
              : null,
        },
      };
    }
  }
  if (selection.disposition !== "single") return { status: "multiple" };
  const isPicture = primary.type === "image";
  const isChart = primary.type === "chart";
  if (
    !isPicture &&
    !isChart &&
    primary.type !== "shape" &&
    !(primary.type === "text" && primary.kind === "textbox")
  ) {
    return { status: "none" };
  }
  const record = primary as PresentationBlock & {
    alt?: unknown;
    decorative?: unknown;
  };
  const source = asRecord((primary as Record<string, unknown>).source);
  const preserved = isPicture
    ? source?.pictureAltTextDisposition === "preserve-only"
    : isChart
      ? source?.chartAltTextDisposition === "preserve-only"
      : isPresentationCameoBlock(primary) &&
        source?.cameoAltTextDisposition === "preserve-only";
  const kind = isPicture ? "image" : isChart ? "chart" : "shape";
  return {
    status: "target",
    target: {
      kind,
      blockId: primary.id,
      key: key(kind, primary.id),
      description: typeof record.alt === "string" ? record.alt : "",
      decorative: record.decorative === true,
      // Lock fixes the object's place, never its Alt Text.
      readOnly: preserved ? "preserved" : null,
    },
  };
}

/**
 * A shape's or text box's Alt Text: a trimmed description, or decorative and
 * no description (the PPTX's `descr` and `adec:decorative`). The block itself
 * when nothing changes, so a keystroke that only adds trailing space commits
 * nothing.
 */
export function withPresentationShapeAltText(
  block: PresentationBlock,
  description: string,
  decorative: boolean,
): PresentationBlock {
  const record = block as PresentationBlock & {
    alt?: unknown;
    decorative?: unknown;
  };
  const text = description.trim();
  const nextAlt = decorative || !text ? undefined : text;
  const nextDecorative = decorative ? true : undefined;
  if (record.alt === nextAlt && record.decorative === nextDecorative) {
    return block;
  }
  const next = { ...block } as PresentationBlock & {
    alt?: string;
    decorative?: boolean;
  };
  if (nextAlt === undefined) delete next.alt;
  else next.alt = nextAlt;
  if (nextDecorative) next.decorative = true;
  else delete next.decorative;
  return next;
}

/**
 * A SmartArt graphic's Alt Text, kept on its carrier (first) member alone:
 * `groupAlt`, or `groupDecorative` and no description. Null when the carrier
 * no longer leads a group every member of which can change, or when nothing
 * changes.
 */
export function withPresentationSmartArtAltText(
  blocks: PresentationBlock[],
  carrierId: string,
  description: string,
  decorative: boolean,
): PresentationBlock[] | null {
  const group = readSmartArtGroupMembers(
    blocks,
    blocks.find((candidate) => candidate.id === carrierId),
  );
  if (
    !group ||
    group.members[0]?.id !== carrierId ||
    group.members.some((block) => !canEditBlockContent(block))
  ) {
    return null;
  }
  const text = description.trim();
  const memberIds = new Set(group.members.map((block) => block.id));
  let changed = false;
  const next = blocks.map((block) => {
    if (!memberIds.has(block.id)) return block;
    const record = block as PresentationBlock & {
      groupAlt?: unknown;
      groupDecorative?: unknown;
    };
    const isCarrier = block.id === carrierId;
    const wantAlt = isCarrier && !decorative && text ? text : undefined;
    const wantDecorative = isCarrier && decorative ? true : undefined;
    if (
      record.groupAlt === wantAlt &&
      record.groupDecorative === wantDecorative
    ) {
      return block;
    }
    changed = true;
    const copy = { ...block } as PresentationBlock & {
      groupAlt?: string;
      groupDecorative?: boolean;
    };
    delete copy.groupAlt;
    delete copy.groupDecorative;
    if (wantDecorative) copy.groupDecorative = true;
    else if (wantAlt) copy.groupAlt = wantAlt;
    return copy;
  });
  return changed ? next : null;
}

