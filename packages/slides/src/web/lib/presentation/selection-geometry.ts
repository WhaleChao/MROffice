/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Selection and gesture geometry of the presentation editor: SmartArt group
 * membership and bounds, which members a gesture carries (native groups,
 * imported groups, SmartArt), a block's CSS geometry and its fallback, the
 * Morph geometry of a block, and the small numeric helpers they share.
 *
 * Moved verbatim out of `components/notebook/contents/PresentationWindow.tsx`
 * (structural split, no behaviour change).
 */

import { smartArtMemberTextItem, smartArtSteppedMember } from "@mrnote/core/web/lib/office-smartart/interaction-model.ts";
import { canChangePresentationChartFrame } from "@mrnote/core/web/lib/presentation-chart-frame-operations.ts";
import type { CSSProperties } from "react";
import { isVerifiedAuthoredSmartArtGroup } from "@mrnote/core/web/lib/presentation-authored-smartart.ts";
import {
  readSmartArtGroupColors,
  readSmartArtGroupFrame,
  readSmartArtGroupMembers,
  type OfficeSmartArtFrame,
  type OfficeSmartArtGroup,
} from "@mrnote/core/web/lib/office-smartart/design-model.ts";
import {
  MAX_SMARTART_OUTLINE_TEXT_LENGTH,
  readSmartArtBlockMarker,
  type SmartArtStyle,
} from "@mrnote/core/web/lib/presentation-smartart.ts";
import type { SmartArtScene3d } from "@mrnote/core/web/lib/office-smartart/styles-model.ts";
import { smartArtCameraCss } from "@mrnote/core/web/lib/office-smartart/canvas-3d.ts";
import {
  smartArtEditIn2dKey,
  smartArtMemberHasScene,
  smartArtShapesLockedIn3d,
} from "@mrnote/core/web/lib/office-smartart/edit-in-2d.ts";
import { blockGeometry } from "@mrnote/core/web/lib/presentation-conveniences.ts";
import type { MorphBlockGeo, MorphGeo } from "../presentation-morph";
import {
  blockClaimsNativeGroup,
  blockFlag,
  exactNativeGroupMembers,
  geometryNumber,
  hasLayoutGeometry,
  isImportedPptxBlock,
} from "@mrnote/core/web/lib/presentation-object-buttons.ts";
import { buildBlockTransform, type BlockGeometry } from "@mrnote/core/web/lib/presentation-drag.ts";
import { blendsWithSlide } from "./metafile-blend-layer";
import { canManipulateBlock } from "@mrnote/core/web/lib/presentation-text-case.ts";
import {
  blockNativeGroupMoveAuthority,
  blockGroupId,
  sameNativeGroupMoveAuthority,
} from "@mrnote/core/web/lib/presentation-crop-group.ts";
import type {
  PresentationBlock,
} from "@mrnote/core/web/components/console/editor/native-document/documentCommands.ts";
import { selectionPaneRows } from "@mrnote/core/web/lib/presentation-selection-pane.ts";
import { asRecord } from "@mrnote/core/web/lib/presentation/deck-model.ts";
import {
  formatPresentationFrameMeasurement,
  type PresentationMeasurementUnit,
} from "@mrnote/core/web/lib/presentation-measurement.ts";
import { presentationShapeBlockTakesText } from "./text-editing-host";

/**
 * One member of a SmartArt diagram: its rect is the shared engine's own output
 * (`lib/office-smartart/engine`), down to the hairline connectors a hierarchy
 * or a timeline draws, so it is exact rather than authored-by-hand.
 */
export function isSmartArtMemberBlock(block: PresentationBlock): boolean {
  return (
    typeof (block as { smartArtRole?: unknown }).smartArtRole === "string" &&
    (block as { smartArtRole?: unknown }).smartArtRole !== ""
  );
}

export function selectionPaneIdsIdentifyOneLiveObject(
  blocks: readonly PresentationBlock[],
  explicitId: string | undefined,
  explicitIds: readonly string[] | undefined,
): boolean {
  if (!explicitIds) return true;
  const row = selectionPaneRows(blocks, () => "").find(
    (candidate) => candidate.block.id === explicitId,
  );
  return Boolean(
    row &&
      row.blockIds.length === explicitIds.length &&
      row.blockIds.every((blockId, index) => blockId === explicitIds[index]),
  );
}

/**
 * The presentation names for the SHARED diagram readers
 * (`lib/office-smartart/design-ops`). Nothing presentation-specific is left in
 * them: a SmartArt group is its members, its frame is their bounding box and
 * its colours are read off the members themselves, identically in every host.
 */
export type PresentationSmartArtGroup = OfficeSmartArtGroup<PresentationBlock>;

export function presentationSmartArtGroup(
  blocks: readonly PresentationBlock[],
  primary: PresentationBlock | null | undefined,
): PresentationSmartArtGroup | null {
  return readSmartArtGroupMembers(blocks, primary);
}

function presentationSmartArtStyle(
  group: PresentationSmartArtGroup,
): SmartArtStyle {
  return readSmartArtGroupColors(group.members);
}

export function presentationSmartArtBounds(
  group: PresentationSmartArtGroup,
): OfficeSmartArtFrame | null {
  return readSmartArtGroupFrame(group.members);
}

/**
 * The bounding box of every block sharing one `groupId`, in the block model's
 * EMU. The SmartArt diagram camera pivots about this box's centre, so each
 * member can compute the shared origin without the group being a real object.
 * Returns null when the id is empty or no member has usable geometry.
 */
export function presentationSmartArtGroupBox(
  blocks: readonly PresentationBlock[],
  groupId: string | null,
): { left: number; top: number; width: number; height: number } | null {
  if (!groupId) return null;
  const members = blocks.filter((block) => blockGroupId(block) === groupId);
  return members.length ? readSmartArtGroupFrame(members) : null;
}

/**
 * The diagram camera ONE SmartArt member needs, as inline style — the same
 * transform the editor stage builds, for the read-only surfaces that render
 * through {@link PresentationSlideView} (present mode and the slide-rail
 * miniature). Without it those surfaces drew a 3-D diagram with its style's
 * colours but none of its tilt, so a thumbnail tracked an edit to a node's text
 * while ignoring the style applied to it. Null for a block that is not a member
 * of a diagram carrying a scene camera (styles 3d5-3d9), which is every other
 * block on every other slide.
 */










function memberPictureCamera(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): SmartArtScene3d | undefined {
  if ((block as { type?: unknown }).type !== "image") return undefined;
  const groupId = blockGroupId(block);
  if (!groupId) return undefined;
  let camera: SmartArtScene3d | undefined;
  let key: string | undefined;
  for (const member of blocks) {
    if (member === block || blockGroupId(member) !== groupId) continue;
    const own = (member as { scene3d?: SmartArtScene3d }).scene3d;
    if (!own) continue;
    const ownKey = JSON.stringify(own);
    if (key !== undefined && ownKey !== key) return undefined;
    camera = own;
    key = ownKey;
  }
  return camera;
}

export function presentationSmartArtCameraStyle(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): {
  transform: string;
  transformOrigin: string;
  transformStyle: "preserve-3d";
} | null {
  // The camera is the diagram's own field: a diagram converted to shapes
  // keeps its members' `scene3d` and their group, and its view with them
  // (PPT-B258); a shape that never was a diagram carries none.
  const scene3d = (block as { scene3d?: SmartArtScene3d }).scene3d ?? memberPictureCamera(blocks, block);
  if (!scene3d) return null;
  const geom = blockGeometry(block);
  if (
    typeof geom.left !== "number" ||
    typeof geom.top !== "number" ||
    typeof geom.width !== "number" ||
    typeof geom.height !== "number"
  ) {
    return null;
  }
  const member = {
    left: geom.left,
    top: geom.top,
    width: geom.width,
    height: geom.height,
  };
  return smartArtCameraCss(
    scene3d,
    presentationSmartArtGroupBox(blocks, blockGroupId(block) ?? null) ?? member,
    member,
  );
}








function ownTurn(
  block: PresentationBlock,
  group: { width: number; height: number },
): { rotation?: number; flipH?: boolean; flipV?: boolean; aspect?: number } {
  const record = block as { rotation?: unknown; flipH?: unknown; flipV?: unknown };
  const rotation =
    typeof record.rotation === "number" && Number.isFinite(record.rotation) && record.rotation % 360 !== 0
      ? record.rotation
      : undefined;
  const flipH = record.flipH === true;
  const flipV = record.flipV === true;
  if (rotation === undefined && !flipH && !flipV) return {};
  return {
    ...(rotation !== undefined ? { rotation } : {}),
    ...(flipH ? { flipH } : {}),
    ...(flipV ? { flipV } : {}),
    ...(group.width > 0 && group.height > 0 ? { aspect: group.width / group.height } : {}),
  };
}

export function presentationSmartArtFaceSpan(
  blocks: readonly PresentationBlock[] | undefined,
  block: PresentationBlock,
): {
  top: number;
  bottom: number;
  left?: number;
  right?: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  aspect?: number;
} | null {
  if (!blocks) return null;
  const group = presentationSmartArtGroupBox(blocks, blockGroupId(block) ?? null);
  const geom = blockGeometry(block);
  if (!group || !(group.height > 0) || typeof geom.top !== "number" || typeof geom.height !== "number") return null;
  return {
    top: (geom.top - group.top) / group.height,
    bottom: (geom.top + geom.height - group.top) / group.height,
    // Across it too: 3d6's light comes from the right.
    ...(group.width > 0 && typeof geom.left === "number" && typeof geom.width === "number"
      ? {
          left: (geom.left - group.left) / group.width,
          right: (geom.left + geom.width - group.left) / group.width,
        }
      : {}),
    // The member's own turn and mirror, which its face is drawn under
    // (`SmartArtFaceSpan`): only where it has one, so an upright member's
    // span is what it was.
    ...ownTurn(block, group),
  };
}

/**
 * Whether one member of a SmartArt diagram may be moved / resized on its own
 * (after the user has entered the group). True for a freshly inserted diagram
 * and for a reopened one whose authored group verifies against its author
 * model; a preserve-only native group (a foreign deck's `p:grpSp`, or native
 * DiagramML) keeps the whole-group-only rule.
 */
function smartArtMemberEditableAlone(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): boolean {
  if (!hasLayoutGeometry(block)) return false;
  const group = presentationSmartArtGroup(blocks, block);
  if (!group) return false;
  if (
    group.members.some(
      (member) =>
        (member as Record<string, unknown>).nativeGroupDisposition ===
        "preserve-only",
    )
  ) {
    return false;
  }
  return (
    !group.members.some(blockClaimsNativeGroup) ||
    isVerifiedAuthoredSmartArtGroup(group.members)
  );
}

type NativeSmartArtTextAuthority = Readonly<{
  dataPart: string;
  dataFingerprint: string;
  points: readonly Readonly<{ modelId: string; text: string }>[];
}>;

export function blockNativeSmartArtTextAuthority(
  block: PresentationBlock | null | undefined,
): NativeSmartArtTextAuthority | null {
  const source = asRecord(block?.source);
  const authority = asRecord(source?.nativeSmartArtAuthority);
  const points = authority?.points;
  if (
    authority?.schemaVersion !== "pptx-smartart-text-v1" ||
    authority.disposition !== "native-text" ||
    typeof authority.dataPart !== "string" ||
    typeof authority.dataFingerprint !== "string" ||
    !/^[0-9a-f]{64}$/u.test(authority.dataFingerprint) ||
    !Array.isArray(points) ||
    points.length === 0 ||
    points.length > 20
  ) {
    return null;
  }
  const normalized = points.flatMap((point) => {
    const record = asRecord(point);
    return record &&
      typeof record.modelId === "string" &&
      /^[A-Za-z0-9{}_-]{1,128}$/u.test(record.modelId) &&
      typeof record.text === "string" &&
      record.text.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH
      ? [{ modelId: record.modelId, text: record.text }]
      : [];
  });
  if (
    normalized.length !== points.length ||
    new Set(normalized.map((point) => point.modelId)).size !== normalized.length
  ) {
    return null;
  }
  return {
    dataPart: authority.dataPart,
    dataFingerprint: authority.dataFingerprint,
    points: normalized,
  };
}

/**
 * Whether `block` belongs to an imported group, whose members the save only
 * accepts moved (or, when authenticated, transformed) together.
 */
export function blockClaimsImportedGroup(block: PresentationBlock): boolean {
  return (
    blockClaimsNativeGroup(block) ||
    Boolean(
      blockGroupId(block) &&
        (isImportedPptxBlock(block) ||
          blockNativeGroupMoveAuthority(
            block as unknown as Record<string, unknown>,
          ) !== null),
    )
  );
}

/**
 * Every block of an imported group the native group authority leaves
 * preserve-only (nested groups, a member in mc:AlternateContent, a lone
 * drawable member) when each one carries the SAME move authority
 * (`source.nativeGroupMove`). Such a group moves whole and only moves: the
 * save writes one translation of its own a:off, and refuses anything else done
 * to it, a member moved alone included. Null when any block sharing the
 * group id lacks that authority, names another group or offset, claims a
 * native group of its own, or has no slide geometry.
 */
export function nativeGroupMoveMembers(
  blocks: readonly PresentationBlock[],
  selected: PresentationBlock,
): PresentationBlock[] | null {
  const authority = blockNativeGroupMoveAuthority(
    selected as unknown as Record<string, unknown>,
  );
  const groupId = blockGroupId(selected);
  if (!authority || !groupId) return null;
  const members = blocks.filter((block) => blockGroupId(block) === groupId);
  return members.every((member) => {
    const candidate = blockNativeGroupMoveAuthority(
      member as unknown as Record<string, unknown>,
    );
    return (
      candidate !== null &&
      sameNativeGroupMoveAuthority(candidate, authority) &&
      !blockClaimsNativeGroup(member) &&
      hasLayoutGeometry(member)
    );
  })
    ? members
    : null;
}

/**
 * The graphic `block` belongs to wears a 3-D scene style and is not in Edit in
 * 2-D: its shapes are not moved, resized or rotated one at a time (Office).
 */
export function presentationSmartArtShapesLockedIn3d(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): boolean {
  const marker = readSmartArtBlockMarker(block);
  if (!marker) return false;
  const group = presentationSmartArtGroup(blocks, block);
  return smartArtShapesLockedIn3d(
    smartArtEditIn2dKey("presentation", marker.groupId),
    smartArtMemberHasScene(block) || Boolean(group?.members.some(smartArtMemberHasScene)),
  );
}

/**
 * Which blocks one move / transform gesture on `block` carries as ONE object:
 * the rule the canvas drag, the diagram frame and the arrow-key nudge share.
 *
 * - `scope === "group"` (the SmartArt diagram frame): every member.
 * - A SmartArt member editable alone: its outline ITEM (the node with its own
 *   label and picture), or null when it stands alone.
 * - An imported group: every member of an authenticated native group, or
 *   every block of a group that carries one move authority (moved, sized or
 *   turned whole; the save writes the group's own frame). An imported group that is neither is refused: a member moved
 *   alone leaves its slide unsavable.
 *
 * `members` is null when the block is not in such a set (it moves alone, or
 * by the persistent-group / selection rules of the drag). Refused as well when
 * any member is locked or its chart frame cannot move, and -- `sceneLocked` --
 * for one shape of a graphic in a 3-D scene style until Edit in 2-D
 * (`lib/office-smartart/edit-in-2d`); its frame still moves the whole graphic.
 */
export function presentationGestureMembers(
  blocks: PresentationBlock[],
  block: PresentationBlock,
  scope: "member" | "group",
  operation: "move" | "transform",
):
  | { refused: true; sceneLocked?: true }
  | {
      refused: false;
      smartArtSolo: boolean;
      members: PresentationBlock[] | null;
    } {
  const smartArtGroup = presentationSmartArtGroup(blocks, block);
  if (scope === "member" && presentationSmartArtShapesLockedIn3d(blocks, block)) {
    return { refused: true, sceneLocked: true };
  }
  const smartArtSolo =
    scope === "member" &&
    smartArtGroup !== null &&
    smartArtMemberEditableAlone(blocks, block);
  const exactMembers = smartArtSolo ? null : exactNativeGroupMembers(blocks, block);
  // A group the file nests (move authority) also sizes and turns whole: the
  // save scales its own `a:ext` or turns its `rot`, child space untouched
  // (`_pptx_native_group_reframe_patch`).
  const moveOnlyMembers =
    smartArtSolo || exactMembers
      ? null
      : nativeGroupMoveMembers(blocks, block);
  if (
    !smartArtSolo &&
    blockClaimsImportedGroup(block) &&
    !exactMembers &&
    !moveOnlyMembers
  ) {
    return { refused: true };
  }
  const smartArtItemBlocks = (() => {
    if (!smartArtSolo || !smartArtGroup) return null;
    const itemId = readSmartArtBlockMarker(block)?.itemId;
    if (!itemId) return null;
    const companions = smartArtGroup.members.filter(
      (member) =>
        member.id !== block.id &&
        hasLayoutGeometry(member) &&
        readSmartArtBlockMarker(member)?.itemId === itemId &&
        canManipulateBlock(member),
    );
    return companions.length > 0 ? [block, ...companions] : null;
  })();
  const members =
    scope === "group" && smartArtGroup
      ? smartArtGroup.members.filter(hasLayoutGeometry)
      : smartArtItemBlocks ?? exactMembers ?? moveOnlyMembers;
  if (
    members?.some(
      (member) =>
        !canManipulateBlock(member) ||
        !canChangePresentationChartFrame(member, "move"),
    )
  ) {
    return { refused: true };
  }
  return { refused: false, smartArtSolo, members };
}

/**
 * The SHAPE whose words a SmartArt label is (PowerPoint: "a shape's text is
 * the shape's own text"). The deck draws an item's words as a separate
 * text strip laid over its node; the strip is never an object of its own:
 * a click on it selects the node, and the node wears the selection frame,
 * the sizing handles and the rotation handle. `null` for anything else --
 * a block that is no label, a label of a package-owned (native) graphic, a
 * caption laid beside its picture rather than on a shape.
 */
export function presentationSmartArtLabelShape(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock | null | undefined,
): PresentationBlock | null {
  if (!block || block.type !== "text") return null;
  const marker = readSmartArtBlockMarker(block);
  if (marker?.role !== "label" || !marker.itemId) return null;
  if (!smartArtMemberEditableAlone(blocks, block)) return null;
  const left = geometryNumber(block, "left");
  const top = geometryNumber(block, "top");
  const width = geometryNumber(block, "width");
  const height = geometryNumber(block, "height");
  if (left === null || top === null || width === null || height === null) return null;
  const centerX = left + width / 2;
  const centerY = top + height / 2;
  return (
    blocks.find((candidate) => {
      if (candidate.id === block.id || candidate.type !== "shape") return false;
      const other = readSmartArtBlockMarker(candidate);
      if (
        other?.role !== "node" ||
        other.itemId !== marker.itemId ||
        other.groupId !== marker.groupId
      ) {
        return false;
      }
      const nodeLeft = geometryNumber(candidate, "left");
      const nodeTop = geometryNumber(candidate, "top");
      const nodeWidth = geometryNumber(candidate, "width");
      const nodeHeight = geometryNumber(candidate, "height");
      return (
        nodeLeft !== null &&
        nodeTop !== null &&
        nodeWidth !== null &&
        nodeHeight !== null &&
        centerX >= nodeLeft - 1 &&
        centerX <= nodeLeft + nodeWidth + 1 &&
        centerY >= nodeTop - 1 &&
        centerY <= nodeTop + nodeHeight + 1
      );
    }) ?? null
  );
}

/**
 * The SmartArt shape Tab (offset 1) or Shift+Tab (-1) moves to inside an
 * activated graphic (Microsoft Support, "Keyboard shortcuts for SmartArt
 * graphics": after Enter activates a graphic, Tab and Shift+Tab select its
 * next and previous elements; Esc returns to the top level): the item shapes
 * in the outline's order, wrapping round. A label stands for its shape, a
 * connecting shape for the item before it. Null when the block is no member
 * of an authored graphic, or the graphic has no item shape. Offset 0 is
 * Enter on the whole graphic, which activates it on its first shape.
 */
export function presentationSmartArtSteppedShape(
  blocks: readonly PresentationBlock[],
  from: PresentationBlock | null | undefined,
  offset: 1 | -1 | 0,
): PresentationBlock | null {
  // The rule the three editors share (`smartArtSteppedMember`).
  return smartArtSteppedMember(blocks, from, offset);
}

/** The label that holds a SmartArt shape's words ({@link presentationSmartArtLabelShape}'s inverse). */
export function presentationSmartArtShapeLabel(
  blocks: readonly PresentationBlock[],
  shape: PresentationBlock | null | undefined,
): PresentationBlock | null {
  if (!shape || shape.type !== "shape") return null;
  const marker = readSmartArtBlockMarker(shape);
  // An item's title box (Name and Title) opens that title's words.
  const titled = (shape as { smartArtTitleIndex?: unknown }).smartArtTitleIndex;
  if (marker?.role === "connector" && typeof titled === "number") {
    return (
      blocks.find(
        (candidate) =>
          candidate.type === "text" &&
          candidate.groupId === shape.groupId &&
          (candidate as { smartArtTitleIndex?: unknown }).smartArtTitleIndex === titled,
      ) ?? null
    );
  }
  // The item whose words a double click on this shape opens, as the three
  // editors share it (`smartArtMemberTextItem`): a node's, never a
  // connecting shape's or a picture placeholder's.
  const itemId =
    marker?.role === "node"
      ? smartArtMemberTextItem({
          smartArtRole: marker.role,
          smartArtItemId: marker.itemId,
          smartArtPicturePlaceholder: (shape as { smartArtPicturePlaceholder?: unknown }).smartArtPicturePlaceholder,
        })
      : null;
  if (!itemId) return null;
  return (
    blocks.find(
      (candidate) =>
        candidate.type === "text" &&
        readSmartArtBlockMarker(candidate)?.itemId === itemId &&
        presentationSmartArtLabelShape(blocks, candidate)?.id === shape.id,
    ) ?? null
  );
}

/**
 * Where a shape's words are typed: its own text, the label of its SmartArt
 * item (an item's words are its label's), or nowhere. A SmartArt graphic's
 * connecting arrow holds no words (Office's text pane has no line for it), so
 * a click on it selects and moves it like any shape and never opens a caret;
 * a freeform item shape (a Pyramid's layer) still takes its item's words.
 */
export function presentationShapeTextHost(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): "own" | "label" | null {
  if (block.type !== "shape") return null;
  const marker = readSmartArtBlockMarker(block);
  if (marker) {
    if (presentationSmartArtShapeLabel(blocks, block)) return "label";
    if (marker.role === "connector") return null;
  }
  return presentationShapeBlockTakesText(block) ? "own" : null;
}

/**
 * What one rotation of a SmartArt shape turns (Microsoft Support, "Rotate or
 * flip a SmartArt graphic shape": each shape "is rotated around its own
 * center"): the pressed item -- its node with its own label and picture --
 * about the node's centre, never the rest of the graphic. `null` when the
 * block is no member of an editable graphic; `locked` when one of the
 * item's blocks cannot turn (the gesture is then refused).
 */
export function presentationSmartArtItemRotation(
  blocks: readonly PresentationBlock[],
  block: PresentationBlock,
): {
  pivot: { x: number; y: number };
  locked: boolean;
  members: Array<{
    blockId: string;
    startGeom: BlockGeometry;
    startRotation: number;
  }>;
} | null {
  if (!presentationSmartArtGroup(blocks, block)) return null;
  if (!smartArtMemberEditableAlone(blocks, block)) return null;
  const itemId = readSmartArtBlockMarker(block)?.itemId;
  const item = itemId
    ? blocks.filter(
        (member) =>
          hasLayoutGeometry(member) &&
          readSmartArtBlockMarker(member)?.groupId ===
            readSmartArtBlockMarker(block)?.groupId &&
          readSmartArtBlockMarker(member)?.itemId === itemId,
      )
    : [block];
  if (!item.some((member) => member.id === block.id)) item.unshift(block);
  const locked = item.some((member) => !canManipulateBlock(member));
  const rectOf = (member: PresentationBlock): BlockGeometry | null => {
    const left = geometryNumber(member, "left");
    const top = geometryNumber(member, "top");
    const width = geometryNumber(member, "width");
    const height = geometryNumber(member, "height");
    return left === null || top === null || width === null || height === null
      ? null
      : { left, top, width, height };
  };
  const members = item.flatMap((member) => {
    const startGeom = rectOf(member);
    return startGeom
      ? [
          {
            blockId: member.id,
            startGeom,
            startRotation: geometryNumber(member, "rotation") ?? 0,
          },
        ]
      : [];
  });
  const shape =
    item.find(
      (member) =>
        member.type === "shape" &&
        readSmartArtBlockMarker(member)?.role === "node",
    ) ??
    presentationSmartArtLabelShape(blocks, block) ??
    block;
  const pivotRect = rectOf(shape) ?? rectOf(block);
  if (!pivotRect || members.length === 0) return null;
  return {
    pivot: {
      x: pivotRect.left + pivotRect.width / 2,
      y: pivotRect.top + pivotRect.height / 2,
    },
    locked,
    members,
  };
}

/** A block's rect for a gesture, filling a missing size as 35% / 18% of the slide. */
export function gestureStartGeom(
  block: PresentationBlock,
  slideBounds: { width: number; height: number },
): BlockGeometry {
  return {
    left: geometryNumber(block, "left") ?? 0,
    top: geometryNumber(block, "top") ?? 0,
    width:
      geometryNumber(block, "width") ?? Math.round(slideBounds.width * 0.35),
    height:
      geometryNumber(block, "height") ?? Math.round(slideBounds.height * 0.18),
  };
}

/** Each follower's start rect translated by the one group delta. */
export function translatedMemberGeoms(
  members: ReadonlyArray<{ blockId: string; startGeom: BlockGeometry }>,
  dx: number,
  dy: number,
): Map<string, BlockGeometry> {
  return new Map(
    members.map((member) => [
      member.blockId,
      {
        ...member.startGeom,
        left: member.startGeom.left + dx,
        top: member.startGeom.top + dy,
      },
    ]),
  );
}

export function hasGeometry(block: PresentationBlock): boolean {
  return (
    hasLayoutGeometry(block) ||
    geometryNumber(block, "left") !== null ||
    geometryNumber(block, "top") !== null ||
    geometryNumber(block, "rotation") !== null
  );
}

/**
 * A rendering-only geometry for a block that has none, on a slide that is
 * otherwise a CANVAS.
 *
 * A slide is a canvas the moment ANY block carries left/top, but each block also
 * decided for itself: one without geometry fell back to the stacked list editor.
 * A slide holding both — which is what you get the instant you drop a text box
 * onto a layout whose placeholders never had geometry — rendered the list editor
 * and the canvas AT ONCE, one above the other, with the same slide's content
 * split between them. Reported from real use as "clicking the lower half enters
 * a strange view".
 *
 * The block model is not touched. This only decides where to draw something the
 * deck never placed, so a title sits at the top and everything else stacks below
 * it in source order, inside the slide's margins.
 */
export function canvasFallbackGeometry(
  block: PresentationBlock,
  index: number,
  bounds: { width: number; height: number },
): { left: number; top: number; width: number; height: number } {
  const margin = Math.round(bounds.width * 0.06);
  const usableWidth = Math.max(bounds.width - margin * 2, 1);
  const role = (block as Record<string, unknown>).role;
  if (role === "title") {
    return {
      left: margin,
      top: margin,
      width: usableWidth,
      height: Math.round(bounds.height * 0.16),
    };
  }
  const bodyTop = margin + Math.round(bounds.height * 0.2);
  const rowHeight = Math.round(bounds.height * 0.16);
  const gap = Math.round(bounds.height * 0.02);
  return {
    left: margin,
    top: Math.min(
      bodyTop + index * (rowHeight + gap),
      Math.max(bounds.height - rowHeight - margin, margin),
    ),
    width: usableWidth,
    height: rowHeight,
  };
}

export function blockGeometryStyle(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  // Optional live override (drag/resize preview): when provided, position/size
  // come from this rect instead of the persisted block geometry while keeping
  // the block's rotation / z-index. Lets the canvas render an uncommitted
  // gesture without mutating slide state.
  override?: { left: number; top: number; width: number; height: number },
  // Optional live rotation override (rotate-handle preview): when provided it
  // replaces the persisted `rotation` so the block visibly rotates during the
  // drag without mutating slide state. Composes with the persisted flip flags.
  rotationOverride?: number,
  // Optional live mirror override (a resize handle dragged through the pinned
  // edge previews the mirrored block before the gesture commits).
  flipOverride?: { flipH: boolean; flipV: boolean },
): CSSProperties | undefined {
  if (!hasLayoutGeometry(block)) return undefined;
  const left = override?.left ?? geometryNumber(block, "left") ?? 0;
  const top = override?.top ?? geometryNumber(block, "top") ?? 0;
  const width =
    override?.width ?? geometryNumber(block, "width") ?? bounds.width * 0.35;
  const height =
    override?.height ?? geometryNumber(block, "height") ?? bounds.height * 0.18;
  const rotation = rotationOverride ?? geometryNumber(block, "rotation") ?? 0;
  const widthPercent = (width / bounds.width) * 100;
  const heightPercent = (height / bounds.height) * 100;
  // The 4%/5% floors keep a block an author dropped without a size clickable.
  // A block whose geometry was COMPUTED is a different thing: an imported
  // native shape carries PowerPoint's own rect, and a SmartArt member carries
  // the one the shared engine laid out. Rounding those up redraws the diagram —
  // a hierarchy's 0.18%-wide connector became a 4% grey slab across the chart,
  // and a timeline's axis a 5% band — so they keep their exact rect.
  // A shape or picture the author sized is the same: its Height box says
  // 0.2 cm, so it draws 0.2 cm (PPT-B044: every height under 5% of the slide
  // drew at 5%, so 0.5 cm and 0.2 cm looked identical). A text frame with a
  // stated size is the same: PowerPoint draws it at that size and its words
  // overflow it (DrawingML's default), as the PDF prints them. Floored, two
  // small boxes a design or Ensure Fit left one above the other (a portrait
  // slide's 5% is 1.7 cm) grew into each other and the title's bottom-anchored
  
  const authorSized =
    (block.type === "shape" || block.type === "image" || block.type === "text") &&
    (override !== undefined ||
      (geometryNumber(block, "width") !== null &&
        geometryNumber(block, "height") !== null));
  const exactGeometry =
    authorSized || isImportedPptxBlock(block) || isSmartArtMemberBlock(block);
  const transform = buildBlockTransform(rotation, flipOverride ?? {
    flipH: blockFlag(block, "flipH"),
    flipV: blockFlag(block, "flipV"),
  });
  return {
    left: `${(left / bounds.width) * 100}%`,
    top: `${(top / bounds.height) * 100}%`,
    width: `${exactGeometry && widthPercent > 0 ? widthPercent : Math.max(4, widthPercent)}%`,
    height: `${exactGeometry && heightPercent > 0 ? heightPercent : Math.max(5, heightPercent)}%`,
    position: "absolute",
    // Rotation + mirror flags compose into ONE transform (flip applied first,
    // then rotate — PowerPoint xfrm order; see buildBlockTransform). The flip is
    // render-only: it never touches the block's geometry/text. `center center`
    // origin makes flip-twice an exact identity.
    transform,
    transformOrigin: "center center",
    // Every object is a stacking layer of its own, drawn in the slide's paint
    // order (`presentationInPaintOrder`): what it holds inside -- a text
    // preview's layers, its grips -- stays with it, and a number it carries
    // never outranks an object in front of it (PPT-B310). A picture with a
    // blend layer (readings row 1021) is not: the layer blends with the slide
    // under it, which a stacking context of its own would hide from it; it
    // is still painted in its place (positioned, in tree order).
    ...(blendsWithSlide(block, transform) ? {} : { zIndex: 0 }),
  };
}

/**
 * Resolve a block's animatable geometry (EMU position/size + degree rotation) for
 * the MORPH planner. Missing size falls back to the SAME proportional defaults
 * `blockGeometryStyle` uses (35% width / 18% height of the slide) so a morph tween
 * lands a matched block exactly where the static render would place it; missing
 * position defaults to the slide origin and missing rotation to 0. Pure read — it
 * never mutates the block.
 */
export function morphGeoOfBlock(
  block: PresentationBlock,
  bounds: { width: number; height: number },
): MorphGeo {
  return {
    left: geometryNumber(block, "left") ?? 0,
    top: geometryNumber(block, "top") ?? 0,
    width: geometryNumber(block, "width") ?? bounds.width * 0.35,
    height: geometryNumber(block, "height") ?? bounds.height * 0.18,
    rotation: geometryNumber(block, "rotation") ?? 0,
  };
}

/** A block as the Morph planner recognises it on the next slide. */
export function morphBlockGeoOf(
  block: PresentationBlock,
  bounds: { width: number; height: number },
): MorphBlockGeo {
  const record = block as Record<string, unknown>;
  return {
    id: block.id,
    geo: morphGeoOfBlock(block, bounds),
    type: block.type,
    ...(typeof record.morphSourceId === "string" && record.morphSourceId
      ? { sourceId: record.morphSourceId }
      : {}),
    ...(typeof record.name === "string" ? { name: record.name } : {}),
  };
}

/**
 * The morph-layer positioning style for one interpolated geometry, mirroring the
 * percentage math + clamps in `blockGeometryStyle` so a matched block tweens along
 * the exact rect the static slide would render. `zIndex` carries the destination
 * block's stacking so the morph layer paints in the settled z-order.
 */
export function morphGeoStyle(
  geo: MorphGeo,
  bounds: { width: number; height: number },
  zIndex?: number,
): CSSProperties {
  return {
    position: "absolute",
    left: `${(geo.left / bounds.width) * 100}%`,
    top: `${(geo.top / bounds.height) * 100}%`,
    width: `${Math.max(4, (geo.width / bounds.width) * 100)}%`,
    height: `${Math.max(5, (geo.height / bounds.height) * 100)}%`,
    transform: geo.rotation ? `rotate(${geo.rotation}deg)` : undefined,
    transformOrigin: "center center",
    zIndex: zIndex !== undefined ? Math.max(0, Math.round(zIndex)) : undefined,
  };
}

/**
 * Where the Morph layer puts one object's frame: the static slide's own rect
 * (`blockGeometryStyle`: an author-sized shape or picture at its exact size, a
 * text frame with its floor, the mirror flags) at the object's place, or at
 * `geo` -- the previous slide's place, with that object's flips -- where a
 * matched object starts its tween. The layer used to floor every frame at
 * 4% x 5% of the slide and drop the flips, so a thin accent bar drew as a
 * thick block and a mirrored arrow turned back while the transition played.
 * An object with no layout geometry keeps the proportional default rect.
 */
export function morphFrameStyle(
  block: PresentationBlock,
  bounds: { width: number; height: number },
  geo?: MorphGeo,
  flipsOf?: PresentationBlock,
): CSSProperties {
  const style = geo
    ? blockGeometryStyle(
        block,
        bounds,
        { left: geo.left, top: geo.top, width: geo.width, height: geo.height },
        geo.rotation,
        flipsOf
          ? { flipH: blockFlag(flipsOf, "flipH"), flipV: blockFlag(flipsOf, "flipV") }
          : undefined,
      )
    : blockGeometryStyle(block, bounds);
  return style ?? morphGeoStyle(geo ?? morphGeoOfBlock(block, bounds), bounds);
}








export function geometryInputValue(
  block: PresentationBlock,
  key: string,
  unit: PresentationMeasurementUnit,
): string {
  const value = geometryNumber(block, key);
  if (key === "rotation") return String(Math.round(value ?? 0));
  if (value === null) return "";
  return formatPresentationFrameMeasurement(value, unit);
}

export function clampNumber(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}

/** How far from the slide's corner, in its own sizes, a stroke is still placed by a translate (see below). */
export const INK_PLACEMENT_MAX_OWN_SIZES = 32;













export function presentationInkBlockPlacement(style: CSSProperties | undefined): CSSProperties | undefined {
  if (!style) return style;
  const share = (value: unknown) =>
    typeof value === "string" && value.endsWith("%") ? Number.parseFloat(value) : Number.NaN;
  const left = share(style.left);
  const top = share(style.top);
  const width = share(style.width);
  const height = share(style.height);
  if (![left, top, width, height].every(Number.isFinite) || width < 1e-6 || height < 1e-6) return style;
  // A translate in shares of the box's own size is taken against the size
  // layout gave the box, snapped to Chromium's 1/64 px layout unit: the error
  // is that snap times the share. A stroke many times narrower than its
  
  // C=O bond (0.7 px wide, 860 of its widths from the corner) -- moved by up to
  // 13 px (TRIAGE 149). Past 32 of its own sizes (half a pixel at most) the
  // stroke keeps its left/top, a whole-pixel snap at worst.
  if (Math.abs(left / width) > INK_PLACEMENT_MAX_OWN_SIZES || Math.abs(top / height) > INK_PLACEMENT_MAX_OWN_SIZES) {
    return style;
  }
  const own = typeof style.transform === "string" && style.transform !== "none" ? ` ${style.transform}` : "";
  return {
    ...style,
    left: 0,
    top: 0,
    transform: `translate(${(left / width) * 100}%, ${(top / height) * 100}%)${own}`,
  };
}
