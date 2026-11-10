/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Multi-selection / group transform math — pure, EMU-space, model-shaped.
 *
 * Behaviour reference: Excalidraw's multi-element resize
 * (github.com/excalidraw/excalidraw @ b2e81e38a6fd,
 * packages/element/src/resizeElements.ts, MIT). No code is copied — the
 * geometry here is written against MRAI's own axis-aligned `BlockGeometry`
 * model — but one behavioural idea is deliberately taken from that file and
 * deserves attribution: when a proportional scale would push ANY member below
 * the minimum size, the SCALE ITSELF is floored (max over members of
 * min/size), so the whole selection stays rigidly uniform instead of one
 * member clamping out of proportion.
 *
 * PowerPoint conventions win where the two references disagree: corner
 * handles free-scale both axes (Shift maintains aspect), edge handles stretch
 * one axis — matching this editor's single-block `applyResize` so a
 * multi-selection feels like a bigger single block, not a different tool.
 */

import {
  DEFAULT_MIN_BLOCK_SIZE,
  type BlockGeometry,
  type DragBounds,
  type ResizeHandle,
} from "./presentation-drag";
import { officeTurnAround } from "./office-object-frame";

export interface TransformMember {
  blockId: string;
  startGeom: BlockGeometry;
}

export interface RotatableTransformMember extends TransformMember {
  startRotation: number;
}

export interface RotatedTransformGeometry extends BlockGeometry {
  rotation: number;
}

/**
 * The axis-aligned box a block's frame covers once rotated about its centre
 * by `rotationDegrees`: what the user sees, which Align, the marquee and the
 * eraser must measure against (PPT-B219, B208, B235).
 */
export function rotatedBlockBounds(
  geometry: BlockGeometry,
  rotationDegrees: number | null | undefined,
): BlockGeometry {
  const rotation = Number.isFinite(rotationDegrees) ? (rotationDegrees as number) % 360 : 0;
  if (!rotation) return geometry;
  const radians = (rotation * Math.PI) / 180;
  const cosine = Math.abs(Math.cos(radians));
  const sine = Math.abs(Math.sin(radians));
  const width = geometry.width * cosine + geometry.height * sine;
  const height = geometry.width * sine + geometry.height * cosine;
  const centerX = geometry.left + geometry.width / 2;
  const centerY = geometry.top + geometry.height / 2;
  return { left: centerX - width / 2, top: centerY - height / 2, width, height };
}

/** The tight axis-aligned bounding box of the members' start geometry. */
export function selectionBoundingBox(
  members: ReadonlyArray<TransformMember>,
): BlockGeometry | null {
  if (members.length === 0) return null;
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const { startGeom } of members) {
    left = Math.min(left, startGeom.left);
    top = Math.min(top, startGeom.top);
    right = Math.max(right, startGeom.left + startGeom.width);
    bottom = Math.max(bottom, startGeom.top + startGeom.height);
  }
  if (!Number.isFinite(left) || right <= left || bottom <= top) return null;
  return { left, top, width: right - left, height: bottom - top };
}

/**
 * Move every member by one delta, clamped so the WHOLE selection stays
 * on-slide — the set-level form of `applyMove`'s single-block invariant. The
 * delta is clamped once against the bounding box, so members keep their exact
 * relative offsets (no per-member squashing at the edges).
 */
export function applyMultiMove(
  members: ReadonlyArray<TransformMember>,
  dx: number,
  dy: number,
  bounds: DragBounds,
): Map<string, BlockGeometry> {
  const result = new Map<string, BlockGeometry>();
  const bbox = selectionBoundingBox(members);
  if (!bbox) return result;
  const clampedDx = Math.min(
    Math.max(dx, -bbox.left),
    Math.max(0, bounds.width - (bbox.left + bbox.width)),
  );
  const clampedDy = Math.min(
    Math.max(dy, -bbox.top),
    Math.max(0, bounds.height - (bbox.top + bbox.height)),
  );
  for (const { blockId, startGeom } of members) {
    result.set(blockId, {
      ...startGeom,
      left: Math.round(startGeom.left + clampedDx),
      top: Math.round(startGeom.top + clampedDy),
    });
  }
  return result;
}

/** Rotate a rigid selection around its shared centre without changing sizes. */
export function applyMultiRotate(
  members: ReadonlyArray<RotatableTransformMember>,
  deltaDegrees: number,
  // The pivot, when it is not the members' bounding-box centre: a SmartArt
  // shape turns about ITS OWN centre and carries its words with it, even
  // where the words' strip reaches past the shape.
  pivot?: { x: number; y: number },
): Map<string, RotatedTransformGeometry> {
  const result = new Map<string, RotatedTransformGeometry>();
  const bbox = selectionBoundingBox(members);
  if (!bbox || !Number.isFinite(deltaDegrees)) return result;
  const centerX =
    pivot && Number.isFinite(pivot.x) ? pivot.x : bbox.left + bbox.width / 2;
  const centerY =
    pivot && Number.isFinite(pivot.y) ? pivot.y : bbox.top + bbox.height / 2;
  for (const member of members) {
    const { x: rotatedCenterX, y: rotatedCenterY } = officeTurnAround(
      { x: member.startGeom.left + member.startGeom.width / 2, y: member.startGeom.top + member.startGeom.height / 2 },
      { x: centerX, y: centerY },
      deltaDegrees,
    );
    result.set(member.blockId, {
      left: Math.round(rotatedCenterX - member.startGeom.width / 2),
      top: Math.round(rotatedCenterY - member.startGeom.height / 2),
      width: member.startGeom.width,
      height: member.startGeom.height,
      rotation: ((member.startRotation + deltaDegrees) % 360 + 360) % 360,
    });
  }
  return result;
}

/** Which axes a handle drives, and the anchor edge it scales away from. */
function handleAxes(handle: ResizeHandle): {
  scalesX: boolean;
  scalesY: boolean;
  fromLeft: boolean;
  fromTop: boolean;
} {
  return {
    scalesX: handle.includes("e") || handle.includes("w"),
    scalesY: handle.includes("n") || handle.includes("s"),
    fromLeft: handle.includes("w"),
    fromTop: handle.includes("n"),
  };
}

export interface MultiResizeOptions {
  /** Shift: corner handles keep the selection's aspect ratio. */
  maintainAspect?: boolean;
  minSize?: number;
}

/**
 * Scale every member around the bounding box's anchor (the corner/edge
 * opposite the dragged handle). Pointer deltas are in EMU. The scale is
 * floored so no member drops below `minSize` on a driven axis, and capped so
 * the scaled bounding box stays on-slide; both caps apply to the SCALE, never
 * to one member, so the selection stays uniform.
 */
export function applyMultiResize(
  members: ReadonlyArray<TransformMember>,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  bounds: DragBounds,
  options: MultiResizeOptions = {},
): Map<string, BlockGeometry> {
  const result = new Map<string, BlockGeometry>();
  const bbox = selectionBoundingBox(members);
  if (!bbox) return result;
  const minSize = options.minSize ?? DEFAULT_MIN_BLOCK_SIZE;
  const { scalesX, scalesY, fromLeft, fromTop } = handleAxes(handle);
  const anchorX = fromLeft ? bbox.left + bbox.width : bbox.left;
  const anchorY = fromTop ? bbox.top + bbox.height : bbox.top;

  let scaleX = scalesX
    ? (bbox.width + (fromLeft ? -dx : dx)) / bbox.width
    : 1;
  let scaleY = scalesY
    ? (bbox.height + (fromTop ? -dy : dy)) / bbox.height
    : 1;
  if (options.maintainAspect && scalesX && scalesY) {
    const uniform = Math.max(scaleX, scaleY);
    scaleX = uniform;
    scaleY = uniform;
  }

  // Floor: no member below minSize on a driven axis (the Excalidraw idea —
  // floor the SCALE, keep the set uniform).
  // A member that is ALREADY thinner than the minimum on an axis (a SmartArt
  // leader line, a connector hairline) is drawn that thin by design and can
  // never be the floor: taking it into the floor pinned every diagram frame
  // holding one at scale 1 on that axis, so the frame grew but never shrank.
  let minScaleX = 0;
  let minScaleY = 0;
  for (const { startGeom } of members) {
    if (scalesX && startGeom.width >= minSize)
      minScaleX = Math.max(minScaleX, minSize / startGeom.width);
    if (scalesY && startGeom.height >= minSize)
      minScaleY = Math.max(minScaleY, minSize / startGeom.height);
  }
  if (scalesX) scaleX = Math.max(scaleX, Math.min(1, minScaleX));
  if (scalesY) scaleY = Math.max(scaleY, Math.min(1, minScaleY));

  // Cap: the scaled box stays in the area (growth toward the moving edge).
  const originX = Number.isFinite(bounds.originX) ? (bounds.originX as number) : 0;
  const originY = Number.isFinite(bounds.originY) ? (bounds.originY as number) : 0;
  if (scalesX && scaleX > 1) {
    const room = fromLeft ? anchorX - originX : originX + bounds.width - anchorX;
    if (room > 0) scaleX = Math.min(scaleX, room / bbox.width);
  }
  if (scalesY && scaleY > 1) {
    const room = fromTop ? anchorY - originY : originY + bounds.height - anchorY;
    if (room > 0) scaleY = Math.min(scaleY, room / bbox.height);
  }
  if (options.maintainAspect && scalesX && scalesY) {
    const uniform = Math.min(scaleX, scaleY);
    scaleX = uniform;
    scaleY = uniform;
  }

  for (const { blockId, startGeom } of members) {
    result.set(blockId, {
      left: Math.round(anchorX + (startGeom.left - anchorX) * scaleX),
      top: Math.round(anchorY + (startGeom.top - anchorY) * scaleY),
      width: Math.max(1, Math.round(startGeom.width * scaleX)),
      height: Math.max(1, Math.round(startGeom.height * scaleY)),
    });
  }
  return result;
}

/**
 * The one turn every member of a group shares (degrees, 0..360), or null when
 * they do not share one or it is none: a group the file turned draws each
 * member turned by the group's turn, so its members share it, and its own
 * frame is turned by it (PowerPoint and ONLYOFFICE frame a turned group by
 * its own turned box, and size it along its own axes).
 */
export function sharedMemberTurn(rotations: ReadonlyArray<number | null | undefined>): number | null {
  if (rotations.length < 2) return null;
  const first = (((rotations[0] ?? 0) % 360) + 360) % 360;
  if (Math.abs(first) < 1e-6 || Math.abs(first - 360) < 1e-6) return null;
  return rotations.every((value) => {
    const turn = (((value ?? 0) % 360) + 360) % 360;
    return Math.abs(turn - first) < 1e-6;
  })
    ? first
    : null;
}

/** Members carried into the frame of their shared `turn` about `pivot`, and back. */
function intoTurnedFrame(geometry: BlockGeometry, pivot: { x: number; y: number }, turn: number): BlockGeometry {
  const center = officeTurnAround(
    { x: geometry.left + geometry.width / 2, y: geometry.top + geometry.height / 2 },
    pivot,
    -turn,
  );
  return { left: center.x - geometry.width / 2, top: center.y - geometry.height / 2, width: geometry.width, height: geometry.height };
}

function outOfTurnedFrame(geometry: BlockGeometry, pivot: { x: number; y: number }, turn: number): BlockGeometry {
  const center = officeTurnAround(
    { x: geometry.left + geometry.width / 2, y: geometry.top + geometry.height / 2 },
    pivot,
    turn,
  );
  return {
    left: Math.round(center.x - geometry.width / 2),
    top: Math.round(center.y - geometry.height / 2),
    width: Math.round(geometry.width),
    height: Math.round(geometry.height),
  };
}

/**
 * The frame of a group whose members share `turn`: their boxes taken into the
 * group's own axes (about the centre of their slide box), bounded there, and
 * placed back -- an unturned box whose centre is the group's, to be drawn
 * turned by `turn` about its centre.
 */
export function turnedGroupFrame(
  geometries: ReadonlyArray<BlockGeometry>,
  turn: number,
): BlockGeometry | null {
  const members = geometries.map((startGeom, index) => ({ blockId: String(index), startGeom }));
  const box = selectionBoundingBox(members);
  if (!box) return null;
  const pivot = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  const local = selectionBoundingBox(
    members.map((member) => ({ blockId: member.blockId, startGeom: intoTurnedFrame(member.startGeom, pivot, turn) })),
  );
  if (!local) return null;
  const center = officeTurnAround({ x: local.left + local.width / 2, y: local.top + local.height / 2 }, pivot, turn);
  return { left: center.x - local.width / 2, top: center.y - local.height / 2, width: local.width, height: local.height };
}

/**
 * `applyMultiResize` along a turned group's own axes: the members and the
 * pointer's movement are carried into the group's frame, scaled there about
 * the edge opposite the grip, and carried back -- PowerPoint sizes a turned
 * group along its own sides, its members keeping their turn.
 */
export function applyTurnedMultiResize(
  members: ReadonlyArray<TransformMember>,
  turn: number,
  handle: ResizeHandle,
  dx: number,
  dy: number,
  options: MultiResizeOptions = {},
): Map<string, BlockGeometry> {
  const box = selectionBoundingBox(members);
  const result = new Map<string, BlockGeometry>();
  if (!box) return result;
  const pivot = { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  const local = members.map((member) => ({
    blockId: member.blockId,
    startGeom: intoTurnedFrame(member.startGeom, pivot, turn),
  }));
  const along = officeTurnAround({ x: dx, y: dy }, { x: 0, y: 0 }, -turn);
  // The group's own frame is not held to the slide's box (its turned box is).
  const free: DragBounds = { width: 1e12, height: 1e12 };
  const shifted = local.map((member) => ({
    blockId: member.blockId,
    startGeom: { ...member.startGeom, left: member.startGeom.left + free.width / 2, top: member.startGeom.top + free.height / 2 },
  }));
  for (const [id, geometry] of applyMultiResize(shifted, handle, along.x, along.y, free, options)) {
    result.set(
      id,
      outOfTurnedFrame({ ...geometry, left: geometry.left - free.width / 2, top: geometry.top - free.height / 2 }, pivot, turn),
    );
  }
  return result;
}
