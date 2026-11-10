/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeTurnAround } from "@mrnote/core/web/lib/office-object-frame.ts";
/**
 * A sheet group turned and flipped as ONE object, as Excel turns a group
 * (Microsoft Support, "Rotate or flip a text box, shape, WordArt, or
 * picture": a group rotates as a unit; ECMA-376 §20.1.7.5 `grpSpPr/a:xfrm`).
 *
 * The page keeps a group's members as they are drawn: each member's place
 * turned about the group's centre and the group's turn added to its own, and
 * the group's turn itself on every member (`groupRotation`), which the
 * workbook writer puts back on the group (`grpSpPr/a:xfrm@rot`) with each
 * child in the unturned group. The group's centre is not stored: it is the
 * one point a turn by the group's own turn about it maps the unturned
 * children's frame (their union) onto the members, `C = R(r)·c0` with `c0`
 * the centre of the members turned back about the origin.
 */

/** One member as drawn: its frame (drawing-layer px) and its own turn (degrees). */
export interface GroupRotationMember {
  readonly id: string;
  readonly frame: { readonly left: number; readonly top: number; readonly width: number; readonly height: number };
  readonly rotation: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

/** A member after the group turned or flipped. */
export interface GroupRotationResult {
  readonly frame: { left: number; top: number; width: number; height: number };
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

/** Degrees in (-180, 180], to the tenth the rotation writer keeps. */
export function normalizeGroupDegrees(value: number): number {
  const turned = Math.round(((((value + 180) % 360) + 360) % 360 - 180) * 10) / 10;
  return turned === -180 ? 180 : turned === 0 ? 0 : turned;
}

/** The turn a page group carries (`groupRotation`, the same on every member), or 0. */
export function groupRotationOf(members: readonly { groupRotation?: unknown }[]): number {
  const values = new Set(members.map((member) => member.groupRotation));
  if (values.size !== 1) return 0;
  const value = [...values][0];
  return typeof value === "number" && Number.isFinite(value) ? normalizeGroupDegrees(value) : 0;
}

// A positive turn is clockwise on the sheet's y-down axes.
function turn(x: number, y: number, degrees: number): { x: number; y: number } {
  return officeTurnAround({ x, y }, { x: 0, y: 0 }, degrees);
}

/** The group's centre: the centre of its unturned frame, turned with it. */
export function groupRotationCentre(members: readonly GroupRotationMember[], groupRotation: number): { x: number; y: number } {
  let left = Infinity;
  let top = Infinity;
  let right = -Infinity;
  let bottom = -Infinity;
  for (const { frame } of members) {
    const back = turn(frame.left + frame.width / 2, frame.top + frame.height / 2, -groupRotation);
    left = Math.min(left, back.x - frame.width / 2);
    right = Math.max(right, back.x + frame.width / 2);
    top = Math.min(top, back.y - frame.height / 2);
    bottom = Math.max(bottom, back.y + frame.height / 2);
  }
  return turn((left + right) / 2, (top + bottom) / 2, groupRotation);
}

/** The members after the group turns by `delta` degrees about its centre. */
export function groupMembersTurned(
  members: readonly GroupRotationMember[],
  groupRotation: number,
  delta: number,
): { members: Map<string, GroupRotationResult>; groupRotation: number } {
  const centre = groupRotationCentre(members, groupRotation);
  const result = new Map<string, GroupRotationResult>();
  for (const member of members) {
    const { frame } = member;
    const moved = turn(frame.left + frame.width / 2 - centre.x, frame.top + frame.height / 2 - centre.y, delta);
    result.set(member.id, {
      frame: {
        left: centre.x + moved.x - frame.width / 2,
        top: centre.y + moved.y - frame.height / 2,
        width: frame.width,
        height: frame.height,
      },
      rotation: normalizeGroupDegrees(member.rotation + delta),
      flipH: member.flipH === true,
      flipV: member.flipV === true,
    });
  }
  return { members: result, groupRotation: normalizeGroupDegrees(groupRotation + delta) };
}

/**
 * The members after the group flips across its own axis (`flipH`: left to
 * right; `flipV`: top to bottom), as Excel flips a group: each member's place
 * mirrored in the group's unturned frame, its flip toggled and its own turn
 * mirrored. The group keeps its turn.
 */
export function groupMembersFlipped(
  members: readonly GroupRotationMember[],
  groupRotation: number,
  axis: "flipH" | "flipV",
): Map<string, GroupRotationResult> {
  const centre = groupRotationCentre(members, groupRotation);
  const result = new Map<string, GroupRotationResult>();
  for (const member of members) {
    const { frame } = member;
    const local = turn(frame.left + frame.width / 2 - centre.x, frame.top + frame.height / 2 - centre.y, -groupRotation);
    const mirrored = axis === "flipH" ? { x: -local.x, y: local.y } : { x: local.x, y: -local.y };
    const moved = turn(mirrored.x, mirrored.y, groupRotation);
    result.set(member.id, {
      frame: {
        left: centre.x + moved.x - frame.width / 2,
        top: centre.y + moved.y - frame.height / 2,
        width: frame.width,
        height: frame.height,
      },
      // Mirrored in the group's frame: the turn relative to the group
      // reverses (2r - θ), and the member's own flip toggles.
      rotation: normalizeGroupDegrees(2 * groupRotation - member.rotation),
      flipH: axis === "flipH" ? member.flipH !== true : member.flipH === true,
      flipV: axis === "flipV" ? member.flipV !== true : member.flipV === true,
    });
  }
  return result;
}
