/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A group's objects as the tree of groups the file holds them in, read from
 * the page's flat `groupMembers` (back to front) and each one's `groupPath`
 * ("" the group itself, "2" the nested group that is its third child, "2/0"
 * deeper).
 *
 * Microsoft Support, "Manage objects with the Selection pane": "For an
 * object that's in a group, you can only move it within the group". A
 * nested group is one object of its own group, so it moves among that
 * group's objects as a whole; an object of a group picked inside it
 * (select the group, then the object) is what Bring Forward / Send Backward
 * and Delete act on. The save writes the same tree back
 * (docx patchback `_restack_group_members`).
 */
import type { DocumentShapeGroupMember } from "./document-shapes";

/** One object of a group: one of its own members, or a nested group whole. */
export interface DocumentGroupUnit {
  /** A member's `memberId`; a nested group's {@link documentGroupUnitKey}. */
  readonly key: string;
  /** A nested group's own path; null for a member of the group itself. */
  readonly path: string | null;
  /** Its members, back to front (a nested group's, deeper ones included). */
  readonly members: readonly DocumentShapeGroupMember[];
}

export type DocumentGroupStep = "front" | "forward" | "backward" | "back";

/** The key a nested group goes by among its group's objects. */
export function documentGroupUnitKey(path: string): string {
  return `#${path}`;
}

const pathOf = (member: DocumentShapeGroupMember): string => member.groupPath ?? "";
const prefixOf = (path: string): string => (path ? `${path}/` : "");
const under = (member: DocumentShapeGroupMember, path: string): boolean => {
  const own = pathOf(member);
  return own === path || own.startsWith(prefixOf(path)) && (path !== "" || own !== "");
};

/**
 * The objects of the group at `path`, back to front. Null when its members
 * are not one run of the flat list, or a nested group's are not.
 */
export function documentGroupUnits(
  members: readonly DocumentShapeGroupMember[],
  path: string,
): DocumentGroupUnit[] | null {
  const prefix = prefixOf(path);
  const units: Array<{ key: string; path: string | null; members: DocumentShapeGroupMember[] }> = [];
  const seen = new Set<string>();
  let state: "before" | "inside" | "after" = "before";
  for (const member of members) {
    const own = pathOf(member);
    const inside = path === "" ? true : own === path || own.startsWith(prefix);
    if (!inside) {
      if (state === "inside") state = "after";
      continue;
    }
    if (state === "after") return null;
    state = "inside";
    let key: string;
    let unitPath: string | null;
    if (own === path) {
      key = member.memberId;
      unitPath = null;
    } else {
      const segment = own.slice(prefix.length).split("/")[0] ?? "";
      if (!/^\d+$/.test(segment)) return null;
      unitPath = `${prefix}${segment}`;
      key = documentGroupUnitKey(unitPath);
    }
    const last = units[units.length - 1];
    if (last && last.key === key && unitPath !== null) {
      last.members.push(member);
      continue;
    }
    if (seen.has(key)) return null;
    seen.add(key);
    units.push({ key, path: unitPath, members: [member] });
  }
  return units;
}

/** `path`, its leading `from` replaced by `to`. */
function moved(path: string, from: string, to: string): string {
  return path === from ? to : `${to}${path.slice(from.length)}`;
}

/**
 * The group's members with the objects of the group at `path` in the order
 * `keys` names them (back to front). A nested group whose place among its
 * group's objects changes is named by its new place, as the file reads it
 * once saved. Null when `keys` are not the group's objects, each once.
 */
export function documentGroupMembersInUnitOrder(
  members: readonly DocumentShapeGroupMember[],
  path: string,
  keys: readonly string[],
): DocumentShapeGroupMember[] | null {
  const units = documentGroupUnits(members, path);
  if (!units || keys.length !== units.length || new Set(keys).size !== keys.length) return null;
  const byKey = new Map(units.map((unit) => [unit.key, unit]));
  const ordered = keys.map((key) => byKey.get(key));
  if (ordered.some((unit) => unit === undefined)) return null;
  const prefix = prefixOf(path);
  const placed = (ordered as DocumentGroupUnit[]).flatMap((unit, index) => {
    if (unit.path === null) return [...unit.members];
    const to = `${prefix}${index}`;
    if (to === unit.path) return [...unit.members];
    const from = unit.path;
    return unit.members.map((member) => ({ ...member, groupPath: moved(pathOf(member), from, to) }));
  });
  const first = members.findIndex((member) => under(member, path));
  if (first < 0) return null;
  return [...members.slice(0, first), ...placed, ...members.slice(first + placed.length)];
}

/**
 * Bring Forward / Send Backward (and to Front / to Back) on one object of a
 * group: it passes its group's neighbouring object, a nested group whole,
 * never leaving its group. Null when it has nowhere to go.
 */
export function documentGroupMembersStepped(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
  step: DocumentGroupStep,
): DocumentShapeGroupMember[] | null {
  const member = members.find((candidate) => candidate.memberId === memberId);
  if (!member) return null;
  const path = pathOf(member);
  const units = documentGroupUnits(members, path);
  if (!units) return null;
  const keys = units.map((unit) => unit.key);
  const at = keys.indexOf(memberId);
  if (at < 0) return null;
  const to =
    step === "front" ? keys.length - 1 : step === "back" ? 0 : step === "forward" ? at + 1 : at - 1;
  if (to < 0 || to >= keys.length || to === at) return null;
  keys.splice(at, 1);
  keys.splice(to, 0, memberId);
  return documentGroupMembersInUnitOrder(members, path, keys);
}

/** Whether the step has nothing to do: the object is at that end of its group. */
export function documentGroupMemberAtBoundary(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
  step: DocumentGroupStep,
): boolean {
  return documentGroupMembersStepped(members, memberId, step) === null;
}

/**
 * The group's members without `memberId` (Delete on an object picked in its
 * group). A nested group left with nothing goes with it, and the objects
 * after it in its own group are named by their new places. Null when the
 * member is not there, when it is the group's last, or when a connector of
 * the group is glued to it.
 */
export function documentGroupMembersWithout(
  members: readonly DocumentShapeGroupMember[],
  memberId: string,
): DocumentShapeGroupMember[] | null {
  const member = members.find((candidate) => candidate.memberId === memberId);
  if (!member || members.length < 2) return null;
  if (
    members.some(
      (other) => other !== member && (other.startMemberId === memberId || other.endMemberId === memberId),
    )
  ) {
    return null;
  }
  const rest = members.filter((candidate) => candidate !== member);
  const parentOf = (path: string) => (path.includes("/") ? path.slice(0, path.lastIndexOf("/")) : "");
  // The object its group loses: the member, or the outermost nested group
  // the deletion empties.
  let gone = memberId;
  let parent = pathOf(member);
  for (let path = pathOf(member); path !== ""; path = parentOf(path)) {
    if (rest.some((other) => under(other, path))) break;
    gone = documentGroupUnitKey(path);
    parent = parentOf(path);
  }
  const before = documentGroupUnits(members, parent);
  if (!before) return null;
  // The others, in the same order, a nested group named by its new place.
  return documentGroupMembersInUnitOrder(
    rest,
    parent,
    before.map((unit) => unit.key).filter((key) => key !== gone),
  );
}

/** What {@link documentGroupMembersReorderable} reads of a node. */
interface GroupNodeLike {
  readonly type: { readonly name: string };
  readonly attrs: Readonly<Record<string, unknown>>;
}

/**
 * Whether a group's objects may change order among themselves, or one of
 * them be deleted on its own:
 * - a group drawn in this editor (as Ungroup asks), whose order is its
 *   `groupMembers`, which the page paints and the save writes in that order;
 *   one holding a nested group's frame keeps its runs;
 * - a Word group of the opened file (`wpg:wgp`) whose import read which of
 *   its groups each object is a child of (`groupPath`): each moves among its
 *   own group's objects, a nested group whole, and the save puts the
 *   group's children in that order (patchback `_restack_group_members`).
 * A SmartArt graphic keeps its layout's order (SmartArt Design > Move Up /
 * Move Down changes the outline, not the stack); a VML group, a canvas and
 * a diagram's drawing are not read with their groups.
 */
export function documentGroupMembersReorderable(node: GroupNodeLike): boolean {
  const members = node.attrs.groupMembers as readonly DocumentShapeGroupMember[] | undefined;
  if (
    node.type.name !== "documentShape" ||
    !Array.isArray(members) ||
    members.length < 2 ||
    typeof node.attrs.smartArtKind === "string"
  ) {
    return false;
  }
  if (node.attrs.source === null) {
    return members.every((member) => !member.scene3dFrame && !member.groupFillFrame && member.groupPath === undefined);
  }
  const source = node.attrs.source as { sourceAnchor?: unknown } | null | undefined;
  const anchor = typeof source?.sourceAnchor === "string" ? source.sourceAnchor : "";
  return (
    /\/group\[\d+\]$/.test(anchor) &&
    members.every((member) => typeof member.groupPath === "string") &&
    documentGroupUnits(members, "") !== null
  );
}
