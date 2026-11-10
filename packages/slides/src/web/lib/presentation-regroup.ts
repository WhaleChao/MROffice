/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Regroup (Arrange ▸ Group Objects ▸ Regroup, Shape Format ▸
 * Group ▾ ▸ Regroup, right-click ▸ Group ▸ Regroup).
 *
 * Microsoft Support ("Group or ungroup shapes, pictures, or other objects"):
 * after ungrouping, "select any one of the shapes, pictures, or other objects
 * that were previously in a group" and choose Regroup -- the group that was
 * taken apart comes back, its members still where the user may have moved
 * them. PowerPoint remembers the group per object; the editor keeps the same
 * memory for the editing session (it is not part of the page: a deck reopened
 * later has nothing to regroup, as PowerPoint after the file is closed).
 *
 * Pure: the window records what an Ungroup dissolved (`ungroupRecords`) and
 * asks which record a selection can restore (`regroupRecordFor`); the group is
 * rebuilt by `groupBlocks` (presentation-object-buttons), the same transform
 * Group runs, under the dissolved group's own id when it is free.
 */

import { blockGroupId, blocksInGroup } from "@mrnote/core/web/lib/presentation-crop-group.ts";

/** One dissolved group: its id and its members when it was taken apart. */
export interface RegroupRecord {
  readonly slideId: string;
  readonly groupId: string;
  readonly memberIds: readonly string[];
}

interface GroupableBlock {
  id: string;
  groupId?: unknown;
  [key: string]: unknown;
}

/** How many dissolved groups the session remembers (newest kept). */
export const REGROUP_MEMORY_LIMIT = 32;

/**
 * The groups an Ungroup took apart: every outermost group of `before` none of
 * whose members carries that group id in `after`.
 */
export function ungroupRecords(
  slideId: string,
  before: readonly GroupableBlock[],
  after: readonly GroupableBlock[],
): RegroupRecord[] {
  const afterById = new Map(after.map((block) => [block.id, block]));
  const seen = new Set<string>();
  const records: RegroupRecord[] = [];
  for (const block of before) {
    const groupId = blockGroupId(block);
    if (!groupId || seen.has(groupId)) continue;
    seen.add(groupId);
    const members = blocksInGroup(before, groupId);
    if (members.length < 2) continue;
    const dissolved = members.every((member) => {
      const next = afterById.get(member.id);
      return next !== undefined && blockGroupId(next) !== groupId;
    });
    if (dissolved) {
      records.push({ slideId, groupId, memberIds: members.map((member) => member.id) });
    }
  }
  return records;
}

/** The memory after an Ungroup: the new records first, a group recorded
 * again replaced, the oldest dropped past the limit. */
export function rememberUngroup(
  memory: readonly RegroupRecord[],
  records: readonly RegroupRecord[],
): RegroupRecord[] {
  if (records.length === 0) return [...memory];
  const key = (record: RegroupRecord) => `${record.slideId}\u0000${record.groupId}`;
  const fresh = new Set(records.map(key));
  return [...records, ...memory.filter((record) => !fresh.has(key(record)))].slice(
    0,
    REGROUP_MEMORY_LIMIT,
  );
}

/**
 * The dissolved group a selection can bring back: the newest record on this
 * slide naming a selected object, whose members all still exist and are not
 * already one group again. Null when there is none (Regroup is unavailable).
 */
export function regroupRecordFor(
  memory: readonly RegroupRecord[],
  slideId: string,
  blocks: readonly GroupableBlock[],
  selectedIds: ReadonlySet<string> | readonly string[],
): RegroupRecord | null {
  const selected = selectedIds instanceof Set ? selectedIds : new Set(selectedIds);
  if (selected.size === 0) return null;
  const byId = new Map(blocks.map((block) => [block.id, block]));
  for (const record of memory) {
    if (record.slideId !== slideId) continue;
    if (!record.memberIds.some((id) => selected.has(id))) continue;
    const members = record.memberIds.map((id) => byId.get(id));
    if (members.some((member) => member === undefined)) continue;
    const firstGroup = blockGroupId(members[0]);
    const alreadyOneGroup =
      firstGroup !== undefined &&
      members.every((member) => blockGroupId(member) === firstGroup);
    if (alreadyOneGroup) continue;
    return record;
  }
  return null;
}

/** The id the restored group takes: its own, unless another object now uses it. */
export function regroupGroupId(
  record: RegroupRecord,
  blocks: readonly GroupableBlock[],
): string | undefined {
  const members = new Set(record.memberIds);
  return blocks.some(
    (block) => !members.has(block.id) && blockGroupId(block) === record.groupId,
  )
    ? undefined
    : record.groupId;
}
