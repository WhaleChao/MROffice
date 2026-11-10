/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Host-neutral SmartArt OUTLINE editing (the Office "Text pane" model).
 *
 * One diagram is one ordered list of items; every item owns exactly one node
 * and one label in the generated scene, and connectors stay structural. The
 * editing verbs here are TOTAL and side-effect-free: each returns a NEW outline
 * or `null` when the edit would break the durable contract, so a host applies a
 * change only when the whole rebuild succeeds (fail-closed relayout).
 *
 * Every function is generic over the host's own block type, so a host gets its
 * own blocks back. The nesting cap lives with the generator
 * (`MAX_SMARTART_OUTLINE_LEVEL` in `./engine`) because the layouts enforce it.
 */

import {
  isSmartArtHierBranch,
  MAX_SMARTART_OUTLINE_LEVEL,
  readSmartArtBlockMarker,
  SMARTART_ASSISTANT_KINDS,
  type OfficeSmartArtHostNode,
  type SmartArtHierBranch,
} from "./model";

export interface SmartArtOutlineItem {
  id: string;
  text: string;
  level: number;
  /**
   * An organization chart's ASSISTANT (Office's Add Assistant; DiagramML
   * `dgm:pt type="asst"`). Always below a parent (never at level 0) and only
   * ever `true`; the Text Pane lists it among its manager's children, and the
   * layout hangs it off the manager's trunk above them.
   */
  assistant?: true;
  /**
   * The item's title (Name and Title Organization Chart: the box under the
   * name), Office's words on the item's following sibling transition. Read
   * from the diagram's title members; absent when the item has none.
   */
  title?: string;
  /**
   * How this box's subordinates hang below it (an organization chart's
   * Layout: Office's `hierBranch`), kept on the item's node member as
   * `smartArtHierBranch`. Absent: the layout's own.
   */
  hierBranch?: SmartArtHierBranch;
}

export const MAX_SMARTART_OUTLINE_ITEMS = 50;
export const MAX_SMARTART_OUTLINE_TEXT_LENGTH = 2_000;

/** Read one complete, ordered outline from durable SmartArt label markers. */
export function readSmartArtOutline(
  blocks: readonly OfficeSmartArtHostNode[],
  groupId: string,
): SmartArtOutlineItem[] | null {
  // Every label of the diagram is a row, or there is no outline: a label
  // left out (its text past the bound, a broken marker) would be dropped by
  // the next rebuild from the outline -- the Text Pane listed two items of
  // three and its next edit deleted the third (PPT-B318).
  let labels = 0;
  const rows = blocks.flatMap((block) => {
    const marker = readSmartArtBlockMarker(block);
    const record = block as Record<string, unknown>;
    if (marker?.groupId === groupId && marker.role === "label") labels += 1;
    return marker?.groupId === groupId &&
      marker.role === "label" &&
      marker.itemId !== null &&
      marker.itemIndex !== null &&
      marker.itemIndex >= 0 &&
      marker.level !== null &&
      marker.level >= 0 &&
      marker.level <= MAX_SMARTART_OUTLINE_LEVEL &&
      typeof record.text === "string" &&
      record.text.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH
      ? [
          {
            id: marker.itemId,
            index: marker.itemIndex,
            level: marker.level,
            text: record.text,
            assistant: record.smartArtAssistant === true,
          },
        ]
      : [];
  });
  if (
    rows.length === 0 ||
    rows.length !== labels ||
    rows.length > MAX_SMARTART_OUTLINE_ITEMS
  ) {
    return null;
  }
  rows.sort((left, right) => left.index - right.index);
  const ids = new Set<string>();
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    if (
      row.index !== index ||
      ids.has(row.id) ||
      (index === 0 && row.level !== 0) ||
      (index > 0 && row.level > rows[index - 1].level + 1)
    ) {
      return null;
    }
    ids.add(row.id);
  }
  // Each box's org chart Layout, stated on its node member.
  const branches = new Map<string, SmartArtHierBranch>();
  for (const block of blocks) {
    const marker = readSmartArtBlockMarker(block);
    const branch = (block as Record<string, unknown>).smartArtHierBranch;
    if (marker?.groupId === groupId && marker.role === "node" && marker.itemId && isSmartArtHierBranch(branch)) {
      branches.set(marker.itemId, branch);
    }
  }
  // Titles: the text of the members that name an item by its index.
  const titles = new Map<number, string>();
  for (const block of blocks) {
    const record = block as Record<string, unknown>;
    const at = record.smartArtTitleIndex;
    if (
      record.groupId === groupId &&
      typeof at === "number" &&
      Number.isInteger(at) &&
      at >= 0 &&
      at < rows.length &&
      typeof record.text === "string" &&
      record.text.length > 0 &&
      record.text.length <= MAX_SMARTART_OUTLINE_TEXT_LENGTH
    ) {
      titles.set(at, record.text);
    }
  }
  return rows.map(({ id, text, level, assistant }, index) => {
    const item: SmartArtOutlineItem =
      assistant && level > 0 ? { id, text, level, assistant: true } : { id, text, level };
    const title = titles.get(index);
    const branch = branches.get(id);
    return {
      ...item,
      ...(title !== undefined ? { title } : {}),
      ...(branch ? { hierBranch: branch } : {}),
    };
  });
}

function copyOutline(
  items: readonly SmartArtOutlineItem[],
): SmartArtOutlineItem[] | null {
  if (items.length === 0 || items.length > MAX_SMARTART_OUTLINE_ITEMS) {
    return null;
  }
  const ids = new Set<string>();
  const next = items.map((item) => ({ ...item }));
  for (let index = 0; index < next.length; index += 1) {
    const item = next[index];
    // An assistant belongs to a manager: at level 0 it is an ordinary item.
    if (item.assistant !== true || item.level === 0) delete item.assistant;
    if (item.hierBranch !== undefined && !isSmartArtHierBranch(item.hierBranch)) delete item.hierBranch;
    if (
      !item.id ||
      ids.has(item.id) ||
      typeof item.text !== "string" ||
      item.text.length > MAX_SMARTART_OUTLINE_TEXT_LENGTH ||
      !Number.isInteger(item.level) ||
      item.level < 0 ||
      item.level > MAX_SMARTART_OUTLINE_LEVEL ||
      (index === 0 && item.level !== 0) ||
      (index > 0 && item.level > next[index - 1].level + 1)
    ) {
      return null;
    }
    ids.add(item.id);
  }
  return next;
}

export function addSmartArtOutlineItem(
  items: readonly SmartArtOutlineItem[],
  anchorId: string,
  position: "before" | "after",
  item: SmartArtOutlineItem,
): SmartArtOutlineItem[] | null {
  const next = copyOutline(items);
  if (!next || next.length >= MAX_SMARTART_OUTLINE_ITEMS) return null;
  const anchorIndex = next.findIndex((candidate) => candidate.id === anchorId);
  if (anchorIndex < 0 || next.some((candidate) => candidate.id === item.id)) {
    return null;
  }
  const insertIndex = anchorIndex + (position === "after" ? 1 : 0);
  const level = next[anchorIndex].level;
  // Enter on an assistant in the Text Pane adds another assistant (Microsoft
  // Support, "Create an organization chart").
  next.splice(insertIndex, 0, {
    ...item,
    level,
    ...(next[anchorIndex].assistant ? { assistant: true as const } : {}),
  });
  return copyOutline(next);
}

/** The index of `index`'s parent item, or -1 for a level-0 item. */
function smartArtOutlineParent(
  items: readonly SmartArtOutlineItem[],
  index: number,
): number {
  const level = items[index]?.level ?? 0;
  for (let cursor = index - 1; cursor >= 0; cursor -= 1) {
    if (items[cursor].level < level) return cursor;
  }
  return -1;
}

/**
 * The index of the last item in `index`'s subtree: the item itself when it has
 * no deeper items directly beneath it.
 */
export function smartArtOutlineSubtreeEnd(
  items: readonly SmartArtOutlineItem[],
  index: number,
): number {
  const level = items[index]?.level ?? 0;
  let end = index;
  while (end + 1 < items.length && items[end + 1].level > level) end += 1;
  return end;
}

/**
 * Where the SmartArt Design tab's Add Shape puts the new shape, relative to the
 * selected one. Microsoft's `MsoSmartArtNodePosition` names the same four:
 * After (the plain Add Shape), Before, Above (a new parent the selected shape
 * moves under) and Below (a new child of the selected shape).
 */
export type SmartArtOutlineInsertPosition =
  | "after"
  | "before"
  | "above"
  | "below"
  /**
   * Add Assistant (organization charts): a new ASSISTANT of the selected
   * shape, listed after its other children, as the Text Pane lists one.
   */
  | "assistant"
  /**
   * Add Bullet: a line of bullet text under the selected shape, its last
   * child, as Below adds one; offered where the layout takes bullets.
   */
  | "bullet";

/**
 * Add one item relative to `anchorId`, in Office's data-model sense rather than
 * the Text Pane's line sense: "after" lands behind the anchor's whole subtree
 * (so the anchor keeps its children), "below" becomes the anchor's LAST child,
 * and "above" inserts a new item at the anchor's level and moves the anchor and
 * its subtree one level down. `null` when the result would break the outline
 * contract (item cap, nesting cap, duplicate id).
 */
export function insertSmartArtOutlineItem(
  items: readonly SmartArtOutlineItem[],
  anchorId: string,
  position: SmartArtOutlineInsertPosition,
  item: SmartArtOutlineItem,
): SmartArtOutlineItem[] | null {
  const next = copyOutline(items);
  if (!next || next.length >= MAX_SMARTART_OUTLINE_ITEMS) return null;
  const anchorIndex = next.findIndex((candidate) => candidate.id === anchorId);
  if (anchorIndex < 0 || next.some((candidate) => candidate.id === item.id)) {
    return null;
  }
  const level = next[anchorIndex].level;
  const end = smartArtOutlineSubtreeEnd(next, anchorIndex);
  const anchorAssistant = next[anchorIndex].assistant === true;
  // A sibling of an assistant is another assistant of the same manager; an
  // assistant has no subordinates (MRAI draws it on the manager's assistant
  // row), so nothing is added beneath or above one.
  const sibling = anchorAssistant ? { assistant: true as const } : {};
  if (position === "before") {
    next.splice(anchorIndex, 0, { ...item, level, ...sibling });
  } else if (position === "after") {
    next.splice(end + 1, 0, { ...item, level, ...sibling });
  } else if (position === "below" || position === "assistant" || position === "bullet") {
    if (anchorAssistant || level + 1 > MAX_SMARTART_OUTLINE_LEVEL) return null;
    next.splice(end + 1, 0, {
      ...item,
      level: level + 1,
      ...(position === "assistant" ? { assistant: true as const } : {}),
    });
  } else {
    if (anchorAssistant) return null;
    for (let cursor = anchorIndex; cursor <= end; cursor += 1) {
      next[cursor].level += 1;
      if (next[cursor].level > MAX_SMARTART_OUTLINE_LEVEL) return null;
    }
    next.splice(anchorIndex, 0, { ...item, level });
  }
  return copyOutline(next);
}

/**
 * Move Up / Move Down: swap the item (with its subtree) with its previous or
 * next SIBLING (with that one's subtree). An item with no sibling on that side
 * is refused, which is where Office greys the command out.
 */
export function moveSmartArtOutlineItem(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
  direction: "up" | "down",
): SmartArtOutlineItem[] | null {
  const next = copyOutline(items);
  const index = next?.findIndex((item) => item.id === itemId) ?? -1;
  if (!next || index < 0) return null;
  const level = next[index].level;
  const end = smartArtOutlineSubtreeEnd(next, index);
  let first: [number, number];
  let second: [number, number];
  if (direction === "up") {
    let sibling = index - 1;
    while (sibling >= 0 && next[sibling].level > level) sibling -= 1;
    if (sibling < 0 || next[sibling].level !== level) return null;
    first = [sibling, index - 1];
    second = [index, end];
  } else {
    const sibling = end + 1;
    if (sibling >= next.length || next[sibling].level !== level) return null;
    first = [index, end];
    second = [sibling, smartArtOutlineSubtreeEnd(next, sibling)];
  }
  return copyOutline([
    ...next.slice(0, first[0]),
    ...next.slice(second[0], second[1] + 1),
    ...next.slice(first[0], first[1] + 1),
    ...next.slice(second[1] + 1),
  ]);
}

export function removeSmartArtOutlineItem(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
): SmartArtOutlineItem[] | null {
  const next = copyOutline(items);
  if (!next || next.length <= 1) return null;
  const index = next.findIndex((item) => item.id === itemId);
  if (index < 0) return null;
  // Office deletes a node by promoting its FIRST child into its place; the
  // other children stay a level down, under that heir (Microsoft Learn,
  // SmartArtNode.Delete: "When the node is deleted, the first child gets
  // promoted"; Microsoft Support, "Delete a SmartArt graphic or a shape":
  // "the first line of Level 2 text is promoted to Level 1"). An assistant
  // below a deleted shape is the one promoted ("the Assistant shape is
  // automatically promoted to the top of the hierarchy"). Every child used to
  // move up a level, so deleting a manager split its reports into separate
  // trees.
  const removedLevel = next[index].level;
  const end = smartArtOutlineSubtreeEnd(next, index);
  const children: number[] = [];
  for (let cursor = index + 1; cursor <= end; cursor += 1) {
    if (next[cursor].level === removedLevel + 1) children.push(cursor);
  }
  if (children.length === 0) {
    next.splice(index, 1);
    return copyOutline(next);
  }
  const heir = children.find((cursor) => next[cursor].assistant) ?? children[0];
  const heirEnd = smartArtOutlineSubtreeEnd(next, heir);
  const promoted = next
    .slice(heir, heirEnd + 1)
    .map((item) => ({ ...item, level: item.level - 1 }));
  // In the deleted shape's place it reports to that shape's own manager.
  delete promoted[0].assistant;
  return copyOutline([
    ...next.slice(0, index),
    ...promoted,
    ...next.slice(index + 1, heir),
    ...next.slice(heirEnd + 1, end + 1),
    ...next.slice(end + 1),
  ]);
}

/**
 * The Text Pane's line merge: Delete at the end of a line, or Backspace at the
 * start of the next one, joins the two lines (Microsoft Support, "Keyboard
 * shortcuts for SmartArt graphics": "Merge two lines of text — Delete at the
 * end of the first line"). `keepId` keeps the joined `text`; `goneId`, the
 * line right after it, leaves as a deleted line does, its first child taking
 * its place ({@link removeSmartArtOutlineItem}).
 */
export function mergeSmartArtOutlineItems(
  items: readonly SmartArtOutlineItem[],
  keepId: string,
  goneId: string,
  text: string,
): SmartArtOutlineItem[] | null {
  if (text.length > MAX_SMARTART_OUTLINE_TEXT_LENGTH) return null;
  const keep = items.findIndex((item) => item.id === keepId);
  if (keep < 0 || items[keep + 1]?.id !== goneId) return null;
  const removed = removeSmartArtOutlineItem(items, goneId);
  return removed ? updateSmartArtOutlineItemText(removed, keepId, text) : null;
}

/**
 * Promote one item a level, carrying its lower-level items with it. Microsoft
 * documents the ribbon's Promote as "Promotes the current node (and all its
 * children) a single level" (Microsoft Learn, SmartArtNode.Promote, which
 * "mimics the Promote button"): A / B / C under B becomes A, B / C under B.
 * The outline keeps its order, so a later sibling of the item follows it and
 * becomes its child, exactly as the Text Pane's Shift+Tab reads.
 */
export function promoteSmartArtOutlineItem(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
): SmartArtOutlineItem[] | null {
  const next = copyOutline(items);
  const index = next?.findIndex((item) => item.id === itemId) ?? -1;
  if (!next || index <= 0 || next[index].level === 0) return null;
  const end = smartArtOutlineSubtreeEnd(next, index);
  for (let cursor = index; cursor <= end; cursor += 1) next[cursor].level -= 1;
  // Promoted, it reports to a different shape: no longer that one's assistant.
  // Its own children still report to it, so theirs stand.
  delete next[index].assistant;
  return copyOutline(next);
}

/**
 * Demote one item a level, alone. Microsoft documents the ribbon's Demote as
 * "Demotes the current node a single level" (Microsoft Learn,
 * SmartArtNode.Demote): A, B / C under B becomes A / B and C under A, so the
 * item's former children keep their level and follow it as its siblings.
 */
export function demoteSmartArtOutlineItem(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
): SmartArtOutlineItem[] | null {
  const next = copyOutline(items);
  const index = next?.findIndex((item) => item.id === itemId) ?? -1;
  if (!next || index <= 0) return null;
  const maximum = Math.min(
    MAX_SMARTART_OUTLINE_LEVEL,
    next[index - 1].level + 1,
  );
  if (next[index].level >= maximum) return null;
  next[index].level += 1;
  // Demoted under an assistant would give the assistant a subordinate, which
  // its row cannot hold; demoted anywhere else it reports to a new shape.
  if (next[smartArtOutlineParent(next, index)]?.assistant) return null;
  delete next[index].assistant;
  return copyOutline(next);
}

/** Replace one outline label without weakening the durable outline bounds. */
export function updateSmartArtOutlineItemText(
  items: readonly SmartArtOutlineItem[],
  itemId: string,
  text: string,
): SmartArtOutlineItem[] | null {
  if (text.length > MAX_SMARTART_OUTLINE_TEXT_LENGTH) return null;
  const next = copyOutline(items);
  if (!next) return null;
  const index = next.findIndex((item) => item.id === itemId);
  if (index < 0) return null;
  next[index].text = text;
  return copyOutline(next);
}

/**
 * Rebind freshly generated SmartArt geometry to one durable outline. Every
 * item must own exactly one node and one label; connectors remain structural
 * and never acquire item identity. This is the fail-closed seam used by the
 * editor after Add/Delete/Promote/Demote rebuilds the diagram.
 */
export function applySmartArtOutline<T extends OfficeSmartArtHostNode>(
  blocks: readonly T[],
  items: readonly SmartArtOutlineItem[],
): T[] | null {
  const outline = copyOutline(items);
  if (!outline) return null;
  const markers = blocks.map(readSmartArtBlockMarker);
  if (markers.some((marker) => marker === null)) return null;
  const first = markers[0];
  if (!first) return null;
  const counts = outline.map(() => ({ node: 0, label: 0 }));
  for (const marker of markers) {
    if (
      !marker ||
      marker.groupId !== first.groupId ||
      marker.kind !== first.kind
    ) {
      return null;
    }
    if (marker.role === "connector") {
      if (marker.itemId !== null || marker.itemIndex !== null) return null;
      continue;
    }
    if (
      marker.itemIndex === null ||
      marker.itemIndex < 0 ||
      marker.itemIndex >= outline.length
    ) {
      return null;
    }
    counts[marker.itemIndex][marker.role] += 1;
  }
  if (counts.some(({ node, label }) => node !== 1 || label !== 1)) {
    return null;
  }
  return blocks.map((block, index) => {
    const marker = markers[index]!;
    if (marker.role === "connector") return block;
    const item = outline[marker.itemIndex!];
    const identity = {
      smartArtItemId: item.id,
      smartArtItemIndex: marker.itemIndex!,
      smartArtLevel: item.level,
    };
    // The assistant marker follows the outline: set on an assistant's node
    // and label in a layout that has the position, absent everywhere else (a
    // layout change to one without it makes the item an ordinary child).
    const {
      smartArtAssistant: _assistant,
      smartArtHierBranch: _branch,
      ...unmarked
    } = block as T & {
      smartArtAssistant?: true;
      smartArtHierBranch?: SmartArtHierBranch;
    };
    const marked = {
      ...unmarked,
      ...identity,
      ...(item.assistant && SMARTART_ASSISTANT_KINDS.has(marker.kind)
        ? { smartArtAssistant: true as const }
        : {}),
      // An org chart box's Layout, on its node, in a layout that reads it.
      ...(item.hierBranch && marker.role === "node" && SMARTART_ASSISTANT_KINDS.has(marker.kind)
        ? { smartArtHierBranch: item.hierBranch }
        : {}),
    };
    if (marker.role === "label" && block.type === "text") {
      const { runs: _runs, ...withoutRuns } = marked as typeof marked & {
        runs?: unknown;
      };
      return { ...withoutRuns, text: item.text } as unknown as T;
    }
    return marked as unknown as T;
  });
}

export type OfficeSmartArtOutlineItem = SmartArtOutlineItem;
