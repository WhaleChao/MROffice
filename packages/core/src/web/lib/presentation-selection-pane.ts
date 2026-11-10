/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Pure logic for the presentation Selection Pane (PowerPoint / WPS "Home ▸
 * Arrange ▸ Selection Pane"): a per-slide panel listing every object on the
 * active slide with a show / hide toggle, an inline rename, and reorder (reusing
 * the existing one-step z-order math).
 *
 * Two additive, migration-safe block fields back the panel:
 *  - `hidden?: true`  — when set, the block is NOT rendered on the editor canvas
 *    or in present mode, and is excluded from marquee / click hit-testing (it can
 *    still be revealed again from the pane). Only a strict `true` is persisted;
 *    every other value (absent / `false`) means visible, so a visible block stays
 *    byte-for-byte as before (reuses `normalizeHiddenFlag`).
 *  - `name?: string`  — an optional user label shown in the pane (PowerPoint's
 *    editable object name). When absent the pane falls back to the block's
 *    accessible name. Only a non-empty trimmed value is persisted; clearing it
 *    DROPS the key (byte-clean).
 *
 * Dependency-free (no React / no DOM) so it is unit-testable in isolation and the
 * component handlers are thin wrappers. Mirrors the conventions of the sibling
 * pure modules (`presentation-z-order`, `presentation-slide-visibility`): no
 * mutation of inputs, deterministic, defensive no-ops.
 */

import { normalizeHiddenFlag } from "./presentation-slide-visibility";
import { presentationRestackedPlaces } from "./presentation-paint-order";

/** Max stored length of a user-assigned object name (matches the schema bound). */
export const MAX_BLOCK_NAME_LENGTH = 200;

/**
 * ASCII / C1 control characters (incl. newline / tab) stripped from a stored
 * object name — a name is single-line and a pasted value must not carry control
 * bytes. Built from explicit code-point ranges (0x00–0x1f, 0x7f–0x9f) so the
 * source carries no literal control chars; mirrors the ranges the sibling text
 * sanitizers use (`presentation-text-format` / `presentation-block-fx`).
 */
const NAME_CONTROL_RE = new RegExp("[\\u0000-\\u001f\\u007f-\\u009f]", "g");

/** The minimal block shape the pane logic needs. */
export interface SelectablePaneBlock {
  id: string;
  hidden?: unknown;
  name?: unknown;
  source?: unknown;
  [key: string]: unknown;
}

function paneRecord(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/**
 * Stable object identity exposed by the pane. Imported PPTX objects use their
 * authenticated native slide/shape identity, which survives a package reorder
 * and reimport; source-free objects use the persisted model id. A partial native
 * identity is unsafe and therefore returns null.
 */
export function selectionPaneObjectIdentity(
  block: SelectablePaneBlock,
): string | null {
  const source = paneRecord(block.source);
  const slideId =
    typeof source?.slideId === "string" ? source.slideId.trim() : "";
  const shapeId =
    typeof source?.shapeId === "string" ? source.shapeId.trim() : "";
  const inheritedFrom =
    source?.inheritedFrom === "layout" || source?.inheritedFrom === "master"
      ? source.inheritedFrom
      : null;
  if (inheritedFrom && slideId && !shapeId) {
    const sourceOrder = Number.isInteger(block.sourceOrder)
      ? block.sourceOrder
      : null;
    const modelId = typeof block.id === "string" ? block.id.trim() : "";
    return sourceOrder !== null && modelId
      ? `inherited:${inheritedFrom}:${slideId}:${sourceOrder}:${modelId}`
      : null;
  }
  if (slideId || shapeId) {
    return slideId && shapeId ? `pptx:${slideId}/${shapeId}` : null;
  }
  const groupId =
    typeof block.groupId === "string" ? block.groupId.trim() : "";
  if (groupId) return `group:${groupId}`;
  const modelId = typeof block.id === "string" ? block.id.trim() : "";
  return modelId ? `model:${modelId}` : null;
}

export interface SelectionPaneObjectGroup<T extends SelectablePaneBlock> {
  objectId: string;
  blocks: T[];
  indexes: number[];
}

/** The slide's objects back-to-front, each a pane row's blocks; null when an
 * id is missing or repeated (the pane then fails closed). */
export function selectionPaneObjectGroups<T extends SelectablePaneBlock>(
  blocks: readonly T[],
): SelectionPaneObjectGroup<T>[] | null {
  // An imported group is one row, as PowerPoint lists a group; a slide whose
  // group members are not side by side keeps one row per member.
  return (
    groupSelectionPaneObjects(blocks, nativeGroupIdentity) ??
    groupSelectionPaneObjects(blocks, selectionPaneObjectIdentity)
  );
}

/** A native group member's group identity, else its own. */
function nativeGroupIdentity(block: SelectablePaneBlock): string | null {
  const source = paneRecord(block.source);
  const groupId = typeof block.groupId === "string" ? block.groupId.trim() : "";
  return source && groupId && source.inheritedFrom === undefined
    ? `group:${groupId}`
    : selectionPaneObjectIdentity(block);
}

function groupSelectionPaneObjects<T extends SelectablePaneBlock>(
  blocks: readonly T[],
  identityOf: (block: T) => string | null,
): SelectionPaneObjectGroup<T>[] | null {
  const modelIds = new Set<string>();
  const seenObjectIds = new Set<string>();
  const groups: SelectionPaneObjectGroup<T>[] = [];
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const modelId = typeof block.id === "string" ? block.id.trim() : "";
    const objectId = identityOf(block);
    if (!modelId || !objectId || modelIds.has(modelId)) return null;
    modelIds.add(modelId);

    const previous = groups[groups.length - 1];
    if (previous?.objectId === objectId) {
      // One native shape can project to several adjacent model blocks. Keep it
      // atomic in the pane; a non-contiguous repeat is ambiguous and rejected.
      previous.blocks.push(block);
      previous.indexes.push(index);
      continue;
    }
    if (seenObjectIds.has(objectId)) return null;
    seenObjectIds.add(objectId);
    groups.push({ objectId, blocks: [block], indexes: [index] });
  }
  return groups;
}

/**
 * Selection Pane commands address objects by their persisted block id, never by
 * their current array position. A missing or duplicate id makes both selection
 * and reorder ambiguous, so the pane fails closed instead of guessing.
 */
export function selectionPaneHasStableObjectIdentity(
  blocks: readonly SelectablePaneBlock[],
): boolean {
  return selectionPaneObjectGroups(blocks) !== null;
}

/**
 * Move one complete pane object by one visual layer. Object groups are stored
 * back-to-front in `blocks[]`; swapping whole adjacent groups keeps every
 * multi-block projection/group contiguous and makes one pane step atomic.
 */
export function moveSelectionPaneObject<T extends SelectablePaneBlock>(
  blocks: readonly T[],
  objectId: string,
  direction: "forward" | "backward",
  /** Both swapped objects must be movable (neither locked); others are not asked. */
  canMove: (block: T) => boolean = () => true,
): T[] | null {
  const groups = selectionPaneObjectGroups(blocks);
  if (!groups) return null;
  const index = groups.findIndex((group) => group.objectId === objectId);
  if (index < 0) return null;
  const otherIndex = direction === "forward" ? index + 1 : index - 1;
  if (otherIndex < 0 || otherIndex >= groups.length) return null;
  if (
    ![...groups[index].blocks, ...groups[otherIndex].blocks].every((block) =>
      canMove(block),
    )
  ) {
    return null;
  }
  const nextGroups = [...groups];
  [nextGroups[index], nextGroups[otherIndex]] = [
    nextGroups[otherIndex],
    nextGroups[index],
  ];
  return presentationRestackedPlaces(nextGroups.flatMap((group) => group.blocks));
}

/** True when a block is explicitly hidden (not rendered / not hit-tested). */
export function isBlockHidden(
  block: SelectablePaneBlock | null | undefined,
): boolean {
  return !!block && block.hidden === true;
}








export function isBlockHiddenByGroup(
  block: SelectablePaneBlock | null | undefined,
): boolean {
  return !!block && paneRecord(block.source)?.nativeHiddenByGroup === true;
}







export function isBlockUnlinkedPlaceholder(
  block: SelectablePaneBlock | null | undefined,
): boolean {
  return !!block && paneRecord(block.source)?.nativePlaceholderUnlinked === true;
}

/**
 * Whether a block is painted at all: neither hidden itself nor by its group,
 * nor a SmartArt item its layout does not draw (`smartArtHidden`: Office's
 * unused text, kept in the outline but never painted, hit-tested or selected),
 * nor a slide placeholder linked to no layout placeholder.
 */
export function isBlockNotDrawn(
  block: SelectablePaneBlock | null | undefined,
): boolean {
  return (
    isBlockHidden(block) ||
    isBlockHiddenByGroup(block) ||
    isBlockUnlinkedPlaceholder(block) ||
    (!!block && block.smartArtHidden === true)
  );
}

/**
 * The VISIBLE blocks of a slide, in array order — every block that is not
 * explicitly hidden. Used to gate the canvas / present-mode render and the
 * marquee hit-test so a hidden object disappears from the slide (but is kept in
 * the deck and still listed in the pane). Inputs untouched.
 */
export function visibleBlocks<T extends SelectablePaneBlock>(
  blocks: readonly T[],
): T[] {
  return blocks.filter((block) => !isBlockNotDrawn(block));
}

/**
 * Sanitize a user-typed object name for STORAGE: strip control characters (so a
 * pasted name can't inject them), collapse to a trimmed single line, and clamp to
 * {@link MAX_BLOCK_NAME_LENGTH}. Returns `undefined` for an empty / whitespace /
 * non-string value so clearing the field drops the key (byte-clean) rather than
 * persisting an empty string.
 */
export function sanitizeBlockName(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const cleaned = value.replace(NAME_CONTROL_RE, " ").trim();
  if (!cleaned) return undefined;
  return cleaned.slice(0, MAX_BLOCK_NAME_LENGTH);
}

/**
 * The DISPLAY label for a block in the pane: its sanitized user `name` when set,
 * else the supplied accessible-name fallback (the editor passes
 * `blockAccessibleName`). Never empty — degrades to the block id if even the
 * fallback is blank, so every row is identifiable.
 */
export function paneDisplayName(
  block: SelectablePaneBlock,
  accessibleFallback: string,
): string {
  const explicit = sanitizeBlockName(block.name);
  if (explicit) return explicit;
  const fallback = accessibleFallback.trim();
  return fallback || block.id;
}

/**
 * Build the minimal patch for the HIDE toggle, applied over the existing block
 * (the caller spreads it through its single mutation entry point). Byte-clean:
 * `hidden: true` is persisted when hiding; showing DROPS the key (reuses
 * `normalizeHiddenFlag`, which maps everything but `true` to `undefined`). Input
 * not mutated.
 */
export function applyBlockHidden<T extends SelectablePaneBlock>(
  block: T,
  hidden: boolean,
): T {
  const updated: Record<string, unknown> = { ...block };
  const flag = normalizeHiddenFlag(hidden);
  if (flag) updated.hidden = true;
  else delete updated.hidden;
  return updated as T;
}

/**
 * Build the minimal patch for a RENAME, applied over the existing block. The
 * value is sanitized; a non-empty result is stored, an empty one DROPS the key
 * (so clearing the name reverts to the accessible-name fallback and serializes
 * byte-clean). Input not mutated.
 */
export function applyBlockName<T extends SelectablePaneBlock>(
  block: T,
  rawName: string,
): T {
  const updated: Record<string, unknown> = { ...block };
  const name = sanitizeBlockName(rawName);
  if (name) updated.name = name;
  else delete updated.name;
  return updated as T;
}

/**
 * The ordered rows the pane renders. PowerPoint lists objects FRONT-to-BACK (the
 * top-most object first); blocks paint in array order (a LATER index paints on
 * top), so the pane shows the array REVERSED. Each row carries the block, its
 * display label, hidden state, stable object id, and the ORIGINAL array index
 * (so a reorder / select maps back to the canonical `blocks[]` without the
 * component re-deriving it). PowerPoint's Selection Pane visual order is
 * front-to-back, while assistive technology consumes slide objects in the
 * reverse order; both one-based positions are carried explicitly so the UI can
 * describe that relationship without pretending they are the same order.
 * Pure; inputs untouched.
 */
export interface SelectionPaneRow<T extends SelectablePaneBlock> {
  block: T;
  blocks: readonly T[];
  blockIds: readonly string[];
  objectId: string;
  index: number;
  label: string;
  hidden: boolean;
  stackingPosition: number;
  readingOrder: number;
  setSize: number;
  canMoveForward: boolean;
  canMoveBackward: boolean;
}

export function selectionPaneRows<T extends SelectablePaneBlock>(
  blocks: readonly T[],
  nameOf: (block: T) => string,
  /** A grouped row's name, "Group 2" for the slide's second group from the
   *  back, as PowerPoint names groups; without it a group shows its first
   *  member's name. */
  groupNameOf?: (ordinal: number) => string,
): SelectionPaneRow<T>[] {
  const groups = selectionPaneObjectGroups(blocks);
  if (!groups) return [];
  const setSize = groups.length;
  let groupOrdinal = 0;
  const rows = groups.map((group, index) => ({
    block: group.blocks[0],
    blocks: group.blocks,
    blockIds: group.blocks.map((block) => block.id),
    objectId: group.objectId,
    index: group.indexes[0],
    label:
      groupNameOf &&
      group.blocks.length > 1 &&
      group.objectId.startsWith("group:")
        ? groupNameOf((groupOrdinal += 1))
        : paneDisplayName(group.blocks[0], nameOf(group.blocks[0])),
    hidden: group.blocks.every(isBlockHidden),
    // The last-painted object is the first visual row.
    stackingPosition: setSize - index,
    // Screen-reader traversal is the reverse of the visual Selection Pane.
    readingOrder: index + 1,
    setSize,
    canMoveForward: index < setSize - 1,
    canMoveBackward: index > 0,
  }));
  // Front-to-back: reverse so the top-most (last-painted) object is listed first.
  rows.reverse();
  return rows;
}
