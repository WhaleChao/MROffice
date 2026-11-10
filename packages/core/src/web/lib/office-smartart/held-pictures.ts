/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pictures a SmartArt graphic keeps while it wears a layout with no
 * picture shapes.
 *
 * Office lets a picture layout become any other layout. What the new layout
 * cannot show is kept, not lost: "Text, pictures, or other content that is
 * not displayed is identified in the Text pane with a red X", it "is still
 * available if you switch to another layout", and it "is not saved" if the
 * file is closed in that layout (Microsoft Support, "Anything above marked
 * with a red X will not appear in this SmartArt graphic and will not be
 * saved"; "Learn more about SmartArt graphics").
 *
 * So the pictures live here, for the editing session and never in the page
 * or a file: a Change Layout to a layout without pictures holds them by the
 * graphic's group id and each picture's durable item id, and a Change Layout
 * back to a picture layout fills each item's slot with its picture again.
 * An item deleted in between takes its picture with it. Each host keeps its
 * own picture record (a deck's image block fields, a worksheet's or a Word
 * group's slot `src` and Picture Format), so the record is opaque here.
 */

import { isSmartArtPictureKind } from "./model";
import {
  smartArtSlotPictureFields,
  smartArtSlotPictureFormatOf,
} from "./slot-picture-format";

type HeldPicture = Readonly<Record<string, unknown>>;

/** Whether a Change Layout from `from` to `to` hides the pictures. */
export function smartArtLayoutChangeHidesPictures(from: unknown, to: unknown): boolean {
  return isSmartArtPictureKind(from) && !isSmartArtPictureKind(to);
}

/** Whether a Change Layout to `to` shows pictures again. */
export function smartArtLayoutChangeShowsPictures(to: unknown): boolean {
  return isSmartArtPictureKind(to);
}

/**
 * The pictures in a worksheet's or a Word group's SLOTS (a slot member:
 * `smartArtPicturePlaceholder`, its picture `src` and its Picture Format), by
 * item id, held for `groupId`: the slot model the two hosts share.
 */
export function holdSmartArtSlotPictures<T>(
  groupId: string,
  members: readonly T[],
  srcOf: (member: T) => unknown,
): void {
  const pictures = new Map<string, HeldPicture>();
  for (const member of members) {
    const record = member as unknown as Record<string, unknown>;
    const src = srcOf(member);
    if (
      record.smartArtPicturePlaceholder !== true ||
      typeof record.smartArtItemId !== "string" ||
      typeof src !== "string" ||
      !src
    ) {
      continue;
    }
    pictures.set(record.smartArtItemId, { src, format: smartArtSlotPictureFormatOf(record) });
  }
  holdSmartArtPictures(groupId, pictures);
}

/** {@link restoreSmartArtHeldPictures} for the slot model: `writeSrc` puts the src where the host keeps it. */
export function restoreSmartArtSlotPictures<T>(
  groupId: string,
  members: readonly T[],
  writeSrc: (member: T, src: string) => T,
): T[] {
  return restoreSmartArtHeldPictures(
    groupId,
    members,
    (member) => {
      const record = member as unknown as Record<string, unknown>;
      return record.smartArtPicturePlaceholder === true && typeof record.smartArtItemId === "string"
        ? record.smartArtItemId
        : null;
    },
    (member, picture) => {
      const format = picture.format as Parameters<typeof smartArtSlotPictureFields>[0] | undefined;
      const written = writeSrc(member, String(picture.src));
      return format && (format.crop || format.style)
        ? ({ ...(written as object), ...smartArtSlotPictureFields(format) } as T)
        : written;
    },
  );
}

/** What a deck picture keeps of its image block: its picture, not its place in the diagram. */
const BLOCK_PICTURE_PLACE_KEYS = new Set([
  "id", "type", "left", "top", "width", "height", "rotation", "groupId", "zIndex", "flipH", "flipV",
]);

/** The pictures in a deck graphic's image blocks (`smartArtRole: "node"`), by item id, held for `groupId`. */
export function holdSmartArtBlockPictures(groupId: string, blocks: readonly unknown[]): void {
  const pictures = new Map<string, HeldPicture>();
  for (const block of blocks) {
    const record = block as Record<string, unknown>;
    if (
      record.type !== "image" ||
      record.smartArtRole !== "node" ||
      typeof record.smartArtItemId !== "string" ||
      !isSmartArtPictureKind(record.smartArtKind)
    ) {
      continue;
    }
    pictures.set(
      record.smartArtItemId,
      Object.fromEntries(
        Object.entries(record).filter(([key]) => !BLOCK_PICTURE_PLACE_KEYS.has(key) && !key.startsWith("smartArt")),
      ),
    );
  }
  holdSmartArtPictures(groupId, pictures);
}

/** {@link restoreSmartArtHeldPictures} for a deck graphic's image blocks. */
export function restoreSmartArtBlockPictures<T>(groupId: string, blocks: readonly T[]): T[] {
  return restoreSmartArtHeldPictures(
    groupId,
    blocks,
    (block) => {
      const record = block as unknown as Record<string, unknown>;
      return record.type === "image" && record.smartArtRole === "node" && typeof record.smartArtItemId === "string"
        ? record.smartArtItemId
        : null;
    },
    (block, picture) => ({ ...(block as object), ...picture }) as T,
  );
}

const held = new Map<string, Map<string, HeldPicture>>();

/** Keep `pictures` (item id → the host's picture record) for `groupId`. */
export function holdSmartArtPictures(
  groupId: string,
  pictures: ReadonlyMap<string, HeldPicture>,
): void {
  if (!groupId || pictures.size === 0) return;
  const kept = held.get(groupId) ?? new Map<string, HeldPicture>();
  for (const [itemId, picture] of pictures) kept.set(itemId, picture);
  held.set(groupId, kept);
}

/** The pictures `groupId` holds, by item id; empty when it holds none. */
export function heldSmartArtPictures(groupId: string): ReadonlyMap<string, HeldPicture> {
  return held.get(groupId) ?? new Map();
}

/** The items of `groupId` whose picture is held (the Text Pane marks them with a red X). */
export function smartArtHeldPictureItemIds(groupId: string | null | undefined): ReadonlySet<string> {
  return new Set(groupId ? held.get(groupId)?.keys() ?? [] : []);
}

/** Forget the pictures `groupId` gave back to `itemIds` (or all of them). */
export function releaseSmartArtPictures(groupId: string, itemIds?: Iterable<string>): void {
  const kept = held.get(groupId);
  if (!kept) return;
  if (itemIds === undefined) {
    held.delete(groupId);
    return;
  }
  for (const itemId of itemIds) kept.delete(itemId);
  if (kept.size === 0) held.delete(groupId);
}

/**
 * Give each held picture back to its item's slot in `members` (a layout
 * change that landed on a picture layout): `restore` writes one picture into
 * one slot member, in the host's own model; members it returns unchanged
 * have no held picture. The pictures given back are released.
 */
export function restoreSmartArtHeldPictures<T>(
  groupId: string,
  members: readonly T[],
  slotItemId: (member: T) => string | null,
  restore: (member: T, picture: HeldPicture) => T,
): T[] {
  const kept = held.get(groupId);
  if (!kept || kept.size === 0) return [...members];
  const given: string[] = [];
  const next = members.map((member) => {
    const itemId = slotItemId(member);
    const picture = itemId ? kept.get(itemId) : undefined;
    if (!itemId || !picture) return member;
    given.push(itemId);
    return restore(member, picture);
  });
  releaseSmartArtPictures(groupId, given);
  return next;
}
