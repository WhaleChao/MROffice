/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  readSmartArtGraphicFormat,
  SMARTART_RESET_GRAPHIC_FORMAT,
  smartArtGraphicFormatRecord,
  type SmartArtGraphicFormatChange,
} from "@mrnote/core/web/lib/spreadsheet/smartart-background.ts";
import { readSmartArtBlockMarker } from "@mrnote/core/web/lib/office-smartart/model.ts";

type Block = { readonly id: string };

/**
 * The members with the graphic's formatting changed: a fill colour, or `null`
 * for No Fill (the historical call), or a {@link SmartArtGraphicFormatChange};
 * `{ clear: true }` takes all of it off. Null when nothing
 * changes.
 */
export function presentationBlocksWithSmartArtBackground<T extends Block>(
  blocks: readonly T[],
  memberIds: ReadonlySet<string>,
  change: string | null | SmartArtGraphicFormatChange | { readonly clear: true },
): T[] | null {
  const asChange: SmartArtGraphicFormatChange | { readonly clear: true } =
    change === null || typeof change === "string" ? { fill: change } : change;
  if ("fill" in asChange && typeof asChange.fill === "string" && !/^#[0-9a-f]{6}$/i.test(asChange.fill)) {
    return null;
  }
  let changed = false;
  const next = blocks.map((block) => {
    if (!memberIds.has(block.id)) return block;
    const current = (block as { smartArtBackground?: unknown }).smartArtBackground;
    const value = "clear" in asChange ? undefined : smartArtGraphicFormatRecord(current, asChange);
    if (JSON.stringify(value ?? null) === JSON.stringify(current ?? null)) return block;
    changed = true;
    const copy = { ...block } as T & { smartArtBackground?: unknown };
    if (value) copy.smartArtBackground = value;
    else delete copy.smartArtBackground;
    return copy;
  });
  return changed ? next : null;
}

/**
 * A relayout (an item added or deleted, Change Layout, a restyle) rebuilds a
 * graphic's members: the graphic's own formatting rides over to the members
 * that came back without it.
 */
export function presentationSmartArtBackgroundsCarried<T extends Block>(
  previous: readonly T[],
  next: readonly T[],
): T[] {
  const backgrounds = new Map<string, unknown>();
  for (const block of previous) {
    const groupId = readSmartArtBlockMarker(block)?.groupId;
    const value = (block as { smartArtBackground?: unknown }).smartArtBackground;
    if (groupId && value && readSmartArtGraphicFormat(block) && !backgrounds.has(groupId)) {
      backgrounds.set(groupId, value);
    }
  }
  if (backgrounds.size === 0) return [...next];
  return next.map((block) => {
    const groupId = readSmartArtBlockMarker(block)?.groupId;
    if (!groupId || !backgrounds.has(groupId) || "smartArtBackground" in block) return block;
    return { ...block, smartArtBackground: backgrounds.get(groupId) };
  });
}

/**
 * Reset Graphic's rebuild: the graphic's background colour rides over to the
 * rebuilt members and its own line and effect come off
 * ({@link SMARTART_RESET_GRAPHIC_FORMAT}, live Word audit WUA 260).
 */
export function presentationSmartArtBackgroundsAfterReset<T extends Block>(
  previous: readonly T[],
  next: readonly T[],
): T[] {
  const kept = new Map<string, Record<string, unknown> | undefined>();
  for (const block of previous) {
    const groupId = readSmartArtBlockMarker(block)?.groupId;
    if (!groupId || kept.has(groupId)) continue;
    const value = (block as { smartArtBackground?: unknown }).smartArtBackground;
    kept.set(groupId, value ? smartArtGraphicFormatRecord(value, SMARTART_RESET_GRAPHIC_FORMAT) : undefined);
  }
  return next.map((block) => {
    const groupId = readSmartArtBlockMarker(block)?.groupId;
    if (!groupId || !kept.has(groupId)) return block;
    const value = kept.get(groupId);
    const current = (block as { smartArtBackground?: unknown }).smartArtBackground;
    if (JSON.stringify(value ?? null) === JSON.stringify(current ?? null)) return block;
    const copy = { ...block } as T & { smartArtBackground?: unknown };
    if (value) copy.smartArtBackground = value;
    else delete copy.smartArtBackground;
    return copy;
  });
}
