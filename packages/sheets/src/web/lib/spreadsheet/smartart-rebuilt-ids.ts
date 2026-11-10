/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Which object ids a rebuilt sheet SmartArt keeps. A pure module of its own:
 * the window imports it, and `smartart-rebuild` (the engine host's) must
 * stay out of the browser bundle.
 */
/**
 * The object id each rebuilt member keeps: the old member's of the same role
 * and item (an item's shape stays the object the author selected and typed
 * into), then the ids left over in their order. By position alone a
 * relayout drawn from Office's definition -- arrows before boxes, as Office
 * orders them -- handed the edited box's id to an arrow.
 */
export function sheetSmartArtRebuiltIds(
  previous: ReadonlyArray<{ id: string; smartArtRole?: unknown; smartArtItemId?: unknown }>,
  rebuilt: ReadonlyArray<{ smartArtRole?: unknown; smartArtItemId?: unknown }>,
): Array<string | undefined> {
  const keyOf = (member: { smartArtRole?: unknown; smartArtItemId?: unknown }) =>
    `${typeof member.smartArtRole === "string" ? member.smartArtRole : ""}|${typeof member.smartArtItemId === "string" ? member.smartArtItemId : ""}`;
  const pool = new Map<string, string[]>();
  for (const member of previous) {
    const key = keyOf(member);
    pool.set(key, [...(pool.get(key) ?? []), member.id]);
  }
  const used = new Set<string>();
  const ids: Array<string | undefined> = rebuilt.map((member) => {
    const id = pool.get(keyOf(member))?.shift();
    if (id) used.add(id);
    return id;
  });
  const leftover = previous.map((member) => member.id).filter((id) => !used.has(id));
  return ids.map((id) => id ?? leftover.shift());
}
