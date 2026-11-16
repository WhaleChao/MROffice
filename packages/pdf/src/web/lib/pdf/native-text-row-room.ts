/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function pdfNativeTextRowRoom(
  box: Readonly<{ x: number; y: number; width: number; height: number }>,
  others: readonly Readonly<{ x: number; y: number; width: number; height: number; text?: unknown }>[],
): number | undefined {
  if (!(box.width > 0) || !(box.height > 0)) return undefined;
  const top = box.y;
  const bottom = box.y + box.height;
  let next = Number.POSITIVE_INFINITY;
  for (const other of others) {
    if (!(other.width > 0) || !(other.height > 0)) continue;
    // (White space the reader reports as an item of its own takes no room.)
    if (typeof other.text === "string" && other.text.trim() === "") continue;
    // On the row: its height and the box's share a third of the smaller one.
    const shared = Math.min(bottom, other.y + other.height) - Math.max(top, other.y);
    if (shared < Math.min(box.height, other.height) / 3) continue;
    // Past the box's middle: text the box already stands over is not what
    // it would grow into.
    if (other.x < box.x + box.width / 2) continue;
    next = Math.min(next, other.x);
  }
  if (!Number.isFinite(next)) return undefined;
  const room = next - box.x;
  return room > 0 ? room : undefined;
}
