/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const KEPT_LAYOUT_BYTES = 32 * 1024 * 1024;
export const KEPT_BYTES_PER_GLYPH = 460;

interface Kept {
  readonly bytes: number;
  readonly evict: () => void;
}

const order = new Map<object, Kept>();
let total = 0;

/** Keeps `token` (a layout) of `glyphs` glyphs; `evict` drops it from its cache when the budget gives it up. */
export function keepLayout(token: object, glyphs: number, evict: () => void): void {
  forgetLayout(token);
  const bytes = Math.max(1, glyphs) * KEPT_BYTES_PER_GLYPH;
  order.set(token, { bytes, evict });
  total += bytes;
  for (const [oldest, kept] of order) {
    if (total <= KEPT_LAYOUT_BYTES || oldest === token) break;
    order.delete(oldest);
    total -= kept.bytes;
    kept.evict();
  }
}

/** A kept layout used again: the most recent. */
export function touchLayout(token: object): void {
  const kept = order.get(token);
  if (!kept) return;
  order.delete(token);
  order.set(token, kept);
}

/** A layout its cache dropped by itself. */
export function forgetLayout(token: object): void {
  const kept = order.get(token);
  if (!kept) return;
  order.delete(token);
  total -= kept.bytes;
}

/** Tests: the layouts kept now and their bytes. */
export function keptLayoutsForTests(): { readonly layouts: number; readonly bytes: number } {
  return { layouts: order.size, bytes: total };
}
