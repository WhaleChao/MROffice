/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A preserved Office diagram's slot PICTURES, drawn from the page's own source
 * package.
 *
 * `docs/office-graphics/SMARTART-IMPORT.md` §5.24 decided a preserved view
 * stores no media (four slots inline is half a Docs page's string budget), and
 * §5.68 projected each picture slot as the neutral slot with a receipt
 * (`pptx_diagram_picture_not_drawn`). §5.68 also named the next step: the
 * preserved RENDER asking the same `/pages/{id}/smartart-slot-media` route the
 * conversion already asks — one authorised, fingerprinted request per
 * diagram — and drawing the real photographs without a byte on the page.
 * This module is that step.
 *
 * The server records, on each host-neutral member of the view whose drawn
 * shape is a slot, `pictureSlot: { embed, fillRect, srcRect }`: the DATA
 * part's relationship id the route answers by, and where Office stretches the
 * picture inside the shape (`a:stretch/a:fillRect`, `a:srcRect`, thousandths
 * of a percent). The renderer clips the picture to the member's own painted
 * geometry, as a picture member of a Word group is clipped.
 *
 * A request that fails leaves the slot drawn as the slot (the page is
 * unchanged and nothing is retried in a loop); the diagram's own export is the
 * original package's, so the picture is never at risk.
 */

import {
  fetchSmartArtSlotMedia,
  smartArtSlotMediaAddress,
  type SmartArtSlotMedia,
  type SmartArtSlotMediaRequest,
} from "./slot-media-client";

/** Insets in thousandths of a percent, every side present. */
export interface SlotPictureRect {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

export interface PreservedSlotPicture {
  readonly embed: string;
  readonly fillRect: SlotPictureRect;
  readonly srcRect: SlotPictureRect;
}

const ZERO_RECT: SlotPictureRect = { l: 0, t: 0, r: 0, b: 0 };
const RECT_LIMIT = 1_000_000;

function readRect(value: unknown): SlotPictureRect | null {
  if (value === undefined) return ZERO_RECT;
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const sides = (["l", "t", "r", "b"] as const).map((side) => record[side] ?? 0);
  if (
    sides.some(
      (side) =>
        typeof side !== "number" ||
        !Number.isInteger(side) ||
        Math.abs(side) > RECT_LIMIT,
    )
  ) {
    return null;
  }
  const [l, t, r, b] = sides as number[];
  if (l + r >= 100_000 || t + b >= 100_000) return null;
  return { l, t, r, b };
}

/** The slot record of one host-neutral member, or `null`. */
export function readPreservedSlotPicture(value: unknown): PreservedSlotPicture | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const embed = record.embed;
  if (typeof embed !== "string" || !/^[A-Za-z0-9_.-]{1,64}$/.test(embed)) return null;
  const fillRect = readRect(record.fillRect);
  const srcRect = readRect(record.srcRect);
  if (!fillRect || !srcRect) return null;
  return { embed, fillRect, srcRect };
}

/**
 * Every slot picture of one preserved diagram, by member id, plus where the
 * diagram's media is addressed. `null` when the view has no addressable slot.
 */
export function preservedSmartArtSlotPictures(view: unknown): {
  readonly address: { readonly dataPart: string; readonly dataFingerprint: string };
  readonly slots: ReadonlyMap<string, PreservedSlotPicture>;
} | null {
  if (!view || typeof view !== "object" || Array.isArray(view)) return null;
  const members = (view as Record<string, unknown>).members;
  if (!Array.isArray(members)) return null;
  const slots = new Map<string, PreservedSlotPicture>();
  for (const member of members) {
    if (!member || typeof member !== "object") continue;
    const record = member as Record<string, unknown>;
    const slot = readPreservedSlotPicture(record.pictureSlot);
    if (slot && typeof record.memberId === "string" && record.memberId) {
      slots.set(record.memberId, slot);
    }
  }
  if (!slots.size) return null;
  const address = smartArtSlotMediaAddress(view);
  return address ? { address, slots } : null;
}

const MEDIA_CACHE_LIMIT = 32;
const mediaCache = new Map<string, Promise<SmartArtSlotMedia | null>>();

/**
 * One diagram's media, asked once per page/part/fingerprint and shared by
 * every render of it (a node view re-renders on every selection change).
 */
export function preservedSmartArtSlotMedia(
  request: SmartArtSlotMediaRequest,
  fetchMedia: (
    request: SmartArtSlotMediaRequest,
  ) => Promise<SmartArtSlotMedia | null> = fetchSmartArtSlotMedia,
): Promise<SmartArtSlotMedia | null> {
  const key = `${request.pageId}\u0000${request.dataPart}\u0000${request.dataFingerprint}`;
  const cached = mediaCache.get(key);
  if (cached) return cached;
  const pending = fetchMedia(request).catch(() => null);
  mediaCache.set(key, pending);
  if (mediaCache.size > MEDIA_CACHE_LIMIT) {
    const oldest = mediaCache.keys().next().value;
    if (oldest !== undefined) mediaCache.delete(oldest);
  }
  return pending;
}

/** Test hook: forget every cached answer. */
export function clearPreservedSmartArtSlotMediaCache(): void {
  mediaCache.clear();
}

/**
 * The box (percent of the member) and the image placement inside it for one
 * slot, as plain numbers so every host can spell them.
 *
 * `fillRect` insets the stretch target inside the shape box; `srcRect` keeps a
 * window of the picture, which is then stretched over that target — the same
 * arithmetic a Word group's picture member uses for its crop.
 */
export function slotPictureLayout(slot: PreservedSlotPicture): {
  readonly box: { left: number; top: number; width: number; height: number };
  readonly image: { left: number; top: number; width: number; height: number };
} {
  const fill = slot.fillRect;
  const box = {
    left: fill.l / 1000,
    top: fill.t / 1000,
    width: 100 - (fill.l + fill.r) / 1000,
    height: 100 - (fill.t + fill.b) / 1000,
  };
  const src = slot.srcRect;
  const keptWidth = Math.max(1e-6, 1 - (src.l + src.r) / 100_000);
  const keptHeight = Math.max(1e-6, 1 - (src.t + src.b) / 100_000);
  return {
    box,
    image: {
      left: ((-src.l / 100_000) / keptWidth) * 100,
      top: ((-src.t / 100_000) / keptHeight) * 100,
      width: (1 / keptWidth) * 100,
      height: (1 / keptHeight) * 100,
    },
  };
}
