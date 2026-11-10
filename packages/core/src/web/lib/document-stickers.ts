/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Draw ▸ Stickers: MRAI's own small set of stickers, drawn for MRAI (no
 * third-party artwork). Each is a die-cut look: the picture, a white rim
 * round it and a soft shadow under it, on a 120 x 120 grid. A sticker is put
 * on the page as an ordinary floating picture (a shape filled with it,
 * In Front Of Text), so it moves, sizes, turns, exports and saves as any
 * picture does.
 */

export const DOCUMENT_STICKER_IDS = [
  "star",
  "heart",
  "check",
  "important",
  "idea",
  "smile",
  "flower",
  "sun",
  "cloud",
  "pin",
  "chat",
  "sparkle",
] as const;
export type DocumentStickerId = (typeof DOCUMENT_STICKER_IDS)[number];

/** The pixels a sticker is drawn at (twice the size it is put on the page at). */
export const DOCUMENT_STICKER_RASTER_PX = 240;
/** How wide a sticker lands on the page, in CSS pixels at 100%. */
export const DOCUMENT_STICKER_PAGE_PX = 96;

/** Each sticker's outline: painted again, white and wide, as its rim. */
interface StickerArt {
  /** The silhouette: rim and shadow follow it. */
  readonly outline: string;
  /** The picture over the silhouette (the silhouette itself filled first). */
  readonly fill: string;
  readonly detail: string;
}

const ART: Readonly<Record<DocumentStickerId, StickerArt>> = {
  star: {
    outline: "M60 12l13.6 30.4 33 3.5-24.7 22.3 7 32.5L60 84.1 31.1 100.7l7-32.5L13.4 45.9l33-3.5Z",
    fill: "#ffc93c",
    detail: '<path d="M53.5 40.5l5-11" fill="none" stroke="#fff3c4" stroke-width="5" stroke-linecap="round"/>',
  },
  heart: {
    outline: "M60 104C28 82 12 64 12 42c0-15 11.5-26 25.5-26 9.7 0 17.6 5.4 22.5 13.6C64.9 21.4 72.8 16 82.5 16 96.5 16 108 27 108 42c0 22-16 40-48 62Z",
    fill: "#ff6b8b",
    detail: '<path d="M30 40c0-7 5-12 11-12" fill="none" stroke="#ffd0da" stroke-width="6" stroke-linecap="round"/>',
  },
  check: {
    outline: "M60 12a48 48 0 1 1 0 96 48 48 0 0 1 0-96Z",
    fill: "#34c38f",
    detail: '<path d="M38 61l15 15 30-32" fill="none" stroke="#fff" stroke-width="11" stroke-linecap="round" stroke-linejoin="round"/>',
  },
  important: {
    outline: "M60 10c5 0 8 2.6 10.6 7.2l38.2 68.4c4.6 8.3-.4 16.4-9.8 16.4H21c-9.4 0-14.4-8.1-9.8-16.4l38.2-68.4C52 12.6 55 10 60 10Z",
    fill: "#ff8a3d",
    detail: '<path d="M60 38v30" fill="none" stroke="#fff" stroke-width="11" stroke-linecap="round"/><circle cx="60" cy="86" r="6.5" fill="#fff"/>',
  },
  idea: {
    outline: "M60 10c21 0 37 15.6 37 35.4 0 13-6.8 21.4-13.4 28.6-4.3 4.7-6.6 9-6.6 14V92H43v-4c0-5-2.3-9.3-6.6-14C29.8 66.8 23 58.4 23 45.4 23 25.6 39 10 60 10ZM43 98h34v4c0 5.5-4.5 10-10 10H53c-5.5 0-10-4.5-10-10Z",
    fill: "#ffd84d",
    detail: '<path d="M43 98h34v4c0 5.5-4.5 10-10 10H53c-5.5 0-10-4.5-10-10Z" fill="#8a94a6"/><path d="M50 92V70l10 8 10-8v22" fill="none" stroke="#e0a800" stroke-width="5" stroke-linejoin="round"/><path d="M42 40c1.5-7 6.5-12 13-14" fill="none" stroke="#fff6c8" stroke-width="6" stroke-linecap="round"/>',
  },
  smile: {
    outline: "M60 12a48 48 0 1 1 0 96 48 48 0 0 1 0-96Z",
    fill: "#ffcf3f",
    detail: '<circle cx="44" cy="50" r="7" fill="#5b3a1a"/><circle cx="76" cy="50" r="7" fill="#5b3a1a"/><path d="M38 70c5 12 13 17 22 17s17-5 22-17" fill="none" stroke="#5b3a1a" stroke-width="7" stroke-linecap="round"/><circle cx="33" cy="68" r="6" fill="#ff9a8a" opacity=".7"/><circle cx="87" cy="68" r="6" fill="#ff9a8a" opacity=".7"/>',
  },
  flower: {
    outline: "M60 14c10 0 17 8 17 17.5 9-4.5 19.5-1 23 8.4 3.6 9.5-1.6 18.7-10.2 21.6 8.6 3 13.8 12.1 10.2 21.6-3.5 9.4-14 12.9-23 8.4 0 9.5-7 17.5-17 17.5S43 101 43 91.5c-9 4.5-19.5 1-23-8.4-3.6-9.5 1.6-18.6 10.2-21.6C21.6 58.6 16.4 49.4 20 39.9c3.5-9.4 14-12.9 23-8.4C43 22 50 14 60 14Z",
    fill: "#c99bff",
    detail: '<circle cx="60" cy="61" r="15" fill="#ffd84d"/><circle cx="55" cy="56" r="4" fill="#fff3c4"/>',
  },
  sun: {
    outline: "M60 8l9 16 17.4-6.4-.6 18.6 18.2 3.6-11.4 14.6L104 69l-17.6 5.8 1.2 18.6-17-7.4L60 102 51.4 86l-17 7.4 1.2-18.6L18 69l11.4-14.6L18 39.8l18.2-3.6-.6-18.6L53 24Z",
    fill: "#ffb020",
    detail: '<circle cx="60" cy="56" r="24" fill="#ffd84d"/><path d="M48 50a14 14 0 0 1 12-10" fill="none" stroke="#fff3c4" stroke-width="5" stroke-linecap="round"/>',
  },
  cloud: {
    outline: "M34 94c-13 0-23-9.6-23-22 0-11 8-20 18.6-21.7C32 36 44 26 58 26c13 0 24 8.2 27.6 20.2C99 47 109 57 109 70c0 13.3-10.7 24-24 24Z",
    fill: "#8fd0ff",
    detail: '<path d="M30 68c0-6 4-10 9-11" fill="none" stroke="#e3f4ff" stroke-width="6" stroke-linecap="round"/><circle cx="48" cy="72" r="4.5" fill="#2d6a9a"/><circle cx="72" cy="72" r="4.5" fill="#2d6a9a"/><path d="M54 82c3.5 3 8.5 3 12 0" fill="none" stroke="#2d6a9a" stroke-width="4" stroke-linecap="round"/>',
  },
  pin: {
    outline: "M60 8c20 0 36 15.4 36 35 0 26-36 67-36 67S24 69 24 43C24 23.4 40 8 60 8Z",
    fill: "#ff5a5f",
    detail: '<circle cx="60" cy="43" r="14" fill="#fff"/><circle cx="60" cy="43" r="6" fill="#ff5a5f"/>',
  },
  chat: {
    outline: "M26 16h68c8.8 0 16 7.2 16 16v38c0 8.8-7.2 16-16 16H58l-24 20V86h-8c-8.8 0-16-7.2-16-16V32c0-8.8 7.2-16 16-16Z",
    fill: "#5b8def",
    detail: '<circle cx="40" cy="51" r="7" fill="#fff"/><circle cx="60" cy="51" r="7" fill="#fff"/><circle cx="80" cy="51" r="7" fill="#fff"/>',
  },
  sparkle: {
    outline: "M52 10c3 22 9 28 30 31-21 3-27 9-30 31-3-22-9-28-30-31 21-3 27-9 30-31ZM88 62c2 13 6 17 18 19-12 2-16 6-18 19-2-13-6-17-18-19 12-2 16-6 18-19Z",
    fill: "#ff7ac6",
    detail: '<path d="M88 62c2 13 6 17 18 19-12 2-16 6-18 19-2-13-6-17-18-19 12-2 16-6 18-19Z" fill="#7a6cff"/>',
  },
};

/** A sticker as an SVG document of its own, at `sizePx` square. */
export function documentStickerSvg(id: DocumentStickerId, sizePx = DOCUMENT_STICKER_RASTER_PX): string {
  const art = ART[id];
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${sizePx}" height="${sizePx}" viewBox="0 0 120 120">` +
    // The shadow, then the white rim, both following the silhouette.
    `<path d="${art.outline}" transform="translate(1.5 3)" fill="#1f2937" opacity=".16" stroke="#1f2937" stroke-width="10" stroke-linejoin="round"/>` +
    `<path d="${art.outline}" fill="#ffffff" stroke="#ffffff" stroke-width="10" stroke-linejoin="round"/>` +
    `<path d="${art.outline}" fill="${art.fill}"/>` +
    art.detail +
    `</svg>`
  );
}

/** The sticker as an image URL (the menu's previews). */
export function documentStickerDataUrl(id: DocumentStickerId, sizePx = DOCUMENT_STICKER_RASTER_PX): string {
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(documentStickerSvg(id, sizePx))}`;
}

/** The sticker as a picture file: the page keeps it as a PNG, as any picture. */
export function documentStickerFile(id: DocumentStickerId): File {
  return new File([documentStickerSvg(id)], `sticker-${id}.svg`, { type: "image/svg+xml" });
}

export function isDocumentStickerId(value: unknown): value is DocumentStickerId {
  return (DOCUMENT_STICKER_IDS as readonly unknown[]).includes(value);
}
