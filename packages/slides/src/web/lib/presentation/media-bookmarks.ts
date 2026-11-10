/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Playback > Bookmarks (PowerPoint: Add Bookmark / Remove Bookmark, PPT-A157).
 * A bookmark marks a time in an audio or video clip, named "Bookmark N" in the
 * UI language; the file holds it as `p14:bmk@name,@time` (the server's writer).
 * PowerPoint refuses a second bookmark at the same time ("A bookmark already
 * exists for this time location.").
 */

import {
  MAX_PRESENTATION_MEDIA_BOOKMARKS,
  type PresentationMediaBookmark,
} from "@mrnote/core/web/lib/presentation-embed.ts";

export type AddPresentationMediaBookmarkResult =
  | {
      readonly ok: true;
      readonly bookmarks: readonly PresentationMediaBookmark[];
      readonly name: string;
    }
  | { readonly ok: false; readonly reason: "exists" | "full" };

/** The bookmark at exactly this time (ms), if any. */
export function presentationMediaBookmarkAt(
  bookmarks: readonly PresentationMediaBookmark[],
  timeMs: number,
): PresentationMediaBookmark | undefined {
  return bookmarks.find((bookmark) => bookmark.timeMs === timeMs);
}

/** Add a bookmark at `timeMs`, named `<base> N` with the first free N, kept in
 *  time order. */
export function addPresentationMediaBookmark(
  bookmarks: readonly PresentationMediaBookmark[],
  timeMs: number,
  baseName: string,
): AddPresentationMediaBookmarkResult {
  const time = Math.max(0, Math.round(timeMs));
  if (presentationMediaBookmarkAt(bookmarks, time)) {
    return { ok: false, reason: "exists" };
  }
  if (bookmarks.length >= MAX_PRESENTATION_MEDIA_BOOKMARKS) {
    return { ok: false, reason: "full" };
  }
  const names = new Set(bookmarks.map((bookmark) => bookmark.name));
  let index = 1;
  while (names.has(`${baseName} ${index}`)) index += 1;
  const name = `${baseName} ${index}`;
  const next = [...bookmarks, { name, timeMs: time }].sort(
    (left, right) => left.timeMs - right.timeMs,
  );
  return { ok: true, bookmarks: next, name };
}

/** Remove the named bookmark; the list unchanged when it is not there. */
export function removePresentationMediaBookmark(
  bookmarks: readonly PresentationMediaBookmark[],
  name: string,
): readonly PresentationMediaBookmark[] {
  return bookmarks.filter((bookmark) => bookmark.name !== name);
}
