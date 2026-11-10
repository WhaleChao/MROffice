/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef, useState } from "react";

import { DECK_PART_ROOM_NOT_SAVED_EVENT } from "@mrnote/core/web/lib/notebook-events.ts";
import type { DeckPartRoomUnsaved, PageItem } from "@mrnote/core/web/lib/notebook-sdk.ts";

/** While a room is listed, the page is read again this often: the entry leaves
 * when any later save of that room lands, and nothing tells this window. */
const READ_AGAIN_MS = 30_000;
const NOTHING: readonly DeckPartRoomUnsaved[] = [];

/**
 * The part rooms of a deck the server could not save for their members, as the
 * page says (INTERFACE-deck-parts). Follows every page this window holds, reads
 * the page when the room says one was refused, and keeps reading while any is
 * listed. A read that fails says nothing either way.
 */
export function useDeckRoomsUnsaved(
  pageId: string | null,
  held: readonly DeckPartRoomUnsaved[] | undefined,
  readPage: (pageId: string) => Promise<Pick<PageItem, "id" | "deck_part_rooms_unsaved">>,
): readonly DeckPartRoomUnsaved[] {
  const [state, setState] = useState<{ pageId: string | null; rooms: readonly DeckPartRoomUnsaved[] }>({ pageId, rooms: NOTHING });
  const readRef = useRef(readPage);
  readRef.current = readPage;

  // Each page read answers with its own array: only a different list is news.
  const take = useRef((forPage: string | null, rooms: readonly DeckPartRoomUnsaved[]) => {
    setState((current) => (current.pageId === forPage && JSON.stringify(current.rooms) === JSON.stringify(rooms)
      ? current : { pageId: forPage, rooms }));
  }).current;

  useEffect(() => {
    if (held !== undefined) take(pageId, held);
  }, [held, pageId, take]);

  useEffect(() => {
    if (!pageId) return;
    let live = true;
    const read = () => {
      void readRef.current(pageId).then((page) => {
        if (live && page.id === pageId) take(pageId, page.deck_part_rooms_unsaved ?? NOTHING);
      }).catch(() => undefined);
    };
    const onNotSaved = (event: Event) => {
      const room = (event as CustomEvent<{ room?: unknown }>).detail?.room;
      if (room === pageId || (typeof room === "string" && room.endsWith(`:${pageId}`))) read();
    };
    window.addEventListener(DECK_PART_ROOM_NOT_SAVED_EVENT, onNotSaved);
    const anyListed = state.pageId === pageId && state.rooms.length > 0;
    const timer = anyListed ? window.setInterval(read, READ_AGAIN_MS) : null;
    return () => {
      live = false;
      window.removeEventListener(DECK_PART_ROOM_NOT_SAVED_EVENT, onNotSaved);
      if (timer !== null) window.clearInterval(timer);
    };
  }, [pageId, state, take]);

  return state.pageId === pageId ? state.rooms : NOTHING;
}
