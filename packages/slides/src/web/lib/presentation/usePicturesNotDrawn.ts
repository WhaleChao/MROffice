/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { useEffect, useRef, useState } from "react";

import { apiGet } from "@mrnote/core/web/lib/api.ts";
import { PAGE_PICTURE_NOT_DRAWN_EVENT } from "@mrnote/core/web/lib/notebook-events.ts";

/** A selection this long after the last read asks again: a room's notice is best effort. */
const READ_AGAIN_MS = 30_000;
const NOTHING: ReadonlySet<string> = new Set();

/**
 * The metafile pictures of a page the server could not draw, by the digest the
 * page names them with (INTERFACE-deck-parts 13.5). Read when a picture is
 * selected, not when the deck opens: only a selected block says anything of it.
 * This is a hint, so a read that fails leaves the picture as it is and the next
 * selection asks again.
 */
export function usePicturesNotDrawn(pageId: string | null, selectedSha: string | null): ReadonlySet<string> {
  const [held, setHeld] = useState<{ pageId: string | null; shas: ReadonlySet<string> }>({ pageId, shas: NOTHING });
  const lastRead = useRef<{ pageId: string | null; at: number } | null>(null);

  const add = useRef((forPage: string, shas: readonly string[]) => {
    setHeld((current) => {
      const known = current.pageId === forPage ? current.shas : NOTHING;
      if (shas.every((sha) => known.has(sha)) && current.pageId === forPage) return current;
      return { pageId: forPage, shas: new Set([...known, ...shas]) };
    });
  }).current;

  useEffect(() => {
    if (!pageId || !selectedSha) return;
    const last = lastRead.current;
    if (last?.pageId === pageId && Date.now() - last.at < READ_AGAIN_MS) return;
    lastRead.current = { pageId, at: Date.now() };
    let live = true;
    void apiGet<{ sha256s?: unknown }>(`/api/v1/pages/${encodeURIComponent(pageId)}/pictures-not-drawn`)
      .then((answer) => {
        if (!live || !Array.isArray(answer?.sha256s)) return;
        add(pageId, answer.sha256s.filter((sha): sha is string => typeof sha === "string"));
      })
      .catch(() => { if (lastRead.current?.pageId === pageId) lastRead.current = null; });
    return () => { live = false; };
  }, [add, pageId, selectedSha]);

  useEffect(() => {
    if (!pageId) return;
    const onNotDrawn = (event: Event) => {
      const detail = (event as CustomEvent<{ room?: unknown; sha256?: unknown }>).detail;
      const room = typeof detail?.room === "string" ? detail.room : "";
      if ((room !== pageId && !room.endsWith(`:${pageId}`)) || typeof detail?.sha256 !== "string") return;
      add(pageId, [detail.sha256]);
    };
    window.addEventListener(PAGE_PICTURE_NOT_DRAWN_EVENT, onNotDrawn);
    return () => window.removeEventListener(PAGE_PICTURE_NOT_DRAWN_EVENT, onNotDrawn);
  }, [add, pageId]);

  return held.pageId === pageId ? held.shas : NOTHING;
}
