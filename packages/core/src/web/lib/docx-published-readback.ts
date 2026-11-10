/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The page's own verdicts on the document its confirmation published.
 *
 * Confirming a generated document answers as soon as the commit is durable;
 * the readback of the stored package settles a moment later and the page
 * carries the outcome (`PageOut.docx_published_readback`). A window that
 * opened on the freshly published page therefore sees `pending` first. The
 * same holds for the real-editor recheck that follows a publication
 * (`PageOut.export_fidelity.editor_recheck`), which takes a minute or more.
 * These hooks follow such a marker with a few spaced re-reads until it
 * settles, so the ribbon can change its word without a reload.
 */
import { useEffect, useState } from "react";

import {
  getPage,
  type DocxPublishedReadbackRead,
  type PageExportFidelityRead,
  type PageItem,
} from "./notebook-sdk";

/** Spaced out: the readback usually lands within the first delay. */
const FOLLOW_UP_DELAYS_MS = [1500, 2500, 4000, 6000, 10000, 15000, 20000];
/** The editor recheck waits for an editor turn and then drives a browser. */
const RECHECK_FOLLOW_UP_DELAYS_MS = [10000, 15000, 20000, 30000, 30000, 45000, 60000, 60000];

interface FollowedMarker<T> {
  pageId: string | null;
  marker: T | null;
}

function useFollowedPageMarker<T>(
  pageId: string | null | undefined,
  read: T | null | undefined,
  options: {
    /** Changes only when the page read carries another marker. */
    seedKey: string;
    pending: (marker: T | null) => boolean;
    select: (page: PageItem) => T | null;
    delays: readonly number[];
  },
): T | null {
  const key = pageId || null;
  const [followed, setFollowed] = useState<FollowedMarker<T>>({ pageId: null, marker: null });
  const { seedKey, pending: isPending, select, delays } = options;

  // The page read seeds the marker; a later read (another publication) reseeds.
  useEffect(() => {
    setFollowed({ pageId: key, marker: key && read ? read : null });
    // `read` is re-derived on every render; `seedKey` carries its identity.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key, seedKey]);

  const marker = followed.pageId === key ? followed.marker : null;
  const pending = isPending(marker);

  useEffect(() => {
    if (!key || !pending) return;
    let disposed = false;
    let attempt = 0;
    let timer: number | undefined;
    let controller: AbortController | null = null;
    const follow = async () => {
      if (disposed) return;
      controller = new AbortController();
      try {
        const remote = await getPage(key, { signal: controller.signal });
        if (disposed || remote.id !== key) return;
        const next = select(remote);
        if (!isPending(next)) {
          setFollowed({ pageId: key, marker: next });
          return;
        }
      } catch {
        if (disposed) return;
      }
      // Past the last delay the server's own staleness rule takes over.
      if (attempt >= delays.length) return;
      timer = window.setTimeout(follow, delays[attempt++]);
    };
    timer = window.setTimeout(follow, delays[attempt++]);
    return () => {
      disposed = true;
      window.clearTimeout(timer);
      controller?.abort();
    };
    // `select`, `isPending` and `delays` are module-level and never change.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key, pending]);

  return marker;
}

const readbackPending = (marker: DocxPublishedReadbackRead | null) =>
  marker?.status === "pending";
const selectReadback = (page: PageItem) => page.docx_published_readback ?? null;

export function useDocxPublishedReadback(
  pageId: string | null | undefined,
  read: DocxPublishedReadbackRead | null | undefined,
): DocxPublishedReadbackRead | null {
  return useFollowedPageMarker(pageId, read, {
    seedKey: `${read?.status ?? ""}|${read?.queued_at ?? ""}`,
    pending: readbackPending,
    select: selectReadback,
    delays: FOLLOW_UP_DELAYS_MS,
  });
}

const recheckPending = (marker: PageExportFidelityRead | null) =>
  marker?.editor_recheck?.status === "pending";
const selectExportFidelity = (page: PageItem) => page.export_fidelity ?? null;

/**
 * A generated document's check record, followed while its editor recheck is
 * still running, so "check before editing" appears without a reload.
 */
export function useDocsExportFidelity(
  pageId: string | null | undefined,
  read: PageExportFidelityRead | null | undefined,
): PageExportFidelityRead | null {
  return useFollowedPageMarker(pageId, read, {
    seedKey: [
      read?.status ?? "",
      read?.notices?.length ?? 0,
      read?.editor_recheck?.status ?? "",
      read?.editor_recheck?.queued_at ?? "",
    ].join("|"),
    pending: recheckPending,
    select: selectExportFidelity,
    delays: RECHECK_FOLLOW_UP_DELAYS_MS,
  });
}
