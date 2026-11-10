"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Playable links to a deck's stored audio: a clip (Insert > Audio) or a
 * transition sound kept beside the deck as a DataItem
 * (presentation-embed's stored audio). The page's own route checks the
 * member may read the page and that the page names the asset, then mints a
 * short-lived private link; one is reused for a minute.
 */

import { useEffect, useMemo, useState } from "react";

import { apiGet } from "@mrnote/core/web/lib/api.ts";
import { getSafeExternalUrl } from "@mrnote/core/web/lib/security.ts";

export interface StoredPresentationAudio {
  readonly assetId: string;
  readonly mimeType: string;
  readonly sizeBytes: number;
}

const REUSE_MS = 60_000;
const links = new Map<string, { url: string; at: number }>();

function custodyKey(pageId: string, audio: StoredPresentationAudio): string {
  return `${pageId}\u0000${audio.assetId}\u0000${audio.mimeType}\u0000${audio.sizeBytes}`;
}

/** Forget every reused link (a test starts from none). */
export function clearStoredPresentationAudioLinks(): void {
  links.clear();
}

/** A link that plays the stored audio, or null when the page does not answer one. */
export async function storedPresentationAudioUrl(
  pageId: string,
  audio: StoredPresentationAudio,
): Promise<string | null> {
  const key = custodyKey(pageId, audio);
  const reusable = links.get(key);
  if (reusable && Date.now() - reusable.at <= REUSE_MS) return reusable.url;
  try {
    const answer = await apiGet<{
      asset_id?: unknown;
      playback_url?: unknown;
      mime_type?: unknown;
      size_bytes?: unknown;
    }>(
      `/api/v1/pages/${encodeURIComponent(pageId)}/presentation-audio/${encodeURIComponent(audio.assetId)}/playback`,
    );
    const raw =
      answer.asset_id === audio.assetId &&
      answer.mime_type === audio.mimeType &&
      answer.size_bytes === audio.sizeBytes &&
      typeof answer.playback_url === "string" &&
      /^https?:\/\//i.test(answer.playback_url)
        ? getSafeExternalUrl(answer.playback_url)
        : null;
    const url = raw && /^https?:/i.test(raw) ? raw : null;
    if (url) links.set(key, { url, at: Date.now() });
    return url;
  } catch {
    return null;
  }
}

/** Links for the stored audio given, by asset id, as they arrive. */
export function useStoredPresentationAudioUrls(
  pageId: string | null,
  audios: readonly StoredPresentationAudio[],
): ReadonlyMap<string, string> {
  const signature = audios
    .map((audio) => `${audio.assetId}\u0000${audio.mimeType}\u0000${audio.sizeBytes}`)
    .sort()
    .join("\u0001");
  const wanted = useMemo(
    () =>
      signature
        ? signature.split("\u0001").map((entry) => {
            const [assetId, mimeType, size] = entry.split("\u0000");
            return { assetId, mimeType, sizeBytes: Number(size) };
          })
        : [],
    [signature],
  );
  const [urls, setUrls] = useState<ReadonlyMap<string, string>>(() => new Map());
  useEffect(() => {
    if (!pageId || wanted.length === 0) return;
    let cancelled = false;
    void Promise.all(
      wanted.map(async (audio) => [audio.assetId, await storedPresentationAudioUrl(pageId, audio)] as const),
    ).then((entries) => {
      if (cancelled) return;
      const next = new Map<string, string>();
      for (const [assetId, url] of entries) if (url) next.set(assetId, url);
      setUrls(next);
    });
    return () => {
      cancelled = true;
    };
  }, [pageId, wanted]);
  return urls;
}
