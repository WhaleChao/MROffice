/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PinnedAsset {
  /** The file's same-origin path. */
  readonly url: string;
  readonly byteLength: number;
  /** Lower-case hex. */
  readonly sha256: string;
}

type AssetFetch = (input: RequestInfo | URL, init?: RequestInit) => Promise<Response>;

const CACHE_NAME = "mrnote-pinned-assets-v1";

async function digestHex(bytes: Uint8Array<ArrayBuffer>): Promise<string | null> {
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) return null;
  const digest = await subtle.digest("SHA-256", bytes.buffer);
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}

async function verified(asset: PinnedAsset, value: ArrayBuffer): Promise<Uint8Array<ArrayBuffer> | null> {
  if (value.byteLength !== asset.byteLength) return null;
  const bytes = new Uint8Array(value);
  return (await digestHex(bytes)) === asset.sha256 ? bytes : null;
}

function keyOf(asset: PinnedAsset): string | null {
  try {
    const origin = (globalThis as { location?: { origin?: string } }).location?.origin;
    if (!origin || !/^https?:/u.test(origin) || !asset.url.startsWith("/")) return null;
    return `${origin}${asset.url}?pinned=${asset.sha256}`;
  } catch {
    return null;
  }
}

async function openCache(): Promise<Cache | null> {
  try {
    const storage = (globalThis as { caches?: CacheStorage }).caches;
    return storage ? await storage.open(CACHE_NAME) : null;
  } catch {
    return null;
  }
}

/**
 * `asset`'s verified bytes (the caller's own copy): from this browser's copy
 * when it holds one that passes the check, else downloaded with `fetcher`,
 * checked, and kept. Null when they cannot be had or are not the pinned
 * bytes. Never throws.
 */
export async function loadPinnedAssetBytes(
  asset: PinnedAsset,
  fetcher: AssetFetch,
  signal?: AbortSignal,
): Promise<Uint8Array<ArrayBuffer> | null> {
  const key = keyOf(asset);
  const cache = key ? await openCache() : null;
  if (cache && key) {
    try {
      const kept = await cache.match(key);
      if (kept) {
        const bytes = await verified(asset, await kept.arrayBuffer());
        if (bytes) return bytes;
        await cache.delete(key);
      }
    } catch {
      // An unreadable copy is no copy: downloaded below.
    }
  }
  try {
    const response = await fetcher(asset.url, signal ? { signal } : undefined);
    if (!response.ok) return null;
    const bytes = await verified(asset, await response.arrayBuffer());
    if (!bytes) return null;
    if (cache && key) {
      // Kept for the next page of this browser; a full disk only costs that.
      void (async () => {
        try {
          // An older pinning of the same path is not read again.
          for (const request of await cache.keys()) {
            const url = new URL(request.url);
            if (url.pathname === asset.url && request.url !== key) await cache.delete(request);
          }
          await cache.put(key, new Response(bytes.slice(), { headers: { "content-type": "application/octet-stream" } }));
        } catch {
          // Not kept.
        }
      })();
    }
    return bytes;
  } catch {
    return null;
  }
}
