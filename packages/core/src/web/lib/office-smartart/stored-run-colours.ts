/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { apiGet } from "../api";
import type { NativeSmartArtOutlineItem } from "@/lib/office-smartart/native-outline";
import { readNativeSmartArtOutline } from "./native-outline-remote";

const answers = new Map<string, Promise<ReadonlyMap<string, NativeSmartArtOutlineItem> | null>>();
const MAX_KEPT = 64;

function record(value: unknown): Readonly<Record<string, unknown>> | null {
  return value && typeof value === "object" && !Array.isArray(value) ? (value as Record<string, unknown>) : null;
}

async function storedItems(pageId: string, part: string, sha256: string): Promise<ReadonlyMap<string, NativeSmartArtOutlineItem> | null> {
  const query = new URLSearchParams({ sha256, part, kind: "data" });
  try {
    const answer = await apiGet<{ readonly sha256?: unknown; readonly xml?: unknown }>(
      `/api/v1/pages/${encodeURIComponent(pageId)}/smartart-layout-part?${query}`,
    );
    if (!answer || answer.sha256 !== sha256 || typeof answer.xml !== "string") return null;
    // Read by the engine host, as every preserved diagram's outline is (the
    // reader is server-only: it stays out of the page's bundle).
    const outline = await readNativeSmartArtOutline(answer.xml);
    return outline ? new Map(outline.map((item) => [item.id, item])) : null;
  } catch {
    return null;
  }
}

/**
 * `outline` with each item's runs as its file states them, theme colours
 * included, when the page's reading predates them; `outline` itself when no
 * item has runs, they already carry their theme colours, or the file cannot
 * be read.
 */
export async function outlineWithStoredRunColours(
  pageId: string | null | undefined,
  outline: readonly NativeSmartArtOutlineItem[],
  slotMediaView: Readonly<Record<string, unknown>> | null | undefined,
): Promise<readonly NativeSmartArtOutlineItem[]> {
  if (!pageId || !outline.some((item) => item.textRuns?.length)) return outline;
  // Read since the reading kept them: nothing to ask.
  if (outline.some((item) => item.textRuns?.some((run) => run.style?.colorScheme))) return outline;
  const parts = record(slotMediaView?.parts);
  const digests = record(slotMediaView?.fingerprints);
  const part = parts?.data;
  const sha256 = digests?.data;
  if (typeof part !== "string" || typeof sha256 !== "string" || !/^[a-f0-9]{64}$/.test(sha256)) return outline;
  const key = `${pageId}|${sha256}`;
  let pending = answers.get(key);
  if (!pending) {
    if (answers.size >= MAX_KEPT) answers.clear();
    pending = storedItems(pageId, part, sha256);
    answers.set(key, pending);
  }
  const stored = await pending;
  if (!stored) return outline;
  let changed = false;
  const next = outline.map((item) => {
    const own = stored.get(item.id);
    const runs = own?.textRuns;
    if (!runs?.some((run) => run.style?.colorScheme) || runs.map((run) => run.text).join("") !== item.text) return item;
    changed = true;
    return { ...item, textRuns: runs };
  });
  return changed ? next : outline;
}
