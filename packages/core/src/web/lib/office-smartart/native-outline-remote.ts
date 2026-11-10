"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A preserved Office diagram's outline, asked of the SmartArt engine.
 *
 * The data part (`dgm:dataModel`) is DiagramML; the reader that walks it
 * (`./native-outline`) runs on the server, like every other Office reader.
 * The browser needs the outline for 转换为可编辑 only: which picture each item
 * carries (the conversion fetches them), and how many items at which levels
 * (a layout that cannot hold them is greyed before the press). Drawing a
 * preserved diagram never needs it.
 *
 * Two ways in, one cache keyed by the data part itself:
 * - {@link readNativeSmartArtOutline} for a command, which can wait;
 * - {@link cachedNativeSmartArtOutline} for a ribbon render, which cannot: an
 *   outline the server already gave comes back at once; one it has not given
 *   yet starts the request and reads `undefined` (nothing known yet, so
 *   nothing is greyed), and the render computes again when it lands
 *   ({@link useNativeSmartArtOutlineVersion}).
 * A failed request (offline, refused) reads as `undefined` too and is asked
 * again after a few seconds; the conversion itself then answers "unavailable",
 * as it does when the diagram's pictures cannot be fetched.
 */

import { useSyncExternalStore } from "react";
import { callEngine } from "../engine-host/client";
import type { NativeSmartArtOutlineItem } from "./native-outline";

export type NativeSmartArtOutline = readonly NativeSmartArtOutlineItem[];

const RETRY_AFTER_MS = 5_000;
const MAX_ENTRIES = 64;

type Entry =
  | { state: "pending"; promise: Promise<NativeSmartArtOutline | null> }
  | { state: "ready"; outline: NativeSmartArtOutline | null }
  | { state: "failed"; at: number };

const entries = new Map<string, Entry>();
const listeners = new Set<() => void>();
let version = 0;

function changed(): void {
  version += 1;
  for (const listener of listeners) listener();
}

function remember(dataXml: string, entry: Entry): void {
  entries.delete(dataXml);
  entries.set(dataXml, entry);
  while (entries.size > MAX_ENTRIES) entries.delete(entries.keys().next().value as string);
}

function ask(dataXml: string): Promise<NativeSmartArtOutline | null> {
  const promise = callEngine("smartart", "native.outline", { dataXml }, { cache: true }).then(
    (answer) => {
      remember(dataXml, { state: "ready", outline: answer.outline });
      changed();
      return answer.outline;
    },
    (error: unknown) => {
      remember(dataXml, { state: "failed", at: Date.now() });
      changed();
      throw error;
    },
  );
  remember(dataXml, { state: "pending", promise });
  return promise;
}

/**
 * The outline of this data part, or `null` when the reader cannot prove one
 * (the caller keeps the diagram preserved). Rejects with the engine's error
 * (`EngineUnavailableError` offline) when the server cannot be asked.
 */
export async function readNativeSmartArtOutline(
  dataXml: string,
): Promise<NativeSmartArtOutline | null> {
  const entry = entries.get(dataXml);
  if (entry?.state === "ready") return entry.outline;
  if (entry?.state === "pending") return entry.promise;
  return ask(dataXml);
}

/**
 * The outline if the server has answered it: the outline, `null` (no outline
 * can be read), or `undefined` while it is not known (asked now, or failed and
 * not yet due again).
 */
export function cachedNativeSmartArtOutline(
  dataXml: string,
): NativeSmartArtOutline | null | undefined {
  const entry = entries.get(dataXml);
  if (entry?.state === "ready") return entry.outline;
  if (entry?.state === "pending") return undefined;
  if (entry?.state === "failed" && Date.now() - entry.at < RETRY_AFTER_MS) return undefined;
  void ask(dataXml).catch(() => undefined);
  return undefined;
}

/** Changes whenever an outline lands, so a ribbon can compute its greying again. */
export function useNativeSmartArtOutlineVersion(): number {
  return useSyncExternalStore(
    (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    () => version,
    () => 0,
  );
}
