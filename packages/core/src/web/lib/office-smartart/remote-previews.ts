"use client";

/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Gallery thumbnails, laid out by the server.
 *
 * Every SmartArt gallery (Insert, Change Layout, the Styles and Colors tiles)
 * draws the scene the engine would insert. The engine is the server's, so a
 * thumbnail asks for its scene: the thumbnails a gallery mounts together are
 * collected for one frame and asked in ONE request (`previews`), and every
 * answer is kept for the session, so reopening a gallery draws at once.
 * Until its answer lands a thumbnail draws its empty frame.
 */

import { useSyncExternalStore } from "react";
import { callEngine, canonicalJson } from "../engine-host/client";
import type { SmartArtPreviewRequest } from "../engine-host/smartart-types";
import type { OfficeSmartArtNode } from "./model";

export type SmartArtPreviewSpec = SmartArtPreviewRequest["previews"][number];

/** A refused or failed thumbnail is asked again after this long. */
const RETRY_AFTER_MS = 10_000;
/** The server's bound for a guest (registry.py ``guest_list_caps``): the
 * largest gallery (39 layouts) fits one request; more are split. */
const MAX_PER_REQUEST = 40;
const MAX_KEPT = 1_000;

type Entry =
  | { state: "queued" | "asked" }
  | { state: "ready"; nodes: OfficeSmartArtNode[] }
  | { state: "failed"; at: number };

const entries = new Map<string, Entry>();
const queue = new Map<string, SmartArtPreviewSpec>();
const listeners = new Set<() => void>();
let version = 0;
let flushTimer: ReturnType<typeof setTimeout> | null = null;

function changed(): void {
  version += 1;
  for (const listener of listeners) listener();
}

function flush(): void {
  flushTimer = null;
  const batch = [...queue.entries()].slice(0, MAX_PER_REQUEST);
  for (const [key] of batch) {
    queue.delete(key);
    entries.set(key, { state: "asked" });
  }
  if (queue.size) flushTimer = setTimeout(flush, 0);
  if (!batch.length) return;
  void callEngine("smartart", "previews", { previews: batch.map(([, spec]) => spec) }).then(
    (answer) => {
      batch.forEach(([key], index) => {
        const nodes = answer.previews[index];
        entries.set(key, nodes ? { state: "ready", nodes } : { state: "failed", at: Date.now() });
      });
      while (entries.size > MAX_KEPT) entries.delete(entries.keys().next().value as string);
      changed();
    },
    () => {
      const at = Date.now();
      for (const [key] of batch) entries.set(key, { state: "failed", at });
      changed();
    },
  );
}

function ask(key: string, spec: SmartArtPreviewSpec): void {
  queue.set(key, spec);
  entries.set(key, { state: "queued" });
  // One frame of collection: a gallery mounting fifty cards asks once.
  flushTimer ??= setTimeout(flush, 16);
}

/** The scene for this thumbnail, or `null` while the server has not drawn it. */
export function smartArtPreviewNodes(spec: SmartArtPreviewSpec | null): OfficeSmartArtNode[] | null {
  if (!spec) return null;
  const key = canonicalJson(spec);
  const entry = entries.get(key);
  if (entry?.state === "ready") return entry.nodes;
  if (entry?.state === "queued" || entry?.state === "asked") return null;
  if (entry?.state === "failed" && Date.now() - entry.at < RETRY_AFTER_MS) return null;
  if (typeof window !== "undefined") ask(key, spec);
  return null;
}

/** Whether the server could not answer this drawing (offline, failed). */
export function smartArtPreviewFailed(spec: SmartArtPreviewSpec | null): boolean {
  if (!spec) return false;
  return entries.get(canonicalJson(spec))?.state === "failed";
}

/**
 * `nodes` with ids of this component's own (`prefix`), so two drawings of one
 * scene on a page never share an SVG id.
 */
export function smartArtNodesWithOwnIds(
  nodes: readonly OfficeSmartArtNode[],
  prefix: string,
): OfficeSmartArtNode[] {
  const ids = new Map<string, string>();
  const own = (id: unknown) => {
    const key = String(id);
    let mapped = ids.get(key);
    if (!mapped) {
      mapped = `${prefix}-${ids.size}`;
      ids.set(key, mapped);
    }
    return mapped;
  };
  return nodes.map((node) => ({ ...node, id: own(node.id), groupId: own(node.groupId) }));
}

/**
 * {@link smartArtPreviewNodes} for a component, and whether the server could
 * not answer it: the drawing renders again when its answer lands. A `null`
 * spec asks nothing (a caller that already holds its scene).
 */
export function useSmartArtPreviewResult(
  spec: SmartArtPreviewSpec | null,
  prefix: string,
): { nodes: OfficeSmartArtNode[] | null; failed: boolean } {
  useSyncExternalStore(
    (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    () => version,
    () => 0,
  );
  const nodes = smartArtPreviewNodes(spec);
  return {
    nodes: nodes ? smartArtNodesWithOwnIds(nodes, prefix) : null,
    failed: !nodes && smartArtPreviewFailed(spec),
  };
}

/**
 * {@link smartArtPreviewNodes} for a component: the thumbnail draws again when
 * its answer lands. Ids are the component's own (`prefix`), so two thumbnails
 * of one layout never share an SVG id.
 */
export function useSmartArtPreview(
  spec: SmartArtPreviewSpec | null,
  prefix: string,
): OfficeSmartArtNode[] | null {
  return useSmartArtPreviewResult(spec, prefix).nodes;
}
