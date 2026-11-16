/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The update a room is applying from the network, readable while Yjs
 * applies it (its observers, `beforeObserverCalls`): what a peer's update
 * says the peer knew, which the document after it no longer tells. A peer
 * that deleted a block deleted every item of it it had; the items it had
 * not are deleted on arrival here just the same, but are absent from the
 * update's own delete set (`OrphanTextEdits`, lib/collab/collab-text-keys).
 */
import * as Y from "yjs";

const applying = new WeakMap<Y.Doc, Uint8Array>();

/** `Y.applyUpdate`, with `payload` readable during it (`remoteUpdateBeingApplied`). */
export function applyRemoteUpdate(doc: Y.Doc, payload: Uint8Array, origin: unknown): void {
  const outer = applying.get(doc);
  applying.set(doc, payload);
  try {
    Y.applyUpdate(doc, payload, origin);
  } finally {
    if (outer) applying.set(doc, outer);
    else applying.delete(doc);
  }
}

/** The update `doc` is applying from the network now, if any. */
export function remoteUpdateBeingApplied(doc: Y.Doc | null | undefined): Uint8Array | undefined {
  return doc ? applying.get(doc) : undefined;
}

/** Whether `payload`'s own delete set holds an item of `client` in [from, to). */
export function updateDeletes(payload: Uint8Array, client: number, from: number, to: number): boolean {
  if (to <= from) return false;
  const ranges = Y.decodeUpdate(payload).ds.clients.get(client) ?? [];
  return ranges.some((range) => range.clock < to && range.clock + range.len > from);
}
