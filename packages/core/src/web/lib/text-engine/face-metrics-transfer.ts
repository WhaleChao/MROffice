/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A face's metrics as they cross from the metrics worker
 * (face-metrics.worker.ts) to the page: the same `readSfntMetrics` reading,
 * with each raw table (kerning, variation) copied out of the decompressed
 * program it was a view into, so the message carries the tables and not the
 * whole program. The page receives what `readSfntMetrics` returns on the
 * page, field for field (tests: face-metrics-offthread.test.ts).
 */
import { readSfntMetrics, type SfntMetrics } from "../pdf/engine/native/sfnt/sfnt-reader";

function ownCopies(tables: ReadonlyMap<string, Uint8Array> | undefined, transfer: ArrayBuffer[]): ReadonlyMap<string, Uint8Array> | undefined {
  if (!tables) return undefined;
  const copies = new Map<string, Uint8Array>();
  for (const [tag, view] of tables) {
    const copy = new Uint8Array(view);
    copies.set(tag, copy);
    transfer.push(copy.buffer);
  }
  return copies;
}

/** The metrics of `bytes` and the buffers a postMessage can hand over without copying. */
export async function faceMetricsForTransfer(bytes: Uint8Array): Promise<{ metrics: SfntMetrics; transfer: ArrayBuffer[] }> {
  const read = await readSfntMetrics(bytes);
  const transfer: ArrayBuffer[] = [];
  const kerningTables = ownCopies(read.kerningTables, transfer);
  const variationTables = ownCopies(read.variationTables, transfer);
  const metrics: SfntMetrics = {
    ...read,
    ...(kerningTables ? { kerningTables } : {}),
    ...(variationTables ? { variationTables } : {}),
  };
  return { metrics, transfer };
}
