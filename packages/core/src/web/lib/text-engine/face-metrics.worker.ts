/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
// Reads served faces' metrics off the page's thread (lib/text-engine/face-metrics-offthread.ts):
// the fetch, the WOFF2 decompression (fontkit, which the page then never loads) and the table

// own compile and set-up 83 ms, the decoding about 75 ms).
import type { SfntMetrics } from "../pdf/engine/native/sfnt/sfnt-reader";
import { faceMetricsForTransfer } from "./face-metrics-transfer";

type Request = { id: number; url: string };
type Response = { ready: true } | { id: number; metrics?: SfntMetrics; error?: string };

const scope = self as unknown as {
  postMessage(value: Response, transfer?: Transferable[]): void;
  onmessage: ((event: MessageEvent<Request>) => void) | null;
};

scope.postMessage({ ready: true });
scope.onmessage = (event) => {
  const { id, url } = event.data;
  void (async () => {
    try {
      const response = await fetch(url, { cache: "force-cache", credentials: "same-origin" });
      if (!response.ok) throw new Error(`font ${response.status}`);
      const { metrics, transfer } = await faceMetricsForTransfer(new Uint8Array(await response.arrayBuffer()));
      scope.postMessage({ id, metrics }, transfer);
    } catch (error) {
      scope.postMessage({ id, error: error instanceof Error ? error.message : String(error) });
    }
  })();
};
