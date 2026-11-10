/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { loadCompatiblePdfJs } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";
import type { PDFWorker } from "pdfjs-dist";

/** One spare PDF.js preview worker per PDF window (the candidate pages are
 * drawn on it). Text preparation itself runs on the server's engine host
 * (`pdf-native/prepare`), so there is no source worker any more. */
export function createPdfNativeTextWorkerSlots(
  factories: {
    preview?: () => Promise<PDFWorker | undefined>;
  } = {},
) {
  let disposed = false;
  let preview: Promise<PDFWorker | undefined> | null = null;
  const createPreview =
    factories.preview ??
    (async () => {
      if (typeof Worker !== "function") return undefined;
      const api = await loadCompatiblePdfJs();
      // Keep the existing runtime's document-owned worker path when a host
      // does not expose an independently owned PDFWorker.
      if (typeof api.PDFWorker !== "function") return undefined;
      const worker = new api.PDFWorker();
      try {
        await worker.promise;
        return worker;
      } catch (error) {
        worker.destroy();
        throw error;
      }
    });
  const warmPreview = () => {
    if (disposed || preview) return;
    const pending = createPreview();
    preview = pending;
    void pending.catch(() => {
      if (preview === pending) preview = null;
    });
  };
  return {
    warm() {
      warmPreview();
    },
    async takePreview(): Promise<PDFWorker | undefined> {
      if (disposed) throw new DOMException("Text session closed", "AbortError");
      const pending = preview ?? createPreview();
      preview = null;
      queueMicrotask(warmPreview);
      const taken = await pending;
      if (disposed) {
        taken?.destroy();
        throw new DOMException("Text session closed", "AbortError");
      }
      return taken;
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      const idle = preview;
      preview = null;
      void idle?.then(
        (worker) => worker?.destroy(),
        () => undefined,
      );
    },
  };
}
