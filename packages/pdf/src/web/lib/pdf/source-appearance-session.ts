/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { MAX_MRAI_NATIVE_SOURCE_BYTES, sha256PdfBytes } from "./engine/native/content-stream-transaction";
import { loadCompatiblePdfJs, PDF_JS_DOCUMENT_ASSETS, type CompatiblePdfJsApi } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";
import { pdfPageSourceAppearanceFromOperators, type PdfPageSourceAppearance, type PdfSourceFont } from "./source-page-appearance";

export type { PdfPageSourceAppearance } from "./source-page-appearance";

const RECENT_PAGES = 6;

/**
 * A read-only view of one exact PDF source: its pages' characters, read with
 * the reader's own engine (PDF.js), so what is returned is what the reader
 * draws. It exposes no save or writer method; the native transaction and the
 * server's validation keep owning publication.
 *
 * `load` and `documentOptions` exist for hosts without the browser's
 * same-origin PDF.js assets (the Node tests).
 */
export async function createPdfSourceAppearanceSession(
  sourceBytes: Uint8Array,
  expectedSha256: string,
  options: {
    load?: () => Promise<CompatiblePdfJsApi>;
    documentOptions?: Readonly<Record<string, unknown>>;
  } = {},
) {
  if (!sourceBytes.length || sourceBytes.length > MAX_MRAI_NATIVE_SOURCE_BYTES) {
    throw new Error("PDF appearance source exceeds its byte budget.");
  }
  const bytes = Uint8Array.from(sourceBytes);
  let task: ReturnType<CompatiblePdfJsApi["getDocument"]> | null = null;
  try {
    if ((await sha256PdfBytes(bytes)) !== expectedSha256) throw new Error("PDF appearance source hash mismatch.");
    const pdfjs = await (options.load ?? loadCompatiblePdfJs)();
    // Fonts are read for their metrics only: nothing is drawn here, so no
    // font face is installed (a worker has no document to install one in).
    task = pdfjs.getDocument({
      data: bytes,
      ...PDF_JS_DOCUMENT_ASSETS,
      useWorkerFetch: true,
      disableFontFace: true,
      useSystemFonts: false,
      ...options.documentOptions,
    } as Parameters<CompatiblePdfJsApi["getDocument"]>[0]);
    const owned = task;
    const document = await owned.promise.catch((error: unknown) => {
      if ((error as { name?: string } | null)?.name === "PasswordException") {
        throw new Error("PDF appearance requires the document password workflow.");
      }
      throw error;
    });
    const ops = pdfjs.OPS as unknown as Readonly<Record<string, number>>;
    const annotationMode = pdfjs.AnnotationMode?.DISABLE ?? 0;
    let disposed = false;
    const appearances = new Map<number, Promise<PdfPageSourceAppearance>>();
    const read = async (pageNumber: number): Promise<PdfPageSourceAppearance> => {
      const page = await document.getPage(pageNumber);
      try {
        // The page's own content only: annotations are not source text.
        const operators = await page.getOperatorList({ annotationMode });
        return pdfPageSourceAppearanceFromOperators({
          sourceSha256: expectedSha256,
          pageNumber,
          operators,
          ops,
          font: (name) => {
            try {
              return page.commonObjs.has(name) ? (page.commonObjs.get(name) as PdfSourceFont) : null;
            } catch {
              return null;
            }
          },
        });
      } finally {
        page.cleanup();
      }
    };
    return {
      snapshot: () => Object.freeze({ sourceSha256: expectedSha256, pageCount: disposed ? 0 : document.numPages, disposed }),
      appearance(pageNumber: number): Promise<PdfPageSourceAppearance> {
        if (disposed) return Promise.reject(new Error("PDF appearance session is closed."));
        if (!Number.isSafeInteger(pageNumber) || pageNumber < 1 || pageNumber > document.numPages) {
          return Promise.reject(new Error("PDF appearance page is out of range."));
        }
        let cached = appearances.get(pageNumber);
        appearances.delete(pageNumber);
        if (!cached) {
          cached = read(pageNumber);
          // A failed read is not kept: the next request reads the page again.
          cached.catch(() => { if (appearances.get(pageNumber) === cached) appearances.delete(pageNumber); });
        }
        appearances.set(pageNumber, cached);
        while (appearances.size > RECENT_PAGES) appearances.delete(appearances.keys().next().value!);
        return cached;
      },
      dispose() {
        if (disposed) return;
        disposed = true;
        appearances.clear();
        void owned.destroy().catch(() => undefined);
      },
    };
  } catch (error) {
    void task?.destroy().catch(() => undefined);
    bytes.fill(0);
    throw error;
  }
}
