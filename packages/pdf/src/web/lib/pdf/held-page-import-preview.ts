/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { loadCompatiblePdfJs, PDF_JS_DOCUMENT_ASSETS } from "@mrnote/core/web/lib/pdf/pdfjs-runtime.ts";

/** One page of a file imported (or set in place of a page) and held for
 * Save (AutoSave off): what the reader shows of it until the saved file
 * carries it. Its size is in points, as the page's own. */
export interface PdfHeldImportedPagePreview {
  readonly width: number;
  readonly height: number;
  /** The page drawn, as an image the reader shows in its place. */
  readonly image: string;
}

/** The pages `pageIndices` (0-based) of `bytes`, drawn for the reader. */
export async function renderPdfHeldImportPreviews(
  bytes: Uint8Array,
  pageIndices: readonly number[],
  scale = 1.5,
): Promise<PdfHeldImportedPagePreview[]> {
  const pdfjs = await loadCompatiblePdfJs();
  const task = pdfjs.getDocument({ data: Uint8Array.from(bytes), ...PDF_JS_DOCUMENT_ASSETS });
  const document_ = await task.promise;
  try {
    const previews: PdfHeldImportedPagePreview[] = [];
    for (const index of pageIndices) {
      const page = await document_.getPage(index + 1);
      const base = page.getViewport({ scale: 1 });
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(viewport.width));
      canvas.height = Math.max(1, Math.round(viewport.height));
      const context = canvas.getContext("2d");
      if (!context) throw new Error("The page could not be drawn.");
      context.fillStyle = "#fff";
      context.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvas, canvasContext: context, viewport }).promise;
      previews.push(Object.freeze({ width: base.width, height: base.height, image: canvas.toDataURL("image/png") }));
      canvas.width = 0;
      canvas.height = 0;
    }
    return previews;
  } finally {
    await document_.destroy().catch(() => undefined);
  }
}
