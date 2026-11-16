/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { pdfLibParseSpeed } from "@mrnote/core/web/lib/pdf/pdf-lib-pacing.ts";
/** Every managed text box of a PDF by page and name: its appearance seal
 * (null without one), how many annotations carry that name, and whether a
 * seal there is malformed. One pdf-lib load for the whole file. */
export type PdfSourceTextBoxSealIndex = {
  readonly pageCount: number;
  /** `${page}\u0000${name}` (page 1-based). */
  readonly boxes: Readonly<Record<string, Readonly<{ matches: number; seal: string | null; invalid: boolean }>>>;
  /** Pages whose annotations could not be read (1-based): a box asked for
   * there is refused, a box elsewhere is not. */
  readonly unreadablePages: readonly number[];
};

export async function scanPdfSourceTextBoxAppearanceSeals(sourceBytes: Uint8Array): Promise<PdfSourceTextBoxSealIndex> {
  const { PDFDocument, PDFArray, PDFDict, PDFName, PDFString, PDFHexString } = await import("pdf-lib");
  const pdf = await PDFDocument.load(sourceBytes, { parseSpeed: pdfLibParseSpeed(), updateMetadata: false });
  const pages = pdf.getPages();
  const boxes: Record<string, { matches: number; seal: string | null; invalid: boolean }> = {};
  const unreadablePages: number[] = [];
  pages.forEach((page, pageIndex) => {
    try {
      scanPage(page, pageIndex);
    } catch {
      unreadablePages.push(pageIndex + 1);
    }
  });
  return { pageCount: pages.length, boxes, unreadablePages };

  function scanPage(page: (typeof pages)[number], pageIndex: number): void {
    const annots = page.node.lookupMaybe(PDFName.of("Annots"), PDFArray);
    for (let index = 0; annots && index < annots.size(); index += 1) {
      const annotation = pdf.context.lookup(annots.get(index));
      if (!(annotation instanceof PDFDict)) continue;
      const nm = annotation.lookup(PDFName.of("NM"));
      if (!(nm instanceof PDFString || nm instanceof PDFHexString)) continue;
      const name = nm.decodeText();
      if (!name.startsWith("MRNote:textbox.")) continue;
      const key = `${pageIndex + 1}\u0000${name}`;
      const entry = boxes[key] ??= { matches: 0, seal: null, invalid: false };
      entry.matches += 1;
      const value = annotation.lookup(PDFName.of("MRNoteSourceTextSeal"));
      if (value instanceof PDFString || value instanceof PDFHexString) {
        const seal = value.decodeText();
        if (/^[a-f0-9]{64}$/.test(seal)) entry.seal = seal;
        else entry.invalid = true;
      }
    }
  }
}
