/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Printing in a browser that cannot show a PDF (user audit PDF 0046). The
 * reader printed by loading the PDF into a hidden frame and printing the
 * frame; a browser without a PDF viewer (an embedded app browser, Chromium
 * with the viewer off: `navigator.pdfViewerEnabled === false`) never shows
 * it there, and every print ended in 「无法打印这个 PDF」 after the paper
 * check passed. Such a browser prints the pages the reader draws itself,
 * one picture per page at its own paper size, as WPS prints with its own
 * renderer.
 */

/** Print resolution of the drawn pages. */
export const PDF_RASTER_PRINT_DPI = 150;

export interface PdfRasterPrintPage {
  /** The page's picture (an object URL). */
  readonly url: string;
  /** Paper size in inches, as the page shows (its /Rotate applied). */
  readonly widthIn: number;
  readonly heightIn: number;
}

/** Whether this browser shows a PDF in a frame it can print. */
export function browserShowsPdf(): boolean {
  if (typeof navigator === "undefined") return true;
  return (navigator as Navigator & { pdfViewerEnabled?: boolean }).pdfViewerEnabled !== false;
}

const inches = (value: number) => `${Math.round(value * 10_000) / 10_000}in`;

/**
 * The document the frame prints: each page on a sheet of its own size (one
 * named @page per size, so a landscape page among portrait ones prints
 * landscape), no margins, nothing else on it.
 */
export function rasterPrintDocumentHtml(pages: readonly PdfRasterPrintPage[]): string {
  const sizes = new Map<string, string>();
  const nameFor = (page: PdfRasterPrintPage) => {
    const key = `${inches(page.widthIn)} ${inches(page.heightIn)}`;
    let name = sizes.get(key);
    if (!name) {
      name = `mrai-sheet-${sizes.size}`;
      sizes.set(key, name);
    }
    return name;
  };
  const sheets = pages
    .map((page) => {
      const name = nameFor(page);
      const url = page.url.replace(/[^A-Za-z0-9:/._%-]/g, "");
      return `<div class="sheet" style="page:${name};width:${inches(page.widthIn)};height:${inches(page.heightIn)}"><img src="${url}" alt=""></div>`;
    })
    .join("");
  const pageRules = [...sizes.entries()]
    .map(([size, name]) => `@page ${name}{size:${size};margin:0}`)
    .join("");
  return (
    "<!doctype html><html><head><meta charset=\"utf-8\"><title></title><style>" +
    `@page{margin:0}${pageRules}` +
    "html,body{margin:0;padding:0;background:#fff}" +
    ".sheet{overflow:hidden;break-after:page;page-break-after:always}" +
    ".sheet:last-child{break-after:auto;page-break-after:auto}" +
    ".sheet img{display:block;width:100%;height:100%;object-fit:fill}" +
    `</style></head><body>${sheets}</body></html>`
  );
}
