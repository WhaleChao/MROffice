/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where the PDF modules shared by the browser and the engine host get their
 * XML DOM. The browser has `DOMParser` / `XMLSerializer`; the engine host's
 * Node workers have neither, so the host installs its own
 * (`xml-dom-mutable.ts`, server-only) here at start-up. Nothing is imported
 * from the host side, so this stays safe for the browser bundle.
 */

export interface PdfXmlDom {
  /** `new DOMParser().parseFromString(text, "application/xml")`. */
  parse(text: string): Document;
  /** `new XMLSerializer().serializeToString(document)`. */
  serialize(document: Document): string;
}

let installed: PdfXmlDom | null = null;

/** The engine host's DOM (engine-host/engines/pdf-document.ts). */
export function installPdfXmlDom(dom: PdfXmlDom | null): void {
  installed = dom;
}

/** The browser's DOM when it has one, else the host's; null when neither. */
export function pdfXmlDom(): PdfXmlDom | null {
  if (typeof DOMParser !== "undefined" && typeof XMLSerializer !== "undefined") {
    return {
      parse: (text) => new DOMParser().parseFromString(text, "application/xml"),
      serialize: (document) => new XMLSerializer().serializeToString(document),
    };
  }
  return installed;
}
