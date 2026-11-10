/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  normalizePdfLanguageTag,
  normalizePdfMetadata,
  readPdfXmpMetadata,
  sanitizeMetadataField,
  type PdfMetadata,
  type PdfXmpMetadataLike,
} from "./document-properties";
import { pdfXmlDom } from "./xml-host";

export const RDF_NS = "http://www.w3.org/1999/02/22-rdf-syntax-ns#";
export const DC_NS = "http://purl.org/dc/elements/1.1/";
export const PDF_NS = "http://ns.adobe.com/pdf/1.3/";
export const XML_NS = "http://www.w3.org/XML/1998/namespace";
export const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
export const XMP_MAX_DECODED_BYTES = 2 * 1024 * 1024;

export function parseXmp(text: string): XMLDocument {
  // The browser's DOM, or the engine host's (xml-host.ts): the export writes
  // the properties on the host, where Node has no DOMParser.
  const dom = pdfXmlDom();
  if (!dom) {
    throw new Error("XML parser is unavailable");
  }
  const parsed = dom.parse(text) as XMLDocument;
  if (parsed.getElementsByTagName("parsererror").length > 0) {
    throw new Error("Malformed XMP metadata");
  }
  if (!parsed.getElementsByTagNameNS(RDF_NS, "RDF").item(0)) {
    throw new Error("XMP metadata has no RDF root");
  }
  return parsed;
}

export type XmpField = keyof PdfMetadata;

export interface XmpFieldLocation {
  namespace: string;
  localName: string;
  kind: "alt" | "seq" | "bag" | "text";
}

export const XMP_FIELDS: Record<XmpField, XmpFieldLocation> = {
  title: { namespace: DC_NS, localName: "title", kind: "alt" },
  author: { namespace: DC_NS, localName: "creator", kind: "seq" },
  subject: { namespace: DC_NS, localName: "description", kind: "alt" },
  keywords: { namespace: PDF_NS, localName: "Keywords", kind: "text" },
  language: { namespace: DC_NS, localName: "language", kind: "bag" },
};

export const XMP_FIELD_ORDER: readonly XmpField[] = [
  "title",
  "author",
  "subject",
  "keywords",
  "language",
];

export function propertyElement(
  document: XMLDocument,
  location: XmpFieldLocation,
): Element | null {
  return (
    document
      .getElementsByTagNameNS(location.namespace, location.localName)
      .item(0) ?? null
  );
}

export function propertyAttribute(
  document: XMLDocument,
  location: XmpFieldLocation,
): string | null {
  for (const description of Array.from(
    document.getElementsByTagNameNS(RDF_NS, "Description"),
  )) {
    if (description.hasAttributeNS(location.namespace, location.localName)) {
      return description.getAttributeNS(
        location.namespace,
        location.localName,
      );
    }
  }
  return null;
}

/** The `rdf:li` members of a property's Alt / Seq / Bag container. */
export function containerItems(property: Element): { container: Element; items: Element[] } | null {
  for (const child of Array.from(property.children)) {
    if (
      child.namespaceURI === RDF_NS &&
      (child.localName === "Alt" ||
        child.localName === "Seq" ||
        child.localName === "Bag")
    ) {
      return {
        container: child,
        items: Array.from(child.children).filter(
          (item) => item.namespaceURI === RDF_NS && item.localName === "li",
        ),
      };
    }
  }
  return null;
}

/** The language alternative a reader shows: `x-default`, else the first. */
export function defaultAltItem(items: readonly Element[]): Element | null {
  return (
    items.find(
      (item) =>
        (
          item.getAttributeNS(XML_NS, "lang") ??
          item.getAttribute("xml:lang") ??
          ""
        ).toLowerCase() === "x-default",
    ) ??
    items[0] ??
    null
  );
}

/**
 * Read one editable field from the XMP packet the way the properties dialog
 * shows it: one language alternative of a title (not every alternative
 * glued together), the creators of a sequence joined by ", ", the first
 * language of a bag.
 */
export function readXmpField(document: XMLDocument, field: XmpField): string {
  const location = XMP_FIELDS[field];
  const property = propertyElement(document, location);
  let value: string;
  if (!property) {
    value = propertyAttribute(document, location) ?? "";
  } else {
    const container = containerItems(property);
    if (!container || location.kind === "text") {
      value = property.textContent ?? "";
    } else if (location.kind === "seq") {
      value = container.items
        .map((item) => (item.textContent ?? "").trim())
        .filter(Boolean)
        .join(", ");
    } else if (location.kind === "alt") {
      value = defaultAltItem(container.items)?.textContent ?? "";
    } else {
      value = container.items[0]?.textContent ?? "";
    }
  }
  const sanitized = sanitizeMetadataField(value);
  if (field !== "language") return sanitized;
  return normalizePdfLanguageTag(sanitized) ?? sanitized;
}

export function readXmpDocumentMetadata(document: XMLDocument): PdfMetadata {
  return normalizePdfMetadata({
    title: readXmpField(document, "title"),
    author: readXmpField(document, "author"),
    subject: readXmpField(document, "subject"),
    keywords: readXmpField(document, "keywords"),
    language: readXmpField(document, "language"),
  });
}







export function readPdfXmpMetadataStructured(
  metadata: (PdfXmpMetadataLike & { getRaw?: () => unknown }) | null | undefined,
): PdfMetadata {
  try {
    const raw = metadata?.getRaw?.();
    if (typeof raw === "string" && raw.trim()) {
      return readXmpDocumentMetadata(parseXmp(raw));
    }
  } catch {
    /* fall back to pdf.js's flattened map */
  }
  return readPdfXmpMetadata(metadata);
}
