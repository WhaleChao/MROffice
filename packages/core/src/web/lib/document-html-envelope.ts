/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Machine-readable payload carried by an MRNote standalone HTML export. */
export const MRNOTE_HTML_DOCUMENT_SCHEMA = "mrnote.document.html";
export const MRNOTE_HTML_DOCUMENT_VERSION = 1;
export const MRNOTE_HTML_DOCUMENT_SCRIPT_ID = "mrnote-document-data";

export interface MrnoteHtmlDocumentEnvelope {
  schema: typeof MRNOTE_HTML_DOCUMENT_SCHEMA;
  version: typeof MRNOTE_HTML_DOCUMENT_VERSION;
  title: string;
  encoding: "base64";
  content: string;
}

function encodeUtf8Base64(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 32_768;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(offset, offset + chunkSize),
    );
  }
  return btoa(binary);
}

export function decodeMrnoteHtmlDocumentContent(value: string): unknown {
  if (
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(
      value,
    )
  ) {
    throw new Error("Invalid base64 document payload");
  }
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes));
}

/**
 * Escape JSON for an inert `<script type="application/json">` text node.
 * Escaping `<` is the critical boundary: document text can contain `</script>`
 * and must never be able to terminate the metadata element.
 */
export function serializeMrnoteHtmlDocumentEnvelope(
  title: string,
  content: unknown,
): string {
  const envelope: MrnoteHtmlDocumentEnvelope = {
    schema: MRNOTE_HTML_DOCUMENT_SCHEMA,
    version: MRNOTE_HTML_DOCUMENT_VERSION,
    title: title.trim().slice(0, 255),
    encoding: "base64",
    // The semantic snapshot is base64-wrapped so visible standalone HTML and
    // source-text assertions only see the static projection, not stale cached
    // labels or other non-rendered JSON attributes. The importer decodes and
    // validates the snapshot before persistence.
    content: encodeUtf8Base64(JSON.stringify(content)),
  };
  return JSON.stringify(envelope).replace(
    /[<>&\u2028\u2029]/g,
    (character) =>
      ({
        "<": "\\u003c",
        ">": "\\u003e",
        "&": "\\u0026",
        "\u2028": "\\u2028",
        "\u2029": "\\u2029",
      })[character] ?? "",
  );
}
