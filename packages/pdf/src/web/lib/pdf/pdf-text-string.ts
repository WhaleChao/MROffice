/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFDict,
  PDFHexString,
  PDFName,
  PDFString,
  type PDFField,
} from "pdf-lib";

const UTF8_BOM = [0xef, 0xbb, 0xbf] as const;

/** Decode a PDF text string, or null when `value` is not a string object. */
export function decodePdfTextString(value: unknown): string | null {
  if (!(value instanceof PDFString || value instanceof PDFHexString)) {
    return null;
  }
  const bytes = value.asBytes();
  if (
    bytes.length >= 3 &&
    bytes[0] === UTF8_BOM[0] &&
    bytes[1] === UTF8_BOM[1] &&
    bytes[2] === UTF8_BOM[2]
  ) {
    try {
      return new TextDecoder("utf-8", { fatal: true }).decode(bytes.subarray(3));
    } catch {
      // Not valid UTF-8 after all: read it the way older readers do.
    }
  }
  return value.decodeText();
}

/**
 * A form field's fully qualified name (its own and every ancestor's /T,
 * joined by "."), with each partial name decoded as a text string. Matches
 * pdf.js and the native writer for UTF-8, UTF-16BE and PDFDocEncoding names.
 */
export function pdfFieldFullName(field: PDFField): string {
  const parts: string[] = [];
  const seen = new Set<PDFDict>();
  let dict: PDFDict | undefined = field.acroField.dict;
  while (dict && !seen.has(dict)) {
    seen.add(dict);
    const partial = decodePdfTextString(dict.lookup(PDFName.of("T")));
    if (partial !== null) parts.unshift(partial);
    const parent: unknown = dict.lookup(PDFName.of("Parent"));
    dict = parent instanceof PDFDict ? parent : undefined;
  }
  return parts.join(".");
}
