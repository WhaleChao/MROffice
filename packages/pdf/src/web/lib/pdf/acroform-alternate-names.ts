/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  PDFDict,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFString,
  type PDFDocument,
} from "pdf-lib";

import type { PdfTextFieldAlignment } from "@mrnote/core/web/lib/pdf/acroform.ts";
import { pdfFieldFullName } from "./pdf-text-string";

/** How many fields to trust from one form. */
const MAX_FIELDS = 16_384;
const MAX_ALTERNATE_NAME = 512;
const CONTROL_CHARS = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g;

function decodeText(value: unknown): string {
  try {
    if (value instanceof PDFHexString || value instanceof PDFString) {
      return value.decodeText();
    }
  } catch {
    /* an undecodable string is no name */
  }
  return "";
}

/**
 * Fully-qualified field name → `/TU` on that terminal field. Fields without a
 * `/TU` are absent. Never throws; an unreadable form yields an empty map.
 */
export function readFieldAlternateNames(doc: PDFDocument): Map<string, string> {
  const out = new Map<string, string>();
  let fields: ReturnType<PDFDocument["getForm"]>["getFields"] extends () => infer R
    ? R
    : never;
  try {
    fields = doc.getForm().getFields();
  } catch {
    return out;
  }
  for (const field of fields.slice(0, MAX_FIELDS)) {
    try {
      const alternate = decodeText(field.acroField.dict.get(PDFName.of("TU")))
        .replace(/\r\n?/g, "\n")
        .replace(CONTROL_CHARS, "")
        .trim()
        .slice(0, MAX_ALTERNATE_NAME);
      if (!alternate) continue;
      const name = pdfFieldFullName(field);
      if (name && !out.has(name)) out.set(name, alternate);
    } catch {
      /* skip a malformed field */
    }
  }
  return out;
}

/** Read the alternate names from raw PDF bytes; empty for anything unreadable. */
export async function readFieldAlternateNamesFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<Map<string, string>> {
  try {
    const { loadPdfForInspection } = await import("./encrypted-guard");
    const doc = await loadPdfForInspection(bytes, { updateMetadata: false });
    if (!doc) return new Map();
    return readFieldAlternateNames(doc);
  } catch {
    return new Map();
  }
}








export function readAcroFormDefaultAlignment(
  doc: PDFDocument,
): PdfTextFieldAlignment | null {
  try {
    const acroForm = doc.catalog.lookup(PDFName.of("AcroForm"), PDFDict);
    const quadding = acroForm?.lookup(PDFName.of("Q"));
    if (!(quadding instanceof PDFNumber)) return null;
    const value = quadding.asNumber();
    return value === 1 ? "center" : value === 2 ? "right" : value === 0 ? "left" : null;
  } catch {
    return null;
  }
}

/** The form facts pdf.js does not surface, read from raw PDF bytes in one
 *  parse: alternate names and the AcroForm default justification. */
/**
 * A push button's caption, as its widget's appearance characteristics say
 * (`/MK /CA`, ISO 32000-1 12.5.6.19): the text WPS and Acrobat draw on a
 * button the file gives no appearance of its own. Keyed by full field name;
 * a field whose widgets disagree has none.
 */
export function readPushButtonCaptions(doc: PDFDocument): Map<string, string> {
  const out = new Map<string, string>();
  const conflicting = new Set<string>();
  let fields: ReturnType<PDFDocument["getForm"]>["getFields"] extends () => infer R
    ? R
    : never;
  try {
    fields = doc.getForm().getFields();
  } catch {
    return out;
  }
  for (const field of fields.slice(0, MAX_FIELDS)) {
    try {
      const name = pdfFieldFullName(field);
      if (!name) continue;
      const dicts = [field.acroField.dict, ...field.acroField.getWidgets().map((widget) => widget.dict)];
      for (const dict of dicts) {
        const characteristics = dict.lookup(PDFName.of("MK"));
        if (!(characteristics instanceof PDFDict)) continue;
        const caption = decodeText(characteristics.get(PDFName.of("CA")))
          .replace(/\r\n?/g, "\n")
          .replace(CONTROL_CHARS, "")
          .trim()
          .slice(0, MAX_ALTERNATE_NAME);
        if (!caption) continue;
        const known = out.get(name);
        if (known !== undefined && known !== caption) conflicting.add(name);
        else out.set(name, caption);
      }
    } catch {
      /* skip a malformed field */
    }
  }
  for (const name of conflicting) out.delete(name);
  return out;
}

export async function readFormSourceFactsFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<{
  names: Map<string, string>;
  defaultAlignment: PdfTextFieldAlignment | null;
  captions: Map<string, string>;
}> {
  try {
    const { loadPdfForInspection } = await import("./encrypted-guard");
    const doc = await loadPdfForInspection(bytes, { updateMetadata: false });
    if (!doc) return { names: new Map(), defaultAlignment: null, captions: new Map() };
    return {
      names: readFieldAlternateNames(doc),
      defaultAlignment: readAcroFormDefaultAlignment(doc),
      captions: readPushButtonCaptions(doc),
    };
  } catch {
    return { names: new Map(), defaultAlignment: null, captions: new Map() };
  }
}
