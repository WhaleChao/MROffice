/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The document's own calculation order (`/AcroForm /CO`).
 *
 * Recalculation always runs a field after the calculated fields it reads, so a
 * missing, partial or contradictory order cannot change the answer. /CO decides
 * only between fields that are ready at the same time, and a complete,
 * consistent /CO is followed exactly.
 *
 * pdf.js does not surface /CO on the annotations it returns, so it is read
 * here straight from the AcroForm dictionary. /CO holds field references, not
 * names; the names come from matching those references against the form's own
 * fields, which is also what makes a reference to something that is not a field
 * (or is no longer in the document) drop out rather than turn into a name that
 * matches nothing.
 */

import {
  PDFArray,
  PDFDict,
  PDFName,
  PDFRef,
  type PDFDocument,
} from "pdf-lib";
import { pdfFieldFullName } from "./pdf-text-string";

/** How many entries to trust from a /CO array. */
const MAX_CALCULATION_ORDER = 16_384;

/**
 * Fully-qualified field names in the document's calculation order.
 *
 * Empty means the document does not say — ties then fall back to document
 * order; it never means "no calculations".
 */
export function readCalculationOrder(doc: PDFDocument): string[] {
  let acroForm: PDFDict | undefined;
  try {
    acroForm = doc.catalog.lookup(PDFName.of("AcroForm"), PDFDict);
  } catch {
    return [];
  }
  if (!acroForm) return [];

  let order: PDFArray | undefined;
  try {
    order = acroForm.lookup(PDFName.of("CO"), PDFArray);
  } catch {
    return [];
  }
  if (!order) return [];

  const nameByRef = new Map<string, string>();
  try {
    for (const field of doc.getForm().getFields()) {
      nameByRef.set(field.ref.tag, pdfFieldFullName(field));
    }
  } catch {
    return [];
  }

  const names: string[] = [];
  const seen = new Set<string>();
  const limit = Math.min(order.size(), MAX_CALCULATION_ORDER);
  for (let index = 0; index < limit; index += 1) {
    const entry = order.get(index);
    if (!(entry instanceof PDFRef)) continue;
    const name = nameByRef.get(entry.tag);
    // A name can legitimately appear twice in /CO; running it twice would be
    // harmless but reporting it twice would make the order look longer than
    // the form is.
    if (!name || seen.has(name)) continue;
    seen.add(name);
    names.push(name);
  }
  return names;
}

/**
 * Read /CO from raw PDF bytes.
 *
 * Returns an empty list for anything unreadable: a form whose order cannot be
 * recovered still calculates in dependency order.
 */
export async function readCalculationOrderFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<string[]> {
  try {
    const { loadPdfForInspection } = await import("./encrypted-guard");
    const doc = await loadPdfForInspection(bytes, { updateMetadata: false });
    if (!doc) return [];
    return readCalculationOrder(doc);
  } catch {
    return [];
  }
}
