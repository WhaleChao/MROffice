/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A picture's alternate description in a tagged PDF: the /Alt of the
 * structure element (a /Figure, as Word and Acrobat write it) whose marked
 * content the picture is painted in (ISO 32000-1 §14.9.3). The element is
 * found the way a reader finds it: the page's /StructParents key in the
 * structure tree root's /ParentTree names the array whose MCID-th entry is
 * the element (§14.7.4.4). Reading only: the write is the engine's
 * (planMraiNativeStructElementAlt).
 */

import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString, type PDFObject } from "pdf-lib";

export type PdfPictureAltTarget = Readonly<{
  element: Readonly<{ objectNumber: number; generation: number }>;
  /** The /Alt the element has now, if any. */
  alt: string | null;
}>;

function lookup(document: PDFDocument, value: PDFObject | undefined): PDFObject | undefined {
  return value instanceof PDFRef ? document.context.lookup(value) : value;
}

/** The value a number tree gives `key` (§7.9.7), following its /Kids. */
function numberTreeValue(document: PDFDocument, node: PDFDict, key: number, depth = 0): PDFObject | undefined {
  if (depth > 32) return undefined;
  const nums = lookup(document, node.get(PDFName.of("Nums")));
  if (nums instanceof PDFArray) {
    for (let index = 0; index + 1 < nums.size(); index += 2) {
      const at = lookup(document, nums.get(index));
      if (at instanceof PDFNumber && at.asNumber() === key) return nums.get(index + 1);
    }
  }
  const kids = lookup(document, node.get(PDFName.of("Kids")));
  if (!(kids instanceof PDFArray)) return undefined;
  for (let index = 0; index < kids.size(); index += 1) {
    const kid = lookup(document, kids.get(index));
    if (!(kid instanceof PDFDict)) continue;
    const limits = lookup(document, kid.get(PDFName.of("Limits")));
    if (limits instanceof PDFArray && limits.size() === 2) {
      const low = lookup(document, limits.get(0));
      const high = lookup(document, limits.get(1));
      if (low instanceof PDFNumber && high instanceof PDFNumber && (key < low.asNumber() || key > high.asNumber())) continue;
    }
    const found = numberTreeValue(document, kid, key, depth + 1);
    if (found !== undefined) return found;
  }
  return undefined;
}

/**
 * The structure element for marked content `mcid` on `page`, or "untagged"
 * when the document has no structure tree (or the page no /StructParents),
 * or null when the tree names no element for it.
 */
export async function readPdfPictureAltTarget(
  bytes: Uint8Array,
  page: Readonly<{ objectNumber: number; generation: number }>,
  mcid: number,
): Promise<PdfPictureAltTarget | "untagged" | null> {
  const document = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
  const root = lookup(document, document.catalog.get(PDFName.of("StructTreeRoot")));
  if (!(root instanceof PDFDict)) return "untagged";
  const pageDict = document.context.lookup(PDFRef.of(page.objectNumber, page.generation));
  const key = pageDict instanceof PDFDict ? lookup(document, pageDict.get(PDFName.of("StructParents"))) : undefined;
  const parentTree = lookup(document, root.get(PDFName.of("ParentTree")));
  if (!(key instanceof PDFNumber) || !(parentTree instanceof PDFDict)) return "untagged";
  const entry = lookup(document, numberTreeValue(document, parentTree, key.asNumber()));
  if (!(entry instanceof PDFArray) || mcid >= entry.size()) return null;
  const elementRef = entry.get(mcid);
  if (!(elementRef instanceof PDFRef)) return null;
  const element = document.context.lookup(elementRef);
  if (!(element instanceof PDFDict)) return null;
  const alt = lookup(document, element.get(PDFName.of("Alt")));
  return Object.freeze({
    element: Object.freeze({ objectNumber: elementRef.objectNumber, generation: elementRef.generationNumber }),
    alt: alt instanceof PDFString || alt instanceof PDFHexString ? alt.decodeText() : null,
  });
}
