/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  inspectMraiNativeVectorDocumentAuthority,
  inspectMraiNativeVectorDocumentAuthorityPages,
  pdfNativeVectorSelectionMatrix,
  type PdfNativeVectorContentShape,
  type PdfNativeVectorDocumentAuthorityInspection,
  type PdfNativeVectorSubpathProof,
} from "./engine/native/vector-content-transaction";
import { inspectMraiNativeVectorObjectStyleDocument } from "./engine/native/vector-style-transaction";
import { inspectPdfSemanticAppearanceOwners } from "./semantic-appearance-owner-inventory";
import {
  joinPdfSemanticAppearanceSelectionEvidence,
  MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES,
  MAX_PDF_SEMANTIC_NATIVE_VECTOR_PROOFS,
  type PdfSemanticAppearanceNativeEvidence,
  type PdfSemanticAppearanceSelectionSourceEvidence,
} from "./semantic-appearance-native-pages";

export async function inspectPdfSemanticAppearanceNativePages(
  sourceBytes: Uint8Array,
): Promise<PdfSemanticAppearanceNativeEvidence> {
  return nativeEvidenceFromAuthority(sourceBytes, await inspectMraiNativeVectorDocumentAuthority(sourceBytes));
}

/** The native evidence of an authority's pages (the whole document, or a
 * slice of it whose `pageCount` is its own length). */
/** How much a line width of each drawing's own space grows on the page:
 * the square root of its paint matrix's determinant, kept where it is not 1
 * (user audit PDF 0058: a drawing scaled to 0.65 read 4 pt and drew 2.6). */
function lineWidthScales(
  proofs: readonly (PdfNativeVectorContentShape | PdfNativeVectorSubpathProof)[],
): Readonly<Record<string, number>> | undefined {
  const scales: Record<string, number> = {};
  for (const proof of proofs) {
    const matrix = pdfNativeVectorSelectionMatrix(proof);
    if (!matrix) continue;
    const scale = Math.sqrt(Math.abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]));
    if (Number.isFinite(scale) && scale > 0 && Math.abs(scale - 1) > 1e-9) {
      scales[proof.id] = scale;
    }
  }
  return Object.keys(scales).length > 0 ? Object.freeze(scales) : undefined;
}

async function nativeEvidenceFromAuthority(
  sourceBytes: Uint8Array,
  nativeAuthority: PdfNativeVectorDocumentAuthorityInspection,
): Promise<PdfSemanticAppearanceNativeEvidence> {
  const native = nativeAuthority.document;
  const styles = await inspectMraiNativeVectorObjectStyleDocument(sourceBytes, nativeAuthority);
  if (
    native.pageCount > MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES ||
    styles.sourceSha256 !== native.sourceSha256 ||
    styles.pageCount !== native.pageCount ||
    styles.pages.some(
      (page, index) =>
        page.pageNumber !== native.pages[index]?.pageNumber ||
        page.proofs.length + page.unsupportedSelectionIds.length !==
          (native.pages[index]?.shapes.length ?? 0) + (native.pages[index]?.subpaths.length ?? 0),
    ) ||
    native.pages.reduce((count, page) => count + page.shapes.length + page.subpaths.length, 0) >
      MAX_PDF_SEMANTIC_NATIVE_VECTOR_PROOFS
  ) {
    throw new Error("inspection-failed");
  }
  return Object.freeze({
    sourceSha256: native.sourceSha256,
    pageCount: native.pageCount,
    pages: Object.freeze(
      native.pages.map((page, index) =>
        Object.freeze({
          pageNumber: page.pageNumber,
          shapes: page.shapes,
          subpaths: page.subpaths,
          styles: styles.pages[index]?.proofs ?? Object.freeze([]),
          ...((scales) => (scales ? { lineWidthScales: scales } : {}))(
            lineWidthScales([...page.shapes, ...page.subpaths]),
          ),
        }),
      ),
    ),
  });
}

async function inspectWholeSelectionSource(
  sourceBytes: Uint8Array,
): Promise<PdfSemanticAppearanceSelectionSourceEvidence> {
  const [inventory, native] = await Promise.all([
    inspectPdfSemanticAppearanceOwners(sourceBytes),
    inspectPdfSemanticAppearanceNativePages(sourceBytes),
  ]);
  return joinPdfSemanticAppearanceSelectionEvidence(inventory, native);
}





export async function inspectPdfSemanticAppearanceNativePageRange(
  sourceBytes: Uint8Array,
  firstPage: number,
  budgetMs: number,
): Promise<Readonly<{
  pages: PdfSemanticAppearanceNativeEvidence["pages"];
  totalPages: number;
  sourceSha256: string;
}>> {
  const slice = await inspectMraiNativeVectorDocumentAuthorityPages(sourceBytes, firstPage, budgetMs);
  const native = await nativeEvidenceFromAuthority(sourceBytes, slice.authority);
  if (native.pages.some((page, index) => page.pageNumber !== firstPage + index)) throw new Error("inspection-failed");
  return Object.freeze({ pages: native.pages, totalPages: slice.totalPages, sourceSha256: native.sourceSha256 });
}

/** The work one chunked read does before it answers. */
const CHUNK_BUDGET_MS = 3_000;
const MIN_CHUNK_BYTES = 1024;
const MAX_CHUNK_BYTES = 6 * 1024 * 1024;











export async function inspectPdfSemanticAppearanceSelectionSource(
  sourceBytes: Uint8Array,
  range?: Readonly<{ firstPage: number; maxBytes: number }>,
): Promise<PdfSemanticAppearanceSelectionSourceEvidence & Readonly<{ nextPage?: number | null }>> {
  if (!range) return inspectWholeSelectionSource(sourceBytes);
  if (!Number.isSafeInteger(range.firstPage) || range.firstPage < 1 || !Number.isFinite(range.maxBytes)) {
    throw new Error("inspection-failed");
  }
  // Each call reads the pages from `firstPage` on for at most
  // CHUNK_BUDGET_MS of work (always one page), then as many of those as fit
  // in `maxBytes`. The whole document in one call ran past the engine call's
  // 15 s deadline on a vector-heavy datasheet (41 s): the read timed out on
  
  const inventory = await inspectPdfSemanticAppearanceOwners(sourceBytes);
  const slice = await inspectMraiNativeVectorDocumentAuthorityPages(sourceBytes, range.firstPage, CHUNK_BUDGET_MS);
  if (slice.totalPages !== inventory.pageCount || inventory.pageCount > MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES ||
    slice.authority.document.sourceSha256 !== inventory.sourceSha256) {
    throw new Error("inspection-failed");
  }
  const native = await nativeEvidenceFromAuthority(sourceBytes, slice.authority);
  const pages = native.pages;
  if (pages.some((page, index) => page.pageNumber !== range.firstPage + index)) throw new Error("inspection-failed");
  const budget = Math.min(MAX_CHUNK_BYTES, Math.max(MIN_CHUNK_BYTES, range.maxBytes));
  let used = 0;
  let end = 0;
  while (end < pages.length) {
    const size = JSON.stringify(pages[end]).length;
    if (end > 0 && used + size > budget) break;
    used += size;
    end += 1;
  }
  const last = range.firstPage + end - 1;
  return Object.freeze({
    inventory,
    nativePages: Object.freeze(pages.slice(0, end)),
    nextPage: last < inventory.pageCount ? last + 1 : null,
  });
}
