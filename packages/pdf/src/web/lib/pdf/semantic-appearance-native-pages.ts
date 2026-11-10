/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type {
  PdfNativeVectorContentShape,
  PdfNativeVectorSubpathProof,
} from "./engine/native/vector-content-transaction";
import type { PdfNativeVectorObjectStyleProof } from "./engine/native/vector-style-transaction";
import type { PdfSemanticAppearanceOwnerInventory } from "./semantic-appearance-owner-inventory";

export const MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES = 4_096;
export const MAX_PDF_SEMANTIC_NATIVE_VECTOR_PROOFS = 200_000;

export interface PdfSemanticAppearanceNativePageProofs {
  readonly pageNumber: number;
  readonly shapes: readonly PdfNativeVectorContentShape[];
  readonly subpaths: readonly PdfNativeVectorSubpathProof[];
  readonly styles: readonly PdfNativeVectorObjectStyleProof[];
  /** Per proof id, how much a line width of the drawing's own space grows
   * on the page (its paint matrix's scale), where that is not 1: the shape
   * tab reads and writes line widths in points on paper (PDF 0058). */
  readonly lineWidthScales?: Readonly<Record<string, number>>;
}

/** One source's native vector pages, as the engine read them. */
export interface PdfSemanticAppearanceNativeEvidence {
  readonly sourceSha256: string;
  readonly pageCount: number;
  readonly pages: readonly PdfSemanticAppearanceNativePageProofs[];
}

export interface PdfSemanticAppearanceSelectionSourceEvidence {
  readonly inventory: PdfSemanticAppearanceOwnerInventory;
  readonly nativePages: readonly PdfSemanticAppearanceNativePageProofs[];
}

/** The owner inventory and the native pages of one source, or a throw. */
export function joinPdfSemanticAppearanceSelectionEvidence(
  inventory: PdfSemanticAppearanceOwnerInventory,
  native: PdfSemanticAppearanceNativeEvidence,
): PdfSemanticAppearanceSelectionSourceEvidence {
  if (
    inventory.pageCount > MAX_PDF_SEMANTIC_NATIVE_VECTOR_PAGES ||
    native.sourceSha256 !== inventory.sourceSha256 ||
    native.pageCount !== inventory.pageCount ||
    native.pages.length !== native.pageCount ||
    native.pages.some((page, index) => page.pageNumber !== index + 1) ||
    native.pages.reduce(
      (count, page) => count + page.shapes.length + page.subpaths.length + page.styles.length,
      0,
    ) > MAX_PDF_SEMANTIC_NATIVE_VECTOR_PROOFS
  ) {
    throw new Error("inspection-failed");
  }
  return Object.freeze({ inventory, nativePages: native.pages });
}
