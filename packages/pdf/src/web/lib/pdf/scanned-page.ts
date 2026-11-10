/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface PdfOperatorListLike {
  readonly fnArray: ArrayLike<number>;
}

export type PdfOpsTable = Readonly<Record<string, number | undefined>>;

const IMAGE_OPS = [
  "paintImageXObject",
  "paintImageXObjectRepeat",
  "paintInlineImageXObject",
  "paintInlineImageXObjectGroup",
  "paintImageMaskXObject",
  "paintImageMaskXObjectRepeat",
  "paintImageMaskXObjectGroup",
  "paintJpegXObject",
] as const;

const TEXT_OPS = ["showText", "showSpacedText", "nextLineShowText", "nextLineSetSpacingShowText"] as const;

const PATH_PAINT_OPS = [
  "fill",
  "eoFill",
  "stroke",
  "closeStroke",
  "fillStroke",
  "eoFillStroke",
  "closeFillStroke",
  "closeEOFillStroke",
] as const;

/** Outlined text paints this many filled paths on even a sparse page. */
export const PDF_SCANNED_PAGE_PATH_PAINT_THRESHOLD = 40;

export function pdfOperatorListSuggestsReadableContent(
  operatorList: PdfOperatorListLike | null | undefined,
  ops: PdfOpsTable,
): boolean {
  if (!operatorList) return false;
  const codes = (names: readonly string[]): Set<number> => {
    const out = new Set<number>();
    for (const name of names) {
      const code = ops[name];
      if (typeof code === "number") out.add(code);
    }
    return out;
  };
  const imageCodes = codes(IMAGE_OPS);
  const textCodes = codes(TEXT_OPS);
  const pathCodes = codes(PATH_PAINT_OPS);
  let paintedPaths = 0;
  const { fnArray } = operatorList;
  for (let index = 0; index < fnArray.length; index += 1) {
    const fn = fnArray[index];
    if (imageCodes.has(fn) || textCodes.has(fn)) return true;
    if (pathCodes.has(fn)) {
      paintedPaths += 1;
      if (paintedPaths >= PDF_SCANNED_PAGE_PATH_PAINT_THRESHOLD) return true;
    }
  }
  return false;
}
