/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Where a source vector stands among everything its page paints, read from
 * the renderer's operator list: Office greys "Bring Forward" once a shape is
 * drawn last and "Send Backward" once it is drawn first. A paint stop is one
 * operation that puts ink on the page -- a painted path, text, a picture, a
 * shading -- outside the annotation appearances the renderer draws after the
 * page content. The writer stays the authority on the move itself.
 */

const TEXT_AND_IMAGE_PAINT_OPERATORS = [
  "showText",
  "showSpacedText",
  "nextLineShowText",
  "nextLineSetSpacingShowText",
  "paintImageXObject",
  "paintImageXObjectRepeat",
  "paintImageMaskXObject",
  "paintImageMaskXObjectGroup",
  "paintImageMaskXObjectRepeat",
  "paintInlineImageXObject",
  "paintInlineImageXObjectGroup",
  "paintJpegXObject",
  "paintSolidColorImageMask",
  "shadingFill",
] as const;

/**
 * The operator indices that paint on the page, ascending: every painted
 * path the inventory lists plus every text, picture and shading operation
 * outside `beginAnnotation`…`endAnnotation`.
 */
export function pdfPagePaintStops(
  fnArray: readonly number[],
  ops: Readonly<Record<string, number | undefined>>,
  paintedPathIndices: readonly number[],
): readonly number[] {
  const inkOperators = new Set<number>();
  for (const name of TEXT_AND_IMAGE_PAINT_OPERATORS) {
    const id = ops[name];
    if (typeof id === "number") inkOperators.add(id);
  }
  const stops = new Set<number>(paintedPathIndices);
  let annotationDepth = 0;
  for (let index = 0; index < fnArray.length; index += 1) {
    const fn = fnArray[index];
    if (fn === ops.beginAnnotation) annotationDepth += 1;
    else if (fn === ops.endAnnotation) annotationDepth = Math.max(0, annotationDepth - 1);
    else if (annotationDepth === 0 && inkOperators.has(fn)) stops.add(index);
  }
  return Object.freeze([...stops].sort((left, right) => left - right));
}

export interface PdfNativeVectorLayerBoundary {
  readonly atFront: boolean;
  readonly atBack: boolean;
}

/** Whether the vector at `operatorIndex` is painted last (first) on its page. */
export function nativeVectorLayerBoundary(
  stops: readonly number[],
  operatorIndex: number,
): PdfNativeVectorLayerBoundary {
  let atFront = true;
  let atBack = true;
  for (const stop of stops) {
    if (stop > operatorIndex) atFront = false;
    if (stop < operatorIndex) atBack = false;
  }
  return Object.freeze({ atFront, atBack });
}
