/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The pictures a PDF page paints, as the picture tools number and frame
 * them: measured from pdf.js's operator list in the stored orientation the
 * way the embedded-image writers count paints (so `order` is their key),
 * then turned into the shown orientation, with the paint's matrix and the
 * page as shown. The reader's select tool (PDFViewer) and the served
 * picture writes (picture-format-plan) read the same targets.
 */

import { placedImagesFromOperatorList, significantImages, type ImageOperatorIds, type OperatorListLike } from "./embedded-images";
import { rotateFractionRect } from "./fraction-rotation";

type Affine = readonly [number, number, number, number, number, number];

export interface PdfPagePictureTarget {
  readonly pageNumber: number;
  readonly order: number;
  readonly operatorIndex?: number;
  readonly name: string;
  readonly rotated: boolean;
  readonly bounds: Readonly<{ x: number; y: number; width: number; height: number }>;
  readonly matrix: Affine;
  readonly display: Readonly<{ transform: Affine; width: number; height: number }>;
  readonly sourceSha256: string;
  readonly sourceOpacity?: number;
  /** The page cuts the picture (a photo bled past the page's edges) and it
   * has no clip of its own: `bounds` is its part on the page, which its
   * frame stands on and every write keeps as the picture's own crop (WPS
   * frames and resizes the visible part, its clip going with it). */
  readonly pageClipped?: boolean;
  /** The clips in force at the paint that are not the picture's own, as one
   * box in page fractions as shown (PlacedImage.enclosingClip): a picture
   * moved past it is taken out from under it, or it is drawn cut there. */
  readonly enclosingClip?: Readonly<{ x: number; y: number; width: number; height: number }>;
}

/** The part of a page-fraction box on the page, when the page cuts it. */
function onPagePart(
  box: Readonly<{ x: number; y: number; width: number; height: number }>,
): Readonly<{ x: number; y: number; width: number; height: number }> | null {
  const left = Math.max(0, box.x);
  const top = Math.max(0, box.y);
  const right = Math.min(1, box.x + box.width);
  const bottom = Math.min(1, box.y + box.height);
  if (!(right - left > 1e-6) || !(bottom - top > 1e-6)) return null;
  const cut = left > box.x + 1e-6 || top > box.y + 1e-6 ||
    right < box.x + box.width - 1e-6 || bottom < box.y + box.height - 1e-6;
  return cut ? { x: left, y: top, width: right - left, height: bottom - top } : null;
}

interface ViewportLike {
  readonly width: number;
  readonly height: number;
  readonly transform: ArrayLike<number>;
}

/** pdf.js's operator ids for the picture count (`pdfjs.OPS`). */
export function pdfPagePictureOperatorIds(ops: Readonly<Record<string, number>>): ImageOperatorIds {
  // Counted as the writers count paints: an inline image (no resource name,
  // which no writer addresses) and an annotation's picture are not offered,
  // and a run pdf.js merged into one Repeat operator is each of its pictures.
  return {
    save: ops.save,
    restore: ops.restore,
    transform: ops.transform,
    paintImageXObject: ops.paintImageXObject,
    paintImageXObjectRepeat: ops.paintImageXObjectRepeat,
    paintImageMaskXObject: ops.paintImageMaskXObject,
    beginAnnotation: ops.beginAnnotation,
    endAnnotation: ops.endAnnotation,
    clip: ops.clip,
    eoClip: ops.eoClip,
    constructPath: ops.constructPath,
    dependency: ops.dependency,
    setGState: ops.setGState,
    beginGroup: ops.beginGroup,
    endGroup: ops.endGroup,
    paintFormXObjectBegin: ops.paintFormXObjectBegin,
    paintFormXObjectEnd: ops.paintFormXObjectEnd,
    beginMarkedContent: ops.beginMarkedContent,
    beginMarkedContentProps: ops.beginMarkedContentProps,
    endMarkedContent: ops.endMarkedContent,
  };
}

export function pdfPagePictureTargets(input: {
  readonly pageNumber: number;
  readonly operatorList: OperatorListLike;
  readonly ops: Readonly<Record<string, number>>;
  /** The page at scale 1, rotation 0 (the stored orientation). */
  readonly storedViewport: ViewportLike;
  /** The page at scale 1 as shown. */
  readonly viewport: ViewportLike;
  /** The page's own /Rotate. */
  readonly rawRotation: number;
  readonly sourceSha256: string;
  /** Whether an optional-content layer is shown (pdf.js OptionalContentConfig). */
  readonly isLayerVisible?: (properties: unknown) => boolean;
}): PdfPagePictureTarget[] {
  const { viewport } = input;
  return significantImages(
    placedImagesFromOperatorList(
      input.operatorList,
      pdfPagePictureOperatorIds(input.ops),
      {
        width: input.storedViewport.width,
        height: input.storedViewport.height,
        transform: input.storedViewport.transform,
      },
      input.isLayerVisible ? { isLayerVisible: input.isLayerVisible } : {},
    ),
  ).map((image) => {
    // Only an upright paint (a and d positive, no turn or mirror) with no
    // clip of its own: its frame is then its visible box, and the crop that
    // keeps it is one rectangle of its own square.
    const [a, b, c, d] = image.matrix;
    const upright = a > 0 && d > 0 && Math.abs(b) <= 1e-9 && Math.abs(c) <= 1e-9;
    const onPage = upright && !image.frame
      ? onPagePart({ x: image.x, y: image.y, width: image.width, height: image.height })
      : null;
    return {
      pageNumber: input.pageNumber,
      order: image.order,
      operatorIndex: image.operatorIndex,
      name: image.name,
      rotated: image.rotated,
      bounds: rotateFractionRect(
        onPage ?? { x: image.x, y: image.y, width: image.width, height: image.height },
        input.rawRotation,
      ),
      ...(onPage ? { pageClipped: true } : {}),
      ...(image.enclosingClip ? { enclosingClip: rotateFractionRect(image.enclosingClip, input.rawRotation) } : {}),
      // What the in-place picture frame and its writes need: the paint, the
      // page as shown, and the bytes both were read from.
      matrix: image.matrix,
      display: {
        transform: [
          viewport.transform[0],
          viewport.transform[1],
          viewport.transform[2],
          viewport.transform[3],
          viewport.transform[4],
          viewport.transform[5],
        ] as const,
        width: viewport.width,
        height: viewport.height,
      },
      sourceSha256: input.sourceSha256,
      ...(image.sourceOpacity === undefined ? {} : { sourceOpacity: image.sourceOpacity }),
    };
  });
}
