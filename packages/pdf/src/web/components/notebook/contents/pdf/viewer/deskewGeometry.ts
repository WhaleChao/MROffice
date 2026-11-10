/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { PdfRotation } from "@mrnote/core/web/lib/pdf/page-ops.ts";

export interface PdfDeskewBox {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface PdfDeskewOutputGeometry {
  readonly media: PdfDeskewBox;
  readonly crop: PdfDeskewBox;
}

export function sameDeskewCoordinate(left: number, right: number): boolean {
  return (
    Math.abs(left - right) <=
    Number.EPSILON * 32 * Math.max(1, Math.abs(left), Math.abs(right))
  );
}

export function readPdfDeskewOutputGeometry(options: {
  readonly media: PdfDeskewBox;
  readonly crop: PdfDeskewBox;
  readonly pdfJsView: readonly number[];
  readonly sourceUserUnit: number;
  readonly rotation: PdfRotation;
}): PdfDeskewOutputGeometry | null {
  const { media, crop, pdfJsView, sourceUserUnit, rotation } = options;
  if (
    !Number.isFinite(sourceUserUnit) ||
    sourceUserUnit <= 0 ||
    ![
      media.x,
      media.y,
      media.width,
      media.height,
      crop.x,
      crop.y,
      crop.width,
      crop.height,
    ].every(Number.isFinite) ||
    media.width <= 0 ||
    media.height <= 0 ||
    crop.width <= 0 ||
    crop.height <= 0 ||
    pdfJsView.length !== 4 ||
    !pdfJsView.every(Number.isFinite)
  ) {
    return null;
  }
  const effectiveCropLeft = Math.max(media.x, crop.x);
  const effectiveCropBottom = Math.max(media.y, crop.y);
  const effectiveCropRight = Math.min(
    media.x + media.width,
    crop.x + crop.width,
  );
  const effectiveCropTop = Math.min(
    media.y + media.height,
    crop.y + crop.height,
  );
  if (
    effectiveCropRight <= effectiveCropLeft ||
    effectiveCropTop <= effectiveCropBottom ||
    !sameDeskewCoordinate(pdfJsView[0], effectiveCropLeft) ||
    !sameDeskewCoordinate(pdfJsView[1], effectiveCropBottom) ||
    !sameDeskewCoordinate(pdfJsView[2], effectiveCropRight) ||
    !sameDeskewCoordinate(pdfJsView[3], effectiveCropTop)
  ) {
    return null;
  }

  const mediaWidth = media.width * sourceUserUnit;
  const mediaHeight = media.height * sourceUserUnit;
  const cropX = (effectiveCropLeft - media.x) * sourceUserUnit;
  const cropY = (effectiveCropBottom - media.y) * sourceUserUnit;
  const cropWidth = (effectiveCropRight - effectiveCropLeft) * sourceUserUnit;
  const cropHeight = (effectiveCropTop - effectiveCropBottom) * sourceUserUnit;
  if (
    ![mediaWidth, mediaHeight, cropX, cropY, cropWidth, cropHeight].every(
      Number.isFinite,
    ) ||
    mediaWidth <= 0 ||
    mediaHeight <= 0 ||
    cropWidth <= 0 ||
    cropHeight <= 0
  ) {
    return null;
  }

  let outputMediaWidth = mediaWidth;
  let outputMediaHeight = mediaHeight;
  let outputCropX = cropX;
  let outputCropY = cropY;
  let outputCropWidth = cropWidth;
  let outputCropHeight = cropHeight;
  if (rotation === 90) {
    outputMediaWidth = mediaHeight;
    outputMediaHeight = mediaWidth;
    outputCropX = cropY;
    outputCropY = mediaWidth - cropX - cropWidth;
    outputCropWidth = cropHeight;
    outputCropHeight = cropWidth;
  } else if (rotation === 180) {
    outputCropX = mediaWidth - cropX - cropWidth;
    outputCropY = mediaHeight - cropY - cropHeight;
  } else if (rotation === 270) {
    outputMediaWidth = mediaHeight;
    outputMediaHeight = mediaWidth;
    outputCropX = mediaHeight - cropY - cropHeight;
    outputCropY = cropX;
    outputCropWidth = cropHeight;
    outputCropHeight = cropWidth;
  }
  const zero = (value: number) => (Object.is(value, -0) ? 0 : value);
  return Object.freeze({
    media: Object.freeze({
      x: 0,
      y: 0,
      width: outputMediaWidth,
      height: outputMediaHeight,
    }),
    crop: Object.freeze({
      x: zero(outputCropX),
      y: zero(outputCropY),
      width: outputCropWidth,
      height: outputCropHeight,
    }),
  });
}
