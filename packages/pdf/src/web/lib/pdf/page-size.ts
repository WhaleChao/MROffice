/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Shared PDF page-size catalog used by blank-page creation and page resizing. */

export const PDF_PAGE_SIZE_MIN_POINTS = 36;
export const PDF_PAGE_SIZE_MAX_POINTS = 14_400;
export const PDF_POINTS_PER_MILLIMETRE = 72 / 25.4;

export const PDF_PAGE_SIZE_PRESETS = Object.freeze({
  letter: Object.freeze({ label: "Letter", width: 612, height: 792 }),
  legal: Object.freeze({ label: "Legal", width: 612, height: 1008 }),
  a0: Object.freeze({ label: "A0", width: 2383.937, height: 3370.394 }),
  a1: Object.freeze({ label: "A1", width: 1683.78, height: 2383.937 }),
  a2: Object.freeze({ label: "A2", width: 1190.551, height: 1683.78 }),
  a3: Object.freeze({ label: "A3", width: 841.89, height: 1190.551 }),
  a4: Object.freeze({ label: "A4", width: 595.28, height: 841.89 }),
  a5: Object.freeze({ label: "A5", width: 419.528, height: 595.276 }),
  a6: Object.freeze({ label: "A6", width: 297.638, height: 419.528 }),
  b1: Object.freeze({ label: "B1", width: 2004.094, height: 2834.646 }),
  b2: Object.freeze({ label: "B2", width: 1417.323, height: 2004.094 }),
  b3: Object.freeze({ label: "B3", width: 1000.63, height: 1417.323 }),
  b4: Object.freeze({ label: "B4", width: 708.661, height: 1000.63 }),
  b5: Object.freeze({ label: "B5", width: 498.898, height: 708.661 }),
  b6: Object.freeze({ label: "B6", width: 354.331, height: 498.898 }),
});

export type PdfPageSizePresetKey = keyof typeof PDF_PAGE_SIZE_PRESETS;
export type PdfPageSizePreset = "original" | "custom" | PdfPageSizePresetKey;
export type PdfPageSizeOrientation = "portrait" | "landscape";
export type PdfPageContentPlacement = "unchanged" | "fit-center";
export type PdfPageSizeScope = "page" | "all" | "range";

export interface PdfPageSizeSelection {
  scope: PdfPageSizeScope;
  rangeSpec: string;
  width: number | null;
  height: number | null;
  contentPlacement: PdfPageContentPlacement;
}

export interface PdfPageSizeDimensions {
  width: number;
  height: number;
}

export function orientPdfPageSize(
  size: PdfPageSizeDimensions,
  orientation: PdfPageSizeOrientation,
): PdfPageSizeDimensions {
  const shortEdge = Math.min(size.width, size.height);
  const longEdge = Math.max(size.width, size.height);
  return orientation === "portrait"
    ? { width: shortEdge, height: longEdge }
    : { width: longEdge, height: shortEdge };
}

export function pdfPointsToMillimetres(points: number): number {
  return points / PDF_POINTS_PER_MILLIMETRE;
}

export function pdfMillimetresToPoints(millimetres: number): number {
  return millimetres * PDF_POINTS_PER_MILLIMETRE;
}

export function isPdfPageSizeDimension(value: number): boolean {
  return (
    Number.isFinite(value) &&
    value >= PDF_PAGE_SIZE_MIN_POINTS &&
    value <= PDF_PAGE_SIZE_MAX_POINTS
  );
}

export function resolvePdfPageSizeTarget(input: {
  preset: PdfPageSizePreset;
  orientation: PdfPageSizeOrientation;
  customWidthMm: string;
  customHeightMm: string;
  sourceSize: PdfPageSizeDimensions | null;
}): PdfPageSizeDimensions | null {
  if (input.preset === "original") return input.sourceSize;
  const base =
    input.preset === "custom"
      ? {
          width: pdfMillimetresToPoints(Number.parseFloat(input.customWidthMm)),
          height: pdfMillimetresToPoints(
            Number.parseFloat(input.customHeightMm),
          ),
        }
      : PDF_PAGE_SIZE_PRESETS[input.preset];
  if (
    !isPdfPageSizeDimension(base.width) ||
    !isPdfPageSizeDimension(base.height)
  ) {
    return null;
  }
  // Custom width and height are literal, as in Word's Page Setup: the
  // orientation follows them (the dialog swaps the two fields when the
  // orientation is switched), so they are never re-sorted here.
  if (input.preset === "custom") return base;
  return orientPdfPageSize(base, input.orientation);
}

/** The orientation a literal width and height describe (square = portrait). */
export function pdfPageSizeOrientationOf(
  size: PdfPageSizeDimensions,
): PdfPageSizeOrientation {
  return size.width > size.height ? "landscape" : "portrait";
}

export function pageSizePresetForDimensions(
  size: PdfPageSizeDimensions | null,
): PdfPageSizePresetKey | null {
  if (!size) return null;
  for (const [key, preset] of Object.entries(PDF_PAGE_SIZE_PRESETS)) {
    const direct =
      Math.abs(size.width - preset.width) <= 0.5 &&
      Math.abs(size.height - preset.height) <= 0.5;
    const rotated =
      Math.abs(size.width - preset.height) <= 0.5 &&
      Math.abs(size.height - preset.width) <= 0.5;
    if (direct || rotated) return key as PdfPageSizePresetKey;
  }
  return null;
}
