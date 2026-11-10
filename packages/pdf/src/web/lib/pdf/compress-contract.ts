/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The PDF compression contract the editor and the writer share: the presets,
 * options and result shape. The lossless writer itself is server-only
 * (`compress-lossless.ts`, engine `pdf-document`).
 */

/**
 * Optimization aggressiveness. Only "basic" exists today — it is the lossless
 * structural pass described above. The enum is here so a future lossy tier
 * (image downsampling, font subsetting) can be added without changing the
 * signature. Unknown values fall back to "basic".
 */
export type PdfCompressLevel = "basic";

export const PDF_COMPRESS_LOSSLESS_PRESET =
  "lossless-structural-v1" as const;
export type PdfCompressLosslessPreset =
  typeof PDF_COMPRESS_LOSSLESS_PRESET;

/**
 * Server-side lossy quality tiers (Ghostscript vocabulary), mapped by the
 * backend to a downsampling target DPI + JPEG re-encode quality:
 * screen = 72 DPI / q60, ebook = 150 / q75, print = 300 / q85.
 */
export type PdfCompressServerPreset = "screen" | "ebook" | "print";

export const PDF_COMPRESS_SERVER_PRESETS: readonly PdfCompressServerPreset[] = [
  "screen",
  "ebook",
  "print",
];

export interface CompressPdfOptions {
  /** Exact local writer contract. Other compression presets are server-owned. */
  preset?: PdfCompressLosslessPreset;
  /** Optimization tier; only "basic" (lossless structural) is implemented. */
  level?: PdfCompressLevel;
  /**
   * pdf-lib serializes large documents in chunks of this many objects per
   * event-loop tick (its `objectsPerTick`). Larger = fewer yields = faster but
   * more main-thread blocking. Defaults to pdf-lib's own default (50). Exposed
   * so a caller on a huge doc can trade responsiveness for throughput.
   */
  objectsPerTick?: number;
}

export interface CompressPdfResult {
  /**
   * The optimized bytes. Guaranteed to be no larger than the input: if the
   * re-save did not shrink the file, this is the ORIGINAL `bytes` (copied), so a
   * download never grows the file.
   */
  bytes: Uint8Array;
  /** Original size in bytes. */
  originalSize: number;
  /** Output size in bytes (=== `bytes.byteLength`). */
  compressedSize: number;
  /**
   * Bytes saved = `originalSize - compressedSize`, clamped to >= 0. Zero is a
   * legitimate, honest outcome for an already-optimized PDF — do NOT surface a
   * fabricated saving when this is 0.
   */
  savedBytes: number;
  /** Fraction saved in [0, 1] (`savedBytes / originalSize`), 0 when input empty. */
  savedRatio: number;
}
