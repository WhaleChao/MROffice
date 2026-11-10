/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export { pdfNativeEmbeddedImagePlacement } from "./engine/native/embedded-image-model";
export type {
  PdfNativeEmbeddedImageTransformFrame,
  PdfNativeEmbeddedImageTransformIntent,
} from "./engine/native/content-stream-transaction";

/** Editor-side geometry an image edit accumulates before a native transform. */
export interface EmbeddedImageNativeTransformState {
  /** The paint CTM the reader walked when the edit started (user space). */
  readonly ctm: readonly [number, number, number, number, number, number];
  /** Clockwise quarter turns applied on screen, 0..3. */
  readonly quarterTurns: 0 | 1 | 2 | 3;
  readonly flipX: boolean;
  readonly flipY: boolean;
}

export function turnedQuarterTurns(
  current: 0 | 1 | 2 | 3,
  direction: "left" | "right",
): 0 | 1 | 2 | 3 {
  return ((current + (direction === "right" ? 1 : 3)) % 4) as 0 | 1 | 2 | 3;
}
