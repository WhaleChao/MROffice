/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";











export function presentationPictureOrientationCss(block: object): CSSProperties {
  return (block as { exifOrientation?: unknown }).exifOrientation === undefined ? { imageOrientation: "none" } : {};
}
