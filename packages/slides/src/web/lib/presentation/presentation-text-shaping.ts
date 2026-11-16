/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";























export function presentationTextShapingCss(
  kerningFrom: unknown,
  runSizesPt: readonly (number | undefined)[] | null,
  autofitScale = 1,
): CSSProperties {
  return {
    fontVariantLigatures: "no-common-ligatures",
    ...(presentationBodyUnkerned(kerningFrom, runSizesPt, autofitScale) ? { fontKerning: "none" as const } : {}),
  };
}

/** Whether PowerPoint kerns none of a body's pairs (see `presentationTextShapingCss`). */
export function presentationBodyUnkerned(
  kerningFrom: unknown,
  runSizesPt: readonly (number | undefined)[] | null,
  autofitScale = 1,
): boolean {
  if (kerningFrom === 0) return true;
  if (typeof kerningFrom !== "number" || !Number.isFinite(kerningFrom) || kerningFrom < 0) return false;
  if (!runSizesPt) return false;
  const scale = Number.isFinite(autofitScale) && autofitScale > 0 ? autofitScale : 1;
  return runSizesPt.every((size) => typeof size === "number" && size * scale < kerningFrom);
}
