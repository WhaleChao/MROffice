/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export function inkLineScale(
  from: { readonly width: number; readonly height: number },
  to: { readonly width: number; readonly height: number },
): number {
  const scale = Math.sqrt((to.width / Math.max(1e-9, from.width)) * (to.height / Math.max(1e-9, from.height)));
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

/** A line width scaled with its stroke's box, as the file keeps it (hundredths, at least 0.1). */
export function inkLineWidthScaled(
  width: number,
  from: { readonly width: number; readonly height: number },
  to: { readonly width: number; readonly height: number },
  max = Number.POSITIVE_INFINITY,
): number {
  if (!(width > 0)) return width;
  const scale = inkLineScale(from, to);
  if (scale === 1) return width;
  return Math.min(max, Math.max(0.1, Math.round(width * scale * 100) / 100));
}
