/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Original point size paired with its legacy CSS projection. Never infer an
 * original point size from a rounded pixel value on an older document. */
export interface NativeFontSize {
  version: 1;
  points: number;
  pixels: number;
}

export function normalizeNativeFontSize(value: unknown): NativeFontSize | undefined {
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const pair = value as Record<string, unknown>;
  if (pair.version !== 1 || Object.keys(pair).some(key => !["version", "points", "pixels"].includes(key)) ||
      typeof pair.points !== "number" || !Number.isFinite(pair.points) || pair.points <= 0 || pair.points > 409 ||
      typeof pair.pixels !== "number" || !Number.isFinite(pair.pixels) || pair.pixels <= 0 || pair.pixels > 546) return undefined;
  // Accept the precise producer or the two-decimal legacy producer only.
  // Their bounds are checked by the importer that created the pair.
  const precise = pair.points / .75;
  // Python and JavaScript legacy producers differ on decimal half ties.
  // Accept a hundredth only at the genuine nearest-rounding boundary.
  const legacy = pair.pixels === Math.round(pair.pixels * 100) / 100 &&
    Math.abs(pair.pixels - precise) <= .005 + Number.EPSILON * Math.max(1, Math.abs(precise));
  if (pair.pixels !== precise && !legacy) return undefined;
  return { version: 1, points: pair.points, pixels: pair.pixels };
}

export function nativeFontSizeCss(pixels: number | undefined, source: unknown): number | undefined {
  const pair = normalizeNativeFontSize(source);
  return pair && pixels === pair.pixels ? pair.points / .75 : pixels;
}

/** An explicit size edit invalidates even an unchanged legacy pixel mirror.
 * Color, weight or family edits alone do not change the source point size. */
export function nativeFontSizeAfterPatch(
  previous: unknown,
  patch: { fontSize?: unknown; nativeFontSize?: unknown },
): NativeFontSize | null | undefined {
  if (Object.prototype.hasOwnProperty.call(patch, "nativeFontSize")) {
    return patch.nativeFontSize === null ? null : normalizeNativeFontSize(patch.nativeFontSize);
  }
  return Object.prototype.hasOwnProperty.call(patch, "fontSize")
    ? previous === undefined && patch.fontSize === undefined ? undefined : null
    : previous === null ? null : normalizeNativeFontSize(previous);
}
