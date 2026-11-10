/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export interface IdentityColor {
  id: string;
  hex: string;
}

export const IDENTITY_COLORS: readonly IdentityColor[] = [
  { id: "red", hex: "#ce514d" },
  { id: "coral", hex: "#cc572a" },
  { id: "orange", hex: "#c0640c" },
  { id: "amber", hex: "#af7112" },
  { id: "gold", hex: "#a1790c" },
  { id: "olive", hex: "#918112" },
  { id: "lime", hex: "#7e8904" },
  { id: "grass", hex: "#5f9205" },
  { id: "emerald", hex: "#31983d" },
  { id: "mint", hex: "#049863" },
  { id: "pine", hex: "#11957c" },
  { id: "teal", hex: "#14938d" },
  { id: "cyan", hex: "#0f919c" },
  { id: "lake", hex: "#158eab" },
  { id: "sky", hex: "#0b8bc0" },
  { id: "blue", hex: "#1a83db" },
  { id: "indigo", hex: "#507adf" },
  { id: "violet", hex: "#7071dd" },
  { id: "purple", hex: "#8968d4" },
  { id: "wisteria", hex: "#9d60c6" },
  { id: "magenta", hex: "#ae59b4" },
  { id: "pink", hex: "#bc549e" },
  { id: "rose", hex: "#c65085" },
  { id: "berry", hex: "#cc4f6a" },
  { id: "cocoa", hex: "#926a55" },
  { id: "stone", hex: "#897e70" },
  { id: "slate", hex: "#6d7c8f" },
  { id: "graphite", hex: "#636363" },
];

/** The palette entry for a stored colour, or null when it is not one of ours
 *  (an older choice made from the theme swatches, for instance). */
export function identityColorByHex(
  hex: string | null | undefined,
): IdentityColor | null {
  if (!hex) return null;
  const wanted = hex.toLowerCase();
  return IDENTITY_COLORS.find((color) => color.hex === wanted) ?? null;
}

/**
 * The colour a person carries before choosing one: a stable pick from the
 * hues (never a neutral) by their email or name, so it is the same on every
 * surface and does not follow the viewer's theme.
 */
export function identityColorFor(seed: string | null | undefined): string {
  const text = (seed ?? "").trim().toLowerCase();
  let hash = 2166136261;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return IDENTITY_COLORS[(hash >>> 0) % 24].hex;
}
