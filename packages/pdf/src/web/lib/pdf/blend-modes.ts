/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
export const PDF_STANDARD_BLEND_MODES = [
  "normal",
  "multiply",
  "screen",
  "overlay",
  "darken",
  "lighten",
  "color-dodge",
  "color-burn",
  "hard-light",
  "soft-light",
  "difference",
  "exclusion",
  "hue",
  "saturation",
  "color",
  "luminosity",
] as const;

export type PdfStandardBlendMode = (typeof PDF_STANDARD_BLEND_MODES)[number];

export const PDF_BLEND_MODE_BY_PDF_NAME: Readonly<
  Record<string, PdfStandardBlendMode>
> = Object.freeze({
  Normal: "normal",
  Compatible: "normal",
  Multiply: "multiply",
  Screen: "screen",
  Overlay: "overlay",
  Darken: "darken",
  Lighten: "lighten",
  ColorDodge: "color-dodge",
  ColorBurn: "color-burn",
  HardLight: "hard-light",
  SoftLight: "soft-light",
  Difference: "difference",
  Exclusion: "exclusion",
  Hue: "hue",
  Saturation: "saturation",
  Color: "color",
  Luminosity: "luminosity",
});

export const PDF_NAME_BY_BLEND_MODE: Readonly<
  Record<PdfStandardBlendMode, string>
> = Object.freeze({
  normal: "Normal",
  multiply: "Multiply",
  screen: "Screen",
  overlay: "Overlay",
  darken: "Darken",
  lighten: "Lighten",
  "color-dodge": "ColorDodge",
  "color-burn": "ColorBurn",
  "hard-light": "HardLight",
  "soft-light": "SoftLight",
  difference: "Difference",
  exclusion: "Exclusion",
  hue: "Hue",
  saturation: "Saturation",
  color: "Color",
  luminosity: "Luminosity",
});

const PDF_STANDARD_BLEND_MODE_SET: ReadonlySet<string> = new Set(
  PDF_STANDARD_BLEND_MODES,
);

export function isPdfStandardBlendMode(
  value: unknown,
): value is PdfStandardBlendMode {
  return typeof value === "string" && PDF_STANDARD_BLEND_MODE_SET.has(value);
}
