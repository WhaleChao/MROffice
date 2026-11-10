/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The hue word Office puts before a theme slot in a swatch's screen tip:
 * "White, Background 1", "Dark Teal, Accent 1", "Orange, Accent 2"
 * (XLS-A005). Office names the colour the slot HOLDS, so the word follows
 * the workbook's theme. The vocabulary is the small one its tips use; the
 * boundaries are read off Office's own themes.
 */
export type ColorHueName =
  | "white"
  | "black"
  | "gray"
  | "darkRed"
  | "red"
  | "brown"
  | "orange"
  | "gold"
  | "yellow"
  | "lightGreen"
  | "green"
  | "teal"
  | "lightBlue"
  | "blue"
  | "darkBlue"
  | "purple"
  | "pink";

export function colorHueName(hex: string): ColorHueName | null {
  const match = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!match) return null;
  const value = Number.parseInt(match[1], 16);
  const r = ((value >> 16) & 255) / 255;
  const g = ((value >> 8) & 255) / 255;
  const b = (value & 255) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  if (lightness >= 0.97 || (lightness >= 0.93 && saturation < 0.2)) return "white";
  if (lightness <= 0.08) return "black";
  if (saturation < 0.12) return "gray";
  let hue =
    max === r ? ((g - b) / delta) % 6 : max === g ? (b - r) / delta + 2 : (r - g) / delta + 4;
  hue = (hue * 60 + 360) % 360;
  if (hue < 15 || hue >= 345) return lightness < 0.3 ? "darkRed" : "red";
  if (hue < 40) return lightness < 0.3 ? "brown" : "orange";
  if (hue < 50) return "gold";
  if (hue < 70) return "yellow";
  if (hue < 160) return lightness > 0.6 ? "lightGreen" : "green";
  if (hue < 190 || (hue < 205 && lightness < 0.35)) return "teal";
  if (hue < 250) return lightness < 0.25 ? "darkBlue" : lightness > 0.65 ? "lightBlue" : "blue";
  if (hue < 300) return "purple";
  return lightness < 0.45 ? "purple" : "pink";
}
