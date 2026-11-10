/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The native table-style GALLERY.
 *
 * PowerPoint ships roughly 74 built-in table styles; this editor shipped six,
 * because the six were the only Office GUIDs we could name with confidence and
 * inventing more would be worse than having fewer — a fabricated GUID makes
 * PowerPoint render something different from what the editor shows.
 *
 * The way out is not to guess Microsoft's identifiers but to AUTHOR our own
 * styles. A `<a:tblStyle>` definition travels inside the package, so a style
 * this app defines renders identically in PowerPoint and WPS without either
 * having heard of it — the same mechanism that lets an imported deck bring its
 * own gallery. Each entry below therefore carries a stable MRAI GUID plus the
 * complete resolved look, which the canvas paints directly and the exporter
 * writes into `ppt/tableStyles.xml`.
 *
 * The looks follow Office's own construction: one accent colour, tinted toward
 * white for the body and banding, at full strength for the header, so a deck
 * built here reads as a coherent family rather than a set of one-off swatches.
 *
 * Pure and deterministic — no React, no DOM, no clock. The GUIDs are literals,
 * never generated, so the same style keeps the same identity across sessions and
 * a re-export produces a byte-identical definition.
 */

import type { TableStyleLook, TableStyleLookPart } from "./presentation-table-style";

/** The six accents of Office's current default theme. */
export const GALLERY_ACCENTS = [
  { key: "accent1", label: "强调 1 / Accent 1", color: "#156082" },
  { key: "accent2", label: "强调 2 / Accent 2", color: "#e97132" },
  { key: "accent3", label: "强调 3 / Accent 3", color: "#196b24" },
  { key: "accent4", label: "强调 4 / Accent 4", color: "#0f9ed5" },
  { key: "accent5", label: "强调 5 / Accent 5", color: "#a02b93" },
  { key: "accent6", label: "强调 6 / Accent 6", color: "#4ea72e" },
] as const;
export type GalleryAccentKey = (typeof GALLERY_ACCENTS)[number]["key"];

/** The three intensity families, matching Office's Light / Medium / Dark. */
export const GALLERY_FAMILIES = ["light", "medium", "dark"] as const;
export type GalleryFamily = (typeof GALLERY_FAMILIES)[number];

const FAMILY_LABELS: Record<GalleryFamily, string> = {
  light: "浅色 / Light",
  medium: "中度 / Medium",
  dark: "深色 / Dark",
};

/**
 * Stable per-style GUIDs. Literal on purpose: a generated id would change the
 * style's identity between sessions, so the same table would export as a
 * different style every time it was saved.
 *
 * The `4D524149` prefix is "MRAI" in ASCII hex — a marker that says these are
 * this app's own styles and not a collision with an Office built-in.
 */
const GALLERY_STYLE_IDS: Record<GalleryFamily, Record<GalleryAccentKey, string>> =
  {
    light: {
      accent1: "{4D524149-0001-4C31-9A01-000000000001}",
      accent2: "{4D524149-0001-4C31-9A01-000000000002}",
      accent3: "{4D524149-0001-4C31-9A01-000000000003}",
      accent4: "{4D524149-0001-4C31-9A01-000000000004}",
      accent5: "{4D524149-0001-4C31-9A01-000000000005}",
      accent6: "{4D524149-0001-4C31-9A01-000000000006}",
    },
    medium: {
      accent1: "{4D524149-0002-4D31-9A02-000000000001}",
      accent2: "{4D524149-0002-4D31-9A02-000000000002}",
      accent3: "{4D524149-0002-4D31-9A02-000000000003}",
      accent4: "{4D524149-0002-4D31-9A02-000000000004}",
      accent5: "{4D524149-0002-4D31-9A02-000000000005}",
      accent6: "{4D524149-0002-4D31-9A02-000000000006}",
    },
    dark: {
      accent1: "{4D524149-0003-4443-9A03-000000000001}",
      accent2: "{4D524149-0003-4443-9A03-000000000002}",
      accent3: "{4D524149-0003-4443-9A03-000000000003}",
      accent4: "{4D524149-0003-4443-9A03-000000000004}",
      accent5: "{4D524149-0003-4443-9A03-000000000005}",
      accent6: "{4D524149-0003-4443-9A03-000000000006}",
    },
  };

function clampChannel(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

const GALLERY_HEX_RE = /^#?[0-9a-fA-F]{6}$/;

/** `null` for anything that is not a six-digit hex colour. */
function toRgb(hex: string): [number, number, number] | null {
  if (typeof hex !== "string" || !GALLERY_HEX_RE.test(hex.trim())) return null;
  const raw = hex.trim().replace("#", "");
  return [
    Number.parseInt(raw.slice(0, 2), 16),
    Number.parseInt(raw.slice(2, 4), 16),
    Number.parseInt(raw.slice(4, 6), 16),
  ];
}

function toHex(rgb: readonly [number, number, number]): string {
  return `#${rgb.map((c) => clampChannel(c).toString(16).padStart(2, "0")).join("")}`;
}

/**
 * DrawingML's own tint: mix toward WHITE by `1 - amount`. `tint(0.2)` is the 20%
 * tint Office uses for a Medium style's body fill, so the values here match what
 * PowerPoint computes from the same accent.
 */
export function tintColor(hex: string, amount: number): string {
  const rgb = toRgb(hex);
  // A colour this cannot read comes back unchanged: `#NaNNaNNaN` is worse than
  // the input, and these helpers are exported.
  if (!rgb || !Number.isFinite(amount)) return hex;
  const mix = (channel: number) => channel * amount + 255 * (1 - amount);
  return toHex([mix(rgb[0]), mix(rgb[1]), mix(rgb[2])]);
}

/** DrawingML's shade: mix toward BLACK. */
export function shadeColor(hex: string, amount: number): string {
  const rgb = toRgb(hex);
  if (!rgb || !Number.isFinite(amount)) return hex;
  return toHex([rgb[0] * amount, rgb[1] * amount, rgb[2] * amount]);
}

function border(color: string, width: number) {
  return { style: "solid" as const, color, width };
}

/** The resolved look for one family / accent pair. */
export function galleryStyleLook(
  family: GalleryFamily,
  accent: (typeof GALLERY_ACCENTS)[number],
): TableStyleLook {
  const styleId = GALLERY_STYLE_IDS[family][accent.key];
  const styleName = `${FAMILY_LABELS[family]} · ${accent.label}`;
  const emphasis: TableStyleLookPart = { bold: true };
  if (family === "light") {
    return {
      styleId,
      styleName,
      parts: {
        wholeTbl: {
          fill: "#ffffff",
          borders: {
            top: border(accent.color, 1.5),
            bottom: border(accent.color, 1.5),
            insideH: border(tintColor(accent.color, 0.4), 1),
          },
        },
        band1H: { fill: tintColor(accent.color, 0.2) },
        firstRow: {
          ...emphasis,
          color: accent.color,
          borders: { bottom: border(accent.color, 1.5) },
        },
        lastRow: { ...emphasis, borders: { top: border(accent.color, 1.5) } },
        firstCol: emphasis,
        lastCol: emphasis,
      },
    };
  }
  if (family === "medium") {
    return {
      styleId,
      styleName,
      parts: {
        wholeTbl: {
          fill: tintColor(accent.color, 0.2),
          borders: {
            left: border("#ffffff", 1),
            right: border("#ffffff", 1),
            top: border("#ffffff", 1),
            bottom: border("#ffffff", 1),
            insideH: border("#ffffff", 1),
            insideV: border("#ffffff", 1),
          },
        },
        band1H: { fill: tintColor(accent.color, 0.4) },
        band1V: { fill: tintColor(accent.color, 0.4) },
        firstRow: {
          ...emphasis,
          fill: accent.color,
          color: "#ffffff",
          borders: { bottom: border("#ffffff", 3) },
        },
        lastRow: {
          ...emphasis,
          fill: accent.color,
          color: "#ffffff",
          borders: { top: border("#ffffff", 3) },
        },
        firstCol: { ...emphasis, fill: accent.color, color: "#ffffff" },
        lastCol: { ...emphasis, fill: accent.color, color: "#ffffff" },
      },
    };
  }
  const deep = shadeColor(accent.color, 0.65);
  return {
    styleId,
    styleName,
    parts: {
      wholeTbl: {
        fill: accent.color,
        color: "#ffffff",
        borders: {
          insideH: border(tintColor(accent.color, 0.6), 1),
          insideV: border(tintColor(accent.color, 0.6), 1),
        },
      },
      band1H: { fill: shadeColor(accent.color, 0.85) },
      firstRow: {
        ...emphasis,
        fill: deep,
        color: "#ffffff",
        borders: { bottom: border("#ffffff", 2) },
      },
      lastRow: {
        ...emphasis,
        fill: deep,
        color: "#ffffff",
        borders: { top: border("#ffffff", 2) },
      },
      firstCol: { ...emphasis, fill: deep, color: "#ffffff" },
      lastCol: { ...emphasis, fill: deep, color: "#ffffff" },
    },
  };
}

export interface GalleryStyle {
  id: string;
  family: GalleryFamily;
  familyLabel: string;
  accentLabel: string;
  look: TableStyleLook;
}

/** The whole authored gallery: three families across six accents. */
export const TABLE_STYLE_GALLERY: readonly GalleryStyle[] = GALLERY_FAMILIES.flatMap(
  (family) =>
    GALLERY_ACCENTS.map((accent) => ({
      id: GALLERY_STYLE_IDS[family][accent.key],
      family,
      familyLabel: FAMILY_LABELS[family],
      accentLabel: accent.label,
      look: galleryStyleLook(family, accent),
    })),
);

const GALLERY_BY_ID = new Map(
  TABLE_STYLE_GALLERY.map((style) => [style.id, style]),
);

/**
 * The authored style for a GUID, or `null`.
 *
 * This is what lets a deck survive a round trip through a tool that rewrote
 * `ppt/tableStyles.xml` without keeping unknown entries: the definition is gone
 * from the package but the GUID is still on the table, and the gallery holds the
 * look — so the table renders as authored rather than as a plain grid.
 */
export function galleryStyleById(value: unknown): GalleryStyle | null {
  return typeof value === "string" ? (GALLERY_BY_ID.get(value) ?? null) : null;
}
