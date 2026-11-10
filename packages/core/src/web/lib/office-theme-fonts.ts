/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Office's theme font schemes as the editors draw them: one heading (major)
 * and one body (minor) face per scheme, with the CSS stacks a browser uses to
 * render them. The Fonts gallery is the same list in every Office host --
 * Excel's Page Layout > Themes > Fonts and PowerPoint's Design > Variants >
 * Fonts both read Office's one "Theme Fonts" folder -- so both editors take
 * their entries from here (lib/office-theme-catalog is the data).
 */

import {
  OFFICE_FONT_SCHEMES,
  type OfficeThemeFontCollection,
  type OfficeThemeLabel,
} from "./office-theme-catalog";

export interface OfficeThemeFontSet {
  readonly id: string;
  readonly label: OfficeThemeLabel;
  /** The scheme's name in its XML. */
  readonly schemeName: string;
  /** Listed in the Fonts gallery (a theme's own fonts are not). */
  readonly gallery: boolean;
  readonly major: OfficeThemeFontCollection;
  readonly minor: OfficeThemeFontCollection;
  readonly headingLatin: string;
  readonly bodyLatin: string;
  /** The face Chinese text uses: the scheme's East Asian face, else its Hans face. */
  readonly headingEastAsia: string;
  readonly bodyEastAsia: string;
  readonly headingFamily: string;
  readonly bodyFamily: string;
}

function quotedFont(face: string): string {
  return `"${face.replace(/["'\\;{}()]/g, "").slice(0, 200)}"`;
}

/** Faces a browser may have when it lacks Office's narrow and display cuts. */
const LATIN_FALLBACKS: Readonly<Record<string, readonly string[]>> = {
  "aptos narrow": ["Aptos"],
  "aptos display": ["Aptos"],
};

function eastAsianFace(collection: OfficeThemeFontCollection): string {
  return collection.ea || collection.scripts.Hans || "";
}

/** The CSS stack that renders one theme font collection. */
export function officeThemeFontStack(collection: OfficeThemeFontCollection): string {
  const faces: string[] = [];
  const seen = new Set<string>();
  for (const face of [
    collection.latin,
    ...(LATIN_FALLBACKS[collection.latin.trim().toLowerCase()] ?? []),
    eastAsianFace(collection),
  ]) {
    const trimmed = face.trim();
    const key = trimmed.toLowerCase();
    if (!trimmed || seen.has(key)) continue;
    seen.add(key);
    faces.push(quotedFont(trimmed));
  }
  return [...faces, '"PingFang SC"', '"Microsoft YaHei"', "sans-serif"].join(", ");
}

export function officeThemeFontSet(
  id: string,
  label: OfficeThemeLabel,
  schemeName: string,
  gallery: boolean,
  major: OfficeThemeFontCollection,
  minor: OfficeThemeFontCollection,
): OfficeThemeFontSet {
  return Object.freeze({
    id,
    label,
    schemeName,
    gallery,
    major,
    minor,
    headingLatin: major.latin,
    bodyLatin: minor.latin,
    headingEastAsia: eastAsianFace(major),
    bodyEastAsia: eastAsianFace(minor),
    headingFamily: officeThemeFontStack(major),
    bodyFamily: officeThemeFontStack(minor),
  });
}

/** Every Office font scheme: the Fonts gallery in Office's order, then each theme's own. */
export const OFFICE_THEME_FONT_SETS: readonly OfficeThemeFontSet[] = Object.freeze(
  OFFICE_FONT_SCHEMES.map((scheme) =>
    officeThemeFontSet(
      scheme.id,
      scheme.label,
      scheme.schemeName,
      scheme.gallery,
      scheme.major,
      scheme.minor,
    ),
  ),
);

/** The Fonts gallery, in Office's order. */
export const OFFICE_THEME_FONT_GALLERY: readonly OfficeThemeFontSet[] = Object.freeze(
  OFFICE_THEME_FONT_SETS.filter((font) => font.gallery),
);

const GALLERY_BY_ID = new Map(OFFICE_THEME_FONT_GALLERY.map((font) => [font.id, font]));

/** One Fonts-gallery entry, or undefined for anything the gallery does not list. */
export function officeThemeFontGalleryEntry(id: unknown): OfficeThemeFontSet | undefined {
  return typeof id === "string" ? GALLERY_BY_ID.get(id) : undefined;
}
