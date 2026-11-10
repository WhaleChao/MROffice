/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { coverArt, isMovingCoverArt, type CoverArtId } from "./cover-art";
import { COVER_PALETTES, type CoverPalette } from "./cover-palette";

// Per-notebook cover customization. New writes are persisted by the API in
// RDS; localStorage is kept as a legacy fallback for covers created before
// the server-side field existed.

// Keep in step with NotebookCoverTemplate in apps/api/app/schemas/notebook.py:
// the API rejects (422) a template it does not know.
export type CoverTemplate =
  | "solid"
  | "halfImage"
  | "fullImage"
  | "framed"
  | "ruled"
  | "cloth"
  | "journal"
  | "spiral"
  | "kraft"
  | "marble"
  | "magazine"
  | "aurora"
  | "borealis"
  | "nebula"
  | "tide"
  | "holo"
  | "rain"
  | "fireflies"
  | "snowfall"
  | "blossom"
  | "inkwash"
  | "koi"
  // a painted cover from lib/cover-art.ts
  | CoverArtId;

export interface CoverPreference {
  template: CoverTemplate;
  color: string;
  spineColor: string;
  textTheme: "light" | "dark";
  imageUrl?: string;
  /** The owner picked the colour (a swatch or the wheel). Until then a
   *  picture cover gives the notebook its own colour (lib/cover-palette).
   *  Absent (null from the API) on covers saved before that. */
  colorPicked?: boolean | null;
}

// Cover images are server-persisted, so the backend
// (_validate_cover_preference_image_url in apps/api/app/schemas/notebook.py)
// is the authority. Mirror its policy exactly to avoid a 422 on create/save:
// https only (blob: is session-scoped/non-persistable, http: is insecure),
// raster data URLs limited to jpeg/png/webp (no gif), 2 MiB cap.
const COVER_IMAGE_ALLOWED_PROTOCOLS = new Set(["https:"]);
const MAX_COVER_DATA_IMAGE_BYTES = 2 * 1024 * 1024;
const SAFE_COVER_DATA_IMAGE_PATTERN = /^data:image\/(?:png|jpe?g|webp);base64,/i;

function getWindowOrigin(): string {
  if (typeof window === "undefined") {
    return "http://localhost";
  }
  return window.location.origin;
}

export function getSafeCoverImageUrl(value?: string | null): string | null {
  const trimmed = value?.trim();
  if (!trimmed) {
    return null;
  }

  if (trimmed.toLowerCase().startsWith("data:")) {
    if (!SAFE_COVER_DATA_IMAGE_PATTERN.test(trimmed)) {
      return null;
    }
    return getDataUrlDecodedByteLength(trimmed) <= MAX_COVER_DATA_IMAGE_BYTES
      ? trimmed
      : null;
  }

  try {
    const url = new URL(trimmed, getWindowOrigin());
    if (!COVER_IMAGE_ALLOWED_PROTOCOLS.has(url.protocol)) {
      return null;
    }
    return url.toString();
  } catch {
    return null;
  }
}

function getDataUrlDecodedByteLength(dataUrl: string): number {
  const commaIndex = dataUrl.indexOf(",");
  if (commaIndex < 0) {
    return Number.POSITIVE_INFINITY;
  }
  const payload = dataUrl.slice(commaIndex + 1).replace(/\s/g, "");
  if (!payload) {
    return 0;
  }
  const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
  return Math.floor((payload.length * 3) / 4) - padding;
}

export interface ColorSwatch {
  id: string;
  labelKey: string;
  color: string;
  spine: string;
  text: "light" | "dark";
}


// red → black in the familiar office-palette order. Spines are the cover
// cut darker; light covers (yellow / cyan / white) flip to dark ink.
export const COLOR_SWATCHES: ColorSwatch[] = [
  { id: "red", labelKey: "notebookCover.swatches.red", color: "#e5402f", spine: "#93241a", text: "light" },
  { id: "orange", labelKey: "notebookCover.swatches.orange", color: "#f0913a", spine: "#9c5a1b", text: "light" },
  { id: "yellow", labelKey: "notebookCover.swatches.yellow", color: "#f6ef53", spine: "#b8ad25", text: "dark" },
  { id: "green", labelKey: "notebookCover.swatches.green", color: "#46d332", spine: "#2b8a1d", text: "light" },
  { id: "cyan", labelKey: "notebookCover.swatches.cyan", color: "#61e8ee", spine: "#3aa9ae", text: "dark" },
  { id: "blue", labelKey: "notebookCover.swatches.blue", color: "#2a3aee", spine: "#16207f", text: "light" },
  { id: "magenta", labelKey: "notebookCover.swatches.magenta", color: "#ea5df0", spine: "#96339b", text: "light" },
  { id: "purple", labelKey: "notebookCover.swatches.purple", color: "#8f2d9e", spine: "#571c62", text: "light" },
  { id: "brown", labelKey: "notebookCover.swatches.brown", color: "#b98a52", spine: "#71512c", text: "light" },
  { id: "white", labelKey: "notebookCover.swatches.white", color: "#ffffff", spine: "#d8d8d8", text: "dark" },
  { id: "gray", labelKey: "notebookCover.swatches.gray", color: "#9e9e9e", spine: "#616161", text: "light" },
  { id: "black", labelKey: "notebookCover.swatches.black", color: "#1a1a1a", spine: "#000000", text: "light" },
];

export interface TemplateOption {
  id: CoverTemplate;
  labelKey: string;
  hintKey: string;
  needsImage: boolean;
  /** A moving cover: the picker lists it under its own heading. */
  dynamic?: boolean;
}


// material covers first, from plain to crafted, then the graphic ones, then
// the two that carry a picture (they open the image studio). The moving

export const TEMPLATE_OPTIONS: TemplateOption[] = [
  {
    id: "solid",
    labelKey: "notebookCover.templates.solid.label",
    hintKey: "notebookCover.templates.solid.hint",
    needsImage: false,
  },
  {
    id: "framed",
    labelKey: "notebookCover.templates.framed.label",
    hintKey: "notebookCover.templates.framed.hint",
    needsImage: false,
  },
  {
    id: "cloth",
    labelKey: "notebookCover.templates.cloth.label",
    hintKey: "notebookCover.templates.cloth.hint",
    needsImage: false,
  },
  {
    id: "journal",
    labelKey: "notebookCover.templates.journal.label",
    hintKey: "notebookCover.templates.journal.hint",
    needsImage: false,
  },
  {
    id: "ruled",
    labelKey: "notebookCover.templates.ruled.label",
    hintKey: "notebookCover.templates.ruled.hint",
    needsImage: false,
  },
  {
    id: "marble",
    labelKey: "notebookCover.templates.marble.label",
    hintKey: "notebookCover.templates.marble.hint",
    needsImage: false,
  },
  {
    id: "kraft",
    labelKey: "notebookCover.templates.kraft.label",
    hintKey: "notebookCover.templates.kraft.hint",
    needsImage: false,
  },
  {
    id: "spiral",
    labelKey: "notebookCover.templates.spiral.label",
    hintKey: "notebookCover.templates.spiral.hint",
    needsImage: false,
  },
  {
    id: "magazine",
    labelKey: "notebookCover.templates.magazine.label",
    hintKey: "notebookCover.templates.magazine.hint",
    needsImage: false,
  },
  
  // on the shelf, slowly. Book3D plays them only while they are on screen
  // and never under reduced motion.
  {
    id: "aurora",
    labelKey: "notebookCover.templates.aurora.label",
    hintKey: "notebookCover.templates.aurora.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "borealis",
    labelKey: "notebookCover.templates.borealis.label",
    hintKey: "notebookCover.templates.borealis.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "nebula",
    labelKey: "notebookCover.templates.nebula.label",
    hintKey: "notebookCover.templates.nebula.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "tide",
    labelKey: "notebookCover.templates.tide.label",
    hintKey: "notebookCover.templates.tide.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "holo",
    labelKey: "notebookCover.templates.holo.label",
    hintKey: "notebookCover.templates.holo.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "rain",
    labelKey: "notebookCover.templates.rain.label",
    hintKey: "notebookCover.templates.rain.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "fireflies",
    labelKey: "notebookCover.templates.fireflies.label",
    hintKey: "notebookCover.templates.fireflies.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "snowfall",
    labelKey: "notebookCover.templates.snowfall.label",
    hintKey: "notebookCover.templates.snowfall.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "blossom",
    labelKey: "notebookCover.templates.blossom.label",
    hintKey: "notebookCover.templates.blossom.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "inkwash",
    labelKey: "notebookCover.templates.inkwash.label",
    hintKey: "notebookCover.templates.inkwash.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "koi",
    labelKey: "notebookCover.templates.koi.label",
    hintKey: "notebookCover.templates.koi.hint",
    needsImage: false,
    dynamic: true,
  },
  {
    id: "halfImage",
    labelKey: "notebookCover.templates.halfImage.label",
    hintKey: "notebookCover.templates.halfImage.hint",
    needsImage: true,
  },
  {
    id: "fullImage",
    labelKey: "notebookCover.templates.fullImage.label",
    hintKey: "notebookCover.templates.fullImage.hint",
    needsImage: true,
  },
];

const STORAGE_KEY = "mrnote_notebook_cover_v1";
const LEGACY_STORAGE_KEYS = [
  "mrai_notebook_cover_v1",
  "mingrun_notebook_cover_v1",
] as const;

// Default cover = the classic black notebook (NOT swatch [0] — the palette
// leads with red, which is a picker ordering, not a sensible default).
const DEFAULT_SWATCH =
  COLOR_SWATCHES.find((s) => s.id === "black") ?? COLOR_SWATCHES[0];

export const DEFAULT_COVER_PREFERENCE: CoverPreference = {
  template: "solid",
  color: DEFAULT_SWATCH.color,
  spineColor: DEFAULT_SWATCH.spine,
  textTheme: DEFAULT_SWATCH.text,
};

/** The colour a picture cover gives its notebook; none for a cover whose
 *  look is the colour itself (solid, cloth, aurora, …). */
export function templatePalette(template: CoverTemplate): CoverPalette | null {
  return COVER_PALETTES[template] ?? null;
}

 
export function preferenceFollowingTemplate(
  pref: CoverPreference,
  palette: CoverPalette,
): CoverPreference {
  return {
    ...pref,
    color: palette.color,
    spineColor: palette.spine,
    textTheme: palette.text,
    colorPicked: false,
  };
}

/**
 * The preference with `template` chosen: a notebook whose owner has not
 * picked a colour takes the new cover's own; a picked colour stays.
 */
export function preferenceWithTemplate(
  pref: CoverPreference,
  template: CoverTemplate,
): CoverPreference {
  const palette = templatePalette(template);
  if (pref.colorPicked || !palette) return { ...pref, template };
  return preferenceFollowingTemplate({ ...pref, template }, palette);
}

/**
 * The preference as the notebook wears it. Following its cover, it wears
 * the cover's colour as the palette has it now. One saved before covers had
 * colours of their own, with a picture cover and still the default black,
 * never had a colour picked for it, and follows too. A picked colour, or
 * any other colour saved before, is the owner's.
 */
export function effectiveCoverPreference(pref: CoverPreference): CoverPreference {
  if (pref.colorPicked) return pref;
  const palette = templatePalette(pref.template);
  if (!palette) return pref;
  const legacyDefault =
    pref.colorPicked == null &&
    pref.color === DEFAULT_COVER_PREFERENCE.color &&
    pref.spineColor === DEFAULT_COVER_PREFERENCE.spineColor;
  if (pref.colorPicked === false || legacyDefault) {
    return preferenceFollowingTemplate(pref, palette);
  }
  return pref;
}

interface CoverStore {
  [notebookId: string]: CoverPreference;
}

function readStore(): CoverStore {
  if (typeof window === "undefined") return {};
  const candidateKeys = [STORAGE_KEY, ...LEGACY_STORAGE_KEYS];
  for (const key of candidateKeys) {
    try {
      const raw = window.localStorage.getItem(key);
      if (!raw) continue;
      const parsed = JSON.parse(raw) as CoverStore;
      if (key !== STORAGE_KEY) {
        window.localStorage.setItem(STORAGE_KEY, raw);
      }
      LEGACY_STORAGE_KEYS.forEach((legacyKey) => window.localStorage.removeItem(legacyKey));
      return parsed;
    } catch {
      continue;
    }
  }
  return {};
}

function writeStore(store: CoverStore): void {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
  } catch {
    /* quota or disabled — silently ignore */
  }
}

export function getCoverPreference(notebookId: string): CoverPreference | null {
  const store = readStore();
  return store[notebookId] ?? null;
}

export function setCoverPreference(notebookId: string, pref: CoverPreference): void {
  const store = readStore();
  store[notebookId] = pref;
  writeStore(store);
}

export function deleteCoverPreference(notebookId: string): void {
  const store = readStore();
  delete store[notebookId];
  writeStore(store);
}

export function makePreferenceFromSwatch(
  template: CoverTemplate,
  swatch: ColorSwatch,
  imageUrl?: string,
): CoverPreference {
  return {
    template,
    color: swatch.color,
    spineColor: swatch.spine,
    textTheme: swatch.text,
    imageUrl,
  };
}

/**
 * Build a preference from a FREE colour (the colour-wheel path): the spine
 * is the cover darkened the same way the preset swatches are cut, and the
 * text theme flips to dark ink on light covers by relative luminance.
 */
export function makePreferenceFromColor(
  template: CoverTemplate,
  color: string,
  imageUrl?: string,
): CoverPreference {
  const rgb = hexToRgb(color);
  const spine = rgb
    ? rgbToHex(rgb.map((c) => Math.round(c * 0.55)) as [number, number, number])
    : color;
  return {
    template,
    color,
    spineColor: spine,
    textTheme: rgb && relativeLuminance(rgb) > 0.55 ? "dark" : "light",
    imageUrl,
  };
}

function hexToRgb(hex: string): [number, number, number] | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return null;
  const v = parseInt(m[1], 16);
  return [(v >> 16) & 255, (v >> 8) & 255, v & 255];
}

function rgbToHex([r, g, b]: [number, number, number]): string {
  return `#${((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1)}`;
}

function relativeLuminance([r, g, b]: [number, number, number]): number {
  const lin = [r, g, b].map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2];
}

/** Contrast below which a cover's title cannot be read at a glance. */
const COVER_INK_MIN_CONTRAST = 3;
/** Luminance of the dark cover ink (Book3D: rgba(20, 20, 19, .88) on the cover). */
const DARK_INK_LUMINANCE = 0.0075;








export function readableCoverInk(color: string, stored: "light" | "dark"): "light" | "dark" {
  const rgb = hexToRgb(color);
  if (!rgb) return stored;
  const cover = relativeLuminance(rgb);
  const onWhite = 1.05 / (cover + 0.05);
  const onDark = (cover + 0.05) / (DARK_INK_LUMINANCE + 0.05);
  if (stored === "light" && onWhite < COVER_INK_MIN_CONTRAST && onDark > onWhite) return "dark";
  if (stored === "dark" && onDark < COVER_INK_MIN_CONTRAST && onWhite > onDark) return "light";
  return stored;
}

export const COVER_PREF_CHANGED_EVENT = "mrnote-cover-pref-changed";

export function dispatchCoverPrefChanged(notebookId: string): void {
  if (typeof window === "undefined") return;
  window.dispatchEvent(
    new CustomEvent(COVER_PREF_CHANGED_EVENT, { detail: { notebookId } }),
  );
}

/** Whether a template moves (Book3D animates it while it is on screen). */
export function isDynamicCoverTemplate(template: CoverTemplate): boolean {
  const art = coverArt(template);
  if (art) return isMovingCoverArt(art);
  return TEMPLATE_OPTIONS.some((option) => option.id === template && option.dynamic);
}
