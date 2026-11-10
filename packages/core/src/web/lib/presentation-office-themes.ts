/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * PowerPoint's Office themes (Design > Themes > Office) as the deck editor
 * draws and applies them: each theme's id, localized name, heading / body
 * faces, the 4 (Integral: 8) colour variants and the placeholder frames of
 * its layouts, in the gallery's order. The data is
 * `presentation-office-themes.json`, generated from the facts of the
 * installed .thmx packages (MRAI-corpus/ppt-themes/tools/build_catalog.py)
 * in a compact positional encoding this module decodes -- names, colours,
 * font names and frame numbers only. The family and variant GUIDs stay on
 * the server (its catalog writes and recognises them). Never hand-edit the
 * file, and never put package XML, artwork or thumbnails into it: the
 * browser draws every tile itself, and the server writes the theme part
 * from its own catalog.
 *
 * Dependency-free apart from the catalog types so the deck theme model
 * (`lib/presentation-theme`) and the zod schema can import it.
 */

import source from "./presentation-office-themes.json";
import {
  officeThemeLabel,
  type OfficeThemeLabel,
} from "./office-theme-catalog";
import { officeTransformedHex } from "./office-color-transform";

/** Variant ids of an Office theme: `v1` is the base variant (the main package). */
export const PRESENTATION_OFFICE_VARIANT_IDS = [
  "v1",
  "v2",
  "v3",
  "v4",
  "v5",
  "v6",
  "v7",
  "v8",
] as const;

export type PresentationOfficeVariantId =
  (typeof PRESENTATION_OFFICE_VARIANT_IDS)[number];

export const PRESENTATION_OFFICE_COLOR_SLOTS = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export type PresentationOfficeColorSlot =
  (typeof PRESENTATION_OFFICE_COLOR_SLOTS)[number];

export type PresentationOfficeColors = Readonly<
  Record<PresentationOfficeColorSlot, string>
>;

/** One font collection (major = headings, minor = body) of a theme. */
export interface PresentationOfficeThemeFontCollection {
  readonly latin: string;
  /** Per-script faces the browser draws with (`Hans`). */
  readonly scripts: Readonly<Record<string, string>>;
}

export interface PresentationOfficeThemeVariant {
  readonly id: PresentationOfficeVariantId;
  /** MRAI layout id -> frame indices for the layouts this variant places differently. */
  readonly layoutFrames: Readonly<Record<string, readonly number[]>>;
  /** 1-based position in the theme's variant manager (16:9 entries). */
  readonly index: number;
  readonly colors: PresentationOfficeColors;
  /**
   * The slide ground the master paints, as a colour: its `bgRef` colour
   * through the master's own colour map (the dark themes map `bg1`/`bg2` to
   * the dark slots) and colour transforms.
   */
  readonly background: string;
  /** The master's title text colour (`p:titleStyle`), readable on the ground. */
  readonly titleColor: string;
  /** The master's body text colour (`p:bodyStyle`), readable on the ground. */
  readonly bodyColor: string;
}

/**
 * One placeholder of a theme layout, as the catalog records it: the `p:ph`
 * identity (`type` / `idx`), its frame in EMU on the 16:9 catalog slide
 * (12192000 x 6858000) and the level-1 text facts the layout states through
 * its master. `clr` is a scheme slot (after the layout's colour map) or a
 * `#rrggbb`; `mods` its colour transforms in order.
 */
export interface PresentationOfficePlaceholderFrame {
  readonly role: "title" | "subtitle" | "body" | "content" | "picture";
  readonly type?: string;
  readonly idx?: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly algn?: "l" | "ctr" | "r" | "just" | "dist";
  readonly anchor?: "t" | "ctr" | "b";
  /** Font size in hundredths of a point (`a:defRPr@sz`). */
  readonly sz?: number;
  readonly b?: boolean;
  readonly cap?: "all" | "small" | "none";
  readonly clr?: string;
  readonly mods?: ReadonlyArray<readonly [string, number]>;
  readonly vert?: boolean;
}

/** The slide every catalog frame was recorded on (all Office themes are 16:9). */
export const PRESENTATION_OFFICE_CATALOG_SLIDE_WIDTH = 12192000;
export const PRESENTATION_OFFICE_CATALOG_SLIDE_HEIGHT = 6858000;

export interface PresentationOfficeTheme {
  readonly id: string;
  readonly label: OfficeThemeLabel;
  readonly fonts: {
    readonly major: PresentationOfficeThemeFontCollection;
    readonly minor: PresentationOfficeThemeFontCollection;
  };
  readonly variants: readonly PresentationOfficeThemeVariant[];
  /** The theme's placeholder frames; `layoutFrames` index into this table. */
  readonly placeholderFrames: readonly PresentationOfficePlaceholderFrame[];
  /** MRAI layout id -> frame indices, in the editor's pour order (base variant). */
  readonly layoutFrames: Readonly<Record<string, readonly number[]>>;
}

/** `c`: the 12 slot colours, `k`: ground / title / body, as packed hex digits. */
interface CatalogVariantSource {
  c: string;
  k: string;
  /** Position in `layoutIds` -> frame indices, for the layouts that differ from the base variant's. */
  l?: Record<string, number[]>;
}

/**
 * `[type, idx, x, y, cx, cy, sz, algn, anchor, cap, b, vert, clr, mods]`,
 * trailing zeros cut; 0 = not stated (`b`: 1 true, 2 false).
 */
type CatalogFrameSource = ReadonlyArray<unknown>;

interface CatalogThemeSource {
  id: string;
  label: OfficeThemeLabel;
  /** `[major latin, minor latin, major Hans, minor Hans]`. */
  fonts: [string, string, string, string];
  frames: CatalogFrameSource[];
  /** Frame indices of each layout of `layoutIds` (base variant). */
  layouts: number[][];
  variants: CatalogVariantSource[];
}

interface CatalogSource {
  defaultThemeId: string;
  order: string[];
  layoutIds: string[];
  themes: CatalogThemeSource[];
}

const catalog = source as unknown as CatalogSource;

const FRAME_ROLE_BY_TYPE: Readonly<
  Record<string, PresentationOfficePlaceholderFrame["role"]>
> = {
  title: "title",
  ctrTitle: "title",
  subTitle: "subtitle",
  body: "body",
  pic: "picture",
};

function frameIndices(raw: unknown): readonly number[] | null {
  return Array.isArray(raw) &&
    raw.every((index) => Number.isInteger(index) && index >= 0)
    ? Object.freeze([...(raw as number[])])
    : null;
}

function freezeFrame(
  raw: CatalogFrameSource,
): PresentationOfficePlaceholderFrame | null {
  const emu = (value: unknown) =>
    typeof value === "number" && Number.isInteger(value) ? value : null;
  const [rawType, rawIdx, , , , , sz, algn, anchor, cap, bold, vert, clr, mods] = raw;
  const x = emu(raw[2]);
  const y = emu(raw[3]);
  const cx = emu(raw[4]);
  const cy = emu(raw[5]);
  if (x === null || y === null || cx === null || cy === null || cx <= 0 || cy <= 0) {
    return null;
  }
  const type = typeof rawType === "string" ? rawType : undefined;
  const frame: {
    -readonly [K in keyof PresentationOfficePlaceholderFrame]: PresentationOfficePlaceholderFrame[K];
  } = { role: FRAME_ROLE_BY_TYPE[type ?? ""] ?? "content", x, y, cx, cy };
  if (type) frame.type = type;
  if (typeof rawIdx === "string") frame.idx = rawIdx;
  if (["l", "ctr", "r", "just", "dist"].includes(String(algn))) {
    frame.algn = algn as PresentationOfficePlaceholderFrame["algn"];
  }
  if (["t", "ctr", "b"].includes(String(anchor))) {
    frame.anchor = anchor as PresentationOfficePlaceholderFrame["anchor"];
  }
  if (typeof sz === "number" && sz > 0) frame.sz = sz;
  if (bold === 1 || bold === 2) frame.b = bold === 1;
  if (["all", "small", "none"].includes(String(cap))) {
    frame.cap = cap as PresentationOfficePlaceholderFrame["cap"];
  }
  if (typeof clr === "string") frame.clr = clr;
  if (Array.isArray(mods)) {
    frame.mods = Object.freeze(
      mods.flatMap((entry) =>
        Array.isArray(entry) &&
        typeof entry[0] === "string" &&
        typeof entry[1] === "number"
          ? [Object.freeze([entry[0], entry[1]] as const)]
          : [],
      ),
    );
  }
  if (vert === 1) frame.vert = true;
  return Object.freeze(frame);
}

const PACKED_COLORS_RE = /^[0-9a-f]{72}$/;
const PACKED_CANVAS_RE = /^[0-9a-f]{18}$/;

function freezeVariant(
  raw: CatalogVariantSource,
  index: number,
): PresentationOfficeThemeVariant | null {
  if (index < 1 || index > 8) return null;
  if (!PACKED_COLORS_RE.test(raw.c) || !PACKED_CANVAS_RE.test(raw.k)) return null;
  const colors = {} as Record<PresentationOfficeColorSlot, string>;
  PRESENTATION_OFFICE_COLOR_SLOTS.forEach((slot, position) => {
    colors[slot] = `#${raw.c.slice(position * 6, position * 6 + 6)}`;
  });
  const layoutFrames: Record<string, readonly number[]> = {};
  for (const [position, indices] of Object.entries(raw.l ?? {})) {
    const layoutId = catalog.layoutIds[Number(position)];
    const frozen = frameIndices(indices);
    if (layoutId && frozen) layoutFrames[layoutId] = frozen;
  }
  return Object.freeze({
    id: `v${index}` as PresentationOfficeVariantId,
    index,
    layoutFrames: Object.freeze(layoutFrames),
    colors: Object.freeze(colors),
    background: `#${raw.k.slice(0, 6)}`,
    titleColor: `#${raw.k.slice(6, 12)}`,
    bodyColor: `#${raw.k.slice(12, 18)}`,
  });
}

function freezeCollection(
  latin: string,
  hans: string,
): PresentationOfficeThemeFontCollection {
  const scripts: Record<string, string> = {};
  if (hans) scripts.Hans = hans;
  return Object.freeze({ latin, scripts: Object.freeze(scripts) });
}

function freezeTheme(raw: CatalogThemeSource): PresentationOfficeTheme | null {
  const variants = raw.variants
    .map((variant, position) => freezeVariant(variant, position + 1))
    .filter((variant): variant is PresentationOfficeThemeVariant => !!variant);
  // Variants are positional (v1 first): one unreadable record would shift the rest.
  if (!variants.length || variants.length !== raw.variants.length) return null;
  const layoutFrames: Record<string, readonly number[]> = {};
  raw.layouts.forEach((indices, position) => {
    const layoutId = catalog.layoutIds[position];
    const frozen = frameIndices(indices);
    if (layoutId && frozen) layoutFrames[layoutId] = frozen;
  });
  return Object.freeze({
    id: raw.id,
    label: raw.label,
    fonts: Object.freeze({
      major: freezeCollection(raw.fonts[0], raw.fonts[2]),
      minor: freezeCollection(raw.fonts[1], raw.fonts[3]),
    }),
    variants: Object.freeze(variants),
    placeholderFrames: Object.freeze(
      raw.frames.map(
        (frame) =>
          freezeFrame(frame) ??
          // A malformed record keeps its index (so the maps stay valid) as a
          // frame the layout cannot use.
          Object.freeze({ role: "content" as const, x: 0, y: 0, cx: 0, cy: 0 }),
      ),
    ),
    layoutFrames: Object.freeze(layoutFrames),
  });
}

const THEME_BY_ID = new Map<string, PresentationOfficeTheme>();
for (const raw of catalog.themes) {
  const theme = freezeTheme(raw);
  if (theme) THEME_BY_ID.set(theme.id, theme);
}

/** The Office themes in the gallery's order (the catalog's `order`). */
export const PRESENTATION_OFFICE_THEMES: readonly PresentationOfficeTheme[] =
  Object.freeze(
    catalog.order.flatMap((id) => {
      const theme = THEME_BY_ID.get(id);
      return theme ? [theme] : [];
    }),
  );

/** The ids of {@link PRESENTATION_OFFICE_THEMES}, in gallery order. */
export const PRESENTATION_OFFICE_THEME_IDS: readonly string[] = Object.freeze(
  PRESENTATION_OFFICE_THEMES.map((theme) => theme.id),
);

/**
 * The width of each Office theme's second line style (`a:lnStyleLst`, EMU):
 * the outline a NEW shape takes, since PowerPoint's own shape style is
 * `a:lnRef idx="2"` (the same in every variant of a theme). Read from the
 * server's theme catalog; `presentation-theme.test.ts` holds the two equal.
 */
export const PRESENTATION_OFFICE_THEME_SHAPE_LINE_EMU: Readonly<
  Record<string, number>
> = Object.freeze({
  office: 19050,
  "office-2013": 12700,
  atlas: 15875,
  badge: 12700,
  berlin: 12700,
  celestial: 19050,
  circuit: 15875,
  crop: 34925,
  dividend: 22225,
  droplet: 15875,
  facet: 19050,
  frame: 10795,
  gallery: 15875,
  integral: 15875,
  ion: 19050,
  "ion-boardroom": 19050,
  madison: 15875,
  "main-event": 19050,
  mesh: 19050,
  organic: 15875,
  parcel: 12700,
  quotable: 15875,
  savon: 12700,
  "vapor-trail": 12700,
  wisp: 15875,
  "wood-type": 12700,
});

/** The catalog's default theme for a new deck ("office"). */
export const PRESENTATION_OFFICE_THEME_DEFAULT_ID: string =
  catalog.defaultThemeId;

const VARIANT_ID_SET: ReadonlySet<string> = new Set(
  PRESENTATION_OFFICE_VARIANT_IDS,
);

export function isOfficeThemeVariantId(
  value: unknown,
): value is PresentationOfficeVariantId {
  return typeof value === "string" && VARIANT_ID_SET.has(value);
}

/** One Office theme by id, or undefined for anything the catalog lacks. */
export function presentationOfficeTheme(
  id: unknown,
): PresentationOfficeTheme | undefined {
  return typeof id === "string" && PRESENTATION_OFFICE_THEME_IDS.includes(id)
    ? THEME_BY_ID.get(id)
    : undefined;
}

/** One variant (`v1`..`v8`) of an Office theme, or undefined. */
export function presentationOfficeThemeVariant(
  id: unknown,
  variantId: unknown,
): PresentationOfficeThemeVariant | undefined {
  if (!isOfficeThemeVariantId(variantId)) return undefined;
  return presentationOfficeTheme(id)?.variants.find(
    (variant) => variant.id === variantId,
  );
}

/** Office's own name for the theme in `locale`, else English. */
export function presentationOfficeThemeLabel(
  theme: PresentationOfficeTheme,
  locale: string | undefined,
): string {
  return officeThemeLabel(theme.label, locale);
}

/** The face a Chinese (Simplified) run of the collection resolves to. */
export function presentationOfficeThemeHansFace(
  collection: PresentationOfficeThemeFontCollection,
): string | undefined {
  const face = collection.scripts.Hans?.trim();
  return face ? face : undefined;
}

/**
 * The placeholder frames of one MRAI layout under an Office theme variant, in
 * the editor's pour order (LAYOUT_DEFINITIONS order), or undefined when the
 * theme, variant or layout is unknown. `blank` has none.
 */
export function presentationOfficeLayoutFrames(
  themeId: unknown,
  variantId: unknown,
  layoutId: unknown,
): readonly PresentationOfficePlaceholderFrame[] | undefined {
  const theme = presentationOfficeTheme(themeId);
  const variant = presentationOfficeThemeVariant(themeId, variantId);
  if (!theme || !variant || typeof layoutId !== "string") return undefined;
  const indices = variant.layoutFrames[layoutId] ?? theme.layoutFrames[layoutId];
  if (!indices) return undefined;
  const frames = indices.map((index) => theme.placeholderFrames[index]);
  return frames.every((frame) => frame && frame.cx > 0 && frame.cy > 0)
    ? frames
    : undefined;
}

/**
 * The `#rrggbb` a frame's text colour resolves to under a variant: its slot
 * through the variant's colours (or its own srgb) and its transforms; null
 * when the frame states no colour or a transform is not modelled.
 */
export function presentationOfficeFrameTextColor(
  frame: PresentationOfficePlaceholderFrame,
  variant: PresentationOfficeThemeVariant,
): string | null {
  if (!frame.clr) return null;
  const base = frame.clr.startsWith("#")
    ? frame.clr
    : (variant.colors as Record<string, string>)[frame.clr];
  if (!base) return null;
  return officeTransformedHex(base, frame.mods ?? []);
}
