/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The decoration of PowerPoint's Office themes as the deck editor draws it:
 * for each theme, variant and slide layout, the decorative shapes of the
 * master and the layout as SVG paths, coloured by references to the theme's
 * colour slots. The paths are in half units of a 1600 x 900 slide (the view
 * box is 3200 x 1800), so every coordinate is an integer.
 *
 * The geometry is MRAI's own drawing (drawn by eye from how PowerPoint
 * renders each theme; MRAI-corpus/ppt-themes/artwork), flattened by
 * MRAI-corpus/ppt-themes/tools/build_web_artwork.py. The data is
 * `presentation-office-theme-artwork.json` in a compact encoding (shared
 * colour, fill, path and shape tables; each path is relative to the start
 * point its shape supplies, so a translated copy is stored once) which this
 * module decodes on first use: never hand-edit it, and never put package XML
 * or Microsoft artwork into it. The server writes the same
 * drawings into an exported package's master and layouts.
 */

import source from "./presentation-office-theme-artwork.json";
import { officeTransformedHex } from "@mrnote/core/web/lib/office-color-transform.ts";

/** `[colour slot or #rrggbb, colour transforms, opacity 0..1]`. */
export type PresentationOfficeArtworkColor = readonly [
  string,
  ReadonlyArray<readonly [string, number | null]>,
  number,
];

export interface PresentationOfficeArtworkFill {
  /** Solid colour. */
  readonly c?: PresentationOfficeArtworkColor;
  /** Gradient stops `[position 0..100, colour]`. */
  readonly g?: ReadonlyArray<readonly [number, PresentationOfficeArtworkColor]>;
  /** Linear gradient angle in degrees, clockwise from left-to-right. */
  readonly ang?: number;
  /** Radial gradient centre in percent of the shape box. */
  readonly radial?: readonly [number, number];
  /** Preset pattern (`pct30`, `dotGrid`, ...) with its foreground / background. */
  readonly pat?: string;
  readonly fg?: PresentationOfficeArtworkColor;
  readonly bg?: PresentationOfficeArtworkColor;
}

/**
 * A preset pattern as an 8 x 8 tile: the dots of a dot screen (`pctNN`
 * covers about NN % of the tile) or the dots of a dot grid. Other presets
 * fall back to a 50 % screen.
 */
export function officeArtworkPatternDots(
  preset: string,
): ReadonlyArray<readonly [number, number, number]> {
  if (preset === "dotGrid") return [[0, 0, 0.7]];
  const percent = /^pct(\d+)$/.exec(preset);
  const coverage = percent ? Number(percent[1]) / 100 : 0.5;
  // Two dots per tile on a staggered lattice; the radius sets the coverage.
  const radius = Math.sqrt((coverage * 32) / Math.PI);
  return [
    [2, 2, radius],
    [6, 6, radius],
  ];
}

export interface PresentationOfficeArtworkShape {
  /** SVG path data in the artwork view box (absolute start, relative segments). */
  readonly d: string;
  readonly f?: PresentationOfficeArtworkFill;
  /** Stroke colour. */
  readonly s?: PresentationOfficeArtworkColor;
  /** Stroke width in view-box units. */
  readonly w?: number;
  /** 1 = the slide ground itself (a slide's own background replaces it). */
  readonly bg?: 1;
}

/** A fill of the file: colours are indices into the colour table. */
interface ArtworkFillSource {
  readonly c?: number;
  readonly g?: ReadonlyArray<readonly [number, number]>;
  readonly ang?: number;
  readonly radial?: readonly [number, number];
  readonly pat?: string;
  readonly fg?: number;
  readonly bg?: number;
}

/**
 * `[path, x, y, fill, stroke colour, stroke width, ground]`: indices into
 * the tables (-1 = none), the path's start point; trailing defaults are cut.
 */
type ArtworkShapeSource = readonly number[];

interface ArtworkThemeSource {
  /** Shape indices of each distinct drawing of the theme. */
  readonly sets: ReadonlyArray<readonly number[]>;
  /** Variant id -> the set each layout of `layoutIds` draws (-1 = none). */
  readonly variants: Readonly<Record<string, readonly number[]>>;
}

interface ArtworkSource {
  readonly viewBox: readonly [number, number];
  readonly layoutIds: readonly string[];
  readonly colors: readonly PresentationOfficeArtworkColor[];
  readonly fills: readonly ArtworkFillSource[];
  readonly paths: readonly string[];
  readonly shapes: readonly ArtworkShapeSource[];
  readonly themes: Readonly<Record<string, ArtworkThemeSource>>;
}

const artwork = source as unknown as ArtworkSource;

export const PRESENTATION_OFFICE_ARTWORK_VIEW_BOX: readonly [number, number] = [
  artwork.viewBox[0],
  artwork.viewBox[1],
];

/** View-box units per unit of the 1600-wide slide the pattern tiles are drawn on. */
export const PRESENTATION_OFFICE_ARTWORK_UNIT = artwork.viewBox[0] / 1600;

const NO_SHAPES: ReadonlyArray<PresentationOfficeArtworkShape> = Object.freeze([]);

const LAYOUT_POSITION: ReadonlyMap<string, number> = new Map(
  artwork.layoutIds.map((layoutId, position) => [layoutId, position]),
);
const FALLBACK_LAYOUT_POSITION = LAYOUT_POSITION.get("title-content") ?? 0;

const decodedShapes: Array<PresentationOfficeArtworkShape | undefined> = [];
const decodedSets = new Map<string, ReadonlyArray<PresentationOfficeArtworkShape>>();

function decodeFill(index: number): PresentationOfficeArtworkFill | undefined {
  const raw = artwork.fills[index];
  if (!raw) return undefined;
  const color = (value: number | undefined) =>
    value === undefined ? undefined : artwork.colors[value];
  const fill: { -readonly [K in keyof PresentationOfficeArtworkFill]: PresentationOfficeArtworkFill[K] } = {};
  const solid = color(raw.c);
  if (solid) fill.c = solid;
  if (raw.g) {
    fill.g = raw.g.flatMap(([position, value]) => {
      const stop = color(value);
      return stop ? [[position, stop] as const] : [];
    });
    if (raw.ang !== undefined) fill.ang = raw.ang;
    if (raw.radial) fill.radial = raw.radial;
  }
  if (raw.pat) {
    fill.pat = raw.pat;
    const foreground = color(raw.fg);
    const background = color(raw.bg);
    if (foreground) fill.fg = foreground;
    if (background) fill.bg = background;
  }
  return Object.freeze(fill);
}

function decodeShape(index: number): PresentationOfficeArtworkShape | undefined {
  const known = decodedShapes[index];
  if (known) return known;
  const raw = artwork.shapes[index];
  const body = raw ? artwork.paths[raw[0]] : undefined;
  if (!raw || body === undefined) return undefined;
  const shape: { -readonly [K in keyof PresentationOfficeArtworkShape]: PresentationOfficeArtworkShape[K] } = {
    d: `M${raw[1]} ${raw[2]}${body}`,
  };
  const fill = raw[3] !== undefined && raw[3] >= 0 ? decodeFill(raw[3]) : undefined;
  if (fill) shape.f = fill;
  const stroke = raw[4] !== undefined && raw[4] >= 0 ? artwork.colors[raw[4]] : undefined;
  if (stroke) {
    shape.s = stroke;
    if (raw[5]) shape.w = raw[5];
  }
  if (raw[6] === 1) shape.bg = 1;
  const frozen = Object.freeze(shape);
  decodedShapes[index] = frozen;
  return frozen;
}

/**
 * The decoration of one Office theme variant on a slide of `layoutId` (an
 * MRAI layout id such as `title` or `title-content`). A layout the theme does
 * not decorate separately shows what a title-and-content slide shows; a
 * theme or variant without artwork gives an empty list.
 */
export function presentationOfficeThemeArtwork(
  themeId: unknown,
  variantId: unknown,
  layoutId: unknown,
): ReadonlyArray<PresentationOfficeArtworkShape> {
  if (typeof themeId !== "string") return NO_SHAPES;
  const theme = Object.prototype.hasOwnProperty.call(artwork.themes, themeId)
    ? artwork.themes[themeId]
    : undefined;
  if (!theme) return NO_SHAPES;
  const variant =
    (typeof variantId === "string" &&
    Object.prototype.hasOwnProperty.call(theme.variants, variantId)
      ? theme.variants[variantId]
      : undefined) ?? theme.variants.v1;
  if (!variant) return NO_SHAPES;
  const position =
    (typeof layoutId === "string" ? LAYOUT_POSITION.get(layoutId) : undefined) ??
    FALLBACK_LAYOUT_POSITION;
  const setIndex = variant[position];
  const set = setIndex === undefined ? undefined : theme.sets[setIndex];
  if (!set) return NO_SHAPES;
  const key = `${themeId}/${setIndex}`;
  const known = decodedSets.get(key);
  if (known) return known;
  const shapes = Object.freeze(
    set.flatMap((shapeIndex) => {
      const shape = decodeShape(shapeIndex);
      return shape ? [shape] : [];
    }),
  );
  decodedSets.set(key, shapes);
  return shapes;
}

/** The SVG path data of one shape. */
export function presentationOfficeArtworkPath(
  shape: PresentationOfficeArtworkShape,
): string {
  return shape.d;
}

/** A drawing colour against the deck's twelve colour slots. */
export function resolveOfficeArtworkColor(
  color: PresentationOfficeArtworkColor,
  slots: Readonly<Record<string, string>> | undefined,
): { color: string; opacity: number } {
  const [reference, transforms, opacity] = color;
  const base = reference.startsWith("#")
    ? reference
    : (slots?.[reference] ?? "#808080");
  const moved = transforms.length ? officeTransformedHex(base, transforms) : base;
  return {
    color: moved ?? base,
    opacity: Number.isFinite(opacity) ? Math.max(0, Math.min(1, opacity)) : 1,
  };
}
