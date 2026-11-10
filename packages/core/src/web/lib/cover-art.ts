/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { COVER_ART } from "./cover-art-catalog";













export type CoverArtId = `art-${string}`;

export type CoverArtLayer =
  /** the upper band drifts sideways (clouds) */
  | "sky"
  /** the water ripples, its reflections breaking and mending */
  | "water"
  /** surf runs up the shore and draws back */
  | "wash"
  /** glints run across the water */
  | "shimmer"
  /** the painting's own points of light twinkle */
  | "twinkle"
  /** the light source breathes (moon, sun, lamp, window) */
  | "glow"
  /** the light source flickers like a flame */
  | "flicker"
  /** lightning flashes now and then */
  | "flash"
  /** the light opens out from its heart and fades (fireworks, ripples) */
  | "burst"
  /** the plants bend in the wind */
  | "sway"
  /** a wave of wind runs over the plants */
  | "gust"
  /** banks of fog, haze or smoke drift through a band */
  | "mist"
  /** the upper band sways like a curtain of light */
  | "aurora"
  /** shafts of light (through water, trees, glass, an aurora) shift */
  | "rays"
  /** white water runs down a waterfall */
  | "fall"
  /** the air over the ground wavers in heat */
  | "heat"
  /** a broad band of light crosses the painting (silk, metal, glass) */
  | "sheen"
  /** the whole (square) painting turns about its centre */
  | "turn";

export type CoverArtParticles =
  | "rain"
  | "snow"
  | "fireflies"
  | "petals"
  | "leaves"
  | "embers"
  | "bubbles"
  | "motes"
  | "meteors"
  | "lanterns";

export type CoverArtGroup =
  | "landscape"
  | "east"
  | "botanical"
  | "pattern"
  | "abstract"
  | "material"
  | "city"
  | "animal"
  | "sky"
  | "night"
  | "water"
  | "weather"
  | "forest"
  | "light"
  | "cosmos"
  | "season"
  | "flow";

/**
 * Where a painting's moving regions lie, in % of the cover's height from
 * its top unless said otherwise. A painting that leaves one out takes the
 * recipe's default (styles/cover-art.css). They keep each moving part to
 * its own region, with the soft edges where a shifted copy has nothing to
 * double.
 */
export interface CoverArtPlaces {
  /** where the drifting sky, or an aurora's curtain, ends */
  sky?: number;
  /** where the water begins */
  water?: number;
  /** where the water ends (a far shore, a quay) */
  shore?: number;
  /** where the swaying plants begin */
  sway?: number;
  /** where they end */
  swayEnd?: number;
  /** what they sway from: their roots below, or a bough above */
  root?: number;
  /** where a gust starts running, if not where the swaying starts */
  gust?: number;
  /** the middle of the mist */
  mist?: number;
  /** the middle of the wavering air */
  heat?: number;
  /** how far down the stars and lamps twinkle */
  stars?: number;
  /** the tone (0–1) above which a point twinkles or glints */
  spark?: number;
  /** how far a light reaches, in % of the cover's width */
  glow?: number;
  /** how far down light rays reach */
  rays?: number;
  /** how far they lean, in degrees (negative: down to the left) */
  tilt?: number;
  /** how far a sheen turns the colours, in degrees (foil) */
  hue?: number;
  /** a waterfall's box: left, top, right, bottom */
  fall?: readonly [number, number, number, number];
  /** where falling rain or snow stops (a window sill) */
  weather?: number;
}

export interface CoverArt {
  id: CoverArtId;
  group: CoverArtGroup;
  /** The title's ink: light over a dark painting, dark over a light one. */
  ink: "light" | "dark";
  /** A moving cover's moving parts. */
  layers?: readonly CoverArtLayer[];
  particles?: CoverArtParticles;
  /** Where the painting's light source is, in % of the cover (x, y). */
  light?: readonly [number, number];
  /** Where its moving regions lie, where they differ from the defaults. */
  places?: CoverArtPlaces;
}

/** The groups in picker order. */
export const STILL_ART_GROUPS: readonly CoverArtGroup[] = [
  "landscape",
  "east",
  "botanical",
  "pattern",
  "abstract",
  "material",
  "city",
  "animal",
  "sky",
];

export const MOVING_ART_GROUPS: readonly CoverArtGroup[] = [
  "night",
  "water",
  "weather",
  "forest",
  "light",
  "cosmos",
  "season",
  "flow",
  "city",
];

const BY_ID = new Map<string, CoverArt>(COVER_ART.map((art) => [art.id, art]));

export function isCoverArtId(template: string): template is CoverArtId {
  return template.startsWith("art-");
}

export function coverArt(template: string): CoverArt | null {
  return BY_ID.get(template) ?? null;
}

export function isMovingCoverArt(art: CoverArt): boolean {
  return Boolean(art.layers?.length || art.particles);
}

export function coverArtSlug(art: CoverArt): string {
  return art.id.slice("art-".length);
}

export function coverArtImage(art: CoverArt): string {
  return `/covers/art/${coverArtSlug(art)}.webp`;
}

/** A moving cover's own light, baked from its painting with the dark ground
 *  transparent (scripts/bake-cover-lights.py): what twinkles, and what
 *  glints on its water. */
export function coverArtLightImage(art: CoverArt, light: "lights" | "glints"): string {
  return `/covers/art/${coverArtSlug(art)}.${light}.webp`;
}

/** The painted covers of one kind, grouped in picker order. */
export function coverArtByGroup(
  moving: boolean,
): { group: CoverArtGroup; covers: CoverArt[] }[] {
  const groups = moving ? MOVING_ART_GROUPS : STILL_ART_GROUPS;
  return groups
    .map((group) => ({
      group,
      covers: COVER_ART.filter(
        (art) => art.group === group && isMovingCoverArt(art) === moving,
      ),
    }))
    .filter((entry) => entry.covers.length > 0);
}

export { COVER_ART };
