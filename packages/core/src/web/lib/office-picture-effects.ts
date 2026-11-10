/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Picture Format > Picture Effects: the page model, PowerPoint's galleries
 * and the canvas rendering.
 *
 * PowerPoint's Picture Effects (the same cascade as Shape Effects) offers
 * Preset, Shadow, Reflection, Glow, Soft Edges, Bevel and 3-D Rotation, each
 * a submenu of "No <effect>" and a gallery of variations; several apply at
 * once and a pick replaces only its own effect (Microsoft Support, "Add or
 * change an effect for a picture"; "Format objects or text with glow or soft
 * edges"). The file side -- reader, from-scratch writer and the in-place patch
 * -- is the server's `app/services/office_picture_effects.py`; a block keeps
 * `pictureEffects`, one key per effect, holding the effect's DrawingML
 * parameters (EMU, 1/1000 %), or `"none"`: "No <effect>" picked on a picture
 * whose Picture Style draws that effect. `shadow` is the shared shadow model
 * (`lib/office-shadow-effects`: Outer, Inner and Perspective); a picture's
 * older `shadow` switch (the exporter's own soft shadow) still draws and
 * saves as before, and a Shadow pick replaces it.
 *
 * Every gallery value below is PowerPoint's own, measured: PowerPoint for Mac
 * 16.86 opened a deck carrying candidate parameters and classified each one
 * as its preset through its scripting model (reflection type 1-9, soft edge
 * type 1-6, glow radius and theme colour, bevel type, camera preset and
 * rotation), `MRAI-corpus/ppt-audit/L11/oracle` (`probe_out.txt`); the glow
 * colour transform (`satMod 175000`, `alpha 40000`) is what every
 * PowerPoint-written gallery glow in the 775-deck corpus carries
 * (`effect_census.txt`). Gallery membership and names are PowerPoint's
 * string tables.
 */

// Byte-identical to the server's `office_camera_rotations.json` (the PDF
// engines turn text and pictures by the same table).
import cameraTable from "./office-camera-rotations.shared.json";
import type { CSSProperties } from "react";

import { officeTransformedHex } from "./office-color-transform";
import {
  isOfficePictureEffectPreset,
  officePictureEffectPresetCss,
  officePictureEffectPresetShadow,
  type OfficePictureEffectPreset,
} from "./office-picture-effect-presets";
import {
  normalizeOfficeShadowEffect,
  officeShadowCss,
  sameOfficeShadowEffect,
  type OfficeShadowEffect,
} from "./office-shadow-effects";

export const OFFICE_PICTURE_EFFECT_KINDS = [
  // Preset 1-12 first: it replaces the effect list, scene and 3-D shape, and
  // the effects picked over it are drawn and written after it.
  "preset",
  "shadow",
  "reflection",
  "glow",
  "softEdge",
  "bevel",
  "rotation3d",
  // Format Shape > 3-D Format beyond the top bevel (`a:sp3d` but `a:bevelT`,
  
  "format3d",
] as const;
export type OfficePictureEffectKind = (typeof OFFICE_PICTURE_EFFECT_KINDS)[number];

export const OFFICE_GLOW_THEME_SLOTS = [
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "bg1",
  "tx1",
  "bg2",
  "tx2",
] as const;
export type OfficeGlowThemeSlot = (typeof OFFICE_GLOW_THEME_SLOTS)[number];

export interface OfficeGlowEffect {
  /** EMU. */
  readonly radius: number;
  /** The `#rrggbb` drawn (theme colour and transforms applied, alpha not). */
  readonly color: string;
  /** Opacity, DrawingML `a:alpha` (100000 = opaque). */
  readonly alpha: number;
  readonly themeColor?: OfficeGlowThemeSlot;
  readonly satMod?: number;
}
export interface OfficeSoftEdgeEffect {
  readonly radius: number;
}
export interface OfficeReflectionEffect {
  readonly blurRad: number;
  readonly stA: number;
  readonly endA: number;
  readonly endPos: number;
  readonly dist: number;
}
export const OFFICE_BEVEL_PRESETS = [
  "circle",
  "relaxedInset",
  "cross",
  "coolSlant",
  "angle",
  "softRound",
  "convex",
  "slope",
  "divot",
  "riblet",
  "hardEdge",
  "artDeco",
] as const;
export type OfficeBevelPreset = (typeof OFFICE_BEVEL_PRESETS)[number];
export interface OfficeBevelEffect {
  readonly preset: OfficeBevelPreset;
  readonly width: number;
  readonly height: number;
}
export interface OfficeRotation3dEffect {
  /** `ST_PresetCameraType`. */
  readonly camera: string;
  readonly lat?: number;
  readonly lon?: number;
  readonly rev?: number;
  /** 3-D Rotation > Perspective: the field of view the file states
   * (`a:camera@fov`, 60000ths of a degree, 0..180°); absent, the preset's own. */
  readonly fov?: number;
}

/** ST_FOVAngle's top: 180 degrees, in 60000ths. */
export const OFFICE_CAMERA_MAX_FOV = 10_800_000;

/**
 * The field of view a camera draws with, in degrees: its stated one, else
 * the preset's (`OFFICE_CAMERA_ROTATIONS`). Only a perspective camera draws
 * one; PowerPoint offers Perspective on those alone.
 */
export function officeCameraFieldOfView(rotation: OfficeRotation3dEffect): number {
  if (typeof rotation.fov === "number") return rotation.fov / 60_000;
  return OFFICE_CAMERA_ROTATIONS[rotation.camera]?.[3] ?? 45;
}

/** `ST_PresetMaterialType`, PowerPoint's Material gallery and the legacy four. */
export const OFFICE_MATERIAL_PRESETS = [
  "matte",
  "warmMatte",
  "plastic",
  "metal",
  "dkEdge",
  "softEdge",
  "flat",
  "legacyWireframe",
  "powder",
  "clear",
  "softmetal",
  "translucentPowder",
  "legacyMatte",
  "legacyPlastic",
  "legacyMetal",
] as const;
export type OfficeMaterialPreset = (typeof OFFICE_MATERIAL_PRESETS)[number];

 
export const OFFICE_LIGHT_RIGS = [
  "threePt",
  "balanced",
  "soft",
  "harsh",
  "flood",
  "contrasting",
  "morning",
  "sunrise",
  "sunset",
  "chilly",
  "freezing",
  "flat",
  "twoPt",
  "glow",
  "brightRoom",
  "legacyFlat1",
  "legacyFlat2",
  "legacyFlat3",
  "legacyFlat4",
  "legacyNormal1",
  "legacyNormal2",
  "legacyNormal3",
  "legacyNormal4",
  "legacyHarsh1",
  "legacyHarsh2",
  "legacyHarsh3",
  "legacyHarsh4",
] as const;
export type OfficeLightRig = (typeof OFFICE_LIGHT_RIGS)[number];

/** `ST_LightRigDirection`. */
export const OFFICE_LIGHT_DIRECTIONS = ["tl", "t", "tr", "l", "r", "bl", "b", "br"] as const;
export type OfficeLightDirection = (typeof OFFICE_LIGHT_DIRECTIONS)[number];










export interface OfficeFormat3dEffect {
  readonly bottom?: OfficeBevelEffect;
  /** EMU, 1..1584 pt. */
  readonly depth?: { readonly height: number; readonly color?: string };
  /** EMU, 1..1584 pt. */
  readonly contour?: { readonly width: number; readonly color?: string };
  readonly material?: OfficeMaterialPreset;
  /** `angle`: 60000ths of a degree. */
  readonly lighting?: { readonly rig: OfficeLightRig; readonly direction?: OfficeLightDirection; readonly angle?: number };
}

export interface OfficePictureEffectValues {
  /** PowerPoint's Preset 1-12 (`lib/office-picture-effect-presets`). */
  readonly preset: OfficePictureEffectPreset;
  readonly shadow: OfficeShadowEffect;
  readonly reflection: OfficeReflectionEffect;
  readonly glow: OfficeGlowEffect;
  readonly softEdge: OfficeSoftEdgeEffect;
  readonly bevel: OfficeBevelEffect;
  readonly rotation3d: OfficeRotation3dEffect;
  readonly format3d: OfficeFormat3dEffect;
}
export type OfficePictureEffects = {
  readonly [K in OfficePictureEffectKind]?: OfficePictureEffectValues[K] | "none";
};

const EMU_PER_PT = 12_700;
const EMU_PER_PX = 9_525;
const MAX = {
  glow: 150 * EMU_PER_PT,
  softEdge: 100 * EMU_PER_PT,
  reflection: 100 * EMU_PER_PT,
  bevel: 1584 * EMU_PER_PT,
};
const DEFAULT_BEVEL = 76_200;

/**
 * `ST_PresetCameraType` without `orthographicFront` (No Rotation), with
 * the rotation PowerPoint reports for each (X, Y, Z degrees and field of
 * view; PowerPoint for Mac, `probe_out.txt`). Legacy cameras are accepted
 * from a file and drawn unrotated.
 */
export const OFFICE_CAMERA_ROTATIONS: Readonly<
  Record<string, readonly [x: number, y: number, z: number, fov: number]>
> = cameraTable.cameras as unknown as Record<string, readonly [number, number, number, number]>;
const LEGACY_CAMERA = /^legacy(?:Oblique|Perspective)(?:TopLeft|Top|TopRight|Left|Front|Right|BottomLeft|Bottom|BottomRight)$/;

export function isOfficeCameraPreset(value: unknown): value is string {
  return (
    typeof value === "string" &&
    (Object.prototype.hasOwnProperty.call(OFFICE_CAMERA_ROTATIONS, value) || LEGACY_CAMERA.test(value))
  );
}

// --- galleries ---------------------------------------------------------------

/** Reflection Variations, PowerPoint's order (reflection type 1-9). */
export const OFFICE_REFLECTION_VARIATIONS: readonly {
  readonly id: string;
  readonly value: OfficeReflectionEffect;
}[] = [
  { id: "tightTouching", value: { blurRad: 6350, stA: 52000, endA: 300, endPos: 35000, dist: 0 } },
  { id: "halfTouching", value: { blurRad: 6350, stA: 50000, endA: 300, endPos: 55000, dist: 0 } },
  { id: "fullTouching", value: { blurRad: 6350, stA: 50000, endA: 300, endPos: 90000, dist: 0 } },
  { id: "tight4pt", value: { blurRad: 6350, stA: 50000, endA: 300, endPos: 38500, dist: 50800 } },
  { id: "half4pt", value: { blurRad: 6350, stA: 50000, endA: 300, endPos: 55500, dist: 50800 } },
  { id: "full4pt", value: { blurRad: 6350, stA: 50000, endA: 300, endPos: 90000, dist: 50800 } },
  { id: "tight8pt", value: { blurRad: 6350, stA: 50000, endA: 275, endPos: 40000, dist: 101600 } },
  { id: "half8pt", value: { blurRad: 6350, stA: 50000, endA: 300, endPos: 55500, dist: 101600 } },
  { id: "full8pt", value: { blurRad: 6350, stA: 50000, endA: 295, endPos: 92000, dist: 101600 } },
];

/** Glow Variations: four sizes (points) by Accent 1-6. */
export const OFFICE_GLOW_SIZES_PT = [5, 8, 11, 18] as const;
export const OFFICE_GLOW_ACCENTS = [
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const;
const GALLERY_GLOW_ALPHA = 40_000;
const GALLERY_GLOW_SAT_MOD = 175_000;

/** One Glow Variations swatch over the deck's accent colour. */
export function officeGlowVariation(
  sizePt: (typeof OFFICE_GLOW_SIZES_PT)[number],
  accent: (typeof OFFICE_GLOW_ACCENTS)[number],
  accentHex: string,
): OfficeGlowEffect {
  return {
    radius: Math.round(sizePt * EMU_PER_PT),
    color: officeTransformedHex(accentHex, [["satMod", GALLERY_GLOW_SAT_MOD]]) ?? accentHex.toLowerCase(),
    alpha: GALLERY_GLOW_ALPHA,
    themeColor: accent,
    satMod: GALLERY_GLOW_SAT_MOD,
  };
}

/** Soft Edge Variations (points), soft edge type 1-6. */
export const OFFICE_SOFT_EDGE_SIZES_PT = [1, 2.5, 5, 10, 25, 50] as const;

export function officeSoftEdgeVariation(sizePt: number): OfficeSoftEdgeEffect {
  return { radius: Math.round(sizePt * EMU_PER_PT) };
}

/** A Bevel gallery pick: the preset at Office's default 6 pt by 6 pt. */
export function officeBevelVariation(preset: OfficeBevelPreset): OfficeBevelEffect {
  return { preset, width: DEFAULT_BEVEL, height: DEFAULT_BEVEL };
}

/**
 * The 3-D Rotation gallery, PowerPoint's sections. Membership is the camera
 * names PowerPoint's own string table carries for the gallery (`Isometric
 * Left Down` ... `Oblique Bottom Right`); `isometricLeftUp`, `obliqueTop` and
 * the like exist in DrawingML but not in the gallery.
 */
export const OFFICE_ROTATION_GALLERY: readonly {
  readonly section: "parallel" | "perspective" | "oblique";
  readonly cameras: readonly string[];
}[] = [
  {
    section: "parallel",
    cameras: [
      "isometricLeftDown",
      "isometricRightUp",
      "isometricTopUp",
      "isometricBottomDown",
      "isometricOffAxis1Left",
      "isometricOffAxis1Right",
      "isometricOffAxis1Top",
      "isometricOffAxis2Left",
      "isometricOffAxis2Right",
      "isometricOffAxis2Top",
    ],
  },
  {
    section: "perspective",
    cameras: [
      "perspectiveFront",
      "perspectiveLeft",
      "perspectiveRight",
      "perspectiveBelow",
      "perspectiveAbove",
      "perspectiveRelaxedModerately",
      "perspectiveRelaxed",
      "perspectiveContrastingLeftFacing",
      "perspectiveContrastingRightFacing",
      "perspectiveHeroicExtremeLeftFacing",
      "perspectiveHeroicExtremeRightFacing",
    ],
  },
  {
    section: "oblique",
    cameras: ["obliqueTopLeft", "obliqueTopRight", "obliqueBottomLeft", "obliqueBottomRight"],
  },
];

// --- the page value ------------------------------------------------------------

function int(value: unknown, low: number, high: number): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= low && value <= high
    ? value
    : null;
}

function record(value: unknown, required: readonly string[], optional: readonly string[] = []) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const keys = Object.keys(value);
  if (!required.every((key) => keys.includes(key))) return null;
  if (keys.some((key) => !required.includes(key) && !optional.includes(key))) return null;
  return value as Record<string, unknown>;
}

function normalizeEffect<K extends OfficePictureEffectKind>(
  kind: K,
  value: unknown,
): OfficePictureEffectValues[K] | "none" | null {
  // A preset is a number; No Presets removes the key (never "none").
  if (kind === "preset") return isOfficePictureEffectPreset(value) ? (value as OfficePictureEffectValues[K]) : null;
  if (value === "none") return "none";
  if (kind === "shadow") {
    return normalizeOfficeShadowEffect(value) as OfficePictureEffectValues[K] | null;
  }
  if (kind === "glow") {
    const raw = record(value, ["radius", "color", "alpha"], ["themeColor", "satMod"]);
    if (!raw) return null;
    const radius = int(raw.radius, 1, MAX.glow);
    const alpha = int(raw.alpha, 0, 100_000);
    const color = typeof raw.color === "string" && /^#[0-9a-fA-F]{6}$/.test(raw.color) ? raw.color.toLowerCase() : null;
    if (radius === null || alpha === null || color === null) return null;
    const glow: { -readonly [P in keyof OfficeGlowEffect]: OfficeGlowEffect[P] } = { radius, color, alpha };
    if ("themeColor" in raw) {
      if (!(OFFICE_GLOW_THEME_SLOTS as readonly unknown[]).includes(raw.themeColor)) return null;
      glow.themeColor = raw.themeColor as OfficeGlowThemeSlot;
    }
    if ("satMod" in raw) {
      const satMod = int(raw.satMod, 0, 1_000_000);
      if (satMod === null || glow.themeColor === undefined) return null;
      glow.satMod = satMod;
    }
    return glow as OfficePictureEffectValues[K];
  }
  if (kind === "softEdge") {
    const raw = record(value, ["radius"]);
    const radius = raw ? int(raw.radius, 1, MAX.softEdge) : null;
    return radius === null ? null : ({ radius } as OfficePictureEffectValues[K]);
  }
  if (kind === "reflection") {
    const raw = record(value, ["blurRad", "stA", "endA", "endPos", "dist"]);
    if (!raw) return null;
    const reflection = {
      blurRad: int(raw.blurRad, 0, MAX.reflection),
      stA: int(raw.stA, 0, 100_000),
      endA: int(raw.endA, 0, 100_000),
      endPos: int(raw.endPos, 0, 100_000),
      dist: int(raw.dist, 0, MAX.reflection),
    };
    return Object.values(reflection).some((number) => number === null)
      ? null
      : (reflection as OfficePictureEffectValues[K]);
  }
  if (kind === "bevel") {
    const raw = record(value, ["preset", "width", "height"]);
    if (!raw || !(OFFICE_BEVEL_PRESETS as readonly unknown[]).includes(raw.preset)) return null;
    const width = int(raw.width, 0, MAX.bevel);
    const height = int(raw.height, 0, MAX.bevel);
    return width === null || height === null
      ? null
      : ({ preset: raw.preset as OfficeBevelPreset, width, height } as OfficePictureEffectValues[K]);
  }
  if (kind === "format3d") {
    const format3d = normalizeOfficeFormat3d(value);
    // Nothing but Office's defaults is no 3-D format (the server refuses `{}`).
    if (format3d && !Object.keys(format3d).length) return "none";
    return format3d as OfficePictureEffectValues[K] | null;
  }
  const raw = record(value, ["camera"], ["lat", "lon", "rev", "fov"]);
  if (!raw || !isOfficeCameraPreset(raw.camera)) return null;
  const angles = (["lat", "lon", "rev"] as const).filter((axis) => axis in raw);
  if (angles.length !== 0 && angles.length !== 3) return null;
  const rotation: Record<string, unknown> = { camera: raw.camera };
  for (const axis of angles) {
    const angle = int(raw[axis], 0, 21_599_999);
    if (angle === null) return null;
    rotation[axis] = angle;
  }
  if ("fov" in raw) {
    const fov = int(raw.fov, 0, OFFICE_CAMERA_MAX_FOV);
    if (fov === null) return null;
    rotation.fov = fov;
  }
  return rotation as unknown as OfficePictureEffectValues[K];
}

/**
 * A `format3d` value in the server's reading, or null when it is not one.
 * Office's defaults it may state (warm matte, the light from the top, a 0°
 * angle, an empty bottom bevel) are dropped, as the server never reads them,
 * so a value may come out empty (no 3-D format: write "none").
 */
export function normalizeOfficeFormat3d(value: unknown): OfficeFormat3dEffect | null {
  const raw = record(value, [], ["bottom", "depth", "contour", "material", "lighting"]);
  if (!raw) return null;
  const hex = (color: unknown): string | null | undefined =>
    color === undefined ? undefined : typeof color === "string" && /^#[0-9a-fA-F]{6}$/.test(color) ? color.toLowerCase() : null;
  const next: {
    -readonly [P in keyof OfficeFormat3dEffect]: OfficeFormat3dEffect[P];
  } = {};
  if ("bottom" in raw) {
    const bottom = normalizeEffect("bevel", raw.bottom);
    if (!bottom || bottom === "none") return null;
    if (bottom.width > 0 || bottom.height > 0) next.bottom = bottom;
  }
  for (const key of ["depth", "contour"] as const) {
    if (!(key in raw)) continue;
    const size = key === "depth" ? "height" : "width";
    const part = record(raw[key], [size], ["color"]);
    const amount = part ? int(part[size], 0, MAX.bevel) : null;
    const color = part ? hex(part.color) : null;
    if (!part || amount === null || color === null) return null;
    if (amount === 0) continue;
    next[key] = (color === undefined ? { [size]: amount } : { [size]: amount, color }) as never;
  }
  if ("material" in raw) {
    if (!(OFFICE_MATERIAL_PRESETS as readonly unknown[]).includes(raw.material)) return null;
    if (raw.material !== "warmMatte") next.material = raw.material as OfficeMaterialPreset;
  }
  if ("lighting" in raw) {
    const lighting = record(raw.lighting, ["rig"], ["direction", "angle"]);
    if (!lighting || !(OFFICE_LIGHT_RIGS as readonly unknown[]).includes(lighting.rig)) return null;
    if ("direction" in lighting && !(OFFICE_LIGHT_DIRECTIONS as readonly unknown[]).includes(lighting.direction)) return null;
    const angle = "angle" in lighting ? int(lighting.angle, 0, 21_599_999) : 0;
    if (angle === null) return null;
    const direction = (lighting.direction ?? "t") as OfficeLightDirection;
    if (lighting.rig !== "threePt" || direction !== "t" || angle !== 0) {
      next.lighting = {
        rig: lighting.rig as OfficeLightRig,
        ...(direction !== "t" ? { direction } : {}),
        ...(angle ? { angle } : {}),
      };
    }
  }
  return next;
}

/** The page's `pictureEffects`, or null when absent or not one (the server's rule). */
export function normalizeOfficePictureEffects(value: unknown): OfficePictureEffects | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const keys = Object.keys(value);
  if (!keys.length || keys.some((key) => !(OFFICE_PICTURE_EFFECT_KINDS as readonly string[]).includes(key))) {
    return null;
  }
  const next: Partial<Record<OfficePictureEffectKind, unknown>> = {};
  for (const kind of OFFICE_PICTURE_EFFECT_KINDS) {
    if (!(kind in value)) continue;
    const effect = normalizeEffect(kind, (value as Record<string, unknown>)[kind]);
    if (effect === null) return null;
    next[kind] = effect;
  }
  return next as OfficePictureEffects;
}

/**
 * The effects after one gallery pick: `value` replaces `kind`; `null` is
 * "No <effect>", which removes it -- or states `"none"` when the picture's
 * Picture Style draws that effect, so the style's own goes too. `undefined`
 * when the picture is left with no effects to name.
 */
export function withOfficePictureEffect<K extends OfficePictureEffectKind>(
  current: unknown,
  kind: K,
  value: OfficePictureEffectValues[K] | null,
  styleHasEffect: boolean,
): OfficePictureEffects | undefined {
  const base: Partial<Record<OfficePictureEffectKind, unknown>> = {
    ...(normalizeOfficePictureEffects(current) ?? {}),
  };
  if (kind === "preset") {
    // PowerPoint (picture-effect-no-preset.pptx): a Preset brings its own
    // scene, 3-D shape and shadow and keeps the glow, reflection and soft
    // edges picked before it. No Presets removes the scene and the 3-D shape
    // only: the preset's shadow stays, now stated, and so does its line (the
    // block's border fields, left alone here).
    const previous = isOfficePictureEffectPreset(base.preset) ? base.preset : undefined;
    delete base.preset;
    delete base.bevel;
    delete base.rotation3d;
    delete base.format3d;
    if (value !== null) {
      delete base.shadow;
      return { preset: value as OfficePictureEffectPreset, ...base } as OfficePictureEffects;
    }
    const presetShadow = previous !== undefined ? officePictureEffectPresetShadow(previous) : undefined;
    if (base.shadow === undefined && presetShadow) base.shadow = presetShadow;
    return Object.keys(base).length ? (base as OfficePictureEffects) : undefined;
  }
  if (value !== null) base[kind] = value;
  // A 3-D format the page does not state is the file's (the server keeps it):
  // taking it off states "none".
  else if (styleHasEffect || kind === "format3d") base[kind] = "none";
  else delete base[kind];
  return Object.keys(base).length ? (base as OfficePictureEffects) : undefined;
}

/** Whether two effect values are the same pick (the menu's checked state). */
export function sameOfficePictureEffect(left: unknown, right: unknown): boolean {
  if (normalizeOfficeShadowEffect(left) && normalizeOfficeShadowEffect(right)) {
    return sameOfficeShadowEffect(left, right);
  }
  if (!left || !right || typeof left !== "object" || typeof right !== "object") return left === right;
  const a = left as Record<string, unknown>;
  const b = right as Record<string, unknown>;
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  for (const key of keys) {
    // The glow's drawn colour follows the deck theme; its reference is the pick.
    if (key === "color" && a.themeColor !== undefined && a.themeColor === b.themeColor) continue;
    if (a[key] !== b[key]) return false;
  }
  return true;
}

// --- rendering ----------------------------------------------------------------

function rgba(hex: string, alpha: number): string {
  const value = Number.parseInt(hex.slice(1), 16);
  return `rgb(${(value >> 16) & 255} ${(value >> 8) & 255} ${value & 255} / ${Math.round(alpha / 1000) / 100})`;
}

function px(emu: number): number {
  return Math.round((emu / EMU_PER_PX) * 100) / 100;
}









export function officeCameraEyePx(size: { readonly width: number; readonly height: number }, fovDegrees: number): number {
  return Math.round((2 * Math.max(size.width, size.height)) / Math.tan((Math.min(179, Math.max(1, fovDegrees)) * Math.PI) / 360));
}

export interface OfficePictureEffectsCss {
  /** On a wrapper around the picture: glow, reflection, soft edges, 3-D rotation, an outer shadow. */
  readonly wrapper: CSSProperties;
  /** An overlay over the picture for the bevel's lit and shaded rims, if any. */
  readonly bevel?: CSSProperties;
  /** A layer under the picture for a perspective shadow, if any. */
  readonly shadowUnder?: CSSProperties;
  /** An overlay over the picture for an inner shadow, if any. */
  readonly shadowOver?: CSSProperties;
}

/**
 * The picture's effects on the canvas, approximated with CSS: a glow is a
 * soft halo in its colour and opacity, a reflection a faded mirror image
 * below (its offset, start/end opacity and length), soft edges an alpha ramp
 * as wide as the radius, a bevel lit and shaded rims, and 3-D rotation the
 * camera's measured X/Y/Z rotation (perspective cameras with their field of
 * view). `size` is the picture's frame in px.
 */
/**
 * `x y blur 0px color` (an outer box shadow with no spread) as the
 * `drop-shadow()` that draws the same shadow from the picture's own alpha;
 * null when it has a spread, which a drop shadow cannot state.
 */
function dropShadowOf(boxShadow: string): string | null {
  const match = /^(-?[\d.]+px) (-?[\d.]+px) ([\d.]+px) (-?[\d.]+)px (.+)$/.exec(boxShadow.trim());
  if (!match || Number(match[4]) !== 0) return null;
  return `drop-shadow(${match[1]} ${match[2]} ${match[3]} ${match[5]})`;
}

export function officePictureEffectsCss(
  value: unknown,
  size: { readonly width: number; readonly height: number },
  options: { readonly scale?: number } = {},
): OfficePictureEffectsCss | null {
  const effects = normalizeOfficePictureEffects(value);
  if (!effects) return null;
  // A Preset draws first; the effects picked over it add to it (a Shadow
  // picked over it replaces the preset's own).
  const preset =
    typeof effects.preset === "number"
      ? officePictureEffectPresetCss(effects.preset, size, {
          scale: options.scale,
          shadow: effects.shadow === undefined,
        })
      : null;
  const wrapper: CSSProperties = { ...(preset?.wrapper ?? {}) };
  const filters: string[] = preset?.wrapper.filter ? [String(preset.wrapper.filter)] : [];
  const glow = effects.glow;
  if (glow && glow !== "none") {
    const blur = Math.max(0.5, px(glow.radius) / 2);
    const color = rgba(glow.color, glow.alpha);
    filters.push(`drop-shadow(0 0 ${blur}px ${color})`, `drop-shadow(0 0 ${blur}px ${color})`);
  }
  const shadow = effects.shadow && effects.shadow !== "none" ? officeShadowCss(effects.shadow, size) : null;
  // An outer shadow of the picture's own size follows what the picture draws,
  // as PowerPoint's does: its transparent pixels cast none, its see-through
  // ones (Transparency) a lighter one, and the shadow shows through them.
  
  
  // land PowerPoint draws grey over its shadow, drawn white).
  const alphaShadow = shadow?.boxShadow ? dropShadowOf(shadow.boxShadow) : null;
  if (alphaShadow) filters.push(alphaShadow);
  if (filters.length) wrapper.filter = filters.join(" ");
  if (shadow?.boxShadow && !alphaShadow) {
    wrapper.boxShadow = [preset?.wrapper.boxShadow, shadow.boxShadow].filter(Boolean).join(", ");
  }
  // The layer under the picture draws behind it, not behind the slide.
  if (shadow?.under) wrapper.isolation = "isolate";
  const reflection = effects.reflection;
  if (reflection && reflection !== "none") {
    const start = 100 - reflection.endPos / 1000;
    wrapper.WebkitBoxReflect =
      `below ${px(reflection.dist)}px linear-gradient(to bottom, transparent ${start}%, ` +
      `rgb(0 0 0 / ${reflection.endA / 100_000}) ${start}%, rgb(0 0 0 / ${reflection.stA / 100_000}) 100%)`;
  }
  const softEdge = effects.softEdge;
  if (softEdge && softEdge !== "none") {
    const ramp = Math.min(px(softEdge.radius), size.width / 2, size.height / 2);
    const mask =
      `linear-gradient(to right, transparent, #000 ${ramp}px, #000 calc(100% - ${ramp}px), transparent), ` +
      `linear-gradient(to bottom, transparent, #000 ${ramp}px, #000 calc(100% - ${ramp}px), transparent)`;
    wrapper.maskImage = mask;
    wrapper.WebkitMaskImage = mask;
    wrapper.maskComposite = "intersect";
    wrapper.WebkitMaskComposite = "source-in";
  }
  const rotation = effects.rotation3d;
  if (rotation && rotation !== "none") {
    const [cameraX, cameraY, cameraZ] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
    // Format Picture > 3-D Rotation's X / Y / Z (`a:rot` on the camera)
    // override the camera's own, as on a shape (office-shape-effects-paint).
    const x = typeof rotation.lon === "number" ? rotation.lon / 60_000 : cameraX;
    const y = typeof rotation.lat === "number" ? rotation.lat / 60_000 : cameraY;
    const z = typeof rotation.rev === "number" ? rotation.rev / 60_000 : cameraZ;
    const perspective = rotation.camera.startsWith("perspective")
      ? `perspective(${officeCameraEyePx(size, officeCameraFieldOfView(rotation))}px) `
      : "";
    if (x || y || z) {
      wrapper.transform = `${perspective}rotateZ(${z}deg) rotateX(${-y}deg) rotateY(${x}deg)`;
    }
  }
  const bevel = effects.bevel;
  const bevelCss: CSSProperties | undefined =
    bevel && bevel !== "none" && (bevel.width > 0 || bevel.height > 0)
      ? (() => {
          const rim = Math.max(1, Math.min(px(bevel.width), size.width / 4, size.height / 4));
          return {
            position: "absolute",
            inset: 0,
            pointerEvents: "none",
            boxShadow: `inset ${rim / 2}px ${rim / 2}px ${rim}px rgb(255 255 255 / 0.55), inset ${-rim / 2}px ${-rim / 2}px ${rim}px rgb(0 0 0 / 0.35)`,
          };
        })()
      : undefined;
  const rim = bevelCss ?? preset?.bevel;
  const under = shadow?.under ?? preset?.shadowUnder;
  const over = shadow?.over ?? preset?.shadowOver;
  if (!Object.keys(wrapper).length && !rim && !under && !over) return null;
  return {
    wrapper,
    ...(rim ? { bevel: rim } : {}),
    ...(under ? { shadowUnder: under } : {}),
    ...(over ? { shadowOver: over } : {}),
  };
}
