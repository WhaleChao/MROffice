/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A shape's Office effects (`shapeEffects`, the model pictures and Docs
 * shapes already keep: lib/office-picture-effects) drawn on the canvas, the
 * same way in every editor. Office draws each effect on the shape's own
 * outline, so this paints on the shape's ALPHA, never its box:
 *
 * - soft edges, an inner shadow and a bevel are SVG filter primitives on the
 *   shape's alpha (`primitives`, rendered inside a host `<filter>` that the
 *   host references from its wrapper's `filter`);
 * - an outer shadow at an even scale and a glow are CSS `drop-shadow`s, which
 *   follow the painted alpha (`cssFilters`, after the SVG filter);
 * - a scaled or skewed outer shadow (Perspective) is a silhouette of the shape
 *   in the shadow colour, blurred, scaled and skewed about its alignment
 *   corner, under the shape (`silhouette`; the transform the picture shadow
 *   uses, lib/office-shadow-effects);
 * - a reflection is `-webkit-box-reflect`, a 3-D rotation the camera's CSS
 *   rotation (lib/office-picture-effects' measured camera table).
 *
 * Lengths are the file's EMU in CSS px (9525 EMU a px), as the picture
 * effects use them. A Preset draws its own shadow, bevel and camera under
 * whatever was picked over it (lib/office-picture-effect-presets).
 */

import {
  normalizeOfficePictureEffects,
  OFFICE_CAMERA_ROTATIONS,
  OFFICE_SOFT_EDGE_INSET,
  OFFICE_SOFT_EDGE_SIGMA,
  officeCameraEyePx,
  officeCameraLonTurn,
  officeCameraRevTurn,
  officeCameraFieldOfView,
  type OfficeFormat3dEffect,
  type OfficeLightRig,
  type OfficeMaterialPreset,
} from "./office-picture-effects";
import {
  isOfficePictureEffectPreset,
  officePictureEffectPresetEffects,
} from "./office-picture-effect-presets";
import { officeShadowCss, type OfficeShadowEffect } from "./office-shadow-effects";

const EMU_PER_PX = 9_525;

/** One SVG filter primitive, as data (each host renders it its own way). */
export interface OfficeShapeEffectPrimitive {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string | number>>;
  readonly children?: readonly OfficeShapeEffectPrimitive[];
}

/** A scaled or skewed outer shadow: the shape's silhouette under it. */
export interface OfficeShapeShadowSilhouette {
  /** The shadow colour with its opacity, as the silhouette's paint. */
  readonly color: string;
  /** The same colour alone, for a renderer that fills the shape in a hex and
   *  takes the opacity on the layer (`officeShapeSilhouetteAlpha`). */
  readonly hex: string;
  readonly transform: string;
  readonly transformOrigin: string;
  /** CSS `blur()`, or "". */
  readonly filter: string;
}

export interface OfficeShapeEffectsPaint {
  /** On the shape's alpha (soft edges, inner shadow, bevel); empty when none. */
  readonly primitives: readonly OfficeShapeEffectPrimitive[];
  /** CSS filter functions after the SVG filter (outer shadow, glow). */
  readonly cssFilters: readonly string[];
  readonly silhouette: OfficeShapeShadowSilhouette | null;
  /** `-webkit-box-reflect`, or "". */
  readonly reflect: string;
  /** The 3-D rotation's CSS transform, or "". */
  readonly transform: string;
  /** How far past the shape the filter must draw, a fraction of its size on
   * each side, when more than the usual half: a deep extrusion's walls
   * (shape3d-probe p6: Isometric Top Up 72 pt deep had its walls cut flat). */
  readonly region?: number;
}

/** The filter region a paint draws in: half the shape past each side, or
 * further for a deep extrusion (`region`). */
export function officeShapeEffectsFilterRegion(paint: Pick<OfficeShapeEffectsPaint, "region">): {
  x: string;
  y: string;
  width: string;
  height: string;
} {
  const margin = Math.round((paint.region ?? 0.5) * 100);
  return { x: `-${margin}%`, y: `-${margin}%`, width: `${100 + 2 * margin}%`, height: `${100 + 2 * margin}%` };
}

function px(emu: number): number {
  return Math.round((emu / EMU_PER_PX) * 100) / 100;
}

function rgba(hex: string, alpha: number): string {
  const value = Number.parseInt(hex.slice(1), 16);
  return `rgb(${(value >> 16) & 255} ${(value >> 8) & 255} ${value & 255} / ${Math.round(alpha / 1000) / 100})`;
}

/**
 * An outer shadow as one CSS `drop-shadow` on the shape's alpha: its offset
 * along `dir`, its blur and colour. A Perspective shadow's scale and skew
 * are not in it (the silhouette draws those).
 */
export function officeShapeShadowDropShadow(shadow: OfficeShadowEffect): string {
  const radians = (shadow.dir / 60_000) * (Math.PI / 180);
  const distance = px(shadow.dist);
  const x = Math.round(Math.cos(radians) * distance * 100) / 100 + 0;
  const y = Math.round(Math.sin(radians) * distance * 100) / 100 + 0;
  return `drop-shadow(${x}px ${y}px ${px(shadow.blurRad) / 2}px ${rgba(shadow.color, shadow.alpha)})`;
}

/** The opacity a Perspective shadow's silhouette is drawn at (its colour's alpha). */
export function officeShapeSilhouetteAlpha(paint: OfficeShapeEffectsPaint): number {
  const alpha = /\/\s*([\d.]+)\)$/.exec(paint.silhouette?.color ?? "")?.[1];
  return alpha === undefined ? 1 : Math.min(1, Math.max(0, Number(alpha)));
}

const p = (
  tag: string,
  attrs: Record<string, string | number>,
  children?: OfficeShapeEffectPrimitive[],
): OfficeShapeEffectPrimitive => (children ? { tag, attrs, children } : { tag, attrs });

/**
 * The effects a shape carries, as paint. Null when it carries none (the host
 * then draws the shape exactly as before).
 */
/** Shape Outline > Sketched, as `a:ln`'s `ask:lineSketchStyleProps` names it. */
export type OfficeLineSketch = "curved" | "freehand" | "scribble";

export function normalizeOfficeLineSketch(value: unknown): OfficeLineSketch | undefined {
  return value === "curved" || value === "freehand" || value === "scribble" ? value : undefined;
}

/**
 * How far a Sketched outline wanders from the true one, in px at 1x, and how
 * fine its wobble is (cycles a px): Curved drifts gently, Freehand more and
 * finer, Scribble most (an approximation: Office's own sketch generator is
 * not published).
 */
const SKETCH_LOOK: Record<OfficeLineSketch, { readonly scale: number; readonly frequency: number; readonly octaves: number }> = {
  curved: { scale: 2.5, frequency: 0.006, octaves: 1 },
  freehand: { scale: 4, frequency: 0.011, octaves: 1 },
  scribble: { scale: 5, frequency: 0.018, octaves: 2 },
};

/** A stable small seed from an object id, so a sketch keeps its shape. */
export function officeSketchSeed(id: string): number {
  let hash = 0;
  for (let index = 0; index < id.length; index += 1) hash = (hash * 31 + id.charCodeAt(index)) | 0;
  return Math.abs(hash % 997) + 1;
}

// --- 3-D Format: the face as PowerPoint lights it ---------------------------
//
// Windows PowerPoint's own rendering (Slide.Export PNG) of a face seen square
// under a circle top bevel 6 x 6 pt (chart24/bevel face-probe5: every material

// #C00000): every channel of the face is a line of its fill, `scale * c +



// highlight and, metals, scales a little more; Matte, Dark Edge and Soft Edge


// Mac lit warm matte warmer (x1.10/1.08/1.06), which this replaced.

type Rgb = readonly [number, number, number];
 
const RIG_FACE_SCALE: Readonly<Partial<Record<OfficeLightRig, Rgb>>> = {
  threePt: [1.036, 1.035, 1.035],
  balanced: [0.945, 0.9459, 0.945],
  soft: [0.9643, 0.9653, 0.965],
  harsh: [0.8344, 0.839, 0.8352],
  flood: [0.8536, 0.8584, 0.8551],
  contrasting: [0.9942, 0.996, 0.995],
  morning: [0.8705, 0.8471, 0.7652],
  sunrise: [0.8397, 0.7889, 0.6601],
  sunset: [0.859, 0.6901, 0.5399],
  chilly: [0.9942, 0.996, 1.045],
  freezing: [0.6687, 0.7502, 0.8699],
  flat: [0.9942, 0.996, 0.995],
  twoPt: [0.9865, 0.9847, 0.985],
  glow: [0.9942, 0.996, 0.995],
  brightRoom: [0.9942, 0.996, 0.995],
};
 
const MATERIAL_SHINE: Readonly<Partial<Record<OfficeMaterialPreset, Readonly<Partial<Record<OfficeLightRig, readonly [k: number, white: number]>>>>>> = {
  warmMatte: { threePt: [1.0000, 5.0], balanced: [1.0000, 6.9], soft: [1.0000, 2.0], twoPt: [1.0000, 5.8] },
  plastic: { threePt: [1.0000, 2.0], balanced: [1.0000, 3.9], soft: [1.0000, 1.0], twoPt: [1.0000, 2.8] },
  metal: { threePt: [1.0043, 4.9], balanced: [0.9953, 9.0], soft: [1.0117, 1.1], twoPt: [0.9860, 7.9] },
  softmetal: { threePt: [1.0472, 7.9], balanced: [1.0583, 12.0], soft: [1.0218, 4.0], twoPt: [1.0426, 9.3] },
};
 
const MATERIAL_WASH: Readonly<Partial<Record<OfficeMaterialPreset, Readonly<Partial<Record<OfficeLightRig, readonly number[]>>>>>> = {
  powder: {
    threePt: [0.5294, 0.5276, 0.5294, 178.0, 178.0, 178.0],
    balanced: [0.464, 0.4632, 0.4608, 164.8, 165.0, 165.0],
    soft: [0.4819, 0.4766, 0.4798, 148.4, 149.2, 149.0],
    harsh: [0.4181, 0.4165, 0.4199, 122.3, 122.0, 122.0],
    flood: [0.4274, 0.4245, 0.425, 131.2, 131.7, 131.9],
    contrasting: [0.4965, 0.4927, 0.4949, 131.4, 131.6, 131.7],
    morning: [0.4297, 0.4245, 0.3799, 115.8, 111.7, 101.1],
    sunrise: [0.4166, 0.3971, 0.3352, 110.2, 103.0, 86.3],
    sunset: [0.4243, 0.3487, 0.275, 113.0, 88.1, 68.0],
    chilly: [0.4965, 0.504, 0.53, 129.4, 127.9, 135.0],
    freezing: [0.3286, 0.3664, 0.435, 89.0, 99.8, 115.7],
    flat: [0.4994, 0.4908, 0.499, 153.8, 154.0, 153.8],
    twoPt: [0.4841, 0.4816, 0.4804, 160.9, 161.0, 161.0],
    glow: [0.4965, 0.4927, 0.4949, 131.4, 131.6, 131.7],
    brightRoom: [0.4965, 0.4927, 0.4949, 147.4, 147.6, 147.7],
  },
  translucentPowder: {
    threePt: [0.7227, 0.7219, 0.7255, 123.7, 124.0, 124.0],
    balanced: [0.6629, 0.6667, 0.6591, 119.8, 120.0, 120.2],
    soft: [0.6714, 0.6707, 0.6701, 102.0, 101.9, 101.9],
    harsh: [0.5857, 0.5872, 0.5851, 93.1, 92.9, 93.0],
    flood: [0.5972, 0.5973, 0.5999, 98.6, 98.8, 98.8],
    contrasting: [0.6937, 0.6901, 0.69, 80.5, 80.8, 80.9],
    morning: [0.608, 0.5892, 0.5399, 80.6, 81.1, 79.8],
    sunrise: [0.5857, 0.5484, 0.46, 80.1, 79.9, 78.8],
    sunset: [0.6026, 0.479, 0.3749, 79.4, 76.8, 77.0],
    chilly: [0.6937, 0.6901, 0.73, 78.5, 78.8, 78.9],
    freezing: [0.4696, 0.5198, 0.6099, 80.0, 80.9, 81.0],
    flat: [0.6937, 0.6901, 0.69, 103.5, 103.8, 103.9],
    twoPt: [0.6822, 0.682, 0.6801, 112.0, 112.1, 112.1],
    glow: [0.6937, 0.6901, 0.69, 80.5, 80.8, 80.9],
    brightRoom: [0.6937, 0.6901, 0.69, 96.5, 96.8, 96.9],
  },
  clear: {
    threePt: [0.0969, 0.0969, 0.0999, 229.8, 229.9, 230.0],
    balanced: [0.0915, 0.0876, 0.0899, 229.7, 229.8, 229.8],
    soft: [0.0915, 0.0876, 0.0899, 229.7, 229.8, 229.8],
    harsh: [0.0776, 0.0775, 0.0799, 229.6, 229.9, 230.0],
    flood: [0.083, 0.0775, 0.0799, 229.4, 229.9, 230.0],
    contrasting: [0.0969, 0.0969, 0.0951, 229.8, 229.9, 230.1],
    morning: [0.083, 0.0775, 0.0699, 229.4, 229.9, 229.8],
    sunrise: [0.0776, 0.0682, 0.0599, 229.6, 229.9, 230.0],
    sunset: [0.083, 0.0581, 0.0499, 229.4, 229.9, 229.8],
    chilly: [0.0969, 0.0969, 0.0999, 229.8, 229.9, 230.0],
    freezing: [0.0607, 0.0682, 0.0799, 229.9, 229.9, 230.0],
    flat: [0.0969, 0.0969, 0.0951, 229.8, 229.9, 230.1],
    twoPt: [0.0915, 0.0876, 0.0899, 229.7, 229.8, 229.8],
    glow: [0.0969, 0.0969, 0.0951, 229.8, 229.9, 230.1],
    brightRoom: [0.0969, 0.0969, 0.0951, 229.8, 229.9, 230.1],
  },
};

/** Wireframe draws grey (as PowerPoint renders it to a picture). */
const WIREFRAME_FACE = 166;



function officeFaceLight(material: string | undefined, rig: string | undefined): { scale: Rgb; offset: Rgb } {
  const lightRig = (rig && rig in RIG_FACE_SCALE ? rig : "threePt") as OfficeLightRig;
  const preset = (material ?? "warmMatte") as OfficeMaterialPreset;
  if (preset === "legacyWireframe") return { scale: [0, 0, 0], offset: [WIREFRAME_FACE, WIREFRAME_FACE, WIREFRAME_FACE] };
  if (preset === "flat") return { scale: RIG_FACE_SCALE.flat!, offset: [0, 0, 0] };
  const wash = MATERIAL_WASH[preset]?.[lightRig];
  if (wash) return { scale: [wash[0]!, wash[1]!, wash[2]!], offset: [wash[3]!, wash[4]!, wash[5]!] };
  const rigScale = RIG_FACE_SCALE[lightRig]!;
  // Matte, Dark Edge and Soft Edge shine not at all; a material not measured
  // (a legacy one) shines as warm matte does.
  const dull = preset === "matte" || preset === "dkEdge" || preset === "softEdge";
  const [k, white] = (dull ? undefined : (MATERIAL_SHINE[preset] ?? MATERIAL_SHINE.warmMatte)?.[lightRig]) ?? [1, 0];
  return { scale: rigScale.map((value) => value * k) as unknown as Rgb, offset: [white, white, white] };
}

type FaceRows = readonly [readonly number[], readonly number[], readonly number[]];
const scaled = (r: number, g: number, b: number): FaceRows => [
  [r, 0, 0, 0, 0],
  [0, g, 0, 0, 0],
  [0, 0, b, 0, 0],
];
/** Where `ST_LightRigDirection` puts the light, clockwise from the top. */
const LIGHT_DIRECTION_DEG: Readonly<Record<string, number>> = {
  t: 0, tr: 45, r: 90, br: 135, b: 180, bl: 225, l: 270, tl: 315,
};

/** A wall of a solid by the way it faces on the slide. */
type WallFacing = "up" | "right" | "down" | "left";










const LEGACY_SOLID_LIGHT: Readonly<
  Record<string, { readonly face: number; readonly walls: Readonly<Record<WallFacing, number>> }>
> = {
  "legacyMatte legacyFlat3 b": { face: 0.76, walls: { up: 0.592, right: 0.875, down: 1, left: 0.875 } },
};

function legacySolidLight(format3d: OfficeFormat3dEffect | null | undefined) {
  const lighting = format3d?.lighting;
  if (!format3d || !lighting || lighting.angle) return undefined;
  return LEGACY_SOLID_LIGHT[`${format3d.material ?? ""} ${lighting.rig} ${lighting.direction ?? "t"}`];
}














export function officeShape3dFaceTiltWhite(
  format3d: OfficeFormat3dEffect | null | undefined,
  tilt: { readonly x: number; readonly y: number } | null | undefined,
): number {
  if (!tilt) return 0;
  const lighting = format3d?.lighting;
  if (lighting && (lighting.rig !== "threePt" || (lighting.direction ?? "t") !== "t" || lighting.angle)) return 0;
  if (format3d?.material && format3d.material !== "warmMatte") return 0;
  // Where the eye stands in the shape's own frame (the lights stand there,
  // as the walls' do): the host's turn rotateZ(z) rotateX(-y) rotateY(x)
  // taken back off the view direction (z, the turn in the face's plane,
  // moves nothing). A highlight is the eye near a light's mirror direction.
  const rad = Math.PI / 180;
  const eye = (x: number, y: number): readonly [number, number, number] => [
    -Math.cos(y * rad) * Math.sin(x * rad),
    -Math.sin(y * rad),
    Math.cos(y * rad) * Math.cos(x * rad),
  ];
  // (Fitted with X as Office's X Rotation: the table's turn the other way.)
  const view = eye(-tilt.x, tilt.y);
  // A raised cosine about each lobe's eye direction, half its peak at
  // `half` degrees off it (the probe's half-height points).
  const lobe = (at: readonly [number, number, number], half: number, peak: number) => {
    const cos = Math.max(-1, Math.min(1, view[0] * at[0] + view[1] * at[1] + view[2] * at[2]));
    const t = Math.acos(cos) / rad / (2 * half);
    return t >= 1 ? 0 : peak * Math.cos((Math.PI / 2) * t) ** 2;
  };
  // The key lobe (Y -42.5: the eye below the face's middle, the key light
  // above it), the fill lobe (Y +65) and the side lights' small tint by X
  // (isometricTopUp's face, its eye 35 degrees off the key lobe, is not lit:
  // p5, +0 .. -3).
  const white = lobe(eye(0, -42.5), 22.5, 55) + lobe(eye(0, 65), 30, 23.5) + Math.max(-5, Math.min(4, -10 * view[0]));
  return Math.round((white / 255) * 1000) / 1000;
}

/** The face's colour matrix (`feColorMatrix` values) under a 3-D format. */
export function officeShape3dFaceMatrix(
  format3d: OfficeFormat3dEffect | null | undefined,
  options: { readonly gradientFill?: boolean; readonly tiltWhite?: number } = {},
): string {
  const legacy = legacySolidLight(format3d);
  if (legacy) {
    
    // reach their stop colours, #33CCCC drawn #34CAC9), a solid one shaded.
    const face = options.gradientFill ? 1 : legacy.face;
    return [...scaled(face, face, face).map((row) => row.join(" ")), "0 0 0 1 0"].join("  ");
  }
  const face = officeFaceLight(format3d?.material, format3d?.lighting?.rig);
  const lit = [0, 1, 2].map((index) => {
    const row = [0, 0, 0, 0, Math.round((face.offset[index]! / 255) * 1000) / 1000];
    row[index] = Math.round(face.scale[index]! * 1000) / 1000;
    return row;
  });
  // The highlight a turned face catches (`officeShape3dFaceTiltWhite`), added.
  const white = options.tiltWhite ?? 0;
  const withWhite = white ? lit.map((row) => row.map((value, column) => (column === 4 ? Math.round((value + white) * 1000) / 1000 : value))) : lit;
  return [...withWhite.map((row) => row.join(" ")), "0 0 0 1 0"].join("  ");
}

/** The camera's X / Y turns (degrees) of the shape's 3-D Rotation, or its preset's. */
function cameraTilt(
  effects: NonNullable<ReturnType<typeof normalizeOfficePictureEffects>>,
  preset: ReturnType<typeof officePictureEffectPresetEffects> | null,
): { x: number; y: number } | null {
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  if (rotation) {
    // Measured through perspective cameras (picture-preset-camera-probe p4,
    // p5's Relaxed); an isometric one's face turned as far caught none (p5).
    if (!rotation.camera.startsWith("perspective")) return null;
    const [x, y] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
    return {
      x: typeof rotation.lon === "number" ? officeCameraLonTurn(rotation.lon) : x,
      y: typeof rotation.lat === "number" ? rotation.lat / 60_000 : y,
    };
  }
  if (effects.rotation3d === undefined && preset?.camera) return { x: officeCameraLonTurn(preset.camera.lon), y: preset.camera.lat / 60_000 };
  return null;
}

/**
 * Office's top-bevel presets as height profiles across the bevel: the height
 * (in units of the bevel's height) at t = 0, 1/16, ..., 1 of its width in
 * from the outline, joined by a monotone cubic (PCHIP). Fitted to Windows
 * PowerPoint's Slide.Export PNG cross-sections of each preset (chart24/bevel
 * fitprofile.py: bevel-probe2 at 12 pt, bevel-profile-probe at 6/12/18 pt,
 * grey and blue) under the Three Point lights below; `circle` (the default)
 * is the quarter circle it is named for.
 */
const BEVEL_HEIGHT: Readonly<Record<string, readonly number[]>> = {
  circle: [0.0, 0.3837, 0.5289, 0.624, 0.6961, 0.755, 0.8079, 0.8547, 0.8956, 0.9297, 0.9552, 0.9735, 0.9859, 0.9941, 0.9991, 1.0016, 1.0023],
  relaxedInset: [0.0, 0.5503, 0.7286, 0.8421, 0.9215, 0.9834, 1.0349, 1.0773, 1.1159, 1.0564, 0.9896, 0.9261, 0.8728, 0.8331, 0.8128, 0.804, 0.8019],
  cross: [0.0, 0.0632, 0.068, 0.068, 0.068, 0.068, 0.068, 0.072, 0.1462, 0.2128, 0.2174, 0.2174, 0.2174, 0.2174, 0.2174, 0.2174, 0.2174],
  coolSlant: [0.0, 0.4392, 0.6174, 0.7729, 0.9218, 1.0393, 1.1026, 1.1475, 1.173, 1.1773, 1.1777, 1.1777, 1.1777, 1.1777, 1.1777, 1.1777, 1.1777],
  angle: [0.0, 0.0606, 0.1215, 0.1824, 0.2431, 0.3039, 0.3647, 0.4255, 0.4863, 0.5471, 0.6079, 0.6687, 0.7295, 0.7903, 0.851, 0.9119, 0.9656],
  softRound: [0.0, 0.7158, 0.9732, 0.9684, 0.8805, 0.9593, 0.8744, 0.8029, 0.7437, 0.6938, 0.6519, 0.6187, 0.5962, 0.5823, 0.5737, 0.5694, 0.5681],
  convex: [0.0, 0.0544, 0.0577, 0.0639, 0.0776, 0.1092, 0.1499, 0.1951, 0.2449, 0.3001, 0.3586, 0.4244, 0.2638, 0.3867, 0.5621, 0.6544, 0.6612],
  slope: [0.0, -0.4886, 0.0417, 0.1812, 0.2223, 0.2272, 0.2279, 0.2279, 0.2274, 0.2239, 0.194, 0.0934, -0.6042, -2.0416, -2.5401, -2.6018, -2.6058],
  divot: [0.0, 0.4512, 0.641, 0.7674, 0.843, 0.8971, 0.9429, 0.9821, 1.0068, 1.0115, 1.0115, 1.0114, 1.0114, 1.0114, 1.0175, 1.053, 1.0587],
  riblet: [0.0, 0.4495, 0.5994, 0.7, 0.7773, 0.8433, 0.9001, 0.9514, 0.9972, 1.0336, 1.0451, 1.0483, 1.0491, 1.0491, 1.0492, 1.0561, 1.0835],
  hardEdge: [0.0, 1.0443, 2.054, 2.2517, 2.2587, 2.2591, 2.2591, 2.2591, 2.2591, 2.2591, 2.2558, 2.1728, 2.1269, 2.1006, 2.09, 2.0855, 2.0842],
  artDeco: [0.0, 0.1205, 0.1752, 0.2166, 0.2312, 0.2318, 0.4109, 0.537, 0.6143, 0.6687, 0.7176, 0.7606, 0.7947, 0.8105, 0.8194, 0.8237, 0.8246],
};










const BEVEL_RIG_LIGHTS: Readonly<Record<string, { readonly ambient: number; readonly lights: readonly (readonly [azimuth: number, elevation: number, strength: number])[] }>> = {
  threePt: { ambient: 0.0, lights: [[299.83, 65.33, 3.757], [98.93, 0.0, 2.557], [221.88, 6.92, 2.686]] },
  balanced: { ambient: 0.281, lights: [[93.36, 40.08, 1.222], [172.85, 26.01, 0.831], [293.2, 0.11, 1.623]] },
  soft: { ambient: 1.372, lights: [[311.47, 0.71, 1.123], [246.19, 8.07, 0.849], [69.05, 81.11, 1.087]] },
  harsh: { ambient: 1.072, lights: [[184.32, 2.92, 2.204], [134.09, 36.52, 2.477], [260.08, 11.19, 1.696]] },
  flood: { ambient: 0.238, lights: [[116.09, 24.67, 2.293], [190.21, 10.84, 3.241], [283.88, 19.8, 2.033]] },
  contrasting: { ambient: 2.133, lights: [[218.54, 11.65, 0.938], [0.15, 23.46, 1.302], [140.97, 11.03, 0.93]] },
  morning: { ambient: 1.054, lights: [[157.92, 25.76, 3.481], [112.16, 0.72, 1.357], [351.59, 15.26, 1.385]] },
  sunrise: { ambient: 0.24, lights: [[337.19, 16.9, 1.343], [160.36, 21.97, 2.511], [82.72, 1.94, 0.892]] },
  sunset: { ambient: 1.56, lights: [[146.67, 31.28, 2.231], [134.97, 2.41, 1.137], [346.96, 22.54, 0.868]] },
  chilly: { ambient: 2.24, lights: [[298.31, 21.4, 0.897], [132.4, 37.31, 2.866], [149.4, 2.23, 1.905]] },
  freezing: { ambient: 0.061, lights: [[25.0, 0.89, 1.284], [151.36, 23.84, 3.312], [315.3, 16.87, 1.425]] },
  flat: { ambient: 0.244, lights: [[321.93, 15.82, 1.413], [188.63, 22.1, 1.301], [86.1, 11.32, 1.232]] },
  twoPt: { ambient: 1.106, lights: [[80.9, 9.68, 1.838], [142.68, 1.12, 2.512], [257.75, 71.04, 3.15]] },
  glow: { ambient: 2.006, lights: [[266.79, 0.0, 0.236], [176.6, 5.71, 1.529], [3.52, 2.55, 1.363]] },
  brightRoom: { ambient: 1.287, lights: [[170.97, 8.06, 1.632], [44.86, 11.12, 1.149], [294.28, 1.35, 1.026]] },
};
const BEVEL_STEEPNESS = 1.1505;
const DEFAULT_BEVEL_HEIGHT_PX = 8;
/** Entries of the table mapping the blurred outline to the profile's height. */
const BEVEL_TABLE_SIZE = 128;

/** A monotone cubic through (k / n, values[k]) at `t` in 0..1 (Fritsch-Carlson). */
function pchip(values: readonly number[], t: number): number {
  const n = values.length - 1;
  const x = Math.min(n, Math.max(0, t * n));
  const i = Math.min(n - 1, Math.floor(x));
  const slope = (k: number) => {
    const left = k > 0 ? values[k]! - values[k - 1]! : null;
    const right = k < n ? values[k + 1]! - values[k]! : null;
    if (left === null) return right!;
    if (right === null) return left;
    return left * right <= 0 ? 0 : (2 * left * right) / (left + right);
  };
  const u = x - i;
  const [h00, h10, h01, h11] = [2 * u ** 3 - 3 * u ** 2 + 1, u ** 3 - 2 * u ** 2 + u, -2 * u ** 3 + 3 * u ** 2, u ** 3 - u ** 2];
  return h00 * values[i]! + h10 * slope(i) + h01 * values[i + 1]! + h11 * slope(i + 1);
}

/** The inverse of the standard normal CDF (Acklam's rational approximation). */
function normalQuantile(p: number): number {
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const low = 0.02425;
  if (p < low) {
    const q = Math.sqrt(-2 * Math.log(p));
    return (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) / ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1);
  }
  if (p > 1 - low) return -normalQuantile(1 - p);
  const q = p - 0.5;
  const r = q * q;
  return ((((((a[0]! * r + a[1]!) * r + a[2]!) * r + a[3]!) * r + a[4]!) * r + a[5]!) * q) / (((((b[0]! * r + b[1]!) * r + b[2]!) * r + b[3]!) * r + b[4]!) * r + 1);
}

const bevelTables = new Map<string, { readonly table: string; readonly range: number }>();
/**
 * The outline blurred by half the bevel's width reads, along a straight edge,
 * Phi(2 t) at t of the width in (0.5 on the outline): the table takes that to
 * the profile's height, normalised into 0..1 (`range`: its span in units of
 * the bevel's height).
 */
function bevelHeightTable(preset: string): { readonly table: string; readonly range: number } {
  const known = bevelTables.get(preset);
  if (known) return known;
  const knots = BEVEL_HEIGHT[preset] ?? BEVEL_HEIGHT.circle!;
  const heights = Array.from({ length: BEVEL_TABLE_SIZE + 1 }, (_, index) => {
    const a = index / BEVEL_TABLE_SIZE;
    if (a <= 0.5) return knots[0]!;
    const t = a >= 1 ? 1 : Math.min(1, normalQuantile(a) / 2);
    return pchip(knots, t);
  });
  const low = Math.min(...heights);
  const high = Math.max(...heights);
  const range = high - low || 1;
  const made = { table: heights.map((value) => fixedPx((value - low) / range)).join(" "), range };
  bevelTables.set(preset, made);
  return made;
}








function officeBevelLightPrimitives(
  bevel: { preset?: string; width: number; height?: number },
  format3d: OfficeFormat3dEffect | null,
  face: string,
  alpha: string,
  size: { readonly width: number; readonly height: number },
): OfficeShapeEffectPrimitive[] {
  const widthPx = Math.max(1, Math.min(px(bevel.width), size.width / 4, size.height / 4));
  const heightPx = bevel.height === undefined ? DEFAULT_BEVEL_HEIGHT_PX : px(bevel.height);
  const { table, range } = bevelHeightTable(bevel.preset ?? "circle");
  const rig = BEVEL_RIG_LIGHTS[format3d?.lighting?.rig ?? "threePt"] ?? BEVEL_RIG_LIGHTS.threePt!;
  const turn = lightTurn(format3d);
  const flat = rig.ambient + rig.lights.reduce((total, [, elevation, strength]) => total + strength * Math.max(0, Math.sin((elevation * Math.PI) / 180)), 0);
  // Half scale (a lighting result is capped at 1), doubled on the face.
  const unit = 1 / (2 * flat);
  const out: OfficeShapeEffectPrimitive[] = [
    p("feGaussianBlur", { in: alpha, stdDeviation: fixedPx(widthPx / 2), result: "bevelSpread" }),
    p("feComponentTransfer", { in: "bevelSpread", result: "bevelHeight" }, [p("feFuncA", { type: "table", tableValues: table })]),
  ];
  rig.lights.forEach(([azimuth, elevation, strength], index) => {
    const gray = Math.round(Math.min(1, strength * unit) * 255);
    out.push(
      p(
        "feDiffuseLighting",
        {
          in: "bevelHeight",
          surfaceScale: fixedPx(BEVEL_STEEPNESS * heightPx * range),
          diffuseConstant: 1,
          "lighting-color": `rgb(${gray},${gray},${gray})`,
          result: `bevelLight${index}`,
        },
        [p("feDistantLight", { azimuth: fixedPx((((azimuth + turn) % 360) + 360) % 360), elevation })],
      ),
    );
    out.push(
      index === 0
        ? p("feComposite", { in: "bevelLight0", in2: "bevelLight0", operator: "arithmetic", k1: 0, k2: 1, k3: 0, k4: fixedPx(rig.ambient * unit), result: "bevelShade0" })
        : p("feComposite", { in: `bevelShade${index - 1}`, in2: `bevelLight${index}`, operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: `bevelShade${index}` }),
    );
  });
  const shade = `bevelShade${rig.lights.length - 1}`;
  out.push(
    // The face made opaque first, so its anti-aliased edge is shaded in its
    // own colour, then given its coverage back by the clip.
    p("feColorMatrix", {
      in: face,
      type: "matrix",
      values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1",
      result: "bevelFace",
    }),
    p("feComposite", { in: "bevelFace", in2: shade, operator: "arithmetic", k1: 2, k2: 0, k3: 0, k4: 0, result: "bevelLit" }),
    p("feComposite", { in: "bevelLit", in2: face, operator: "in", result: "beveled" }),
  );
  return out;
}

const fixedPx = (value: number) => Math.round(value * 1000) / 1000;









const WALL_LIGHTS: readonly (readonly [azimuth: number, strength: number])[] = [
  [270, 0.63],
  [90, 0.845],
  [180, 0.47],
];
const WALL_AMBIENT = 0.115;





















const WALL_LIFTS: readonly (readonly [azimuth: number, table: readonly (readonly [number, number])[], crossed: boolean])[] = [
  [90, [[0, 0], [20, 36], [40, 54], [50.4, 43], [70, 20], [90, 0]], true],
  [180, [[0, 0], [20, 16], [40, 11], [60, 2], [90, 0]], true],
  [0, [[0, 0], [20, 3], [40, 0]], false],
];






export function officeShape3dWallLifts(
  format3d: OfficeFormat3dEffect | null | undefined,
  tilt: { readonly x: number; readonly y: number } | null | undefined,
): (readonly [azimuth: number, strength: number])[] {
  if (!tilt) return [];
  const lighting = format3d?.lighting;
  if (lighting && lighting.rig !== "threePt") return [];
  if (format3d?.material && format3d.material !== "warmMatte") return [];
  const turn = lightTurn(format3d);
  const rad = Math.PI / 180;
  const lifts: (readonly [number, number])[] = [];
  for (const [azimuth, table, crossed] of WALL_LIFTS) {
    const facing = (azimuth + turn) % 360;
    // How far the eye leans toward the wall, and across it: a turn of the
    // camera by X shows the right wall, by -Y the bottom one (y down).
    // (The camera table's X turns the other way: its +20 shows the left wall.)
    const lean = -tilt.x * Math.cos(facing * rad) - tilt.y * Math.sin(facing * rad);
    const across = Math.abs(tilt.x * Math.sin(facing * rad) - tilt.y * Math.cos(facing * rad));
    if (!(lean > 0)) continue;
    let lift = 0;
    for (let index = 1; index < table.length; index += 1) {
      const [a0, v0] = table[index - 1], [a1, v1] = table[index];
      if (lean <= a1) {
        lift = v0 + ((v1 - v0) * (lean - a0)) / (a1 - a0);
        break;
      }
    }
    if (crossed) lift *= Math.cos(Math.min(90, across) * rad) ** 20;
    if (lift > 0.5) lifts.push([facing, Math.round((lift / 255) * 1000) / 1000]);
  }
  return lifts;
}

/**
 * The camera's lean in the shape's frame (degrees about the vertical and
 * the horizontal axis) for the walls' lift: an explicit 3-D Rotation, else
 * its preset camera's, whatever its projection (an orthographic camera's
 * walls are lit as a perspective one's, wall-brightness-probe p1-p8).
 */
function cameraLean(
  effects: NonNullable<ReturnType<typeof normalizeOfficePictureEffects>>,
  preset: ReturnType<typeof officePictureEffectPresetEffects> | null,
): { x: number; y: number } | null {
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  if (rotation) {
    const [x, y] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
    const signed = (value: number) => ((((value + 180) % 360) + 360) % 360) - 180;
    return {
      x: signed(typeof rotation.lon === "number" ? officeCameraLonTurn(rotation.lon) : x),
      y: signed(typeof rotation.lat === "number" ? rotation.lat / 60_000 : y),
    };
  }
  if (effects.rotation3d === undefined && preset?.camera) {
    const signed = (value: number) => ((((value + 180) % 360) + 360) % 360) - 180;
    return { x: signed(officeCameraLonTurn(preset.camera.lon)), y: signed(preset.camera.lat / 60_000) };
  }
  return null;
}

/** The light's turn (degrees clockwise) the bevel's lit rims follow. */
function lightTurn(format3d: OfficeFormat3dEffect | null | undefined): number {
  const lighting = format3d?.lighting;
  return (LIGHT_DIRECTION_DEG[lighting?.direction ?? "t"] ?? 0) + (lighting?.angle ?? 0) / 60_000;
}

const OBLIQUE_DEPTH_SCALE = 0.39;






const LEGACY_OBLIQUE_DEPTH_SCALE = 0.55;
/** Oblique cameras look at the face square on and slant the depth (unit direction). */
const OBLIQUE_DEPTH: Readonly<Record<string, readonly [number, number]>> = {
  obliqueTopLeft: [-1, -1],
  obliqueTop: [0, -1],
  obliqueTopRight: [1, -1],
  obliqueLeft: [-1, 0],
  obliqueRight: [1, 0],
  obliqueBottomLeft: [-1, 1],
  obliqueBottom: [0, 1],
  obliqueBottomRight: [1, 1],
  legacyObliqueTopLeft: [-1, -1],
  legacyObliqueTop: [0, -1],
  legacyObliqueTopRight: [1, -1],
  legacyObliqueLeft: [-1, 0],
  legacyObliqueRight: [1, 0],
  legacyObliqueBottomLeft: [-1, 1],
  legacyObliqueBottom: [0, 1],
  legacyObliqueBottomRight: [1, 1],
};

/**
 * Where the back of a solid `depthPx` deep lands, in the shape's own plane
 * before the camera turns it, and how many copies of the face draw the walls
 * between: the turned depth vector taken back through the turn of the
 * plane, so that once the CSS rotation is applied the copies trail exactly
 * along the depth the camera sees (exact for an orthographic camera). Null
 * when the camera faces the shape (the walls are hidden behind it).
 */
function extrusionStep(
  effects: NonNullable<ReturnType<typeof normalizeOfficePictureEffects>>,
  preset: ReturnType<typeof officePictureEffectPresetEffects> | null,
  depthPx: number,
  frame: OfficeShapeFrame,
): { dx: number; dy: number; count: number } | null {
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  let angles: [number, number, number] | null = null;
  if (rotation) {
    // A legacy perspective camera is a vanishing point, not a slant: its
    // sweep is `legacyPerspectiveSweep`'s.
    if (LEGACY_PERSPECTIVE_VIEW[rotation.camera]) return null;
    const oblique = OBLIQUE_DEPTH[rotation.camera];
    if (oblique && rotation.lon === undefined && rotation.lat === undefined && rotation.rev === undefined) {
      // PowerPoint's oblique projection draws the depth at about 0.39 of its
      // length (measured: 30 pt deep, Oblique: Bottom Right, 7.5 / 9 pt).
      const length = Math.hypot(oblique[0], oblique[1]);
      const slant = rotation.camera.startsWith("legacy") ? LEGACY_OBLIQUE_DEPTH_SCALE : OBLIQUE_DEPTH_SCALE;
      const scale = (depthPx * slant) / length;
      // The camera slants the depth on the SLIDE: a turned or flipped shape
      // draws it in its own frame, so the slant is taken back through the
      
      // back up and right like the level ones).
      const [dx, dy] = slideToShape(frame, oblique[0] * scale, oblique[1] * scale);
      return wallCount(dx, dy);
    }
    const [x, y, z] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
    angles = [
      typeof rotation.lon === "number" ? officeCameraLonTurn(rotation.lon) : x,
      typeof rotation.lat === "number" ? rotation.lat / 60_000 : y,
      typeof rotation.rev === "number" ? officeCameraRevTurn(rotation.rev) : z,
    ];
  } else if (effects.rotation3d === undefined && preset?.camera) {
    angles = [officeCameraLonTurn(preset.camera.lon), preset.camera.lat / 60_000, officeCameraRevTurn(preset.camera.rev)];
  }
  if (!angles) return null;
  // The turn of the layer the walls are drawn in: inside the host's flips.
  angles = flippedCameraTurn(frame, angles);
  const rad = Math.PI / 180;
  // The host's CSS turn: rotateZ(rz) rotateX(-ry) rotateY(rx).
  const beta = angles[0] * rad;
  const alpha = -angles[1] * rad;
  const gamma = angles[2] * rad;
  const [sb, cb, sa, ca, sg, cg] = [Math.sin(beta), Math.cos(beta), Math.sin(alpha), Math.cos(alpha), Math.sin(gamma), Math.cos(gamma)];
  // The depth vector (into the page) on the screen.
  const sx = -depthPx * (sb * cg + cb * sa * sg);
  const sy = -depthPx * (sb * sg - cb * sa * cg);
  if (Math.hypot(sx, sy) < 0.5) return null;
  // The plane's own axes on the screen.
  const ax = cb * cg - sb * sa * sg;
  const ay = cb * sg + sb * sa * cg;
  const bx = -ca * sg;
  const by = ca * cg;
  const det = ax * by - bx * ay;
  // Seen edge on, the plane's turn cannot be undone; the walls are not drawn.
  if (Math.abs(det) < 0.05) return null;
  const lx = (by * sx - bx * sy) / det;
  const ly = (-ay * sx + ax * sy) / det;
  return wallCount(lx, ly);
}















const LEGACY_PERSPECTIVE_VIEW: Readonly<Record<string, readonly [number, number]>> = {
  legacyPerspectiveTopLeft: [-1, -1],
  legacyPerspectiveTop: [0, -1],
  legacyPerspectiveTopRight: [1, -1],
  legacyPerspectiveLeft: [-1, 0],
  legacyPerspectiveFront: [0, 0],
  legacyPerspectiveRight: [1, 0],
  legacyPerspectiveBottomLeft: [-1, 1],
  legacyPerspectiveBottom: [0, 1],
  legacyPerspectiveBottomRight: [1, 1],
};
const LEGACY_PERSPECTIVE_VIEWPOINT_EMU = 1_250_000;
const LEGACY_PERSPECTIVE_EYE_EMU = 8_600_000;

/** A legacy perspective solid's sweep: the vanishing point in the shape's own
 * pixels, how much the back shrinks toward it, and how many copies draw the
 * walls between (about a pixel of movement each, at most 160). */
export interface LegacyPerspectiveSweep {
  readonly vx: number;
  readonly vy: number;
  /** The back's size over the face's (`D / (D + depth)`). */
  readonly back: number;
  readonly count: number;
}

export function legacyPerspectiveSweep(
  effects: NonNullable<ReturnType<typeof normalizeOfficePictureEffects>>,
  depthEmu: number,
  size: { readonly width: number; readonly height: number },
  frame: OfficeShapeFrame,
): LegacyPerspectiveSweep | null {
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  const view = rotation ? LEGACY_PERSPECTIVE_VIEW[rotation.camera] : undefined;
  if (!view || (view[0] === 0 && view[1] === 0) || !(depthEmu > 0)) return null;
  if (rotation!.lon !== undefined || rotation!.lat !== undefined || rotation!.rev !== undefined) return null;
  // The viewpoint stands off the shape's edge on the SLIDE (as the oblique
  // slant does): the edge of the turned shape's box, taken back into the
  // shape's own frame.
  const turn = ((Number.isFinite(frame.rotation) ? (frame.rotation as number) : 0) * Math.PI) / 180;
  const cos = Math.abs(Math.cos(turn));
  const sin = Math.abs(Math.sin(turn));
  const boxWidth = size.width * cos + size.height * sin;
  const boxHeight = size.width * sin + size.height * cos;
  const off = px(LEGACY_PERSPECTIVE_VIEWPOINT_EMU);
  const [ox, oy] = slideToShape(
    frame,
    view[0] * (boxWidth / 2 + off),
    view[1] * (boxHeight / 2 + off),
  );
  const vx = size.width / 2 + ox;
  const vy = size.height / 2 + oy;
  const back = LEGACY_PERSPECTIVE_EYE_EMU / (LEGACY_PERSPECTIVE_EYE_EMU + depthEmu);
  const reach = Math.max(
    ...[
      [0, 0],
      [size.width, 0],
      [0, size.height],
      [size.width, size.height],
    ].map(([x, y]) => Math.hypot(x! - vx, y! - vy)),
  );
  const travel = reach * (1 - back);
  if (travel < 0.5) return null;
  return { vx, vy, back, count: Math.max(1, Math.min(160, Math.ceil(travel))) };
}

/**
 * A perspective camera's walls (Format Shape > 3-D Rotation, a `perspective*`
 * camera): the back is the face moved the depth behind it, which the eye sees
 * smaller. Seen from the eye, a plane parallel to the face lands on the face's
 * plane scaled about the eye's foot there by `h / (h + depth)` (`h` the eye's
 * height over the face), so each copy of the walls is the face scaled about
 * that point before the camera's CSS turn and perspective draw it -- exact
 * for the CSS camera. Copies only offset in the face's plane (the orthographic
 * sweep) came nearer the eye as they went "behind" a face tilted back and
 * drew the walls flaring open; PowerPoint draws them closing toward the back
 * (shape3d-probe p6, Perspective Relaxed 72 pt deep). Null for any other
 * camera, or an eye that does not stand in front of the face.
 */
export function officePerspectiveCameraSweep(
  effects: NonNullable<ReturnType<typeof normalizeOfficePictureEffects>>,
  depthPx: number,
  size: { readonly width: number; readonly height: number },
  frame: OfficeShapeFrame = {},
): LegacyPerspectiveSweep | null {
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  if (!rotation || !rotation.camera.startsWith("perspective") || !(depthPx > 0)) return null;
  const [cx, cy] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
  // The turn of the layer the walls are drawn in: inside the host's flips.
  const [rx, ry] = flippedCameraTurn(frame, [
    typeof rotation.lon === "number" ? officeCameraLonTurn(rotation.lon) : cx,
    typeof rotation.lat === "number" ? rotation.lat / 60_000 : cy,
    0,
  ]);
  const rad = Math.PI / 180;
  // The host's CSS turn rotateZ(rz) rotateX(-ry) rotateY(rx) taken back from
  // the eye at (0, 0, D): the eye in the face's own frame (z toward it). The
  // turn about the view axis moves nothing of it.
  const beta = rx * rad;
  const alpha = -ry * rad;
  const eye = officeCameraEyePx(size, officeCameraFieldOfView(rotation));
  const footX = -Math.sin(beta) * Math.cos(alpha) * eye;
  const footY = Math.sin(alpha) * eye;
  const height = Math.cos(beta) * Math.cos(alpha) * eye;
  // Facing the camera the walls hide behind the face (the orthographic
  // sweep's own rule: the depth seen end on).
  if (!(height > depthPx * 0.01)) return null;
  const turned = Math.hypot(Math.sin(beta), Math.sin(alpha) * Math.cos(beta));
  if (turned * depthPx < 0.5) return null;
  const vx = size.width / 2 + footX;
  const vy = size.height / 2 + footY;
  const back = height / (height + depthPx);
  const reach = Math.max(
    ...[
      [0, 0],
      [size.width, 0],
      [0, size.height],
      [size.width, size.height],
    ].map(([x, y]) => Math.hypot(x! - vx, y! - vy)),
  );
  const travel = reach * (1 - back);
  if (travel < 0.5) return null;
  return { vx, vy, back, count: Math.max(1, Math.min(160, Math.ceil(travel))) };
}

/**
 * A linear ramp over the filter's area as an `feImage` (an SVG gradient, one
 * channel), for a displacement map: 0.5 at `centre`, 0.5 +- span / (2 * half)
 * at the ends. A gradient is interpolated in sRGB, and the filters read it in
 * sRGB (`color-interpolation-filters`), so the channel is linear in place.
 */
function rampImage(
  axis: "x" | "y",
  from: number,
  to: number,
  centre: number,
  half: number,
  across: readonly [number, number],
  result: string,
): OfficeShapeEffectPrimitive {
  const level = (value: number) => Math.round(Math.min(1, Math.max(0, 0.5 + (value - centre) / (2 * half))) * 255);
  const channel = (value: number) => (axis === "x" ? `rgb(${level(value)},0,0)` : `rgb(0,${level(value)},0)`);
  const length = to - from;
  const breadth = across[1] - across[0];
  const [width, height] = axis === "x" ? [length, breadth] : [breadth, length];
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fixedPx(width)}" height="${fixedPx(height)}">` +
    `<linearGradient id="r" x1="0" y1="0" x2="${axis === "x" ? 1 : 0}" y2="${axis === "x" ? 0 : 1}">` +
    `<stop offset="0" stop-color="${channel(from)}"/><stop offset="1" stop-color="${channel(to)}"/>` +
    `</linearGradient><rect width="100%" height="100%" fill="url(#r)"/></svg>`;
  return p("feImage", {
    href: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`,
    x: fixedPx(axis === "x" ? from : across[0]),
    y: fixedPx(axis === "x" ? across[0] : from),
    width: fixedPx(width),
    height: fixedPx(height),
    preserveAspectRatio: "none",
    result,
  });
}

/**
 * The perspective walls: the face's rim (`source`) drawn ever smaller toward
 * the vanishing point, one copy for about each pixel the farthest corner
 * moves. A copy at `t` of the depth is the rim scaled by `D / (D + t depth)`
 * about the point: an `feDisplacementMap` over a ramp that is 0.5 at the
 * point, whose scale grows with `t` (the pull `(x - v) (1 / s - 1)` is
 * linear in the ramp and in `t`). The farthest copy is the back.
 */
function legacyPerspectiveWallPrimitives(
  sweep: LegacyPerspectiveSweep,
  size: { readonly width: number; readonly height: number },
  source: string,
  /** The filter's reach beyond the shape on each side, in shape sizes (at least 1). */
  region = 1,
  /** A ring drawn around the back, under the walls (a Contour's far edge). */
  backRing?: string,
): { primitives: OfficeShapeEffectPrimitive[]; result: string } {
  // The filter's area: the ramps must cover all of it (outside an feImage
  // reads 0, a full pull that drew stray copies of the face).
  const reach = Math.max(1, region);
  const x0 = -reach * size.width;
  const x1 = (1 + reach) * size.width;
  const y0 = -reach * size.height;
  const y1 = (1 + reach) * size.height;
  // The ramps are centred on the shape, not on the point: a far point (a
  // modern perspective camera's eye stands well off the face) spread them so
  // thin that eight bits of the ramp stepped the walls by a pixel. The pull
  // toward the point is split into the part about the centre (the ramps) and
  // a constant shift (an offset after it).
  const cx = size.width / 2;
  const cy = size.height / 2;
  const half = Math.max(Math.abs(x0 - cx), Math.abs(x1 - cx), Math.abs(y0 - cy), Math.abs(y1 - cy)) + 1;
  const out: OfficeShapeEffectPrimitive[] = [
    rampImage("x", x0, x1, cx, half, [y0, y1], "perspectiveRampX"),
    rampImage("y", y0, y1, cy, half, [x0, x1], "perspectiveRampY"),
    p("feComposite", {
      in: "perspectiveRampX",
      in2: "perspectiveRampY",
      operator: "arithmetic",
      k1: 0,
      k2: 1,
      k3: 1,
      k4: 0,
      result: "perspectiveRamp",
    }),
  ];
  // A copy at t of the depth shows the face at `v + (x - v) (1 + c)`, with
  // 1 / (1 + c) its size: c = t (1 / back - 1). About the centre that is the
  // ramp's pull `(x - centre) c`, then a shift of `(v - centre) c / (1 + c)`.
  const pull = 1 / sweep.back - 1;
  const copies: string[] = [];
  for (let index = 1; index <= sweep.count; index += 1) {
    const c = pull * (index / sweep.count);
    const pulled = `perspectivePull${index}`;
    const result = `perspectiveWall${index}`;
    out.push(
      p("feDisplacementMap", {
        in: source,
        in2: "perspectiveRamp",
        scale: fixedPx(2 * half * c),
        xChannelSelector: "R",
        yChannelSelector: "G",
        "color-interpolation-filters": "sRGB",
        result: pulled,
      }),
      p("feOffset", {
        in: pulled,
        dx: fixedPx(((sweep.vx - cx) * c) / (1 + c)),
        dy: fixedPx(((sweep.vy - cy) * c) / (1 + c)),
        result,
      }),
    );
    copies.push(result);
  }
  const rings: string[] = [];
  if (backRing) {
    // The ring at the back: the farthest copy's own pull and shift.
    const c = pull;
    out.push(
      p("feDisplacementMap", {
        in: backRing,
        in2: "perspectiveRamp",
        scale: fixedPx(2 * half * c),
        xChannelSelector: "R",
        yChannelSelector: "G",
        "color-interpolation-filters": "sRGB",
        result: "perspectiveRingPull",
      }),
      p("feOffset", {
        in: "perspectiveRingPull",
        dx: fixedPx(((sweep.vx - cx) * c) / (1 + c)),
        dy: fixedPx(((sweep.vy - cy) * c) / (1 + c)),
        result: "contourBack",
      }),
    );
    rings.push("contourBack");
  }
  out.push(
    p("feMerge", { result: "depthBack" }, [
      ...rings.map((name) => p("feMergeNode", { in: name })),
      ...copies.reverse().map((name) => p("feMergeNode", { in: name })),
      p("feMergeNode", { in: source }),
    ]),
  );
  return { primitives: out, result: "depthBack" };
}

/** The shape's turn and flips on the slide (DrawingML: flip, then turn). */
export interface OfficeShapeFrame {
  /** Clockwise degrees. */
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
}

/**
 * The 3-D turn of the layer a host draws INSIDE the shape's flips, as
 * `[x, y, z]` of the CSS turn `rotateZ(z) rotateX(-y) rotateY(x)`.
 * PowerPoint flips a shape in its own plane first and turns it in 3-D after
 * (text3d-flip-probe.pdf, perspectiveFront X/Y/Z 40/0/0, 0/40/0, 40/20/0 and
 * 320/10/15: a flipH or flipV rectangle is drawn exactly as the plain one --
 * the same edge nearer, the same lean -- its words a half turn for flipV);
 * its 2-D turn comes last, about the projected shape. Every host wears the
 * flips OUTSIDE the turned layer (`rotate(r) scaleX(-1) scaleY(-1)`, then the
 * layer's turn), so the layer takes the turn conjugated by the flips
 * (F·T·F): a horizontal mirror turns the other way about the vertical axis
 * and in the plane (x and z negated), a vertical one about the horizontal
 * axis and in the plane (y and z negated); both flips are a half turn in
 * the plane (x and y negated). A host that draws the flips inside the turn
 * passes none.
 */
export function flippedCameraTurn(
  frame: OfficeShapeFrame | null | undefined,
  [x, y, z]: readonly [number, number, number],
): [number, number, number] {
  const h = frame?.flipH === true;
  const v = frame?.flipV === true;
  const turn = (value: number, negate: boolean) => (negate && value ? -value : value);
  return [turn(x, h), turn(y, v), turn(z, h !== v)];
}

/**
 * A light's or a wall's SVG azimuth (degrees clockwise from +x, y down) on
 * the shape's own plane, as the layer inside the host's flips sees it. The
 * lights stand in the scene, not on the shape's mirror image: flipped in its
 * own plane, a solid is lit as the same solid unflipped (a mirrored
 * rectangle IS the rectangle), so the layer takes them mirrored back.
 */
function flippedAzimuth(frame: OfficeShapeFrame, azimuth: number): number {
  let value = azimuth;
  if (frame.flipH) value = 180 - value;
  if (frame.flipV) value = -value;
  return ((value % 360) + 360) % 360;
}

/** A slide-space vector in the shape's own frame: the turn undone, then the flips. */
function slideToShape(frame: OfficeShapeFrame, x: number, y: number): [number, number] {
  const turn = (-(Number.isFinite(frame.rotation) ? (frame.rotation as number) : 0) * Math.PI) / 180;
  const cos = Math.cos(turn);
  const sin = Math.sin(turn);
  const at = (value: number) => Math.round(value * 1e6) / 1e6 + 0;
  let rx = at(x * cos - y * sin);
  let ry = at(x * sin + y * cos);
  if (frame.flipH) rx = -rx;
  if (frame.flipV) ry = -ry;
  return [rx, ry];
}

/** Where a face coordinate lands on the back: scaled toward the point by the back's size. */
function sweepBack(sweep: LegacyPerspectiveSweep, value: number, centre: number): number {
  return centre + (value - centre) * sweep.back;
}

/** How far past the shape a perspective sweep's back reaches, in shape sizes. */
function perspectiveReach(sweep: LegacyPerspectiveSweep, size: { readonly width: number; readonly height: number }): number {
  const reach = Math.max(
    ...[0, size.width].map((x) => Math.abs(sweepBack(sweep, x, sweep.vx) - x) / Math.max(1, size.width)),
    ...[0, size.height].map((y) => Math.abs(sweepBack(sweep, y, sweep.vy) - y) / Math.max(1, size.height)),
  );
  return Math.round((reach + 0.1) * 100) / 100;
}

/** The walls' sweep: `count` steps of about a pixel (at most 160). */
function wallCount(dx: number, dy: number): { dx: number; dy: number; count: number } | null {
  const length = Math.hypot(dx, dy);
  if (length < 0.5) return null;
  return { dx, dy, count: Math.max(1, Math.min(160, Math.ceil(length))) };
}

/**
 * A member's own effects as drawn inside a group that has a 3-D Rotation
 * (`grpSpPr/a:scene3d`): PowerPoint, Word and Excel turn every member by the
 * GROUP's camera alone, and a member's own camera -- which they keep in the
 * file -- is not drawn (group-scene3d-probe V3: member B's own Perspective
 * Left drawn exactly as its unturned siblings, in all three hosts' PDFs).
 * Before, the member turned twice. Unchanged without a group camera.
 */
export function officeEffectsUnderGroupCamera<T>(effects: T, groupCamera: unknown): T {
  if (!groupCamera || !effects || typeof effects !== "object") return effects;
  const rotation = (effects as { rotation3d?: unknown }).rotation3d;
  if (rotation === undefined || rotation === "none") return effects;
  return { ...(effects as Record<string, unknown>), rotation3d: "none" } as T;
}

export function officeShapeEffectsPaint(
  value: unknown,
  size: { readonly width: number; readonly height: number },
  options: {
    /** Shape Outline > Sketched: the drawn shape wanders as a hand drew it. */
    readonly sketch?: OfficeLineSketch;
    readonly seed?: number;
    /**
     * The shape's own turn and flips, which an oblique camera's slant
     * ignores. The flips are the ones the host wears OUTSIDE the drawn layer:
     * its 3-D turn is taken through them (`flippedCameraTurn`).
     */
    readonly frame?: OfficeShapeFrame;
    /** The shape is filled with a gradient (a 2003 3-D Style face keeps it unshaded). */
    readonly gradientFill?: boolean;
  } = {},
): OfficeShapeEffectsPaint | null {
  const sketch = normalizeOfficeLineSketch(options.sketch);
  const effects = normalizeOfficePictureEffects(value) ?? (sketch ? {} : null);
  if (!effects) return null;
  const preset = isOfficePictureEffectPreset(effects.preset)
    ? officePictureEffectPresetEffects(effects.preset)
    : null;
  const primitives: OfficeShapeEffectPrimitive[] = [];
  const cssFilters: string[] = [];
  let silhouette: OfficeShapeShadowSilhouette | null = null;
  let base = "SourceGraphic";
  let alpha = "SourceAlpha";

  // Sketched: the outline (and the fill it bounds, as Office draws the
  // sketched path) displaced by smooth noise, before every other effect.
  if (sketch) {
    const look = SKETCH_LOOK[sketch];
    primitives.push(
      p("feTurbulence", {
        type: "fractalNoise",
        baseFrequency: look.frequency,
        numOctaves: look.octaves,
        seed: options.seed ?? 1,
        result: "sketchNoise",
      }),
      p("feDisplacementMap", {
        in: base,
        in2: "sketchNoise",
        scale: look.scale * 2,
        xChannelSelector: "R",
        yChannelSelector: "G",
        result: "sketched",
      }),
      // Every effect after it follows the sketched outline, not the true one.
      p("feColorMatrix", {
        in: "sketched",
        type: "matrix",
        values: "0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 1 0",
        result: "sketchedAlpha",
      }),
    );
    base = "sketched";
    alpha = "sketchedAlpha";
  }

  // Soft edges: PowerPoint's fade (`officeSoftEdgeAlpha`; look-fit
  
  // 0.9 of the radius, then blurred a third of the radius wide, multiplied into
  // what the shape draws -- clear at the outline, two thirds at the radius,
  // whole at twice it (it was blurred half the radius and ramped 2a - 1: whole
  // at the radius, and twice too strong at the edge). The fade is taken on
  // the coverage, not the fill's own alpha: a fill at 50 % transparency
  
  // gradient veils, soft-edged, vanished and left the slide dark).
  const softEdge = effects.softEdge && effects.softEdge !== "none" ? effects.softEdge : null;
  if (softEdge && softEdge.radius > 0) {
    primitives.push(
      p("feComponentTransfer", { in: alpha, result: "softCover" }, [
        p("feFuncA", { type: "linear", slope: 255, intercept: 0 }),
      ]),
      p("feMorphology", { in: "softCover", operator: "erode", radius: px(softEdge.radius) * OFFICE_SOFT_EDGE_INSET, result: "softInset" }),
      p("feGaussianBlur", { in: "softInset", stdDeviation: px(softEdge.radius) * OFFICE_SOFT_EDGE_SIGMA, result: "softMask" }),
      p("feComposite", { in: base, in2: "softMask", operator: "in", result: "softened" }),
    );
    base = "softened";
  }

  // Bevel: lit and shaded rims inside the outline (the bands SmartArt's 3-D
  // styles measure), merged over the face.
  const bevel =
    effects.bevel && effects.bevel !== "none"
      ? effects.bevel
      : effects.bevel === undefined && preset && preset.bevel > 0
        ? { preset: "circle" as const, width: preset.bevel, height: preset.bevel }
        : null;
  const format3d = effects.format3d && effects.format3d !== "none" ? effects.format3d : null;

  // Any 3-D format lights the face as its material and light do (PowerPoint
  // draws the face of a bevelled, extruded or contoured shape lit, not in
  // its flat fill).
  if (bevel || (format3d && Object.keys(format3d).length)) {
    primitives.push(
      p("feColorMatrix", {
        in: base,
        type: "matrix",
        values: officeShape3dFaceMatrix(format3d, {
          gradientFill: options.gradientFill,
          tiltWhite: legacySolidLight(format3d) ? 0 : officeShape3dFaceTiltWhite(format3d, cameraTilt(effects, preset)),
        }),
        result: "lit",
      }),
    );
    base = "lit";
  }
  // Contour: a band of its colour on the outline, seen face on, half
  // outside and half over the face (PowerPoint: 6 pt reaches 3 pt each way);
  // Automatic draws black, as PowerPoint draws a shape with no line.
  const contoured = Boolean(format3d?.contour && format3d.contour.width > 0);
  if (format3d?.contour && format3d.contour.width > 0) {
    const half = px(format3d.contour.width) / 2;
    primitives.push(
      p("feMorphology", { in: alpha, operator: "dilate", radius: half, result: "contourOuter" }),
      p("feMorphology", { in: alpha, operator: "erode", radius: half, result: "contourInner" }),
      p("feComposite", { in: "contourOuter", in2: "contourInner", operator: "out", result: "contourBand" }),
      p("feFlood", { "flood-color": format3d.contour.color ?? "#000000", result: "contourInk" }),
      p("feComposite", { in: "contourInk", in2: "contourBand", operator: "in", result: "contourRing" }),
    );
  }
  if (bevel && bevel.width > 0) {
    primitives.push(...officeBevelLightPrimitives(bevel, format3d, base, alpha, size));
    base = "beveled";
  }
  // The contour is its own colour all round, never shaded by the bevel it
  // stands over (shape3d-probe p5: the 2 pt orange contour is as bright
  // along the shaded lower edge of a circle bevel as along the lit top).
  if (contoured) {
    primitives.push(
      p("feMerge", { result: "contoured" }, [p("feMergeNode", { in: base }), p("feMergeNode", { in: "contourRing" })]),
    );
    base = "contoured";
  }

  // Depth (`sp3d@extrusionH`): the solid's side walls, seen once the camera
  // turns it. Facing the camera they are behind the face and not drawn.
  const extrusion = format3d?.depth && format3d.depth.height > 0 ? format3d.depth : null;
  const frame = options.frame ?? {};
  const perspective = extrusion
    ? (legacyPerspectiveSweep(effects, extrusion.height, size, frame) ??
      officePerspectiveCameraSweep(effects, px(extrusion.height), size, frame))
    : null;
  const wall = extrusion && !perspective ? extrusionStep(effects, preset, px(extrusion.height), frame) : null;
  if (extrusion && (wall || perspective)) {
    // The walls take Depth's colour, or (Automatic) the face's own fill
    // shaded, as PowerPoint shades the sides of a solid -- the fill, not
    // the outline: the face's colour is softened inward first, so a thin
    // outline at the edge does not become the walls' colour.
    if (!extrusion.color) {
      primitives.push(
        p("feGaussianBlur", { in: base, stdDeviation: 4, result: "depthBodyBlur" }),
        p("feComponentTransfer", { in: "depthBodyBlur", result: "depthBodyOpaque" }, [p("feFuncA", { type: "linear", slope: 8 })]),
        p("feComposite", { in: "depthBodyOpaque", in2: alpha, operator: "in", result: "depthBody" }),
      );
    }
    primitives.push(
      extrusion.color
        ? p("feFlood", { "flood-color": extrusion.color, result: "depthInk" })
        : p("feColorMatrix", {
            in: "depthBody",
            type: "matrix",
            values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 1 0",
            result: "depthShade",
          }),
    );
    if (extrusion.color) {
      primitives.push(p("feComposite", { in: "depthInk", in2: alpha, operator: "in", result: "depthShade" }));
    }
    // Each wall takes the light of the edge it stands on, as PowerPoint's
    // flat-shaded walls do. Measured on PowerPoint's own render
    
    // under the face's lower edge is the brightest, about the depth colour
    // (Automatic: the face's fill, not darkened) or a little over it; a
    // side wall about half of it; never the near-black a top light left
    // under the face. The face is lit by its outline's facing (the slope of the
    // softened outline, steep only at the edge) and swept; the nearest copy
    // is drawn last, so every wall pixel shows the face's edge pixel it was
    // swept from. (Sweeping a thin rim instead left gaps between copies in
    // WebKit, which offsets by rounded pixels.)
    const legacy = legacySolidLight(format3d);
    if (legacy) {
      // Lit flat (PowerPoint 2003's 3-D Style): each wall one shade of the
      // depth colour by the way it faces on the slide. The face's edge facing
      // that way (its pixels whose neighbour that way is outside) takes the
      // wall's shade, and the sweep carries it along the wall.
      const facings: readonly [WallFacing, number, number][] = [
        ["right", 1, 0],
        ["left", -1, 0],
        ["down", 0, 1],
        ["up", 0, -1],
      ];
      const shade = (name: string, factor: number) =>
        p("feComponentTransfer", { in: "depthShade", result: name }, [
          p("feFuncR", { type: "linear", slope: factor }),
          p("feFuncG", { type: "linear", slope: factor }),
          p("feFuncB", { type: "linear", slope: factor }),
        ]);
      primitives.push(shade("depthEdgeBody", legacy.walls.right));
      const layers = ["depthEdgeBody"];
      for (const [facing, x, y] of facings) {
        const [ux, uy] = slideToShape(frame, x * 2, y * 2);
        primitives.push(
          p("feOffset", { in: alpha, dx: 0 - ux, dy: 0 - uy, result: `depthEdgeNear-${facing}` }),
          p("feComposite", { in: alpha, in2: `depthEdgeNear-${facing}`, operator: "out", result: `depthEdgeBand-${facing}` }),
          shade(`depthEdgeInk-${facing}`, legacy.walls[facing]),
          p("feComposite", { in: `depthEdgeInk-${facing}`, in2: `depthEdgeBand-${facing}`, operator: "in", result: `depthEdge-${facing}` }),
        );
        layers.push(`depthEdge-${facing}`);
      }
      primitives.push(
        p("feMerge", { result: "depthEdges" }, layers.map((layer) => p("feMergeNode", { in: layer }))),
        p("feComposite", { in: "depthEdges", in2: alpha, operator: "in", result: "depthSource" }),
      );
    } else {
      // The lights stand in the solid's own frame, not the camera's: a wall
      // takes the light of the way it faces on the shape, however the camera
      
      // walls of an Automatic depth against the depth colour: the top wall
      // 0.97-1.18 under Oblique Top Right and Contrasting Left, the bottom
      // wall 0.98-1.30 under Isometric Top Up and Relaxed, the left wall
      // 0.86, the right wall 0.16-0.27 -- dark even where it faces the
      // camera, as under Contrasting Left. A light fixed on the slide
      // explains none of that: the one from below this replaced drew the
      // top wall of an oblique solid dark and its right wall half lit).
      const turn = lightTurn(format3d);
      const lifts = officeShape3dWallLifts(format3d, cameraLean(effects, preset));
      const lit = (name: string, azimuth: number, strength: number) =>
        p(
          "feDiffuseLighting",
          {
            in: "depthRelief",
            surfaceScale: 8,
            diffuseConstant: strength,
            "lighting-color": "#ffffff",
            result: name,
          },
          [p("feDistantLight", { azimuth: flippedAzimuth(frame, (azimuth + turn) % 360), elevation: 0 })],
        );
      primitives.push(
        p("feGaussianBlur", { in: alpha, stdDeviation: 1, result: "depthRelief" }),
        ...WALL_LIGHTS.map(([azimuth, strength], index) => lit(`depthLight${index}`, azimuth, strength)),
        p("feComposite", { in: "depthLight0", in2: "depthLight1", operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: "depthLight01" }),
        p("feComposite", { in: "depthLight01", in2: "depthLight2", operator: "arithmetic", k1: 0, k2: 1, k3: 1, k4: 0, result: "depthLight" }),
        p("feComposite", {
          in: "depthShade",
          in2: "depthLight",
          operator: "arithmetic",
          k1: 1,
          k2: WALL_AMBIENT,
          k3: 0,
          k4: 0,
          result: lifts.length ? "depthLitDiffuse" : "depthLitShade",
        }),
        // The walls' lift (`WALL_LIFTS`): an amount of white added to the
        // walls the eye leans toward, whatever the depth colour.
        ...lifts.flatMap(([azimuth, strength], index) => [
          p(
            "feDiffuseLighting",
            { in: "depthRelief", surfaceScale: 8, diffuseConstant: strength, "lighting-color": "#ffffff", result: `depthLift${index}` },
            [p("feDistantLight", { azimuth: flippedAzimuth(frame, azimuth), elevation: 0 })],
          ),
          p("feComposite", {
            in: index === 0 ? "depthLitDiffuse" : `depthLitLift${index - 1}`,
            in2: `depthLift${index}`,
            operator: "arithmetic",
            k1: 0,
            k2: 1,
            k3: 1,
            k4: 0,
            result: index === lifts.length - 1 ? "depthLitShade" : `depthLitLift${index}`,
          }),
        ]),
        p("feComposite", { in: "depthLitShade", in2: alpha, operator: "in", result: "depthSource" }),
      );
    }
    // The walls are the rim swept along the depth: a union of its copies
    // one pixel apart (a browser offsets by whole pixels, so each copy is a
    // pixel of a straight line from the face to the back).
    if (perspective) {
      primitives.push(
        ...legacyPerspectiveWallPrimitives(
          perspective,
          size,
          "depthSource",
          perspectiveReach(perspective, size),
          contoured ? "contourRing" : undefined,
        ).primitives,
      );
    } else if (wall) {
      const step = { dx: wall.dx / wall.count, dy: wall.dy / wall.count };
      const at = (value: number) => Math.round(value * 1000) / 1000;
      const copies = Array.from({ length: wall.count - 1 }, (_, index) =>
        p("feOffset", { in: "depthSource", dx: at(step.dx * (index + 1)), dy: at(step.dy * (index + 1)), result: `depthWall${index}` }),
      );
      const swept = "depthSwept";
      primitives.push(
        ...copies,
        p("feMerge", { result: swept }, [
          ...copies.map((_, index) => p("feMergeNode", { in: `depthWall${copies.length - 1 - index}` })),
          p("feMergeNode", { in: "depthSource" }),
        ]),
        // ...one more step, so the back sits the full depth behind the face.
        p("feOffset", { in: swept, dx: at(step.dx), dy: at(step.dy), result: contoured ? "depthBackWalls" : "depthBack" }),
      );
      if (contoured) {
        // A Contour surrounds the whole solid: its band shows at the back's
        // outline too, outside the walls (shape3d-probe p5: the orange edge
        // under each wall, as at the face).
        primitives.push(
          p("feOffset", { in: "contourRing", dx: at(wall.dx), dy: at(wall.dy), result: "contourBack" }),
          p("feMerge", { result: "depthBack" }, [
            p("feMergeNode", { in: "contourBack" }),
            p("feMergeNode", { in: "depthBackWalls" }),
          ]),
        );
      }
    }
    primitives.push(
      // A browser offsets by whole pixels: the slanted back edge steps, and
      // the turn magnifies the steps. A faint blur draws it straight.
      p("feGaussianBlur", { in: "depthBack", stdDeviation: 0.6, result: "depthWalls" }),
      p("feMerge", { result: "extruded" }, [p("feMergeNode", { in: "depthWalls" }), p("feMergeNode", { in: base })]),
    );
    base = "extruded";
  }

  const shadow =
    effects.shadow && effects.shadow !== "none"
      ? effects.shadow
      : effects.shadow === undefined
        ? (preset?.shadow ?? null)
        : null;
  if (shadow?.inner) {
    // PowerPoint darkens the side the inner shadow is offset toward.
    const radians = (shadow.dir / 60_000) * (Math.PI / 180);
    const distance = px(shadow.dist);
    primitives.push(
      p("feComponentTransfer", { in: alpha, result: "outside" }, [
        p("feFuncA", { type: "table", tableValues: "1 0" }),
      ]),
      p("feOffset", {
        in: "outside",
        dx: Math.round(Math.cos(radians) * distance * 100) / 100,
        dy: Math.round(Math.sin(radians) * distance * 100) / 100,
        result: "innerShifted",
      }),
      p("feGaussianBlur", { in: "innerShifted", stdDeviation: px(shadow.blurRad) / 2, result: "innerSpread" }),
      p("feFlood", { "flood-color": shadow.color, "flood-opacity": shadow.alpha / 100_000, result: "innerInk" }),
      p("feComposite", { in: "innerInk", in2: "innerSpread", operator: "in", result: "innerShade" }),
      p("feComposite", { in: "innerShade", in2: alpha, operator: "in", result: "inner" }),
      p("feMerge", { result: "innerShadowed" }, [p("feMergeNode", { in: base }), p("feMergeNode", { in: "inner" })]),
    );
    base = "innerShadowed";
  } else if (shadow) {
    const css = officeShadowCss(shadow, size);
    if (css?.under) {
      silhouette = {
        color: rgba(shadow.color, shadow.alpha),
        hex: shadow.color,
        transform: String(css.under.transform ?? ""),
        transformOrigin: String(css.under.transformOrigin ?? "50% 100%"),
        filter: css.under.filter ? String(css.under.filter) : "",
      };
    } else {
      cssFilters.push(officeShapeShadowDropShadow(shadow));
    }
  }

  const glow = effects.glow && effects.glow !== "none" ? effects.glow : null;
  if (glow && glow.radius > 0) {
    const blur = Math.max(0.5, px(glow.radius) / 2);
    const color = rgba(glow.color, glow.alpha);
    cssFilters.push(`drop-shadow(0 0 ${blur}px ${color})`, `drop-shadow(0 0 ${blur}px ${color})`);
  }

  const reflection = effects.reflection && effects.reflection !== "none" ? effects.reflection : null;
  const reflect = reflection
    ? `below ${px(reflection.dist)}px linear-gradient(to bottom, transparent ${100 - reflection.endPos / 1000}%, ` +
      `rgb(0 0 0 / ${reflection.endA / 100_000}) ${100 - reflection.endPos / 1000}%, rgb(0 0 0 / ${reflection.stA / 100_000}) 100%)`
    : "";

  let transform = "";
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  const presetCamera = !rotation && effects.rotation3d === undefined ? preset?.camera : undefined;
  if (presetCamera) {
    // Inside the host's flips (`flippedCameraTurn`).
    const [rx, ry, rz] = flippedCameraTurn(frame, [
      officeCameraLonTurn(presetCamera.lon),
      presetCamera.lat / 60_000,
      officeCameraRevTurn(presetCamera.rev),
    ]);
    const fov = presetCamera.fov / 60_000;
    const lean = (degrees: number) => ((degrees + 180) % 360) - 180;
    transform =
      `perspective(${officeCameraEyePx(size, fov)}px) ` +
      `rotateZ(${lean(rz)}deg) rotateX(${-lean(ry)}deg) rotateY(${lean(rx)}deg)`;
  }
  if (rotation) {
    const [x, y, z] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
    const fov = officeCameraFieldOfView(rotation);
    // An explicit rotation (Format Shape > 3-D Rotation's X/Y/Z) overrides
    // the camera's own.
    // Inside the host's flips (`flippedCameraTurn`).
    const [rx, ry, rz] = flippedCameraTurn(frame, [
      typeof rotation.lon === "number" ? officeCameraLonTurn(rotation.lon) : x,
      typeof rotation.lat === "number" ? rotation.lat / 60_000 : y,
      typeof rotation.rev === "number" ? officeCameraRevTurn(rotation.rev) : z,
    ]);
    const perspective = rotation.camera.startsWith("perspective")
      ? `perspective(${officeCameraEyePx(size, fov)}px) `
      : "";
    // 3-D Rotation > Distance from ground (`a:sp3d@z`): the shape stands
    // that far off its ground along its own normal, toward the eye -- nearer,
    // and so larger, through a perspective camera; a parallel camera draws it
    // the same size.
    const format3d = effects.format3d && effects.format3d !== "none" ? effects.format3d : null;
    const lift = perspective && format3d?.z ? ` translateZ(${Math.round((format3d.z / 9525) * 100) / 100}px)` : "";
    if (rx || ry || rz || lift) transform = `${perspective}rotateZ(${rz}deg) rotateX(${-ry}deg) rotateY(${rx}deg)${lift}`;
  }

  // A filter draws its LAST primitive: end on the drawn shape, never on an
  // intermediate (the sketched alpha the effects after it read).
  const last = primitives[primitives.length - 1];
  if (last && last.attrs.result !== base) {
    primitives.push(p("feMerge", { result: "drawn" }, [p("feMergeNode", { in: base })]));
  }
  if (!primitives.length && !cssFilters.length && !silhouette && !reflect && !transform) return null;
  // The walls trail the whole depth from the face: the filter draws that far.
  const reach = wall
    ? Math.max(Math.abs(wall.dx) / Math.max(1, size.width), Math.abs(wall.dy) / Math.max(1, size.height)) + 0.1
    : perspective
      ? perspectiveReach(perspective, size)
      : 0;
  return { primitives, cssFilters, silhouette, reflect, transform, ...(reach > 0.5 ? { region: Math.round(reach * 100) / 100 } : {}) };
}

/**
 * A picture's 3-D Format (its top bevel and `format3d`: bottom bevel, depth,
 * contour, material, lighting), drawn on the picture's own alpha as a
 * shape's is: Office lights a bevelled or extruded picture with the same
 * DrawingML 3-D a shape wears (`p:pic`, `xdr:pic` and `pic:pic` all take
 * `a:sp3d` and `a:scene3d`). The picture's other effects stay the CSS of
 * lib/office-picture-effects; only the 3-D filter comes from here, with the
 * 3-D Rotation it is seen through (the walls' direction) -- the rotation
 * itself is the CSS camera turn. Null when the picture states no 3-D Format.
 */
export function officePictureThreeDPaint(
  value: unknown,
  size: { readonly width: number; readonly height: number },
): Pick<OfficeShapeEffectsPaint, "primitives" | "region"> | null {
  const effects = normalizeOfficePictureEffects(value);
  if (!effects) return null;
  const bevel = effects.bevel && effects.bevel !== "none" && (effects.bevel.width > 0 || effects.bevel.height > 0) ? effects.bevel : undefined;
  const format3d =
    effects.format3d && effects.format3d !== "none" && Object.keys(effects.format3d).length ? effects.format3d : undefined;
  if (!bevel && !format3d) return null;
  const paint = officeShapeEffectsPaint(
    {
      ...(bevel ? { bevel } : {}),
      ...(format3d ? { format3d } : {}),
      ...(effects.rotation3d && effects.rotation3d !== "none" ? { rotation3d: effects.rotation3d } : {}),
    },
    size,
  );
  if (!paint || !paint.primitives.length) return null;
  // A picture's face is lit neutrally (a photo keeps its hues), not as a
  // shape's warm-matte fill.
  const face = pictureFaceMap(format3d ?? null);
  const primitives = paint.primitives.map((primitive) =>
    primitive.tag === "feColorMatrix" && primitive.attrs.result === "lit"
      ? { ...primitive, attrs: { ...primitive.attrs, values: face } }
      : primitive,
  );
  return { primitives, ...(paint.region ? { region: paint.region } : {}) };
}











const PICTURE_FACE: Readonly<Record<string, readonly [number, number]>> = {
  twoPt: [0.979, 6.6],
  threePt: [1.037, 5.3],
  contrasting: [1, 0],
  "matte/twoPt": [0.984, 0],
};

function pictureFaceMap(format3d: OfficeFormat3dEffect | null): string {
  const rig = format3d?.lighting?.rig ?? "threePt";
  const material = format3d?.material ?? "warmMatte";
  const [k, c] = PICTURE_FACE[`${material}/${rig}`] ?? (material === "warmMatte" || material === "plastic" ? PICTURE_FACE[rig] : undefined) ?? [1, 0];
  const offset = fixedPx(c / 255);
  return `${k} 0 0 0 ${offset}  0 ${k} 0 0 ${offset}  0 0 ${k} 0 ${offset}  0 0 0 1 0`;
}

/** The CSS a shape's words wear to turn with it (`officeShapeTextTurn`). */
export interface OfficeShapeTextTurn {
  readonly transform: string;
  readonly transformStyle: "preserve-3d";
}

/**
 * Format Shape > 3-D Rotation > Keep text flat (`keepTextFlat`, the text
 * body's `a:flatTx`): whether a shape's words turn with it. Unchecked, as
 * Office leaves it, the words lie on the shape's front and its 3-D rotation
 * turns them with it: the rotation's own transform, which the host applies
 * to a layer about the SHAPE's centre (the shape's frame turns about it).
 * Checked, or while the words are typed into (PowerPoint lays them flat to
 * edit), they stay flat: null. Null too for a shape its rotation leaves
 * facing the eye.
 */
export function officeShapeTextTurn(
  effects: unknown,
  size: { readonly width: number; readonly height: number },
  options: {
    readonly keepTextFlat?: unknown;
    readonly editing?: boolean;
    /**
     * The flips the host wears OUTSIDE the words' turned layer
     * (`flippedCameraTurn`): the words lie on the shape's front, flipped in
     * its plane before the 3-D turn as the shape is. A host that leaves the
     * words outside the flips passes none.
     */
    readonly frame?: OfficeShapeFrame;
  } = {},
): OfficeShapeTextTurn | null {
  if (options.keepTextFlat === true || options.editing || !effects) return null;
  const transform = officeShapeEffectsPaint(effects, size, options.frame ? { frame: options.frame } : {})?.transform;
  return transform ? { transform, transformStyle: "preserve-3d" } : null;
}
