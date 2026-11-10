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
  officeCameraEyePx,
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
// PowerPoint for Mac's own rendering of one #4472C4 rectangle per material and

// the face: any 3-D format lights the face (warm matte under three-point

// colour, applied to every colour as the same per-channel change.

/** The face under three-point light, warm matte (any 3-D format). */
const LIT_FACE: readonly [number, number, number] = [1.103, 1.079, 1.061];

type FaceRows = readonly [readonly number[], readonly number[], readonly number[]];
const scaled = (r: number, g: number, b: number): FaceRows => [
  [r, 0, 0, 0, 0],
  [0, g, 0, 0, 0],
  [0, 0, b, 0, 0],
];
const MATERIAL_FACE: Partial<Record<OfficeMaterialPreset, FaceRows>> = {
  matte: scaled(1.029, 1.035, 1.036),
  plastic: scaled(1.059, 1.053, 1.046),
  metal: scaled(1.118, 1.088, 1.056),
  dkEdge: scaled(1.029, 1.035, 1.036),
  softEdge: scaled(1.029, 1.035, 1.036),
  flat: scaled(1, 1, 1),
  softmetal: scaled(1.191, 1.158, 1.117),
  // Powder washes the face almost to white; Clear greys it; Wireframe draws
  // grey (as PowerPoint renders it to a picture).
  powder: [
    [0.22, 0, 0, 0, 0.78],
    [0, 0.12, 0, 0, 0.88],
    [0, 0, 0.05, 0, 0.95],
  ],
  translucentPowder: [
    [0.22, 0, 0, 0, 0.78],
    [0, 0.12, 0, 0, 0.88],
    [0, 0, 0.05, 0, 0.95],
  ],
  clear: [
    [0.106, 0.358, 0.036, 0, 0.396],
    [0.106, 0.358, 0.036, 0, 0.412],
    [0.106, 0.358, 0.036, 0, 0.451],
  ],
  legacyWireframe: [
    [0, 0, 0, 0, 0.651],
    [0, 0, 0, 0, 0.651],
    [0, 0, 0, 0, 0.651],
  ],
};
 
const RIG_FACE: Partial<Record<OfficeLightRig, readonly [number, number, number]>> = {
  balanced: [1.044, 1.009, 0.98],
  soft: [1, 0.982, 0.974],
  harsh: [0.824, 0.842, 0.832],
  flood: [0.853, 0.86, 0.852],
  contrasting: [1, 1, 1],
  morning: [0.853, 0.842, 0.76],
  sunrise: [0.824, 0.789, 0.658],
  sunset: [0.853, 0.684, 0.541],
  chilly: [1, 1, 1.046],
  freezing: [0.676, 0.754, 0.872],
  flat: [1, 1, 1],
  twoPt: [1.059, 1.035, 1.015],
  glow: [1, 1, 1],
  brightRoom: [1, 1, 1],
};
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

/** The face's colour matrix (`feColorMatrix` values) under a 3-D format. */
export function officeShape3dFaceMatrix(
  format3d: OfficeFormat3dEffect | null | undefined,
  options: { readonly gradientFill?: boolean } = {},
): string {
  const legacy = legacySolidLight(format3d);
  if (legacy) {
    
    // reach their stop colours, #33CCCC drawn #34CAC9), a solid one shaded.
    const face = options.gradientFill ? 1 : legacy.face;
    return [...scaled(face, face, face).map((row) => row.join(" ")), "0 0 0 1 0"].join("  ");
  }
  const material = format3d?.material ? MATERIAL_FACE[format3d.material] : undefined;
  const rows = material ?? scaled(...LIT_FACE);
  const rig = format3d?.lighting ? RIG_FACE[format3d.lighting.rig] : undefined;
  const light = rig ?? [1, 1, 1];
  
  const lit = rows.map((row, index) => row.map((value) => Math.round(value * light[index]! * 1000) / 1000));
  return [...lit.map((row) => row.join(" ")), "0 0 0 1 0"].join("  ");
}

/**
 * A bevel's profile as a height field over the shape: steps of its alpha,
 * each eroded `inset` of the bevel's width in from the outline and blurred
 * `sigma` of it, added with `weight` (a negative step falls inward: the edge
 * of a groove or a recessed frame). Fitted by eye to Excel for Mac's
 * rendering of the Office 2007 bevel sample (O15_Bevel-O12-XL-OartEffects,
 * X-155); `circle` (the default) was fitted to its pixels: the lit top
 * band peaks 1.36x the face 2 px in and settles by 6 px for a 6 pt bevel,
 * the bottom falls to half the face (ours 1.31x and 0.4-0.55).
 */
const BEVEL_PROFILE: Readonly<Record<string, readonly (readonly [inset: number, sigma: number, weight: number])[]>> = {
  circle: [[0, 0.33, 1]],
  convex: [[0.4, 0.2, 0.7]],
  angle: [[0.5, 0.25, 0.8]],
  coolSlant: [[0.5, 0.3, 0.8]],
  slope: [[0.125, 0.1, 0.18], [0.375, 0.1, 0.18], [0.625, 0.1, 0.18], [0.875, 0.1, 0.18]],
  softRound: [[0, 0.08, 0.4], [0.5, 0.3, -0.5]],
  relaxedInset: [[0, 0.1, 0.4], [0.8, 0.2, -0.5]],
  hardEdge: [[0, 0.06, 0.3], [0.9, 0.12, -0.45]],
  artDeco: [[0.15, 0.08, 0.3], [0.85, 0.3, 0.6]],
  cross: [[0.3, 0.1, 0.7], [0.8, 0.1, -0.5]],
  divot: [[0.2, 0.1, 0.6], [0.5, 0.1, -0.6], [0.8, 0.1, 0.6]],
  riblet: [[0.17, 0.07, 0.5], [0.5, 0.07, -0.5], [0.83, 0.07, 0.5]],
};
/**
 * The light on a bevel (fitted with the profile above): from the upper left
 * (SVG azimuth 258, clockwise from +x with y down) 40 degrees up, 40 % of a
 * flat face's light not depending on the slope; a facet facing it brightens
 * the face, one facing away darkens it, a flat face keeps its colour.
 */
const BEVEL_LIGHT = { azimuth: 258, elevation: 40, ambient: 0.4, steepness: 0.8 } as const;
const DEFAULT_BEVEL_HEIGHT_PX = 8;









function officeBevelLightPrimitives(
  bevel: { preset?: string; width: number; height?: number },
  format3d: OfficeFormat3dEffect | null,
  face: string,
  alpha: string,
  size: { readonly width: number; readonly height: number },
): OfficeShapeEffectPrimitive[] {
  const widthPx = Math.max(1, Math.min(px(bevel.width), size.width / 4, size.height / 4));
  const heightPx = bevel.height === undefined ? DEFAULT_BEVEL_HEIGHT_PX : px(bevel.height);
  const steps = BEVEL_PROFILE[bevel.preset ?? "circle"] ?? BEVEL_PROFILE.circle!;
  // The steps add up within 0..1 around a 0.5 bias, so no partial sum is
  // clamped; the surface scale undoes the squeeze.
  const span = steps.reduce((total, [, , weight]) => total + Math.abs(weight), 0);
  const squeeze = 1 / (2 * span);
  const out: OfficeShapeEffectPrimitive[] = [];
  steps.forEach(([inset, sigma, weight], index) => {
    const source = inset > 0 ? `bevelErode${index}` : alpha;
    if (inset > 0) {
      out.push(p("feMorphology", { in: alpha, operator: "erode", radius: fixedPx(inset * widthPx), result: source }));
    }
    out.push(p("feGaussianBlur", { in: source, stdDeviation: fixedPx(Math.max(0.5, sigma * widthPx)), result: `bevelStep${index}` }));
    if (index === 0) {
      out.push(
        p("feComponentTransfer", { in: `bevelStep${index}`, result: "bevelHeight0" }, [
          p("feFuncA", { type: "linear", slope: fixedPx(weight * squeeze), intercept: 0.5 }),
        ]),
      );
    } else {
      out.push(
        p("feComposite", {
          in: `bevelHeight${index - 1}`,
          in2: `bevelStep${index}`,
          operator: "arithmetic",
          k1: 0,
          k2: 1,
          k3: fixedPx(weight * squeeze),
          k4: 0,
          result: `bevelHeight${index}`,
        }),
      );
    }
  });
  const height = `bevelHeight${steps.length - 1}`;
  const elevation = BEVEL_LIGHT.elevation;
  const sin = Math.sin((elevation * Math.PI) / 180);
  const gain = (1 - BEVEL_LIGHT.ambient) / sin / 2;
  const azimuth = (((BEVEL_LIGHT.azimuth + lightTurn(format3d)) % 360) + 360) % 360;
  out.push(
    p(
      "feDiffuseLighting",
      {
        in: height,
        surfaceScale: fixedPx((BEVEL_LIGHT.steepness * heightPx) / squeeze),
        diffuseConstant: 1,
        "lighting-color": "#ffffff",
        result: "bevelDiffuse",
      },
      [p("feDistantLight", { azimuth: fixedPx(azimuth), elevation })],
    ),
    // Half scale (a lighting result is capped at 1), doubled on the face.
    p("feColorMatrix", {
      in: "bevelDiffuse",
      type: "matrix",
      values: `${fixedPx(gain)} 0 0 0 ${BEVEL_LIGHT.ambient / 2}  0 ${fixedPx(gain)} 0 0 ${BEVEL_LIGHT.ambient / 2}  0 0 ${fixedPx(gain)} 0 ${BEVEL_LIGHT.ambient / 2}  0 0 0 0 1`,
      result: "bevelShade",
    }),
    // The face made opaque first, so its anti-aliased edge is shaded in its
    // own colour, then given its coverage back by the clip.
    p("feColorMatrix", {
      in: face,
      type: "matrix",
      values: "1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1",
      result: "bevelFace",
    }),
    p("feComposite", { in: "bevelFace", in2: "bevelShade", operator: "arithmetic", k1: 2, k2: 0, k3: 0, k4: 0, result: "bevelLit" }),
    p("feComposite", { in: "bevelLit", in2: face, operator: "in", result: "beveled" }),
  );
  return out;
}

const fixedPx = (value: number) => Math.round(value * 1000) / 1000;

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
      typeof rotation.lon === "number" ? rotation.lon / 60_000 : x,
      typeof rotation.lat === "number" ? rotation.lat / 60_000 : y,
      typeof rotation.rev === "number" ? rotation.rev / 60_000 : z,
    ];
  } else if (effects.rotation3d === undefined && preset?.camera) {
    angles = [preset.camera.lon / 60_000, preset.camera.lat / 60_000, preset.camera.rev / 60_000];
  }
  if (!angles) return null;
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
): LegacyPerspectiveSweep | null {
  const rotation = effects.rotation3d && effects.rotation3d !== "none" ? effects.rotation3d : null;
  if (!rotation || !rotation.camera.startsWith("perspective") || !(depthPx > 0)) return null;
  const [cx, cy] = OFFICE_CAMERA_ROTATIONS[rotation.camera] ?? [0, 0, 0, 45];
  const rx = typeof rotation.lon === "number" ? rotation.lon / 60_000 : cx;
  const ry = typeof rotation.lat === "number" ? rotation.lat / 60_000 : cy;
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

export function officeShapeEffectsPaint(
  value: unknown,
  size: { readonly width: number; readonly height: number },
  options: {
    /** Shape Outline > Sketched: the drawn shape wanders as a hand drew it. */
    readonly sketch?: OfficeLineSketch;
    readonly seed?: number;
    /** The shape's own turn and flips, which an oblique camera's slant ignores. */
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

  // Soft edges: the shape's outline feathered inward over the radius (its
  // coverage blurred, ramped so the edge reaches 0 and the inside stays
  // whole), multiplied into what the shape draws. The ramp is taken on the
  // coverage, not the fill's own alpha: a fill at 50 % transparency ramped
  
  // gradient veils, soft-edged, vanished and left the slide dark).
  const softEdge = effects.softEdge && effects.softEdge !== "none" ? effects.softEdge : null;
  if (softEdge && softEdge.radius > 0) {
    primitives.push(
      p("feComponentTransfer", { in: alpha, result: "softCover" }, [
        p("feFuncA", { type: "linear", slope: 255, intercept: 0 }),
      ]),
      p("feGaussianBlur", { in: "softCover", stdDeviation: px(softEdge.radius) / 2, result: "softBlur" }),
      p("feComponentTransfer", { in: "softBlur", result: "softMask" }, [
        p("feFuncA", { type: "linear", slope: 2, intercept: -1 }),
      ]),
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
        values: officeShape3dFaceMatrix(format3d, { gradientFill: options.gradientFill }),
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
      officePerspectiveCameraSweep(effects, px(extrusion.height), size))
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
      primitives.push(
        p("feGaussianBlur", { in: alpha, stdDeviation: 1, result: "depthRelief" }),
        p(
          "feDiffuseLighting",
          { in: "depthRelief", surfaceScale: 8, diffuseConstant: 1, "lighting-color": "#ffffff", result: "depthLight" },
          [p("feDistantLight", { azimuth: (90 + lightTurn(format3d)) % 360, elevation: 30 })],
        ),
        p("feComposite", { in: "depthShade", in2: "depthLight", operator: "arithmetic", k1: 0.8, k2: 0.4, k3: 0, k4: 0, result: "depthLitShade" }),
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
    const rx = presetCamera.lon / 60_000;
    const ry = presetCamera.lat / 60_000;
    const rz = presetCamera.rev / 60_000;
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
    const rx = typeof rotation.lon === "number" ? rotation.lon / 60_000 : x;
    const ry = typeof rotation.lat === "number" ? rotation.lat / 60_000 : y;
    const rz = typeof rotation.rev === "number" ? rotation.rev / 60_000 : z;
    const perspective = rotation.camera.startsWith("perspective")
      ? `perspective(${officeCameraEyePx(size, fov)}px) `
      : "";
    if (rx || ry || rz) transform = `${perspective}rotateZ(${rz}deg) rotateX(${-ry}deg) rotateY(${rx}deg)`;
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
  return { primitives: paint.primitives, ...(paint.region ? { region: paint.region } : {}) };
}
