/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { OFFICE_MATERIAL_PRESETS, type OfficeMaterialPreset } from "../office-picture-effects";
import { officeGradientAngle } from "../office-shape-paint";

/**
 * The fourteen PowerPoint "SmartArt Styles": the catalog a gallery lists (ids,
 * names, sections, the quick-style urn), the 3-D payload types and their
 * persisted guards, and the two paint conversions a canvas draws a stored
 * member with. What each style MEANS -- the colour modifiers, the line, shadow
 * and 3-D it resolves to -- is `./styles`, which the engine host runs and the
 * browser never loads; it re-exports everything here.
 */


/** Gallery order: five "best match for document" styles, then the nine 3-D. */
export type SmartArtStyleId =
  | "simple1"
  | "simple2"
  | "simple3"
  | "simple4"
  | "simple5"
  | "3d1"
  | "3d2"
  | "3d3"
  | "3d4"
  | "3d5"
  | "3d6"
  | "3d7"
  | "3d8"
  | "3d9";

/** The two labelled sections PowerPoint's Design gallery shows. */
export type SmartArtStyleSection = "bestMatch" | "threeD";

export type SmartArtMaterial =
  | "dkEdge"
  | "plastic"
  | "metal"
  | "translucentPowder"
  | "matte";

export type SmartArtBevelPreset =
  | "angle"
  | "relaxedInset"
  | "softRound"
  | "coolSlant";

export type SmartArtLightRigType =
  | "threePt"
  | "flat"
  | "contrasting"
  | "chilly"
  | "balanced"
  | "morning"
  | "soft";

/**
 * The camera presets the fourteen styles actually use. `orthographicFront` is
 * the implicit camera of `simple1`–`3d4` (never serialised as a diagram
 * `scene3d`); the other five are the diagram cameras of `3d5`–`3d9`.
 */
export type SmartArtCameraPreset =
  | "orthographicFront"
  | "isometricOffAxis2Left"
  | "perspectiveRelaxed"
  | "perspectiveRelaxedModerately"
  | "perspectiveLeft"
  | "perspectiveHeroicExtremeRightFacing";

/** `widthEmu` / `heightEmu` are EMU; an absent preset means DrawingML's `circle`. */
export interface SmartArtBevel {
  preset?: SmartArtBevelPreset;
  widthEmu: number;
  heightEmu?: number;
}

/** `revolution` is in 60000ths of a degree, as DrawingML's `a:rot/@rev`. */
export interface SmartArtLightRig {
  rig: SmartArtLightRigType;
  direction: "t";
  revolution?: number;
}

 
export interface SmartArtShape3d {
  material?: SmartArtMaterial;
  bevelTop?: SmartArtBevel;
  bevelBottom?: SmartArtBevel;
  extrusionHeightEmu?: number;
  contourWidthEmu?: number;
  contourColor?: string;
  lightRig?: SmartArtLightRig;
}

/**
 * The diagram camera (`3d5`–`3d9` only). Identical on every member of one
 * SmartArt; the PPTX writer emits it once, on the group.
 * Angles are 60000ths of a degree; `zoom` is thousandths of a percent.
 */
export interface SmartArtScene3d {
  camera: {
    preset: SmartArtCameraPreset;
    zoom?: number;
    /** The camera's own field of view (`a:camera@fov`, 60000ths of a degree,
     *  0..180 degrees): Format Shape > 3-D Rotation > Perspective; absent,
     *  the preset's own. */
    fov?: number;
    latitude?: number;
    longitude?: number;
    revolution?: number;
  };
  lightRig: SmartArtLightRig;
  backdrop?: {
    anchor: [number, number, number];
    norm: [number, number, number];
    up: [number, number, number];
  };
}







export interface SmartArtTextShape3d {
  extrusionHeightEmu: number;
  material: OfficeMaterialPreset;
}

/**
 * Resolved linear gradient, in the CSS degrees every host's gradient field and
 * every writer already speaks: `angle: 180` is top→bottom, which the DrawingML
 * writers turn into `a:lin ang="5400000"` (SMARTART-STYLES.md §2). Stating it
 * in DrawingML degrees here wrote a left→right gradient into every file.
 */
/**
 * A path gradient's shading (`a:path@path`) and focus box (`a:fillToRect`
 * insets, 1/1000 %), as the shared paint states it (`OfficeShapeGradient.path`).
 */
export interface SmartArtGradientPath {
  kind: "circle" | "rect" | "shape";
  fillToRect?: { l?: number; t?: number; r?: number; b?: number };
}

export interface SmartArtFillGradient {
  /** The style tables' ramps are 180; a host theme's own fill style states its own (`theme-effects`). */
  angle: number;
  stops: { position: number; color: string }[];
  /** `a:lin@scaled` and `a:gradFill@rotWithShape` as the host theme's fill style states them. */
  scaled?: boolean;
  rotWithShape?: boolean;
  




  path?: SmartArtGradientPath;
}

/**
 * Resolved drop shadow, in CSS pixels — the unit the shadow field and the
 * DrawingML writers already use (`blur * 9525` EMU). The direction is always
 * straight down.
 */
export interface SmartArtShapeShadow {
  color: string;
  blur: number;
  distance: number;
  opacity: number;
}

/** What one shape member (node or connector) gets from a style. */
export interface SmartArtStyleShapeProps {
  fill: string;
  fillGradient?: SmartArtFillGradient;
  /** `"none"` means the style draws no line. */
  stroke: string;
  /** Points; only present alongside a real line. */
  strokeWidth?: number;
  shapeShadow?: SmartArtShapeShadow;
  shape3d?: SmartArtShape3d;
  scene3d?: SmartArtScene3d;
  smartArtStyle: SmartArtStyleId;
}

/** What one label member gets from a style. Labels never get a plate. */
export interface SmartArtStyleLabelProps {
  color: string;
  textShape3d?: SmartArtTextShape3d;
  scene3d?: SmartArtScene3d;
  smartArtStyle: SmartArtStyleId;
}

export interface ResolvedSmartArtStyle {
  node: SmartArtStyleShapeProps;
  connector: SmartArtStyleShapeProps;
  label: SmartArtStyleLabelProps;
}

export interface SmartArtStyleCatalogEntry {
  id: SmartArtStyleId;
  section: SmartArtStyleSection;
  /** Chinese gallery name, as PowerPoint shows it. */
  zh: string;
  /** English gallery name. */
  en: string;
  /** The DiagramML quick-style urn the deck carries. */
  urn: string;
}

const QUICK_STYLE_2005 = "urn:microsoft.com/office/officeart/2005/8/quickstyle/";
// 3d8 is the one style introduced in the 2009 namespace.
const QUICK_STYLE_2009 = "urn:microsoft.com/office/officeart/2009/2/quickstyle/";

/** The catalog, in gallery order. */
export const SMARTART_STYLES: readonly SmartArtStyleCatalogEntry[] = [
  { id: "simple1", section: "bestMatch", zh: "简单填充", en: "Simple Fill", urn: `${QUICK_STYLE_2005}simple1` },
  { id: "simple2", section: "bestMatch", zh: "白色轮廓", en: "White Outline", urn: `${QUICK_STYLE_2005}simple2` },
  { id: "simple3", section: "bestMatch", zh: "细微效果", en: "Subtle Effect", urn: `${QUICK_STYLE_2005}simple3` },
  { id: "simple4", section: "bestMatch", zh: "中等效果", en: "Moderate Effect", urn: `${QUICK_STYLE_2005}simple4` },
  { id: "simple5", section: "bestMatch", zh: "强烈效果", en: "Intense Effect", urn: `${QUICK_STYLE_2005}simple5` },
  { id: "3d1", section: "threeD", zh: "优雅", en: "Polished", urn: `${QUICK_STYLE_2005}3d1` },
  { id: "3d2", section: "threeD", zh: "嵌入", en: "Inset", urn: `${QUICK_STYLE_2005}3d2` },
  { id: "3d3", section: "threeD", zh: "卡通", en: "Cartoon", urn: `${QUICK_STYLE_2005}3d3` },
  { id: "3d4", section: "threeD", zh: "粉末", en: "Powder", urn: `${QUICK_STYLE_2005}3d4` },
  { id: "3d5", section: "threeD", zh: "砖块场景", en: "Brick Scene", urn: `${QUICK_STYLE_2005}3d5` },
  { id: "3d6", section: "threeD", zh: "鸟瞰场景", en: "Bird's Eye Scene", urn: `${QUICK_STYLE_2005}3d6` },
  { id: "3d7", section: "threeD", zh: "日落场景", en: "Sunset Scene", urn: `${QUICK_STYLE_2005}3d7` },
  { id: "3d8", section: "threeD", zh: "平面场景", en: "Flat Scene", urn: `${QUICK_STYLE_2009}3d8` },
  { id: "3d9", section: "threeD", zh: "金属场景", en: "Metallic Scene", urn: `${QUICK_STYLE_2005}3d9` },
];

export const SMARTART_STYLE_IDS: readonly SmartArtStyleId[] = SMARTART_STYLES.map((entry) => entry.id);

const SMARTART_STYLE_ID_SET: ReadonlySet<string> = new Set<string>(SMARTART_STYLE_IDS);

/** PowerPoint's default for a new SmartArt. */
export const DEFAULT_SMARTART_STYLE_ID: SmartArtStyleId = "simple1";

export function isSmartArtStyleId(value: unknown): value is SmartArtStyleId {
  return typeof value === "string" && SMARTART_STYLE_ID_SET.has(value);
}

/** Narrow an untrusted value to a style id, falling back to the default. */
export function coerceSmartArtStyleId(value: unknown): SmartArtStyleId {
  return isSmartArtStyleId(value) ? value : DEFAULT_SMARTART_STYLE_ID;
}

export function smartArtStyleEntry(id: SmartArtStyleId): SmartArtStyleCatalogEntry {
  return SMARTART_STYLES.find((entry) => entry.id === id) ?? SMARTART_STYLES[0];
}


/**
 * Turn a persisted `fillGradient` into the two companions the presentation
 * canvas needs but never stores: the renderer only switches a gradient on when
 * `fill` AND `fillTo` are set, and it aims the ramp with `gradientAngle` in CSS
 * degrees (0 points up, clockwise). `fillGradient.angle` is in DrawingML
 * degrees, where 0 points right, hence the +90.
 */
export function smartArtGradientCanvasCompanions(
  // Structural, not {@link SmartArtFillGradient}: hosts type their stored
  // gradient with a plain `number` angle, and this works for any of them.
  gradient: {
    // A path gradient has none (`OfficeShapeGradient.angle`).
    angle?: number;
    stops: readonly { position: number; color: string }[];
  },
): { fillTo: string; gradientAngle: number } {
  // The stored angle is already the CSS one every host's gradient field and
  // every writer speaks, so the companions only name the far stop.
  return {
    fillTo: gradient.stops[gradient.stops.length - 1].color,
    gradientAngle: officeGradientAngle(gradient) % 360,
  };
}

/**
 * The same gradient as the host-neutral `{angle, stops}` paint the Docs shape
 * model, the worksheet shape and BOTH DOCX writers already speak
 * (`OfficeShapeGradient`), whose angle is CSS's — 0 points up, clockwise —
 * rather than DrawingML's. `officeGradientXml` / `gradient_xml` turn 180 back
 * into the `a:lin ang="5400000"` of §2, and `officeGradientCss` /
 * `gradientVector` aim the same ramp on the canvas.
 *
 * It is the one conversion, shared with the canvas companions above, so the
 * two can never disagree about which way a style's ramp runs.
 */
export function smartArtGradientOfficePaint(gradient: SmartArtFillGradient): {
  angle: number;
  stops: { position: number; color: string }[];
  scaled?: boolean;
  rotWithShape?: boolean;
  path?: SmartArtGradientPath;
} {
  return {
    angle: smartArtGradientCanvasCompanions(gradient).gradientAngle,
    stops: gradient.stops.map((stop) => ({ ...stop })),
    ...(gradient.scaled !== undefined ? { scaled: gradient.scaled } : {}),
    ...(gradient.rotWithShape !== undefined ? { rotWithShape: gradient.rotWithShape } : {}),
    // A path gradient keeps its shading and focus (the host paint draws and
    // writes `a:path`; the angle beside it is dropped there).
    ...(gradient.path
      ? {
          path: {
            kind: gradient.path.kind,
            ...(gradient.path.fillToRect ? { fillToRect: { ...gradient.path.fillToRect } } : {}),
          },
        }
      : {}),
  };
}

/**
 * The same shadow as the host-neutral `OfficeShapeShadow`. A style already
 * resolves its shadow in CSS PIXELS — what `officeShadowCss` draws and what
 * `officeShadowXml` / `shadow_xml` multiply by 9525 to get back the
 * `blurRad="57150" dist="19050"` of §2 — so this only copies it.
 */
export function smartArtShadowOfficePaint(shadow: SmartArtShapeShadow): {
  color: string;
  blur: number;
  distance: number;
  opacity: number;
} {
  return {
    color: shadow.color,
    blur: shadow.blur,
    distance: shadow.distance,
    opacity: shadow.opacity,
  };
}


/* --------------------------------------------------------- persisted guards */

/**
 * Structural guards for the persisted 3-D payloads, mirroring the bounds the
 * API validator enforces: EMU are integers in 0..9144000 (one slide width),
 * angles are 60000ths of a degree in 0..21600000 (one turn), camera zoom is
 * thousandths of a percent in 1..400000. They exist so the page schema can
 * FAIL CLOSED on a malformed payload rather than passing it through to a writer.
 */
const MAX_EMU = 9144000;
const MAX_ANGLE = 21600000;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isEmu = (value: unknown): boolean =>
  typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= MAX_EMU;

const isAngle = (value: unknown): boolean =>
  typeof value === "number" && Number.isInteger(value) && value >= 0 && value <= MAX_ANGLE;

const isHexColor = (value: unknown): boolean =>
  typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value);

const isOffset = (value: unknown): boolean =>
  typeof value === "number" &&
  Number.isInteger(value) &&
  value >= -MAX_EMU &&
  value <= MAX_EMU;

const isVector3 = (value: unknown): boolean =>
  Array.isArray(value) && value.length === 3 && value.every(isOffset);

const MATERIALS: ReadonlySet<string> = new Set<SmartArtMaterial>([
  "dkEdge", "plastic", "metal", "translucentPowder", "matte",
]);
const BEVEL_PRESETS: ReadonlySet<string> = new Set<SmartArtBevelPreset>([
  "angle", "relaxedInset", "softRound", "coolSlant",
]);
const LIGHT_RIGS: ReadonlySet<string> = new Set<SmartArtLightRigType>([
  "threePt", "flat", "contrasting", "chilly", "balanced", "morning", "soft",
]);
const CAMERA_PRESETS: ReadonlySet<string> = new Set<SmartArtCameraPreset>([
  "orthographicFront", "isometricOffAxis2Left", "perspectiveRelaxed",
  "perspectiveRelaxedModerately", "perspectiveLeft",
  "perspectiveHeroicExtremeRightFacing",
]);

function isBevel(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (!isEmu(value.widthEmu)) return false;
  if (value.heightEmu !== undefined && !isEmu(value.heightEmu)) return false;
  if (value.preset !== undefined && !BEVEL_PRESETS.has(String(value.preset))) return false;
  return true;
}

export function isSmartArtLightRig(value: unknown): value is SmartArtLightRig {
  return isLightRig(value);
}

function isLightRig(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (!LIGHT_RIGS.has(String(value.rig))) return false;
  if (value.direction !== "t") return false;
  if (value.revolution !== undefined && !isAngle(value.revolution)) return false;
  return true;
}

export function isSmartArtShape3d(value: unknown): value is SmartArtShape3d {
  if (!isRecord(value) || Object.keys(value).length === 0) return false;
  const { material, bevelTop, bevelBottom, extrusionHeightEmu, contourWidthEmu, contourColor, lightRig } = value;
  if (material !== undefined && !MATERIALS.has(String(material))) return false;
  if (bevelTop !== undefined && !isBevel(bevelTop)) return false;
  if (bevelBottom !== undefined && !isBevel(bevelBottom)) return false;
  if (extrusionHeightEmu !== undefined && !isEmu(extrusionHeightEmu)) return false;
  if (contourWidthEmu !== undefined && !isEmu(contourWidthEmu)) return false;
  if (contourColor !== undefined && !isHexColor(contourColor)) return false;
  if (lightRig !== undefined && !isLightRig(lightRig)) return false;
  return true;
}

export function isSmartArtScene3d(value: unknown): value is SmartArtScene3d {
  if (!isRecord(value)) return false;
  const camera = value.camera;
  if (!isRecord(camera) || !CAMERA_PRESETS.has(String(camera.preset))) return false;
  if (
    camera.zoom !== undefined &&
    !(typeof camera.zoom === "number" && Number.isInteger(camera.zoom) && camera.zoom >= 1 && camera.zoom <= 400000)
  ) {
    return false;
  }
  if (
    camera.fov !== undefined &&
    !(typeof camera.fov === "number" && Number.isInteger(camera.fov) && camera.fov >= 0 && camera.fov <= 10_800_000)
  ) {
    return false;
  }
  for (const key of ["latitude", "longitude", "revolution"] as const) {
    if (camera[key] !== undefined && !isAngle(camera[key])) return false;
  }
  if (!isLightRig(value.lightRig)) return false;
  const backdrop = value.backdrop;
  if (backdrop !== undefined) {
    if (!isRecord(backdrop)) return false;
    if (!isVector3(backdrop.anchor) || !isVector3(backdrop.norm) || !isVector3(backdrop.up)) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// An ORDINARY shape's own `a:sp3d` / `a:scene3d`, display-only.
//
// The guards above admit exactly what the fourteen SmartArt Styles write: four

// reader (`office_shape_paint`) admits the literal ECMA enumerations, because
// an Office-authored shape may wear any of them — and the sheet's display-only
// `nativeShape3d` slot was held to the SmartArt guards. Corpus S054 (six
// bevelled rectangles: `convex`, `slope`, `hardEdge`, `artDeco`…) then failed
// the page schema, and a schema failure is not "this shape draws flat": the
// WHOLE workbook opened read-only and empty. Seen in a real browser.
//
// A display-only value is bounded here by SHAPE — numbers in range, names that
// are plain identifiers — and the server stays the authority on which names
// exist. The canvas draws a bevel from its width and height and falls back for

// fidelity, never the page.
// ---------------------------------------------------------------------------
const isPresetName = (value: unknown): boolean =>
  typeof value === "string" && /^[A-Za-z][A-Za-z0-9]{0,47}$/.test(value);

const LIGHT_DIRECTIONS: ReadonlySet<string> = new Set([
  "tl", "t", "tr", "l", "r", "bl", "b", "br",
]);

function isNativeBevel(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (!isEmu(value.widthEmu)) return false;
  if (value.heightEmu !== undefined && !isEmu(value.heightEmu)) return false;
  return value.preset === undefined || isPresetName(value.preset);
}

function isNativeLightRig(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (!isPresetName(value.rig)) return false;
  if (!LIGHT_DIRECTIONS.has(String(value.direction))) return false;
  return value.revolution === undefined || isAngle(value.revolution);
}

export function isNativeShape3d(value: unknown): value is SmartArtShape3d {
  if (!isRecord(value) || Object.keys(value).length === 0) return false;
  const { material, bevelTop, bevelBottom, extrusionHeightEmu, contourWidthEmu, contourColor, lightRig } = value;
  if (material !== undefined && !isPresetName(material)) return false;
  if (bevelTop !== undefined && !isNativeBevel(bevelTop)) return false;
  if (bevelBottom !== undefined && !isNativeBevel(bevelBottom)) return false;
  if (extrusionHeightEmu !== undefined && !isEmu(extrusionHeightEmu)) return false;
  if (contourWidthEmu !== undefined && !isEmu(contourWidthEmu)) return false;
  if (contourColor !== undefined && !isHexColor(contourColor)) return false;
  if (lightRig !== undefined && !isNativeLightRig(lightRig)) return false;
  return true;
}

export function isNativeScene3d(value: unknown): value is SmartArtScene3d {
  if (!isRecord(value)) return false;
  const camera = value.camera;
  if (!isRecord(camera) || !isPresetName(camera.preset)) return false;
  if (
    camera.zoom !== undefined &&
    !(typeof camera.zoom === "number" && Number.isInteger(camera.zoom) && camera.zoom >= 1 && camera.zoom <= 400000)
  ) {
    return false;
  }
  if (
    camera.fov !== undefined &&
    !(typeof camera.fov === "number" && Number.isInteger(camera.fov) && camera.fov >= 0 && camera.fov <= 10_800_000)
  ) {
    return false;
  }
  for (const key of ["latitude", "longitude", "revolution"] as const) {
    if (camera[key] !== undefined && !isAngle(camera[key])) return false;
  }
  if (!isNativeLightRig(value.lightRig)) return false;
  const backdrop = value.backdrop;
  if (backdrop !== undefined) {
    if (!isRecord(backdrop)) return false;
    if (!isVector3(backdrop.anchor) || !isVector3(backdrop.norm) || !isVector3(backdrop.up)) return false;
  }
  return true;
}

const TEXT_SHAPE_3D_MATERIALS: ReadonlySet<string> = new Set(OFFICE_MATERIAL_PRESETS);

export function isSmartArtTextShape3d(value: unknown): value is SmartArtTextShape3d {
  return (
    isRecord(value) &&
    isEmu(value.extrusionHeightEmu) &&
    typeof value.material === "string" &&
    TEXT_SHAPE_3D_MATERIALS.has(value.material)
  );
}

/** Read the quick style a persisted member carries, if it carries a valid one. */
export function readSmartArtStyleId(block: unknown): SmartArtStyleId | null {
  if (!isRecord(block)) return null;
  return isSmartArtStyleId(block.smartArtStyle) ? block.smartArtStyle : null;
}

/** Every field a style owns, so a restyle can clear what the new style omits. */
export const SMARTART_STYLE_SHAPE_FIELDS = [
  "fillGradient",
  "stroke",
  "strokeWidth",
  "shapeShadow",
  "shape3d",
  "scene3d",
  "smartArtStyle",
] as const;

export const SMARTART_STYLE_LABEL_FIELDS = [
  "textShape3d",
  "scene3d",
  "smartArtStyle",
] as const;
