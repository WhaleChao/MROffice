/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The words' whole 3-D, as Format Text Effects > 3-D Format and 3-D Rotation
 * edit it (Microsoft Support, "Format Text Effects - 3-D Format pane"): the top
 * and bottom bevels, Depth and Contour with their colours, Material,
 * Lighting and its Angle; X / Y / Z Rotation and Perspective. Shared by every
 * host whose words carry Office's statement (`office-text-paint-stated`):
 * the deck's and the worksheet's text paint (DrawingML, the body's
 * `a:scene3d` / `a:sp3d`) and a Docs run's text effects (Word 2010's
 * `w14:scene3d` / `w14:props3d`, whose camera states its preset alone).
 *
 * An edit of the 3-D is an edit of the statement: the parts it does not name
 * stay as Office stated them (a harsh light, a metal surface, a dark depth
 * colour, a bottom bevel...), and the display fields the paint draws are
 * re-derived from it, so the writers put the statement back verbatim
 * (`stated_is_authentic`, the server's twin of `officeTextPaintStatedIsAuthentic`).
 * A paint without a statement is first stated as its writer writes it.
 */

import {
  OFFICE_BEVEL_PRESETS,
  OFFICE_LIGHT_DIRECTIONS,
  OFFICE_LIGHT_RIGS,
  OFFICE_MATERIAL_PRESETS,
  type OfficeBevelEffect,
  type OfficeBevelPreset,
  type OfficeFormat3dEffect,
  type OfficeLightDirection,
  type OfficeLightRig,
  type OfficeMaterialPreset,
  type OfficeRotation3dEffect,
} from "./office-picture-effects";
import { officeGradientAngle, resolveOfficeStopColor, type OfficeColorSource } from "./office-shape-paint";
import { normalizeOfficeTextPaint, type OfficeTextPaint } from "./office-text-paint";
import {
  normalizeOfficeTextPaintStated,
  officeTextPaintDisplayFromStated,
  type OfficeStatedDisplay,
  type OfficeTextPaintStated,
} from "./office-text-paint-stated";

/** The 3-D kinds of the shared effects model a statement holds. */
export type OfficeText3dKind = "bevel" | "format3d" | "rotation3d";
export interface OfficeText3dValues {
  readonly bevel: OfficeBevelEffect;
  readonly format3d: OfficeFormat3dEffect;
  readonly rotation3d: OfficeRotation3dEffect;
}
/** DrawingML (deck, worksheet) or Word 2010's w14 (Docs runs). */
export type OfficeText3dDialect = "drawing" | "word";

type Scene = NonNullable<OfficeTextPaintStated["scene3d"]>;
type Shape = NonNullable<OfficeTextPaintStated["shape3d"]>;

/** CT_Bevel's defaults: 6 pt by 6 pt, a circle. */
const BEVEL_DEFAULT_EMU = 76_200;
/** The display paint's bevel bound (px), as `text_paint` holds it. */
const DISPLAY_BEVEL_MAX_PX = 16;
const FRONT = "orthographicFront";
/**
 * The light stated with 3-D the words had none of: Three Point, from the top
 * left as the deck's writer states a gallery bevel (`text_paint_xml`), from
 * the top as Word writes a run's (`w14:lightRig w14:dir="t"`).
 */
const DEFAULT_LIGHT = { rig: "threePt", direction: "tl" } as const;
const lightOf = (dialect: OfficeText3dDialect) =>
  dialect === "word" ? { rig: DEFAULT_LIGHT.rig, direction: "t" } : { ...DEFAULT_LIGHT };
const TURN = 21_600_000;

const hexOf = (color: OfficeColorSource | undefined): string | undefined => {
  const hex = color ? resolveOfficeStopColor(color) : null;
  return hex ? hex.toLowerCase() : undefined;
};

function srgb(hex: string, opacity?: number): OfficeColorSource {
  const base = hex.toLowerCase();
  return {
    kind: "srgb",
    base,
    val: base.slice(1).toUpperCase(),
    ...(opacity !== undefined ? { transforms: [["alpha", Math.round(opacity * 100_000)]] } : {}),
  };
}

/** A colour the pane picked: the statement's own source while it still draws it. */
function colorSource(hex: string, held: OfficeColorSource | undefined): OfficeColorSource {
  return held && hexOf(held) === hex.toLowerCase() ? held : srgb(hex);
}

const isBevelPreset = (value: unknown): value is OfficeBevelPreset =>
  typeof value === "string" && (OFFICE_BEVEL_PRESETS as readonly string[]).includes(value);

function bevelEffect(bevel: Shape["bevelTop"]): OfficeBevelEffect | undefined {
  if (!bevel) return undefined;
  const preset = bevel.preset ?? "circle";
  if (!isBevelPreset(preset)) return undefined;
  return { preset, width: bevel.widthEmu ?? BEVEL_DEFAULT_EMU, height: bevel.heightEmu ?? BEVEL_DEFAULT_EMU };
}

const bevelStated = (bevel: OfficeBevelEffect): NonNullable<Shape["bevelTop"]> => ({
  widthEmu: Math.max(0, Math.round(bevel.width)),
  heightEmu: Math.max(0, Math.round(bevel.height)),
  preset: bevel.preset,
});

/**
 * The statement's 3-D in the shared effects model (`office-picture-effects`):
 * its top bevel, its 3-D Format beyond it, and its 3-D Rotation with the
 * camera's own X / Y / Z and field of view.
 */
export function officeTextStated3dEffects(
  stated: OfficeTextPaintStated | null | undefined,
): Partial<OfficeText3dValues> {
  const result: { -readonly [K in OfficeText3dKind]?: OfficeText3dValues[K] } = {};
  const shape = stated?.shape3d;
  const scene = stated?.scene3d;
  const top = bevelEffect(shape?.bevelTop);
  if (top) result.bevel = top;
  const format: { -readonly [K in keyof OfficeFormat3dEffect]: OfficeFormat3dEffect[K] } = {};
  const bottom = bevelEffect(shape?.bevelBottom);
  if (bottom) format.bottom = bottom;
  if (shape?.extrusionHeightEmu) {
    const color = hexOf(shape.extrusionColor);
    format.depth = { height: shape.extrusionHeightEmu, ...(color ? { color } : {}) };
  }
  if (shape?.contourWidthEmu) {
    const color = hexOf(shape.contourColor);
    format.contour = { width: shape.contourWidthEmu, ...(color ? { color } : {}) };
  }
  if (shape?.material && (OFFICE_MATERIAL_PRESETS as readonly string[]).includes(shape.material))
    format.material = shape.material as OfficeMaterialPreset;
  // 3-D Rotation > Distance from ground (`a:sp3d@z`): the words' own, as a
  // shape's (Text Options' 3-D Rotation shows the shape's controls).
  if (shape?.zEmu) format.z = shape.zEmu;
  const rig = scene?.lightRig;
  if (rig && (OFFICE_LIGHT_RIGS as readonly string[]).includes(rig.rig)) {
    format.lighting = {
      rig: rig.rig as OfficeLightRig,
      ...((OFFICE_LIGHT_DIRECTIONS as readonly string[]).includes(rig.direction)
        ? { direction: rig.direction as OfficeLightDirection }
        : {}),
      ...(rig.revolution ? { angle: rig.revolution } : {}),
    };
  }
  if (Object.keys(format).length) result.format3d = format;
  if (scene && (scene.camera !== FRONT || scene.cameraRotation)) {
    result.rotation3d = {
      camera: scene.camera,
      ...(scene.cameraRotation
        ? {
            lat: scene.cameraRotation.latitude,
            lon: scene.cameraRotation.longitude,
            rev: scene.cameraRotation.revolution,
          }
        : {}),
      ...(scene.fieldOfView !== undefined ? { fov: scene.fieldOfView } : {}),
    };
  }
  return result;
}

const turn = (value: number | undefined) => ((Math.round(value ?? 0) % TURN) + TURN) % TURN;

/**
 * The statement with one 3-D pick applied (`value` null: "No Bevel", "No
 * Rotation", Reset of the 3-D Format); everything else it states is kept.
 * A camera preset picked alone replaces the camera and its own rotation,
 * as PowerPoint's 3-D Rotation gallery does; typed angles turn the camera
 * the preset names. Word's w14 camera states its preset alone.
 */
export function withOfficeTextStated3d<K extends OfficeText3dKind>(
  stated: OfficeTextPaintStated | null | undefined,
  kind: K,
  value: OfficeText3dValues[K] | null,
  dialect: OfficeText3dDialect = "drawing",
): OfficeTextPaintStated | null {
  const next: OfficeTextPaintStated = stated ? JSON.parse(JSON.stringify(stated)) : {};
  const shape: Shape = next.shape3d ?? {};
  const scene: Scene = next.scene3d ?? { camera: FRONT, lightRig: lightOf(dialect) };
  if (kind === "bevel") {
    const bevel = value as OfficeBevelEffect | null;
    if (bevel) shape.bevelTop = bevelStated(bevel);
    else delete shape.bevelTop;
  } else if (kind === "format3d") {
    const format = (value as OfficeFormat3dEffect | null) ?? {};
    if (format.bottom) shape.bevelBottom = bevelStated(format.bottom);
    else delete shape.bevelBottom;
    if (format.depth && format.depth.height > 0) {
      shape.extrusionHeightEmu = Math.round(format.depth.height);
      if (format.depth.color) shape.extrusionColor = colorSource(format.depth.color, shape.extrusionColor);
      else delete shape.extrusionColor;
    } else {
      delete shape.extrusionHeightEmu;
      delete shape.extrusionColor;
    }
    if (format.contour && format.contour.width > 0) {
      shape.contourWidthEmu = Math.round(format.contour.width);
      if (format.contour.color) shape.contourColor = colorSource(format.contour.color, shape.contourColor);
      else delete shape.contourColor;
    } else {
      delete shape.contourWidthEmu;
      delete shape.contourColor;
    }
    // Distance from ground: DrawingML's `a:sp3d@z`; Word's w14 states none.
    if (format.z && dialect === "drawing") shape.zEmu = Math.round(format.z);
    else delete shape.zEmu;
    if (format.material && format.material !== "warmMatte") shape.material = format.material;
    else if (format.material === "warmMatte" && shape.material) shape.material = "warmMatte";
    else delete shape.material;
    const lighting = format.lighting;
    const rig: Scene["lightRig"] = {
      ...scene.lightRig,
      rig: lighting?.rig ?? DEFAULT_LIGHT.rig,
      direction: lighting?.direction ?? (lighting ? scene.lightRig.direction : lightOf(dialect).direction),
    };
    if (lighting?.angle !== undefined || rig.latitude !== undefined || rig.longitude !== undefined) {
      rig.revolution = turn(lighting?.angle);
      rig.latitude = rig.latitude ?? 0;
      rig.longitude = rig.longitude ?? 0;
    } else {
      delete rig.revolution;
    }
    if (!lighting) {
      // Reset: Office's own light, unturned.
      delete rig.latitude;
      delete rig.longitude;
      delete rig.revolution;
    }
    scene.lightRig = rig;
  } else {
    const rotation = value as OfficeRotation3dEffect | null;
    delete scene.cameraRotation;
    delete scene.fieldOfView;
    delete scene.zoom;
    if (!rotation) {
      scene.camera = FRONT;
    } else {
      scene.camera = rotation.camera;
      const turned = typeof rotation.lat === "number" || typeof rotation.lon === "number" || typeof rotation.rev === "number";
      if (dialect === "drawing" && turned) {
        scene.cameraRotation = { latitude: turn(rotation.lat), longitude: turn(rotation.lon), revolution: turn(rotation.rev) };
      }
      if (dialect === "drawing" && typeof rotation.fov === "number") scene.fieldOfView = Math.round(rotation.fov);
    }
  }
  if (Object.keys(shape).length) next.shape3d = shape;
  else delete next.shape3d;
  const light = scene.lightRig;
  const plainScene =
    scene.camera === FRONT &&
    !scene.cameraRotation &&
    scene.fieldOfView === undefined &&
    scene.zoom === undefined &&
    light.rig === DEFAULT_LIGHT.rig &&
    (light.direction === "t" || light.direction === "tl") &&
    !light.latitude &&
    !light.longitude &&
    !light.revolution;
  
  // scene that turns and lights nothing of the plain front view is none.
  if (next.shape3d || !plainScene) next.scene3d = scene;
  else delete next.scene3d;
  if (!Object.keys(next).length) return null;
  try {
    return normalizeOfficeTextPaintStated(next) ?? null;
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// A paint and its statement

/** `office_text_paint._px_emu`: a px length as the EMU Office writes. */
function pxEmu(value: number): number {
  const emu = value * 9525;
  const step = Math.round(emu / 127) * 127;
  return Math.abs(emu - step) <= 5 ? step : Math.round(emu);
}

const degrees60k = (dx: number, dy: number) => {
  const degrees = ((((Math.atan2(dy, dx) * 180) / Math.PI) % 360) + 360) % 360;
  return Math.round(degrees * 60_000) % TURN;
};

/**
 * The statement a paint's writer states (`text_paint_xml(..., "drawing")`),
 * or null where the paint holds what a statement does not draw back (a
 * transparent flat fill, a scaled or skewed shadow).
 */
function statementOfPaint(paint: OfficeTextPaint): OfficeTextPaintStated | null {
  if (paint.fillOpacity !== undefined) return null;
  const stated: OfficeTextPaintStated = {};
  const gradient = paint.fillGradient;
  if (gradient) {
    stated.fill = {
      kind: "gradient",
      stops: gradient.stops.map((stop) => ({
        position: Math.round(stop.position * 100_000),
        color: srgb(stop.color, stop.alpha !== undefined && stop.alpha < 1 ? stop.alpha : undefined),
      })),
      ...(gradient.path
        ? { path: gradient.path.fillToRect ? { kind: gradient.path.kind as "circle", fillToRect: { ...gradient.path.fillToRect } } : { kind: gradient.path.kind as "circle" } }
        : { angle: Math.round(((officeGradientAngle(gradient) + 270) % 360) * 60_000) % TURN, scaled: true }),
      rotateWithShape: true,
    };
  } else if (paint.fill === "none") stated.fill = { kind: "none" };
  else if (paint.fillPattern)
    stated.fill = {
      kind: "pattern",
      preset: paint.fillPattern.preset,
      foreground: srgb(paint.fillPattern.foreground),
      background: srgb(paint.fillPattern.background),
    };
  else stated.fill = { kind: "solid", color: srgb(paint.fill) };
  if (paint.outline !== "none" && paint.outlineWidthPt > 0)
    stated.line = { widthEmu: Math.round(paint.outlineWidthPt * 12_700), fill: { kind: "solid", color: srgb(paint.outline) } };
  const effects: NonNullable<OfficeTextPaintStated["effects"]> = {};
  const shadow = paint.shadow;
  if (shadow) {
    if (shadow.scale !== undefined || shadow.scaleY !== undefined || shadow.skewX || shadow.skewY || shadow.align) return null;
    effects.outerShadow = {
      blurEmu: pxEmu(shadow.blur),
      distanceEmu: pxEmu(Math.hypot(shadow.offsetX, shadow.offsetY)),
      direction: degrees60k(shadow.offsetX, shadow.offsetY),
      align: "ctr",
      rotateWithShape: false,
      color: srgb(shadow.color, shadow.opacity),
    };
  }
  const inner = paint.innerShadow;
  if (inner)
    effects.innerShadow = {
      blurEmu: pxEmu(inner.blur),
      distanceEmu: pxEmu(Math.hypot(inner.offsetX, inner.offsetY)),
      direction: degrees60k(inner.offsetX, inner.offsetY),
      color: srgb(inner.color, inner.opacity),
    };
  if (paint.glow) effects.glow = { radiusEmu: pxEmu(paint.glow.radius), color: srgb(paint.glow.color, paint.glow.opacity) };
  const reflection = paint.reflection;
  if (reflection)
    effects.reflection = {
      blurEmu: pxEmu(reflection.blur),
      startAlpha: Math.round(reflection.opacity * 100_000),
      endAlpha: 0,
      endPosition: reflection.size !== undefined ? reflection.size * 1000 : REFLECTION_END_DEFAULT,
      distanceEmu: pxEmu(reflection.distance),
      direction: 5_400_000,
      fadeDirection: 5_400_000,
      scaleY: -100_000,
      align: "bl",
      rotateWithShape: false,
    };
  if (Object.keys(effects).length) stated.effects = effects;
  if (paint.bevel) {
    stated.shape3d = {
      extrusionHeightEmu: pxEmu(paint.bevel.depth),
      material: "warmMatte",
      bevelTop: { widthEmu: pxEmu(paint.bevel.width), heightEmu: pxEmu(paint.bevel.height), preset: paint.bevel.preset ?? "circle" },
    };
  }
  if (paint.bevel || paint.rotation3d)
    stated.scene3d = { camera: paint.rotation3d?.camera ?? FRONT, lightRig: { rig: "threePt", direction: "tl" } };
  try {
    return normalizeOfficeTextPaintStated(stated) ?? null;
  } catch {
    return null;
  }
}

/** The gallery's reflection end (Tight, 40 %): a reflection without a size. */
const REFLECTION_END_DEFAULT = 40_000;

/**
 * What a statement draws, as the server's `display_from_stated(...,
 * fitted=True)` derives it: a bevel or depth past the display's bound at the
 * bound, a reflection's size only off the gallery's default, the front
 * camera no rotation.
 */
function fittedDisplay(stated: OfficeTextPaintStated, fallbackFill: string): OfficeStatedDisplay {
  const display = officeTextPaintDisplayFromStated(stated, fallbackFill);
  if (display.bevel) {
    display.bevel = {
      ...display.bevel,
      width: Math.min(DISPLAY_BEVEL_MAX_PX, display.bevel.width),
      height: Math.min(DISPLAY_BEVEL_MAX_PX, display.bevel.height),
      depth: Math.min(DISPLAY_BEVEL_MAX_PX, display.bevel.depth),
    };
  }
  const end = stated.effects?.reflection?.endPosition;
  if (display.reflection && (end === undefined || end === REFLECTION_END_DEFAULT)) {
    delete (display.reflection as { size?: number }).size;
  }
  if (display.rotation3d?.camera === FRONT) delete display.rotation3d;
  return display;
}

function sameDisplay(a: unknown, b: unknown): boolean {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= 0.01;
  if (typeof a === "string" && typeof b === "string") return a.toLowerCase() === b.toLowerCase();
  if (Array.isArray(a) && Array.isArray(b)) return a.length === b.length && a.every((item, index) => sameDisplay(item, b[index]));
  if (a && b && typeof a === "object" && typeof b === "object") {
    const left = a as Record<string, unknown>;
    const right = b as Record<string, unknown>;
    const keys = Object.keys(left);
    return keys.length === Object.keys(right).length && keys.every((key) => key in right && sameDisplay(left[key], right[key]));
  }
  return a === b;
}

const DISPLAY_KEYS = [
  "fill", "outline", "outlineWidthPt", "fillGradient", "fillPattern",
  "shadow", "innerShadow", "glow", "reflection", "bevel", "rotation3d",
] as const;

/**
 * Whether the paint's statement still draws its display fields: the twin of
 * the server's `stated_is_authentic`, which decides whether the writers put
 * the statement back verbatim.
 */
export function officeTextPaintStatedIsAuthentic(paint: OfficeTextPaint): boolean {
  const stated = paint.stated;
  if (!stated) return false;
  let display: OfficeStatedDisplay;
  try {
    display = fittedDisplay(stated, paint.fill);
  } catch {
    return false;
  }
  const shown = display as unknown as Record<string, unknown>;
  if (!("rotation3d" in paint)) delete shown.rotation3d;
  for (const [key, detail] of [["reflection", "size"], ["bevel", "preset"]] as const) {
    const own = paint[key] as Record<string, unknown> | undefined;
    const theirs = shown[key] as Record<string, unknown> | undefined;
    if (own && !(detail in own) && theirs) {
      const rest = { ...theirs };
      delete rest[detail];
      shown[key] = rest;
    }
  }
  for (const key of DISPLAY_KEYS) {
    if (key === "fill" && !stated.fill && stated.inheritedFill === undefined) continue;
    let mine = paint[key] as unknown;
    const theirs = shown[key];
    if (key === "fillGradient" && paint.fillGradient) {
      const g = paint.fillGradient;
      mine = {
        angle: g.angle,
        stops: g.stops.map((stop) => ({ position: stop.position, color: stop.color })),
        ...(g.path ? { path: g.path } : {}),
      };
    }
    if ((mine === undefined) !== (theirs === undefined)) return false;
    if (mine !== undefined && !sameDisplay(mine, theirs)) return false;
  }
  return true;
}

/**
 * The paint's statement: Office's own while it still draws the paint, else
 * the one its writer would state; null when the paint holds what no
 * statement draws back.
 */
export function officeTextPaintStatement(paint: OfficeTextPaint): OfficeTextPaintStated | null {
  if (paint.stated && officeTextPaintStatedIsAuthentic(paint)) return paint.stated;
  const made = statementOfPaint(paint);
  if (!made) return null;
  // It must draw what the paint draws (an outline that is off has no width).
  const probe: OfficeTextPaint = {
    ...paint,
    outlineWidthPt: paint.outline === "none" ? 0 : paint.outlineWidthPt,
    stated: made,
  };
  return officeTextPaintStatedIsAuthentic(probe) ? made : null;
}

/** The paint a statement draws, its statement beside it. */
function paintOfStatement(paint: OfficeTextPaint, stated: OfficeTextPaintStated): OfficeTextPaint | null {
  try {
    const display = fittedDisplay(stated, paint.fill === "none" ? "#000000" : paint.fill);
    return normalizeOfficeTextPaint({ ...display, stated }) ?? null;
  } catch {
    return null;
  }
}

/**
 * The paint with one 3-D pick applied to its statement, every other part of
 * it kept; null when the paint cannot be stated (the caller keeps its own
 * display edit then).
 */
export function officeTextPaintWith3d<K extends OfficeText3dKind>(
  paint: OfficeTextPaint,
  kind: K,
  value: OfficeText3dValues[K] | null,
): OfficeTextPaint | null {
  const base = officeTextPaintStatement(paint);
  if (!base) return null;
  const stated = withOfficeTextStated3d(base, kind, value, "drawing");
  if (!stated) return null;
  const next = paintOfStatement(paint, stated);
  if (!next || !officeTextPaintStatedIsAuthentic(next)) return null;
  // A paint stated without a fill keeps the colour its words inherit.
  if (!stated.fill && stated.inheritedFill === undefined && next.fill !== paint.fill) {
    return normalizeOfficeTextPaint({ ...next, fill: paint.fill }) ?? null;
  }
  return next;
}

/** The paint's 3-D as the shared effects model: its statement's whole 3-D where it still draws it. */
export function officeTextPaint3dEffects(paint: OfficeTextPaint | null | undefined): Partial<OfficeText3dValues> | null {
  if (!paint?.stated || !officeTextPaintStatedIsAuthentic(paint)) return null;
  return officeTextStated3dEffects(paint.stated);
}
