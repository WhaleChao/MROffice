/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  normalizeOfficeColorSource,
  officeStopOpacity,
  resolveOfficeStopColor,
  type OfficeColorSource,
} from "./office-shape-paint";
import { OFFICE_CHART_FILL_PATTERNS } from "./office-chart-shape-paint";
import { isOfficeCameraPreset } from "./office-picture-effects";

export type OfficeStatedFill =
  | { kind: "none" }
  | { kind: "solid"; color: OfficeColorSource }
  | {
      kind: "gradient";
      stops: { position: number; color: OfficeColorSource }[];
      angle?: number;
      scaled?: boolean;
      rotateWithShape?: boolean;
      /**
       * A path gradient (DrawingML `a:path`, Word's `w14:path`): it spreads
       * from the focus rectangle `fillToRect` (insets from each edge of the
       * box, in thousandths of a percent, negative past the edge) to the box,
       * as a circle, a rectangle or the shape. It has no linear angle.
       */
      path?: { kind: "circle" | "rect" | "shape"; fillToRect?: { l?: number; t?: number; r?: number; b?: number } };
    }
  | {
      kind: "pattern";
      preset: (typeof OFFICE_CHART_FILL_PATTERNS)[number];
      foreground: OfficeColorSource;
      background: OfficeColorSource;
    };

export interface OfficeStatedShadow {
  blurEmu?: number;
  distanceEmu?: number;
  direction?: number;
  scaleX?: number;
  scaleY?: number;
  skewX?: number;
  skewY?: number;
  align?: string;
  rotateWithShape?: boolean;
  color: OfficeColorSource;
}

export interface OfficeTextPaintStated {
  /** The glyph outline as written; `{}` is a bare `<a:ln/>`. */
  line?: {
    widthEmu?: number;
    cap?: string;
    compound?: string;
    align?: string;
    dash?: string;
    join?: { kind: "round" | "bevel" | "miter"; limit?: number };
    /** A gradient outline is Word's Text Outline > Gradient line. */
    fill?: Extract<OfficeStatedFill, { kind: "none" | "solid" | "gradient" }>;
  };
  /** Absent: the statement names no fill and the text keeps its own colour. */
  fill?: OfficeStatedFill;
  /** Not Office's: the colour a fill-less statement was made over. */
  inheritedFill?: string;
  effects?: {
    glow?: { radiusEmu?: number; color: OfficeColorSource };
    innerShadow?: Omit<OfficeStatedShadow, "align" | "rotateWithShape">;
    outerShadow?: OfficeStatedShadow;
    reflection?: {
      blurEmu?: number;
      startAlpha?: number;
      startPosition?: number;
      endAlpha?: number;
      endPosition?: number;
      distanceEmu?: number;
      direction?: number;
      fadeDirection?: number;
      scaleX?: number;
      scaleY?: number;
      skewX?: number;
      skewY?: number;
      align?: string;
      rotateWithShape?: boolean;
    };
  };
  scene3d?: {
    camera: string;
    /** The camera's own rotation (`a:camera/a:rot`), 60 000ths of a degree. */
    cameraRotation?: { latitude: number; longitude: number; revolution: number };
    /** `a:camera@fov` (60 000ths of a degree) and `@zoom` (1 000ths of a percent). */
    fieldOfView?: number;
    zoom?: number;
    lightRig: {
      rig: string;
      direction: string;
      latitude?: number;
      longitude?: number;
      revolution?: number;
    };
  };
  shape3d?: {
    extrusionHeightEmu?: number;
    contourWidthEmu?: number;
    material?: string;
    bevelTop?: { widthEmu?: number; heightEmu?: number; preset?: string };
    bevelBottom?: { widthEmu?: number; heightEmu?: number; preset?: string };
    extrusionColor?: OfficeColorSource;
    contourColor?: OfficeColorSource;
    /** 3-D Rotation > Distance from ground (`a:sp3d@z`): EMU toward the
     * viewer, -4000..4000 pt; DrawingML only (Word's w14 has none). */
    zEmu?: number;
  };
}

const ERROR = "invalid_office_text_paint_statement";
const TOKEN = /^[A-Za-z][A-Za-z0-9]{0,63}$/;
const HEX = /^#[0-9a-f]{6}$/i;
const PATTERNS = new Set<string>(OFFICE_CHART_FILL_PATTERNS);

function record(
  value: unknown,
  allowed: readonly string[],
  required: readonly string[] = [],
): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw Error(ERROR);
  const keys = Object.keys(value);
  if (keys.some((key) => !allowed.includes(key)) || required.some((key) => !keys.includes(key)))
    throw Error(ERROR);
  return value as Record<string, unknown>;
}

function integer(value: unknown, lo: number, hi: number): number {
  if (typeof value !== "number" || !Number.isInteger(value) || value < lo || value > hi)
    throw Error(ERROR);
  return value;
}

function token(value: unknown): string {
  if (typeof value !== "string" || !TOKEN.test(value)) throw Error(ERROR);
  return value;
}

function bool(value: unknown): boolean {
  if (typeof value !== "boolean") throw Error(ERROR);
  return value;
}

function source(value: unknown): OfficeColorSource {
  let color: OfficeColorSource;
  try {
    color = normalizeOfficeColorSource(value);
  } catch {
    throw Error(ERROR);
  }
  if (resolveOfficeStopColor(color) === null) throw Error(ERROR);
  return color;
}

function optional<T>(
  from: Record<string, unknown>,
  to: Record<string, unknown>,
  key: string,
  read: (value: unknown) => T,
) {
  if (key in from) to[key] = read(from[key]);
}

const MAX_EMU = 12_700_000;
const MAX_ANGLE = 21_599_999;
const PERCENT = 100_000;
/** A focus inset may lie past the box (corpus R78-D02: `t=130000`, `b=-30000`). */
const MAX_FOCUS_INSET = 1_000_000;
const MAX_SCALE = 10_000_000;
const MAX_SKEW = 5_400_000;
const LINE_JOINS = new Set(["round", "bevel", "miter"]);

function fill(value: unknown, line = false): OfficeStatedFill {
  const raw = record(
    value,
    ["kind", "color", "stops", "angle", "scaled", "rotateWithShape", "path", "preset", "foreground", "background"],
    ["kind"],
  );
  if (raw.kind === "none") {
    record(raw, ["kind"]);
    return { kind: "none" };
  }
  if (raw.kind === "solid") {
    record(raw, ["kind", "color"], ["color"]);
    return { kind: "solid", color: source(raw.color) };
  }
  if (raw.kind === "gradient") {
    // Stops stay in file order (Word writes a gallery outline's 70% stop
    // before its 0% stop); drawing sorts them.
    record(raw, ["kind", "stops", "angle", "scaled", "rotateWithShape", "path"], ["stops"]);
    if (!Array.isArray(raw.stops) || raw.stops.length < 2 || raw.stops.length > 16) throw Error(ERROR);
    const stops = raw.stops.map((stop) => {
      const s = record(stop, ["position", "color"], ["position", "color"]);
      return { position: integer(s.position, 0, PERCENT), color: source(s.color) };
    });
    const result: Record<string, unknown> = { kind: "gradient", stops };
    optional(raw, result, "angle", (v) => integer(v, 0, MAX_ANGLE));
    optional(raw, result, "scaled", bool);
    optional(raw, result, "rotateWithShape", bool);
    if (raw.path !== undefined) {
      if (raw.angle !== undefined || raw.scaled !== undefined) throw Error(ERROR);
      const path = record(raw.path, ["kind", "fillToRect"], ["kind"]);
      if (path.kind !== "circle" && path.kind !== "rect" && path.kind !== "shape") throw Error(ERROR);
      const next: Record<string, unknown> = { kind: path.kind };
      if (path.fillToRect !== undefined) {
        const rect = record(path.fillToRect, ["l", "t", "r", "b"]);
        const sides: Record<string, number> = {};
        for (const side of ["l", "t", "r", "b"] as const)
          if (rect[side] !== undefined) sides[side] = integer(rect[side], -MAX_FOCUS_INSET, MAX_FOCUS_INSET);
        next.fillToRect = sides;
      }
      result.path = next;
    }
    return result as OfficeStatedFill;
  }
  if (raw.kind === "pattern" && !line) {
    record(raw, ["kind", "preset", "foreground", "background"], ["kind", "preset", "foreground", "background"]);
    if (typeof raw.preset !== "string" || !PATTERNS.has(raw.preset)) throw Error(ERROR);
    return {
      kind: "pattern",
      preset: raw.preset as (typeof OFFICE_CHART_FILL_PATTERNS)[number],
      foreground: source(raw.foreground),
      background: source(raw.background),
    };
  }
  throw Error(ERROR);
}

function shadow(value: unknown, outer: boolean): OfficeStatedShadow {
  const raw = record(
    value,
    outer
      ? [
          "blurEmu", "distanceEmu", "direction", "scaleX", "scaleY", "skewX", "skewY",
          "align", "rotateWithShape", "color",
        ]
      : ["blurEmu", "distanceEmu", "direction", "color"],
    ["color"],
  );
  const result: Record<string, unknown> = {};
  optional(raw, result, "blurEmu", (v) => integer(v, 0, MAX_EMU));
  optional(raw, result, "distanceEmu", (v) => integer(v, 0, MAX_EMU));
  optional(raw, result, "direction", (v) => integer(v, 0, MAX_ANGLE));
  for (const key of ["scaleX", "scaleY"]) optional(raw, result, key, (v) => integer(v, -MAX_SCALE, MAX_SCALE));
  for (const key of ["skewX", "skewY"]) optional(raw, result, key, (v) => integer(v, -MAX_SKEW, MAX_SKEW));
  optional(raw, result, "align", token);
  optional(raw, result, "rotateWithShape", bool);
  result.color = source(raw.color);
  return result as unknown as OfficeStatedShadow;
}

function bevel(value: unknown) {
  const raw = record(value, ["widthEmu", "heightEmu", "preset"]);
  const result: Record<string, unknown> = {};
  optional(raw, result, "widthEmu", (v) => integer(v, 0, 9_144_000));
  optional(raw, result, "heightEmu", (v) => integer(v, 0, 9_144_000));
  optional(raw, result, "preset", token);
  return result;
}

/**
 * The statement's shape, validated (the server owns the enumerations and
 * refuses a token it does not know). Throws on anything else.
 */
export function normalizeOfficeTextPaintStated(value: unknown): OfficeTextPaintStated | undefined {
  if (value == null) return undefined;
  const raw = record(value, ["line", "fill", "inheritedFill", "effects", "scene3d", "shape3d"]);
  if (!Object.keys(raw).length) throw Error(ERROR);
  const result: OfficeTextPaintStated = {};
  if ("line" in raw) {
    const line = record(raw.line, ["widthEmu", "cap", "compound", "align", "dash", "join", "fill"]);
    const next: Record<string, unknown> = {};
    optional(line, next, "widthEmu", (v) => integer(v, 0, 20_116_800));
    optional(line, next, "cap", token);
    optional(line, next, "compound", token);
    optional(line, next, "align", token);
    optional(line, next, "dash", token);
    optional(line, next, "join", (v) => {
      const join = record(v, ["kind", "limit"], ["kind"]);
      if (typeof join.kind !== "string" || !LINE_JOINS.has(join.kind)) throw Error(ERROR);
      if ("limit" in join && join.kind !== "miter") throw Error(ERROR);
      return {
        kind: join.kind,
        ...("limit" in join ? { limit: integer(join.limit, 0, 100_000_000) } : {}),
      };
    });
    optional(line, next, "fill", (v) => fill(v, true));
    if ((next.fill as OfficeStatedFill | undefined)?.kind === "pattern") throw Error(ERROR);
    result.line = next as OfficeTextPaintStated["line"];
  }
  if ("fill" in raw) result.fill = fill(raw.fill);
  if ("inheritedFill" in raw) {
    if ("fill" in raw || typeof raw.inheritedFill !== "string" || !HEX.test(raw.inheritedFill))
      throw Error(ERROR);
    result.inheritedFill = raw.inheritedFill.toLowerCase();
  }
  if ("effects" in raw) {
    const effects = record(raw.effects, ["glow", "innerShadow", "outerShadow", "reflection"]);
    const next: NonNullable<OfficeTextPaintStated["effects"]> = {};
    if ("glow" in effects) {
      const glow = record(effects.glow, ["radiusEmu", "color"], ["color"]);
      next.glow = {
        ...("radiusEmu" in glow ? { radiusEmu: integer(glow.radiusEmu, 0, MAX_EMU) } : {}),
        color: source(glow.color),
      };
    }
    if ("innerShadow" in effects) next.innerShadow = shadow(effects.innerShadow, false);
    if ("outerShadow" in effects) next.outerShadow = shadow(effects.outerShadow, true);
    if ("reflection" in effects) {
      const r = record(effects.reflection, [
        "blurEmu", "startAlpha", "startPosition", "endAlpha", "endPosition", "distanceEmu",
        "direction", "fadeDirection", "scaleX", "scaleY", "skewX", "skewY", "align",
        "rotateWithShape",
      ]);
      const reflection: Record<string, unknown> = {};
      for (const key of ["blurEmu", "distanceEmu"]) optional(r, reflection, key, (v) => integer(v, 0, MAX_EMU));
      for (const key of ["startAlpha", "startPosition", "endAlpha", "endPosition"])
        optional(r, reflection, key, (v) => integer(v, 0, PERCENT));
      for (const key of ["direction", "fadeDirection"]) optional(r, reflection, key, (v) => integer(v, 0, MAX_ANGLE));
      for (const key of ["scaleX", "scaleY"]) optional(r, reflection, key, (v) => integer(v, -MAX_SCALE, MAX_SCALE));
      for (const key of ["skewX", "skewY"]) optional(r, reflection, key, (v) => integer(v, -MAX_SKEW, MAX_SKEW));
      optional(r, reflection, "align", token);
      optional(r, reflection, "rotateWithShape", bool);
      next.reflection = reflection;
    }
    result.effects = next;
  }
  if ("scene3d" in raw) {
    const scene = record(
      raw.scene3d,
      ["camera", "lightRig", "cameraRotation", "fieldOfView", "zoom"],
      ["camera", "lightRig"],
    );
    const rig = record(
      scene.lightRig,
      ["rig", "direction", "latitude", "longitude", "revolution"],
      ["rig", "direction"],
    );
    const lightRig: Record<string, unknown> = { rig: token(rig.rig), direction: token(rig.direction) };
    for (const key of ["latitude", "longitude", "revolution"])
      optional(rig, lightRig, key, (v) => integer(v, 0, MAX_ANGLE));
    const camera: Record<string, unknown> = {};
    if (scene.cameraRotation !== undefined) {
      const turned = record(
        scene.cameraRotation,
        ["latitude", "longitude", "revolution"],
        ["latitude", "longitude", "revolution"],
      );
      camera.cameraRotation = {
        latitude: integer(turned.latitude, 0, MAX_ANGLE),
        longitude: integer(turned.longitude, 0, MAX_ANGLE),
        revolution: integer(turned.revolution, 0, MAX_ANGLE),
      };
    }
    optional(scene, camera, "fieldOfView", (v) => integer(v, 0, 10_800_000));
    optional(scene, camera, "zoom", (v) => integer(v, 1, 2_147_483_647));
    result.scene3d = {
      camera: token(scene.camera),
      lightRig: lightRig as NonNullable<OfficeTextPaintStated["scene3d"]>["lightRig"],
      ...(camera as Pick<NonNullable<OfficeTextPaintStated["scene3d"]>, "cameraRotation" | "fieldOfView" | "zoom">),
    };
  }
  if ("shape3d" in raw) {
    const shape = record(raw.shape3d, [
      "extrusionHeightEmu", "contourWidthEmu", "material", "bevelTop", "bevelBottom",
      "extrusionColor", "contourColor", "zEmu",
    ]);
    const next: Record<string, unknown> = {};
    optional(shape, next, "zEmu", (v) => integer(v, -50_800_000, 50_800_000));
    optional(shape, next, "extrusionHeightEmu", (v) => integer(v, 0, 9_144_000));
    optional(shape, next, "contourWidthEmu", (v) => integer(v, 0, 9_144_000));
    optional(shape, next, "material", token);
    optional(shape, next, "bevelTop", bevel);
    optional(shape, next, "bevelBottom", bevel);
    optional(shape, next, "extrusionColor", source);
    optional(shape, next, "contourColor", source);
    result.shape3d = next as OfficeTextPaintStated["shape3d"];
  }
  return result;
}

/** What a statement draws, in the shared text paint's display fields. */
export interface OfficeStatedDisplay {
  fill: string;
  outline: string;
  outlineWidthPt: number;
  fillGradient?: {
    /** A path gradient has none (`OfficeShapeGradient.angle`). */
    angle?: number;
    stops: { position: number; color: string }[];
    path?: { kind: string; fillToRect?: { l?: number; t?: number; r?: number; b?: number } };
  };
  fillPattern?: { preset: (typeof OFFICE_CHART_FILL_PATTERNS)[number]; foreground: string; background: string };
  shadow?: { color: string; blur: number; offsetX: number; offsetY: number; opacity: number };
  innerShadow?: { color: string; blur: number; offsetX: number; offsetY: number; opacity: number };
  glow?: { color: string; radius: number; opacity: number };
  reflection?: { opacity: number; blur: number; distance: number };
  bevel?: { width: number; height: number; depth: number; preset?: string };
  rotation3d?: { camera: string };
}

const round = (value: number, digits: number) => {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor + 0;
};

/** A gradient's stops in position order (the statement keeps file order). */
export function sortedStops<T extends { position: number }>(fill: { stops: T[] }): T[] {
  return [...fill.stops].sort((a, b) => a.position - b.position);
}

/** The colour a glyph outline paints in: its solid colour, or a gradient outline's first stop. */
export function officeStatedLineColor(line: OfficeTextPaintStated["line"]): OfficeColorSource | null {
  const fill = line?.fill;
  if (fill?.kind === "solid") return fill.color;
  if (fill?.kind === "gradient") return sortedStops(fill)[0].color;
  return null;
}

function resolved(color: OfficeColorSource): string {
  const hex = resolveOfficeStopColor(color);
  if (!hex) throw Error(ERROR);
  return hex;
}

function shadowDisplay(s: OfficeStatedShadow | Omit<OfficeStatedShadow, "align" | "rotateWithShape">) {
  const distance = (s.distanceEmu ?? 0) / 9525;
  const angle = (((s.direction ?? 0) / 60000) * Math.PI) / 180;
  return {
    color: resolved(s.color),
    blur: round((s.blurEmu ?? 0) / 9525, 3),
    offsetX: round(distance * Math.cos(angle), 3),
    offsetY: round(distance * Math.sin(angle), 3),
    opacity: round(officeStopOpacity(s.color.transforms), 4),
  };
}

/**
 * The display paint a statement draws -- the twin of the server's
 * `display_from_stated`, whose result the writers check the paint against.
 * `fallbackFill` is the colour of text whose statement names no fill.
 */
export function officeTextPaintDisplayFromStated(
  stated: OfficeTextPaintStated,
  fallbackFill: string,
): OfficeStatedDisplay {
  const result: OfficeStatedDisplay = {
    fill: (stated.inheritedFill ?? fallbackFill).toLowerCase(),
    outline: "none",
    outlineWidthPt: 0,
  };
  const f = stated.fill;
  if (f?.kind === "none") result.fill = "none";
  else if (f?.kind === "solid") result.fill = resolved(f.color);
  else if (f?.kind === "pattern") {
    result.fillPattern = {
      preset: f.preset,
      foreground: resolved(f.foreground),
      background: resolved(f.background),
    };
    result.fill = result.fillPattern.foreground;
  } else if (f?.kind === "gradient") {
    const stops = sortedStops(f).map((stop) => ({ position: stop.position / PERCENT, color: resolved(stop.color) }));
    // A path gradient has no angle (`OfficeShapeGradient.angle`).
    result.fillGradient = f.path ? { stops } : { angle: ((f.angle ?? 0) / 60000 + 90) % 360, stops };
    if (f.path) result.fillGradient.path = f.path.fillToRect ? { kind: f.path.kind, fillToRect: { ...f.path.fillToRect } } : { kind: f.path.kind };
    result.fill = stops[0].color;
  }
  const line = stated.line;
  const lineColor = officeStatedLineColor(line);
  if (lineColor) {
    result.outline = resolved(lineColor);
    result.outlineWidthPt = Math.min(12, round((line!.widthEmu ?? 9525) / 12700, 4));
  }
  const effects = stated.effects ?? {};
  if (effects.outerShadow) result.shadow = shadowDisplay(effects.outerShadow);
  if (effects.innerShadow) result.innerShadow = shadowDisplay(effects.innerShadow);
  if (effects.glow) {
    result.glow = {
      color: resolved(effects.glow.color),
      radius: round((effects.glow.radiusEmu ?? 0) / 9525, 3),
      opacity: round(officeStopOpacity(effects.glow.color.transforms), 4),
    };
  }
  if (effects.reflection) {
    result.reflection = {
      opacity: round((effects.reflection.startAlpha ?? PERCENT) / PERCENT, 4),
      // As the importer reads them: a theme change re-derives this display, and a
      
      blur: Math.min(32, round((effects.reflection.blurEmu ?? 0) / 9525, 3)),
      distance: round((effects.reflection.distanceEmu ?? 0) / 9525, 3),
      ...(effects.reflection.endPosition === undefined
        ? {}
        : { size: Math.max(0, Math.min(100, Math.round(effects.reflection.endPosition / 1000))) }),
    };
  }
  if (stated.shape3d) {
    const top = stated.shape3d.bevelTop;
    result.bevel = {
      width: round((top ? top.widthEmu ?? 76_200 : 0) / 9525, 3),
      height: round((top ? top.heightEmu ?? 76_200 : 0) / 9525, 3),
      depth: round((stated.shape3d.extrusionHeightEmu ?? 0) / 9525, 3),
    };
    // The bevel's type, as the importer reads it (CT_Bevel's default `prst`
    // is `circle`, which the display leaves unsaid).
    const preset = top?.preset ?? "circle";
    if (top && preset !== "circle") result.bevel.preset = preset;
  }
  // A 3-D Rotation of the words: a preset camera on the text's scene (No
  // Rotation, `orthographicFront`, is none). Left out, a theme change that
  // re-derives this display read as the words' rotation taken off, which the
  
  // perspectiveHeroicExtremeRightFacing and an Art Deco bevel).
  if (stated.scene3d && isOfficeCameraPreset(stated.scene3d.camera)) {
    result.rotation3d = { camera: stated.scene3d.camera };
  }
  return result;
}

/** A DrawingML scheme colour's slot in a deck's ten-colour theme grid. */
const GRID_SLOT: Readonly<Record<string, string>> = {
  bg1: "bg1", lt1: "bg1", tx1: "tx1", dk1: "tx1", bg2: "bg2", lt2: "bg2", tx2: "tx2", dk2: "tx2",
  accent1: "accent1", accent2: "accent2", accent3: "accent3",
  accent4: "accent4", accent5: "accent5", accent6: "accent6",
};

function rebound(color: OfficeColorSource, theme: Readonly<Record<string, string>>): OfficeColorSource {
  if (color.kind !== "scheme" || !color.val) return color;
  const slot = GRID_SLOT[color.val];
  const base = slot ? theme[slot] : undefined;
  return base && HEX.test(base) ? { ...color, base: base.toLowerCase() } : color;
}

/**
 * The statement under a deck's theme: every theme colour takes that theme's
 * colour as its base (Office's gallery follows the theme; a colour the grid
 * does not name keeps the base it had).
 */
export function officeTextPaintStatedUnderTheme(
  stated: OfficeTextPaintStated,
  theme: Readonly<Record<string, string>>,
): OfficeTextPaintStated {
  const bind = (color: OfficeColorSource) => rebound(color, theme);
  const bindFill = (f: OfficeStatedFill): OfficeStatedFill => {
    if (f.kind === "solid") return { ...f, color: bind(f.color) };
    if (f.kind === "pattern") return { ...f, foreground: bind(f.foreground), background: bind(f.background) };
    if (f.kind === "gradient") return { ...f, stops: f.stops.map((stop) => ({ ...stop, color: bind(stop.color) })) };
    return f;
  };
  const next: OfficeTextPaintStated = { ...stated };
  const lineFill = stated.line?.fill;
  if (stated.line)
    next.line = {
      ...stated.line,
      ...(lineFill ? { fill: bindFill(lineFill) as NonNullable<OfficeTextPaintStated["line"]>["fill"] } : {}),
    };
  if (stated.fill) next.fill = bindFill(stated.fill);
  if (stated.effects) {
    const e = stated.effects;
    next.effects = {
      ...e,
      ...(e.glow ? { glow: { ...e.glow, color: bind(e.glow.color) } } : {}),
      ...(e.innerShadow ? { innerShadow: { ...e.innerShadow, color: bind(e.innerShadow.color) } } : {}),
      ...(e.outerShadow ? { outerShadow: { ...e.outerShadow, color: bind(e.outerShadow.color) } } : {}),
    };
  }
  if (stated.shape3d?.contourColor || stated.shape3d?.extrusionColor)
    next.shape3d = {
      ...stated.shape3d,
      ...(stated.shape3d.contourColor ? { contourColor: bind(stated.shape3d.contourColor) } : {}),
      ...(stated.shape3d.extrusionColor ? { extrusionColor: bind(stated.shape3d.extrusionColor) } : {}),
    };
  return next;
}

/** Two statements equal but for their theme colours' bases (and the colour a fill-less one was made over). */
export function sameOfficeTextPaintStatement(a: unknown, b: unknown): boolean {
  const strip = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(strip);
    if (value && typeof value === "object") {
      return Object.fromEntries(
        Object.entries(value as Record<string, unknown>)
          .filter(([key]) => key !== "base" && key !== "inheritedFill")
          .sort(([x], [y]) => (x < y ? -1 : x > y ? 1 : 0))
          .map(([key, inner]) => [key, strip(inner)]),
      );
    }
    return value;
  };
  return JSON.stringify(strip(a)) === JSON.stringify(strip(b));
}
