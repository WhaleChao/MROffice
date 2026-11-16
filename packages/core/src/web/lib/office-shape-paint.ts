/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  isOfficeColorTransform,
  officeColorChannelsFromHex,
  officeColorHexFromChannels,
  officeColorTransformChannels,
} from "./office-color-transform";

/** `a:alpha`/`a:alphaMod`/`a:alphaOff` are opacity, not colour. */
export const OFFICE_ALPHA_TRANSFORMS = ["alpha", "alphaMod", "alphaOff"] as const;
/** The members CT_EG_ColorTransform declares without a `val`. */
const VALUELESS = new Set(["inv", "gray", "comp"]);
const ALPHA = new Set<string>(OFFICE_ALPHA_TRANSFORMS);
const isModelledTransform = (kind: string) =>
  isOfficeColorTransform(kind) || ALPHA.has(kind);

/** ECMA-376 ST_SchemeColorVal. */
const SCHEME_COLORS = new Set([
  "bg1", "tx1", "bg2", "tx2", "dk1", "lt1", "dk2", "lt2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
  "hlink", "folHlink", "phClr",
]);
/** ECMA-376 ST_PathShadeType / ST_TileFlipMode. */
const GRADIENT_PATHS = new Set(["circle", "rect", "shape"]);
const GRADIENT_FLIPS = new Set(["none", "x", "y", "xy"]);
/** PowerPoint's own editor caps a fill at ten stops; the schema does not. */
export const MAX_OFFICE_GRADIENT_STOPS = 16;
const RECT_EDGES = ["l", "t", "r", "b"] as const;

export type OfficeColorSourceKind = "srgb" | "scheme" | "sys" | "prst" | "scrgb";
export type OfficeColorTransformEntry = [string, number | null];

export interface OfficeColorSource {
  kind: OfficeColorSourceKind;
  /** The identity's own `#rrggbb` BEFORE its transforms. */
  base: string;
  val?: string;
  lastColor?: string;
  scrgb?: number[];
  transforms?: OfficeColorTransformEntry[];
}
export interface OfficeGradientRect {
  l?: number;
  t?: number;
  r?: number;
  b?: number;
}
export interface OfficeShapeGradientStop {
  position: number;
  /** What to DRAW: `source.base` through its ordered colour transforms. */
  color: string;
  /** 0..1, from the stop's `a:alpha` family. Absent means fully opaque. */
  alpha?: number;
  /** What to WRITE BACK. Absent for a gradient this product authored. */
  source?: OfficeColorSource;
}
export interface OfficeShapeGradient {
  /**
   * A linear gradient's direction (`a:lin@ang` plus a quarter turn, CSS's 0 =
   * upward). A path gradient (`path`: radial, rectangular or path shading)
   * has none: it states no `a:lin`, so no angle is read, compared or written
   * (`office_shape_paint.gradient`); one a page stored beside it is dropped.
   */
  angle?: number;
  stops: OfficeShapeGradientStop[];
  /** `a:lin@scaled`; absent means the `"1"` this module writes. */
  scaled?: boolean;
  rotWithShape?: boolean;
  flip?: string;
  tileRect?: OfficeGradientRect;
  path?: { kind: string; fillToRect?: OfficeGradientRect };
}
export interface OfficeShapeShadow {
  color: string;
  blur: number;
  distance: number;
  opacity: number;
  /**
   * DrawingML's `a:outerShdw@dir` in degrees -- clockwise from the positive
   * x-axis, y downward. Absent means 90 (straight down), the one direction
   * this product could write before, so an older shadow is byte-identical.
   *
   * Measured in real PowerPoint for Mac (audit SS36.3): `dist="457200"` (36pt)
   * drew the shadow at exactly (dist*cos dir, dist*sin dir) for dir 0, 45, 90
   * and 315 degrees, and `algn` moved it not at all.
   */
  direction?: number;
  /** `a:outerShdw@algn`, carried so a rewrite keeps the file's spelling. */
  align?: string;
}

/** The direction a shadow with no `direction` is drawn at. */
export const OFFICE_SHADOW_DEFAULT_DIRECTION = 90;

/** ST_RectAlignment, the values `a:outerShdw@algn` may take. */
export const OFFICE_SHADOW_ALIGNMENTS = [
  "tl",
  "t",
  "tr",
  "l",
  "ctr",
  "r",
  "bl",
  "b",
  "br",
] as const;

/**
 * Whether this shadow is the one THIS PRODUCT'S writers emit for a style:
 * straight down and centre-aligned. The server twin is
 * `office_shape_paint.shadow_is_straight_down`.
 *
 * It is the only shadow this module could read before the direction was
 * measured (audit §36.3), so it is the only one a host may assume when it
 * hard-codes a NAME for what it read. Every other direction belongs to that
 * host's own named vocabulary -- the sheet's 3x3 Shape Effects gallery calls
 * `dir="2700000"` `outer-bottom-right` and offsets each axis by the LEG, not
 * the hypotenuse the file records; Word's "Shadow" preset writes
 * `dir="2700000" algn="br"`. The general case must not cost those names.
 */
export const officeShadowIsStraightDown = (
  s: OfficeShapeShadow | null | undefined,
): boolean => Boolean(s) && s!.direction === undefined && s!.align === undefined;

/** The shadow's `(dx, dy)` offset, y DOWNWARD. The measured rule, once. */
export function officeShadowOffset(s: OfficeShapeShadow): {
  dx: number;
  dy: number;
} {
  const radians =
    (((s.direction ?? OFFICE_SHADOW_DEFAULT_DIRECTION) % 360) * Math.PI) / 180;
  return {
    dx: Math.round(s.distance * Math.cos(radians) * 10000) / 10000 + 0,
    dy: Math.round(s.distance * Math.sin(radians) * 10000) / 10000 + 0,
  };
}
const hex = (v: unknown): v is string =>
  typeof v === "string" && /^#[0-9a-f]{6}$/i.test(v);
const isInt = (v: unknown, lo: number, hi: number): v is number =>
  typeof v === "number" && Number.isInteger(v) && v >= lo && v <= hi;

function normalizeRect(value: unknown): OfficeGradientRect {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw Error("invalid_shape_gradient");
  const rect = value as Record<string, unknown>;
  for (const key of Object.keys(rect))
    if (!(RECT_EDGES as readonly string[]).includes(key))
      throw Error("invalid_shape_gradient");
  for (const edge of RECT_EDGES)
    if (edge in rect && !isInt(rect[edge], -1_000_000, 1_000_000))
      throw Error("invalid_shape_gradient");
  return rect as OfficeGradientRect;
}

function normalizeTransforms(value: unknown): OfficeColorTransformEntry[] {
  if (!Array.isArray(value) || value.length > 16)
    throw Error("invalid_shape_gradient_color");
  return value.map((item) => {
    if (!Array.isArray(item) || item.length !== 2)
      throw Error("invalid_shape_gradient_color");
    const [kind, raw] = item as [unknown, unknown];
    if (typeof kind !== "string" || !isModelledTransform(kind))
      throw Error("invalid_shape_gradient_color");
    if (VALUELESS.has(kind)) {
      if (raw !== null) throw Error("invalid_shape_gradient_color");
    } else if (
      typeof raw !== "number" ||
      !Number.isFinite(raw) ||
      raw < -21_600_000 ||
      raw > 21_600_000
    )
      throw Error("invalid_shape_gradient_color");
    return [kind, raw] as OfficeColorTransformEntry;
  });
}

/**
 * A stop's colour identity: what the file said, before resolution.
 *
 * `base` is that identity's own colour BEFORE its transforms; the stop's
 * `color` is the same colour AFTER them. Keeping both is what lets a writer
 * tell an untouched stop from an edited one.
 */
export function normalizeOfficeColorSource(value: unknown): OfficeColorSource {
  const error = "invalid_shape_gradient_color";
  if (!value || typeof value !== "object" || Array.isArray(value)) throw Error(error);
  const s = value as Record<string, unknown>;
  for (const key of Object.keys(s))
    if (!["kind", "val", "lastColor", "scrgb", "base", "transforms"].includes(key))
      throw Error(error);
  const kind = s.kind;
  if (
    typeof kind !== "string" ||
    !["srgb", "scheme", "sys", "prst", "scrgb"].includes(kind) ||
    !hex(s.base)
  )
    throw Error(error);
  const result: OfficeColorSource = {
    kind: kind as OfficeColorSourceKind,
    base: s.base.toLowerCase(),
  };
  if (kind === "scrgb") {
    if ("val" in s || !Array.isArray(s.scrgb) || s.scrgb.length !== 3) throw Error(error);
    if (s.scrgb.some((c) => !isInt(c, 0, 100_000))) throw Error(error);
    result.scrgb = [...(s.scrgb as number[])];
  } else {
    if ("scrgb" in s || typeof s.val !== "string") throw Error(error);
    if (kind === "srgb" && !/^[0-9a-fA-F]{6}$/.test(s.val)) throw Error(error);
    if (kind === "scheme" && !SCHEME_COLORS.has(s.val)) throw Error(error);
    // `a:sysClr`/`a:prstClr` carry a long, version-dependent token list; the
    // drawn colour comes from `base`, so bounding the token's shape is enough.
    if (
      (kind === "sys" || kind === "prst") &&
      !/^[A-Za-z][A-Za-z0-9]{0,31}$/.test(s.val)
    )
      throw Error(error);
    result.val = s.val;
  }
  if ("lastColor" in s) {
    if (kind !== "sys" || !hex(s.lastColor)) throw Error(error);
    result.lastColor = s.lastColor.toLowerCase();
  }
  if (s.transforms !== undefined) result.transforms = normalizeTransforms(s.transforms);
  return result;
}

/**
 * The stop's drawn `#rrggbb`: `base` through its ordered transforms, by the one
 * measured rule. The opacity family is skipped; an unmodelled member makes the
 * colour unclaimable and returns null.
 */
export function resolveOfficeStopColor(source: OfficeColorSource): string | null {
  let channels = officeColorChannelsFromHex(source.base);
  if (!channels) return null;
  for (const [kind, raw] of source.transforms ?? []) {
    if (ALPHA.has(kind)) continue;
    const moved = officeColorTransformChannels(channels, kind, raw);
    if (!moved) return null;
    channels = moved;
  }
  return officeColorHexFromChannels(channels);
}

/** The stop's 0..1 opacity from its `a:alpha`/`alphaMod`/`alphaOff`. */
export function officeStopOpacity(
  transforms: readonly OfficeColorTransformEntry[] | undefined,
): number {
  let opacity = 1;
  for (const [kind, raw] of transforms ?? []) {
    if (kind === "alpha") opacity = (raw ?? 0) / 100_000;
    else if (kind === "alphaMod") opacity *= (raw ?? 0) / 100_000;
    else if (kind === "alphaOff") opacity += (raw ?? 0) / 100_000;
  }
  return Math.max(0, Math.min(1, opacity));
}

/** The direction a renderer that draws only linear gradients gives a path
 * gradient, which has none of its own (`a:lin@ang` 0, as a reader of a fill
 * stating none takes it). Never the model's, never written. */
export const OFFICE_PATH_GRADIENT_DRAWN_ANGLE = 90;

/** ``g``'s angle, or `OFFICE_PATH_GRADIENT_DRAWN_ANGLE` for a path gradient. */
export function officeGradientAngle(g: Pick<OfficeShapeGradient, "angle">): number {
  return g.angle ?? OFFICE_PATH_GRADIENT_DRAWN_ANGLE;
}

export function normalizeOfficeShapeGradient(
  value: unknown,
): OfficeShapeGradient | undefined {
  if (value == null) return undefined;
  const g = value as OfficeShapeGradient;
  const fields = ["angle", "stops", "scaled", "rotWithShape", "flip", "tileRect", "path"];
  if (
    typeof g !== "object" ||
    Object.keys(g).some((key) => !fields.includes(key)) ||
    // A path gradient has no angle; one stored beside it is dropped below.
    (!("angle" in g) && !("path" in g)) ||
    !("stops" in g) ||
    (!("path" in g) && (!Number.isFinite(g.angle) || (g.angle as number) < 0 || (g.angle as number) >= 360)) ||
    !Array.isArray(g.stops) ||
    g.stops.length < 2 ||
    g.stops.length > MAX_OFFICE_GRADIENT_STOPS ||
    // Positions need only be non-decreasing: two stops at the same position are
    // DrawingML's (and CSS's) hard colour break, and a gradient is free not to
    // reach the edges of its own box.
    g.stops.some(
      (s, i) =>
        !s ||
        Object.keys(s).some(
          (key) => !["position", "color", "alpha", "source"].includes(key),
        ) ||
        !hex(s.color) ||
        !Number.isFinite(s.position) ||
        s.position < 0 ||
        s.position > 1 ||
        (i > 0 && s.position < g.stops[i - 1].position) ||
        (s.alpha !== undefined &&
          (!Number.isFinite(s.alpha) || s.alpha < 0 || s.alpha > 1)),
    )
  )
    throw Error("invalid_shape_gradient");
  for (const key of ["scaled", "rotWithShape"] as const)
    if (key in g && typeof g[key] !== "boolean") throw Error("invalid_shape_gradient");
  if ("flip" in g && !GRADIENT_FLIPS.has(String(g.flip)))
    throw Error("invalid_shape_gradient");
  if ("path" in g) {
    const p = g.path as Record<string, unknown>;
    if (
      !p ||
      typeof p !== "object" ||
      Object.keys(p).some((key) => !["kind", "fillToRect"].includes(key)) ||
      !GRADIENT_PATHS.has(String(p.kind))
    )
      throw Error("invalid_shape_gradient");
  }
  return {
    ...(g.path === undefined ? { angle: g.angle } : {}),
    stops: g.stops.map((s) => ({
      position: s.position,
      color: s.color.toLowerCase(),
      ...(s.alpha !== undefined ? { alpha: s.alpha } : {}),
      ...(s.source !== undefined
        ? { source: normalizeOfficeColorSource(s.source) }
        : {}),
    })),
    ...(g.scaled !== undefined ? { scaled: g.scaled } : {}),
    ...(g.rotWithShape !== undefined ? { rotWithShape: g.rotWithShape } : {}),
    ...(g.flip !== undefined ? { flip: g.flip } : {}),
    ...(g.tileRect !== undefined ? { tileRect: normalizeRect(g.tileRect) } : {}),
    ...(g.path !== undefined
      ? {
          path: {
            kind: String((g.path as { kind: string }).kind),
            ...((g.path as { fillToRect?: unknown }).fillToRect !== undefined
              ? { fillToRect: normalizeRect((g.path as { fillToRect: unknown }).fillToRect) }
              : {}),
          },
        }
      : {}),
  };
}
/** A shadow's blur and distance are ST_PositiveCoordinate (ECMA-376
 * 20.1.10.42: up to 27273042316900 EMU), and Word and PowerPoint keep and draw
 * any such value as written; the API model (`office_shape_paint.shadow`)
 * holds the same bound, in CSS px. */
export const OFFICE_SHADOW_MAX_BLUR_PX = 27_273_042_316_900 / 9525;
export const OFFICE_SHADOW_MAX_DISTANCE_PX = 27_273_042_316_900 / 9525;

export function normalizeOfficeShapeShadow(
  value: unknown,
): OfficeShapeShadow | undefined {
  if (value == null) return undefined;
  const s = value as OfficeShapeShadow;
  if (
    typeof s !== "object" ||
    Object.keys(s).some(
      (key) =>
        !["blur", "color", "distance", "opacity", "direction", "align"].includes(
          key,
        ),
    ) ||
    !["blur", "color", "distance", "opacity"].every((key) => key in s) ||
    !hex(s.color) ||
    ![s.blur, s.distance, s.opacity].every(Number.isFinite) ||
    s.blur < 0 ||
    s.blur > OFFICE_SHADOW_MAX_BLUR_PX ||
    s.distance < 0 ||
    s.distance > OFFICE_SHADOW_MAX_DISTANCE_PX ||
    s.opacity < 0 ||
    s.opacity > 1 ||
    (s.direction !== undefined &&
      (!Number.isFinite(s.direction) || s.direction < 0 || s.direction > 360)) ||
    (s.align !== undefined &&
      !(OFFICE_SHADOW_ALIGNMENTS as readonly string[]).includes(s.align))
  )
    throw Error("invalid_shape_shadow");
  return { ...s, color: s.color.toLowerCase() };
}

/**
 * The CSS angle this gradient actually draws at inside a `width`×`height` box.
 *
 * `a:lin@scaled="1"` — which is what this module WRITES, and what the corpus's
 * gradients carry — states its angle in the shape's own normalised space, so it
 * comes out aspect-scaled: measured in real PowerPoint, `ang="2700000"` on a
 * 16:3 box drew at 79.380°, and `atan(width/height)` there is 79.380° (§32.2).
 * CSS has no such mode, so drawing one at its literal angle is wrong by up to
 * 45° on a flat shape. Without a box the literal angle is all there is.
 */
export function officeGradientAngleForBox(
  gradient: OfficeShapeGradient | number,
  width?: number,
  height?: number,
): number {
  const angle = typeof gradient === "number" ? gradient : officeGradientAngle(gradient);
  const scaled = typeof gradient === "number" ? true : gradient.scaled !== false;
  if (
    !scaled ||
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    !(width! > 0) ||
    !(height! > 0)
  )
    return ((angle % 360) + 360) % 360;
  // The model's angle is CSS (ppt + 90); the scaling is defined on the ppt one.
  const ppt = ((angle - 90) * Math.PI) / 180;
  const turned = Math.atan2(Math.sin(ppt) * width!, Math.cos(ppt) * height!);
  return (((turned * 180) / Math.PI + 90) % 360 + 360) % 360;
}
































export const OFFICE_EASED_GRADIENT_STEPS = 16;
/** How many runs the eased gradient's first and last run are split into. */
export const OFFICE_EASED_GRADIENT_END_SPLIT = 4;

/** Windows PowerPoint's ease of a two-colour gradient, in linear light. */
export function officeGradientEase(t: number): number {
  return t - ((2 / 3) * Math.sin(2 * Math.PI * t)) / (2 * Math.PI);
}

/** The positions the eased run is laid out at: even runs, the end runs split finer. */
function officeEasedGradientPositions(): number[] {
  const steps = OFFICE_EASED_GRADIENT_STEPS;
  const split = OFFICE_EASED_GRADIENT_END_SPLIT;
  const out: number[] = [];
  for (let index = 0; index < split; index += 1) out.push(index / (steps * split));
  for (let step = 1; step < steps; step += 1) out.push(step / steps);
  for (let index = split - 1; index >= 0; index -= 1) out.push(1 - index / (steps * split));
  return out;
}

const linearLight = (channel: number) => (channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4);
const encodedLight = (value: number) => {
  const v = Math.min(1, Math.max(0, value));
  return v <= 0.0031308 ? 12.92 * v : 1.055 * v ** (1 / 2.4) - 0.055;
};

export function officeGradientDrawnStops<S extends { position: number; color: string; alpha?: number }>(
  stops: readonly S[],
): readonly S[] {
  if (stops.length !== 2) return stops;
  const [first, last] = stops;
  if (Math.abs(first.position) > 1e-9 || Math.abs(last.position - 1) > 1e-9) return stops;
  const from = officeColorChannelsFromHex(first.color);
  const to = officeColorChannelsFromHex(last.color);
  if (!from || !to || from.every((channel, index) => Math.abs(channel - to[index]) < 1e-9)) return stops;
  const fromLight = from.map(linearLight);
  const toLight = to.map(linearLight);
  const alphaStated = first.alpha !== undefined || last.alpha !== undefined;
  const out: S[] = [];
  const positions = officeEasedGradientPositions();
  for (const [index, t] of positions.entries()) {
    const eased = officeGradientEase(t);
    const channels = fromLight.map((value, at) => encodedLight(value + (toLight[at] - value) * eased)) as [number, number, number];
    const { source: _source, ...base } = (index === positions.length - 1 ? last : first) as S & { source?: unknown };
    const stop = { ...base, position: t, color: officeColorHexFromChannels(channels) } as S;
    if (alphaStated) {
      const a = first.alpha ?? 1;
      const b = last.alpha ?? 1;
      (stop as { alpha?: number }).alpha = a + (b - a) * t;
    }
    out.push(stop);
  }
  return out;
}

const stopCss = (s: OfficeShapeGradientStop) => {
  const position = `${s.position * 100}%`;
  if (s.alpha === undefined || s.alpha >= 1) return `${s.color} ${position}`;
  const channels = officeColorChannelsFromHex(s.color) ?? [0, 0, 0];
  const rgb = channels.map((c) => Math.round(c * 255)).join(",");
  return `rgba(${rgb},${Number(s.alpha.toFixed(4))}) ${position}`;
};
















export function officePathGradientFocus(
  rect: { l?: number; t?: number; r?: number; b?: number } | undefined,
): { x: number; y: number } {
  const axis = (near: number | undefined, far: number | undefined) => {
    const a = (rect ? near ?? 0 : 50_000) / 100_000;
    const b = (rect ? far ?? 0 : 50_000) / 100_000;
    return a + b > 1e-9 && a + b < 1 - 1e-9 ? a / (a + b) : a;
  };
  return { x: axis(rect?.l, rect?.r), y: axis(rect?.t, rect?.b) };
}

/**
 * The CSS `background` for a gradient. Pass the shape's box so an aspect-scaled
 * angle can be turned into the angle CSS means (§32.2); an `a:path` gradient
 * becomes the radial gradient it is, instead of a linear stand-in.
 */
export function officeGradientCss(
  g: OfficeShapeGradient,
  width?: number,
  height?: number,
) {
  const stops = officeGradientDrawnStops(g.stops).map(stopCss).join(",");
  if (g.path) {
    // From the focus point PowerPoint starts it at (`officePathGradientFocus`).
    const focus = officePathGradientFocus(g.path.fillToRect);
    const at = ` at ${(focus.x * 100).toFixed(2)}% ${(focus.y * 100).toFixed(2)}%`;
    // DrawingML fills the shape FROM the focus rect outward; CSS reads the
    // first stop at the centre, which is the same order.
    return `radial-gradient(${g.path.kind === "circle" ? "circle" : "ellipse"}${at},${stops})`;
  }
  return `linear-gradient(${officeGradientAngleForBox(g, width, height)}deg,${stops})`;
}
export const officeShadowCss = (s: OfficeShapeShadow) => {
  const { dx, dy } = officeShadowOffset(s);
  return `${dx}px ${dy}px ${s.blur}px ${s.color}${Math.round(s.opacity * 255)
    .toString(16)
    .padStart(2, "0")}`;
};
