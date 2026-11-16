/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "./office-units";
import { OFFICE_BEVEL_PRESETS, OFFICE_CAMERA_ROTATIONS, officeCameraLonTurn, officeCameraRevTurn } from "./office-picture-effects";
import type { CSSProperties } from "react";
import {
  normalizeOfficeShapeGradient,
  officeGradientCss,
  type OfficeShapeGradient, officeGradientAngle, officeGradientDrawnStops } from "./office-shape-paint";
import {
  normalizeOfficeChartFillPattern,
  officeFillPatternBackground,
  type OfficeChartFillPattern,
} from "./office-chart-shape-paint";
import {
  normalizeOfficeTextPaintStated,
  type OfficeTextPaintStated,
} from "./office-text-paint-stated";
import { cssFontFamilyList } from "./css-font-family";
import {
  OFFICE_TEXT_OUTLINE_CLEARED,
  OFFICE_TEXT_OUTLINE_RINGS,
  officeTextOutlineLineOfStated,
  officeTextOutlineMarks,
  officeTextOutlineStroke,
  officeTextOutlineSvgGlyphs,
  type OfficeTextOutlineStroke,
} from "./office-text-outline/office-text-outline-css";

/** `ST_RectAlignment` without the centre, which is said by saying nothing. */
export const OFFICE_TEXT_SHADOW_ALIGNS = ["tl", "t", "tr", "l", "r", "bl", "b", "br"] as const;
export type OfficeTextShadowAlign = (typeof OFFICE_TEXT_SHADOW_ALIGNS)[number];

export interface OfficeTextShadow {
  color: string;
  blur: number;
  offsetX: number;
  offsetY: number;
  opacity: number;
  /**
   * The outer shadow's own transform, DrawingML's (`a:outerShdw` /
   * `w14:shadow`; the server's `office_text_paint.SHADOW_TRANSFORM_KEYS`,
   * where what Word draws for it is written down): the shadow is a copy of
   * the words scaled (`scale` = `@sx`, and `scaleY` = `@sy` when it differs,
   * percent; a negative `scaleY` is the copy flipped) and skewed (`skewX` =
   * `@kx`, `skewY` = `@ky`, degrees) about one of the nine anchors of the
   * words' box (`align` = `@algn`; absent = the centre), then moved by its
   * offset.
   */
  scale?: number;
  scaleY?: number;
  skewX?: number;
  skewY?: number;
  align?: OfficeTextShadowAlign;
}

/** Text paint only: applying a gallery recipe never owns font, text or geometry. */
export interface OfficeTextPaint {
  fill: string;
  outline: string;
  outlineWidthPt: number;
  fillGradient?: OfficeShapeGradient;
  /** Text Fill > Transparency: how opaque a flat `fill` is (absent = 1).
   * Never beside a gradient or a pattern. */
  fillOpacity?: number;
  shadow?: OfficeTextShadow;
  /** `size`: Office's reflection size (`a:reflection@endPos`, percent):
   * Tight 35, Half 55, Full 90; absent draws the paint's own default. */
  reflection?: { opacity: number; blur: number; distance: number; size?: number };
  glow?: { color: string; radius: number; opacity: number };
  /** `preset`: the bevel's type (`ST_BevelPresetType`, Text Effects >
   * Bevel); absent is the paint's own (a circle). */
  bevel?: { width: number; height: number; depth: number; preset?: string };
  /** Text Effects > 3-D Rotation: the text's camera preset. */
  rotation3d?: { camera: string };
  /** Text Fill > Pattern (PowerPoint's WordArt items 16-20); `fill` is its foreground. */
  fillPattern?: OfficeChartFillPattern;
  /** Text Effects > Shadow > Inner. */
  innerShadow?: OfficeTextShadow;
  /**
   * Office's own statement of the paint (`office-text-paint-stated`): theme
   * colours, exact effect geometry, 3-D. The deck's writers put it back
   * verbatim while it still draws the fields above; an edit drops it.
   */
  stated?: OfficeTextPaintStated;
}

const hex = (v: unknown): v is string =>
  typeof v === "string" && /^#[0-9a-f]{6}$/i.test(v);
const number = (v: unknown, min: number, max: number): v is number =>
  typeof v === "number" && Number.isFinite(v) && v >= min && v <= max;
const record = (
  value: unknown,
  keys: readonly string[],
): Record<string, unknown> => {
  if (
    !value ||
    typeof value !== "object" ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !keys.includes(key))
  ) {
    throw Error("invalid_office_text_paint");
  }
  return value as Record<string, unknown>;
};
const rgb = (value: unknown): string => {
  if (!hex(value)) throw Error("invalid_office_text_paint_color");
  return value.toLowerCase();
};

export function normalizeOfficeTextPaint(
  value: unknown,
): OfficeTextPaint | undefined {
  if (value == null) return undefined;
  const p = record(value, [
    "fill",
    "outline",
    "outlineWidthPt",
    "fillGradient",
    "shadow",
    "reflection",
    "glow",
    "bevel",
    "fillPattern",
    "innerShadow",
    "stated",
    "fillOpacity",
    "rotation3d",
  ]);
  if (!number(p.outlineWidthPt, 0, 12))
    throw Error("invalid_office_text_outline");
  const next: OfficeTextPaint = {
    fill: p.fill === "none" ? "none" : rgb(p.fill),
    outline: p.outline === "none" ? "none" : rgb(p.outline),
    outlineWidthPt: p.outlineWidthPt,
  };
  if (p.fillGradient != null) {
    if (next.fill === "none") throw Error("conflicting_office_text_fill");
    next.fillGradient = normalizeOfficeShapeGradient(p.fillGradient);
    if (next.fillGradient!.stops[0].color !== next.fill)
      throw Error("conflicting_office_text_fill");
  }
  if (p.fillPattern != null) {
    let pattern: OfficeChartFillPattern | undefined;
    try {
      pattern = normalizeOfficeChartFillPattern(p.fillPattern);
    } catch {
      throw Error("invalid_office_text_pattern");
    }
    if (!pattern || next.fill !== pattern.foreground || next.fillGradient)
      throw Error("conflicting_office_text_fill");
    next.fillPattern = pattern;
  }
  if (p.fillOpacity != null) {
    if (!number(p.fillOpacity, 0, 1) || next.fill === "none" || next.fillGradient || next.fillPattern)
      throw Error("invalid_office_text_fill_opacity");
    if (p.fillOpacity < 1) next.fillOpacity = p.fillOpacity;
  }
  for (const key of ["shadow", "innerShadow"] as const) {
    if (p[key] == null) continue;
    const s = record(p[key], [
      "color",
      "blur",
      "offsetX",
      "offsetY",
      "opacity",
      ...(key === "shadow" ? ["scale", "scaleY", "skewX", "skewY", "align"] : []),
    ]);
    const align = s.align == null || s.align === "ctr" ? undefined : s.align;
    if (
      !number(s.blur, 0, 64) ||
      !number(s.offsetX, -64, 64) ||
      !number(s.offsetY, -64, 64) ||
      !number(s.opacity, 0, 1) ||
      (s.scale != null && !number(s.scale, 10, 400)) ||
      (s.scaleY != null && !(number(s.scaleY, -400, 400) && Math.abs(s.scaleY) >= 10)) ||
      (s.skewX != null && !number(s.skewX, -80, 80)) ||
      (s.skewY != null && !number(s.skewY, -80, 80)) ||
      (align !== undefined && !(OFFICE_TEXT_SHADOW_ALIGNS as readonly unknown[]).includes(align))
    )
      throw Error("invalid_office_text_shadow");
    const scale = typeof s.scale === "number" && s.scale !== 100 ? s.scale : undefined;
    next[key] = {
      color: rgb(s.color),
      blur: s.blur,
      offsetX: s.offsetX,
      offsetY: s.offsetY,
      opacity: s.opacity,
      // Only what says something: 100 %, the same height as width, no skew
      // and the centre are said by saying nothing (as the server keeps them).
      ...(scale !== undefined ? { scale } : {}),
      ...(typeof s.scaleY === "number" && s.scaleY !== (scale ?? 100) ? { scaleY: s.scaleY } : {}),
      ...(typeof s.skewX === "number" && s.skewX !== 0 ? { skewX: s.skewX } : {}),
      ...(typeof s.skewY === "number" && s.skewY !== 0 ? { skewY: s.skewY } : {}),
      ...(align !== undefined ? { align: align as OfficeTextShadowAlign } : {}),
    };
  }
  if (p.reflection != null) {
    const r = record(p.reflection, ["opacity", "blur", "distance", "size"]);
    if (
      !number(r.opacity, 0, 1) ||
      !number(r.blur, 0, 32) ||
      !number(r.distance, 0, 32) ||
      (r.size != null && !number(r.size, 0, 100))
    )
      throw Error("invalid_office_text_reflection");
    next.reflection = {
      opacity: r.opacity,
      blur: r.blur,
      distance: r.distance,
      ...(typeof r.size === "number" ? { size: r.size } : {}),
    };
  }
  if (p.glow != null) {
    const g = record(p.glow, ["color", "radius", "opacity"]);
    if (!number(g.radius, 0, 32) || !number(g.opacity, 0, 1))
      throw Error("invalid_office_text_glow");
    next.glow = { color: rgb(g.color), radius: g.radius, opacity: g.opacity };
  }
  if (p.bevel != null) {
    const b = record(p.bevel, ["width", "height", "depth", "preset"]);
    if (
      !number(b.width, 0, 16) ||
      !number(b.height, 0, 16) ||
      !number(b.depth, 0, 16) ||
      (b.preset != null && !(OFFICE_BEVEL_PRESETS as readonly unknown[]).includes(b.preset))
    )
      throw Error("invalid_office_text_bevel");
    next.bevel = {
      width: b.width,
      height: b.height,
      depth: b.depth,
      ...(typeof b.preset === "string" ? { preset: b.preset } : {}),
    };
  }
  if (p.rotation3d != null) {
    const r = record(p.rotation3d, ["camera"]);
    // The server keeps a text camera only when it turns the words (never
    // the front view, `orthographicFront`).
    if (
      typeof r.camera !== "string" ||
      r.camera === "orthographicFront" ||
      !Object.hasOwn(OFFICE_CAMERA_ROTATIONS, r.camera)
    )
      throw Error("invalid_office_text_rotation");
    next.rotation3d = { camera: r.camera };
  }
  if (p.stated != null) next.stated = normalizeOfficeTextPaintStated(p.stated);
  return next;
}

/**
 * The CSS turn of the words' 3-D Rotation: the camera preset's (the shape
 * painter's table), or the X / Y / Z and Perspective Office states on the
 * camera itself (`a:camera/a:rot`, `@fov`), as Format Text Effects types them.
 */
function officeTextRotationCss(paint: OfficeTextPaint): CSSProperties {
  const scene = paint.stated?.scene3d;
  const camera = paint.rotation3d?.camera ?? scene?.camera ?? "orthographicFront";
  const [presetX, presetY, presetZ] = OFFICE_CAMERA_ROTATIONS[camera] ?? [0, 0, 0, 45];
  const turned = scene?.cameraRotation;
  const x = turned ? officeCameraLonTurn(turned.longitude) : presetX;
  const y = turned ? turned.latitude / 60_000 : presetY;
  const z = turned ? officeCameraRevTurn(turned.revolution) : presetZ;
  // A stated field of view stands the eye nearer or farther than the 600 px
  // every preset is drawn from (45 degrees).
  const fov = Math.min(179, Math.max(1, scene?.fieldOfView !== undefined ? scene.fieldOfView / 60_000 : 45));
  const eye = Math.round((600 * Math.tan((22.5 * Math.PI) / 180)) / Math.tan((fov * Math.PI) / 360));
  const perspective = camera.startsWith("perspective") ? `perspective(${eye}px) ` : "";
  // 3-D Rotation > Distance from ground (`a:sp3d@z`): the words stand that
  // far off their ground toward the eye -- nearer, and so larger, through a
  // perspective camera; a parallel one draws them the same size (as a
  // shape's, `officeShapeEffectsPaint`).
  const ground = paint.stated?.shape3d?.zEmu;
  const lift = perspective && ground ? ` translateZ(${Math.round((ground / 9525) * 100) / 100}px)` : "";
  if (!x && !y && !z && !lift) return {};
  return {
    display: "inline-block",
    transform: `${perspective}rotateZ(${z}deg) rotateX(${-y}deg) rotateY(${x}deg)${lift}`,
  };
}

const alpha = (color: string, opacity: number) =>
  color +
  Math.round(opacity * 255)
    .toString(16)
    .padStart(2, "0");



















/** The custom property painted words carry (inherited by their runs). */
export const OFFICE_TEXT_PAINT_MARK = "--office-text-paint";







export const OFFICE_TEXT_PAINT_INK = "--office-text-paint-ink";
/**
 * The ink, stated only where the fill is clipped to the glyphs (a gradient,
 * a pattern): an equation among the words is drawn in it. Its glyphs sit in
 * KaTeX's inline blocks, which Safari leaves out of an ancestor's text clip,
 * and its rules, roots and boxes are drawn in `currentColor`, which such a
 * paint makes transparent: the equation drew nothing at all.
 */
export const OFFICE_TEXT_PAINT_CLIPPED_INK = "--office-text-paint-clipped-ink";

function officeTextPaintInk(paint: OfficeTextPaint): string | null {
  const stop = paint.fillGradient?.stops.slice().sort((a, b) => a.position - b.position)[0];
  if (stop) return stop.color;
  if (paint.fill !== "none") return paint.fill;
  return paint.outline !== "none" && paint.outlineWidthPt > 0 ? paint.outline : null;
}

/**
 * Whether a text shadow is more than a moved copy: scaled or skewed
 * (Text Effects > Shadow > Perspective). Such a shadow is drawn by the shadow
 * copy (`office-text-shadow-copy`), not by `drop-shadow`.
 */
export function officeTextShadowTransformed(
  shadow: { scale?: number; scaleY?: number; skewX?: number; skewY?: number } | null | undefined,
): boolean {
  if (!shadow) return false;
  const scale = shadow.scale ?? 100;
  const scaleY = shadow.scaleY ?? scale;
  return scale !== 100 || scaleY !== 100 || Boolean(shadow.skewX) || Boolean(shadow.skewY);
}

export function officeTextPaintCss(value: OfficeTextPaint): CSSProperties {
  const paint = normalizeOfficeTextPaint(value)!;
  const surface: string[] = [];
  if (paint.bevel) {
    surface.push("-0.6px -0.6px 0 #ffffff80", "0.6px 0.6px 0 #00000050");
    if (paint.bevel.depth)
      surface.push(
        `${paint.bevel.depth / 2}px ${paint.bevel.depth / 2}px 0 #00000038`,
      );
  }
  if (paint.innerShadow) {
    // CSS has no inner text shadow: drawn as a faint shadow on the side the
    // light comes from, which reads as the glyphs pressed in.
    const s = paint.innerShadow;
    surface.push(
      `${-s.offsetX / 2}px ${-s.offsetY / 2}px ${Math.min(s.blur, 6) / 2}px ${alpha(s.color, s.opacity * 0.6)}`,
    );
  }
  // A pattern fill is the shared pattern drawing, clipped to the glyphs.
  // Longhands only: the same object states `backgroundClip`, and React
  // refuses the `background` shorthand beside it (a WordArt changed from a
  // pattern to a gradient or to none left a warning and a stale layer).
  const textBackground = paint.fillPattern
    ? officeFillPatternBackground(paint.fillPattern)
    : {
        backgroundImage: paint.fillGradient
          ? officeGradientCss(paint.fillGradient)
          : "none",
      };
  const clipped = Boolean(paint.fillGradient || paint.fillPattern);
  const filled = paint.fill !== "none";
  const seeThrough = filled && paint.fillOpacity !== undefined && paint.fillOpacity < 1;
  const outlined = paint.outline !== "none" && paint.outlineWidthPt > 0;
  const width = paint.outlineWidthPt;
  const effects: string[] = [];
  let stroke = "0 transparent";
  let paintOrder: string | undefined;
  let rings: string | undefined;
  if (outlined) {
    // The outline Office draws (one centred, round stroke of each glyph's
    // merged outline, over the fill) is the overlay's
    // (office-text-outline-css.ts); the strokes and rings below stand in
    // for it until it has drawn, and where it cannot (a textarea).
    if (clipped || seeThrough) {
      rings = [[width, 0], [-width, 0], [0, width], [0, -width]]
        .map(([x, y]) => `drop-shadow(${x}pt ${y}pt 0 ${paint.outline})`)
        .join(" ");
      effects.push(`var(${OFFICE_TEXT_OUTLINE_RINGS})`);
    } else if (filled) {
      stroke = `${width * 2}pt ${paint.outline}`;
      paintOrder = "stroke fill";
    } else {
      stroke = `${width}pt ${paint.outline}`;
    }
  }
  if (paint.shadow && !officeTextShadowTransformed(paint.shadow)) {
    // A perspective shadow (scaled or skewed) is the shadow copy's to draw
    // (`office-text-shadow-copy`): a drop-shadow can only move the words.
    const s = paint.shadow;
    effects.push(`drop-shadow(${s.offsetX}px ${s.offsetY}px ${s.blur}px ${alpha(s.color, s.opacity)})`);
  }
  if (paint.glow) {
    effects.push(`drop-shadow(0 0 ${paint.glow.radius}px ${alpha(paint.glow.color, paint.glow.opacity)})`);
  }
  return {
    ...textBackground,
    color:
      clipped || !filled
        ? "transparent"
        : paint.fillOpacity !== undefined
          ? alpha(paint.fill, paint.fillOpacity)
          : paint.fill,
    backgroundClip: clipped ? "text" : "border-box",
    WebkitBackgroundClip: clipped ? "text" : "border-box",
    WebkitTextStroke: stroke,
    ...(paintOrder ? { paintOrder } : {}),
    textShadow: surface.length ? surface.join(", ") : "none",
    // Composited: a caret blinking inside filtered words (and a style
    // switched under it) repaints a sliver of the layer, which Safari drew
    
    
    ...(effects.length ? { filter: effects.join(" "), willChange: "filter" } : {}),
    ...(rings ? { [OFFICE_TEXT_OUTLINE_RINGS as string]: rings } : {}),
    // Words with a paint and no outline say so, so the outline of words
    // round them does not draw theirs.
    ...(outlined ? officeTextOutlineMarks(paint.outline, width, officeTextOutlineLineOfStated(paint.stated?.line)) : OFFICE_TEXT_OUTLINE_CLEARED),
    // Marks painted words for the shared caret (`office-painted-caret.ts`),
    // which draws a caret of its own over them: one colour whatever the
    // WordArt, never filtered, never transparent with a clipped fill.
    [OFFICE_TEXT_PAINT_MARK as string]: "1",
    ...(officeTextPaintInk(paint) ? { [OFFICE_TEXT_PAINT_INK as string]: officeTextPaintInk(paint) } : {}),
    // A paint of words inside a clipped one states its own: none.
    [OFFICE_TEXT_PAINT_CLIPPED_INK as string]: clipped && officeTextPaintInk(paint) ? officeTextPaintInk(paint) : "initial",
    // A reflection's size is how much of the words it shows (Office's Tight,
    // Half and Full); without one, the reflection's lower 40 %.
    WebkitBoxReflect: paint.reflection
      ? `below ${paint.reflection.distance}px linear-gradient(transparent ${100 - (paint.reflection.size ?? 40)}%, rgba(0,0,0,${paint.reflection.opacity}))`
      : "unset",
    // Text Effects > 3-D Rotation: the words turned as the camera sees them.
    ...(paint.rotation3d || paint.stated?.scene3d?.cameraRotation || paint.stated?.shape3d?.zEmu
      ? officeTextRotationCss(paint)
      : {}),
  } as CSSProperties;
}









export const OFFICE_TEXT_PAINT_HIDDEN_CSS = {
  color: "transparent",
  WebkitTextFillColor: "transparent",
  backgroundImage: "none",
  backgroundClip: "border-box",
  WebkitBackgroundClip: "border-box",
  WebkitTextStroke: "0 transparent",
  paintOrder: "normal",
  textShadow: "none",
  filter: "none",
  WebkitBoxReflect: "unset",
  caretColor: "transparent",
  willChange: "auto",
  [OFFICE_TEXT_PAINT_MARK as string]: "0",
  [OFFICE_TEXT_PAINT_INK as string]: "initial",
  [OFFICE_TEXT_PAINT_CLIPPED_INK as string]: "initial",
  ...OFFICE_TEXT_OUTLINE_CLEARED,
} as CSSProperties;

export function sameOfficeTextPaint(a: unknown, b: unknown): boolean {
  return (
    JSON.stringify(normalizeOfficeTextPaint(a)) ===
    JSON.stringify(normalizeOfficeTextPaint(b))
  );
}

/** Vector title surface for canvas/print hosts that cannot paint gradient text.
 * It is ephemeral display data only; the native document retains real text. */

/** An outline's join, miter limit and cap as SVG stroke attributes. */
function svgOutlineStroke(line: OfficeTextOutlineStroke): string {
  return ` stroke-linejoin="${line.join}"${line.join === "miter" ? ` stroke-miterlimit="${line.miterLimit}"` : ""}`
    + ` stroke-linecap="${line.dash ? line.cap : "round"}"`;
}

/** A dashed outline's dash array for a line `unit` wide (none when solid). */
function svgOutlineDash(line: OfficeTextOutlineStroke, unit: number): string {
  return line.dash ? ` stroke-dasharray="${line.dash.map((length) => +(length * unit).toPrecision(6)).join(" ")}"` : "";
}

export function officeTextPaintSvg(
  text: string,
  value: OfficeTextPaint,
  fontSize = 14,
  /**
   * The face and weight the title draws in without its WordArt. The picture
   * stands in for that text, so it keeps its font: a 36 px bold title came
   * back as 14 px regular sans-serif (PPT-B120).
   */
  font: { family?: string; weight?: string | number; italic?: boolean } = {},
) {
  if (
    typeof text !== "string" ||
    text.length > 2000 ||
    !Number.isFinite(fontSize) ||
    fontSize < 8 ||
    fontSize > 96
  )
    throw Error("office_text_preview_budget");
  const p = normalizeOfficeTextPaint(value)!;
  const escape = (value: string) =>
    value.replace(
      /[&<>"']/g,
      (ch) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&apos;",
        })[ch]!,
    );
  const padding = 16;
  const width = Math.max(
    64,
    Array.from(text).length * fontSize * 1.3 + padding * 2,
  );
  const height = fontSize * (p.reflection ? 2 : 1) + padding * 2;
  if (width * height > 4_000_000) throw Error("office_text_preview_budget");
  const y = padding + fontSize * 0.85;
  const defs: string[] = [];
  if (p.fillGradient?.path) {
    // A path gradient spreads from its focus rectangle's centre out to the
    // box, as the editor's CSS radial gradient draws it (officeGradientCss).
    const rect = p.fillGradient.path.fillToRect ?? {};
    const cx = ((rect.l ?? 0) + (100_000 - (rect.r ?? 0))) / 200_000;
    const cy = ((rect.t ?? 0) + (100_000 - (rect.b ?? 0))) / 200_000;
    defs.push(
      `<radialGradient id="fill" cx="${cx}" cy="${cy}" r="0.75">${officeGradientDrawnStops(p.fillGradient.stops).map((s) => `<stop offset="${s.position}" stop-color="${s.color}"/>`).join("")}</radialGradient>`,
    );
  } else if (p.fillGradient) {
    const angle = ((officeGradientAngle(p.fillGradient) - 90) * Math.PI) / 180;
    const dx = Math.cos(angle),
      dy = Math.sin(angle);
    defs.push(
      `<linearGradient id="fill" x1="${(1 - dx) / 2}" y1="${(1 - dy) / 2}" x2="${(1 + dx) / 2}" y2="${(1 + dy) / 2}">${officeGradientDrawnStops(p.fillGradient.stops).map((s) => `<stop offset="${s.position}" stop-color="${s.color}"/>`).join("")}</linearGradient>`,
    );
  }
  const shadows: Array<{
    color: string;
    opacity: number;
    blur: number;
    offsetX: number;
    offsetY: number;
  }> = [];
  if (p.bevel) {
    shadows.push(
      { color: "#ffffff", opacity: 0.5, blur: 0, offsetX: -0.6, offsetY: -0.6 },
      { color: "#000000", opacity: 0.31, blur: 0, offsetX: 0.6, offsetY: 0.6 },
    );
    if (p.bevel.depth)
      shadows.push({
        color: "#000000",
        opacity: 0.22,
        blur: 0,
        offsetX: p.bevel.depth / 2,
        offsetY: p.bevel.depth / 2,
      });
  }
  if (p.shadow) shadows.push(p.shadow);
  if (p.glow)
    shadows.push({
      color: p.glow.color,
      opacity: p.glow.opacity,
      blur: p.glow.radius,
      offsetX: 0,
      offsetY: 0,
    });
  if (shadows.length)
    defs.push(
      `<filter id="effects" filterUnits="userSpaceOnUse" x="0" y="0" width="${width}" height="${height}">${shadows.map((s) => `<feDropShadow dx="${s.offsetX}" dy="${s.offsetY}" stdDeviation="${s.blur / 2}" flood-color="${s.color}" flood-opacity="${s.opacity}"/>`).join("")}</filter>`,
    );
  const family = font.family?.trim() ? escape(cssFontFamilyList(font.family.trim())) : "sans-serif";
  const weight =
    font.weight === "bold" || (typeof font.weight === "number" && font.weight >= 600)
      ? "bold"
      : "normal";
  // The outline as Office draws it: one centred, round-joined stroke of
  // each glyph's merged outline, over the fill (office-text-outline-css.ts).
  // With the glyphs' outlines to hand (the page's served faces), the words
  // and the band are both drawn from them, so they meet exactly; until then
  // the picture's text is stroked round, centred, over its fill.
  const outlined = p.outline !== "none" && p.outlineWidthPt > 0;
  const widthPx = ptToPx(p.outlineWidthPt);
  // Its join, cap and dash as Office states them (round, solid where none is).
  const line = officeTextOutlineStroke(officeTextOutlineLineOfStated(p.stated?.line));
  const fillPaint = p.fillGradient ? "url(#fill)" : p.fill;
  const placed = outlined && font.family?.trim()
    ? officeTextOutlineSvgGlyphs(text, font.family.trim(), weight === "bold" ? 700 : 400, Boolean(font.italic), fontSize)
    : null;
  let glyph: string;
  if (placed) {
    const start = width / 2 - placed.width / 2;
    const ids = new Map<string, string>();
    let uses = "";
    for (const g of placed.glyphs) {
      let id = ids.get(g.key);
      if (!id) {
        id = `g${ids.size}`;
        ids.set(g.key, id);
        defs.push(`<path id="${id}" d="${g.d}"/>`);
      }
      uses += `<use href="#${id}" transform="matrix(${+g.scale.toPrecision(6)} 0 ${+(g.skew * g.scale).toPrecision(6)} ${+(-g.scale).toPrecision(6)} ${+(start + g.x).toFixed(2)} ${+y.toFixed(2)})" stroke-width="${+(widthPx / g.scale).toPrecision(6)}"${svgOutlineDash(line, widthPx / g.scale)}/>`;
    }
    glyph = `<g class="office-text-glyphs"${shadows.length ? ' filter="url(#effects)"' : ""}>` +
      (p.fill !== "none" || p.fillGradient ? `<g fill="${fillPaint}">${uses}</g>` : "") +
      `<g fill="none" stroke="${p.outline}"${svgOutlineStroke(line)}>${uses}</g></g>`;
  } else {
    glyph = `<text x="${width / 2}" y="${y}" text-anchor="middle" font-family="${family}" font-size="${fontSize}" font-weight="${weight}"${font.italic ? ' font-style="italic"' : ""} fill="${fillPaint}" stroke="${outlined ? p.outline : "none"}"${outlined ? ` stroke-width="${widthPx}"${svgOutlineStroke(line)}${svgOutlineDash(line, widthPx)}` : ""}${shadows.length ? ' filter="url(#effects)"' : ""}>${escape(text)}</text>`;
  }
  let reflection = "";
  if (p.reflection) {
    const top = padding + fontSize + p.reflection.distance;
    defs.push(
      `<linearGradient id="fade" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="white" stop-opacity="${p.reflection.opacity}"/><stop offset="0.4" stop-color="white" stop-opacity="0"/></linearGradient><mask id="reflection"><rect x="0" y="${top}" width="${width}" height="${fontSize}" fill="url(#fade)"/></mask>`,
    );
    reflection = `<g mask="url(#reflection)"><g transform="translate(0 ${2 * (padding + fontSize) + p.reflection.distance}) scale(1 -1)">${glyph}</g></g>`;
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><defs>${defs.join("")}</defs>${glyph}${reflection}</svg>`;
  return {
    uri: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`,
    svg,
    width,
    height,
    padding,
    /** The middle of the words' em box in the picture: what stands on the
     * middle of the title's own (transparent) words. */
    glyphCenter: { x: width / 2, y: padding + fontSize / 2 },
  };
}

export function patchOfficeTextPaint(
  value: OfficeTextPaint,
  patch: {
    fill?: string;
    outline?: string;
    outlineWidthPt?: number;
    effect?: string;
  },
): OfficeTextPaint {
  const next = { ...normalizeOfficeTextPaint(value)! };
  // An edit is the user's own paint now: Office's statement of the style it
  // came from no longer says it (the writers put back what is drawn).
  delete next.stated;
  if (patch.fill !== undefined) {
    next.fill = patch.fill;
    delete next.fillGradient;
    delete next.fillPattern;
  }
  if (patch.outline !== undefined) {
    next.outline = patch.outline;
    next.outlineWidthPt =
      patch.outline === "none" ? 0 : next.outlineWidthPt || 1;
  }
  if (patch.outlineWidthPt !== undefined)
    next.outlineWidthPt = patch.outlineWidthPt;
  if (patch.effect !== undefined) {
    if (
      !["none", "shadow", "reflection", "glow", "bevel", "threeD"].includes(
        patch.effect,
      )
    ) {
      throw Error("unsupported_office_text_paint_effect");
    }
    delete next.shadow;
    delete next.innerShadow;
    delete next.reflection;
    delete next.glow;
    delete next.bevel;
    if (patch.effect === "shadow")
      next.shadow = {
        color: "#000000",
        blur: 3,
        offsetX: 2,
        offsetY: 2,
        opacity: 0.4,
      };
    if (patch.effect === "reflection")
      next.reflection = { opacity: 0.35, blur: 0, distance: 1 };
    if (patch.effect === "glow")
      next.glow = { color: "#5b9bd5", radius: 5, opacity: 0.8 };
    if (patch.effect === "bevel" || patch.effect === "threeD")
      next.bevel = {
        width: 2,
        height: 2,
        depth: patch.effect === "threeD" ? 4 : 1,
      };
  }
  return normalizeOfficeTextPaint(next)!;
}

/** One effect of a text paint, or null to take it off (`patchOfficeTextPaintParts`). */
/**
 * Font Color on characters that wear a paint: the colour IS their fill.
 *
 * Home's Font Color in PowerPoint, Word and Excel sets the characters'
 * solid fill and keeps their outline and effects (the WordArt still has
 * its outline, in the new colour's fill). The deck, the document's shapes
 * and the worksheet's labels used to write the plain colour beside the
 * paint, which goes on drawing the glyphs: the menu changed, the hover
 * preview showed the colour, and the words stayed as they were.
 *
 * The colour replaces a gradient, a pattern and a transparency (a solid
 * colour, as the menu shows it). `color` undefined is Automatic: the fill
 * goes back to `automaticFill` (the host's default text colour) where one
 * is given. A paint left with nothing but a flat fill is null, so the
 * words wear the plain colour again; `keepBare` keeps it a paint instead,
 * where the words sit inside a paint of their container's (a WordArt cell
 * or shape) and only a paint of their own shows over it.
 */
export function officeTextPaintWithFontColor(
  paint: OfficeTextPaint,
  color: string | undefined,
  options: { readonly automaticFill?: string; readonly keepBare?: boolean } = {},
): OfficeTextPaint | null {
  const fill = color ?? options.automaticFill;
  const next = { ...(fill === undefined ? normalizeOfficeTextPaint(paint)! : patchOfficeTextPaint(paint, { fill })) };
  if (fill !== undefined) delete next.fillOpacity;
  // Font Color changes the ink and nothing else: what Office states of the
  
  
  const stated = color !== undefined ? normalizeOfficeTextPaint(paint)?.stated : undefined;
  if (stated && color !== undefined) {
    const base = color.toLowerCase();
    next.stated = { ...stated, fill: { kind: "solid", color: { kind: "srgb", base, val: base.slice(1).toUpperCase() } } };
    delete next.stated.inheritedFill;
  }
  if (options.keepBare) return next;
  const fillOnly =
    (next.outline === "none" || !(next.outlineWidthPt > 0)) &&
    next.fillGradient === undefined &&
    next.fillPattern === undefined &&
    (next.fillOpacity === undefined || next.fillOpacity >= 1) &&
    next.shadow === undefined &&
    next.innerShadow === undefined &&
    next.reflection === undefined &&
    next.glow === undefined &&
    next.bevel === undefined &&
    next.rotation3d === undefined &&
    next.stated === undefined;
  return fillOnly ? null : next;
}

export interface OfficeTextPaintParts {
  shadow?: OfficeTextShadow | null;
  innerShadow?: OfficeTextShadow | null;
  reflection?: OfficeTextPaint["reflection"] | null;
  glow?: OfficeTextPaint["glow"] | null;
  bevel?: OfficeTextPaint["bevel"] | null;
}

/**
 * The paint with some of its effects set or cleared and the rest kept: the
 * Format Shape pane's Text Effects page edits Shadow, Reflection, Glow and
 * 3-D Format one at a time, where the ribbon's Text Effects menu
 * (`patchOfficeTextPaint`'s `effect`) picks one preset for the whole paint.
 * An edit is the user's own paint, so Office's statement is dropped.
 */
export function patchOfficeTextPaintParts(
  value: OfficeTextPaint,
  parts: OfficeTextPaintParts,
): OfficeTextPaint {
  const next = { ...normalizeOfficeTextPaint(value)! } as OfficeTextPaint &
    Record<string, unknown>;
  delete next.stated;
  for (const key of ["shadow", "innerShadow", "reflection", "glow", "bevel"] as const) {
    if (!(key in parts)) continue;
    const part = parts[key];
    if (part == null) delete next[key];
    else next[key] = { ...part } as never;
  }
  return normalizeOfficeTextPaint(next)!;
}
