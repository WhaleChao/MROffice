/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { OfficeBevelEffect, OfficeFormat3dEffect } from "../office-picture-effects";

const EMU_PER_PX = 9_525;

/** The ECharts item style key a mark's 3-D Format rides under (the engine draws none). */
export const OFFICE_CHART_BEVEL = "officeBevel3d";
/** The option key the chart area's and the plot area's 3-D Formats ride under (`{ chart?, plot? }`). */
export const OFFICE_CHART_AREA_BEVELS = "officeAreaBevels3d";

export interface OfficeChartBevel {
  readonly bevel?: OfficeBevelEffect;
  readonly format3d?: OfficeFormat3dEffect;
  /** The mark's own fill, for a Contour in Automatic (its line's colour, else black). */
  readonly automaticContour?: string;
  /** The mark's own fill (a plain colour): a box's bevel is then drawn in
   * the colours PowerPoint draws (`officeChartLitBoxBevel`). */
  readonly fill?: string;
}

export type OfficeChartBevelPiece =
  | {
      readonly kind: "polygon";
      readonly points: readonly (readonly [number, number])[];
      readonly fill: string;
      readonly opacity: number;
    }
  | {
      readonly kind: "sector";
      readonly cx: number;
      readonly cy: number;
      readonly r0: number;
      readonly r: number;
      readonly startAngle: number;
      readonly endAngle: number;
      readonly clockwise: boolean;
      readonly fill: string;
      readonly opacity: number;
    }
  | {
      readonly kind: "outline";
      readonly points: readonly (readonly [number, number])[];
      readonly stroke: string;
      readonly width: number;
    }
  | {
      readonly kind: "arcOutline";
      readonly cx: number;
      readonly cy: number;
      readonly r0: number;
      readonly r: number;
      readonly startAngle: number;
      readonly endAngle: number;
      readonly clockwise: boolean;
      readonly stroke: string;
      readonly width: number;
    };

/** Clockwise from the top, where `ST_LightRigDirection` puts the light. */
const LIGHT_DIRECTION_DEG: Readonly<Record<string, number>> = {
  t: 0, tr: 45, r: 90, br: 135, b: 180, bl: 225, l: 270, tl: 315,
};
 
const KEY_LIGHT_TILT_DEG = -35;

/** How hard each material shows the light: [lit, shaded, the whole rim darkened]. */
const MATERIAL_CONTRAST: Readonly<Record<string, readonly [number, number, number]>> = {
  warmMatte: [0.42, 0.34, 0],
  matte: [0.36, 0.3, 0],
  legacyMatte: [0.36, 0.3, 0],
  plastic: [0.55, 0.38, 0],
  legacyPlastic: [0.55, 0.38, 0],
  metal: [0.62, 0.48, 0],
  legacyMetal: [0.62, 0.48, 0],
  softmetal: [0.5, 0.4, 0],
  dkEdge: [0.3, 0.45, 0.28],
  softEdge: [0.28, 0.22, 0],
  powder: [0.32, 0.26, 0],
  translucentPowder: [0.3, 0.22, 0],
  clear: [0.5, 0.3, 0],
  flat: [0, 0, 0],
  legacyWireframe: [0, 0, 0],
};
 
const FLAT_RIGS = new Set(["flat", "legacyFlat1", "legacyFlat2", "legacyFlat3", "legacyFlat4"]);
const SOFT_RIGS = new Set(["soft", "balanced", "flood", "brightRoom", "glow"]);
/** The bevel's steps across its band (a print has no gradient of opacity). */
const STEPS = 6;

const round = (value: number) => Math.round(value * 100) / 100;

/** The light's direction in screen coordinates (y down), a unit vector towards it. */
export function officeChartBevelLight(format3d: OfficeFormat3dEffect | undefined): readonly [number, number] {
  const lighting = format3d?.lighting;
  const base = LIGHT_DIRECTION_DEG[lighting?.direction ?? "t"] ?? 0;
  const angle = lighting?.angle ? lighting.angle / 60_000 : 0;
  const rig = lighting?.rig ?? "threePt";
  const tilt = SOFT_RIGS.has(rig) ? KEY_LIGHT_TILT_DEG / 2 : KEY_LIGHT_TILT_DEG;
  const degrees = ((base + angle + tilt) * Math.PI) / 180;
  return [Math.sin(degrees), -Math.cos(degrees)];
}

 
function contrast(value: OfficeChartBevel): readonly [number, number, number] {
  const material = MATERIAL_CONTRAST[value.format3d?.material ?? "warmMatte"] ?? MATERIAL_CONTRAST.warmMatte!;
  const rig = value.format3d?.lighting?.rig ?? "threePt";
  if (FLAT_RIGS.has(rig)) return [0, 0, material[2]];
  const scale = SOFT_RIGS.has(rig) ? 0.7 : rig === "harsh" || rig === "contrasting" ? 1.25 : 1;
  // A taller bevel turns its rim further from the face: it shows the light harder.
  const bevel = value.bevel;
  const steep = bevel && bevel.width > 0 ? Math.max(0.5, Math.min(1.5, Math.sqrt(bevel.height / bevel.width))) : 1;
  return [material[0] * scale * steep, material[1] * scale * steep, material[2]];
}

/** The rim's paint at `t` (0 at the outline .. 1 at the band's inner edge) facing `facing` (-1..1). */
function rimPaint(
  facing: number,
  t: number,
  [lit, shaded, dark]: readonly [number, number, number],
  profile: string,
): { fill: string; opacity: number } | null {
  // A rounded bevel ("circle", the default) turns most at its outer edge and
  // settles into the face; an angle or a slope keeps one slant across.
  const falloff =
    profile === "angle" || profile === "slope" || profile === "coolSlant" || profile === "hardEdge"
      ? 1
      : profile === "convex" || profile === "softRound"
        ? 1 - t * t
        : 1 - t;
  const along = facing >= 0 ? lit * facing : shaded * -facing;
  const white = facing >= 0;
  const amount = along * falloff;
  if (dark > 0) {
    // Dark Edge: the rim darkens whatever it faces, more at the outline.
    const darkness = dark * (1 - t) + (white ? 0 : amount);
    const opacity = Math.min(0.85, darkness);
    return opacity > 0.004 ? { fill: "#000000", opacity: round(opacity * 100) / 100 } : null;
  }
  return amount > 0.004 ? { fill: white ? "#ffffff" : "#000000", opacity: Math.round(Math.min(0.85, amount) * 1000) / 1000 } : null;
}

function bevelWidthPx(bevel: OfficeBevelEffect | undefined, most: number): number {
  if (!bevel) return 0;
  return Math.max(0, Math.min(most, bevel.width / EMU_PER_PX));
}





















type LitProfile = readonly (readonly [s: number, k: number])[];
const LIT_FACE: readonly [k: number, white: number] = [1.039, 0.017];
const ROUND_PROFILES: readonly (readonly [facing: number, profile: LitProfile])[] = [
  [0.82, [[0, 0.85], [0.2, 1.08], [0.6, 1.07], [0.9, 1.065], [1, 1.039]]],
  [0.57, [[0, 0.586], [0.15, 0.695], [0.35, 0.883], [0.55, 0.977], [0.9, 1.055], [1, 1.039]]],
  [-0.57, [[0, 0.17], [0.18, 0.44], [0.36, 0.66], [0.55, 0.81], [0.73, 0.89], [0.91, 1], [1, 1.039]]],
  [-0.82, [[0, 0.6], [0.14, 0.5], [0.23, 0.516], [0.41, 0.69], [0.59, 0.836], [0.77, 0.945], [0.95, 1.016], [1, 1.039]]],
];
const STRAIGHT_PROFILES: readonly (readonly [facing: number, k: number])[] = [
  [0.82, 1.1],
  [0.57, 0.86],
  [-0.57, 0.6],
  [-0.82, 0.516],
];
/**
 * The white on the edges facing toward and away from the light: a lobe
 * (peak, where across the band, its spread toward the outline and toward
 * the face) and a glint at the outline itself. The top's lobe 0.62 across;
 * the bottom catches a reflection 0.45 across (0.095 white on warm matte,
 * 0.37 on metal) and a glint at its outline (0.14, 0.61); metal's top a
 * glint of 0.37 too.
 */
type Lobe = readonly [peak: number, at: number, outward: number, inward: number];
const HIGHLIGHT: Readonly<Record<"matte" | "metal", { readonly top: Lobe; readonly bottom: Lobe; readonly glint: readonly [top: number, bottom: number] }>> = {
  matte: { top: [0.24, 0.62, 0.19, 0.26], bottom: [0.095, 0.45, 0.2, 0.2], glint: [0, 0.14] },
  metal: { top: [1.2, 0.62, 0.15, 0.2], bottom: [0.37, 0.45, 0.2, 0.2], glint: [0.37, 0.61] },
};
const ANGLE_BOTTOM_WHITE = 0.038;
/** The materials the measured profiles stand for (the rest keep the rims). */
const LIT_MATERIALS = new Set(["warmMatte", "matte", "legacyMatte", "metal", "legacyMetal"]);
const STRAIGHT_PRESETS = new Set(["angle", "slope", "coolSlant", "hardEdge"]);

function between(table: readonly (readonly [number, number])[], at: number): number {
  if (at <= table[0]![0]) return table[0]![1];
  for (let index = 1; index < table.length; index += 1) {
    const [x1, y1] = table[index]!;
    const [x0, y0] = table[index - 1]!;
    if (at <= x1) return y0 + ((y1 - y0) * (at - x0)) / (x1 - x0 || 1);
  }
  return table[table.length - 1]![1];
}

/** k at `s` across the band of an edge facing `facing`, between the measured edges. */
function litK(facing: number, s: number, straight: boolean): number {
  const rows = straight
    ? STRAIGHT_PROFILES.map(([f, k]) => [f, k] as const)
    : ROUND_PROFILES.map(([f, profile]) => [f, between(profile, s)] as const);
  // Facing runs from 1 (the light) to -1: the table from its top down.
  const descending = [...rows].sort((a, b) => b[0] - a[0]);
  if (facing >= descending[0]![0]) return descending[0]![1];
  for (let index = 1; index < descending.length; index += 1) {
    const [f1, k1] = descending[index]!;
    const [f0, k0] = descending[index - 1]!;
    if (facing >= f1) return k1 + ((k0 - k1) * (facing - f1)) / (f0 - f1 || 1);
  }
  return descending[descending.length - 1]![1];
}

function litColor(fill: readonly [number, number, number], k: number, white: number): string {
  const channel = (c: number) => Math.round(Math.max(0, Math.min(255, c * k + 255 * white))).toString(16).padStart(2, "0");
  return `#${channel(fill[0])}${channel(fill[1])}${channel(fill[2])}`;
}

function plainRgb(color: string | undefined): [number, number, number] | null {
  const hex = typeof color === "string" ? /^#([0-9a-f]{6})$/i.exec(color.trim()) ?? /^#([0-9a-f]{3})$/i.exec(color.trim()) : null;
  if (hex) {
    const digits = hex[1]!.length === 3 ? hex[1]!.split("").map((c) => c + c).join("") : hex[1]!;
    return [parseInt(digits.slice(0, 2), 16), parseInt(digits.slice(2, 4), 16), parseInt(digits.slice(4, 6), 16)];
  }
  const rgb = typeof color === "string" ? /^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*1\s*)?\)$/i.exec(color.trim()) : null;
  return rgb ? [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])] : null;
}






function officeChartLitBoxBevel(
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  value: OfficeChartBevel,
  width: number,
): OfficeChartBevelPiece[] | null {
  const fill = plainRgb(value.fill);
  const material = value.format3d?.material ?? "warmMatte";
  const rig = value.format3d?.lighting?.rig ?? "threePt";
  if (!fill || !LIT_MATERIALS.has(material) || FLAT_RIGS.has(rig) || SOFT_RIGS.has(rig)) return null;
  const preset = value.bevel?.preset ?? "circle";
  const straight = STRAIGHT_PRESETS.has(preset);
  const highlight = HIGHLIGHT[material === "metal" || material === "legacyMetal" ? "metal" : "matte"];
  const [lx, ly] = officeChartBevelLight(value.format3d);
  const x0 = box.x, y0 = box.y, x1 = box.x + box.width, y1 = box.y + box.height;
  const pieces: OfficeChartBevelPiece[] = [];
  // The face inside the band.
  pieces.push({
    kind: "polygon",
    points: [[x0 + width, y0 + width], [x1 - width, y0 + width], [x1 - width, y1 - width], [x0 + width, y1 - width]].map(([x, y]) => [round(x!), round(y!)] as const),
    fill: litColor(fill, LIT_FACE[0], LIT_FACE[1]),
    opacity: 1,
  });
  const edges: readonly {
    readonly normal: readonly [number, number];
    readonly at: (d0: number, d1: number) => (readonly [number, number])[];
  }[] = [
    { normal: [0, -1], at: (d0, d1) => [[x0 + d0, y0 + d0], [x1 - d0, y0 + d0], [x1 - d1, y0 + d1], [x0 + d1, y0 + d1]] },
    { normal: [1, 0], at: (d0, d1) => [[x1 - d0, y0 + d0], [x1 - d0, y1 - d0], [x1 - d1, y1 - d1], [x1 - d1, y0 + d1]] },
    { normal: [0, 1], at: (d0, d1) => [[x1 - d0, y1 - d0], [x0 + d0, y1 - d0], [x0 + d1, y1 - d1], [x1 - d1, y1 - d1]] },
    { normal: [-1, 0], at: (d0, d1) => [[x0 + d0, y1 - d0], [x0 + d0, y0 + d0], [x0 + d1, y0 + d1], [x0 + d1, y1 - d1]] },
  ];
  const steps = straight ? 1 : Math.max(2, Math.min(24, Math.round(width * 2)));
  for (const edge of edges) {
    const facing = edge.normal[0] * lx + edge.normal[1] * ly;
    // The lobes belong to the edges facing toward the light (the measured
    // top) and away from it (the bottom); the sides, at +-0.57, show none.
    const toward = Math.max(0, Math.min(1, (facing - 0.57) / (0.82 - 0.57)));
    const away = Math.max(0, Math.min(1, (-0.57 - facing) / (0.82 - 0.57)));
    const lobe = ([peak, at, outward, inward]: Lobe, across: number) =>
      peak * Math.exp(-(((across - at) / (across < at ? outward : inward)) ** 2));
    for (let step = 0; step < steps; step += 1) {
      const sMid = (step + 0.5) / steps;
      const k = litK(facing, sMid, straight);
      const glint = Math.exp(-(((sMid - 0.05) / 0.06) ** 2));
      const white = straight
        ? away * ANGLE_BOTTOM_WHITE
        : toward * (lobe(highlight.top, sMid) + highlight.glint[0] * glint) +
          away * (lobe(highlight.bottom, sMid) + highlight.glint[1] * glint);
      pieces.push({
        kind: "polygon",
        points: edge.at((width * step) / steps, (width * (step + 1)) / steps).map(([x, y]) => [round(x), round(y)] as const),
        fill: litColor(fill, k, white),
        opacity: 1,
      });
    }
  }
  return pieces;
}

/**
 * A rectangular mark's (a column's, a bar's, an area's box) bevel rims and
 * contour, px.
 */
export function officeChartBoxBevel(
  box: { readonly x: number; readonly y: number; readonly width: number; readonly height: number },
  value: OfficeChartBevel,
): OfficeChartBevelPiece[] {
  const pieces: OfficeChartBevelPiece[] = [];
  if (!(box.width > 0 && box.height > 0)) return pieces;
  const width = bevelWidthPx(value.bevel, Math.min(box.width, box.height) / 2);
  const lit = width >= 0.5 ? officeChartLitBoxBevel(box, value, width) : null;
  if (lit) pieces.push(...lit);
  const strength = contrast(value);
  const [lx, ly] = officeChartBevelLight(value.format3d);
  const profile = value.bevel?.preset ?? "circle";
  if (width >= 0.5 && !lit) {
    const x0 = box.x, y0 = box.y, x1 = box.x + box.width, y1 = box.y + box.height;
    // Each edge: its outward normal and its band's quadrilateral at a depth
    // d in from the outline, mitred at the corners.
    const edges: readonly {
      readonly normal: readonly [number, number];
      readonly at: (d0: number, d1: number) => (readonly [number, number])[];
    }[] = [
      { normal: [0, -1], at: (d0, d1) => [[x0 + d0, y0 + d0], [x1 - d0, y0 + d0], [x1 - d1, y0 + d1], [x0 + d1, y0 + d1]] },
      { normal: [1, 0], at: (d0, d1) => [[x1 - d0, y0 + d0], [x1 - d0, y1 - d0], [x1 - d1, y1 - d1], [x1 - d1, y0 + d1]] },
      { normal: [0, 1], at: (d0, d1) => [[x1 - d0, y1 - d0], [x0 + d0, y1 - d0], [x0 + d1, y1 - d1], [x1 - d1, y1 - d1]] },
      { normal: [-1, 0], at: (d0, d1) => [[x0 + d0, y1 - d0], [x0 + d0, y0 + d0], [x0 + d1, y0 + d1], [x0 + d1, y1 - d1]] },
    ];
    const steps = Math.max(1, Math.min(STEPS, Math.round(width)));
    for (const edge of edges) {
      const facing = edge.normal[0] * lx + edge.normal[1] * ly;
      for (let step = 0; step < steps; step += 1) {
        const d0 = (width * step) / steps;
        const d1 = (width * (step + 1)) / steps;
        const paint = rimPaint(facing, (step + 0.5) / steps, strength, profile);
        if (!paint) continue;
        pieces.push({
          kind: "polygon",
          points: edge.at(d0, d1).map(([x, y]) => [round(x), round(y)] as const),
          ...paint,
        });
      }
    }
  }
  const contour = value.format3d?.contour;
  if (contour && contour.width > 0) {
    // Half outside the outline and half over the face, as PowerPoint draws
    // a contour face on.
    pieces.push({
      kind: "outline",
      points: [
        [round(box.x), round(box.y)],
        [round(box.x + box.width), round(box.y)],
        [round(box.x + box.width), round(box.y + box.height)],
        [round(box.x), round(box.y + box.height)],
      ],
      stroke: contour.color ?? value.automaticContour ?? "#000000",
      width: round(contour.width / EMU_PER_PX),
    });
  }
  return pieces;
}

/**
 * A pie or doughnut slice's bevel rims and contour, px. The slice is the
 * engine's sector (centre, radii, angles in radians as ZRender draws them:
 * x = cx + r cos a, y = cy + r sin a).
 */
export function officeChartSectorBevel(
  sector: {
    readonly cx: number;
    readonly cy: number;
    readonly r0: number;
    readonly r: number;
    readonly startAngle: number;
    readonly endAngle: number;
    readonly clockwise: boolean;
  },
  value: OfficeChartBevel,
): OfficeChartBevelPiece[] {
  const pieces: OfficeChartBevelPiece[] = [];
  const { cx, cy, r0, r, startAngle, endAngle, clockwise } = sector;
  if (!(r > r0) || startAngle === endAngle) return pieces;
  const sweep = endAngle - startAngle;
  const width = bevelWidthPx(value.bevel, (r - r0) / 2);
  const strength = contrast(value);
  const [lx, ly] = officeChartBevelLight(value.format3d);
  const profile = value.bevel?.preset ?? "circle";
  if (width >= 0.5) {
    const steps = Math.max(1, Math.min(STEPS, Math.round(width)));
    // The outer (and a doughnut's inner) arc, in pieces of about 4 degrees,
    // each lit by the way its normal faces.
    const arcs = Math.max(1, Math.min(90, Math.ceil(Math.abs(sweep) / (Math.PI / 45))));
    for (let piece = 0; piece < arcs; piece += 1) {
      const a0 = startAngle + (sweep * piece) / arcs;
      const a1 = startAngle + (sweep * (piece + 1)) / arcs;
      const mid = (a0 + a1) / 2;
      for (const outer of r0 > 0 ? [true, false] : [true]) {
        const facing = (outer ? 1 : -1) * (Math.cos(mid) * lx + Math.sin(mid) * ly);
        for (let step = 0; step < steps; step += 1) {
          const paint = rimPaint(facing, (step + 0.5) / steps, strength, profile);
          if (!paint) continue;
          const d0 = (width * step) / steps;
          const d1 = (width * (step + 1)) / steps;
          pieces.push({
            kind: "sector",
            cx: round(cx),
            cy: round(cy),
            r0: round(outer ? r - d1 : r0 + d0),
            r: round(outer ? r - d0 : r0 + d1),
            startAngle: a0,
            endAngle: a1,
            clockwise,
            ...paint,
          });
        }
      }
    }
    // The two radial edges of a slice short of a whole turn.
    if (Math.abs(sweep) < Math.PI * 2 - 1e-6) {
      const mid = startAngle + sweep / 2;
      for (const angle of [startAngle, endAngle]) {
        const along: readonly [number, number] = [Math.cos(angle), Math.sin(angle)];
        // Outward: away from the slice's middle.
        let normal: [number, number] = [-along[1], along[0]];
        const toMiddle = [Math.cos(mid) - along[0], Math.sin(mid) - along[1]];
        if (normal[0] * toMiddle[0]! + normal[1] * toMiddle[1]! > 0) normal = [-normal[0], -normal[1]];
        const facing = normal[0] * lx + normal[1] * ly;
        const inner = Math.max(r0, width);
        for (let step = 0; step < steps; step += 1) {
          const paint = rimPaint(facing, (step + 0.5) / steps, strength, profile);
          if (!paint) continue;
          const d0 = (width * step) / steps;
          const d1 = (width * (step + 1)) / steps;
          const at = (radius: number, depth: number) =>
            [round(cx + along[0] * radius - normal[0] * depth), round(cy + along[1] * radius - normal[1] * depth)] as const;
          pieces.push({
            kind: "polygon",
            points: [at(inner, d0), at(r - d0, d0), at(r - d1, d1), at(inner, d1)],
            ...paint,
          });
        }
      }
    }
  }
  const contour = value.format3d?.contour;
  if (contour && contour.width > 0) {
    pieces.push({
      kind: "arcOutline",
      cx: round(cx),
      cy: round(cy),
      r0,
      r,
      startAngle,
      endAngle,
      clockwise,
      stroke: contour.color ?? value.automaticContour ?? "#000000",
      width: round(contour.width / EMU_PER_PX),
    });
  }
  return pieces;
}

/**
 * A polygon's bevel rims and contour (a 3-D column's face, a 3-D slice's
 * top, as the camera sees it): each edge's band inside the outline, mitred where the
 * edges meet, lit by the way the edge faces the light; the contour along the
 * outline. Points in drawing order, px.
 */
export function officeChartPolygonBevel(
  points: readonly (readonly [number, number])[],
  value: OfficeChartBevel,
): OfficeChartBevelPiece[] {
  const pieces: OfficeChartBevelPiece[] = [];
  const count = points.length;
  if (count < 3) return pieces;
  // Signed area: which way round the points go, so each edge's outward side is known.
  let area = 0;
  for (let index = 0; index < count; index += 1) {
    const [x0, y0] = points[index]!;
    const [x1, y1] = points[(index + 1) % count]!;
    area += x0 * y1 - x1 * y0;
  }
  if (Math.abs(area) < 1) return pieces;
  const turn = area > 0 ? 1 : -1;
  const edges = points.map((from, index) => {
    const to = points[(index + 1) % count]!;
    const length = Math.hypot(to[0] - from[0], to[1] - from[1]);
    // Outward normal (screen y down): the side away from the inside.
    const normal: [number, number] = length > 0 ? [(turn * (to[1] - from[1])) / length, (-turn * (to[0] - from[0])) / length] : [0, 0];
    return { from, to, length, normal };
  });
  // How far in a band may reach. Each edge's own room: its mitred ends close
  // in by tan(half the turn) per unit of depth at a convex corner, so a box's
  // side has half its length, while a short step of a sampled arc (a 3-D
  // slice's top), between near-straight turns, has plenty. A reflex corner
  // (a slice past half the pie) lengthens its edges instead.
  const closing = (index: number): number => {
    const before = edges[(index - 1 + count) % count]!.normal;
    const after = edges[index]!.normal;
    const dot = Math.max(-1, Math.min(1, before[0] * after[0] + before[1] * after[1]));
    const cross = before[0] * after[1] - before[1] * after[0];
    return cross * turn > 1e-9 ? Math.sqrt((1 - dot) / Math.max(1e-9, 1 + dot)) : 0;
  };
  let room = Infinity;
  edges.forEach((edge, index) => {
    const shut = closing(index) + closing((index + 1) % count);
    if (shut > 1e-9) room = Math.min(room, edge.length / shut);
  });
  // And the polygon's narrowest breadth (its extent off each edge): no band past half of it.
  let across = Infinity;
  for (const edge of edges) {
    if (edge.length <= 0) continue;
    let extent = 0;
    for (const point of points) {
      extent = Math.max(extent, (edge.from[0] - point[0]) * edge.normal[0] + (edge.from[1] - point[1]) * edge.normal[1]);
    }
    across = Math.min(across, extent);
  }
  const width = bevelWidthPx(value.bevel, Math.min(room, Number.isFinite(across) ? across / 2 : 0));
  if (width >= 0.5) {
    const strength = contrast(value);
    const [lx, ly] = officeChartBevelLight(value.format3d);
    const profile = value.bevel?.preset ?? "circle";
    // A vertex moved `depth` in from both of its edges (the edges' offset lines meet there).
    const inset = (index: number, depth: number): readonly [number, number] => {
      const before = edges[(index - 1 + count) % count]!;
      const after = edges[index]!;
      const point = points[index]!;
      const nx = before.normal[0] + after.normal[0];
      const ny = before.normal[1] + after.normal[1];
      const dot = before.normal[0] * after.normal[0] + before.normal[1] * after.normal[1];
      const scale = depth / Math.max(0.2, 1 + dot);
      return [point[0] - nx * scale, point[1] - ny * scale];
    };
    const steps = Math.max(1, Math.min(STEPS, Math.round(width)));
    edges.forEach((edge, index) => {
      const facing = edge.normal[0] * lx + edge.normal[1] * ly;
      for (let step = 0; step < steps; step += 1) {
        const paint = rimPaint(facing, (step + 0.5) / steps, strength, profile);
        if (!paint) continue;
        const d0 = (width * step) / steps;
        const d1 = (width * (step + 1)) / steps;
        const next = (index + 1) % count;
        pieces.push({
          kind: "polygon",
          points: [inset(index, d0), inset(next, d0), inset(next, d1), inset(index, d1)].map(
            ([x, y]) => [round(x), round(y)] as const,
          ),
          ...paint,
        });
      }
    });
  }
  const contour = value.format3d?.contour;
  if (contour && contour.width > 0) {
    pieces.push({
      kind: "outline",
      points: points.map(([x, y]) => [round(x), round(y)] as const),
      stroke: contour.color ?? value.automaticContour ?? "#000000",
      width: round(contour.width / EMU_PER_PX),
    });
  }
  return pieces;
}
