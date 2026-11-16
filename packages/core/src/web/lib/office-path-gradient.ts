/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/** Insets in 1/1000 % of the box (DrawingML's `ST_Percentage` 100000 = 100 %). */
export interface OfficeGradientInsets {
  readonly l?: number;
  readonly t?: number;
  readonly r?: number;
  readonly b?: number;
}

export interface OfficeRectPathRegion {
  /** Polygon corners in the box's unit square (0..1 each way; the bounds may lie past it). */
  readonly points: readonly (readonly [number, number])[];
  /** Where the stops start (position 0) and end (position 1), the same unit square. */
  readonly from: readonly [number, number];
  readonly to: readonly [number, number];
}

export interface OfficeRectPathGradient {
  /** The focus rectangle [left, top, right, bottom], drawn in the first stop's colour (empty when a point). */
  readonly focus: readonly [number, number, number, number];
  /** The bounds [left, top, right, bottom]: what the last stop's colour reaches. */
  readonly bounds: readonly [number, number, number, number];
  readonly regions: readonly OfficeRectPathRegion[];
}

const round = (value: number) => Math.round(value * 1e6) / 1e6;

function share(value: number | undefined): number {
  return typeof value === "number" && Number.isFinite(value) ? value / 100_000 : 0;
}

export function officeRectPathGradient(
  fillToRect: OfficeGradientInsets | undefined,
  tileRect?: OfficeGradientInsets,
): OfficeRectPathGradient {
  const bl = share(tileRect?.l);
  const bt = share(tileRect?.t);
  const br = 1 - share(tileRect?.r);
  const bb = 1 - share(tileRect?.b);
  // A focus inset past its own opposite side (l + r > 100 %) is a line at
  // its near inset, as the radial focus reads a crossed rectangle
  
  // the box's bottom edge).
  const axis = (near: number, far: number): [number, number] => {
    const a = near;
    const b = 1 - far;
    return a <= b ? [a, b] : [a, a];
  };
  const [fl, fr] = axis(share(fillToRect?.l), share(fillToRect?.r));
  const [ft, fb] = axis(share(fillToRect?.t), share(fillToRect?.b));
  const regions: OfficeRectPathRegion[] = [];
  const push = (points: [number, number][], from: [number, number], to: [number, number]) =>
    regions.push({
      points: points.map(([x, y]) => [round(x), round(y)] as const),
      from: [round(from[0]), round(from[1])],
      to: [round(to[0]), round(to[1])],
    });
  if (fl > bl) push([[bl, bt], [fl, ft], [fl, fb], [bl, bb]], [fl, ft], [bl, ft]);
  if (br > fr) push([[fr, ft], [br, bt], [br, bb], [fr, fb]], [fr, ft], [br, ft]);
  if (ft > bt) push([[bl, bt], [br, bt], [fr, ft], [fl, ft]], [fl, ft], [fl, bt]);
  if (bb > fb) push([[fl, fb], [fr, fb], [br, bb], [bl, bb]], [fl, fb], [fl, bb]);
  return {
    focus: [round(fl), round(ft), round(fr), round(fb)],
    bounds: [round(bl), round(bt), round(br), round(bb)],
    regions,
  };
}

/** Where a point of the unit square falls along the stops (0 at the focus, 1 at the bounds and past them). */
export function officeRectPathGradientPosition(
  gradient: OfficeRectPathGradient,
  x: number,
  y: number,
): number {
  const [fl, ft, fr, fb] = gradient.focus;
  const [bl, bt, br, bb] = gradient.bounds;
  const along = (value: number, near: number, far: number, outNear: number, outFar: number) => {
    if (value < near) return near > outNear ? (near - value) / (near - outNear) : 1;
    if (value > far) return outFar > far ? (value - far) / (outFar - far) : 1;
    return 0;
  };
  return Math.min(1, Math.max(along(x, fl, fr, bl, br), along(y, ft, fb, bt, bb)));
}

/**
 * The CSS `background` of a box filled with a rectangular path gradient: an
 * SVG stretched over the box. It paints the last stop's colour, the focus
 * rectangle in the first, and each trapezoid as a linear gradient across it.
 * Each trapezoid is stroked with its own gradient so that no seam shows.
 * `stops` are as drawn (`officeGradientDrawnStops`).
 */
export function officeRectPathGradientBackground(
  stops: readonly { readonly color: string; readonly position: number; readonly alpha?: number }[],
  fillToRect: OfficeGradientInsets | undefined,
  tileRect?: OfficeGradientInsets,
): string {
  const shape = officeRectPathGradient(fillToRect, tileRect);
  const unit = (value: number) => Number((value * 1000).toFixed(2));
  const opacity = (alpha: number | undefined) => (alpha !== undefined && alpha < 1 ? ` stop-opacity="${alpha}"` : "");
  const stopTags = stops
    .map((stop) => `<stop offset="${Number(stop.position.toFixed(5))}" stop-color="${stop.color}"${opacity(stop.alpha)}/>`)
    .join("");
  const defs: string[] = [];
  const shapes: string[] = [];
  shape.regions.forEach((region, index) => {
    defs.push(
      `<linearGradient id="g${index}" gradientUnits="userSpaceOnUse" x1="${unit(region.from[0])}" y1="${unit(region.from[1])}" x2="${unit(region.to[0])}" y2="${unit(region.to[1])}">${stopTags}</linearGradient>`,
    );
    const points = region.points.map(([x, y]) => `${unit(x)},${unit(y)}`).join(" ");
    shapes.push(`<polygon points="${points}" fill="url(#g${index})" stroke="url(#g${index})" stroke-width="1.5" stroke-linejoin="round"/>`);
  });
  const [fl, ft, fr, fb] = shape.focus;
  const first = stops[0];
  const last = stops[stops.length - 1];
  const fillOpacity = (alpha: number | undefined) => (alpha !== undefined && alpha < 1 ? ` fill-opacity="${alpha}"` : "");
  const focus = first && fr > fl && fb > ft
    ? `<rect x="${unit(fl)}" y="${unit(ft)}" width="${unit(fr - fl)}" height="${unit(fb - ft)}" fill="${first.color}"${fillOpacity(first.alpha)}/>`
    : "";
  const ground = last ? `<rect width="1000" height="1000" fill="${last.color}"${fillOpacity(last.alpha)}/>` : "";
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 1000" preserveAspectRatio="none">` +
    `<defs>${defs.join("")}</defs>${ground}${shapes.join("")}${focus}</svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 0 0 / 100% 100% no-repeat`;
}


















































export interface OfficeCirclePathGradient {
  readonly focus: readonly [number, number];
  readonly centre: readonly [number, number];
  readonly radius: number;
  /** The focus lies outside the circle: each ray takes the circle's far side. */
  readonly outside: boolean;
}

export function officeCirclePathGradient(
  fillToRect: OfficeGradientInsets | undefined,
  tileRect: OfficeGradientInsets | undefined,
  width = 1,
  height = 1,
): OfficeCirclePathGradient {
  const w = Number.isFinite(width) && width > 0 ? width : 1;
  const h = Number.isFinite(height) && height > 0 ? height : 1;
  const stated = (value: number | undefined) => (fillToRect ? share(value) : 0.5);
  const axis = (near: number | undefined, far: number | undefined) => {
    const a = stated(near);
    const b = stated(far);
    if (a + b <= 1e-9) return 0.5;
    return a + b < 1 - 1e-9 ? a / (a + b) : a;
  };
  const fx = axis(fillToRect?.l, fillToRect?.r) * w;
  const fy = axis(fillToRect?.t, fillToRect?.b) * h;
  const bl = share(tileRect?.l) * w;
  const bt = share(tileRect?.t) * h;
  const br = (1 - share(tileRect?.r)) * w;
  const bb = (1 - share(tileRect?.b)) * h;
  const cx = (bl + br) / 2;
  const cy = (bt + bb) / 2;
  const radius = Math.hypot(br - bl, bb - bt) / 2;
  const outside = Math.hypot(fx - cx, fy - cy) > radius * (1 + 1e-9);
  return { focus: [round(fx), round(fy)], centre: [round(cx), round(cy)], radius: round(radius), outside };
}

/** Where a point (the box's units) falls along the stops of `officeCirclePathGradient`: 0 at the focus, 1 on the circle and past it. */
export function officeCirclePathGradientPosition(gradient: OfficeCirclePathGradient, x: number, y: number): number {
  const [fx, fy] = gradient.focus;
  const [cx, cy] = gradient.centre;
  const ex = x - fx;
  const ey = y - fy;
  const length = Math.hypot(ex, ey);
  if (length < 1e-9) return 0;
  const ux = ex / length;
  const uy = ey / length;
  const ox = fx - cx;
  const oy = fy - cy;
  const b = ox * ux + oy * uy;
  const c = ox * ox + oy * oy - gradient.radius * gradient.radius;
  const disc = b * b - c;
  if (disc < 0) return 1;
  const far = -b + Math.sqrt(disc);
  return far > 1e-9 ? Math.min(1, length / far) : 1;
}

/**
 * `officeCirclePathGradient` as an SVG `radialGradient`'s geometry, in the
 * same units (`gradientUnits="userSpaceOnUse"`, offset by `x` / `y`) and its
 * stops in the order to write them.
 * - Focus inside (or on) the circle: the focal radial, `fx` / `fy` the focus.
 * - Focus outside: SVG and a PDF shading take, of the circles through a
 *   point, the one farthest along (the cone's larger root), which is the
 *   circle's NEAR side. Drawn backwards, from the whole circle (`fx` / `fy` /
 *   `fr`) in to the focus point (`cx` / `cy`, a hair of radius: r = 0 would
 *   paint the last stop everywhere) with the stops reversed, the farthest
 *   along is the far side's (Chromium and WebKit draw it so; the PDF export
 *   draws the same two circles as its shading).
 * `reversed` draws the stops the other way round as well (a negative bar's,
 * `chart-render-server`).
 */
export function officeCirclePathGradientSvg<Stop extends { readonly position: number }>(
  gradient: OfficeCirclePathGradient,
  stops: readonly Stop[],
  x = 0,
  y = 0,
  reversed = false,
): { attributes: { cx: number; cy: number; r: number; fx: number; fy: number; fr?: number }; stops: Stop[] } {
  const turn = gradient.outside !== reversed;
  const ordered = turn
    ? [...stops].reverse().map((stop) => ({ ...stop, position: round(1 - stop.position) }))
    : [...stops];
  const at = (value: number, offset: number) => round(value + offset);
  return gradient.outside
    ? {
        attributes: {
          cx: at(gradient.focus[0], x),
          cy: at(gradient.focus[1], y),
          r: 0.001,
          fx: at(gradient.centre[0], x),
          fy: at(gradient.centre[1], y),
          fr: gradient.radius,
        },
        stops: ordered,
      }
    : {
        attributes: {
          cx: at(gradient.centre[0], x),
          cy: at(gradient.centre[1], y),
          r: gradient.radius,
          fx: at(gradient.focus[0], x),
          fy: at(gradient.focus[1], y),
        },
        stops: ordered,
      };
}

/**
 * A circular path gradient (`a:path path="circle"`) as Windows Word draws a
 * shape's fill: concentric circles about the middle of its bounds (the tile
 * rectangle, or the box), out to half the bounds' diagonal, round in the
 * box's own units. Word does not start it at the `a:fillToRect` focus the way
 * PowerPoint does (`officeCirclePathGradient`).
 *
 * Measured on Windows Word's PDF (M365 16.0.20527) of docx-look/probes
 * focus.docx, focus3.docx, focus4.docx: foci stated at a corner, at the middle
 * of an edge, partly stated or absent all drew about the box's centre under an
 * empty tile; a tile out 100 % right and below drew about the box's
 * bottom-right corner whatever the focus said (top-left or bottom-right); a
 * 6 x 2.5 in box drew circles, not an ellipse (0.96 / 0.30 along the stops at
 * the middles of its left and top edges, against 0.96 / 0.30 for circles and
 * 0.76 / 0.74 for the stretched ellipse), and the colour runs along
 * `officeGradientDrawnStops`' eased ramp (within 0.02 at ten radii).
 *
 * The result is in the box's own units (`width` x `height`, y down); the
 * focus is the centre.
 */
export function wordCirclePathGradient(
  tileRect: OfficeGradientInsets | undefined,
  width = 1,
  height = 1,
): OfficeCirclePathGradient {
  const w = Number.isFinite(width) && width > 0 ? width : 1;
  const h = Number.isFinite(height) && height > 0 ? height : 1;
  const bl = share(tileRect?.l) * w;
  const bt = share(tileRect?.t) * h;
  const br = (1 - share(tileRect?.r)) * w;
  const bb = (1 - share(tileRect?.b)) * h;
  const centre = [round((bl + br) / 2), round((bt + bb) / 2)] as const;
  return { focus: centre, centre, radius: round(Math.hypot(br - bl, bb - bt) / 2), outside: false };
}

/**
 * An SVG `radialGradient`'s attributes (objectBoundingBox units) for a circle
 * gradient in a `width` x `height` box: measured in widths, with the
 * transform that keeps the circle round on a box that is not square.
 */
export function officeCircleRadialSvgAttrs(
  circle: OfficeCirclePathGradient,
  width = 1,
  height = 1,
): { cx: number; cy: number; r: number; fx: number; fy: number; gradientTransform?: string } {
  const w = Number.isFinite(width) && width > 0 ? width : 1;
  const h = Number.isFinite(height) && height > 0 ? height : 1;
  return {
    cx: circle.centre[0] / w,
    cy: circle.centre[1] / w,
    r: circle.radius / w,
    fx: circle.focus[0] / w,
    fy: circle.focus[1] / w,
    ...(Math.abs(h / w - 1) > 1e-9 ? { gradientTransform: `scale(1 ${w / h})` } : {}),
  };
}
