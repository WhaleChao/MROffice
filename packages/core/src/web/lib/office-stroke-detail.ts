/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * What a shape's outline states beyond its paint and its dash: Format Shape >
 * Line > Compound type, Cap type and Join type (`a:ln@cmpd`, `a:ln@cap`, and
 * the `a:round` / `a:bevel` / `a:miter` child with `a:miter@lim`).
 *
 * A shape carries each only when its source line STATES it, so every shape
 * authored here and every earlier page holds none and draws as it always did
 * (round caps and joins, one line). Keep in step with the server's
 * `office_shape_paint.read_stroke_detail`.
 */

export const OFFICE_STROKE_COMPOUNDS = ["double", "thickThin", "thinThick", "triple"] as const;
export type OfficeStrokeCompound = (typeof OFFICE_STROKE_COMPOUNDS)[number];
export const OFFICE_STROKE_CAPS = ["flat", "round", "square"] as const;
export type OfficeStrokeCap = (typeof OFFICE_STROKE_CAPS)[number];
export const OFFICE_STROKE_JOINS = ["round", "bevel", "miter"] as const;
export type OfficeStrokeJoin = (typeof OFFICE_STROKE_JOINS)[number];

/** `a:miter@lim` is in 1000ths of a percent of the line width. */
export const OFFICE_STROKE_MAX_MITER_LIMIT = 100_000_000;

export interface OfficeStrokeDetail {
  outlineCompound?: OfficeStrokeCompound;
  outlineCap?: OfficeStrokeCap;
  outlineJoin?: OfficeStrokeJoin;
  outlineMiterLimit?: number;
}

function oneOf<T extends string>(values: readonly T[], value: unknown): T | undefined {
  return typeof value === "string" && (values as readonly string[]).includes(value)
    ? (value as T)
    : undefined;
}

/** The stated values among an attrs object's four keys; unknown ones say nothing. */
export function officeStrokeDetail(attrs: {
  outlineCompound?: unknown;
  outlineCap?: unknown;
  outlineJoin?: unknown;
  outlineMiterLimit?: unknown;
}): OfficeStrokeDetail {
  const compound = oneOf(OFFICE_STROKE_COMPOUNDS, attrs.outlineCompound);
  const cap = oneOf(OFFICE_STROKE_CAPS, attrs.outlineCap);
  const join = oneOf(OFFICE_STROKE_JOINS, attrs.outlineJoin);
  const limit =
    join === "miter" &&
    typeof attrs.outlineMiterLimit === "number" &&
    Number.isInteger(attrs.outlineMiterLimit) &&
    attrs.outlineMiterLimit >= 0 &&
    attrs.outlineMiterLimit <= OFFICE_STROKE_MAX_MITER_LIMIT
      ? attrs.outlineMiterLimit
      : undefined;
  return {
    ...(compound ? { outlineCompound: compound } : {}),
    ...(cap ? { outlineCap: cap } : {}),
    ...(join ? { outlineJoin: join } : {}),
    ...(limit !== undefined ? { outlineMiterLimit: limit } : {}),
  };
}

/**
 * The SVG stroke attributes a detail states. A line that states no cap or
 * join keeps the canvas's own round ones (what every shape was drawn with
 * before these were read).
 */
export function officeStrokeDetailSvgAttributes(detail: OfficeStrokeDetail): {
  "stroke-linecap": "butt" | "round" | "square";
  "stroke-linejoin": "round" | "bevel" | "miter";
  "stroke-miterlimit"?: number;
} {
  const cap =
    detail.outlineCap === "flat" ? "butt" : detail.outlineCap === "square" ? "square" : "round";
  const join = detail.outlineJoin ?? "round";
  return {
    "stroke-linecap": cap,
    "stroke-linejoin": join,
    // `lim` is the miter length over the line width in 1000ths of a percent
    // (Word writes 800000: 8 widths); SVG states the same ratio as a number.
    ...(join === "miter"
      ? { "stroke-miterlimit": (detail.outlineMiterLimit ?? 800_000) / 100_000 }
      : {}),
  };
}

/** One layer of a compound line's mask: a stroke of `width` line widths. */
export interface OfficeCompoundStrokeLayer {
  readonly ink: boolean;
  readonly width: number;
  /** Painted only inside the shape's own outline (its fill region). */
  readonly inside?: boolean;
}

/**
 * A compound line as layered strokes of one mask, widest first: white layers
 * are ink, black layers the gaps. Measured in Word 16.86 (corpus R14-D05,
 * PDF at 770 px/in): Thin Thick is 1/5 line, 1/5 gap, 3/5 line from the
 * outside in; Triple 1/6, 1/6, 1/3, 1/6, 1/6; the line is centred on the
 * outline, so the outer gap shows the page and the inner one the fill.
 *
 * The two uneven pairs need the outline's inside to tell their sides apart.
 * An OPEN path has none, so they are drawn there as the even Double.
 */
export function officeCompoundStrokeLayers(
  compound: OfficeStrokeCompound,
  closed: boolean,
): readonly OfficeCompoundStrokeLayer[] {
  if (compound === "triple") {
    return [
      { ink: true, width: 1 },
      { ink: false, width: 2 / 3 },
      { ink: true, width: 1 / 3 },
    ];
  }
  if (compound === "thinThick" && closed) {
    return [
      { ink: true, width: 1 },
      { ink: false, width: 0.6 },
      { ink: true, width: 0.2 },
      { ink: true, width: 0.6, inside: true },
    ];
  }
  if (compound === "thickThin" && closed) {
    return [
      { ink: true, width: 1 },
      { ink: false, width: 0.6, inside: true },
      { ink: true, width: 0.2 },
    ];
  }
  return [
    { ink: true, width: 1 },
    { ink: false, width: 1 / 3 },
  ];
}
