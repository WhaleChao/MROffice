/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";

/**
 * PDF text-markup styles — the four standard Acrobat text markups.
 *
 * - "highlight" fills the selection rect (a translucent wash).
 * - "underline" draws a straight line along the baseline.
 * - "strikethrough" draws a straight line through the middle.
 * - "squiggly" draws a wavy (zig-zag) line along the baseline — Acrobat's
 *   "squiggly underline" markup.
 *
 * All four reuse the same selection rects — only the CSS / baked band differ.
 */
export type HighlightMarkupStyle =
  | "highlight"
  | "underline"
  | "strikethrough"
  | "squiggly";

export const HIGHLIGHT_MARKUP_STYLES: readonly HighlightMarkupStyle[] = [
  "highlight",
  "underline",
  "strikethrough",
  "squiggly",
];

export function isHighlightMarkupStyle(value: unknown): value is HighlightMarkupStyle {
  return (
    value === "highlight" ||
    value === "underline" ||
    value === "strikethrough" ||
    value === "squiggly"
  );
}

/**
 * A repeating wavy band rendered with two diagonal gradients, drawn at the
 * bottom of the rect so it reads as a squiggly underline (the SVG-free CSS
 * equivalent of Acrobat's squiggly markup). Tile width tracks {@link color}
 * only; the band height is a fraction of an em so it scales with the text.
 */
function squigglyBackgroundImage(color: string): string {
  // Two 45°/135° gradients offset by half a tile form a continuous zig-zag.
  return [
    `linear-gradient(135deg, transparent 0, transparent 40%, ${color} 40%, ${color} 60%, transparent 60%)`,
    `linear-gradient(45deg, transparent 0, transparent 40%, ${color} 40%, ${color} 60%, transparent 60%)`,
  ].join(", ");
}

/**
 * CSS for one selection rect given its color and markup style. Underline,
 * strikethrough, and squiggly keep the background transparent (so the text
 * stays readable) and render a thin line via box-shadow / a centered gradient
 * band / a repeating wavy band.
 */
export function highlightMarkupRectStyle(
  color: string,
  style: HighlightMarkupStyle = "highlight",
): CSSProperties {
  if (style === "underline") {
    return { background: "transparent", boxShadow: `inset 0 -0.12em 0 ${color}` };
  }
  if (style === "strikethrough") {
    return {
      background: "transparent",
      backgroundImage: `linear-gradient(${color}, ${color})`,
      backgroundRepeat: "no-repeat",
      backgroundPosition: "center",
      backgroundSize: "100% 0.12em",
    };
  }
  if (style === "squiggly") {
    // A wavy band pinned to the bottom of the rect. The tile is square-ish
    // (height drives the wave amplitude); it repeats horizontally only.
    return {
      background: "transparent",
      backgroundImage: squigglyBackgroundImage(color),
      backgroundRepeat: "repeat-x",
      backgroundPosition: "left bottom",
      backgroundSize: "0.4em 0.22em",
    };
  }
  return { background: color };
}

/** A point in whichever plane a quad is expressed in. */
export interface HighlightQuadPoint {
  x: number;
  y: number;
}










export type HighlightQuad = readonly number[];

/** Values per quad. */
export const HIGHLIGHT_QUAD_VALUES = 8;

/**
 * Validate optional persisted quads against their rects. `undefined` means
 * "none" (an axis-aligned legacy highlight); `null` means the value is
 * present but malformed, which a payload reader must refuse rather than drop.
 */
export function normalizeHighlightQuads(
  value: unknown,
  rectCount: number,
): number[][] | undefined | null {
  if (value === undefined) return undefined;
  if (!Array.isArray(value) || value.length !== rectCount) return null;
  const quads: number[][] = [];
  for (const quad of value) {
    if (!Array.isArray(quad) || quad.length !== HIGHLIGHT_QUAD_VALUES) {
      return null;
    }
    const numbers = quad.map(Number);
    if (numbers.some((n) => !Number.isFinite(n) || n < 0 || n > 100)) {
      return null;
    }
    quads.push(numbers);
  }
  return quads;
}

/** The quad's four vertices, in its stored order. */
export function highlightQuadVertices(
  quad: HighlightQuad,
): [HighlightQuadPoint, HighlightQuadPoint, HighlightQuadPoint, HighlightQuadPoint] {
  return [
    { x: quad[0], y: quad[1] },
    { x: quad[2], y: quad[3] },
    { x: quad[4], y: quad[5] },
    { x: quad[6], y: quad[7] },
  ];
}

/** True when the quad is (to `tolerance`) an axis-aligned rectangle. */
export function isAxisAlignedHighlightQuad(
  quad: HighlightQuad,
  tolerance = 1e-3,
): boolean {
  const xs = [quad[0], quad[2], quad[4], quad[6]];
  const ys = [quad[1], quad[3], quad[5], quad[7]];
  const onTwo = (values: number[]) => {
    const low = Math.min(...values);
    const high = Math.max(...values);
    return values.every(
      (v) => Math.abs(v - low) <= tolerance || Math.abs(v - high) <= tolerance,
    );
  };
  return onTwo(xs) && onTwo(ys);
}

/** Axis-aligned bounds of a quad. */
export function highlightQuadBounds(quad: HighlightQuad): {
  x: number;
  y: number;
  width: number;
  height: number;
} {
  const xs = [quad[0], quad[2], quad[4], quad[6]];
  const ys = [quad[1], quad[3], quad[5], quad[7]];
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
}

/**
 * The quad as a simple (non-self-intersecting) outline. Producers disagree on
 * vertex order (the spec's figure is a Z, its prose says counter-clockwise),
 * so the outline orders the vertices around their centroid instead of
 * trusting either, which never draws a bow-tie.
 */
export function highlightQuadOutline(quad: HighlightQuad): HighlightQuadPoint[] {
  const vertices = highlightQuadVertices(quad);
  const cx = vertices.reduce((sum, v) => sum + v.x, 0) / 4;
  const cy = vertices.reduce((sum, v) => sum + v.y, 0) / 4;
  return [...vertices].sort(
    (a, b) => Math.atan2(a.y - cy, a.x - cx) - Math.atan2(b.y - cy, b.x - cx),
  );
}

function lerpPoint(
  a: HighlightQuadPoint,
  b: HighlightQuadPoint,
  t: number,
): HighlightQuadPoint {
  return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
}

/** Band thickness of an underline / strikeout, as a fraction of the quad height. */
const QUAD_BAND_FRACTION = 0.08;

/**
 * The painted geometry of one quad for a markup style, in the quad's own
 * plane: a filled outline for a highlight, a thin filled band along the
 * baseline (underline) or through the middle (strikeout), or a zig-zag
 * polyline along the baseline (squiggly). Bands follow the quad's side edges,
 * so they stay parallel to slanted text.
 */
export function highlightQuadMarkupGeometry(
  quad: HighlightQuad,
  style: HighlightMarkupStyle = "highlight",
):
  | { kind: "fill"; points: HighlightQuadPoint[] }
  | { kind: "stroke"; points: HighlightQuadPoint[] } {
  if (style === "highlight") {
    return { kind: "fill", points: highlightQuadOutline(quad) };
  }
  // Stored order: top-left, top-right, bottom-left, bottom-right.
  const [topLeft, topRight, bottomLeft, bottomRight] =
    highlightQuadVertices(quad);
  const along = (t: number) => ({
    left: lerpPoint(bottomLeft, topLeft, t),
    right: lerpPoint(bottomRight, topRight, t),
  });
  if (style === "underline" || style === "strikethrough") {
    const from = style === "underline" ? 0 : 0.5 - QUAD_BAND_FRACTION / 2;
    const low = along(from);
    const high = along(from + QUAD_BAND_FRACTION);
    return {
      kind: "fill",
      points: [low.left, low.right, high.right, high.left],
    };
  }
  // Squiggly: a zig-zag between 0 and 2 band-heights above the baseline,
  // its period about half the line height.
  const base = along(0);
  const crest = along(QUAD_BAND_FRACTION * 2);
  const length = Math.hypot(base.right.x - base.left.x, base.right.y - base.left.y);
  const height = Math.hypot(topLeft.x - bottomLeft.x, topLeft.y - bottomLeft.y);
  const halfPeriods = Math.max(
    2,
    Math.min(400, Math.round(height > 0 ? length / (height * 0.25) : 2)),
  );
  const points: HighlightQuadPoint[] = [];
  for (let step = 0; step <= halfPeriods; step += 1) {
    const t = step / halfPeriods;
    points.push(
      step % 2 === 0
        ? lerpPoint(base.left, base.right, t)
        : lerpPoint(crest.left, crest.right, t),
    );
  }
  return { kind: "stroke", points };
}

/**
 * A quad markup laid out for the reader's DOM: the element box is the quad's
 * page-percent bounds, `clipPath` confines painting AND hit-testing to the
 * quad itself, and `paint` is the style geometry in that box's local 0..100
 * space (an SVG `viewBox="0 0 100 100"` with `preserveAspectRatio="none"`).
 */
export function highlightQuadLayout(
  quad: HighlightQuad,
  style: HighlightMarkupStyle = "highlight",
): {
  bounds: { x: number; y: number; width: number; height: number };
  clipPath: string;
  paint: { kind: "fill" | "stroke"; points: string };
} | null {
  const bounds = highlightQuadBounds(quad);
  if (!(bounds.width > 0) || !(bounds.height > 0)) return null;
  const local = (point: HighlightQuadPoint) => ({
    x: ((point.x - bounds.x) / bounds.width) * 100,
    y: ((point.y - bounds.y) / bounds.height) * 100,
  });
  const round = (n: number) => Math.round(n * 1000) / 1000;
  const clipPath = `polygon(${highlightQuadOutline(quad)
    .map(local)
    .map((p) => `${round(p.x)}% ${round(p.y)}%`)
    .join(", ")})`;
  const geometry = highlightQuadMarkupGeometry(quad, style);
  return {
    bounds,
    clipPath,
    paint: {
      kind: geometry.kind,
      points: geometry.points
        .map(local)
        .map((p) => `${round(p.x)},${round(p.y)}`)
        .join(" "),
    },
  };
}
