/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import data from "./native.json";
import {
  officePictureStyle,
  type OfficePictureStyleId,
} from "./catalog";

/**
 * One Office picture style as the renderer reads it (`geometry` .. `threeD`).
 * The package recipe a writer emits (the complete `spPr` DrawingML) is the
 * server's (`apps/api/app/services/file_adapters/data/office-picture-styles.json`,
 * read and written by `office_picture_styles.py`); the browser keeps what it
 * draws and `outlineIdentity`: the one Picture Border that redraws the style's
 * own outline exactly (`"none"` for a style without one, `null` when no
 * border can, e.g. a recipe line without a dash preset).
 */
export interface NativePictureStyle {
  id: OfficePictureStyleId;
  geometry: string;
  adjustments: Record<string, number>;
  line?: { color: string; opacity: number; width: number; cmpd?: string };
  fill?: { color: string; opacity: number };
  effects?: Array<Record<string, string | number>>;
  camera?: Record<string, string>;
  threeD?: {
    bevel?: Record<string, string>;
    contour?: { color: string; opacity: number };
    contourW?: string;
    extrusion?: { color: string; opacity: number };
    extrusionH?: string;
  };
  outlineIdentity: OfficePictureStyleOutline | null;
}

const styles = new Map(
  (data.styles as unknown as NativePictureStyle[]).map((style) => [
    style.id,
    style,
  ]),
);

export function nativePictureStyle(id: unknown): NativePictureStyle | null {
  return typeof id === "string"
    ? (styles.get(id as OfficePictureStyleId) ?? null)
    : null;
}

/**
 * A Picture Border set over a style (Word keeps the style and replaces only
 * the outline's paint, width and dash), or `"none"` for "No Outline".
 */
export type OfficePictureStyleOutline =
  | "none"
  | {
      readonly widthEmu: number;
      /** `#rrggbb` */
      readonly color: string;
      readonly dash: "solid" | "dashed" | "dotted";
    };

function sameOutline(
  left: OfficePictureStyleOutline,
  right: OfficePictureStyleOutline,
): boolean {
  if (left === "none" || right === "none") return left === right;
  return (
    Math.round(left.widthEmu) === Math.round(right.widthEmu) &&
    left.color.toLowerCase() === right.color.toLowerCase() &&
    left.dash === right.dash
  );
}

/** Null when the override draws exactly the style's own outline. */
export function canonicalOfficePictureStyleOutline(
  id: unknown,
  outline: OfficePictureStyleOutline | null | undefined,
): OfficePictureStyleOutline | null {
  const style = officePictureStyle(id) ? nativePictureStyle(id) : null;
  if (!style || !outline) return null;
  return style.outlineIdentity && sameOutline(style.outlineIdentity, outline)
    ? null
    : outline;
}

export function pictureStylePath(
  style: NativePictureStyle,
  w: number,
  h: number,
  inset = 0,
): string {
  const x = inset,
    y = inset,
    right = w - inset,
    bottom = h - inset,
    small = Math.min(w, h);
  const r = Math.max(
    0,
    Math.min(
      (small * (style.adjustments.adj ?? 16667)) / 100000,
      (right - x) / 2,
      (bottom - y) / 2,
    ),
  );
  if (style.geometry === "ellipse")
    return `M${x},${h / 2} A${(right - x) / 2},${(bottom - y) / 2} 0 1 0 ${right},${h / 2} A${(right - x) / 2},${(bottom - y) / 2} 0 1 0 ${x},${h / 2} Z`;
  if (style.geometry === "roundRect")
    return `M${x + r},${y} H${right - r} Q${right},${y} ${right},${y + r} V${bottom - r} Q${right},${bottom} ${right - r},${bottom} H${x + r} Q${x},${bottom} ${x},${bottom - r} V${y + r} Q${x},${y} ${x + r},${y} Z`;
  // The two-corner presets' adjust values (ECMA-376 presetShapeDefinitions):
  // round2DiagRect rounds top-left/bottom-right by adj1 (16667) and the other
  // two by adj2 (0); snip2DiagRect snips top-left/bottom-right by adj1 (0) and
  // the other two by adj2 (16667); round2SameRect rounds the top two by adj1
  // (16667) and the bottom two by adj2 (0). Lengths are of the shorter side.
  const corner = (value: number | undefined, fallback: number) =>
    Math.max(0, Math.min((small * (value ?? fallback)) / 100000, (right - x) / 2, (bottom - y) / 2));
  if (style.geometry === "round2DiagRect") {
    const a = corner(style.adjustments.adj1 ?? style.adjustments.adj, 16667);
    const b = corner(style.adjustments.adj2, 0);
    return `M${x + a},${y} H${right - b} Q${right},${y} ${right},${y + b} V${bottom - a} Q${right},${bottom} ${right - a},${bottom} H${x + b} Q${x},${bottom} ${x},${bottom - b} V${y + a} Q${x},${y} ${x + a},${y} Z`;
  }
  if (style.geometry === "snip2DiagRect") {
    const a = corner(style.adjustments.adj1, 0);
    const b = corner(style.adjustments.adj2, 16667);
    return `M${x + a},${y} H${right - b} L${right},${y + b} V${bottom - a} L${right - a},${bottom} H${x + b} L${x},${bottom - b} V${y + a} Z`;
  }
  if (style.geometry === "round2SameRect") {
    const a = corner(style.adjustments.adj1, 16667);
    const b = corner(style.adjustments.adj2, 0);
    return `M${x + a},${y} H${right - a} Q${right},${y} ${right},${y + a} V${bottom - b} Q${right},${bottom} ${right - b},${bottom} H${x + b} Q${x},${bottom} ${x},${bottom - b} V${y + a} Q${x},${y} ${x + a},${y} Z`;
  }
  return `M${x},${y} H${right} V${bottom} H${x} Z`;
}

/**
 * The Picture Effects a style draws (Shadow, Reflection, Glow, Soft Edges,
 * Bevel, 3-D Rotation): the effects whose "No <effect>" must take the
 * style's own off.
 */
export function officePictureStyleEffectKinds(id: unknown): ReadonlySet<string> {
  const style = nativePictureStyle(id);
  const kinds = new Set<string>();
  if (!style) return kinds;
  for (const effect of style.effects ?? []) {
    if (effect.kind === "outerShdw" || effect.kind === "innerShdw") kinds.add("shadow");
    else if (effect.kind === "reflection" || effect.kind === "glow" || effect.kind === "softEdge") {
      kinds.add(String(effect.kind));
    }
  }
  if (style.threeD?.bevel) kinds.add("bevel");
  if (style.camera && style.camera.prst !== "orthographicFront") kinds.add("rotation3d");
  return kinds;
}
