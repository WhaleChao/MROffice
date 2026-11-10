/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Graphics Format > Graphics Fill / Graphics Outline on an SVG graphic
 * (PPT-A161). PowerPoint applies the colour "to the entire image" (Microsoft
 * Support, "Edit SVG images in Microsoft 365"): it rewrites the SVG part and
 * redraws its PNG beside it. The page keeps the result as the graphic's one
 * fill and one outline -- `graphicFill` / `graphicOutline`, `#rrggbb` or
 * `"none"` -- and a theme swatch's reference in `themeColors`
 * (`lib/presentation-theme-color-binding`). The server writes the SVG part
 * (`office_svg_recolor`); this module only edits the two fields and DRAWS a
 * recoloured copy of the block's `src` for the canvas. The package's parts
 * are never rewritten in the browser.
 */

import type { PresentationImageBlock } from "../components/console/editor/native-document/documentCommands";
import { OFFICE_CHART_LINE_DASH_PATTERN } from "./office-chart-shape-paint";
import { normalizeHexColor } from "./office-color-palette";
import {
  officeThemeRefHex,
  type OfficeThemeColorRef,
  type OfficeThemeGridColors,
  type OfficeThemeGridSlot,
} from "./office-theme-color-grid";

export type OfficeGraphicPaintField = "graphicFill" | "graphicOutline";

/**
 * Graphics Outline > Dashes: PowerPoint's eight, in its order (Solid, Round
 * Dot, Square Dot, Dash, Dash Dot, Long Dash, Long Dash Dot, Long Dash Dot
 * Dot), by DrawingML name -- the names the shared dash icons draw.
 */
export const OFFICE_GRAPHIC_OUTLINE_DASHES = [
  "solid",
  "sysDot",
  "sysDash",
  "dash",
  "dashDot",
  "lgDash",
  "lgDashDot",
  "lgDashDotDot",
] as const;
export type OfficeGraphicOutlineDash = (typeof OFFICE_GRAPHIC_OUTLINE_DASHES)[number];

const EMU_PER_POINT = 12_700;


export interface OfficeGraphicDefaultOutline {
  readonly color: string;
}

/** A picked paint: a colour, or null for No Fill / No Outline. */
export type OfficeGraphicPaintPick = string | null;

const SVG_DATA_PREFIX = "data:image/svg+xml;base64,";
const PAINTED = new Set(["path", "rect", "circle", "ellipse", "line", "polyline", "polygon"]);
const NOT_DRAWN = new Set(["defs", "clipPath", "mask", "linearGradient", "radialGradient", "style"]);

/** The paint a field states: a colour, "none", or undefined (the graphic's own). */
export function officeGraphicPaint(
  block: PresentationImageBlock,
  field: OfficeGraphicPaintField,
): string | undefined {
  const value = (block as Record<string, unknown>)[field];
  if (value === "none") return "none";
  return normalizeHexColor(value) ?? undefined;
}

/**
 * The block with one of Graphics Fill / Outline set (`null` = No Fill / No
 * Outline), or null when nothing changes. A theme swatch's reference is kept
 * by the window's commit (`reconcilePresentationThemeColors`).
 */
export function withOfficeGraphicPaint(
  block: PresentationImageBlock,
  field: OfficeGraphicPaintField,
  pick: OfficeGraphicPaintPick,
): PresentationImageBlock | null {
  if (block.officeGraphicKind !== "svg") return null;
  const value = pick === null ? "none" : normalizeHexColor(pick);
  if (!value || (block as Record<string, unknown>)[field] === value) return null;
  return { ...block, [field]: value } as PresentationImageBlock;
}

/** The graphic's own width in SVG user units (its viewBox, else width). */
export function officeGraphicUserWidth(block: PresentationImageBlock): number | null {
  const text = decodeSvg(block.src);
  const root = text ? /<svg\b[^>]*>/i.exec(text)?.[0] : undefined;
  if (root) {
    const viewBox = /\sviewBox\s*=\s*["']([^"']+)["']/i.exec(root)?.[1];
    const parts = viewBox?.trim().split(/[\s,]+/).map(Number);
    if (parts && parts.length === 4 && parts[2]! > 0) return parts[2]!;
    const width = Number.parseFloat(/\swidth\s*=\s*["']([^"']+)["']/i.exec(root)?.[1] ?? "");
    if (width > 0) return width;
  }
  return typeof block.pixelWidth === "number" && block.pixelWidth > 0 ? block.pixelWidth : null;
}

/** Points per SVG user unit at the picture's current size. */
function pointsPerUnit(block: PresentationImageBlock): number | null {
  const units = officeGraphicUserWidth(block);
  const width = typeof block.width === "number" && block.width > 0 ? block.width : null;
  return units && width ? width / EMU_PER_POINT / units : null;
}

/** The outline's weight in points as Graphics Outline > Weight shows it. */
export function officeGraphicOutlineWeightPt(block: PresentationImageBlock): number | null {
  const outline = officeGraphicPaint(block, "graphicOutline");
  if (!outline || outline === "none") return null;
  const units = typeof block.graphicOutlineWidth === "number" ? block.graphicOutlineWidth : 1;
  const scale = pointsPerUnit(block);
  return scale ? Math.round(units * scale * 100) / 100 : null;
}

/** The outline's dash as Graphics Outline > Dashes shows it. */
export function officeGraphicOutlineDash(
  block: PresentationImageBlock,
): OfficeGraphicOutlineDash | null {
  const outline = officeGraphicPaint(block, "graphicOutline");
  if (!outline || outline === "none") return null;
  const dash = block.graphicOutlineDash;
  return (OFFICE_GRAPHIC_OUTLINE_DASHES as readonly string[]).includes(dash ?? "")
    ? (dash as OfficeGraphicOutlineDash)
    : "solid";
}

/** A graphic with no outline takes PowerPoint's default one for a weight or dash. */
function withVisibleOutline(
  block: PresentationImageBlock,
  fallback: OfficeGraphicDefaultOutline,
): PresentationImageBlock {
  const outline = officeGraphicPaint(block, "graphicOutline");
  if (outline && outline !== "none") return block;
  const color = normalizeHexColor(fallback.color) ?? "#000000";
  return {
    ...block,
    graphicOutline: color,
    themeColors: {
      ...((block.themeColors as Record<string, unknown> | undefined) ?? {}),
      graphicOutline: { slot: "tx1", rgb: color },
    },
  } as PresentationImageBlock;
}

/**
 * Graphics Outline > Weight: `points` at the picture's current size, kept as
 * the SVG's own `stroke-width` (PowerPoint wrote 3 pt on a 96-unit graphic
 * shown 120 pt wide as 2.4). Null when nothing changes or the size is unknown.
 */
export function withOfficeGraphicOutlineWeight(
  block: PresentationImageBlock,
  points: number,
  fallback: OfficeGraphicDefaultOutline,
): PresentationImageBlock | null {
  const scale = pointsPerUnit(block);
  if (block.officeGraphicKind !== "svg" || !scale || !(points > 0)) return null;
  const units = Math.round((points / scale) * 10_000) / 10_000;
  const next = withVisibleOutline(block, fallback);
  if (next === block && block.graphicOutlineWidth === units) return null;
  return { ...next, graphicOutlineWidth: units } as PresentationImageBlock;
}

/** Graphics Outline > Dashes. Null when nothing changes. */
export function withOfficeGraphicOutlineDash(
  block: PresentationImageBlock,
  dash: OfficeGraphicOutlineDash,
  fallback: OfficeGraphicDefaultOutline,
): PresentationImageBlock | null {
  if (block.officeGraphicKind !== "svg") return null;
  const next = withVisibleOutline(block, fallback);
  if (next === block && (block.graphicOutlineDash ?? "solid") === dash) return null;
  return { ...next, graphicOutlineDash: dash } as PresentationImageBlock;
}

function decodeSvg(src: string): string | null {
  if (!src.toLowerCase().startsWith(SVG_DATA_PREFIX)) return null;
  try {
    const binary = atob(src.slice(SVG_DATA_PREFIX.length));
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return null;
  }
}

function encodeSvg(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  return SVG_DATA_PREFIX + btoa(binary);
}

/** An Office class naming the fill (`Fill`) or stroke (`Stroke`) colour. */
function isOfficeClass(name: string, target: "Fill" | "Stroke"): boolean {
  if (name.startsWith("MsftOfcThm_")) {
    return name.endsWith(`_${target}_v2`) || name.endsWith(`_${target}`);
  }
  return name.startsWith(`MsftOfcResponsive_${target}_`);
}

interface OutlineStyle {
  readonly width?: number;
  readonly dash?: OfficeGraphicOutlineDash;
}

function formatLength(value: number): string {
  return String(Math.round(value * 10_000) / 10_000);
}

function paintElement(
  element: Element,
  fill: string | undefined,
  outline: string | undefined,
  style: OutlineStyle,
): void {
  if (style.width !== undefined) element.setAttribute("stroke-width", formatLength(style.width));
  if (style.dash !== undefined) {
    const width = Number.parseFloat(element.getAttribute("stroke-width") ?? "") || style.width || 1;
    if (style.dash === "solid") element.removeAttribute("stroke-dasharray");
    else {
      element.setAttribute(
        "stroke-dasharray",
        OFFICE_CHART_LINE_DASH_PATTERN[style.dash].map((step) => formatLength(step * width)).join(" "),
      );
    }
  }
  const classes = (element.getAttribute("class") ?? "").split(/\s+/).filter(Boolean);
  let kept = classes;
  if (fill !== undefined) {
    element.setAttribute("fill", fill);
    // An Office class rule outranks the attribute; the new colour has none.
    kept = kept.filter((name) => !isOfficeClass(name, "Fill"));
  }
  if (outline !== undefined) {
    element.setAttribute("stroke", outline);
    kept = kept.filter((name) => !isOfficeClass(name, "Stroke"));
  }
  if (kept.length !== classes.length) {
    if (kept.length) element.setAttribute("class", kept.join(" "));
    else element.removeAttribute("class");
  }
}

const PAINTED_CACHE_LIMIT = 64;
const paintedCache = new Map<string, string>();

/**
 * The SVG data URL drawn with Graphics Fill / Outline over the whole image,
 * for display only. The source string comes back when there is nothing to
 * paint, when it is not an SVG data URL, or where no DOM parser exists.
 */
export function officeGraphicPaintedSrc(
  src: string,
  fill: string | undefined,
  outline: string | undefined,
  style: OutlineStyle = {},
): string {
  if (fill === undefined && outline === undefined && style.width === undefined && style.dash === undefined) {
    return src;
  }
  if (typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") return src;
  const key = `${fill ?? ""}|${outline ?? ""}|${style.width ?? ""}|${style.dash ?? ""}|${src}`;
  const cached = paintedCache.get(key);
  if (cached !== undefined) return cached;
  const text = decodeSvg(src);
  if (text === null) return src;
  const document = new DOMParser().parseFromString(text, "image/svg+xml");
  const root = document.documentElement;
  if (!root || root.localName !== "svg" || document.getElementsByTagName("parsererror").length) {
    return src;
  }
  const walk = (element: Element) => {
    for (const child of Array.from(element.children)) {
      if (NOT_DRAWN.has(child.localName)) continue;
      if (PAINTED.has(child.localName)) paintElement(child, fill, outline, style);
      else walk(child);
    }
  };
  walk(root);
  const painted = encodeSvg(new XMLSerializer().serializeToString(document));
  if (paintedCache.size >= PAINTED_CACHE_LIMIT) {
    const oldest = paintedCache.keys().next().value;
    if (oldest !== undefined) paintedCache.delete(oldest);
  }
  paintedCache.set(key, painted);
  return painted;
}

const STRETCHED_CACHE_LIMIT = 64;
const stretchedCache = new Map<string, string>();












export function officeGraphicStretchedSrc(src: string): string {
  if (typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") return src;
  const cached = stretchedCache.get(src);
  if (cached !== undefined) return cached;
  const text = decodeSvg(src);
  if (text === null) return src;
  const document = new DOMParser().parseFromString(text, "image/svg+xml");
  const root = document.documentElement;
  let stretched = src;
  if (
    root &&
    root.localName === "svg" &&
    !document.getElementsByTagName("parsererror").length &&
    root.hasAttribute("viewBox") &&
    !root.hasAttribute("preserveAspectRatio")
  ) {
    root.setAttribute("preserveAspectRatio", "none");
    stretched = encodeSvg(new XMLSerializer().serializeToString(document));
  }
  if (stretchedCache.size >= STRETCHED_CACHE_LIMIT) {
    const oldest = stretchedCache.keys().next().value;
    if (oldest !== undefined) stretchedCache.delete(oldest);
  }
  stretchedCache.set(src, stretched);
  return stretched;
}

/** The block as the canvas draws it: `src` recoloured when it states a paint. */
export function officeGraphicDisplayBlock(block: PresentationImageBlock): PresentationImageBlock {
  if (block.officeGraphicKind !== "svg") return block;
  const fill = officeGraphicPaint(block, "graphicFill");
  const outline = officeGraphicPaint(block, "graphicOutline");
  const style: OutlineStyle = {
    ...(typeof block.graphicOutlineWidth === "number" ? { width: block.graphicOutlineWidth } : {}),
    ...((OFFICE_GRAPHIC_OUTLINE_DASHES as readonly string[]).includes(block.graphicOutlineDash ?? "")
      ? { dash: block.graphicOutlineDash as OfficeGraphicOutlineDash }
      : {}),
  };
  if (fill === undefined && outline === undefined && !style.width && !style.dash) return block;
  const src = officeGraphicPaintedSrc(block.src, fill, outline, style);
  return src === block.src ? block : { ...block, src };
}

/**
 * Graphics Format > Graphics Styles (PPT-A161): presets that set "both a fill
 * color and an outline" over the whole graphic (Microsoft Support, "Edit SVG
 * images in Microsoft 365"). Seven theme columns -- Dark 1 and Accent 1-6, as
 * Office's style galleries run -- by four rows: a colour fill with no
 * outline, a colour outline with no fill, a colour fill with a Dark 1
 * outline, and a light (80% lighter) fill with a colour outline. Each colour
 * keeps its theme reference, so the graphic follows a theme change as a
 * Graphics Fill / Outline swatch does.
 */
export const OFFICE_GRAPHIC_STYLE_ROWS = [
  "fill",
  "outline",
  "fillDarkOutline",
  "lightFillOutline",
] as const;
export type OfficeGraphicStyleRow = (typeof OFFICE_GRAPHIC_STYLE_ROWS)[number];

const OFFICE_GRAPHIC_STYLE_SLOTS = [
  "tx1",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const satisfies readonly OfficeThemeGridSlot[];

export interface OfficeGraphicStylePaint {
  /** `#rrggbb`, or "none" for No Fill / No Outline. */
  readonly color: string;
  readonly ref: OfficeThemeColorRef | null;
}

export interface OfficeGraphicStyle {
  readonly id: string;
  readonly row: OfficeGraphicStyleRow;
  /** 0 = Dark 1, 1-6 = Accent 1-6. */
  readonly column: number;
  readonly fill: OfficeGraphicStylePaint;
  readonly outline: OfficeGraphicStylePaint;
}

const NO_PAINT: OfficeGraphicStylePaint = { color: "none", ref: null };

function themePaint(
  colors: OfficeThemeGridColors,
  ref: OfficeThemeColorRef,
): OfficeGraphicStylePaint {
  return { color: officeThemeRefHex(ref, colors) ?? "#000000", ref };
}

/** The gallery's 28 styles under the deck's theme colours, row by row. */
export function officeGraphicStyles(
  colors: OfficeThemeGridColors,
): readonly OfficeGraphicStyle[] {
  return OFFICE_GRAPHIC_STYLE_ROWS.flatMap((row) =>
    OFFICE_GRAPHIC_STYLE_SLOTS.map((slot, column) => {
      const solid = themePaint(colors, { slot });
      const dark = themePaint(colors, { slot: "tx1" });
      const light = themePaint(colors, { slot, lumMod: 20_000, lumOff: 80_000 });
      const paints: Record<
        OfficeGraphicStyleRow,
        readonly [OfficeGraphicStylePaint, OfficeGraphicStylePaint]
      > = {
        fill: [solid, NO_PAINT],
        outline: [NO_PAINT, solid],
        fillDarkOutline: [solid, dark],
        lightFillOutline: [light, solid],
      };
      const [fill, outline] = paints[row];
      return { id: `${row}-${column}`, row, column, fill, outline };
    }),
  );
}

/** The style the graphic's fill and outline match now, if any. */
export function officeGraphicStyleValue(
  block: PresentationImageBlock,
  styles: readonly OfficeGraphicStyle[],
): string | null {
  const fill = officeGraphicPaint(block, "graphicFill");
  const outline = officeGraphicPaint(block, "graphicOutline");
  if (fill === undefined || outline === undefined) return null;
  return (
    styles.find(
      (style) => style.fill.color === fill && style.outline.color === outline,
    )?.id ?? null
  );
}

/**
 * The block with a Graphics Style applied: its fill and outline, each with
 * its theme reference (or none). Null when nothing changes.
 */
export function withOfficeGraphicStyle(
  block: PresentationImageBlock,
  style: OfficeGraphicStyle,
): PresentationImageBlock | null {
  if (block.officeGraphicKind !== "svg") return null;
  const record = block as Record<string, unknown>;
  const bindings: Record<string, unknown> = {
    ...((record.themeColors as Record<string, unknown> | undefined) ?? {}),
  };
  for (const [field, paint] of [
    ["graphicFill", style.fill],
    ["graphicOutline", style.outline],
  ] as const) {
    if (paint.ref && paint.color !== "none") {
      bindings[field] = { ...paint.ref, rgb: paint.color };
    } else {
      delete bindings[field];
    }
  }
  if (
    record.graphicFill === style.fill.color &&
    record.graphicOutline === style.outline.color &&
    JSON.stringify(record.themeColors ?? {}) === JSON.stringify(bindings)
  ) {
    return null;
  }
  const next: Record<string, unknown> = {
    ...block,
    graphicFill: style.fill.color,
    graphicOutline: style.outline.color,
  };
  if (Object.keys(bindings).length) next.themeColors = bindings;
  else delete next.themeColors;
  return next as PresentationImageBlock;
}
