/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import type { CSSProperties } from "react";
import { smartArtPictureSrc } from "../office-smartart/picture";

/** The graphic's background as the sheet draws it: a resolved colour or a picture. */
export type SheetSmartArtBackground = { fill: string; alpha?: number } | { src: string };

/** The effects a graphic's background takes: the effect-list presets (no Bevel / 3-D, which `dgm:bg` cannot hold). */
export const SMARTART_GRAPHIC_EFFECTS = ["shadow", "reflection", "glow", "softEdge"] as const;
export type SmartArtGraphicEffect = (typeof SMARTART_GRAPHIC_EFFECTS)[number];

/** The graphic's own line (`dgm:whole/a:ln`): a solid colour and a weight in points. */
export interface SmartArtGraphicLine {
  readonly color: string;
  readonly widthPt: number;
}

/** Everything the graphic itself wears; every part optional. */
export interface SmartArtGraphicFormat {
  readonly background: SheetSmartArtBackground | null;
  readonly line: SmartArtGraphicLine | null;
  readonly effect: SmartArtGraphicEffect | null;
}

const HEX = /^#[0-9a-f]{6}$/i;
/** Office's line weights stop at 1584 pt; a graphic's line is a frame border. */
const MAX_LINE_PT = 1584;

function record(value: unknown): Record<string, unknown> | null {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

/** A line value, checked; null for none or anything this reader does not trust. */
export function readSmartArtGraphicLine(value: unknown): SmartArtGraphicLine | null {
  const line = record(value);
  if (!line || typeof line.color !== "string" || !HEX.test(line.color)) return null;
  const width = typeof line.widthPt === "number" && Number.isFinite(line.widthPt) ? line.widthPt : 0.75;
  if (width <= 0) return null;
  return { color: line.color.toLowerCase(), widthPt: Math.min(MAX_LINE_PT, width) };
}

/** An effect value, checked. */
export function readSmartArtGraphicEffect(value: unknown): SmartArtGraphicEffect | null {
  return typeof value === "string" && (SMARTART_GRAPHIC_EFFECTS as readonly string[]).includes(value)
    ? (value as SmartArtGraphicEffect)
    : null;
}

/** The background a member carries, or `null` for none / one this reader does not trust. */
export function readSheetSmartArtBackground(member: unknown): SheetSmartArtBackground | null {
  const value = record(record(member)?.smartArtBackground);
  if (!value) return null;
  const { fill, alpha, src } = value;
  if (typeof fill === "string" && HEX.test(fill)) {
    return typeof alpha === "number" && alpha >= 0 && alpha < 1 ? { fill, alpha } : { fill };
  }
  const picture = smartArtPictureSrc(src);
  return picture ? { src: picture } : null;
}

/** Fill, line and effect a member carries for its graphic, or `null` when it wears none of them. */
export function readSmartArtGraphicFormat(member: unknown): SmartArtGraphicFormat | null {
  const value = record(record(member)?.smartArtBackground);
  if (!value) return null;
  const background = readSheetSmartArtBackground(member);
  const line = readSmartArtGraphicLine(value.line);
  const effect = readSmartArtGraphicEffect(value.effect);
  return background || line || effect ? { background, line, effect } : null;
}

/** One change to the graphic's own formatting; `null` removes that part (No Fill, No Outline, No Effect). */
export interface SmartArtGraphicFormatChange {
  readonly fill?: string | null;
  /** A colour (keeping the weight), a weight (keeping the colour), or null for No Outline. */
  readonly line?: { readonly color?: string; readonly widthPt?: number } | null;
  readonly effect?: SmartArtGraphicEffect | null;
}

/**
 * The `smartArtBackground` record after one change, or `undefined` when the
 * graphic is left wearing nothing (the field then comes off). Parts the change
 * does not name are kept as they are.
 */
export function smartArtGraphicFormatRecord(
  current: unknown,
  change: SmartArtGraphicFormatChange,
): Record<string, unknown> | undefined {
  const base = { ...(record(current) ?? {}) };
  if (change.fill !== undefined) {
    delete base.fill;
    delete base.alpha;
    delete base.src;
    if (change.fill && HEX.test(change.fill)) base.fill = change.fill.toLowerCase();
  }
  if (change.line !== undefined) {
    if (change.line === null) {
      delete base.line;
    } else {
      const was = readSmartArtGraphicLine(base.line);
      const color = change.line.color && HEX.test(change.line.color) ? change.line.color.toLowerCase() : was?.color ?? "#000000";
      const widthPt = typeof change.line.widthPt === "number" && change.line.widthPt > 0
        ? Math.min(MAX_LINE_PT, change.line.widthPt)
        : was?.widthPt ?? 0.75;
      base.line = { color, widthPt };
    }
  }
  if (change.effect !== undefined) {
    if (change.effect) base.effect = change.effect;
    else delete base.effect;
  }
  return readSmartArtGraphicFormat({ smartArtBackground: base }) ? base : undefined;
}

/**
 * What Reset Graphic does to the graphic's own formatting, in the three
 * editors: Microsoft Support's "Restore the default layout and color of your
 * SmartArt graphic" lists "Background color applied to the entire SmartArt
 * graphic" among the customizations NOT removed, and line and effects among
 * those removed. So the fill (a colour or a picture) stays, the rest comes off
 * (live Word audit WUA 260).
 */
export const SMARTART_RESET_GRAPHIC_FORMAT: SmartArtGraphicFormatChange = { line: null, effect: null };

const PX_PER_PT = 96 / 72;

/**
 * The CSS that draws a graphic's own formatting across its frame, under its
 * shapes: the fill, the line on the frame's edge, and the effect on what that
 * paints (an effect on a graphic with neither fill nor line shows nothing, as
 * an effect on an unfilled, unlined shape does). `null` for none.
 */
export function smartArtGraphicFormatStyle(format: SmartArtGraphicFormat | null): CSSProperties | null {
  if (!format) return null;
  const style: Record<string, string | number> = {};
  const background = format.background;
  if (background) {
    if ("src" in background) {
      // `a:blipFill/a:stretch`: the picture is stretched to the frame.
      style.backgroundImage = `url("${background.src}")`;
      style.backgroundSize = "100% 100%";
      style.backgroundRepeat = "no-repeat";
    } else if (background.alpha === undefined) {
      style.backgroundColor = background.fill;
    } else {
      const [r, g, b] = [1, 3, 5].map((index) => parseInt(background.fill.slice(index, index + 2), 16));
      style.backgroundColor = `rgba(${r}, ${g}, ${b}, ${background.alpha})`;
    }
  }
  if (format.line) {
    const width = Math.max(0.5, format.line.widthPt * PX_PER_PT);
    // `a:ln` is centred on the frame's edge.
    style.outline = `${width}px solid ${format.line.color}`;
    style.outlineOffset = `${-width / 2}px`;
  }
  switch (format.effect) {
    case "shadow":
      // The Shadow preset's `a:outerShdw` (blur 4 pt, 3 pt at 45°, 35 % black).
      style.filter = "drop-shadow(2.8px 2.8px 2.7px rgba(0, 0, 0, 0.35))";
      break;
    case "glow":
      // The Glow preset's `a:glow` (5 pt, accent blue at 72 %).
      style.filter = "drop-shadow(0 0 3.3px rgba(59, 130, 246, 0.72)) drop-shadow(0 0 3.3px rgba(59, 130, 246, 0.72))";
      break;
    case "softEdge": {
      // The Soft Edges preset's `a:softEdge` (2 pt): every edge fades out.
      const fade = "linear-gradient(to right, transparent, #000 2.7px, #000 calc(100% - 2.7px), transparent)";
      const fadeDown = "linear-gradient(to bottom, transparent, #000 2.7px, #000 calc(100% - 2.7px), transparent)";
      style.maskImage = `${fade}, ${fadeDown}`;
      style.WebkitMaskImage = `${fade}, ${fadeDown}`;
      style.maskComposite = "intersect";
      style.WebkitMaskComposite = "source-in";
      break;
    }
    case "reflection":
      // The Reflection preset's `a:reflection` (a half-strength mirror below).
      style.WebkitBoxReflect = "below 2.7px linear-gradient(transparent 50%, rgba(0, 0, 0, 0.5))";
      break;
    default:
      break;
  }
  return Object.keys(style).length ? (style as CSSProperties) : null;
}

/** The CSS that paints a member's graphic formatting across the frame, or `null`. */
export function sheetSmartArtBackgroundStyle(member: unknown): CSSProperties | null {
  return smartArtGraphicFormatStyle(readSmartArtGraphicFormat(member));
}
