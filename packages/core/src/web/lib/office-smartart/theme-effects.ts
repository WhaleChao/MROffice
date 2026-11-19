/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The theme's effect styles a SmartArt quick style names (`a:effectRef idx`).
 *
 * A quick style does not carry its shadows: its labels name an entry of the
 * host theme's `a:fmtScheme/a:effectStyleLst` (ECMA-376 §20.1.4.1.12, one-based
 * `idx`; 0 is no effect), and Office draws that entry. 3-D Polished (`3d2`)
 * names the second: the 2007 Office theme's is a shadow (blur 40000, distance
 * 23000 EMU, straight down, black at 35 %), the 2013+ Office theme's is empty.
 * A converted diagram that drew the style's own table instead lost the shadow
 * Word drew on a 2007-theme document, and its export wrote an empty
 * `a:effectLst`.
 *
 * Shared by the three hosts (Docs, Presentation, Spreadsheet), each from its
 * own theme: an Office effect scheme by id (`OFFICE_EFFECT_SCHEMES`), or the
 * file's own `a:fmtScheme` as it was read (`nativeEffects.xml`).
 */

import { officeTransformedHex } from "../office-color-transform";
import { officeEffectGalleryScheme } from "../office-theme-catalog";
import { emuToPx } from "../office-units";
import { child, childrenNamed, descendants, parseXml, type XmlElement } from "./ms-engine/xml";
import type { SmartArtFillGradient, SmartArtGradientPath, SmartArtShapeShadow } from "./styles-model";

/** One entry per `a:effectStyle`, in order: its outer shadow, or null for none. */
export type SmartArtThemeEffectShadows = readonly (SmartArtShapeShadow | null)[];

/** Where a host's theme effects come from. */
export interface SmartArtThemeEffectSource {
  /** An Office effect scheme (`OFFICE_EFFECT_SCHEMES` id; "none" is Office's own). */
  readonly id?: string | null;
  /** The file's own `a:fmtScheme` (a theme Office does not ship), as read. */
  readonly xml?: string | null;
}

const MAX_FMT_SCHEME_CHARS = 64_000;

function hexColor(element: XmlElement | undefined): { color: string; alpha: number } | null {
  if (!element) return null;
  let color: string | null = null;
  if (element.name === "srgbClr" && /^[0-9a-fA-F]{6}$/.test(element.attrs.val ?? "")) color = `#${element.attrs.val!.toLowerCase()}`;
  else if (element.name === "prstClr" && element.attrs.val === "black") color = "#000000";
  else if (element.name === "schemeClr" && (element.attrs.val === "phClr" || element.attrs.val === "dk1" || element.attrs.val === "tx1")) color = "#000000";
  if (!color) return null;
  const alpha = child(element, "alpha");
  const value = alpha ? Number(alpha.attrs.val) : 100000;
  return { color, alpha: Number.isFinite(value) ? Math.min(1, Math.max(0, value / 100000)) : 1 };
}

function shadowOf(effectStyle: XmlElement): SmartArtShapeShadow | null {
  const list = child(effectStyle, "effectLst");
  const outer = list ? child(list, "outerShdw") : undefined;
  if (!outer) return null;
  const paint = hexColor(outer.children[0]);
  if (!paint) return null;
  const blur = Number(outer.attrs.blurRad ?? "0");
  const distance = Number(outer.attrs.dist ?? "0");
  return {
    color: paint.color,
    blur: emuToPx(Number.isFinite(blur) ? blur : 0),
    distance: emuToPx(Number.isFinite(distance) ? distance : 0),
    opacity: Math.round(paint.alpha * 100) / 100,
  };
}

/** The theme's effect styles' shadows, or undefined when the source names none it can read. */
export function smartArtThemeEffectShadows(
  source: SmartArtThemeEffectSource | null | undefined,
): SmartArtThemeEffectShadows | undefined {
  if (!source) return undefined;
  if (typeof source.xml === "string" && source.xml && source.xml.length <= MAX_FMT_SCHEME_CHARS) {
    try {
      const root = parseXml(source.xml);
      const lists = root.name === "effectStyleLst" ? [root] : descendants(root, "effectStyleLst");
      const styles = childrenNamed(lists[0], "effectStyle");
      if (styles.length) return Object.freeze(styles.map(shadowOf));
    } catch {
      // An unreadable scheme names nothing; the id (if any) still does.
    }
  }
  const scheme = officeEffectGalleryScheme(source.id ?? undefined);
  if (!scheme) return undefined;
  return Object.freeze(
    scheme.effects.map((effect) =>
      effect.shadow
        ? {
            color: "#000000",
            blur: emuToPx(effect.shadow.blur),
            distance: emuToPx(effect.shadow.distance),
            opacity: Math.round(effect.shadow.alpha / 1000) / 100,
          }
        : null,
    ),
  );
}












export type SmartArtThemeFillStyle =
  | { readonly kind: "solid" }
  | {
      readonly kind: "gradient";
      /** CSS degrees, the `SmartArtFillGradient` unit (DrawingML `a:lin ang` + 90). */
      readonly angle: number;
      readonly stops: readonly { readonly position: number; readonly mods: readonly (readonly [string, number])[] }[];
      readonly scaled?: boolean;
      readonly rotWithShape?: boolean;
      /** A path gradient's shading and focus box; its `angle` is then unused (90). */
      readonly path?: SmartArtGradientPath;
    }
  | null;

export type SmartArtThemeFillStyles = readonly SmartArtThemeFillStyle[];

const FILL_STOP_MODS = new Set(["tint", "shade", "lumMod", "lumOff", "satMod", "satOff", "hueMod", "hueOff"]);

function booleanAttr(value: string | undefined): boolean | undefined {
  if (value === "1" || value === "true") return true;
  if (value === "0" || value === "false") return false;
  return undefined;
}

const PATH_KINDS = new Set(["circle", "rect", "shape"]);
const RECT_EDGES = ["l", "t", "r", "b"] as const;

/** `a:path` as the shared paint states it (`OfficeShapeGradient.path`), or null when unreadable. */
function pathOf(element: XmlElement): SmartArtGradientPath | null {
  const kind = element.attrs.path ?? "";
  if (!PATH_KINDS.has(kind)) return null;
  const toRect = child(element, "fillToRect");
  if (!toRect) return { kind: kind as SmartArtGradientPath["kind"] };
  const rect: { l?: number; t?: number; r?: number; b?: number } = {};
  for (const edge of RECT_EDGES) {
    const raw = toRect.attrs[edge];
    if (raw === undefined) continue;
    const value = Number(raw);
    if (!Number.isInteger(value) || Math.abs(value) > 1_000_000) return null;
    rect[edge] = value;
  }
  return { kind: kind as SmartArtGradientPath["kind"], fillToRect: rect };
}

function fillStyleOf(fill: XmlElement): SmartArtThemeFillStyle {
  if (fill.name === "solidFill") return { kind: "solid" };
  if (fill.name !== "gradFill") return null;
  const list = child(fill, "gsLst");
  const linear = child(fill, "lin");
  const pathElement = child(fill, "path");
  if (!list || child(fill, "tileRect") || (pathElement && linear)) return null;
  const path = pathElement ? pathOf(pathElement) : undefined;
  if (path === null) return null;
  const stops: { position: number; mods: (readonly [string, number])[] }[] = [];
  for (const stop of childrenNamed(list, "gs")) {
    const position = Number(stop.attrs.pos ?? "");
    const color = stop.children[0];
    // A stop on the shape's own colour; any other colour is not the style's.
    if (!Number.isFinite(position) || !color || color.name !== "schemeClr" || color.attrs.val !== "phClr") return null;
    const mods: (readonly [string, number])[] = [];
    for (const mod of color.children) {
      const value = Number(mod.attrs.val ?? "");
      if (!FILL_STOP_MODS.has(mod.name) || !Number.isFinite(value)) return null;
      mods.push([mod.name, value]);
    }
    stops.push({ position: Math.min(1, Math.max(0, position / 100000)), mods });
  }
  if (stops.length < 2) return null;
  const angle = Number(linear?.attrs.ang ?? "0");
  if (!Number.isFinite(angle)) return null;
  const scaled = booleanAttr(linear?.attrs.scaled);
  const rotWithShape = booleanAttr(fill.attrs.rotWithShape);
  return {
    kind: "gradient",
    angle: (((Math.round(angle / 60000) + 90) % 360) + 360) % 360,
    stops,
    ...(scaled !== undefined ? { scaled } : {}),
    ...(rotWithShape !== undefined ? { rotWithShape } : {}),
    ...(path ? { path } : {}),
  };
}

/**
 * The theme's fill styles as a quick style's `a:fillRef idx` draws them, from
 * the file's own `a:fmtScheme` (an Office scheme by id names only the kinds of
 * its fills, so its gradients stay the style's table); undefined when the
 * source states none.
 */
export function smartArtThemeFillStyles(
  source: SmartArtThemeEffectSource | null | undefined,
): SmartArtThemeFillStyles | undefined {
  if (!source || typeof source.xml !== "string" || !source.xml || source.xml.length > MAX_FMT_SCHEME_CHARS) return undefined;
  try {
    const root = parseXml(source.xml);
    const lists = root.name === "fillStyleLst" ? [root] : descendants(root, "fillStyleLst");
    if (!lists.length) return undefined;
    const fills = lists[0].children;
    return fills.length ? Object.freeze(fills.map(fillStyleOf)) : undefined;
  } catch {
    return undefined;
  }
}

/**
 * A theme fill style's gradient on a shape's colour, or undefined when the
 * style is not one this reader draws (the style's own table then).
 */
export function smartArtThemeFillGradient(
  style: SmartArtThemeFillStyle | undefined,
  color: string,
): SmartArtFillGradient | null | undefined {
  if (style === undefined || style === null) return undefined;
  if (style.kind === "solid") return null;
  const stops: { position: number; color: string }[] = [];
  for (const stop of style.stops) {
    const hex = officeTransformedHex(color, stop.mods);
    if (!hex) return undefined;
    stops.push({ position: stop.position, color: hex });
  }
  return {
    angle: style.angle,
    stops,
    ...(style.scaled !== undefined ? { scaled: style.scaled } : {}),
    // Office's diagram engine draws the ramp unturned whatever the theme's
    // style states: Word's own drawing of SmartArtWordSaved's 3-D Polished
    // writes `a:gradFill rotWithShape="0"` under the 2007 theme's
    // `rotWithShape="1"` fill style, as every SmartArt drawing in the corpus
    // does; the drawing part repeats what Word wrote.
    rotWithShape: false,
    ...(style.path ? { path: { ...style.path, ...(style.path.fillToRect ? { fillToRect: { ...style.path.fillToRect } } : {}) } } : {}),
  };
}

/**
 * Which theme a converted diagram draws its quick style through: the host's
 * Design > Effects pick when it names one the file's theme is not, else the
 * file's own theme as its view carries it (`hostTheme`, read at import),
 * else nothing (the style's own table).
 */
export function smartArtHostThemeSource(
  host: SmartArtThemeEffectSource | null | undefined,
  view: { readonly effects?: string | null; readonly xml?: string | null } | null | undefined,
): SmartArtThemeEffectSource | null {
  const viewXml = view && typeof view.xml === "string" && view.xml ? view.xml : null;
  if (host?.xml) return host;
  if (host?.id && (!viewXml || host.id !== view?.effects)) return { id: host.id };
  if (viewXml) return { ...(view?.effects ? { id: view.effects } : {}), xml: viewXml };
  return host?.id ? { id: host.id } : null;
}
