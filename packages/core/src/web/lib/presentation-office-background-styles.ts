/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Design ▸ Variants ▸ Background Styles for an Office theme deck (PPT-A073).
 *
 * PowerPoint's gallery offers twelve styles: the three background fills of
 * the theme's format scheme (`a:bgFillStyleLst`, which a master's
 * `p:bgRef@idx` 1001..1003 names), each in four colours. A master's bgRef
 * only ever names the background slots bg1 / bg2 -- the 26 Office themes'
 * masters and variants in the server catalog all do -- so the four colours are
 * bg1 and bg2 under the light colour map (lt1, lt2) and bg2 and bg1 under the
 * dark one (dk2, dk1), the map dark variants use: bg1 = dk1, tx1 = lt1,
 * bg2 = dk2, tx2 = lt2. A dark style therefore turns the text light too.
 *
 * The fills are `presentation-office-background-styles.json`, generated from
 * the server catalog (MRAI-corpus/ppt-audit/L17/scripts/
 * build_background_styles.py): transform names and numbers only. The canvas
 * draws a style through the deck theme's background fields (a picture fill,
 * which the browser has no picture for, as its placeholder colour); the
 * download writes the style itself (`p:bgRef` and the master's `p:clrMap`).
 *
 * The gallery's column order (light bg1, light bg2, dark bg2, dark bg1) is
 * PowerPoint's as remembered from its Windows gallery, not yet measured on
 * this machine's PowerPoint.
 */

import source from "./presentation-office-background-styles.json";
import { officeTransformedHex } from "./office-color-transform";
import {
  normalizeDeckOfficeMaster,
  OFFICE_BACKGROUND_STYLE_COUNT,
} from "./presentation-office-master-model";
import { presentationOfficeThemeVariant } from "./presentation-office-themes";
import type { BgGradient } from "./presentation-slide-bg";
import {
  normalizeTheme,
  normalizeThemeColorMap,
  type DeckTheme,
  type DeckThemeColorMap,
  type PresentationThemeColorSlot,
} from "./presentation-theme";

type Transforms = ReadonlyArray<readonly [string, number]>;

type RawFill =
  | { t: "s"; m: Array<[string, number]> }
  | { t: "g"; s: Array<[number, Array<[string, number]>]>; a?: number; p?: string }
  | { t: "b" };

interface RawThemeStyles {
  fills: RawFill[];
  dark: boolean;
  variants?: Record<string, { fills?: RawFill[]; dark?: boolean }>;
}

const catalog = source as unknown as { themes: Record<string, RawThemeStyles> };

/** One cell of the gallery. */
export interface OfficeBackgroundStyle {
  /** 1..12, the gallery's order (row by row). */
  readonly index: number;
  /** `p:bgRef@idx` - 1000: the format scheme's background fill, 1..3. */
  readonly fill: 1 | 2 | 3;
  /** The slot `p:bgRef` names. */
  readonly slot: "bg1" | "bg2";
  /** Whether the master's colour map is the dark one. */
  readonly dark: boolean;
  /** How the canvas draws it: a colour, and a gradient where the fill is one. */
  readonly background: string;
  readonly backgroundGradient?: BgGradient;
}

const COLUMNS: ReadonlyArray<{ slot: "bg1" | "bg2"; dark: boolean }> = [
  { slot: "bg1", dark: false },
  { slot: "bg2", dark: false },
  { slot: "bg2", dark: true },
  { slot: "bg1", dark: true },
];

function themeStyles(themeId: string, variantId: string): { fills: RawFill[]; dark: boolean } | null {
  const theme = Object.prototype.hasOwnProperty.call(catalog.themes, themeId)
    ? catalog.themes[themeId]
    : undefined;
  if (!theme) return null;
  const variant = theme.variants?.[variantId];
  return { fills: variant?.fills ?? theme.fills, dark: variant?.dark ?? theme.dark };
}

/** The slot colours the background styles resolve in: the theme's, re-coloured by a Colors pick. */
function slotColors(theme: DeckTheme): Record<string, string> | null {
  const identity = theme.identity;
  if (!identity) return null;
  const variant = presentationOfficeThemeVariant(identity.themeId, identity.variantId);
  if (!variant) return null;
  return { ...variant.colors, ...(theme.colorSlots ?? {}) };
}

/**
 * The twelve Background Styles of an Office theme deck's theme and variant,
 * or [] for any other deck (an imported deck keeps its package's master).
 */
export function officeBackgroundStyles(theme: DeckTheme | null | undefined): OfficeBackgroundStyle[] {
  const identity = theme?.identity;
  if (!theme || !identity || identity.selectionKind !== "office" || identity.source) return [];
  const styles = themeStyles(identity.themeId, identity.variantId);
  const colors = slotColors(theme);
  if (!styles || !colors || styles.fills.length < 3) return [];
  const result: OfficeBackgroundStyle[] = [];
  styles.fills.slice(0, 3).forEach((fill, row) => {
    COLUMNS.forEach((column, columnIndex) => {
      const base = column.dark
        ? colors[column.slot === "bg1" ? "dk1" : "dk2"]
        : colors[column.slot === "bg1" ? "lt1" : "lt2"];
      if (!base) return;
      const transformed = (mods: Transforms) => officeTransformedHex(base, mods) ?? base;
      const style: {
        -readonly [K in keyof OfficeBackgroundStyle]: OfficeBackgroundStyle[K];
      } = {
        index: row * 4 + columnIndex + 1,
        fill: (row + 1) as 1 | 2 | 3,
        slot: column.slot,
        dark: column.dark,
        background: base,
      };
      if (fill.t === "s") {
        style.background = transformed(fill.m);
      } else if (fill.t === "g" && fill.s.length >= 2) {
        const stops = fill.s.map(([position, mods]) => ({
          color: transformed(mods),
          position: Math.min(1, Math.max(0, position)),
        }));
        // DrawingML's lin@ang runs clockwise from left-to-right; the page
        // stores the CSS angle (the writers turn it back: CSS - 90). A path
        // (radial) fill is drawn as its top-to-bottom ramp.
        const angle = ((fill.a ?? 90) + 90) % 360;
        style.background = stops[0].color;
        style.backgroundGradient = {
          from: stops[0].color,
          to: stops[stops.length - 1].color,
          angle,
          stops,
        };
      }
      result.push(style);
    });
  });
  return result;
}

/** Whether the deck theme's own master (its variant) uses the dark colour map. */
export function officeThemeMasterIsDark(theme: DeckTheme | null | undefined): boolean {
  const identity = theme?.identity;
  if (!identity) return false;
  return themeStyles(identity.themeId, identity.variantId)?.dark ?? false;
}

/**
 * Whether the master's colour map, after a Background Styles pick, is the
 * opposite of the theme variant's own -- so text the theme draws in tx1/tx2
 * draws in the other slot of its pair.
 */
export function officeMasterColorMapFlipped(theme: DeckTheme | null | undefined): boolean {
  const index = normalizeDeckOfficeMaster(theme?.officeMaster)?.backgroundStyle;
  if (index === undefined) return false;
  const style = COLUMNS[(index - 1) % 4];
  return style.dark !== officeThemeMasterIsDark(theme);
}

const LIGHT_COLOR_MAP: DeckThemeColorMap = Object.freeze({ bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" });

/**
 * The slot each of Background 1 / Text 1 / Background 2 / Text 2 names on the
 * deck's slides (`p:clrMap`): the map an imported deck's master states
 * (`colorMap`, reading row 826), else the Office theme variant's own master
 * (a dark variant maps tx1 to lt1), else the light map; then flipped when a
 * Background Styles pick swapped the master's map. One resolution for every
 * reader of Text 1 -- the slide's ink, the Theme Colors grid.
 */
export function presentationThemeColorMap(theme: DeckTheme | null | undefined): DeckThemeColorMap {
  const identity = theme?.identity;
  const stated = normalizeThemeColorMap(theme?.colorMap);
  const own =
    stated ??
    (identity?.selectionKind === "office"
      ? presentationOfficeThemeVariant(identity.themeId, identity.variantId)?.colorMap
      : undefined) ??
    LIGHT_COLOR_MAP;
  if (!officeMasterColorMapFlipped(theme)) return own;
  return {
    bg1: flippedColorSlot(own.bg1) as PresentationThemeColorSlot,
    tx1: flippedColorSlot(own.tx1) as PresentationThemeColorSlot,
    bg2: flippedColorSlot(own.bg2) as PresentationThemeColorSlot,
    tx2: flippedColorSlot(own.tx2) as PresentationThemeColorSlot,
  };
}

/** The pair slot of a colour slot under the flipped map (dk1 <-> lt1, dk2 <-> lt2). */
export function flippedColorSlot(slot: string): string {
  return ({ dk1: "lt1", lt1: "dk1", dk2: "lt2", lt2: "dk2" } as Record<string, string>)[slot] ?? slot;
}

/**
 * The theme with Background Styles `index` picked: the master's background is
 * that style (and the master's own Format Background edit, if any, is
 * replaced by it), the canvas draws it through the theme's background fields,
 * and the default title / body colours follow the colour map. Null for a
 * no-op or a deck without an Office master.
 */
export function themeWithOfficeBackgroundStyle(
  theme: DeckTheme | null | undefined,
  index: number,
): DeckTheme | null {
  if (!theme || !Number.isInteger(index) || index < 1 || index > OFFICE_BACKGROUND_STYLE_COUNT) return null;
  const style = officeBackgroundStyles(theme).find((candidate) => candidate.index === index);
  const colors = slotColors(theme);
  if (!style || !colors) return null;
  const master = { ...(normalizeDeckOfficeMaster(theme.officeMaster) ?? {}) };
  const wasFlipped = officeMasterColorMapFlipped(theme);
  delete master.background;
  master.backgroundStyle = index;
  const next: DeckTheme = { ...theme, officeMaster: master, background: style.background };
  delete next.backgroundGradient;
  delete next.backgroundPattern;
  delete next.backgroundImage;
  if (style.backgroundGradient) next.backgroundGradient = style.backgroundGradient;
  const isFlipped = officeMasterColorMapFlipped(next);
  if (wasFlipped !== isFlipped) {
    // tx1 / tx2 now map to the other slot of their pair: the theme's default
    // title and body colours move with them.
    for (const key of ["titleColor", "bodyTextColor"] as const) {
      const current = theme[key]?.toLowerCase();
      const slot = Object.keys(colors).find((name) => colors[name]?.toLowerCase() === current);
      if (slot) next[key] = colors[flippedColorSlot(slot)] ?? theme[key];
    }
  }
  const result = normalizeTheme(next);
  return result && JSON.stringify(result) !== JSON.stringify(normalizeTheme(theme)) ? result : null;
}
