/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import {
  resolveThemeColor,
  type SpreadsheetThemeColor,
} from "@mrnote/core/web/lib/spreadsheet/theme-colors.ts";
import type { Sparkline, SparklineColorField } from "@mrnote/core/web/lib/spreadsheet/sparkline.ts";

/** One style colour: a theme slot (with Excel's tint) or a fixed rgb. */
type StyleColor = { readonly theme: number; readonly tint?: number } | { readonly rgb: string };

/** The colour fields a style sets, in x14 element order. */
export const SPARKLINE_STYLE_FIELDS = [
  "color",
  "negativeColor",
  "markerColor",
  "firstColor",
  "lastColor",
  "highColor",
  "lowColor",
] as const satisfies readonly SparklineColorField[];
type StyleField = (typeof SPARKLINE_STYLE_FIELDS)[number];

export type SparklineStyleRow = "darker50" | "darker25" | "accent" | "lighter40" | "dark" | "colorful";
export const SPARKLINE_STYLE_ROWS: readonly SparklineStyleRow[] = [
  "darker50",
  "darker25",
  "accent",
  "lighter40",
  "dark",
  "colorful",
];

export interface SparklineStyle {
  /** "darker50-1" ... "colorful-6". */
  readonly id: string;
  readonly row: SparklineStyleRow;
  /** 1-6: the accent of an accent row, the number of a Dark/Colorful style. */
  readonly index: number;
  readonly colors: Readonly<Record<StyleField, StyleColor>>;
}

// Excel's tint doubles (the theme ladder it writes).
const DARKER_50 = /* own data: DARKER_50 */ -0.5;
const DARKER_25 = /* own data: DARKER_25 */ -0.25;
const DARKER_35_TEXT = /* own data: DARKER_35_TEXT */ 0.35;
const LIGHTER_25 = /* own data: LIGHTER_25 */ 0.25;
const LIGHTER_40 = /* own data: LIGHTER_40 */ 0.4;
const LIGHTER_50 = /* own data: LIGHTER_50 */ 0.5;
const LIGHTER_80 = /* own data: LIGHTER_80 */ 0.8;

const theme = (slot: number, tint?: number): StyleColor => (tint ? { theme: slot, tint } : { theme: slot });
const rgb = (hex: string): StyleColor => ({ rgb: `#${hex.toLowerCase()}` });
const seven = (
  series: StyleColor,
  negative: StyleColor,
  markers: StyleColor,
  first: StyleColor,
  last: StyleColor,
  high: StyleColor,
  low: StyleColor,
): Record<StyleField, StyleColor> => ({
  color: series,
  negativeColor: negative,
  markerColor: markers,
  firstColor: first,
  lastColor: last,
  highColor: high,
  lowColor: low,
});

function accentRow(row: "darker50" | "darker25" | "accent" | "lighter40", index: number): SparklineStyle {
  const accent = 3 + index; // Accent 1 is theme slot 4
  const next = index === 6 ? 4 : accent + 1; // the negative points take the next accent
  const colors =
    row === "darker50"
      ? seven(
          theme(accent, DARKER_50),
          theme(next),
          theme(accent, DARKER_50),
          theme(accent, LIGHTER_40),
          theme(accent, LIGHTER_40),
          theme(accent),
          theme(accent),
        )
      : row === "darker25"
        ? seven(
            theme(accent, DARKER_25),
            theme(next),
            theme(next, DARKER_25),
            theme(next, DARKER_25),
            theme(next, DARKER_25),
            theme(next, DARKER_25),
            theme(next, DARKER_25),
          )
        : row === "accent"
          ? seven(
              theme(accent),
              theme(next),
              theme(accent, DARKER_25),
              theme(accent, DARKER_25),
              theme(accent, DARKER_25),
              theme(accent, DARKER_25),
              theme(accent, DARKER_25),
            )
          : seven(
              theme(accent, LIGHTER_40),
              theme(0, DARKER_50),
              theme(accent, LIGHTER_80),
              theme(accent, DARKER_25),
              theme(accent, DARKER_25),
              theme(accent, DARKER_50),
              theme(accent, DARKER_50),
            );
  return { id: `${row}-${index}`, row, index, colors };
}

const repeat = (series: StyleColor, other: StyleColor) =>
  seven(series, other, other, other, other, other, other);

const DARK: readonly Record<StyleField, StyleColor>[] = /* own data: DARK */ [{"color":{"theme":1,"tint":0.55},"negativeColor":{"theme":4,"tint":-0.25},"markerColor":{"theme":4,"tint":-0.25},"firstColor":{"theme":4,"tint":-0.25},"lastColor":{"theme":4,"tint":-0.25},"highColor":{"theme":4,"tint":-0.25},"lowColor":{"theme":4,"tint":-0.25}},{"color":{"theme":1,"tint":0.4},"negativeColor":{"theme":5,"tint":-0.25},"markerColor":{"theme":5,"tint":-0.25},"firstColor":{"theme":5,"tint":-0.25},"lastColor":{"theme":5,"tint":-0.25},"highColor":{"theme":5,"tint":-0.25},"lowColor":{"theme":5,"tint":-0.25}},{"color":{"theme":1,"tint":0.25},"negativeColor":{"theme":6,"tint":-0.25},"markerColor":{"theme":6,"tint":-0.25},"firstColor":{"theme":6,"tint":-0.25},"lastColor":{"theme":6,"tint":-0.25},"highColor":{"theme":6,"tint":-0.25},"lowColor":{"theme":6,"tint":-0.25}},{"color":{"theme":1,"tint":0.1},"negativeColor":{"theme":7,"tint":-0.25},"markerColor":{"theme":7,"tint":-0.25},"firstColor":{"theme":7,"tint":-0.25},"lastColor":{"theme":7,"tint":-0.25},"highColor":{"theme":7,"tint":-0.25},"lowColor":{"theme":7,"tint":-0.25}},{"color":{"theme":1},"negativeColor":{"theme":8,"tint":-0.25},"markerColor":{"theme":8,"tint":-0.25},"firstColor":{"theme":8,"tint":-0.25},"lastColor":{"theme":8,"tint":-0.25},"highColor":{"theme":8,"tint":-0.25},"lowColor":{"theme":8,"tint":-0.25}},{"color":{"theme":3},"negativeColor":{"theme":9,"tint":-0.25},"markerColor":{"theme":9,"tint":-0.25},"firstColor":{"theme":9,"tint":-0.25},"lastColor":{"theme":9,"tint":-0.25},"highColor":{"theme":9,"tint":-0.25},"lowColor":{"theme":9,"tint":-0.25}}];

const COLORFUL: readonly Record<StyleField, StyleColor>[] = /* own data: COLORFUL */ [{"color":{"rgb":"#8d803f"},"negativeColor":{"rgb":"#5fd42b"},"markerColor":{"rgb":"#2bd488"},"firstColor":{"rgb":"#2b90d4"},"lastColor":{"rgb":"#572bd4"},"highColor":{"rgb":"#d42bc0"},"lowColor":{"rgb":"#d42b2f"}},{"color":{"rgb":"#598d3f"},"negativeColor":{"rgb":"#2bd484"},"markerColor":{"rgb":"#2b94d4"},"firstColor":{"rgb":"#532bd4"},"lastColor":{"rgb":"#d42bc4"},"highColor":{"rgb":"#d42b33"},"lowColor":{"rgb":"#d4b42b"}},{"color":{"rgb":"#3f8d66"},"negativeColor":{"rgb":"#2b98d4"},"markerColor":{"rgb":"#4f2bd4"},"firstColor":{"rgb":"#d42bc8"},"lastColor":{"rgb":"#d42b37"},"highColor":{"rgb":"#d4b02b"},"lowColor":{"rgb":"#67d42b"}},{"color":{"rgb":"#3f738d"},"negativeColor":{"rgb":"#4b2bd4"},"markerColor":{"rgb":"#d42bcc"},"firstColor":{"rgb":"#d42b3b"},"lastColor":{"rgb":"#d4ac2b"},"highColor":{"rgb":"#6bd42b"},"lowColor":{"rgb":"#2bd47b"}},{"color":{"rgb":"#4c3f8d"},"negativeColor":{"rgb":"#d42bd0"},"markerColor":{"rgb":"#d42b3f"},"firstColor":{"rgb":"#d4a82b"},"lastColor":{"rgb":"#6fd42b"},"highColor":{"rgb":"#2bd477"},"lowColor":{"rgb":"#2ba0d4"}},{"color":{"rgb":"#8d3f8d"},"negativeColor":{"rgb":"#d42b43"},"markerColor":{"rgb":"#d4a42b"},"firstColor":{"rgb":"#73d42b"},"lastColor":{"rgb":"#2bd473"},"highColor":{"rgb":"#2ba4d4"},"lowColor":{"rgb":"#432bd4"}}];

/** The gallery in Excel's order: row by row, left to right. */
export const SPARKLINE_STYLES: readonly SparklineStyle[] = /* own data: SPARKLINE_STYLES */ [{"id":"darker50-1","row":"darker50","index":1,"colors":{"color":{"theme":4,"tint":-0.5},"negativeColor":{"theme":7,"tint":-0.5},"markerColor":{"theme":4,"tint":-0.5},"firstColor":{"theme":4,"tint":0.3},"lastColor":{"theme":4,"tint":0.3},"highColor":{"theme":4,"tint":0.3},"lowColor":{"theme":7,"tint":0.3}}},{"id":"darker50-2","row":"darker50","index":2,"colors":{"color":{"theme":5,"tint":-0.5},"negativeColor":{"theme":8,"tint":-0.5},"markerColor":{"theme":5,"tint":-0.5},"firstColor":{"theme":5,"tint":0.3},"lastColor":{"theme":5,"tint":0.3},"highColor":{"theme":5,"tint":0.3},"lowColor":{"theme":8,"tint":0.3}}},{"id":"darker50-3","row":"darker50","index":3,"colors":{"color":{"theme":6,"tint":-0.5},"negativeColor":{"theme":9,"tint":-0.5},"markerColor":{"theme":6,"tint":-0.5},"firstColor":{"theme":6,"tint":0.3},"lastColor":{"theme":6,"tint":0.3},"highColor":{"theme":6,"tint":0.3},"lowColor":{"theme":9,"tint":0.3}}},{"id":"darker50-4","row":"darker50","index":4,"colors":{"color":{"theme":7,"tint":-0.5},"negativeColor":{"theme":4,"tint":-0.5},"markerColor":{"theme":7,"tint":-0.5},"firstColor":{"theme":7,"tint":0.3},"lastColor":{"theme":7,"tint":0.3},"highColor":{"theme":7,"tint":0.3},"lowColor":{"theme":4,"tint":0.3}}},{"id":"darker50-5","row":"darker50","index":5,"colors":{"color":{"theme":8,"tint":-0.5},"negativeColor":{"theme":5,"tint":-0.5},"markerColor":{"theme":8,"tint":-0.5},"firstColor":{"theme":8,"tint":0.3},"lastColor":{"theme":8,"tint":0.3},"highColor":{"theme":8,"tint":0.3},"lowColor":{"theme":5,"tint":0.3}}},{"id":"darker50-6","row":"darker50","index":6,"colors":{"color":{"theme":9,"tint":-0.5},"negativeColor":{"theme":6,"tint":-0.5},"markerColor":{"theme":9,"tint":-0.5},"firstColor":{"theme":9,"tint":0.3},"lastColor":{"theme":9,"tint":0.3},"highColor":{"theme":9,"tint":0.3},"lowColor":{"theme":6,"tint":0.3}}},{"id":"darker25-1","row":"darker25","index":1,"colors":{"color":{"theme":4,"tint":-0.25},"negativeColor":{"theme":7,"tint":-0.25},"markerColor":{"theme":4,"tint":-0.25},"firstColor":{"theme":4,"tint":0.4},"lastColor":{"theme":4,"tint":0.4},"highColor":{"theme":4,"tint":0.4},"lowColor":{"theme":7,"tint":0.4}}},{"id":"darker25-2","row":"darker25","index":2,"colors":{"color":{"theme":5,"tint":-0.25},"negativeColor":{"theme":8,"tint":-0.25},"markerColor":{"theme":5,"tint":-0.25},"firstColor":{"theme":5,"tint":0.4},"lastColor":{"theme":5,"tint":0.4},"highColor":{"theme":5,"tint":0.4},"lowColor":{"theme":8,"tint":0.4}}},{"id":"darker25-3","row":"darker25","index":3,"colors":{"color":{"theme":6,"tint":-0.25},"negativeColor":{"theme":9,"tint":-0.25},"markerColor":{"theme":6,"tint":-0.25},"firstColor":{"theme":6,"tint":0.4},"lastColor":{"theme":6,"tint":0.4},"highColor":{"theme":6,"tint":0.4},"lowColor":{"theme":9,"tint":0.4}}},{"id":"darker25-4","row":"darker25","index":4,"colors":{"color":{"theme":7,"tint":-0.25},"negativeColor":{"theme":4,"tint":-0.25},"markerColor":{"theme":7,"tint":-0.25},"firstColor":{"theme":7,"tint":0.4},"lastColor":{"theme":7,"tint":0.4},"highColor":{"theme":7,"tint":0.4},"lowColor":{"theme":4,"tint":0.4}}},{"id":"darker25-5","row":"darker25","index":5,"colors":{"color":{"theme":8,"tint":-0.25},"negativeColor":{"theme":5,"tint":-0.25},"markerColor":{"theme":8,"tint":-0.25},"firstColor":{"theme":8,"tint":0.4},"lastColor":{"theme":8,"tint":0.4},"highColor":{"theme":8,"tint":0.4},"lowColor":{"theme":5,"tint":0.4}}},{"id":"darker25-6","row":"darker25","index":6,"colors":{"color":{"theme":9,"tint":-0.25},"negativeColor":{"theme":6,"tint":-0.25},"markerColor":{"theme":9,"tint":-0.25},"firstColor":{"theme":9,"tint":0.4},"lastColor":{"theme":9,"tint":0.4},"highColor":{"theme":9,"tint":0.4},"lowColor":{"theme":6,"tint":0.4}}},{"id":"accent-1","row":"accent","index":1,"colors":{"color":{"theme":4},"negativeColor":{"theme":7},"markerColor":{"theme":4},"firstColor":{"theme":4,"tint":-0.5},"lastColor":{"theme":4,"tint":-0.5},"highColor":{"theme":4,"tint":-0.5},"lowColor":{"theme":7,"tint":-0.5}}},{"id":"accent-2","row":"accent","index":2,"colors":{"color":{"theme":5},"negativeColor":{"theme":8},"markerColor":{"theme":5},"firstColor":{"theme":5,"tint":-0.5},"lastColor":{"theme":5,"tint":-0.5},"highColor":{"theme":5,"tint":-0.5},"lowColor":{"theme":8,"tint":-0.5}}},{"id":"accent-3","row":"accent","index":3,"colors":{"color":{"theme":6},"negativeColor":{"theme":9},"markerColor":{"theme":6},"firstColor":{"theme":6,"tint":-0.5},"lastColor":{"theme":6,"tint":-0.5},"highColor":{"theme":6,"tint":-0.5},"lowColor":{"theme":9,"tint":-0.5}}},{"id":"accent-4","row":"accent","index":4,"colors":{"color":{"theme":7},"negativeColor":{"theme":4},"markerColor":{"theme":7},"firstColor":{"theme":7,"tint":-0.5},"lastColor":{"theme":7,"tint":-0.5},"highColor":{"theme":7,"tint":-0.5},"lowColor":{"theme":4,"tint":-0.5}}},{"id":"accent-5","row":"accent","index":5,"colors":{"color":{"theme":8},"negativeColor":{"theme":5},"markerColor":{"theme":8},"firstColor":{"theme":8,"tint":-0.5},"lastColor":{"theme":8,"tint":-0.5},"highColor":{"theme":8,"tint":-0.5},"lowColor":{"theme":5,"tint":-0.5}}},{"id":"accent-6","row":"accent","index":6,"colors":{"color":{"theme":9},"negativeColor":{"theme":6},"markerColor":{"theme":9},"firstColor":{"theme":9,"tint":-0.5},"lastColor":{"theme":9,"tint":-0.5},"highColor":{"theme":9,"tint":-0.5},"lowColor":{"theme":6,"tint":-0.5}}},{"id":"lighter40-1","row":"lighter40","index":1,"colors":{"color":{"theme":4,"tint":0.4},"negativeColor":{"theme":7,"tint":0.4},"markerColor":{"theme":4,"tint":0.4},"firstColor":{"theme":4,"tint":-0.5},"lastColor":{"theme":4,"tint":-0.5},"highColor":{"theme":4,"tint":-0.5},"lowColor":{"theme":7,"tint":-0.5}}},{"id":"lighter40-2","row":"lighter40","index":2,"colors":{"color":{"theme":5,"tint":0.4},"negativeColor":{"theme":8,"tint":0.4},"markerColor":{"theme":5,"tint":0.4},"firstColor":{"theme":5,"tint":-0.5},"lastColor":{"theme":5,"tint":-0.5},"highColor":{"theme":5,"tint":-0.5},"lowColor":{"theme":8,"tint":-0.5}}},{"id":"lighter40-3","row":"lighter40","index":3,"colors":{"color":{"theme":6,"tint":0.4},"negativeColor":{"theme":9,"tint":0.4},"markerColor":{"theme":6,"tint":0.4},"firstColor":{"theme":6,"tint":-0.5},"lastColor":{"theme":6,"tint":-0.5},"highColor":{"theme":6,"tint":-0.5},"lowColor":{"theme":9,"tint":-0.5}}},{"id":"lighter40-4","row":"lighter40","index":4,"colors":{"color":{"theme":7,"tint":0.4},"negativeColor":{"theme":4,"tint":0.4},"markerColor":{"theme":7,"tint":0.4},"firstColor":{"theme":7,"tint":-0.5},"lastColor":{"theme":7,"tint":-0.5},"highColor":{"theme":7,"tint":-0.5},"lowColor":{"theme":4,"tint":-0.5}}},{"id":"lighter40-5","row":"lighter40","index":5,"colors":{"color":{"theme":8,"tint":0.4},"negativeColor":{"theme":5,"tint":0.4},"markerColor":{"theme":8,"tint":0.4},"firstColor":{"theme":8,"tint":-0.5},"lastColor":{"theme":8,"tint":-0.5},"highColor":{"theme":8,"tint":-0.5},"lowColor":{"theme":5,"tint":-0.5}}},{"id":"lighter40-6","row":"lighter40","index":6,"colors":{"color":{"theme":9,"tint":0.4},"negativeColor":{"theme":6,"tint":0.4},"markerColor":{"theme":9,"tint":0.4},"firstColor":{"theme":9,"tint":-0.5},"lastColor":{"theme":9,"tint":-0.5},"highColor":{"theme":9,"tint":-0.5},"lowColor":{"theme":6,"tint":-0.5}}},{"id":"dark-1","row":"dark","index":1,"colors":{"color":{"theme":1,"tint":0.55},"negativeColor":{"theme":4,"tint":-0.25},"markerColor":{"theme":4,"tint":-0.25},"firstColor":{"theme":4,"tint":-0.25},"lastColor":{"theme":4,"tint":-0.25},"highColor":{"theme":4,"tint":-0.25},"lowColor":{"theme":4,"tint":-0.25}}},{"id":"dark-2","row":"dark","index":2,"colors":{"color":{"theme":1,"tint":0.4},"negativeColor":{"theme":5,"tint":-0.25},"markerColor":{"theme":5,"tint":-0.25},"firstColor":{"theme":5,"tint":-0.25},"lastColor":{"theme":5,"tint":-0.25},"highColor":{"theme":5,"tint":-0.25},"lowColor":{"theme":5,"tint":-0.25}}},{"id":"dark-3","row":"dark","index":3,"colors":{"color":{"theme":1,"tint":0.25},"negativeColor":{"theme":6,"tint":-0.25},"markerColor":{"theme":6,"tint":-0.25},"firstColor":{"theme":6,"tint":-0.25},"lastColor":{"theme":6,"tint":-0.25},"highColor":{"theme":6,"tint":-0.25},"lowColor":{"theme":6,"tint":-0.25}}},{"id":"dark-4","row":"dark","index":4,"colors":{"color":{"theme":1,"tint":0.1},"negativeColor":{"theme":7,"tint":-0.25},"markerColor":{"theme":7,"tint":-0.25},"firstColor":{"theme":7,"tint":-0.25},"lastColor":{"theme":7,"tint":-0.25},"highColor":{"theme":7,"tint":-0.25},"lowColor":{"theme":7,"tint":-0.25}}},{"id":"dark-5","row":"dark","index":5,"colors":{"color":{"theme":1},"negativeColor":{"theme":8,"tint":-0.25},"markerColor":{"theme":8,"tint":-0.25},"firstColor":{"theme":8,"tint":-0.25},"lastColor":{"theme":8,"tint":-0.25},"highColor":{"theme":8,"tint":-0.25},"lowColor":{"theme":8,"tint":-0.25}}},{"id":"dark-6","row":"dark","index":6,"colors":{"color":{"theme":3},"negativeColor":{"theme":9,"tint":-0.25},"markerColor":{"theme":9,"tint":-0.25},"firstColor":{"theme":9,"tint":-0.25},"lastColor":{"theme":9,"tint":-0.25},"highColor":{"theme":9,"tint":-0.25},"lowColor":{"theme":9,"tint":-0.25}}},{"id":"colorful-1","row":"colorful","index":1,"colors":{"color":{"rgb":"#8d803f"},"negativeColor":{"rgb":"#5fd42b"},"markerColor":{"rgb":"#2bd488"},"firstColor":{"rgb":"#2b90d4"},"lastColor":{"rgb":"#572bd4"},"highColor":{"rgb":"#d42bc0"},"lowColor":{"rgb":"#d42b2f"}}},{"id":"colorful-2","row":"colorful","index":2,"colors":{"color":{"rgb":"#598d3f"},"negativeColor":{"rgb":"#2bd484"},"markerColor":{"rgb":"#2b94d4"},"firstColor":{"rgb":"#532bd4"},"lastColor":{"rgb":"#d42bc4"},"highColor":{"rgb":"#d42b33"},"lowColor":{"rgb":"#d4b42b"}}},{"id":"colorful-3","row":"colorful","index":3,"colors":{"color":{"rgb":"#3f8d66"},"negativeColor":{"rgb":"#2b98d4"},"markerColor":{"rgb":"#4f2bd4"},"firstColor":{"rgb":"#d42bc8"},"lastColor":{"rgb":"#d42b37"},"highColor":{"rgb":"#d4b02b"},"lowColor":{"rgb":"#67d42b"}}},{"id":"colorful-4","row":"colorful","index":4,"colors":{"color":{"rgb":"#3f738d"},"negativeColor":{"rgb":"#4b2bd4"},"markerColor":{"rgb":"#d42bcc"},"firstColor":{"rgb":"#d42b3b"},"lastColor":{"rgb":"#d4ac2b"},"highColor":{"rgb":"#6bd42b"},"lowColor":{"rgb":"#2bd47b"}}},{"id":"colorful-5","row":"colorful","index":5,"colors":{"color":{"rgb":"#4c3f8d"},"negativeColor":{"rgb":"#d42bd0"},"markerColor":{"rgb":"#d42b3f"},"firstColor":{"rgb":"#d4a82b"},"lastColor":{"rgb":"#6fd42b"},"highColor":{"rgb":"#2bd477"},"lowColor":{"rgb":"#2ba0d4"}}},{"id":"colorful-6","row":"colorful","index":6,"colors":{"color":{"rgb":"#8d3f8d"},"negativeColor":{"rgb":"#d42b43"},"markerColor":{"rgb":"#d4a42b"},"firstColor":{"rgb":"#73d42b"},"lastColor":{"rgb":"#2bd473"},"highColor":{"rgb":"#2ba4d4"},"lowColor":{"rgb":"#432bd4"}}}];

/** The style a new sparkline gets: Sparkline Style Accent 1, Darker 50%. */
export const DEFAULT_SPARKLINE_STYLE_ID = "darker50-1";

export function sparklineStyleById(id: string): SparklineStyle | null {
  return SPARKLINE_STYLES.find((style) => style.id === id) ?? null;
}

/**
 * The colours a style writes, resolved in the workbook's theme
 * (`paletteId`), with each theme colour's identity so a save writes the x14
 * element as `theme`/`tint` and the colour follows a later theme change.
 */
export function sparklineStylePatch(
  style: SparklineStyle,
  paletteId: string,
): Record<StyleField, string> & Pick<Sparkline, "themeColors"> {
  const patch: Partial<Record<StyleField, string>> = {};
  const themeColors: Partial<Record<SparklineColorField, SpreadsheetThemeColor>> = {};
  for (const field of SPARKLINE_STYLE_FIELDS) {
    const color = style.colors[field];
    if ("rgb" in color) {
      patch[field] = color.rgb;
      continue;
    }
    const resolved = resolveThemeColor({ theme: color.theme, tint: color.tint, rgb: "#000000" }, paletteId).toLowerCase();
    patch[field] = resolved;
    themeColors[field] = { theme: color.theme, ...(color.tint ? { tint: color.tint } : {}), rgb: resolved };
  }
  return { ...(patch as Record<StyleField, string>), themeColors };
}

/**
 * The built-in style a sparkline carries, or null: every style colour
 * matches, a theme colour by its slot and tint (whatever theme it resolves
 * in), an rgb one by value. The gallery marks it as the current style.
 */
export function sparklineStyleIdOf(sparkline: Sparkline): string | null {
  for (const style of SPARKLINE_STYLES) {
    const matches = SPARKLINE_STYLE_FIELDS.every((field) => {
      const color = style.colors[field];
      const carried = sparkline[field];
      if (!carried) return false;
      if ("rgb" in color) return !sparkline.themeColors?.[field] && carried.toLowerCase() === color.rgb;
      const ref = sparkline.themeColors?.[field];
      return Boolean(ref && ref.theme === color.theme && Math.abs((ref.tint ?? 0) - (color.tint ?? 0)) < 0.0005);
    });
    if (matches) return style.id;
  }
  return null;
}

/**
 * A sparkline as the workbook's current theme draws it: each colour that
 * carries a theme identity resolved in `paletteId` (Excel redraws a theme
 * colour after a theme change). The stored colours never change, so the
 * saved x14 element keeps its theme reference. Unchanged without a palette
 * or without theme colours.
 */
export function sparklineInTheme(sparkline: Sparkline, paletteId: string | undefined): Sparkline {
  const refs = sparkline.themeColors;
  if (!paletteId || !refs) return sparkline;
  let next: Sparkline | null = null;
  for (const [field, ref] of Object.entries(refs) as [SparklineColorField, SpreadsheetThemeColor][]) {
    if (!ref) continue;
    const resolved = resolveThemeColor(ref, paletteId);
    if (resolved.toLowerCase() === sparkline[field]?.toLowerCase()) continue;
    next ??= { ...sparkline };
    next[field] = resolved;
  }
  return next ?? sparkline;
}

const themedCache = new WeakMap<Sparkline, { readonly paletteId: string; readonly value: Sparkline }>();

/** {@link sparklineInTheme}, the same object for the same sparkline and
 *  palette (a cell's glyph is not redrawn for nothing). */
export function sparklineInThemeCached(sparkline: Sparkline, paletteId: string | undefined): Sparkline {
  if (!paletteId || !sparkline.themeColors) return sparkline;
  const hit = themedCache.get(sparkline);
  if (hit && hit.paletteId === paletteId) return hit.value;
  const value = sparklineInTheme(sparkline, paletteId);
  themedCache.set(sparkline, { paletteId, value });
  return value;
}
