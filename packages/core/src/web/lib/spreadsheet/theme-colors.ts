/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { officeThemeTintSteps } from "../office-color-palette";
import { paletteById } from "./color-palette";

export interface SpreadsheetThemeColor {
  theme: number;
  tint?: number;
  /** The import/pick-time fallback authenticates this binding after style edits. */
  rgb: string;
}
export type SpreadsheetThemeColors = Record<string, SpreadsheetThemeColor>;
const SLOT =
  /^(?:color|backgroundColor|borderEdges\.(?:top|right|bottom|left|diagonal)\.color|fillPattern\.(?:foregroundColor|backgroundColor)|gradientFill\.stops\.(?:[0-9]|1[0-5])\.color)$/;

export function normalizeThemeColor(
  value: unknown,
): SpreadsheetThemeColor | null {
  if (!value || typeof value !== "object") return null;
  const ref = value as SpreadsheetThemeColor;
  if (
    !Number.isInteger(ref.theme) ||
    ref.theme < 0 ||
    ref.theme > 11 ||
    typeof ref.rgb !== "string" ||
    !/^#[0-9a-f]{6}$/i.test(ref.rgb) ||
    (ref.tint !== undefined &&
      (typeof ref.tint !== "number" ||
        !Number.isFinite(ref.tint) ||
        Math.abs(ref.tint) > 1))
  )
    return null;
  return {
    theme: ref.theme,
    ...(ref.tint ? { tint: ref.tint } : {}),
    rgb: ref.rgb.toLowerCase(),
  };
}

function atPath(root: object, path: string): unknown {
  return path
    .split(".")
    .reduce<unknown>(
      (value, key) =>
        value && typeof value === "object"
          ? (value as Record<string, unknown>)[key]
          : undefined,
      root,
    );
}

export function validThemeColors(style: object): SpreadsheetThemeColors {
  const raw = (style as { themeColors?: unknown }).themeColors;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return {};
  if (Object.keys(raw).length > 26) return {};
  const result: SpreadsheetThemeColors = {};
  for (const [path, value] of Object.entries(raw)) {
    if (!SLOT.test(path)) continue;
    const ref = normalizeThemeColor(value);
    const color = atPath(style, path);
    if (ref && typeof color === "string" && color.toLowerCase() === ref.rgb)
      result[path] = ref;
  }
  return result;
}

/** SpreadsheetML tint changes HSL luminance, not individual RGB channels. */
const tinted = new Map<string, string>();

/** Remembered: a table style's bands ask this for every cell drawn. */
export function tintSpreadsheetColor(color: string, tint: number): string {
  if (!tint) return color;
  const key = `${color}\u0000${tint}`;
  let result = tinted.get(key);
  if (result === undefined) {
    if (tinted.size >= 5_000) tinted.clear();
    result = computeTint(color, tint);
    tinted.set(key, result);
  }
  return result;
}

function computeTint(color: string, tint: number): string {
  const [r, g, b] = [1, 3, 5].map(
    (i) => parseInt(color.slice(i, i + 2), 16) / 255,
  );
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b),
    delta = max - min;
  const light = (max + min) / 2;
  const saturation =
    delta === 0 ? 0 : delta / (light > 0.5 ? 2 - max - min : max + min);
  let hue =
    delta === 0
      ? 0
      : max === r
        ? ((g - b) / delta) % 6
        : max === g
          ? (b - r) / delta + 2
          : (r - g) / delta + 4;
  hue = (hue / 6 + 1) % 1;
  const l = Math.max(
    0,
    Math.min(1, tint < 0 ? light * (1 + tint) : light * (1 - tint) + tint),
  );
  const q = l < 0.5 ? l * (1 + saturation) : l + saturation - l * saturation,
    p = 2 * l - q;
  const channel = (h: number) => {
    h = (h + 1) % 1;
    return h < 1 / 6
      ? p + (q - p) * 6 * h
      : h < 0.5
        ? q
        : h < 2 / 3
          ? p + (q - p) * (2 / 3 - h) * 6
          : p;
  };
  return (
    "#" +
    (saturation === 0
      ? [l, l, l]
      : [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)]
    )
      .map((c) =>
        Math.round(c * 255)
          .toString(16)
          .padStart(2, "0"),
      )
      .join("")
  );
}

export function resolveThemeColor(
  ref: SpreadsheetThemeColor,
  paletteId: string,
): string {
  const palette = paletteById(paletteId);
  const slots = [
    "#ffffff",
    "#000000",
    palette.light2,
    palette.dark2,
    ...palette.accents,
    palette.hyperlink,
    palette.followedHyperlink,
  ];
  return tintSpreadsheetColor(slots[ref.theme] ?? ref.rgb, ref.tint ?? 0);
}

/**
 * The tints Excel lists under theme slot `theme` (0-9, Background 1 first):
 * each column's own ladder, which depends on how light the slot colour is.
 */
export function themeTintSteps(
  theme: number,
  paletteId: string,
): readonly number[] {
  return officeThemeTintSteps(
    resolveThemeColor({ theme, rgb: "#000000" }, paletteId),
  );
}

/** The grid row (0 = the colour itself) of a theme binding, or -1 when none. */
export function themeTintRow(
  ref: Pick<SpreadsheetThemeColor, "theme" | "tint">,
  paletteId: string,
): number {
  if (!ref.tint) return 0;
  const tint = ref.tint;
  // Excel writes the ladder as doubles (0.79998168889431442 for 80%).
  const index = themeTintSteps(ref.theme, paletteId).findIndex(
    (step) => Math.abs(step - tint) < 0.005,
  );
  return index < 0 ? -1 : index + 1;
}

export function themeColorGrid(paletteId: string): string[][] {
  return Array.from({ length: 10 }, (_, theme) =>
    [0, ...themeTintSteps(theme, paletteId)].map((tint) =>
      resolveThemeColor({ theme, tint, rgb: "#000000" }, paletteId),
    ),
  );
}

export function themeColorFromEntry(
  color: string,
  entry: { kind: string; occurrenceId: string } | undefined,
  paletteId: string,
): SpreadsheetThemeColor | undefined {
  const match =
    entry?.kind === "theme" && /^theme-(\d+)-(\d+)$/.exec(entry.occurrenceId);
  if (!match) return undefined;
  const theme = Number(match[1]),
    row = Number(match[2]);
  if (theme > 9 || row > 5) return undefined;
  const tint = row === 0 ? 0 : themeTintSteps(theme, paletteId)[row - 1];
  return { theme, ...(tint ? { tint } : {}), rgb: color };
}

/** Clone only bound paths. Stored fallback colors and native identities never change on a theme switch. */
export function resolveThemeStyle<T extends object>(
  style: T,
  paletteId?: string,
): T {
  if (!paletteId) return style;
  const bindings = validThemeColors(style);
  if (!Object.keys(bindings).length) return style;
  const result = { ...style } as Record<string, unknown>;
  for (const [path, ref] of Object.entries(bindings)) {
    const parts = path.split(".");
    let target = result;
    for (const key of parts.slice(0, -1)) {
      const value = target[key];
      target[key] = Array.isArray(value)
        ? [...value]
        : { ...(value as object) };
      target = target[key] as Record<string, unknown>;
    }
    target[parts.at(-1)!] = resolveThemeColor(ref, paletteId);
  }
  return result as T;
}

/** A direct fill/font/border command releases just the slots it owns. */
export function themeColorsAfterPatch(
  previous: object,
  patch: object,
): SpreadsheetThemeColors | undefined {
  const result = validThemeColors(previous);
  const explicit = validThemeColors(patch);
  for (const path of Object.keys(result)) {
    if (Object.hasOwn(patch, path.split(".")[0])) delete result[path];
  }
  Object.assign(result, explicit);
  return Object.keys(result).length ? result : undefined;
}

/**
 * A cell style as it lays over a table style: a font colour that is the
 * workbook's Text 1 (theme 1, untinted -- the Normal style's own colour) is
 * no colour of the cell's, so the table style's text colour shows. Measured
 * on Excel for Mac 16.86 (TableStyleMedium2 headers): theme 1 header text
 * draws the style's white, an rgb FF000000 one black (P27_56274's white
 * "Header A..E" on orange).
 */
export function cellStyleOverTableStyle<T extends object | undefined>(style: T): T {
  if (!style) return style;
  // One answer per style object: what is drawn from it is kept by its identity.
  const known = stylesOverTables.get(style);
  if (known) return known as T;
  const colors = validThemeColors(style);
  const text = colors.color;
  let over: object = style;
  if (text && text.theme === 1 && !text.tint) {
    const { color: _color, themeColors: _themeColors, ...rest } = style as Record<string, unknown>;
    const { color: _text, ...otherColors } = colors;
    over = Object.keys(otherColors).length ? { ...rest, themeColors: otherColors } : rest;
  }
  stylesOverTables.set(style, over);
  return over as T;
}
const stylesOverTables = new WeakMap<object, object>();
