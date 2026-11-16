/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { ptToPx } from "./office-units";
import {
  normalizeHexColor,
  OFFICE_THEME_PALETTES,
} from "./office-color-palette";
import {
  resolveOfficeStopColor,
  type OfficeColorSource,
  type OfficeColorTransformEntry,
  type OfficeShapeGradient,
  type OfficeShapeShadow,
} from "./office-shape-paint";
import type { OfficeShapeStylePresetRow } from "./office-shape-style-presets";

export const OFFICE_WEB_SHAPE_ACCENTS = /* own data: OFFICE_WEB_SHAPE_ACCENTS */ ["#2f6fd6","#e07b39","#3a9d6b","#8c5cc9","#d04b6b","#2c9aa8"];














const COLORED_FILL_LINE: readonly OfficeColorTransformEntry[] = /* own data: COLORED_FILL_LINE */ [["lumMod",70000]];
const SUBTLE_STOPS: readonly (readonly OfficeColorTransformEntry[])[] = /* own data: SUBTLE_STOPS */ [[["lumMod",25000],["lumOff",75000]],[["lumMod",40000],["lumOff",60000]],[["lumMod",55000],["lumOff",45000]]];
const MODERATE_STOPS: readonly (readonly OfficeColorTransformEntry[])[] = /* own data: MODERATE_STOPS */ [[["lumMod",80000],["lumOff",20000]],[["lumMod",100000]],[["lumMod",80000]]];
/** `a:effectStyleLst`'s third: blur 57150, distance 19050 EMU, straight down,
 * black at alpha 63000. */
const INTENSE_SHADOW: OfficeShapeShadow = /* own data: INTENSE_SHADOW */ {"color":"#000000","blur":8,"distance":3,"opacity":0.4};

function transformed(
  base: string,
  transforms: readonly OfficeColorTransformEntry[],
): string {
  return (
    resolveOfficeStopColor({ kind: "srgb", base, transforms: transforms.map(([kind, value]): OfficeColorTransformEntry => [kind, value]) } satisfies OfficeColorSource) ??
    base
  );
}
const names = {
  en: [
    "Colored Outline",
    "Colored Fill",
    "Light 1 Outline",
    "Subtle Effect",
    "Moderate Effect",
    "Intense Effect",
  ],
  zh: [
    "彩色轮廓",
    "彩色填充",
    "浅色 1 轮廓",
    "细微效果",
    "中等效果",
    "强烈效果",
  ],
};
/** Each column's colour name and slot, as the gallery labels it (own data). */
const SHAPE_STYLE_COLOR_NAMES: { readonly en: readonly string[]; readonly zh: readonly string[] } = /* own data: SHAPE_STYLE_COLOR_NAMES */ {"en":["Black, Dark 1","Blue, Accent 1","Orange, Accent 2","Green, Accent 3","Purple, Accent 4","Red, Accent 5","Teal, Accent 6"],"zh":["黑色，深色 1","蓝色，强调颜色 1","橙色，强调颜色 2","绿色，强调颜色 3","紫色，强调颜色 4","红色，强调颜色 5","青色，强调颜色 6"]};
export function officeShapeStyleLabel(
  recipe: OfficeShapeStyleRecipe,
  zh: boolean,
  labelFor?: (recipe: OfficeShapeStyleRecipe) => string,
): string {
  if (labelFor) return labelFor(recipe);
  const colors = zh ? SHAPE_STYLE_COLOR_NAMES.zh : SHAPE_STYLE_COLOR_NAMES.en;
  return `${names[zh ? "zh" : "en"][(OFFICE_SHAPE_STYLE_ROWS as readonly string[]).indexOf(recipe.row)]} - ${colors[recipe.column]}`;
}

/**
 * The style's name as `console-notebooks` catalogue keys: Office's
 * `office.shapeStyle.name` ("{effect} - {color}") takes the row's effect and
 * the column's theme colour.
 */
export function officeShapeStyleLabelKeys(recipe: OfficeShapeStyleRecipe): {
  effect: string;
  color: string;
} {
  return {
    effect: `office.shapeStyle.row.${recipe.row}`,
    color: `office.shapeStyle.color.${recipe.column}`,
  };
}

export const OFFICE_SHAPE_STYLE_ROWS = [
  "outline",
  "outlined-fill",
  "solid",
  "light-gradient",
  "gradient",
  "shadow",
] as const;
export type OfficeShapeStyleRow = (typeof OFFICE_SHAPE_STYLE_ROWS)[number];

/** Paint-only values: editor adapters retain their own objects and writers. */
export interface OfficeShapeStyleRecipe {
  readonly id: string;
  /** A theme-style row, or one of the gallery's measured Presets rows. */
  readonly row: OfficeShapeStyleRow | OfficeShapeStylePresetRow;
  readonly column: number;
  readonly fill: string;
  readonly fillTo?: string;
  readonly gradientAngle?: number;
  readonly stroke: string;
  readonly strokeWidth: number;
  readonly textColor: string;
  readonly shadow: boolean;
  readonly fillGradient?: OfficeShapeGradient;
  readonly shapeShadow?: OfficeShapeShadow;
  /** The solid fill's Transparency, a whole percentage (the Semitransparent
   * presets' `a:alpha val="50000"`); absent is opaque. */
  readonly fillTransparency?: number;
}

/**
 * The theme slots a Shape Styles recipe paints its fill and line in (the
 * Office gallery's `fillRef` / `lnRef` colours): a column's colour is Dark 1
 * (column 0) or Accent 1-6, a Colored Outline fills in Light 1 and a Colored
 * Fill outlines in it. `undefined` where the recipe paints no line.
 */
export function officeShapeStyleRecipeSlots(recipe: OfficeShapeStyleRecipe): {
  readonly fill?: string;
  readonly stroke?: string;
} {
  const column = recipe.column === 0 ? "dk1" : `accent${recipe.column}`;
  // The measured Presets rows (`office-shape-style-presets`): Transparent
  // paints neither, Colored Outline its line, the others their fill, in the
  // column's colour.
  if (recipe.row === "preset-transparent") return {};
  if (recipe.row === "preset-outline") return { stroke: column };
  if (String(recipe.row).startsWith("preset-")) return { fill: column };
  if (!(OFFICE_SHAPE_STYLE_ROWS as readonly string[]).includes(recipe.row)) return {};
  if (recipe.row === "outline") return { fill: "lt1", stroke: column };
  if (recipe.row === "solid") return { fill: column, stroke: "lt1" };
  if (recipe.row === "shadow") return { fill: column };
  return { fill: column, stroke: column };
}

/** Seven slot identities are preserved even when a theme repeats a colour. */
export function officeShapeStyleRecipes(
  accents: readonly string[],
): readonly OfficeShapeStyleRecipe[] {
  const columns = [
    "#000000",
    ...Array.from(
      { length: 6 },
      (_, index) =>
        normalizeHexColor(accents[index]) ??
        OFFICE_THEME_PALETTES[0].accents[index],
    ),
  ];
  return OFFICE_SHAPE_STYLE_ROWS.flatMap((row) =>
    columns.map((color, column) => {
      const base = { id: `${row}-${column}`, row, column, shadow: false };
      if (row === "outline")
        return {
          ...base,
          fill: "#ffffff",
          stroke: color,
          strokeWidth: 2,
          textColor: "#000000",
        };
      if (row === "outlined-fill")
        return {
          ...base,
          fill: color,
          stroke: transformed(color, COLORED_FILL_LINE),
          strokeWidth: 2,
          textColor: "#ffffff",
        };
      if (row === "solid")
        return {
          ...base,
          fill: color,
          stroke: "#ffffff",
          strokeWidth: 8 / 3,
          textColor: "#ffffff",
        };
      const isSubtle = row === "light-gradient";
      const stops = (isSubtle ? SUBTLE_STOPS : MODERATE_STOPS).map((transforms) =>
        transformed(color, transforms),
      );
      return {
        ...base,
        fill: stops[0],
        fillTo: stops[2],
        fillGradient: {
          angle: 180,
          stops: stops.map((color, i) => ({ position: i / 2, color })),
        },
        gradientAngle: 180,
        stroke: row === "shadow" ? "none" : color,
        strokeWidth: row === "shadow" ? 0 : ptToPx(1),
        textColor: isSubtle ? "#000000" : "#ffffff",
        shadow: row === "shadow",
        ...(row === "shadow"
          ? {
              shapeShadow: { ...INTENSE_SHADOW },
            }
          : {}),
      };
    }),
  );
}
