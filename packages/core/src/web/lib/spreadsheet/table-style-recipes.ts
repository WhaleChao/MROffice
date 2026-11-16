/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SpreadsheetML built-in table differential formats (not copied SVG artwork).
 * Data adapted from Apache POI presetTableStyles.xml, Apache License 2.0:
 * https://github.com/apache/poi/blob/trunk/poi-ooxml/src/main/resources/org/apache/poi/xssf/usermodel/presetTableStyles.xml
 * https://www.apache.org/licenses/LICENSE-2.0
 * Upstream's dxfId is one-based. Theme/tint references stay semantic here.
 * Shared by the worksheet renderer and both Spreadsheet style galleries.
 */
import { paletteById, normalizeHexColor } from "./color-palette";
import { tintSpreadsheetColor } from "./theme-colors";
// The built-in styles' elements (ATOMS) and which element each style uses
// (RECIPES); the API's file_adapters/xlsx/table_style_recipes.json is the same data.
import RECIPE_DATA from "./table-style-recipes.json";

type Color = string | readonly [number, number];
type Edge = "top" | "right" | "bottom" | "left";
/** SpreadsheetML line styles a table style element may draw (§18.18.3). */
type BorderStyle =
  | "none"
  | "thin"
  | "medium"
  | "thick"
  | "double"
  | "dashed"
  | "dotted"
  | "hair"
  | "dashDot"
  | "dashDotDot"
  | "mediumDashed"
  | "mediumDashDot"
  | "mediumDashDotDot"
  | "slantDashDot";
type Atom = {
  backgroundColor?: Color; color?: Color; bold?: boolean; italic?: boolean; strike?: boolean;
  border?: Partial<Record<Edge | "horizontal" | "vertical", { style: BorderStyle; color?: Color }>>;
};
/**
 * One table style element's differential format: fill, font colour and
 * emphasis, and the element's outer edges plus the inner horizontal /
 * vertical lines. A colour is `#rrggbb` or `[themeIndex, tint]` (0 Background
 * 1, 1 Text 1, 4-9 Accent 1-6), resolved against the workbook theme.
 */
export type TableStyleAtom = Atom;
export type TableStyleBorderStyle = BorderStyle;
export type TableStyleElementType =
  | "wholeTable"
  | "firstColumnStripe"
  | "secondColumnStripe"
  | "firstRowStripe"
  | "secondRowStripe"
  | "lastColumn"
  | "firstColumn"
  | "headerRow"
  | "totalRow"
  | "firstHeaderCell"
  | "lastHeaderCell"
  | "firstTotalCell"
  | "lastTotalCell";
export const TABLE_STYLE_ELEMENT_TYPES: readonly TableStyleElementType[] = [
  "wholeTable",
  "firstColumnStripe",
  "secondColumnStripe",
  "firstRowStripe",
  "secondRowStripe",
  "lastColumn",
  "firstColumn",
  "headerRow",
  "totalRow",
  "firstHeaderCell",
  "lastHeaderCell",
  "firstTotalCell",
  "lastTotalCell",
];
export type TableStyleStripeType =
  | "firstRowStripe"
  | "secondRowStripe"
  | "firstColumnStripe"
  | "secondColumnStripe";
/** A table style as its elements paint it (built-in or a workbook's own). */
export interface TableStyleRecipe {
  readonly elements: Readonly<Partial<Record<TableStyleElementType, TableStyleAtom>>>;
  /** Band sizes other than 1 (`tableStyleElement@size`, 1-9). */
  readonly stripeSizes?: Readonly<Partial<Record<TableStyleStripeType, number>>>;
}

export type TableStyleCellPaint = {
  backgroundColor?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  borderEdges?: Partial<Record<Edge, { style: Exclude<BorderStyle, "none">; color: string }>>;
};

export interface TableStyleLook {
  range: { r1: number; c1: number; r2: number; c2: number };
  headerRow?: boolean;
  totalsRow?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
}

const builtInRecipeCache = new Map<string, TableStyleRecipe | null>();

function builtInRecipe(name: string): TableStyleRecipe | null {
  const cached = builtInRecipeCache.get(name);
  if (cached !== undefined) return cached;
  const ids = RECIPES[name];
  const recipe = ids
    ? {
        elements: Object.fromEntries(
          Object.entries(ids).map(([role, id]) => [role, ATOMS[id]]),
        ) as TableStyleRecipe["elements"],
      }
    : null;
  builtInRecipeCache.set(name, recipe);
  return recipe;
}

/** Native precedence: whole, column stripes, row stripes, edge columns, header/total. */
export function builtInTableCellPaint(
  name: string,
  accents: readonly string[],
  table: TableStyleLook,
  row: number,
  col: number,
): TableStyleCellPaint {
  const recipe = builtInRecipe(name);
  return recipe ? tableRecipeCellPaint(recipe, accents, table, row, col) : {};
}

function stripeSize(recipe: TableStyleRecipe, type: TableStyleStripeType): number {
  const size = recipe.stripeSizes?.[type];
  return typeof size === "number" && Number.isInteger(size) && size >= 1 && size <= 9 ? size : 1;
}

/**
 * Paint one cell of a table from a style's elements, in ECMA-376's order
 * (§18.8.41): whole table, column stripes, row stripes, last then first
 * column, header row, total row, then the four corner cells; a later element
 * wins. Stripes run in bands of their `size` rows / columns.
 */
export function tableRecipeCellPaint(
  recipe: TableStyleRecipe,
  accents: readonly string[],
  table: TableStyleLook,
  row: number,
  col: number,
): TableStyleCellPaint {
  const { r1, c1, r2, c2 } = table.range;
  const lastRow = r2 + (table.totalsRow ? 1 : 0);
  if (row < r1 || row > lastRow || col < c1 || col > c2) return {};
  const defaults = paletteById("office").accents;
  const resolve = (color: Color): string => {
    if (typeof color === "string") return color;
    const [slot, tint] = color;
    const base = slot === 0 ? "#ffffff" : slot === 1 ? "#000000"
      : normalizeHexColor(accents[slot - 4]) ?? defaults[slot - 4] ?? "#000000";
    return tintSpreadsheetColor(base, tint);
  };
  const result: TableStyleCellPaint = {};
  const apply = (role: TableStyleElementType, top: number, left: number, bottom: number, right: number) => {
    const atom = recipe.elements[role];
    if (!atom) return;
    if (atom.backgroundColor) result.backgroundColor = resolve(atom.backgroundColor);
    if (atom.color) result.color = resolve(atom.color);
    for (const key of ["bold", "italic", "strike"] as const) {
      if (atom[key] !== undefined) result[key] = atom[key];
    }
    if (!atom.border) return;
    const edges = {
      top: row === top ? "top" : "horizontal",
      bottom: row === bottom ? "bottom" : "horizontal",
      left: col === left ? "left" : "vertical",
      right: col === right ? "right" : "vertical",
    } as const;
    for (const edge of ["top", "right", "bottom", "left"] as const) {
      const border = atom.border[edges[edge]];
      if (border) {
        result.borderEdges ??= {};
        if (border.style === "none") delete result.borderEdges[edge];
        else result.borderEdges[edge] = { style: border.style, color: border.color ? resolve(border.color) : "#000000" };
      }
    }
  };
  apply("wholeTable", r1, c1, lastRow, c2);
  const dataStart = r1 + (table.headerRow ? 1 : 0);
  if (row >= dataStart && row <= r2) {
    if (table.bandedColumns) {
      const first = stripeSize(recipe, "firstColumnStripe");
      const period = first + stripeSize(recipe, "secondColumnStripe");
      const position = (col - c1) % period;
      const bandStart = col - (position < first ? position : position - first);
      const bandEnd = Math.min(c2, bandStart + (position < first ? first : period - first) - 1);
      apply(position < first ? "firstColumnStripe" : "secondColumnStripe", dataStart, bandStart, r2, bandEnd);
    }
    if (table.bandedRows) {
      const first = stripeSize(recipe, "firstRowStripe");
      const period = first + stripeSize(recipe, "secondRowStripe");
      const position = (row - dataStart) % period;
      const bandStart = row - (position < first ? position : position - first);
      const bandEnd = Math.min(r2, bandStart + (position < first ? first : period - first) - 1);
      apply(position < first ? "firstRowStripe" : "secondRowStripe", bandStart, c1, bandEnd, c2);
    }
  }
  if (table.lastColumn && col === c2) apply("lastColumn", r1, c2, lastRow, c2);
  if (table.firstColumn && col === c1) apply("firstColumn", r1, c1, lastRow, c1);
  if (table.headerRow && row === r1) {
    apply("headerRow", r1, c1, r1, c2);
    if (col === c1) apply("firstHeaderCell", r1, c1, r1, c1);
    if (col === c2) apply("lastHeaderCell", r1, c2, r1, c2);
  }
  if (table.totalsRow && row === lastRow) {
    apply("totalRow", lastRow, c1, lastRow, c2);
    if (col === c1) apply("firstTotalCell", lastRow, c1, lastRow, c1);
    if (col === c2) apply("lastTotalCell", lastRow, c2, lastRow, c2);
  }
  return result;
}
const ATOMS: readonly Atom[] = (RECIPE_DATA as unknown as { atoms: Atom[] }).atoms;


const RECIPES: Readonly<Record<string, Readonly<Record<string, number>>>> =
  (RECIPE_DATA as unknown as { recipes: Record<string, Record<string, number>> }).recipes;
