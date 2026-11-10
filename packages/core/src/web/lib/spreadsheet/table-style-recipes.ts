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
const ATOMS: readonly Atom[] = [
  {"color":[1,0.0],"border":{"top":{"style":"thin","color":[1,0.0]},"bottom":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[1,0.0]},
  {"backgroundColor":[0,-0.1499984740745262]},
  {"color":[4,-0.249977111117893],"border":{"top":{"style":"thin","color":[4,0.0]},"bottom":{"style":"thin","color":[4,0.0]}}},
  {"bold":true,"color":[4,-0.249977111117893],"border":{"bottom":{"style":"thin","color":[4,0.0]}}},
  {"bold":true,"color":[4,-0.249977111117893],"border":{"top":{"style":"thin","color":[4,0.0]}}},
  {"bold":true,"color":[4,-0.249977111117893]},
  {"backgroundColor":[4,0.7999816888943144]},
  {"color":[5,-0.249977111117893],"border":{"top":{"style":"thin","color":[5,0.0]},"bottom":{"style":"thin","color":[5,0.0]}}},
  {"bold":true,"color":[5,-0.249977111117893],"border":{"bottom":{"style":"thin","color":[5,0.0]}}},
  {"bold":true,"color":[5,-0.249977111117893],"border":{"top":{"style":"thin","color":[5,0.0]}}},
  {"bold":true,"color":[5,-0.249977111117893]},
  {"backgroundColor":[5,0.7999816888943144]},
  {"color":[6,-0.249977111117893],"border":{"top":{"style":"thin","color":[6,0.0]},"bottom":{"style":"thin","color":[6,0.0]}}},
  {"bold":true,"color":[6,-0.249977111117893],"border":{"bottom":{"style":"thin","color":[6,0.0]}}},
  {"bold":true,"color":[6,-0.249977111117893],"border":{"top":{"style":"thin","color":[6,0.0]}}},
  {"bold":true,"color":[6,-0.249977111117893]},
  {"backgroundColor":[6,0.7999816888943144]},
  {"color":[7,-0.249977111117893],"border":{"top":{"style":"thin","color":[7,0.0]},"bottom":{"style":"thin","color":[7,0.0]}}},
  {"bold":true,"color":[7,-0.249977111117893],"border":{"bottom":{"style":"thin","color":[7,0.0]}}},
  {"bold":true,"color":[7,-0.249977111117893],"border":{"top":{"style":"thin","color":[7,0.0]}}},
  {"bold":true,"color":[7,-0.249977111117893]},
  {"backgroundColor":[7,0.7999816888943144]},
  {"color":[8,-0.249977111117893],"border":{"top":{"style":"thin","color":[8,0.0]},"bottom":{"style":"thin","color":[8,0.0]}}},
  {"bold":true,"color":[8,-0.249977111117893],"border":{"bottom":{"style":"thin","color":[8,0.0]}}},
  {"bold":true,"color":[8,-0.249977111117893],"border":{"top":{"style":"thin","color":[8,0.0]}}},
  {"bold":true,"color":[8,-0.249977111117893]},
  {"backgroundColor":[8,0.7999816888943144]},
  {"color":[9,-0.249977111117893],"border":{"top":{"style":"thin","color":[9,0.0]},"bottom":{"style":"thin","color":[9,0.0]}}},
  {"bold":true,"color":[9,-0.249977111117893],"border":{"bottom":{"style":"thin","color":[9,0.0]}}},
  {"bold":true,"color":[9,-0.249977111117893],"border":{"top":{"style":"thin","color":[9,0.0]}}},
  {"bold":true,"color":[9,-0.249977111117893]},
  {"backgroundColor":[9,0.7999816888943144]},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[1,0.0]},"right":{"style":"thin","color":[1,0.0]},"top":{"style":"thin","color":[1,0.0]},"bottom":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[1,0.0]}}},
  {"border":{"top":{"style":"thin","color":[1,0.0]}}},
  {"border":{"left":{"style":"thin","color":[1,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[4,0.0]},"right":{"style":"thin","color":[4,0.0]},"top":{"style":"thin","color":[4,0.0]},"bottom":{"style":"thin","color":[4,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[4,0.0]}}},
  {"border":{"top":{"style":"thin","color":[4,0.0]}}},
  {"border":{"left":{"style":"thin","color":[4,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[5,0.0]},"right":{"style":"thin","color":[5,0.0]},"top":{"style":"thin","color":[5,0.0]},"bottom":{"style":"thin","color":[5,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[5,0.0]}}},
  {"border":{"top":{"style":"thin","color":[5,0.0]}}},
  {"border":{"left":{"style":"thin","color":[5,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[6,0.0]},"right":{"style":"thin","color":[6,0.0]},"top":{"style":"thin","color":[6,0.0]},"bottom":{"style":"thin","color":[6,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[6,0.0]}}},
  {"border":{"top":{"style":"thin","color":[6,0.0]}}},
  {"border":{"left":{"style":"thin","color":[6,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[7,0.0]},"right":{"style":"thin","color":[7,0.0]},"top":{"style":"thin","color":[7,0.0]},"bottom":{"style":"thin","color":[7,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[7,0.0]}}},
  {"border":{"top":{"style":"thin","color":[7,0.0]}}},
  {"border":{"left":{"style":"thin","color":[7,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[8,0.0]},"right":{"style":"thin","color":[8,0.0]},"top":{"style":"thin","color":[8,0.0]},"bottom":{"style":"thin","color":[8,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[8,0.0]}}},
  {"border":{"top":{"style":"thin","color":[8,0.0]}}},
  {"border":{"left":{"style":"thin","color":[8,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[9,0.0]},"right":{"style":"thin","color":[9,0.0]},"top":{"style":"thin","color":[9,0.0]},"bottom":{"style":"thin","color":[9,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,0.0]},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"double","color":[9,0.0]}}},
  {"border":{"top":{"style":"thin","color":[9,0.0]}}},
  {"border":{"left":{"style":"thin","color":[9,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[1,0.0]},"right":{"style":"thin","color":[1,0.0]},"top":{"style":"thin","color":[1,0.0]},"bottom":{"style":"thin","color":[1,0.0]},"vertical":{"style":"thin","color":[1,0.0]},"horizontal":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[4,0.0]},"right":{"style":"thin","color":[4,0.0]},"top":{"style":"thin","color":[4,0.0]},"bottom":{"style":"thin","color":[4,0.0]},"vertical":{"style":"thin","color":[4,0.0]},"horizontal":{"style":"thin","color":[4,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[4,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[5,0.0]},"right":{"style":"thin","color":[5,0.0]},"top":{"style":"thin","color":[5,0.0]},"bottom":{"style":"thin","color":[5,0.0]},"vertical":{"style":"thin","color":[5,0.0]},"horizontal":{"style":"thin","color":[5,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[5,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[6,0.0]},"right":{"style":"thin","color":[6,0.0]},"top":{"style":"thin","color":[6,0.0]},"bottom":{"style":"thin","color":[6,0.0]},"vertical":{"style":"thin","color":[6,0.0]},"horizontal":{"style":"thin","color":[6,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[6,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[7,0.0]},"right":{"style":"thin","color":[7,0.0]},"top":{"style":"thin","color":[7,0.0]},"bottom":{"style":"thin","color":[7,0.0]},"vertical":{"style":"thin","color":[7,0.0]},"horizontal":{"style":"thin","color":[7,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[7,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[8,0.0]},"right":{"style":"thin","color":[8,0.0]},"top":{"style":"thin","color":[8,0.0]},"bottom":{"style":"thin","color":[8,0.0]},"vertical":{"style":"thin","color":[8,0.0]},"horizontal":{"style":"thin","color":[8,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[8,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[9,0.0]},"right":{"style":"thin","color":[9,0.0]},"top":{"style":"thin","color":[9,0.0]},"bottom":{"style":"thin","color":[9,0.0]},"vertical":{"style":"thin","color":[9,0.0]},"horizontal":{"style":"thin","color":[9,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"bottom":{"style":"medium","color":[9,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[1,0.0]},"right":{"style":"thin","color":[1,0.0]},"top":{"style":"thin","color":[1,0.0]},"bottom":{"style":"thin","color":[1,0.0]},"horizontal":{"style":"thin","color":[1,0.0]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[4,0.3999755851924192]},"right":{"style":"thin","color":[4,0.3999755851924192]},"top":{"style":"thin","color":[4,0.3999755851924192]},"bottom":{"style":"thin","color":[4,0.3999755851924192]},"horizontal":{"style":"thin","color":[4,0.3999755851924192]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[5,0.3999755851924192]},"right":{"style":"thin","color":[5,0.3999755851924192]},"top":{"style":"thin","color":[5,0.3999755851924192]},"bottom":{"style":"thin","color":[5,0.3999755851924192]},"horizontal":{"style":"thin","color":[5,0.3999755851924192]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[6,0.3999755851924192]},"right":{"style":"thin","color":[6,0.3999755851924192]},"top":{"style":"thin","color":[6,0.3999755851924192]},"bottom":{"style":"thin","color":[6,0.3999755851924192]},"horizontal":{"style":"thin","color":[6,0.3999755851924192]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[7,0.3999755851924192]},"right":{"style":"thin","color":[7,0.3999755851924192]},"top":{"style":"thin","color":[7,0.3999755851924192]},"bottom":{"style":"thin","color":[7,0.3999755851924192]},"horizontal":{"style":"thin","color":[7,0.3999755851924192]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[8,0.3999755851924192]},"right":{"style":"thin","color":[8,0.3999755851924192]},"top":{"style":"thin","color":[8,0.3999755851924192]},"bottom":{"style":"thin","color":[8,0.3999755851924192]},"horizontal":{"style":"thin","color":[8,0.3999755851924192]}}},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[9,0.3999755851924192]},"right":{"style":"thin","color":[9,0.3999755851924192]},"top":{"style":"thin","color":[9,0.3999755851924192]},"bottom":{"style":"thin","color":[9,0.3999755851924192]},"horizontal":{"style":"thin","color":[9,0.3999755851924192]}}},
  {"color":[1,0.0],"backgroundColor":[0,-0.1499984740745262],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[0,-0.3499862666707358]},
  {"color":[1,0.0],"backgroundColor":[4,0.7999816888943144],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[4,0.5999938962981048]},
  {"color":[1,0.0],"backgroundColor":[5,0.7999816888943144],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[5,0.5999938962981048]},
  {"color":[1,0.0],"backgroundColor":[6,0.7999816888943144],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[6,0.5999938962981048]},
  {"color":[1,0.0],"backgroundColor":[7,0.7999816888943144],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[7,0.5999938962981048]},
  {"color":[1,0.0],"backgroundColor":[8,0.7999816888943144],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[8,0.5999938962981048]},
  {"color":[1,0.0],"backgroundColor":[9,0.7999816888943144],"border":{"vertical":{"style":"thin","color":[0,0.0]},"horizontal":{"style":"thin","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,0.0],"border":{"bottom":{"style":"thick","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,0.0],"border":{"top":{"style":"thick","color":[0,0.0]}}},
  {"backgroundColor":[9,0.5999938962981048]},
  {"color":[1,0.0],"border":{"left":{"style":"thin","color":[1,0.0]},"right":{"style":"thin","color":[1,0.0]},"top":{"style":"medium","color":[1,0.0]},"bottom":{"style":"medium","color":[1,0.0]},"vertical":{"style":"thin","color":[1,0.0]},"horizontal":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"border":{"top":{"style":"double","color":[1,0.0]}}},
  {"color":[1,0.0],"border":{"top":{"style":"medium","color":[1,0.0]},"bottom":{"style":"medium","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,0.0],"border":{"bottom":{"style":"medium","color":[1,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[0,-0.1499984740745262],"border":{"left":{"style":"thin","color":[1,0.0]},"right":{"style":"thin","color":[1,0.0]},"top":{"style":"thin","color":[1,0.0]},"bottom":{"style":"thin","color":[1,0.0]},"vertical":{"style":"thin","color":[1,0.0]},"horizontal":{"style":"thin","color":[1,0.0]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[1,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[4,0.7999816888943144],"border":{"left":{"style":"thin","color":[4,0.3999755851924192]},"right":{"style":"thin","color":[4,0.3999755851924192]},"top":{"style":"thin","color":[4,0.3999755851924192]},"bottom":{"style":"thin","color":[4,0.3999755851924192]},"vertical":{"style":"thin","color":[4,0.3999755851924192]},"horizontal":{"style":"thin","color":[4,0.3999755851924192]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[4,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[5,0.7999816888943144],"border":{"left":{"style":"thin","color":[5,0.3999755851924192]},"right":{"style":"thin","color":[5,0.3999755851924192]},"top":{"style":"thin","color":[5,0.3999755851924192]},"bottom":{"style":"thin","color":[5,0.3999755851924192]},"vertical":{"style":"thin","color":[5,0.3999755851924192]},"horizontal":{"style":"thin","color":[5,0.3999755851924192]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[5,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[6,0.7999816888943144],"border":{"left":{"style":"thin","color":[6,0.3999755851924192]},"right":{"style":"thin","color":[6,0.3999755851924192]},"top":{"style":"thin","color":[6,0.3999755851924192]},"bottom":{"style":"thin","color":[6,0.3999755851924192]},"vertical":{"style":"thin","color":[6,0.3999755851924192]},"horizontal":{"style":"thin","color":[6,0.3999755851924192]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[6,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[7,0.7999816888943144],"border":{"left":{"style":"thin","color":[7,0.3999755851924192]},"right":{"style":"thin","color":[7,0.3999755851924192]},"top":{"style":"thin","color":[7,0.3999755851924192]},"bottom":{"style":"thin","color":[7,0.3999755851924192]},"vertical":{"style":"thin","color":[7,0.3999755851924192]},"horizontal":{"style":"thin","color":[7,0.3999755851924192]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[7,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[8,0.7999816888943144],"border":{"left":{"style":"thin","color":[8,0.3999755851924192]},"right":{"style":"thin","color":[8,0.3999755851924192]},"top":{"style":"thin","color":[8,0.3999755851924192]},"bottom":{"style":"thin","color":[8,0.3999755851924192]},"vertical":{"style":"thin","color":[8,0.3999755851924192]},"horizontal":{"style":"thin","color":[8,0.3999755851924192]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[8,0.0]}}},
  {"color":[1,0.0],"backgroundColor":[9,0.7999816888943144],"border":{"left":{"style":"thin","color":[9,0.3999755851924192]},"right":{"style":"thin","color":[9,0.3999755851924192]},"top":{"style":"thin","color":[9,0.3999755851924192]},"bottom":{"style":"thin","color":[9,0.3999755851924192]},"vertical":{"style":"thin","color":[9,0.3999755851924192]},"horizontal":{"style":"thin","color":[9,0.3999755851924192]}}},
  {"bold":true,"color":[1,0.0],"border":{"top":{"style":"medium","color":[9,0.0]}}},
  {"color":[0,0.0],"backgroundColor":[1,0.4499954222235786]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.0],"border":{"bottom":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.1499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[1,0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[1,0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[4,0.0]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,-0.499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,-0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[4,-0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[4,-0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[5,0.0]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,-0.499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,-0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[5,-0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[5,-0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[6,0.0]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,-0.499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,-0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[6,-0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[6,-0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[7,0.0]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,-0.499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,-0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[7,-0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[7,-0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[8,0.0]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,-0.499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,-0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[8,-0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[8,-0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[9,0.0]},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,-0.499984740745262],"border":{"top":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,-0.249977111117893],"border":{"right":{"style":"medium","color":[0,0.0]}}},
  {"bold":true,"color":[0,0.0],"backgroundColor":[9,-0.249977111117893],"border":{"left":{"style":"medium","color":[0,0.0]}}},
  {"backgroundColor":[9,-0.249977111117893]},
  {"color":[0,0.0],"backgroundColor":[1,0.0]},
];

const RECIPES: Readonly<Record<string, Readonly<Record<string, number>>>> = {
  "TableStyleLight1": {"wholeTable":0,"headerRow":1,"totalRow":2,"firstColumn":3,"lastColumn":3,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleLight2": {"wholeTable":5,"headerRow":6,"totalRow":7,"firstColumn":8,"lastColumn":8,"firstRowStripe":9,"firstColumnStripe":9},
  "TableStyleLight3": {"wholeTable":10,"headerRow":11,"totalRow":12,"firstColumn":13,"lastColumn":13,"firstRowStripe":14,"firstColumnStripe":14},
  "TableStyleLight4": {"wholeTable":15,"headerRow":16,"totalRow":17,"firstColumn":18,"lastColumn":18,"firstRowStripe":19,"firstColumnStripe":19},
  "TableStyleLight5": {"wholeTable":20,"headerRow":21,"totalRow":22,"firstColumn":23,"lastColumn":23,"firstRowStripe":24,"firstColumnStripe":24},
  "TableStyleLight6": {"wholeTable":25,"headerRow":26,"totalRow":27,"firstColumn":28,"lastColumn":28,"firstRowStripe":29,"firstColumnStripe":29},
  "TableStyleLight7": {"wholeTable":30,"headerRow":31,"totalRow":32,"firstColumn":33,"lastColumn":33,"firstRowStripe":34,"firstColumnStripe":34},
  "TableStyleLight8": {"wholeTable":35,"headerRow":36,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":38,"secondRowStripe":38,"firstColumnStripe":39,"secondColumnStripe":39},
  "TableStyleLight9": {"wholeTable":40,"headerRow":41,"totalRow":42,"firstColumn":3,"lastColumn":3,"firstRowStripe":43,"secondRowStripe":43,"firstColumnStripe":44,"secondColumnStripe":44},
  "TableStyleLight10": {"wholeTable":45,"headerRow":46,"totalRow":47,"firstColumn":3,"lastColumn":3,"firstRowStripe":48,"secondRowStripe":48,"firstColumnStripe":49,"secondColumnStripe":49},
  "TableStyleLight11": {"wholeTable":50,"headerRow":51,"totalRow":52,"firstColumn":3,"lastColumn":3,"firstRowStripe":53,"secondRowStripe":53,"firstColumnStripe":54,"secondColumnStripe":54},
  "TableStyleLight12": {"wholeTable":55,"headerRow":56,"totalRow":57,"firstColumn":3,"lastColumn":3,"firstRowStripe":58,"secondRowStripe":58,"firstColumnStripe":59,"secondColumnStripe":59},
  "TableStyleLight13": {"wholeTable":60,"headerRow":61,"totalRow":62,"firstColumn":3,"lastColumn":3,"firstRowStripe":63,"secondRowStripe":63,"firstColumnStripe":64,"secondColumnStripe":64},
  "TableStyleLight14": {"wholeTable":65,"headerRow":66,"totalRow":67,"firstColumn":3,"lastColumn":3,"firstRowStripe":68,"secondRowStripe":68,"firstColumnStripe":69,"secondColumnStripe":69},
  "TableStyleLight15": {"wholeTable":70,"headerRow":71,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleLight16": {"wholeTable":72,"headerRow":73,"totalRow":42,"firstColumn":3,"lastColumn":3,"firstRowStripe":9,"firstColumnStripe":9},
  "TableStyleLight17": {"wholeTable":74,"headerRow":75,"totalRow":47,"firstColumn":3,"lastColumn":3,"firstRowStripe":14,"firstColumnStripe":14},
  "TableStyleLight18": {"wholeTable":76,"headerRow":77,"totalRow":52,"firstColumn":3,"lastColumn":3,"firstRowStripe":19,"firstColumnStripe":19},
  "TableStyleLight19": {"wholeTable":78,"headerRow":79,"totalRow":57,"firstColumn":3,"lastColumn":3,"firstRowStripe":24,"firstColumnStripe":24},
  "TableStyleLight20": {"wholeTable":80,"headerRow":81,"totalRow":62,"firstColumn":3,"lastColumn":3,"firstRowStripe":29,"firstColumnStripe":29},
  "TableStyleLight21": {"wholeTable":82,"headerRow":83,"totalRow":67,"firstColumn":3,"lastColumn":3,"firstRowStripe":34,"firstColumnStripe":34},
  "TableStyleMedium1": {"wholeTable":84,"headerRow":36,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium2": {"wholeTable":85,"headerRow":41,"totalRow":42,"firstColumn":3,"lastColumn":3,"firstRowStripe":9,"firstColumnStripe":9},
  "TableStyleMedium3": {"wholeTable":86,"headerRow":46,"totalRow":47,"firstColumn":3,"lastColumn":3,"firstRowStripe":14,"firstColumnStripe":14},
  "TableStyleMedium4": {"wholeTable":87,"headerRow":51,"totalRow":52,"firstColumn":3,"lastColumn":3,"firstRowStripe":19,"firstColumnStripe":19},
  "TableStyleMedium5": {"wholeTable":88,"headerRow":56,"totalRow":57,"firstColumn":3,"lastColumn":3,"firstRowStripe":24,"firstColumnStripe":24},
  "TableStyleMedium6": {"wholeTable":89,"headerRow":61,"totalRow":62,"firstColumn":3,"lastColumn":3,"firstRowStripe":29,"firstColumnStripe":29},
  "TableStyleMedium7": {"wholeTable":90,"headerRow":66,"totalRow":67,"firstColumn":3,"lastColumn":3,"firstRowStripe":34,"firstColumnStripe":34},
  "TableStyleMedium8": {"wholeTable":91,"headerRow":92,"totalRow":93,"firstColumn":36,"lastColumn":36,"firstRowStripe":94,"firstColumnStripe":94},
  "TableStyleMedium9": {"wholeTable":95,"headerRow":96,"totalRow":97,"firstColumn":41,"lastColumn":41,"firstRowStripe":98,"firstColumnStripe":98},
  "TableStyleMedium10": {"wholeTable":99,"headerRow":100,"totalRow":101,"firstColumn":46,"lastColumn":46,"firstRowStripe":102,"firstColumnStripe":102},
  "TableStyleMedium11": {"wholeTable":103,"headerRow":104,"totalRow":105,"firstColumn":51,"lastColumn":51,"firstRowStripe":106,"firstColumnStripe":106},
  "TableStyleMedium12": {"wholeTable":107,"headerRow":108,"totalRow":109,"firstColumn":56,"lastColumn":56,"firstRowStripe":110,"firstColumnStripe":110},
  "TableStyleMedium13": {"wholeTable":111,"headerRow":112,"totalRow":113,"firstColumn":61,"lastColumn":61,"firstRowStripe":114,"firstColumnStripe":114},
  "TableStyleMedium14": {"wholeTable":115,"headerRow":116,"totalRow":117,"firstColumn":66,"lastColumn":66,"firstRowStripe":118,"firstColumnStripe":118},
  "TableStyleMedium15": {"wholeTable":119,"headerRow":120,"totalRow":121,"firstColumn":36,"lastColumn":36,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium16": {"wholeTable":122,"headerRow":123,"totalRow":121,"firstColumn":41,"lastColumn":41,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium17": {"wholeTable":122,"headerRow":124,"totalRow":121,"firstColumn":46,"lastColumn":46,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium18": {"wholeTable":122,"headerRow":125,"totalRow":121,"firstColumn":51,"lastColumn":51,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium19": {"wholeTable":122,"headerRow":126,"totalRow":121,"firstColumn":56,"lastColumn":56,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium20": {"wholeTable":122,"headerRow":127,"totalRow":121,"firstColumn":61,"lastColumn":61,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium21": {"wholeTable":122,"headerRow":128,"totalRow":121,"firstColumn":66,"lastColumn":66,"firstRowStripe":4,"firstColumnStripe":4},
  "TableStyleMedium22": {"wholeTable":129,"headerRow":3,"totalRow":130,"firstColumn":3,"lastColumn":3,"firstRowStripe":94,"firstColumnStripe":94},
  "TableStyleMedium23": {"wholeTable":131,"headerRow":3,"totalRow":132,"firstColumn":3,"lastColumn":3,"firstRowStripe":98,"firstColumnStripe":98},
  "TableStyleMedium24": {"wholeTable":133,"headerRow":3,"totalRow":134,"firstColumn":3,"lastColumn":3,"firstRowStripe":102,"firstColumnStripe":102},
  "TableStyleMedium25": {"wholeTable":135,"headerRow":3,"totalRow":136,"firstColumn":3,"lastColumn":3,"firstRowStripe":106,"firstColumnStripe":106},
  "TableStyleMedium26": {"wholeTable":137,"headerRow":3,"totalRow":138,"firstColumn":3,"lastColumn":3,"firstRowStripe":110,"firstColumnStripe":110},
  "TableStyleMedium27": {"wholeTable":139,"headerRow":3,"totalRow":140,"firstColumn":3,"lastColumn":3,"firstRowStripe":114,"firstColumnStripe":114},
  "TableStyleMedium28": {"wholeTable":141,"headerRow":3,"totalRow":142,"firstColumn":3,"lastColumn":3,"firstRowStripe":118,"firstColumnStripe":118},
  "TableStyleDark1": {"wholeTable":143,"headerRow":144,"totalRow":145,"firstColumn":146,"lastColumn":147,"firstRowStripe":148,"firstColumnStripe":148},
  "TableStyleDark2": {"wholeTable":149,"headerRow":144,"totalRow":150,"firstColumn":151,"lastColumn":152,"firstRowStripe":153,"firstColumnStripe":153},
  "TableStyleDark3": {"wholeTable":154,"headerRow":144,"totalRow":155,"firstColumn":156,"lastColumn":157,"firstRowStripe":158,"firstColumnStripe":158},
  "TableStyleDark4": {"wholeTable":159,"headerRow":144,"totalRow":160,"firstColumn":161,"lastColumn":162,"firstRowStripe":163,"firstColumnStripe":163},
  "TableStyleDark5": {"wholeTable":164,"headerRow":144,"totalRow":165,"firstColumn":166,"lastColumn":167,"firstRowStripe":168,"firstColumnStripe":168},
  "TableStyleDark6": {"wholeTable":169,"headerRow":144,"totalRow":170,"firstColumn":171,"lastColumn":172,"firstRowStripe":173,"firstColumnStripe":173},
  "TableStyleDark7": {"wholeTable":174,"headerRow":144,"totalRow":175,"firstColumn":176,"lastColumn":177,"firstRowStripe":178,"firstColumnStripe":178},
  "TableStyleDark8": {"wholeTable":4,"headerRow":179,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":94,"firstColumnStripe":94},
  "TableStyleDark9": {"wholeTable":9,"headerRow":154,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":98,"firstColumnStripe":98},
  "TableStyleDark10": {"wholeTable":19,"headerRow":164,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":106,"firstColumnStripe":106},
  "TableStyleDark11": {"wholeTable":29,"headerRow":174,"totalRow":37,"firstColumn":3,"lastColumn":3,"firstRowStripe":114,"firstColumnStripe":114},
};
