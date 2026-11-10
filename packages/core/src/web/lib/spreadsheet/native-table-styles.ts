/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * A workbook's own table styles (styles.xml `<tableStyles>`, ECMA-376
 * §18.8.40-41), as the page carries them: per style its name, the elements it
 * paints in the shared table-style recipe model (table-style-recipes.ts) and
 * its stripe band sizes. Excel lists them in the Table Styles gallery's Custom
 * section ("Custom table styles are only stored in the current workbook",
 * Microsoft Support, "Format an Excel table"). The definitions are package
 * data: the editor paints and chooses them, the XLSX writer only ever writes
 * a table's choice (`tableStyleInfo@name`), never the definition.
 */
import {
  TABLE_STYLE_ELEMENT_TYPES,
  type TableStyleAtom,
  type TableStyleBorderStyle,
  type TableStyleElementType,
  type TableStyleRecipe,
  type TableStyleStripeType,
} from "./table-style-recipes";

export interface SpreadsheetNativeTableStyle extends TableStyleRecipe {
  readonly name: string;
  /**
   * False when a member of one of its formats has no place in the recipe
   * model (a pattern fill, an underline, a font size): painted without it,
   * and never materialized onto cells (Convert to Range refuses).
   */
  readonly exact?: false;
}

export const MAX_NATIVE_TABLE_STYLES = 256;
const ELEMENT_TYPES = new Set<string>(TABLE_STYLE_ELEMENT_TYPES);
const STRIPE_TYPES = new Set<string>([
  "firstRowStripe",
  "secondRowStripe",
  "firstColumnStripe",
  "secondColumnStripe",
]);
const BORDER_EDGES = ["top", "right", "bottom", "left", "horizontal", "vertical"] as const;
const BORDER_STYLES = new Set<string>([
  "none",
  "thin",
  "medium",
  "thick",
  "double",
  "dashed",
  "dotted",
  "hair",
  "dashDot",
  "dashDotDot",
  "mediumDashed",
  "mediumDashDot",
  "mediumDashDotDot",
  "slantDashDot",
]);
const THEME_SLOTS = new Set([0, 1, 4, 5, 6, 7, 8, 9]);

type Color = NonNullable<TableStyleAtom["color"]>;

function normalizeColor(raw: unknown): Color | undefined {
  if (typeof raw === "string") {
    return /^#[0-9a-f]{6}$/i.test(raw) ? raw.toLowerCase() : undefined;
  }
  if (
    Array.isArray(raw) &&
    raw.length === 2 &&
    typeof raw[0] === "number" &&
    THEME_SLOTS.has(raw[0]) &&
    typeof raw[1] === "number" &&
    Number.isFinite(raw[1]) &&
    raw[1] >= -1 &&
    raw[1] <= 1
  ) {
    return [raw[0], raw[1]] as const;
  }
  return undefined;
}

function normalizeAtom(raw: unknown): TableStyleAtom | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const source = raw as Record<string, unknown>;
  const atom: {
    -readonly [K in keyof TableStyleAtom]: TableStyleAtom[K];
  } = {};
  const backgroundColor = normalizeColor(source.backgroundColor);
  if (backgroundColor) atom.backgroundColor = backgroundColor;
  const color = normalizeColor(source.color);
  if (color) atom.color = color;
  for (const key of ["bold", "italic", "strike"] as const) {
    if (typeof source[key] === "boolean") atom[key] = source[key] as boolean;
  }
  const rawBorder = source.border;
  if (rawBorder && typeof rawBorder === "object" && !Array.isArray(rawBorder)) {
    const border: NonNullable<TableStyleAtom["border"]> = {};
    for (const edge of BORDER_EDGES) {
      const entry = (rawBorder as Record<string, unknown>)[edge];
      if (!entry || typeof entry !== "object") continue;
      const { style, color: edgeColor } = entry as Record<string, unknown>;
      if (typeof style !== "string" || !BORDER_STYLES.has(style)) continue;
      const normalized = normalizeColor(edgeColor);
      border[edge] = {
        style: style as TableStyleBorderStyle,
        ...(normalized ? { color: normalized } : {}),
      };
    }
    if (Object.keys(border).length > 0) atom.border = border;
  }
  return Object.keys(atom).length > 0 ? atom : null;
}

/** The page's `nativeTableStyles`, normalized; [] when absent or invalid. */
export function normalizeNativeTableStyles(
  raw: unknown,
): SpreadsheetNativeTableStyle[] {
  if (!Array.isArray(raw)) return [];
  const out: SpreadsheetNativeTableStyle[] = [];
  const seen = new Set<string>();
  for (const candidate of raw) {
    if (out.length >= MAX_NATIVE_TABLE_STYLES) break;
    if (!candidate || typeof candidate !== "object") continue;
    const source = candidate as Record<string, unknown>;
    const name = source.name;
    if (
      typeof name !== "string" ||
      name.length === 0 ||
      name.length > 255 ||
      /[\u0000-\u001f]/.test(name) ||
      seen.has(name.trim().toLowerCase())
    ) {
      continue;
    }
    seen.add(name.trim().toLowerCase());
    const elements: Partial<Record<TableStyleElementType, TableStyleAtom>> = {};
    if (source.elements && typeof source.elements === "object") {
      for (const [type, value] of Object.entries(
        source.elements as Record<string, unknown>,
      )) {
        if (!ELEMENT_TYPES.has(type)) continue;
        const atom = normalizeAtom(value);
        if (atom) elements[type as TableStyleElementType] = atom;
      }
    }
    const stripeSizes: Partial<Record<TableStyleStripeType, number>> = {};
    if (source.stripeSizes && typeof source.stripeSizes === "object") {
      for (const [type, value] of Object.entries(
        source.stripeSizes as Record<string, unknown>,
      )) {
        if (
          STRIPE_TYPES.has(type) &&
          typeof value === "number" &&
          Number.isInteger(value) &&
          value >= 2 &&
          value <= 9
        ) {
          stripeSizes[type as TableStyleStripeType] = value;
        }
      }
    }
    out.push({
      name,
      elements,
      ...(Object.keys(stripeSizes).length > 0 ? { stripeSizes } : {}),
      ...(source.exact === false ? { exact: false as const } : {}),
    });
  }
  return out;
}

/** The workbook style a table names, matched as Excel matches style names. */
export function findNativeTableStyle(
  styles: readonly SpreadsheetNativeTableStyle[] | undefined,
  name: string | undefined,
): SpreadsheetNativeTableStyle | undefined {
  if (!styles?.length || !name) return undefined;
  const folded = name.trim().toLowerCase();
  return styles.find((style) => style.name.trim().toLowerCase() === folded);
}
