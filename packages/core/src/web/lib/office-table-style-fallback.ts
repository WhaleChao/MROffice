/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Table Styles gallery while the measured looks are away.
 *
 * The looks measured from real Office are served to a signed-in page only
 * (`lib/office-data/store.ts`): a guest page, the marketing site's editors
 * among them, never has them, and a signed-in page may open before they
 * land. Until this recipe the gallery then opened as 74 named, empty tiles
 * and no style could be applied.
 *
 * This is MRAI's own recipe of the gallery's fourteen families in their
 * colour columns: which regions a Light, Medium, Dark or Themed style fills,
 * bolds and rules, drawn with the standard 20 / 40 / 60 % tints and shades.
 * A readable stand-in, not the measured catalogue: a Themed style sits on a
 * flat accent instead of Office's gradient and shadow, the corner cells take
 * no exceptions, and a line is 1.5 or 2.5 pt. Office's own name and GUID
 * (`office-table-style-index.ts`) identify every style, so a table styled
 * from here exports as the built-in style it names (the writers resolve the
 * GUID or the name against their own catalogue), a Docs table draws from the
 * gallery by id and repaints when the catalogue lands, and a deck's table
 * that carries this look takes the style's own when the looks arrive
 * (`slidesWithServedTableStyleLooks`).
 *
 * One recipe, two projections: the Presentation catalogue's own shape
 * (scheme-colour tokens with DrawingML transforms, resolved by
 * `presentation-table-web-gallery.ts`), and Word table styles with theme
 * references (`office-table-style-word.ts`), whose tint and shade bytes are
 * fitted to the same resolved colours as the generated Word twin's are.
 */
import {
  officeColorChannelsFromHex,
  officeColorHexFromChannels,
  officeColorTransformChannels,
  type OfficeColorChannels,
} from "./office-color-transform";
import {
  PRESENTATION_TABLE_STYLE_INDEX,
  WORD_TABLE_STYLE_DEFAULT_THEME,
  WORD_TABLE_STYLE_INDEX,
  type OfficeTableStyleGroup,
} from "./office-table-style-index";

/** The scheme colours the recipe names (PowerPoint's slots). */
export type FallbackThemeSlot =
  | "dk1"
  | "lt1"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6";

/**
 * A colour of the recipe: a scheme slot, lightened or darkened by one
 * DrawingML transform (per-cent thousandths, as `a:tint val="20000"`), or
 * laid over what is under it at an alpha.
 */
export interface FallbackColor {
  readonly slot: FallbackThemeSlot;
  readonly tint?: number;
  readonly shade?: number;
  readonly alpha?: number;
}

export type FallbackEdge = "left" | "right" | "top" | "bottom" | "insideH" | "insideV";

export interface FallbackLine {
  readonly style: "solid" | "double" | "none";
  readonly color?: FallbackColor;
  /** Points, as a Presentation look states a width. */
  readonly width: number;
}

export interface FallbackPart {
  /** A fill, or `"none"` for an explicit clear. */
  readonly fill?: FallbackColor | "none";
  readonly color?: FallbackColor;
  readonly bold?: true;
  readonly borders?: Partial<Record<FallbackEdge, FallbackLine>>;
}

export type FallbackPartName =
  | "wholeTbl"
  | "band1H"
  | "band2H"
  | "band1V"
  | "band2V"
  | "firstRow"
  | "lastRow"
  | "firstCol"
  | "lastCol"
  | "nwCell"
  | "neCell"
  | "swCell"
  | "seCell";

export interface FallbackTableStyle {
  readonly slot: number;
  readonly group: OfficeTableStyleGroup;
  /** The built-in style's GUID (PowerPoint's identity). */
  readonly presentationStyleId: string;
  /** PowerPoint's (zh) name, what the gallery tile is labelled from. */
  readonly presentationStyleName: string;
  /** The preferred `w:styleId` and `w:name` (Word's identity). */
  readonly wordStyleId: string;
  readonly name: string;
  /** A flat table background under the cells (the Themed styles). */
  readonly background?: FallbackColor;
  readonly parts: Partial<Record<FallbackPartName, FallbackPart>>;
}

/** The theme the recipe resolves in (Office 2023, the gallery's own). */
export const FALLBACK_TABLE_STYLE_THEME: Readonly<Record<string, string>> = Object.freeze({
  dk1: WORD_TABLE_STYLE_DEFAULT_THEME.text1,
  lt1: WORD_TABLE_STYLE_DEFAULT_THEME.background1,
  dk2: WORD_TABLE_STYLE_DEFAULT_THEME.text2,
  lt2: WORD_TABLE_STYLE_DEFAULT_THEME.background2,
  accent1: WORD_TABLE_STYLE_DEFAULT_THEME.accent1,
  accent2: WORD_TABLE_STYLE_DEFAULT_THEME.accent2,
  accent3: WORD_TABLE_STYLE_DEFAULT_THEME.accent3,
  accent4: WORD_TABLE_STYLE_DEFAULT_THEME.accent4,
  accent5: WORD_TABLE_STYLE_DEFAULT_THEME.accent5,
  accent6: WORD_TABLE_STYLE_DEFAULT_THEME.accent6,
  hlink: "#467886",
  folHlink: "#96607d",
});

const DK1: FallbackColor = { slot: "dk1" };
const LT1: FallbackColor = { slot: "lt1" };
const EDGES: readonly FallbackEdge[] = ["left", "right", "top", "bottom", "insideH", "insideV"];
const NONE: FallbackLine = { style: "none", width: 1 };

const tint = (color: FallbackColor, value: number): FallbackColor => ({ slot: color.slot, tint: value });
const shade = (color: FallbackColor, value: number): FallbackColor => ({ slot: color.slot, shade: value });
const alpha = (color: FallbackColor, value: number): FallbackColor => ({ slot: color.slot, alpha: value });
const solid = (color: FallbackColor, width = 1.5): FallbackLine => ({ style: "solid", color, width });
const double = (color: FallbackColor): FallbackLine => ({ style: "double", color, width: 5.5 });
const every = (line: FallbackLine): Record<FallbackEdge, FallbackLine> =>
  Object.fromEntries(EDGES.map((edge) => [edge, line])) as Record<FallbackEdge, FallbackLine>;
/** The outline only: the inside lines cleared. */
const outline = (line: FallbackLine): Record<FallbackEdge, FallbackLine> => ({
  ...every(NONE),
  left: line,
  right: line,
  top: line,
  bottom: line,
});
const bold = { bold: true as const };

type Recipe = Pick<FallbackTableStyle, "background" | "parts">;

/**
 * The family's recipe in a colour column: `accent` is the column's accent
 * (null in the neutral column), `second` the header accent of a Dark Style 2
 * pair ("Dark Style 2 - Accent 1/Accent 2").
 */
function familyRecipe(
  family: string,
  accent: FallbackColor | null,
  second: FallbackColor | null,
): Recipe {
  // The column's colour where the neutral column draws in text colour.
  const c = accent ?? DK1;
  switch (family) {
    case "No Style - No Grid":
      return { parts: { wholeTbl: { fill: "none", color: DK1, borders: every(NONE) } } };
    case "No Style - Table Grid":
      return { parts: { wholeTbl: { fill: "none", color: DK1, borders: every(solid(DK1)) } } };
    case "Themed Style 1":
      return {
        background: tint(c, 73000),
        parts: {
          wholeTbl: { fill: "none", color: DK1, borders: every(solid(c)) },
          band1H: { fill: alpha(c, 40000) },
          band1V: { fill: alpha(c, 40000) },
          firstRow: { fill: c, color: LT1, ...bold, borders: { bottom: solid(LT1, 2.5) } },
          lastRow: { ...bold, borders: { top: solid(c, 2.5) } },
          firstCol: { ...bold, borders: { right: solid(c, 2.5) } },
          lastCol: { ...bold, borders: { left: solid(c, 2.5) } },
        },
      };
    case "Themed Style 2":
      return {
        background: c,
        parts: {
          wholeTbl: { fill: "none", color: LT1, borders: outline(solid(tint(c, 50000))) },
          band1H: { fill: alpha(LT1, 20000) },
          band1V: { fill: alpha(LT1, 20000) },
          firstRow: { ...bold, borders: { bottom: solid(LT1, 2.5) } },
          lastRow: { ...bold, borders: { top: solid(LT1, 2.5) } },
          firstCol: { ...bold, borders: { right: solid(LT1, 2.5) } },
          lastCol: { ...bold, borders: { left: solid(LT1, 2.5) } },
        },
      };
    case "Light Style 1":
      return {
        parts: {
          wholeTbl: { fill: "none", color: DK1, borders: { ...every(NONE), top: solid(c), bottom: solid(c) } },
          band1H: { fill: alpha(c, 20000) },
          band1V: { fill: alpha(c, 20000) },
          firstRow: { ...bold, borders: { bottom: solid(c) } },
          lastRow: { ...bold, borders: { top: solid(c) } },
          firstCol: bold,
          lastCol: bold,
        },
      };
    case "Light Style 2":
      return {
        parts: {
          wholeTbl: { fill: "none", color: DK1, borders: outline(solid(c)) },
          band1H: { borders: { top: solid(c), bottom: solid(c) } },
          band1V: { borders: { left: solid(c), right: solid(c) } },
          band2V: { borders: { left: solid(c), right: solid(c) } },
          // Office's header here is the theme's first fill style in the
          // column's colour (`a:fillRef idx="1"`): a solid accent bar.
          firstRow: { fill: c, color: LT1, ...bold },
          lastRow: { ...bold, borders: { top: double(c) } },
          firstCol: bold,
          lastCol: bold,
        },
      };
    case "Light Style 3":
      return {
        parts: {
          wholeTbl: { fill: "none", color: DK1, borders: every(solid(c)) },
          band1H: { fill: alpha(c, 20000) },
          band1V: { fill: alpha(c, 20000) },
          firstRow: { ...bold, borders: { bottom: solid(c, 2.5) } },
          lastRow: { ...bold, borders: { top: double(c) } },
          firstCol: bold,
          lastCol: bold,
        },
      };
    case "Medium Style 1":
      return {
        parts: {
          wholeTbl: { fill: LT1, color: DK1, borders: { ...outline(solid(c)), insideH: solid(c) } },
          band1H: { fill: tint(c, 20000) },
          band1V: { fill: tint(c, 20000) },
          firstRow: { fill: c, color: LT1, ...bold },
          lastRow: { fill: LT1, ...bold, borders: { top: double(c) } },
          firstCol: bold,
          lastCol: bold,
        },
      };
    case "Medium Style 2":
      return {
        parts: {
          wholeTbl: { fill: tint(c, 20000), color: DK1, borders: every(solid(LT1)) },
          band1H: { fill: tint(c, 40000) },
          band1V: { fill: tint(c, 40000) },
          firstRow: { fill: c, color: LT1, ...bold, borders: { bottom: solid(LT1, 2.5) } },
          lastRow: { fill: c, color: LT1, ...bold, borders: { top: solid(LT1, 2.5) } },
          firstCol: { fill: c, color: LT1, ...bold },
          lastCol: { fill: c, color: LT1, ...bold },
        },
      };
    case "Medium Style 3":
      return {
        parts: {
          wholeTbl: { fill: LT1, color: DK1, borders: { ...every(NONE), top: solid(DK1, 2.5), bottom: solid(DK1, 2.5) } },
          band1H: { fill: tint(DK1, 20000) },
          band1V: { fill: tint(DK1, 20000) },
          firstRow: { fill: c, color: LT1, ...bold, borders: { bottom: solid(DK1, 2.5) } },
          lastRow: { fill: LT1, ...bold, borders: { top: double(DK1) } },
          firstCol: { fill: c, color: LT1, ...bold },
          lastCol: { fill: c, color: LT1, ...bold },
          swCell: { color: DK1, ...bold },
          seCell: { color: DK1, ...bold },
        },
      };
    case "Medium Style 4":
      return {
        parts: {
          wholeTbl: { fill: tint(c, 20000), color: DK1, borders: every(solid(c)) },
          band1H: { fill: tint(c, 40000) },
          band1V: { fill: tint(c, 40000) },
          firstRow: { fill: tint(c, 20000), ...bold },
          lastRow: { fill: tint(c, 20000), ...bold, borders: { top: solid(c, 2.5) } },
          firstCol: bold,
          lastCol: bold,
        },
      };
    case "Dark Style 1": {
      // The neutral column lightens text colour; an accent column darkens
      // the accent.
      const body = accent ?? tint(DK1, 20000);
      const band = accent ? shade(accent, 60000) : tint(DK1, 40000);
      const foot = accent ? shade(accent, 40000) : tint(DK1, 60000);
      const side = accent ? shade(accent, 60000) : tint(DK1, 60000);
      return {
        parts: {
          wholeTbl: { fill: body, color: LT1, borders: every(NONE) },
          band1H: { fill: band },
          band1V: { fill: band },
          firstRow: { fill: DK1, ...bold, borders: { bottom: solid(LT1, 2.5) } },
          lastRow: { fill: foot, ...bold, borders: { top: solid(LT1, 2.5) } },
          firstCol: { fill: side, ...bold, borders: { right: solid(LT1, 2.5) } },
          lastCol: { fill: side, ...bold, borders: { left: solid(LT1, 2.5) } },
        },
      };
    }
    case "Dark Style 2":
      return {
        parts: {
          wholeTbl: { fill: tint(c, 20000), color: DK1, borders: every(NONE) },
          band1H: { fill: tint(c, 40000) },
          band1V: { fill: tint(c, 40000) },
          firstRow: { fill: second ?? DK1, color: LT1, ...bold },
          lastRow: { fill: tint(c, 20000), ...bold, borders: { top: double(DK1) } },
          firstCol: bold,
          lastCol: bold,
        },
      };
    default:
      throw new Error(`office-table-style-fallback: no recipe for "${family}"`);
  }
}

/** Word's name of a gallery style: its family and its colour column. */
const FAMILY_NAME =
  /^(No Style - No Grid|No Style - Table Grid|Themed Style [12]|Light Style [123]|Medium Style [1-4]|Dark Style [12])(?: - Accent ([1-6])(?:\/Accent ([1-6]))?)?$/;

let styles: readonly FallbackTableStyle[] | null = null;

/** The 74 gallery styles as the recipe draws them, by slot. */
export function fallbackTableStyles(): readonly FallbackTableStyle[] {
  if (styles) return styles;
  const presentationBySlot = new Map(
    PRESENTATION_TABLE_STYLE_INDEX.map(([slot, , styleId, styleName]) => [slot, { styleId, styleName }]),
  );
  styles = WORD_TABLE_STYLE_INDEX.map(([slot, group, wordStyleId, name]) => {
    const match = FAMILY_NAME.exec(name);
    const presentation = presentationBySlot.get(slot);
    if (!match || !presentation) {
      throw new Error(`office-table-style-fallback: no recipe for slot ${slot} "${name}"`);
    }
    const accent = match[2] ? ({ slot: `accent${match[2]}` } as FallbackColor) : null;
    const second = match[3] ? ({ slot: `accent${match[3]}` } as FallbackColor) : null;
    return {
      slot,
      group,
      presentationStyleId: presentation.styleId,
      presentationStyleName: presentation.styleName,
      wordStyleId,
      name,
      ...familyRecipe(match[1], accent, second),
    };
  });
  return styles;
}

export function fallbackTableStyleBySlot(slot: number): FallbackTableStyle | null {
  return fallbackTableStyles().find((style) => style.slot === slot) ?? null;
}

/**
 * A recipe colour as a hex, resolved in a theme (the gallery's own by
 * default): the slot's colour through its transform, an alpha laid over
 * `under` (what the cell shows beneath the part: the table's fill, its
 * background, or the white page).
 */
export function fallbackColorHex(
  color: FallbackColor,
  under: string | null = null,
  theme: Readonly<Record<string, string>> = FALLBACK_TABLE_STYLE_THEME,
): string {
  let channels: OfficeColorChannels =
    officeColorChannelsFromHex(theme[color.slot] ?? FALLBACK_TABLE_STYLE_THEME[color.slot]) ?? [0, 0, 0];
  if (color.tint !== undefined) channels = officeColorTransformChannels(channels, "tint", color.tint) ?? channels;
  if (color.shade !== undefined) channels = officeColorTransformChannels(channels, "shade", color.shade) ?? channels;
  if (color.alpha !== undefined) {
    const opacity = Math.min(Math.max(color.alpha / 100000, 0), 1);
    const beneath = officeColorChannelsFromHex(under ?? "#ffffff") ?? [1, 1, 1];
    channels = [0, 1, 2].map((i) => channels[i] * opacity + beneath[i] * (1 - opacity)) as unknown as OfficeColorChannels;
  }
  return officeColorHexFromChannels(channels);
}

/** A colour token of the Presentation catalogue's shape. */
export interface TableStyleCatalogToken {
  readonly type: string;
  readonly value: Record<string, string>;
  readonly transforms: readonly (readonly string[])[];
  readonly base?: string;
}

/** A style of the Presentation catalogue's shape: its colours are `$n`
 * references into `colors`, resolved against a theme when read. */
export interface TableStyleCatalogStyle {
  readonly slot: number;
  readonly group: string;
  readonly label: string;
  readonly styleId: string;
  readonly styleName: string;
  readonly background?: string;
  readonly parts: Record<string, Record<string, unknown>>;
  readonly [key: string]: unknown;
}

/** The served Presentation catalogue's shape, which the fallback shares. */
export interface TableStyleCatalog {
  readonly source?: string;
  readonly defaults: Readonly<Record<string, string>>;
  readonly colors: readonly TableStyleCatalogToken[];
  readonly styles: readonly TableStyleCatalogStyle[];
}

let presentationCatalog: TableStyleCatalog | null = null;

/**
 * The recipe in the Presentation catalogue's shape, built once: the gallery
 * reads it exactly as it reads the served catalogue, so a fallback look is
 * resolved against a deck's accents by the same arithmetic.
 */
export function fallbackPresentationTableCatalog(): TableStyleCatalog {
  if (presentationCatalog) return presentationCatalog;
  const colors: TableStyleCatalogToken[] = [];
  const indexByKey = new Map<string, number>();
  const ref = (color: FallbackColor): string => {
    const key = `${color.slot}|${color.tint ?? ""}|${color.shade ?? ""}|${color.alpha ?? ""}`;
    let index = indexByKey.get(key);
    if (index === undefined) {
      index = colors.length;
      indexByKey.set(key, index);
      colors.push({
        type: "schemeClr",
        value: { val: color.slot },
        transforms: [
          ...(color.tint !== undefined ? [["tint", String(color.tint)]] : []),
          ...(color.shade !== undefined ? [["shade", String(color.shade)]] : []),
          ...(color.alpha !== undefined ? [["alpha", String(color.alpha)]] : []),
        ],
      });
    }
    return `$${index}`;
  };
  const line = (value: FallbackLine): Record<string, unknown> =>
    value.style === "none" || !value.color
      ? { style: "none", color: "#000000", width: 1 }
      : { style: value.style, color: ref(value.color), width: value.width };
  const part = (value: FallbackPart): Record<string, unknown> => ({
    ...(value.fill === "none" ? { fill: "none" } : value.fill ? { fill: ref(value.fill) } : {}),
    ...(value.color ? { color: ref(value.color) } : {}),
    ...(value.bold ? { bold: true } : {}),
    ...(value.borders
      ? {
          borders: Object.fromEntries(
            Object.entries(value.borders).map(([edge, border]) => [edge, line(border)]),
          ),
        }
      : {}),
  });
  const catalogStyles = fallbackTableStyles().map(
    (style): TableStyleCatalogStyle => ({
      slot: style.slot,
      group: style.group,
      label: style.presentationStyleName,
      styleId: style.presentationStyleId,
      styleName: style.presentationStyleName,
      ...(style.background ? { background: ref(style.background) } : {}),
      parts: Object.fromEntries(
        Object.entries(style.parts).map(([name, value]) => [name, part(value)]),
      ),
    }),
  );
  presentationCatalog = {
    source: "office-table-style-fallback",
    defaults: FALLBACK_TABLE_STYLE_THEME,
    colors,
    styles: catalogStyles,
  };
  return presentationCatalog;
}

/** Whether a catalogue is the recipe's (not the served, measured one). */
export function isFallbackPresentationTableCatalog(catalog: unknown): boolean {
  return presentationCatalog !== null && catalog === presentationCatalog;
}
