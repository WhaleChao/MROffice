/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * SmartArt Design > Change Colors: Office's own colour variations.
 *
 * Microsoft Support ("Learn more about SmartArt graphics"): the Change Colors
 * gallery "provides a range of different color options for a SmartArt
 * graphic, each one applying one or more theme colors in a different way to
 * the shapes"; use Colorful to "emphasize distinct steps" of a process and the
 * Accent options for a cycle. The gallery is grouped Primary Theme Colors,
 * Colorful, then Accent 1 to Accent 6.
 *
 * Each variation is one of Office's colour definitions (`dgm:colorsDef`,
 * ECMA-376 Part 1 §21.4.4, urn `.../2005/8/colors/<id>`). What a definition
 * says for the item shapes (`styleLbl name="node1"`), the flow arrows
 * (`sibTrans2D1`) and the text is reproduced below as data, read off the 38
 * definitions PowerPoint itself wrote (see `DEFINITIONS`):
 *
 * - `accentN_1` Colored Outline, `accentN_2` Colored Fill (Office's default,
 *   `accent1_2`), `accentN_3` Gradient Range (fills SPAN a shade of the accent
 *   to a tint of it), `accentN_4` Gradient Loop (fills CYCLE: out to the tint
 *   and back), `accentN_5` Transparent Gradient Range (fills span 90 % to 50 %
 *   opacity of the accent).
 * - `accent0_1` / `accent0_2` Dark 1 / Dark 2 Outline, `accent0_3` Dark 2 Fill.
 * - `colorful1` Colorful - Accent Colors: fills REPEAT `accent2` .. `accent6`;
 *   `colorfulK` (K = 2..5) Colorful Range - Accent Colors K to K+1: fills SPAN
 *   `accentK` to `accentK+1`.
 *
 * How a list is spread across the shapes was measured on Office's own cached
 * drawings (`dsp:drawing`), which state every shape's colour as the list's
 * first colour plus HSL offsets: see {@link blend} and {@link listColor}.
 */

import {
  officeColorChannelsFromHex,
  officeColorHexFromChannels,
  officeColorTransformChannels,
  rgbToHls,
  type OfficeColorChannels,
} from "../office-color-transform";

import {
  DEFAULT_SMARTART_STYLE,
  type SmartArtNodePaint,
  type SmartArtStyle,
} from "./model";
import type { SmartArtThemeEffectShadows, SmartArtThemeFillStyles } from "./theme-effects";

/** The theme colours a variation is resolved against. */
export interface SmartArtColorTheme {
  readonly dk1: string;
  readonly lt1: string;
  readonly dk2: string;
  readonly lt2: string;
  /** accent1 .. accent6. */
  readonly accents: readonly string[];
  /** The theme's body face (`+mn-lt`), which sets how tall a diagram's lines stand. */
  readonly bodyFace?: string;
  /**
   * The face the theme gives Chinese words (its minor `Hans` script face, else
   * its `ea` face; "" for none). Not an East Asian face -- Calibri, as Excel's
   * every-layout workbook states -- means the words fall back to the system's
   * CJK face, whose line stands taller (`ms-engine/text` FALLBACK_CJK_LINES).
   */
  readonly eastAsianFace?: string;
  /** The theme's effect styles' shadows (`theme-effects`), what an `a:effectRef` draws. */
  readonly effectShadows?: SmartArtThemeEffectShadows;
  /** The theme's fill styles (`theme-effects`), what an `a:fillRef` draws. */
  readonly fillStyles?: SmartArtThemeFillStyles;
}

/** Office's own theme (2013 - 2022), whose accent 1 is the palette's #4472c4. */
export const OFFICE_SMARTART_COLOR_THEME: SmartArtColorTheme = /* own data: OFFICE_SMARTART_COLOR_THEME */ {"dk1":"#000000","lt1":"#ffffff","dk2":"#1f2a37","lt2":"#eef1f4","accents":["#2f6fd6","#e07b39","#3a9d6b","#8c5cc9","#d04b6b","#2c9aa8"]};

/**
 * The theme a diagram is drawn through: the host's own (its colours, faces and
 * style matrix -- `effectShadows` / `fillStyles`, `theme-effects`) over
 * Office's, with `accent` as accent 1 when the host states no accents. The one
 * merge the conversion, every block host's rebuild and the Docs rebuild share,
 * so a diagram converted through its file's theme is rebuilt through it too.
 */
export function smartArtHostColorTheme(
  theme: Partial<SmartArtColorTheme> | null | undefined,
  accent: string | undefined,
): SmartArtColorTheme {
  const base = OFFICE_SMARTART_COLOR_THEME;
  return {
    dk1: theme?.dk1 ?? base.dk1,
    lt1: theme?.lt1 ?? base.lt1,
    dk2: theme?.dk2 ?? base.dk2,
    lt2: theme?.lt2 ?? base.lt2,
    accents: theme?.accents?.length ? theme.accents : [accent ?? base.accents[0], ...base.accents.slice(1)],
    ...(theme?.bodyFace ? { bodyFace: theme.bodyFace } : {}),
    ...(typeof theme?.eastAsianFace === "string" ? { eastAsianFace: theme.eastAsianFace } : {}),
    ...(theme?.effectShadows ? { effectShadows: theme.effectShadows } : {}),
    ...(theme?.fillStyles ? { fillStyles: theme.fillStyles } : {}),
  };
}

type SchemeName = "dk1" | "lt1" | "dk2" | "lt2" | `accent${1 | 2 | 3 | 4 | 5 | 6}`;
/** A colour member Office's definitions use: two colour transforms and opacity. */
type ColorMod = readonly [kind: "shade" | "tint" | "alpha", value: number];
/** One `a:schemeClr`: the scheme slot, then its members in document order. */
type SchemeColor = readonly [scheme: SchemeName, ...mods: ColorMod[]];
/**
 * A `dgm:fillClrLst` / `dgm:linClrLst` (ECMA-376 §21.4.4.7), with its
 * `ST_ClrAppMethod`: `span` (the default) spreads the list across the shapes,
 * `cycle` runs to the list's end and back, `repeat` steps through it.
 */
interface ColorList {
  readonly method: "span" | "cycle" | "repeat";
  readonly colors: readonly SchemeColor[];
}
interface VariantDefinition {
  /** The item shapes (`styleLbl name="node1"`). */
  readonly fill: ColorList;
  readonly line: ColorList;
  /** The item text; absent is the light text of a filled shape (`lt1`). */
  readonly text: ColorList | null;
  /** The flow arrows between items (`sibTrans2D1`). */
  readonly arrow: ColorList;
}

export type SmartArtColorGroupId =
  | "primary"
  | "colorful"
  | `accent${1 | 2 | 3 | 4 | 5 | 6}`;

export interface SmartArtColorVariant {
  /** Office's own id, the last segment of the colour definition's urn. */
  readonly id: string;
  readonly group: SmartArtColorGroupId;
  /** Which label the gallery shows (with `accent` / `from` / `to`). */
  readonly name:
    | "dark1Outline"
    | "dark2Outline"
    | "dark2Fill"
    | "colorful1"
    | "colorfulRange"
    | "coloredOutline"
    | "coloredFill"
    | "gradientRange"
    | "gradientLoop"
    | "transparentGradientRange";
  readonly accent?: number;
  readonly from?: number;
  readonly to?: number;
  readonly definition: VariantDefinition;
}













const DEFINITIONS: Readonly<Record<string, VariantDefinition>> = /* own data: DEFINITIONS */ {"accent0_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["dk1"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["dk1",["tint",50000]]]}},"accent0_2":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["dk2"]]},"text":{"method":"repeat","colors":[["dk2"]]},"arrow":{"method":"repeat","colors":[["dk2",["tint",50000]]]}},"accent0_3":{"fill":{"method":"repeat","colors":[["dk2"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["dk2",["tint",50000]]]}},"colorful1":{"fill":{"method":"repeat","colors":[["accent1"],["accent2"],["accent3"],["accent4"],["accent5"],["accent6"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent1",["tint",50000]],["accent2",["tint",50000]],["accent3",["tint",50000]],["accent4",["tint",50000]],["accent5",["tint",50000]],["accent6",["tint",50000]]]}},"colorful2":{"fill":{"method":"span","colors":[["accent2"],["accent3"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent2",["tint",50000]],["accent3",["tint",50000]]]}},"colorful3":{"fill":{"method":"span","colors":[["accent3"],["accent4"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent3",["tint",50000]],["accent4",["tint",50000]]]}},"colorful4":{"fill":{"method":"span","colors":[["accent4"],["accent5"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent4",["tint",50000]],["accent5",["tint",50000]]]}},"colorful5":{"fill":{"method":"span","colors":[["accent5"],["accent6"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent5",["tint",50000]],["accent6",["tint",50000]]]}},"accent1_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["accent1"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent1",["tint",50000]]]}},"accent1_2":{"fill":{"method":"repeat","colors":[["accent1"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent1",["tint",50000]]]}},"accent1_3":{"fill":{"method":"span","colors":[["accent1",["shade",70000]],["accent1",["tint",60000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent1",["shade",70000]],["accent1",["tint",60000]]]}},"accent1_4":{"fill":{"method":"cycle","colors":[["accent1",["shade",60000]],["accent1",["tint",50000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"cycle","colors":[["accent1",["shade",60000]],["accent1",["tint",50000]]]}},"accent1_5":{"fill":{"method":"span","colors":[["accent1",["alpha",85000]],["accent1",["alpha",40000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent1",["alpha",85000]],["accent1",["alpha",40000]]]}},"accent2_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["accent2"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent2",["tint",50000]]]}},"accent2_2":{"fill":{"method":"repeat","colors":[["accent2"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent2",["tint",50000]]]}},"accent2_3":{"fill":{"method":"span","colors":[["accent2",["shade",70000]],["accent2",["tint",60000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent2",["shade",70000]],["accent2",["tint",60000]]]}},"accent2_4":{"fill":{"method":"cycle","colors":[["accent2",["shade",60000]],["accent2",["tint",50000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"cycle","colors":[["accent2",["shade",60000]],["accent2",["tint",50000]]]}},"accent2_5":{"fill":{"method":"span","colors":[["accent2",["alpha",85000]],["accent2",["alpha",40000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent2",["alpha",85000]],["accent2",["alpha",40000]]]}},"accent3_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["accent3"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent3",["tint",50000]]]}},"accent3_2":{"fill":{"method":"repeat","colors":[["accent3"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent3",["tint",50000]]]}},"accent3_3":{"fill":{"method":"span","colors":[["accent3",["shade",70000]],["accent3",["tint",60000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent3",["shade",70000]],["accent3",["tint",60000]]]}},"accent3_4":{"fill":{"method":"cycle","colors":[["accent3",["shade",60000]],["accent3",["tint",50000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"cycle","colors":[["accent3",["shade",60000]],["accent3",["tint",50000]]]}},"accent3_5":{"fill":{"method":"span","colors":[["accent3",["alpha",85000]],["accent3",["alpha",40000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent3",["alpha",85000]],["accent3",["alpha",40000]]]}},"accent4_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["accent4"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent4",["tint",50000]]]}},"accent4_2":{"fill":{"method":"repeat","colors":[["accent4"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent4",["tint",50000]]]}},"accent4_3":{"fill":{"method":"span","colors":[["accent4",["shade",70000]],["accent4",["tint",60000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent4",["shade",70000]],["accent4",["tint",60000]]]}},"accent4_4":{"fill":{"method":"cycle","colors":[["accent4",["shade",60000]],["accent4",["tint",50000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"cycle","colors":[["accent4",["shade",60000]],["accent4",["tint",50000]]]}},"accent4_5":{"fill":{"method":"span","colors":[["accent4",["alpha",85000]],["accent4",["alpha",40000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent4",["alpha",85000]],["accent4",["alpha",40000]]]}},"accent5_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["accent5"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent5",["tint",50000]]]}},"accent5_2":{"fill":{"method":"repeat","colors":[["accent5"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent5",["tint",50000]]]}},"accent5_3":{"fill":{"method":"span","colors":[["accent5",["shade",70000]],["accent5",["tint",60000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent5",["shade",70000]],["accent5",["tint",60000]]]}},"accent5_4":{"fill":{"method":"cycle","colors":[["accent5",["shade",60000]],["accent5",["tint",50000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"cycle","colors":[["accent5",["shade",60000]],["accent5",["tint",50000]]]}},"accent5_5":{"fill":{"method":"span","colors":[["accent5",["alpha",85000]],["accent5",["alpha",40000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent5",["alpha",85000]],["accent5",["alpha",40000]]]}},"accent6_1":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["accent6"]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent6",["tint",50000]]]}},"accent6_2":{"fill":{"method":"repeat","colors":[["accent6"]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"repeat","colors":[["accent6",["tint",50000]]]}},"accent6_3":{"fill":{"method":"span","colors":[["accent6",["shade",70000]],["accent6",["tint",60000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent6",["shade",70000]],["accent6",["tint",60000]]]}},"accent6_4":{"fill":{"method":"cycle","colors":[["accent6",["shade",60000]],["accent6",["tint",50000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"cycle","colors":[["accent6",["shade",60000]],["accent6",["tint",50000]]]}},"accent6_5":{"fill":{"method":"span","colors":[["accent6",["alpha",85000]],["accent6",["alpha",40000]]]},"line":{"method":"repeat","colors":[["lt1"]]},"text":null,"arrow":{"method":"span","colors":[["accent6",["alpha",85000]],["accent6",["alpha",40000]]]}}};

const NAME_BY_SUFFIX = {
  1: "coloredOutline",
  2: "coloredFill",
  3: "gradientRange",
  4: "gradientLoop",
  5: "transparentGradientRange",
} as const;

const PRIMARY: readonly SmartArtColorVariant[] = (
  [
    ["accent0_1", "dark1Outline"],
    ["accent0_2", "dark2Outline"],
    ["accent0_3", "dark2Fill"],
  ] as const
).map(
  ([id, name]): SmartArtColorVariant => ({ id, group: "primary", name, definition: DEFINITIONS[id] }),
);

const COLORFUL: readonly SmartArtColorVariant[] = [
  { id: "colorful1", group: "colorful", name: "colorful1", definition: DEFINITIONS.colorful1 },
  ...[2, 3, 4, 5].map(
    (from): SmartArtColorVariant => ({
      id: `colorful${from}`,
      group: "colorful",
      name: "colorfulRange",
      from,
      to: from + 1,
      definition: DEFINITIONS[`colorful${from}`],
    }),
  ),
];

/** Each accent's five variations, in PowerPoint's order (the definitions' `pri`). */
const ACCENT_VARIANTS: readonly SmartArtColorVariant[] = [1, 2, 3, 4, 5, 6].flatMap(
  (accent) =>
    ([1, 2, 3, 4, 5] as const).map(
      (suffix): SmartArtColorVariant => ({
        id: `accent${accent}_${suffix}`,
        group: `accent${accent}` as SmartArtColorGroupId,
        name: NAME_BY_SUFFIX[suffix],
        accent,
        definition: DEFINITIONS[`accent${accent}_${suffix}`],
      }),
    ),
);

/**
 * PowerPoint's whole Change Colors gallery: 38 variations, in its order --
 * Primary Theme Colors, Colorful, then Accent 1 to 6 with Colored Outline,
 * Colored Fill, Gradient Range, Gradient Loop and Transparent Gradient Range.
 */
export const OFFICE_SMARTART_POWERPOINT_COLOR_VARIANTS: readonly SmartArtColorVariant[] =
  Object.freeze([...PRIMARY, ...COLORFUL, ...ACCENT_VARIANTS]);

/**
 * The variations the Docs and worksheet galleries offer: every one but
 * Gradient Loop and Transparent Gradient Range, in gallery order. The deck
 * offers PowerPoint's whole gallery ({@link OFFICE_SMARTART_POWERPOINT_COLOR_VARIANTS});
 * the other two hosts keep the list they shipped with until their own owners
 * take the rest (a translucent fill is not a field every host writes yet).
 */
export const OFFICE_SMARTART_COLOR_VARIANTS: readonly SmartArtColorVariant[] =
  Object.freeze(
    OFFICE_SMARTART_POWERPOINT_COLOR_VARIANTS.filter(
      (variant) => variant.name !== "gradientLoop" && variant.name !== "transparentGradientRange",
    ),
  );

export type SmartArtColorGroup = {
  readonly id: SmartArtColorGroupId;
  readonly accent?: number;
  readonly variants: readonly SmartArtColorVariant[];
};

function groupsOf(variants: readonly SmartArtColorVariant[]): readonly SmartArtColorGroup[] {
  return Object.freeze(
    (["primary", "colorful", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const).map(
      (id) => ({
        id,
        ...(id.startsWith("accent") ? { accent: Number(id.slice(6)) } : {}),
        variants: variants.filter((variant) => variant.group === id),
      }),
    ),
  );
}

/** The gallery's groups, in Office's order (the Docs / worksheet list). */
export const OFFICE_SMARTART_COLOR_GROUPS: readonly SmartArtColorGroup[] =
  groupsOf(OFFICE_SMARTART_COLOR_VARIANTS);

/** PowerPoint's whole gallery, grouped: Primary Theme Colors, Colorful, Accent 1..6. */
export const OFFICE_SMARTART_POWERPOINT_COLOR_GROUPS: readonly SmartArtColorGroup[] =
  groupsOf(OFFICE_SMARTART_POWERPOINT_COLOR_VARIANTS);

/**
 * Office's colour variations for its ICON layouts (`…/2018/5/colors/
 * Iconchunking_…`): PowerPoint's Change Colors gallery does not list them --
 * Designer and Convert to SmartArt write them with the icon layouts -- so no
 * gallery offers them either, but a converted icon diagram keeps the one it
 * wears, painted from its definition (`ms-definitions.json`) and written
 * back under its own urn (the server's `ICON_COLOR_VARIATIONS`). Each
 * definition below is the part's `node1` (fill, line, words) and
 * `sibTrans2D1` (arrows): the circles are `bgShp` (accent 2 to 6, or the
 * background), which only the definition-drawn scene paints.
 */
const ICON_ACCENTS: readonly SchemeColor[] = [["accent2"], ["accent3"], ["accent4"], ["accent5"], ["accent6"]];
const CLEAR_LINE: ColorList = { method: "repeat", colors: [["lt1", ["alpha", 0]]] };
export const OFFICE_SMARTART_ICON_COLOR_VARIANTS: readonly SmartArtColorVariant[] = /* own data: OFFICE_SMARTART_ICON_COLOR_VARIANTS */ [{"id":"Iconchunking_neutralicon_colorful1","group":"colorful","name":"colorful1","definition":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["lt1",["alpha",0]]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent1"],["accent2"],["accent3"],["accent4"],["accent5"]]}}},{"id":"Iconchunking_neutralicontext_colorful1","group":"colorful","name":"colorful1","definition":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["lt1",["alpha",0]]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"repeat","colors":[["accent1"],["accent2"],["accent3"],["accent4"],["accent5"]]}}},{"id":"Iconchunking_neutralicontext_colorful2","group":"colorful","name":"colorful1","definition":{"fill":{"method":"repeat","colors":[["lt1"]]},"line":{"method":"repeat","colors":[["lt1",["alpha",0]]]},"text":{"method":"repeat","colors":[["dk1"]]},"arrow":{"method":"span","colors":[["accent1"],["accent2"]]}}},{"id":"Iconchunking_neutralbg_colorful1","group":"colorful","name":"colorful1","definition":{"fill":{"method":"repeat","colors":[["accent1"],["accent2"],["accent3"],["accent4"],["accent5"]]},"line":{"method":"repeat","colors":[["lt1",["alpha",0]]]},"text":null,"arrow":{"method":"repeat","colors":[["accent1"],["accent2"],["accent3"],["accent4"],["accent5"]]}}}];

/**
 * One variation by Office's id. `all` admits PowerPoint's whole gallery and
 * the icon variations a converted diagram may wear; the default is the list
 * the Docs and worksheet commands accept.
 */
export function smartArtColorVariant(
  id: unknown,
  options?: { readonly all?: boolean },
): SmartArtColorVariant | null {
  const list = options?.all
    ? OFFICE_SMARTART_POWERPOINT_COLOR_VARIANTS
    : OFFICE_SMARTART_COLOR_VARIANTS;
  return (
    list.find((variant) => variant.id === id) ??
    (options?.all ? OFFICE_SMARTART_ICON_COLOR_VARIANTS.find((variant) => variant.id === id) : undefined) ??
    null
  );
}

/**
 * The variation id a colours urn names, as Office writes it in a data part's
 * `csTypeId` or a colours part's `uniqueId`: `…/colors/<id>`, where PowerPoint
 * appends `#N` to a definition whose content is not its built-in one (the
 * same colours, a different byte order). Null for anything that is not one of
 * the 38.
 */
export function smartArtColorVariantIdFromUrn(urn: unknown): string | null {
  if (typeof urn !== "string") return null;
  const matched = /\/colors\/([A-Za-z0-9_]+)(?:#\d+)?$/.exec(urn.trim());
  const id = matched?.[1] ?? (/^[A-Za-z0-9_]+$/.test(urn) ? urn : null);
  return id && smartArtColorVariant(id, { all: true }) ? id : null;
}

function normalizedTheme(theme?: Partial<SmartArtColorTheme> | null): SmartArtColorTheme {
  const hex = (value: unknown, fallback: string) =>
    typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : fallback;
  const base = OFFICE_SMARTART_COLOR_THEME;
  return {
    dk1: hex(theme?.dk1, base.dk1),
    lt1: hex(theme?.lt1, base.lt1),
    dk2: hex(theme?.dk2, base.dk2),
    lt2: hex(theme?.lt2, base.lt2),
    accents: base.accents.map((fallback, index) => hex(theme?.accents?.[index], fallback)),
  };
}

/**
 * One resolved list entry: its channels (gamma-encoded, NOT rounded to a hex
 * yet -- Office blends the exact values: Accent 2's `shade 50000` and `tint
 * 45000` are hue 357.1° and 359.2°, a +2.1° step, while their rounded hexes
 * would put the tint at hue 0°) and its opacity in percent.
 */
interface ResolvedColor {
  readonly channels: OfficeColorChannels;
  readonly alpha: number;
}

const hexOf = (color: ResolvedColor) => officeColorHexFromChannels(color.channels);

function resolveScheme(color: SchemeColor, theme: SmartArtColorTheme): ResolvedColor {
  const [scheme, ...mods] = color;
  const base = scheme.startsWith("accent")
    ? theme.accents[Number(scheme.slice(6)) - 1]
    : theme[scheme as "dk1" | "lt1" | "dk2" | "lt2"];
  let channels: OfficeColorChannels = officeColorChannelsFromHex(base) ?? [0, 0, 0];
  let alpha = 100;
  for (const [kind, value] of mods) {
    if (kind === "alpha") alpha = value / 1000;
    else channels = officeColorTransformChannels(channels, kind, value) ?? channels;
  }
  return { channels, alpha };
}

/**
 * Two list colours blended at `t` the way Office's own cache states a blend:
 * the first colour plus `hueOff` / `satOff` / `lumOff` (and `alphaOff`) of
 * `t` times the HSL difference -- measured on Office's cached drawings, where
 * Gradient Loop's 9 shapes on accent `#DA1F28` carry lumOff 0, 11174, 22348 …
 * (2/9 of the 50.28 % lightness gap per step) and Colorful Range's hue moves
 * by the RAW difference (+233.9°, not the short way round).
 */
function blend(from: ResolvedColor, to: ResolvedColor, t: number): ResolvedColor {
  if (t === 0) return from;
  const [h1, l1, s1] = rgbToHls(...from.channels);
  const [h2, l2, s2] = rgbToHls(...to.channels);
  let channels = from.channels;
  for (const [kind, value] of [
    ["hueOff", Math.round((h2 - h1) * 21_600_000 * t)],
    ["satOff", Math.round((s2 - s1) * 100_000 * t)],
    ["lumOff", Math.round((l2 - l1) * 100_000 * t)],
  ] as const) {
    channels = officeColorTransformChannels(channels, kind, value) ?? channels;
  }
  return { channels, alpha: from.alpha + (to.alpha - from.alpha) * t };
}

/** Item `index` of `count` under one colour list (`ST_ClrAppMethod`). */
function listColor(
  list: ColorList,
  index: number,
  count: number,
  theme: SmartArtColorTheme,
): ResolvedColor {
  const colors = list.colors.map((color) => resolveScheme(color, theme));
  if (colors.length === 1) return colors[0];
  if (list.method === "repeat") return colors[index % colors.length];
  // Span: first to last across the shapes, t = i / (n − 1) (Office's caches:
  // the last of five shapes carries the whole difference). Cycle: out and
  // back, t = 2·min(i, n − i) / n (nine shapes: 0,1,2,3,4,4,3,2,1 steps of 2/9).
  const t =
    list.method === "cycle"
      ? count <= 1
        ? 0
        : (2 * Math.min(index, count - index)) / count
      : count <= 1
        ? 0
        : index / (count - 1);
  const position = t * (colors.length - 1);
  const lower = Math.min(colors.length - 2, Math.floor(position));
  return blend(colors[lower], colors[lower + 1], position - lower);
}

/** One `a:schemeClr` of a colour list read off Office's own colours part: its slot and members. */
export type SmartArtSchemeColor = { scheme: string; mods: readonly (readonly [kind: string, value: number])[] };

const SCHEME_ALIASES: Readonly<Record<string, SchemeName>> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };

/**
 * Shape `index` of `count` under one colour list of Office's own colours part
 * (`dgm:fillClrLst` / `linClrLst` / `txFillClrLst` and its `meth`), resolved
 * on `theme` exactly as a variation's lists are: Office's colour transforms
 * and its span / cycle / repeat blending. Null for an empty list or a slot
 * the theme has no colour for.
 */
export function resolveSmartArtListColor(
  method: string | undefined,
  colors: readonly SmartArtSchemeColor[],
  index: number,
  count: number,
  theme?: Partial<SmartArtColorTheme> | null,
): { color: string; alpha: number } | null {
  const scheme = (name: string): SchemeName | null =>
    SCHEME_ALIASES[name] ?? (/^(dk1|lt1|dk2|lt2|accent[1-6])$/.test(name) ? (name as SchemeName) : null);
  const resolved: SchemeColor[] = [];
  for (const color of colors) {
    const slot = scheme(color.scheme);
    if (!slot) return null;
    resolved.push([slot, ...(color.mods as unknown as ColorMod[])]);
  }
  if (!resolved.length) return null;
  const list: ColorList = {
    method: method === "cycle" || method === "repeat" ? method : "span",
    colors: resolved,
  };
  const picked = listColor(list, Math.max(0, index), Math.max(1, count), normalizedTheme(theme));
  return { color: hexOf(picked), alpha: picked.alpha / 100 };
}

/**
 * One variation, resolved for a diagram of `itemCount` items: the diagram's
 * colour set (the first shape's fill, the line, the arrows' colour and the
 * text) and each item's own fill and line -- and its opacity, for the
 * Transparent Gradient Range variations, whose fills are translucent.
 */
export function resolveSmartArtColorVariant(
  variant: SmartArtColorVariant,
  itemCount: number,
  theme?: Partial<SmartArtColorTheme> | null,
): { colors: SmartArtStyle; nodePaint: SmartArtNodePaint[] } {
  const resolved = normalizedTheme(theme);
  const count = Math.max(1, Math.floor(itemCount));
  const { definition } = variant;
  const line = hexOf(listColor(definition.line, 0, 1, resolved));
  const nodePaint = Array.from({ length: count }, (_, index): SmartArtNodePaint => {
    const fill = listColor(definition.fill, index, count, resolved);
    return {
      fill: hexOf(fill),
      stroke: line,
      ...(fill.alpha < 100 ? { opacity: Math.round(fill.alpha * 100) / 100 } : {}),
    };
  });
  return {
    colors: {
      nodeFill: nodePaint[0].fill!,
      nodeStroke: line,
      connector: hexOf(listColor(definition.arrow, 0, count, resolved)),
      textColor: definition.text
        ? hexOf(listColor(definition.text, 0, 1, resolved))
        : resolved.lt1,
      fontSize: DEFAULT_SMARTART_STYLE.fontSize,
    },
    nodePaint,
  };
}

/**
 * The colours a NEW diagram takes in a document on `theme`: Office inserts
 * SmartArt in "Colored Fill - Accent 1" (`accent1_2`), resolved against the
 * document's own theme -- a deck on the Office theme gets Dark Teal, not the
 * 2013 palette's blue. Omitted, the editor's historical default.
 */
export function defaultSmartArtColors(
  theme?: Partial<SmartArtColorTheme> | null,
): SmartArtStyle {
  const variant = smartArtColorVariant("accent1_2");
  return variant && theme
    ? resolveSmartArtColorVariant(variant, 1, theme).colors
    : DEFAULT_SMARTART_STYLE;
}

/** The swatch colours a gallery cell draws for one variation (three shapes). */
export function smartArtColorVariantSwatch(
  variant: SmartArtColorVariant,
  theme?: Partial<SmartArtColorTheme> | null,
): { fills: string[]; line: string; arrow: string; opacities?: number[] } {
  const { colors, nodePaint } = resolveSmartArtColorVariant(variant, 3, theme);
  const opacities = nodePaint.map((paint) => paint.opacity ?? 100);
  return {
    fills: nodePaint.map((paint) => paint.fill!),
    line: colors.nodeStroke,
    arrow: colors.connector,
    ...(opacities.some((value) => value < 100) ? { opacities } : {}),
  };
}
