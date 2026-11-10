/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The generated Office theme ladder table (`presentation-office-body-ladder.json`,
 * MRAI-corpus/ppt-audit/L17/scripts/build_body_ladder.py, from the server's
 * theme catalog) decoded: per Office theme, variant, MRAI layout and
 * placeholder slot, the five outline levels a placeholder paragraph inherits
 * from its layout and master, and which run attributes the placeholder's own
 * list style states (those a master text-style edit does not reach).
 *
 * Dependency-free apart from the JSON so the layout module can read it.
 */

import source from "./presentation-office-body-ladder.json";

/** One outline level of an inherited ladder. EMU for indents, points for sizes. */
export interface OfficeBodyLevel {
  readonly marL: number;
  readonly indent: number;
  /** `a:defRPr@sz` in points. */
  readonly sizePt?: number;
  /** Space before / after, in points (`a:spcPts`). */
  readonly spaceBeforePt?: number;
  readonly spaceAfterPt?: number;
  /** `a:lnSpc/a:spcPct` as a multiple (0.9 = 90 %). */
  readonly lineSpacing?: number;
  /** `a:lnSpc/a:spcPts` in points. */
  readonly lineSpacingPt?: number;
  /** The level's marker: a glyph (with its face), a numbering scheme, or none. */
  readonly marker?:
    | { readonly char: string; readonly typeface?: string }
    | { readonly autoNum: string };
  /** `a:buClr` as a scheme slot (after the layout's colour map) or `#rrggbb`. */
  readonly markerColor?: string;
  readonly markerColorMods?: ReadonlyArray<readonly [string, number]>;
  /** `a:buSzPct`, thousandths of a percent. */
  readonly markerSizePct?: number;
}

type RawLevel = ReadonlyArray<unknown>;

type RawRect = [number, number, number, number];

interface RawTheme {
  ladders: number[][];
  layouts: Record<string, Array<number | null>>;
  own: Record<string, string[][]>;
  xfrm: Record<string, number[]>;
  master: { title?: RawRect; body?: RawRect };
  variants?: Record<
    string,
    {
      layouts?: Record<string, Array<number | null>>;
      own?: Record<string, string[][]>;
      xfrm?: Record<string, number[]>;
      master?: { title?: RawRect; body?: RawRect };
    }
  >;
}

const catalog = source as unknown as {
  levels: RawLevel[];
  themes: Record<string, RawTheme>;
};

function spacingPoints(raw: unknown): number | undefined {
  return typeof raw === "number" && raw > 0 ? raw / 100 : undefined;
}

function decodeLevel(raw: RawLevel): OfficeBodyLevel {
  const [marL, indent, sz, , before, after, line, bu, buFont, buClr, buClrMods, buSzPct, buAutoNum] = raw;
  const level: {
    -readonly [K in keyof OfficeBodyLevel]: OfficeBodyLevel[K];
  } = {
    marL: typeof marL === "number" ? marL : 0,
    indent: typeof indent === "number" ? indent : 0,
  };
  if (typeof sz === "number" && sz > 0) level.sizePt = sz / 100;
  const spaceBeforePt = spacingPoints(before);
  if (spaceBeforePt !== undefined) level.spaceBeforePt = spaceBeforePt;
  const spaceAfterPt = spacingPoints(after);
  if (spaceAfterPt !== undefined) level.spaceAfterPt = spaceAfterPt;
  if (typeof line === "number" && line < 0) level.lineSpacing = -line / 100_000;
  else if (typeof line === "number" && line > 0) level.lineSpacingPt = line / 100;
  // A blank glyph (some themes state `char=" "`) draws nothing either.
  if (typeof bu === "string" && bu.trim()) {
    level.marker =
      typeof buFont === "string" && buFont ? { char: bu, typeface: buFont } : { char: bu };
  } else if (typeof buAutoNum === "string" && buAutoNum) {
    level.marker = { autoNum: buAutoNum };
  }
  if (typeof buClr === "string" && buClr) level.markerColor = buClr;
  if (Array.isArray(buClrMods)) {
    level.markerColorMods = buClrMods.flatMap((entry) =>
      Array.isArray(entry) && typeof entry[0] === "string" && typeof entry[1] === "number"
        ? [[entry[0], entry[1]] as const]
        : [],
    );
  }
  if (typeof buSzPct === "number" && buSzPct > 0) level.markerSizePct = buSzPct;
  return Object.freeze(level);
}

const LEVELS: readonly OfficeBodyLevel[] = catalog.levels.map(decodeLevel);

/**
 * The ladder of one slot of an MRAI layout under an Office theme variant, in
 * the editor's pour order (the same slots `presentationOfficeLayoutFrames`
 * returns), or null for a title slot / anything the catalog lacks.
 */
export function officeBodyLadder(
  themeId: string,
  variantId: string,
  layoutId: string,
  slot: number,
): readonly OfficeBodyLevel[] | null {
  const theme = Object.prototype.hasOwnProperty.call(catalog.themes, themeId)
    ? catalog.themes[themeId]
    : undefined;
  if (!theme) return null;
  const slots = theme.variants?.[variantId]?.layouts?.[layoutId] ?? theme.layouts[layoutId];
  const index = slots?.[slot];
  if (typeof index !== "number") return null;
  const ladder = theme.ladders[index];
  if (!ladder) return null;
  const levels = ladder.map((position) => LEVELS[position]);
  return levels.every(Boolean) ? levels : null;
}

/** A run attribute a master text style states: size, bold, italic, colour. */
export type OfficeRunAttribute = "s" | "b" | "i" | "c";

/**
 * Per level (one for a title slot, five for a body slot), the run attributes
 * the slot's own placeholder list style states, which a master text-style
 * edit therefore does not reach. Empty sets where the catalog lacks the slot.
 */
export function officeSlotOwnAttributes(
  themeId: string,
  variantId: string,
  layoutId: string,
  slot: number,
): ReadonlyArray<ReadonlySet<OfficeRunAttribute>> {
  const theme = Object.prototype.hasOwnProperty.call(catalog.themes, themeId)
    ? catalog.themes[themeId]
    : undefined;
  const flags =
    theme?.variants?.[variantId]?.own?.[layoutId]?.[slot] ?? theme?.own?.[layoutId]?.[slot] ?? [];
  return flags.map(
    (entry) => new Set(entry.split("").filter((flag): flag is OfficeRunAttribute => "sbic".includes(flag))),
  );
}

function catalogTheme(themeId: string): RawTheme | undefined {
  return Object.prototype.hasOwnProperty.call(catalog.themes, themeId)
    ? catalog.themes[themeId]
    : undefined;
}

/**
 * Whether the slot's layout placeholder states its own position and size
 * (`a:xfrm` on the layout): a slot that does not inherits the master
 * placeholder's frame, so moving the master's title or body moves it.
 */
export function officeSlotOwnsGeometry(
  themeId: string,
  variantId: string,
  layoutId: string,
  slot: number,
): boolean {
  const theme = catalogTheme(themeId);
  const flags = theme?.variants?.[variantId]?.xfrm?.[layoutId] ?? theme?.xfrm?.[layoutId];
  return flags?.[slot] !== 0;
}

/** A frame in EMU on the 16:9 catalog slide (12192000 x 6858000). */
export interface OfficeCatalogRect {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/** The master's own title and body placeholder frames. */
export function officeMasterCatalogFrames(
  themeId: string,
  variantId: string,
): { title?: OfficeCatalogRect; body?: OfficeCatalogRect } {
  const theme = catalogTheme(themeId);
  const raw = theme?.variants?.[variantId]?.master ?? theme?.master ?? {};
  const rect = (value: RawRect | undefined) =>
    value ? { x: value[0], y: value[1], cx: value[2], cy: value[3] } : undefined;
  const result: { title?: OfficeCatalogRect; body?: OfficeCatalogRect } = {};
  const title = rect(raw.title);
  const body = rect(raw.body);
  if (title) result.title = title;
  if (body) result.body = body;
  return result;
}
