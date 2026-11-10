/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * The Theme Colors grid of an Office colour picker, built from a document's
 * OWN theme, and the theme reference each swatch stands for.
 *
 * Office lists ten theme colours -- Background 1, Text 1, Background 2,
 * Text 2, Accent 1..6 -- each above five tints or shades whose ladder
 * depends on how light the colour is (`officeThemeTintSteps`). A swatch is
 * not a colour but a REFERENCE: PowerPoint writes
 * `<a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/>`
 * for "Accent 1, Lighter 80%" (ISO/IEC 29500-1 §20.1.2.3.29 schemeClr,
 * §20.1.2.3.20/21 lumMod/lumOff), so the colour follows the theme when the
 * theme changes, while a Standard or custom colour stays put (Microsoft
 * Support, "Change the colors in a text box or shape").
 *
 * Lighter p% is `lumMod 1-p, lumOff p`; Darker p% is `lumMod 1-p`. The drawn
 * colour is the shared DrawingML transform (`officeTransformedHex`), the same
 * algorithm the server resolves the written reference with.
 */

import {
  normalizeHexColor,
  officeThemeTintSteps,
} from "./office-color-palette";
import { officeTransformedHex } from "./office-color-transform";

/** The ten picker columns, as DrawingML `a:schemeClr@val` names them. */
export const OFFICE_THEME_GRID_SLOTS = [
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
] as const;

export type OfficeThemeGridSlot = (typeof OFFICE_THEME_GRID_SLOTS)[number];

/** The theme's colour for each picker column (`bg1` = `lt1` under the
 * default colour map, `tx1` = `dk1`, `bg2` = `lt2`, `tx2` = `dk2`). */
export type OfficeThemeGridColors = Readonly<Record<OfficeThemeGridSlot, string>>;

/** A theme colour reference: the slot and its luminance transform, in
 * DrawingML's thousandths of a percent (`20000` = 20 %). */
export interface OfficeThemeColorRef {
  readonly slot: OfficeThemeGridSlot;
  readonly lumMod?: number;
  readonly lumOff?: number;
}

export interface OfficeThemeColorGrid {
  /** Ten columns, each the colour then its five tints/shades. */
  readonly columns: readonly (readonly string[])[];
  /** The reference behind each swatch, in the same shape as `columns`. */
  readonly refs: readonly (readonly OfficeThemeColorRef[])[];
  /** Each swatch's step: 0 for the colour, positive lighter, negative darker. */
  readonly steps: readonly (readonly number[])[];
}

const PERCENT = 100000;

export function isOfficeThemeGridSlot(value: unknown): value is OfficeThemeGridSlot {
  return (
    typeof value === "string" &&
    (OFFICE_THEME_GRID_SLOTS as readonly string[]).includes(value)
  );
}

/** The reference Office writes for one step of a slot's ladder. */
export function officeThemeStepRef(
  slot: OfficeThemeGridSlot,
  step: number,
): OfficeThemeColorRef {
  if (!step) return { slot };
  const amount = Math.round(Math.abs(step) * PERCENT);
  return step > 0
    ? { slot, lumMod: PERCENT - amount, lumOff: amount }
    : { slot, lumMod: PERCENT - amount };
}

/** The step a reference names (0 for the plain colour), or `null` when its
 * transform is not one a picker swatch writes. */
export function officeThemeRefStep(ref: OfficeThemeColorRef): number | null {
  const lumMod = ref.lumMod ?? PERCENT;
  const lumOff = ref.lumOff ?? 0;
  if (lumOff === 0) return lumMod === PERCENT ? 0 : -(PERCENT - lumMod) / PERCENT;
  return lumMod + lumOff === PERCENT ? lumOff / PERCENT : null;
}

/** The colour a reference draws in a theme. */
export function officeThemeRefHex(
  ref: OfficeThemeColorRef,
  colors: OfficeThemeGridColors,
): string | null {
  const base = normalizeHexColor(colors[ref.slot]);
  if (!base) return null;
  const transforms: (readonly [string, number])[] = [];
  if (ref.lumMod !== undefined && ref.lumMod !== PERCENT) {
    transforms.push(["lumMod", ref.lumMod]);
  }
  if (ref.lumOff) transforms.push(["lumOff", ref.lumOff]);
  if (!transforms.length) return base;
  return officeTransformedHex(base, transforms)?.toLowerCase() ?? null;
}

/** The whole Theme Colors grid of a theme. */
export function officeThemeColorGrid(
  colors: OfficeThemeGridColors,
): OfficeThemeColorGrid {
  const columns: string[][] = [];
  const refs: OfficeThemeColorRef[][] = [];
  const steps: number[][] = [];
  for (const slot of OFFICE_THEME_GRID_SLOTS) {
    const base = normalizeHexColor(colors[slot]) ?? "#000000";
    const ladder = [0, ...officeThemeTintSteps(base)];
    const columnRefs = ladder.map((step) => officeThemeStepRef(slot, step));
    refs.push(columnRefs);
    steps.push(ladder);
    columns.push(
      columnRefs.map((ref) => officeThemeRefHex(ref, { ...colors, [slot]: base }) ?? base),
    );
  }
  return { columns, refs, steps };
}

/** A stored reference, validated; `null` for anything else. */
export function normalizeOfficeThemeColorRef(
  value: unknown,
): OfficeThemeColorRef | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  if (!isOfficeThemeGridSlot(record.slot)) return null;
  const percent = (raw: unknown): number | undefined | null => {
    if (raw === undefined) return undefined;
    return typeof raw === "number" &&
      Number.isInteger(raw) &&
      raw >= 0 &&
      raw <= PERCENT
      ? raw
      : null;
  };
  const lumMod = percent(record.lumMod);
  const lumOff = percent(record.lumOff);
  if (lumMod === null || lumOff === null) return null;
  return {
    slot: record.slot,
    ...(lumMod !== undefined && lumMod !== PERCENT ? { lumMod } : {}),
    ...(lumOff ? { lumOff } : {}),
  };
}

export function sameOfficeThemeColorRef(
  left: OfficeThemeColorRef | null | undefined,
  right: OfficeThemeColorRef | null | undefined,
): boolean {
  if (!left || !right) return left === right;
  return (
    left.slot === right.slot &&
    (left.lumMod ?? PERCENT) === (right.lumMod ?? PERCENT) &&
    (left.lumOff ?? 0) === (right.lumOff ?? 0)
  );
}

/**
 * Office names a theme colour after the colour itself ("Dark Teal, Accent 1")
 * from a fixed list of 48 named colours. How a colour maps to a name is not
 * documented, so only colours whose names were read from PowerPoint's own
 * picker are named here: the Office theme of a new presentation (2023+) and
 * the 2013-2022 Office theme. Any other theme colour is named by its slot
 * alone ("Accent 1, Lighter 80%"), never by a guess.
 */
export const OFFICE_NAMED_THEME_COLORS: Readonly<Record<string, OfficeNamedColorId>> =
  /* own data: OFFICE_NAMED_THEME_COLORS */ {"#ffffff":"white","#000000":"black","#1f2a37":"blueGray","#eef1f4":"lightGray","#2f6fd6":"blue","#e07b39":"orange","#3a9d6b":"green","#2c9aa8":"turquoise","#2b3440":"blueGray","#e6e9ed":"lightGray","#4a6382":"blueGray","#8da4bf":"blueGray","#6e8b74":"gray","#b08a57":"orange","#5c7f8f":"turquoise","#12324a":"darkBlue","#e3eef4":"lightGray","#1c6e9c":"blue","#35a3c4":"turquoise","#f2a541":"orange","#4b8f5e":"green","#c9553d":"orange","#6b6fa8":"blueGray","#263a22":"darkGreen","#eef3e6":"lightGray","#5a9a3c":"green","#a7c957":"green","#d9a13b":"orange","#3c7f7a":"darkTeal","#b5654a":"orange","#f6ece4":"lightGray","#c8462b":"orange","#e8913a":"orange","#f2c14e":"gold","#5c6b73":"gray","#a0785a":"orange","#f2ecf5":"lightGray","#c76ba6":"plum","#5a7bc2":"blue","#e39a6b":"orange","#4fa39a":"turquoise","#b8a04a":"gold","#222222":"gray","#ededed":"lightGray","#4d4d4d":"gray","#7a7a7a":"gray","#a6a6a6":"gray","#fbf3e4":"lightGray","#f0a12c":"orange","#e5652e":"orange","#3d8fc7":"blue","#52a86b":"green"};

export const OFFICE_NAMED_COLOR_IDS = [
  "white",
  "black",
  "lightGray",
  "gray",
  "darkBlue",
  "blueGray",
  "blue",
  "darkTeal",
  "orange",
  "darkGreen",
  "turquoise",
  "plum",
  "green",
  "gold",
] as const;

export type OfficeNamedColorId = (typeof OFFICE_NAMED_COLOR_IDS)[number];

export function officeNamedThemeColor(hex: unknown): OfficeNamedColorId | null {
  const normalized = normalizeHexColor(hex);
  return normalized ? (OFFICE_NAMED_THEME_COLORS[normalized] ?? null) : null;
}

/**
 * Text Highlight Color: Office's fifteen highlighter colours, in the order
 * PowerPoint and Word list them (five per row), with their message-key ids.
 */
export const OFFICE_HIGHLIGHT_COLORS = [
  { id: "yellow", color: "#ffff00" },
  { id: "brightGreen", color: "#00ff00" },
  { id: "turquoise", color: "#00ffff" },
  { id: "pink", color: "#ff00ff" },
  { id: "blue", color: "#0000ff" },
  { id: "red", color: "#ff0000" },
  { id: "darkBlue", color: "#000080" },
  { id: "teal", color: "#008080" },
  { id: "green", color: "#008000" },
  { id: "violet", color: "#800080" },
  { id: "darkRed", color: "#800000" },
  { id: "darkYellow", color: "#808000" },
  { id: "gray50", color: "#808080" },
  { id: "gray25", color: "#c0c0c0" },
  { id: "black", color: "#000000" },
] as const;

export type OfficeHighlightColorId = (typeof OFFICE_HIGHLIGHT_COLORS)[number]["id"];
